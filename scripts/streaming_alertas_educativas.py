# -*- coding: utf-8 -*-
"""
streaming_alertas_educativas.py

Módulo de Procesamiento de Eventos en Tiempo Real (Corte 3)
Sistema Integrado de Análisis y Calidad de Datos — Ecosistema Educativo Bogotá D.C.

Implementa los conceptos del Tercer Corte aplicados al monitoreo continuo de
inasistencia escolar, riesgo de deserción y alertas psicosociales en Bogotá:
  1. Arquitectura Orientada a Eventos (EDA) con formato CloudEvents 1.0,
     Productor, Broker particionado por localidad, Consumidor idempotente y DLQ.
  2. Diferenciación entre Event Time (cuándo ocurre el reporte en el colegio) y
     Processing Time (cuándo llega al servidor de la SED), demostrando el impacto
     de la latencia de red en colegios periféricos o rurales.
  3. Cálculo dinámico de Watermarks: WM = max(event_time) - umbral_tolerancia.
  4. Agrupación sobre flujos mediante tres tipos de Ventanas de Tiempo:
     - Tumbling Windows (10 min, sin traslape): consolidado periódico por localidad.
     - Sliding / Hopping Windows (10 min cada 2 min): detección de ráfagas críticas.
     - Session Windows (gap de 4 min): sesiones activas de reporte por colegio (DANE).
  5. Tratamiento de Eventos Tardíos (Late Events): aceptación dentro del watermark,
     re-emisión por Allowed Lateness y desvío a Side Output / Dead Letter Queue (DLQ)
     para conciliación sin pérdida de alertas estudiantiles.

Autores:
    David Santiago Castelblanco Artunduaga (67001182) — Ventanas de tiempo (Tumbling,
        Sliding, Session), arquitectura de streaming e integración con capa Gold.
    Dairo José Narváez Morales (67002032) — Broker EDA idempotente, Event Time vs.
        Processing Time, calibración de Watermarks y manejo de eventos tardíos / DLQ.

Asignatura: Ingeniería de Datos — Universidad Católica de Colombia
Docente: Ing. Diana Marcela Toquica Rodríguez
Versión: v2.0.0 (Entrega Final — Corte 3)
Fecha: 22/11/2026
"""

import os
import sys
import json
import hashlib
from datetime import datetime, timedelta
import pandas as pd

# Asegurar codificación UTF-8 en consola de Windows
if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
if sys.stderr and hasattr(sys.stderr, "reconfigure"):
    try:
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

# ─────────────────────────────────────────────────────────────────────────────
# Rutas y constantes de configuración de Streaming
# ─────────────────────────────────────────────────────────────────────────────
BASE_DIR      = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RAW_DIR       = os.path.join(BASE_DIR, "data", "RAW")
STAGING_DIR   = os.path.join(BASE_DIR, "data", "STAGING")
ANALYTICS_DIR = os.path.join(BASE_DIR, "data", "ANALYTICS")

ARCHIVO_STREAM_JSONL = os.path.join(RAW_DIR, "eventos_asistencia_desercion_stream.jsonl")

# Parámetros temporales del motor de streaming
UMBRAL_WATERMARK_MIN: int   = 5   # Tolerancia del watermark: 5 minutos
ALLOWED_LATENESS_MIN: int   = 5   # Tolerancia adicional (Allowed Lateness) tras cierre
TUMBLING_SIZE_MIN: int      = 10  # Ventana fija de 10 minutos sin traslape
SLIDING_SIZE_MIN: int       = 10  # Ventana deslizante de 10 minutos
SLIDING_SLIDE_MIN: int      = 2   # Avance (slide) cada 2 minutos -> 5 ventanas/evento
SESSION_GAP_MIN: int        = 4   # Brecha de inactividad de 4 minutos por colegio
UMBRAL_RAFAGA_CRITICA: int  = 3   # 3 o más eventos críticos en ventana sliding dispara alerta


# ─────────────────────────────────────────────────────────────────────────────
# 1. Arquitectura Orientada a Eventos (EDA): Broker, Idempotencia y CloudEvents
# ─────────────────────────────────────────────────────────────────────────────

class BrokerEventosEducativos:
    """
    Simula un Broker distribuido orientado a eventos (estilo Apache Kafka) con:
      - Tópicos particionados por código de localidad (garantiza orden por partición).
      - Registro inmutable (append-only log) con offsets secuenciales.
      - Control de idempotencia en el consumidor para descartar eventos duplicados
        producto de reintentos de red en colegios con conectividad inestable.
      - Cola de mensajes muertos (Dead Letter Queue - DLQ) / Side Output para
        eventos que llegan después del cierre definitivo de su ventana.
    """

    def __init__(self):
        self.topicos: dict[str, list[dict]] = {
            "sed.alertas.asistencia": [],
            "sed.alertas.dlq": [],
        }
        self.suscriptores: dict[str, list] = {
            "sed.alertas.asistencia": [],
            "sed.alertas.dlq": [],
        }
        self.ids_procesados: set[str] = set()
        self.duplicados_ignorados: list[dict] = []

    def suscribir(self, topico: str, consumidor_fn) -> None:
        """Registra una función consumidora en un tópico específico."""
        if topico not in self.suscriptores:
            self.topicos[topico] = []
            self.suscriptores[topico] = []
        self.suscriptores[topico].append(consumidor_fn)

    def publicar(self, topico: str, evento_cloudevent: dict) -> bool:
        """
        Publica un evento inmutable en el tópico indicado y notifica a los
        consumidores aplicando verificación de idempotencia por event ID.
        """
        if topico not in self.topicos:
            self.topicos[topico] = []
            self.suscriptores[topico] = []

        event_id = evento_cloudevent.get("id", "")
        # Verificación de idempotencia (evita doble conteo ante reintentos HTTP/TCP)
        if topico == "sed.alertas.asistencia" and event_id in self.ids_procesados:
            self.duplicados_ignorados.append(evento_cloudevent)
            return False

        if topico == "sed.alertas.asistencia":
            self.ids_procesados.add(event_id)

        offset = len(self.topicos[topico])
        evento_con_meta = {**evento_cloudevent, "_offset": offset, "_topico": topico}
        self.topicos[topico].append(evento_con_meta)

        for consumidor in self.suscriptores.get(topico, []):
            consumidor(evento_con_meta)
        return True


def anonimizar_estudiante(codigo_est: str) -> str:
    """
    Aplica hash SHA-256 truncado al identificador del estudiante para cumplir
    con la Ley 1581 de 2012 (Habeas Data y principio de minimización en menores).
    """
    return "EST-" + hashlib.sha256(codigo_est.encode("utf-8")).hexdigest()[:10].upper()


def generar_flujo_eventos_educativos() -> list[dict]:
    """
    Construye y persiste en data/RAW/eventos_asistencia_desercion_stream.jsonl
    un flujo realista de eventos educativos bajo el estándar CloudEvents 1.0 (CNCF).

    Incluye casos reales de operación en Bogotá D.C.:
      - Eventos en orden con baja latencia (fibra óptica urbana).
      - Ráfagas de alertas en BOSA y CIUDAD BOLÍVAR para pruebas de Sliding Windows.
      - Sesiones continuas de reporte de orientación escolar en colegios DANE.
      - Eventos fuera de orden (Out-of-Order) que llegan dentro del Watermark (5 min).
      - Eventos tardíos moderados cubiertos por Allowed Lateness (+5 min adicionales).
      - Eventos extremadamente tardíos (ej. colegio rural en SUMAPAZ / USME tras
        caída de enlace satelital de 28 minutos) enviados a Side Output / DLQ.
      - Un evento duplicado por reintento de red para demostrar idempotencia.
    """
    os.makedirs(RAW_DIR, exist_ok=True)

    # Catálogo base de eventos (event_time vs processing_time entre 08:00 y 08:35)
    # Formato: (id, event_time_str, proc_time_str, cod_loc, localidad, cod_dane, tipo, est_raw, dias_inasist, score_riesgo, causal)
    especificaciones = [
        # --- Ventana [08:00, 08:10) ---
        ("EVT-001", "2026-11-22 08:01:10", "2026-11-22 08:01:15", 19, "CIUDAD BOLÍVAR", "110010190012",
         "alerta.riesgo_desercion", "1001", 6, 88.5, "Inasistencia reiterada y vulnerabilidad económica"),
        ("EVT-002", "2026-11-22 08:02:30", "2026-11-22 08:02:40", 19, "CIUDAD BOLÍVAR", "110010190012",
         "solicitud.apoyo_psicosocial", "1002", 5, 84.0, "Alerta familiar reportada por orientación escolar"),
        ("EVT-003", "2026-11-22 08:04:15", "2026-11-22 08:04:22", 19, "CIUDAD BOLÍVAR", "110010190012",
         "alerta.riesgo_desercion", "1003", 8, 92.0, "Riesgo alto de abandono escolar en educación media"),
        ("EVT-004", "2026-11-22 08:05:00", "2026-11-22 08:05:10", 7, "BOSA", "110010070045",
         "inasistencia.critica", "1004", 4, 76.0, "Cuatro días consecutivos sin registro biométrico"),
        ("EVT-005", "2026-11-22 08:06:20", "2026-11-22 08:06:35", 13, "TEUSAQUILLO", "110010130088",
         "actualizacion.matricula_simat", "1005", 1, 22.0, "Actualización preventiva de datos de acudiente"),
        ("EVT-006", "2026-11-22 08:08:40", "2026-11-22 08:08:50", 7, "BOSA", "110010070045",
         "alerta.riesgo_desercion", "1006", 7, 86.5, "Cambio abrupto de rendimiento e inasistencia"),

        # Reintento duplicado de red del EVT-004 (mismo ID, llega 30s después)
        ("EVT-004", "2026-11-22 08:05:00", "2026-11-22 08:05:40", 7, "BOSA", "110010070045",
         "inasistencia.critica", "1004", 4, 76.0, "Cuatro días consecutivos sin registro biométrico (REINTENTO)"),

        # --- Avanza el reloj a [08:10, 08:20) ---
        ("EVT-007", "2026-11-22 08:11:10", "2026-11-22 08:11:20", 7, "BOSA", "110010070045",
         "alerta.riesgo_desercion", "1007", 9, 94.0, "Tercer caso crítico en menos de 10 minutos en Bosa"),
        ("EVT-008", "2026-11-22 08:13:00", "2026-11-22 08:13:10", 9, "FONTIBÓN", "110010090031",
         "alerta.riesgo_desercion", "1008", 6, 81.5, "Deserción potencial por traslado familiar"),

        # Evento FUERA DE ORDEN pero DENTRO DEL WATERMARK:
        # Ocurrió a las 08:09:10 (pertenece a [08:00, 08:10)), llega a las 08:13:30.
        # En ese instante max(event_time)=08:13:00 -> WM = 08:08:00 (< 08:10:00), ventana sigue abierta.
        ("EVT-009", "2026-11-22 08:09:10", "2026-11-22 08:13:30", 5, "USME", "110010050019",
         "alerta.riesgo_desercion", "1009", 7, 89.0, "Reporte retenido 4.3 min en buffer local por lluvia en Usme"),

        # Evento que hace avanzar el Event Time a 08:16:00 -> WM = 08:11:00 (CIERRA [08:00, 08:10))
        ("EVT-010", "2026-11-22 08:16:00", "2026-11-22 08:16:08", 8, "KENNEDY", "110010080077",
         "inasistencia.critica", "1010", 5, 79.0, "Inasistencia en jornada tarde reportada al cierre de bloque"),
        ("EVT-011", "2026-11-22 08:17:30", "2026-11-22 08:17:40", 8, "KENNEDY", "110010080077",
         "solicitud.apoyo_psicosocial", "1011", 6, 83.0, "Solicitud de visita domiciliaria en Kennedy"),
        ("EVT-012", "2026-11-22 08:19:00", "2026-11-22 08:19:12", 8, "KENNEDY", "110010080077",
         "alerta.riesgo_desercion", "1012", 8, 91.0, "Alerta crítica en colegio oficial de Kennedy"),

        # Evento TARDÍO MODERADO (cubierto por Allowed Lateness):
        # Ocurrió a las 08:07:30 (ventana [08:00, 08:10)), llega a las 08:19:30 cuando WM=08:14:00.
        # La ventana [08:00, 08:10) cerró en WM=08:10, pero está dentro de Allowed Lateness (hasta WM<=08:15).
        ("EVT-013", "2026-11-22 08:07:30", "2026-11-22 08:19:30", 5, "USME", "110010050019",
         "alerta.riesgo_desercion", "1013", 10, 95.0, "Sincronización tardía de tableta rural en Usme (12 min retraso)"),

        # --- Avanza el reloj a [08:20, 08:30) ---
        ("EVT-014", "2026-11-22 08:22:00", "2026-11-22 08:22:10", 11, "SUBA", "110010110054",
         "inasistencia.critica", "1014", 4, 72.0, "Alerta temprana en básica secundaria Suba"),
        ("EVT-015", "2026-11-22 08:26:30", "2026-11-22 08:26:40", 1, "USAQUÉN", "110010010022",
         "actualizacion.matricula_simat", "1015", 2, 31.0, "Seguimiento preventivo de asistencia"),
        ("EVT-016", "2026-11-22 08:29:00", "2026-11-22 08:29:05", 19, "CIUDAD BOLÍVAR", "110010190012",
         "alerta.riesgo_desercion", "1016", 7, 87.0, "Nueva sesión de reporte en IED Ciudad Bolívar tras pausa"),

        # Evento EXTREMADAMENTE TARDÍO (Enviado a Side Output / DLQ):
        # Ocurrió a las 08:03:00 (ventana [08:00, 08:10)), pero llega a las 08:31:00 (28 min tarde)
        # cuando max(event_time)=08:29:00 y WM=08:24:00 (superó cierre 08:10 + 5 min Allowed Lateness).
        ("EVT-017", "2026-11-22 08:03:00", "2026-11-22 08:31:00", 20, "SUMAPAZ", "110010200003",
         "alerta.riesgo_desercion", "1017", 12, 97.5, "Colegio rural Sumapaz sin enlace satelital durante 28 minutos"),
    ]

    eventos_cloudevents = []
    with open(ARCHIVO_STREAM_JSONL, "w", encoding="utf-8") as f:
        for (ev_id, et_str, pt_str, cod_loc, loc, dane, tipo, est_id, dias, score, causal) in especificaciones:
            evento = {
                "specversion": "1.0",
                "id": ev_id,
                "source": f"/sed/colegios/{dane}",
                "type": f"co.gov.educacionbogota.{tipo}",
                "event_time": et_str,
                "processing_time": pt_str,
                "partition_key": f"LOC-{cod_loc:02d}",
                "data": {
                    "cod_localidad": cod_loc,
                    "localidad": loc,
                    "cod_dane": dane,
                    "estudiante_hash": anonimizar_estudiante(est_id),
                    "tipo_evento": tipo,
                    "dias_inasistencia": dias,
                    "puntaje_riesgo": score,
                    "causal": causal,
                },
            }
            eventos_cloudevents.append(evento)
            f.write(json.dumps(evento, ensure_ascii=False) + "\n")

    return eventos_cloudevents


# ─────────────────────────────────────────────────────────────────────────────
# 2. Motor de Ventanas de Tiempo, Watermarks y Tratamiento de Eventos Tardíos
# ─────────────────────────────────────────────────────────────────────────────

def obtener_inicio_ventana_fija(dt: datetime, size_min: int) -> datetime:
    """Calcula el límite inferior de un intervalo semiabierto [inicio, fin) alineado al reloj."""
    minuto_base = (dt.minute // size_min) * size_min
    return dt.replace(minute=minuto_base, second=0, microsecond=0)


def procesar_flujo_con_watermarks_y_ventanas(eventos: list[dict]) -> dict:
    """
    Procesa secuencialmente el flujo de eventos aplicando:
      1. Descarte idempotente de duplicados en el Broker EDA.
      2. Actualización del Watermark: WM = max(event_time) - 5 min.
      3. Clasificación de cada evento según su marca temporal frente al Watermark:
         - ACEPTADO_EN_VENTANA: llega mientras WM < fin_ventana.
         - ALLOWED_LATENESS_CORRECCION: llega con fin_ventana <= WM <= fin_ventana + 5 min.
         - DESVIADO_A_DLQ_SIDE_OUTPUT: llega con WM > fin_ventana + 5 min.
      4. Cálculo comparativo Event Time vs. Processing Time en Tumbling Windows (10 min).
      5. Detección de ráfagas críticas en Sliding Windows (10 min cada 2 min).
      6. Agrupación dinámica de sesiones de reporte por colegio en Session Windows (gap 4 min).
    """
    broker = BrokerEventosEducativos()

    max_event_time: datetime | None = None
    watermark_actual: datetime | None = None

    traza_eventos = []
    eventos_validos_et = []
    eventos_dlq = []

    # State Store para Tumbling Windows por Event Time: clave = (inicio_str, fin_str, localidad)
    tumbling_et_store: dict[tuple, list[dict]] = {}
    # State Store para Tumbling Windows por Processing Time (para contrastar distorsión)
    tumbling_pt_store: dict[tuple, list[dict]] = {}

    def consumidor_principal(ev: dict) -> None:
        nonlocal max_event_time, watermark_actual

        et = datetime.strptime(ev["event_time"], "%Y-%m-%d %H:%M:%S")
        pt = datetime.strptime(ev["processing_time"], "%Y-%m-%d %H:%M:%S")
        retraso_min = round((pt - et).total_seconds() / 60.0, 2)

        # Inicio y fin de la ventana Tumbling [inicio, fin) a la que pertenece por Event Time
        w_start = obtener_inicio_ventana_fija(et, TUMBLING_SIZE_MIN)
        w_end = w_start + timedelta(minutes=TUMBLING_SIZE_MIN)

        # Evaluar estado frente al Watermark vigente en el instante de llegada
        wm_previo = watermark_actual
        if wm_previo is None or wm_previo < w_end:
            estado_arrivo = "ACEPTADO_EN_VENTANA"
            accion_tomada = "Acumulado en State Store de ventana abierta"
        elif wm_previo <= w_end + timedelta(minutes=ALLOWED_LATENESS_MIN):
            estado_arrivo = "ALLOWED_LATENESS"
            accion_tomada = "Ventana reabierta en memoria; se re-emite versión corregida"
        else:
            estado_arrivo = "RECHAZADO_EXTREMO_DLQ"
            accion_tomada = "Spark lo descartaría (Drop); desviado a Side Output / DLQ para conciliación SED"

        # Actualizar max_event_time y hacer avanzar el Watermark de forma monótona creciente
        if max_event_time is None or et > max_event_time:
            max_event_time = et
            watermark_actual = max_event_time - timedelta(minutes=UMBRAL_WATERMARK_MIN)

        registro_traza = {
            "id": ev["id"],
            "localidad": ev["data"]["localidad"],
            "cod_dane": ev["data"]["cod_dane"],
            "tipo_evento": ev["data"]["tipo_evento"],
            "puntaje_riesgo": ev["data"]["puntaje_riesgo"],
            "event_time": ev["event_time"],
            "processing_time": ev["processing_time"],
            "retraso_min": retraso_min,
            "ventana_et": f"[{w_start.strftime('%H:%M')}, {w_end.strftime('%H:%M')})",
            "watermark_al_llegar": wm_previo.strftime("%H:%M:%S") if wm_previo else "08:00:00 (Inicial)",
            "nuevo_watermark": watermark_actual.strftime("%H:%M:%S"),
            "clasificacion_arrivo": estado_arrivo,
            "accion_motor": accion_tomada,
        }
        traza_eventos.append(registro_traza)

        # Siempre se contabiliza en Processing Time (así operaría un sistema ingenuo sin Event Time)
        pt_start = obtener_inicio_ventana_fija(pt, TUMBLING_SIZE_MIN)
        pt_end = pt_start + timedelta(minutes=TUMBLING_SIZE_MIN)
        k_pt = (pt_start.strftime("%H:%M"), pt_end.strftime("%H:%M"))
        tumbling_pt_store.setdefault(k_pt, []).append(ev)

        if estado_arrivo in ("ACEPTADO_EN_VENTANA", "ALLOWED_LATENESS"):
            k_et = (w_start.strftime("%H:%M"), w_end.strftime("%H:%M"), ev["data"]["localidad"])
            tumbling_et_store.setdefault(k_et, []).append(ev)
            eventos_validos_et.append((et, pt, ev, estado_arrivo))
        else:
            broker.publicar("sed.alertas.dlq", {**ev, "motivo_dlq": accion_tomada, "retraso_min": retraso_min})
            eventos_dlq.append(registro_traza)

    broker.suscribir("sed.alertas.asistencia", consumidor_principal)

    for ev in eventos:
        broker.publicar("sed.alertas.asistencia", ev)

    # ── A. Consolidar Tumbling Windows (10 min, sin traslape) por Event Time ──
    filas_tumbling = []
    for (ini, fin, loc), lista_ev in sorted(tumbling_et_store.items()):
        puntajes = [e["data"]["puntaje_riesgo"] for e in lista_ev]
        dias_ina = [e["data"]["dias_inasistencia"] for e in lista_ev]
        hubo_correccion = any(
            t["clasificacion_arrivo"] == "ALLOWED_LATENESS"
            for t in traza_eventos
            if t["ventana_et"] == f"[{ini}, {fin})" and t["localidad"] == loc
        )
        filas_tumbling.append({
            "ventana_tumbling": f"[{ini}, {fin})",
            "localidad": loc,
            "total_eventos": len(lista_ev),
            "riesgo_promedio": round(sum(puntajes) / len(puntajes), 2),
            "dias_inasistencia_prom": round(sum(dias_ina) / len(dias_ina), 1),
            "estado_emision": "CORREGIDA (Allowed Lateness)" if hubo_correccion else "CERRADA POR WATERMARK",
        })
    df_tumbling = pd.DataFrame(filas_tumbling)

    # ── B. Comparación Event Time vs. Processing Time por bloque de 10 min ──
    ventanas_10m = [("08:00", "08:10"), ("08:10", "08:20"), ("08:20", "08:30"), ("08:30", "08:40")]
    filas_comp_relojes = []
    for ini, fin in ventanas_10m:
        conteo_et = sum(
            len(v) for (i, f, _), v in tumbling_et_store.items() if i == ini and f == fin
        )
        conteo_pt = len(tumbling_pt_store.get((ini, fin), []))
        filas_comp_relojes.append({
            "ventana_10_min": f"[{ini}, {fin})",
            "conteo_event_time_real": conteo_et,
            "conteo_processing_time_servidor": conteo_pt,
            "diferencia_distorsion": conteo_pt - conteo_et,
        })
    df_comparacion_relojes = pd.DataFrame(filas_comp_relojes)

    # ── C. Sliding / Hopping Windows (10 min de ancho, avance cada 2 min) ──
    # Rango de evaluación: desde 07:52 hasta 08:30 cada 2 minutos
    t_base = datetime(2026, 11, 22, 8, 0, 0)
    inicios_sliding = [t_base + timedelta(minutes=2 * k) for k in range(-4, 13)]
    filas_sliding = []
    for s_start in inicios_sliding:
        s_end = s_start + timedelta(minutes=SLIDING_SIZE_MIN)
        por_localidad: dict[str, list[dict]] = {}
        for et, _, ev, _ in eventos_validos_et:
            if s_start <= et < s_end:
                loc = ev["data"]["localidad"]
                por_localidad.setdefault(loc, []).append(ev)
        for loc, evs in por_localidad.items():
            if len(evs) >= UMBRAL_RAFAGA_CRITICA:
                punt_prom = round(sum(e["data"]["puntaje_riesgo"] for e in evs) / len(evs), 2)
                filas_sliding.append({
                    "ventana_sliding": f"[{s_start.strftime('%H:%M')}, {s_end.strftime('%H:%M')})",
                    "localidad": loc,
                    "eventos_en_ventana": len(evs),
                    "riesgo_promedio": punt_prom,
                    "alerta_rafaga": f"🚨 RÁFAGA CRÍTICA ({len(evs)} alertas en 10 min) -> Despachar unidad móvil SED",
                })
    df_sliding = pd.DataFrame(filas_sliding)

    # ── D. Session Windows (brecha de inactividad de 4 min por colegio DANE) ──
    # Ordenamos los eventos válidos por Event Time dentro de cada colegio
    por_colegio: dict[str, list[tuple[datetime, dict]]] = {}
    for et, _, ev, _ in sorted(eventos_validos_et, key=lambda x: x[0]):
        dane = ev["data"]["cod_dane"]
        por_colegio.setdefault(dane, []).append((et, ev))

    filas_session = []
    for dane, lista in sorted(por_colegio.items()):
        sesiones_colegio = []
        sesion_actual = [lista[0]]
        for item in lista[1:]:
            et_actual = item[0]
            et_anterior = sesion_actual[-1][0]
            if (et_actual - et_anterior) <= timedelta(minutes=SESSION_GAP_MIN):
                sesion_actual.append(item)
            else:
                sesiones_colegio.append(sesion_actual)
                sesion_actual = [item]
        sesiones_colegio.append(sesion_actual)

        for idx, ses in enumerate(sesiones_colegio, start=1):
            ini_ses = ses[0][0]
            fin_ses = ses[-1][0] + timedelta(minutes=SESSION_GAP_MIN)
            duracion_min = round((fin_ses - ini_ses).total_seconds() / 60.0, 2)
            loc_nombre = ses[0][1]["data"]["localidad"]
            riesgo_max = max(x[1]["data"]["puntaje_riesgo"] for x in ses)
            filas_session.append({
                "cod_dane": dane,
                "localidad": loc_nombre,
                "sesion_num": idx,
                "inicio_sesion": ini_ses.strftime("%H:%M:%S"),
                "cierre_por_inactividad": fin_ses.strftime("%H:%M:%S"),
                "duracion_sesion_min": duracion_min,
                "alertas_agrupadas": len(ses),
                "riesgo_maximo_sesion": riesgo_max,
            })
    df_session = pd.DataFrame(filas_session)

    return {
        "broker": broker,
        "df_traza": pd.DataFrame(traza_eventos),
        "df_tumbling": df_tumbling,
        "df_comparacion_relojes": df_comparacion_relojes,
        "df_sliding": df_sliding,
        "df_session": df_session,
        "df_dlq": pd.DataFrame(eventos_dlq),
    }


def obtener_codigo_pyspark_structured_streaming() -> str:
    """
    Retorna la definición declarativa en PySpark Structured Streaming equivalente
    a las tres ventanas y watermark implementados, lista para clústeres Spark 4.x.
    """
    return '''# Definición equivalente en PySpark Structured Streaming (Spark 3.5+ / 4.x)
from pyspark.sql.functions import window, session_window, col, count, avg, round

stream_base = (
    spark.readStream
    .schema(esquema_cloudevents_sed)
    .json("data/RAW/eventos_asistencia_desercion_stream.jsonl")
    .withWatermark("event_time", "5 minutes")
)

# 1. Tumbling Window (10 minutos, sin traslape)
tumbling_df = (
    stream_base
    .groupBy(window(col("event_time"), "10 minutes"), col("data.localidad"))
    .agg(count("*").alias("total_eventos"), round(avg("data.puntaje_riesgo"), 2).alias("riesgo_promedio"))
)

# 2. Sliding Window (10 minutos de ancho, avance cada 2 minutos)
sliding_rafagas_df = (
    stream_base
    .groupBy(window(col("event_time"), "10 minutes", "2 minutes"), col("data.localidad"))
    .agg(count("*").alias("eventos_en_ventana"), round(avg("data.puntaje_riesgo"), 2).alias("riesgo_promedio"))
    .where(col("eventos_en_ventana") >= 3)
)

# 3. Session Window (brecha dinámica de inactividad de 4 minutos por colegio DANE)
session_colegios_df = (
    stream_base
    .groupBy(session_window(col("event_time"), "4 minutes"), col("data.cod_dane"), col("data.localidad"))
    .agg(count("*").alias("alertas_agrupadas"))
)'''


# ─────────────────────────────────────────────────────────────────────────────
# 3. Orquestador del Módulo de Streaming y Persistencia en Capa Gold
# ─────────────────────────────────────────────────────────────────────────────

def ejecutar_streaming_educativo() -> dict:
    """
    Ejecuta de principio a fin la simulación y procesamiento de eventos en
    tiempo real del ecosistema educativo de Bogotá D.C., persistiendo las tablas
    resultantes en data/ANALYTICS/.
    """
    os.makedirs(ANALYTICS_DIR, exist_ok=True)

    print("\n" + "=" * 72)
    print("  ETAPA 6 (CORTE 3) — PROCESAMIENTO DE EVENTOS EN TIEMPO REAL (EDA & STREAMING)")
    print("  Event Time · Watermarks · Ventanas (Tumbling/Sliding/Session) · DLQ")
    print("=" * 72)

    eventos = generar_flujo_eventos_educativos()
    print(f"\n📡 1. Ingesta de Flujo de Eventos (CloudEvents 1.0 en JSONL):")
    print(f"   Archivo fuente: {ARCHIVO_STREAM_JSONL}")
    print(f"   Eventos recibidos en el stream: {len(eventos)} (incluye reintentos y tardíos)")

    res = procesar_flujo_con_watermarks_y_ventanas(eventos)
    broker = res["broker"]

    print(f"   Eventos únicos procesados (Idempotencia activa): {len(broker.ids_procesados)}")
    print(f"   Duplicados descartados por ID repetido: {len(broker.duplicados_ignorados)} "
          f"-> {[e['id'] for e in broker.duplicados_ignorados]}")

    print("\n💧 2. Traza de Event Time vs. Processing Time y Evolución del Watermark (WM = max(ET) - 5 min):")
    cols_traza = [
        "id", "localidad", "event_time", "processing_time",
        "retraso_min", "ventana_et", "watermark_al_llegar", "clasificacion_arrivo"
    ]
    print(res["df_traza"][cols_traza].to_string(index=False))

    print("\n⚖️  3. Comparación de Relojes (Event Time vs. Processing Time):")
    print(res["df_comparacion_relojes"].to_string(index=False))

    print("\n🕒 4. Resultados Tumbling Windows (10 min, sin traslape — Reporte Periódico SED):")
    print(res["df_tumbling"].to_string(index=False))

    print("\n🚨 5. Resultados Sliding Windows (10 min cada 2 min — Detección de Ráfagas >= 3 alertas):")
    print(res["df_sliding"].to_string(index=False))

    print("\n🏫 6. Resultados Session Windows (Gap de 4 min — Sesiones Activas por Colegio DANE):")
    print(res["df_session"].to_string(index=False))

    print("\n📥 7. Tratamiento de Eventos Tardíos Extremos — Side Output / Dead Letter Queue (DLQ):")
    if not res["df_dlq"].empty:
        cols_dlq = ["id", "localidad", "cod_dane", "event_time", "processing_time", "retraso_min", "accion_motor"]
        print(res["df_dlq"][cols_dlq].to_string(index=False))
    else:
        print("   Sin eventos en DLQ.")

    # Persistir resultados en data/ANALYTICS/
    rutas_salida = {
        "traza": os.path.join(ANALYTICS_DIR, "streaming_traza_watermarks.csv"),
        "relojes": os.path.join(ANALYTICS_DIR, "streaming_comparacion_event_vs_processing.csv"),
        "tumbling": os.path.join(ANALYTICS_DIR, "streaming_ventanas_tumbling.csv"),
        "sliding": os.path.join(ANALYTICS_DIR, "streaming_ventanas_sliding_rafagas.csv"),
        "session": os.path.join(ANALYTICS_DIR, "streaming_ventanas_session_colegios.csv"),
        "dlq": os.path.join(ANALYTICS_DIR, "streaming_eventos_tardios_dlq.csv"),
    }

    res["df_traza"].to_csv(rutas_salida["traza"], index=False, encoding="utf-8")
    res["df_comparacion_relojes"].to_csv(rutas_salida["relojes"], index=False, encoding="utf-8")
    res["df_tumbling"].to_csv(rutas_salida["tumbling"], index=False, encoding="utf-8")
    res["df_sliding"].to_csv(rutas_salida["sliding"], index=False, encoding="utf-8")
    res["df_session"].to_csv(rutas_salida["session"], index=False, encoding="utf-8")
    res["df_dlq"].to_csv(rutas_salida["dlq"], index=False, encoding="utf-8")

    print("\n💾 [ANALYTICS - STREAMING] Archivos generados en Capa Gold:")
    for k, r in rutas_salida.items():
        print(f"   ✅ {os.path.basename(r)}")

    return res


if __name__ == "__main__":
    ejecutar_streaming_educativo()
