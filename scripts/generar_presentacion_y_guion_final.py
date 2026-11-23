# -*- coding: utf-8 -*-
"""
generar_presentacion_y_guion_final.py

Genera para la Entrega Final (Corte 3):
  1. Presentacion_Sustentacion_Entrega_Final.pptx (10 diapositivas 16:9 con figuras,
     código, evidencias, conceptos de los Cortes 1-3 y aportes individuales).
  2. Presentacion_Interactiva_Sustentacion.html (Presentación web interactiva
     autocontenida con navegación por teclado para proyectar en el video).
  3. Guion_Video_Sustentacion_Entrega_Final.docx (Guion conversacional minuto a
     minuto para el video de 8 minutos [rango 5-10 min], con indicaciones de
     cámara, transición a demostración funcional en vivo y lista de chequeo YouTube).
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor as PptRGB
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

from docx import Document
from docx.shared import Inches as DocInches, Pt as DocPt, RGBColor as DocRGB
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

BASE_REPO = r"C:\Users\dsant\Desktop\U CATOLICA\ING SISTEMAS\SEMESTRES SISTEMAS\noveno semestre\Ingenieria de datos\Proyecto ING Datos\ing datos\Proyecto-ING-DATO"
FIGS_DIR  = os.path.join(BASE_REPO, "data", "ANALYTICS", "figuras")

DEST_DIRS = [
    os.path.join(BASE_REPO, "entregables", "entrega final"),
    r"c:\Users\dsant\Desktop\U CATOLICA\Programas\Data\entregables\entrega final",
]
for d in DEST_DIRS:
    os.makedirs(d, exist_ok=True)

# ══════════════════════════════════════════════════════════════════════════════
# 1. GENERACIÓN DE LA PRESENTACIÓN POWERPOINT (.PPTX)
# ══════════════════════════════════════════════════════════════════════════════
NAVY_DARK   = PptRGB(15, 23, 42)     # #0F172A
SLATE_DARK  = PptRGB(30, 41, 59)     # #1E293B
SLATE_TEXT  = PptRGB(51, 65, 85)     # #334155
BG_LIGHT    = PptRGB(248, 250, 252)  # #F8FAFC
WHITE       = PptRGB(255, 255, 255)
BLUE_ACCENT = PptRGB(37, 99, 235)    # #2563EB
BLUE_SOFT   = PptRGB(219, 234, 254)  # #DBEAFE
GREEN_ACC   = PptRGB(22, 163, 74)    # #16A34A
GREEN_SOFT  = PptRGB(220, 252, 231)  # #DCFCE7
AMBER_ACC   = PptRGB(217, 119, 6)    # #D97706
AMBER_SOFT  = PptRGB(254, 243, 199)  # #FEF3C7
RED_ACC     = PptRGB(220, 38, 38)    # #DC2626
RED_SOFT    = PptRGB(254, 226, 226)  # #FEE2E2
PURPLE_ACC  = PptRGB(124, 58, 237)   # #7C3AED
PURPLE_SOFT = PptRGB(237, 233, 254)  # #EDE9FE

FONT_MAIN = "Segoe UI"
FONT_CODE = "Consolas"


def set_slide_bg(slide, color):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_header(slide, title_text, subtitle_text, slide_tag, speaker_name, time_badge):
    header_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.15))
    header_box.fill.solid()
    header_box.fill.fore_color.rgb = NAVY_DARK
    header_box.line.fill.background()

    stripe = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(1.11), Inches(13.333), Inches(0.05))
    stripe.fill.solid()
    stripe.fill.fore_color.rgb = BLUE_ACCENT
    stripe.line.fill.background()

    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.10), Inches(9.4), Inches(0.55))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.name = FONT_MAIN
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = WHITE

    tb_sub = slide.shapes.add_textbox(Inches(0.5), Inches(0.64), Inches(9.4), Inches(0.38))
    tf_sub = tb_sub.text_frame
    tf_sub.word_wrap = True
    p2 = tf_sub.paragraphs[0]
    p2.text = subtitle_text
    p2.font.name = FONT_MAIN
    p2.font.size = Pt(12.5)
    p2.font.color.rgb = PptRGB(148, 163, 184)

    badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(10.05), Inches(0.18), Inches(2.85), Inches(0.76))
    badge.fill.solid()
    badge.fill.fore_color.rgb = SLATE_DARK
    badge.line.color.rgb = BLUE_ACCENT
    badge.line.width = Pt(1.5)
    tf_b = badge.text_frame
    tf_b.word_wrap = True
    pb = tf_b.paragraphs[0]
    pb.alignment = PP_ALIGN.CENTER
    pb.text = f"{slide_tag} ({time_badge})\nExpone: {speaker_name}"
    pb.font.name = FONT_MAIN
    pb.font.size = Pt(11)
    pb.font.bold = True
    pb.font.color.rgb = WHITE

    footer = slide.shapes.add_textbox(Inches(0.5), Inches(7.12), Inches(12.3), Inches(0.28))
    tf_f = footer.text_frame
    pf = tf_f.paragraphs[0]
    pf.text = "Castelblanco (67001182) · Narváez (67002032)  |  Ingeniería de Datos — Universidad Católica de Colombia  |  Entrega Final 2026"
    pf.font.name = FONT_MAIN
    pf.font.size = Pt(10)
    pf.font.color.rgb = PptRGB(100, 116, 139)


def add_card(slide, left, top, width, height, title, bullets, bg_color=WHITE, border_color=BLUE_ACCENT, title_color=NAVY_DARK, is_code=False):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    card.adjustments[0] = 0.04
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    card.line.color.rgb = border_color
    card.line.width = Pt(1.5)

    tb = slide.shapes.add_textbox(Inches(left + 0.16), Inches(top + 0.12), Inches(width - 0.32), Inches(height - 0.24))
    tf = tb.text_frame
    tf.word_wrap = True

    p0 = tf.paragraphs[0]
    p0.text = title
    p0.font.name = FONT_MAIN
    p0.font.size = Pt(14)
    p0.font.bold = True
    p0.font.color.rgb = title_color
    p0.space_after = Pt(6)

    for b in bullets:
        p = tf.add_paragraph()
        p.text = b
        p.font.name = FONT_CODE if is_code else FONT_MAIN
        p.font.size = Pt(9.8 if is_code else 11.5)
        p.font.color.rgb = WHITE if bg_color == NAVY_DARK else SLATE_TEXT
        p.space_after = Pt(4)


def construir_pptx():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # ── SLIDE 1: PORTADA ──
    s1 = prs.slides.add_slide(blank)
    set_slide_bg(s1, NAVY_DARK)

    banner = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.65), Inches(11.733), Inches(6.15))
    banner.adjustments[0] = 0.03
    banner.fill.solid()
    banner.fill.fore_color.rgb = SLATE_DARK
    banner.line.color.rgb = BLUE_ACCENT
    banner.line.width = Pt(2)

    tb1 = s1.shapes.add_textbox(Inches(1.2), Inches(0.95), Inches(10.9), Inches(5.5))
    tf1 = tb1.text_frame
    tf1.word_wrap = True

    p_kicker = tf1.paragraphs[0]
    p_kicker.text = "UNIVERSIDAD CATÓLICA DE COLOMBIA · FACULTAD DE INGENIERÍA · INGENIERÍA DE DATOS (ENTREGA FINAL CORTE 3)"
    p_kicker.font.name = FONT_MAIN
    p_kicker.font.size = Pt(12)
    p_kicker.font.bold = True
    p_kicker.font.color.rgb = PptRGB(96, 165, 250)
    p_kicker.space_after = Pt(14)

    p_main = tf1.add_paragraph()
    p_main.text = "Sistema Integrado de Análisis y Calidad de Datos para el Ecosistema Educativo Distrital de Bogotá D.C."
    p_main.font.name = FONT_MAIN
    p_main.font.size = Pt(28)
    p_main.font.bold = True
    p_main.font.color.rgb = WHITE
    p_main.space_after = Pt(12)

    p_sub = tf1.add_paragraph()
    p_sub.text = "Arquitectura Híbrida Data Lakehouse (Medallón Bronze/Silver/Gold) + Procesamiento de Eventos en Tiempo Real (CloudEvents, Watermarks, Ventanas Tumbling/Sliding/Session y DLQ)"
    p_sub.font.name = FONT_MAIN
    p_sub.font.size = Pt(15)
    p_sub.font.color.rgb = PptRGB(203, 213, 225)
    p_sub.space_after = Pt(22)

    for line in [
        "👨‍💻 Integrante 1: David Santiago Castelblanco Artunduaga (Código: 67001182) — Lakehouse, PySpark/Pandas, Joins y Ventanas de Tiempo",
        "👨‍💻 Integrante 2: Dairo José Narváez Morales (Código: 67002032) — Ingesta RAW, Contrato Calidad Silver, Reglas y Watermarks/DLQ",
        "👩‍🏫 Docente: Ing. Diana Marcela Toquica Rodríguez   |   📅 Fecha: 22 de noviembre de 2026   |   🔗 Repo: github.com/HCA-3/Proyecto-ING-DATO",
    ]:
        p_a = tf1.add_paragraph()
        p_a.text = line
        p_a.font.name = FONT_MAIN
        p_a.font.size = Pt(13)
        p_a.font.bold = True
        p_a.font.color.rgb = WHITE
        p_a.space_after = Pt(8)

    # ── SLIDE 2: PROBLEMÁTICA SOCIOEDUCATIVA Y TÉCNICA ──
    s2 = prs.slides.add_slide(blank)
    set_slide_bg(s2, BG_LIGHT)
    add_header(s2, "1. Problemática Socioeducativa y Patologías de Datos en Bogotá",
               "Articulación entre brechas territoriales (DANE / SaluData) y problemas técnicos en fuentes abiertas",
               "Diapositiva 2/10", "Dairo Narváez", "00:40 - 01:30")

    add_card(s2, 0.5, 1.35, 6.0, 2.65, "📊 Contexto Social y Educativo (Bogotá D.C.)", [
        "• Pobreza Multidimensional (DANE, 2025): subió de 3.6% a 5.4% (~431,000 personas en Bogotá).",
        "• Rezago Escolar: alcanzó el 17.5% en 2024, siendo un factor crítico de vulnerabilidad.",
        "• Asimetría Territorial (SaluData): Pobreza monetaria en Usme (57.81%), Ciudad Bolívar (57.37%) y Bosa (53.18%) frente a Teusaquillo (7.90%) y Chapinero (12.83%).",
    ], WHITE, BLUE_ACCENT)

    add_card(s2, 6.8, 1.35, 6.0, 2.65, "⚠️ Cuatro Patologías Técnicas en los Datos Crudos", [
        "1. Silos desconectados: Saber 11 (70k), Colegios (2.5k), Jóvenes a la E (15k) y Alertas (20 loc) no comparten esquema unificado.",
        "2. Valores centinela en puntajes: 1,070 registros (1.53%) con '-', 'N/A', 'ND' que rompen el tipado numérico.",
        "3. Pérdida de llave foránea DANE: 2,934 estudiantes (4.2%) sin COLE_COD_DANE_INST.",
        "4. Latencia de red en alertas: colegios rurales/periféricos reportan con retraso.",
    ], RED_SOFT, RED_ACC, PptRGB(153, 27, 27))

    add_card(s2, 0.5, 4.20, 12.3, 2.70, "⚖️ Marco Normativo y Ético Aplicado (Privacy by Design)", [
        "• Ley 1712 de 2014 (Transparencia y Datos Abiertos): Fundamenta la ingesta de activos públicos de la Secretaría de Educación del Distrito (SED).",
        "• Ley 1581 de 2012 (Habeas Data y Protección de Menores): Exclusión de datos identificables en reportes Gold y anonimización criptográfica SHA-256 (función anonimizar_estudiante() -> EST-B8F9A21C4D) en el flujo de eventos en tiempo real.",
        "• Documento CONPES 3920 (Política Nacional de Explotación de Datos - Big Data): Respalda el uso de arquitecturas analíticas para focalizar inversión pública.",
    ], BLUE_SOFT, BLUE_ACCENT, NAVY_DARK)

    # ── SLIDE 3: ARQUITECTURA HÍBRIDA (CORTES 1, 2 Y 3) ──
    s3 = prs.slides.add_slide(blank)
    set_slide_bg(s3, BG_LIGHT)
    add_header(s3, "2. Arquitectura Híbrida: Data Lakehouse (Medallón) + Streaming EDA",
               "Aplicación integrada de los conceptos vistos en los Cortes 1, 2 y 3 de Ingeniería de Datos",
               "Diapositiva 3/10", "David Castelblanco", "01:30 - 02:25")

    fig1_path = os.path.join(FIGS_DIR, "fig1_arquitectura_lakehouse_streaming.png")
    if os.path.exists(fig1_path):
        s3.shapes.add_picture(fig1_path, Inches(0.5), Inches(1.32), width=Inches(8.3))

    add_card(s3, 8.95, 1.32, 3.88, 5.65, "🧭 Conceptos Aplicados", [
        "• Corte 1 (Almacenamiento y Modelo): Transición de OLTP a OLAP mediante Data Lakehouse Medallón (Bronze inmutable, Silver validada, Gold Estrella).",
        "• Corte 2 (Cómputo Distribuido): Procesamiento con PySpark SQL y motor resiliente Pandas/NumPy.",
        "• Corte 3 (Pipelines y Streaming): Pipeline ELT/ETL de 6 etapas + Arquitectura Orientada a Eventos (CloudEvents 1.0, Broker idempotente, Watermarks, Ventanas Tumbling/Sliding/Session y DLQ).",
    ], WHITE, PURPLE_ACC, NAVY_DARK)

    # ── SLIDE 4: CONTRATO DE CALIDAD (DAMA-DMBOK) Y CÓDIGO SILVER ──
    s4 = prs.slides.add_slide(blank)
    set_slide_bg(s4, BG_LIGHT)
    add_header(s4, "3. Contrato de Calidad de Datos (DAMA-DMBOK) — Capa Silver",
               "Explicación técnica de la limpieza de centinelas O(1) e imputación estadística por moda DANE",
               "Diapositiva 4/10", "Dairo Narváez", "02:25 - 03:20")

    fig2_path = os.path.join(FIGS_DIR, "fig2_calidad_datos_bronze_silver.png")
    if os.path.exists(fig2_path):
        s4.shapes.add_picture(fig2_path, Inches(0.5), Inches(1.35), width=Inches(6.1))

    add_card(s4, 6.75, 1.35, 6.08, 3.15, "💻 Código: scripts/pipeline_educativo.py", [
        "VALORES_CENTINELA = {'-', 'N/A', 'ND', '', 'nan', 'none'}",
        "def validar_rango_puntaje(valor_str) -> float | None:",
        "    val = str(valor_str).strip()",
        "    if val.lower() in VALORES_CENTINELA: return None",
        "    numero = float(val)",
        "    return numero if 0.0 <= numero <= 500.0 else None",
        "",
        "# Imputación por moda para salvar 2,934 registros sin DANE:",
        "df = limpiar_valores_nulos(df, 'COLE_COD_DANE_INST', 'moda')",
    ], NAVY_DARK, BLUE_ACCENT, PptRGB(96, 165, 250), is_code=True)

    add_card(s4, 6.75, 4.65, 6.08, 2.30, "✅ Evidencia de Mitigación Técnica", [
        "• Validez: 1,070 puntajes con texto inválido convertidos a NaN seguro.",
        "• Completitud: 2,934 códigos DANE imputados con la moda (0 nulos).",
        "• Pertinencia: Filtrado de 2,262 colegios con ESTADO == 'ACTIVO'.",
    ], GREEN_SOFT, GREEN_ACC, PptRGB(22, 101, 52))

    # ── SLIDE 5: MOTOR HÍBRIDO PYSPARK/PANDAS, LEFT JOIN Y EQUIDAD GOLD ──
    s5 = prs.slides.add_slide(blank)
    set_slide_bg(s5, BG_LIGHT)
    add_header(s5, "4. Procesamiento Híbrido PySpark/Pandas y Datamart Capa Gold",
               "Conmutación automática de entorno, cruce relacional LEFT JOIN y regla de alertas (> 8.0%)",
               "Diapositiva 5/10", "David Castelblanco", "03:20 - 04:15")

    fig3_path = os.path.join(FIGS_DIR, "fig3_equidad_desercion_localidades.png")
    if os.path.exists(fig3_path):
        s5.shapes.add_picture(fig3_path, Inches(0.5), Inches(1.35), width=Inches(6.35))

    add_card(s5, 7.0, 1.35, 5.8, 2.85, "💻 Resiliencia y Cruce Dimensional", [
        "def crear_spark_session():",
        "    if not comprobar_java_compatible():",
        "        # Java 8 detectado -> activa motor Pandas sin fallar",
        "        return None",
        "    return SparkSession.builder.master('local[*]').getOrCreate()",
        "",
        "# LEFT JOIN: 70,000 estudiantes × 2,262 colegios activos",
        "df_enriq = df_saber11.join(dir_sel, on='COD_DANE', how='left')",
    ], NAVY_DARK, BLUE_ACCENT, PptRGB(96, 165, 250), is_code=True)

    add_card(s5, 7.0, 4.35, 5.8, 2.60, "🚨 Hallazgos Reales en Capa Gold (ANALYTICS)", [
        "• 63,498 estudiantes vinculados exitosamente a colegio activo.",
        "• 9 localidades superan el umbral crítico del 8.0% de deserción: Usme (12.1%), Fontibón (11.8%), Engativá (11.3%), Los Mártires (11.2%), Tunjuelito (10.9%), Usaquén (9.7%), Kennedy (9.1%), La Candelaria (8.9%) y Antonio Nariño (8.1%).",
        "• Absorción IES: UNIMINUTO (38.21%) y UNAL (36.53%) lideran graduación.",
    ], AMBER_SOFT, AMBER_ACC, PptRGB(146, 64, 14))

    # ── SLIDE 6: CORTE 3 — EDA, EVENT TIME VS PROCESSING TIME Y WATERMARKS ──
    s6 = prs.slides.add_slide(blank)
    set_slide_bg(s6, BG_LIGHT)
    add_header(s6, "5. Tercer Corte (Parte I): EDA, Event Time, Watermarks y Cola DLQ",
               "Implementación en scripts/streaming_alertas_educativas.py con CloudEvents 1.0 e idempotencia",
               "Diapositiva 6/10", "Dairo Narváez", "04:15 - 05:10")

    fig4_path = os.path.join(FIGS_DIR, "fig4_streaming_watermarks_ventanas.png")
    if os.path.exists(fig4_path):
        s6.shapes.add_picture(fig4_path, Inches(0.5), Inches(1.35), width=Inches(6.6))

    add_card(s6, 7.25, 1.35, 5.55, 2.75, "💧 Fórmula del Watermark y 3 Políticas de Tardíos", [
        "• Watermark dinámico: WM = max(event_time) - 5 min.",
        "• 1. Dentro del WM (EVT-009 Usme, 4.33m tarde): WM=08:08 < 08:10 -> ACEPTADO_EN_VENTANA.",
        "• 2. Allowed Lateness (+5m) (EVT-013 Usme, 12m tarde): WM=08:14 <= 08:15 -> Reabre estado y emite CORRECCIÓN.",
        "• 3. Side Output / DLQ (EVT-017 Sumapaz, 28m tarde): WM=08:24 > 08:15 -> Desviado a sed.alertas.dlq (0 pérdida).",
    ], WHITE, PURPLE_ACC, NAVY_DARK)

    add_card(s6, 7.25, 4.25, 5.55, 2.70, "🛡️ Broker EDA Idempotente (Deduplicación O(1))", [
        "• Recibidos en stream JSONL: 18 eventos CloudEvents 1.0.",
        "• El evento EVT-004 (Bosa) llega dos veces por reintento de red.",
        "• BrokerEventosEducativos verifica event_id in self.ids_procesados y descarta el duplicado -> 17 eventos únicos procesados.",
        "• Evidencia empírica: Processing Time subestima [-2] en [08:00, 08:10) y sobreestima [+2] en [08:10, 08:20).",
    ], PURPLE_SOFT, PURPLE_ACC, PptRGB(91, 33, 182))

    # ── SLIDE 7: CORTE 3 — VENTANAS TUMBLING, SLIDING Y SESSION + MOTORES ──
    s7 = prs.slides.add_slide(blank)
    set_slide_bg(s7, BG_LIGHT)
    add_header(s7, "6. Tercer Corte (Parte II): Ventanas Tumbling, Sliding y Session",
               "Segmentación del flujo infinito en bloques finitos y comparación Spark / Flink / Kafka Streams",
               "Diapositiva 7/10", "David Castelblanco", "05:10 - 06:00")

    add_card(s7, 0.5, 1.35, 4.0, 3.10, "🕒 1. Tumbling Window (10 min)", [
        "• Tamaño fijo (10 min), SIN traslape.",
        "• Costo en RAM: 1 ventana por evento.",
        "• Uso SED: Reporte periódico por localidad.",
        "• Evidencia: [08:00, 08:10) C. Bolívar = 3 alertas (riesgo 88.17); Usme = 2 alertas (92.00, corregida por Allowed Lateness).",
    ], WHITE, BLUE_ACCENT)

    add_card(s7, 4.65, 1.35, 4.0, 3.10, "🚨 2. Sliding Window (10m / 2m)", [
        "• Tamaño fijo (10 min), avanza cada 2 min.",
        "• Costo en RAM: 10 ÷ 2 = 5 ventanas/evento.",
        "• Uso SED: Detectar ráfagas (>= 3 alertas).",
        "• Evidencia: Detecta ráfaga de 3 alertas en BOSA en [08:02, 08:12) que la Tumbling partía en dos bloques distintos.",
    ], RED_SOFT, RED_ACC, PptRGB(153, 27, 27))

    add_card(s7, 8.8, 1.35, 4.0, 3.10, "🏫 3. Session Window (Gap 4m)", [
        "• Tamaño dinámico por actividad de clave (código DANE del colegio).",
        "• Cierra tras 4 min de inactividad.",
        "• Evidencia: En IED 110010190012 (C. Bolívar) fusiona 3 alertas en Sesión #1 (7.08 min) y abre Sesión #2 a las 08:29.",
    ], GREEN_SOFT, GREEN_ACC, PptRGB(22, 101, 52))

    add_card(s7, 0.5, 4.60, 12.3, 2.35, "⚙️ Código PySpark Structured Streaming & Comparativa con Apache Flink y Kafka Streams", [
        "stream_base = spark.readStream.schema(esquema).json('data/RAW/eventos_asistencia_desercion_stream.jsonl').withWatermark('event_time', '5 minutes')",
        "tumbling = stream_base.groupBy(window(col('event_time'), '10 minutes'), col('data.localidad')).count()",
        "sliding  = stream_base.groupBy(window(col('event_time'), '10 minutes', '2 minutes'), col('data.localidad')).count().where('count >= 3')",
        "session  = stream_base.groupBy(session_window(col('event_time'), '4 minutes'), col('data.cod_dane')).count()",
        "• Decisión Arquitectónica: Spark unifica Batch + Streaming por micro-lotes; Flink 2.0 aporta estado desagregado y Side Outputs nativos; Kafka Streams opera como librería ligera con KStream/KTable.",
    ], NAVY_DARK, BLUE_ACCENT, PptRGB(96, 165, 250), is_code=True)

    # ── SLIDE 8: IDENTIFICACIÓN ARGUMENTADA DE APORTES POR INTEGRANTE ──
    s8 = prs.slides.add_slide(blank)
    set_slide_bg(s8, BG_LIGHT)
    add_header(s8, "7. Identificación Clara y Argumentada de Aportes por Integrante",
               "Qué desarrolló cada autor, cómo lo implementó en código y en qué sección del manual se documenta",
               "Diapositiva 8/10", "Castelblanco y Narváez", "06:00 - 06:45")

    add_card(s8, 0.5, 1.35, 6.0, 5.55, "👨‍💻 David Santiago Castelblanco Artunduaga (67001182)", [
        "• ¿Qué desarrolló?:",
        "  1. Arquitectura Lakehouse Medallón (Bronze/Silver/Gold).",
        "  2. Motor de resiliencia híbrida PySpark 4.x / Pandas.",
        "  3. Cruce relacional LEFT JOIN y Datamart de Equidad Gold.",
        "  4. Algoritmos de Ventanas Tumbling, Sliding y Session (Corte 3) y generador de figuras APA 7.",
        "• ¿Cómo lo implementó?:",
        "  En pipeline_educativo.py implementó comprobar_java_compatible(), crear_spark_session(), cruzar_saber11_con_directorio() y generar_reporte_equidad_educativa(). En streaming_alertas_educativas.py programó la aritmética de intervalos [inicio, fin), el barrido 5x de Sliding Windows y la fusión por gap de 4 min en Session Windows.",
        "• ¿Dónde está en el Manual Técnico APA 7?:",
        "  Secciones 1, 3, 5.1, 5.2, 5.6, 5.7, 7 (Figura 1), 9.2.1, 9.2.3 (Tabla 7, Figura 3), 9.3 (Tabla 10) y 10.1.",
    ], WHITE, BLUE_ACCENT, NAVY_DARK)

    add_card(s8, 6.8, 1.35, 6.0, 5.55, "👨‍💻 Dairo José Narváez Morales (67002032)", [
        "• ¿Qué desarrolló?:",
        "  1. Generador reproducible RAW (87,520 registros, semilla 42).",
        "  2. Contrato de calidad Silver (centinelas y moda DANE).",
        "  3. Regla de alertas de deserción (> 8.0%) y absorción IES.",
        "  4. Broker EDA idempotente CloudEvents, Watermarks, Allowed Lateness y cola DLQ (Corte 3) + Manual de Usuario.",
        "• ¿Cómo lo implementó?:",
        "  En generar_datos_raw.py programó puntaje_raw() y get_dane(). En pipeline_educativo.py implementó validar_rango_puntaje() con búsqueda O(1) en hash set, limpiar_valores_nulos() por moda y evaluar_alerta_localidad(). En streaming_alertas_educativas.py implementó BrokerEventosEducativos, anonimizar_estudiante() (SHA-256) y la regla tripartita de Watermark / DLQ.",
        "• ¿Dónde está en el Manual Técnico APA 7?:",
        "  Secciones 2, 4 (Tabla 1), 5.3, 5.4, 5.5, 6 (Tabla 4), 8 (Tabla 6, Figura 2), 9.1, 9.2.2, 9.2.3 (Tabla 8), 9.3 (Tabla 9, Figura 4), 9.4 (Tabla 11), 10.2 y 11.",
    ], WHITE, GREEN_ACC, NAVY_DARK)

    # ── SLIDE 9: DEMOSTRACIÓN FUNCIONAL EN VIVO (PASO A PASO) ──
    s9 = prs.slides.add_slide(blank)
    set_slide_bg(s9, BG_LIGHT)
    add_header(s9, "8. Demostración Funcional en Vivo de la Solución (Terminal CLI)",
               "Ejecución en tiempo real de las 6 etapas del pipeline y verificación de entregables en disco",
               "Diapositiva 9/10", "Demostración Conjunta", "06:45 - 07:35")

    add_card(s9, 0.5, 1.35, 12.3, 2.75, "🖥️ Comandos de Ejecución en Vivo (Compartir Pantalla — Terminal PowerShell)", [
        "# Paso 1: Generar los 4 datasets crudos inmutables en data/RAW/ (87,520 registros con anomalías controladas)",
        "PS Proyecto-ING-DATO> python scripts/generar_datos_raw.py",
        "",
        "# Paso 2: Ejecutar el pipeline completo (Etapas 1 a 5 Lakehouse Batch + Etapa 6 Streaming Corte 3)",
        "PS Proyecto-ING-DATO> python scripts/pipeline_educativo.py",
    ], NAVY_DARK, BLUE_ACCENT, PptRGB(96, 165, 250), is_code=True)

    add_card(s9, 0.5, 4.25, 6.0, 2.65, "📂 Verificación en data/STAGING/ (Capa Silver)", [
        "• Saber11_2025_Bogota_clean.csv: 70,000 filas × 35 cols (1,070 centinelas limpiados y 2,934 códigos DANE imputados -> 0 nulos).",
        "• Directorio_Colegios_clean.csv: 2,262 colegios activos filtrados de los 2,500 originales.",
    ], BLUE_SOFT, BLUE_ACCENT, NAVY_DARK)

    add_card(s9, 6.8, 4.25, 6.0, 2.65, "📊 Verificación en data/ANALYTICS/ (Capa Gold)", [
        "• 5 Reportes Batch: reporte_equidad_educativa.csv, alertas_criticas_localidad.csv, absorcion_educacion_superior.csv, etc.",
        "• 6 Reportes Streaming Corte 3: streaming_traza_watermarks.csv, streaming_ventanas_tumbling.csv, sliding_rafagas.csv, session_colegios.csv y eventos_tardios_dlq.csv.",
    ], GREEN_SOFT, GREEN_ACC, PptRGB(22, 101, 52))

    # ── SLIDE 10: CONCLUSIONES Y REFERENCIAS APA 7 ──
    s10 = prs.slides.add_slide(blank)
    set_slide_bg(s10, BG_LIGHT)
    add_header(s10, "9. Conclusiones y Referencias Bibliográficas (Normas APA 7)",
               "Síntesis del impacto técnico y fuentes académicas que respaldan el proyecto integrador",
               "Diapositiva 10/10", "Castelblanco y Narváez", "07:35 - 08:00")

    add_card(s10, 0.5, 1.35, 6.0, 5.55, "🎯 Conclusiones Técnicas Principales", [
        "1. Calidad e Inmutabilidad Garantizadas: La arquitectura Medallón permitió conservar intacto el origen (RAW) mientras el contrato Silver recuperó 2,934 estudiantes sin DANE y eliminó 1,070 centinelas.",
        "2. Superioridad del Event Time y Watermarks: Se demostró cuantitativamente que Processing Time distorsiona las alertas ante caídas de red, mientras que Event Time + Watermark (5 min) + Allowed Lateness (+5 min) + DLQ garantiza 0 pérdida de alertas estudiantiles.",
        "3. Complementariedad de Ventanas: Tumbling entrega el consolidado oficial cada 10 min, Sliding (10m/2m) captura ráfagas entre fronteras (caso Bosa) y Session (gap 4m) modela la actividad real por colegio.",
    ], WHITE, BLUE_ACCENT, NAVY_DARK)

    add_card(s10, 6.8, 1.35, 6.0, 5.55, "📚 Referencias Seleccionadas (APA 7.ª Edición)", [
        "• Akidau, T., et al. (2015). The dataflow model: A practical approach to balancing correctness, latency, and cost. PVLDB, 8(12), 1792–1803.",
        "• Akidau, T., Chernyak, S., & Lax, R. (2018). Streaming systems. O'Reilly Media.",
        "• Apache Software Foundation. (2026). Apache Spark 4.0, Apache Flink 2.0 & Apache Kafka 4.0 Official Documentation.",
        "• Cloud Native Computing Foundation [CNCF]. (2024). CloudEvents specification (v1.0.2).",
        "• Congreso de la República de Colombia. (2012, 2014). Ley 1581 de 2012 (Habeas Data) y Ley 1712 de 2014 (Datos Abiertos).",
        "• DAMA International. (2017). DAMA-DMBOK (2nd ed.).",
        "• DANE (2025) & DNP (2018). Pobreza Multidimensional y CONPES 3920.",
    ], WHITE, PURPLE_ACC, NAVY_DARK)

    for d in DEST_DIRS:
        out_pptx = os.path.join(d, "Presentacion_Sustentacion_Entrega_Final.pptx")
        prs.save(out_pptx)
        print("✅ Presentación PPTX guardada en:", out_pptx)


# ══════════════════════════════════════════════════════════════════════════════
# 2. GENERACIÓN DE LA PRESENTACIÓN INTERACTIVA (.HTML)
# ══════════════════════════════════════════════════════════════════════════════
def construir_html_interactivo():
    html_content = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Sustentación Final — Sistema Integrado de Análisis y Calidad de Datos Educativos Bogotá D.C.</title>
<style>
  :root {
    --navy: #0F172A; --slate: #1E293B; --blue: #2563EB; --green: #16A34A;
    --amber: #D97706; --red: #DC2626; --purple: #7C3AED; --bg: #F8FAFC;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', system-ui, sans-serif; background: var(--navy); color: #0F172A; overflow: hidden; }
  .deck { width: 100vw; height: 100vh; display: flex; flex-direction: column; }
  .slide { display: none; flex: 1; background: var(--bg); padding: 0; flex-direction: column; }
  .slide.active { display: flex; }
  .header { background: var(--navy); color: white; padding: 18px 36px; display: flex; justify-content: space-between; align-items: center; border-bottom: 4px solid var(--blue); }
  .header h1 { font-size: 1.45rem; margin-bottom: 4px; }
  .header p { font-size: 0.92rem; color: #94A3B8; }
  .badge { background: var(--slate); border: 1.5px solid var(--blue); padding: 8px 16px; border-radius: 10px; text-align: center; font-size: 0.85rem; font-weight: 700; color: #E2E8F0; }
  .content { flex: 1; padding: 24px 36px; display: grid; gap: 18px; overflow-y: auto; }
  .grid-2 { grid-template-columns: 1fr 1fr; }
  .grid-3 { grid-template-columns: 1fr 1fr 1fr; }
  .card { background: white; border-radius: 12px; padding: 18px 22px; border-left: 5px solid var(--blue); box-shadow: 0 4px 14px rgba(15,23,42,0.06); }
  .card.red { border-left-color: var(--red); background: #FEF2F2; }
  .card.green { border-left-color: var(--green); background: #F0FDF4; }
  .card.purple { border-left-color: var(--purple); background: #FAF5FF; }
  .card.amber { border-left-color: var(--amber); background: #FFFBEB; }
  .card h3 { font-size: 1.08rem; margin-bottom: 10px; color: var(--navy); }
  .card ul { padding-left: 18px; line-height: 1.55; font-size: 0.95rem; color: #334155; }
  pre { background: var(--navy); color: #E2E8F0; padding: 14px 18px; border-radius: 10px; font-family: Consolas, monospace; font-size: 0.84rem; line-height: 1.45; overflow-x: auto; border: 1px solid #334155; }
  table { width: 100%; border-collapse: collapse; font-size: 0.88rem; background: white; border-radius: 8px; overflow: hidden; }
  th { background: var(--navy); color: white; text-align: left; padding: 9px 12px; }
  td { padding: 8px 12px; border-bottom: 1px solid #E2E8F0; }
  .footer { background: #E2E8F0; padding: 10px 36px; display: flex; justify-content: space-between; align-items: center; font-size: 0.85rem; color: #475569; }
  .controls button { background: var(--navy); color: white; border: none; padding: 6px 14px; border-radius: 6px; cursor: pointer; font-weight: 600; margin-left: 6px; }
  .controls button:hover { background: var(--blue); }
</style>
</head>
<body>
<div class="deck">
  <!-- SLIDE 1 -->
  <div class="slide active" style="background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%); color: white; justify-content: center; align-items: center; padding: 48px;">
    <div style="max-width: 980px; background: rgba(30,41,59,0.85); border: 2px solid #2563EB; border-radius: 18px; padding: 42px;">
      <p style="color: #60A5FA; font-weight: 700; letter-spacing: 1px; margin-bottom: 12px;">UNIVERSIDAD CATÓLICA DE COLOMBIA · FACULTAD DE INGENIERÍA · INGENIERÍA DE DATOS (ENTREGA FINAL CORTE 3)</p>
      <h1 style="font-size: 2.3rem; line-height: 1.25; margin-bottom: 16px;">Sistema Integrado de Análisis y Calidad de Datos para el Ecosistema Educativo Distrital de Bogotá D.C.</h1>
      <p style="font-size: 1.12rem; color: #CBD5E1; margin-bottom: 26px;">Arquitectura Híbrida Data Lakehouse (Medallón Bronze/Silver/Gold) + Procesamiento de Eventos en Tiempo Real (CloudEvents 1.0, Watermarks, Ventanas Tumbling/Sliding/Session y DLQ)</p>
      <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-bottom: 20px;">
        <div style="background: #0F172A; padding: 14px 18px; border-radius: 10px; border-left: 4px solid #3B82F6;">
          <strong>David Santiago Castelblanco Artunduaga</strong><br><span style="color:#94A3B8;">Código: 67001182 · Lakehouse, PySpark/Pandas, Joins y Ventanas de Tiempo</span>
        </div>
        <div style="background: #0F172A; padding: 14px 18px; border-radius: 10px; border-left: 4px solid #10B981;">
          <strong>Dairo José Narváez Morales</strong><br><span style="color:#94A3B8;">Código: 67002032 · Ingesta RAW, Contrato Silver, Reglas y Watermarks/DLQ</span>
        </div>
      </div>
      <p style="font-size: 0.95rem; color: #94A3B8;">Docente: <strong>Ing. Diana Marcela Toquica Rodríguez</strong> &nbsp;|&nbsp; Fecha: <strong>22 de noviembre de 2026</strong> &nbsp;|&nbsp; Repositorio: <strong>github.com/HCA-3/Proyecto-ING-DATO</strong></p>
    </div>
  </div>

  <!-- SLIDE 2 -->
  <div class="slide">
    <div class="header">
      <div><h1>1. Problemática Socioeducativa y Patologías de Datos</h1><p>Brechas territoriales en Bogotá D.C. (DANE / SaluData) y anomalías técnicas en fuentes abiertas</p></div>
      <div class="badge">Diapositiva 2/8 · 00:40 - 01:40<br>Expone: Dairo Narváez</div>
    </div>
    <div class="content grid-2">
      <div class="card">
        <h3>📊 Contexto Socioeducativo (DANE, 2025; SaluData, 2025)</h3>
        <ul>
          <li><strong>Pobreza Multidimensional:</strong> Ascendió del 3.6% (2023) al 5.4% (2024) en Bogotá (~431,000 personas).</li>
          <li><strong>Rezago Escolar:</strong> Alcanzó el <strong>17.5%</strong> en 2024, siendo la dimensión educativa más crítica.</li>
          <li><strong>Brecha Territorial Extrema:</strong> Pobreza monetaria en <strong>Usme (57.81%)</strong>, <strong>Ciudad Bolívar (57.37%)</strong> y <strong>Bosa (53.18%)</strong> frente a <strong>Teusaquillo (7.90%)</strong> y <strong>Chapinero (12.83%)</strong>.</li>
        </ul>
      </div>
      <div class="card red">
        <h3>⚠️ Patologías Técnicas Verificadas en Capa RAW</h3>
        <ul>
          <li><strong>Silos fragmentados:</strong> 4 fuentes tabulares (87,520 registros) + flujo de eventos sin llave común saneada.</li>
          <li><strong>Valores centinela en Saber 11:</strong> <strong>1,070 registros (1.53%)</strong> con <code>"-"</code>, <code>"N/A"</code>, <code>"ND"</code> en puntajes numéricos.</li>
          <li><strong>Orfandad de Código DANE:</strong> <strong>2,934 estudiantes (4.20%)</strong> carecen de <code>COLE_COD_DANE_INST</code>.</li>
          <li><strong>Marco Ético (Ley 1581/2012):</strong> Minimización y hash SHA-256 (<code>EST-B8F9A21C4D</code>) en menores de edad.</li>
        </ul>
      </div>
    </div>
  </div>

  <!-- SLIDE 3 -->
  <div class="slide">
    <div class="header">
      <div><h1>2. Arquitectura Híbrida: Lakehouse Medallón + Streaming (Cortes 1, 2 y 3)</h1><p>Integración de procesamiento por lotes (Bronze, Silver, Gold) y flujo de eventos en tiempo real</p></div>
      <div class="badge">Diapositiva 3/8 · 01:40 - 02:45<br>Expone: David Castelblanco</div>
    </div>
    <div class="content grid-3">
      <div class="card amber">
        <h3>🥉 Capa Bronze (data/RAW/)</h3>
        <ul>
          <li><strong>Inmutabilidad estricta</strong> (filosofía ELT / Data Lakehouse).</li>
          <li><code>Saber11_2025_Bogota.csv</code> (70,000 filas × 35 cols).</li>
          <li><code>Directorio_Colegios_Bogota.csv</code> (2,500 filas × 16 cols).</li>
          <li><code>Beneficiarios_Jovenes_a_la_E.csv</code> (15,000 filas × 10 cols).</li>
          <li><code>Reporte_Alertas_Localidad.csv</code> (20 localidades × 8 cols).</li>
          <li><code>eventos_asistencia_desercion_stream.jsonl</code> (18 eventos CloudEvents 1.0).</li>
        </ul>
      </div>
      <div class="card">
        <h3>🥈 Capa Silver (data/STAGING/)</h3>
        <ul>
          <li><strong>Contrato de Calidad DAMA-DMBOK:</strong></li>
          <li>Búsqueda O(1) en conjunto hash <code>VALORES_CENTINELA</code> y validación de rango <code>[0.0, 500.0]</code>.</li>
          <li>Imputación estadística por <strong>moda</strong> en <code>COLE_COD_DANE_INST</code> (2,934 recuperados &rarr; 0 nulos).</li>
          <li>Filtrado de <strong>2,262 colegios ACTIVOS</strong>.</li>
          <li>Resiliencia híbrida automática <strong>PySpark 4.x &harr; Pandas</strong>.</li>
        </ul>
      </div>
      <div class="card green">
        <h3>🥇 Capa Gold (data/ANALYTICS/)</h3>
        <ul>
          <li><strong>Esquema en Estrella + Sinks de Streaming:</strong></li>
          <li><code>LEFT JOIN</code> preserva los 70,000 estudiantes (63,498 vinculados a sede activa).</li>
          <li>Regla automática de deserción (<code>tasa &gt; 8.0%</code>) detecta <strong>9 localidades en ALERTA CRÍTICA</strong>.</li>
          <li>Salidas de Streaming Corte 3: Watermarks, ventanas <strong>Tumbling</strong>, <strong>Sliding</strong>, <strong>Session</strong> y cola <strong>DLQ</strong>.</li>
        </ul>
      </div>
    </div>
  </div>

  <!-- SLIDE 4 -->
  <div class="slide">
    <div class="header">
      <div><h1>3. Implementación Batch: Limpieza Silver, LEFT JOIN y Reglas Gold</h1><p>Fragmentos de código clave en scripts/pipeline_educativo.py y evidencias cuantitativas</p></div>
      <div class="badge">Diapositiva 4/8 · 02:45 - 04:00<br>Exponen: Narváez y Castelblanco</div>
    </div>
    <div class="content grid-2">
      <pre>
# 1. Limpieza O(1) de centinelas (Dairo Narváez)
VALORES_CENTINELA = {"-", "N/A", "ND", "", "nan", "none"}
def validar_rango_puntaje(valor_str) -> float | None:
    if valor_str is None or pd.isna(valor_str): return None
    val = str(valor_str).strip()
    if val.lower() in VALORES_CENTINELA: return None
    numero = float(val)
    return numero if 0.0 &lt;= numero &lt;= 500.0 else None

# 2. Cruce dimensional LEFT JOIN PySpark/Pandas (David Castelblanco)
df_enriquecido = df_saber11.join(
    dir_sel,
    df_saber11["COLE_COD_DANE_INST"] == dir_sel["COD_DANE"],
    how="left"
).drop("COD_DANE")</pre>
      <div>
        <table>
          <thead><tr><th>Localidad</th><th>Puntaje Saber 11</th><th>Deserción (%)</th><th>Dictamen Regla (&gt;8.0%)</th></tr></thead>
          <tbody>
            <tr><td><strong>Usme</strong></td><td>251.89</td><td>12.07%</td><td>🚨 ALERTA CRÍTICA</td></tr>
            <tr><td><strong>Fontibón</strong></td><td>250.15</td><td>11.77%</td><td>🚨 ALERTA CRÍTICA</td></tr>
            <tr><td><strong>Engativá</strong></td><td>249.67</td><td>11.26%</td><td>🚨 ALERTA CRÍTICA</td></tr>
            <tr><td><strong>Los Mártires</strong></td><td>248.38</td><td>11.22%</td><td>🚨 ALERTA CRÍTICA</td></tr>
            <tr><td><strong>Tunjuelito</strong></td><td>246.13</td><td>10.92%</td><td>🚨 ALERTA CRÍTICA</td></tr>
            <tr><td><strong>Kennedy</strong></td><td>249.95</td><td>9.08%</td><td>🚨 ALERTA CRÍTICA</td></tr>
            <tr><td><strong>Ciudad Bolívar</strong></td><td>244.63 (Mín)</td><td>3.16%</td><td>✅ Dentro de umbral</td></tr>
          </tbody>
        </table>
      </div>
    </div>
  </div>

  <!-- SLIDE 5 -->
  <div class="slide">
    <div class="header">
      <div><h1>4. Tercer Corte (Parte I): EDA, Event Time vs. Processing Time y Watermarks</h1><p>Implementación funcional en scripts/streaming_alertas_educativas.py</p></div>
      <div class="badge">Diapositiva 5/8 · 04:00 - 05:10<br>Expone: Dairo Narváez</div>
    </div>
    <div class="content grid-2">
      <div class="card purple">
        <h3>💧 Broker EDA Idempotente y Fórmula del Watermark</h3>
        <ul>
          <li><strong>Estándar CloudEvents 1.0:</strong> Cada alerta viaja con <code>id</code>, <code>event_time</code> (colegio) y <code>processing_time</code> (servidor SED).</li>
          <li><strong>Idempotencia O(1):</strong> De 18 eventos recibidos, el reintento duplicado <code>EVT-004</code> (Bosa) es bloqueado por <code>ids_procesados</code> &rarr; <strong>17 eventos únicos</strong>.</li>
          <li><strong>Watermark:</strong> <code>WM = max(event_time) - 5 min</code>.</li>
          <li><strong>Distorsión probada en [08:00, 08:10):</strong> Ocurrieron <strong>8 eventos reales (Event Time)</strong>, pero <em>Processing Time</em> solo vio <strong>6 (-2)</strong> e infló el bloque [08:10, 08:20) a <strong>7 (+2)</strong>.</li>
        </ul>
      </div>
      <div>
        <table>
          <thead><tr><th>Evento</th><th>Localidad</th><th>Event Time</th><th>Proc. Time</th><th>WM al Llegar</th><th>Política Aplicada</th></tr></thead>
          <tbody>
            <tr><td><strong>EVT-009</strong></td><td>Usme</td><td>08:09:10</td><td>08:13:30 (+4.3m)</td><td>08:08:00</td><td>✅ <strong>ACEPTADO_EN_VENTANA</strong> (WM &lt; 08:10)</td></tr>
            <tr><td><strong>EVT-013</strong></td><td>Usme Rural</td><td>08:07:30</td><td>08:19:30 (+12m)</td><td>08:14:00</td><td>🔄 <strong>ALLOWED_LATENESS</strong> (WM &le; 08:15; corrige)</td></tr>
            <tr><td><strong>EVT-017</strong></td><td>Sumapaz</td><td>08:03:00</td><td>08:31:00 (+28m)</td><td>08:24:00</td><td>📥 <strong>SIDE OUTPUT / DLQ</strong> (Evita Drop de Spark)</td></tr>
          </tbody>
        </table>
      </div>
    </div>
  </div>

  <!-- SLIDE 6 -->
  <div class="slide">
    <div class="header">
      <div><h1>5. Tercer Corte (Parte II): Ventanas Tumbling, Sliding y Session</h1><p>Resultados experimentales en Capa Gold y código declarativo en PySpark Structured Streaming</p></div>
      <div class="badge">Diapositiva 6/8 · 05:10 - 06:10<br>Expone: David Castelblanco</div>
    </div>
    <div class="content grid-3">
      <div class="card">
        <h3>🕒 Tumbling (10 min fija)</h3>
        <ul>
          <li>Sin traslape (1 ventana/evento).</li>
          <li><code>[08:00, 08:10)</code> C. Bolívar: 3 alertas (riesgo 88.17).</li>
          <li><code>[08:00, 08:10)</code> Usme: 2 alertas (riesgo 92.00, <strong>re-emitida por Allowed Lateness</strong>).</li>
        </ul>
      </div>
      <div class="card red">
        <h3>🚨 Sliding (10 min / slide 2 min)</h3>
        <ul>
          <li>Con traslape (10 &divide; 2 = 5 ventanas/evento).</li>
          <li>Detecta <strong>ráfaga crítica de 3 alertas en Bosa en <code>[08:02, 08:12)</code></strong> que Tumbling partía entre [08:00, 08:10) y [08:10, 08:20).</li>
        </ul>
      </div>
      <div class="card green">
        <h3>🏫 Session (Gap 4 min por DANE)</h3>
        <ul>
          <li>Duración dinámica por actividad del colegio.</li>
          <li>En <code>110010190012</code> (C. Bolívar): une 3 alertas en <strong>Sesión #1 (7.08 min)</strong> y abre <strong>Sesión #2</strong> tras 24 min de pausa.</li>
        </ul>
      </div>
    </div>
  </div>

  <!-- SLIDE 7 -->
  <div class="slide">
    <div class="header">
      <div><h1>6. Matriz Argumentada de Aportes por Integrante</h1><p>Qué desarrolló cada autor, cómo lo implementó y dónde está documentado en el manual APA 7</p></div>
      <div class="badge">Diapositiva 7/8 · 06:10 - 06:50<br>Exponen: Ambos Integrantes</div>
    </div>
    <div class="content grid-2">
      <div class="card">
        <h3>👨‍💻 David Santiago Castelblanco Artunduaga (67001182)</h3>
        <ul>
          <li><strong>Qué desarrolló:</strong> Arquitectura Lakehouse Medallón, motor híbrido PySpark/Pandas, cruce <code>LEFT JOIN</code>, reporte de equidad Gold, ventanas <em>Tumbling, Sliding y Session</em> (Corte 3) y figuras APA 7.</li>
          <li><strong>Cómo lo implementó:</strong> Funciones <code>crear_spark_session()</code>, <code>cruzar_saber11_con_directorio()</code>, <code>generar_reporte_equidad_educativa()</code> y motor de ventanas en <code>streaming_alertas_educativas.py</code>.</li>
          <li><strong>Secciones en el Manual:</strong> Secciones 1, 3, 5.1, 5.2, 5.6, 5.7, 7 (Figura 1), 9.2.1, 9.2.3 (Tabla 7, Figura 3), 9.3 (Tabla 10) y 10.1.</li>
        </ul>
      </div>
      <div class="card green">
        <h3>👨‍💻 Dairo José Narváez Morales (67002032)</h3>
        <ul>
          <li><strong>Qué desarrolló:</strong> Generador sintético RAW (87,520 filas, semilla 42), contrato Silver (centinelas y moda DANE), reglas de deserción (&gt;8.0%), absorción IES, <code>BrokerEventosEducativos</code> idempotente, Watermarks y cola DLQ (Corte 3).</li>
          <li><strong>Cómo lo implementó:</strong> Funciones <code>validar_rango_puntaje()</code>, <code>limpiar_valores_nulos()</code>, <code>evaluar_alerta_localidad()</code>, hash SHA-256 <code>anonimizar_estudiante()</code> y decisión tripartita WM/DLQ.</li>
          <li><strong>Secciones en el Manual:</strong> Secciones 2, 4 (Tabla 1), 5.3, 5.4, 5.5, 6 (Tabla 4), 8 (Tabla 6, Figura 2), 9.1, 9.2.2, 9.2.3 (Tabla 8), 9.3 (Tabla 9, Figura 4), 9.4 (Tabla 11), 10.2 y 11.</li>
        </ul>
      </div>
    </div>
  </div>

  <!-- SLIDE 8 -->
  <div class="slide">
    <div class="header">
      <div><h1>7. Demostración Funcional en Vivo y Conclusiones</h1><p>Ejecución en terminal de las 6 etapas del sistema y verificación de archivos en STAGING y ANALYTICS</p></div>
      <div class="badge">Diapositiva 8/8 · 06:50 - 08:00<br>Demostración Conjunta</div>
    </div>
    <div class="content grid-2">
      <pre>
# 1. Generación reproducible de la Capa Bronze (RAW)
PS Proyecto-ING-DATO&gt; python scripts/generar_datos_raw.py

# 2. Ejecución integral (Etapas 1-5 Batch + Etapa 6 Streaming Corte 3)
PS Proyecto-ING-DATO&gt; python scripts/pipeline_educativo.py

# 3. Ejecución aislada del motor de Streaming (Corte 3)
PS Proyecto-ING-DATO&gt; python scripts/streaming_alertas_educativas.py</pre>
      <div class="card green">
        <h3>🎯 Conclusiones del Proyecto Integrador</h3>
        <ul>
          <li><strong>Gobierno y Calidad:</strong> 0 archivos RAW alterados, 1,070 centinelas filtrados y 2,934 códigos DANE recuperados por moda estadística.</li>
          <li><strong>Streaming Resiliente (Corte 3):</strong> <code>Event Time</code> + <code>Watermark (5m)</code> + <code>Allowed Lateness (+5m)</code> + <code>Side Output / DLQ</code> garantiza que ninguna alerta estudiantil rural (como Sumapaz con 28 min de retraso) se pierda.</li>
        </ul>
      </div>
    </div>
  </div>

  <div class="footer">
    <div>Ingeniería de Datos — Universidad Católica de Colombia | Castelblanco (67001182) &amp; Narváez (67002032)</div>
    <div class="controls">
      <span id="counter">1 / 8</span>
      <button onclick="move(-1)">&#8592; Anterior</button>
      <button onclick="move(1)">Siguiente &#8594;</button>
    </div>
  </div>
</div>
<script>
  let idx = 0;
  const slides = document.querySelectorAll('.slide');
  const counter = document.getElementById('counter');
  function move(step) {
    slides[idx].classList.remove('active');
    idx = (idx + step + slides.length) % slides.length;
    slides[idx].classList.add('active');
    counter.textContent = (idx + 1) + ' / ' + slides.length;
  }
  document.addEventListener('keydown', (e) => {
    if (e.key === 'ArrowRight' || e.key === ' ') move(1);
    if (e.key === 'ArrowLeft') move(-1);
  });
</script>
</body>
</html>"""

    for d in DEST_DIRS:
        out_html = os.path.join(d, "Presentacion_Interactiva_Sustentacion.html")
        with open(out_html, "w", encoding="utf-8") as f:
            f.write(html_content)
        print("✅ Presentación HTML interactiva guardada en:", out_html)


# ══════════════════════════════════════════════════════════════════════════════
# 3. GENERACIÓN DEL GUION DE SUSTENTACIÓN PARA EL VIDEO (5-10 MINUTOS)
# ══════════════════════════════════════════════════════════════════════════════
def set_doc_cell_bg(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    tcPr.append(shd)


def set_doc_cell_margins(cell, top=90, bottom=90, left=130, right=130):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for m, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        node = OxmlElement(f"w:{m}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def construir_guion_docx():
    doc = Document()
    for s in doc.sections:
        s.top_margin = DocInches(0.8)
        s.bottom_margin = DocInches(0.8)
        s.left_margin = DocInches(0.85)
        s.right_margin = DocInches(0.85)

    st = doc.styles["Normal"]
    st.font.name = "Segoe UI"
    st.font.size = DocPt(10.5)
    st.font.color.rgb = DocRGB(30, 41, 59)

    p_hdr = doc.add_paragraph()
    p_hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p_hdr.add_run("UNIVERSIDAD CATÓLICA DE COLOMBIA · INGENIERÍA DE DATOS (ENTREGA FINAL CORTE 3)\n")
    r1.font.size = DocPt(10)
    r1.font.bold = True
    r1.font.color.rgb = DocRGB(37, 99, 235)

    r2 = p_hdr.add_run("GUION DE SUSTENTACIÓN EN VIDEO Y DEMOSTRACIÓN FUNCIONAL (DURACIÓN: 8 MINUTOS)\n")
    r2.font.size = DocPt(16)
    r2.font.bold = True
    r2.font.color.rgb = DocRGB(15, 23, 42)

    r3 = p_hdr.add_run(
        "Integrantes en cámara: David Santiago Castelblanco Artunduaga (67001182) y Dairo José Narváez Morales (67002032)\n"
        "Rango oficial exigido: Mínimo 5 y máximo 10 minutos | Visibilidad en YouTube: Pública"
    )
    r3.font.size = DocPt(10)
    r3.font.italic = True
    r3.font.color.rgb = DocRGB(100, 116, 139)

    h1 = doc.add_heading("1. Distribución Equitativa de Tiempos y Cumplimiento de la Rúbrica", level=1)
    h1.runs[0].font.color.rgb = DocRGB(15, 23, 42)

    t_cron = doc.add_table(rows=9, cols=4)
    t_cron.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers = ["Bloque / Minuto", "Diapositiva / Pantalla", "Integrante en Cámara", "Objetivo Clave Evaluado por la Docente"]
    for j, h in enumerate(headers):
        c = t_cron.rows[0].cells[j]
        set_doc_cell_bg(c, "0F172A")
        set_doc_cell_margins(c)
        p = c.paragraphs[0]
        r = p.add_run(h)
        r.font.bold = True
        r.font.color.rgb = DocRGB(255, 255, 255)

    cron_data = [
        ("00:00 – 00:40 (40 s)", "Diapositiva 1 (Portada)", "David Castelblanco y Dairo Narváez", "Presentación en cámara de ambos integrantes y propósito general del sistema."),
        ("00:40 – 01:40 (60 s)", "Diapositiva 2 (Problemática y Ética)", "Dairo José Narváez", "Cifras DANE/SaluData, 4 patologías de calidad y cumplimiento Ley 1581 (SHA-256)."),
        ("01:40 – 02:45 (65 s)", "Diapositiva 3 (Arquitectura Híbrida)", "David Santiago Castelblanco", "Conceptos Cortes 1, 2 y 3: Lakehouse Medallón (Bronze/Silver/Gold) + Streaming EDA."),
        ("02:45 – 04:00 (75 s)", "Diapositivas 4 y 5 (Código Batch)", "Dairo Narváez y David Castelblanco", "Explicación técnica de código: centinelas O(1), moda DANE, resiliencia Spark y LEFT JOIN."),
        ("04:00 – 05:10 (70 s)", "Diapositiva 6 (Corte 3: EDA y WM)", "Dairo José Narváez", "CloudEvents, idempotencia, Event Time vs. Processing Time, Watermark y cola DLQ."),
        ("05:10 – 06:05 (55 s)", "Diapositiva 7 (Corte 3: Ventanas)", "David Santiago Castelblanco", "Ventanas Tumbling (10m), Sliding (10m/2m ráfaga Bosa), Session (gap 4m) y Spark vs. Flink."),
        ("06:05 – 06:45 (40 s)", "Diapositiva 8 (Aportes Individuales)", "David Castelblanco y Dairo Narváez", "Argumentación explícita de qué desarrolló cada integrante y en qué sección del manual está."),
        ("06:45 – 08:00 (75 s)", "Diapositivas 9–10 + Terminal en Vivo", "Demostración Conjunta en Pantalla", "Ejecución real en consola de pipeline_educativo.py (Etapas 1 a 6) y cierre."),
    ]
    for i, row in enumerate(cron_data, start=1):
        bg = "F8FAFC" if i % 2 == 1 else "EFF6FF"
        for j, val in enumerate(row):
            c = t_cron.rows[i].cells[j]
            set_doc_cell_bg(c, bg)
            set_doc_cell_margins(c)
            p = c.paragraphs[0]
            r = p.add_run(val)
            if j == 0:
                r.font.bold = True

    h2 = doc.add_heading("2. Guion Conversacional Paso a Paso (Listo para Grabar)", level=1)
    h2.runs[0].font.color.rgb = DocRGB(15, 23, 42)

    bloques = [
        (
            "BLOQUE 1 — APERTURA Y PORTADA (00:00 – 00:40) | Diapositiva 1",
            "Ambos integrantes con cámara encendida y presentación proyectada.",
            "David Santiago Castelblanco (00:00 – 00:25):",
            "«Buenas tardes, ingeniera Diana Marcela. Mi nombre es David Santiago Castelblanco Artunduaga (código 67001182) y junto con mi compañero Dairo José Narváez Morales (código 67002032) presentamos la sustentación final de nuestro proyecto integrador: Sistema Integrado de Análisis y Calidad de Datos para el Ecosistema Educativo Distrital de Bogotá D.C., en su versión final 2.0, que integra nuestra arquitectura Data Lakehouse por capas con el motor de procesamiento de eventos en tiempo real del Tercer Corte.»\n\n"
            "Dairo José Narváez (00:25 – 00:40):\n"
            "«Buenas tardes, profesora. A lo largo de esta sustentación mostraremos cómo construimos la solución, qué conceptos de los tres cortes aplicamos, cuál fue el aporte específico de cada uno en el código y en el manual técnico bajo normas APA 7, y realizaremos la demostración funcional en vivo.»"
        ),
        (
            "BLOQUE 2 — PROBLEMÁTICA, CALIDAD Y MARCO ÉTICO (00:40 – 01:40) | Diapositiva 2",
            "Cámara de Dairo activa señalando la Diapositiva 2.",
            "Dairo José Narváez:",
            "«Nuestro proyecto parte de una realidad crítica: según el DANE (2025), la pobreza multidimensional en Bogotá subió al 5.4% y el rezago escolar alcanzó el 17.5%, con brechas territoriales severas donde localidades como Usme y Ciudad Bolívar superan el 57% de pobreza monetaria frente al 7.9% en Teusaquillo. Aunque la Secretaría de Educación publica datos abiertos valiosos, al perfilar los 87,520 registros crudos de la capa RAW identificamos cuatro patologías técnicas: primero, silos desconectados entre Saber 11, directorio de colegios, Jóvenes a la E y alertas; segundo, 1,070 registros con valores centinela como guiones o 'N/A' en puntajes numéricos; tercero, 2,934 estudiantes (el 4.2%) sin código DANE institucional; y cuarto, retrasos de red en reportes de colegios periféricos. Además, cumpliendo la Ley 1581 de 2012 de Habeas Data en menores de edad, aplicamos minimización en la capa Gold y anonimización criptográfica SHA-256 en el stream de eventos.»"
        ),
        (
            "BLOQUE 3 — ARQUITECTURA HÍBRIDA LAKEHOUSE + STREAMING (01:40 – 02:45) | Diapositiva 3",
            "Cámara de David Castelblanco señalando el diagrama de la Figura 1 en la Diapositiva 3.",
            "David Santiago Castelblanco:",
            "«Para resolver este problema aplicamos los conceptos de los tres cortes en una arquitectura híbrida, documentada en la Sección 5 y en la Figura 1 del manual técnico. Del Primer Corte aplicamos la transición de sistemas OLTP a OLAP mediante un Data Lakehouse con Arquitectura Medallón en tres zonas físicas: la capa Bronze (data/RAW/), estrictamente inmutable siguiendo el paradigma ELT; la capa Silver (data/STAGING/), donde se ejecuta el contrato de calidad DAMA-DMBOK; y la capa Gold (data/ANALYTICS/), estructurada como un modelo dimensional en Esquema de Estrella donde la tabla de hechos de Saber 11 se enriquece con las dimensiones de Colegio, Localidad y Programa Social. Y del Tercer Corte integramos una Arquitectura Orientada a Eventos (EDA) que ingiere alertas en formato CloudEvents 1.0, las procesa por Event Time con Watermarks y alimenta tres tipos de ventanas de tiempo y una cola de conciliación DLQ.»"
        ),
        (
            "BLOQUE 4 — IMPLEMENTACIÓN TÉCNICA BATCH: SILVER Y GOLD (02:45 – 04:00) | Diapositivas 4 y 5",
            "Intervienen Dairo (Diapositiva 4) y David (Diapositiva 5) explicando los fragmentos de código.",
            "Dairo José Narváez (02:45 – 03:20 — Diapositiva 4):",
            "«En la Diapositiva 4 vemos el código del contrato de calidad que implementé en scripts/pipeline_educativo.py. En la función validar_rango_puntaje() definimos VALORES_CENTINELA como un conjunto hash (set), logrando búsquedas en tiempo constante O(1) para convertir los 1,070 textos inválidos en nulos seguros y validar que cada puntaje esté entre 0 y 500. Luego, en limpiar_valores_nulos(), imputamos mediante la moda estadística los 2,934 códigos DANE faltantes, reduciendo los nulos a cero sin descartar estudiantes.»\n\n"
            "David Santiago Castelblanco (03:20 – 04:00 — Diapositiva 5):\n"
            "«En la Diapositiva 5 vemos dos decisiones clave que programé: primero, en crear_spark_session() implementé resiliencia híbrida que verifica con comprobar_java_compatible() si el sistema tiene Java 8; como PySpark 4 requiere Java 17+, el orquestador conmuta automáticamente al motor vectorizado de Pandas sin interrumpir la ejecución. Segundo, en cruzar_saber11_con_directorio() aplicamos un LEFT JOIN contra los 2,262 colegios activos, vinculando 63,498 estudiantes sin perder registros. Al cruzar con la regla de negocio de deserción superior al 8.0%, el sistema identifica automáticamente 9 localidades en alerta crítica, encabezadas por Usme (12.1%) y Fontibón (11.8%).»"
        ),
        (
            "BLOQUE 5 — TERCER CORTE (PARTE I): EDA, WATERMARKS Y EVENTOS TARDÍOS (04:00 – 05:10) | Diapositiva 6",
            "Cámara de Dairo explicando la Figura 4 y la traza de Watermarks.",
            "Dairo José Narváez:",
            "«Pasando a la implementación funcional del Tercer Corte en scripts/streaming_alertas_educativas.py, construí la clase BrokerEventosEducativos y la lógica de Watermarks, documentadas en la Sección 9.3 y la Figura 4 del manual. Primero, para garantizar idempotencia ante reintentos de red, el broker registra cada ID en un conjunto hash: de los 18 eventos CloudEvents recibidos en el stream JSONL, detecta y bloquea automáticamente el reintento duplicado del evento EVT-004 en Bosa. Segundo, demostramos cuantitativamente por qué no debemos agrupar por Processing Time: en el bloque de 08:00 a 08:10 ocurrieron 8 alertas reales en los colegios (Event Time), pero por retrasos de red el servidor solo vio 6 en Processing Time. Por eso implementamos un Watermark dinámico WM = max(event_time) - 5 minutos y tres políticas de eventos tardíos: el evento EVT-009 de Usme llega 4.3 minutos tarde pero con el Watermark en 08:08, por lo que entra normalmente a su ventana; el EVT-013 llega 12 minutos tarde cuando la ventana ya cerró, pero entra en nuestra holgura de Allowed Lateness (+5 min), reabriendo el estado y emitiendo una versión corregida; y el EVT-017 de un colegio rural en Sumapaz sin enlace satelital llega 28 minutos tarde: mientras Spark haría un Drop descartándolo, nuestro motor lo desvía al tópico Side Output / DLQ para que la Secretaría de Educación nunca pierda la alerta de un estudiante vulnerable.»"
        ),
        (
            "BLOQUE 6 — TERCER CORTE (PARTE II): VENTANAS TUMBLING, SLIDING Y SESSION (05:10 – 06:05) | Diapositiva 7",
            "Cámara de David explicando los 3 tipos de ventanas y el código PySpark.",
            "David Santiago Castelblanco:",
            "«En la Diapositiva 7 vemos los tres operadores de ventanas de tiempo que implementé sobre ese flujo y su equivalencia en PySpark Structured Streaming, documentados en la Tabla 10 del manual. Primero, las Tumbling Windows de 10 minutos sin traslape consolidan el reporte oficial periódico con costo de 1 ventana por evento en memoria. Segundo, las Sliding Windows de 10 minutos que avanzan cada 2 minutos (donde cada evento cae en 10 dividido 2 igual a 5 ventanas) nos permitieron detectar una ráfaga crítica de 3 alertas en Bosa en el intervalo [08:02, 08:12), la cual pasaba desapercibida en la ventana Tumbling porque dos eventos caían antes de las 08:10 y el tercero a las 08:11. Y tercero, las Session Windows con brecha de inactividad (gap) de 4 minutos por código DANE agruparon dinámicamente las sesiones activas de cada orientador escolar, separando en el colegio de Ciudad Bolívar una primera sesión continua de 7.08 minutos con 3 alertas y una segunda sesión independiente a las 08:29.»"
        ),
        (
            "BLOQUE 7 — IDENTIFICACIÓN ARGUMENTADA DE APORTES INDIVIDUALES (06:05 – 06:45) | Diapositiva 8",
            "Ambos integrantes sintetizan su matriz de aportes de la Sección 10 del manual.",
            "David Santiago Castelblanco (06:05 – 06:25):",
            "«Como se detalla en la Sección 10 y la Tabla 12 del manual técnico: mi aporte se centró en la arquitectura Lakehouse, la resiliencia híbrida PySpark/Pandas, el cruce LEFT JOIN, la consolidación del reporte de equidad en la capa Gold, la programación de las ventanas Tumbling, Sliding y Session del Corte 3 y el script de figuras APA 7.»\n\n"
            "Dairo José Narváez (06:25 – 06:45):\n"
            "«Por mi parte, desarrollé el generador reproducible de la capa RAW, el contrato de limpieza e imputación por moda en la capa Silver, la regla automática de alertas de deserción y absorción IES, el broker EDA idempotente con anonimización SHA-256, el cálculo de Watermarks con desvío a la cola DLQ y el manual de usuario en la Sección 11.»"
        ),
        (
            "BLOQUE 8 — DEMOSTRACIÓN FUNCIONAL EN VIVO Y CIERRE (06:45 – 08:00) | Terminal + Diapositivas 9 y 10",
            "Se comparte la ventana de PowerShell en la carpeta Proyecto-ING-DATO y se ejecuta el comando en vivo.",
            "Demostración Conjunta (David ejecuta y comenta Etapas 1–4; Dairo comenta Etapas 5–6 y archivos CSV):",
            "1. Ejecutar en vivo en la terminal: python scripts/pipeline_educativo.py\n"
            "2. Señalar en consola la Etapa 1 (70,000 registros Saber 11 y diagnóstico de 2,934 nulos en DANE) y la Etapa 2 (1,070 centinelas filtrados e imputación por moda dejando 0 nulos en data/STAGING/).\n"
            "3. Señalar las Etapas 3, 4 y 5 (las 9 localidades con ALERTA CRÍTICA > 8.0%, el LEFT JOIN de 63,498 estudiantes vinculados y la tabla de equidad en data/ANALYTICS/).\n"
            "4. Señalar la Etapa 6 del Corte 3 en consola: 18 eventos CloudEvents leídos, 1 duplicado descartado por idempotencia (EVT-004), tabla comparativa Event Time vs. Processing Time, ráfagas detectadas por Sliding Windows en Ciudad Bolívar, Bosa y Kennedy, sesiones por colegio DANE y el evento EVT-017 de Sumapaz resguardado en streaming_eventos_tardios_dlq.csv.\n"
            "5. Despedida en cámara de ambos integrantes agradeciendo a la docente."
        ),
    ]

    for titulo_b, escena, orador, parlamento in bloques:
        tbl = doc.add_table(rows=2, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        c0 = tbl.rows[0].cells[0]
        set_doc_cell_bg(c0, "1E293B")
        set_doc_cell_margins(c0, top=70, bottom=70, left=120, right=120)
        p0 = c0.paragraphs[0]
        r0 = p0.add_run(f"{titulo_b}\n🎬 Indicación visual: {escena}")
        r0.font.bold = True
        r0.font.size = DocPt(9.5)
        r0.font.color.rgb = DocRGB(255, 255, 255)

        c1 = tbl.rows[1].cells[0]
        set_doc_cell_bg(c1, "F8FAFC")
        set_doc_cell_margins(c1, top=90, bottom=90, left=120, right=120)
        p1 = c1.paragraphs[0]
        ro = p1.add_run(orador + "\n")
        ro.font.bold = True
        ro.font.color.rgb = DocRGB(37, 99, 235)
        rp = p1.add_run(parlamento)
        rp.font.size = DocPt(10)

        doc.add_paragraph()

    h3 = doc.add_heading("3. Lista de Verificación Final antes de Enviar (Checklist)", level=1)
    h3.runs[0].font.color.rgb = DocRGB(15, 23, 42)
    checks = [
        "☑ Ambos integrantes (David Santiago Castelblanco y Dairo José Narváez) aparecen en cámara durante el video.",
        "☑ El video dura entre 5 y 10 minutos (este guion está calibrado exactamente en 8 minutos).",
        "☑ Se utiliza la presentación de apoyo (Presentacion_Sustentacion_Entrega_Final.pptx o el HTML interactivo) y se ejecuta el pipeline en vivo.",
        "☑ Al subir el video a YouTube, configurar la visibilidad como PÚBLICA (no Oculto ni Privado).",
        "☑ Copiar el enlace de YouTube y pegarlo en el recuadro amarillo de la Página 2 de 'Documento Integrado del Proyecto.docx'.",
    ]
    for ch in checks:
        p = doc.add_paragraph(ch)
        p.paragraph_format.space_after = DocPt(4)

    for d in DEST_DIRS:
        out_doc = os.path.join(d, "Guion_Video_Sustentacion_Entrega_Final.docx")
        doc.save(out_doc)
        print("✅ Guion de sustentación guardado en:", out_doc)


if __name__ == "__main__":
    construir_pptx()
    construir_html_interactivo()
    construir_guion_docx()
