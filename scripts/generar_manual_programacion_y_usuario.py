# -*- coding: utf-8 -*-
"""
generar_manual_programacion_y_usuario.py

Genera y actualiza el documento 'Plantillas_Manual_Programacion_Codigo_y_Usuario.docx'
(Manual de Programación y Código Fuente + Manual de Usuario) en su versión final v2.0.0
(22/11/2026), integrando los módulos Batch Lakehouse y Streaming del Corte 3, el enlace
de YouTube en la segunda página, los aportes argumentados de cada integrante, las
evidencias reales y cumpliendo el formato solicitado:
  - Fuente Arial 12 pt en todo el documento
  - Texto justificado
  - Todo en color negro (#000000)
  - Todas las tablas con bordes negros completos y SIN colores de fondo
  - Cero guiones largos ("—")
"""

import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

BASE_REPO = r"C:\Users\dsant\Desktop\U CATOLICA\ING SISTEMAS\SEMESTRES SISTEMAS\noveno semestre\Ingenieria de datos\Proyecto ING Datos\ing datos\Proyecto-ING-DATO"
FIGS_DIR  = os.path.join(BASE_REPO, "data", "ANALYTICS", "figuras")

BLACK = RGBColor(0, 0, 0)
FONT_NAME = "Arial"
FONT_SIZE = Pt(12)


def limpiar_guiones(texto: str) -> str:
    return texto.replace(" — ", ", ").replace("—", "-").replace(" – ", " a ").replace("–", "-")


def set_cell_margins(cell, top=75, bottom=75, left=105, right=105):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for m, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        node = OxmlElement(f"w:{m}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def set_full_black_borders(table):
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for b_name in ["top", "bottom", "left", "right", "insideH", "insideV"]:
        b = OxmlElement(f"w:{b_name}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "6")
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), "000000")
        borders.append(b)
    tblPr.append(borders)


def add_page_number(section):
    header = section.header
    p = header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run()
    run.font.name = FONT_NAME
    run.font.size = FONT_SIZE
    run.font.color.rgb = BLACK
    fld1 = OxmlElement("w:fldChar")
    fld1.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    fld2 = OxmlElement("w:fldChar")
    fld2.set(qn("w:fldCharType"), "separate")
    fld3 = OxmlElement("w:fldChar")
    fld3.set(qn("w:fldCharType"), "end")
    run._r.extend([fld1, instr, fld2, fld3])


def add_heading_apa(doc, text, level=1):
    text = limpiar_guiones(text)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.first_line_indent = Inches(0)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if level == 1 else WD_ALIGN_PARAGRAPH.LEFT
    r = p.add_run(text)
    r.font.name = FONT_NAME
    r.font.size = FONT_SIZE
    r.font.bold = True
    r.font.italic = (level >= 3)
    r.font.color.rgb = BLACK
    return p


def add_para(doc, text, indent=True, bold_prefix=None):
    text = limpiar_guiones(text)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.first_line_indent = Inches(0.5) if indent else Inches(0)
    if bold_prefix:
        rb = p.add_run(limpiar_guiones(bold_prefix))
        rb.font.name = FONT_NAME
        rb.font.size = FONT_SIZE
        rb.font.bold = True
        rb.font.color.rgb = BLACK
    r = p.add_run(text)
    r.font.name = FONT_NAME
    r.font.size = FONT_SIZE
    r.font.color.rgb = BLACK
    return p


def add_table_full(doc, headers, rows, bold_first_col=False):
    tbl = doc.add_table(rows=len(rows) + 1, cols=len(headers))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_full_black_borders(tbl)
    for j, h in enumerate(headers):
        c = tbl.rows[0].cells[j]
        set_cell_margins(c)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.line_spacing = 1.15
        r = p.add_run(limpiar_guiones(str(h)))
        r.font.name = FONT_NAME
        r.font.size = FONT_SIZE
        r.font.bold = True
        r.font.color.rgb = BLACK
    for i, row in enumerate(rows, start=1):
        for j, val in enumerate(row):
            c = tbl.rows[i].cells[j]
            set_cell_margins(c)
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.line_spacing = 1.15
            r = p.add_run(limpiar_guiones(str(val)))
            r.font.name = FONT_NAME
            r.font.size = FONT_SIZE
            r.font.bold = (bold_first_col and j == 0)
            r.font.color.rgb = BLACK
    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_after = Pt(4)
    return tbl


def add_code_table(doc, label, code_str):
    tbl = doc.add_table(rows=2, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_full_black_borders(tbl)
    c0 = tbl.rows[0].cells[0]
    set_cell_margins(c0)
    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r0 = p0.add_run(limpiar_guiones(label))
    r0.font.name = FONT_NAME
    r0.font.size = FONT_SIZE
    r0.font.bold = True
    r0.font.color.rgb = BLACK

    c1 = tbl.rows[1].cells[0]
    set_cell_margins(c1)
    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p1.paragraph_format.line_spacing = 1.15
    r1 = p1.add_run(limpiar_guiones(code_str.strip()))
    r1.font.name = FONT_NAME
    r1.font.size = FONT_SIZE
    r1.font.color.rgb = BLACK
    doc.add_paragraph()


def construir_manual_programacion_y_usuario():
    doc = Document()
    for s in doc.sections:
        s.top_margin = Inches(1.0)
        s.bottom_margin = Inches(1.0)
        s.left_margin = Inches(1.0)
        s.right_margin = Inches(1.0)
        add_page_number(s)

    st = doc.styles["Normal"]
    st.font.name = FONT_NAME
    st.font.size = FONT_SIZE
    st.font.color.rgb = BLACK

    # ── PÁGINA 1: PORTADA ──
    for _ in range(3):
        doc.add_paragraph()

    p_t = doc.add_paragraph()
    p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t.paragraph_format.line_spacing = 1.5
    rt = p_t.add_run(
        "DOCUMENTACIÓN TÉCNICA DE SOFTWARE: MANUAL DE PROGRAMACIÓN Y CÓDIGO FUENTE + MANUAL DE USUARIO\n"
        "Sistema Integrado de Análisis y Calidad de Datos para el Ecosistema Educativo Distrital de Bogotá D.C. (Versión 2.0.0)"
    )
    rt.font.name = FONT_NAME
    rt.font.size = FONT_SIZE
    rt.font.bold = True
    rt.font.color.rgb = BLACK

    for _ in range(3):
        doc.add_paragraph()

    for line in [
        "David Santiago Castelblanco Artunduaga (Código: 67001182)",
        "Dairo José Narváez Morales (Código: 67002032)",
        "",
        "Facultad de Ingeniería, Programa de Ingeniería de Sistemas",
        "Universidad Católica de Colombia",
        "Asignatura: Ingeniería de Datos",
        "Docente: Ing. Diana Marcela Toquica Rodríguez",
        "22 de noviembre de 2026",
    ]:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.line_spacing = 1.5
        r = p.add_run(line)
        r.font.name = FONT_NAME
        r.font.size = FONT_SIZE
        r.font.color.rgb = BLACK

    doc.add_page_break()

    # ── PÁGINA 2: ENLACE AL VIDEO DE YOUTUBE + FICHA GENERAL DEL SISTEMA ──
    add_heading_apa(doc, "Enlace del Video de Sustentación en YouTube y Ficha Técnica del Sistema", level=1)
    add_para(
        doc,
        "De acuerdo con los lineamientos de la entrega final del proyecto de Ingeniería de Datos, "
        "en esta segunda página se registra el enlace público de YouTube del video de sustentación "
        "(duración entre 5 y 10 minutos, con aparición en cámara de ambos integrantes, apoyo en "
        "presentación y demostración funcional de la solución), seguido de la ficha general del software.",
        indent=True,
    )

    add_table_full(
        doc,
        headers=["Atributo de Entrega", "Información Oficial del Proyecto"],
        rows=[
            ["Enlace Público de YouTube (Sustentación)", "https://youtu.be/[PEGAR_AQUI_ENLACE_PUBLICO_YOUTUBE]"],
            ["Nombre del sistema", "Sistema Integrado de Análisis y Calidad de Datos Educativos Bogotá D.C."],
            ["Versión / etiqueta del código", "v2.0.0 (Entrega Final, Integración Batch Lakehouse + Streaming Corte 3)"],
            ["Lenguaje(s) y librerías", "Python 3.10+, PySpark 4.x, Pandas, NumPy, Matplotlib, python-docx"],
            ["Repositorio oficial GitHub", "https://github.com/HCA-3/Proyecto-ING-DATO.git"],
            ["Autores y roles principales", "David Santiago Castelblanco Artunduaga (67001182): Lakehouse, PySpark/Pandas, Joins y Ventanas de Tiempo | Dairo José Narváez Morales (67002032): Ingesta RAW, Limpieza Silver, Reglas de Alerta, Watermarks/DLQ y Manual de Usuario"],
            ["Fecha de entrega", "22 de noviembre de 2026"],
        ],
        bold_first_col=True,
    )

    add_heading_apa(doc, "Control de Cambios del Software", level=2)
    add_table_full(
        doc,
        headers=["Versión", "Fecha", "Autores", "Cambios Principales Implementados"],
        rows=[
            ["v0.5.0", "18/08/2026", "Castelblanco / Narváez", "Creación de scripts/generar_datos_raw.py y definición de esquemas de los 4 datasets sintéticos de Bogotá D.C. en data/RAW/."],
            ["v1.0.0", "02/11/2026", "Castelblanco / Narváez", "Construcción de scripts/pipeline_educativo.py (RAW -> STAGING -> ANALYTICS), limpieza de centinelas, imputación por moda DANE, cruces LEFT JOIN y alertas (> 8.0%)."],
            ["v2.0.0", "22/11/2026", "Castelblanco / Narváez", "Implementación de scripts/streaming_alertas_educativas.py (Etapa 6 Corte 3: Broker EDA idempotente, Event Time vs. Processing Time, Watermarks, ventanas Tumbling/Sliding/Session y cola DLQ)."],
        ],
    )

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════════
    # PARTE 1: MANUAL DE PROGRAMACIÓN Y CÓDIGO FUENTE
    # ══════════════════════════════════════════════════════════════════════════
    add_heading_apa(doc, "PARTE 1. MANUAL DE PROGRAMACIÓN Y CÓDIGO FUENTE", level=1)

    add_heading_apa(doc, "1. Propósito, Alcance y Conceptos de la Asignatura Aplicados", level=2)
    add_para(
        doc,
        "El software resuelve la dispersión, falta de estandarización y problemas de calidad "
        "que afectan los datos abiertos del ecosistema educativo en Bogotá D.C., integrando en "
        "una sola solución reproducible tanto el procesamiento histórico por lotes (Batch Lakehouse) "
        "como el monitoreo continuo de alertas de deserción e inasistencia en tiempo real (Streaming).",
        indent=True,
    )
    add_para(
        doc,
        "En su construcción se aplicaron directamente los conceptos de la asignatura que aportan "
        "valor técnico al proyecto: (1) Arquitectura Data Lakehouse bajo el patrón Medallón "
        "(capas Bronze/RAW inmutable, Silver/STAGING depurada y Gold/ANALYTICS dimensional); "
        "(2) Modelado Dimensional en Esquema de Estrella (tabla de hechos de rendimiento Saber 11 "
        "cruzada con dimensiones de Colegio, Localidad y Programa Social); (3) Gestión de Calidad "
        "de Datos DAMA-DMBOK (validez mediante filtrado O(1) de valores centinela, completitud "
        "mediante imputación estadística por moda de códigos DANE y consistencia referencial); "
        "(4) Procesamiento Distribuido con PySpark y resiliencia híbrida con Pandas; y "
        "(5) Procesamiento de Eventos en Tiempo Real del Tercer Corte (Arquitectura Orientada a "
        "Eventos con estándar CloudEvents 1.0, consumidor idempotente, procesamiento por Event Time "
        "con Watermarks de 5 minutos, ventanas Tumbling, Sliding y Session, y cola DLQ para "
        "resguardar reportes rurales tardíos).",
        indent=True,
    )

    add_heading_apa(doc, "2. Estructura del Proyecto y Responsabilidad de cada Archivo", level=2)
    add_table_full(
        doc,
        headers=["Carpeta / Archivo", "Tipo de Elemento", "Responsabilidad Técnica en el Sistema", "Relación con Otros Componentes"],
        rows=[
            ["data/RAW/", "Directorio (Capa Bronze)", "Almacena de forma estrictamente inmutable los 4 archivos CSV crudos (87,520 filas) y el archivo JSONL de eventos (18 eventos).", "Generado por generar_datos_raw.py y streaming_alertas_educativas.py; leído por pipeline_educativo.py."],
            ["data/STAGING/", "Directorio (Capa Silver)", "Alberga los datasets limpios Saber11_2025_Bogota_clean.csv (0 nulos en DANE, 0 centinelas) y Directorio_Colegios_clean.csv (2,262 activos).", "Generado por la Etapa 2 de pipeline_educativo.py."],
            ["data/ANALYTICS/", "Directorio (Capa Gold)", "Consolida los 5 reportes CSV de equidad, absorción IES y alertas batch + los 6 reportes CSV de streaming, Watermarks, ventanas y DLQ.", "Generado por las Etapas 5 y 6 de pipeline_educativo.py."],
            ["scripts/generar_datos_raw.py", "Módulo generador RAW", "Genera con semilla fija (seed=42) los 4 archivos CSV crudos simulando las 20 localidades de Bogotá e inyectando anomalías reales.", "Escribe en data/RAW/ e imprime el perfilamiento inicial."],
            ["scripts/pipeline_educativo.py", "Módulo orquestador principal", "Coordina las Etapas 1 a 6: detección Java/Spark, perfilamiento, limpieza Silver, reglas de negocio, LEFT JOIN Gold e invocación de streaming.", "Consume data/RAW/, produce data/STAGING/ y data/ANALYTICS/, e invoca streaming_alertas_educativas.py."],
            ["scripts/streaming_alertas_educativas.py", "Módulo de Streaming (Corte 3)", "Implementa BrokerEventosEducativos (EDA idempotente), cálculo de Watermarks, ventanas Tumbling/Sliding/Session y desvío a cola DLQ.", "Lee y escribe eventos en data/RAW/ y persiste 6 tablas analíticas en data/ANALYTICS/."],
            ["scripts/generar_graficos_evidencias.py", "Módulo de visualización", "Construye a partir de data/ANALYTICS/ las 4 figuras técnicas de arquitectura, calidad, equidad y Watermarks.", "Guarda imágenes PNG en data/ANALYTICS/figuras/."],
        ],
        bold_first_col=True,
    )

    add_heading_apa(doc, "3. Arquitectura Lógica y Flujo de Ejecución (Etapas 1 a 6)", level=2)
    add_para(
        doc,
        "Al ejecutar scripts/pipeline_educativo.py, el sistema avanza secuencialmente por seis "
        "etapas controladas: Etapa 1 (Inicialización híbrida Spark/Pandas, lectura desde data/RAW/ "
        "y diagnóstico automático de nulos); Etapa 2 (Contrato de calidad en capa Silver: conversión "
        "de 1,070 centinelas en PUNT_GLOBAL, imputación por moda de 2,934 códigos DANE faltantes "
        "y filtrado de 2,262 colegios ACTIVOS hacia data/STAGING/); Etapa 3 (Cálculo de promedios "
        "Saber 11 por localidad, evaluación de la regla de alertas de deserción > 8.0% y tasa de "
        "graduación en 10 IES); Etapa 4 (Cruce relacional LEFT JOIN entre los 70,000 estudiantes "
        "y los colegios activos); Etapa 5 (Exportación del datamart de equidad en data/ANALYTICS/); "
        "y Etapa 6 (Ejecución del motor de streaming del Corte 3 sobre el flujo de alertas en "
        "formato CloudEvents 1.0 con Watermarks, ventanas Tumbling, Sliding y Session, y cola DLQ).",
        indent=True,
    )

    add_heading_apa(doc, "4. Fichas Técnicas de Módulos, Clases y Funciones Principales", level=2)
    add_table_full(
        doc,
        headers=["Módulo / Clase / Función", "Parámetros de Entrada", "Lógica Interna y Algoritmo", "Salida / Efecto en el Sistema"],
        rows=[
            ["comprobar_java_compatible() y crear_spark_session()", "Ninguno", "Ejecuta 'java -version' en subproceso. Si detecta Java 8 (incompatible con Spark 4.x que exige Java 17+), retorna None y activa Pandas; si hay Java 17+, inicia SparkSession local[*].", "Retorna objeto SparkSession activo o None (resiliencia híbrida transparente)."],
            ["validar_rango_puntaje(valor_str)", "valor_str: str, float o None", "Limpia espacios y verifica en tiempo constante O(1) si el texto está en el conjunto hash VALORES_CENTINELA. Si es numérico, valida que pertenezca al rango [0.0, 500.0].", "Retorna float válido entre 0.0 y 500.0, o None si es centinela o fuera de rango."],
            ["limpiar_valores_nulos(df, columna, metodo)", "df: DataFrame, columna: str, metodo: 'moda' | 'mediana'", "Cuenta nulos iniciales. Si metodo=='moda', calcula valores_validos.mode()[0] y rellena los vacíos con fillna(), salvando los 2,934 estudiantes sin código DANE.", "Retorna DataFrame con 0 nulos en la columna indicada."],
            ["evaluar_alerta_localidad(nombre_localidad, tasa_desercion)", "nombre_localidad: str, tasa_desercion: float", "Valida tipo numérico con bloque try/except (ValueError, TypeError) y compara la tasa contra UMBRAL_ALERTA (8.0%).", "Retorna mensaje de ALERTA CRÍTICA si supera 8.0% o confirmación dentro del límite."],
            ["cruzar_saber11_con_directorio(df_saber11, df_dir, es_spark)", "df_saber11, df_directorio, es_spark: bool", "Selecciona atributos clave del directorio de colegios activos y ejecuta un LEFT JOIN por COLE_COD_DANE_INST == COD_DANE en PySpark o Pandas.", "Retorna DataFrame enriquecido de 70,000 filas (63,498 vinculadas a colegio activo)."],
            ["BrokerEventosEducativos (Clase en streaming_alertas_educativas.py)", "topico: str, evento_cloudevent: dict", "Administra tópicos en memoria con offsets secuenciales. Verifica si event_id ya existe en el conjunto self.ids_procesados para descartar reintentos duplicados en O(1).", "Garantiza procesamiento idempotente (bloquea duplicado EVT-004) y enruta a suscriptores o DLQ."],
            ["procesar_flujo_con_watermarks_y_ventanas(eventos)", "eventos: list[dict] (CloudEvents 1.0)", "Calcula Watermark WM = max(event_time) - 5 min; clasifica cada arribo (En Ventana, Allowed Lateness +5m o DLQ); agrupa en ventanas Tumbling (10m), Sliding (10m/2m) y Session (gap 4m).", "Retorna diccionario con 6 DataFrames analíticos de streaming y los persiste en data/ANALYTICS/."],
        ],
        bold_first_col=True,
    )

    add_heading_apa(doc, "5. Explicación Detallada de Fragmentos de Código e Implementación", level=2)

    add_heading_apa(doc, "5.1 Limpieza O(1) de Valores Centinela e Imputación por Moda (Capa Silver)", level=3)
    add_para(
        doc,
        "En scripts/pipeline_educativo.py se implementó el contrato de calidad de la capa Silver. "
        "El uso de un conjunto hash (set) para VALORES_CENTINELA permite que las 70,000 filas de "
        "Saber 11 se evalúen en complejidad O(1) por registro, convirtiendo los 1,070 valores "
        "corruptos ('-', 'N/A', 'ND') en nulos numéricos seguros antes de calcular promedios. "
        "Seguidamente, limpiar_valores_nulos() aplica imputación por moda sobre COLE_COD_DANE_INST, "
        "recuperando los 2,934 registros huérfanos.",
        indent=True,
    )
    add_code_table(
        doc,
        "Archivo: scripts/pipeline_educativo.py | Funciones: validar_rango_puntaje() y limpiar_valores_nulos()",
        """VALORES_CENTINELA: set = {"-", "N/A", "ND", "", "nan", "none"}

def validar_rango_puntaje(valor_str) -> float | None:
    if valor_str is None or pd.isna(valor_str):
        return None
    val_limpio = str(valor_str).strip()
    if val_limpio.lower() in VALORES_CENTINELA:
        return None
    try:
        numero = float(val_limpio)
        return numero if 0.0 <= numero <= 500.0 else None
    except ValueError:
        return None

def limpiar_valores_nulos(df: pd.DataFrame, columna: str, metodo: str = "moda") -> pd.DataFrame:
    nulos_antes = df[columna].isna().sum()
    if nulos_antes == 0:
        return df
    if metodo == "moda":
        valores_validos = df[columna].dropna()
        valor_fill = valores_validos.mode()[0] if len(valores_validos) > 0 else "SIN_DATO"
    elif metodo == "mediana":
        valor_fill = df[columna].median()
    df[columna] = df[columna].fillna(valor_fill)
    return df""",
    )

    add_heading_apa(doc, "5.2 Regla de Alertas de Deserción (> 8.0%) y Cruce Dimensional LEFT JOIN (Capa Gold)", level=3)
    add_para(
        doc,
        "La función evaluar_alerta_localidad() automatiza la regla de negocio distrital fijada "
        "en UMBRAL_ALERTA = 8.0%. Al ejecutarse sobre las 20 localidades, detecta 9 localidades "
        "en estado crítico (Usme con 12.1%, Fontibón con 11.8%, Engativá con 11.3%, Los Mártires "
        "con 11.2%, Tunjuelito con 10.9%, Usaquén con 9.7%, Kennedy con 9.1%, La Candelaria con "
        "8.9% y Antonio Nariño con 8.1%). Por su parte, cruzar_saber11_con_directorio() ejecuta "
        "un LEFT JOIN tanto en PySpark como en Pandas para enriquecer los resultados académicos "
        "sin descartar estudiantes.",
        indent=True,
    )
    add_code_table(
        doc,
        "Archivo: scripts/pipeline_educativo.py | Funciones: evaluar_alerta_localidad() y cruzar_saber11_con_directorio()",
        """UMBRAL_ALERTA: float = 8.0

def evaluar_alerta_localidad(nombre_localidad: str, tasa_desercion) -> str:
    if tasa_desercion is None or (isinstance(tasa_desercion, (int, float)) and tasa_desercion < 0):
        return f"[{nombre_localidad}] Dato inválido: tasa de deserción no disponible."
    try:
        tasa = float(tasa_desercion)
    except (ValueError, TypeError):
        return f"[{nombre_localidad}] Error: formato de tasa no numérico."
    if tasa > UMBRAL_ALERTA:
        return (
            f"ALERTA CRÍTICA: {nombre_localidad}, "
            f"Tasa de deserción {tasa:.1f}% supera el umbral de {UMBRAL_ALERTA}%. "
            "Requiere intervención inmediata."
        )
    return f"OK: {nombre_localidad}, Tasa de deserción {tasa:.1f}% dentro del límite aceptable ({UMBRAL_ALERTA}%)."

def cruzar_saber11_con_directorio(df_saber11, df_directorio, es_spark=False):
    if es_spark:
        dir_sel = df_directorio.select(
            col("COD_DANE"), col("NOMBRE_INST").alias("DIR_NOMBRE"),
            col("LOCALIDAD").alias("DIR_LOCALIDAD"), col("SECTOR").alias("DIR_SECTOR"),
            col("JORNADA").alias("DIR_JORNADA"), col("ESTADO").alias("DIR_ESTADO")
        )
        return df_saber11.join(dir_sel, df_saber11["COLE_COD_DANE_INST"] == dir_sel["COD_DANE"], how="left").drop("COD_DANE")""",
    )

    add_heading_apa(doc, "5.3 Motor de Streaming: Broker Idempotente, Watermarks, Ventanas y DLQ (Corte 3)", level=3)
    add_para(
        doc,
        "En scripts/streaming_alertas_educativas.py se implementó el procesamiento continuo de "
        "alertas escolares de inasistencia y riesgo de deserción. La clase BrokerEventosEducativos "
        "garantiza idempotencia descartando reintentos de red con el mismo ID (bloquea la segunda "
        "copia de EVT-004 en Bosa). El motor temporal calcula el Watermark dinámico "
        "WM = max(event_time) - 5 minutos, evalúa si cada evento entra en su ventana abierta "
        "(EVT-009), si activa una re-emisión corregida por Allowed Lateness de +5 minutos (EVT-013 "
        "en Usme con 12 min de retraso) o si se desvía a la cola DLQ (EVT-017 en Sumapaz con "
        "28 min de retraso), y calcula las agregaciones sobre ventanas Tumbling (10 min), "
        "Sliding (10 min cada 2 min) y Session (gap de 4 min por código DANE).",
        indent=True,
    )
    add_code_table(
        doc,
        "Archivo: scripts/streaming_alertas_educativas.py | Clase: BrokerEventosEducativos y Motor de Watermarks/Ventanas",
        """class BrokerEventosEducativos:
    def __init__(self):
        self.topicos = {"sed.alertas.asistencia": [], "sed.alertas.dlq": []}
        self.suscriptores = {"sed.alertas.asistencia": [], "sed.alertas.dlq": []}
        self.ids_procesados: set[str] = set()
        self.duplicados_ignorados: list[dict] = []

    def publicar(self, topico: str, evento_cloudevent: dict) -> bool:
        event_id = evento_cloudevent.get("id", "")
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

# Clasificación de arribo frente al Watermark (WM = max(event_time) - 5 min):
if wm_previo is None or wm_previo < w_end:
    estado_arrivo = "ACEPTADO_EN_VENTANA"
elif wm_previo <= w_end + timedelta(minutes=ALLOWED_LATENESS_MIN):
    estado_arrivo = "ALLOWED_LATENESS"
else:
    estado_arrivo = "RECHAZADO_EXTREMO_DLQ"
    broker.publicar("sed.alertas.dlq", {**ev, "retraso_min": retraso_min})""",
    )

    add_heading_apa(doc, "6. Pruebas Ejecutadas y Verificación del Código", level=2)
    add_table_full(
        doc,
        headers=["Prueba", "Método / Componente", "Entrada", "Resultado Esperado", "Resultado Obtenido en Consola"],
        rows=[
            ["Puntaje centinela", "validar_rango_puntaje()", "'-' o 'N/A'", "Retornar None sin romper cálculos", "None retornado (1,070 centinelas limpiados)"],
            ["Imputación DANE", "limpiar_valores_nulos()", "2,934 nulos en COLE_COD_DANE_INST", "0 nulos tras imputar moda", "0 nulos residuales (moda: 1100127417599)"],
            ["Alerta crítica", "evaluar_alerta_localidad()", "loc='FONTIBÓN', tasa=11.77", "Mensaje ALERTA CRÍTICA (> 8.0%)", "ALERTA CRÍTICA emitida correctamente"],
            ["Idempotencia EDA", "BrokerEventosEducativos.publicar()", "Reintento duplicado de EVT-004", "Retornar False e ignorar duplicado", "Duplicado descartado; 17 eventos únicos"],
            ["Eventos tardíos y DLQ", "procesar_flujo_con_watermarks_y_ventanas()", "EVT-013 (+12m) y EVT-017 (+28m)", "EVT-013 corrige ventana; EVT-017 va a DLQ", "Ventana Usme corregida y EVT-017 en DLQ"],
        ],
        bold_first_col=True,
    )

    add_heading_apa(doc, "7. Identificación Clara y Argumentada de los Aportes de Cada Integrante", level=2)
    add_table_full(
        doc,
        headers=["Integrante y Código", "¿Qué Desarrolló?", "¿Cómo lo Implementó?", "Sección donde se Documenta"],
        rows=[
            [
                "David Santiago Castelblanco Artunduaga\n(Código: 67001182)",
                "1. Arquitectura Lakehouse (Bronze, Silver, Gold).\n2. Resiliencia híbrida PySpark 4.x / Pandas.\n3. Cruce LEFT JOIN y reporte de equidad Gold.\n4. Ventanas Tumbling (10m), Sliding (10m/2m) y Session (gap 4m) en streaming + figuras.",
                "Programó comprobar_java_compatible(), crear_spark_session(), cruzar_saber11_con_directorio() y generar_reporte_equidad_educativa() en pipeline_educativo.py. En streaming_alertas_educativas.py implementó obtener_inicio_ventana_fija(), el barrido 5x de Sliding Windows para ráfagas (>=3 alertas) y la agrupación de sesiones por código DANE.",
                "Parte 1: Secciones 1, 2, 3, 4, 5.2 y 5.3 del presente manual; y Secciones 5.1, 5.2, 5.6, 7, 9.2.1, 9.2.3, 9.3 y 10.1 del Documento Integrado."
            ],
            [
                "Dairo José Narváez Morales\n(Código: 67002032)",
                "1. Generador sintético RAW (87,520 filas, semilla 42) y perfilamiento.\n2. Contrato de calidad Silver (centinelas y moda DANE).\n3. Reglas de alertas (>8.0%) y absorción IES.\n4. Broker EDA idempotente, Watermarks, Allowed Lateness y cola DLQ + Manual de Usuario.",
                "Programó puntaje_raw() y get_dane() en generar_datos_raw.py. En pipeline_educativo.py implementó validar_rango_puntaje() con búsqueda O(1) en hash set, limpiar_valores_nulos() por moda y evaluar_alerta_localidad(). En streaming_alertas_educativas.py programó BrokerEventosEducativos, anonimizar_estudiante() (SHA-256) y la lógica de Watermark y DLQ.",
                "Parte 1: Secciones 4, 5.1, 5.2, 5.3 y 6; Parte 2 completa (Manual de Usuario); y Secciones 2, 4, 5.3, 5.4, 5.5, 6, 8, 9.1, 9.2.2, 9.3, 10.2 y 11 del Documento Integrado."
            ],
        ],
        bold_first_col=True,
    )

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════════
    # PARTE 2: MANUAL DE USUARIO (OPERACIÓN DEL SISTEMA)
    # ══════════════════════════════════════════════════════════════════════════
    add_heading_apa(doc, "PARTE 2. MANUAL DE USUARIO (OPERACIÓN PASO A PASO)", level=1)

    add_heading_apa(doc, "1. Introducción y Público Objetivo", level=2)
    add_para(
        doc,
        "Este manual explica de forma clara y paso a paso cómo operar el Sistema Integrado de "
        "Análisis y Calidad de Datos Educativos de Bogotá D.C. Está dirigido a analistas de la "
        "Secretaría de Educación del Distrito (SED), directivos docentes, investigadores y "
        "usuarios académicos que requieren ejecutar la limpieza de datos, generar los reportes "
        "de equidad y monitorear las alertas de deserción escolar sin necesidad de modificar el "
        "código fuente.",
        indent=True,
    )

    add_heading_apa(doc, "2. Requisitos para Utilizar el Sistema", level=2)
    add_table_full(
        doc,
        headers=["Requisito", "Especificación Mínima Requerida"],
        rows=[
            ["Equipo y Sistema Operativo", "Computador personal o servidor con Windows 10/11, Linux o macOS, mínimo 4 GB de memoria RAM."],
            ["Intérprete y Terminal", "Terminal de comandos (PowerShell o CMD) y Python 3.9 o superior instalado."],
            ["Librerías requeridas", "Ejecutar una sola vez en consola: pip install pandas numpy pyspark python-docx python-pptx matplotlib"],
            ["Permisos y conectividad", "Permisos estándar de lectura/escritura en la carpeta del proyecto. No requiere conexión a internet ni claves privadas."],
        ],
        bold_first_col=True,
    )

    add_heading_apa(doc, "3. Navegación y Directorios Visibles para el Usuario", level=2)
    add_table_full(
        doc,
        headers=["Elemento / Directorio", "Función para el Usuario", "Acción que Puede Realizar la Persona Usuaria"],
        rows=[
            ["Terminal / PowerShell", "Interfaz de ejecución y monitoreo en vivo.", "Ejecutar los comandos python y observar el diagnóstico y las alertas por etapa."],
            ["Carpeta data/RAW/", "Almacén de archivos fuente crudos (Capa Bronze).", "Verificar la existencia de los 4 archivos CSV y el flujo JSONL de entrada."],
            ["Carpeta data/STAGING/", "Almacén de datos depurados (Capa Silver).", "Consultar Saber11_2025_Bogota_clean.csv y Directorio_Colegios_clean.csv ya saneados."],
            ["Carpeta data/ANALYTICS/", "Zona de reportes finales y gráficas (Capa Gold).", "Abrir los 11 reportes CSV en Excel o Power BI y visualizar las 4 figuras en figuras/."],
        ],
        bold_first_col=True,
    )

    add_heading_apa(doc, "4. Procedimientos de Uso Paso a Paso", level=2)

    add_para(
        doc,
        "1. Abrir PowerShell y ubicarse en la carpeta raíz del proyecto: cd \"Proyecto-ING-DATO\". "
        "2. Ejecutar el comando: python scripts/generar_datos_raw.py y presionar Enter. "
        "3. Resultado esperado: El sistema crea los 4 archivos CSV en data/RAW/ (70,000 registros "
        "de Saber 11, 2,500 colegios, 15,000 beneficiarios de Jóvenes a la E y 20 localidades) "
        "y muestra en pantalla el resumen de perfilamiento inicial.",
        indent=False,
        bold_prefix="Tarea 1: Generar o actualizar los datos crudos (Capa Bronze / RAW). ",
    )

    add_para(
        doc,
        "1. En la misma ventana de terminal, ejecutar: python scripts/pipeline_educativo.py y "
        "presionar Enter. "
        "2. El sistema ejecutará automáticamente las 6 etapas: diagnóstico de nulos (Etapa 1), "
        "limpieza de centinelas e imputación por moda en data/STAGING/ (Etapa 2), evaluación de "
        "alertas críticas de deserción > 8.0% y graduación IES (Etapa 3), cruce LEFT JOIN con "
        "colegios activos (Etapa 4), generación del reporte de equidad en data/ANALYTICS/ (Etapa 5) "
        "y procesamiento del flujo de alertas en tiempo real con Watermarks, ventanas Tumbling, "
        "Sliding y Session, y cola DLQ (Etapa 6). "
        "3. Resultado esperado: Mensaje final confirmando que el pipeline integral fue ejecutado con éxito.",
        indent=False,
        bold_prefix="Tarea 2: Ejecutar el pipeline completo (Limpieza Silver, Cruces Gold y Streaming Corte 3). ",
    )

    add_para(
        doc,
        "1. Si desea correr únicamente la simulación de alertas en tiempo real del Tercer Corte, "
        "ejecutar: python scripts/streaming_alertas_educativas.py. "
        "2. Resultado esperado: Visualización en consola del descarte del evento duplicado EVT-004, "
        "la evolución del Watermark, la detección de ráfagas críticas en Ciudad Bolívar, Bosa y "
        "Kennedy, las sesiones por colegio DANE y el envío del reporte rural tardío EVT-017 "
        "(Sumapaz) a la cola DLQ.",
        indent=False,
        bold_prefix="Tarea 3: Ejecutar de manera individual el monitoreo de eventos en tiempo real (Corte 3). ",
    )

    add_para(
        doc,
        "1. Abrir el Explorador de Archivos de Windows e ingresar a la carpeta data/ANALYTICS/. "
        "2. Abrir en Microsoft Excel o importar en Power BI los archivos reporte_equidad_educativa.csv, "
        "alertas_criticas_localidad.csv, absorcion_educacion_superior.csv, "
        "streaming_ventanas_sliding_rafagas.csv y streaming_eventos_tardios_dlq.csv. "
        "3. Ingresar a data/ANALYTICS/figuras/ para consultar las 4 gráficas generadas.",
        indent=False,
        bold_prefix="Tarea 4: Consultar los reportes finales y gráficas para toma de decisiones. ",
    )

    add_heading_apa(doc, "5. Guía de Mensajes del Sistema y Solución de Problemas Frecuentes", level=2)
    add_table_full(
        doc,
        headers=["Mensaje o Problema en Consola", "Causa", "Solución Paso a Paso"],
        rows=[
            ["Error: Archivo data/RAW/... no encontrado", "No se han generado aún los archivos crudos de entrada.", "Ejecutar primero: python scripts/generar_datos_raw.py"],
            ["Entorno Java local (Java 8) no compatible con PySpark 4.x", "El computador tiene Java 8 y Spark 4 requiere Java 17+.", "No requiere acción: el programa activa automáticamente el motor Pandas y genera todos los reportes."],
            ["No module named 'pandas' / 'docx'", "Faltan librerías en el entorno Python activo.", "Ejecutar: pip install pandas numpy pyspark python-docx python-pptx matplotlib"],
            ["ALERTA CRÍTICA: [Localidad], Tasa supera 8.0%", "La tasa de deserción de la localidad supera el umbral permitido.", "Priorizar la localidad en los planes de intervención de la SED."],
            ["RÁFAGA CRÍTICA (3 alertas en 10 min)", "La ventana Sliding detectó concentración rápida de inasistencia.", "Consultar streaming_ventanas_sliding_rafagas.csv para focalizar apoyo psicosocial."],
        ],
        bold_first_col=True,
    )

    rutas_salida = [
        r"c:\Users\dsant\Desktop\U CATOLICA\Programas\Data\Documentos\Plantillas_Manual_Programacion_Codigo_y_Usuario.docx",
        r"c:\Users\dsant\Desktop\U CATOLICA\Programas\Data\entregables\entrega final\Plantillas_Manual_Programacion_Codigo_y_Usuario.docx",
        os.path.join(BASE_REPO, "entregables", "entrega final", "Plantillas_Manual_Programacion_Codigo_y_Usuario.docx"),
        os.path.join(BASE_REPO, "programas", "Plantillas_Manual_Programacion_Codigo_y_Usuario.docx"),
    ]
    for r in rutas_salida:
        os.makedirs(os.path.dirname(r), exist_ok=True)
        doc.save(r)
        print("Manual de Programación y Usuario guardado en:", r)


if __name__ == "__main__":
    construir_manual_programacion_y_usuario()
