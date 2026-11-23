# -*- coding: utf-8 -*-
"""
generar_documento_integrado_apa7.py

Genera el 'Documento Integrado del Proyecto.docx' y su respectivo PDF con:
  - Fuente Arial 12 pt en todo el documento
  - Texto justificado (excepto portada y títulos Nivel 1 centrados, títulos Nivel 2/3 y bloques de código alineados a la izquierda)
  - Todo el texto y bordes estrictamente en color negro (#000000) sobre fondo blanco (únicamente las imágenes conservan color)
  - Todas las tablas con todos sus bordes completos (exteriores e interiores) en negro
  - Cero guiones largos ("—") en todo el documento
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
WORKSPACE_DOCS = r"c:\Users\dsant\Desktop\U CATOLICA\Programas\Data\Documentos"

BLACK = RGBColor(0, 0, 0)
FONT_NAME = "Arial"
FONT_SIZE = Pt(12)


def limpiar_guiones(texto: str) -> str:
    """Garantiza que no exista ningún guion largo '—' ni '–' en el texto."""
    return texto.replace(" — ", ", ").replace("—", "-").replace(" – ", " a ").replace("–", "-")


def set_cell_margins(cell, top=80, bottom=80, left=110, right=110):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for m, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        node = OxmlElement(f"w:{m}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def set_full_black_table_borders(table):
    """Aplica todos los bordes (top, bottom, left, right, insideH, insideV) en color negro."""
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


def add_page_number_to_header(section):
    header = section.header
    p = header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run()
    run.font.name = FONT_NAME
    run.font.size = FONT_SIZE
    run.font.color.rgb = BLACK
    fldChar1 = OxmlElement("w:fldChar")
    fldChar1.set(qn("w:fldCharType"), "begin")
    instrText = OxmlElement("w:instrText")
    instrText.set(qn("xml:space"), "preserve")
    instrText.text = "PAGE"
    fldChar2 = OxmlElement("w:fldChar")
    fldChar2.set(qn("w:fldCharType"), "separate")
    fldChar3 = OxmlElement("w:fldChar")
    fldChar3.set(qn("w:fldCharType"), "end")
    r = run._r
    r.append(fldChar1)
    r.append(instrText)
    r.append(fldChar2)
    r.append(fldChar3)


def add_apa_heading(doc, text, level=1):
    text = limpiar_guiones(text)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.first_line_indent = Inches(0)
    if level == 1:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(text)
        r.font.name = FONT_NAME
        r.font.size = FONT_SIZE
        r.font.bold = True
        r.font.color.rgb = BLACK
    elif level == 2:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(text)
        r.font.name = FONT_NAME
        r.font.size = FONT_SIZE
        r.font.bold = True
        r.font.color.rgb = BLACK
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(text)
        r.font.name = FONT_NAME
        r.font.size = FONT_SIZE
        r.font.bold = True
        r.font.italic = True
        r.font.color.rgb = BLACK
    return p


def add_apa_paragraph(doc, text, indent=True, bold_prefix=None):
    text = limpiar_guiones(text)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.first_line_indent = Inches(0.5) if indent else Inches(0)
    if bold_prefix:
        bold_prefix = limpiar_guiones(bold_prefix)
        rb = p.add_run(bold_prefix)
        rb.font.name = FONT_NAME
        rb.font.size = FONT_SIZE
        rb.font.bold = True
        rb.font.color.rgb = BLACK
    r = p.add_run(text)
    r.font.name = FONT_NAME
    r.font.size = FONT_SIZE
    r.font.color.rgb = BLACK
    return p


def add_apa_bullet(doc, bold_lead, text):
    text = limpiar_guiones(text)
    p = doc.add_paragraph(style="List Bullet")
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(4)
    if bold_lead:
        bold_lead = limpiar_guiones(bold_lead)
        rb = p.add_run(bold_lead)
        rb.font.name = FONT_NAME
        rb.font.size = FONT_SIZE
        rb.font.bold = True
        rb.font.color.rgb = BLACK
    r = p.add_run(text)
    r.font.name = FONT_NAME
    r.font.size = FONT_SIZE
    r.font.color.rgb = BLACK
    return p


def add_apa_table(doc, num, title, headers, rows, note_text):
    title = limpiar_guiones(title)
    note_text = limpiar_guiones(note_text)

    p_num = doc.add_paragraph()
    p_num.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_num.paragraph_format.space_before = Pt(10)
    p_num.paragraph_format.space_after = Pt(2)
    p_num.paragraph_format.first_line_indent = Inches(0)
    r_num = p_num.add_run(f"Tabla {num}")
    r_num.font.name = FONT_NAME
    r_num.font.size = FONT_SIZE
    r_num.font.bold = True
    r_num.font.color.rgb = BLACK

    p_tit = doc.add_paragraph()
    p_tit.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_tit.paragraph_format.space_before = Pt(0)
    p_tit.paragraph_format.space_after = Pt(6)
    p_tit.paragraph_format.first_line_indent = Inches(0)
    r_tit = p_tit.add_run(title)
    r_tit.font.name = FONT_NAME
    r_tit.font.size = FONT_SIZE
    r_tit.font.italic = True
    r_tit.font.color.rgb = BLACK

    table = doc.add_table(rows=len(rows) + 1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_full_black_table_borders(table)

    for j, h in enumerate(headers):
        cell = table.rows[0].cells[j]
        set_cell_margins(cell, top=70, bottom=70, left=90, right=90)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.15
        r = p.add_run(limpiar_guiones(str(h)))
        r.font.name = FONT_NAME
        r.font.size = FONT_SIZE
        r.font.bold = True
        r.font.color.rgb = BLACK

    for i, row_data in enumerate(rows, start=1):
        for j, val in enumerate(row_data):
            cell = table.rows[i].cells[j]
            set_cell_margins(cell, top=60, bottom=60, left=90, right=90)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.15
            r = p.add_run(limpiar_guiones(str(val)))
            r.font.name = FONT_NAME
            r.font.size = FONT_SIZE
            r.font.color.rgb = BLACK

    p_note = doc.add_paragraph()
    p_note.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_note.paragraph_format.space_before = Pt(4)
    p_note.paragraph_format.space_after = Pt(10)
    p_note.paragraph_format.first_line_indent = Inches(0)
    p_note.paragraph_format.line_spacing = 1.5
    rn1 = p_note.add_run("Nota. ")
    rn1.font.name = FONT_NAME
    rn1.font.size = FONT_SIZE
    rn1.font.italic = True
    rn1.font.color.rgb = BLACK
    rn2 = p_note.add_run(note_text)
    rn2.font.name = FONT_NAME
    rn2.font.size = FONT_SIZE
    rn2.font.color.rgb = BLACK


def add_apa_figure(doc, num, title, img_path, note_text, width_inches=6.2):
    title = limpiar_guiones(title)
    note_text = limpiar_guiones(note_text)

    p_num = doc.add_paragraph()
    p_num.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_num.paragraph_format.space_before = Pt(10)
    p_num.paragraph_format.space_after = Pt(2)
    p_num.paragraph_format.first_line_indent = Inches(0)
    r_num = p_num.add_run(f"Figura {num}")
    r_num.font.name = FONT_NAME
    r_num.font.size = FONT_SIZE
    r_num.font.bold = True
    r_num.font.color.rgb = BLACK

    p_tit = doc.add_paragraph()
    p_tit.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_tit.paragraph_format.space_before = Pt(0)
    p_tit.paragraph_format.space_after = Pt(6)
    p_tit.paragraph_format.first_line_indent = Inches(0)
    r_tit = p_tit.add_run(title)
    r_tit.font.name = FONT_NAME
    r_tit.font.size = FONT_SIZE
    r_tit.font.italic = True
    r_tit.font.color.rgb = BLACK

    if os.path.exists(img_path):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.first_line_indent = Inches(0)
        run_img = p_img.add_run()
        run_img.add_picture(img_path, width=Inches(width_inches))

    p_note = doc.add_paragraph()
    p_note.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_note.paragraph_format.space_before = Pt(4)
    p_note.paragraph_format.space_after = Pt(10)
    p_note.paragraph_format.first_line_indent = Inches(0)
    p_note.paragraph_format.line_spacing = 1.5
    rn1 = p_note.add_run("Nota. ")
    rn1.font.name = FONT_NAME
    rn1.font.size = FONT_SIZE
    rn1.font.italic = True
    rn1.font.color.rgb = BLACK
    rn2 = p_note.add_run(note_text)
    rn2.font.name = FONT_NAME
    rn2.font.size = FONT_SIZE
    rn2.font.color.rgb = BLACK


def add_code_block(doc, header_label, code_text):
    header_label = limpiar_guiones(header_label)
    code_text = limpiar_guiones(code_text)

    tbl = doc.add_table(rows=2, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_full_black_table_borders(tbl)

    c0 = tbl.rows[0].cells[0]
    set_cell_margins(c0, top=60, bottom=60, left=110, right=110)
    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p0.paragraph_format.space_before = Pt(0)
    p0.paragraph_format.space_after = Pt(0)
    r0 = p0.add_run(header_label)
    r0.font.name = FONT_NAME
    r0.font.size = FONT_SIZE
    r0.font.bold = True
    r0.font.color.rgb = BLACK

    c1 = tbl.rows[1].cells[0]
    set_cell_margins(c1, top=80, bottom=80, left=110, right=110)
    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p1.paragraph_format.space_before = Pt(0)
    p1.paragraph_format.space_after = Pt(0)
    p1.paragraph_format.line_spacing = 1.15
    r1 = p1.add_run(code_text.strip())
    r1.font.name = FONT_NAME
    r1.font.size = FONT_SIZE
    r1.font.color.rgb = BLACK

    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_before = Pt(0)
    p_sp.paragraph_format.space_after = Pt(4)


def add_reference_apa(doc, ref_text):
    ref_text = limpiar_guiones(ref_text)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.left_indent = Inches(0.5)
    p.paragraph_format.first_line_indent = Inches(-0.5)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(ref_text)
    r.font.name = FONT_NAME
    r.font.size = FONT_SIZE
    r.font.color.rgb = BLACK


def construir_documento():
    doc = Document()

    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        add_page_number_to_header(section)

    style_normal = doc.styles["Normal"]
    style_normal.font.name = FONT_NAME
    style_normal.font.size = FONT_SIZE
    style_normal.font.color.rgb = BLACK

    # ══════════════════════════════════════════════════════════════════════════
    # PÁGINA 1: PORTADA ESTUDIANTIL APA 7.ª EDICIÓN
    # ══════════════════════════════════════════════════════════════════════════
    for _ in range(3):
        doc.add_paragraph()

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.line_spacing = 1.5
    r_t = p_title.add_run(
        "Sistema Integrado de Análisis y Calidad de Datos para el Ecosistema Educativo "
        "Distrital de Bogotá D.C.\nManual Técnico, Fundamentación Conceptual y Operación (Versión Final 2.0)"
    )
    r_t.font.name = FONT_NAME
    r_t.font.size = FONT_SIZE
    r_t.font.bold = True
    r_t.font.color.rgb = BLACK

    for _ in range(3):
        doc.add_paragraph()

    datos_portada = [
        "David Santiago Castelblanco Artunduaga (Código: 67001182)",
        "Dairo José Narváez Morales (Código: 67002032)",
        "",
        "Facultad de Ingeniería, Programa de Ingeniería de Sistemas",
        "Universidad Católica de Colombia",
        "Ingeniería de Datos",
        "Ing. Diana Marcela Toquica Rodríguez",
        "22 de noviembre de 2026",
    ]
    for linea in datos_portada:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(linea)
        r.font.name = FONT_NAME
        r.font.size = FONT_SIZE
        r.font.color.rgb = BLACK

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════════
    # PÁGINA 2: ENLACE DE SUSTENTACIÓN EN YOUTUBE (OBLIGATORIO) + RESUMEN EJECUTIVO
    # ══════════════════════════════════════════════════════════════════════════
    add_apa_heading(doc, "Enlace del Video de Sustentación y Demostración Funcional", level=1)

    add_apa_paragraph(
        doc,
        "En cumplimiento de los lineamientos establecidos para la entrega final de la asignatura "
        "Ingeniería de Datos (Tercer Corte), a continuación se registra el enlace público en la "
        "plataforma YouTube correspondiente a la sustentación integral del proyecto, en la cual "
        "ambos integrantes participan en cámara con presentación de apoyo y realizan la demostración "
        "funcional en vivo del pipeline por capas (Lakehouse Batch) y del motor de procesamiento de "
        "eventos en tiempo real (Streaming, Watermarks y Ventanas de Tiempo):",
        indent=True,
    )

    tbl_yt = doc.add_table(rows=5, cols=2)
    tbl_yt.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_full_black_table_borders(tbl_yt)
    yt_rows = [
        ("Enlace Público de YouTube (Video de Sustentación):", "https://youtu.be/[PEGAR_AQUI_ENLACE_PUBLICO_YOUTUBE]"),
        ("Visibilidad y Duración:", "Pública (Acceso abierto sin restricción) | Duración: 8 minutos (Rango 5 a 10 min)"),
        ("Integrantes en Cámara:", "David Santiago Castelblanco Artunduaga (67001182) y Dairo José Narváez Morales (67002032)"),
        ("Repositorio Oficial GitHub:", "https://github.com/HCA-3/Proyecto-ING-DATO.git"),
        ("Contenido Articulado:", "Fundamentación teórica (Cortes 1, 2 y 3), arquitectura Lakehouse + Streaming, aportes individuales y demostración funcional en terminal"),
    ]
    for idx, (k, v) in enumerate(yt_rows):
        c_k = tbl_yt.rows[idx].cells[0]
        c_v = tbl_yt.rows[idx].cells[1]
        set_cell_margins(c_k, top=70, bottom=70, left=100, right=100)
        set_cell_margins(c_v, top=70, bottom=70, left=100, right=100)
        pk = c_k.paragraphs[0]
        pk.alignment = WD_ALIGN_PARAGRAPH.LEFT
        rk = pk.add_run(limpiar_guiones(k))
        rk.font.name = FONT_NAME
        rk.font.size = FONT_SIZE
        rk.font.bold = True
        rk.font.color.rgb = BLACK
        pv = c_v.paragraphs[0]
        pv.alignment = WD_ALIGN_PARAGRAPH.LEFT
        rv = pv.add_run(limpiar_guiones(v))
        rv.font.name = FONT_NAME
        rv.font.size = FONT_SIZE
        rv.font.color.rgb = BLACK
        if idx == 0:
            rv.font.bold = True

    add_apa_heading(doc, "Resumen Ejecutivo", level=1)
    add_apa_paragraph(
        doc,
        "El presente manual técnico e informe integrado aborda la problemática de dispersión, "
        "heterogeneidad e inconsistencia de calidad en los datos del ecosistema educativo de "
        "Bogotá D.C., articulando cifras oficiales sobre pobreza multidimensional, rezago y "
        "deserción escolar con una solución de ingeniería de datos completamente funcional. Se "
        "consolida la fundamentación teórica abordada a lo largo de los tres cortes de la "
        "asignatura Ingeniería de Datos, abarcando desde arquitecturas Data Warehouse, Data Lake y Data "
        "Lakehouse (arquitectura Medallón Bronze, Silver y Gold), modelado dimensional en estrella, "
        "gobierno de calidad DAMA-DMBOK y marco normativo colombiano (Leyes 1712 de 2014 y 1581 de "
        "2012, CONPES 3920), hasta los componentes de ingeniería implementados en el Tercer Corte: "
        "orquestación de pipelines ETL/ELT, Arquitectura Orientada a Eventos (EDA) para alertas "
        "escolares, procesamiento por Event Time con marcas de agua (Watermarks), agrupación por "
        "ventanas de tiempo (Tumbling, Sliding y Session Windows) con PySpark y manejo de reportes "
        "tardíos mediante holgura (Allowed Lateness) y cola de conciliación (DLQ).",
        indent=True,
    )
    add_apa_paragraph(
        doc,
        "A nivel de implementación técnica (versión 2.0.0), se documenta con fragmentos de código "
        "comentados y evidencias reales el funcionamiento de tres módulos principales en Python y "
        "PySpark: (a) generar_datos_raw.py, que sintetiza de forma reproducible 87,520 registros "
        "distribuidos en las 20 localidades de Bogotá con anomalías realistas de calidad; "
        "(b) pipeline_educativo.py, que orquesta la resiliencia híbrida entre PySpark y Pandas, "
        "ejecuta el contrato de limpieza en la capa Silver (eliminación de 1,070 valores centinela "
        "e imputación por moda de 2,934 códigos DANE faltantes) y construye el datamart analítico "
        "Gold con alertas tempranas de deserción (> 8.0%); y (c) streaming_alertas_educativas.py, "
        "que procesa flujos de eventos CloudEvents 1.0 en tiempo real aplicando idempotencia, "
        "Watermarks de 5 minutos, ventanas Tumbling, Sliding y Session, y desvío de eventos "
        "extremos a una cola DLQ para evitar la pérdida de alertas estudiantiles.",
        indent=True,
    )

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════════
    # 1. INTRODUCCIÓN
    # ══════════════════════════════════════════════════════════════════════════
    add_apa_heading(doc, "1. Introducción", level=1)
    add_apa_paragraph(
        doc,
        "Este documento corresponde a la versión final 2.0 del proyecto integrador Sistema "
        "Integrado de Análisis y Calidad de Datos para el Ecosistema Educativo Distrital de "
        "Bogotá D.C. Concebido desde el inicio del semestre como un documento vivo y evolutivo, "
        "este manual técnico articula en una sola obra la justificación socioeducativa, el marco "
        "normativo y ético, los conceptos teóricos de los tres cortes de la asignatura, el manual "
        "de programación con código fuente real y evidencias verificables, la identificación "
        "argumentada del aporte individual de cada autor y el manual de usuario operativo.",
        indent=True,
    )
    add_apa_paragraph(
        doc,
        "El proyecto parte de una paradoja crítica en la gestión pública distrital: Bogotá D.C. "
        "dispone de un volumen significativo de datos abiertos relacionados con la educación, "
        "tales como resultados de las pruebas estandarizadas Saber 11, directorios institucionales de "
        "colegios oficiales y privados, registros de beneficiarios de programas de acceso a la "
        "educación superior como Jóvenes a la E y La U en tu Colegio, y reportes zonales de "
        "deserción y repitencia, pero dichos activos se encuentran fragmentados en silos "
        "independientes y presentan deficiencias severas de calidad (valores centinela en "
        "campos numéricos, ausencia de llaves foráneas DANE y rezagos en la transmisión de alertas).",
        indent=True,
    )
    add_apa_paragraph(
        doc,
        "Para resolver esta problemática de extremo a extremo, se diseñó e implementó una "
        "arquitectura híbrida que combina un Data Lakehouse por capas (Bronze/RAW, Silver/STAGING "
        "y Gold/ANALYTICS) para el procesamiento histórico por lotes (Batch) con un módulo de "
        "procesamiento de eventos en tiempo real (Streaming) fundamentado en los temas del Tercer "
        "Corte, permitiendo tanto el análisis estructural de equidad educativa como la detección "
        "inmediata de ráfagas de inasistencia y riesgo de abandono escolar en las 20 localidades.",
        indent=True,
    )

    # ══════════════════════════════════════════════════════════════════════════
    # 2. PLANTEAMIENTO DEL PROBLEMA Y JUSTIFICACIÓN
    # ══════════════════════════════════════════════════════════════════════════
    add_apa_heading(doc, "2. Planteamiento del Problema y Justificación", level=1)
    add_apa_heading(doc, "2.1 Problemática Identificada", level=2)
    add_apa_paragraph(
        doc,
        "Aunque Bogotá D.C. registra uno de los índices de pobreza multidimensional más bajos a "
        "nivel nacional, las brechas intraurbanas siguen afectando profundamente la trayectoria "
        "escolar de miles de niños, niñas y jóvenes. De acuerdo con el Departamento Administrativo "
        "Nacional de Estadística (DANE, 2025), el índice de pobreza multidimensional en Bogotá "
        "pasó del 3.6% en 2023 al 5.4% en 2024, afectando a cerca de 431,000 personas. Dentro de "
        "este indicador compuesto, la dimensión de rezago escolar alcanzó el 17.5% en 2024, "
        "posicionándose como uno de los factores estructurales de mayor peso en la vulnerabilidad "
        "de los hogares capitalinos.",
        indent=True,
    )
    add_apa_paragraph(
        doc,
        "Adicionalmente, estas condiciones adversas se distribuyen de manera altamente asimétrica "
        "en el territorio. Localidades del sur y suroccidente como Usme, Ciudad Bolívar y Bosa "
        "registran niveles de pobreza monetaria e inasistencia escolar sustancialmente superiores "
        "a los observados en localidades como Teusaquillo, Chapinero y Usaquén (SaluData, 2025). "
        "En coherencia con ello, las estrategias de búsqueda activa de población desescolarizada "
        "implementadas por la Alcaldía Mayor de Bogotá (2024) priorizaron localidades críticas "
        "como Bosa, Suba, Kennedy, Ciudad Bolívar, Engativá, Rafael Uribe Uribe, San Cristóbal, "
        "Usme, Fontibón y Tunjuelito.",
        indent=True,
    )
    add_apa_paragraph(
        doc,
        "Desde la perspectiva de la Ingeniería de Datos, el obstáculo central no radica en la "
        "ausencia de registros, sino en cuatro patologías técnicas verificadas en las fuentes: "
        "(a) heterogeneidad de esquemas y formatos entre conjuntos de datos; (b) presencia de "
        "cadenas centinela ('-', 'N/A', 'ND') incrustadas en columnas de puntajes académicos que "
        "corrompen la inferencia de tipos numéricos; (c) pérdida del identificador institucional "
        "COLE_COD_DANE_INST en el 4.2% de los estudiantes evaluados, lo que rompe la integridad "
        "referencial al intentar cruzar resultados con el directorio de colegios; y (d) latencia "
        "en el reporte de alertas de deserción desde colegios con conectividad intermitente.",
        indent=True,
    )

    add_apa_heading(doc, "2.2 Justificación", level=2)
    add_apa_paragraph(
        doc,
        "La construcción de una plataforma de ingeniería de datos gobernada que unifique en una "
        "capa Gold (datamart dimensional) el desempeño en Saber 11, la caracterización de "
        "instituciones educativas, la absorción en programas de educación superior y el flujo "
        "en tiempo real de alertas de deserción permite superar el paradigma tradicional de "
        "reportes estáticos y aislados. Este enfoque se alinea directamente con la Política "
        "Nacional de Explotación de Datos, Big Data (Departamento Nacional de Planeación [DNP], "
        "2018, Documento CONPES 3920), la cual establece que los datos de las entidades públicas "
        "constituyen activos estratégicos para formular políticas basadas en evidencia y optimizar "
        "la focalización de recursos estatales.",
        indent=True,
    )

    # ══════════════════════════════════════════════════════════════════════════
    # 3. OBJETIVOS
    # ══════════════════════════════════════════════════════════════════════════
    add_apa_heading(doc, "3. Objetivos", level=1)
    add_apa_heading(doc, "3.1 Objetivo General", level=2)
    add_apa_paragraph(
        doc,
        "Diseñar, implementar y validar una arquitectura híbrida de ingeniería de datos, integrada "
        "por un Data Lakehouse multicapa (Bronze, Silver y Gold) y un motor de procesamiento de "
        "eventos en tiempo real con ventanas de tiempo y marcas de agua, que permita ingerir, "
        "depurar, cruzar y analizar información del ecosistema educativo de Bogotá D.C. para "
        "apoyar la toma de decisiones sobre equidad académica y prevención de la deserción escolar.",
        indent=True,
    )

    add_apa_heading(doc, "3.2 Objetivos Específicos", level=2)
    add_apa_bullet(
        doc,
        "Ingesta inmutable (Capa Bronze / RAW): ",
        "Construir un mecanismo reproducible de generación e ingesta de conjuntos de datos "
        "tabulares (.csv) y flujos de eventos (.jsonl en formato CloudEvents 1.0) que preserve "
        "intactos los datos crudos de las 20 localidades de Bogotá D.C.",
    )
    add_apa_bullet(
        doc,
        "Contrato de calidad y limpieza (Capa Silver / STAGING): ",
        "Implementar algoritmos de perfilamiento, filtrado de valores centinela en tiempo O(1) "
        "e imputación estadística por moda para sanear puntajes Saber 11 y recuperar 2,934 "
        "registros con códigos DANE faltantes sin alterar la distribución poblacional.",
    )
    add_apa_bullet(
        doc,
        "Modelado dimensional y reglas de negocio (Capa Gold / ANALYTICS): ",
        "Consolidar un datamart analítico mediante cruces relacionales (LEFT JOIN) y agregaciones "
        "distribuidas (PySpark/Pandas) que evalúe automáticamente la regla crítica de deserción "
        "escolar (> 8.0%) y la tasa de graduación en educación superior.",
    )
    add_apa_bullet(
        doc,
        "Procesamiento de eventos en tiempo real (Corte 3, Streaming): ",
        "Desarrollar un módulo funcional orientado a eventos que implemente deduplicación "
        "idempotente, diferenciación entre Event Time y Processing Time, cálculo de Watermarks, "
        "agrupación mediante ventanas Tumbling, Sliding y Session, y desvío de eventos tardíos "
        "extremos hacia una cola de conciliación (Side Output / DLQ).",
    )
    add_apa_bullet(
        doc,
        "Documentación técnica y transferencia operativa: ",
        "Documentar bajo normas APA (7.ª edición) los conceptos de la asignatura, la explicación "
        "detallada del código con sus evidencias, los aportes individuales de cada autor y los "
        "procedimientos de operación para usuarios finales.",
    )

    # ══════════════════════════════════════════════════════════════════════════
    # 4. VALOR ESPERADO Y CONTEXTO SOCIOEDUCATIVO DE BOGOTÁ D.C.
    # ══════════════════════════════════════════════════════════════════════════
    add_apa_heading(doc, "4. Valor Esperado y Contexto Socioeducativo de Bogotá D.C.", level=1)
    add_apa_paragraph(
        doc,
        "El contexto territorial de Bogotá D.C. fundamenta las preguntas analíticas que resuelve "
        "la arquitectura desarrollada. Los resultados de la Encuesta Multipropósito evidencian "
        "contrastes socioeconómicos pronunciados entre localidades que inciden directamente en la "
        "permanencia escolar y en el rendimiento académico en la educación media. Como se observa "
        "en la Tabla 1, la incidencia de la pobreza monetaria en localidades del sur supera en más "
        "de siete veces la registrada en localidades del corredor centro-norte.",
        indent=True,
    )

    add_apa_table(
        doc,
        num=1,
        title="Pobreza monetaria por localidad en Bogotá D.C., comparación de valores extremos",
        headers=["Localidad", "Código Distrital", "Pobreza Monetaria (%)", "Clasificación de Incidencia", "Prioridad de Intervención SED"],
        rows=[
            ["Usme", "05", "57.81%", "Alta incidencia", "Prioridad máxima (Búsqueda activa y apoyo)"],
            ["Ciudad Bolívar", "19", "57.37%", "Alta incidencia", "Prioridad máxima (Búsqueda activa y apoyo)"],
            ["Bosa", "07", "53.18%", "Alta incidencia", "Prioridad máxima (Búsqueda activa y apoyo)"],
            ["Usaquén", "01", "15.96%", "Baja incidencia", "Monitoreo focalizado por sectores"],
            ["Chapinero", "02", "12.83%", "Baja incidencia", "Monitoreo preventivo"],
            ["Teusaquillo", "13", "7.90%", "Baja incidencia", "Monitoreo preventivo"],
        ],
        note_text="Adaptado de los indicadores socioeconómicos de la Encuesta Multipropósito y reportes territoriales de SaluData (2025) y DANE (2025).",
    )

    add_apa_paragraph(
        doc,
        "A partir de esta realidad territorial, la solución técnica aporta valor directo a tres "
        "perfiles institucionales: (a) a los equipos de planeación de la Secretaría de Educación "
        "del Distrito (SED), al automatizar la detección de localidades cuya tasa de deserción "
        "supera el umbral del 8.0%; (b) a la Agencia Distrital para la Educación Superior, la "
        "Ciencia y la Tecnología (Atenea), al cruzar los beneficiarios activos de Jóvenes a la E "
        "y La U en tu Colegio frente a los promedios globales de Saber 11; y (c) a los comités de "
        "convivencia y orientación escolar, al procesar en tiempo real las ráfagas de inasistencia "
        "sin perder reportes provenientes de zonas con conectividad inestable.",
        indent=True,
    )

    # ══════════════════════════════════════════════════════════════════════════
    # 5. MARCO TEÓRICO Y CONCEPTUAL DE LA ASIGNATURA Y SU APLICACIÓN
    # ══════════════════════════════════════════════════════════════════════════
    add_apa_heading(doc, "5. Marco Teórico y Conceptual de la Asignatura y su Aplicación en el Proyecto", level=1)
    add_apa_paragraph(
        doc,
        "En esta sección se documentan de manera sistemática los conceptos y tecnologías "
        "abordados durante los tres cortes de la asignatura Ingeniería de Datos y se explica "
        "cómo cada uno de ellos fue aplicado técnicamente dentro del proyecto.",
        indent=True,
    )

    add_apa_heading(doc, "5.1 Sistemas OLTP vs. OLAP y Evolución hacia el Data Lakehouse (Arquitectura Medallón)", level=2)
    add_apa_paragraph(
        doc,
        "En la ingeniería de datos moderna se distingue entre los sistemas transaccionales en "
        "línea (OLTP, Online Transaction Processing), orientados a registrar operaciones "
        "atómicas del día a día (como el registro individual de matrícula de un estudiante en el "
        "sistema SIMAT), y los sistemas analíticos en línea (OLAP, Online Analytical Processing), "
        "diseñados para ejecutar consultas agregadas complejas sobre grandes volúmenes históricos "
        "(Kleppmann, 2017). Históricamente, las organizaciones consolidaron sus sistemas OLAP "
        "mediante almacenes de datos (Data Warehouses) con esquemas rígidos de escritura "
        "(schema-on-write), y posteriormente mediante lagos de datos (Data Lakes) capaces de "
        "almacenar archivos crudos en cualquier formato (schema-on-read), pero vulnerables a "
        "degradarse en pantanos de datos (data swamps) por falta de gobierno y calidad.",
        indent=True,
    )
    add_apa_paragraph(
        doc,
        "El paradigma Data Lakehouse fusiona la flexibilidad y bajo costo de almacenamiento de un "
        "Data Lake con las garantías de calidad, gobierno y estructura analítica de un Data "
        "Warehouse (Databricks, 2024). Dentro de este enfoque, nuestro proyecto implementa la "
        "Arquitectura Medallón (Medallion Architecture) estructurada en tres capas físicas: "
        "(a) Capa Bronze (data/RAW/), donde se depositan los archivos CSV y JSONL originales de "
        "manera estrictamente inmutable; (b) Capa Silver (data/STAGING/), donde se aplican reglas "
        "de limpieza, tipado, filtrado de activos e imputación estadística; y (c) Capa Gold "
        "(data/ANALYTICS/), donde se disponen las tablas dimensionales y agregaciones listas para "
        "el consumo en inteligencia de negocios.",
        indent=True,
    )

    add_apa_heading(doc, "5.2 Modelado Dimensional (Esquema en Estrella) y Pipelines ETL vs. ELT", level=2)
    add_apa_paragraph(
        doc,
        "Para estructurar la capa Gold se adoptó el modelado dimensional en Esquema de Estrella "
        "(Star Schema), el cual organiza la información en una tabla de hechos central "
        "(Fact_Rendimiento_Saber11) con las métricas cuantitativas de desempeño (PUNT_GLOBAL y "
        "puntajes por área), rodeada por tablas de dimensiones desnormalizadas que aportan el "
        "contexto descriptivo: Dim_Colegio (código DANE, nombre, sector, jornada, estado), "
        "Dim_Localidad (código, nombre de localidad, tasa de deserción y nivel de alerta), "
        "Dim_Programa_Social (Jóvenes a la E, La U en tu Colegio, IES admitida y estado del "
        "beneficio) y Dim_Tiempo (periodo académico y marcas temporales).",
        indent=True,
    )
    add_apa_paragraph(
        doc,
        "En el Tercer Corte se profundizó en la construcción de pipelines de datos y en la "
        "diferenciación entre los enfoques ETL (Extract, Transform, Load) y ELT (Extract, Load, "
        "Transform). Mientras que en ETL la transformación ocurre antes de almacenar los datos en "
        "el repositorio central, en ELT los datos extraídos se cargan primero de forma cruda e "
        "inmutable en el almacenamiento (Capa Bronze) y posteriormente se transforman en el motor "
        "de procesamiento hacia las capas Silver y Gold. Nuestro proyecto adopta un diseño "
        "híbrido ELT/ETL: preserva la carga inmutable en data/RAW/ (filosofía ELT de linaje y "
        "reprocesabilidad) y orquesta transformaciones modulares deterministas en memoria con "
        "PySpark y Pandas hacia data/STAGING/ y data/ANALYTICS/.",
        indent=True,
    )

    add_apa_heading(doc, "5.3 Gestión de Calidad de Datos (DAMA-DMBOK) y Procesamiento Distribuido (Apache Spark)", level=2)
    add_apa_paragraph(
        doc,
        "Siguiendo el marco internacional DAMA-DMBOK (DAMA International, 2017), la calidad de "
        "los datos se operacionalizó en el proyecto mediante cuatro dimensiones verificables "
        "programáticamente en el paso de Bronze a Silver: (a) Validez y Conformidad, verificando "
        "que todo puntaje académico sea numérico y pertenezca al intervalo cerrado [0.0, 500.0] "
        "tras descartar códigos centinela; (b) Completitud, remediando el 4.2% de registros "
        "huérfanos en COLE_COD_DANE_INST mediante imputación por moda institucional; "
        "(c) Consistencia e Integridad Referencial, asegurando que el cruce relacional LEFT JOIN "
        "preserve la totalidad de los 70,000 estudiantes evaluados; y (d) Actualidad y Pertinencia, "
        "filtrando exclusivamente los 2,262 establecimientos educativos con estado ACTIVO.",
        indent=True,
    )
    add_apa_paragraph(
        doc,
        "En cuanto al motor de cómputo, se aplicaron los principios de procesamiento distribuido "
        "en memoria de Apache Spark (PySpark SQL y DataFrames), diferenciando entre "
        "transformaciones perezosas (lazy transformations como filter, select, groupBy, join y "
        "withColumn) y acciones (count, collect, show, write) que desencadenan la ejecución del "
        "grafo acíclico dirigido (DAG) (Apache Software Foundation, 2026). Adicionalmente, para "
        "garantizar la confiabilidad y mantenibilidad en estaciones de trabajo con Java 8 "
        "(incompatible con el bytecode Java 17+ exigido por Spark 4.x), se implementó un patrón "
        "de resiliencia híbrida que detecta la máquina virtual en tiempo de arranque y conmuta "
        "transparentemente al motor vectorizado de Pandas/NumPy manteniendo idéntica semántica.",
        indent=True,
    )

    add_apa_heading(doc, "5.4 Arquitectura Orientada a Eventos (EDA), CloudEvents e Idempotencia (Corte 3)", level=2)
    add_apa_paragraph(
        doc,
        "Como se estudió en el Tercer Corte, una Arquitectura Orientada a Eventos (Event-Driven "
        "Architecture, EDA) reemplaza el acoplamiento síncrono de petición/respuesta (REST) por "
        "una comunicación asíncrona basada en la producción, enrutamiento y reacción ante eventos "
        "(hechos inmutables que ya ocurrieron y cambian el estado del dominio, nombrados en tiempo "
        "pasado o como sustantivos de notificación, p. ej., alerta.riesgo_desercion). La "
        "arquitectura se compone de cuatro elementos esenciales: (a) Productor, que detecta el "
        "hecho en origen y publica el evento sin saber quién lo consumirá; (b) Broker o "
        "intermediario (como Apache Kafka), que almacena los eventos en un log ordenado e "
        "inmutable; (c) Tópico, canal lógico particionado que agrupa eventos de la misma "
        "categoría; y (d) Consumidor, que se suscribe al tópico y procesa cada registro.",
        indent=True,
    )
    add_apa_paragraph(
        doc,
        "En el módulo scripts/streaming_alertas_educativas.py aplicamos este paradigma adoptando "
        "la especificación estándar CloudEvents 1.0 de la Cloud Native Computing Foundation "
        "(CNCF, 2024), estructurando cada alerta escolar con metadatos de cabeceras (specversion, id, "
        "source, type, event_time, processing_time, partition_key) y carga útil (data). Asimismo, "
        "para mitigar el riesgo clásico de duplicidad de eventos ocasionado por reintentos de red "
        "(at-least-once delivery), implementamos un consumidor idempotente en la clase "
        "BrokerEventosEducativos que registra en un conjunto hash O(1) los identificadores ya "
        "procesados (ids_procesados), descartando automáticamente el reintento duplicado del "
        "evento EVT-004 en la localidad de Bosa sin alterar los conteos.",
        indent=True,
    )

    add_apa_heading(doc, "5.5 Event Time, Processing Time y Marcas de Agua (Watermarks) (Corte 3)", level=2)
    add_apa_paragraph(
        doc,
        "Un pilar fundamental del procesamiento de flujos abordado en el Tercer Corte es la "
        "distinción entre dos referencias temporales: el Tiempo del Evento (Event Time), que "
        "corresponde al instante exacto en que el hecho ocurrió en el mundo real y viaja "
        "incrustado dentro del dato, y el Tiempo de Procesamiento (Processing Time), que "
        "corresponde al reloj del servidor en el momento en que el motor recibe y procesa el "
        "registro (Akidau et al., 2015). En escenarios reales, ambos relojes divergen debido a "
        "latencia de red, desconexiones temporales en zonas periféricas o rurales, colas con "
        "contrapresión (backpressure) y reintentos de transmisión.",
        indent=True,
    )
    add_apa_paragraph(
        doc,
        "Si un sistema educativo agrupa las alertas usando Processing Time, una caída temporal de "
        "conectividad en colegios de Usme o Sumapaz genera un falso silencio seguido de un pico "
        "artificial cuando retorna el enlace, distorsionando las métricas y produciendo resultados "
        "no reproducibles. Por ello, nuestro motor procesa por Event Time e implementa una Marca "
        "de Agua (Watermark) definida formalmente como WM = max(event_time visto) - T_umbral. El "
        "Watermark actúa como un reloj lógico monótonamente creciente que indica hasta qué "
        "instante de tiempo de evento el sistema asume que no llegarán más registros atrasados, "
        "permitiendo cerrar ventanas finitas y equilibrar el compromiso (trade-off) entre "
        "completitud de datos y baja latencia.",
        indent=True,
    )

    add_apa_heading(doc, "5.6 Ventanas de Tiempo (Tumbling, Sliding y Session) y Eventos Tardíos (Corte 3)", level=2)
    add_apa_paragraph(
        doc,
        "Dado que un flujo de eventos (stream) es potencialmente infinito, para calcular "
        "agregaciones (conteos, sumas o promedios) es indispensable segmentar el tiempo en "
        "intervalos semiabiertos [inicio, fin) respaldados por un almacén de estado (State Store) "
        "(Akidau et al., 2018). En el proyecto se implementaron y compararon las tres tipologías "
        "de ventanas estudiadas en el curso:",
        indent=True,
    )
    add_apa_bullet(
        doc,
        "Ventanas Fijas (Tumbling Windows): ",
        "Dividen el tiempo en bloques de tamaño fijo (10 minutos en nuestro sistema, p. ej., "
        "[08:00, 08:10) y [08:10, 08:20)) sin ningún traslape. Cada evento pertenece exactamente "
        "a una sola ventana (factor en memoria = 1). Se aplican en el proyecto para emitir el "
        "reporte periódico consolidado de alertas y puntaje de riesgo por localidad.",
    )
    add_apa_bullet(
        doc,
        "Ventanas Deslizantes (Sliding / Hopping Windows): ",
        "Tienen una duración fija (10 minutos) pero avanzan con un paso o slide menor que su "
        "tamaño (cada 2 minutos). Como consecuencia, las ventanas se solapan y cada evento "
        "pertenece simultáneamente a Duración / Slide = 10 / 2 = 5 ventanas en el State Store. "
        "Se aplican en el proyecto para detectar ráfagas críticas de inasistencia (3 o más "
        "alertas en cualquier intervalo móvil de 10 minutos, incluso si ocurren a caballo entre "
        "las 08:08 y las 08:11 como sucedió en Bosa).",
    )
    add_apa_bullet(
        doc,
        "Ventanas de Sesión (Session Windows): ",
        "Poseen duración dinámica y sin tamaño predefinido: agrupan eventos consecutivos de una "
        "misma clave (código DANE del colegio) mientras el tiempo transcurrido entre ellos no "
        "supere una brecha de inactividad (session gap de 4 minutos). Si llega un evento antes de "
        "cumplirse el gap, la sesión se extiende; si transcurren más de 4 minutos sin actividad, "
        "la ventana se cierra. Se aplican en el proyecto para medir la duración y severidad de las "
        "sesiones activas de reporte de orientación escolar en cada institución educativa.",
    )
    add_apa_paragraph(
        doc,
        "Asimismo, cuando un reporte escolar llega fuera de orden después de que el Watermark ha "
        "superado el fin de su ventana (WM >= fin_ventana), nuestro pipeline implementa un "
        "tratamiento de eventos tardíos en tres niveles adaptado a las condiciones de conectividad "
        "de Bogotá D.C.: (a) aceptación directa en el State Store cuando el evento llega antes del "
        "cierre del Watermark (como el reporte EVT-009 de Usme); (b) holgura de actualización "
        "(Allowed Lateness de +5 minutos adicionales tras el cierre de la ventana), que permite "
        "reabrir el agregado y emitir una versión corregida cuando una sede sincroniza datos con "
        "retraso moderado (como el evento EVT-013 de Usme); y (c) desvío a una cola de conciliación "
        "(Dead Letter Queue, DLQ) para reportes extremadamente tardíos (como el evento EVT-017 del "
        "colegio rural de Sumapaz con 28 minutos de retraso), evitando el descarte silencioso (Drop) "
        "y garantizando que ninguna alerta de un estudiante vulnerable se pierda.",
        indent=True,
    )

    add_apa_heading(doc, "5.7 Síntesis de Conceptos de la Asignatura Implementados en el Proyecto", level=2)
    add_apa_paragraph(
        doc,
        "Para evidenciar de forma directa cómo se articulan los fundamentos de la asignatura "
        "con el software construido, la Tabla 2 relaciona cada concepto técnico seleccionado con "
        "el problema real del ecosistema educativo de Bogotá D.C. que resuelve y el módulo o "
        "función del repositorio donde fue programado.",
        indent=True,
    )

    add_apa_table(
        doc,
        num=2,
        title="Articulación de los conceptos de la asignatura con su implementación en el proyecto educativo",
        headers=["Concepto de la Asignatura", "Problema del Proyecto que Resuelve", "Implementación en Nuestro Sistema", "Archivo y Función en el Repositorio"],
        rows=[
            ["Data Lakehouse y Arquitectura Medallón (Bronze, Silver, Gold)", "Dispersión de fuentes abiertas de la SED y riesgo de alterar los archivos originales durante la limpieza.", "Separación física en tres capas: data/RAW/ (inmutable), data/STAGING/ (datos validados) y data/ANALYTICS/ (reportes finales).", "scripts/pipeline_educativo.py (RAW_DIR, STAGING_DIR, ANALYTICS_DIR)"],
            ["Modelado Dimensional (Esquema en Estrella)", "Falta de cruce entre puntajes Saber 11, características del colegio, programas de acceso superior y deserción por localidad.", "Tabla de hechos Saber 11 enriquecida mediante LEFT JOIN con el directorio de colegios activos y alertas por localidad.", "cruzar_saber11_con_directorio() y generar_reporte_equidad_educativa()"],
            ["Gestión de Calidad de Datos (Marco DAMA-DMBOK)", "Presencia de 1,070 puntajes con texto centinela ('-', 'N/A') y 2,934 estudiantes (4.2%) sin código DANE.", "Contrato de calidad en capa Silver: filtrado O(1) con conjunto hash VALORES_CENTINELA e imputación estadística por moda.", "validar_rango_puntaje(), limpiar_valores_nulos() y filtrar_colegios_activos()"],
            ["Procesamiento Distribuido (PySpark) y Resiliencia Híbrida", "Volumen de 87,520 registros y fallos de ejecución local cuando la máquina cuenta con Java 8 en lugar de Java 17+.", "Transformaciones distribuidas en PySpark con conmutación automática a Pandas vectorizado si se detecta incompatibilidad de JVM.", "comprobar_java_compatible() y crear_spark_session()"],
            ["Pipelines de Datos (Enfoque Híbrido ELT / ETL)", "Ejecución manual fragmentada y propensa a errores humanos al transformar múltiples conjuntos de datos.", "Orquestador automatizado de 6 etapas secuenciales desde la validación de archivos crudos hasta la persistencia analítica.", "ejecutar_pipeline() en scripts/pipeline_educativo.py"],
            ["Arquitectura Orientada a Eventos (EDA) e Idempotencia", "Reintentos de red desde colegios con internet inestable que duplican alertas de inasistencia escolar.", "Ingesta de eventos estandarizados y broker con registro hash O(1) de IDs procesados que descarta duplicados automáticamente.", "BrokerEventosEducativos y generar_flujo_eventos_educativos()"],
            ["Event Time vs. Processing Time y Watermarks", "Retrasos de transmisión en colegios periféricos o rurales que distorsionan el conteo si se usa la hora del servidor.", "Procesamiento gobernado por la hora real del reporte en el colegio (event_time) con marca de agua WM = max(ET) - 5 min.", "procesar_flujo_con_watermarks_y_ventanas() en streaming_alertas_educativas.py"],
            ["Ventanas de Tiempo (Tumbling, Sliding y Session) y Cola DLQ", "Necesidad de reportes periódicos, detección de ráfagas de deserción entre bloques y rescate de reportes rurales muy tardíos.", "Ventanas Tumbling (10m), Sliding (10m cada 2m para ráfagas >= 3 alertas), Session (gap 4m por colegio DANE), Allowed Lateness y DLQ.", "scripts/streaming_alertas_educativas.py (reportes en data/ANALYTICS/)"],
        ],
        note_text="Elaboración propia. Todos los conceptos listados se encuentran implementados y operativos en los scripts del proyecto.",
    )

    # ══════════════════════════════════════════════════════════════════════════
    # 6. MARCO NORMATIVO, ÉTICO Y MINIMIZACIÓN DE DATOS
    # ══════════════════════════════════════════════════════════════════════════
    add_apa_heading(doc, "6. Marco Normativo, Ético y Minimización de Datos", level=1)
    add_apa_paragraph(
        doc,
        "El procesamiento de información educativa en Colombia debe sujetarse a un marco jurídico "
        "y ético riguroso que armonice el derecho ciudadano de acceso a los datos públicos con la "
        "protección integral de los datos personales, especialmente tratándose de población "
        "estudiantil menor de edad. La Tabla 3 resume las disposiciones normativas que rigen "
        "la arquitectura del sistema.",
        indent=True,
    )

    add_apa_table(
        doc,
        num=3,
        title="Marco normativo colombiano aplicable al sistema de ingeniería de datos educativos",
        headers=["Norma / Documento", "Año", "Objeto Principal", "Aplicación y Cumplimiento en el Proyecto"],
        rows=[
            ["Ley 1712 de 2014", "2014", "Ley de Transparencia y del Derecho de Acceso a la Información Pública Nacional.", "Fundamenta el aprovechamiento de los conjuntos del portal Datos Abiertos Bogotá expedidos por la SED y el ICFES."],
            ["Ley 1581 de 2012", "2012", "Régimen General de Protección de Datos Personales (Habeas Data).", "Exige el principio de minimización y anonimización para impedir la reidentificación de estudiantes menores de edad."],
            ["CONPES 3920", "2018", "Política Nacional de Explotación de Datos (Big Data).", "Respalda la creación de arquitecturas interoperables y datamarts analíticos para focalizar la inversión pública distrital."],
        ],
        note_text="Elaboración propia con base en Congreso de la República de Colombia (2012, 2014) y DNP (2018).",
    )

    add_apa_paragraph(
        doc,
        "En cumplimiento estricto de la Ley 1581 de 2012, el sistema aplica dos controles técnicos "
        "de privacidad desde el diseño (Privacy by Design): (a) en la capa analítica Gold "
        "(data/ANALYTICS/), ningún reporte expone registros individuales de estudiantes, sino "
        "únicamente métricas agregadas a nivel de institución educativa, IES o localidad; y "
        "(b) en el flujo de eventos en tiempo real (data/RAW/eventos_asistencia_desercion_stream.jsonl), "
        "el código del estudiante es transformado inmediatamente mediante una función criptográfica "
        "hash SHA-256 truncada (función anonimizar_estudiante), generando pseudónimos irreversibles "
        "como EST-B8F9A21C4D.",
        indent=True,
    )

    # ══════════════════════════════════════════════════════════════════════════
    # 7. FUENTES DE DATOS Y ARQUITECTURA LÓGICA DEL SOFTWARE
    # ══════════════════════════════════════════════════════════════════════════
    add_apa_heading(doc, "7. Fuentes de Datos y Arquitectura Lógica del Software", level=1)
    add_apa_paragraph(
        doc,
        "El sistema procesa cinco activos de datos de entrada depositados en la capa Bronze "
        "(data/RAW/): cuatro conjuntos tabulares en formato CSV que modelan los reportes oficiales "
        "del portal Datos Abiertos Bogotá y un flujo continuo de eventos en formato JSON Lines "
        "(.jsonl) que modela las alertas tempranas emitidas en tiempo real por los colegios "
        "distritales. La Tabla 4 detalla su inventario físico y su rol dimensional.",
        indent=True,
    )

    add_apa_table(
        doc,
        num=4,
        title="Inventario de conjuntos de datos de entrada en la capa Bronze (data/RAW/)",
        headers=["Archivo en data/RAW/", "Modalidad", "Registros × Columnas", "Entidad Referente", "Rol en el Modelo Dimensional y Streaming"],
        rows=[
            ["Saber11_2025_Bogota.csv", "Batch (CSV)", "70,000 × 35 (18.88 MB)", "ICFES / SED", "Fuente central para Fact_Rendimiento_Saber11; contiene puntajes globales y por área."],
            ["Directorio_Colegios_Bogota.csv", "Batch (CSV)", "2,500 × 16 (0.49 MB)", "SED / IDECA", "Alimenta Dim_Colegio y Dim_Localidad; aporta código DANE, sector, jornada y estado."],
            ["Beneficiarios_Jovenes_a_la_E.csv", "Batch (CSV)", "15,000 × 10 (1.89 MB)", "SED / Atenea", "Alimenta Dim_Programa_Social; evalúa retención y tasa de graduación en 10 IES."],
            ["Reporte_Alertas_Localidad.csv", "Batch (CSV)", "20 × 8 (< 0.01 MB)", "SED", "Insumo de tasas de deserción y repitencia en las 20 localidades de Bogotá D.C."],
            ["eventos_asistencia_desercion_stream.jsonl", "Streaming (JSONL)", "18 eventos CloudEvents", "Colegios IED / SED", "Flujo continuo con event_time y processing_time para ventanas Tumbling, Sliding, Session y DLQ."],
        ],
        note_text="Elaboración propia a partir del perfilamiento ejecutado por generar_datos_raw.py y streaming_alertas_educativas.py.",
    )

    add_apa_paragraph(
        doc,
        "La Figura 1 ilustra la arquitectura híbrida implementada, mostrando cómo convergen el "
        "flujo por lotes de la Arquitectura Medallón (fila superior) y el flujo orientado a "
        "eventos en tiempo real del Tercer Corte (fila inferior) hacia la capa Gold de analítica "
        "y toma de decisiones.",
        indent=True,
    )

    add_apa_figure(
        doc,
        num=1,
        title="Arquitectura híbrida del sistema: Data Lakehouse (Medallón) y Streaming Orientado a Eventos",
        img_path=os.path.join(FIGS_DIR, "fig1_arquitectura_lakehouse_streaming.png"),
        note_text="Diagrama de arquitectura generado programáticamente por scripts/generar_graficos_evidencias.py. Integra las capas Bronze, Silver y Gold con el pipeline de streaming (CloudEvents, Broker idempotente, Watermarks, ventanas y DLQ).",
    )

    add_apa_paragraph(
        doc,
        "A nivel de sistema de archivos, el repositorio Proyecto-ING-DATO mantiene una estructura "
        "modular que separa estrictamente el código fuente (scripts/), los talleres de clase "
        "(programas/), las tres zonas de almacenamiento de datos (data/RAW/, data/STAGING/ y "
        "data/ANALYTICS/) y los paquetes de entrega final (entregables/entrega final/).",
        indent=True,
    )

    add_code_block(
        doc,
        "Estructura física del repositorio Proyecto-ING-DATO",
        """Proyecto-ING-DATO/
├── data/
│   ├── RAW/          (Capa Bronze inmutable: 4 archivos CSV + 1 stream JSONL CloudEvents)
│   ├── STAGING/      (Capa Silver depurada: Saber11_2025_Bogota_clean.csv, Directorio_Colegios_clean.csv)
│   └── ANALYTICS/    (Capa Gold: 5 reportes Batch CSV + 6 reportes Streaming CSV + figuras/)
├── scripts/
│   ├── generar_datos_raw.py             (Generador determinista de datasets RAW con semilla 42)
│   ├── pipeline_educativo.py            (Orquestador principal Etapas 1 a 6: PySpark / Pandas)
│   ├── streaming_alertas_educativas.py  (Motor EDA, Watermarks, Ventanas y DLQ, Corte 3)
│   └── generar_graficos_evidencias.py   (Generador automatizado de figuras técnicas APA 7)
├── entregables/
│   └── entrega final/                   (Manual Técnico APA 7, Presentación PPTX/HTML y Guion)
└── README.md                            (Documentación general del repositorio)""",
    )

    # ══════════════════════════════════════════════════════════════════════════
    # 8. DIAGNÓSTICO DE CALIDAD, LIMPIEZA Y CONTRATO DE DATOS
    # ══════════════════════════════════════════════════════════════════════════
    add_apa_heading(doc, "8. Diagnóstico de Calidad, Limpieza y Contrato de Datos (Capa Silver)", level=1)
    add_apa_paragraph(
        doc,
        "Durante la Etapa 1 del pipeline se ejecuta un perfilamiento automático en memoria sobre "
        "los archivos de la capa Bronze. Este diagnóstico puso al descubierto cuatro anomalías "
        "críticas que impedían el análisis directo y que fueron subsanadas mediante el contrato "
        "de datos en la Etapa 2 (capa Silver / STAGING), según se detalla en la Tabla 5 y se "
        "visualiza en la Figura 2.",
        indent=True,
    )

    add_apa_table(
        doc,
        num=5,
        title="Diagnóstico de anomalías de calidad en capa RAW y mitigación técnica en capa STAGING",
        headers=["Anomalía Detectada", "Volumen / Porcentaje Afectado", "Impacto Técnico en el Análisis", "Acción del Pipeline (Contrato Silver)"],
        rows=[
            ["Valores centinela en PUNT_GLOBAL ('-', 'N/A', 'ND', '')", "1,070 registros (1.53% de 70,000)", "Fuerza la inferencia de columna como texto (object) y provoca ValueError al calcular promedios.", "Filtrado en tiempo O(1) con conjunto hash VALORES_CENTINELA y conversión segura a float/NaN en validar_rango_puntaje()."],
            ["Falta de código DANE (COLE_COD_DANE_INST nulo)", "2,934 registros (4.20% de 70,000)", "Registros huérfanos que se perderían en un INNER JOIN con el directorio de colegios.", "Imputación estadística por moda institucional en limpiar_valores_nulos() + cruce LEFT JOIN (0 registros perdidos)."],
            ["Colegios con estado INACTIVO o EN PROCESO", "238 colegios (9.52% de 2,500)", "Distorsiona la oferta educativa vigente al incluir sedes cerradas o no operativas.", "Filtrado estricto ESTADO == 'ACTIVO' en filtrar_colegios_activos(), conservando 2,262 colegios activos."],
            ["Nulos en IES_ADMITIDA y FECHA_ADMISION", "1,693 (11.29%) y 2,227 (14.85%)", "Sesgo en el denominador al calcular la tasa de absorción y graduación universitaria.", "Exclusión focalizada (dropna(subset=['IES_ADMITIDA'])) únicamente en el cálculo de graduación por IES."],
            ["Eventos duplicados en flujo de streaming", "1 evento repetido (EVT-004 en Bosa)", "Sobredimensiona las alertas por localidad ante reintentos automáticos de red.", "Deduplicación idempotente por ID en BrokerEventosEducativos (conjunto ids_procesados)."],
        ],
        note_text="Elaboración propia a partir de las métricas reales arrojadas por la ejecución de pipeline_educativo.py.",
    )

    add_apa_figure(
        doc,
        num=2,
        title="Impacto cuantitativo del contrato de calidad de datos entre la capa Bronze y la capa Silver",
        img_path=os.path.join(FIGS_DIR, "fig2_calidad_datos_bronze_silver.png"),
        note_text="Comparación de registros con incidencias antes y después de la ejecución de la Etapa 2 del pipeline. Los 2,934 códigos DANE faltantes y los 1,070 valores centinela quedan completamente saneados.",
    )

    # ══════════════════════════════════════════════════════════════════════════
    # 9. MANUAL TÉCNICO DE PROGRAMACIÓN, CÓDIGO Y EVIDENCIAS
    # ══════════════════════════════════════════════════════════════════════════
    add_apa_heading(doc, "9. Manual Técnico de Programación, Explicación de Código y Evidencias", level=1)
    add_apa_paragraph(
        doc,
        "Este apartado documenta la arquitectura interna del código fuente para ingenieros de "
        "datos, docentes y desarrolladores encargados de auditar, mantener o escalar el sistema. "
        "Cada componente incluye la identificación exacta del archivo y función, el fragmento de "
        "código relevante, su explicación técnica paso a paso y las evidencias reales obtenidas "
        "tras su ejecución.",
        indent=True,
    )

    add_apa_heading(doc, "9.1 Módulo Generador de Datos Crudos (scripts/generar_datos_raw.py)", level=2)
    add_apa_paragraph(
        doc,
        "El script scripts/generar_datos_raw.py construye los cuatro archivos CSV de la capa "
        "Bronze fijando las semillas pseudoaleatorias (random.seed(42) y np.random.seed(42)) "
        "para garantizar absoluta reproducibilidad científica en cualquier máquina donde se "
        "clona el repositorio. Las funciones puntaje_raw() y get_dane() inyectan de forma "
        "controlada los valores centinela y la tasa del 4.2% de códigos DANE ausentes que "
        "simulan la suciedad real de las fuentes primarias.",
        indent=True,
    )

    add_code_block(
        doc,
        "Archivo: scripts/generar_datos_raw.py | Funciones: puntaje_raw() y get_dane()",
        """# Semilla fija para garantizar reproducibilidad exacta en auditorías
random.seed(42)
np.random.seed(42)

def puntaje_raw(p_null=0.042):
    \"\"\"Genera puntaje entre 0 y 500 con inyección controlada de valores centinela.\"\"\"
    if random.random() < p_null / 2:
        return random.choice(["-", "N/A", "", "ND"])
    return str(random.randint(0, 500))

DANE_NULL_RATE = 0.042  # 4.2% de registros sin código DANE institucional
def get_dane():
    if random.random() < DANE_NULL_RATE:
        return np.nan
    return random.choice(colegios_dane)""",
    )

    add_apa_heading(doc, "9.2 Módulo Principal del Pipeline Lakehouse (scripts/pipeline_educativo.py)", level=2)
    add_apa_heading(doc, "9.2.1 Inicialización Híbrida y Tolerancia a Fallos de Entorno", level=3)
    add_apa_paragraph(
        doc,
        "Las funciones comprobar_java_compatible() y crear_spark_session() evalúan mediante un "
        "subproceso la versión del entorno Java local. Dado que PySpark 4.x compila sus clases "
        "con versión de classfile 61.0 (Java 17+), si el computador local cuenta con Java 8 "
        "(1.8.x), intentar levantar la JVM lanzaría UnsupportedClassVersionError. Nuestra función "
        "anticipa esta condición y activa automáticamente el motor local en Pandas sin abortar "
        "la ejecución.",
        indent=True,
    )

    add_code_block(
        doc,
        "Archivo: scripts/pipeline_educativo.py | Funciones: comprobar_java_compatible() y crear_spark_session()",
        """def comprobar_java_compatible() -> bool:
    \"\"\"Verifica si la versión de Java instalada es compatible con Spark 4 (requiere Java 17+).\"\"\"
    try:
        import subprocess
        res = subprocess.run(["java", "-version"], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=5)
        salida = res.stderr or res.stdout
        if "1.8." in salida or "build 1.8" in salida:
            return False
        return True
    except Exception:
        return False

def crear_spark_session():
    if not HAY_PYSPARK or not comprobar_java_compatible():
        print("Entorno Java local (Java 8) no compatible con PySpark 4.x (requiere Java 17+).")
        print("Se activa el motor de procesamiento analítico local en Pandas.\\n")
        return None
    spark = (
        SparkSession.builder
        .appName("Pipeline_Educativo_Bogota")
        .master("local[*]")
        .config("spark.sql.shuffle.partitions", "4")
        .getOrCreate()
    )
    return spark""",
    )

    add_apa_heading(doc, "9.2.2 Validación de Puntajes y Saneamiento de Valores Centinela e Imputación DANE", level=3)
    add_apa_paragraph(
        doc,
        "En la Etapa 2 (capa Silver), la función validar_rango_puntaje() consulta en tiempo "
        "constante O(1) el conjunto hash VALORES_CENTINELA = {'-', 'N/A', 'ND', '', 'nan', 'none'}, "
        "convirtiendo las 1,070 cadenas inválidas en None/NaN y verificando que todo número "
        "pertenezca al rango oficial del ICFES [0.0, 500.0]. Enseguida, limpiar_valores_nulos() "
        "calcula la moda estadística de COLE_COD_DANE_INST e imputa los 2,934 vacíos para evitar "
        "la pérdida del 4.2% de la muestra estudiantil.",
        indent=True,
    )

    add_code_block(
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

    add_apa_heading(doc, "9.2.3 Regla de Negocio de Alertas de Deserción y Cruce Dimensional Gold", level=3)
    add_apa_paragraph(
        doc,
        "En las Etapas 3, 4 y 5 se ejecutan las reglas de negocio y la construcción del datamart. "
        "La función evaluar_alerta_localidad() compara la tasa de deserción de cada una de las 20 "
        "localidades frente a la constante UMBRAL_ALERTA = 8.0%, atrapando posibles excepciones "
        "de tipo (ValueError, TypeError) sin interrumpir el flujo. Por su parte, "
        "cruzar_saber11_con_directorio() aplica un LEFT JOIN entre los 70,000 registros de Saber 11 "
        "y los 2,262 colegios activos, logrando vincular exitosamente 63,498 estudiantes con los "
        "atributos institucionales de su colegio, mientras que generar_reporte_equidad_educativa() "
        "consolida la tabla maestra en data/ANALYTICS/reporte_equidad_educativa.csv.",
        indent=True,
    )

    add_code_block(
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
        return df_saber11.join(dir_sel, df_saber11["COLE_COD_DANE_INST"] == dir_sel["COD_DANE"], how="left").drop("COD_DANE")
    # Ejecución equivalente en Pandas con pd.merge(..., how="left")""",
    )

    add_apa_paragraph(
        doc,
        "Como evidencia verificable del procesamiento en la capa Gold, la Tabla 6 presenta los "
        "resultados reales del archivo reporte_equidad_educativa.csv para las localidades que "
        "superan el umbral crítico del 8.0% o presentan los menores puntajes globales, la "
        "Figura 3 grafica la distribución completa en las 20 localidades y la Tabla 7 resume la "
        "tasa de absorción y graduación en las 10 Instituciones de Educación Superior (IES).",
        indent=True,
    )

    add_apa_table(
        doc,
        num=6,
        title="Evidencia de resultados en Capa Gold: localidades con alerta crítica de deserción o menor puntaje Saber 11",
        headers=["Localidad", "Puntaje Global Promedio", "Estudiantes Evaluados", "Nivel Alerta Origen", "Tasa de Deserción (%)", "Dictamen Regla Automática (> 8.0%)"],
        rows=[
            ["Usme", "251.89", "3,494", "BAJO", "12.07%", "ALERTA CRÍTICA (Supera umbral 8.0%)"],
            ["Fontibón", "250.15", "3,453", "CRÍTICO", "11.77%", "ALERTA CRÍTICA (Supera umbral 8.0%)"],
            ["Engativá", "249.67", "3,413", "MEDIO", "11.26%", "ALERTA CRÍTICA (Supera umbral 8.0%)"],
            ["Los Mártires", "248.38", "3,422", "BAJO", "11.22%", "ALERTA CRÍTICA (Supera umbral 8.0%)"],
            ["Tunjuelito", "246.13", "3,464", "ALTO", "10.92%", "ALERTA CRÍTICA (Supera umbral 8.0%)"],
            ["Usaquén", "250.68", "3,412", "MEDIO", "9.72%", "ALERTA CRÍTICA (Supera umbral 8.0%)"],
            ["Kennedy", "249.95", "3,385", "BAJO", "9.08%", "ALERTA CRÍTICA (Supera umbral 8.0%)"],
            ["La Candelaria", "250.91", "3,536", "ALTO", "8.94%", "ALERTA CRÍTICA (Supera umbral 8.0%)"],
            ["Antonio Nariño", "250.29", "3,418", "ALTO", "8.12%", "ALERTA CRÍTICA (Supera umbral 8.0%)"],
            ["Ciudad Bolívar", "244.63", "3,397", "CRÍTICO", "3.16%", "Menor puntaje global del Distrito (244.63)"],
        ],
        note_text="Datos extraídos directamente de data/ANALYTICS/reporte_equidad_educativa.csv tras la ejecución de pipeline_educativo.py.",
    )

    add_apa_figure(
        doc,
        num=3,
        title="Evaluación automática de tasas de deserción escolar y puntaje promedio Saber 11 en las 20 localidades",
        img_path=os.path.join(FIGS_DIR, "fig3_equidad_desercion_localidades.png"),
        note_text="Las barras en rojo identifican las 9 localidades que superan el umbral crítico del 8.0% de deserción escolar y disparan alerta operativa automática.",
    )

    add_apa_table(
        doc,
        num=7,
        title="Evidencia en Capa Gold: tasa de absorción y graduación en Educación Superior (Jóvenes a la E y La U en tu Colegio)",
        headers=["Institución de Educación Superior (IES)", "Total Admitidos", "Graduados", "Activos", "Tasa de Graduación (%)"],
        rows=[
            ["Corporación Universitaria Minuto de Dios (UNIMINUTO)", "1,293", "494", "521", "38.21%"],
            ["Universidad Nacional de Colombia", "1,377", "503", "565", "36.53%"],
            ["Servicio Nacional de Aprendizaje (SENA)", "1,357", "483", "532", "35.59%"],
            ["Universidad de la Salle", "1,320", "464", "500", "35.15%"],
            ["Universidad Santo Tomás", "1,327", "466", "536", "35.12%"],
            ["Universidad de los Andes", "1,312", "459", "545", "34.98%"],
            ["Universidad Distrital Francisco José de Caldas", "1,390", "477", "543", "34.32%"],
            ["Universidad Externado de Colombia", "1,360", "465", "558", "34.19%"],
            ["Politécnico Grancolombiano", "1,237", "419", "513", "33.87%"],
            ["Pontificia Universidad Javeriana", "1,334", "438", "559", "32.83%"],
        ],
        note_text="Datos extraídos de data/ANALYTICS/absorcion_educacion_superior.csv (13,307 registros válidos tras excluir 1,693 nulos en IES_ADMITIDA).",
    )

    # ══════════════════════════════════════════════════════════════════════════
    # 9.3 MÓDULO DE STREAMING DEL CORTE 3 (scripts/streaming_alertas_educativas.py)
    # ══════════════════════════════════════════════════════════════════════════
    add_apa_heading(doc, "9.3 Módulo de Procesamiento de Eventos en Tiempo Real, Corte 3 (scripts/streaming_alertas_educativas.py)", level=2)
    add_apa_paragraph(
        doc,
        "Para materializar en código funcional los temas del Tercer Corte, se implementó el "
        "módulo scripts/streaming_alertas_educativas.py, integrado como Etapa 6 del pipeline "
        "principal. El módulo consta de tres bloques algorítmicos fundamentales: (a) el "
        "intermediario orientado a eventos BrokerEventosEducativos con control de idempotencia "
        "y cola de mensajes muertos (DLQ); (b) el procesador de flujo con cálculo dinámico de "
        "Watermarks y ventanas Tumbling, Sliding y Session; y (c) la especificación declarativa "
        "equivalente para PySpark Structured Streaming.",
        indent=True,
    )

    add_code_block(
        doc,
        "Archivo: scripts/streaming_alertas_educativas.py | Clase: BrokerEventosEducativos (EDA e Idempotencia)",
        """class BrokerEventosEducativos:
    \"\"\"Broker EDA con tópicos particionados, offsets secuenciales, idempotencia y DLQ.\"\"\"
    def __init__(self):
        self.topicos = {"sed.alertas.asistencia": [], "sed.alertas.dlq": []}
        self.suscriptores = {"sed.alertas.asistencia": [], "sed.alertas.dlq": []}
        self.ids_procesados: set[str] = set()
        self.duplicados_ignorados: list[dict] = []

    def publicar(self, topico: str, evento_cloudevent: dict) -> bool:
        event_id = evento_cloudevent.get("id", "")
        # Control de idempotencia O(1): descarta reintentos duplicados de red
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
        return True""",
    )

    add_apa_paragraph(
        doc,
        "En el motor temporal procesar_flujo_con_watermarks_y_ventanas(), cada vez que ingresa un "
        "evento único se extrae su marca event_time (ET) y processing_time (PT), se determina su "
        "intervalo semiabierto [w_start, w_end) y se compara contra el Watermark vigente "
        "(WM = max(ET) - 5 minutos). Como se observa en el fragmento siguiente, el algoritmo "
        "distingue entre eventos que llegan con la ventana abierta (WM < w_end), eventos tardíos "
        "moderados cubiertos por Allowed Lateness (w_end <= WM <= w_end + 5 min) y eventos "
        "extremadamente tardíos (WM > w_end + 5 min), los cuales se desvían al tópico "
        "sed.alertas.dlq.",
        indent=True,
    )

    add_code_block(
        doc,
        "Archivo: scripts/streaming_alertas_educativas.py | Función: consumidor_principal() y expresión PySpark",
        """# 1. Evaluación del evento frente al Watermark vigente (WM = max(event_time) - 5 min)
w_start = obtener_inicio_ventana_fija(et, TUMBLING_SIZE_MIN)
w_end = w_start + timedelta(minutes=TUMBLING_SIZE_MIN)

if wm_previo is None or wm_previo < w_end:
    estado_arrivo = "ACEPTADO_EN_VENTANA"
elif wm_previo <= w_end + timedelta(minutes=ALLOWED_LATENESS_MIN):
    estado_arrivo = "ALLOWED_LATENESS"  # Reabre estado y emite corrección
else:
    estado_arrivo = "RECHAZADO_EXTREMO_DLQ"  # Desvío a Side Output / DLQ sin borrar alerta

if max_event_time is None or et > max_event_time:
    max_event_time = et
    watermark_actual = max_event_time - timedelta(minutes=UMBRAL_WATERMARK_MIN)

# 2. Equivalencia declarativa en PySpark Structured Streaming (Spark 4.x)
stream_base = (
    spark.readStream.schema(esquema_cloudevents_sed)
    .json("data/RAW/eventos_asistencia_desercion_stream.jsonl")
    .withWatermark("event_time", "5 minutes")
)
tumbling_df = stream_base.groupBy(window(col("event_time"), "10 minutes"), col("data.localidad")).count()
sliding_df  = stream_base.groupBy(window(col("event_time"), "10 minutes", "2 minutes"), col("data.localidad")).count().where("count >= 3")
session_df  = stream_base.groupBy(session_window(col("event_time"), "4 minutes"), col("data.cod_dane")).count()""",
    )

    add_apa_paragraph(
        doc,
        "La ejecución real del módulo de streaming procesó 18 eventos recibidos en el archivo "
        "JSONL: descartó 1 evento duplicado por idempotencia (reintento de EVT-004 en Bosa) y "
        "procesó 17 eventos únicos. La Tabla 8 muestra la evidencia empírica de la distorsión "
        "que ocurre al agrupar por Processing Time frente a Event Time, junto con el comportamiento "
        "de los tres eventos fuera de orden más representativos del flujo (EVT-009, EVT-013 y "
        "EVT-017), lo cual se ilustra gráficamente en la Figura 4.",
        indent=True,
    )

    add_apa_table(
        doc,
        num=8,
        title="Evidencia experimental (Corte 3): comparación Event Time vs. Processing Time y trazabilidad de eventos tardíos",
        headers=["Caso / Ventana", "Event Time (Colegio)", "Processing Time (Servidor SED)", "Watermark al Llegar", "Retraso (min)", "Resultado y Política Aplicada"],
        rows=[
            ["Ventana [08:00, 08:10)", "8 eventos reales", "6 eventos vistos", "Cierra en WM=08:11", "N/A", "Processing Time subestima en -2 alertas el primer bloque matutino."],
            ["Ventana [08:10, 08:20)", "5 eventos reales", "7 eventos vistos", "Cierra en WM=08:21", "N/A", "Processing Time infla en +2 alertas el segundo bloque por retrasos de red."],
            ["EVT-009 (Usme)", "08:09:10", "08:13:30", "08:08:00 (< 08:10)", "4.33 min", "ACEPTADO_EN_VENTANA: llega fuera de orden pero dentro del umbral WM (5 min)."],
            ["EVT-013 (Usme Rural)", "08:07:30", "08:19:30", "08:14:00 (<= 08:15)", "12.00 min", "ALLOWED_LATENESS: ventana [08:00, 08:10) ya cerrada, pero dentro de holgura (+5m); re-emite corrección."],
            ["EVT-017 (Sumapaz Rural)", "08:03:00", "08:31:00", "08:24:00 (> 08:15)", "28.00 min", "RECHAZADO_EXTREMO_DLQ: se desvía a cola de conciliación DLQ de la SED sin perder el reporte."],
        ],
        note_text="Datos extraídos de data/ANALYTICS/streaming_traza_watermarks.csv y streaming_comparacion_event_vs_processing.csv.",
    )

    add_apa_figure(
        doc,
        num=4,
        title="Distorsión entre Event Time y Processing Time y clasificación de eventos según Watermark y Allowed Lateness",
        img_path=os.path.join(FIGS_DIR, "fig4_streaming_watermarks_ventanas.png"),
        note_text="El panel A demuestra cómo el retraso de red desplaza eventos del bloque [08:00, 08:10) hacia bloques posteriores en Processing Time. El panel B muestra en verde los eventos aceptados en ventana, en ámbar el evento corregido por Allowed Lateness (EVT-013) y en rojo el evento desviado a DLQ (EVT-017).",
    )

    add_apa_paragraph(
        doc,
        "Finalmente, la Tabla 9 resume las salidas generadas por las tres tipologías de ventanas "
        "de tiempo en la capa Gold, demostrando cómo la ventana Sliding (10 min cada 2 min) logra "
        "detectar la ráfaga crítica de 3 alertas en Bosa en el intervalo [08:02, 08:12), la cual "
        "pasaba desapercibida en la ventana Tumbling fija porque 2 eventos cayeron en [08:00, 08:10) "
        "y el tercero en [08:10, 08:20), y cómo las Session Windows separaron automáticamente "
        "dos sesiones distintas de reporte en el colegio 110010190012 de Ciudad Bolívar tras una "
        "pausa superior al gap de 4 minutos.",
        indent=True,
    )

    add_apa_table(
        doc,
        num=9,
        title="Síntesis de resultados de Ventanas Tumbling, Sliding y Session en la capa Gold (data/ANALYTICS/)",
        headers=["Tipo de Ventana", "Intervalo / Clave", "Localidad / Colegio DANE", "Eventos Agrupados", "Riesgo Prom. / Máx.", "Hallazgo Operativo para la SED"],
        rows=[
            ["Tumbling (10 min fija)", "[08:00, 08:10)", "Ciudad Bolívar", "3 alertas", "88.17 (Prom.)", "Cerrada normalmente por avance del Watermark."],
            ["Tumbling (10 min fija)", "[08:00, 08:10)", "Usme", "2 alertas", "92.00 (Prom.)", "CORREGIDA dinámicamente al ingresar EVT-013 vía Allowed Lateness."],
            ["Sliding (10m / slide 2m)", "[08:02, 08:12)", "Bosa", "3 alertas", "85.50 (Prom.)", "RÁFAGA CRÍTICA detectada entre bloques fijos -> Despachar unidad móvil."],
            ["Sliding (10m / slide 2m)", "[08:10, 08:20)", "Kennedy", "3 alertas", "84.33 (Prom.)", "RÁFAGA CRÍTICA confirmada en 4 ventanas deslizantes consecutivas."],
            ["Session (gap 4 min)", "Sesión #1 (08:01 a 08:08)", "DANE 110010190012 (C. Bolívar)", "3 alertas", "92.00 (Máx.)", "Sesión continua de 7.08 min del orientador escolar fusionando 3 casos."],
            ["Session (gap 4 min)", "Sesión #2 (08:29 a 08:33)", "DANE 110010190012 (C. Bolívar)", "1 alerta", "87.00 (Máx.)", "Nueva sesión independiente abierta tras 24 min de inactividad (> gap 4m)."],
            ["Session (gap 4 min)", "Sesión #1 (08:05 a 08:15)", "DANE 110010070045 (Bosa)", "3 alertas", "94.00 (Máx.)", "Sesión activa de 10.17 min que agrupa 3 reportes consecutivos en Bosa."],
        ],
        note_text="Datos extraídos de streaming_ventanas_tumbling.csv, streaming_ventanas_sliding_rafagas.csv y streaming_ventanas_session_colegios.csv.",
    )

    add_apa_heading(doc, "9.4 Pruebas de Verificación y Casos Límite", level=2)
    add_apa_paragraph(
        doc,
        "La Tabla 10 documenta las pruebas ejecutadas sobre las funciones críticas del sistema "
        "para validar su comportamiento frente a entradas anómalas y casos de borde.",
        indent=True,
    )

    add_apa_table(
        doc,
        num=10,
        title="Matriz de pruebas funcionales y verificación de casos límite del software",
        headers=["ID Prueba", "Componente / Función", "Entrada de Prueba", "Resultado Esperado", "Resultado Obtenido"],
        rows=[
            ["PRB-01", "validar_rango_puntaje()", "valor_str = '-' y 'N/A'", "Retornar None sin lanzar ValueError", "None retornado; 1,070 centinelas saneados correctamente."],
            ["PRB-02", "validar_rango_puntaje()", "valor_str = '540.0' (fuera de rango)", "Retornar None por superar límite 500.0", "None retornado; protege media estadística."],
            ["PRB-03", "limpiar_valores_nulos()", "Columna COLE_COD_DANE_INST (2,934 nulos)", "0 nulos residuales tras imputar moda", "0 nulos residuales (moda imputada: 1100127417599)."],
            ["PRB-04", "evaluar_alerta_localidad()", "loc='FONTIBÓN', tasa=11.77", "Cadena de ALERTA CRÍTICA (> 8.0%)", "ALERTA CRÍTICA emitida para 9 localidades."],
            ["PRB-05", "BrokerEventosEducativos.publicar()", "Segundo envío de EVT-004 (duplicado)", "Retornar False e ignorar duplicado", "Duplicado bloqueado en O(1); 17 eventos únicos."],
            ["PRB-06", "procesar_flujo_con_watermarks_y_ventanas()", "EVT-013 (12m tarde) y EVT-017 (28m tarde)", "EVT-013 en Allowed Lateness; EVT-017 en DLQ", "Ventana Usme corregida y EVT-017 persistido en DLQ."],
        ],
        note_text="Elaboración propia. Todas las pruebas fueron verificadas sobre la versión 2.0.0 del repositorio.",
    )

    # ══════════════════════════════════════════════════════════════════════════
    # 10. IDENTIFICACIÓN CLARA Y ARGUMENTADA DE LOS APORTES DE CADA INTEGRANTE
    # ══════════════════════════════════════════════════════════════════════════
    add_apa_heading(doc, "10. Identificación y Argumentación de los Aportes de Cada Integrante", level=1)
    add_apa_paragraph(
        doc,
        "En cumplimiento estricto de las orientaciones de la entrega final, a continuación se "
        "presenta de forma individual, detallada y argumentada el aporte técnico y conceptual "
        "de cada uno de los dos integrantes del equipo, especificando qué desarrolló, cómo lo "
        "implementó algorítmicamente y en qué secciones del presente manual técnico se encuentra "
        "documentado.",
        indent=True,
    )

    add_apa_heading(doc, "10.1 Aportes de David Santiago Castelblanco Artunduaga (Código: 67001182)", level=2)
    add_apa_paragraph(
        doc,
        "David Santiago Castelblanco Artunduaga lideró el diseño arquitectónico del Data Lakehouse "
        "(capas Bronze, Silver y Gold), la programación del motor de resiliencia híbrida "
        "PySpark/Pandas, los cruces relacionales y agregaciones del modelo dimensional en estrella, "
        "y la implementación de los algoritmos de agrupación por Ventanas de Tiempo (Tumbling, "
        "Sliding y Session Windows) del Tercer Corte, así como la generación automatizada de "
        "visualizaciones técnicas.",
        indent=True,
    )
    add_apa_bullet(
        doc,
        "¿Qué desarrolló?: ",
        "Desarrolló (1) el subsistema de inicialización distribuida y conmutación automática "
        "Spark/Pandas (comprobar_java_compatible y crear_spark_session); (2) las funciones de "
        "cruce relacional dimensional y agregación analítica en la capa Gold "
        "(calcular_puntaje_promedio_por_localidad, cruzar_saber11_con_directorio, "
        "comparar_metricas_por_localidad y generar_reporte_equidad_educativa en "
        "scripts/pipeline_educativo.py); (3) la lógica de segmentación de flujos infinitos en "
        "bloques finitos mediante Tumbling Windows (10 min), Sliding Windows (10 min con avance "
        "de 2 min para detección de ráfagas) y Session Windows (brecha dinámica de 4 min por "
        "colegio DANE) junto con su formulación en PySpark Structured Streaming dentro de "
        "scripts/streaming_alertas_educativas.py; y (4) el script de graficación "
        "scripts/generar_graficos_evidencias.py.",
    )
    add_apa_bullet(
        doc,
        "¿Cómo lo implementó?: ",
        "Implementó la detección preventiva del runtime de Java mediante subprocess.run(['java', '-version']), "
        "evitando fallos fatales de la JVM cuando Spark 4.x se ejecuta sobre máquinas con Java 8, "
        "y diseñó cada función analítica con doble rama de ejecución (es_spark=True usando la API "
        "declarativa pyspark.sql.functions.col, groupBy, agg, join y orderBy; y es_spark=False "
        "usando operaciones vectorizadas pd.merge(..., how='left') y groupby en Pandas). En el "
        "módulo de streaming del Corte 3, implementó la aritmética de intervalos semiabiertos "
        "[inicio, fin) en obtener_inicio_ventana_fija(), el barrido deslizante con factor de "
        "multiplicación 5x (Duración / Slide = 10 / 2) que evalúa la condición de ráfaga "
        "len(evs) >= 3, y el algoritmo de fusión dinámica de sesiones por clave DANE cuando "
        "(et_actual - et_anterior) <= timedelta(minutes=4).",
    )
    add_apa_bullet(
        doc,
        "¿Dónde está documentado en el manual?: ",
        "Se encuentra documentado en las Secciones 5.1 y 5.2 (fundamentación Lakehouse, modelo "
        "estrella y pipelines ETL/ELT), Sección 5.6 (teoría y costo en memoria de ventanas "
        "Tumbling, Sliding y Session), Sección 7 (estructura lógica y Figura 1), Sección 9.2.1 "
        "(código de resiliencia híbrida Spark/Pandas), Sección 9.2.3 (código del LEFT JOIN y "
        "Tabla 6 / Figura 3 de equidad educativa) y Sección 9.3 (código de ventanas en streaming "
        "y Tabla 9 de resultados).",
    )

    add_apa_heading(doc, "10.2 Aportes de Dairo José Narváez Morales (Código: 67002032)", level=2)
    add_apa_paragraph(
        doc,
        "Dairo José Narváez Morales lideró la ingeniería de generación y perfilamiento de la capa "
        "Bronze, el diseño e implementación del contrato de calidad y limpieza estadística en la "
        "capa Silver bajo el marco DAMA-DMBOK, el motor de reglas de negocio de alertas tempranas "
        "y absorción universitaria, y la implementación de la Arquitectura Orientada a Eventos "
        "(Broker idempotente CloudEvents 1.0), cálculo de Watermarks y tratamiento de eventos "
        "tardíos (Allowed Lateness y Side Output / DLQ) del Tercer Corte, además de la "
        "estructuración del manual de usuario.",
        indent=True,
    )
    add_apa_bullet(
        doc,
        "¿Qué desarrolló?: ",
        "Desarrolló (1) el generador reproducible de datos crudos sintéticos "
        "scripts/generar_datos_raw.py (87,520 registros en 4 archivos CSV con semilla 42 y "
        "perfilamiento automático en consola); (2) las funciones de saneamiento y contrato de "
        "calidad en la capa Silver (validar_rango_puntaje, limpiar_valores_nulos, "
        "filtrar_colegios_activos) y las reglas de negocio (evaluar_alerta_localidad, "
        "analizar_alertas_criticas, calcular_beneficiarios_por_localidad_programa y "
        "analizar_absorcion_educacion_superior en scripts/pipeline_educativo.py); y (3) en "
        "scripts/streaming_alertas_educativas.py, la clase BrokerEventosEducativos, la función "
        "criptográfica anonimizar_estudiante() (Ley 1581 de 2012), el generador de flujos "
        "CloudEvents 1.0 generar_flujo_eventos_educativos(), la lógica de evolución del "
        "Watermark (WM = max(ET) - 5 min), la comparativa Event Time vs. Processing Time y el "
        "enrutamiento de eventos tardíos hacia Allowed Lateness y la cola DLQ.",
    )
    add_apa_bullet(
        doc,
        "¿Cómo lo implementó?: ",
        "Implementó la limpieza de valores centinela apoyándose en una estructura hash set "
        "(VALORES_CENTINELA) que garantiza búsquedas en tiempo constante O(1) por registro y "
        "programó la imputación estadística de los 2,934 códigos DANE faltantes mediante el "
        "cálculo dinámico de la moda (valores_validos.mode()[0]). En el motor de reglas, utilizó "
        "bloques try/except (ValueError, TypeError) para blindar la comparación frente a "
        "UMBRAL_ALERTA (8.0%). En el módulo de streaming del Corte 3, implementó la verificación "
        "de idempotencia guardando cada event_id en el conjunto self.ids_procesados antes de "
        "asignar el offset del tópico, y programó la regla de decisión tripartita que contrasta "
        "el Watermark vigente frente al límite superior de la ventana (w_end) y la holgura "
        "adicional (w_end + 5 min) para decidir entre acumulación normal, re-emisión por "
        "Allowed Lateness o publicación en el tópico sed.alertas.dlq.",
    )
    add_apa_bullet(
        doc,
        "¿Dónde está documentado en el manual?: ",
        "Se encuentra documentado en las Secciones 5.3, 5.4 y 5.5 (calidad DAMA-DMBOK, "
        "Arquitectura Orientada a Eventos, Event Time vs. Processing Time y Watermarks), Sección 6 "
        "(marco normativo y Tabla 3), Sección 8 (diagnóstico de calidad, Tabla 5 y Figura 2), "
        "Sección 9.1 (código de generar_datos_raw.py), Sección 9.2.2 (código de limpieza e "
        "imputación), Sección 9.2.3 (código de evaluar_alerta_localidad y Tabla 7 de absorción "
        "IES), Sección 9.3 (código de BrokerEventosEducativos, Tabla 8 y Figura 4 de Watermarks "
        "y DLQ) y Sección 11 (Manual de Usuario).",
    )

    add_apa_table(
        doc,
        num=11,
        title="Matriz síntesis de aportes individuales por integrante, archivos implementados y localización en el manual",
        headers=["Integrante y Código", "Componentes Desarrollados", "Implementación Técnica Clave (Archivos y Funciones)", "Secciones del Manual donde se Documenta"],
        rows=[
            [
                "David Santiago Castelblanco Artunduaga\n(Cód. 67001182)",
                "• Arquitectura Lakehouse Medallón\n• Motor híbrido PySpark / Pandas\n• Cruces LEFT JOIN y Datamart Gold\n• Ventanas Tumbling, Sliding y Session (Corte 3)\n• Figuras técnicas APA 7",
                "• pipeline_educativo.py: comprobar_java_compatible(), crear_spark_session(), cruzar_saber11_con_directorio(), calcular_puntaje_promedio_por_localidad(), generar_reporte_equidad_educativa()\n• streaming_alertas_educativas.py: obtener_inicio_ventana_fija(), bloques Tumbling/Sliding/Session y código PySpark\n• generar_graficos_evidencias.py",
                "Secciones 1, 3, 5.1, 5.2, 5.6, 5.7 (Tabla 2), 7 (Figura 1), 9.2.1, 9.2.3 (Tabla 6, Figura 3), 9.3 (Tabla 9) y 10.1."
            ],
            [
                "Dairo José Narváez Morales\n(Cód. 67002032)",
                "• Generador sintético RAW y perfilamiento\n• Contrato de calidad Silver (centinelas y moda DANE)\n• Reglas de alertas (>8.0%) y absorción IES\n• Broker EDA idempotente, Watermarks y políticas de eventos tardíos / DLQ (Corte 3)\n• Manual de Usuario",
                "• generar_datos_raw.py: puntaje_raw(), get_dane(), perfilamiento RAW\n• pipeline_educativo.py: validar_rango_puntaje(), limpiar_valores_nulos(), filtrar_colegios_activos(), evaluar_alerta_localidad(), analizar_absorcion_educacion_superior()\n• streaming_alertas_educativas.py: BrokerEventosEducativos, anonimizar_estudiante(), evolución Watermark y DLQ",
                "Secciones 2, 4 (Tabla 1), 5.3, 5.4, 5.5, 6 (Tabla 3), 8 (Tabla 5, Figura 2), 9.1, 9.2.2, 9.2.3 (Tabla 7), 9.3 (Tabla 8, Figura 4), 9.4 (Tabla 10), 10.2 y 11."
            ],
        ],
        note_text="Elaboración propia. Ambos integrantes participaron activamente en la integración final, pruebas y sustentación.",
    )

    # ══════════════════════════════════════════════════════════════════════════
    # 11. MANUAL DE USUARIO (OPERACIÓN DEL SISTEMA PASO A PASO)
    # ══════════════════════════════════════════════════════════════════════════
    add_apa_heading(doc, "11. Manual de Usuario (Operación del Sistema)", level=1)
    add_apa_paragraph(
        doc,
        "Este apartado está dirigido a analistas de la Secretaría de Educación del Distrito (SED), "
        "directivos docentes e investigadores que deseen ejecutar el sistema de forma autónoma "
        "desde la interfaz de línea de comandos (CLI).",
        indent=True,
    )

    add_apa_heading(doc, "11.1 Requisitos del Sistema e Instalación", level=2)
    add_apa_bullet(
        doc,
        "Sistema operativo y hardware: ",
        "Computador con Windows 10/11, macOS o Linux, mínimo 4 GB de RAM y 200 MB de espacio libre en disco.",
    )
    add_apa_bullet(
        doc,
        "Entorno Python y dependencias: ",
        "Contar con Python 3.9 o superior instalado y ejecutar en la terminal: "
        "pip install pandas numpy pyspark python-docx python-pptx matplotlib",
    )

    add_apa_heading(doc, "11.2 Procedimientos de Uso Paso a Paso", level=2)
    add_apa_paragraph(
        doc,
        "1. Abrir PowerShell o la terminal del sistema, ubicarse en la carpeta raíz "
        "Proyecto-ING-DATO y ejecutar: python scripts/generar_datos_raw.py. Verificar en pantalla "
        "la confirmación de creación de los 4 archivos CSV en data/RAW/ junto con el resumen de "
        "perfilamiento inicial.",
        indent=False,
        bold_prefix="Tarea 1, Generar o actualizar los datos crudos de la Capa Bronze (RAW): ",
    )
    add_apa_paragraph(
        doc,
        "2. Ejecutar el comando: python scripts/pipeline_educativo.py. El orquestador ejecutará "
        "de manera automatizada las 6 etapas del sistema: (Etapa 1) carga y diagnóstico de nulos; "
        "(Etapa 2) limpieza de centinelas, imputación por moda DANE y guardado en data/STAGING/; "
        "(Etapa 3) evaluación de la regla de alertas de deserción (> 8.0%) y absorción IES; "
        "(Etapa 4) cruce relacional LEFT JOIN con los 2,262 colegios activos; (Etapa 5) generación "
        "del reporte de equidad educativa en data/ANALYTICS/; y (Etapa 6, Corte 3) procesamiento "
        "del flujo de eventos en tiempo real con Watermarks, ventanas Tumbling/Sliding/Session y "
        "cola DLQ.",
        indent=False,
        bold_prefix="Tarea 2, Ejecutar el pipeline integral (Batch Lakehouse + Streaming Corte 3): ",
    )
    add_apa_paragraph(
        doc,
        "3. Si se desea ejecutar únicamente el motor de procesamiento de eventos en tiempo real "
        "del Tercer Corte para inspeccionar en detalle la traza de Watermarks y la cola DLQ, "
        "ejecutar: python scripts/streaming_alertas_educativas.py.",
        indent=False,
        bold_prefix="Tarea 3, Ejecutar de forma independiente la simulación de Streaming (Corte 3): ",
    )
    add_apa_paragraph(
        doc,
        "4. Dirigirse al directorio data/ANALYTICS/ desde el Explorador de Archivos o importar "
        "los 11 archivos CSV generados directamente en Microsoft Excel o Power BI para visualizar "
        "los tableros de equidad, ráfagas críticas de inasistencia y sesiones por colegio.",
        indent=False,
        bold_prefix="Tarea 4, Consultar los reportes analíticos y figuras en la Capa Gold: ",
    )

    add_apa_heading(doc, "11.3 Guía de Mensajes y Solución de Problemas Frecuentes", level=2)
    add_apa_table(
        doc,
        num=12,
        title="Guía rápida de mensajes del sistema y solución de incidentes frecuentes para el usuario",
        headers=["Mensaje o Incidente en Consola", "Causa Técnica", "Acción Recomendada para el Usuario"],
        rows=[
            ["Error: Archivo data/RAW/... no encontrado", "Se ejecutó el pipeline sin haber creado primero los datasets crudos.", "Ejecutar primero: python scripts/generar_datos_raw.py"],
            ["Entorno Java local (Java 8) no compatible con PySpark 4.x", "La estación local tiene Java 8 y Spark 4 requiere Java 17+.", "No requiere acción: el sistema activa automáticamente el motor Pandas y completa las 6 etapas."],
            ["ALERTA CRÍTICA: [Localidad], Tasa supera 8.0%", "La tasa de deserción de la localidad rebasa el límite distrital.", "Priorizar la localidad en los planes de intervención y búsqueda activa de la SED."],
            ["RÁFAGA CRÍTICA (3 alertas en 10 min)", "La ventana Sliding detectó acumulación rápida de alertas en una localidad.", "Despachar unidad móvil de atención psicosocial a los colegios reportados."],
            ["RECHAZADO_EXTREMO_DLQ (EVT-017 Sumapaz)", "Un evento llegó 28 min tarde tras cerrarse la ventana y su Allowed Lateness.", "Consultar data/ANALYTICS/streaming_eventos_tardios_dlq.csv para conciliación administrativa."],
        ],
        note_text="Elaboración propia para soporte operativo de analistas de la SED y usuarios académicos.",
    )

    # ══════════════════════════════════════════════════════════════════════════
    # 12. RIESGOS, LIMITACIONES Y PLAN DE TRABAJO POR FASES
    # ══════════════════════════════════════════════════════════════════════════
    add_apa_heading(doc, "12. Riesgos, Limitaciones y Plan de Trabajo por Fases", level=1)
    add_apa_paragraph(
        doc,
        "El proyecto se ejecutó a lo largo de seis fases incrementales que culminaron con éxito "
        "en la presente entrega final: Fase 1 (Definición del problema socioeducativo y "
        "arquitectura Lakehouse), Fase 2 (Generación e ingesta inmutable en capa Bronze), Fase 3 "
        "(Perfilamiento y contrato de calidad en capa Silver), Fase 4 (Modelado dimensional Gold "
        "y reglas de alerta), Fase 5 (Implementación del motor de streaming orientado a eventos, "
        "Watermarks, ventanas de tiempo y DLQ del Tercer Corte) y Fase 6 (Validación integral, "
        "documentación APA 7 y sustentación).",
        indent=True,
    )
    add_apa_paragraph(
        doc,
        "Como limitaciones actuales de la versión 2.0.0 y líneas de evolución para un despliegue "
        "productivo en la Secretaría de Educación del Distrito, se identifican tres aspectos: "
        "(a) el procesamiento se ejecuta actualmente en nodo local (local[*]), por lo que su "
        "paso a producción distrital requerirá aprovisionar un clúster administrado como Google "
        "Cloud Dataproc o Databricks; (b) los archivos de entrada replican estadísticamente el "
        "comportamiento de Bogotá mediante simulación controlada, debiendo conectarse en "
        "producción directamente a los endpoints REST de la API CKAN de Datos Abiertos Bogotá y "
        "al sistema SIMAT; y (c) la construcción futura de tableros visuales interactivos en "
        "Power BI conectados directamente a las tablas generadas en la capa Gold.",
        indent=True,
    )

    # ══════════════════════════════════════════════════════════════════════════
    # 13. CONCLUSIONES
    # ══════════════════════════════════════════════════════════════════════════
    add_apa_heading(doc, "13. Conclusiones", level=1)
    add_apa_paragraph(
        doc,
        "1. La implementación de la arquitectura Data Lakehouse bajo el patrón Medallón (Bronze, "
        "Silver y Gold) demostró ser una solución eficaz y auditable para resolver la "
        "fragmentación de los datos abiertos educativos de Bogotá D.C. Al preservar intactos los "
        "archivos originales en data/RAW/ y aplicar contratos programáticos de calidad en "
        "data/STAGING/, se logró sanear el 100% de los 1,070 valores centinela en puntajes "
        "Saber 11 y recuperar mediante imputación por moda los 2,934 registros estudiantiles que "
        "carecían de código DANE, garantizando cruces dimensionales sin pérdida de muestra.",
        indent=True,
    )
    add_apa_paragraph(
        doc,
        "2. La incorporación práctica de los conceptos del Tercer Corte (Arquitectura Orientada a "
        "Eventos, diferenciación entre Event Time y Processing Time, Watermarks, ventanas "
        "Tumbling, Sliding y Session, y manejo de eventos tardíos) permitió extender el alcance "
        "del proyecto desde el análisis histórico por lotes hacia el monitoreo preventivo en "
        "tiempo real. Quedó demostrado experimentalmente que procesar por Processing Time "
        "distorsiona los conteos por ventana cuando existen retrasos de red en colegios "
        "periféricos, mientras que el uso de Event Time con un Watermark de 5 minutos, "
        "complementado con Allowed Lateness (+5 min) y una cola de conciliación DLQ, evita que "
        "alertas críticas de zonas rurales como Sumapaz o Usme se pierdan silenciosamente.",
        indent=True,
    )
    add_apa_paragraph(
        doc,
        "3. La articulación de las tres tipologías de ventanas de tiempo demostró que cada "
        "operador temporal responde a una necesidad operativa distinta e irreemplazable dentro "
        "de la Secretaría de Educación: las ventanas fijas (Tumbling) aseguran reportes "
        "periódicos sin doble conteo; las ventanas deslizantes (Sliding) detectan ráfagas de "
        "inasistencia que ocurren en la frontera entre dos bloques fijos (como el caso de Bosa "
        "entre las 08:02 y las 08:12); y las ventanas de sesión (Session) permiten caracterizar "
        "el flujo real de trabajo de los orientadores escolares en cada institución educativa.",
        indent=True,
    )

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════════
    # REFERENCIAS (FORMATO APA 7.ª EDICIÓN CON SANGRÍA FRANCESA)
    # ══════════════════════════════════════════════════════════════════════════
    add_apa_heading(doc, "Referencias", level=1)

    referencias = [
        "Akidau, T., Bradshaw, R., Chambers, C., Chernyak, S., Fernández-Moctezuma, R. J., Lax, R., McVeety, S., Mills, D., Perry, F., Schmidt, E., & Whittle, S. (2015). The dataflow model: A practical approach to balancing correctness, latency, and cost in massive-scale, unbounded, out-of-order data processing. Proceedings of the VLDB Endowment, 8(12), 1792-1803. https://doi.org/10.14778/2824032.2824076",
        "Akidau, T., Chernyak, S., & Lax, R. (2018). Streaming systems: The what, where, when, and how of large-scale data processing. O'Reilly Media.",
        "Alcaldía Mayor de Bogotá D.C. (2024). Cupos escolares en Bogotá: Distrito busca población desescolarizada en localidades priorizadas. Secretaría de Educación del Distrito. https://www.educacionbogota.edu.co/",
        "Apache Software Foundation. (2026). Spark Structured Streaming programming guide: Window operations on event time and handling late data (Version 4.0). https://spark.apache.org/docs/latest/structured-streaming-programming-guide.html",
        "Cloud Native Computing Foundation [CNCF]. (2024). CloudEvents specification (Version 1.0.2). https://cloudevents.io/",
        "Congreso de la República de Colombia. (2012). Ley Estatutaria 1581 de 2012: Por la cual se dictan disposiciones generales para la protección de datos personales. Diario Oficial No. 48.587.",
        "Congreso de la República de Colombia. (2014). Ley 1712 de 2014: Por medio de la cual se crea la Ley de Transparencia y del Derecho de Acceso a la Información Pública Nacional. Diario Oficial No. 49.084.",
        "DAMA International. (2017). DAMA-DMBOK: Data management body of knowledge (2nd ed.). Technics Publications.",
        "Databricks. (2024). What is the medallion lakehouse architecture? Databricks Documentation. https://www.databricks.com/glossary/medallion-architecture",
        "Departamento Administrativo Nacional de Estadística [DANE]. (2025). Boletín técnico: Pobreza multidimensional en Colombia y Bogotá D.C., resultados año 2024. https://www.dane.gov.co/",
        "Departamento Nacional de Planeación [DNP]. (2018). Documento CONPES 3920: Política Nacional de Explotación de Datos (Big Data). Consejo Nacional de Política Económica y Social.",
        "Kleppmann, M. (2017). Designing data-intensive applications: The big ideas behind reliable, scalable, and maintainable systems. O'Reilly Media.",
        "SaluData, Observatorio de Salud de Bogotá. (2025). Pobreza monetaria y condiciones de vida por localidad en Bogotá D.C. Secretaría Distrital de Salud. https://saludata.saludcapital.gov.co/",
        "Secretaría de Educación del Distrito [SED]. (2025). Portal de Datos Abiertos de Bogotá D.C.: Conjuntos de datos del sector educación. https://datosabiertos.bogota.gov.co/",
    ]
    for ref in referencias:
        add_reference_apa(doc, ref)

    rutas_destino = [
        os.path.join(WORKSPACE_DOCS, "Documento Integrado del Proyecto.docx"),
        os.path.join(BASE_REPO, "entregables", "entrega final", "Documento Integrado del Proyecto.docx"),
        r"c:\Users\dsant\Desktop\U CATOLICA\Programas\Data\entregables\entrega final\Documento Integrado del Proyecto.docx",
    ]
    for r in rutas_destino:
        os.makedirs(os.path.dirname(r), exist_ok=True)
        doc.save(r)
        print("Documento guardado en:", r)


if __name__ == "__main__":
    construir_documento()
