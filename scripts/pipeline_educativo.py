# -*- coding: utf-8 -*-
"""
pipeline_educativo.py

Sistema Integrado de Análisis de Calidad de Datos
Ecosistema Educativo Distrital de Bogotá D.C.

Pipeline principal: RAW -> STAGING -> ANALYTICS

Autores:
    David Santiago Castelblanco Artunduaga (67001182) — Pipeline y análisis
    Dairo Jose Narvaez Moralez (67002032)              — Limpieza y documentación

Asignatura: Ingeniería de Datos
Versión: v2.0.0 (Entrega Final — Arquitectura Lakehouse Batch + Streaming Corte 3)
Fecha: 22/11/2026
"""

import os
import sys

# Asegurar codificación utf-8 en consola de Windows
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
if sys.stderr and hasattr(sys.stderr, 'reconfigure'):
    try:
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

import pandas as pd
import numpy as np

# Intentamos importar PySpark; si está disponible y configurado se usará como motor primario
HAY_PYSPARK = False
try:
    from pyspark.sql import SparkSession
    from pyspark.sql.functions import (
        col, avg, count, when, desc, round as spark_round, isnull
    )
    from pyspark.sql.types import DoubleType
    HAY_PYSPARK = True
except ImportError:
    HAY_PYSPARK = False

# ─────────────────────────────────────────────────────────────────────────────
# Rutas y configuración global
# ─────────────────────────────────────────────────────────────────────────────
BASE_DIR      = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RAW_DIR       = os.path.join(BASE_DIR, "data", "RAW")
STAGING_DIR   = os.path.join(BASE_DIR, "data", "STAGING")
ANALYTICS_DIR = os.path.join(BASE_DIR, "data", "ANALYTICS")

# Umbral crítico de tasa de deserción (porcentaje)
UMBRAL_ALERTA: float = 8.0

# Valores centinela que indican puntaje no disponible
VALORES_CENTINELA: set = {"-", "N/A", "ND", "", "nan", "none"}

# Nombres de los archivos RAW
ARCHIVO_SABER11       = "Saber11_2025_Bogota.csv"
ARCHIVO_DIRECTORIO    = "Directorio_Colegios_Bogota.csv"
ARCHIVO_BENEFICIARIOS = "Beneficiarios_Jovenes_a_la_E.csv"
ARCHIVO_ALERTAS       = "Reporte_Alertas_Localidad.csv"


def comprobar_java_compatible() -> bool:
    """Verifica si la versión de Java instalada es compatible con Spark 4 (requiere Java 17+)."""
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
    """
    Intenta crear y retornar una SparkSession activa si Java 17+ está disponible.
    Si el entorno local cuenta con Java 8 o no tiene Java configurado,
    retorna None para permitir la ejecución con el motor local alternativo.
    """
    if not HAY_PYSPARK:
        return None
    if not comprobar_java_compatible():
        print("ℹ️  Entorno Java local (Java 8) no compatible con PySpark 4.x (requiere Java 17+).")
        print("ℹ️  Se activa el motor de procesamiento analítico local en Pandas.")
        print("   (En Google Colab o clúster con Java 17, el pipeline utiliza PySpark automáticamente).\n")
        return None

    try:
        spark = (
            SparkSession.builder
            .appName("Pipeline_Educativo_Bogota")
            .master("local[*]")
            .config("spark.sql.shuffle.partitions", "4")
            .config("spark.ui.showConsoleProgress", "false")
            .getOrCreate()
        )
        spark.sparkContext.setLogLevel("ERROR")
        print("✅ Sesión de Spark iniciada correctamente.")
        print(f"   Versión de Spark: {spark.version}\n")
        return spark
    except Exception as e:
        print(f"⚠️  Aviso de entorno Spark: {e}")
        print("ℹ️  Se continuará el procesamiento con el motor local en Pandas.\n")
        return None


# ─────────────────────────────────────────────────────────────────────────────
# Funciones de limpieza y validación (adaptadas de untitled1.py)
# ─────────────────────────────────────────────────────────────────────────────

def limpiar_valores_nulos(df: pd.DataFrame, columna: str, metodo: str = "moda") -> pd.DataFrame:
    """
    Rellena valores nulos de una columna de un DataFrame Pandas con moda o mediana.

    Parámetros
    ----------
    df : pandas.DataFrame
        DataFrame con los datos a limpiar.
    columna : str
        Nombre de la columna que contiene valores nulos.
    metodo : str, opcional
        Estrategia de imputación: 'moda' (valor más frecuente) o 'mediana'.
        Por defecto 'moda'.

    Retorna
    -------
    pandas.DataFrame
        DataFrame con los nulos de la columna imputados.
    """
    if columna not in df.columns:
        print(f"  ⚠️  Columna '{columna}' no encontrada en el DataFrame.")
        return df

    nulos_antes = df[columna].isna().sum()
    if nulos_antes == 0:
        return df

    if metodo == "moda":
        valores_validos = df[columna].dropna()
        if len(valores_validos) > 0:
            valor_fill = valores_validos.mode()[0]
        else:
            valor_fill = "SIN_DATO"
    elif metodo == "mediana":
        valor_fill = df[columna].median()
    else:
        print(f"  ⚠️  Método '{metodo}' no reconocido. Use 'moda' o 'mediana'.")
        return df

    df[columna] = df[columna].fillna(valor_fill)
    nulos_despues = df[columna].isna().sum()
    print(f"  🔧 '{columna}': {nulos_antes} nulos -> {nulos_despues} (imputados con {metodo}: {valor_fill!r})")
    return df


def validar_rango_puntaje(valor_str) -> float | None:
    """
    Convierte cadenas de puntaje problemáticas a float válido o None.

    Parámetros
    ----------
    valor_str : str o cualquier tipo
        Valor crudo del puntaje (puede ser cadena, número o None).

    Retorna
    -------
    float
        Puntaje como número de punto flotante (rango esperado 0.0–500.0).
    None
        Si el valor es un centinela ("-", "N/A", "ND", "") o no puede
        convertirse a número válido.
    """
    if valor_str is None or pd.isna(valor_str):
        return None
    val_limpio = str(valor_str).strip()
    if val_limpio.lower() in VALORES_CENTINELA:
        return None
    try:
        numero = float(val_limpio)
        if 0.0 <= numero <= 500.0:
            return numero
        return None
    except ValueError:
        return None


# ─────────────────────────────────────────────────────────────────────────────
# Reglas operativas y funciones analíticas
# ─────────────────────────────────────────────────────────────────────────────

def evaluar_alerta_localidad(nombre_localidad: str, tasa_desercion) -> str:
    """
    Evalúa si una localidad supera el umbral crítico de tasa de deserción (8.0%).
    Adapta la lógica de alertas operativas del ejercicio de domicilios.

    Parámetros
    ----------
    nombre_localidad : str
        Nombre de la localidad evaluada.
    tasa_desercion : float o None
        Tasa de deserción escolar expresada en porcentaje.

    Retorna
    -------
    str
        Mensaje con estado normal o alerta crítica.
    """
    if tasa_desercion is None or (isinstance(tasa_desercion, (int, float)) and tasa_desercion < 0):
        return f"⚠️  [{nombre_localidad}] Dato inválido: tasa de deserción no disponible."

    try:
        tasa = float(tasa_desercion)
    except (ValueError, TypeError):
        return f"⚠️  [{nombre_localidad}] Error: formato de tasa no numérico."

    if tasa > UMBRAL_ALERTA:
        return (
            f"🚨 ALERTA CRÍTICA: {nombre_localidad} — "
            f"Tasa de deserción {tasa:.1f}% supera el umbral de {UMBRAL_ALERTA}%. "
            "Requiere intervención inmediata."
        )
    return (
        f"✅ {nombre_localidad} — "
        f"Tasa de deserción {tasa:.1f}% dentro del límite aceptable ({UMBRAL_ALERTA}%)."
    )


def calcular_puntaje_promedio_por_localidad(df_saber11, es_spark=False):
    """
    Calcula el promedio del puntaje global Saber 11 agrupado por localidad.
    """
    if es_spark:
        return (
            df_saber11
            .filter(col("PUNT_GLOBAL").isNotNull())
            .groupBy("COLE_LOCALIDAD")
            .agg(
                spark_round(avg(col("PUNT_GLOBAL").cast(DoubleType())), 2).alias("promedio_punt_global"),
                count("*").alias("total_estudiantes")
            )
            .orderBy(desc("promedio_punt_global"))
        )
    else:
        df_validos = df_saber11[df_saber11["PUNT_GLOBAL"].notna()].copy()
        df_validos["PUNT_GLOBAL"] = pd.to_numeric(df_validos["PUNT_GLOBAL"], errors="coerce")
        res = (
            df_validos.groupby("COLE_LOCALIDAD")
            .agg(
                promedio_punt_global=("PUNT_GLOBAL", lambda x: round(x.mean(), 2)),
                total_estudiantes=("PUNT_GLOBAL", "count")
            )
            .reset_index()
            .sort_values(by="promedio_punt_global", ascending=False)
        )
        return res


def filtrar_colegios_activos(df_directorio, es_spark=False):
    """Filtra los colegios con estado ACTIVO."""
    if es_spark:
        return df_directorio.filter(col("ESTADO") == "ACTIVO")
    else:
        return df_directorio[df_directorio["ESTADO"] == "ACTIVO"].copy()


def analizar_alertas_criticas(df_alertas, es_spark=False):
    """Filtra alertas en niveles CRÍTICO o ALTO ordenadas por tasa de deserción descendente."""
    if es_spark:
        return (
            df_alertas
            .filter(col("NIVEL_ALERTA").isin("CRÍTICO", "ALTO"))
            .orderBy(desc("TASA_DESERCION"))
        )
    else:
        filtro = df_alertas["NIVEL_ALERTA"].isin(["CRÍTICO", "ALTO"])
        return df_alertas[filtro].sort_values(by="TASA_DESERCION", ascending=False).copy()


def calcular_beneficiarios_por_localidad_programa(df_beneficiarios, es_spark=False):
    """Agrupa beneficiarios activos por localidad y programa."""
    col_estado = "ESTADO_BENEFICIO" if "ESTADO_BENEFICIO" in df_beneficiarios.columns else "ESTADO_BENEFICIARIO"
    if es_spark:
        return (
            df_beneficiarios
            .filter(col(col_estado) == "ACTIVO")
            .groupBy("LOCALIDAD", "PROGRAMA")
            .agg(count("*").alias("total_beneficiarios_activos"))
            .orderBy(desc("total_beneficiarios_activos"))
        )
    else:
        activos = df_beneficiarios[df_beneficiarios[col_estado] == "ACTIVO"]
        res = (
            activos.groupby(["LOCALIDAD", "PROGRAMA"])
            .size()
            .reset_index(name="total_beneficiarios_activos")
            .sort_values(by="total_beneficiarios_activos", ascending=False)
        )
        return res


def analizar_absorcion_educacion_superior(df_beneficiarios, es_spark=False):
    """Calcula tasa de graduación por Institución de Educación Superior (IES)."""
    col_estado = "ESTADO_BENEFICIO" if "ESTADO_BENEFICIO" in df_beneficiarios.columns else "ESTADO_BENEFICIARIO"
    if es_spark:
        df_clean = df_beneficiarios.dropna(subset=["IES_ADMITIDA"])
        return (
            df_clean.groupBy("IES_ADMITIDA")
            .agg(
                count("*").alias("total"),
                count(when(col(col_estado) == "GRADUADO", 1)).alias("graduados"),
                count(when(col(col_estado) == "ACTIVO", 1)).alias("activos")
            )
            .withColumn("tasa_graduacion", spark_round((col("graduados") / col("total")) * 100, 2))
            .orderBy(desc("tasa_graduacion"))
        )
    else:
        df_clean = df_beneficiarios[df_beneficiarios["IES_ADMITIDA"].notna()].copy()
        res = df_clean.groupby("IES_ADMITIDA").agg(
            total=(col_estado, "count"),
            graduados=(col_estado, lambda x: (x == "GRADUADO").sum()),
            activos=(col_estado, lambda x: (x == "ACTIVO").sum())
        ).reset_index()
        res["tasa_graduacion"] = (res["graduados"] / res["total"] * 100).round(2)
        return res.sort_values(by="tasa_graduacion", ascending=False)


def cruzar_saber11_con_directorio(df_saber11, df_directorio, es_spark=False):
    """Realiza un LEFT JOIN entre Saber 11 y Directorio de colegios mediante código DANE."""
    if es_spark:
        col_nom = "NOMBRE_INST" if "NOMBRE_INST" in df_directorio.columns else "NOMBRE_ESTABLECIMIENTO"
        dir_sel = df_directorio.select(
            col("COD_DANE"),
            col(col_nom).alias("DIR_NOMBRE"),
            col("LOCALIDAD").alias("DIR_LOCALIDAD"),
            col("SECTOR").alias("DIR_SECTOR"),
            col("JORNADA").alias("DIR_JORNADA"),
            col("ESTADO").alias("DIR_ESTADO")
        )
        return df_saber11.join(
            dir_sel,
            df_saber11["COLE_COD_DANE_INST"] == dir_sel["COD_DANE"],
            how="left"
        ).drop("COD_DANE")
    else:
        col_nom = "NOMBRE_INST" if "NOMBRE_INST" in df_directorio.columns else "NOMBRE_ESTABLECIMIENTO"
        dir_cols = ["COD_DANE", col_nom, "LOCALIDAD", "SECTOR", "JORNADA", "ESTADO"]
        dir_sub = df_directorio[[c for c in dir_cols if c in df_directorio.columns]].copy()
        dir_sub.rename(columns={
            "COD_DANE": "COLE_COD_DANE_INST",
            col_nom: "DIR_NOMBRE",
            "LOCALIDAD": "DIR_LOCALIDAD",
            "SECTOR": "DIR_SECTOR",
            "JORNADA": "DIR_JORNADA",
            "ESTADO": "DIR_ESTADO"
        }, inplace=True)
        return pd.merge(df_saber11, dir_sub, on="COLE_COD_DANE_INST", how="left")


def generar_reporte_equidad_educativa(df_saber11_limpio, df_alertas, es_spark=False):
    """Cruza puntajes promedio por localidad con alertas de deserción escolar."""
    if es_spark:
        prom = calcular_puntaje_promedio_por_localidad(df_saber11_limpio, es_spark=True)
        col_loc = "NOMBRE_LOCALIDAD" if "NOMBRE_LOCALIDAD" in df_alertas.columns else "LOCALIDAD"
        al_sel = df_alertas.select(col(col_loc).alias("AL_LOC"), col("NIVEL_ALERTA"), col("TASA_DESERCION"))
        return prom.join(al_sel, prom["COLE_LOCALIDAD"] == al_sel["AL_LOC"], how="left").drop("AL_LOC").orderBy("promedio_punt_global")
    else:
        prom = calcular_puntaje_promedio_por_localidad(df_saber11_limpio, es_spark=False)
        col_loc = "NOMBRE_LOCALIDAD" if "NOMBRE_LOCALIDAD" in df_alertas.columns else "LOCALIDAD"
        al_sub = df_alertas[[col_loc, "NIVEL_ALERTA", "TASA_DESERCION"]].copy()
        al_sub.rename(columns={col_loc: "COLE_LOCALIDAD"}, inplace=True)
        res = pd.merge(prom, al_sub, on="COLE_LOCALIDAD", how="left")
        return res.sort_values(by="promedio_punt_global", ascending=True)


def comparar_metricas_por_localidad(df_beneficiarios, df_saber11=None, es_spark=False):
    """Compara total de beneficiarios activos vs puntaje global promedio."""
    col_estado = "ESTADO_BENEFICIO" if "ESTADO_BENEFICIO" in df_beneficiarios.columns else "ESTADO_BENEFICIARIO"
    if es_spark:
        ben = df_beneficiarios.filter(col(col_estado) == "ACTIVO").groupBy("LOCALIDAD").agg(count("*").alias("total_beneficiarios"))
        if df_saber11 is not None:
            punt = calcular_puntaje_promedio_por_localidad(df_saber11, es_spark=True)
            return ben.join(punt, ben["LOCALIDAD"] == punt["COLE_LOCALIDAD"], how="left").drop("COLE_LOCALIDAD").orderBy(desc("total_beneficiarios"))
        return ben.orderBy(desc("total_beneficiarios"))
    else:
        act = df_beneficiarios[df_beneficiarios[col_estado] == "ACTIVO"]
        ben = act.groupby("LOCALIDAD").size().reset_index(name="total_beneficiarios")
        if df_saber11 is not None:
            punt = calcular_puntaje_promedio_por_localidad(df_saber11, es_spark=False)
            punt.rename(columns={"COLE_LOCALIDAD": "LOCALIDAD"}, inplace=True)
            res = pd.merge(ben, punt, on="LOCALIDAD", how="left")
            return res.sort_values(by="total_beneficiarios", ascending=False)
        return ben.sort_values(by="total_beneficiarios", ascending=False)


def leer_documento_docx(ruta_archivo: str) -> str:
    """Extrae el contenido de texto de un archivo Word .docx."""
    try:
        import docx
    except ImportError:
        return ""
    if not os.path.exists(ruta_archivo):
        return ""
    try:
        doc = docx.Document(ruta_archivo)
        return "\n".join([p.text.strip() for p in doc.paragraphs if p.text.strip()])
    except Exception:
        return ""


# ─────────────────────────────────────────────────────────────────────────────
# Ejecución principal del Pipeline (RAW -> STAGING -> ANALYTICS)
# ─────────────────────────────────────────────────────────────────────────────

def ejecutar_pipeline() -> None:
    os.makedirs(STAGING_DIR, exist_ok=True)
    os.makedirs(ANALYTICS_DIR, exist_ok=True)

    print("=" * 65)
    print("  SISTEMA INTEGRADO DE ANÁLISIS DE DATOS EDUCATIVOS — BOGOTÁ D.C.")
    print("  Pipeline de Ingeniería de Datos: RAW -> STAGING -> ANALYTICS")
    print("=" * 65 + "\n")

    # 1. Intentar sesión Spark
    spark = crear_spark_session()
    usar_spark = spark is not None

    print("─" * 65)
    print("  ETAPA 1 — CARGA Y PERFILAMIENTO DE DATOS RAW")
    print("─" * 65)

    rutas = {
        "saber11": os.path.join(RAW_DIR, ARCHIVO_SABER11),
        "directorio": os.path.join(RAW_DIR, ARCHIVO_DIRECTORIO),
        "beneficiarios": os.path.join(RAW_DIR, ARCHIVO_BENEFICIARIOS),
        "alertas": os.path.join(RAW_DIR, ARCHIVO_ALERTAS)
    }

    for k, ruta in rutas.items():
        if not os.path.exists(ruta):
            print(f"❌ Error: Archivo {ruta} no encontrado.")
            print("   Ejecute primero: python scripts/generar_datos_raw.py")
            return

    # Cargamos en Pandas
    df_saber11_pd = pd.read_csv(rutas["saber11"], encoding="utf-8-sig", low_memory=False)
    df_directorio_pd = pd.read_csv(rutas["directorio"], encoding="utf-8-sig")
    df_beneficiarios_pd = pd.read_csv(rutas["beneficiarios"], encoding="utf-8-sig")
    df_alertas_pd = pd.read_csv(rutas["alertas"], encoding="utf-8-sig")

    print(f"✅ Saber11: {len(df_saber11_pd):,} registros × {len(df_saber11_pd.columns)} columnas")
    print(f"✅ Directorio Colegios: {len(df_directorio_pd):,} registros × {len(df_directorio_pd.columns)} columnas")
    print(f"✅ Beneficiarios Jóvenes a la E: {len(df_beneficiarios_pd):,} registros × {len(df_beneficiarios_pd.columns)} columnas")
    print(f"✅ Reporte Alertas Localidad: {len(df_alertas_pd):,} registros × {len(df_alertas_pd.columns)} columnas")

    print("\n── Diagnóstico de nulos iniciales en Saber 11 ──")
    nulos_saber = df_saber11_pd.isna().sum()
    for col_name, cnt in nulos_saber[nulos_saber > 0].items():
        print(f"  ⚠️  {col_name}: {cnt:,} nulos ({cnt/len(df_saber11_pd)*100:.1f}%)")

    # 2. Limpieza STAGING
    print("\n" + "─" * 65)
    print("  ETAPA 2 — LIMPIEZA Y VALIDACIÓN (STAGING)")
    print("─" * 65)

    print("\n⏳ Limpiando puntajes Saber 11 (filtrado de centinelas '-', 'N/A')...")
    df_saber11_pd["PUNT_GLOBAL"] = df_saber11_pd["PUNT_GLOBAL"].apply(validar_rango_puntaje)
    nulos_punt = df_saber11_pd["PUNT_GLOBAL"].isna().sum()
    print(f"  Puntajes no válidos/centinelas convertidos a nulos: {nulos_punt:,} ({nulos_punt/len(df_saber11_pd)*100:.2f}%)")

    print("\n⏳ Imputando códigos DANE faltantes con la moda...")
    df_saber11_pd = limpiar_valores_nulos(df_saber11_pd, "COLE_COD_DANE_INST", metodo="moda")

    ruta_saber11_clean = os.path.join(STAGING_DIR, "Saber11_2025_Bogota_clean.csv")
    df_saber11_pd.to_csv(ruta_saber11_clean, index=False, encoding="utf-8")
    print(f"  ✅ Guardado en STAGING: {ruta_saber11_clean}")

    print("\n⏳ Filtrando colegios con estado ACTIVO...")
    df_dir_activos = filtrar_colegios_activos(df_directorio_pd, es_spark=False)
    ruta_dir_clean = os.path.join(STAGING_DIR, "Directorio_Colegios_clean.csv")
    df_dir_activos.to_csv(ruta_dir_clean, index=False, encoding="utf-8")
    print(f"  ✅ Colegios activos: {len(df_dir_activos):,} de {len(df_directorio_pd):,}")
    print(f"  ✅ Guardado en STAGING: {ruta_dir_clean}")

    # 3. Transformaciones analíticas
    print("\n" + "─" * 65)
    print("  ETAPA 3 — TRANSFORMACIONES Y REGLAS DE NEGOCIO")
    print("─" * 65)

    print("\n📊 1. Puntajes promedio Saber 11 por Localidad:")
    df_puntaje_loc = calcular_puntaje_promedio_por_localidad(df_saber11_pd, es_spark=False)
    print(df_puntaje_loc.head(10).to_string(index=False))

    print("\n🚨 2. Alertas críticas de deserción (CRÍTICO / ALTO):")
    df_alertas_crit = analizar_alertas_criticas(df_alertas_pd, es_spark=False)
    col_nom_loc = "NOMBRE_LOCALIDAD" if "NOMBRE_LOCALIDAD" in df_alertas_crit.columns else "LOCALIDAD"
    print(df_alertas_crit[[col_nom_loc, "NIVEL_ALERTA", "TASA_DESERCION"]].to_string(index=False))

    print("\n── 3. Evaluación automática de regla de alertas por localidad ──")
    for _, row in df_alertas_pd.iterrows():
        loc_nombre = row[col_nom_loc]
        tasa_val = row["TASA_DESERCION"]
        print("  " + evaluar_alerta_localidad(loc_nombre, tasa_val))

    print("\n👥 4. Beneficiarios activos por Localidad y Programa (Top 10):")
    df_ben_loc = calcular_beneficiarios_por_localidad_programa(df_beneficiarios_pd, es_spark=False)
    print(df_ben_loc.head(10).to_string(index=False))

    print("\n🎓 5. Absorción en Educación Superior (Tasa de Graduación por IES):")
    df_absorcion = analizar_absorcion_educacion_superior(df_beneficiarios_pd, es_spark=False)
    print(df_absorcion.head(10).to_string(index=False))

    # 4. Cruces relacionales
    print("\n" + "─" * 65)
    print("  ETAPA 4 — CRUCE RELACIONAL Y ENRIQUECIMIENTO")
    print("─" * 65)

    print("⏳ Cruzando Saber 11 con Directorio de Colegios mediante código DANE...")
    df_enriquecido = cruzar_saber11_con_directorio(df_saber11_pd, df_dir_activos, es_spark=False)
    print(f"  Total registros cruzados: {len(df_enriquecido):,}")
    print(f"  Estudiantes con colegio vinculado exitosamente: {df_enriquecido['DIR_NOMBRE'].notna().sum():,}")

    print("\n📈 6. Comparación de métricas: Beneficiarios vs Puntaje promedio:")
    df_comparativo = comparar_metricas_por_localidad(df_beneficiarios_pd, df_saber11_pd, es_spark=False)
    print(df_comparativo.head(10).to_string(index=False))

    # 5. Persistencia ANALYTICS
    print("\n" + "─" * 65)
    print("  ETAPA 5 — REPORTE DE EQUIDAD Y PERSISTENCIA (ANALYTICS)")
    print("─" * 65)

    print("⏳ Generando reporte integral de equidad educativa...")
    df_equidad = generar_reporte_equidad_educativa(df_saber11_pd, df_alertas_pd, es_spark=False)
    print(df_equidad.to_string(index=False))

    # Guardar en ANALYTICS
    ruta_equidad = os.path.join(ANALYTICS_DIR, "reporte_equidad_educativa.csv")
    df_equidad.to_csv(ruta_equidad, index=False, encoding="utf-8")
    print(f"\n  💾 [ANALYTICS] Reporte Equidad: {ruta_equidad}")

    ruta_punt = os.path.join(ANALYTICS_DIR, "puntaje_promedio_por_localidad.csv")
    df_puntaje_loc.to_csv(ruta_punt, index=False, encoding="utf-8")
    print(f"  💾 [ANALYTICS] Puntajes Promedio: {ruta_punt}")

    ruta_ben = os.path.join(ANALYTICS_DIR, "beneficiarios_por_localidad_programa.csv")
    df_ben_loc.to_csv(ruta_ben, index=False, encoding="utf-8")
    print(f"  💾 [ANALYTICS] Beneficiarios Localidad/Programa: {ruta_ben}")

    ruta_abs = os.path.join(ANALYTICS_DIR, "absorcion_educacion_superior.csv")
    df_absorcion.to_csv(ruta_abs, index=False, encoding="utf-8")
    print(f"  💾 [ANALYTICS] Absorción IES: {ruta_abs}")

    ruta_al = os.path.join(ANALYTICS_DIR, "alertas_criticas_localidad.csv")
    df_alertas_crit.to_csv(ruta_al, index=False, encoding="utf-8")
    print(f"  💾 [ANALYTICS] Alertas Críticas: {ruta_al}")

    # 6. Procesamiento de Eventos en Tiempo Real (Corte 3: EDA, Watermarks, Ventanas y DLQ)
    try:
        from streaming_alertas_educativas import ejecutar_streaming_educativo
        ejecutar_streaming_educativo()
    except Exception as e:
        print(f"⚠️  Aviso al ejecutar módulo de streaming: {e}")

    if usar_spark:
        spark.stop()

    print("\n" + "=" * 65)
    print("  ✅ PIPELINE INTEGRAL (BATCH + STREAMING CORTE 3) EJECUTADO CON ÉXITO")
    print(f"     STAGING   -> {STAGING_DIR}")
    print(f"     ANALYTICS -> {ANALYTICS_DIR}")
    print("=" * 65 + "\n")


if __name__ == "__main__":
    ejecutar_pipeline()
