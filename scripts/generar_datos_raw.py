"""
Generador de Datos Sintéticos — Capa RAW
=========================================
Genera los 4 archivos CSV de la capa RAW con datos realistas
del ecosistema educativo de Bogotá D.C., incluyendo los
problemas de calidad descritos en el perfilamiento inicial.

Archivos generados:
  - data/RAW/Saber11_2025_Bogota.csv        (~70,000 filas, 35 cols)
  - data/RAW/Directorio_Colegios_Bogota.csv (~2,500  filas, 15 cols)
  - data/RAW/Beneficiarios_Jovenes_a_la_E.csv (~15,000 filas, 10 cols)
  - data/RAW/Reporte_Alertas_Localidad.csv  (~20     filas,  8 cols)

Problemas de calidad INTENCIONALES (perfilamiento):
  - Puntajes Saber11 con "-" o "N/A" (números leídos como texto)
  - ~4.2% de registros sin COLE_COD_DANE_INST
  - Valores nulos dispersos por columnas de caracterización
"""

import os
import random
import string
import pandas as pd
import numpy as np
from datetime import datetime, timedelta

# ─── Semilla para reproducibilidad ────────────────────────────────────────────
random.seed(42)
np.random.seed(42)

# ─── Directorios ──────────────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RAW_DIR  = os.path.join(BASE_DIR, "data", "RAW")
os.makedirs(RAW_DIR, exist_ok=True)

# ─── Catálogos maestros ───────────────────────────────────────────────────────
LOCALIDADES = [
    (1,  "USAQUÉN"),     (2,  "CHAPINERO"),    (3,  "SANTA FE"),
    (4,  "SAN CRISTÓBAL"),(5, "USME"),         (6,  "TUNJUELITO"),
    (7,  "BOSA"),        (8,  "KENNEDY"),       (9,  "FONTIBÓN"),
    (10, "ENGATIVÁ"),    (11, "SUBA"),          (12, "BARRIOS UNIDOS"),
    (13, "TEUSAQUILLO"), (14, "LOS MÁRTIRES"),  (15, "ANTONIO NARIÑO"),
    (16, "PUENTE ARANDA"),(17,"LA CANDELARIA"), (18, "RAFAEL URIBE URIBE"),
    (19, "CIUDAD BOLÍVAR"),(20,"SUMAPAZ"),
]
LOC_DICT = {cod: nombre for cod, nombre in LOCALIDADES}

JORNADAS    = ["MAÑANA", "TARDE", "NOCHE", "ÚNICA", "FIN DE SEMANA"]
NATURALEZAS = ["OFICIAL", "NO OFICIAL"]
SECTORES    = ["OFICIAL", "PRIVADO"]
CALENDARIOS = ["A", "B"]
AREAS       = ["URBANA", "RURAL"]
GENEROS     = ["M", "F"]
DOCUMENTOS  = ["CC", "TI", "CE"]
NIVELES_SIS = ["1", "2", "3", "4", "Sin Sisben", np.nan]
ESTRATOS    = ["1", "2", "3", "4", "5", "6", np.nan]
DESEMPEÑO   = ["A1", "A2", "B1", "B2", "C1", "C2"]
PROGRAMAS   = ["Jóvenes a la E", "La U en tu Colegio"]
IES_LIST    = [
    "Universidad Nacional de Colombia", "Universidad Distrital FJC",
    "Universidad de los Andes", "Pontificia Universidad Javeriana",
    "Universidad Externado de Colombia", "Universidad de la Salle",
    "UNIMINUTO", "Politécnico Grancolombiano", "SENA",
    "Universidad Santo Tomás",
]
NOMBRES = ["CARLOS","MARIA","JUAN","ANA","PEDRO","LAURA","LUIS","CAROLINA",
           "JORGE","ANDREA","ANDRÉS","DIANA","FELIPE","PAULA","SERGIO","CAMILA",
           "MIGUEL","VALENTINA","DAVID","NATALIA","ALEJANDRO","ISABELLA"]
APELLIDOS = ["GARCÍA","RODRÍGUEZ","MARTÍNEZ","LÓPEZ","SÁNCHEZ","PÉREZ","GÓMEZ",
             "HERNÁNDEZ","DÍAZ","TORRES","VARGAS","RAMÍREZ","CASTILLO","MORENO",
             "RAMOS","SILVA","FLORES","GUERRERO","REYES","MEDINA","MENDOZA"]

# ─── Helpers ──────────────────────────────────────────────────────────────────
def dane_code():
    """Código DANE institucional de 12 dígitos."""
    return f"11001{random.randint(10000000, 99999999)}"

def rand_date(start_year=1995, end_year=2007):
    start = datetime(start_year, 1, 1)
    end   = datetime(end_year,  12, 31)
    return (start + timedelta(days=random.randint(0, (end - start).days))).strftime("%Y-%m-%d")

def puntaje_raw(p_null=0.042):
    """Puntaje entre 0-500, con ~4.2% de valores problemáticos."""
    if random.random() < p_null / 2:
        return random.choice(["-", "N/A", "", "ND"])
    return str(random.randint(0, 500))

def nombre_colegio(loc_nombre):
    prefijos = ["COLEGIO", "IED", "INSTITUCIÓN EDUCATIVA", "LICEO", "GIMNASIO"]
    temas    = ["DISTRITAL", "NACIONAL", "MODERNO", "TÉCNICO", "CIUDADELA",
                "INTEGRAL", "NUEVO HORIZONTE", "COLOMBIA", "LIBERTAD", "PAZ"]
    return f"{random.choice(prefijos)} {loc_nombre.split()[0]} {random.choice(temas)}"

# ══════════════════════════════════════════════════════════════════════════════
# 1. Directorio_Colegios_Bogota.csv  (~2,500 filas, 15 columnas)
# ══════════════════════════════════════════════════════════════════════════════
print("⏳ Generando Directorio_Colegios_Bogota.csv ...")

N_COLEGIOS = 2500
colegios_dane = [dane_code() for _ in range(N_COLEGIOS)]
loc_asignadas = [random.choice(LOCALIDADES) for _ in range(N_COLEGIOS)]

directorio = pd.DataFrame({
    "COD_DANE": colegios_dane,
    "NOMBRE_INST": [nombre_colegio(loc[1]) for loc in loc_asignadas],
    "SECTOR": [random.choice(SECTORES) for _ in range(N_COLEGIOS)],
    "NATURALEZA": [random.choice(NATURALEZAS) for _ in range(N_COLEGIOS)],
    "JORNADA": [random.choice(JORNADAS) for _ in range(N_COLEGIOS)],
    "LOCALIDAD": [loc[1] for loc in loc_asignadas],
    "COD_LOCALIDAD": [loc[0] for loc in loc_asignadas],
    "DIRECCION": [f"CL {random.randint(1,180)} #{random.randint(1,120)}-{random.randint(1,99)}" for _ in range(N_COLEGIOS)],
    "TELEFONO": [f"601{random.randint(1000000,9999999)}" for _ in range(N_COLEGIOS)],
    "EMAIL": [f"rectoria{random.randint(100,9999)}@educacionbogota.edu.co" for _ in range(N_COLEGIOS)],
    "RECTOR": [f"{random.choice(NOMBRES)} {random.choice(APELLIDOS)} {random.choice(APELLIDOS)}" for _ in range(N_COLEGIOS)],
    "NIVELES_EDUCATIVOS": [random.choice(["PREESCOLAR,BÁSICA,MEDIA", "BÁSICA,MEDIA", "MEDIA", "BÁSICA"]) for _ in range(N_COLEGIOS)],
    "CALENDARIO": [random.choice(CALENDARIOS) for _ in range(N_COLEGIOS)],
    "AREA_UBICACION": [random.choice(AREAS) for _ in range(N_COLEGIOS)],
    "ESTADO": [random.choices(["ACTIVO","INACTIVO","EN PROCESO"], weights=[90,5,5])[0] for _ in range(N_COLEGIOS)],
    "FECHA_REGISTRO": [rand_date(2000, 2024) for _ in range(N_COLEGIOS)],
})

ruta_colegios = os.path.join(RAW_DIR, "Directorio_Colegios_Bogota.csv")
directorio.to_csv(ruta_colegios, index=False, encoding="utf-8-sig")
print(f"   ✅ {ruta_colegios}  ({len(directorio):,} filas × {len(directorio.columns)} cols)")

# ══════════════════════════════════════════════════════════════════════════════
# 2. Saber11_2025_Bogota.csv  (~70,000 filas, 35 columnas)
# ══════════════════════════════════════════════════════════════════════════════
print("⏳ Generando Saber11_2025_Bogota.csv (puede tardar ~30 seg) ...")

N_SABER = 70000
DANE_NULL_RATE = 0.042  # 4.2% sin código DANE

# Asignar colegios DANE con posibilidad de nulo
def get_dane():
    if random.random() < DANE_NULL_RATE:
        return np.nan
    return random.choice(colegios_dane)

loc_estud = [random.choice(LOCALIDADES) for _ in range(N_SABER)]

saber11 = pd.DataFrame({
    "ESTU_CONSECUTIVO":       range(100000, 100000 + N_SABER),
    "ESTU_TIPODOCUMENTO":     [random.choice(DOCUMENTOS) for _ in range(N_SABER)],
    "ESTU_DOCUMENTO":         [str(random.randint(1000000000, 9999999999)) for _ in range(N_SABER)],
    "ESTU_NOMBRES":           [random.choice(NOMBRES) + " " + random.choice(NOMBRES) for _ in range(N_SABER)],
    "ESTU_PRIMERAPE":         [random.choice(APELLIDOS) for _ in range(N_SABER)],
    "ESTU_SEGUNDOAPE":        [random.choice(APELLIDOS) for _ in range(N_SABER)],
    "ESTU_GENERO":            [random.choice(GENEROS) for _ in range(N_SABER)],
    "ESTU_FECHANACIMIENTO":   [rand_date(1995, 2007) for _ in range(N_SABER)],
    "ESTU_EDAD":              [random.randint(15, 22) for _ in range(N_SABER)],
    "ESTU_NACIONALIDAD":      [random.choices(["COLOMBIANO","VENEZOLANO","OTRO"], weights=[94,4,2])[0] for _ in range(N_SABER)],
    "ESTU_PAIS_RESIDE":       ["COLOMBIA"] * N_SABER,
    "ESTU_DEPTO_RESIDE":      ["BOGOTÁ D.C."] * N_SABER,
    "ESTU_MCPIO_RESIDE":      ["BOGOTÁ D.C."] * N_SABER,
    "ESTU_LOCALIDAD_RESIDE":  [loc[1] for loc in loc_estud],
    "COLE_CODIGO_ICFES":      [f"BOG{random.randint(10000,99999)}" for _ in range(N_SABER)],
    "COLE_COD_DANE_INST":     [get_dane() for _ in range(N_SABER)],   # 4.2% nulos
    "COLE_NOMBRE_INST":       [nombre_colegio(loc[1]) for loc in loc_estud],
    "COLE_DEPTO_UBICACION":   ["BOGOTÁ D.C."] * N_SABER,
    "COLE_MCPIO_UBICACION":   ["BOGOTÁ D.C."] * N_SABER,
    "COLE_AREA_UBICACION":    [random.choice(AREAS) for _ in range(N_SABER)],
    "COLE_JORNADA":           [random.choice(JORNADAS) for _ in range(N_SABER)],
    "COLE_NATURALEZA":        [random.choice(NATURALEZAS) for _ in range(N_SABER)],
    "COLE_SECTOR":            [random.choice(SECTORES) for _ in range(N_SABER)],
    "COLE_CALENDARIO":        [random.choice(CALENDARIOS) for _ in range(N_SABER)],
    "COLE_LOCALIDAD":         [loc[1] for loc in loc_estud],
    # Puntajes con problemas de calidad intencionales ("-", "N/A", etc.)
    "PUNT_INGLES":            [puntaje_raw() for _ in range(N_SABER)],
    "PUNT_MATEMATICAS":       [puntaje_raw() for _ in range(N_SABER)],
    "PUNT_SOCIALES_CIUDADANAS":[puntaje_raw() for _ in range(N_SABER)],
    "PUNT_CIENCIAS_NATURALES":[puntaje_raw() for _ in range(N_SABER)],
    "PUNT_LECTURA_CRITICA":   [puntaje_raw() for _ in range(N_SABER)],
    "PUNT_GLOBAL":            [puntaje_raw(p_null=0.03) for _ in range(N_SABER)],
    "PERIODO":                ["2025-1"] * N_SABER,
    "DESEMP_INGLES":          [random.choices(DESEMPEÑO + [np.nan], weights=[15,20,25,20,15,4,1])[0] for _ in range(N_SABER)],
    "ESTU_NIVEL_SISBEN":      [random.choice(NIVELES_SIS) for _ in range(N_SABER)],
    "ESTU_ESTRATO_VIVIENDA":  [random.choice(ESTRATOS) for _ in range(N_SABER)],
})

ruta_saber = os.path.join(RAW_DIR, "Saber11_2025_Bogota.csv")
saber11.to_csv(ruta_saber, index=False, encoding="utf-8-sig")
null_dane = saber11["COLE_COD_DANE_INST"].isna().sum()
print(f"   ✅ {ruta_saber}  ({len(saber11):,} filas × {len(saber11.columns)} cols)")
print(f"   ℹ️  COLE_COD_DANE_INST nulos: {null_dane:,} ({null_dane/N_SABER*100:.2f}%)")

# ══════════════════════════════════════════════════════════════════════════════
# 3. Beneficiarios_Jovenes_a_la_E.csv  (~15,000 filas, 10 columnas)
# ══════════════════════════════════════════════════════════════════════════════
print("⏳ Generando Beneficiarios_Jovenes_a_la_E.csv ...")

N_BENEF = 15000
loc_benef = [random.choice(LOCALIDADES) for _ in range(N_BENEF)]

beneficiarios = pd.DataFrame({
    "ID_BENEFICIARIO":   [f"JE-{str(i).zfill(6)}" for i in range(1, N_BENEF+1)],
    "DOCUMENTO":         [str(random.randint(1000000000, 9999999999)) for _ in range(N_BENEF)],
    "NOMBRE_COMPLETO":   [f"{random.choice(NOMBRES)} {random.choice(APELLIDOS)} {random.choice(APELLIDOS)}" for _ in range(N_BENEF)],
    "LOCALIDAD":         [loc[1] for loc in loc_benef],
    "COLEGIO_DANE":      [random.choice(colegios_dane) for _ in range(N_BENEF)],
    "PROGRAMA":          [random.choice(PROGRAMAS) for _ in range(N_BENEF)],
    "PERIODO":           [random.choice(["2023","2024","2025"]) for _ in range(N_BENEF)],
    "ESTADO_BENEFICIO":  [random.choices(["ACTIVO","GRADUADO","RETIRADO","EN PROCESO"], weights=[40,35,15,10])[0] for _ in range(N_BENEF)],
    "IES_ADMITIDA":      [random.choices(IES_LIST + [np.nan], weights=[8,8,8,8,8,8,8,8,8,8,10])[0] for _ in range(N_BENEF)],
    "FECHA_ADMISION":    [rand_date(2020, 2025) if random.random() > 0.15 else np.nan for _ in range(N_BENEF)],
})

ruta_benef = os.path.join(RAW_DIR, "Beneficiarios_Jovenes_a_la_E.csv")
beneficiarios.to_csv(ruta_benef, index=False, encoding="utf-8-sig")
print(f"   ✅ {ruta_benef}  ({len(beneficiarios):,} filas × {len(beneficiarios.columns)} cols)")

# ══════════════════════════════════════════════════════════════════════════════
# 4. Reporte_Alertas_Localidad.csv  (~20 filas, 8 columnas)
# ══════════════════════════════════════════════════════════════════════════════
print("⏳ Generando Reporte_Alertas_Localidad.csv ...")

alertas_data = []
for cod, nombre in LOCALIDADES:
    tasa_desercion = round(random.uniform(1.5, 12.8), 2)
    alertas_data.append({
        "COD_LOCALIDAD":    cod,
        "NOMBRE_LOCALIDAD": nombre,
        "PERIODO":          "2025-1",
        "NIVEL_ALERTA":     random.choice(["CRÍTICO","ALTO","MEDIO","BAJO"]),
        "NUM_DESERCION":    random.randint(50, 1800),
        "TASA_DESERCION":   tasa_desercion,
        "NUM_REPITENCIA":   random.randint(20, 950),
        "DESCRIPCION_ALERTA": random.choice([
            "Alta tasa de deserción en zona rural",
            "Incremento de repitencia en básica secundaria",
            "Alertas por situación socioeconómica crítica",
            "Riesgo de deserción por desplazamiento",
            "Condiciones estables, monitoreo preventivo",
            "Brecha significativa en acceso a educación superior",
        ]),
    })

alertas = pd.DataFrame(alertas_data)
ruta_alertas = os.path.join(RAW_DIR, "Reporte_Alertas_Localidad.csv")
alertas.to_csv(ruta_alertas, index=False, encoding="utf-8-sig")
print(f"   ✅ {ruta_alertas}  ({len(alertas):,} filas × {len(alertas.columns)} cols)")

# ══════════════════════════════════════════════════════════════════════════════
# Reporte de perfilamiento
# ══════════════════════════════════════════════════════════════════════════════
print("\n" + "═"*60)
print("📊 PERFILAMIENTO INICIAL — CAPA RAW")
print("═"*60)

archivos = [
    (ruta_saber,     "Saber11_2025_Bogota.csv"),
    (ruta_colegios,  "Directorio_Colegios_Bogota.csv"),
    (ruta_benef,     "Beneficiarios_Jovenes_a_la_E.csv"),
    (ruta_alertas,   "Reporte_Alertas_Localidad.csv"),
]

for ruta, nombre in archivos:
    df  = pd.read_csv(ruta, encoding="utf-8-sig")
    tam = os.path.getsize(ruta) / (1024*1024)
    print(f"\n📄 {nombre}")
    print(f"   Filas: {len(df):,}  |  Cols: {len(df.columns)}  |  Tamaño: {tam:.2f} MB")
    nulos = df.isna().sum()
    nulos_total = nulos.sum()
    print(f"   Nulos totales: {nulos_total:,}  ({nulos_total/(len(df)*len(df.columns))*100:.2f}%)")
    if nulos_total > 0:
        top_nulos = nulos[nulos > 0].sort_values(ascending=False).head(5)
        for col, cnt in top_nulos.items():
            print(f"   ⚠️  {col}: {cnt:,} nulos ({cnt/len(df)*100:.1f}%)")

print("\n" + "═"*60)
print("✅ Todos los archivos generados exitosamente en data/RAW/")
print("   Los archivos son INMUTABLES. Transformaciones → data/STAGING/")
print("═"*60)
