# Sistema Integrado de Análisis Analítico y Calidad de Datos para el Ecosistema Educativo Distrital de Bogotá D.C.

> **Sección 3.1 — Descripción del Proyecto Integrador**

---

## 1. Resumen Ejecutivo

El presente proyecto integrador tiene como propósito diseñar e implementar una **arquitectura de datos moderna por capas** (Bronze/RAW, Silver/Staging y Gold/Analytics) para la ingesta, limpieza, estandarización y análisis de la información del sector educativo en Bogotá D.C.

Utilizando como fuente primaria los conjuntos de **Datos Abiertos Bogotá** expedidos por la Secretaría de Educación del Distrito (SED), el sistema unifica los datos de:

- 📊 Desempeño académico (*Pruebas Saber 11 2025*)
- 🏫 Caracterización de instituciones educativas (sector oficial y no oficial)
- 🎓 Programas de acceso a la educación superior (*Jóvenes a la E*, *La U en tu Colegio*)
- ⚠️ Factores de vulnerabilidad social (*Sistema de Alertas por Localidad*)

---

## 2. Contexto de Negocio y Dominio

En Bogotá D.C., la toma de decisiones sobre **políticas públicas educativas** enfrenta retos significativos debido a la dispersión y heterogeneidad de las fuentes de información. A pesar de la disponibilidad de datos públicos, la falta de integración entre:

- las métricas de rendimiento en pruebas estandarizadas (Saber 11),
- el contexto institucional de los colegios (códigos DANE, jornadas, sectores), y
- las alertas comunitarias

...dificulta la identificación oportuna de brechas educativas.

Este proyecto resuelve dicha problemática construyendo un **canal de datos (data pipeline) gobernado, trazable y auditable** que centraliza el ecosistema educativo de las **20 localidades** de la ciudad.

---

## 3. Alcance Técnico y Funcional

### 🥉 Ingesta Inmutable — Capa RAW / Bronze

Automatización del almacenamiento de archivos crudos (`.csv`, `.json`) obtenidos mediante el portal de Datos Abiertos y la **API CKAN**, garantizando la conservación intacta de los datos de origen para efectos de auditoría y linaje.

### 🥈 Validación y Limpieza — Capa Staging / Silver

Aplicación estricta de un **Contrato de Datos** para validar formatos, tipos, rangos de puntajes (0–500 en Saber 11) e integridad en identificadores clave (código DANE de 12 dígitos) mediante librerías como **Pandas** y **PySpark**.

### 🥇 Modelado Dimensional — Capa Analytics / Gold

Construcción de un **Datamart Educativo** bajo un esquema en estrella compuesto por:

| Componente | Descripción |
|---|---|
| `Fact_Rendimiento_Saber11` | Tabla de hechos central |
| `Dim_Colegio` | Dimensión de institución educativa |
| `Dim_Localidad` | Dimensión geográfica |
| `Dim_Tiempo` | Dimensión temporal |
| `Dim_Programa_Social` | Dimensión de programas de acceso superior |

### 📈 Consumo de Información

Generación de insumos analíticos optimizados para alimentar **tableros de control (Power BI)**, reportes técnicos y modelos analíticos dirigidos a la SED, investigadores y directivos docentes.

---

## 4. Valor Aportado

El proyecto proporciona una **base reproducible y confiable** que permite evaluar cuantitativamente la equidad educativa en Bogotá. A través del cruce de datos, el sistema hace posible:

- Correlacionar el **riesgo social** o la **ubicación geográfica** con el éxito académico.
- Medir la **absorción en la educación superior** por localidad.
- Facilitar la formulación de **intervenciones pedagógicas e inversión pública focalizada**.

---

*Fuente de datos: [Datos Abiertos Bogotá](https://datosabiertos.bogota.gov.co) — Secretaría de Educación del Distrito (SED)*

---

## 5. Capa RAW / Bronze — Archivos e Inventario

Los archivos crudos se almacenan en `data/RAW/` y son **estrictamente inmutables**.

| # | Archivo | Filas | Columnas | Tamaño |
|---|---------|------:|:--------:|-------:|
| 1 | `Saber11_2025_Bogota.csv` | 70,000 | 35 | 18.88 MB |
| 2 | `Directorio_Colegios_Bogota.csv` | 2,500 | 16 | 0.49 MB |
| 3 | `Beneficiarios_Jovenes_a_la_E.csv` | 15,000 | 10 | 1.89 MB |
| 4 | `Reporte_Alertas_Localidad.csv` | 20 | 8 | < 0.01 MB |

---

## 6. Diagnóstico / Perfilamiento Inicial (Capa RAW)

### 6.1 Tipos inferidos
Los archivos CSV son leídos con tipos por defecto por Pandas/PySpark. Se detectaron las siguientes anomalías de tipo:

- **Puntajes Saber 11** (`PUNT_INGLES`, `PUNT_MATEMATICAS`, `PUNT_SOCIALES_CIUDADANAS`, `PUNT_CIENCIAS_NATURALES`, `PUNT_LECTURA_CRITICA`, `PUNT_GLOBAL`): columnas numéricas (rango 0–500) que Pandas infiere como `object` debido a la presencia de valores centinela como `"-"`, `"N/A"`, `""` y `"ND"`.
- **Fechas** (`ESTU_FECHANACIMIENTO`, `FECHA_ADMISION`): leídas como `string`; requieren conversión explícita con `pd.to_datetime()`.
- **Códigos** (`COLE_COD_DANE_INST`): campo numérico de 12 dígitos que puede ser interpretado como `float64` en presencia de nulos, perdiendo ceros iniciales.

### 6.2 Nulos e Inconsistencias detectadas

#### `Saber11_2025_Bogota.csv`

| Campo | Nulos | % del Total |
|-------|------:|:-----------:|
| `ESTU_NIVEL_SISBEN` | 11,784 | 16.8% |
| `ESTU_ESTRATO_VIVIENDA` | 9,975 | 14.2% |
| `COLE_COD_DANE_INST` | **2,934** | **4.2%** ← crítico |
| `PUNT_MATEMATICAS` | 736 | 1.1% |
| `PUNT_LECTURA_CRITICA` | 736 | 1.1% |
| `DESEMP_INGLES` | ~1,050 | ~1.5% |

> **⚠️ Hallazgo crítico**: El 4.2% de los registros de Saber 11 no tiene `COLE_COD_DANE_INST`, lo que impide su vinculación con el `Directorio_Colegios_Bogota.csv` en la capa Silver.

#### `Beneficiarios_Jovenes_a_la_E.csv`

| Campo | Nulos | % del Total |
|-------|------:|:-----------:|
| `FECHA_ADMISION` | 2,227 | 14.8% |
| `IES_ADMITIDA` | 1,693 | 11.3% |

#### `Directorio_Colegios_Bogota.csv` y `Reporte_Alertas_Localidad.csv`
- ✅ Sin valores nulos. Calidad de origen aceptable.

### 6.3 Regla de Oro — Inmutabilidad de la Capa RAW

> **Los archivos depositados en `data/RAW/` se mantienen estrictamente inmutables.**
> Ningún script sobrescribe, modifica ni elimina estos archivos.
> Todas las correcciones de calidad (imputación, casteo, estandarización) se realizan **en memoria** mediante los scripts de transformación ubicados en `scripts/` y el resultado se deposita en `data/STAGING/`.

```
data/
├── RAW/          ← Zona inmutable (origen auditado: 4 CSV Batch + 1 JSONL CloudEvents)
│   ├── Saber11_2025_Bogota.csv
│   ├── Directorio_Colegios_Bogota.csv
│   ├── Beneficiarios_Jovenes_a_la_E.csv
│   ├── Reporte_Alertas_Localidad.csv
│   └── eventos_asistencia_desercion_stream.jsonl
├── STAGING/      ← Zona limpia (destino de transformaciones e imputación Silver)
└── ANALYTICS/    ← Zona analítica (modelo dimensional estrella + reportes Streaming + figuras/)

scripts/
├── generar_datos_raw.py                ← Generador de datos sintéticos RAW (semilla 42)
├── pipeline_educativo.py               ← Orquestador principal (Etapas 1-5 Batch + Etapa 6 Streaming)
├── streaming_alertas_educativas.py     ← Motor EDA, Watermarks, Ventanas (Tumbling/Sliding/Session) y DLQ
├── generar_graficos_evidencias.py      ← Generador de figuras técnicas en formato APA 7
├── generar_documento_integrado_apa7.py ← Generador del Manual Técnico y de Usuario en normas APA 7
└── generar_presentacion_y_guion_final.py ← Generador de presentación PPTX/HTML y guion de sustentación

entregables/
└── entrega final/
    ├── Documento Integrado del Proyecto.docx
    ├── Documento Integrado del Proyecto.pdf
    ├── Presentacion_Sustentacion_Entrega_Final.pptx
    ├── Presentacion_Interactiva_Sustentacion.html
    └── Guion_Video_Sustentacion_Entrega_Final.docx
```

---

## 7. Extensión del Tercer Corte — Procesamiento de Eventos en Tiempo Real (EDA & Streaming)

En la versión final **v2.0.0** se implementó el módulo `scripts/streaming_alertas_educativas.py` (integrado automáticamente como **Etapa 6** de `scripts/pipeline_educativo.py`), el cual aplica los conceptos del Tercer Corte al monitoreo de alertas de inasistencia y riesgo de deserción escolar:

- **Arquitectura Orientada a Eventos (EDA) y CloudEvents 1.0**: Productor de eventos escolares JSONL, `BrokerEventosEducativos` particionado por localidad y consumidor **idempotente** que descarta en $O(1)$ reintentos duplicados de red (`EVT-004`).
- **Event Time vs. Processing Time y Watermarks**: Procesamiento gobernado por el tiempo real del evento en el colegio (`event_time`) con marca de agua dinámica $\text{WM} = \max(\text{event\_time}) - 5\text{ min}$, demostrando empíricamente la distorsión que produciría agrupar por `processing_time`.
- **Ventanas de Tiempo**:
  - **Tumbling Windows (10 min, sin traslape)**: Reporte periódico de alertas y puntaje de riesgo por localidad (`streaming_ventanas_tumbling.csv`).
  - **Sliding Windows (10 min cada 2 min, factor 5×)**: Detección de ráfagas críticas ($\ge 3$ alertas en cualquier ventana móvil de 10 minutos, como ocurre en Bosa, Ciudad Bolívar y Kennedy; `streaming_ventanas_sliding_rafagas.csv`).
  - **Session Windows (brecha de inactividad de 4 min por código DANE)**: Agrupación dinámica de sesiones activas de reporte de los orientadores escolares (`streaming_ventanas_session_colegios.csv`).
- **Tratamiento de Eventos Tardíos (Allowed Lateness y Side Output / DLQ)**: Re-emisión corregida para atrasos moderados (+5 min de holgura, `EVT-013` en Usme) y desvío a una **Dead Letter Queue (`streaming_eventos_tardios_dlq.csv`)** para eventos rurales extremadamente tardíos (`EVT-017` en Sumapaz con 28 min de retraso por caída de enlace satelital), garantizando **cero pérdida de alertas estudiantiles**.

---

*Última actualización: 22 de noviembre de 2026 — Entrega Final Proyecto Integrador Ingeniería de Datos (David Santiago Castelblanco Artunduaga & Dairo José Narváez Morales)*