# -*- coding: utf-8 -*-
"""
generar_graficos_evidencias.py

Genera las figuras técnicas y gráficas de evidencia a partir de los resultados
reales en data/ANALYTICS/ para su inclusión bajo normas APA (7.ª edición) en el
Documento Integrado del Proyecto y en la Presentación de Sustentación Final.
"""

import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as patches

plt.rcParams["font.sans-serif"] = "Segoe UI"
plt.rcParams["axes.edgecolor"] = "#CBD5E1"
plt.rcParams["axes.linewidth"] = 0.8

BASE_DIR      = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ANALYTICS_DIR = os.path.join(BASE_DIR, "data", "ANALYTICS")
ASSETS_DIR    = os.path.join(BASE_DIR, "data", "ANALYTICS", "figuras")
os.makedirs(ASSETS_DIR, exist_ok=True)


def fig1_arquitectura():
    fig, ax = plt.subplots(figsize=(11, 5.2), dpi=220)
    ax.set_xlim(0, 11)
    ax.set_ylim(0, 5.2)
    ax.axis("off")

    def draw_box(x, y, w, h, title, body, bg, border, title_color="#0F172A"):
        rect = patches.FancyBboxPatch(
            (x, y), w, h, boxstyle="round,pad=0.12,rounding_size=0.15",
            facecolor=bg, edgecolor=border, linewidth=1.8
        )
        ax.add_patch(rect)
        ax.text(x + w / 2, y + h - 0.28, title, ha="center", va="top",
                fontsize=10.5, fontweight="bold", color=title_color)
        ax.text(x + w / 2, y + h / 2 - 0.12, body, ha="center", va="center",
                fontsize=8.7, color="#334155", linespacing=1.35)

    # Banner superior
    ax.text(5.5, 4.92, "Arquitectura Híbrida del Sistema: Data Lakehouse (Medallón) + Streaming Orientado a Eventos (Corte 3)",
            ha="center", va="center", fontsize=12, fontweight="bold", color="#0F172A")

    # Fila Superior: Flujo Batch Medallón
    draw_box(0.3, 2.85, 2.2, 1.55, "Fuentes Datos Abiertos",
             "• Pruebas Saber 11 (70k)\n• Directorio Colegios (2.5k)\n• Jóvenes a la E (15k)\n• Alertas Localidad (20)",
             "#F8FAFC", "#64748B")
    draw_box(3.0, 2.85, 2.2, 1.55, "Capa Bronze (RAW)",
             "Inmutabilidad Estricta\ndata/RAW/*.csv\nSin alteración de origen\nTrazabilidad y auditoría",
             "#FEF3C7", "#D97706", "#92400E")
    draw_box(5.7, 2.85, 2.3, 1.55, "Capa Silver (STAGING)",
             "Contrato de Calidad DAMA\n• Filtro centinelas ('-', 'N/A')\n• Imputación moda DANE\n• Colegios ACTIVOS (2,262)",
             "#E0F2FE", "#0284C7", "#075985")
    draw_box(8.5, 2.85, 2.2, 1.55, "Capa Gold (ANALYTICS)",
             "Datamart Estrella + Alertas\n• LEFT JOIN Saber11-DANE\n• Regla Deserción > 8.0%\n• Absorción IES y Equidad",
             "#DCFCE7", "#16A34A", "#166534")

    # Fila Inferior: Flujo Streaming Corte 3
    draw_box(0.3, 0.55, 2.2, 1.65, "Productores IED (EDA)",
             "Colegios Distritales (20 loc)\nEventos CloudEvents 1.0\nevent_time vs processing_time\nHash SHA-256 (Ley 1581)",
             "#F3E8FF", "#9333EA", "#6B21A8")
    draw_box(3.0, 0.55, 2.2, 1.65, "Broker & Idempotencia",
             "Tópico sed.alertas.asistencia\nParticionado por Localidad\nDeduplicación O(1) por ID\nLog ordenado con Offsets",
             "#EDE9FE", "#7C3AED", "#5B21B6")
    draw_box(5.7, 0.55, 2.3, 1.65, "Motor Streaming & WM",
             "WM = max(ET) - 5 min\n• Tumbling (10m sin solape)\n• Sliding (10m / 2m ráfagas)\n• Session (gap 4m por DANE)",
             "#DBEAFE", "#2563EB", "#1E40AF")
    draw_box(8.5, 0.55, 2.2, 1.65, "Políticas Tardíos & DLQ",
             "• En ventana: State Store\n• Allowed Lateness (+5m)\n• Extremos (>10m): desvío\n  a Side Output / DLQ SED",
             "#FEE2E2", "#DC2626", "#991B1B")

    # Flechas horizontales
    arrow_kw = dict(arrowstyle="->,head_width=0.25,head_length=0.25", color="#1E293B", lw=1.8)
    for x1, x2 in [(2.52, 2.98), (5.22, 5.68), (8.02, 8.48)]:
        ax.annotate("", xy=(x2, 3.62), xytext=(x1, 3.62), arrowprops=arrow_kw)
        ax.annotate("", xy=(x2, 1.38), xytext=(x1, 1.38), arrowprops=arrow_kw)

    # Flecha vertical de convergencia hacia Gold
    ax.annotate("", xy=(9.6, 2.83), xytext=(9.6, 2.22), arrowprops=dict(arrowstyle="->,head_width=0.25,head_length=0.25", color="#16A34A", lw=1.8))

    plt.tight_layout()
    out = os.path.join(ASSETS_DIR, "fig1_arquitectura_lakehouse_streaming.png")
    plt.savefig(out, dpi=220, bbox_inches="tight")
    plt.close()
    return out


def fig2_calidad_datos():
    fig, ax = plt.subplots(figsize=(8.8, 4.4), dpi=220)

    categorias = [
        "Códigos DANE Nulos\n(Saber 11)",
        "Centinelas en Puntaje\n('-', 'N/A', 'ND')",
        "Colegios Inactivos / En Proceso\n(Directorio SED)",
        "Eventos Duplicados en Stream\n(Reintentos de Red)"
    ]
    antes_bronze = [2934, 1070, 238, 1]
    despues_silver = [0, 0, 0, 0]

    x = np.arange(len(categorias))
    width = 0.36

    b1 = ax.bar(x - width/2, antes_bronze, width, label="Incidencias Detectadas en Capa Bronze (RAW)", color="#DC2626", alpha=0.88)
    b2 = ax.bar(x + width/2, despues_silver, width, label="Incidencias Residuales en Capa Silver (STAGING)", color="#16A34A", alpha=0.92)

    ax.set_ylabel("Cantidad de Registros Afectados", fontsize=10, fontweight="bold", color="#1E293B")
    ax.set_title("Impacto del Contrato de Calidad de Datos (Paso de Capa Bronze a Capa Silver)", fontsize=11.5, fontweight="bold", color="#0F172A", pad=12)
    ax.set_xticks(x)
    ax.set_xticklabels(categorias, fontsize=9.2, color="#1E293B")
    ax.legend(frameon=True, facecolor="#F8FAFC", edgecolor="#CBD5E1", fontsize=9)
    ax.grid(axis="y", linestyle="--", alpha=0.45)

    for bar in b1:
        h = bar.get_height()
        ax.annotate(f"{h:,}", xy=(bar.get_x() + bar.get_width()/2, h),
                    xytext=(0, 4), textcoords="offset points", ha="center", va="bottom",
                    fontsize=9, fontweight="bold", color="#991B1B")
    for bar in b2:
        h = bar.get_height()
        ax.annotate("0 (Saneado)", xy=(bar.get_x() + bar.get_width()/2, h),
                    xytext=(0, 4), textcoords="offset points", ha="center", va="bottom",
                    fontsize=8.5, fontweight="bold", color="#166534")

    ax.set_ylim(0, 3450)
    plt.tight_layout()
    out = os.path.join(ASSETS_DIR, "fig2_calidad_datos_bronze_silver.png")
    plt.savefig(out, dpi=220, bbox_inches="tight")
    plt.close()
    return out


def fig3_equidad_localidades():
    df_eq = pd.read_csv(os.path.join(ANALYTICS_DIR, "reporte_equidad_educativa.csv"))
    df_eq = df_eq.sort_values(by="TASA_DESERCION", ascending=True)

    fig, ax = plt.subplots(figsize=(9.5, 5.4), dpi=220)
    colores = ["#DC2626" if t > 8.0 else "#2563EB" for t in df_eq["TASA_DESERCION"]]

    bars = ax.barh(df_eq["COLE_LOCALIDAD"], df_eq["TASA_DESERCION"], color=colores, alpha=0.88, height=0.68)
    ax.axvline(8.0, color="#991B1B", linestyle="--", linewidth=1.8, label="Umbral Crítico de Deserción (8.0%)")

    ax.set_xlabel("Tasa de Deserción Escolar (%)", fontsize=10, fontweight="bold", color="#1E293B")
    ax.set_title("Evaluación Automática de Alertas de Deserción y Puntaje Global Saber 11 por Localidad",
                 fontsize=11.2, fontweight="bold", color="#0F172A", pad=12)
    ax.tick_params(axis="y", labelsize=8.5)
    ax.grid(axis="x", linestyle="--", alpha=0.45)

    for bar, (_, row) in zip(bars, df_eq.iterrows()):
        w = bar.get_width()
        estado = "ALERTA" if w > 8.0 else "OK"
        ax.text(w + 0.18, bar.get_y() + bar.get_height()/2,
                f"{w:.1f}% | Saber11: {row['promedio_punt_global']:.1f} ({estado})",
                va="center", ha="left", fontsize=7.8,
                color="#991B1B" if w > 8.0 else "#1E3A8A", fontweight="bold" if w > 8.0 else "normal")

    ax.set_xlim(0, 16.5)
    ax.legend(loc="lower right", frameon=True, facecolor="#F8FAFC", edgecolor="#CBD5E1", fontsize=9)
    plt.tight_layout()
    out = os.path.join(ASSETS_DIR, "fig3_equidad_desercion_localidades.png")
    plt.savefig(out, dpi=220, bbox_inches="tight")
    plt.close()
    return out


def fig4_streaming_watermarks():
    df_rel = pd.read_csv(os.path.join(ANALYTICS_DIR, "streaming_comparacion_event_vs_processing.csv"))
    df_traza = pd.read_csv(os.path.join(ANALYTICS_DIR, "streaming_traza_watermarks.csv"))

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(10.5, 4.3), dpi=220)

    # Panel 1: Event Time vs Processing Time
    x = np.arange(len(df_rel))
    w = 0.35
    b1 = ax1.bar(x - w/2, df_rel["conteo_event_time_real"], w, label="Event Time (Realidad en Colegio)", color="#2563EB")
    b2 = ax1.bar(x + w/2, df_rel["conteo_processing_time_servidor"], w, label="Processing Time (Reloj Servidor)", color="#F59E0B")
    ax1.set_xticks(x)
    ax1.set_xticklabels(df_rel["ventana_10_min"], fontsize=8.8)
    ax1.set_ylabel("Eventos Contabilizados", fontsize=9.5, fontweight="bold")
    ax1.set_title("A. Distorsión: Event Time vs. Processing Time", fontsize=10.5, fontweight="bold", color="#0F172A")
    ax1.legend(fontsize=8.2, frameon=True)
    ax1.grid(axis="y", linestyle="--", alpha=0.45)
    ax1.set_ylim(0, 10)
    for bar in list(b1) + list(b2):
        h = bar.get_height()
        ax1.annotate(f"{int(h)}", xy=(bar.get_x() + bar.get_width()/2, h),
                     xytext=(0, 3), textcoords="offset points", ha="center", fontsize=8.5, fontweight="bold")

    # Panel 2: Retraso por evento y clasificación frente al Watermark
    colores_map = {
        "ACEPTADO_EN_VENTANA": "#16A34A",
        "ALLOWED_LATENESS": "#D97706",
        "RECHAZADO_EXTREMO_DLQ": "#DC2626"
    }
    c_bars = [colores_map.get(c, "#64748B") for c in df_traza["clasificacion_arrivo"]]
    ax2.bar(df_traza["id"], df_traza["retraso_min"], color=c_bars, alpha=0.9)
    ax2.axhline(5.0, color="#2563EB", linestyle="--", lw=1.4, label="Umbral Watermark (5 min)")
    ax2.axhline(10.0, color="#D97706", linestyle=":", lw=1.5, label="Límite Allowed Lateness (+5 min)")
    ax2.set_ylabel("Retraso de Red / Sincronización (min)", fontsize=9.5, fontweight="bold")
    ax2.set_title("B. Retraso por Evento y Política Aplicada (WM / DLQ)", fontsize=10.5, fontweight="bold", color="#0F172A")
    ax2.tick_params(axis="x", rotation=65, labelsize=7.8)
    ax2.grid(axis="y", linestyle="--", alpha=0.45)
    ax2.legend(fontsize=8, loc="upper left", frameon=True)

    plt.tight_layout()
    out = os.path.join(ASSETS_DIR, "fig4_streaming_watermarks_ventanas.png")
    plt.savefig(out, dpi=220, bbox_inches="tight")
    plt.close()
    return out


if __name__ == "__main__":
    f1 = fig1_arquitectura()
    f2 = fig2_calidad_datos()
    f3 = fig3_equidad_localidades()
    f4 = fig4_streaming_watermarks()
    print("Figuras generadas exitosamente:")
    for f in (f1, f2, f3, f4):
        print(" ->", f)
