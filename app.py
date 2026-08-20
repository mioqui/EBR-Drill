from __future__ import annotations

from io import BytesIO
from pathlib import Path
import hashlib
import math
import tempfile
import zipfile

import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

from procesador import procesar_pdf


# ==========================================================
# CONFIGURACIÓN DE LA APP
# ==========================================================

st.set_page_config(
    page_title="EBR Drill Analytics",
    page_icon="⛏️",
    layout="wide",
)

st.title("EBR Drill Analytics")
st.caption("Sandvik iSURE® Round Report Analysis · El Brocal · v15.3-Beeswarm")

st.info(
    "Carga uno o varios reportes PDF de iSURE®. "
    "La aplicación procesa únicamente los archivos nuevos, muestra los "
    "indicadores de perforación y uso automático del movimiento de brazos, "
    "y genera un Excel consolidado."
)


# ==========================================================
# ESTADO DE SESIÓN
# ==========================================================

if "uploader_version" not in st.session_state:
    st.session_state.uploader_version = 0

if "procesados" not in st.session_state:
    st.session_state.procesados = {}


def limpiar_analisis():
    st.session_state.procesados = {}
    st.session_state.uploader_version += 1


# ==========================================================
# ESTILO DEL FILE UPLOADER
# ==========================================================

st.markdown(
    """
    <style>
    .block-container { padding-top: 0.8rem; padding-bottom: 0.8rem; }
    [data-testid="stFileUploader"] { font-size: 0.90rem; }
    [data-testid="stFileUploader"] button { min-width: 40px !important; min-height: 40px !important; font-size: 1.0rem !important; border-radius: 8px !important; }
    [data-testid="stFileUploader"] button svg { width: 21px !important; height: 21px !important; }
    [data-testid="stFileUploaderFile"] { margin-right: 4px !important; margin-bottom: 4px !important; }
    [data-testid="stExpander"] summary { font-size: 0.88rem !important; font-weight: 600 !important; padding-top: 0.35rem !important; padding-bottom: 0.35rem !important; }
    [data-testid="stMetric"] { padding: 0.05rem 0 !important; }
    [data-testid="stMetricLabel"] { font-size: 0.72rem !important; line-height: 1.05 !important; margin-bottom: 0.06rem !important; }
    [data-testid="stMetricValue"] { font-size: 1.20rem !important; line-height: 1.05 !important; font-weight: 600 !important; }
    h1 { font-size: 1.70rem !important; margin-bottom: 0.2rem !important; }
    h2 { font-size: 1.15rem !important; margin-top: 0.65rem !important; margin-bottom: 0.3rem !important; }
    h3 { font-size: 0.98rem !important; margin-top: 0.50rem !important; margin-bottom: 0.22rem !important; }
    h4 { font-size: 0.90rem !important; margin-top: 0.42rem !important; margin-bottom: 0.18rem !important; }
    [data-testid="stCaptionContainer"] { font-size: 0.70rem !important; line-height: 1.20 !important; }
    [data-testid="stAlert"] { padding: 0.42rem 0.60rem !important; font-size: 0.76rem !important; }
    [data-testid="stDataFrame"] div[role="grid"] { font-size: 0.76rem !important; }
    .stDownloadButton button, .stButton button { min-height: 2.05rem !important; padding: 0.30rem 0.65rem !important; font-size: 0.78rem !important; }
    div[data-testid="stVerticalBlock"] > div { gap: 0.36rem !important; }
    hr { margin-top: 0.30rem !important; margin-bottom: 0.40rem !important; }
    </style>
    """,
    unsafe_allow_html=True,
)


# ==========================================================
# FUNCIONES AUXILIARES
# ==========================================================

CACHE_SCHEMA_VERSION = "v15_3_beeswarm"


def hash_archivo(uploaded_file) -> str:
    contenido = uploaded_file.getvalue()

    return hashlib.sha256(
        contenido
        + CACHE_SCHEMA_VERSION.encode("utf-8")
    ).hexdigest()


def crear_excel(
    df_reportes: pd.DataFrame,
    df_resumen: pd.DataFrame,
    df_detalle: pd.DataFrame,
    df_validacion: pd.DataFrame,
    df_extras: pd.DataFrame,
    df_automatico: pd.DataFrame,
) -> bytes:
    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_reportes.to_excel(
            writer,
            sheet_name="Resumen_Reportes",
            index=False,
        )
        df_resumen.to_excel(
            writer,
            sheet_name="Resumen_Ciclos",
            index=False,
        )
        df_detalle.to_excel(
            writer,
            sheet_name="Detalle_Barrenos",
            index=False,
        )
        df_validacion.to_excel(
            writer,
            sheet_name="Validacion",
            index=False,
        )
        df_extras.to_excel(
            writer,
            sheet_name="Barrenos_Extra",
            index=False,
        )
        df_automatico.to_excel(
            writer,
            sheet_name="Uso_Automatico",
            index=False,
        )

    buffer.seek(0)
    wb = load_workbook(buffer)

    fill_header = PatternFill(
        fill_type="solid",
        fgColor="1F4E78",
    )
    font_header = Font(
        color="FFFFFF",
        bold=True,
    )

    for ws in wb.worksheets:
        ws.freeze_panes = "A2"

        if ws.max_row >= 1 and ws.max_column >= 1:
            ws.auto_filter.ref = ws.dimensions

        for cell in ws[1]:
            cell.fill = fill_header
            cell.font = font_header
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
            )

        for columna in ws.columns:
            letra = columna[0].column_letter
            max_length = 0

            for cell in columna:
                if cell.value is None:
                    continue

                max_length = max(
                    max_length,
                    len(str(cell.value)),
                )

            ws.column_dimensions[letra].width = min(
                max_length + 2,
                35,
            )

    salida = BytesIO()
    wb.save(salida)
    salida.seek(0)

    return salida.getvalue()


def figura_a_png(fig) -> bytes:
    buffer = BytesIO()

    fig.savefig(
        buffer,
        format="png",
        dpi=250,
        bbox_inches="tight",
    )

    buffer.seek(0)
    return buffer.getvalue()


def crear_zip_graficos(graficos) -> bytes:
    buffer = BytesIO()

    with zipfile.ZipFile(
        buffer,
        mode="w",
        compression=zipfile.ZIP_DEFLATED,
    ) as zip_file:
        for nombre, contenido in graficos:
            zip_file.writestr(
                nombre,
                contenido,
            )

    buffer.seek(0)
    return buffer.getvalue()


def nombre_seguro(texto: str) -> str:
    if not texto:
        return "sin_nombre"

    texto = str(texto)

    for origen, destino in {
        "/": "-",
        "\\": "-",
        ":": "-",
        " ": "_",
    }.items():
        texto = texto.replace(
            origen,
            destino,
        )

    return texto


def fmt_pct(valor):
    if valor is None or pd.isna(valor):
        return "-"
    return f"{valor:.1f}%"


def clasificar_tipo_disparo(barrenos_realizados):
    """
    Clasificación operativa por número de barrenos realizados:

    FRENTE: >= 45
    SELLADA: 25 a 44
    ESTOCADA Y/O CORRECCIONES: < 25
    """
    if barrenos_realizados is None or pd.isna(barrenos_realizados):
        return "SIN CLASIFICAR"

    n = int(barrenos_realizados)

    if n >= 45:
        return "FRENTE"
    elif n >= 25:
        return "SELLADA"
    else:
        return "ESTOCADA Y/O CORRECCIONES"


def calcular_kpi_movimiento_desde_brazos(
    auto_b1,
    auto_b2,
    manual_b1,
    manual_b2,
):
    """
    KPI reconciliado a partir de los tiempos individuales de ambos brazos.

    % Auto Jumbo =
        (Auto B1 + Auto B2)
        / (Auto B1 + Auto B2 + Manual B1 + Manual B2) * 100
    """
    valores = [
        auto_b1,
        auto_b2,
        manual_b1,
        manual_b2,
    ]

    if any(
        v is None or pd.isna(v)
        for v in valores
    ):
        return None, None, None, None

    auto_total_brazos = (
        float(auto_b1)
        + float(auto_b2)
    )

    manual_total_brazos = (
        float(manual_b1)
        + float(manual_b2)
    )

    total = (
        auto_total_brazos
        + manual_total_brazos
    )

    if total <= 0:
        return (
            auto_total_brazos,
            manual_total_brazos,
            None,
            None,
        )

    pct_auto = (
        auto_total_brazos
        / total
        * 100
    )

    pct_manual = (
        manual_total_brazos
        / total
        * 100
    )

    return (
        auto_total_brazos,
        manual_total_brazos,
        pct_auto,
        pct_manual,
    )


# ==========================================================
# HELPERS PARA GRÁFICOS DE TENDENCIA
# ==========================================================

def preparar_eje_fechas(df_plot: pd.DataFrame):
    """
    Asigna una posición X por fecha y distribuye varios ciclos
    dentro del mismo día alrededor del centro de esa fecha.
    """
    fechas = list(
        df_plot["Fecha"]
        .drop_duplicates()
        .sort_values()
    )

    posiciones = {}

    for indice_dia, fecha in enumerate(fechas):
        grupo_dia = df_plot[
            df_plot["Fecha"] == fecha
        ].sort_values(
            ["FechaHora", "Jumbo", "Ciclo"]
        )

        n = len(grupo_dia)

        if n == 1:
            offsets = [0.0]
        else:
            paso = 0.64 / (n - 1)
            offsets = [
                -0.32 + (i * paso)
                for i in range(n)
            ]

        for offset, idx in zip(
            offsets,
            grupo_dia.index,
        ):
            posiciones[idx] = (
                indice_dia + offset
            )

    return fechas, posiciones


def aplicar_formato_eje_fechas(
    ax,
    fechas,
):
    centros_dia = list(
        range(len(fechas))
    )

    ax.set_xticks(
        centros_dia
    )

    ax.set_xticklabels(
        [
            pd.Timestamp(
                fecha
            ).strftime("%d/%m")
            for fecha in fechas
        ],
        fontsize=9,
    )

    for i in range(
        len(fechas) - 1
    ):
        ax.axvline(
            i + 0.5,
            linewidth=0.8,
            alpha=0.18,
        )

    ax.set_xlim(
        -0.55,
        len(fechas) - 0.45,
    )


def anotar_etiquetas_sin_solape(
    ax,
    puntos,
):
    """
    Distribuye las etiquetas de TODOS los puntos del gráfico evitando
    solapes entre series.

    Cada elemento de ``puntos`` debe contener:
        x, y, texto

    La lógica prueba posiciones alrededor del punto (arriba, abajo y
    desplazamientos laterales) y elige la primera que no colisiona con
    las etiquetas ya ubicadas. Si todas colisionan, escoge la posición
    con menor área de solape.
    """
    if not puntos:
        return

    fig = ax.figure

    # Necesario para medir el tamaño real de cada etiqueta en pantalla.
    fig.canvas.draw()
    renderer = fig.canvas.get_renderer()
    bbox_ejes = ax.get_window_extent(renderer=renderer)

    # Separación visual extra entre etiquetas, expresada como expansión
    # de su caja en píxeles.
    cajas_ocupadas = []

    candidatos_arriba = [
        (0, 12),
        (-20, 14),
        (20, 14),
        (0, 27),
        (-34, 24),
        (34, 24),
        (0, -18),
        (-20, -20),
        (20, -20),
        (0, -34),
        (-36, -32),
        (36, -32),
    ]

    candidatos_abajo = [
        (0, -18),
        (-20, -20),
        (20, -20),
        (0, -34),
        (-34, -32),
        (34, -32),
        (0, 12),
        (-20, 14),
        (20, 14),
        (0, 27),
        (-36, 24),
        (36, 24),
    ]

    def area_solape(a, b):
        ancho = max(
            0.0,
            min(a.x1, b.x1)
            - max(a.x0, b.x0),
        )
        alto = max(
            0.0,
            min(a.y1, b.y1)
            - max(a.y0, b.y0),
        )
        return ancho * alto

    # Orden estable: primero por X y luego por Y. De esta manera los
    # puntos de una misma fecha se resuelven juntos.
    puntos_ordenados = sorted(
        puntos,
        key=lambda p: (
            float(p["x"]),
            float(p["y"]),
            str(p.get("serie", "")),
        ),
    )

    for i, punto in enumerate(puntos_ordenados):
        x = float(punto["x"])
        y = float(punto["y"])
        texto = str(punto["texto"])

        # Alternar la preferencia inicial ayuda a separar etiquetas de
        # dos series que caen exactamente sobre el mismo punto.
        candidatos = (
            candidatos_arriba
            if i % 2 == 0
            else candidatos_abajo
        )

        anotacion = ax.annotate(
            texto,
            xy=(x, y),
            xytext=candidatos[0],
            textcoords="offset points",
            ha="center",
            va="bottom",
            fontsize=8.5,
            fontweight="medium",
            annotation_clip=True,
            arrowprops={
                "arrowstyle": "-",
                "linewidth": 0.55,
                "alpha": 0.38,
            },
        )

        mejor = None
        menor_solape = None

        for dx, dy in candidatos:
            anotacion.set_position(
                (dx, dy)
            )
            anotacion.set_va(
                "bottom"
                if dy >= 0
                else "top"
            )

            caja = anotacion.get_window_extent(
                renderer=renderer
            )

            # Amplía un poco la caja para dejar aire entre textos.
            caja_segura = caja.expanded(
                1.10,
                1.18,
            )

            # Evitar que el texto se salga del área útil del gráfico.
            margen = 3.0
            dentro_ejes = (
                caja_segura.x0
                >= bbox_ejes.x0 + margen
                and caja_segura.x1
                <= bbox_ejes.x1 - margen
                and caja_segura.y0
                >= bbox_ejes.y0 + margen
                and caja_segura.y1
                <= bbox_ejes.y1 - margen
            )

            if not dentro_ejes:
                continue

            solape_total = sum(
                area_solape(
                    caja_segura,
                    ocupada,
                )
                for ocupada
                in cajas_ocupadas
            )

            if solape_total == 0:
                mejor = (
                    dx,
                    dy,
                    caja_segura,
                )
                break

            if (
                menor_solape is None
                or solape_total
                < menor_solape[0]
            ):
                menor_solape = (
                    solape_total,
                    dx,
                    dy,
                    caja_segura,
                )

        if mejor is None:
            if menor_solape is not None:
                _, dx, dy, caja_segura = (
                    menor_solape
                )
                mejor = (
                    dx,
                    dy,
                    caja_segura,
                )
            else:
                # Fallback muy defensivo.
                dx, dy = candidatos[0]
                anotacion.set_position(
                    (dx, dy)
                )
                anotacion.set_va(
                    "bottom"
                    if dy >= 0
                    else "top"
                )
                caja_segura = (
                    anotacion
                    .get_window_extent(
                        renderer=renderer
                    )
                    .expanded(
                        1.10,
                        1.18,
                    )
                )
                mejor = (
                    dx,
                    dy,
                    caja_segura,
                )

        dx, dy, caja_final = mejor

        anotacion.set_position(
            (dx, dy)
        )
        anotacion.set_va(
            "bottom"
            if dy >= 0
            else "top"
        )

        # La línea guía solo se muestra cuando la etiqueta fue desplazada
        # de forma apreciable.
        if anotacion.arrow_patch is not None:
            anotacion.arrow_patch.set_visible(
                abs(dx) >= 18
                or abs(dy) >= 24
            )

        cajas_ocupadas.append(
            caja_final
        )


def crear_puntos_etiquetas(
    df_plot,
    x_col,
    y_col,
    etiqueta_extra_col=None,
    serie=None,
):
    """Convierte una serie del DataFrame en etiquetas para el solver."""
    puntos = []

    for _, fila in df_plot.iterrows():
        x = fila.get(x_col)
        y = fila.get(y_col)

        if (
            x is None
            or y is None
            or pd.isna(x)
            or pd.isna(y)
        ):
            continue

        texto = f"{float(y):.1f}%"

        if (
            etiqueta_extra_col
            and etiqueta_extra_col
            in df_plot.columns
        ):
            valor_extra = fila.get(
                etiqueta_extra_col
            )

            if (
                valor_extra is not None
                and not pd.isna(valor_extra)
            ):
                texto += (
                    f"\n{int(valor_extra)} tal."
                )

        puntos.append(
            {
                "x": float(x),
                "y": float(y),
                "texto": texto,
                "serie": serie or "",
            }
        )

    return puntos


# ==========================================================
# GRÁFICO GENERAL: USO AUTOMÁTICO POR EQUIPO
# ==========================================================

def generar_grafico_tendencia_automatico(
    df_auto: pd.DataFrame,
):
    df_plot = df_auto.copy()

    if "Considerado_KPI_Automatizacion" in df_plot.columns:
        df_plot = df_plot[
            df_plot["Considerado_KPI_Automatizacion"] == True
        ].copy()

    df_plot = df_plot[
        df_plot[
            "Pct_Movimiento_Automatico_Brazos"
        ].notna()
    ].copy()

    if df_plot.empty:
        return None

    df_plot["FechaHora"] = pd.to_datetime(
        (
            df_plot[
                "Fecha_Inicio"
            ].fillna("")
            + " "
            + df_plot[
                "Hora_Inicio"
            ].fillna("")
        ),
        format="%d/%m/%Y %H:%M:%S",
        errors="coerce",
    )

    df_plot = df_plot[
        df_plot[
            "FechaHora"
        ].notna()
    ].copy()

    if df_plot.empty:
        return None

    df_plot["Fecha"] = (
        df_plot[
            "FechaHora"
        ].dt.normalize()
    )

    df_plot = df_plot.sort_values(
        [
            "FechaHora",
            "Jumbo",
            "Ciclo",
        ]
    ).reset_index(
        drop=True
    )

    fechas, posiciones = (
        preparar_eje_fechas(
            df_plot
        )
    )

    df_plot["X"] = [
        posiciones[idx]
        for idx in df_plot.index
    ]

    fig, ax = plt.subplots(
        figsize=(14, 7.0)
    )

    # ------------------------------------------------------
    # KPI GLOBAL PONDERADO POR MINUTOS DE MOVIMIENTO
    # % Auto global = Σ Auto / (Σ Auto + Σ Manual) × 100
    # ------------------------------------------------------

    kpis_globales = []

    for jumbo, grupo in df_plot.groupby(
        "Jumbo",
        dropna=False,
    ):
        auto_total = (
            grupo[
                "Auto_Brazo1_min"
            ].sum(
                min_count=1
            )
            +
            grupo[
                "Auto_Brazo2_min"
            ].sum(
                min_count=1
            )
        )

        manual_total = (
            grupo[
                "Manual_Brazo1_min"
            ].sum(
                min_count=1
            )
            +
            grupo[
                "Manual_Brazo2_min"
            ].sum(
                min_count=1
            )
        )

        if (
            pd.notna(auto_total)
            and pd.notna(manual_total)
            and (
                auto_total
                + manual_total
            ) > 0
        ):
            pct_global = (
                auto_total
                / (
                    auto_total
                    + manual_total
                )
                * 100
            )

            kpis_globales.append(
                (
                    str(jumbo),
                    pct_global,
                )
            )

    puntos_etiquetas = []

    for jumbo, grupo in df_plot.groupby(
        "Jumbo",
        dropna=False,
    ):
        grupo = grupo.sort_values(
            "FechaHora"
        )

        ax.plot(
            grupo["X"],
            grupo[
                "Pct_Movimiento_Automatico_Brazos"
            ],
            marker="o",
            markersize=6,
            linewidth=2.0,
            label=str(jumbo),
        )

        puntos_etiquetas.extend(
            crear_puntos_etiquetas(
                grupo,
                "X",
                "Pct_Movimiento_Automatico_Brazos",
                etiqueta_extra_col="Barrenos_Realizados",
                serie=str(jumbo),
            )
        )

    aplicar_formato_eje_fechas(
        ax,
        fechas,
    )

    max_y = float(
        df_plot[
            "Pct_Movimiento_Automatico_Brazos"
        ].max()
    )

    limite_superior = max(
        50,
        math.ceil(
            (max_y + 10)
            / 10
        )
        * 10,
    )

    limite_superior = min(
        100,
        limite_superior,
    )

    ax.set_ylim(
        0,
        limite_superior,
    )

    ax.set_yticks(
        range(
            0,
            int(
                limite_superior
            )
            + 1,
            10,
        )
    )

    ax.set_title(
        "Evolución del uso automático del movimiento de brazos",
        fontsize=15,
        pad=18,
    )

    ax.set_xlabel(
        "Fecha",
        labelpad=10,
    )

    ax.set_ylabel(
        "Movimiento automático (%)",
        labelpad=10,
    )

    ax.grid(
        axis="y",
        alpha=0.22,
    )

    ax.grid(
        axis="x",
        visible=False,
    )

    ax.legend(
        title="Jumbo",
        frameon=False,
        loc="upper right",
    )

    ax.spines[
        "top"
    ].set_visible(False)

    ax.spines[
        "right"
    ].set_visible(False)

    # ------------------------------------------------------
    # KPI GLOBAL EN LA PARTE SUPERIOR DEL GRÁFICO
    # ------------------------------------------------------

    if kpis_globales:
        if len(kpis_globales) == 1:
            posiciones_x = [0.50]
        else:
            posiciones_x = [
                0.34,
                0.66,
            ]

        for x, (
            jumbo,
            pct_global,
        ) in zip(
            posiciones_x,
            kpis_globales,
        ):
            fig.text(
                x,
                0.945,
                (
                    f"{jumbo}  |  "
                    f"Automático global: "
                    f"{pct_global:.1f}%"
                ),
                ha="center",
                va="center",
                fontsize=10,
                fontweight="bold",
                bbox={
                    "boxstyle":
                        "round,pad=0.45",
                    "facecolor":
                        "white",
                    "edgecolor":
                        "0.75",
                    "linewidth":
                        0.8,
                },
            )

    # ------------------------------------------------------
    # NOTA METODOLÓGICA
    # ------------------------------------------------------

    fig.text(
        0.5,
        0.018,
        (
            "Nota: el KPI considera solo disparos FRENTE (>=45 barrenos). "
            "El porcentaje global del Jumbo se calcula con los tiempos individuales "
            "de ambos brazos: Σ(Auto B1 + Auto B2) / "
            "Σ(Auto B1 + Auto B2 + Manual B1 + Manual B2) × 100."
        ),
        ha="center",
        va="bottom",
        fontsize=8.5,
        color="dimgray",
    )

    # Deja espacio para los KPI superiores y la nota inferior.
    fig.tight_layout(
        rect=[
            0,
            0.055,
            1,
            0.89,
        ]
    )

    # Resolver conjuntamente las etiquetas de todos los Jumbos.
    anotar_etiquetas_sin_solape(
        ax,
        puntos_etiquetas,
    )

    return fig


# ==========================================================
# GRÁFICO: EVOLUCIÓN DE LA MEDIANA DE BARRENOS CUT
# ==========================================================

def generar_grafico_mediana_cut(
    df_resumen: pd.DataFrame,
):
    """
    Muestra la evolución de la mediana de longitud perforada
    de los barrenos Cut para cada ciclo.

    Cada Jumbo se representa como una serie independiente.
    """
    if df_resumen is None or df_resumen.empty:
        return None

    columnas_requeridas = {
        "Tipo",
        "Mediana",
        "Fecha_Inicio",
        "Hora_Inicio",
        "Jumbo",
        "Ciclo",
    }

    if not columnas_requeridas.issubset(df_resumen.columns):
        return None

    df_plot = df_resumen.copy()

    # Tipo puede ser categórico; convertir a texto evita problemas
    # al filtrar y conserva únicamente los barrenos Cut.
    df_plot = df_plot[
        df_plot["Tipo"].astype(str).str.lower() == "cut"
    ].copy()

    df_plot = df_plot[
        df_plot["Mediana"].notna()
    ].copy()

    if df_plot.empty:
        return None

    df_plot["FechaHora"] = pd.to_datetime(
        (
            df_plot["Fecha_Inicio"].fillna("")
            + " "
            + df_plot["Hora_Inicio"].fillna("")
        ),
        format="%d/%m/%Y %H:%M:%S",
        errors="coerce",
    )

    # Fallback si no existe hora válida.
    faltantes = df_plot["FechaHora"].isna()
    if faltantes.any():
        df_plot.loc[faltantes, "FechaHora"] = pd.to_datetime(
            df_plot.loc[faltantes, "Fecha_Inicio"],
            format="%d/%m/%Y",
            errors="coerce",
        )

    df_plot = df_plot[
        df_plot["FechaHora"].notna()
    ].copy()

    if df_plot.empty:
        return None

    df_plot["Fecha"] = df_plot["FechaHora"].dt.normalize()

    df_plot = df_plot.sort_values(
        ["FechaHora", "Jumbo", "Ciclo"]
    ).reset_index(drop=True)

    fechas, posiciones = preparar_eje_fechas(df_plot)
    df_plot["X"] = [posiciones[idx] for idx in df_plot.index]

    fig, ax = plt.subplots(figsize=(14, 6.2))
    puntos_etiquetas = []

    for jumbo, grupo in df_plot.groupby("Jumbo", dropna=False):
        grupo = grupo.sort_values("FechaHora")

        ax.plot(
            grupo["X"],
            grupo["Mediana"],
            marker="o",
            markersize=6,
            linewidth=2.0,
            label=str(jumbo),
        )

        for _, fila in grupo.iterrows():
            ciclo = fila.get("Ciclo")
            ciclo_txt = (
                f"Ciclo {int(ciclo)}"
                if pd.notna(ciclo)
                else "Ciclo -"
            )
            puntos_etiquetas.append({
                "x": fila["X"],
                "y": fila["Mediana"],
                "texto": f"{fila['Mediana']:.2f} m\n{ciclo_txt}",
                "serie": str(jumbo),
            })

    aplicar_formato_eje_fechas(ax, fechas)

    min_y = float(df_plot["Mediana"].min())
    max_y = float(df_plot["Mediana"].max())
    rango = max(max_y - min_y, 0.20)
    margen = max(0.12, rango * 0.20)

    ax.set_ylim(
        max(0, min_y - margen),
        max_y + margen,
    )

    ax.set_title(
        "Evolución de la mediana de longitud de barrenos Cut",
        fontsize=15,
        pad=18,
    )
    ax.set_xlabel("Fecha", labelpad=10)
    ax.set_ylabel("Mediana longitud Cut (m)", labelpad=10)

    ax.grid(axis="y", alpha=0.22)
    ax.grid(axis="x", visible=False)

    if df_plot["Jumbo"].nunique(dropna=False) > 1:
        ax.legend(
            title="Jumbo",
            frameon=False,
            loc="upper right",
        )

    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    fig.tight_layout(rect=[0, 0.03, 1, 0.96])

    # Reutiliza la lógica de distribución de etiquetas para evitar
    # que medianas cercanas se solapen entre ciclos o Jumbos.
    anotar_etiquetas_sin_solape(
        ax,
        puntos_etiquetas,
    )

    return fig


# ==========================================================
# GRÁFICO POR JUMBO: BRAZO 1 VS BRAZO 2
# ==========================================================

def generar_grafico_brazos_por_jumbo(
    df_auto: pd.DataFrame,
    jumbo: str,
):
    """
    Compara el porcentaje automático de Brazo 1 y Brazo 2
    para un Jumbo específico.
    """
    df_plot = df_auto[
        df_auto["Jumbo"] == jumbo
    ].copy()

    if "Considerado_KPI_Automatizacion" in df_plot.columns:
        df_plot = df_plot[
            df_plot["Considerado_KPI_Automatizacion"] == True
        ].copy()

    df_plot = df_plot[
        df_plot[
            "Pct_Automatico_Brazo1"
        ].notna()
        |
        df_plot[
            "Pct_Automatico_Brazo2"
        ].notna()
    ].copy()

    if df_plot.empty:
        return None

    df_plot["FechaHora"] = pd.to_datetime(
        (
            df_plot[
                "Fecha_Inicio"
            ].fillna("")
            + " "
            + df_plot[
                "Hora_Inicio"
            ].fillna("")
        ),
        format="%d/%m/%Y %H:%M:%S",
        errors="coerce",
    )

    df_plot = df_plot[
        df_plot[
            "FechaHora"
        ].notna()
    ].copy()

    if df_plot.empty:
        return None

    df_plot["Fecha"] = (
        df_plot[
            "FechaHora"
        ].dt.normalize()
    )

    df_plot = df_plot.sort_values(
        [
            "FechaHora",
            "Ciclo",
        ]
    ).reset_index(
        drop=True
    )

    fechas, posiciones = (
        preparar_eje_fechas(
            df_plot
        )
    )

    df_plot["X"] = [
        posiciones[idx]
        for idx in df_plot.index
    ]

    # Gap absoluto en puntos porcentuales
    df_plot[
        "Gap_Brazos_pp"
    ] = (
        df_plot[
            "Pct_Automatico_Brazo1"
        ]
        - df_plot[
            "Pct_Automatico_Brazo2"
        ]
    ).abs()

    fig, ax = plt.subplots(
        figsize=(14, 6.2)
    )

    series = [
        (
            "Brazo 1",
            "Pct_Automatico_Brazo1",
        ),
        (
            "Brazo 2",
            "Pct_Automatico_Brazo2",
        ),
    ]

    puntos_etiquetas = []

    for nombre_serie, columna in series:
        grupo = df_plot[
            df_plot[
                columna
            ].notna()
        ].copy()

        if grupo.empty:
            continue

        ax.plot(
            grupo["X"],
            grupo[columna],
            marker="o",
            markersize=6,
            linewidth=2.0,
            label=nombre_serie,
        )

        puntos_etiquetas.extend(
            crear_puntos_etiquetas(
                grupo,
                "X",
                columna,
                etiqueta_extra_col="Barrenos_Realizados",
                serie=nombre_serie,
            )
        )

    aplicar_formato_eje_fechas(
        ax,
        fechas,
    )

    max_y = float(
        pd.concat(
            [
                df_plot[
                    "Pct_Automatico_Brazo1"
                ],
                df_plot[
                    "Pct_Automatico_Brazo2"
                ],
            ]
        ).max()
    )

    limite_superior = max(
        50,
        math.ceil(
            (max_y + 10)
            / 10
        )
        * 10,
    )

    limite_superior = min(
        100,
        limite_superior,
    )

    ax.set_ylim(
        0,
        limite_superior,
    )

    ax.set_yticks(
        range(
            0,
            int(
                limite_superior
            )
            + 1,
            10,
        )
    )

    ax.set_title(
        f"{jumbo} · Uso automático por brazo",
        fontsize=15,
        pad=18,
    )

    ax.set_xlabel(
        "Fecha",
        labelpad=10,
    )

    ax.set_ylabel(
        "Uso automático (%)",
        labelpad=10,
    )

    ax.grid(
        axis="y",
        alpha=0.22,
    )

    ax.grid(
        axis="x",
        visible=False,
    )

    ax.legend(
        title="Brazo",
        frameon=False,
        loc="upper right",
    )

    ax.spines[
        "top"
    ].set_visible(False)

    ax.spines[
        "right"
    ].set_visible(False)

    # ------------------------------------------------------
    # KPI GLOBAL PONDERADO POR BRAZO
    # Brazo 1 = Σ Auto B1 / (Σ Auto B1 + Σ Manual B1)
    # Brazo 2 = Σ Auto B2 / (Σ Auto B2 + Σ Manual B2)
    # ------------------------------------------------------

    auto_b1 = df_plot[
        "Auto_Brazo1_min"
    ].sum(
        min_count=1
    )

    manual_b1 = df_plot[
        "Manual_Brazo1_min"
    ].sum(
        min_count=1
    )

    auto_b2 = df_plot[
        "Auto_Brazo2_min"
    ].sum(
        min_count=1
    )

    manual_b2 = df_plot[
        "Manual_Brazo2_min"
    ].sum(
        min_count=1
    )

    pct_global_b1 = None
    pct_global_b2 = None

    if (
        pd.notna(auto_b1)
        and pd.notna(manual_b1)
        and (
            auto_b1
            + manual_b1
        ) > 0
    ):
        pct_global_b1 = (
            auto_b1
            / (
                auto_b1
                + manual_b1
            )
            * 100
        )

    if (
        pd.notna(auto_b2)
        and pd.notna(manual_b2)
        and (
            auto_b2
            + manual_b2
        ) > 0
    ):
        pct_global_b2 = (
            auto_b2
            / (
                auto_b2
                + manual_b2
            )
            * 100
        )

    # Gap promedio por ciclo
    gap_promedio = (
        df_plot[
            "Gap_Brazos_pp"
        ].mean()
    )

    if pd.notna(
        gap_promedio
    ):
        ax.text(
            0.01,
            0.98,
            f"Gap promedio entre brazos: {gap_promedio:.1f} pp",
            transform=ax.transAxes,
            ha="left",
            va="top",
            fontsize=9,
        )

    # KPI ponderado de cada brazo en la parte superior
    kpis_brazos = []

    if pct_global_b1 is not None:
        kpis_brazos.append(
            (
                "Brazo 1",
                pct_global_b1,
            )
        )

    if pct_global_b2 is not None:
        kpis_brazos.append(
            (
                "Brazo 2",
                pct_global_b2,
            )
        )

    if kpis_brazos:
        if len(kpis_brazos) == 1:
            posiciones_x = [
                0.50
            ]
        else:
            posiciones_x = [
                0.34,
                0.66,
            ]

        for x, (
            brazo,
            pct_global,
        ) in zip(
            posiciones_x,
            kpis_brazos,
        ):
            fig.text(
                x,
                0.945,
                (
                    f"{brazo}  |  "
                    f"Automático global: "
                    f"{pct_global:.1f}%"
                ),
                ha="center",
                va="center",
                fontsize=10,
                fontweight="bold",
                bbox={
                    "boxstyle":
                        "round,pad=0.45",
                    "facecolor":
                        "white",
                    "edgecolor":
                        "0.75",
                    "linewidth":
                        0.8,
                },
            )

    # Nota metodológica
    fig.text(
        0.5,
        0.018,
        (
            "Nota: el KPI considera solo disparos FRENTE (>=45 barrenos). "
            "El porcentaje global por brazo está ponderado por sus minutos de movimiento: "
            "Σ Automático / (Σ Automático + Σ Manual) × 100."
        ),
        ha="center",
        va="bottom",
        fontsize=8.5,
        color="dimgray",
    )

    fig.tight_layout(
        rect=[
            0,
            0.055,
            1,
            0.89,
        ]
    )

    # Resolver conjuntamente las etiquetas de Brazo 1 y Brazo 2.
    # Los puntos muy próximos se distribuyen arriba/abajo/lateralmente
    # y se agrega una línea guía solo cuando el desplazamiento es mayor.
    anotar_etiquetas_sin_solape(
        ax,
        puntos_etiquetas,
    )

    return fig


# ==========================================================
# GUARDAR RESULTADO DE PDF EN CACHÉ
# ==========================================================

def guardar_resultado_en_cache(
    archivo,
    clave_archivo: str,
):
    temp_path = None

    try:
        with tempfile.NamedTemporaryFile(
            delete=False,
            suffix=".pdf",
        ) as temp:
            temp.write(
                archivo.getbuffer()
            )
            temp_path = Path(
                temp.name
            )

        resultado = procesar_pdf(
            temp_path,
            nombre_archivo=archivo.name,
        )

        fig = resultado["fig"]
        png_bytes = figura_a_png(
            fig
        )
        plt.close(
            fig
        )

        metadata = resultado[
            "metadata"
        ]

        fecha_segura = nombre_seguro(
            metadata.get(
                "Fecha_Inicio"
            )
            or "sin_fecha"
        )

        jumbo_seguro = nombre_seguro(
            metadata.get(
                "Jumbo"
            )
            or "JUMBO"
        )

        ciclo_seguro = nombre_seguro(
            metadata.get(
                "Ciclo"
            )
            or "sin_ciclo"
        )

        nombre_png = (
            f"{jumbo_seguro}_"
            f"Ciclo_{ciclo_seguro}_"
            f"{fecha_segura}.png"
        )

        resultado_cache = {
            "nombre_archivo":
                archivo.name,
            "metadata":
                metadata,
            "movimiento":
                resultado[
                    "movimiento"
                ],
            "detalle":
                resultado[
                    "detalle"
                ],
            "validacion":
                resultado[
                    "validacion"
                ],
            "resumen_ciclo":
                resultado[
                    "resumen_ciclo"
                ],
            "resumen_reporte":
                resultado[
                    "resumen_reporte"
                ],
            "extras":
                resultado[
                    "extras"
                ],
            "png_bytes":
                png_bytes,
            "plano_nav_png":
                resultado.get(
                    "plano_nav_png"
                ),
            "nombre_png":
                nombre_png,
            "error":
                None,
        }

        st.session_state.procesados[
            clave_archivo
        ] = resultado_cache

    except Exception as exc:
        st.session_state.procesados[
            clave_archivo
        ] = {
            "nombre_archivo":
                archivo.name,
            "error":
                str(exc),
        }

    finally:
        if (
            temp_path
            and temp_path.exists()
        ):
            temp_path.unlink(
                missing_ok=True
            )


# ==========================================================
# CONTROLES DE CARGA / LIMPIEZA
# ==========================================================

st.subheader(
    "Reportes"
)

col_uploader, col_limpiar = st.columns(
    [5, 1]
)

with col_uploader:
    archivos = st.file_uploader(
        "Seleccionar o agregar reportes PDF",
        type=["pdf"],
        accept_multiple_files=True,
        key=(
            "uploader_pdfs_"
            f"{st.session_state.uploader_version}"
        ),
    )

with col_limpiar:
    st.write("")
    st.write("")

    st.button(
        "Limpiar análisis",
        use_container_width=True,
        on_click=limpiar_analisis,
    )


# ==========================================================
# PROCESAMIENTO SOLO DE ARCHIVOS NUEVOS
# ==========================================================

if archivos:
    st.write(
        f"**Archivos seleccionados:** "
        f"{len(archivos)}"
    )

    archivos_actuales = []

    for archivo in archivos:
        clave = hash_archivo(
            archivo
        )

        archivos_actuales.append(
            (
                clave,
                archivo,
            )
        )

    nuevos = [
        (
            clave,
            archivo,
        )
        for clave, archivo
        in archivos_actuales
        if clave
        not in st.session_state.procesados
    ]

    if nuevos:
        progreso = st.progress(
            0
        )

        for i, (
            clave,
            archivo,
        ) in enumerate(
            nuevos,
            start=1,
        ):
            with st.spinner(
                "Procesando archivo nuevo: "
                f"{archivo.name}"
            ):
                guardar_resultado_en_cache(
                    archivo,
                    clave,
                )

            progreso.progress(
                i / len(nuevos)
            )

        progreso.empty()

    resultados_validos = []
    errores_actuales = []

    for clave, archivo in archivos_actuales:
        resultado = (
            st.session_state.procesados.get(
                clave
            )
        )

        if not resultado:
            continue

        if resultado.get(
            "error"
        ):
            errores_actuales.append(
                {
                    "Archivo_PDF":
                        archivo.name,
                    "Error":
                        resultado[
                            "error"
                        ],
                }
            )
            continue

        resultados_validos.append(
            resultado
        )

        metadata = resultado[
            "metadata"
        ]

        movimiento = resultado[
            "movimiento"
        ]

        validacion = resultado[
            "validacion"
        ]

        extras = resultado[
            "extras"
        ]

        png_bytes = resultado[
            "png_bytes"
        ]

        # Imagen del plano de navegación extraída de la
        # primera página del reporte iSURE.
        plano_nav_png = resultado.get(
            "plano_nav_png"
        )

        nombre_png = resultado[
            "nombre_png"
        ]

        # Asignación defensiva del plano de navegación.
        # Se repite aquí para evitar cualquier variable sin definir
        # incluso si cambia la estructura previa del resultado/cache.
        plano_nav_png = resultado.get(
            "plano_nav_png",
            None,
        )

        titulo_expander = (
            f"{metadata['Jumbo']} | "
            f"Ciclo {metadata['Ciclo']} | "
            f"{metadata['Fecha_Inicio']} | "
            f"{archivo.name}"
        )

        with st.expander(
            titulo_expander,
            expanded=False,
        ):
            # ==================================================
            # CABECERA COMPACTA
            # La información continúa en la columna izquierda
            # mientras el plano de navegación ocupa la derecha.
            # Esto evita que la altura de la miniatura genere
            # un gran espacio vacío antes de la clasificación.
            # ==================================================

            col_info, col_plano = st.columns(
                [4.6, 1.25],
                gap="medium",
            )

            with col_info:
                c1, c2, c3, c4 = st.columns(
                    [1.0, 1.0, 0.85, 1.0]
                )

                c1.metric(
                    "Jumbo",
                    metadata[
                        "Jumbo"
                    ]
                    or "-",
                )

                c2.metric(
                    "Nº de serie",
                    metadata[
                        "Numero_Serie"
                    ]
                    or "-",
                )

                c3.metric(
                    "Ciclo",
                    metadata[
                        "Ciclo"
                    ]
                    or "-",
                )

                c4.metric(
                    "Metros perforados",
                    (
                        f"{metadata['Metros_Perforados']:.2f} m"
                        if metadata[
                            "Metros_Perforados"
                        ]
                        is not None
                        else "-"
                    ),
                )

                barrenos_realizados = metadata.get(
                    "Barrenos_Realizados"
                )

                tipo_disparo = clasificar_tipo_disparo(
                    barrenos_realizados
                )

                considerado_kpi = (
                    tipo_disparo == "FRENTE"
                )

                st.markdown("### Clasificación del disparo")

                d1, d2, d3 = st.columns(
                    [1.15, 1.0, 1.0]
                )

                d1.metric(
                    "Tipo de disparo",
                    tipo_disparo,
                )

                d2.metric(
                    "Barrenos realizados",
                    (
                        int(barrenos_realizados)
                        if barrenos_realizados is not None
                        else "-"
                    ),
                )

                d3.metric(
                    "Considerado en KPI Auto",
                    "Sí" if considerado_kpi else "No",
                )

                fuente_barrenos = metadata.get(
                    "Fuente_Barrenos_Realizados"
                )

                if fuente_barrenos:
                    st.caption(
                        f"Fuente del número de barrenos: {fuente_barrenos}. "
                        "Para la clasificación se excluyen los barrenos Reaming."
                    )

                st.markdown("### Uso automático del movimiento de brazos")

                m1, m2, m3, m4 = st.columns(
                    [1.0, 1.0, 1.0, 1.0]
                )

                (
                    auto_brazos_min,
                    manual_brazos_min,
                    pct_auto_brazos,
                    pct_manual_brazos,
                ) = calcular_kpi_movimiento_desde_brazos(
                    movimiento[
                        "Auto_Brazo1_min"
                    ],
                    movimiento[
                        "Auto_Brazo2_min"
                    ],
                    movimiento[
                        "Manual_Brazo1_min"
                    ],
                    movimiento[
                        "Manual_Brazo2_min"
                    ],
                )

                m1.metric(
                    "Movimiento automático",
                    fmt_pct(
                        pct_auto_brazos
                    ),
                )

                m2.metric(
                    "Movimiento manual",
                    fmt_pct(
                        pct_manual_brazos
                    ),
                )

                m3.metric(
                    "Brazo 1 automático",
                    fmt_pct(
                        movimiento[
                            "Pct_Automatico_Brazo1"
                        ]
                    ),
                )

                m4.metric(
                    "Brazo 2 automático",
                    fmt_pct(
                        movimiento[
                            "Pct_Automatico_Brazo2"
                        ]
                    ),
                )

                st.caption(
                    "Base de cálculo reconciliada: tiempos individuales de ambos brazos. "
                    "% Auto = (Auto B1 + Auto B2) / "
                    "(Auto B1 + Auto B2 + Manual B1 + Manual B2). "
                    "Los gráficos y KPI consolidados consideran únicamente "
                    "disparos FRENTE (>=45 barrenos)."
                )

                if not considerado_kpi:
                    st.info(
                        f"Este reporte está clasificado como {tipo_disparo} "
                        f"({barrenos_realizados if barrenos_realizados is not None else '-'} barrenos) "
                        "y no se considera en el KPI consolidado de automatización."
                    )

            with col_plano:
                if plano_nav_png is not None:
                    st.image(
                        plano_nav_png,
                        use_container_width=True,
                    )
                else:
                    st.empty()

            col_graf_izq, col_graf, col_graf_der = st.columns(
                [0.08, 0.84, 0.08]
            )

            with col_graf:
                st.image(
                    png_bytes,
                    use_container_width=True,
                )

            st.download_button(
                label="Descargar gráfico PNG",
                data=png_bytes,
                file_name=nombre_png,
                mime="image/png",
                key=f"grafico_{clave}",
                on_click="ignore",
            )

            with st.expander(
                "Validación",
                expanded=False,
            ):
                st.dataframe(
                    validacion[
                        [
                            "Tipo",
                            "Esperado",
                            "Encontrado",
                            "Diferencia",
                            "Estado",
                        ]
                    ],
                    use_container_width=True,
                    hide_index=True,
                    height=220,
                )

                if (validacion["Estado"] == "REVISAR").any():
                    st.warning(
                        "Este ciclo presenta diferencias entre el resumen "
                        "'TIPOS DE BARRENO' y los barrenos extraídos."
                    )
                else:
                    st.success(
                        "La extracción reconcilia con el resumen 'TIPOS DE BARRENO'."
                    )

            with st.expander(
                f"Barrenos extra ({len(extras)})",
                expanded=False,
            ):
                if not extras.empty:
                    columnas_extras = [
                        col
                        for col in [
                            "ID",
                            "Tipo",
                            "Longitud_roca_m",
                            "Beta_grados",
                        ]
                        if col in extras.columns
                    ]

                    st.dataframe(
                        extras[columnas_extras],
                        use_container_width=True,
                        hide_index=True,
                        height=165,
                    )
                else:
                    st.caption(
                        "No se identificaron barrenos extra en este ciclo."
                    )

    # ======================================================
    # CONSOLIDADO
    # ======================================================

    if resultados_validos:
        df_detalle = pd.concat(
            [
                r["detalle"]
                for r
                in resultados_validos
            ],
            ignore_index=True,
        )

        df_resumen = pd.concat(
            [
                r["resumen_ciclo"]
                for r
                in resultados_validos
            ],
            ignore_index=True,
        )

        df_reportes = pd.DataFrame(
            [
                r[
                    "resumen_reporte"
                ]
                for r
                in resultados_validos
            ]
        )

        df_validacion = pd.concat(
            [
                r[
                    "validacion"
                ]
                for r
                in resultados_validos
            ],
            ignore_index=True,
        )

        extras_validos = [
            r["extras"]
            for r
            in resultados_validos
            if not r[
                "extras"
            ].empty
        ]

        df_extras = (
            pd.concat(
                extras_validos,
                ignore_index=True,
            )
            if extras_validos
            else pd.DataFrame(
                columns=df_detalle.columns
            )
        )

        columnas_auto = [
            "Archivo_PDF",
            "Fecha_Inicio",
            "Hora_Inicio",
            "Jumbo",
            "Numero_Serie",
            "Ciclo",
            "Barrenos_Realizados",
            "Fuente_Barrenos_Realizados",
            "Auto_Brazo1_min",
            "Auto_Brazo2_min",
            "Auto_Total_min",
            "Manual_Brazo1_min",
            "Manual_Brazo2_min",
            "Manual_Total_min",
            "Pct_Movimiento_Automatico",
            "Pct_Movimiento_Manual",
            "Pct_Automatico_Brazo1",
            "Pct_Automatico_Brazo2",
            "Pagina_Movimiento_Brazos",
        ]

        df_automatico = (
            df_reportes[
                [
                    c
                    for c
                    in columnas_auto
                    if c
                    in df_reportes.columns
                ]
            ].copy()
        )

        # --------------------------------------------------
        # CLASIFICACIÓN DEL TIPO DE DISPARO
        # --------------------------------------------------

        if "Barrenos_Realizados" in df_automatico.columns:
            df_automatico[
                "Tipo_Disparo"
            ] = df_automatico[
                "Barrenos_Realizados"
            ].apply(
                clasificar_tipo_disparo
            )

            df_automatico[
                "Considerado_KPI_Automatizacion"
            ] = (
                df_automatico[
                    "Tipo_Disparo"
                ] == "FRENTE"
            )
        else:
            df_automatico[
                "Tipo_Disparo"
            ] = "SIN CLASIFICAR"

            df_automatico[
                "Considerado_KPI_Automatizacion"
            ] = False

        # --------------------------------------------------
        # KPI RECONCILIADO DESDE LOS TIEMPOS DE LOS BRAZOS
        # --------------------------------------------------

        columnas_tiempo_brazos = [
            "Auto_Brazo1_min",
            "Auto_Brazo2_min",
            "Manual_Brazo1_min",
            "Manual_Brazo2_min",
        ]

        if all(
            c in df_automatico.columns
            for c in columnas_tiempo_brazos
        ):
            df_automatico[
                "Auto_Total_Brazos_min"
            ] = (
                df_automatico[
                    "Auto_Brazo1_min"
                ]
                +
                df_automatico[
                    "Auto_Brazo2_min"
                ]
            )

            df_automatico[
                "Manual_Total_Brazos_min"
            ] = (
                df_automatico[
                    "Manual_Brazo1_min"
                ]
                +
                df_automatico[
                    "Manual_Brazo2_min"
                ]
            )

            denominador_brazos = (
                df_automatico[
                    "Auto_Total_Brazos_min"
                ]
                +
                df_automatico[
                    "Manual_Total_Brazos_min"
                ]
            )

            df_automatico[
                "Pct_Movimiento_Automatico_Brazos"
            ] = (
                df_automatico[
                    "Auto_Total_Brazos_min"
                ]
                /
                denominador_brazos
                * 100
            ).where(
                denominador_brazos > 0
            )

            df_automatico[
                "Pct_Movimiento_Manual_Brazos"
            ] = (
                df_automatico[
                    "Manual_Total_Brazos_min"
                ]
                /
                denominador_brazos
                * 100
            ).where(
                denominador_brazos > 0
            )

            # Diferencias vs la columna "Suma" reportada por iSURE.
            if "Auto_Total_min" in df_automatico.columns:
                df_automatico[
                    "Dif_Auto_Reporte_vs_Brazos_min"
                ] = (
                    df_automatico[
                        "Auto_Total_min"
                    ]
                    -
                    df_automatico[
                        "Auto_Total_Brazos_min"
                    ]
                )

            if "Manual_Total_min" in df_automatico.columns:
                df_automatico[
                    "Dif_Manual_Reporte_vs_Brazos_min"
                ] = (
                    df_automatico[
                        "Manual_Total_min"
                    ]
                    -
                    df_automatico[
                        "Manual_Total_Brazos_min"
                    ]
                )

        # Gap absoluto por ciclo
        if (
            "Pct_Automatico_Brazo1"
            in df_automatico.columns
            and
            "Pct_Automatico_Brazo2"
            in df_automatico.columns
        ):
            df_automatico[
                "Gap_Automatico_Brazos_pp"
            ] = (
                df_automatico[
                    "Pct_Automatico_Brazo1"
                ]
                -
                df_automatico[
                    "Pct_Automatico_Brazo2"
                ]
            ).abs()

        st.divider()
        st.header(
            "Consolidado"
        )

        # --------------------------------------------------
        # GRÁFICO GENERAL
        # --------------------------------------------------

        st.subheader(
            "Evolución del movimiento automático"
        )

        fig_auto = (
            generar_grafico_tendencia_automatico(
                df_automatico
            )
        )

        tendencia_png = None

        if fig_auto is not None:
            st.pyplot(
                fig_auto,
                use_container_width=True,
            )

            tendencia_png = (
                figura_a_png(
                    fig_auto
                )
            )

            plt.close(
                fig_auto
            )
        else:
            st.info(
                "No se encontraron datos suficientes "
                "de movimiento automático/manual para "
                "generar la tendencia."
            )

        # --------------------------------------------------
        # GRÁFICOS POR EQUIPO: BRAZO 1 VS BRAZO 2
        # --------------------------------------------------

        st.subheader(
            "Uso automático por brazo"
        )

        graficos_brazos_png = []

        jumbos_disponibles = [
            j
            for j
            in sorted(
                df_automatico[
                    "Jumbo"
                ].dropna().unique()
            )
        ]

        for jumbo in jumbos_disponibles:
            fig_brazos = (
                generar_grafico_brazos_por_jumbo(
                    df_automatico,
                    jumbo,
                )
            )

            if fig_brazos is None:
                continue

            st.markdown(
                f"#### {jumbo}"
            )

            st.pyplot(
                fig_brazos,
                use_container_width=True,
            )

            png_brazos = (
                figura_a_png(
                    fig_brazos
                )
            )

            nombre_brazos = (
                f"{jumbo}_"
                "Uso_Automatico_Brazo1_vs_Brazo2.png"
            )

            graficos_brazos_png.append(
                (
                    nombre_brazos,
                    png_brazos,
                )
            )

            plt.close(
                fig_brazos
            )

        # --------------------------------------------------
        # TABLA USO AUTOMÁTICO
        # --------------------------------------------------

        st.markdown(
            "#### Clasificación de disparos"
        )

        if "Tipo_Disparo" in df_automatico.columns:
            resumen_disparos = (
                df_automatico[
                    "Tipo_Disparo"
                ]
                .value_counts(
                    dropna=False
                )
                .rename_axis(
                    "Tipo_Disparo"
                )
                .reset_index(
                    name="N_Reportes"
                )
            )

            st.dataframe(
                resumen_disparos,
                use_container_width=True,
                hide_index=True,
            )

        st.markdown(
            "#### Uso automático por ciclo"
        )

        columnas_vista_auto = [
            c
            for c
            in [
                "Fecha_Inicio",
                "Jumbo",
                "Ciclo",
                "Barrenos_Realizados",
                "Fuente_Barrenos_Realizados",
                "Tipo_Disparo",
                "Considerado_KPI_Automatizacion",
                "Auto_Total_Brazos_min",
                "Manual_Total_Brazos_min",
                "Pct_Movimiento_Automatico_Brazos",
                "Pct_Movimiento_Manual_Brazos",
                "Pct_Automatico_Brazo1",
                "Pct_Automatico_Brazo2",
                "Gap_Automatico_Brazos_pp",
                "Auto_Total_min",
                "Manual_Total_min",
                "Dif_Auto_Reporte_vs_Brazos_min",
                "Dif_Manual_Reporte_vs_Brazos_min",
            ]
            if c
            in df_automatico.columns
        ]

        st.dataframe(
            df_automatico[
                columnas_vista_auto
            ],
            use_container_width=True,
            hide_index=True,
        )

        # --------------------------------------------------
        # EVOLUCIÓN MEDIANA DE BARRENOS CUT
        # --------------------------------------------------

        st.markdown(
            "#### Evolución de la mediana de barrenos Cut"
        )

        fig_cut = generar_grafico_mediana_cut(
            df_resumen
        )

        mediana_cut_png = None

        if fig_cut is not None:
            st.pyplot(
                fig_cut,
                use_container_width=True,
            )

            mediana_cut_png = figura_a_png(
                fig_cut
            )

            plt.close(
                fig_cut
            )
        else:
            st.info(
                "No se encontraron datos suficientes de barrenos Cut "
                "para generar la evolución de la mediana."
            )

        # --------------------------------------------------
        # DESCARGAS
        # --------------------------------------------------

        excel_bytes = crear_excel(
            df_reportes,
            df_resumen,
            df_detalle,
            df_validacion,
            df_extras,
            df_automatico,
        )

        graficos_zip = [
            (
                r[
                    "nombre_png"
                ],
                r[
                    "png_bytes"
                ],
            )
            for r
            in resultados_validos
        ]

        if (
            tendencia_png
            is not None
        ):
            graficos_zip.append(
                (
                    "Evolucion_Movimiento_Automatico.png",
                    tendencia_png,
                )
            )

        if mediana_cut_png is not None:
            graficos_zip.append(
                (
                    "Evolucion_Mediana_Barrenos_Cut.png",
                    mediana_cut_png,
                )
            )

        graficos_zip.extend(
            graficos_brazos_png
        )

        col_excel, col_graficos = (
            st.columns(
                2
            )
        )

        with col_excel:
            st.download_button(
                label="Descargar Excel consolidado",
                data=excel_bytes,
                file_name="EBR_Drill_Consolidado.xlsx",
                mime=(
                    "application/"
                    "vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                type="primary",
                use_container_width=True,
                on_click="ignore",
            )

        with col_graficos:
            zip_bytes = (
                crear_zip_graficos(
                    graficos_zip
                )
            )

            st.download_button(
                label="Descargar todos los gráficos",
                data=zip_bytes,
                file_name="EBR_Drill_Graficos.zip",
                mime="application/zip",
                use_container_width=True,
                on_click="ignore",
            )

    # ======================================================
    # ERRORES
    # ======================================================

    if errores_actuales:
        st.divider()

        st.subheader(
            "Archivos con error"
        )

        st.dataframe(
            pd.DataFrame(
                errores_actuales
            ),
            use_container_width=True,
            hide_index=True,
        )

else:
    st.write(
        "Selecciona uno o varios PDF para iniciar el análisis."
    )
