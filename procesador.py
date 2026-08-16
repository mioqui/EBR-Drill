from __future__ import annotations

from pathlib import Path
import math
import re
from typing import Dict, List, Tuple, Optional

import pdfplumber
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
from matplotlib.patches import Patch


# ==========================================================
# CONFIGURACIÓN
# ==========================================================

ORDEN_TIPOS = [
    "Bottom",
    "Easer",
    "Cut",
    "Contour",
    "Reaming",
    "Casing",
]

JUMBOS = {
    "125D114796": "JUMB001",
    "125D98943": "JUMB002",
}

NUM = r"-?\d+(?:\.\d+)?"


# ==========================================================
# FUNCIONES AUXILIARES
# ==========================================================

def limpiar_texto(texto) -> str:
    if texto is None:
        return ""
    return " ".join(str(texto).split())


def identificar_jumbo(numero_serie: Optional[str]) -> str:
    if not numero_serie:
        return "JUMBO_NO_IDENTIFICADO"
    return JUMBOS.get(numero_serie, "JUMBO_NO_IDENTIFICADO")


def _normalizar_serie(serie_completa: str) -> str:
    # Ejemplo: 125D114796-1 -> 125D114796
    return re.sub(r"-\d+$", "", serie_completa.strip())


# ==========================================================
# METADATOS
# ==========================================================

def leer_metadatos(pdf_path: Path, nombre_archivo: Optional[str] = None) -> Dict:
    datos = {
        "Archivo_PDF": nombre_archivo or pdf_path.name,
        "Ciclo": None,
        "Fecha_Inicio": None,
        "Hora_Inicio": None,
        "Jumbo": None,
        "Numero_Serie": None,
        "Plan_Perforacion": None,
        "Metros_Perforados": None,
        "Barrenos_Realizados": None,
        "Barrenos_Planificados": None,
    }

    with pdfplumber.open(pdf_path) as pdf:
        texto = pdf.pages[0].extract_text() or ""

    m = re.search(r"Ciclo\s+(\d+)", texto, re.IGNORECASE)
    if m:
        datos["Ciclo"] = int(m.group(1))

    m = re.search(
        r"N[º°o]?\s*de\s*serie\s+([A-Za-z0-9-]+)",
        texto,
        re.IGNORECASE,
    )
    if m:
        serie_base = _normalizar_serie(m.group(1))
        datos["Numero_Serie"] = serie_base
        datos["Jumbo"] = identificar_jumbo(serie_base)

    m = re.search(
        r"Iniciado\s+(\d{1,2}/\d{1,2}/\d{4})\s+(\d{2}:\d{2}:\d{2})",
        texto,
        re.IGNORECASE,
    )
    if m:
        datos["Fecha_Inicio"] = m.group(1)
        datos["Hora_Inicio"] = m.group(2)

    m = re.search(r"Plan de perforación\s+([^\n]+)", texto, re.IGNORECASE)
    if m:
        plan = limpiar_texto(m.group(1))
        # Evitar capturar el encabezado repetido si aparece a continuación.
        plan = re.split(r"\s+Plan de bulonaje\b", plan, maxsplit=1, flags=re.IGNORECASE)[0]
        datos["Plan_Perforacion"] = plan

    m = re.search(r"Metros perforados\s*\[m\]\s*([0-9.]+)", texto, re.IGNORECASE)
    if m:
        datos["Metros_Perforados"] = float(m.group(1))

    m = re.search(
        r"Barrenos de perforación\s+en frentes.*?"
        r"Planificado:\s*(\d+).*?"
        r"Realizado:\s*(\d+)",
        texto,
        re.IGNORECASE | re.DOTALL,
    )
    if m:
        datos["Barrenos_Planificados"] = int(m.group(1))
        datos["Barrenos_Realizados"] = int(m.group(2))

    return datos


# ==========================================================
# RESUMEN "TIPOS DE BARRENO"
# ==========================================================

def leer_totales_tipos_barreno(pdf_path: Path) -> Tuple[Dict[str, int], int]:
    totales = {tipo: 0 for tipo in ORDEN_TIPOS}
    pagina_encontrada = None

    with pdfplumber.open(pdf_path) as pdf:
        for numero_pagina, pagina in enumerate(pdf.pages, start=1):
            texto_pagina = pagina.extract_text() or ""

            # Evita confundir el índice con la página real.
            if (
                "TIPOS DE BARRENO" not in texto_pagina
                or "Longitud total en la roca" not in texto_pagina
            ):
                continue

            pagina_encontrada = numero_pagina

            tablas = pagina.extract_tables() or []
            for tabla in tablas:
                for fila in tabla:
                    celdas = [limpiar_texto(celda) for celda in fila]
                    primera = next((c for c in celdas if c), "")

                    if primera not in ORDEN_TIPOS:
                        continue

                    cantidades = []
                    for celda in celdas:
                        m = re.search(r"(\d+)\s+barrenos?", celda, re.IGNORECASE)
                        if m:
                            cantidades.append(int(m.group(1)))

                    if cantidades:
                        # Último conteo de la fila = columna "Suma".
                        totales[primera] = cantidades[-1]

            break

    if pagina_encontrada is None:
        raise ValueError("No se encontró la página real 'TIPOS DE BARRENO'.")

    return totales, pagina_encontrada


# ==========================================================
# PATRONES DE LA TABLA iSURE
# ==========================================================

TIPOS_REGEX = r"Bottom|Easer|Cut|Contour|Casing"

patron_normal = re.compile(
    rf"Torque\s+"
    rf"(?P<boom>\d+)\s+"
    rf"(?P<sec>\d+)\s+"
    rf"(?P<x>{NUM})\s+"
    rf"(?P<y>{NUM})\s+"
    rf"(?P<z>{NUM})\s+"
    rf"(?P<alpha>{NUM})\s+"
    rf"(?P<beta>{NUM})\s+"
    rf"(?P<tilt>{NUM})\s+"
    rf"(?P<depth>{NUM})\s+"
    rf"(?P<length>{NUM})\s+"
    rf"(?P<type>{TIPOS_REGEX})\s*$"
)

# Reaming suele aparecer quebrado como "Reamin g".
patron_reaming = re.compile(
    rf"Torque\s+"
    rf"(?:g\s+)?"
    rf"Reamin\s+"
    rf"(?P<boom>\d+)\s+"
    rf"(?P<sec>\d+)\s+"
    rf"(?P<x>{NUM})\s+"
    rf"(?P<y>{NUM})\s+"
    rf"(?P<z>{NUM})\s+"
    rf"(?P<alpha>{NUM})\s+"
    rf"(?P<beta>{NUM})\s+"
    rf"(?P<tilt>{NUM})\s+"
    rf"(?P<depth>{NUM})\s+"
    rf"(?P<length>{NUM})\s+"
    rf"g\s*$"
)

# Fallback: fila ejecutada que quedó separada del ID planificado al cambiar de página.
patron_continuacion = re.compile(
    rf"^"
    rf"(?P<boom>\d+)\s+"
    rf"(?P<sec>\d+)\s+"
    rf"(?P<x>{NUM})\s+"
    rf"(?P<y>{NUM})\s+"
    rf"(?P<z>{NUM})\s+"
    rf"(?P<alpha>{NUM})\s+"
    rf"(?P<beta>{NUM})\s+"
    rf"(?P<tilt>{NUM})\s+"
    rf"(?P<depth>{NUM})\s+"
    rf"(?P<length>{NUM})\s+"
    rf"(?P<type>{TIPOS_REGEX}|Reaming)\s*$"
)


# ==========================================================
# EXTRAER BARRENOS REALES
# ==========================================================

def extraer_barrenos(
    pdf_path: Path,
    metadatos: Dict,
) -> Tuple[pd.DataFrame, List[int]]:
    registros: List[Dict] = []
    paginas_procesadas: List[int] = []

    # En algunos PDF la primera parte de una fila queda al final de una página.
    id_pendiente = None
    tipo_pendiente = None

    with pdfplumber.open(pdf_path) as pdf:
        for numero_pagina, pagina in enumerate(pdf.pages, start=1):
            texto_pagina = pagina.extract_text() or ""

            if "BARRENOS DE BULONAJE" in texto_pagina:
                continue

            if "ID de" not in texto_pagina or "barreno" not in texto_pagina:
                continue

            paginas_procesadas.append(numero_pagina)

            tablas = pagina.extract_tables() or []

            for tabla in tablas:
                for fila in tabla:
                    texto = limpiar_texto(
                        " ".join(c for c in fila if c is not None)
                    )
                    if not texto:
                        continue

                    # Si una fila contiene solo un ID, conservarlo para una posible
                    # continuación en la siguiente fila/página.
                    if re.fullmatch(r"(?:E\d+|\d+)", texto):
                        id_pendiente = texto
                        continue

                    # Intento de capturar el ID/tipo planificado para continuidad.
                    # No se usa como dato ejecutado; solo como ayuda para fila partida.
                    m_plan = re.match(
                        rf"^(?P<id>E\d+|\d+)\s+.*?\s(?P<type>"
                        rf"Bottom|Easer|Cut|Contour|Casing|Reamin(?:\s+g)?|Reaming)"
                        rf"(?:\s+Torque|\s+-|$)",
                        texto,
                        re.IGNORECASE,
                    )
                    if m_plan:
                        id_pendiente = m_plan.group("id")
                        t = m_plan.group("type")
                        tipo_pendiente = (
                            "Reaming" if t.lower().startswith("reamin") else t.title()
                        )

                    match = patron_normal.search(texto)
                    tipo = None
                    id_barreno = None
                    datos = None

                    if match:
                        datos = match.groupdict()
                        tipo = datos["type"]
                        id_barreno = texto.split()[0]

                    if match is None:
                        match = patron_reaming.search(texto)
                        if match:
                            datos = match.groupdict()
                            tipo = "Reaming"
                            id_barreno = texto.split()[0]

                    if match is None:
                        match = patron_continuacion.fullmatch(texto)
                        if match and id_pendiente:
                            datos = match.groupdict()
                            tipo_detectado = datos["type"]
                            if tipo_detectado.lower().startswith("reamin"):
                                tipo_detectado = "Reaming"

                            tipo = tipo_detectado or tipo_pendiente
                            id_barreno = id_pendiente
                            id_pendiente = None
                            tipo_pendiente = None

                    if match is None or datos is None:
                        continue

                    longitud = float(datos["length"])
                    beta = float(datos["beta"])

                    extra = str(id_barreno).upper().startswith("E")

                    longitud_axial = longitud * math.cos(math.radians(beta))

                    registros.append(
                        {
                            "Archivo_PDF": metadatos["Archivo_PDF"],
                            "Ciclo": metadatos["Ciclo"],
                            "Fecha_Inicio": metadatos["Fecha_Inicio"],
                            "Hora_Inicio": metadatos["Hora_Inicio"],
                            "Jumbo": metadatos["Jumbo"],
                            "Numero_Serie": metadatos["Numero_Serie"],
                            "Plan_Perforacion": metadatos["Plan_Perforacion"],
                            "ID": str(id_barreno),
                            "Tipo": tipo,
                            "Longitud_roca_m": longitud,
                            "Beta_grados": beta,
                            "Longitud_axial_m": longitud_axial,
                            "Extra": extra,
                            "Pagina_PDF": numero_pagina,
                        }
                    )

    df = pd.DataFrame(registros)

    if df.empty:
        return df, paginas_procesadas

    df = df.drop_duplicates(
        subset=["Archivo_PDF", "ID", "Tipo", "Longitud_roca_m"]
    ).copy()

    df["Tipo"] = pd.Categorical(
        df["Tipo"],
        categories=ORDEN_TIPOS,
        ordered=True,
    )

    df = df.sort_values(["Tipo", "ID"]).reset_index(drop=True)

    return df, paginas_procesadas


# ==========================================================
# VALIDACIÓN / RESÚMENES
# ==========================================================

def construir_validacion(
    df: pd.DataFrame,
    esperados: Dict[str, int],
    metadatos: Dict,
) -> pd.DataFrame:
    filas = []

    for tipo in ORDEN_TIPOS:
        esperado = int(esperados.get(tipo, 0))
        encontrado = int((df["Tipo"] == tipo).sum()) if not df.empty else 0

        filas.append(
            {
                "Archivo_PDF": metadatos["Archivo_PDF"],
                "Fecha_Inicio": metadatos["Fecha_Inicio"],
                "Ciclo": metadatos["Ciclo"],
                "Jumbo": metadatos["Jumbo"],
                "Numero_Serie": metadatos["Numero_Serie"],
                "Tipo": tipo,
                "Esperado": esperado,
                "Encontrado": encontrado,
                "Diferencia": encontrado - esperado,
                "Estado": "OK" if encontrado == esperado else "REVISAR",
            }
        )

    return pd.DataFrame(filas)


def construir_resumen_ciclo(
    df: pd.DataFrame,
    esperados: Dict[str, int],
    metadatos: Dict,
) -> pd.DataFrame:
    filas = []

    for tipo in ORDEN_TIPOS:
        grupo = df[df["Tipo"] == tipo]
        if grupo.empty:
            continue

        esperado = int(esperados.get(tipo, 0))
        n = len(grupo)

        filas.append(
            {
                "Archivo_PDF": metadatos["Archivo_PDF"],
                "Fecha_Inicio": metadatos["Fecha_Inicio"],
                "Hora_Inicio": metadatos["Hora_Inicio"],
                "Ciclo": metadatos["Ciclo"],
                "Jumbo": metadatos["Jumbo"],
                "Numero_Serie": metadatos["Numero_Serie"],
                "Plan_Perforacion": metadatos["Plan_Perforacion"],
                "Tipo": tipo,
                "N": n,
                "Min": grupo["Longitud_roca_m"].min(),
                "Max": grupo["Longitud_roca_m"].max(),
                "Promedio": grupo["Longitud_roca_m"].mean(),
                "Mediana": grupo["Longitud_roca_m"].median(),
                "Esperado": esperado,
                "Estado": "OK" if n == esperado else "REVISAR",
            }
        )

    return pd.DataFrame(filas)


def construir_resumen_reporte(
    metadatos: Dict,
    esperados: Dict[str, int],
    df: pd.DataFrame,
    pagina_tipos: int,
    paginas_detalle: List[int],
) -> Dict:
    total_esperado = int(sum(esperados.values()))
    total_encontrado = int(len(df))

    return {
        **metadatos,
        "Pagina_Tipos_Barreno": pagina_tipos,
        "Paginas_Detalle": ", ".join(map(str, paginas_detalle)),
        "Total_Tipos_Reporte": total_esperado,
        "Total_Extraido": total_encontrado,
        "Diferencia": total_encontrado - total_esperado,
        "Estado": "OK" if total_encontrado == total_esperado else "REVISAR",
    }


# ==========================================================
# GRÁFICO
# ==========================================================

def generar_grafico(
    df: pd.DataFrame,
    metadatos: Dict,
):
    if df.empty:
        raise ValueError("No hay datos para generar el gráfico.")

    tipos_grafico = [
        tipo for tipo in ORDEN_TIPOS
        if not df[df["Tipo"] == tipo].empty
    ]

    fig, ax = plt.subplots(figsize=(14, 8))

    datos_boxplot = [
        df.loc[df["Tipo"] == tipo, "Longitud_roca_m"].values
        for tipo in tipos_grafico
    ]

    ax.boxplot(
        datos_boxplot,
        tick_labels=[
            f"{tipo}\n(n={len(df[df['Tipo'] == tipo])})"
            for tipo in tipos_grafico
        ],
        widths=0.50,
        showmeans=False,
        showfliers=False,
        whis=(0, 100),
        medianprops={"color": "tab:orange", "linewidth": 1.5},
        boxprops={"color": "black"},
        whiskerprops={"color": "black", "linewidth": 1.2},
        capprops={"color": "black", "linewidth": 1.2},
    )

    # Puntos individuales:
    # - valor Y único: centrado
    # - valores repetidos: distribuidos simétricamente
    for posicion, tipo in enumerate(tipos_grafico, start=1):
        grupo = df[df["Tipo"] == tipo].copy()

        for _, subgrupo in grupo.groupby("Longitud_roca_m", sort=True):
            n = len(subgrupo)
            if n == 1:
                offsets = [0.0]
            else:
                offsets = np.linspace(-0.06, 0.06, n)

            for offset, (_, fila) in zip(offsets, subgrupo.iterrows()):
                x = posicion + offset
                y = fila["Longitud_roca_m"]

                if fila["Extra"]:
                    ax.scatter(
                        x, y,
                        s=95,
                        facecolor="yellow",
                        edgecolor="black",
                        linewidth=1.5,
                        zorder=5,
                    )
                    ax.annotate(
                        f"{fila['ID']} extra",
                        xy=(x, y),
                        xytext=(8, 0),
                        textcoords="offset points",
                        va="center",
                        fontsize=9,
                    )
                else:
                    ax.scatter(
                        x, y,
                        s=60,
                        facecolor="black",
                        edgecolor="gray",
                        linewidth=0.8,
                        zorder=4,
                    )

    # Estadísticos
    for posicion, tipo in enumerate(tipos_grafico, start=1):
        valores = df.loc[df["Tipo"] == tipo, "Longitud_roca_m"]
        minimo = valores.min()
        maximo = valores.max()
        promedio = valores.mean()
        mediana = valores.median()

        ax.text(
            posicion,
            maximo + 0.08,
            (
                f"Min {minimo:.2f} | Máx {maximo:.2f}\n"
                f"Prom {promedio:.2f} | Med {mediana:.2f}"
            ),
            ha="center",
            va="bottom",
            fontsize=9,
        )

    leyenda = [
        Line2D(
            [0], [0],
            marker="o",
            linestyle="None",
            markersize=7,
            markerfacecolor="black",
            markeredgecolor="black",
            label="Punto: valor individual de cada barreno",
        ),
        Patch(
            facecolor="white",
            edgecolor="black",
            label="Caja: 50% central de los datos (Q1-Q3)",
        ),
        Line2D(
            [0], [0],
            linestyle="-",
            linewidth=2,
            color="tab:orange",
            label="Línea dentro de la caja: mediana",
        ),
        Line2D(
            [0], [0],
            linestyle="-",
            linewidth=1.2,
            color="black",
            label="Bigotes: mínimo y máximo real",
        ),
        Line2D(
            [0], [0],
            marker="o",
            markerfacecolor="yellow",
            markeredgecolor="black",
            linestyle="None",
            markersize=9,
            label="Punto resaltado: barreno extra (no programado)",
        ),
    ]

    ax.legend(
        handles=leyenda,
        loc="lower left",
        fontsize=9,
        title="Leyenda del gráfico",
    )

    min_global = df["Longitud_roca_m"].min()
    max_global = df["Longitud_roca_m"].max()
    rango = max_global - min_global

    margen_inf = max(0.15, rango * 0.08)
    margen_sup = max(0.25, rango * 0.12)

    ax.set_ylim(min_global - margen_inf, max_global + margen_sup)

    jumbo = metadatos.get("Jumbo") or "JUMBO"
    serie = metadatos.get("Numero_Serie") or "-"
    ciclo = metadatos.get("Ciclo") or "-"
    fecha = metadatos.get("Fecha_Inicio") or "-"

    ax.set_title(
        "Distribución de longitud perforada en roca por tipo de barreno\n"
        f"{jumbo} | Serie {serie} | Ciclo {ciclo} | {fecha}",
        fontsize=14,
    )

    ax.set_xlabel("Tipo de barreno")
    ax.set_ylabel("Longitud perforada en roca (m)")
    ax.grid(True, alpha=0.5)

    fig.tight_layout()
    return fig


# ==========================================================
# FUNCIÓN PRINCIPAL PARA STREAMLIT
# ==========================================================

def procesar_pdf(
    pdf_path: Path,
    nombre_archivo: Optional[str] = None,
) -> Dict:
    metadatos = leer_metadatos(pdf_path, nombre_archivo=nombre_archivo)
    esperados, pagina_tipos = leer_totales_tipos_barreno(pdf_path)
    df, paginas_detalle = extraer_barrenos(pdf_path, metadatos)

    if df.empty:
        raise ValueError(
            "No se encontraron barrenos ejecutados en la tabla de detalle."
        )

    validacion = construir_validacion(df, esperados, metadatos)
    resumen_ciclo = construir_resumen_ciclo(df, esperados, metadatos)
    resumen_reporte = construir_resumen_reporte(
        metadatos,
        esperados,
        df,
        pagina_tipos,
        paginas_detalle,
    )
    extras = df[df["Extra"]].copy()
    fig = generar_grafico(df, metadatos)

    return {
        "metadata": metadatos,
        "esperados": esperados,
        "detalle": df,
        "validacion": validacion,
        "resumen_ciclo": resumen_ciclo,
        "resumen_reporte": resumen_reporte,
        "extras": extras,
        "fig": fig,
    }
