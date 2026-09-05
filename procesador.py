from __future__ import annotations

from pathlib import Path
from io import BytesIO
import math
import re
import struct
import zipfile
import unicodedata
from datetime import datetime, timezone
from typing import Dict, List, Tuple, Optional

import pdfplumber
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
from matplotlib.patches import Patch

ORDEN_TIPOS = ["Bottom", "Easer", "Cut", "Contour", "Reaming", "Casing"]
JUMBOS = {"125D114796": "JUMB001", "125D98943": "JUMB002"}
NUM = r"-?\d+(?:\.\d+)?"

VERSION_PROCESADOR = "V34.46-Python-MASIVO-OPERADORES-ZDA-NORMALIZADOS"


def limpiar_texto(texto) -> str:
    if texto is None:
        return ""
    return " ".join(str(texto).split())


def identificar_jumbo(numero_serie: Optional[str]) -> str:
    """
    Devuelve el alias conocido del jumbo.

    Para una serie nueva que todavía no tenga alias corporativo asignado,
    se conserva la identidad del equipo como "Serie <número>" para evitar
    agrupar varios jumbos distintos bajo JUMBO_NO_IDENTIFICADO.
    """
    if not numero_serie:
        return "JUMBO_NO_IDENTIFICADO"

    serie = str(numero_serie).strip()
    return JUMBOS.get(serie, f"Serie {serie}")


def _normalizar_serie(serie_completa: str) -> str:
    return re.sub(r"-\d+$", "", serie_completa.strip())


def tiempo_a_minutos(valor: Optional[str]) -> Optional[float]:
    if not valor:
        return None
    m = re.fullmatch(r"\s*(\d+):(\d{2})\s*", valor)
    if not m:
        return None
    return int(m.group(1)) * 60 + int(m.group(2))


def pct(parte: Optional[float], total: Optional[float]) -> Optional[float]:
    if parte is None or total is None or total <= 0:
        return None
    return (parte / total) * 100.0


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
        "Fuente_Barrenos_Realizados": None,
        "Operario": None,
    }

    textos_paginas = []

    with pdfplumber.open(pdf_path) as pdf:
        for pagina in pdf.pages[:4]:
            textos_paginas.append(
                pagina.extract_text() or ""
            )

    # BD-PERFO: OPERADOR = código numérico del campo "Operario"
    # ubicado en la página física 2 del PDF.
    if len(textos_paginas) >= 2:
        texto_pagina_2 = limpiar_texto(textos_paginas[1])
        m_operario = re.search(
            r"\bOperario\b\s*:?\s*(\d+)",
            texto_pagina_2,
            re.IGNORECASE,
        )
        if m_operario:
            datos["Operario"] = m_operario.group(1)

    texto = "\n".join(textos_paginas)
    texto_normalizado = re.sub(
        r"\s+",
        " ",
        texto,
    )

    m = re.search(
        r"Ciclo\s*:?\s*(\d+)",
        texto_normalizado,
        re.IGNORECASE,
    )
    if m:
        datos["Ciclo"] = int(m.group(1))

    m = re.search(
        r"N[º°o]?\s*de\s*serie\s*:?\s*([A-Za-z0-9-]+)",
        texto_normalizado,
        re.IGNORECASE,
    )
    if m:
        serie_base = _normalizar_serie(
            m.group(1)
        )
        datos["Numero_Serie"] = serie_base
        datos["Jumbo"] = identificar_jumbo(
            serie_base
        )

    m = re.search(
        r"Iniciado\s*:?\s*(\d{1,2}/\d{1,2}/\d{4})\s+(\d{2}:\d{2}:\d{2})",
        texto_normalizado,
        re.IGNORECASE,
    )
    if m:
        datos["Fecha_Inicio"] = m.group(1)
        datos["Hora_Inicio"] = m.group(2)

    m = re.search(
        r"Plan\s+de\s+perforaci[oó]n\s*:?\s*(.+?)(?=\s+Plan\s+de\s+bulonaje\b|\s+Metros\s+perforados\b|$)",
        texto_normalizado,
        re.IGNORECASE,
    )
    if m:
        datos["Plan_Perforacion"] = limpiar_texto(
            m.group(1)
        )

    m = re.search(
        r"Metros\s+perforados\s*\[m\]\s*:?\s*([0-9]+(?:\.[0-9]+)?)",
        texto_normalizado,
        re.IGNORECASE,
    )
    if m:
        datos["Metros_Perforados"] = float(
            m.group(1)
        )

    # ------------------------------------------------------
    # Barrenos de perforación en frentes
    # ------------------------------------------------------
    # Se toleran saltos de línea y pequeñas variaciones
    # en la extracción de texto de iSURE.
    patrones_inicio = [
        r"Barrenos?\s+de\s+perforaci[oó]n\s+en\s+frentes?",
        r"Barrenos?\s+de\s+perforaci[oó]n.*?\bfrentes?\b",
        r"Barrenos?.{0,80}\bfrentes?\b",
    ]

    bloque = None

    for patron in patrones_inicio:
        m_inicio = re.search(
            patron,
            texto_normalizado,
            re.IGNORECASE,
        )

        if m_inicio:
            bloque = texto_normalizado[
                m_inicio.start():
                m_inicio.start() + 800
            ]
            break

    if bloque:
        m_plan = re.search(
            r"Planificad[oa]\s*:?\s*(\d+)",
            bloque,
            re.IGNORECASE,
        )

        m_real = re.search(
            r"Realizad[oa]\s*:?\s*(\d+)",
            bloque,
            re.IGNORECASE,
        )

        if m_plan:
            datos["Barrenos_Planificados"] = int(
                m_plan.group(1)
            )

        if m_real:
            datos["Barrenos_Realizados"] = int(
                m_real.group(1)
            )
            datos[
                "Fuente_Barrenos_Realizados"
            ] = (
                "Reporte iSURE - "
                "barrenos de perforación en frentes"
            )

    return datos

def leer_movimiento_brazos(pdf_path: Path) -> Dict:
    resultado = {
        "Auto_Brazo1_min": None,
        "Auto_Brazo2_min": None,
        "Auto_Total_min": None,
        "Manual_Brazo1_min": None,
        "Manual_Brazo2_min": None,
        "Manual_Total_min": None,
        "Pct_Movimiento_Automatico": None,
        "Pct_Movimiento_Manual": None,
        "Pct_Automatico_Brazo1": None,
        "Pct_Automatico_Brazo2": None,
        "Pagina_Movimiento_Brazos": None,
    }

    patron_auto = re.compile(
        r"Autom[aá]tico\s*\[h\]\s+(?P<b1>\d+:\d{2})\s+(?P<b2>\d+:\d{2})\s+(?P<suma>\d+:\d{2})(?:\s+(?P<prom>\d+:\d{2}))?",
        re.IGNORECASE,
    )
    patron_manual = re.compile(
        r"Manual\s*\[h\]\s+(?P<b1>\d+:\d{2})\s+(?P<b2>\d+:\d{2})\s+(?P<suma>\d+:\d{2})(?:\s+(?P<prom>\d+:\d{2}))?",
        re.IGNORECASE,
    )

    with pdfplumber.open(pdf_path) as pdf:
        for numero_pagina, pagina in enumerate(pdf.pages, start=1):
            texto = pagina.extract_text() or ""
            if "Tiempo de movimiento del brazo" not in texto:
                continue

            m_auto = patron_auto.search(texto)
            m_manual = patron_manual.search(texto)
            if not (m_auto and m_manual):
                continue

            auto_b1 = tiempo_a_minutos(m_auto.group("b1"))
            auto_b2 = tiempo_a_minutos(m_auto.group("b2"))
            auto_total = tiempo_a_minutos(m_auto.group("suma"))
            manual_b1 = tiempo_a_minutos(m_manual.group("b1"))
            manual_b2 = tiempo_a_minutos(m_manual.group("b2"))
            manual_total = tiempo_a_minutos(m_manual.group("suma"))

            total_mov = auto_total + manual_total if auto_total is not None and manual_total is not None else None
            total_b1 = auto_b1 + manual_b1 if auto_b1 is not None and manual_b1 is not None else None
            total_b2 = auto_b2 + manual_b2 if auto_b2 is not None and manual_b2 is not None else None

            resultado.update({
                "Auto_Brazo1_min": auto_b1,
                "Auto_Brazo2_min": auto_b2,
                "Auto_Total_min": auto_total,
                "Manual_Brazo1_min": manual_b1,
                "Manual_Brazo2_min": manual_b2,
                "Manual_Total_min": manual_total,
                "Pct_Movimiento_Automatico": pct(auto_total, total_mov),
                "Pct_Movimiento_Manual": pct(manual_total, total_mov),
                "Pct_Automatico_Brazo1": pct(auto_b1, total_b1),
                "Pct_Automatico_Brazo2": pct(auto_b2, total_b2),
                "Pagina_Movimiento_Brazos": numero_pagina,
            })
            break

    return resultado


def leer_totales_tipos_barreno(pdf_path: Path) -> Tuple[Dict[str, int], int]:
    totales = {tipo: 0 for tipo in ORDEN_TIPOS}
    pagina_encontrada = None

    with pdfplumber.open(pdf_path) as pdf:
        for numero_pagina, pagina in enumerate(pdf.pages, start=1):
            texto_pagina = pagina.extract_text() or ""
            if "TIPOS DE BARRENO" not in texto_pagina or "Longitud total en la roca" not in texto_pagina:
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
                        totales[primera] = cantidades[-1]
            break

    if pagina_encontrada is None:
        raise ValueError("No se encontró la página real 'TIPOS DE BARRENO'.")

    return totales, pagina_encontrada


TIPOS_REGEX = r"Bottom|Easer|Cut|Contour|Casing"
patron_normal = re.compile(
    rf"Torque\s+(?P<boom>\d+)\s+(?P<sec>\d+)\s+(?P<x>{NUM})\s+(?P<y>{NUM})\s+(?P<z>{NUM})\s+(?P<alpha>{NUM})\s+(?P<beta>{NUM})\s+(?P<tilt>{NUM})\s+(?P<depth>{NUM})\s+(?P<length>{NUM})\s+(?P<type>{TIPOS_REGEX})\s*$"
)
patron_reaming = re.compile(
    rf"Torque\s+(?:g\s+)?Reamin\s+(?P<boom>\d+)\s+(?P<sec>\d+)\s+(?P<x>{NUM})\s+(?P<y>{NUM})\s+(?P<z>{NUM})\s+(?P<alpha>{NUM})\s+(?P<beta>{NUM})\s+(?P<tilt>{NUM})\s+(?P<depth>{NUM})\s+(?P<length>{NUM})\s+g\s*$"
)
patron_continuacion = re.compile(
    rf"^(?P<boom>\d+)\s+(?P<sec>\d+)\s+(?P<x>{NUM})\s+(?P<y>{NUM})\s+(?P<z>{NUM})\s+(?P<alpha>{NUM})\s+(?P<beta>{NUM})\s+(?P<tilt>{NUM})\s+(?P<depth>{NUM})\s+(?P<length>{NUM})\s+(?P<type>{TIPOS_REGEX}|Reaming)\s*$"
)


def extraer_barrenos(pdf_path: Path, metadatos: Dict) -> Tuple[pd.DataFrame, List[int]]:
    registros: List[Dict] = []
    paginas_procesadas: List[int] = []
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
                    texto = limpiar_texto(" ".join(c for c in fila if c is not None))
                    if not texto:
                        continue

                    if re.fullmatch(r"(?:E\d+|\d+)", texto):
                        id_pendiente = texto
                        continue

                    m_plan = re.match(
                        rf"^(?P<id>E\d+|\d+)\s+.*?\s(?P<type>Bottom|Easer|Cut|Contour|Casing|Reamin(?:\s+g)?|Reaming)(?:\s+Torque|\s+-|$)",
                        texto,
                        re.IGNORECASE,
                    )
                    if m_plan:
                        id_pendiente = m_plan.group("id")
                        t = m_plan.group("type")
                        tipo_pendiente = "Reaming" if t.lower().startswith("reamin") else t.title()

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

                    registros.append({
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
                    })

    df = pd.DataFrame(registros)
    if df.empty:
        return df, paginas_procesadas

    df = df.drop_duplicates(subset=["Archivo_PDF", "ID", "Tipo", "Longitud_roca_m"]).copy()
    df["Tipo"] = pd.Categorical(df["Tipo"], categories=ORDEN_TIPOS, ordered=True)
    df = df.sort_values(["Tipo", "ID"]).reset_index(drop=True)
    return df, paginas_procesadas


def construir_validacion(df: pd.DataFrame, esperados: Dict[str, int], metadatos: Dict) -> pd.DataFrame:
    filas = []
    for tipo in ORDEN_TIPOS:
        esperado = int(esperados.get(tipo, 0))
        encontrado = int((df["Tipo"] == tipo).sum()) if not df.empty else 0
        filas.append({
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
        })
    return pd.DataFrame(filas)


def construir_resumen_ciclo(df: pd.DataFrame, esperados: Dict[str, int], metadatos: Dict) -> pd.DataFrame:
    filas = []
    for tipo in ORDEN_TIPOS:
        grupo = df[df["Tipo"] == tipo]
        if grupo.empty:
            continue
        esperado = int(esperados.get(tipo, 0))
        n = len(grupo)
        filas.append({
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
        })
    return pd.DataFrame(filas)


def construir_resumen_reporte(metadatos: Dict, esperados: Dict[str, int], df: pd.DataFrame, pagina_tipos: int, paginas_detalle: List[int], movimiento: Dict) -> Dict:
    total_esperado = int(sum(esperados.values()))
    total_encontrado = int(len(df))
    return {
        **metadatos,
        **movimiento,
        "Pagina_Tipos_Barreno": pagina_tipos,
        "Paginas_Detalle": ", ".join(map(str, paginas_detalle)),
        "Total_Tipos_Reporte": total_esperado,
        "Total_Extraido": total_encontrado,
        "Diferencia": total_encontrado - total_esperado,
        "Estado": "OK" if total_encontrado == total_esperado else "REVISAR",
    }




def generar_grafico(df: pd.DataFrame, metadatos: Dict):
    """
    Versión BEESWARM SIMPLE.

    Los puntos se distribuyen horizontalmente solo cuando hay
    riesgo de solape visual. La posición Y permanece intacta.

    La lógica usa carriles simétricos:
    0, +1, -1, +2, -2...
    """
    if df.empty:
        raise ValueError("No hay datos para generar el gráfico.")

    tipos_grafico = [
        tipo
        for tipo in ORDEN_TIPOS
        if not df[df["Tipo"] == tipo].empty
    ]

    fig, ax = plt.subplots(figsize=(10.8, 5.35), dpi=160)

    datos_boxplot = [
        df.loc[
            df["Tipo"] == tipo,
            "Longitud_roca_m",
        ].values
        for tipo in tipos_grafico
    ]

    ax.boxplot(
        datos_boxplot,
        tick_labels=[
            f"{tipo}\n(n={len(df[df['Tipo'] == tipo])})"
            for tipo in tipos_grafico
        ],
        widths=0.44,
        showmeans=False,
        showfliers=False,
        whis=(0, 100),
        medianprops={
            "color": "tab:orange",
            "linewidth": 1.5,
        },
        boxprops={
            "color": "black",
        },
        whiskerprops={
            "color": "black",
            "linewidth": 1.0,
        },
        capprops={
            "color": "black",
            "linewidth": 1.0,
        },
    )

    def carriles_beeswarm(
        valores_y,
        separacion_y,
        paso_x=0.038,
        max_x=0.16,
    ):
        """
        Asigna un desplazamiento horizontal a cada Y evitando
        que puntos cercanos queden en el mismo carril.
        """
        indices = np.argsort(
            valores_y
        )

        offsets = np.zeros(
            len(valores_y),
            dtype=float,
        )

        asignados = []

        secuencia = [0]

        for k in range(1, 20):
            secuencia.extend(
                [
                    k,
                    -k,
                ]
            )

        for idx in indices:
            y = valores_y[idx]

            vecinos = [
                (
                    y_prev,
                    lane_prev,
                )
                for (
                    y_prev,
                    lane_prev,
                )
                in asignados
                if abs(
                    y
                    - y_prev
                )
                < separacion_y
            ]

            usados = {
                lane_prev
                for (
                    _,
                    lane_prev,
                )
                in vecinos
            }

            lane = next(
                lane_candidate
                for lane_candidate
                in secuencia
                if lane_candidate
                not in usados
            )

            x_offset = (
                lane
                * paso_x
            )

            x_offset = max(
                -max_x,
                min(
                    max_x,
                    x_offset,
                ),
            )

            offsets[idx] = (
                x_offset
            )

            asignados.append(
                (
                    y,
                    lane,
                )
            )

        return offsets

    min_global = df[
        "Longitud_roca_m"
    ].min()

    max_global = df[
        "Longitud_roca_m"
    ].max()

    rango_global = max(
        0.5,
        max_global
        - min_global,
    )

    # Distancia vertical aproximada a partir de la cual
    # visualmente los marcadores empiezan a tocarse.
    separacion_y = max(
        0.045,
        rango_global * 0.017,
    )

    for posicion, tipo in enumerate(
        tipos_grafico,
        start=1,
    ):
        grupo = df[
            df["Tipo"] == tipo
        ].copy()

        grupo = grupo.sort_values(
            [
                "Longitud_roca_m",
                "Extra",
                "ID",
            ]
        ).reset_index(
            drop=True
        )

        valores_y = grupo[
            "Longitud_roca_m"
        ].to_numpy(
            dtype=float
        )

        offsets = carriles_beeswarm(
            valores_y,
            separacion_y=separacion_y,
            paso_x=0.040,
            max_x=0.16,
        )

        for i, fila in grupo.iterrows():
            x = (
                posicion
                + offsets[i]
            )

            y = fila[
                "Longitud_roca_m"
            ]

            if fila["Extra"]:
                ax.scatter(
                    x,
                    y,
                    s=58,
                    facecolor="yellow",
                    edgecolor=(
                        0,
                        0,
                        0,
                        0.90,
                    ),
                    linewidth=0.45,
                    alpha=0.95,
                    zorder=5,
                )

                ax.annotate(
                    f"{fila['ID']} extra",
                    xy=(x, y),
                    xytext=(4, 0),
                    textcoords="offset points",
                    va="center",
                    ha="left",
                    fontsize=6.1,
                    color="#2b2b2b",
                )

            else:
                ax.scatter(
                    x,
                    y,
                    s=38,
                    facecolor=(
                        0,
                        0,
                        0,
                        0.58,
                    ),
                    edgecolor=(
                        0.0,
                        0.45,
                        0.12,
                        0.85,
                    ),
                    linewidth=0.35,
                    zorder=4,
                )

    for posicion, tipo in enumerate(
        tipos_grafico,
        start=1,
    ):
        valores = df.loc[
            df["Tipo"] == tipo,
            "Longitud_roca_m",
        ]

        ax.text(
            posicion,
            valores.max() + 0.08,
            (
                f"Min {valores.min():.2f} | "
                f"Máx {valores.max():.2f}\n"
                f"Prom {valores.mean():.2f} | "
                f"Med {valores.median():.2f}"
            ),
            ha="center",
            va="bottom",
            fontsize=7.0,
        )

    leyenda = [
        Line2D(
            [0],
            [0],
            marker="o",
            linestyle="None",
            markersize=6.4,
            markerfacecolor=(
                0,
                0,
                0,
                0.58,
            ),
            markeredgecolor=(
                0.0,
                0.45,
                0.12,
                0.85,
            ),
            markeredgewidth=0.35,
            label="Punto: valor de cada barreno",
        ),
        Patch(
            facecolor="white",
            edgecolor="black",
            label="Caja: 50% central de los datos (Q1-Q3)",
        ),
        Line2D(
            [0],
            [0],
            linestyle="-",
            linewidth=1.6,
            color="tab:orange",
            label="Mediana",
        ),
        Line2D(
            [0],
            [0],
            linestyle="-",
            linewidth=1.0,
            color="black",
            label="Bigotes: mínimo y máximo real",
        ),
        Line2D(
            [0],
            [0],
            marker="o",
            linestyle="None",
            markersize=7.0,
            markerfacecolor="yellow",
            markeredgecolor="black",
            markeredgewidth=0.45,
            label="Barreno extra (no programado)",
        ),
    ]

    ax.legend(
        handles=leyenda,
        loc="lower left",
        fontsize=6.5,
        title="Leyenda",
        title_fontsize=6.5,
        framealpha=0.95,
        borderpad=0.5,
        handletextpad=0.6,
        labelspacing=0.35,
    )

    rango = (
        max_global
        - min_global
    )

    ax.set_ylim(
        min_global
        - max(
            0.15,
            rango * 0.08,
        ),
        max_global
        + max(
            0.25,
            rango * 0.12,
        ),
    )

    ax.set_title(
        "Distribución de longitud perforada en roca por tipo de barreno\n"
        f"{metadatos.get('Jumbo') or 'JUMBO'} | "
        f"Serie {metadatos.get('Numero_Serie') or '-'} | "
        f"Ciclo {metadatos.get('Ciclo') or '-'} | "
        f"{metadatos.get('Fecha_Inicio') or '-'}",
        fontsize=10.5,
    )

    ax.set_xlabel(
        "Tipo de barreno",
        fontsize=8.0,
    )

    ax.set_ylabel(
        "Longitud perforada en roca (m)",
        fontsize=8.0,
    )

    ax.tick_params(
        axis="both",
        labelsize=7.4,
    )

    ax.grid(
        True,
        alpha=0.35,
    )

    fig.tight_layout()
    return fig



def extraer_plano_navegacion_png(pdf_path: Path, resolution: int = 170) -> Optional[bytes]:
    """
    Extrae de la primera hoja la imagen "Barrenos perforados, Plano de navegación"
    para mostrarla como miniatura en la cabecera del reporte.

    Se usa un recorte relativo porque la ubicación del plano es bastante
    consistente en los reportes iSURE revisados.
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages:
                return None

            page = pdf.pages[0]

            bbox = (
                page.width * 0.515,
                page.height * 0.150,
                page.width * 0.958,
                page.height * 0.482,
            )

            recorte = page.crop(bbox)
            imagen = recorte.to_image(resolution=resolution)

            buffer = BytesIO()
            imagen.save(buffer, format='PNG')
            buffer.seek(0)
            return buffer.getvalue()
    except Exception:
        return None

def procesar_pdf(
    pdf_path: Path,
    nombre_archivo: Optional[str] = None,
    generar_visuales: bool = True,
) -> Dict:
    metadatos = leer_metadatos(pdf_path, nombre_archivo=nombre_archivo)
    movimiento = leer_movimiento_brazos(pdf_path)
    esperados, pagina_tipos = leer_totales_tipos_barreno(pdf_path)
    df, paginas_detalle = extraer_barrenos(pdf_path, metadatos)

    if df.empty:
        raise ValueError("No se encontraron barrenos ejecutados en la tabla de detalle.")

    # Fallback para clasificación del disparo:
    # si el PDF no permite leer claramente el campo
    # "Barrenos de perforación en frentes - Realizado",
    # se usa la misma regla validada para ZDA:
    # Bottom + Easer + Cut + Contour.
    # Reaming y Casing NO forman parte de Barrenos_Realizados.
    if metadatos.get("Barrenos_Realizados") is None:
        tipos_frente = {"Bottom", "Easer", "Cut", "Contour"}
        metadatos["Barrenos_Realizados"] = int(
            df["Tipo"].astype(str).isin(tipos_frente).sum()
        )

        metadatos[
            "Fuente_Barrenos_Realizados"
        ] = "Detalle extraído - Bottom/Easer/Cut/Contour"

    validacion = construir_validacion(df, esperados, metadatos)
    resumen_ciclo = construir_resumen_ciclo(df, esperados, metadatos)
    resumen_reporte = construir_resumen_reporte(metadatos, esperados, df, pagina_tipos, paginas_detalle, movimiento)
    extras = df[df["Extra"]].copy()
    fig = generar_grafico(df, metadatos) if generar_visuales else None
    plano_nav_png = extraer_plano_navegacion_png(pdf_path) if generar_visuales else None

    return {
        "metadata": metadatos,
        "movimiento": movimiento,
        "esperados": esperados,
        "detalle": df,
        "validacion": validacion,
        "resumen_ciclo": resumen_ciclo,
        "resumen_reporte": resumen_reporte,
        "extras": extras,
        "fig": fig,
        "plano_nav_png": plano_nav_png,
        "plano_nav_origen": "PDF",
    }


# ==========================================================
# EXTENSION V33: SALIDA ESTANDAR + LECTOR ZDA
# ==========================================================

TIPOS_ZDA = ORDEN_TIPOS
ZDA_TIPO_CODES = {
    0: "Reaming",
    1: "Contour",
    4: "Cut",
    5: "Easer",
    8: "Bottom",
    9: "Casing",
}


def clasificar_tipo_disparo_v33(barrenos_realizados):
    """Criterio exacto de la HTML V33: FRENTE >45; SELLADA 25-45; resto <25."""
    if barrenos_realizados is None or pd.isna(barrenos_realizados):
        return "SIN CLASIFICAR"
    n = int(barrenos_realizados)
    if n > 45:
        return "FRENTE"
    if n >= 25:
        return "SELLADA"
    return "ESTOCADA Y/O CORRECCIONES"


def _enriquecer_resultado_estandar(resultado: Dict, fuente: str) -> Dict:
    """Normaliza la salida de PDF/ZDA para que app_v33 consuma el mismo esquema."""
    metadata = dict(resultado.get("metadata") or {})
    movimiento = dict(resultado.get("movimiento") or {})
    resumen_reporte = dict(resultado.get("resumen_reporte") or {})
    detalle = resultado.get("detalle")
    validacion = resultado.get("validacion")
    resumen_ciclo = resultado.get("resumen_ciclo")
    extras = resultado.get("extras")

    if detalle is None:
        detalle = pd.DataFrame()
    if validacion is None:
        validacion = pd.DataFrame()
    if resumen_ciclo is None:
        resumen_ciclo = pd.DataFrame()
    if extras is None:
        extras = pd.DataFrame()

    metadata["Fuente"] = fuente
    metadata["Archivo_Fuente"] = metadata.get("Archivo_PDF") or metadata.get("Archivo_ZDA")

    auto_b1 = movimiento.get("Auto_Brazo1_min")
    auto_b2 = movimiento.get("Auto_Brazo2_min")
    man_b1 = movimiento.get("Manual_Brazo1_min")
    man_b2 = movimiento.get("Manual_Brazo2_min")
    vals = [auto_b1, auto_b2, man_b1, man_b2]
    if all(v is not None and not pd.isna(v) for v in vals):
        auto_total = float(auto_b1) + float(auto_b2)
        manual_total = float(man_b1) + float(man_b2)
        den = auto_total + manual_total
        pct_auto = auto_total / den * 100 if den > 0 else None
        pct_manual = manual_total / den * 100 if den > 0 else None
    else:
        auto_total = manual_total = pct_auto = pct_manual = None

    barrenos = metadata.get("Barrenos_Realizados")
    tipo = clasificar_tipo_disparo_v33(barrenos)
    conteo_ok = str(resumen_reporte.get("Estado") or "REVISAR") == "OK"

    # La versión PDF histórica no reconcilia metros por tipo. Se conserva N/A
    # en vez de inventar una validación que el parser no soporta.
    estado_metros = resumen_reporte.get("Estado_Metros_Tipos") or "N/A"
    if estado_metros == "REVISAR":
        lectura_ok = False
    else:
        lectura_ok = conteo_ok

    resumen_reporte.update(metadata)
    resumen_reporte.update(movimiento)
    resumen_reporte.update({
        "Fuente": fuente,
        "Archivo_Fuente": metadata.get("Archivo_Fuente"),
        "Tipo_Disparo": tipo,
        "Considerado_KPI_Automatizacion": tipo == "FRENTE",
        "Auto_Total_Brazos_min": auto_total,
        "Manual_Total_Brazos_min": manual_total,
        "Pct_Movimiento_Automatico_Brazos": pct_auto,
        "Pct_Movimiento_Manual_Brazos": pct_manual,
        "Gap_Automatico_Brazos_pp": (
            abs(float(movimiento.get("Pct_Automatico_Brazo1")) - float(movimiento.get("Pct_Automatico_Brazo2")))
            if movimiento.get("Pct_Automatico_Brazo1") is not None and movimiento.get("Pct_Automatico_Brazo2") is not None
            else None
        ),
        "Estado_Conteo": "OK" if conteo_ok else "REVISAR",
        "Estado_Metros_Tipos": estado_metros,
        "Lectura_Confiable": "OK" if lectura_ok else "REVISAR",
    })

    atipicos = detalle[
        (pd.to_numeric(detalle.get("Longitud_roca_m"), errors="coerce") < 0.20)
        | (pd.to_numeric(detalle.get("Longitud_roca_m"), errors="coerce") > 7.00)
    ].copy() if not detalle.empty and "Longitud_roca_m" in detalle.columns else pd.DataFrame()
    if not atipicos.empty:
        atipicos["Clasificacion_Atipico"] = atipicos["Longitud_roca_m"].apply(
            lambda x: "Barreno largo" if float(x) > 7 else "Longitud muy corta"
        )
        atipicos["Tratamiento"] = f"Dato {fuente} conservado; excluido sólo del boxplot estándar"

    for df in (detalle, validacion, resumen_ciclo, extras, atipicos):
        if not df.empty:
            if "Fuente" not in df.columns:
                df["Fuente"] = fuente
            if "Archivo_Fuente" not in df.columns:
                df["Archivo_Fuente"] = metadata.get("Archivo_Fuente")

    resultado.update({
        "metadata": metadata,
        "movimiento": movimiento,
        "resumen_reporte": resumen_reporte,
        "detalle": detalle,
        "validacion": validacion,
        "validacion_metros": resultado.get("validacion_metros", pd.DataFrame()),
        "resumen_ciclo": resumen_ciclo,
        "extras": extras,
        "atipicos": atipicos,
        "mwd_barrenos": resultado.get("mwd_barrenos", pd.DataFrame()),
        "fuente": fuente,
    })
    return resultado


def procesar_pdf_v33(
    pdf_path: Path,
    nombre_archivo: Optional[str] = None,
    generar_visuales: bool = True,
) -> Dict:
    """Usa el parser PDF Python existente y normaliza su salida al esquema V33."""
    return _enriquecer_resultado_estandar(
        procesar_pdf(
            pdf_path,
            nombre_archivo=nombre_archivo,
            generar_visuales=generar_visuales,
        ),
        "PDF",
    )


def _zda_kv(texto: str) -> Dict[str, str]:
    out = {}
    for line in str(texto or "").splitlines():
        if ":" not in line:
            continue
        k, v = line.split(":", 1)
        out[k.strip()] = v.strip()
    return out


CATALOGO_OPERADORES_ZDA = {
    "RIVERA": "Josue Rivera",
    "CELIS": "Nilton Celis",
    "CASAS": "Abraham Casas",
    "SOLIS": "Roy Solis",
    "OSORIO": "John Osorio",
    "CUCHULA": "Rogelio Cuchula",
}


def _zda_texto_normalizado(valor: Optional[str]) -> str:
    texto = str(valor or "").strip()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(
        c for c in texto
        if not unicodedata.combining(c)
    )
    return re.sub(r"\s+", " ", texto).strip().upper()


def _zda_normalizar_operador(valor: Optional[str]) -> Optional[str]:
    """
    Normaliza un nombre/apellido contra el catálogo conocido.
    Si OP: trae un valor nuevo, conserva el texto detectado.
    """
    texto = str(valor or "").strip().strip(",;:-")
    if not texto:
        return None

    norm = _zda_texto_normalizado(texto)

    for apellido, nombre_completo in CATALOGO_OPERADORES_ZDA.items():
        if re.search(rf"\b{re.escape(apellido)}\b", norm):
            return nombre_completo

    return " ".join(
        palabra.capitalize()
        for palabra in re.split(r"\s+", texto)
        if palabra
    ) or None


def _zda_operador_desde_tunnel_id(tunnel_id: Optional[str]) -> Optional[str]:
    """
    Extrae el operador desde tunnel_id / ID Auxiliar de round.txt.

    Prioridad:
      1) Campo explícito OP:...
      2) Apellido conocido en cualquier parte del texto

    Ejemplos:
      GL:898 NV:4055 OP:RIVERA T:N -> Josue Rivera
      nv 4055 Gl. 7939w Celis     -> Nilton Celis
    """
    texto = str(tunnel_id or "").strip()
    if not texto:
        return None

    m = re.search(
        r"(?:^|\s)OP\s*:\s*(.+?)(?=\s+(?:T|TURN|NV|VN|GL|RMR|B|BLOCK)\s*:|$)",
        texto,
        re.IGNORECASE,
    )
    if m:
        operador = _zda_normalizar_operador(m.group(1))
        if operador:
            return operador

    norm = _zda_texto_normalizado(texto)
    for apellido, nombre_completo in CATALOGO_OPERADORES_ZDA.items():
        if re.search(rf"\b{re.escape(apellido)}\b", norm):
            return nombre_completo

    return None


def _zda_parse_ts(valor: Optional[str]) -> Optional[int]:
    if not valor:
        return None
    try:
        dt = datetime.strptime(valor.strip(), "%Y-%m-%d %H:%M:%S").replace(tzinfo=timezone.utc)
        return int(dt.timestamp())
    except Exception:
        return None


def _zda_fmt_date(ts: Optional[int]) -> Optional[str]:
    if ts is None:
        return None
    return datetime.fromtimestamp(int(ts), tz=timezone.utc).strftime("%d/%m/%Y")


def _zda_fmt_time(ts: Optional[int]) -> Optional[str]:
    if ts is None:
        return None
    return datetime.fromtimestamp(int(ts), tz=timezone.utc).strftime("%H:%M:%S")


def _zda_fmt_datetime(ts: Optional[int]) -> Optional[str]:
    if ts is None:
        return None
    return datetime.fromtimestamp(int(ts), tz=timezone.utc).strftime("%d/%m/%Y %H:%M:%S")


def _zda_duration_hms(sec: Optional[float]) -> Optional[str]:
    if sec is None or not np.isfinite(sec) or sec < 0:
        return None
    sec = int(round(sec))
    h, rem = divmod(sec, 3600)
    m, s = divmod(rem, 60)
    return f"{h:02d}:{m:02d}:{s:02d}"


def _zda_base_serie(rig: Optional[str]) -> Optional[str]:
    if not rig:
        return None
    return re.sub(r"-(?:\d+|L)$", "", rig, flags=re.IGNORECASE)


def _zda_ascii(data: bytes, start: int, length: int) -> str:
    raw = data[start:start + length].split(b"\x00", 1)[0]
    return "".join(chr(b) for b in raw if 32 <= b <= 126).strip()


def _parse_zda_boom(data: bytes, nombre_archivo: str, metadata: Dict) -> Tuple[pd.DataFrame, Dict]:
    record_size = 297
    first = 4
    rows = []
    raw_records = invalid = 0
    unknown = {}

    for start in range(first, len(data) - record_size + 1, record_size):
        raw_records += 1
        try:
            boom0 = data[start + 159]
            sec = data[start + 160]
            start_ts = struct.unpack_from("<I", data, start + 163)[0]
            end_ts = struct.unpack_from("<I", data, start + 240)[0]
            type_code = data[start + 175]
            tipo = ZDA_TIPO_CODES.get(type_code)
            x, y, z = struct.unpack_from("<ddd", data, start + 183)
            x2, y2, z2 = struct.unpack_from("<ddd", data, start + 257)
            length = math.sqrt((x2-x)**2 + (y2-y)**2 + (z2-z)**2)

            ts_ok = 1577836800 < start_ts < 2051222400 and start_ts <= end_ts < 2051222400
            geom_ok = all(np.isfinite(v) for v in [x,y,z,x2,y2,z2,length]) and 0.10 < length < 20
            boom_ok = boom0 in (0,1) and 0 < sec < 100
            if not tipo:
                if ts_ok and geom_ok and boom_ok:
                    unknown[type_code] = unknown.get(type_code, 0) + 1
                continue
            if not (ts_ok and geom_ok and boom_ok):
                invalid += 1
                continue

            boom = boom0 + 1
            ident = _zda_ascii(data, start + 26, 30)
            if not ident:
                ident = f"R{boom}-{sec}" if tipo == "Reaming" else f"S{boom}-{sec}"
            length2 = round(length, 2)
            depth2 = round(y2, 2)
            rows.append({
                "Fuente": "ZDA",
                "Archivo_Fuente": nombre_archivo,
                "Archivo_ZDA": nombre_archivo,
                "Archivo_PDF": None,
                "Ciclo": metadata.get("Ciclo"),
                "Fecha_Inicio": metadata.get("Fecha_Inicio"),
                "Hora_Inicio": metadata.get("Hora_Inicio"),
                "Jumbo": metadata.get("Jumbo"),
                "Numero_Serie": metadata.get("Numero_Serie"),
                "Plan_Perforacion": metadata.get("Plan_Perforacion"),
                "ID": str(ident),
                "Tipo": tipo,
                "Boom": boom,
                "Secuencia": int(sec),
                "X": round(x,2), "Y": round(y,2), "Z": round(z,2),
                "X2": round(x2,2), "Y2": round(y2,2), "Z2": round(z2,2),
                "Alpha_grados": None, "Beta_grados": None, "Tilt_grados": None,
                "Profundidad_m": depth2,
                "Longitud_roca_m": length2,
                "Longitud_axial_m": length2,
                "Extra": str(ident).upper().startswith("E"),
                "Pagina_PDF": None,
                "Fuente_Parser": "ZDA boom.dat",
                "Inicio_Barreno_TS": int(start_ts),
                "Fin_Barreno_TS": int(end_ts),
                "Inicio_Barreno": _zda_fmt_datetime(start_ts),
                "Fin_Barreno": _zda_fmt_datetime(end_ts),
                "Codigo_Tipo_ZDA": int(type_code),
            })
        except Exception:
            invalid += 1

    df = pd.DataFrame(rows)
    if not df.empty:
        df["Tipo"] = pd.Categorical(df["Tipo"], categories=ORDEN_TIPOS, ordered=True)
        df = df.sort_values(["Tipo", "ID"]).reset_index(drop=True)
    diag = {
        "boom_records": raw_records,
        "boom_valid": len(df),
        "boom_invalid": invalid,
        "unknown_codes": [{"Codigo": k, "N": v} for k,v in sorted(unknown.items())],
    }
    return df, diag


def _parse_zda_counters(data: bytes) -> Tuple[Dict, Dict]:
    header = 25
    expected_values = 464
    if len(data) < header + expected_values * 8:
        return {}, {"ok": False, "motivo": f"counters.dat corto: {len(data)} bytes"}

    def val(i):
        return struct.unpack_from("<d", data, header + i*8)[0]

    def h_to_min(i):
        h = val(i)
        return int(round(h * 60)) if np.isfinite(h) and 0 <= h < 24 else None

    ab1, mb1, ab2, mb2 = h_to_min(43), h_to_min(47), h_to_min(131), h_to_min(135)
    if any(v is None for v in [ab1, mb1, ab2, mb2]):
        return {}, {"ok": False, "motivo": "No se pudieron recuperar los cuatro tiempos de movimiento."}
    auto = ab1 + ab2
    manual = mb1 + mb2
    total = auto + manual
    mov = {
        "Auto_Brazo1_min": ab1, "Auto_Brazo2_min": ab2, "Auto_Total_min": auto,
        "Manual_Brazo1_min": mb1, "Manual_Brazo2_min": mb2, "Manual_Total_min": manual,
        "Pct_Movimiento_Automatico": auto/total*100 if total else None,
        "Pct_Movimiento_Manual": manual/total*100 if total else None,
        "Pct_Automatico_Brazo1": ab1/(ab1+mb1)*100 if (ab1+mb1) else None,
        "Pct_Automatico_Brazo2": ab2/(ab2+mb2)*100 if (ab2+mb2) else None,
        "Pagina_Movimiento_Brazos": None,
    }
    return mov, {"ok": True}


def _zda_validacion(df: pd.DataFrame, metadata: Dict) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    valid_rows = []
    meter_rows = []
    resumen = []
    for tipo in ORDEN_TIPOS:
        g = df[df["Tipo"].astype(str) == tipo].copy() if not df.empty else pd.DataFrame()
        n = len(g)
        metros = float(g["Longitud_roca_m"].sum()) if n else 0.0
        valid_rows.append({
            **metadata, "Tipo": tipo, "Esperado": n, "Encontrado": n,
            "Diferencia": 0, "Estado": "OK", "Fuente": "ZDA",
        })
        meter_rows.append({
            **metadata, "Tipo": tipo, "N": n, "Metros_Reporte_m": round(metros,2),
            "Metros_Extraidos_m": round(metros,2), "Diferencia_m": 0.0,
            "Tolerancia_m": 0.35, "Estado": "OK", "Fuente": "ZDA",
        })
        if n:
            vals = pd.to_numeric(g["Longitud_roca_m"], errors="coerce")
            resumen.append({
                "Archivo_PDF": None, "Archivo_ZDA": metadata.get("Archivo_ZDA"),
                "Archivo_Fuente": metadata.get("Archivo_Fuente"), "Fuente": "ZDA",
                "Fecha_Inicio": metadata.get("Fecha_Inicio"), "Hora_Inicio": metadata.get("Hora_Inicio"),
                "Ciclo": metadata.get("Ciclo"), "Jumbo": metadata.get("Jumbo"),
                "Numero_Serie": metadata.get("Numero_Serie"), "Plan_Perforacion": metadata.get("Plan_Perforacion"),
                "Tipo": tipo, "N": n, "Min": vals.min(), "Max": vals.max(),
                "Promedio": vals.mean(), "Mediana": vals.median(), "Esperado": n,
                "Estado": "OK", "Metros_Reporte_m": round(metros,2),
                "Metros_Extraidos_m": round(metros,2), "Diferencia_Metros_m": 0.0,
                "Estado_Metros": "OK", "Estado_Reporte": "OK",
            })
    return pd.DataFrame(valid_rows), pd.DataFrame(meter_rows), pd.DataFrame(resumen)


def _parse_zda_mwd(zf: zipfile.ZipFile, names: List[str], metadata: Dict) -> Tuple[pd.DataFrame, Dict]:
    pat = re.compile(r"-mwd-(\d+)-(\d+)\.dat$", re.IGNORECASE)
    mwd_names = [n for n in names if pat.search(n)]
    mwd_names.sort(key=lambda n: tuple(map(int, pat.search(n).groups())))
    rows = []
    diffs = []
    first_ts = last_ts = None
    total_m = 0.0
    samples_total = valid_holes = short = empty = bad_layout = 0

    for name in mwd_names:
        m = pat.search(name)
        boom = int(m.group(1)) + 1
        seq = int(m.group(2))
        data = zf.read(name)
        if len(data) < 131:
            bad_layout += 1
            continue
        payload = len(data) - 131
        if payload % 122 != 0:
            bad_layout += 1
        nrec = payload // 122
        h_first = h_last = max_pos = prev_pos = None
        samples = 0
        for i in range(nrec):
            off = 131 + i*122
            if off + 122 > len(data):
                break
            lo, hi = struct.unpack_from("<II", data, off)
            ts = lo + hi * 4294967296
            pos = struct.unpack_from("<f", data, off + 8)[0]
            if not (1577836800 < ts < 2051222400 and np.isfinite(pos) and 0 <= pos < 30):
                continue
            samples += 1
            samples_total += 1
            h_first = ts if h_first is None else min(h_first, ts)
            h_last = ts if h_last is None else max(h_last, ts)
            max_pos = pos if max_pos is None else max(max_pos, pos)
            if prev_pos is not None:
                d = pos - prev_pos
                if 0.002 < d < 0.20:
                    diffs.append(d)
            prev_pos = pos

        status = "Sin muestras"
        if max_pos is not None and max_pos > 0:
            total_m += max_pos
            if max_pos > 1:
                valid_holes += 1; status = "Completo"
            else:
                short += 1; status = "Intento corto"
            first_ts = h_first if first_ts is None else min(first_ts, h_first)
            last_ts = h_last if last_ts is None else max(last_ts, h_last)
        else:
            empty += 1
        rows.append({
            "Fuente": "ZDA", "Archivo_ZDA": metadata.get("Archivo_ZDA"),
            "Jumbo": metadata.get("Jumbo"), "Ciclo": metadata.get("Ciclo"),
            "Brazo": boom, "Secuencia": seq, "Muestras_MWD": samples,
            "Profundidad_Max_MWD_m": round(max_pos,4) if max_pos is not None else None,
            "Inicio_MWD": _zda_fmt_datetime(h_first), "Fin_MWD": _zda_fmt_datetime(h_last),
            "Duracion_MWD_s": (h_last-h_first) if h_first is not None and h_last is not None else None,
            "Estado_MWD": status, "Archivo_Interno": name,
        })
    paso = float(np.median(diffs)) if diffs else None
    return pd.DataFrame(rows), {
        "first_ts": first_ts, "last_ts": last_ts, "archivos_mwd": len(mwd_names),
        "muestras_mwd": samples_total, "barrenos_mwd_validos": valid_holes,
        "intentos_mwd_cortos": short, "archivos_mwd_vacios": empty,
        "layouts_mwd_no_estandar": bad_layout, "metros_mwd": round(total_m,3),
        "paso_mwd_mediana_m": paso,
    }



def _convex_hull_2d(points):
    """Convex hull 2D por monotone chain, sin dependencias externas."""
    pts = sorted(set((float(x), float(y)) for x, y in points))
    if len(pts) <= 1:
        return pts

    def cross(o, a, b):
        return (
            (a[0] - o[0]) * (b[1] - o[1])
            - (a[1] - o[1]) * (b[0] - o[0])
        )

    lower = []
    for p in pts:
        while len(lower) >= 2 and cross(lower[-2], lower[-1], p) <= 0:
            lower.pop()
        lower.append(p)

    upper = []
    for p in reversed(pts):
        while len(upper) >= 2 and cross(upper[-2], upper[-1], p) <= 0:
            upper.pop()
        upper.append(p)

    return lower[:-1] + upper[:-1]




def _seccion_desde_plan(plan_perforacion: Optional[str]) -> Tuple[float, float, str]:
    """
    Detecta la sección nominal desde Plan_Perforacion.

    Ejemplos admitidos:
      MALLA E.E. 4.5x4.5 III-B
      MALLA E.E. 5.0x5.0 III-A
      4,5 X 4,5

    Retorna:
      ancho_m, alto_m, etiqueta

    Si no se detecta una sección, usa 4.5 x 4.5 m como referencia
    conservadora para no dejar el gráfico sin contorno.
    """
    texto = str(plan_perforacion or "")
    m = re.search(
        r"(\d+(?:[.,]\d+)?)\s*[xX×]\s*(\d+(?:[.,]\d+)?)",
        texto,
    )

    if not m:
        return 4.5, 4.5, "4.5 x 4.5 (referencia)"

    ancho = float(m.group(1).replace(",", "."))
    alto = float(m.group(2).replace(",", "."))

    # Evitar geometrías absurdas por una lectura defectuosa.
    if not (2.0 <= ancho <= 10.0 and 2.0 <= alto <= 10.0):
        return 4.5, 4.5, "4.5 x 4.5 (referencia)"

    return ancho, alto, f"{ancho:.1f} x {alto:.1f}"


def generar_plano_zda_png(
    detalle: pd.DataFrame,
    metadata: Dict,
    resolution: int = 180,
) -> Optional[bytes]:
    """
    Versión 2:
    - Usa un contorno fijo/referencial del frente.
    - Superpone los puntos y segmentos reconstruidos desde ZDA.
    - Mantiene escala y encuadre constantes entre ciclos.

    Proyección usada:
      X -> horizontal del frente
      Z -> vertical del frente
      Y -> profundidad longitudinal (no se grafica directamente)
    """
    requeridas = {"X", "Z", "X2", "Z2", "Tipo"}
    if detalle is None or detalle.empty or not requeridas.issubset(detalle.columns):
        return None

    work = detalle.copy()

    for c in ["X", "Z", "X2", "Z2"]:
        work[c] = pd.to_numeric(work[c], errors="coerce")

    work = work.dropna(subset=["X", "Z", "X2", "Z2"]).copy()
    if work.empty:
        return None

    fig, ax = plt.subplots(figsize=(5.2, 5.2), dpi=resolution)
    ax.set_facecolor("white")
    ax.set_axisbelow(True)

    # Sin numeración visible: estilo más cercano al PDF.
    ax.tick_params(
        axis="both",
        which="both",
        length=0,
        labelbottom=False,
        labelleft=False,
    )

    for spine in ax.spines.values():
        spine.set_visible(False)

    # Ejes de referencia.
    ax.axhline(0, color="#707070", linewidth=0.75, alpha=0.85, zorder=1)
    ax.axvline(0, color="#707070", linewidth=0.75, alpha=0.85, zorder=1)

    # Flechas del sistema de referencia, parecidas al PDF.
    ax.annotate(
        "",
        xy=(0.55, 0.0),
        xytext=(0.0, 0.0),
        arrowprops=dict(arrowstyle="->", color="#111111", lw=0.9),
        zorder=2,
    )
    ax.annotate(
        "",
        xy=(0.0, 0.55),
        xytext=(0.0, 0.0),
        arrowprops=dict(arrowstyle="->", color="#111111", lw=0.9),
        zorder=2,
    )

    # ------------------------------------------------------
    # Contorno dinámico según la sección nominal del plan
    # ------------------------------------------------------
    ancho_seccion, alto_seccion, etiqueta_seccion = _seccion_desde_plan(
        metadata.get("Plan_Perforacion")
    )

    # Calibración visual contra los planos PDF:
    # el contorno dibujado queda ligeramente dentro del ancho nominal.
    #
    # 4.5 x 4.5 -> laterales aprox. ±2.15 m
    #               (dentro del bloque 5, cerca del límite del bloque 4)
    # 5.0 x 5.0 -> laterales aprox. ±2.40 m
    #               (dentro del bloque 5, próximo al límite exterior)
    #
    # Para otras secciones se conserva el mismo retiro visual
    # aproximado de 0.10 m por lado respecto al ancho nominal.
    if abs(ancho_seccion - 4.5) < 0.06:
        half_width_visual = 2.15
    elif abs(ancho_seccion - 5.0) < 0.06:
        half_width_visual = 2.40
    else:
        half_width_visual = max(0.5, ancho_seccion / 2.0 - 0.10)

    left_x = -half_width_visual
    right_x = half_width_visual
    base_z = 0.0

    # Calibración visual del borde superior según referencia PDF:
    # 4.5 x 4.5 -> arriba en el bloque 9, un poco antes de llegar al 10
    # 5.0 x 5.0 -> arriba en el bloque 10, un poco antes de llegar al 11
    if abs(ancho_seccion - 4.5) < 0.06 and abs(alto_seccion - 4.5) < 0.06:
        crown_top_z = 4.42
    elif abs(ancho_seccion - 5.0) < 0.06 and abs(alto_seccion - 5.0) < 0.06:
        crown_top_z = 4.92
    else:
        crown_top_z = max(0.5, alto_seccion - 0.08)

    # Calibración separada del contorno superior:
    # - corner_rx controla hasta dónde llega el borde superior plano.
    # - corner_rz controla la transición vertical de la esquina.
    #
    # Con esto se ajusta mejor lo que se observa en el PDF:
    # 4.5 x 4.5 -> borde superior más largo
    # 5.0 x 5.0 -> borde superior también más largo, sin mover laterales.
    if abs(ancho_seccion - 4.5) < 0.06 and abs(alto_seccion - 4.5) < 0.06:
        corner_rx = 0.48   # top flat hasta aprox. ±1.67 m
        corner_rz = 0.55
    elif abs(ancho_seccion - 5.0) < 0.06 and abs(alto_seccion - 5.0) < 0.06:
        corner_rx = 0.55   # top flat hasta aprox. ±1.85 m
        corner_rz = 0.70
    else:
        corner_rx = min(
            0.75,
            max(
                0.40,
                min(ancho_seccion, alto_seccion) * 0.11,
            ),
        )
        corner_rz = min(
            0.85,
            max(
                0.50,
                min(ancho_seccion, alto_seccion) * 0.13,
            ),
        )

    wall_top_z = crown_top_z - corner_rz

    # Laterales
    ax.plot(
        [left_x, left_x],
        [base_z, wall_top_z],
        color="#303030",
        linewidth=0.95,
        alpha=0.95,
        zorder=2,
    )
    ax.plot(
        [right_x, right_x],
        [base_z, wall_top_z],
        color="#303030",
        linewidth=0.95,
        alpha=0.95,
        zorder=2,
    )

    # Base
    ax.plot(
        [left_x, right_x],
        [base_z, base_z],
        color="#303030",
        linewidth=0.95,
        alpha=0.95,
        zorder=2,
    )

    # ------------------------------------------------------
    # Encuadre dinámico: mostrar siempre todos los barrenos
    # ------------------------------------------------------
    # Mantener un encuadre estándar, pero expandirlo cuando
    # algún barreno quede más arriba, más abajo o más afuera.
    default_x_half = 3.6
    default_y_min = -2.2
    default_y_max = 5.6
    margin_x = 0.35
    margin_y = 0.35

    data_x = pd.concat([work["X"], work["X2"]], ignore_index=True).dropna()
    data_z = pd.concat([work["Z"], work["Z2"]], ignore_index=True).dropna()

    x_candidates = [left_x, right_x]
    z_candidates = [base_z, crown_top_z]
    if not data_x.empty:
        x_candidates.extend(data_x.tolist())
    if not data_z.empty:
        z_candidates.extend(data_z.tolist())

    max_abs_x = max(abs(float(v)) for v in x_candidates) + margin_x
    x_half = max(default_x_half, np.ceil(max_abs_x / 0.5) * 0.5)

    z_min = min(float(v) for v in z_candidates) - margin_y
    z_max = max(float(v) for v in z_candidates) + margin_y
    y_min = min(default_y_min, np.floor(z_min / 0.5) * 0.5)
    y_max = max(default_y_max, np.ceil(z_max / 0.5) * 0.5)

    xlim = (-x_half, x_half)
    ylim = (y_min, y_max)
    xticks = np.arange(np.floor(xlim[0] / 0.5) * 0.5, np.ceil(xlim[1] / 0.5) * 0.5 + 0.001, 0.5)
    yticks = np.arange(np.floor(ylim[0] / 0.5) * 0.5, np.ceil(ylim[1] / 0.5) * 0.5 + 0.001, 0.5)

    ax.set_xlim(*xlim)
    ax.set_ylim(*ylim)
    ax.set_aspect("equal", adjustable="box")
    ax.set_xticks(xticks)
    ax.set_yticks(yticks)
    ax.grid(True, color="#b8cbe6", linewidth=0.6, alpha=0.9)

    # Corona similar al PDF:
    # tramo superior casi plano + esquinas elípticas
    # para controlar mejor hasta dónde llega el borde superior.
    theta_left = np.linspace(np.pi, np.pi / 2.0, 60)
    cx_left = left_x + corner_rx
    cz_left = wall_top_z
    x_left_arc = cx_left + corner_rx * np.cos(theta_left)
    z_left_arc = cz_left + corner_rz * np.sin(theta_left)

    ax.plot(
        x_left_arc,
        z_left_arc,
        color="#303030",
        linewidth=0.95,
        alpha=0.95,
        zorder=2,
    )

    ax.plot(
        [left_x + corner_rx, right_x - corner_rx],
        [crown_top_z, crown_top_z],
        color="#303030",
        linewidth=0.95,
        alpha=0.95,
        zorder=2,
    )

    theta_right = np.linspace(np.pi / 2.0, 0.0, 60)
    cx_right = right_x - corner_rx
    cz_right = wall_top_z
    x_right_arc = cx_right + corner_rx * np.cos(theta_right)
    z_right_arc = cz_right + corner_rz * np.sin(theta_right)

    ax.plot(
        x_right_arc,
        z_right_arc,
        color="#303030",
        linewidth=0.95,
        alpha=0.95,
        zorder=2,
    )

    # Segmentos y puntos del ZDA superpuestos.
    for _, r in work.iterrows():
        x0, z0 = float(r["X"]), float(r["Z"])
        x1, z1 = float(r["X2"]), float(r["Z2"])

        # Segmento
        ax.plot(
            [x0, x1],
            [z0, z1],
            color="#ff4b4b",
            linewidth=0.8,
            alpha=0.95,
            zorder=3,
        )

        # Collar
        ax.scatter(
            [x0],
            [z0],
            s=18,
            facecolor="#ff0000",
            edgecolor="#111111",
            linewidth=0.35,
            zorder=4,
        )

        # Extremo
        ax.scatter(
            [x1],
            [z1],
            s=7,
            facecolor="#ff4b4b",
            edgecolor="none",
            zorder=3,
        )

    # Título similar al PDF.
    ax.set_title(
        "Barrenos perforados, Plano de navegación",
        fontsize=10,
        fontstyle="italic",
        loc="left",
        pad=6,
    )

    # Subnota técnica discreta.
    fig.text(
        0.5,
        0.012,
        f"Plano reconstruido desde ZDA · Sección {etiqueta_seccion}",
        ha="center",
        va="bottom",
        fontsize=6.7,
        color="#667085",
    )

    fig.tight_layout(rect=[0.02, 0.03, 1, 1])

    buffer = BytesIO()
    fig.savefig(
        buffer,
        format="PNG",
        dpi=resolution,
        bbox_inches="tight",
        facecolor="white",
    )
    plt.close(fig)
    buffer.seek(0)
    return buffer.getvalue()



def procesar_zda(
    zda_path: Path,
    nombre_archivo: Optional[str] = None,
    generar_visuales: bool = True,
) -> Dict:
    nombre = nombre_archivo or zda_path.name
    with zipfile.ZipFile(zda_path, "r") as zf:
        names = [n for n in zf.namelist() if not n.endswith("/")]
        round_name = next((n for n in names if re.match(r"^round-.*\.txt$", n, re.I) and not re.search(r"hole_comment", n, re.I)), None)
        if not round_name:
            raise ValueError("El ZDA no contiene el archivo round-*.txt esperado.")
        kv = _zda_kv(zf.read(round_name).decode("utf-8", errors="replace"))
        if not kv.get("rig") or not kv.get("round"):
            raise ValueError("El round.txt no contiene rig/round; formato ZDA no reconocido.")

        serie = _zda_base_serie(kv.get("rig"))
        operador_zda = _zda_operador_desde_tunnel_id(
            kv.get("tunnel_id")
        )
        nav_ts = _zda_parse_ts(kv.get("navigation"))
        decl_start = _zda_parse_ts(kv.get("start"))
        decl_end = _zda_parse_ts(kv.get("end"))
        cycle_start = nav_ts if nav_ts is not None else decl_start
        metadata = {
            "Fuente": "ZDA", "Archivo_Fuente": nombre, "Archivo_ZDA": nombre, "Archivo_PDF": None,
            "Ciclo": int(kv["round"]), "Fecha_Inicio": _zda_fmt_date(cycle_start),
            "Hora_Inicio": _zda_fmt_time(cycle_start), "Numero_Serie": serie,
            "Jumbo": identificar_jumbo(serie), "Plan_Perforacion": kv.get("drill_plan") or None,
            "Operador_ZDA": operador_zda,
            "Operador": operador_zda,
            "Operador_ZDA_Raw": kv.get("tunnel_id") or None,
            "Fuente_Operador": (
                "ZDA round.txt · tunnel_id / ID Auxiliar"
                if operador_zda else None
            ),
        }

        boom_name = next((n for n in names if re.search(r"-boom\.dat$", n, re.I)), None)
        if not boom_name:
            raise ValueError("El ZDA no contiene boom.dat; no se puede reconstruir la tabla de barrenos.")
        detalle, boom_diag = _parse_zda_boom(zf.read(boom_name), nombre, metadata)
        if detalle.empty:
            raise ValueError("No se encontraron barrenos válidos en boom.dat.")

        counters_name = next((n for n in names if re.search(r"-counters\.dat$", n, re.I)), None)
        movimiento, counter_diag = _parse_zda_counters(zf.read(counters_name)) if counters_name else ({}, {"ok":False,"motivo":"No se encontró counters.dat"})
        for key in [
            "Auto_Brazo1_min","Auto_Brazo2_min","Auto_Total_min","Manual_Brazo1_min","Manual_Brazo2_min","Manual_Total_min",
            "Pct_Movimiento_Automatico","Pct_Movimiento_Manual","Pct_Automatico_Brazo1","Pct_Automatico_Brazo2","Pagina_Movimiento_Brazos"
        ]:
            movimiento.setdefault(key, None)

        planned = int(kv["planned_face_holes"]) if kv.get("planned_face_holes") not in (None,"") else None
        drilled = int(kv["drilled_holes"]) if kv.get("drilled_holes") not in (None,"") else None
        front_types = {"Bottom","Easer","Cut","Contour"}
        front_count = int(detalle["Tipo"].astype(str).isin(front_types).sum())
        total_m = float(pd.to_numeric(detalle["Longitud_roca_m"], errors="coerce").fillna(0).sum())
        metadata.update({
            "Metros_Perforados": round(total_m,2), "Barrenos_Planificados": planned,
            "Barrenos_Realizados": front_count,
            "Fuente_Barrenos_Realizados": "ZDA boom.dat - Bottom/Easer/Cut/Contour",
        })

        validacion, validacion_metros, resumen = _zda_validacion(detalle, metadata)
        mwd_df, mwd_diag = _parse_zda_mwd(zf, names, metadata)
        actual_start = mwd_diag["first_ts"] if mwd_diag["first_ts"] is not None else (decl_start if decl_start is not None else nav_ts)
        actual_end = mwd_diag["last_ts"] if mwd_diag["last_ts"] is not None else decl_end
        sec = actual_end - actual_start if actual_start is not None and actual_end is not None and actual_end >= actual_start else None

        boom_count_ok = drilled is None or drilled == len(detalle)
        counters_ok = bool(counter_diag.get("ok"))
        unknown_ok = len(boom_diag.get("unknown_codes", [])) == 0
        lectura_ok = bool(kv.get("rig") and kv.get("round")) and len(detalle) > 0 and boom_count_ok and counters_ok and unknown_ok

        report = {
            **metadata, **movimiento,
            "Pagina_Tipos_Barreno": None, "Paginas_Detalle": "ZDA boom.dat",
            "Total_Tipos_Reporte": len(detalle), "Total_Extraido": len(detalle), "Diferencia": 0,
            "Estado": "OK" if boom_count_ok else "REVISAR", "Estado_Conteo": "OK" if boom_count_ok else "REVISAR",
            "Estado_Metros_Tipos": "OK", "Lectura_Confiable": "OK" if lectura_ok else "REVISAR",
            "Rig_ZDA": kv.get("rig"), "Labor": kv.get("tunnel_id") or None,
            "Operador_ZDA": operador_zda,
            "Operador": operador_zda,
            "Operador_ZDA_Raw": kv.get("tunnel_id") or None,
            "Fuente_Operador": (
                "ZDA round.txt · tunnel_id / ID Auxiliar"
                if operador_zda else None
            ),
            "Tabla_Curvas": kv.get("curve_table") or None, "PEG": float(kv["peg"]) if kv.get("peg") not in (None,"") else None,
            "Navegacion_ZDA": kv.get("navigation") or None, "Inicio_Declarado_ZDA": kv.get("start") or None,
            "Fin_Declarado_ZDA": kv.get("end") or None, "Inicio_Perforacion": _zda_fmt_datetime(actual_start),
            "Fin_Perforacion": _zda_fmt_datetime(actual_end), "Inicio_Perforacion_TS": actual_start,
            "Fin_Perforacion_TS": actual_end, "Tiempo_Perforacion_s": sec, "Tiempo_Perforacion_hms": _zda_duration_hms(sec),
            "Barrenos_Planificados_ZDA": planned, "Barrenos_ZDA": drilled, "Barrenos_Boom_Validos": len(detalle),
            "Archivos_MWD": mwd_diag["archivos_mwd"], "Muestras_MWD": mwd_diag["muestras_mwd"],
            "Barrenos_MWD_Validos": mwd_diag["barrenos_mwd_validos"], "Intentos_MWD_Cortos": mwd_diag["intentos_mwd_cortos"],
            "Archivos_MWD_Vacios": mwd_diag["archivos_mwd_vacios"], "Layouts_MWD_No_Estandar": mwd_diag["layouts_mwd_no_estandar"],
            "Paso_MWD_Mediana_m": mwd_diag["paso_mwd_mediana_m"], "Metros_MWD": mwd_diag["metros_mwd"],
            "Fuente_Tiempo_Perforacion": "MWD: primer y último registro" if mwd_diag["first_ts"] is not None and mwd_diag["last_ts"] is not None else "start/end de round.txt",
            "Counters_Decodificado": counters_ok, "Boom_Decodificado": True,
            "Codigos_Tipo_No_Reconocidos": ", ".join(f"{x['Codigo']} ({x['N']})" for x in boom_diag.get("unknown_codes", [])),
        }

        # En modo masivo, los visuales se difieren hasta que el usuario
        # abre/solicita el detalle de un ciclo. Esto evita crear cientos
        # de figuras e imágenes durante el procesamiento inicial.
        fig = generar_grafico(detalle, metadata) if generar_visuales else None

        plano_zda_png = (
            generar_plano_zda_png(detalle, metadata)
            if generar_visuales
            else None
        )

        extras = detalle[detalle["Extra"] == True].copy() if "Extra" in detalle.columns else pd.DataFrame()
        result = {
            "metadata": metadata, "movimiento": movimiento, "esperados": {t:int((detalle["Tipo"].astype(str)==t).sum()) for t in ORDEN_TIPOS},
            "detalle": detalle, "validacion": validacion, "validacion_metros": validacion_metros,
            "resumen_ciclo": resumen, "resumen_reporte": report, "extras": extras,
            "fig": fig,
            "plano_nav_png": plano_zda_png,
            "plano_nav_origen": "ZDA_RECONSTRUIDO",
            "mwd_barrenos": mwd_df,
            "diagnostico_zda": {**boom_diag, **counter_diag, **mwd_diag, "boom_count_ok": boom_count_ok},
        }
    return _enriquecer_resultado_estandar(result, "ZDA")


def procesar_archivo(
    path: Path,
    nombre_archivo: Optional[str] = None,
    generar_visuales: bool = True,
) -> Dict:
    """
    Despachador PDF/ZDA con una única interfaz de salida.

    ``generar_visuales=False`` está pensado para procesamiento masivo:
    extrae y valida los datos, pero difiere boxplots/planos hasta que
    realmente se soliciten en la interfaz.
    """
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        return procesar_pdf_v33(
            path,
            nombre_archivo=nombre_archivo,
            generar_visuales=generar_visuales,
        )
    if suffix == ".zda":
        return procesar_zda(
            path,
            nombre_archivo=nombre_archivo,
            generar_visuales=generar_visuales,
        )
    raise ValueError("Formato no soportado. Use archivos .PDF o .ZDA.")
