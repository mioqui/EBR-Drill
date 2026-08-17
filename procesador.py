from __future__ import annotations

from pathlib import Path
from io import BytesIO
import math
import re
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


def limpiar_texto(texto) -> str:
    if texto is None:
        return ""
    return " ".join(str(texto).split())


def identificar_jumbo(numero_serie: Optional[str]) -> str:
    if not numero_serie:
        return "JUMBO_NO_IDENTIFICADO"
    return JUMBOS.get(numero_serie, "JUMBO_NO_IDENTIFICADO")


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
    }

    textos_paginas = []

    with pdfplumber.open(pdf_path) as pdf:
        for pagina in pdf.pages[:4]:
            textos_paginas.append(
                pagina.extract_text() or ""
            )

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
    if df.empty:
        raise ValueError("No hay datos para generar el gráfico.")

    tipos_grafico = [tipo for tipo in ORDEN_TIPOS if not df[df["Tipo"] == tipo].empty]
    fig, ax = plt.subplots(figsize=(14, 8))
    datos_boxplot = [df.loc[df["Tipo"] == tipo, "Longitud_roca_m"].values for tipo in tipos_grafico]

    ax.boxplot(
        datos_boxplot,
        tick_labels=[f"{tipo}\n(n={len(df[df['Tipo'] == tipo])})" for tipo in tipos_grafico],
        widths=0.50,
        showmeans=False,
        showfliers=False,
        whis=(0, 100),
        medianprops={"color": "tab:orange", "linewidth": 1.5},
        boxprops={"color": "black"},
        whiskerprops={"color": "black", "linewidth": 1.2},
        capprops={"color": "black", "linewidth": 1.2},
    )

    for posicion, tipo in enumerate(tipos_grafico, start=1):
        grupo = df[df["Tipo"] == tipo].copy()
        for _, subgrupo in grupo.groupby("Longitud_roca_m", sort=True):
            n = len(subgrupo)
            offsets = [0.0] if n == 1 else np.linspace(-0.06, 0.06, n)
            for offset, (_, fila) in zip(offsets, subgrupo.iterrows()):
                x = posicion + offset
                y = fila["Longitud_roca_m"]
                if fila["Extra"]:
                    ax.scatter(x, y, s=95, facecolor="yellow", edgecolor="black", linewidth=1.5, zorder=5)
                    ax.annotate(f"{fila['ID']} extra", xy=(x, y), xytext=(8, 0), textcoords="offset points", va="center", fontsize=9)
                else:
                    ax.scatter(x, y, s=60, facecolor="black", edgecolor="gray", linewidth=0.8, zorder=4)

    for posicion, tipo in enumerate(tipos_grafico, start=1):
        valores = df.loc[df["Tipo"] == tipo, "Longitud_roca_m"]
        ax.text(
            posicion,
            valores.max() + 0.08,
            f"Min {valores.min():.2f} | Máx {valores.max():.2f}\nProm {valores.mean():.2f} | Med {valores.median():.2f}",
            ha="center",
            va="bottom",
            fontsize=9,
        )

    leyenda = [
        Line2D([0], [0], marker="o", linestyle="None", markersize=7, markerfacecolor="black", markeredgecolor="black", label="Punto: valor individual de cada barreno"),
        Patch(facecolor="white", edgecolor="black", label="Caja: 50% central de los datos (Q1-Q3)"),
        Line2D([0], [0], linestyle="-", linewidth=2, color="tab:orange", label="Línea dentro de la caja: mediana"),
        Line2D([0], [0], linestyle="-", linewidth=1.2, color="black", label="Bigotes: mínimo y máximo real"),
        Line2D([0], [0], marker="o", markerfacecolor="yellow", markeredgecolor="black", linestyle="None", markersize=9, label="Punto resaltado: barreno extra (no programado)"),
    ]
    ax.legend(handles=leyenda, loc="lower left", fontsize=9, title="Leyenda del gráfico")

    min_global = df["Longitud_roca_m"].min()
    max_global = df["Longitud_roca_m"].max()
    rango = max_global - min_global
    ax.set_ylim(min_global - max(0.15, rango * 0.08), max_global + max(0.25, rango * 0.12))

    ax.set_title(
        "Distribución de longitud perforada en roca por tipo de barreno\n"
        f"{metadatos.get('Jumbo') or 'JUMBO'} | Serie {metadatos.get('Numero_Serie') or '-'} | Ciclo {metadatos.get('Ciclo') or '-'} | {metadatos.get('Fecha_Inicio') or '-'}",
        fontsize=14,
    )
    ax.set_xlabel("Tipo de barreno")
    ax.set_ylabel("Longitud perforada en roca (m)")
    ax.grid(True, alpha=0.5)
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

def procesar_pdf(pdf_path: Path, nombre_archivo: Optional[str] = None) -> Dict:
    metadatos = leer_metadatos(pdf_path, nombre_archivo=nombre_archivo)
    movimiento = leer_movimiento_brazos(pdf_path)
    esperados, pagina_tipos = leer_totales_tipos_barreno(pdf_path)
    df, paginas_detalle = extraer_barrenos(pdf_path, metadatos)

    if df.empty:
        raise ValueError("No se encontraron barrenos ejecutados en la tabla de detalle.")

    # Fallback para clasificación del disparo:
    # si iSURE no expone claramente "Barrenos realizados en frentes",
    # contar los barrenos ejecutados excluyendo Reaming.
    if metadatos.get("Barrenos_Realizados") is None:
        metadatos["Barrenos_Realizados"] = int(
            (
                df["Tipo"].astype(str)
                != "Reaming"
            ).sum()
        )

        metadatos[
            "Fuente_Barrenos_Realizados"
        ] = "Detalle extraído - excluye Reaming"

    validacion = construir_validacion(df, esperados, metadatos)
    resumen_ciclo = construir_resumen_ciclo(df, esperados, metadatos)
    resumen_reporte = construir_resumen_reporte(metadatos, esperados, df, pagina_tipos, paginas_detalle, movimiento)
    extras = df[df["Extra"]].copy()
    fig = generar_grafico(df, metadatos)
    plano_nav_png = extraer_plano_navegacion_png(pdf_path)

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
    }
