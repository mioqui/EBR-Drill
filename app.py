from __future__ import annotations

from io import BytesIO
from pathlib import Path
from datetime import datetime, timezone, timedelta
import hashlib
import tempfile
import re
import gc
import shutil
import uuid
import zipfile
import struct

import pandas as pd
import numpy as np
import streamlit as st
import matplotlib.pyplot as plt
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from plotly.colors import qualitative
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

from reporte_perforacion import generar_reporte_pdf
import asignaciones_operador as asig

from procesador import (
    procesar_archivo,
    clasificar_tipo_disparo_v33,
    generar_grafico,
    generar_plano_zda_png,
    process_zda_bytes as procesar_eficiencia_zda,
    volumen_unico_ciclo,
    resumenes_eficiencia_lote,
    VERSION as PROCESADOR_EFICIENCIA_VERSION,
    CATALOGO_OPERADORES_ZDA,
    parse_round_dat_profile,
)


# ==========================================================
# CONFIGURACIÓN
# ==========================================================

APP_VERSION_INTERNAL = "V35.22-ORDEN-BALANCE"

# Turnos del filtro lateral: Día = 07:00 a 19:00, Noche = 19:00 a 07:00 (hora de inicio del ciclo).
TURNOS_FILTRO = ["Día", "Noche"]
PUBLIC_VERSION = "v1.0"
CACHE_SCHEMA_VERSION = "v35_22_orden_balance_20260917"
TIPOS_DISPARO = ["FRENTE", "SELLADA", "ESTOCADA Y/O CORRECCIONES"]
COLORES = qualitative.Plotly

st.set_page_config(page_title=f"EBR Drill Analytics · Piloto {PUBLIC_VERSION}", page_icon="⛏️", layout="wide")
st.title("EBR Drill Analytics")
st.caption(f"Piloto {PUBLIC_VERSION} · Análisis de archivos ZDA de equipos Jumbo")
st.markdown(
    """
    <style>
    /* Reserva espacio debajo de la barra superior de Streamlit
       para evitar que el título quede recortado/oculto por el header. */
    [data-testid="stAppViewBlockContainer"],
    .block-container {
        padding-top: 3.25rem !important;
        padding-bottom: 6rem !important;
    }

    h1 {
        font-size: 1.75rem !important;
        line-height: 1.20 !important;
        margin-top: 0 !important;
        margin-bottom: 0.15rem !important;
        padding-top: 0.15rem !important;
        overflow: visible !important;
    }
    /* Los <style> de primer nivel son elementos vacíos que igual suman el hueco entre elementos
       (16 px cada uno): se ocultan. Solo afecta al nivel superior de la página. */
    [data-testid="stAppViewBlockContainer"] > [data-testid="stVerticalBlock"] > [data-testid="stElementContainer"]:has(> [data-testid="stMarkdown"] style) {
        display: none;
    }
    /* Tarjeta "Cargar ciclos": menos aire arriba y separación uniforme entre sus elementos. */
    .st-key-zona_carga { gap: 0.6rem !important; padding-bottom: 0.6rem !important; }
    .st-key-zona_carga h3 { padding: 0 0 0.55rem 0 !important; margin: 0 !important; }
    h2 { font-size: 1.22rem !important; margin-top: 0.7rem !important; }
    h3 { font-size: 1.00rem !important; margin-top: 0.55rem !important; margin-bottom: 0.25rem !important; }
    [data-testid="stMetricLabel"] { font-size: 0.75rem !important; }
    [data-testid="stMetricValue"] { font-size: 1.20rem !important; }
    [data-testid="stDataFrame"] div[role="grid"] { font-size: 0.76rem !important; }
    [data-testid="stExpander"] summary { font-size: 0.88rem !important; font-weight: 600 !important; }
    .stDownloadButton button, .stButton button { min-height: 2.15rem !important; font-size: 0.80rem !important; }

    /* Panel de filtros inspirado en la versión web: solo estética, sin logos. */
    [data-testid="stSidebar"] {
        background: linear-gradient(155deg, #194D48 0%, #153F3C 66%, #103936 100%) !important;
        border-right: 1px solid rgba(255,255,255,.10);
    }
    [data-testid="stSidebar"] > div:first-child { background: transparent !important; }
    [data-testid="stSidebar"] [data-testid="stSidebarUserContent"] {
        padding-top: 1.45rem !important;
    }
    [data-testid="stSidebar"] h2 {
        color: #fff !important;
        font-size: 1.25rem !important;
        letter-spacing: -.02em;
        padding-bottom: .9rem !important;
        border-bottom: 1px solid rgba(255,255,255,.18);
        margin-bottom: 1.1rem !important;
    }
    [data-testid="stSidebar"] h4,
    [data-testid="stSidebar"] h5 {
        color: #D8E7E4 !important;
        font-size: .78rem !important;
        font-weight: 800 !important;
        letter-spacing: .105em !important;
        text-transform: uppercase;
        margin-top: .85rem !important;
        margin-bottom: .55rem !important;
    }
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
    [data-testid="stSidebar"] [data-testid="stCaptionContainer"],
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] label p {
        color: #F4F8F7 !important;
    }
    [data-testid="stSidebar"] [data-testid="stCaptionContainer"] {
        color: #BDD0CD !important;
        line-height: 1.5;
    }
    /* Signo de interrogación de ayuda (help=): claro para que se vea sobre el verde del panel. */
    /* El trazo del ícono lo fija Streamlit con un gris oscuro (no usa currentColor): se fuerza aquí. */
    [data-testid="stSidebar"] [data-testid="stTooltipIcon"] button {
        color: #BDD0CD !important;
        opacity: 1 !important;
    }
    [data-testid="stSidebar"] [data-testid="stTooltipIcon"] svg,
    [data-testid="stSidebar"] [data-testid="stTooltipIcon"] svg * {
        stroke: #BDD0CD !important;
    }
    [data-testid="stSidebar"] [data-testid="stTooltipIcon"] button:hover svg,
    [data-testid="stSidebar"] [data-testid="stTooltipIcon"] button:hover svg * {
        stroke: #FFFFFF !important;
    }
    [data-testid="stSidebar"] [data-testid="stCheckbox"] {
        margin: 0 !important;
        padding: .02rem 0 !important;
    }
    [data-testid="stSidebar"] [data-testid="stCheckbox"] label p {
        font-size: .88rem !important;
        font-weight: 500 !important;
    }
    [data-testid="stSidebar"] [data-testid="stCheckbox"] input[type="checkbox"] { accent-color: #95C123 !important; }
    [data-testid="stSidebar"] [data-testid="stCheckbox"] label[data-checked="true"] > div:first-child,
    [data-testid="stSidebar"] [data-testid="stCheckbox"] label > span:first-child {
        accent-color: #95C123 !important;
    }
    [data-testid="stSidebar"] [data-baseweb="checkbox"] [aria-checked="true"] {
        background-color: #95C123 !important;
        border-color: #95C123 !important;
    }
    [data-testid="stSidebar"] [data-testid="stDateInput"] input {
        background: #fff !important;
        color: #203342 !important;
        border-radius: .7rem !important;
        font-size: .81rem !important;
    }
    [data-testid="stSidebar"] [data-testid="stDateInput"] > label p {
        font-size: .74rem !important;
        color: #DCEAE7 !important;
    }
    [data-testid="stSidebar"] [data-testid="stDateInput"] [data-baseweb="input"] {
        border-radius: .7rem !important;
        background-color: #fff !important;
    }
    [data-testid="stSidebar"] hr {
        border-color: rgba(255,255,255,.16) !important;
        margin: 1.15rem 0 !important;
    }
    [data-testid="stSidebar"] button[kind="secondary"] {
        border-color: rgba(255,255,255,.35) !important;
        color: #fff !important;
    }
    @media (min-width: 900px) {
        [data-testid="stSidebar"] { min-width: 340px !important; max-width: 370px !important; }
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ==========================================================
# ESTADO
# ==========================================================

if st.session_state.get("cache_schema_version") != CACHE_SCHEMA_VERSION:
    st.session_state.procesados = {}
    st.session_state.uploader_version = st.session_state.get("uploader_version", 0) + 1
    st.session_state.cache_schema_version = CACHE_SCHEMA_VERSION

if "uploader_version" not in st.session_state:
    st.session_state.uploader_version = 0
if "procesados" not in st.session_state:
    st.session_state.procesados = {}
if "session_disk_id" not in st.session_state:
    st.session_state.session_disk_id = uuid.uuid4().hex
if "staged_queue" not in st.session_state:
    st.session_state.staged_queue = []
if "auto_process_staged" not in st.session_state:
    st.session_state.auto_process_staged = False


def _session_work_dir() -> Path:
    root = Path(tempfile.gettempdir()) / "ebr_drill_massive"
    path = root / st.session_state.session_disk_id
    path.mkdir(parents=True, exist_ok=True)
    (path / "uploads").mkdir(exist_ok=True)
    (path / "visuals").mkdir(exist_ok=True)
    return path


# Streamlit vuelve a ejecutar el script completo ante cualquier cambio de widget.
# st.fragment permite que filtros y checkboxes vuelvan a ejecutar SOLO su bloque,
# evitando recalcular Excel, otros gráficos y resultados individuales.
if hasattr(st, "fragment"):
    fragment = st.fragment
else:
    def fragment(func):
        return func


def limpiar_analisis():
    work_dir = None
    try:
        work_dir = _session_work_dir()
    except Exception:
        pass

    st.session_state.procesados = {}
    st.session_state.staged_queue = []
    st.session_state.auto_process_staged = False
    st.session_state.uploader_version += 1

    if work_dir and work_dir.exists():
        shutil.rmtree(work_dir, ignore_errors=True)

    # Nuevo directorio limpio para la misma sesión.
    st.session_state.session_disk_id = uuid.uuid4().hex

    for key in list(st.session_state.keys()):
        if key.startswith((
            "global_",
            "_global_",
            "opt_",
            "auto_",
            "cut_",
            "zda_",
            "class_",
            "lbl_",
            "desglosar_",
            "detalle_ciclo_",
            "sidebar_check_",
        )):
            del st.session_state[key]


def tipo_roca_desde_plan_texto(plan_perforacion):
    """
    Extrae el tipo/clase de roca desde Plan_Perforacion / drill_plan.
    Ejemplos:
      "MALLA E.E. 4.5x4.5 III-B" -> "III-B"
      "MALLA E.B. 5.0x5.0 IV-A"   -> "IV-A"

    Se acepta cualquier número romano seguido de guion y código
    alfanumérico para no limitar futuros valores.
    """
    texto = str(plan_perforacion or "").upper()
    m = re.search(
        r"\b([IVXLCDM]+)\s*[-–—]\s*([A-Z0-9]+)\b",
        texto,
        re.IGNORECASE,
    )
    if not m:
        return "SIN DATO"

    return f"{m.group(1).upper()}-{m.group(2).upper()}"


# ==========================================================
# FILTROS Y PARÁMETROS GLOBALES
# ==========================================================

def _asignaciones_operador() -> dict:
    """Asignaciones manuales de operador. Se leen del disco una vez por sesión."""
    if "_asig_operadores" not in st.session_state:
        st.session_state["_asig_operadores"] = asig.cargar()
    return st.session_state["_asig_operadores"]


def _guardar_asignaciones_operador() -> None:
    """Persiste las asignaciones y recuerda la ruta realmente usada."""
    st.session_state["_asig_ruta"] = str(asig.guardar(_asignaciones_operador()))


def _valores_detectados_desde_cache():
    """
    Obtiene jumbos, tipos de disparo y tipos de roca desde los
    resultados ya procesados. No presupone cuántos equipos ni
    cuántas clases de roca existen.
    """
    jumbos = []
    tipos = []
    rocas = []
    operadores = []

    for r in st.session_state.procesados.values():
        if r.get("error"):
            continue

        rep = r.get("resumen_reporte") or {}
        jumbo = rep.get("Jumbo")
        tipo = rep.get("Tipo_Disparo")
        roca = tipo_roca_desde_plan_texto(
            rep.get("Plan_Perforacion")
        )
        operador = (
            rep.get("Operador_ZDA")
            or rep.get("Operador")
            or None
        )
        # Conservar explícitamente los ciclos sin operador como opción de filtro.
        if operador is None or str(operador).strip().upper() in (
            "", "SIN DATO", "NONE", "NAN", "NULL"
        ):
            operador = "SIN DATO"

        if jumbo and str(jumbo).strip():
            jumbos.append(str(jumbo).strip())
        if tipo and str(tipo).strip():
            tipos.append(str(tipo).strip())
        if roca and str(roca).strip():
            rocas.append(str(roca).strip())
        if operador and str(operador).strip():
            operadores.append(str(operador).strip())

    # Orden natural para JUMB001, JUMB002... y luego series/otros.
    def jumbo_sort_key(valor):
        m = re.fullmatch(r"JUMB(\d+)", str(valor), re.IGNORECASE)
        if m:
            return (0, int(m.group(1)))
        return (1, str(valor))

    jumbos = sorted(set(jumbos), key=jumbo_sort_key)
    tipos_presentes = set(tipos)
    tipos = [t for t in TIPOS_DISPARO if t in tipos_presentes]

    # Si apareciera una clasificación futura no contemplada, no se pierde.
    tipos_extra = sorted(tipos_presentes - set(TIPOS_DISPARO))
    tipos.extend(tipos_extra)

    # SIN DATO al final; los demás valores alfabéticamente.
    rocas_presentes = set(rocas)
    rocas = sorted(r for r in rocas_presentes if r != "SIN DATO")
    if "SIN DATO" in rocas_presentes:
        rocas.append("SIN DATO")

    operadores_presentes = set(operadores)
    operadores = sorted(o for o in operadores_presentes if o != "SIN DATO")
    if "SIN DATO" in operadores_presentes:
        operadores.append("SIN DATO")

    return jumbos, tipos, rocas, operadores


def _sincronizar_multiselect_dinamico(key, options, previous_options_key):
    """
    Mantiene la selección del usuario y agrega automáticamente
    cualquier opción nueva detectada tras procesar más archivos.
    """
    options = list(options)
    prev_options = list(st.session_state.get(previous_options_key, []))

    if key not in st.session_state:
        st.session_state[key] = options.copy()
    else:
        actual = [
            x for x in st.session_state.get(key, [])
            if x in options
        ]
        nuevos = [
            x for x in options
            if x not in prev_options
        ]
        for x in nuevos:
            if x not in actual:
                actual.append(x)
        st.session_state[key] = actual

    st.session_state[previous_options_key] = options.copy()


# Las asignaciones manuales se aplican ANTES de detectar filtros: así el panel lateral, los
# gráficos, el Excel y "Resultados por archivo" ven el operador ya asignado.
asig.aplicar_a_resultados(st.session_state.procesados.values(), _asignaciones_operador())

jumbos_detectados, tipos_detectados, rocas_detectadas, operadores_detectados = _valores_detectados_desde_cache()

# Variables siempre definidas aunque aún no existan datos.
global_turnos = list(TURNOS_FILTRO)
global_jumbos = []
global_tipos = []
global_rocas = []
global_operadores = []
sidebar_fecha_container = None

with st.sidebar:
    st.markdown(
        """
        <div style="position:sticky; top:0; z-index:999;
                    background:transparent;
                    padding:0.25rem 0 0.85rem 0;
                    margin-bottom:0.45rem;
                    border-bottom:1px solid rgba(255,255,255,0.22);">
            <div style="font-size:1.18rem; font-weight:700; line-height:1.2;
                        color:#ffffff; letter-spacing:-0.01em;">
                EBR Drill Analytics
            </div>
            <div style="font-size:0.78rem; color:rgba(255,255,255,0.72);
                        margin-top:0.20rem; font-weight:400;">
                ZDA Analytics
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.header("Filtros y parámetros")
    sidebar_fecha_container = st.container()

    if not jumbos_detectados:
        st.info(
            "Los filtros se habilitarán automáticamente después de procesar "
            "los primeros archivos ZDA."
        )
    else:
        _sincronizar_multiselect_dinamico(
            "global_jumbos",
            jumbos_detectados,
            "_global_jumbos_options_prev",
        )
        _sincronizar_multiselect_dinamico(
            "global_tipos",
            tipos_detectados,
            "_global_tipos_options_prev",
        )
        _sincronizar_multiselect_dinamico(
            "global_rocas",
            rocas_detectadas,
            "_global_rocas_options_prev",
        )
        _sincronizar_multiselect_dinamico(
            "global_operadores",
            operadores_detectados,
            "_global_operadores_options_prev",
        )

        # Igual que la interfaz HTML: opciones visibles sin desplegables.
        # El estado global_* sigue siendo una lista y mantiene los filtros originales.
        def _grupo_checks(titulo, clave, opciones, ayuda=None):
            st.divider()
            # `ayuda` muestra un signo de interrogación con el mensaje al pasar el mouse.
            st.markdown(f"#### {titulo}", help=ayuda)
            seleccion_previa = set(st.session_state.get(clave, opciones))
            elegidos = []
            for opcion in opciones:
                check_key = ("sidebar_check_" + clave + "_" +
                             hashlib.md5(str(opcion).encode("utf-8")).hexdigest()[:12])
                if check_key not in st.session_state:
                    st.session_state[check_key] = opcion in seleccion_previa
                # Las opciones descubiertas en una carga nueva quedan activadas
                # sin sobrescribir selecciones que el usuario haya desmarcado.
                if st.checkbox(str(opcion), key=check_key):
                    elegidos.append(opcion)
            st.session_state[clave] = elegidos
            return elegidos

        global_turnos = _grupo_checks(
            "Turnos", "global_turnos", TURNOS_FILTRO,
            ayuda=("Aplica a Uso Automático, Longitud de Perforación, Primer Golpe, Eficiencia de "
                   "Perforación, Clasificación y ROP por barreno. Día: 07:00 a 19:00 · Noche: 19:00 a "
                   "07:00, según la hora de inicio del ciclo (igual que el TURNO del Excel)."),
        )
        global_jumbos = _grupo_checks("Jumbos", "global_jumbos", jumbos_detectados)
        global_tipos = _grupo_checks(
            "Tipo de disparo", "global_tipos", tipos_detectados,
            ayuda=("Clasificación por barrenos realizados: Frente ≥45 | Sellada 25–44 | "
                   "Estocada y/o correcciones <25."),
        )
        global_rocas = _grupo_checks("Tipo de roca", "global_rocas", rocas_detectadas)
        global_operadores = _grupo_checks("Operadores", "global_operadores", operadores_detectados)

        st.divider()
        # Las opciones de etiquetas, tipo de gráfico y línea curva ahora están sobre
        # cada gráfico (ver `_controles_grafico`).

        st.caption(
            "Los filtros no eliminan datos del Excel exportado; la exportación conserva "
            "todos los archivos procesados."
        )


# ==========================================================
# HELPERS
# ==========================================================

def seccion_desde_plan_texto(plan_perforacion):
    """
    Devuelve una etiqueta corta de sección desde Plan_Perforacion.
    Ej.: "MALLA E.E. 4.5x4.5 III-B" -> "4.5 x 4.5"
    """
    texto = str(plan_perforacion or "")
    m = re.search(
        r"(\d+(?:[.,]\d+)?)\s*[xX×]\s*(\d+(?:[.,]\d+)?)",
        texto,
    )
    if not m:
        return "-"

    a = float(m.group(1).replace(",", "."))
    b = float(m.group(2).replace(",", "."))
    return f"{a:.1f} x {b:.1f}"


def hash_archivo_en_disco(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            h.update(chunk)
    h.update(CACHE_SCHEMA_VERSION.encode("utf-8"))
    return h.hexdigest()


def preparar_archivos_en_disco(uploaded_files):
    """
    Primera fase del modo masivo.

    Copia los UploadedFile a disco temporal SIN procesarlos todavía.
    Después se fuerza un rerun para que Streamlit libere del uploader los
    bytes de los 100–150 ZDA antes de comenzar el parsing intensivo.
    """
    uploads_dir = _session_work_dir() / "uploads"
    nuevos = []

    for n, archivo in enumerate(uploaded_files, start=1):
        safe_name = re.sub(r"[^A-Za-z0-9._-]+", "_", Path(archivo.name).name)
        disk_path = uploads_dir / f"{uuid.uuid4().hex[:10]}_{safe_name}"

        archivo.seek(0)
        with disk_path.open("wb") as out:
            shutil.copyfileobj(archivo, out, length=1024 * 1024)

        clave = hash_archivo_en_disco(disk_path)
        if clave in st.session_state.procesados:
            disk_path.unlink(missing_ok=True)
            continue

        # Evitar duplicar el mismo archivo en una cola ya preparada.
        if any(x.get("clave") == clave for x in st.session_state.staged_queue):
            disk_path.unlink(missing_ok=True)
            continue

        item = {
            "clave": clave,
            "nombre": archivo.name,
            "path": str(disk_path),
            "size": disk_path.stat().st_size,
        }
        st.session_state.staged_queue.append(item)
        nuevos.append(item)

    return nuevos


def preparar_carpeta_local(ruta_carpeta: str):
    """Carga todos los .ZDA de una carpeta LOCAL, sin depender del diálogo
    del navegador. Se copia al temporal de la sesión para permitir ROP diferido;
    nunca se elimina el archivo original del usuario.
    """
    carpeta = Path(ruta_carpeta.strip().strip('"').strip("'")).expanduser()
    if not carpeta.is_dir():
        raise ValueError("La ruta indicada no existe o no es una carpeta accesible en el equipo que ejecuta Python.")
    archivos = sorted(p for p in carpeta.rglob("*")
                      if p.is_file() and not p.is_symlink() and p.suffix.lower() == ".zda")
    if not archivos:
        raise ValueError("La carpeta no contiene archivos .ZDA (tampoco en sus subcarpetas).")
    if len(archivos) > 2000:
        raise ValueError("La carpeta contiene más de 2000 ZDA. Selecciona una subcarpeta para esta carga.")
    uploads_dir = _session_work_dir() / "uploads"
    nuevos = []
    vistos = set(st.session_state.procesados)
    vistos.update(x.get("clave") for x in st.session_state.staged_queue)
    for archivo in archivos:
        clave = hash_archivo_en_disco(archivo)
        if clave in vistos:
            continue
        destino = uploads_dir / f"{uuid.uuid4().hex[:10]}_{re.sub(r'[^A-Za-z0-9._-]+', '_', archivo.name)}"
        shutil.copyfile(archivo, destino)
        item = {"clave": clave, "nombre": archivo.name,
                "path": str(destino), "size": destino.stat().st_size}
        st.session_state.staged_queue.append(item)
        nuevos.append(item)
        vistos.add(clave)
    return len(archivos), len(nuevos)


def fmt(valor, dec=1, sufijo=""):
    if valor is None or pd.isna(valor):
        return "-"
    return f"{float(valor):.{dec}f}{sufijo}"


def concatenar_dataframes(resultados, key, solo_ok=False):
    dfs = []
    for r in resultados:
        if solo_ok and str(r.get("resumen_reporte", {}).get("Estado")) != "OK":
            continue
        df = r.get(key)
        if isinstance(df, pd.DataFrame) and not df.empty:
            dfs.append(df)
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()


def asegurar_fechahora(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["FechaHora"] = pd.to_datetime(
        out.get("Fecha_Inicio", pd.Series(index=out.index, dtype=object)).fillna("")
        + " "
        + out.get("Hora_Inicio", pd.Series(index=out.index, dtype=object)).fillna(""),
        format="%d/%m/%Y %H:%M:%S",
        errors="coerce",
    )
    faltan = out["FechaHora"].isna()
    if faltan.any() and "Fecha_Inicio" in out.columns:
        out.loc[faltan, "FechaHora"] = pd.to_datetime(
            out.loc[faltan, "Fecha_Inicio"], format="%d/%m/%Y", errors="coerce"
        )
    return out[out["FechaHora"].notna()].sort_values(["FechaHora", "Jumbo", "Ciclo"])


def smart_annotations(points, x_window_hours=18, y_window=3.0, font_size=11):
    """Equivalente del buildSmartAnnotations de la HTML V33."""
    out, used = [], []
    points = sorted(points, key=lambda p: (pd.Timestamp(p["x"]), float(p["y"])))
    x_window = pd.Timedelta(hours=x_window_hours)
    for i, p in enumerate(points):
        base_sign = -1 if int(p.get("rank", i)) % 2 == 0 else 1
        for level in range(8):
            yshift = base_sign * (18 + level * 14)
            clash = False
            for u in used:
                if (
                    abs(pd.Timestamp(u["x"]) - pd.Timestamp(p["x"])) <= x_window
                    and abs(float(u["y"]) - float(p["y"])) <= y_window
                    and abs(u["yshift"] - yshift) < 14
                ):
                    clash = True
                    break
            if not clash:
                used.append({"x": p["x"], "y": p["y"], "yshift": yshift})
                out.append(
                    dict(
                        x=p["x"], y=p["y"], xref="x", yref="y",
                        showarrow=False, yshift=yshift, text=p["text"],
                        align="center", font=dict(size=font_size, color="#334155"),
                        bgcolor="rgba(255,255,255,0.88)",
                        bordercolor="rgba(203,213,225,0.95)", borderpad=3,
                    )
                )
                break
    return out


def _pct_entero_txt(v):
    """Porcentaje entero sin decimales (redondeo comercial: 44.5 -> 45%)."""
    try:
        return f"{int(np.floor(float(v) + 0.5))}%"
    except (TypeError, ValueError):
        return "-"


def eje_fechas_dias(fechas, max_etiquetas=31):
    """Eje X de fechas que muestra TODOS los días del rango cuando caben.

    - Hasta `max_etiquetas` días: una marca por día (incluye días sin ciclos).
    - Rangos mayores: paso automático de 2, 3, 7, 14 o 30 días para no saturar.
    Con paso de 1 día la etiqueta se centra en el día ("period") y las líneas de la
    cuadrícula separan un día del siguiente.
    """
    base = dict(title="Fecha", tickformat="%d/%m", gridcolor="#eef2f7", showgrid=True)
    f = pd.to_datetime(pd.Series(fechas), errors="coerce").dropna()
    if f.empty:
        return base
    d0, d1 = f.min().normalize(), f.max().normalize()
    n_dias = (d1 - d0).days + 1
    paso = next((p for p in (1, 2, 3, 7, 14, 30) if n_dias / p <= max_etiquetas), 30)
    base.update(
        dtick=paso * 86_400_000,                         # milisegundos
        tick0=d0.strftime("%Y-%m-%d"),
        range=[d0.strftime("%Y-%m-%d"), (d1 + pd.Timedelta(days=1)).strftime("%Y-%m-%d")],
    )
    if paso == 1:
        base["ticklabelmode"] = "period"
    if n_dias / paso > 16:
        base["tickangle"] = -45
    return base


def _eje_barras_por_dia(fechas, max_etiquetas=31, hueco=0.6):
    """Posiciones X uniformes para barras "una por ciclo", agrupadas por día.

    Sobre un eje de tiempo continuo los ciclos caen a horas irregulares y las barras salen
    de anchos distintos, pegadas o muy finas. Aquí cada ciclo ocupa una posición de igual
    ancho, en orden cronológico; los ciclos de un mismo día quedan juntos, los días sin
    ciclos dejan un espacio vacío y la etiqueta (dd/mm) va centrada bajo cada día.

    Devuelve (posiciones alineadas con `fechas`, dict con la configuración del eje X y los
    separadores entre días). Requiere un índice posicional (0..n-1).
    """
    f = pd.to_datetime(pd.Series(list(fechas)), errors="coerce")
    dia = f.dt.normalize()
    ok = f.notna().to_numpy()
    d0, d1 = dia[ok].min(), dia[ok].max()
    dias = pd.date_range(d0, d1, freq="D")
    n_dias = len(dias)
    paso = next((p for p in (1, 2, 3, 7, 14, 30) if n_dias / p <= max_etiquetas), 30)

    por_dia = {d: [] for d in dias}
    for i in np.argsort(f.to_numpy(), kind="stable"):
        if ok[i]:
            por_dia[dia.iloc[i]].append(i)

    pos = np.full(len(f), np.nan)
    tickvals, ticktext, separadores = [], [], []
    cursor = 0.0
    for k, d in enumerate(dias):
        idxs = por_dia[d]
        n = max(len(idxs), 1)
        for j, i in enumerate(idxs):
            pos[i] = cursor + j
        if k % paso == 0:
            tickvals.append(cursor + (n - 1) / 2)
            ticktext.append(d.strftime("%d/%m"))
        fin = cursor + n - 1
        if k < n_dias - 1 and n_dias <= max_etiquetas:
            separadores.append(fin + 0.5 + hueco / 2)
        cursor = fin + 1 + hueco
    ultimo = cursor - hueco - 1
    eje = dict(
        title="Fecha", type="linear", tickmode="array", tickvals=tickvals, ticktext=ticktext,
        range=[-0.7, ultimo + 0.7], showgrid=False, zeroline=False,
    )
    if len(tickvals) > 16:
        eje["tickangle"] = -45
    shapes = [
        dict(type="line", xref="x", yref="paper", x0=s, x1=s, y0=0, y1=1,
             line=dict(color="#e2e8f0", width=1, dash="dot"), layer="below")
        for s in separadores
    ]
    return pos, dict(xaxis=eje, shapes=shapes)


def _font_etiqueta_barras(n_barras):
    """Tamaño de la etiqueta de las barras según la cantidad de barras del gráfico."""
    if n_barras <= 50:
        return 11
    if n_barras <= 70:
        return 10
    if n_barras <= 95:
        return 9
    return 8


# Suavizado de las curvas de los gráficos de evolución (Plotly: 0 = tramos rectos, 1 = muy curvo).
# Bajarlo reduce las ondulaciones y que la curva se pase por debajo de 0 entre dos puntos.
SUAVIZADO_LINEA = 0.5


def _estilo_linea(color, curva=True, simbolo="circle", ancho=3.0):
    """Estilo de línea de los gráficos de evolución: curva suave con marcadores huecos
    (círculo blanco con borde del color de la serie). Sin relleno bajo la línea.
    Sin `curva` solo se dibujan los marcadores."""
    marcador = dict(size=9, symbol=simbolo, color="#ffffff", line=dict(color=color, width=2.5))
    if not curva:
        return dict(mode="markers", marker=marcador)
    return dict(
        mode="lines+markers",
        line=dict(width=ancho, color=color, shape="spline", smoothing=SUAVIZADO_LINEA),
        marker=marcador,
    )


def _es_barras(tipo_grafico):
    return str(tipo_grafico).strip().lower().startswith("barra")


def base_layout(height=450, **kwargs):
    layout = dict(
        height=height,
        paper_bgcolor="#ffffff",
        plot_bgcolor="#ffffff",
        font=dict(family="Arial, sans-serif", size=12, color="#334155"),
        margin=dict(l=70, r=30, t=70, b=70),
        hovermode="closest",
        legend=dict(orientation="h", y=-0.18, x=0),
        xaxis=dict(showgrid=True, gridcolor="#eef2f7", zeroline=False),
        yaxis=dict(showgrid=True, gridcolor="#eef2f7", zeroline=False),
    )
    layout.update(kwargs)
    return layout


# ==========================================================
# EXCEL / BD-PERFO - PUBLICACIÓN V1.0
# ==========================================================

BD_PERFO_COLUMNS = [
    "MES","SEMANA","AÑO","FECHA","TURNO","JEFE DE TURNO","OPERADOR","JUMBO",
    "NIVEL","BLOCK","LABOR","TIPO DISPARO","SECCIÓN","RMR","TIP. DE ROCA",
    "TIPO EXP.","Hora 1° taladro","Tiempo de perforación efec.","Tiempo de perforación",
    "Tiempo de movimiento ","T. mov. Manual","T.mov. Autom.","T. AUTOMÁTICO",
    "# TALADROS","T. MANUAL","Av. Perf","Av.Top","Av.Real","Eficiencia",
    "Obj. Automático","CONLABOR","Tipo Mat","Ciclo","%AutoDrill","% Auto B1",
    "% Auto B2","Mediana Cut (m)","Promedio Cut (m)","Hora 1° martillo"
]

BD_JUMBO_ALIAS = {"JUMB001": "JF01", "JUMB002": "JF02"}
BD_MESES_OPERATIVOS = [
    "Enero","Febrero","Marzo","Abril","Mayo","Junio",
    "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre",
]


def _bd_num(v):
    if v is None or pd.isna(v):
        return None
    try:
        n = float(v)
        return n if pd.notna(n) else None
    except Exception:
        return None


def _bd_parse_datetime(row):
    candidatos = []
    inicio_perf = row.get("Inicio_Perforacion")
    if inicio_perf is not None and not pd.isna(inicio_perf):
        candidatos.append(str(inicio_perf))
    fecha = row.get("Fecha_Inicio")
    hora = row.get("Hora_Inicio")
    if fecha is not None and not pd.isna(fecha):
        candidatos.append(
            f"{fecha} {hora if hora is not None and not pd.isna(hora) else '00:00:00'}"
        )

    for s in candidatos:
        dt = pd.to_datetime(s, format="%d/%m/%Y %H:%M:%S", errors="coerce")
        if pd.notna(dt):
            return dt.to_pydatetime()
    return None


def _bd_mes_operativo(dt):
    """
    Mes operativo para la hoja BD-PERFO.

    Regla:
    - Del día 01 al 25 -> pertenece al mes calendario actual.
    - Del día 26 al último día del mes -> pertenece al mes operativo siguiente.

    Ejemplos:
    - 26/07 al 25/08 = Agosto
    - 26/08 al 25/09 = Septiembre
    - 26/12 al 25/01 = Enero
    """
    if dt is None:
        return None

    mes_operativo = dt.month + (1 if dt.day >= 26 else 0)
    if mes_operativo > 12:
        mes_operativo = 1

    return BD_MESES_OPERATIVOS[mes_operativo - 1]


def _bd_turno(dt):
    if dt is None:
        return None
    return "DIA" if 7 <= dt.hour < 19 else "NOCHE"


def _bd_hora_excel(dt):
    if dt is None:
        return None
    return (dt.hour * 3600 + dt.minute * 60 + dt.second) / 86400.0


def _bd_plan_fields(plan):
    p = "" if plan is None or pd.isna(plan) else str(plan)
    seccion = roca = tipo_exp = None

    m = re.search(r"(\d+(?:[.,]\d+)?)\s*[xX×]\s*(\d+(?:[.,]\d+)?)", p)
    if m:
        seccion = f"{m.group(1).replace(',', '.')}x{m.group(2).replace(',', '.')}"

    m = re.search(r"\b(VI|IV|V|III|II|I)\s*-\s*([A-C])\b", p, re.IGNORECASE)
    if m:
        roca = f"{m.group(1).upper()}-{m.group(2).upper()}"

    if re.search(r"\bE\s*\.\s*E\s*\.?", p, re.IGNORECASE):
        tipo_exp = "Encartuchada"
    elif re.search(r"\bE\s*\.\s*B\s*\.?", p, re.IGNORECASE):
        tipo_exp = "Bombeable"

    return seccion, roca, tipo_exp


def _bd_labor_fields(row):
    curve_raw = row.get("Tabla_Curvas")
    tunnel_raw = row.get("Labor")
    curve = "" if curve_raw is None or pd.isna(curve_raw) else str(curve_raw).strip()
    tunnel = "" if tunnel_raw is None or pd.isna(tunnel_raw) else str(tunnel_raw).strip()
    src = curve or tunnel
    nivel = block = labor = None

    if src:
        m = re.search(r"\bNV\s*:?\s*(\d+)\b", src, re.IGNORECASE)
        if m:
            nivel = int(m.group(1))

        m = re.search(r"\b(?:BLOCK|B)\s*:?\s*(\d+)\b", src, re.IGNORECASE)
        if m:
            block = int(m.group(1))

        if curve:
            for token in [t for t in re.split(r"\s+", curve) if t]:
                if re.match(r"^(?:NV|B|BLOCK)\s*:?\d+$", token, re.IGNORECASE):
                    continue
                if re.match(r"^\d+$", token):
                    continue
                labor = token.upper()
                break

        if not labor and tunnel:
            for key, value in re.findall(
                r"\b([A-Za-z]{1,8})\s*:\s*([A-Za-z0-9._-]+)", tunnel
            ):
                if key.upper() not in {"NV","OP","T","TURN","RMR"}:
                    labor = f"{key}{value}".upper()
                    break

    return nivel, block, labor


def _bd_operador_exportado(row):
    """Operador leído directamente desde el ZDA."""
    for campo in ("Operador_ZDA", "Operador"):
        valor = row.get(campo)
        if valor is not None and not pd.isna(valor):
            texto = str(valor).strip()
            if texto:
                return texto
    return None


def _bd_cut_stats(df_detalle, jumbo, ciclo):
    if df_detalle.empty or "Longitud_roca_m" not in df_detalle.columns:
        return None, None

    mask = pd.Series(True, index=df_detalle.index)
    if "Fuente" in df_detalle.columns:
        mask &= df_detalle["Fuente"].astype(str).str.upper().eq("ZDA")
    if "Jumbo" in df_detalle.columns:
        mask &= df_detalle["Jumbo"].astype(str).eq(str(jumbo))
    if "Ciclo" in df_detalle.columns:
        mask &= df_detalle["Ciclo"].astype(str).eq(str(ciclo))
    if "Tipo" in df_detalle.columns:
        mask &= df_detalle["Tipo"].astype(str).str.upper().eq("CUT")

    vals = pd.to_numeric(
        df_detalle.loc[mask, "Longitud_roca_m"], errors="coerce"
    ).dropna()

    if vals.empty:
        return None, None
    return float(vals.median()), float(vals.mean())


def _bd_primer_martillo(df_detalle, jumbo, ciclo):
    if df_detalle.empty or "Inicio_Barreno_TS" not in df_detalle.columns:
        return None

    mask = pd.Series(True, index=df_detalle.index)
    if "Fuente" in df_detalle.columns:
        mask &= df_detalle["Fuente"].astype(str).str.upper().eq("ZDA")
    if "Jumbo" in df_detalle.columns:
        mask &= df_detalle["Jumbo"].astype(str).eq(str(jumbo))
    if "Ciclo" in df_detalle.columns:
        mask &= df_detalle["Ciclo"].astype(str).eq(str(ciclo))

    ts = pd.to_numeric(
        df_detalle.loc[mask, "Inicio_Barreno_TS"], errors="coerce"
    ).dropna()
    if ts.empty:
        return None

    d = datetime.fromtimestamp(int(ts.min()), tz=timezone.utc)
    return (d.hour * 3600 + d.minute * 60 + d.second) / 86400.0


def construir_bd_perfo(df_reportes, df_detalle):
    if df_reportes.empty or "Fuente" not in df_reportes.columns:
        return pd.DataFrame(columns=BD_PERFO_COLUMNS)

    zda = df_reportes[
        df_reportes["Fuente"].astype(str).str.upper().eq("ZDA")
    ].copy()
    rows = []

    for _, r in zda.iterrows():
        dt = _bd_parse_datetime(r)
        seccion, roca, tipo_exp = _bd_plan_fields(r.get("Plan_Perforacion"))
        nivel, block, labor = _bd_labor_fields(r)

        auto = _bd_num(r.get("Auto_Total_Brazos_min"))
        manual = _bd_num(r.get("Manual_Total_Brazos_min"))

        if auto is None:
            a1 = _bd_num(r.get("Auto_Brazo1_min"))
            a2 = _bd_num(r.get("Auto_Brazo2_min"))
            if a1 is not None and a2 is not None:
                auto = a1 + a2

        if manual is None:
            m1 = _bd_num(r.get("Manual_Brazo1_min"))
            m2 = _bd_num(r.get("Manual_Brazo2_min"))
            if m1 is not None and m2 is not None:
                manual = m1 + m2

        movimiento = (
            auto + manual
            if auto is not None and manual is not None
            else None
        )
        taladros = _bd_num(r.get("Barrenos_Realizados"))
        t_automatico = auto * 60 / 26.5 if auto is not None else None
        t_manual = (
            taladros - t_automatico
            if taladros is not None and t_automatico is not None
            else None
        )
        obj_automatico = (
            t_automatico / taladros
            if taladros not in (None, 0) and t_automatico is not None
            else None
        )

        auto_b1 = _bd_num(r.get("Auto_Brazo1_min"))
        man_b1 = _bd_num(r.get("Manual_Brazo1_min"))
        auto_b2 = _bd_num(r.get("Auto_Brazo2_min"))
        man_b2 = _bd_num(r.get("Manual_Brazo2_min"))

        pct_auto_drill = (
            auto / (auto + manual)
            if auto is not None and manual is not None and (auto + manual) > 0
            else None
        )
        pct_auto_b1 = (
            auto_b1 / (auto_b1 + man_b1)
            if auto_b1 is not None and man_b1 is not None and (auto_b1 + man_b1) > 0
            else None
        )
        pct_auto_b2 = (
            auto_b2 / (auto_b2 + man_b2)
            if auto_b2 is not None and man_b2 is not None and (auto_b2 + man_b2) > 0
            else None
        )

        med_cut, prom_cut = _bd_cut_stats(
            df_detalle, r.get("Jumbo"), r.get("Ciclo")
        )
        primer_martillo = _bd_primer_martillo(
            df_detalle, r.get("Jumbo"), r.get("Ciclo")
        )

        conlabor = (
            f"{labor} {block} {nivel}"
            if labor and block is not None and nivel is not None
            else None
        )

        tiempo_perf_s = _bd_num(r.get("Tiempo_Perforacion_s"))

        rows.append({
            "MES": _bd_mes_operativo(dt),
            "SEMANA": None,
            "AÑO": dt.year if dt else None,
            "FECHA": datetime(dt.year, dt.month, dt.day) if dt else None,
            "TURNO": _bd_turno(dt),
            "JEFE DE TURNO": None,
            "OPERADOR": _bd_operador_exportado(r),
            "JUMBO": BD_JUMBO_ALIAS.get(r.get("Jumbo"), r.get("Jumbo")),
            "NIVEL": nivel,
            "BLOCK": block,
            "LABOR": labor,
            "TIPO DISPARO": r.get("Tipo_Disparo"),
            "SECCIÓN": seccion,
            "RMR": None,
            "TIP. DE ROCA": roca,
            "TIPO EXP.": tipo_exp,
            "Hora 1° taladro": _bd_hora_excel(dt),
            "Tiempo de perforación efec.": None,
            "Tiempo de perforación": (
                tiempo_perf_s / 86400.0 if tiempo_perf_s is not None else None
            ),
            "Tiempo de movimiento ": (
                movimiento / 1440.0 if movimiento is not None else None
            ),
            "T. mov. Manual": (
                manual / 1440.0 if manual is not None else None
            ),
            "T.mov. Autom.": (
                auto / 1440.0 if auto is not None else None
            ),
            "T. AUTOMÁTICO": t_automatico,
            "# TALADROS": taladros,
            "T. MANUAL": t_manual,
            "Av. Perf": None,
            "Av.Top": None,
            "Av.Real": None,
            "Eficiencia": None,
            "Obj. Automático": obj_automatico,
            "CONLABOR": conlabor,
            "Tipo Mat": None,
            "Ciclo": r.get("Ciclo"),
            "%AutoDrill": pct_auto_drill,
            "% Auto B1": pct_auto_b1,
            "% Auto B2": pct_auto_b2,
            "Mediana Cut (m)": med_cut,
            "Promedio Cut (m)": prom_cut,
            "Hora 1° martillo": primer_martillo,
        })

    return pd.DataFrame(rows, columns=BD_PERFO_COLUMNS)


@st.cache_data(show_spinner=False, max_entries=1)
def crear_excel_publicacion(df_bd_perfo, df_reportes, df_resumen):
    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_bd_perfo.to_excel(writer, sheet_name="BD-PERFO", index=False)
        df_reportes.to_excel(writer, sheet_name="Resumen_Reportes", index=False)
        df_resumen.to_excel(writer, sheet_name="Resumen_Ciclos", index=False)

    buffer.seek(0)
    wb = load_workbook(buffer)

    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = "auto"

    fill = PatternFill(fill_type="solid", fgColor="1F4E78")
    font = Font(color="FFFFFF", bold=True)

    for ws_any in wb.worksheets:
        ws_any.freeze_panes = "A2"
        if ws_any.max_row >= 1 and ws_any.max_column >= 1:
            ws_any.auto_filter.ref = ws_any.dimensions
            for c in ws_any[1]:
                c.fill = fill
                c.font = font
                c.alignment = Alignment(
                    horizontal="center",
                    vertical="center",
                )

    ws = wb["BD-PERFO"]

    formatos = {
        "D": "dd/mm/yyyy",
        "Q": "hh:mm:ss",
        "R": "[h]:mm",
        "S": "[h]:mm",
        "T": "[h]:mm",
        "U": "[h]:mm",
        "V": "[h]:mm",
        "W": "0",
        "Y": "0",
        "AC": "0.0%",
        "AD": "0.0%",
        "AH": "0.0%",
        "AI": "0.0%",
        "AJ": "0.0%",
        "AK": "0.00",
        "AL": "0.00",
        "AM": "hh:mm:ss",
    }

    for row in range(2, ws.max_row + 1):
        ws[f"C{row}"] = f"=YEAR(D{row})"
        ws[f"V{row}"] = f"=T{row}-U{row}"
        ws[f"W{row}"] = f"=V{row}*86400/26.5"
        ws[f"Y{row}"] = f"=X{row}-W{row}"
        ws[f"AC{row}"] = f"=AB{row}/Z{row}"
        ws[f"AD{row}"] = f"=W{row}/X{row}"
        ws[f"AE{row}"] = f'=K{row}&" "&J{row}&" "&I{row}'
        ws[f"AH{row}"] = f"=V{row}/T{row}"

        for col, fmt_code in formatos.items():
            ws[f"{col}{row}"].number_format = fmt_code

    widths = [
        12,10,8,12,10,18,18,10,10,10,16,24,12,10,14,16,16,24,22,22,
        18,18,16,12,14,12,12,12,12,18,24,14,10,14,14,14,16,16,16,
    ]
    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[
            ws.cell(row=1, column=idx).column_letter
        ].width = width

    for nombre_hoja in ["Resumen_Reportes", "Resumen_Ciclos"]:
        ws_tabla = wb[nombre_hoja]
        for col in ws_tabla.columns:
            letra = col[0].column_letter
            ancho = max(
                (len(str(c.value)) for c in col if c.value is not None),
                default=8,
            )
            ws_tabla.column_dimensions[letra].width = min(ancho + 2, 35)

    salida = BytesIO()
    wb.save(salida)
    salida.seek(0)
    return salida.getvalue()

# ==========================================================
# GRÁFICOS CONSOLIDADOS
# ==========================================================


def grafico_auto(df_auto: pd.DataFrame, mostrar_etiquetas: bool, mostrar_linea: bool,
                 tipo_grafico: str = "Líneas"):
    if df_auto.empty:
        return None
    df = df_auto.copy()
    # La población ya llega filtrada desde el bloque de automatización.
    # No se fuerza FRENTE aquí para que SELLADA y ESTOCADA Y/O CORRECCIONES
    # tengan efecto real en el filtro.
    df = df[df["Pct_Movimiento_Automatico_Brazos"].notna()].copy()
    df = asegurar_fechahora(df)
    if df.empty:
        return None

    barras = _es_barras(tipo_grafico)
    fig = go.Figure()
    annotations = []
    points = []
    jumbos = list(df["Jumbo"].dropna().astype(str).unique())
    if barras:
        df = df.reset_index(drop=True)
        df["_x"], cfg_barras = _eje_barras_por_dia(df["FechaHora"])
        df["_fecha_txt"] = df["FechaHora"].dt.strftime("%d/%m/%Y %H:%M")
    for idx, jumbo in enumerate(jumbos):
        g = df[df["Jumbo"].astype(str) == jumbo].sort_values("FechaHora")
        color = COLORES[idx % len(COLORES)]
        custom = g[["Barrenos_Realizados", "Tipo_Roca"]].to_numpy()
        hover = (f"{jumbo}<br>%{{x|%d/%m %H:%M}}<br>Automático: %{{y:.1f}}%"
                 "<br>Barrenos: %{customdata[0]} tal."
                 "<br>Tipo de roca: %{customdata[1]}<extra></extra>")
        if barras:
            # Una barra por ciclo, en orden cronológico y con el mismo ancho.
            fig.add_trace(go.Bar(
                x=g["_x"], y=g["Pct_Movimiento_Automatico_Brazos"], name=jumbo, width=0.8,
                marker=dict(color=color, line=dict(color="#ffffff", width=0.6)),
                text=[_pct_entero_txt(v) for v in g["Pct_Movimiento_Automatico_Brazos"]] if mostrar_etiquetas else None,
                # constraintext="none": la etiqueta conserva su tamaño aunque la barra sea angosta
                # (por defecto Plotly la encoge u oculta cuando no cabe en el ancho de la barra).
                textposition="outside", cliponaxis=False, constraintext="none",
                textfont=dict(size=_font_etiqueta_barras(len(df)), color="#334155"),
                customdata=g[["Barrenos_Realizados", "Tipo_Roca", "_fecha_txt"]].to_numpy(),
                hovertemplate=(f"{jumbo}<br>%{{customdata[2]}}<br>Automático: %{{y:.1f}}%"
                               "<br>Barrenos: %{customdata[0]} tal."
                               "<br>Tipo de roca: %{customdata[1]}<extra></extra>"),
            ))
        else:
            fig.add_trace(go.Scatter(
                x=g["FechaHora"], y=g["Pct_Movimiento_Automatico_Brazos"],
                name=jumbo,
                customdata=custom,
                hovertemplate=hover,
                **_estilo_linea(color, curva=mostrar_linea, ancho=3.0),
            ))
        auto = pd.to_numeric(g["Auto_Total_Brazos_min"], errors="coerce").sum(min_count=1)
        manual = pd.to_numeric(g["Manual_Total_Brazos_min"], errors="coerce").sum(min_count=1)
        kpi = auto / (auto + manual) * 100 if pd.notna(auto) and pd.notna(manual) and (auto + manual) > 0 else None
        annotations.append(dict(
            xref="paper", yref="paper", x=0.01, y=1.18 - idx*0.08, xanchor="left",
            text=f"<b>{jumbo}</b> · Automático global: <b>{fmt(kpi,1,'%')}</b>",
            showarrow=False, bgcolor="rgba(255,255,255,0.92)", bordercolor="#dbe3ea", borderpad=5,
        ))
        if mostrar_etiquetas and not barras:
            # Etiqueta: solo el porcentaje, entero (el detalle con decimales queda en el hover).
            for i, (_, r) in enumerate(g.iterrows()):
                points.append(dict(
                    x=r["FechaHora"], y=r["Pct_Movimiento_Automatico_Brazos"],
                    text=_pct_entero_txt(r["Pct_Movimiento_Automatico_Brazos"]),
                    rank=i+idx,
                ))
    if mostrar_etiquetas and not barras:
        annotations.extend(smart_annotations(points, x_window_hours=18, y_window=4, font_size=11))
    yaxis = dict(title="Movimiento automático (%)", rangemode="tozero", gridcolor="#eef2f7")
    extra = {}
    if barras:
        ymax = float(df["Pct_Movimiento_Automatico_Brazos"].max())
        yaxis["range"] = [0, max(ymax * (1.18 if mostrar_etiquetas else 1.08), 10)]
        extra.update(barmode="overlay", shapes=cfg_barras["shapes"])
    fig.update_layout(**base_layout(
        470, annotations=annotations, margin=dict(l=78,r=35,t=120,b=72),
        yaxis=yaxis,
        xaxis=cfg_barras["xaxis"] if barras else eje_fechas_dias(df["FechaHora"]),
        **extra,
    ))
    return fig


COLOR_OPERADOR_AZUL = "#4F67F2"
COLOR_OPERADOR_VERDE = "#16C48A"
COLOR_OPERADOR_NARANJA = "#F59E0B"
COLOR_OPERADOR_ROJO = "#EF4444"
COLOR_OPERADOR_SIN_REGISTRO = "#111111"


def construir_colores_operador_por_ranking(df_auto: pd.DataFrame):
    """
    Colores dinámicos según horas automáticas acumuladas del conjunto visible.

    Regla:
    - Sin registrar -> negro
    - Mayor horas   -> azul
    - Segundo       -> verde
    - Tercero       -> naranja
    - Menor horas   -> rojo

    Para 1, 2 o 3 operadores se conserva el sentido de los extremos:
    1 -> azul
    2 -> azul / rojo
    3 -> azul / verde / rojo

    Si en el futuro hubiera más de 4 operadores, los intermedios adicionales
    se muestran en naranja y el de menor acumulado permanece rojo.
    """
    color_map = {
        "Sin registrar": COLOR_OPERADOR_SIN_REGISTRO,
    }

    if (
        df_auto is None
        or df_auto.empty
        or "Operador_Filtro" not in df_auto.columns
        or "Auto_Total_Brazos_min" not in df_auto.columns
    ):
        return color_map

    work = df_auto.copy()
    work["_Operador_Color"] = (
        work["Operador_Filtro"]
        .fillna("SIN DATO")
        .astype(str)
        .str.strip()
    )
    work["_Auto_min_color"] = pd.to_numeric(
        work["Auto_Total_Brazos_min"],
        errors="coerce",
    )

    op_upper = work["_Operador_Color"].str.upper()
    work = work[
        ~op_upper.isin(["", "SIN DATO", "NONE", "NAN"])
        & work["_Auto_min_color"].notna()
    ].copy()

    if work.empty:
        return color_map

    ranking = (
        work.groupby("_Operador_Color", as_index=False)
        .agg(Horas_auto=("_Auto_min_color", "sum"))
        .sort_values(
            ["Horas_auto", "_Operador_Color"],
            ascending=[False, True],
        )
        .reset_index(drop=True)
    )

    operadores = ranking["_Operador_Color"].astype(str).tolist()
    n = len(operadores)

    if n == 1:
        colores = [COLOR_OPERADOR_AZUL]
    elif n == 2:
        colores = [
            COLOR_OPERADOR_AZUL,
            COLOR_OPERADOR_ROJO,
        ]
    elif n == 3:
        colores = [
            COLOR_OPERADOR_AZUL,
            COLOR_OPERADOR_VERDE,
            COLOR_OPERADOR_ROJO,
        ]
    elif n == 4:
        colores = [
            COLOR_OPERADOR_AZUL,
            COLOR_OPERADOR_VERDE,
            COLOR_OPERADOR_NARANJA,
            COLOR_OPERADOR_ROJO,
        ]
    else:
        colores = (
            [COLOR_OPERADOR_AZUL, COLOR_OPERADOR_VERDE]
            + [COLOR_OPERADOR_NARANJA] * max(0, n - 3)
            + [COLOR_OPERADOR_ROJO]
        )

    color_map.update(dict(zip(operadores, colores)))
    return color_map


def grafico_auto_por_operador(
    df_auto: pd.DataFrame,
    mostrar_etiquetas: bool,
    mostrar_linea: bool,
    tipo_grafico: str = "Líneas",
):
    """
    Evolución del movimiento automático por operador.

    - Cada línea/color representa un operador.
    - Cada punto representa un ciclo/round.
    - El % global mostrado en la leyenda se calcula a partir de la suma
      de tiempos automáticos y manuales del operador:
          Auto / (Auto + Manual)
      y no como promedio simple de porcentajes por ciclo.
    """
    if df_auto.empty:
        return None

    requeridas = {
        "Operador_Filtro",
        "Pct_Movimiento_Automatico_Brazos",
    }
    if not requeridas.issubset(df_auto.columns):
        return None

    df = df_auto.copy()

    df["Operador_Filtro"] = (
        df["Operador_Filtro"]
        .fillna("SIN DATO")
        .astype(str)
        .str.strip()
    )

    # Para una comparación realmente "por operador", no dibujar
    # registros donde no se logró identificar a la persona.
    df = df[
        df["Operador_Filtro"].ne("")
        & df["Operador_Filtro"].str.upper().ne("SIN DATO")
        & df["Pct_Movimiento_Automatico_Brazos"].notna()
    ].copy()

    df = asegurar_fechahora(df)
    if df.empty:
        return None

    operadores = sorted(
        df["Operador_Filtro"]
        .dropna()
        .astype(str)
        .unique()
    )

    fig = go.Figure()
    annotations = []
    points = []
    etiquetas_finales = []

    # Símbolo permite reconocer el jumbo sin competir con el color,
    # que queda reservado para distinguir operadores.
    jumbo_symbols = {
        "JUMB001": "circle",
        "JUMB002": "square",
    }

    # La paleta se asigna dinámicamente según las horas automáticas
    # acumuladas del rango/filtros actualmente visibles.
    color_map_operador = construir_colores_operador_por_ranking(df)

    # Modo barras: una barra por ciclo, en orden cronológico, del color de su operador.
    barras = _es_barras(tipo_grafico)
    if barras:
        df = df.reset_index(drop=True)
        df["_x"], cfg_barras = _eje_barras_por_dia(df["FechaHora"])
        df["_fecha_txt"] = df["FechaHora"].dt.strftime("%d/%m/%Y %H:%M")

    for idx, operador in enumerate(operadores):
        color_operador = color_map_operador.get(
            str(operador),
            "#64748b",
        )

        g = df[
            df["Operador_Filtro"].astype(str).eq(str(operador))
        ].sort_values("FechaHora").copy()

        if g.empty:
            continue

        auto = pd.to_numeric(
            g.get("Auto_Total_Brazos_min"),
            errors="coerce",
        ).sum(min_count=1)

        manual = pd.to_numeric(
            g.get("Manual_Total_Brazos_min"),
            errors="coerce",
        ).sum(min_count=1)

        global_pct = (
            auto / (auto + manual) * 100
            if pd.notna(auto)
            and pd.notna(manual)
            and (auto + manual) > 0
            else None
        )

        ciclos = (
            g["Ciclo"]
            if "Ciclo" in g.columns
            else pd.Series(["-"] * len(g), index=g.index)
        )
        jumbos = (
            g["Jumbo"].fillna("-").astype(str)
            if "Jumbo" in g.columns
            else pd.Series(["-"] * len(g), index=g.index)
        )
        tipos = (
            g["Tipo_Disparo"].fillna("-").astype(str)
            if "Tipo_Disparo" in g.columns
            else pd.Series(["-"] * len(g), index=g.index)
        )
        rocas = (
            g["Tipo_Roca"].fillna("SIN DATO").astype(str)
            if "Tipo_Roca" in g.columns
            else pd.Series(["SIN DATO"] * len(g), index=g.index)
        )
        barrenos = (
            g["Barrenos_Realizados"]
            if "Barrenos_Realizados" in g.columns
            else pd.Series([None] * len(g), index=g.index)
        )

        cols_custom = [
            ciclos.astype(object),
            jumbos.astype(object),
            tipos.astype(object),
            rocas.astype(object),
            barrenos.astype(object),
        ]
        if barras:
            cols_custom.append(g["_fecha_txt"].astype(object))
        custom = np.column_stack(cols_custom)

        symbols = [
            jumbo_symbols.get(str(j), "diamond")
            for j in jumbos
        ]

        nombre_leyenda = (
            f"{operador} · Global {global_pct:.1f}%"
            if global_pct is not None and pd.notna(global_pct)
            else operador
        )

        hover_op = (
            f"<b>{operador}</b>"
            "<br>Fecha: " + ("%{customdata[5]}" if barras else "%{x|%d/%m/%Y %H:%M}") +
            "<br>Jumbo: %{customdata[1]}"
            "<br>Ciclo: %{customdata[0]}"
            "<br>Automático: %{y:.1f}%"
            "<br>Barrenos: %{customdata[4]}"
            "<br>Tipo de disparo: %{customdata[2]}"
            "<br>Tipo de roca: %{customdata[3]}"
            "<extra></extra>"
        )
        if barras:
            fig.add_trace(
                go.Bar(
                    x=g["_x"],
                    y=g["Pct_Movimiento_Automatico_Brazos"],
                    name=nombre_leyenda,
                    width=0.8,
                    marker=dict(
                        color=color_operador,
                        line=dict(color="#ffffff", width=0.6),
                    ),
                    text=(
                        [_pct_entero_txt(v) for v in g["Pct_Movimiento_Automatico_Brazos"]]
                        if mostrar_etiquetas else None
                    ),
                    textposition="outside",
                    cliponaxis=False,
                    constraintext="none",
                    textfont=dict(size=_font_etiqueta_barras(len(df)), color="#334155"),
                    customdata=custom,
                    hovertemplate=hover_op,
                )
            )
        else:
            fig.add_trace(
                go.Scatter(
                    x=g["FechaHora"],
                    y=g["Pct_Movimiento_Automatico_Brazos"],
                    name=nombre_leyenda,
                    customdata=custom,
                    hovertemplate=hover_op,
                    **_estilo_linea(color_operador, curva=mostrar_linea, simbolo=symbols, ancho=2.8),
                )
            )

        if mostrar_etiquetas and not barras:
            for i, (_, r) in enumerate(g.iterrows()):
                points.append({
                    "x": r["FechaHora"],
                    "y": r["Pct_Movimiento_Automatico_Brazos"],
                    "text": (
                        f"{operador.split()[-1]}<br>"
                        f"{_pct_entero_txt(r['Pct_Movimiento_Automatico_Brazos'])}"
                    ),
                    "rank": i + idx * 100,
                })

        # Guardar la etiqueta final para posicionarla después.
        # El ajuste conjunto permite evitar superposición entre apellidos.
        ultimo = g.iloc[-1]
        if not barras:
            etiquetas_finales.append({
                "x": ultimo["FechaHora"],
                "y": float(ultimo["Pct_Movimiento_Automatico_Brazos"]),
                "texto": operador.split()[-1],
                "color": color_operador,
            })

    # ------------------------------------------------------
    # Etiquetas finales sin superposición
    # ------------------------------------------------------
    # Si dos o más operadores terminan con valores muy próximos,
    # sus apellidos se desplazan verticalmente en píxeles.
    # Los puntos y las curvas permanecen exactamente en su valor real.
    if etiquetas_finales:
        etiquetas_ordenadas = sorted(
            etiquetas_finales,
            key=lambda e: e["y"],
            reverse=True,
        )

        # Umbral en puntos porcentuales para considerar que dos etiquetas
        # podrían superponerse visualmente.
        umbral_pp = 6.0
        clusters = []
        cluster_actual = []

        for etiqueta in etiquetas_ordenadas:
            if not cluster_actual:
                cluster_actual = [etiqueta]
                continue

            if abs(cluster_actual[-1]["y"] - etiqueta["y"]) <= umbral_pp:
                cluster_actual.append(etiqueta)
            else:
                clusters.append(cluster_actual)
                cluster_actual = [etiqueta]

        if cluster_actual:
            clusters.append(cluster_actual)

        # Desplazamientos simétricos alrededor de la posición real.
        # Ejemplos:
        # 2 etiquetas -> +11 / -11 px
        # 3 etiquetas -> +22 / 0 / -22 px
        # 4 etiquetas -> +33 / +11 / -11 / -33 px
        for cluster in clusters:
            n = len(cluster)
            paso_px = 22
            centro = (n - 1) / 2.0

            for pos, etiqueta in enumerate(cluster):
                yshift = int(round((centro - pos) * paso_px))

                annotations.append(
                    dict(
                        x=etiqueta["x"],
                        y=etiqueta["y"],
                        xref="x",
                        yref="y",
                        text=f'<b>{etiqueta["texto"]}</b>',
                        showarrow=False,
                        xanchor="left",
                        yanchor="middle",
                        xshift=14,
                        yshift=yshift,
                        font=dict(
                            size=11,
                            color=etiqueta["color"],
                        ),
                        bgcolor="rgba(255,255,255,0.90)",
                        bordercolor="rgba(0,0,0,0)",
                        borderpad=2,
                    )
                )

    if mostrar_etiquetas and points:
        annotations.extend(
            smart_annotations(
                points,
                x_window_hours=18,
                y_window=4,
                font_size=10,
            )
        )

    yaxis_op = dict(
        title="Movimiento automático (%)",
        rangemode="tozero",
        gridcolor="#eef2f7",
    )
    extra_op = {}
    if barras:
        ymax_op = float(df["Pct_Movimiento_Automatico_Brazos"].max())
        yaxis_op["range"] = [0, max(ymax_op * (1.18 if mostrar_etiquetas else 1.08), 10)]
        extra_op.update(barmode="overlay", shapes=cfg_barras["shapes"])

    fig.update_layout(
        **base_layout(
            500,
            annotations=annotations,
            margin=dict(
                l=78,
                # En líneas se reserva espacio para el apellido al final de cada serie.
                r=35 if barras else 145,
                t=55,
                b=82,
            ),
            yaxis=yaxis_op,
            xaxis=cfg_barras["xaxis"] if barras else eje_fechas_dias(df["FechaHora"]),
            legend=dict(
                orientation="h",
                y=-0.20,
                x=0,
            ),
            hovermode="closest",
            **extra_op,
        )
    )

    return fig


def grafico_uso_auto_promedio_operador(
    df_auto: pd.DataFrame,
    operadores_visibles=None,
):
    """
    Uso promedio del movimiento automático (%) por operador.

    - % = horas en automático / (automático + manual) de TODOS los ciclos visibles del
      operador. Es el mismo "Global" que figura en la leyenda de la evolución por operador,
      de modo que ambos gráficos muestran la misma cifra.
    - En el hover se agrega el promedio simple de los porcentajes de cada ciclo.
    - Solo operadores identificados: los ciclos sin operador no se muestran en este gráfico.
    - Se muestra solo el apellido.
    - Cada operador conserva el color del gráfico de evolución por operador.
    """
    if df_auto is None or df_auto.empty:
        return None

    if not {"Auto_Total_Brazos_min", "Manual_Total_Brazos_min"}.issubset(df_auto.columns):
        return None

    df = df_auto.copy()

    # Operador normalizado para agrupación.
    if "Operador_Filtro" in df.columns:
        operador_raw = (
            df["Operador_Filtro"]
            .fillna("SIN DATO")
            .astype(str)
            .str.strip()
        )
    else:
        operador_raw = pd.Series(
            ["SIN DATO"] * len(df),
            index=df.index,
            dtype=object,
        )

    es_sin_registro = operador_raw.str.upper().isin(["", "SIN DATO", "NONE", "NAN"])

    df["_Operador_Agrupado"] = operador_raw
    df = df[~es_sin_registro].copy()

    if operadores_visibles is not None:
        operadores_sel = {str(x).strip() for x in operadores_visibles}
        df = df[df["_Operador_Agrupado"].isin(operadores_sel)].copy()

    df["_Auto_min"] = pd.to_numeric(df["Auto_Total_Brazos_min"], errors="coerce")
    df["_Manual_min"] = pd.to_numeric(df["Manual_Total_Brazos_min"], errors="coerce")
    df["_Pct_ciclo"] = (
        pd.to_numeric(df["Pct_Movimiento_Automatico_Brazos"], errors="coerce")
        if "Pct_Movimiento_Automatico_Brazos" in df.columns else np.nan
    )
    df = df[df["_Auto_min"].notna() & df["_Manual_min"].notna()].copy()

    if df.empty:
        return None

    resumen = (
        df.groupby("_Operador_Agrupado", as_index=False)
        .agg(
            Auto_min=("_Auto_min", "sum"),
            Manual_min=("_Manual_min", "sum"),
            Pct_promedio_ciclos=("_Pct_ciclo", "mean"),
            Ciclos=("_Auto_min", "size"),
        )
    )
    total_min = resumen["Auto_min"] + resumen["Manual_min"]
    resumen["Pct_uso"] = np.where(total_min > 0, resumen["Auto_min"] / total_min.where(total_min > 0) * 100.0, np.nan)
    resumen = resumen[resumen["Pct_uso"].notna()].copy()

    if resumen.empty:
        return None

    # Misma paleta que el gráfico de curvas (identidad de color por operador).
    color_map = construir_colores_operador_por_ranking(df)

    def _apellido(nombre):
        partes = str(nombre).split()
        return partes[-1] if partes else str(nombre)

    resumen["Etiqueta"] = resumen["_Operador_Agrupado"].apply(_apellido)
    resumen["Color"] = (
        resumen["_Operador_Agrupado"]
        .map(color_map)
        .fillna("#64748b")
    )

    # Orden visual: de mayor a menor porcentaje.
    resumen = resumen.sort_values(
        ["Pct_uso", "Etiqueta"],
        ascending=[False, True],
    ).reset_index(drop=True)

    # Fondo tipo "track": representa el 100 %.
    fig = go.Figure()

    fig.add_trace(
        go.Bar(
            y=resumen["Etiqueta"],
            x=[100.0] * len(resumen),
            orientation="h",
            marker=dict(
                color="#F2F1EC",
                line=dict(width=0),
            ),
            width=0.22,
            hoverinfo="skip",
            showlegend=False,
        )
    )

    fig.add_trace(
        go.Bar(
            y=resumen["Etiqueta"],
            x=resumen["Pct_uso"],
            orientation="h",
            marker=dict(
                color=resumen["Color"].tolist(),
                line=dict(width=0),
            ),
            width=0.22,
            text=[_pct_entero_txt(v) for v in resumen["Pct_uso"]],
            textposition="outside",
            textfont=dict(
                size=13,
                color="#6b6b66",
            ),
            cliponaxis=False,
            customdata=np.column_stack([
                resumen["_Operador_Agrupado"].astype(object),
                resumen["Ciclos"].astype(object),
                resumen["Pct_promedio_ciclos"].map(
                    lambda v: f"{v:.1f}%" if pd.notna(v) else "N/D"
                ).astype(object),
            ]),
            hovertemplate=(
                "<b>%{customdata[0]}</b>"
                "<br>Uso automático promedio: %{x:.1f}%"
                "<br>Promedio simple de los ciclos: %{customdata[2]}"
                "<br>Ciclos con dato: %{customdata[1]}"
                "<extra></extra>"
            ),
            showlegend=False,
        )
    )

    fig.update_layout(
        **base_layout(
            max(330, 110 + len(resumen) * 58),
            margin=dict(
                l=125,
                r=95,
                t=20,
                b=35,
            ),
            barmode="overlay",
            bargap=0.46,
            xaxis=dict(
                range=[0, 112],
                showgrid=False,
                showticklabels=False,
                zeroline=False,
                title=None,
                fixedrange=True,
            ),
            yaxis=dict(
                title=None,
                categoryorder="array",
                categoryarray=resumen["Etiqueta"].tolist(),
                autorange="reversed",
                showgrid=False,
                zeroline=False,
                tickfont=dict(
                    size=14,
                    color="#73726d",
                ),
                fixedrange=True,
            ),
            plot_bgcolor="#ffffff",
            paper_bgcolor="#ffffff",
            hovermode="closest",
            showlegend=False,
        )
    )

    return fig


def grafico_brazos(df_auto: pd.DataFrame, jumbo: str, mostrar_etiquetas: bool,
                   tipo_grafico: str = "Líneas"):
    df = df_auto[df_auto["Jumbo"].astype(str) == str(jumbo)].copy()
    # La población ya llega filtrada desde el bloque de automatización.
    df = df[df["Pct_Automatico_Brazo1"].notna() | df["Pct_Automatico_Brazo2"].notna()].copy()
    df = asegurar_fechahora(df)
    if df.empty:
        return None

    barras = _es_barras(tipo_grafico)
    fig = go.Figure()
    annotations = []
    points = []
    series = [
        ("Brazo 1", "Pct_Automatico_Brazo1", "Auto_Brazo1_min", "Manual_Brazo1_min", 0),
        ("Brazo 2", "Pct_Automatico_Brazo2", "Auto_Brazo2_min", "Manual_Brazo2_min", 1),
    ]
    if barras:
        df = df.reset_index(drop=True)
        df["_x"], cfg_barras = _eje_barras_por_dia(df["FechaHora"])
        df["_fecha_txt"] = df["FechaHora"].dt.strftime("%d/%m/%Y %H:%M")
    for nombre, col, auto_col, man_col, idx in series:
        g = df[df[col].notna()].sort_values("FechaHora")
        if g.empty:
            continue
        auto = pd.to_numeric(g[auto_col], errors="coerce").sum(min_count=1)
        manual = pd.to_numeric(g[man_col], errors="coerce").sum(min_count=1)
        global_pct = auto / (auto + manual) * 100 if pd.notna(auto) and pd.notna(manual) and (auto + manual) > 0 else None
        hover = f"{nombre}<br>%{{x|%d/%m %H:%M}}<br>Automático: %{{y:.1f}}%<extra></extra>"
        if barras:
            # Brazo 1 y Brazo 2 del mismo ciclo, lado a lado dentro de su posición.
            fig.add_trace(go.Bar(
                x=g["_x"] + (idx - 0.5) * 0.4, y=g[col],
                name=f"{nombre} · Global {fmt(global_pct,1,'%')}", width=0.38,
                marker=dict(color=COLORES[idx], line=dict(color="#ffffff", width=0.6)),
                text=[_pct_entero_txt(v) for v in g[col]] if mostrar_etiquetas else None,
                # Etiqueta vertical: las dos barras de un ciclo quedan a ~15 px una de otra.
                textposition="outside", cliponaxis=False, constraintext="none", textangle=-90,
                textfont=dict(size=_font_etiqueta_barras(2 * len(df)), color="#334155"),
                customdata=g[["_fecha_txt"]].to_numpy(),
                hovertemplate=f"{nombre}<br>%{{customdata[0]}}<br>Automático: %{{y:.1f}}%<extra></extra>",
            ))
        else:
            fig.add_trace(go.Scatter(
                x=g["FechaHora"], y=g[col],
                name=f"{nombre} · Global {fmt(global_pct,1,'%')}",
                hovertemplate=hover,
                **_estilo_linea(COLORES[idx], curva=True, ancho=2.8),
            ))
        annotations.append(dict(
            xref="paper", yref="paper", x=0.01 + idx*0.25, y=1.16, xanchor="left",
            text=f"<b>{nombre} global: {fmt(global_pct,1,'%')}</b>", showarrow=False,
            bgcolor="rgba(255,255,255,0.92)", bordercolor="#dbe3ea", borderpad=5,
        ))
        if mostrar_etiquetas and not barras:
            for i, (_, r) in enumerate(g.iterrows()):
                points.append(dict(x=r["FechaHora"], y=r[col], text=_pct_entero_txt(r[col]), rank=i+idx*100))
    if mostrar_etiquetas and not barras:
        annotations.extend(smart_annotations(points, x_window_hours=18, y_window=4, font_size=10))
    yaxis = dict(title="Movimiento automático (%)", rangemode="tozero", gridcolor="#eef2f7")
    extra = {}
    if barras:
        ymax = float(pd.concat([pd.to_numeric(df[c], errors="coerce") for _, c, _, _, _ in series]).max())
        yaxis["range"] = [0, max(ymax * (1.25 if mostrar_etiquetas else 1.08), 10)]
        extra.update(barmode="overlay", shapes=cfg_barras["shapes"])
    fig.update_layout(**base_layout(
        430, title=dict(text=f"<b>{jumbo}</b>", x=0.01), annotations=annotations,
        margin=dict(l=72,r=30,t=92,b=72),
        yaxis=yaxis,
        xaxis=cfg_barras["xaxis"] if barras else eje_fechas_dias(df["FechaHora"]),
        **extra,
    ))
    return fig


def preparar_cut(df_resumen: pd.DataFrame, df_reportes: pd.DataFrame) -> pd.DataFrame:
    if df_resumen.empty:
        return pd.DataFrame()
    cut = df_resumen[df_resumen["Tipo"].astype(str).str.lower() == "cut"].copy()
    cut = cut[cut["Mediana"].notna()].copy()
    if cut.empty:
        return cut
    cols = [c for c in ["Jumbo","Ciclo","Fecha_Inicio","Tipo_Disparo","Tipo_Roca","Operador_Filtro"] if c in df_reportes.columns]
    tipos = df_reportes[cols].drop_duplicates(subset=[c for c in ["Jumbo","Ciclo","Fecha_Inicio"] if c in cols])
    cut = cut.merge(tipos, on=["Jumbo","Ciclo","Fecha_Inicio"], how="left")
    cut["Tipo_Disparo"] = cut["Tipo_Disparo"].fillna("SIN CLASIFICAR")
    return asegurar_fechahora(cut)


def grafico_cut(
    df_cut: pd.DataFrame,
    jumbos_visibles,
    tipos_visibles,
    rocas_visibles,
    operadores_visibles,
    mostrar_etiquetas: bool,
    tipo_grafico: str = "Líneas",
):
    if df_cut.empty:
        return None

    df = df_cut[
        df_cut["Jumbo"].astype(str).isin([str(x) for x in jumbos_visibles])
        & df_cut["Tipo_Disparo"].isin(tipos_visibles)
        & df_cut["Tipo_Roca"].isin(rocas_visibles)
        & df_cut["Operador_Filtro"].isin(operadores_visibles)
    ].copy()

    if df.empty:
        return None

    barras = _es_barras(tipo_grafico)
    all_jumbos = sorted(df_cut["Jumbo"].dropna().astype(str).unique())
    visibles = sorted(df["Jumbo"].dropna().astype(str).unique())

    if barras:
        # Una barra por ciclo, en orden cronológico y agrupadas por día (ver _eje_barras_por_dia).
        df = df.sort_values(["FechaHora", "Jumbo", "Ciclo"]).reset_index(drop=True)
        df["_x"], cfg_barras = _eje_barras_por_dia(df["FechaHora"])
        df["_fecha_txt"] = df["FechaHora"].dt.strftime("%d/%m/%Y %H:%M")

    fig = go.Figure()
    points = []
    annotations = []

    for pos_visible, jumbo in enumerate(visibles):
        idx = all_jumbos.index(jumbo)
        g = df[df["Jumbo"].astype(str) == jumbo].sort_values("FechaHora")
        color = COLORES[idx % len(COLORES)]

        if barras:
            fig.add_trace(go.Bar(
                x=g["_x"],
                y=g["Mediana"],
                name=jumbo,
                width=0.8,
                marker=dict(color=color, line=dict(color="#ffffff", width=0.6)),
                # En barras la etiqueta va sin unidad (el eje ya dice "m") para que quepa.
                text=[f"{v:.2f}" for v in g["Mediana"]] if mostrar_etiquetas else None,
                textposition="outside", cliponaxis=False, constraintext="none",
                textfont=dict(size=_font_etiqueta_barras(len(df)), color="#334155"),
                customdata=g[["Ciclo", "Tipo_Disparo", "_fecha_txt"]].to_numpy(),
                hovertemplate=(
                    f"{jumbo}<br>%{{customdata[2]}}"
                    "<br>Ciclo: %{customdata[0]}"
                    "<br>Tipo: %{customdata[1]}"
                    "<br>Mediana Cut: %{y:.2f} m<extra></extra>"
                ),
            ))
        else:
            fig.add_trace(go.Scatter(
                x=g["FechaHora"],
                y=g["Mediana"],
                name=jumbo,
                customdata=g[["Ciclo", "Tipo_Disparo"]].to_numpy(),
                **_estilo_linea(color, curva=True, ancho=3.0),
                hovertemplate=(
                    f"{jumbo}<br>%{{x|%d/%m %H:%M}}"
                    "<br>Ciclo: %{customdata[0]}"
                    "<br>Tipo: %{customdata[1]}"
                    "<br>Mediana Cut: %{y:.2f} m<extra></extra>"
                ),
            ))

        # Global coherente con la gráfica:
        # mediana de las medianas Cut de los ciclos visibles.
        global_cut = pd.to_numeric(g["Mediana"], errors="coerce").dropna().median()
        if pd.notna(global_cut):
            annotations.append(dict(
                xref="paper",
                yref="paper",
                x=0.01 + pos_visible * 0.32,
                y=1.16,
                xanchor="left",
                showarrow=False,
                text=f"<b>{jumbo}</b> · Longitud global Cut: <b>{global_cut:.2f} m</b>",
                font=dict(
                    size=12,
                    color=color,
                ),
                bgcolor="rgba(255,255,255,0.95)",
                bordercolor="#dbe3ea",
                borderpad=5,
            ))

        if mostrar_etiquetas and not barras:
            for i, (_, r) in enumerate(g.iterrows()):
                points.append(dict(
                    x=r["FechaHora"],
                    y=r["Mediana"],
                    text=f"{r['Mediana']:.2f} m",
                    rank=i + idx * 50,
                ))

    if mostrar_etiquetas and not barras:
        annotations.extend(
            smart_annotations(
                points,
                x_window_hours=18,
                y_window=0.18,
                font_size=11,
            )
        )

    yaxis = dict(
        title="Mediana de longitud perforada Cut (m)",
        gridcolor="#eef2f7",
    )
    extra = {}
    if barras:
        # Barras: el eje parte de 0 (una barra con eje truncado exagera las diferencias).
        ymax = float(pd.to_numeric(df["Mediana"], errors="coerce").max())
        yaxis["range"] = [0, ymax * (1.15 if mostrar_etiquetas else 1.05)]
        extra.update(barmode="overlay", shapes=cfg_barras["shapes"])

    fig.update_layout(**base_layout(
        450,
        annotations=annotations,
        margin=dict(l=80, r=30, t=88, b=70),
        yaxis=yaxis,
        xaxis=cfg_barras["xaxis"] if barras else eje_fechas_dias(df["FechaHora"]),
        **extra,
    ))
    return fig

# ==========================================================
# ZDA: RESÚMENES Y TIMELINE
# ==========================================================


def format_duration_hms(segundos) -> str:
    """Duración en segundos -> "HH:MM:SS" (mismo formato que Tiempo_Perforacion_hms)."""
    total = int(round(float(segundos)))
    h, resto = divmod(max(total, 0), 3600)
    m, s = divmod(resto, 60)
    return f"{h:02d}:{m:02d}:{s:02d}"


def _utc_dt(ts):
    return datetime.fromtimestamp(int(ts), tz=timezone.utc)


def zda_operational_date(ts):
    d = _utc_dt(ts)
    base = datetime(d.year, d.month, d.day, tzinfo=timezone.utc)
    if d.hour < 7:
        base -= timedelta(days=1)
    return base


def zda_operational_hour(ts, op_date):
    d = _utc_dt(ts)
    day_diff = (datetime(d.year,d.month,d.day,tzinfo=timezone.utc) - op_date).days
    return d.hour + d.minute/60 + d.second/3600 + 24*day_diff


def zda_turno(ts):
    d = _utc_dt(ts)
    h = d.hour + d.minute/60 + d.second/3600
    if 7 <= h < 19:
        return "Día", h
    return "Noche", h+24 if h < 7 else h


def turno_ciclo(df: pd.DataFrame) -> pd.Series:
    """Turno de cada ciclo según su HORA DE INICIO (misma regla que `zda_turno`):
    Día = 07:00 a 19:00; Noche = 19:00 a 07:00. Devuelve None donde no hay hora."""
    if "FechaHora" in df.columns:
        ts = pd.to_datetime(df["FechaHora"], errors="coerce")
    elif "Hora_Inicio" in df.columns:
        ts = pd.to_datetime(df["Hora_Inicio"].astype(str), format="%H:%M:%S", errors="coerce")
    else:
        return pd.Series([None] * len(df), index=df.index, dtype=object)
    h = ts.dt.hour + ts.dt.minute / 60 + ts.dt.second / 3600
    turno = np.select([h.isna(), (h >= 7) & (h < 19)], [None, "Día"], default="Noche")
    return pd.Series(turno, index=df.index, dtype=object)


def turno_de_hora(hora) -> "str | None":
    """Turno de una hora "HH:MM:SS" (o de un datetime): Día 07:00-19:00, Noche el resto.
    None si no hay una hora válida. Misma regla que `turno_ciclo` y `zda_turno`."""
    if hora is None:
        return None
    try:
        t = hora if hasattr(hora, "hour") else pd.to_datetime(str(hora), format="%H:%M:%S")
    except (ValueError, TypeError):
        return None
    if pd.isna(t):
        return None
    h = t.hour + t.minute / 60 + t.second / 3600
    return "Día" if 7 <= h < 19 else "Noche"


def turno_permitido(hora, sel_turnos) -> bool:
    """¿Un ciclo que inició a `hora` entra en los turnos elegidos? (con ambos turnos, siempre)."""
    if sel_turnos is None or set(TURNOS_FILTRO) <= set(sel_turnos):
        return True
    return turno_de_hora(hora) in set(sel_turnos)


def filtrar_por_turno(df: pd.DataFrame, sel_turnos, col_ts=None) -> pd.DataFrame:
    """Filtra por turno. Con Día y Noche marcados (o sin filtro) devuelve el df intacto;
    con uno solo, únicamente sus ciclos; con ninguno, un df vacío.

    `col_ts`: columna con el instante de inicio de la perforación (epoch). Solo la usa
    Primer Golpe, que etiqueta sus turnos con `zda_turno` sobre ese instante; el resto usa
    la hora de inicio del ciclo (igual que el TURNO del Excel exportado)."""
    if df is None or df.empty or sel_turnos is None:
        return df
    elegidos = set(sel_turnos)
    if set(TURNOS_FILTRO) <= elegidos:
        return df
    if col_ts and col_ts in df.columns:
        turno = df[col_ts].map(lambda x: zda_turno(x)[0] if pd.notna(x) else None)
    else:
        turno = turno_ciclo(df)
    return df[turno.isin(elegidos)].copy()


def fmt_hora_decimal(h):
    if h is None or pd.isna(h):
        return "-"
    total = int(round(float(h)*60)) % (24*60)
    return f"{total//60:02d}:{total%60:02d}"


def resumen_turnos_zda(rows: pd.DataFrame) -> pd.DataFrame:
    if rows.empty:
        return pd.DataFrame()
    work = rows.copy()
    work["_turno"] = work["Inicio_Perforacion_TS"].apply(lambda x: zda_turno(x)[0])
    work["_opHour"] = work["Inicio_Perforacion_TS"].apply(lambda x: zda_turno(x)[1])
    work["_opDate"] = work["Inicio_Perforacion_TS"].apply(zda_operational_date)

    # Primer round por jumbo, fecha operativa y turno para indicadores de inicio.
    first_idx = work.groupby(["Jumbo","_opDate","_turno"])["Inicio_Perforacion_TS"].idxmin()
    first = work.loc[first_idx].copy()
    salida = []
    for jumbo in sorted(work["Jumbo"].dropna().astype(str).unique()):
        allj = work[work["Jumbo"].astype(str)==jumbo]
        firstj = first[first["Jumbo"].astype(str)==jumbo]
        day_all = allj[allj["_turno"]=="Día"]
        night_all = allj[allj["_turno"]=="Noche"]
        day = firstj[firstj["_turno"]=="Día"]
        night = firstj[firstj["_turno"]=="Noche"]
        salida.append({
            "Jumbo": jumbo,
            "Ciclos día": len(day_all),
            "Inicio prom. día": fmt_hora_decimal(day["_opHour"].mean()) if not day.empty else "-",
            "Inicio más temprano día": fmt_hora_decimal(day["_opHour"].min()) if not day.empty else "-",
            "Inicio más tarde día": fmt_hora_decimal(day["_opHour"].max()) if not day.empty else "-",
            "Ciclos noche": len(night_all),
            "Inicio prom. noche": fmt_hora_decimal(night["_opHour"].mean()) if not night.empty else "-",
            "Inicio más temprano noche": fmt_hora_decimal(night["_opHour"].min()) if not night.empty else "-",
            "Inicio más tarde noche": fmt_hora_decimal(night["_opHour"].max()) if not night.empty else "-",
        })
    return pd.DataFrame(salida)


def resumen_fin_turnos_zda(rows: pd.DataFrame) -> pd.DataFrame:
    """
    Resume la hora de término del último round por jumbo y turno.

    - Ciclos día/noche: cuenta todos los rounds iniciados en cada turno.
    - Fin prom./más temprano/más tarde: usa únicamente el FIN DEL ÚLTIMO
      round de cada Jumbo + Fecha operativa + Turno.

    Para el turno noche, la hora se calcula sobre una escala continua
    19:00 -> 07:00 (por ejemplo, 02:00 = 26.0) y luego se vuelve a mostrar
    como hora reloj mediante fmt_hora_decimal().
    """
    if rows.empty:
        return pd.DataFrame()

    requeridas = {
        "Jumbo",
        "Inicio_Perforacion_TS",
        "Fin_Perforacion_TS",
    }
    if not requeridas.issubset(rows.columns):
        return pd.DataFrame()

    work = rows[
        rows["Inicio_Perforacion_TS"].notna()
        & rows["Fin_Perforacion_TS"].notna()
        & rows["Jumbo"].notna()
    ].copy()

    if work.empty:
        return pd.DataFrame()

    # El turno y la fecha operativa se determinan por el INICIO del round.
    work["_turno"] = work["Inicio_Perforacion_TS"].apply(
        lambda x: zda_turno(x)[0]
    )
    work["_opDate"] = work["Inicio_Perforacion_TS"].apply(
        zda_operational_date
    )

    # Hora de fin continua respecto de la fecha operativa.
    # Esto evita errores al cruzar medianoche en el turno noche.
    work["_finHour"] = work.apply(
        lambda r: zda_operational_hour(
            r["Fin_Perforacion_TS"],
            r["_opDate"],
        ),
        axis=1,
    )

    # Último round por Jumbo + Fecha operativa + Turno.
    last_idx = work.groupby(
        ["Jumbo", "_opDate", "_turno"]
    )["Fin_Perforacion_TS"].idxmax()
    last = work.loc[last_idx].copy()

    salida = []

    for jumbo in sorted(work["Jumbo"].dropna().astype(str).unique()):
        allj = work[work["Jumbo"].astype(str) == jumbo]
        lastj = last[last["Jumbo"].astype(str) == jumbo]

        day_all = allj[allj["_turno"] == "Día"]
        night_all = allj[allj["_turno"] == "Noche"]

        day = lastj[lastj["_turno"] == "Día"]
        night = lastj[lastj["_turno"] == "Noche"]

        salida.append({
            "Jumbo": jumbo,
            "Ciclos día": len(day_all),
            "Fin prom. día": (
                fmt_hora_decimal(day["_finHour"].mean())
                if not day.empty else "-"
            ),
            "Fin más temprano día": (
                fmt_hora_decimal(day["_finHour"].min())
                if not day.empty else "-"
            ),
            "Fin más tarde día": (
                fmt_hora_decimal(day["_finHour"].max())
                if not day.empty else "-"
            ),
            "Ciclos noche": len(night_all),
            "Fin prom. noche": (
                fmt_hora_decimal(night["_finHour"].mean())
                if not night.empty else "-"
            ),
            "Fin más temprano noche": (
                fmt_hora_decimal(night["_finHour"].min())
                if not night.empty else "-"
            ),
            "Fin más tarde noche": (
                fmt_hora_decimal(night["_finHour"].max())
                if not night.empty else "-"
            ),
        })

    return pd.DataFrame(salida)


def resumen_tipos_zda(rows: pd.DataFrame) -> pd.DataFrame:
    if rows.empty:
        return pd.DataFrame()
    jumbos = sorted(rows["Jumbo"].dropna().astype(str).unique())
    salida = []
    for tipo in TIPOS_DISPARO:
        r = {"Tipo": tipo, "Total": int((rows["Tipo_Disparo"]==tipo).sum())}
        for j in jumbos:
            r[j] = int(((rows["Tipo_Disparo"]==tipo) & (rows["Jumbo"].astype(str)==j)).sum())
        if r["Total"] > 0:
            salida.append(r)
    return pd.DataFrame(salida)


def _turno_inicio_ciclo(ts):
    d = _utc_dt(ts)
    h = d.hour + d.minute / 60 + d.second / 3600
    return "Día" if 7 <= h < 19 else "Noche"


def preparar_timeline_ciclos_turno(rows: pd.DataFrame) -> pd.DataFrame:
    """
    Prepara ciclos físicos únicos para el timeline por día/turno.

    Identidad del ciclo:
        Jumbo + número de Ciclo

    Cada fila resultante representa UN round/frente perforado y contiene:
        - fecha operativa,
        - turno de inicio,
        - hora relativa de inicio,
        - hora relativa de fin,
        - duración,
        - jumbo,
        - número de ciclo.

    No se asigna 1.er/2.º/3.º ciclo. Si existen dos rounds reales del mismo
    jumbo dentro de un turno, aparecerán naturalmente como dos segmentos.
    """
    if rows.empty:
        return pd.DataFrame()

    requeridas = {
        "Inicio_Perforacion_TS",
        "Fin_Perforacion_TS",
        "Jumbo",
        "Ciclo",
    }
    if not requeridas.issubset(rows.columns):
        return pd.DataFrame()

    work = rows[
        rows["Inicio_Perforacion_TS"].notna()
        & rows["Fin_Perforacion_TS"].notna()
        & rows["Jumbo"].notna()
        & rows["Ciclo"].notna()
    ].copy()

    if work.empty:
        return pd.DataFrame()

    work["_Inicio"] = pd.to_numeric(
        work["Inicio_Perforacion_TS"],
        errors="coerce",
    )
    work["_Fin"] = pd.to_numeric(
        work["Fin_Perforacion_TS"],
        errors="coerce",
    )

    work = work[
        work["_Inicio"].notna()
        & work["_Fin"].notna()
        & (work["_Fin"] >= work["_Inicio"])
    ].copy()

    if work.empty:
        return pd.DataFrame()

    work["_Jumbo_Key"] = (
        work["Jumbo"]
        .astype(str)
        .str.strip()
    )
    work["_Ciclo_Key"] = (
        work["Ciclo"]
        .astype(str)
        .str.strip()
    )

    work["_Duracion_s"] = (
        work["_Fin"] - work["_Inicio"]
    )

    # Un mismo número de ciclo del mismo jumbo se considera el mismo round.
    # Si aparece duplicado, conservamos el registro con mayor ventana temporal.
    work = (
        work.sort_values(
            [
                "_Jumbo_Key",
                "_Ciclo_Key",
                "_Duracion_s",
            ],
            ascending=[True, True, False],
        )
        .drop_duplicates(
            subset=[
                "_Jumbo_Key",
                "_Ciclo_Key",
            ],
            keep="first",
        )
        .copy()
    )

    salida = []

    for _, r in work.iterrows():
        start_ts = float(r["_Inicio"])
        end_ts = float(r["_Fin"])

        op_date = zda_operational_date(start_ts)
        turno = _turno_inicio_ciclo(start_ts)

        if turno == "Día":
            shift_start = op_date + timedelta(hours=7)
        else:
            shift_start = op_date + timedelta(hours=19)

        shift_start_ts = shift_start.timestamp()

        x_inicio = (
            start_ts - shift_start_ts
        ) / 3600.0

        x_fin = (
            end_ts - shift_start_ts
        ) / 3600.0

        salida.append({
            "Fecha_Operativa_DT": op_date,
            "Fecha_Operativa": op_date.strftime("%d/%m/%Y"),
            "Turno": turno,
            "Dia_Turno": (
                f"{op_date.strftime('%d/%m')} · {turno}"
            ),
            "Jumbo": str(r.get("Jumbo") or "-"),
            "Ciclo": r.get("Ciclo"),
            "X_Inicio": x_inicio,
            "X_Fin": x_fin,
            "Inicio": (
                r.get("Inicio_Perforacion")
                or _utc_dt(start_ts).strftime("%d/%m/%Y %H:%M:%S")
            ),
            "Fin": (
                r.get("Fin_Perforacion")
                or _utc_dt(end_ts).strftime("%d/%m/%Y %H:%M:%S")
            ),
            "Duracion": (
                r.get("Tiempo_Perforacion_hms")
                or format_duration_hms(end_ts - start_ts)
            ),
            "Duracion_h": (
                end_ts - start_ts
            ) / 3600.0,
            "Tipo_Disparo": r.get("Tipo_Disparo") or "-",
            "Tipo_Roca": r.get("Tipo_Roca") or "SIN DATO",
            "Labor": r.get("Labor") or "-",
            "Operador": (
                r.get("Operador_ZDA")
                or r.get("Operador")
                or "-"
            ),
            "Barrenos": (
                r.get("Barrenos_ZDA")
                if pd.notna(r.get("Barrenos_ZDA"))
                else r.get("Barrenos_Realizados")
            ),
            "Sobrepasa_Turno": bool(x_fin > 12),
        })

    if not salida:
        return pd.DataFrame()

    return (
        pd.DataFrame(salida)
        .sort_values(
            [
                "Fecha_Operativa_DT",
                "Turno",
                "Jumbo",
                "X_Inicio",
            ]
        )
        .reset_index(drop=True)
    )


def _kmeans_1d(valores, k, max_iter=100):
    """
    K-Means determinístico para una sola variable.
    Devuelve etiquetas y centros ordenados de menor a mayor.
    """
    x = np.asarray(valores, dtype=float)
    x = x[np.isfinite(x)]

    if x.size < k or k < 1:
        return None, None

    # Inicialización determinística por cuantiles.
    qs = np.linspace(0, 1, k + 2)[1:-1]
    centros = np.quantile(x, qs).astype(float)

    for _ in range(max_iter):
        dist = np.abs(x[:, None] - centros[None, :])
        etiquetas = np.argmin(dist, axis=1)

        nuevos = centros.copy()
        for j in range(k):
            vals_j = x[etiquetas == j]
            if vals_j.size:
                nuevos[j] = float(vals_j.mean())
            else:
                # Reubicar un centro vacío en el punto más alejado
                # de su centro actualmente asignado.
                dist_min = np.min(dist, axis=1)
                nuevos[j] = float(x[int(np.argmax(dist_min))])

        if np.allclose(nuevos, centros, atol=1e-6):
            centros = nuevos
            break
        centros = nuevos

    # Reordenar centros y etiquetas de temprano -> tardío.
    orden = np.argsort(centros)
    mapa = {old: new for new, old in enumerate(orden)}
    centros_ordenados = centros[orden]
    etiquetas_ordenadas = np.array([mapa[int(e)] for e in etiquetas], dtype=int)

    return etiquetas_ordenadas, centros_ordenados


def _silhouette_1d(valores, etiquetas):
    """
    Silhouette promedio para clustering 1D usando distancia absoluta.
    No requiere scikit-learn.
    """
    x = np.asarray(valores, dtype=float)
    labels = np.asarray(etiquetas, dtype=int)

    if x.size < 3 or len(np.unique(labels)) < 2:
        return None

    scores = []

    for i in range(len(x)):
        same = labels == labels[i]
        same[i] = False

        # Para clusters singleton se usa 0, equivalente a una
        # observación sin soporte interno suficiente.
        if same.sum() == 0:
            scores.append(0.0)
            continue

        a = float(np.mean(np.abs(x[i] - x[same])))

        b_candidates = []
        for lab in np.unique(labels):
            if lab == labels[i]:
                continue
            other = labels == lab
            if other.sum():
                b_candidates.append(
                    float(np.mean(np.abs(x[i] - x[other])))
                )

        if not b_candidates:
            scores.append(0.0)
            continue

        b = min(b_candidates)
        denom = max(a, b)
        scores.append((b - a) / denom if denom > 0 else 0.0)

    return float(np.mean(scores)) if scores else None


def analizar_clusters_primer_inicio(
    ciclos: pd.DataFrame,
    turno: str,
    k_min: int = 2,
    k_max: int = 4,
):
    """
    Detecta clusters naturales de hora de primer inicio.

    Población:
      primer inicio por Fecha operativa + Turno + Jumbo.

    Selección de K:
      prueba K=2..4 y elige el mayor Silhouette Score.
    """
    if ciclos is None or ciclos.empty:
        return None

    requeridas = {
        "Fecha_Operativa_DT",
        "Turno",
        "Jumbo",
        "X_Inicio",
    }
    if not requeridas.issubset(ciclos.columns):
        return None

    work = ciclos[ciclos["Turno"].eq(turno)].copy()
    if work.empty:
        return None

    work["X_Inicio"] = pd.to_numeric(
        work["X_Inicio"],
        errors="coerce",
    )
    work = work[
        work["X_Inicio"].notna()
        & work["X_Inicio"].between(0, 12, inclusive="both")
    ].copy()

    if work.empty:
        return None

    # Primer arranque real de cada equipo por fecha-turno.
    primeros = (
        work.sort_values(
            [
                "Fecha_Operativa_DT",
                "Jumbo",
                "X_Inicio",
                "Ciclo",
            ]
        )
        .drop_duplicates(
            subset=[
                "Fecha_Operativa_DT",
                "Turno",
                "Jumbo",
            ],
            keep="first",
        )
        .copy()
    )

    valores = primeros["X_Inicio"].to_numpy(dtype=float)
    n = len(valores)

    # Con muy pocos puntos la segmentación no es estable.
    if n < 4 or np.unique(np.round(valores, 4)).size < 2:
        return {
            "ok": False,
            "motivo": "Se requieren al menos 4 primeros inicios con variación.",
            "n": n,
            "primeros": primeros,
        }

    mejor = None
    k_sup = min(k_max, n - 1, np.unique(np.round(valores, 4)).size)

    for k in range(k_min, k_sup + 1):
        labels, centros = _kmeans_1d(valores, k)
        if labels is None:
            continue

        # Evitar soluciones con clusters vacíos.
        counts = np.bincount(labels, minlength=k)
        if np.any(counts == 0):
            continue

        sil = _silhouette_1d(valores, labels)
        if sil is None:
            continue

        candidato = {
            "k": k,
            "labels": labels,
            "centros": centros,
            "silhouette": sil,
            "counts": counts,
        }

        if mejor is None or sil > mejor["silhouette"]:
            mejor = candidato

    if mejor is None:
        return {
            "ok": False,
            "motivo": "No fue posible obtener una segmentación estable.",
            "n": n,
            "primeros": primeros,
        }

    k = mejor["k"]
    centros = np.asarray(mejor["centros"], dtype=float)

    # Límites de zona = punto medio entre centroides consecutivos.
    limites = [0.0]
    for a, b in zip(centros[:-1], centros[1:]):
        limites.append(float((a + b) / 2.0))
    limites.append(12.0)

    if k == 2:
        nombres = ["Temprano", "Tardío"]
    elif k == 3:
        nombres = ["Temprano", "Intermedio", "Tardío"]
    elif k == 4:
        nombres = [
            "Muy temprano",
            "Temprano",
            "Tardío",
            "Muy tardío",
        ]
    else:
        nombres = [f"Cluster {i+1}" for i in range(k)]

    resumen = []
    for i in range(k):
        n_i = int(mejor["counts"][i])
        resumen.append({
            "Cluster": nombres[i],
            "Centro_h": float(centros[i]),
            "Desde_h": float(limites[i]),
            "Hasta_h": float(limites[i + 1]),
            "N": n_i,
            "Pct": n_i / n * 100.0 if n else 0.0,
        })

    primeros = primeros.copy()
    primeros["Cluster_ID"] = mejor["labels"]
    primeros["Cluster"] = [
        nombres[int(i)]
        for i in mejor["labels"]
    ]

    return {
        "ok": True,
        "k": k,
        "silhouette": float(mejor["silhouette"]),
        "centros": centros,
        "limites": limites,
        "resumen": pd.DataFrame(resumen),
        "primeros": primeros,
        "n": n,
    }


def render_resumen_clusters_primer_inicio(
    cluster_info,
    turno: str,
):
    """
    Muestra una tabla compacta y una lectura automática del clustering
    del primer inicio para un turno.
    """
    if not cluster_info:
        return

    if not cluster_info.get("ok"):
        st.info(
            cluster_info.get(
                "motivo",
                f"No hay datos suficientes para resumir clusters del Turno {turno}.",
            )
        )
        return

    resumen = cluster_info.get("resumen")
    if resumen is None or resumen.empty:
        return

    st.markdown(
        f"##### Resumen de clusters · Turno {turno}"
    )
    st.caption(
        f"K={cluster_info.get('k', '-')} · "
        f"Silhouette={cluster_info.get('silhouette', 0):.2f}"
    )

    tabla = resumen.copy()
    tabla["Centro"] = tabla["Centro_h"].apply(
        lambda v: _hora_relativa_turno_a_texto(v, turno)
    )
    tabla["Ciclos"] = pd.to_numeric(
        tabla["N"],
        errors="coerce",
    ).fillna(0).astype(int)
    tabla["Participación"] = tabla["Pct"].apply(
        lambda v: f"{float(v):.0f}%"
    )

    tabla = tabla[
        [
            "Cluster",
            "Centro",
            "Ciclos",
            "Participación",
        ]
    ].rename(
        columns={
            "Cluster": "Grupo",
        }
    )

    st.dataframe(
        tabla,
        width="stretch",
        hide_index=True,
        height=min(
            60 + 36 * len(tabla),
            235,
        ),
    )

    # Lectura compacta del patrón dominante y del patrón más tardío.
    dominante = resumen.sort_values(
        ["Pct", "Centro_h"],
        ascending=[False, True],
    ).iloc[0]

    tardio = resumen.sort_values(
        "Centro_h",
        ascending=False,
    ).iloc[0]

    centro_dom = _hora_relativa_turno_a_texto(
        dominante["Centro_h"],
        turno,
    )
    centro_tardio = _hora_relativa_turno_a_texto(
        tardio["Centro_h"],
        turno,
    )

    pct_dom = float(dominante["Pct"])
    pct_tardio = float(tardio["Pct"])

    if str(dominante["Cluster"]) != str(tardio["Cluster"]):
        lectura = (
            f"**El {pct_dom:.0f}% de los primeros martillos del Turno {turno} "
            f"se concentra alrededor de las {centro_dom} "
            f"({dominante['Cluster']}), mientras que el patrón más tardío "
            f"se concentra alrededor de las {centro_tardio} y representa "
            f"el {pct_tardio:.0f}% de los inicios.**"
        )
    else:
        lectura = (
            f"**El patrón con mayor participación del Turno {turno} es "
            f"{dominante['Cluster']}: concentra el {pct_dom:.0f}% de los primeros "
            f"martillos alrededor de las {centro_dom}.**"
        )

    st.markdown(
        f"""
        <div style="
            border-left: 4px solid #d1d5db;
            padding: 0.55rem 0.85rem;
            margin: 0.35rem 0 1.10rem 0;
            color: #1f2937;
            line-height: 1.45;
        ">
            {lectura}
        </div>
        """,
        unsafe_allow_html=True,
    )


def grafico_timeline_ciclos_turno(
    ciclos: pd.DataFrame,
    turno: str,
    solo_puntos_inicio: bool = False,
    solo_primer_inicio: bool = False,
    mostrar_clusters: bool = False,
):
    """
    Timeline horizontal por fecha operativa.

    X:
        horas transcurridas desde el inicio del turno.

    Y:
        fecha operativa + turno.

    Dentro de cada fila:
        JUMB001 se dibuja ligeramente arriba,
        JUMB002 ligeramente abajo.

    Cada round real es un segmento completo Inicio -> Fin.
    Si un jumbo ejecuta dos rounds, aparecen dos segmentos consecutivos
    sobre la misma pista de ese jumbo.

    Por defecto se muestran barras horizontales Inicio -> Fin.
    Si solo_puntos_inicio=True, las barras se ocultan completamente y
    se muestran únicamente los puntos de inicio para analizar patrones horarios.
    Si además solo_primer_inicio=True, se muestra únicamente el primer
    inicio de cada equipo por cada combinación fecha operativa + turno.
    """
    if ciclos.empty:
        return None

    g = ciclos[
        ciclos["Turno"].eq(turno)
    ].copy()

    if g.empty:
        return None

    cluster_info = None
    if mostrar_clusters and solo_puntos_inicio:
        cluster_info = analizar_clusters_primer_inicio(
            ciclos,
            turno,
        )

    # Vista de patrones de arranque:
    # al mostrar clusters se usa automáticamente el primer inicio
    # de cada Jumbo + fecha operativa + turno, porque esa es la
    # población sobre la que se calculan los centroides.
    if solo_puntos_inicio and (solo_primer_inicio or mostrar_clusters):
        g = (
            g.sort_values(
                [
                    "Fecha_Operativa_DT",
                    "Jumbo",
                    "X_Inicio",
                    "Ciclo",
                ]
            )
            .drop_duplicates(
                subset=[
                    "Fecha_Operativa_DT",
                    "Turno",
                    "Jumbo",
                ],
                keep="first",
            )
            .copy()
        )

    fechas = sorted(
        g["Fecha_Operativa_DT"].dropna().unique()
    )

    if not fechas:
        return None

    y_map = {
        pd.Timestamp(fecha): i
        for i, fecha in enumerate(fechas)
    }

    offsets = {
        "JUMB001": -0.16,
        "JUMB002": 0.16,
    }

    estilos = {
        "JUMB001": {
            "line_color": "#64748b",
            "dash": "solid",
            "marker_color": "#64748b",
            "legend_name": "JUMB001",
        },
        "JUMB002": {
            "line_color": "#111827",
            "dash": "solid",
            "marker_color": "#111827",
            "legend_name": "JUMB002",
        },
    }

    fig = go.Figure()
    legend_done = set()

    for _, r in g.iterrows():
        jumbo = str(r["Jumbo"])
        estilo = estilos.get(
            jumbo,
            estilos["JUMB001"],
        )

        fecha_ts = pd.Timestamp(
            r["Fecha_Operativa_DT"]
        )
        y_base = y_map[fecha_ts]
        y = y_base + offsets.get(jumbo, 0)

        x1 = float(r["X_Inicio"])
        x2 = float(r["X_Fin"])

        custom = [[
            jumbo,
            r.get("Ciclo"),
            r.get("Inicio"),
            r.get("Fin"),
            r.get("Duracion"),
            r.get("Tipo_Disparo"),
            r.get("Tipo_Roca"),
            r.get("Labor"),
            r.get("Barrenos"),
            r.get("Fecha_Operativa"),
            "Sí" if r.get("Sobrepasa_Turno") else "No",
            r.get("Operador") or "-",
        ]] * 2

        if solo_puntos_inicio:
            # Modo análisis de inicios:
            # ocultar completamente la barra y mostrar solo el inicio.
            if jumbo == "JUMB001":
                # En símbolos "open", Plotly usa marker.color como
                # color principal del contorno. No debe ser blanco.
                symbol = "square-open"
                marker_color = "#64748b"
                marker_line_color = "#64748b"
            else:
                symbol = "square"
                marker_color = "#111827"
                marker_line_color = "#111827"

            fig.add_trace(
                go.Scatter(
                    x=[x1],
                    y=[y],
                    mode="markers",
                    name=estilo["legend_name"],
                    legendgroup=jumbo,
                    showlegend=jumbo not in legend_done,
                    marker=dict(
                        size=11,
                        symbol=symbol,
                        color=marker_color,
                        line=dict(
                            color=marker_line_color,
                            width=1.8,
                        ),
                    ),
                    customdata=[custom[0]],
                    hovertemplate=(
                        "<b>%{customdata[0]}</b>"
                        "<br>Ciclo / round: %{customdata[1]}"
                        "<br>Fecha operativa: %{customdata[9]}"
                        "<br>Inicio: %{customdata[2]}"
                        "<br>Fin: %{customdata[3]}"
                        "<br>Duración: %{customdata[4]}"
                        "<br>Tipo: %{customdata[5]}"
                        "<br>Tipo de roca: %{customdata[6]}"
                        "<br>Labor: %{customdata[7]}"
                        "<br>Operador: %{customdata[11]}"
                        "<br>Barrenos: %{customdata[8]}"
                        "<br>Sobrepasa turno: %{customdata[10]}"
                        "<extra></extra>"
                    ),
                )
            )
        else:
            # Modo timeline completo, sincronizado con la vista de puntos:
            # JUMB001 = barra hueca con contorno cerrado.
            # JUMB002 = barra negra sólida.
            hover_barra = (
                "<b>%{customdata[0]}</b>"
                "<br>Ciclo / round: %{customdata[1]}"
                "<br>Fecha operativa: %{customdata[9]}"
                "<br>Inicio: %{customdata[2]}"
                "<br>Fin: %{customdata[3]}"
                "<br>Duración: %{customdata[4]}"
                "<br>Tipo: %{customdata[5]}"
                "<br>Tipo de roca: %{customdata[6]}"
                "<br>Labor: %{customdata[7]}"
                "<br>Operador: %{customdata[11]}"
                "<br>Barrenos: %{customdata[8]}"
                "<br>Sobrepasa turno: %{customdata[10]}"
                "<extra></extra>"
            )

            duracion = max(float(x2 - x1), 0.02)
            ancho_barra = 0.15
            show_legend_actual = jumbo not in legend_done

            if jumbo == "JUMB001":
                fig.add_trace(
                    go.Bar(
                        x=[duracion],
                        y=[y],
                        base=[x1],
                        orientation="h",
                        width=ancho_barra,
                        name=estilo["legend_name"],
                        legendgroup=jumbo,
                        showlegend=show_legend_actual,
                        marker=dict(
                            color="rgba(255,255,255,0)",
                            line=dict(
                                color="#64748b",
                                width=1.8,
                            ),
                        ),
                        customdata=[custom[0]],
                        hovertemplate=hover_barra,
                    )
                )
            else:
                fig.add_trace(
                    go.Bar(
                        x=[duracion],
                        y=[y],
                        base=[x1],
                        orientation="h",
                        width=ancho_barra,
                        name=estilo["legend_name"],
                        legendgroup=jumbo,
                        showlegend=show_legend_actual,
                        marker=dict(
                            color="#111827",
                            line=dict(
                                color="#111827",
                                width=0.8,
                            ),
                        ),
                        customdata=[custom[0]],
                        hovertemplate=hover_barra,
                    )
                )

        legend_done.add(jumbo)

    tickvals = list(range(0, 13))

    if turno == "Día":
        ticktext = [f"{(7 + h) % 24:02d}:00" for h in tickvals]
        titulo = "Turno Día · 07:00–19:00"
    else:
        ticktext = [f"{(19 + h) % 24:02d}:00" for h in tickvals]
        titulo = "Turno Noche · 19:00–07:00"

    y_tickvals = [y_map[pd.Timestamp(fecha)] for fecha in fechas]
    y_ticktext = [f"{pd.Timestamp(fecha).strftime('%d/%m')} · {turno}" for fecha in fechas]

    height = max(420, min(980, 150 + len(fechas) * 30))

    fig.update_layout(
        **base_layout(
            height,
            title=dict(text=f"<b>{titulo}</b>", x=0.01, xanchor="left"),
            margin=dict(l=125, r=30, t=105 if (mostrar_clusters and solo_puntos_inicio) else 58, b=75),
            xaxis=dict(
                title="Hora del turno",
                range=[0, 12],
                tickmode="array",
                tickvals=tickvals,
                ticktext=ticktext,
                gridcolor="#e5e7eb",
                zeroline=False,
                fixedrange=True,
            ),
            yaxis=dict(
                title="Fecha operativa · turno",
                tickmode="array",
                tickvals=y_tickvals,
                ticktext=y_ticktext,
                gridcolor="#eef2f7",
                zeroline=False,
                fixedrange=True,
                # Margen extra arriba para que la primera barra no quede recortada.
                range=[len(fechas) - 0.5, -0.85],
            ),
            legend=dict(orientation="h", y=-0.12, x=0),
            hovermode="closest",
        )
    )

    # ------------------------------------------------------
    # Capa analítica de clusters de hora de primer inicio
    # ------------------------------------------------------
    if cluster_info and cluster_info.get("ok"):
        resumen_cluster = cluster_info["resumen"]

        # Colores muy suaves para no competir con la codificación
        # JUMB001/JUMB002 de los puntos.
        zonas_rgba = [
            "rgba(59,130,246,0.055)",   # azul suave
            "rgba(16,185,129,0.050)",   # verde suave
            "rgba(245,158,11,0.050)",   # naranja suave
            "rgba(239,68,68,0.045)",    # rojo suave
        ]
        lineas_cluster = [
            "#3B82F6",
            "#10B981",
            "#F59E0B",
            "#EF4444",
        ]

        for i, row in resumen_cluster.iterrows():
            color_fill = zonas_rgba[i % len(zonas_rgba)]
            color_line = lineas_cluster[i % len(lineas_cluster)]

            fig.add_vrect(
                x0=float(row["Desde_h"]),
                x1=float(row["Hasta_h"]),
                fillcolor=color_fill,
                line_width=0,
                layer="below",
            )

            centro = float(row["Centro_h"])
            fig.add_vline(
                x=centro,
                line_width=1.6,
                line_dash="dash",
                line_color=color_line,
                opacity=0.82,
            )

            fig.add_annotation(
                x=centro,
                y=1.035,
                xref="x",
                yref="paper",
                text=(
                    f"<b>{row['Cluster']}</b><br>"
                    f"Centro {_hora_relativa_turno_a_texto(centro, turno)}"
                    f"<br>{int(row['N'])} inicios · {row['Pct']:.0f}%"
                ),
                showarrow=False,
                xanchor="center",
                yanchor="bottom",
                align="center",
                font=dict(
                    size=10,
                    color=color_line,
                ),
                bgcolor="rgba(255,255,255,0.82)",
                bordercolor="rgba(0,0,0,0)",
                borderpad=2,
            )

        fig.add_annotation(
            x=0.995,
            y=1.105,
            xref="paper",
            yref="paper",
            text=(
                f"K={cluster_info['k']} · "
                f"Silhouette={cluster_info['silhouette']:.2f}"
            ),
            showarrow=False,
            xanchor="right",
            yanchor="top",
            font=dict(size=10, color="#64748b"),
            bgcolor="rgba(255,255,255,0.80)",
        )

    elif mostrar_clusters and solo_puntos_inicio and cluster_info:
        fig.add_annotation(
            x=0.995,
            y=1.075,
            xref="paper",
            yref="paper",
            text=cluster_info.get(
                "motivo",
                "No hay datos suficientes para clustering.",
            ),
            showarrow=False,
            xanchor="right",
            yanchor="top",
            font=dict(size=10, color="#64748b"),
        )

    fig.add_vline(
        x=12,
        line_width=1.6,
        line_dash="dash",
        line_color="#b45309",
        annotation_text="Fin turno",
        annotation_position="top",
    )

    for x_ref in [2, 4, 6, 8, 10]:
        fig.add_vline(
            x=x_ref,
            line_width=0.9,
            line_dash="dot",
            line_color="#cbd5e1",
        )

    return fig


def preparar_primeros_inicios_distribucion(ciclos: pd.DataFrame) -> pd.DataFrame:
    """
    Conserva únicamente el primer inicio de cada Jumbo + fecha operativa + turno.
    X_Inicio está en horas relativas desde el inicio del turno (0..12), por lo
    que el turno noche se maneja correctamente aunque cruce medianoche.
    """
    if ciclos.empty:
        return pd.DataFrame()

    requeridas = {
        "Fecha_Operativa_DT",
        "Turno",
        "Jumbo",
        "X_Inicio",
    }
    if not requeridas.issubset(ciclos.columns):
        return pd.DataFrame()

    work = ciclos.copy()
    work["X_Inicio"] = pd.to_numeric(work["X_Inicio"], errors="coerce")
    work = work[
        work["Fecha_Operativa_DT"].notna()
        & work["Turno"].notna()
        & work["Jumbo"].notna()
        & work["X_Inicio"].notna()
        & work["X_Inicio"].between(0, 12, inclusive="both")
    ].copy()

    if work.empty:
        return pd.DataFrame()

    primeros = (
        work.sort_values(
            [
                "Fecha_Operativa_DT",
                "Turno",
                "Jumbo",
                "X_Inicio",
                "Ciclo",
            ]
        )
        .drop_duplicates(
            subset=[
                "Fecha_Operativa_DT",
                "Turno",
                "Jumbo",
            ],
            keep="first",
        )
        .copy()
    )

    return primeros.reset_index(drop=True)


def _densidad_gaussiana_horas(valores, grid):
    """
    KDE gaussiana simple sin scipy.
    El ancho de banda usa una adaptación de Silverman y se limita para evitar
    curvas demasiado irregulares cuando existen pocos ciclos.
    """
    vals = np.asarray(valores, dtype=float)
    vals = vals[np.isfinite(vals)]
    grid = np.asarray(grid, dtype=float)

    if vals.size == 0:
        return np.zeros_like(grid, dtype=float)

    if vals.size == 1:
        bandwidth = 0.40
    else:
        std = float(np.std(vals, ddof=1))
        q75, q25 = np.percentile(vals, [75, 25])
        iqr_sigma = float((q75 - q25) / 1.349) if q75 > q25 else std
        escala = min(std, iqr_sigma) if std > 0 and iqr_sigma > 0 else max(std, iqr_sigma)
        if not np.isfinite(escala) or escala <= 0:
            escala = 0.50
        bandwidth = 0.9 * escala * (vals.size ** (-1 / 5))
        bandwidth = float(np.clip(bandwidth, 0.25, 0.85))

    z = (grid[:, None] - vals[None, :]) / bandwidth
    densidad = np.exp(-0.5 * z * z).sum(axis=1)
    densidad /= vals.size * bandwidth * np.sqrt(2 * np.pi)
    return densidad


def grafico_distribucion_primeros_inicios(
    primeros: pd.DataFrame,
    turno: str,
):
    """
    Histograma de primeros inicios + curva de densidad para un turno.

    Histograma: frecuencia en intervalos de 30 min.
    Curva: KDE gaussiana en eje Y secundario.
    Incluye Promedio, Mediana y Pico aproximado de la densidad.
    """
    if primeros.empty:
        return None

    g = primeros[primeros["Turno"].eq(turno)].copy()
    if g.empty:
        return None

    valores = pd.to_numeric(g["X_Inicio"], errors="coerce").dropna().to_numpy(dtype=float)
    valores = valores[(valores >= 0) & (valores <= 12)]
    if len(valores) == 0:
        return None

    # Intervalos de 30 minutos.
    edges = np.arange(0, 12.0001 + 0.5, 0.5)
    counts, _ = np.histogram(valores, bins=edges)
    centers = (edges[:-1] + edges[1:]) / 2

    # Curva de densidad suave.
    grid = np.linspace(0, 12, 241)
    densidad = _densidad_gaussiana_horas(valores, grid)

    promedio = float(np.mean(valores))
    mediana = float(np.median(valores))
    pico = float(grid[int(np.argmax(densidad))]) if len(densidad) else mediana

    fig = make_subplots(specs=[[{"secondary_y": True}]])

    fig.add_trace(
        go.Bar(
            x=centers,
            y=counts,
            width=0.46,
            name="Frecuencia · 30 min",
            marker=dict(
                color="rgba(100,116,139,0.42)",
                line=dict(color="#64748b", width=0.8),
            ),
            customdata=[
                [
                    _hora_relativa_turno_a_texto(max(0, c - 0.25), turno),
                    _hora_relativa_turno_a_texto(min(12, c + 0.25), turno),
                    int(n),
                ]
                for c, n in zip(centers, counts)
            ],
            hovertemplate=(
                "<b>Primeros inicios</b>"
                "<br>Rango: %{customdata[0]}–%{customdata[1]}"
                "<br>Cantidad: %{customdata[2]}"
                "<extra></extra>"
            ),
        ),
        secondary_y=False,
    )

    fig.add_trace(
        go.Scatter(
            x=grid,
            y=densidad,
            mode="lines",
            name="Densidad suavizada",
            line=dict(color="#0f766e", width=3),
            hovertemplate=(
                "<b>Densidad</b>"
                "<br>Hora: %{customdata}"
                "<br>Densidad: %{y:.3f}"
                "<extra></extra>"
            ),
            customdata=[_hora_relativa_turno_a_texto(v, turno) for v in grid],
        ),
        secondary_y=True,
    )

    # Promedio y mediana.
    fig.add_vline(
        x=promedio,
        line_width=1.6,
        line_dash="dash",
        line_color="#111827",
        annotation_text=f"Promedio {_hora_relativa_turno_a_texto(promedio, turno)}",
        annotation_position="top right",
    )
    fig.add_vline(
        x=mediana,
        line_width=1.8,
        line_dash="dot",
        line_color="#2563eb",
        annotation_text=f"Mediana {_hora_relativa_turno_a_texto(mediana, turno)}",
        annotation_position="top left",
    )

    # Pico aproximado de la densidad.
    fig.add_vline(
        x=pico,
        line_width=1.1,
        line_dash="dot",
        line_color="#0f766e",
        opacity=0.65,
    )

    tickvals = list(range(0, 13))
    if turno == "Día":
        ticktext = [f"{(7 + h) % 24:02d}:00" for h in tickvals]
        titulo = "Turno Día · Distribución de primeros inicios"
    else:
        ticktext = [f"{(19 + h) % 24:02d}:00" for h in tickvals]
        titulo = "Turno Noche · Distribución de primeros inicios"

    equipos = int(g["Jumbo"].nunique())
    observaciones = int(len(g))

    fig.update_layout(
        **base_layout(
            480,
            title=dict(
                text=f"<b>{titulo}</b>",
                x=0.01,
                xanchor="left",
            ),
            margin=dict(l=70, r=70, t=80, b=70),
            xaxis=dict(
                title="Hora del primer inicio",
                range=[0, 12],
                tickmode="array",
                tickvals=tickvals,
                ticktext=ticktext,
                fixedrange=True,
                gridcolor="#e5e7eb",
                zeroline=False,
            ),
            legend=dict(
                orientation="h",
                y=-0.18,
                x=0,
            ),
            hovermode="closest",
        )
    )

    fig.update_yaxes(
        title_text="Cantidad de primeros inicios",
        rangemode="tozero",
        fixedrange=True,
        secondary_y=False,
    )
    fig.update_yaxes(
        title_text="Densidad",
        rangemode="tozero",
        fixedrange=True,
        showgrid=False,
        secondary_y=True,
    )

    # Nota compacta dentro del gráfico.
    fig.add_annotation(
        x=0.995,
        y=1.11,
        xref="paper",
        yref="paper",
        xanchor="right",
        yanchor="top",
        showarrow=False,
        align="right",
        text=(
            f"Pico aprox.: <b>{_hora_relativa_turno_a_texto(pico, turno)}</b> · "
            f"{observaciones} inicios · {equipos} equipo(s)"
        ),
        font=dict(size=11, color="#475569"),
    )

    return fig


def preparar_tendencia_inicio_diario(ciclos: pd.DataFrame) -> pd.DataFrame:
    """
    Construye un valor diario consolidado de primer inicio por turno.

    1) Para cada Fecha operativa + Turno + Jumbo conserva únicamente
       el primer inicio del equipo.
    2) Para cada Fecha operativa + Turno calcula:
       - Promedio de los primeros inicios de los equipos disponibles.
       - Mediana de los primeros inicios de los equipos disponibles.

    X_Inicio está expresado como horas transcurridas desde el inicio
    del turno, por lo que el cálculo del turno noche es continuo y no
    se distorsiona al cruzar medianoche.
    """
    if ciclos.empty:
        return pd.DataFrame()

    requeridas = {
        "Fecha_Operativa_DT",
        "Turno",
        "Jumbo",
        "X_Inicio",
    }
    if not requeridas.issubset(ciclos.columns):
        return pd.DataFrame()

    work = ciclos.copy()
    work["X_Inicio"] = pd.to_numeric(
        work["X_Inicio"],
        errors="coerce",
    )
    work = work[
        work["Fecha_Operativa_DT"].notna()
        & work["Turno"].notna()
        & work["Jumbo"].notna()
        & work["X_Inicio"].notna()
    ].copy()

    if work.empty:
        return pd.DataFrame()

    # Primer inicio real de cada equipo en cada fecha-turno.
    primeros = (
        work.sort_values(
            [
                "Fecha_Operativa_DT",
                "Turno",
                "Jumbo",
                "X_Inicio",
                "Ciclo",
            ]
        )
        .drop_duplicates(
            subset=[
                "Fecha_Operativa_DT",
                "Turno",
                "Jumbo",
            ],
            keep="first",
        )
        .copy()
    )

    resumen = (
        primeros.groupby(
            ["Fecha_Operativa_DT", "Turno"],
            as_index=False,
        )
        .agg(
            Promedio_h=("X_Inicio", "mean"),
            Mediana_h=("X_Inicio", "median"),
            Equipos=("Jumbo", "nunique"),
        )
        .sort_values(["Turno", "Fecha_Operativa_DT"])
        .reset_index(drop=True)
    )

    return resumen


def _hora_relativa_turno_a_texto(valor_h, turno: str) -> str:
    """Convierte hora relativa del turno a HH:MM."""
    if valor_h is None or pd.isna(valor_h):
        return "-"

    minutos = int(round(float(valor_h) * 60))
    inicio_h = 7 if turno == "Día" else 19
    total_min = (inicio_h * 60 + minutos) % (24 * 60)
    hh = total_min // 60
    mm = total_min % 60
    return f"{hh:02d}:{mm:02d}"


def grafico_tendencia_inicio_diario(
    resumen: pd.DataFrame,
    turno: str,
    mostrar_promedio: bool = True,
    mostrar_mediana: bool = True,
):
    """
    Curva diaria de la hora consolidada del primer inicio.

    Eje X: fecha operativa.
    Eje Y: hora dentro del turno (0 a 12 h desde el inicio).
    """
    if resumen.empty:
        return None

    g = resumen[
        resumen["Turno"].eq(turno)
    ].copy()

    if g.empty or (not mostrar_promedio and not mostrar_mediana):
        return None

    g = g.sort_values("Fecha_Operativa_DT")

    fig = go.Figure()

    fechas_txt = [
        pd.Timestamp(x).strftime("%d/%m/%Y")
        for x in g["Fecha_Operativa_DT"]
    ]

    if mostrar_promedio:
        horas_prom = [
            _hora_relativa_turno_a_texto(v, turno)
            for v in g["Promedio_h"]
        ]
        custom_prom = list(
            zip(
                fechas_txt,
                horas_prom,
                g["Equipos"].astype(int).tolist(),
            )
        )

        fig.add_trace(
            go.Scatter(
                x=g["Fecha_Operativa_DT"],
                y=g["Promedio_h"],
                mode="lines+markers",
                name="Promedio",
                line=dict(
                    color="#111827",
                    width=2.4,
                    dash="solid",
                ),
                marker=dict(
                    size=7,
                    symbol="circle",
                    color="#111827",
                ),
                customdata=custom_prom,
                hovertemplate=(
                    "<b>Promedio</b>"
                    "<br>Fecha: %{customdata[0]}"
                    "<br>Hora: %{customdata[1]}"
                    "<br>Equipos considerados: %{customdata[2]}"
                    "<extra></extra>"
                ),
            )
        )

    if mostrar_mediana:
        horas_med = [
            _hora_relativa_turno_a_texto(v, turno)
            for v in g["Mediana_h"]
        ]
        custom_med = list(
            zip(
                fechas_txt,
                horas_med,
                g["Equipos"].astype(int).tolist(),
            )
        )

        # La mediana se dibuja después del promedio y con línea punteada.
        # Con 2 equipos ambos valores coinciden; el patrón punteado permite
        # reconocer que ambas series están superpuestas.
        fig.add_trace(
            go.Scatter(
                x=g["Fecha_Operativa_DT"],
                y=g["Mediana_h"],
                mode="lines+markers",
                name="Mediana",
                line=dict(
                    color="#64748b",
                    width=2.4,
                    dash="dash",
                ),
                marker=dict(
                    size=8,
                    symbol="square-open",
                    color="#64748b",
                    line=dict(
                        color="#64748b",
                        width=1.5,
                    ),
                ),
                customdata=custom_med,
                hovertemplate=(
                    "<b>Mediana</b>"
                    "<br>Fecha: %{customdata[0]}"
                    "<br>Hora: %{customdata[1]}"
                    "<br>Equipos considerados: %{customdata[2]}"
                    "<extra></extra>"
                ),
            )
        )

    tickvals_y = list(range(0, 13))
    if turno == "Día":
        ticktext_y = [
            f"{(7 + h) % 24:02d}:00"
            for h in tickvals_y
        ]
        titulo = "Turno Día · Tendencia del primer inicio consolidado"
    else:
        ticktext_y = [
            f"{(19 + h) % 24:02d}:00"
            for h in tickvals_y
        ]
        titulo = "Turno Noche · Tendencia del primer inicio consolidado"

    fechas = g["Fecha_Operativa_DT"].tolist()
    tickvals_x = fechas[::2] if len(fechas) > 12 else fechas
    ticktext_x = [
        pd.Timestamp(x).strftime("%d/%m")
        for x in tickvals_x
    ]

    fig.update_layout(
        **base_layout(
            460,
            title=dict(
                text=f"<b>{titulo}</b>",
                x=0.01,
                xanchor="left",
            ),
            margin=dict(l=90, r=30, t=60, b=80),
            xaxis=dict(
                title="Fecha operativa",
                tickmode="array",
                tickvals=tickvals_x,
                ticktext=ticktext_x,
                tickangle=-35,
                gridcolor="#eef2f7",
                fixedrange=True,
            ),
            yaxis=dict(
                title="Hora de primer inicio",
                range=[0, 12],
                tickmode="array",
                tickvals=tickvals_y,
                ticktext=ticktext_y,
                gridcolor="#e5e7eb",
                zeroline=False,
                fixedrange=True,
            ),
            legend=dict(
                orientation="h",
                y=1.09,
                x=0.70,
                xanchor="left",
            ),
            hovermode="x unified",
        )
    )

    # Guías cada 2 horas para mantener consistencia con el timeline.
    for y_ref in [2, 4, 6, 8, 10]:
        fig.add_hline(
            y=y_ref,
            line_width=0.8,
            line_dash="dot",
            line_color="#cbd5e1",
            layer="below",
        )

    return fig


def preparar_tendencia_ultimo_fin_diario(ciclos: pd.DataFrame) -> pd.DataFrame:
    """
    Construye un valor diario consolidado de la hora de término del último
    ciclo/round por turno.

    1) Para cada Fecha operativa + Turno + Jumbo conserva el ciclo cuyo
       X_Fin sea más tardío dentro de ese turno. Si un equipo realizó dos
       o más rounds, se toma el Fin del último round ejecutado.
    2) Para cada Fecha operativa + Turno calcula:
       - Promedio de las últimas horas de término de los equipos disponibles.
       - Mediana de las últimas horas de término de los equipos disponibles.

    X_Fin está expresado como horas transcurridas desde el inicio del turno,
    por lo que el cálculo del turno noche permanece continuo al cruzar
    medianoche.
    """
    if ciclos.empty:
        return pd.DataFrame()

    requeridas = {
        "Fecha_Operativa_DT",
        "Turno",
        "Jumbo",
        "X_Fin",
    }
    if not requeridas.issubset(ciclos.columns):
        return pd.DataFrame()

    work = ciclos.copy()
    work["X_Fin"] = pd.to_numeric(
        work["X_Fin"],
        errors="coerce",
    )
    work = work[
        work["Fecha_Operativa_DT"].notna()
        & work["Turno"].notna()
        & work["Jumbo"].notna()
        & work["X_Fin"].notna()
    ].copy()

    if work.empty:
        return pd.DataFrame()

    # Último término real de cada equipo en cada fecha-turno.
    ultimos = (
        work.sort_values(
            [
                "Fecha_Operativa_DT",
                "Turno",
                "Jumbo",
                "X_Fin",
                "Ciclo",
            ],
            ascending=[True, True, True, False, False],
        )
        .drop_duplicates(
            subset=[
                "Fecha_Operativa_DT",
                "Turno",
                "Jumbo",
            ],
            keep="first",
        )
        .copy()
    )

    resumen = (
        ultimos.groupby(
            ["Fecha_Operativa_DT", "Turno"],
            as_index=False,
        )
        .agg(
            Promedio_h=("X_Fin", "mean"),
            Mediana_h=("X_Fin", "median"),
            Equipos=("Jumbo", "nunique"),
        )
        .sort_values(["Turno", "Fecha_Operativa_DT"])
        .reset_index(drop=True)
    )

    return resumen


def grafico_tendencia_ultimo_fin_diario(
    resumen: pd.DataFrame,
    turno: str,
    mostrar_promedio: bool = True,
    mostrar_mediana: bool = True,
):
    """
    Curva diaria de la hora consolidada de término del último ciclo.

    Eje X: fecha operativa.
    Eje Y: hora dentro del turno medida desde el inicio del turno.
    """
    if resumen.empty:
        return None

    g = resumen[
        resumen["Turno"].eq(turno)
    ].copy()

    if g.empty or (not mostrar_promedio and not mostrar_mediana):
        return None

    g = g.sort_values("Fecha_Operativa_DT")

    fig = go.Figure()

    fechas_txt = [
        pd.Timestamp(x).strftime("%d/%m/%Y")
        for x in g["Fecha_Operativa_DT"]
    ]

    if mostrar_promedio:
        horas_prom = [
            _hora_relativa_turno_a_texto(v, turno)
            for v in g["Promedio_h"]
        ]
        custom_prom = list(
            zip(
                fechas_txt,
                horas_prom,
                g["Equipos"].astype(int).tolist(),
            )
        )

        fig.add_trace(
            go.Scatter(
                x=g["Fecha_Operativa_DT"],
                y=g["Promedio_h"],
                mode="lines+markers",
                name="Promedio",
                line=dict(
                    color="#111827",
                    width=2.4,
                    dash="solid",
                ),
                marker=dict(
                    size=7,
                    symbol="circle",
                    color="#111827",
                ),
                customdata=custom_prom,
                hovertemplate=(
                    "<b>Promedio</b>"
                    "<br>Fecha: %{customdata[0]}"
                    "<br>Último término: %{customdata[1]}"
                    "<br>Equipos considerados: %{customdata[2]}"
                    "<extra></extra>"
                ),
            )
        )

    if mostrar_mediana:
        horas_med = [
            _hora_relativa_turno_a_texto(v, turno)
            for v in g["Mediana_h"]
        ]
        custom_med = list(
            zip(
                fechas_txt,
                horas_med,
                g["Equipos"].astype(int).tolist(),
            )
        )

        fig.add_trace(
            go.Scatter(
                x=g["Fecha_Operativa_DT"],
                y=g["Mediana_h"],
                mode="lines+markers",
                name="Mediana",
                line=dict(
                    color="#64748b",
                    width=2.4,
                    dash="dash",
                ),
                marker=dict(
                    size=8,
                    symbol="square-open",
                    color="#64748b",
                    line=dict(
                        color="#64748b",
                        width=1.5,
                    ),
                ),
                customdata=custom_med,
                hovertemplate=(
                    "<b>Mediana</b>"
                    "<br>Fecha: %{customdata[0]}"
                    "<br>Último término: %{customdata[1]}"
                    "<br>Equipos considerados: %{customdata[2]}"
                    "<extra></extra>"
                ),
            )
        )

    # Mantener referencia del turno completo, pero permitir visualizar
    # ciclos que terminen ligeramente después de las 12 horas.
    max_val = pd.to_numeric(
        g[["Promedio_h", "Mediana_h"]].stack(),
        errors="coerce",
    ).max()
    y_max = max(12.0, float(max_val) if pd.notna(max_val) else 12.0)
    y_max = min(max(12.0, y_max + 0.25), 16.0)

    tickvals_y = list(range(0, int(y_max) + 1))
    if turno == "Día":
        ticktext_y = [
            f"{(7 + h) % 24:02d}:00"
            for h in tickvals_y
        ]
        titulo = "Turno Día · Tendencia del término del último ciclo"
    else:
        ticktext_y = [
            f"{(19 + h) % 24:02d}:00"
            for h in tickvals_y
        ]
        titulo = "Turno Noche · Tendencia del término del último ciclo"

    fechas = g["Fecha_Operativa_DT"].tolist()
    tickvals_x = fechas[::2] if len(fechas) > 12 else fechas
    ticktext_x = [
        pd.Timestamp(x).strftime("%d/%m")
        for x in tickvals_x
    ]

    fig.update_layout(
        **base_layout(
            460,
            title=dict(
                text=f"<b>{titulo}</b>",
                x=0.01,
                xanchor="left",
            ),
            margin=dict(l=90, r=30, t=60, b=80),
            xaxis=dict(
                title="Fecha operativa",
                tickmode="array",
                tickvals=tickvals_x,
                ticktext=ticktext_x,
                tickangle=-35,
                gridcolor="#eef2f7",
                fixedrange=True,
            ),
            yaxis=dict(
                title="Hora de término del último ciclo",
                range=[0, y_max],
                tickmode="array",
                tickvals=tickvals_y,
                ticktext=ticktext_y,
                gridcolor="#e5e7eb",
                zeroline=False,
                fixedrange=True,
            ),
            legend=dict(
                orientation="h",
                y=1.09,
                x=0.70,
                xanchor="left",
            ),
            hovermode="x unified",
        )
    )

    # Línea de referencia del fin nominal del turno.
    fig.add_hline(
        y=12,
        line_width=1.4,
        line_dash="dash",
        line_color="#b45309",
        annotation_text="Fin turno",
        annotation_position="top right",
    )

    for y_ref in [2, 4, 6, 8, 10]:
        fig.add_hline(
            y=y_ref,
            line_width=0.8,
            line_dash="dot",
            line_color="#cbd5e1",
            layer="below",
        )

    return fig


# ==========================================================
# PROCESAMIENTO MASIVO / CACHE
# ==========================================================


def guardar_resultado_desde_disco(item):
    """Procesa un archivo ya liberado del uploader y conserva solo resultados."""
    clave = item["clave"]
    path = Path(item["path"])
    try:
        resultado = procesar_archivo(
            path,
            nombre_archivo=item["nombre"],
            generar_visuales=False,
        )

        # Visuales diferidos: nunca se crean durante el lote masivo.
        fig = resultado.pop("fig", None)
        if fig is not None:
            plt.close(fig)
            del fig
        resultado.pop("plano_nav_png", None)
        resultado.pop("png_bytes", None)

        resultado["nombre_archivo"] = item["nombre"]
        resultado["error"] = None
        resultado["_cache_key"] = clave
        resultado["_source_path"] = str(path)

        st.session_state.procesados[clave] = resultado
    except Exception as exc:
        st.session_state.procesados[clave] = {
            "nombre_archivo": item["nombre"],
            "_cache_key": clave,
            "_source_path": str(path),
            "error": str(exc),
        }


def _visual_paths(cache_key):
    visual_dir = _session_work_dir() / "visuals"
    return (
        visual_dir / f"{cache_key}_boxplot.png",
        visual_dir / f"{cache_key}_plano.png",
    )


def asegurar_visuales_resultado(r):
    """Genera boxplot y plano ZDA solo cuando el usuario solicita ese ciclo."""
    cache_key = r.get("_cache_key") or hashlib.sha1(
        str(r.get("nombre_archivo", "")).encode("utf-8")
    ).hexdigest()
    box_path, nav_path = _visual_paths(cache_key)

    detalle = r.get("detalle")
    rep = r.get("resumen_reporte") or {}
    metadata = r.get("metadata") or rep

    if not box_path.exists() and isinstance(detalle, pd.DataFrame) and not detalle.empty:
        fig = generar_grafico(detalle, metadata)
        fig.savefig(box_path, format="png", dpi=170, bbox_inches="tight", facecolor="white")
        plt.close(fig)
        del fig
        gc.collect()

    if not nav_path.exists() and isinstance(detalle, pd.DataFrame) and not detalle.empty:
        nav_bytes = generar_plano_zda_png(detalle, metadata, resolution=150)
        if nav_bytes:
            nav_path.write_bytes(nav_bytes)
            del nav_bytes
            gc.collect()

    return box_path if box_path.exists() else None, nav_path if nav_path.exists() else None


# ==========================================================
# CARGA MASIVA EN DOS FASES
# ==========================================================


def _streamlit_version_tuple():
    """Versión numérica simple para habilitar carga de carpetas."""
    nums = re.findall(r"\d+", str(getattr(st, "__version__", "0.0.0")))
    vals = [int(x) for x in nums[:3]]
    while len(vals) < 3:
        vals.append(0)
    return tuple(vals)


with st.container(border=True, key="zona_carga"):
    st.subheader("Cargar ciclos (.ZDA)")
    st.caption(
        "Puedes agregar archivos individuales o seleccionar una carpeta completa. "
        "Los ciclos ya procesados permanecen cargados y puedes seguir incorporando "
        "nuevos archivos a medida que avanza el mes."
    )

    col_files, col_folder, col_clear = st.columns([1.15, 1.15, 1.45])

    with col_files:
        with st.popover(
            "📄 Elegir archivos",
            use_container_width=True,
            help="Selecciona uno o varios archivos ZDA.",
        ):
            archivos_individuales = st.file_uploader(
                "Archivos ZDA",
                type=["zda"],
                accept_multiple_files=True,
                key=f"uploader_archivos_{st.session_state.uploader_version}",
                label_visibility="collapsed",
                help=(
                    "Puedes seleccionar uno o varios archivos. "
                    "La carga se procesa automáticamente."
                ),
            )

    with col_folder:
        if _streamlit_version_tuple() >= (1, 49, 0):
            with st.popover(
                "📁 Elegir carpeta",
                use_container_width=True,
                help="Selector de directorio del navegador. Si solo permite seleccionar archivos, usa la opción de ruta local que aparece debajo.",
            ):
                archivos_carpeta = st.file_uploader(
                    "Carpeta con archivos ZDA",
                    type=["zda"],
                    accept_multiple_files="directory",
                    key=f"uploader_carpeta_{st.session_state.uploader_version}",
                    label_visibility="collapsed",
                    help=(
                        "Selecciona una carpeta. Solo se cargarán archivos ZDA; "
                        "también se consideran sus subcarpetas."
                    ),
                )
        else:
            archivos_carpeta = []
            st.button(
                "📁 Elegir carpeta",
                width="stretch",
                disabled=True,
                help="Requiere Streamlit 1.49 o superior.",
                key="folder_upload_disabled",
            )

    with col_clear:
        st.button(
            "🗑️ Borrar datos cargados",
            width="stretch",
            on_click=limpiar_analisis,
            help=(
                "Elimina de la sesión todos los ciclos procesados, "
                "archivos temporales y filtros asociados."
            ),
        )

    # Alternativa fiable para uso local: el navegador (sobre todo Safari)
    # puede abrir una carpeta mostrando su contenido y obligar a seleccionar
    # archivos. Esta ruta lee la carpeta completa directamente desde Python.
    with st.expander("📂 Cargar carpeta completa desde esta computadora", expanded=False):
        st.caption(
            "Pega la ruta de una carpeta del equipo donde se ejecuta Streamlit. "
            "Se incluyen sus subcarpetas y solo archivos .ZDA. En Finder (Mac): "
            "selecciona la carpeta y pulsa ⌥⌘C para copiar su ruta. "
            "Si Streamlit está publicado en la nube, esta opción NO puede leer carpetas de tu Mac."
        )
        ruta_zda_local = st.text_input(
            "Ruta de la carpeta local", key="ruta_carpeta_zda_local",
            placeholder="/Users/usuario/Documentos/ZDA",
        )
        if st.button("Cargar todos los ZDA de esta carpeta", key="btn_cargar_ruta_zda"):
            try:
                total_local, nuevos_local = preparar_carpeta_local(ruta_zda_local)
                if nuevos_local:
                    st.session_state.auto_process_staged = True
                    st.success(f"Detectados {total_local} ZDA; {nuevos_local} nuevos en cola de procesamiento.")
                else:
                    st.info(f"Se detectaron {total_local} ZDA; todos estaban cargados anteriormente.")
            except (ValueError, OSError, PermissionError) as exc:
                st.error(str(exc))

    archivos_individuales = archivos_individuales or []
    archivos_carpeta = archivos_carpeta or []
    seleccion = list(archivos_individuales) + list(archivos_carpeta)

    total_upload_mb = sum(
        int(getattr(a, "size", 0) or 0)
        for a in seleccion
    ) / (1024 * 1024)

    procesados_ok_actual = sum(
        1
        for r in st.session_state.procesados.values()
        if not r.get("error")
    )
    procesados_error_actual = sum(
        1
        for r in st.session_state.procesados.values()
        if r.get("error")
    )

    estado_carga = (
        f"{procesados_ok_actual} ciclo(s) cargado(s) correctamente."
    )
    if procesados_error_actual:
        estado_carga += f" · {procesados_error_actual} archivo(s) con error."
    if seleccion:
        estado_carga += (
            f" · Nuevos seleccionados: {len(seleccion)} "
            f"({total_upload_mb:,.1f} MB)."
        )

    st.markdown(
        f"<div style='font-size:0.98rem; color:#475467; padding:0.25rem 0 0.7rem 0;'>"
        f"{estado_carga}</div>",
        unsafe_allow_html=True,
    )

if seleccion and total_upload_mb >= 700:
    st.warning(
        "El lote seleccionado supera aproximadamente 700 MB. El modo masivo "
        "libera los archivos antes del parsing, pero el navegador debe transferir "
        "primero todo el lote. Si la carga supera la memoria disponible, conviene "
        "dividir únicamente la selección en dos carpetas o lotes."
    )

# Primera fase automática: al seleccionar archivos o una carpeta, se copian
# a disco temporal y se libera inmediatamente el uploader antes del parsing.
if seleccion:
    estado_stage = st.empty()
    estado_stage.write(
        f"Preparando {len(seleccion)} archivo(s) en disco temporal..."
    )
    nuevos = preparar_archivos_en_disco(seleccion)
    estado_stage.empty()

    # Cambiar la versión de ambos uploaders hace que Streamlit libere sus bytes
    # antes de comenzar el procesamiento intensivo del lote.
    st.session_state.uploader_version += 1
    st.session_state.auto_process_staged = bool(nuevos)
    st.rerun()

# Segunda fase: ya sin UploadedFile en RAM.
if st.session_state.auto_process_staged and st.session_state.staged_queue:
    cola = list(st.session_state.staged_queue)
    barra = st.progress(0)
    estado = st.empty()

    for i, item in enumerate(cola, start=1):
        estado.write(f"Procesando {i}/{len(cola)}: {item['nombre']}")
        guardar_resultado_desde_disco(item)
        st.session_state.staged_queue = [
            x for x in st.session_state.staged_queue
            if x.get("clave") != item.get("clave")
        ]
        gc.collect()
        barra.progress(i / len(cola))

    barra.empty()
    estado.empty()
    st.session_state.auto_process_staged = False
    st.rerun()

resultados_validos = [
    r
    for r in st.session_state.procesados.values()
    if not r.get("error")
]
errores = [
    r
    for r in st.session_state.procesados.values()
    if r.get("error")
]

if not resultados_validos:
    if errores:
        st.error(
            "No hay archivos procesados correctamente todavía."
        )
        st.dataframe(
            pd.DataFrame([
                {
                    "Archivo": r.get("nombre_archivo"),
                    "Error": r.get("error"),
                }
                for r in errores
            ]),
            width="stretch",
            hide_index=True,
        )
    else:
        st.write(
            "Selecciona uno o varios archivos o una carpeta completa para "
            "comenzar el análisis."
        )
    st.stop()


# ==========================================================
# CONSOLIDACIÓN
# ==========================================================

report_rows = [dict(r["resumen_reporte"]) for r in resultados_validos]
df_reportes = pd.DataFrame(report_rows)

# Clasificación uniforme para archivos ZDA.
df_reportes["Tipo_Disparo"] = df_reportes["Barrenos_Realizados"].apply(clasificar_tipo_disparo_v33)
df_reportes["Tipo_Roca"] = df_reportes["Plan_Perforacion"].apply(tipo_roca_desde_plan_texto)

def _operador_filtro_row(r):
    for campo in ("Operador_ZDA", "Operador", "Operario"):
        valor = r.get(campo)
        if valor is not None and not pd.isna(valor):
            texto = str(valor).strip()
            if texto and texto.upper() not in ("SIN DATO", "NONE", "NAN", "NULL"):
                return texto
    return "SIN DATO"

df_reportes["Operador_Filtro"] = df_reportes.apply(_operador_filtro_row, axis=1)
df_reportes["Considerado_KPI_Automatizacion"] = df_reportes["Tipo_Disparo"].eq("FRENTE")
for r in resultados_validos:
    rr = r["resumen_reporte"]
    rr["Tipo_Disparo"] = clasificar_tipo_disparo_v33(rr.get("Barrenos_Realizados"))
    rr["Tipo_Roca"] = tipo_roca_desde_plan_texto(rr.get("Plan_Perforacion"))
    rr["Operador_Filtro"] = _operador_filtro_row(rr)
    rr["Considerado_KPI_Automatizacion"] = rr["Tipo_Disparo"] == "FRENTE"

# HTML V33 solo agrega Resumen_Ciclos de reportes cuyo conteo está OK.
df_resumen = concatenar_dataframes(resultados_validos, "resumen_ciclo", solo_ok=True)
df_detalle = concatenar_dataframes(resultados_validos, "detalle")
df_atipicos = concatenar_dataframes(resultados_validos, "atipicos")
df_automatico = df_reportes.copy()

df_zda = df_reportes[df_reportes["Fuente"].eq("ZDA")].copy() if "Fuente" in df_reportes.columns else pd.DataFrame()

# ==========================================================
# FILTRO GLOBAL DE FECHAS EN SIDEBAR
# ==========================================================
# El contenedor fue creado al inicio dentro de st.sidebar para conservar
# la ubicación del filtro junto con Jumbos / Tipo / Roca / Operadores.
# Se llena aquí, cuando df_zda ya está consolidado, independientemente
# de qué sección del dashboard esté seleccionada.

global_fecha_inicio_zda = None
global_fecha_fin_zda = None

if sidebar_fecha_container is not None:
    with sidebar_fecha_container:
        st.markdown("#### Rango de fechas")
        st.caption("Aplica a Uso Automático y Primer Golpe.")

        if (
            not df_zda.empty
            and "Inicio_Perforacion_TS" in df_zda.columns
            and df_zda["Inicio_Perforacion_TS"].notna().any()
        ):
            fechas_sidebar = (
                df_zda.loc[
                    df_zda["Inicio_Perforacion_TS"].notna(),
                    "Inicio_Perforacion_TS",
                ]
                .apply(zda_operational_date)
            )

            fechas_sidebar = pd.to_datetime(
                fechas_sidebar,
                errors="coerce",
                utc=True,
            ).dropna()

            if not fechas_sidebar.empty:
                fecha_min_sidebar = fechas_sidebar.min().date()
                fecha_max_sidebar = fechas_sidebar.max().date()

                # Recuperar valores previos y ajustarlos al rango disponible.
                fecha_inicio_previa = st.session_state.get(
                    "fecha_inicio_zda_global",
                    fecha_min_sidebar,
                )
                fecha_fin_previa = st.session_state.get(
                    "fecha_fin_zda_global",
                    fecha_max_sidebar,
                )

                try:
                    fecha_inicio_previa = max(
                        fecha_min_sidebar,
                        min(fecha_inicio_previa, fecha_max_sidebar),
                    )
                except Exception:
                    fecha_inicio_previa = fecha_min_sidebar

                try:
                    fecha_fin_previa = max(
                        fecha_min_sidebar,
                        min(fecha_fin_previa, fecha_max_sidebar),
                    )
                except Exception:
                    fecha_fin_previa = fecha_max_sidebar

                # Si el rango previo queda invertido, volver al rango completo.
                if fecha_inicio_previa > fecha_fin_previa:
                    fecha_inicio_previa = fecha_min_sidebar
                    fecha_fin_previa = fecha_max_sidebar

                col_fecha_desde, col_fecha_hasta = st.columns(2, gap="small")
                with col_fecha_desde:
                    global_fecha_inicio_zda = st.date_input(
                        "Desde",
                        value=fecha_inicio_previa,
                        min_value=fecha_min_sidebar,
                        max_value=fecha_max_sidebar,
                        key="fecha_inicio_zda_global",
                        format="DD/MM/YYYY",
                    )
                with col_fecha_hasta:
                    global_fecha_fin_zda = st.date_input(
                        "Hasta",
                        value=fecha_fin_previa,
                        min_value=fecha_min_sidebar,
                        max_value=fecha_max_sidebar,
                        key="fecha_fin_zda_global",
                        format="DD/MM/YYYY",
                    )

                st.caption(
                    f"Rango mostrado: "
                    f"{global_fecha_inicio_zda.strftime('%d/%m/%Y')} "
                    f"→ {global_fecha_fin_zda.strftime('%d/%m/%Y')}"
                )
            else:
                st.caption("Sin fechas ZDA válidas.")
        else:
            st.caption("Sin fechas ZDA disponibles.")


# ==========================================================
# EXPORTACIÓN EXCEL - BD-PERFO + Resumen_Reportes + Resumen_Ciclos
# ==========================================================

df_bd_perfo = construir_bd_perfo(df_reportes, df_detalle)
excel_bytes = crear_excel_publicacion(df_bd_perfo, df_reportes, df_resumen)
nombre_excel = f"EBR Drill Analytics {datetime.now().strftime('%d-%m-%Y')}.xlsx"

st.download_button(
    "Descargar Excel consolidado",
    data=excel_bytes,
    file_name=nombre_excel,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)


def aplicar_filtro_fechas_global(df: pd.DataFrame) -> pd.DataFrame:
    """
    Aplica Fecha inicio / Fecha fin del sidebar a un DataFrame
    usando la columna Fecha_Inicio (dd/mm/YYYY).

    Si no hay rango seleccionado o no existe Fecha_Inicio,
    devuelve el DataFrame sin cambios.
    """
    if df is None or df.empty or "Fecha_Inicio" not in df.columns:
        return df.copy() if isinstance(df, pd.DataFrame) else pd.DataFrame()

    fecha_inicio = st.session_state.get("fecha_inicio_zda_global")
    fecha_fin = st.session_state.get("fecha_fin_zda_global")

    if fecha_inicio is None or fecha_fin is None:
        return df.copy()

    fechas = pd.to_datetime(
        df["Fecha_Inicio"],
        format="%d/%m/%Y",
        errors="coerce",
    )

    mask = (
        fechas.notna()
        & (fechas.dt.date >= fecha_inicio)
        & (fechas.dt.date <= fecha_fin)
    )

    return df.loc[mask].copy()


def _nombre_jumbo_resumen(valor):
    texto = str(valor or "").strip()
    m = re.fullmatch(r"JUMB0*(\d+)", texto, re.IGNORECASE)
    if m:
        return f"Jumbo {int(m.group(1))}"
    return texto or "-"


def _fecha_corta_es(valor):
    if valor is None or pd.isna(valor):
        return "-"
    ts = pd.Timestamp(valor)
    meses = [
        "ene.", "feb.", "mar.", "abr.", "may.", "jun.",
        "jul.", "ago.", "sep.", "oct.", "nov.", "dic.",
    ]
    return f"{ts.day:02d}-{meses[ts.month - 1]}"


def render_kpis_uso_automatico(df_auto: pd.DataFrame, df_pendientes: pd.DataFrame = None):
    """
    Resumen superior de la sección Uso Automático.

    Los indicadores se recalculan con el rango Fecha inicio / Fecha fin
    seleccionado en el panel lateral.
    """
    if df_auto is None or df_auto.empty:
        return

    base = aplicar_filtro_fechas_global(df_auto)
    total_ciclos = len(base)

    # ------------------------------------------------------
    # Ciclos cargados + desglose por jumbo
    # ------------------------------------------------------
    if "Jumbo" in base.columns:
        vc = (
            base["Jumbo"]
            .fillna("SIN DATO")
            .astype(str)
            .value_counts()
        )

        def _sort_jumbo(item):
            nombre, _ = item
            m = re.fullmatch(r"JUMB0*(\d+)", str(nombre), re.IGNORECASE)
            if m:
                return (0, int(m.group(1)))
            return (1, str(nombre))

        jumbo_detalle = " · ".join(
            f"{_nombre_jumbo_resumen(jumbo)}: {int(n)}"
            for jumbo, n in sorted(vc.items(), key=_sort_jumbo)
        )
    else:
        jumbo_detalle = f"{total_ciclos} ciclos"

    # ------------------------------------------------------
    # Rango de fechas
    # ------------------------------------------------------
    fechas = pd.Series(dtype="datetime64[ns]")
    if "Fecha_Inicio" in base.columns:
        fechas = pd.to_datetime(
            base["Fecha_Inicio"],
            format="%d/%m/%Y",
            errors="coerce",
        ).dropna()

    fecha_inicio_sel = st.session_state.get("fecha_inicio_zda_global")
    fecha_fin_sel = st.session_state.get("fecha_fin_zda_global")

    def _rango_compacto(d0, d1):
        # "24 ago – 03 sep": sin guiones ni puntos para que quepa en una sola línea.
        f0 = _fecha_corta_es(d0).replace("-", " ").replace(".", "")
        f1 = _fecha_corta_es(d1).replace("-", " ").replace(".", "")
        return f"{f0} – {f1}"

    if fecha_inicio_sel is not None and fecha_fin_sel is not None:
        rango_fecha = _rango_compacto(pd.Timestamp(fecha_inicio_sel), pd.Timestamp(fecha_fin_sel))
        fechas_sub = f"{len(fechas)} {'ciclo' if len(fechas) == 1 else 'ciclos'} en el rango"
    elif not fechas.empty:
        rango_fecha = _rango_compacto(fechas.min(), fechas.max())
        fechas_sub = f"{len(fechas)} {'ciclo' if len(fechas) == 1 else 'ciclos'} con fecha"
    else:
        rango_fecha = "-"
        fechas_sub = "Sin fechas válidas"

    # ------------------------------------------------------
    # Horas automáticas totales
    # ------------------------------------------------------
    auto_min = (
        pd.to_numeric(
            base.get(
                "Auto_Total_Brazos_min",
                pd.Series(index=base.index, dtype=float),
            ),
            errors="coerce",
        )
    )
    manual_min = (
        pd.to_numeric(
            base.get(
                "Manual_Total_Brazos_min",
                pd.Series(index=base.index, dtype=float),
            ),
            errors="coerce",
        )
    )

    mask_binario = auto_min.notna() & manual_min.notna()
    ciclos_binarios = int(mask_binario.sum())
    horas_auto = float(auto_min[mask_binario].sum()) / 60.0 if ciclos_binarios else 0.0

    horas_auto_txt = f"{horas_auto:,.2f} h".replace(",", "X").replace(".", ",").replace("X", ".")

    # ------------------------------------------------------
    # Sin operador registrado
    # ------------------------------------------------------
    if "Operador_Filtro" in base.columns:
        op = (
            base["Operador_Filtro"]
            .fillna("SIN DATO")
            .astype(str)
            .str.strip()
            .str.upper()
        )
        sin_operador = int(
            op.isin(["", "SIN DATO", "NONE", "NAN"]).sum()
        )
    else:
        campos_op = [
            c for c in ["Operador_ZDA", "Operador", "Operario"]
            if c in base.columns
        ]
        if campos_op:
            tiene_op = pd.Series(False, index=base.index)
            for campo in campos_op:
                vals = base[campo].fillna("").astype(str).str.strip()
                tiene_op = tiene_op | vals.ne("")
            sin_operador = int((~tiene_op).sum())
        else:
            sin_operador = total_ciclos

    n_manuales = (
        int(base["Operador_Asignado_Manual"].fillna(False).astype(bool).sum())
        if "Operador_Asignado_Manual" in base.columns else 0
    )
    pct_sin_operador = (
        sin_operador / total_ciclos * 100
        if total_ciclos > 0
        else 0.0
    )

    # ------------------------------------------------------
    # Tarjetas
    # ------------------------------------------------------
    st.markdown(
        """
        <style>
        /* Las 4 tarjetas viven en UNA grilla: todas toman la altura de la más alta. */
        .ebr-kpi-grid {
            display: grid;
            grid-template-columns: repeat(4, minmax(0, 1fr));
            grid-auto-rows: 1fr;
            gap: 1rem;
            align-items: stretch;
            margin-bottom: 0.45rem;
        }
        @media (max-width: 1100px) {
            .ebr-kpi-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
        }
        .ebr-kpi-card {
            display: flex;
            flex-direction: column;
            box-sizing: border-box;
            min-height: 150px;
            min-width: 0;
            border: 1px solid #e2e8f0;
            border-radius: 18px;
            background: #ffffff;
            padding: 1.15rem 1.25rem 1.05rem 1.25rem;
            box-shadow: 0 4px 14px rgba(15, 23, 42, 0.06);
            container-type: inline-size;   /* permite dimensionar la fuente según el ancho REAL de la tarjeta */
        }
        .ebr-kpi-label {
            font-size: 0.78rem;
            letter-spacing: 0.055em;
            text-transform: uppercase;
            color: #8a8a84;
            font-weight: 500;
            margin-bottom: 0.50rem;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }
        .ebr-kpi-value {
            font-size: clamp(1.15rem, 2.0vw, 1.85rem);          /* respaldo si no hay container queries */
            font-size: min(1.85rem, calc(100cqw / (var(--n, 8) * 0.62)));   /* --n = nº de caracteres */
            height: 2.2rem;              /* alto fijo: el texto de abajo queda alineado entre tarjetas */
            line-height: 2.2rem;
            color: #111111;
            font-weight: 750;
            margin-bottom: 0.30rem;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }
        .ebr-kpi-sub {
            font-size: 0.92rem;
            line-height: 1.30;
            color: #66645f;
            display: -webkit-box;
            -webkit-line-clamp: 2;
            -webkit-box-orient: vertical;
            overflow: hidden;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    cards = [
        (
            "Ciclos cargados",
            f"{total_ciclos}",
            jumbo_detalle,
        ),
        (
            "Rango de fechas",
            rango_fecha,
            fechas_sub,
        ),
        (
            "Horas automático (total)",
            horas_auto_txt,
            f"{ciclos_binarios} {'ciclo' if ciclos_binarios == 1 else 'ciclos'} con dato automático/manual",
        ),
        (
            "Sin operador registrado",
            f"{sin_operador}",
            f"{pct_sin_operador:.0f}% de los ciclos" + (f" · {n_manuales} {'asignado' if n_manuales == 1 else 'asignados'} a mano" if n_manuales else ""),
        ),
    ]

    from html import escape as _esc
    tarjetas_html = "".join(
        f'<div class="ebr-kpi-card" title="{_esc(str(label))}: {_esc(str(value))} · {_esc(str(sub))}">'
        f'<div class="ebr-kpi-label">{_esc(str(label))}</div>'
        f'<div class="ebr-kpi-value" style="--n:{max(len(str(value)), 4)}">{_esc(str(value))}</div>'
        f'<div class="ebr-kpi-sub">{_esc(str(sub))}</div>'
        f'</div>'
        for label, value, sub in cards
    )
    st.markdown(f'<div class="ebr-kpi-grid">{tarjetas_html}</div>', unsafe_allow_html=True)

    # Acceso directo al módulo donde se asigna el operador a los ciclos que no lo traen.
    # El módulo lista TODOS los ciclos cargados (sin filtros de fecha ni turno), por eso el botón
    # cuenta los pendientes totales y no solo los de la tarjeta.
    df_total = df_pendientes if df_pendientes is not None else df_auto
    pendientes_total = int(
        df_total["Operador_Filtro"].map(asig.es_sin_operador).sum()
        if "Operador_Filtro" in df_total.columns else sin_operador
    )
    if pendientes_total > 0:
        col_boton = st.columns(4)[3]
        if col_boton.button(
            f"Asignar operadores ({pendientes_total}) →", key="kpi_ir_asignar_operadores",
            width="stretch", help=f"{pendientes_total} ciclo(s) sin operador en total (todos los cargados, "
                                  "sin filtro de fechas ni de turno). Abre el módulo para escribir el operador.",
        ):
            st.session_state["seccion_analisis_principal"] = "Asignar operadores"
            st.rerun(scope="app")


# ==========================================================
# ASIGNAR OPERADORES (ciclos cuyo ZDA no trae el operador)
# ==========================================================

def _tabla_operadores(df_reportes: pd.DataFrame) -> pd.DataFrame:
    """Un renglón por ciclo con el operador vigente, su origen (ZDA / Manual / Pendiente),
    una sugerencia y la clave que identifica el ciclo."""
    base = asegurar_fechahora(df_reportes.copy()).reset_index(drop=True)
    if base.empty:
        return base
    efectivo = base["Operador_Filtro"].map(
        lambda v: None if asig.es_sin_operador(v) else str(v).strip()
    )
    manual = (
        base["Operador_Asignado_Manual"].fillna(False).astype(bool)
        if "Operador_Asignado_Manual" in base.columns else pd.Series(False, index=base.index)
    )
    base["_operador"] = efectivo
    base["_origen"] = np.where(manual, "Manual", np.where(efectivo.notna(), "ZDA", "Pendiente"))
    base["_sugerencia"] = asig.sugerir(base.assign(operador=efectivo), col_operador="operador")
    base["_clave"] = [asig.clave_ciclo(r) for r in base.to_dict("records")]
    return base


def _aplicar_cambios_operador(cambios: dict, filas: dict) -> int:
    """Aplica {clave: operador | None} a las asignaciones y guarda. Devuelve cuántas cambiaron."""
    data = _asignaciones_operador()
    n = 0
    for clave, nombre in cambios.items():
        rep = filas.get(clave)
        if rep is None:
            continue
        n += bool(asig.quitar(data, clave) if nombre is None else asig.poner(data, rep, nombre))
    if n:
        _guardar_asignaciones_operador()
    return n


def _terminar_cambios_operador(n: int) -> None:
    """Mensaje, reinicio del editor y recarga completa para que todo el app vea el cambio."""
    st.session_state["_msg_operadores"] = (
        f"Se guardaron {n} asignación(es). Ya se reflejan en filtros, gráficos y Excel."
        if n else "No había cambios que guardar."
    )
    st.session_state["_op_editor_version"] = st.session_state.get("_op_editor_version", 0) + 1
    st.rerun(scope="app")


@fragment
def render_asignar_operadores_section(df_reportes: pd.DataFrame):
    st.subheader("Asignar operadores")
    st.caption(
        "El ZDA guarda el operador como texto libre (campo tunnel_id). Cuando el equipo no lo "
        "trae, el ciclo queda \"SIN DATO\". Aquí puedes asignarlo a mano: se guarda en este equipo "
        "y se aplica a filtros, gráficos, Excel y Resultados por archivo, sin modificar los ZDA."
    )
    if df_reportes is None or df_reportes.empty:
        st.info("No hay ciclos cargados.")
        return

    base = _tabla_operadores(df_reportes)
    if base.empty:
        st.info("Ningún ciclo tiene fecha de inicio válida.")
        return
    data = _asignaciones_operador()

    aviso = st.session_state.pop("_msg_operadores", None)
    if aviso:
        st.success(aviso)

    n_pend = int((base["_origen"] == "Pendiente").sum())
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Ciclos cargados", len(base))
    m2.metric("Operador del ZDA", int((base["_origen"] == "ZDA").sum()))
    m3.metric("Asignados a mano", int((base["_origen"] == "Manual").sum()))
    m4.metric("Pendientes", n_pend)
    if n_pend == 0:
        st.success("Todos los ciclos tienen operador.")

    # ---- Filtros ----------------------------------------------------------------------
    f1, f2, f3 = st.columns([2, 3.2, 2.4], vertical_alignment="bottom")
    jumbos = sorted(base["Jumbo"].dropna().astype(str).unique())
    sel_jumbos = f1.multiselect("Jumbo", jumbos, default=jumbos, key="op_filtro_jumbo")
    estado = f2.radio(
        "Mostrar", ["Pendientes", "Asignados a mano", "Todos"],
        index=0, horizontal=True, key="op_filtro_estado",
    )
    marcar_todas = f3.toggle("Marcar todas las filas visibles", key="op_marcar_todas")

    mask = base["Jumbo"].astype(str).isin(sel_jumbos)
    if estado == "Pendientes":
        mask &= base["_origen"].eq("Pendiente")
    elif estado == "Asignados a mano":
        mask &= base["_origen"].eq("Manual")
    visibles = base[mask].reset_index(drop=True)

    conocidos = asig.lista_operadores(
        data,
        extras=list(CATALOGO_OPERADORES_ZDA.values()) + [o for o in base["_operador"].dropna().unique()],
    )
    filas = {
        r["_clave"]: {k: r.get(k) for k in ("Jumbo", "Ciclo", "Fecha_Inicio", "Hora_Inicio")}
        for r in base.to_dict("records")
    }

    # ---- Tabla editable ---------------------------------------------------------------
    if visibles.empty:
        if not (estado == "Pendientes" and n_pend == 0):      # si no hay pendientes, ya se avisó arriba
            st.info("No hay ciclos con estos filtros.")
        vista = editado = None
    else:
        st.caption(
            "Elige el operador en la columna **Operador** (o marca varias filas y usa "
            "\"Asignar a varias filas\" más abajo). **Texto del ZDA** es lo que el operador "
            "escribió en el equipo: úsalo como pista. **Sugerencia** es el operador del ciclo más "
            "cercano (hasta 6 h) del mismo jumbo; solo se aplica si tú lo confirmas."
        )
        vista = pd.DataFrame({
            "Sel": bool(marcar_todas),
            "Jumbo": visibles["Jumbo"].astype(str),
            "Ciclo": visibles["Ciclo"],
            "Fecha": visibles["Fecha_Inicio"],
            "Hora": visibles["Hora_Inicio"],
            "Barrenos": visibles["Barrenos_Realizados"] if "Barrenos_Realizados" in visibles.columns else None,
            "Texto del ZDA": (visibles["Operador_ZDA_Raw"].fillna("") if "Operador_ZDA_Raw" in visibles.columns else ""),
            "Origen": visibles["_origen"],
            "Sugerencia": visibles["_sugerencia"],
            "Operador": visibles["_operador"],
            "clave": visibles["_clave"],
        })
        version = st.session_state.get("_op_editor_version", 0)
        editado = st.data_editor(
            vista,
            key=f"op_editor_{version}_{int(marcar_todas)}_{estado}_{'-'.join(sel_jumbos)}",
            hide_index=True, width="stretch", num_rows="fixed",
            height=int(min(38 * (len(vista) + 1) + 3, 560)),
            disabled=[c for c in vista.columns if c not in ("Sel", "Operador")],
            column_config={
                "Sel": st.column_config.CheckboxColumn("Sel.", width="small"),
                "Ciclo": st.column_config.NumberColumn("Ciclo", format="%d", width="small"),
                "Barrenos": st.column_config.NumberColumn("Barrenos", format="%d", width="small"),
                "Operador": st.column_config.SelectboxColumn(
                    "Operador", options=conocidos, required=False, width="medium",
                    help="Elige un operador de la lista. Vacío = sin operador.",
                ),
                "clave": None,
            },
        )
        if st.button("Guardar asignaciones", type="primary", key="op_guardar"):
            cambios = asig.cambios_desde_editor(vista, editado)
            _terminar_cambios_operador(_aplicar_cambios_operador(cambios, filas))

        with st.expander("Asignar a varias filas (marca las filas en la columna Sel.)"):
            marcadas = editado.loc[editado["Sel"].fillna(False).astype(bool)]
            st.caption(f"{len(marcadas)} fila(s) marcada(s). Las ediciones sueltas de la tabla también se guardan.")
            nombre_bulk = st.selectbox(
                "Operador para las filas marcadas", conocidos, index=None,
                placeholder="Elegir operador…", key="op_bulk_nombre",
            )
            c1, c2 = st.columns(2)
            if c1.button("Aplicar a las filas marcadas", key="op_bulk_aplicar", disabled=nombre_bulk is None or marcadas.empty):
                cambios = asig.cambios_desde_editor(vista, editado)
                cambios.update({c: nombre_bulk for c in marcadas["clave"]})
                _terminar_cambios_operador(_aplicar_cambios_operador(cambios, filas))
            con_sug = marcadas[marcadas["Sugerencia"].astype(str).str.strip().ne("")]
            if c2.button(
                f"Aplicar la sugerencia a las marcadas ({len(con_sug)})", key="op_bulk_sugerencia",
                disabled=con_sug.empty,
            ):
                cambios = asig.cambios_desde_editor(vista, editado)
                cambios.update(dict(zip(con_sug["clave"], con_sug["Sugerencia"])))
                _terminar_cambios_operador(_aplicar_cambios_operador(cambios, filas))

    # ---- Lista de operadores ---------------------------------------------------------
    with st.expander("Lista de operadores"):
        st.caption("Disponibles para elegir: " + ", ".join(conocidos))
        with st.form("op_form_nuevo", clear_on_submit=True):
            nuevo = st.text_input("Agregar un operador nuevo a la lista", placeholder="Nombre y apellido")
            if st.form_submit_button("Agregar"):
                if not asig.normalizar_nombre(nuevo):
                    st.warning("Escribe un nombre.")
                elif asig.agregar_operador(data, nuevo):
                    _guardar_asignaciones_operador()
                    st.rerun(scope="fragment")
                else:
                    st.warning("Ese nombre ya está en la lista.")

    # ---- Copia de seguridad ----------------------------------------------------------
    with st.expander("Copia de seguridad y compartir"):
        ruta = st.session_state.get("_asig_ruta") or str(asig.ruta_asignaciones())
        st.caption(
            f"Las asignaciones se guardan en: `{ruta}`. Si publicas la app en la nube ese archivo "
            "puede borrarse al reiniciar: descarga una copia y reimpórtala cuando la necesites."
        )
        st.download_button(
            "Descargar asignaciones (CSV)", data=asig.a_csv(data),
            file_name="operadores_asignados.csv", mime="text/csv",
            disabled=not data.get("asignaciones"), key="op_descargar",
        )
        subido = st.file_uploader("Importar asignaciones (CSV)", type=["csv"], key="op_importar_archivo")
        if subido is not None and st.button("Importar", key="op_importar_boton"):
            n_ok, avisos = asig.importar_csv(data, subido.getvalue())
            if n_ok:
                _guardar_asignaciones_operador()
            for texto in avisos[:5]:
                st.warning(texto)
            if n_ok:
                _terminar_cambios_operador(n_ok)
            elif not avisos:
                st.info("El archivo no tenía filas para importar.")


# ==========================================================
# BLOQUE 1 - AUTOMATIZACIÓN
# ==========================================================

def _controles_grafico(clave: str, con_curva: bool = False) -> dict:
    """Controles que van SOBRE cada gráfico (reemplazan las opciones del panel lateral).

    - Tipo de gráfico: Líneas / Barras (solo cambia la presentación).
    - Etiquetas: muestra u oculta los valores sobre puntos o barras.
    - Línea curva (opcional): suaviza la línea; se desactiva en barras.
    Devuelve {"tipo": "Líneas"|"Barras", "etiquetas": bool, "curva": bool}.
    """
    opciones = ["Líneas", "Barras"]
    anchos = [2.4, 1.3, 1.7, 4.0] if con_curva else [2.4, 1.3, 5.7]
    cols = st.columns(anchos, vertical_alignment="center")
    with cols[0]:
        valor = st.segmented_control(
            "Tipo de gráfico", opciones, default="Líneas",
            key=f"{clave}_tipo", label_visibility="collapsed",
        )
    tipo = valor if valor in opciones else "Líneas"      # None si se des-selecciona
    with cols[1]:
        etiquetas = st.checkbox(
            "Etiquetas", value=False, key=f"{clave}_lbl",
            help="Muestra u oculta los valores sobre los puntos o las barras.",
        )
    curva = True
    if con_curva:
        with cols[2]:
            curva = st.checkbox(
                "Línea curva", value=True, key=f"{clave}_curva",
                disabled=(tipo == "Barras"),
                help="Solo en líneas. Desmarcado: solo puntos. Marcado: curva suavizada que "
                     "conecta los ciclos sin modificar los valores reales.",
            )
    return {"tipo": tipo, "etiquetas": bool(etiquetas), "curva": bool(curva)}


@fragment
def render_automation_section(
    df_automatico: pd.DataFrame,
    sel_jumbos,
    sel_tipos,
    sel_rocas,
    sel_operadores,
    sel_turnos=None,
):
    if df_automatico.empty:
        st.info("Sin datos suficientes de automatización.")
        return

    # Las tarjetas también respetan el turno; el botón "Asignar operadores" sigue contando
    # todos los ciclos cargados (el módulo los lista sin filtros).
    render_kpis_uso_automatico(filtrar_por_turno(df_automatico, sel_turnos), df_pendientes=df_automatico)
    st.markdown("<div style='height:0.35rem'></div>", unsafe_allow_html=True)

    # Base visible por Jumbo / Tipo / Roca / Fecha.
    # Se conserva aparte para el gráfico de horas acumuladas, que además
    # debe poder mostrar los ciclos "Sin registrar".
    df_visible_base = df_automatico[
        df_automatico["Jumbo"].astype(str).isin(
            [str(x) for x in sel_jumbos]
        )
        & df_automatico["Tipo_Disparo"].isin(sel_tipos)
        & df_automatico["Tipo_Roca"].isin(sel_rocas)
    ].copy()

    df_visible_base = filtrar_por_turno(df_visible_base, sel_turnos)
    df_visible_base = aplicar_filtro_fechas_global(df_visible_base)

    # Los gráficos de evolución por operador siguen respetando
    # el filtro explícito de operadores del panel lateral.
    df_visible = df_visible_base[
        df_visible_base["Operador_Filtro"].isin(sel_operadores)
    ].copy()

    st.subheader("Evolución del movimiento automático")
    ctl_evol = _controles_grafico("auto_evol", con_curva=True)

    fig_auto = grafico_auto(
        df_visible,
        ctl_evol["etiquetas"],
        ctl_evol["curva"],
        ctl_evol["tipo"],
    )

    if fig_auto is not None:
        st.plotly_chart(
            fig_auto,
            width="stretch",
            config={"displaylogo": False},
        )
    else:
        st.info(
            "No hay ciclos visibles con datos de movimiento automático "
            "para los filtros globales seleccionados."
        )

    st.subheader("Evolución del movimiento automático por operador")
    st.caption(
        "Cada línea representa un operador y cada punto un ciclo/round. "
        "El color se asigna por horas automáticas acumuladas del rango visible: "
        "azul = más horas, verde = segundo, naranja = tercero y rojo = menos horas. "
        "El símbolo identifica el jumbo: círculo = JUMB001, cuadrado = JUMB002. "
        "En barras, los ciclos simultáneos de dos jumbos se muestran lado a lado "
        "(el jumbo y la hora exacta figuran en el hover)."
    )
    ctl_oper = _controles_grafico("auto_oper", con_curva=True)

    fig_auto_operador = grafico_auto_por_operador(
        df_visible,
        ctl_oper["etiquetas"],
        ctl_oper["curva"],
        ctl_oper["tipo"],
    )

    if fig_auto_operador is not None:
        st.plotly_chart(
            fig_auto_operador,
            width="stretch",
            config={"displaylogo": False},
        )
    else:
        st.info(
            "No hay ciclos visibles con operador identificado y datos de "
            "movimiento automático para los filtros seleccionados."
        )

    st.subheader("Uso automático promedio por operador")
    st.caption(
        "Porcentaje de movimiento automático de cada operador: horas en automático ÷ "
        "(automático + manual) de todos sus ciclos visibles. Es el mismo «Global» de la leyenda "
        "de la evolución por operador. Cada operador conserva el color de ese gráfico. "
        "No incluye los ciclos sin operador registrado."
    )

    fig_horas_operador = grafico_uso_auto_promedio_operador(
        df_visible_base,
        sel_operadores,
    )

    if fig_horas_operador is not None:
        st.plotly_chart(
            fig_horas_operador,
            width="stretch",
            config={
                "displaylogo": False,
                "scrollZoom": False,
            },
        )
    else:
        st.info(
            "No hay ciclos visibles con operador identificado y datos de movimiento "
            "automático y manual para los filtros seleccionados."
        )

    st.subheader("Uso automático por brazo")
    ctl_brazo = _controles_grafico("auto_brazo")

    jumbos_visibles = sorted(
        df_visible.get(
            "Jumbo",
            pd.Series(dtype=object),
        ).dropna().astype(str).unique()
    )

    if not jumbos_visibles:
        st.info(
            "No hay jumbos visibles con los filtros globales seleccionados."
        )
        return

    hubo_grafico = False

    for jumbo in jumbos_visibles:
        fig_arm = grafico_brazos(
            df_visible,
            jumbo,
            ctl_brazo["etiquetas"],
            ctl_brazo["tipo"],
        )
        if fig_arm is not None:
            hubo_grafico = True
            st.plotly_chart(
                fig_arm,
                width="stretch",
                config={"displaylogo": False},
            )

    if not hubo_grafico:
        st.info(
            "No hay datos por brazo para los filtros globales seleccionados."
        )


# ==========================================================
# BLOQUE 2 - BARRENOS CUT
# ==========================================================

@fragment
def render_cut_section(
    df_resumen: pd.DataFrame,
    df_reportes: pd.DataFrame,
    sel_jumbos,
    sel_tipos,
    sel_rocas,
    sel_operadores,
    sel_turnos=None,
):
    st.subheader("Evolución de la longitud perforada en barrenos Cut")

    df_cut = preparar_cut(
        df_resumen,
        df_reportes,
    )

    if df_cut.empty:
        st.info("Sin datos suficientes de barrenos Cut.")
        return

    df_cut = filtrar_por_turno(df_cut, sel_turnos)

    st.caption(
        "Cada punto (o barra) representa la mediana de la longitud perforada "
        "de los barrenos Cut de cada ciclo. Se aplican los filtros "
        "globales del panel lateral. En barras el eje parte de 0."
    )
    ctl_cut = _controles_grafico("cut")

    fig_cut = grafico_cut(
        df_cut,
        sel_jumbos,
        sel_tipos,
        sel_rocas,
        sel_operadores,
        ctl_cut["etiquetas"],
        ctl_cut["tipo"],
    )

    if fig_cut is not None:
        st.plotly_chart(
            fig_cut,
            width="stretch",
            config={"displaylogo": False},
        )
    else:
        st.info(
            "No hay ciclos visibles con los filtros globales seleccionados."
        )


# ==========================================================
# BLOQUE 3 - TIEMPOS DE CICLO DE PERFORACIÓN
# ==========================================================

@fragment
def _mostrar_tabla_por_turno(tabla: pd.DataFrame) -> None:
    """Muestra la tabla de indicadores por turno como DOS tablas lado a lado (día | noche).

    Antes día y noche compartían una sola tabla de 9 columnas y se leían como un mismo bloque.
    Cada turno lleva su rótulo con el horario y columnas más cortas; la primera columna
    (Jumbo) se repite para leer cada fila sin mirar la otra tabla.
    """
    def _bloque(sufijo):
        cols = [c for c in tabla.columns if str(c).endswith(f" {sufijo}")]
        t = tabla[["Jumbo"] + cols].copy()
        # "Inicio más temprano día" -> "Más temprano"; "Ciclos día" -> "Ciclos"
        nombres = []
        for c in cols:
            n = str(c)[: -len(sufijo) - 1]
            for pref in ("Inicio ", "Fin "):
                if n.startswith(pref):
                    n = n[len(pref):]
            nombres.append(n[:1].upper() + n[1:])
        t.columns = ["Jumbo"] + nombres
        return t

    dia, noche = _bloque("día"), _bloque("noche")
    col_dia, col_noche = st.columns(2, gap="large")
    with col_dia:
        st.markdown("**Turno día** · 07:00–19:00")
        st.dataframe(dia, width="stretch", hide_index=True)
    with col_noche:
        st.markdown("**Turno noche** · 19:00–07:00")
        st.dataframe(noche, width="stretch", hide_index=True)


def render_zda_section(
    df_zda: pd.DataFrame,
    sel_jumbos,
    sel_tipos,
    sel_rocas,
    sel_operadores,
    sel_turnos=None,
):
    st.subheader("Tiempos de ciclo de perforación")

    if (
        df_zda.empty
        or not {
            "Inicio_Perforacion_TS",
            "Fin_Perforacion_TS",
        }.issubset(df_zda.columns)
    ):
        st.info(
            "Sin datos ZDA suficientes para mostrar tiempos de ciclo."
        )
        return

    zda_all = df_zda[
        df_zda["Inicio_Perforacion_TS"].notna()
        & df_zda["Fin_Perforacion_TS"].notna()
    ].copy()

    if zda_all.empty:
        st.info("Sin ventanas de perforación ZDA válidas.")
        return

    st.caption(
        "Análisis de las ventanas reales de perforación por fecha operativa y turno. "
        "Turno día: 07:00–19:00 · Turno noche: 19:00–07:00."
    )

    # Clasificación unificada: Bottom + Easer + Cut + Contour.
    zda_all["Tipo_Disparo"] = (
        zda_all["Barrenos_Realizados"]
        .apply(clasificar_tipo_disparo_v33)
    )

    zda_rows = zda_all[
        zda_all["Jumbo"].astype(str).isin(
            [str(x) for x in sel_jumbos]
        )
        & zda_all["Tipo_Disparo"].isin(sel_tipos)
        & zda_all["Tipo_Roca"].isin(sel_rocas)
        & zda_all["Operador_Filtro"].isin(sel_operadores)
    ].copy()
    zda_rows = filtrar_por_turno(zda_rows, sel_turnos, col_ts="Inicio_Perforacion_TS")

    # ------------------------------------------------------
    # FILTRO GLOBAL DE FECHAS DEL PANEL LATERAL.
    # Los widgets se muestran siempre en el sidebar; aquí únicamente
    # se aplican los valores seleccionados a Tiempos de Ciclo.
    # ------------------------------------------------------
    fecha_inicio_zda = st.session_state.get("fecha_inicio_zda_global")
    fecha_fin_zda = st.session_state.get("fecha_fin_zda_global")

    if not zda_rows.empty:
        zda_rows["_Fecha_Operativa_Filtro"] = (
            zda_rows["Inicio_Perforacion_TS"]
            .apply(zda_operational_date)
        )

        fecha_op_date = pd.to_datetime(
            zda_rows["_Fecha_Operativa_Filtro"],
            errors="coerce",
            utc=True,
        ).dt.date

        if fecha_inicio_zda is not None and fecha_fin_zda is not None:
            if fecha_inicio_zda > fecha_fin_zda:
                st.error("La Fecha inicio no puede ser posterior a la Fecha fin.")
                zda_rows = zda_rows.iloc[0:0].copy()
            else:
                mask_fecha_zda = (
                    (fecha_op_date >= fecha_inicio_zda)
                    & (fecha_op_date <= fecha_fin_zda)
                )
                zda_rows = zda_rows[mask_fecha_zda].copy()

    st.caption(
        "Los filtros globales del panel lateral actualizan la gráfica "
        "y los resúmenes de esta sección."
    )

    # Expanders (no st.markdown("####...") fijo): permiten contraer cada tabla para no
    # empujar el resto de "Primer Golpe" hacia abajo cuando no se necesitan.
    with st.expander("Hora promedio de inicio por jumbo y turno", expanded=True):
        shift = resumen_turnos_zda(
            zda_rows
        )

        if not shift.empty:
            _mostrar_tabla_por_turno(shift)
            st.caption(
                "“Ciclos” cuenta todos los rounds del turno. "
                "Para el promedio, el inicio más temprano y el más tarde "
                "se considera solo el primer round de cada jumbo por "
                "fecha operativa y turno."
            )
        else:
            st.info(
                "No hay ciclos visibles para calcular los indicadores de inicio."
            )

    with st.expander("Hora promedio de fin por jumbo y turno", expanded=True):
        shift_fin = resumen_fin_turnos_zda(
            zda_rows
        )

        if not shift_fin.empty:
            _mostrar_tabla_por_turno(shift_fin)
            st.caption(
                "“Ciclos” cuenta todos los rounds iniciados en el turno. "
                "Para el promedio, el fin más temprano y el más tarde se considera "
                "únicamente el término del último round de cada jumbo por "
                "fecha operativa y turno."
            )
        else:
            st.info(
                "No hay ciclos visibles para calcular los indicadores de fin."
            )

    with st.expander(
        "Tipo de disparo",
        expanded=False,
    ):
        type_summary = resumen_tipos_zda(
            zda_rows
        )

        if not type_summary.empty:
            st.dataframe(
                type_summary,
                width="stretch",
                hide_index=True,
            )
        else:
            st.info(
                "No hay ciclos visibles para resumir el tipo de disparo."
            )

        st.caption(
            "Criterio sobre barrenos de frente "
            "(Bottom + Easer + Cut + Contour). "
            "Reaming y Casing no se consideran para la clasificación."
        )


    ciclos_turno = preparar_timeline_ciclos_turno(
        zda_rows
    )

    if ciclos_turno.empty:
        st.markdown(
            "#### Piloto · Timeline de ciclos por día y turno"
        )
        st.info(
            "No hay ciclos suficientes para construir el timeline por turno."
        )
    else:
        st.markdown(
            "#### Piloto · Timeline de ciclos por día y turno"
        )
        st.caption(
            "Cada segmento representa un ciclo/round físico completo desde "
            "Inicio hasta Fin. JUMB001 se representa como contorno cerrado y "
            "JUMB002 como sólido, manteniendo la misma codificación en barras "
            "y puntos. El filtro de fechas de la cabecera aplica a ambos turnos."
        )

        ciclos_turno_filtrado = ciclos_turno.copy()

        solo_puntos_inicio_timeline = st.checkbox(
            "Mostrar solo puntos de inicio",
            value=False,
            key="solo_puntos_inicio_timeline",
            help=(
                "Oculta las barras completas y muestra únicamente la hora "
                "de inicio de cada ciclo para facilitar la identificación "
                "de patrones de arranque."
            ),
        )

        solo_primer_inicio_timeline = st.checkbox(
            "Mostrar solo el primer inicio por equipo y turno",
            value=False,
            key="solo_primer_inicio_timeline",
            disabled=not solo_puntos_inicio_timeline,
            help=(
                "Disponible cuando se activa la vista de puntos. "
                "Si un jumbo tiene dos o más ciclos en el mismo turno, "
                "solo se muestra el primer inicio y se ocultan los demás."
            ),
        )

        mostrar_clusters_inicio_timeline = st.checkbox(
            "Mostrar clusters de hora de inicio",
            value=False,
            key="mostrar_clusters_inicio_timeline",
            disabled=not solo_puntos_inicio_timeline,
            help=(
                "Identifica automáticamente entre 2 y 4 grupos de hora de "
                "primer inicio mediante K-Means 1D. Al activarlo se usa "
                "automáticamente solo el primer inicio de cada jumbo por turno. "
                "Las zonas muestran los clusters y la línea punteada su centro."
            ),
        )

        if solo_puntos_inicio_timeline and mostrar_clusters_inicio_timeline:
            st.caption(
                "Clusters calculados por separado para Día y Noche sobre el "
                "primer inicio de cada jumbo por fecha operativa. K se selecciona "
                "automáticamente entre 2 y 4 usando Silhouette Score."
            )

        # Orden de lectura fijo: primero Turno Día y luego Turno Noche.
        fig_turno_dia = grafico_timeline_ciclos_turno(
            ciclos_turno_filtrado,
            "Día",
            solo_puntos_inicio=solo_puntos_inicio_timeline,
            solo_primer_inicio=solo_primer_inicio_timeline,
            mostrar_clusters=(
                solo_puntos_inicio_timeline
                and mostrar_clusters_inicio_timeline
            ),
        )

        if fig_turno_dia is not None:
            st.plotly_chart(
                fig_turno_dia,
                width="stretch",
                config={
                    "displaylogo": False,
                    "scrollZoom": False,
                },
            )
        else:
            st.info(
                "No hay ciclos visibles en el Turno Día."
            )

        if (
            solo_puntos_inicio_timeline
            and mostrar_clusters_inicio_timeline
        ):
            cluster_dia = analizar_clusters_primer_inicio(
                ciclos_turno_filtrado,
                "Día",
            )
            render_resumen_clusters_primer_inicio(
                cluster_dia,
                "Día",
            )

        fig_turno_noche = grafico_timeline_ciclos_turno(
            ciclos_turno_filtrado,
            "Noche",
            solo_puntos_inicio=solo_puntos_inicio_timeline,
            solo_primer_inicio=solo_primer_inicio_timeline,
            mostrar_clusters=(
                solo_puntos_inicio_timeline
                and mostrar_clusters_inicio_timeline
            ),
        )

        if fig_turno_noche is not None:
            st.plotly_chart(
                fig_turno_noche,
                width="stretch",
                config={
                    "displaylogo": False,
                    "scrollZoom": False,
                },
            )
        else:
            st.info(
                "No hay ciclos visibles en el Turno Noche."
            )

        if (
            solo_puntos_inicio_timeline
            and mostrar_clusters_inicio_timeline
        ):
            cluster_noche = analizar_clusters_primer_inicio(
                ciclos_turno_filtrado,
                "Noche",
            )
            render_resumen_clusters_primer_inicio(
                cluster_noche,
                "Noche",
            )

        # ------------------------------------------------------
        # DISTRIBUCIÓN DE PRIMEROS INICIOS
        # ------------------------------------------------------
        st.markdown(
            "#### Distribución de primeros inicios"
        )
        st.caption(
            "El histograma agrupa el primer inicio de cada jumbo por fecha y turno "
            "en intervalos de 30 minutos. La curva suavizada permite identificar "
            "visualmente las horas de mayor concentración."
        )

        primeros_distribucion = preparar_primeros_inicios_distribucion(
            ciclos_turno_filtrado
        )

        if primeros_distribucion.empty:
            st.info(
                "No hay información suficiente para construir la distribución "
                "de primeros inicios."
            )
        else:
            fig_dist_dia = grafico_distribucion_primeros_inicios(
                primeros_distribucion,
                "Día",
            )
            if fig_dist_dia is not None:
                st.plotly_chart(
                    fig_dist_dia,
                    width="stretch",
                    config={
                        "displaylogo": False,
                        "scrollZoom": False,
                    },
                )

            fig_dist_noche = grafico_distribucion_primeros_inicios(
                primeros_distribucion,
                "Noche",
            )
            if fig_dist_noche is not None:
                st.plotly_chart(
                    fig_dist_noche,
                    width="stretch",
                    config={
                        "displaylogo": False,
                        "scrollZoom": False,
                    },
                )

        # ------------------------------------------------------
        # TENDENCIA DIARIA DEL PRIMER INICIO CONSOLIDADO
        # ------------------------------------------------------
        st.markdown(
            "#### Tendencia diaria de la hora de primer inicio"
        )
        st.caption(
            "Para cada fecha y turno se toma únicamente el primer inicio "
            "de cada jumbo disponible. Luego se calcula el promedio y la "
            "mediana entre los equipos. Con 2 equipos ambos valores son "
            "iguales; con 3 o más pueden diferenciarse."
        )

        ctrl_prom, ctrl_med, _ = st.columns([1.0, 1.0, 3.2])
        with ctrl_prom:
            mostrar_promedio_inicio = st.checkbox(
                "Mostrar promedio",
                value=False,
                key="mostrar_promedio_inicio_diario",
            )
        with ctrl_med:
            mostrar_mediana_inicio = st.checkbox(
                "Mostrar mediana",
                value=True,
                key="mostrar_mediana_inicio_diario",
            )

        resumen_inicio_diario = preparar_tendencia_inicio_diario(
            ciclos_turno_filtrado
        )

        if not mostrar_promedio_inicio and not mostrar_mediana_inicio:
            st.info(
                "Activa Promedio y/o Mediana para mostrar la tendencia."
            )
        elif resumen_inicio_diario.empty:
            st.info(
                "No hay información suficiente para calcular la tendencia diaria."
            )
        else:
            fig_inicio_dia = grafico_tendencia_inicio_diario(
                resumen_inicio_diario,
                "Día",
                mostrar_promedio=mostrar_promedio_inicio,
                mostrar_mediana=mostrar_mediana_inicio,
            )
            if fig_inicio_dia is not None:
                st.plotly_chart(
                    fig_inicio_dia,
                    width="stretch",
                    config={
                        "displaylogo": False,
                        "scrollZoom": False,
                    },
                )

            fig_inicio_noche = grafico_tendencia_inicio_diario(
                resumen_inicio_diario,
                "Noche",
                mostrar_promedio=mostrar_promedio_inicio,
                mostrar_mediana=mostrar_mediana_inicio,
            )
            if fig_inicio_noche is not None:
                st.plotly_chart(
                    fig_inicio_noche,
                    width="stretch",
                    config={
                        "displaylogo": False,
                        "scrollZoom": False,
                    },
                )

        # ------------------------------------------------------
        # TENDENCIA DIARIA DEL TÉRMINO DEL ÚLTIMO CICLO
        # ------------------------------------------------------
        st.markdown(
            "#### Tendencia diaria de la hora de término del último ciclo"
        )
        st.caption(
            "Para cada fecha y turno se toma el Fin más tardío de cada jumbo. "
            "Si un equipo realizó dos o más rounds, se considera el término "
            "del último round. Luego se calcula el promedio y la mediana "
            "entre los equipos disponibles."
        )

        ctrl_prom_fin, ctrl_med_fin, _ = st.columns([1.0, 1.0, 3.2])
        with ctrl_prom_fin:
            mostrar_promedio_fin = st.checkbox(
                "Mostrar promedio",
                value=False,
                key="mostrar_promedio_ultimo_fin_diario",
            )
        with ctrl_med_fin:
            mostrar_mediana_fin = st.checkbox(
                "Mostrar mediana",
                value=True,
                key="mostrar_mediana_ultimo_fin_diario",
            )

        resumen_ultimo_fin_diario = preparar_tendencia_ultimo_fin_diario(
            ciclos_turno_filtrado
        )

        if not mostrar_promedio_fin and not mostrar_mediana_fin:
            st.info(
                "Activa Promedio y/o Mediana para mostrar la tendencia."
            )
        elif resumen_ultimo_fin_diario.empty:
            st.info(
                "No hay información suficiente para calcular el término "
                "diario del último ciclo."
            )
        else:
            fig_fin_dia = grafico_tendencia_ultimo_fin_diario(
                resumen_ultimo_fin_diario,
                "Día",
                mostrar_promedio=mostrar_promedio_fin,
                mostrar_mediana=mostrar_mediana_fin,
            )
            if fig_fin_dia is not None:
                st.plotly_chart(
                    fig_fin_dia,
                    width="stretch",
                    config={
                        "displaylogo": False,
                        "scrollZoom": False,
                    },
                )

            fig_fin_noche = grafico_tendencia_ultimo_fin_diario(
                resumen_ultimo_fin_diario,
                "Noche",
                mostrar_promedio=mostrar_promedio_fin,
                mostrar_mediana=mostrar_mediana_fin,
            )
            if fig_fin_noche is not None:
                st.plotly_chart(
                    fig_fin_noche,
                    width="stretch",
                    config={
                        "displaylogo": False,
                        "scrollZoom": False,
                    },
                )

        with st.expander(
            "Detalle de ciclos mostrados en el timeline",
            expanded=False,
        ):
            detalle_timeline = ciclos_turno_filtrado[
                [
                    "Fecha_Operativa",
                    "Turno",
                    "Jumbo",
                    "Ciclo",
                    "Inicio",
                    "Fin",
                    "Duracion",
                    "Tipo_Disparo",
                    "Tipo_Roca",
                    "Labor",
                    "Operador",
                    "Barrenos",
                    "Sobrepasa_Turno",
                ]
            ].copy()

            st.dataframe(
                detalle_timeline,
                width="stretch",
                hide_index=True,
            )


# ==========================================================
# BLOQUE 4 - CLASIFICACIÓN DE DISPAROS
# ==========================================================

@fragment
def render_classification_section(
    df_reportes: pd.DataFrame,
    df_automatico: pd.DataFrame,
    df_atipicos: pd.DataFrame,
    sel_jumbos,
    sel_tipos,
    sel_rocas,
    sel_operadores,
    sel_turnos=None,
):
    filtrados = df_reportes[
        df_reportes["Jumbo"].astype(str).isin(
            [str(x) for x in sel_jumbos]
        )
        & df_reportes["Tipo_Disparo"].isin(sel_tipos)
        & df_reportes["Tipo_Roca"].isin(sel_rocas)
        & df_reportes["Operador_Filtro"].isin(sel_operadores)
    ].copy()
    filtrados = filtrar_por_turno(filtrados, sel_turnos)

    st.caption(
        "Esta sección utiliza los mismos filtros globales del panel lateral."
    )

    col_class, col_read = st.columns(
        [1, 1.25]
    )

    with col_class:
        if filtrados.empty:
            st.info(
                "No hay reportes visibles con los filtros globales seleccionados."
            )
        else:
            resumen_clase = (
                filtrados["Tipo_Disparo"]
                .value_counts()
                .reindex(
                    TIPOS_DISPARO,
                    fill_value=0,
                )
                .rename_axis(
                    "Tipo de disparo"
                )
                .reset_index(
                    name="N reportes"
                )
            )

            resumen_clase = resumen_clase[
                resumen_clase["N reportes"] > 0
            ]

            st.dataframe(
                resumen_clase,
                width="stretch",
                hide_index=True,
            )

    with col_read:
        st.markdown("#### Resumen de lectura")


        zda_rows_all = (
            filtrados[
                filtrados["Fuente"].eq("ZDA")
            ]
            if "Fuente" in filtrados.columns
            else pd.DataFrame()
        )

        lectura_ok = int(
            filtrados.get(
                "Lectura_Confiable",
                pd.Series(dtype=object),
            ).eq("OK").sum()
        )


        zda_ok = (
            int(
                zda_rows_all.get(
                    "Lectura_Confiable",
                    pd.Series(dtype=object),
                ).eq("OK").sum()
            )
            if not zda_rows_all.empty
            else 0
        )

        conteo_ok = int(
            filtrados.get(
                "Estado_Conteo",
                pd.Series(dtype=object),
            ).eq("OK").sum()
        )

        metros_ok = int(
            filtrados.get(
                "Estado_Metros_Tipos",
                pd.Series(dtype=object),
            ).eq("OK").sum()
        )

        # Atípicos pertenecientes a ciclos actualmente visibles.
        n_atipicos_visible = 0
        if (
            not df_atipicos.empty
            and {"Jumbo", "Ciclo"}.issubset(df_atipicos.columns)
            and {"Jumbo", "Ciclo"}.issubset(filtrados.columns)
        ):
            claves_visibles = {
                (str(j), str(c))
                for j, c in zip(
                    filtrados["Jumbo"],
                    filtrados["Ciclo"],
                )
            }

            n_atipicos_visible = sum(
                (str(j), str(c)) in claves_visibles
                for j, c in zip(
                    df_atipicos["Jumbo"],
                    df_atipicos["Ciclo"],
                )
            )

        r1, r2, r3 = st.columns(3)

        r1.metric("Archivos ZDA", len(filtrados))
        r2.metric("Lectura ZDA OK", f"{zda_ok}/{len(zda_rows_all)}")
        r3.metric("Revisar lectura", f"{len(zda_rows_all) - zda_ok}")

        s1, s2, s3, s4 = st.columns(4)

        s1.metric(
            "Conteo OK",
            f"{conteo_ok}/{len(filtrados)}",
        )
        s2.metric(
            "Metros por tipo OK",
            f"{metros_ok}/{len(filtrados)}",
        )
        s3.metric(
            "Revisar lectura",
            len(filtrados) - lectura_ok,
        )
        s4.metric(
            "Atípicos",
            n_atipicos_visible,
        )

    st.markdown(
        "#### Uso automático por ciclo"
    )

    auto_filtrado = df_automatico[
        df_automatico["Jumbo"].astype(str).isin(
            [str(x) for x in sel_jumbos]
        )
        & df_automatico["Tipo_Disparo"].isin(sel_tipos)
        & df_automatico["Tipo_Roca"].isin(sel_rocas)
        & df_automatico["Operador_Filtro"].isin(sel_operadores)
    ].copy()
    auto_filtrado = filtrar_por_turno(auto_filtrado, sel_turnos)

    cols_auto = [
        c
        for c in [
            "Fecha_Inicio",
            "Jumbo",
            "Ciclo",
            "Operador_Filtro",
            "Barrenos_Realizados",
            "Tipo_Disparo",
            "Considerado_KPI_Automatizacion",
            "Auto_Total_Brazos_min",
            "Manual_Total_Brazos_min",
            "Pct_Movimiento_Automatico_Brazos",
            "Pct_Automatico_Brazo1",
            "Pct_Automatico_Brazo2",
            "Gap_Automatico_Brazos_pp",
        ]
        if c in auto_filtrado.columns
    ]

    tabla_auto = auto_filtrado[
        cols_auto
    ].copy()

    tabla_auto = tabla_auto.rename(
        columns={
            "Operador_Filtro": "Operador",
            "Considerado_KPI_Automatizacion": "KPI Auto",
            "Auto_Total_Brazos_min": "Auto total min",
            "Manual_Total_Brazos_min": "Manual total min",
            "Pct_Movimiento_Automatico_Brazos": "Auto Jumbo %",
            "Pct_Automatico_Brazo1": "Brazo 1 %",
            "Pct_Automatico_Brazo2": "Brazo 2 %",
            "Gap_Automatico_Brazos_pp": "Gap brazos pp",
        }
    )

    if "KPI Auto" in tabla_auto.columns:
        tabla_auto["KPI Auto"] = (
            tabla_auto["KPI Auto"]
            .map(
                {
                    True: "Sí",
                    False: "No",
                }
            )
        )

    st.dataframe(
        tabla_auto,
        width="stretch",
        hide_index=True,
        height=360,
    )


# ==========================================================
# BLOQUE 5 - RESULTADOS POR ARCHIVO
# ==========================================================

@fragment

def _extraer_rop_mwd_desde_zda(
    source_path: Path,
    archivo_interno: str,
) -> pd.DataFrame:
    """
    Extrae la progresión del barreno y la tasa de penetración ROP
    registrada en el archivo MWD interno del ZDA.

    Layout MWD usado por los ZDA DD322i analizados:
      timestamp  -> bytes 0..7 del registro
      posición   -> float32 en offset +8
      ROP        -> float32 en offset +44

    Cada registro ocupa 122 bytes y el encabezado del archivo 131 bytes.
    """
    cols = [
        "Timestamp",
        "Profundidad_m",
        "ROP_m_min",
    ]

    if (
        source_path is None
        or not source_path.exists()
        or not archivo_interno
        or source_path.suffix.lower() != ".zda"
    ):
        return pd.DataFrame(columns=cols)

    try:
        with zipfile.ZipFile(source_path, "r") as zf:
            if archivo_interno not in zf.namelist():
                return pd.DataFrame(columns=cols)

            data = zf.read(archivo_interno)
    except Exception:
        return pd.DataFrame(columns=cols)

    if len(data) < 131 + 122:
        return pd.DataFrame(columns=cols)

    filas = []
    nrec = (len(data) - 131) // 122

    for i in range(nrec):
        off = 131 + i * 122
        if off + 122 > len(data):
            break

        try:
            lo, hi = struct.unpack_from("<II", data, off)
            ts = lo + hi * 4294967296
            profundidad = struct.unpack_from("<f", data, off + 8)[0]
            rop = struct.unpack_from("<f", data, off + 44)[0]
        except Exception:
            continue

        if not (
            1577836800 < ts < 2051222400
            and np.isfinite(profundidad)
            and np.isfinite(rop)
            and 0 <= profundidad < 30
            and 0 <= rop < 20
        ):
            continue

        filas.append(
            {
                "Timestamp": pd.to_datetime(
                    ts,
                    unit="s",
                    utc=True,
                    errors="coerce",
                ),
                "Profundidad_m": float(profundidad),
                "ROP_m_min": float(rop),
            }
        )

    df = pd.DataFrame(filas, columns=cols)
    if df.empty:
        return df

    df = (
        df.dropna(
            subset=[
                "Profundidad_m",
                "ROP_m_min",
            ]
        )
        .sort_values(
            [
                "Profundidad_m",
                "Timestamp",
            ]
        )
        .reset_index(drop=True)
    )

    # El MWD puede registrar más de una muestra en la misma progresión.
    # Consolidar por profundidad evita líneas verticales artificiales.
    df = (
        df.groupby(
            "Profundidad_m",
            as_index=False,
        )
        .agg(
            Timestamp=("Timestamp", "min"),
            ROP_m_min=("ROP_m_min", "median"),
        )
        .sort_values("Profundidad_m")
        .reset_index(drop=True)
    )

    # Suavizado visual robusto. El valor original continúa disponible
    # en el tooltip; no se reemplaza el dato fuente.
    df["ROP_suave_m_min"] = (
        df["ROP_m_min"]
        .rolling(
            window=5,
            center=True,
            min_periods=1,
        )
        .median()
    )

    return df


def _catalogo_rop_resultados(
    resultados_validos,
    df_reportes: pd.DataFrame,
    sel_jumbos,
    sel_tipos,
    sel_rocas,
    sel_operadores,
    sel_turnos=None,
):
    """
    Construye catálogo de ciclos ZDA visibles para la viñeta ROP.
    Respeta filtros globales y rango de fechas.
    """
    if not resultados_validos:
        return []

    permitidos = None

    if (
        isinstance(df_reportes, pd.DataFrame)
        and not df_reportes.empty
    ):
        rep = df_reportes.copy()

        if "Fuente" in rep.columns:
            rep = rep[
                rep["Fuente"]
                .fillna("")
                .astype(str)
                .str.upper()
                .eq("ZDA")
            ].copy()

        if "Jumbo" in rep.columns:
            rep = rep[
                rep["Jumbo"]
                .astype(str)
                .isin([str(x) for x in sel_jumbos])
            ]

        if "Tipo_Disparo" in rep.columns:
            rep = rep[
                rep["Tipo_Disparo"].isin(sel_tipos)
            ]

        if "Tipo_Roca" in rep.columns:
            rep = rep[
                rep["Tipo_Roca"].isin(sel_rocas)
            ]

        if "Operador_Filtro" in rep.columns:
            rep = rep[
                rep["Operador_Filtro"].isin(sel_operadores)
            ]

        rep = filtrar_por_turno(rep, sel_turnos)
        rep = aplicar_filtro_fechas_global(rep)

        if {
            "Jumbo",
            "Ciclo",
        }.issubset(rep.columns):
            permitidos = set(
                zip(
                    rep["Jumbo"].astype(str),
                    rep["Ciclo"].astype(str),
                )
            )

    catalogo = []

    for idx, r in enumerate(resultados_validos):
        rep = r.get("resumen_reporte") or {}
        meta = r.get("metadata") or {}

        fuente = str(
            rep.get("Fuente")
            or meta.get("Fuente")
            or r.get("fuente")
            or ""
        ).upper()

        if fuente != "ZDA":
            continue

        mwd = r.get("mwd_barrenos")
        if (
            not isinstance(mwd, pd.DataFrame)
            or mwd.empty
            or "Archivo_Interno" not in mwd.columns
        ):
            continue

        jumbo = str(
            rep.get("Jumbo")
            or meta.get("Jumbo")
            or ""
        )
        ciclo = str(
            rep.get("Ciclo")
            or meta.get("Ciclo")
            or ""
        )

        if (
            permitidos is not None
            and (jumbo, ciclo) not in permitidos
        ):
            continue

        fecha = (
            rep.get("Fecha_Inicio")
            or meta.get("Fecha_Inicio")
            or "-"
        )
        operador = (
            rep.get("Operador_Filtro")
            or meta.get("Operador")
            or meta.get("Operador_ZDA")
            or "SIN DATO"
        )

        # Total de barrenos del ciclo.
        # Prioridad: Barrenos_Realizados -> Barrenos_ZDA -> cantidad MWD.
        barrenos_ciclo = pd.to_numeric(
            pd.Series([
                rep.get("Barrenos_Realizados")
                if rep.get("Barrenos_Realizados") is not None
                else rep.get("Barrenos_ZDA")
            ]),
            errors="coerce",
        ).iloc[0]

        if pd.isna(barrenos_ciclo):
            barrenos_ciclo = pd.to_numeric(
                pd.Series([
                    meta.get("Barrenos_Realizados")
                    if meta.get("Barrenos_Realizados") is not None
                    else meta.get("Barrenos_ZDA")
                ]),
                errors="coerce",
            ).iloc[0]

        if pd.isna(barrenos_ciclo):
            barrenos_ciclo = len(mwd)

        barrenos_ciclo = int(round(float(barrenos_ciclo)))

        catalogo.append(
            {
                "idx": idx,
                "resultado": r,
                "Jumbo": jumbo,
                "Ciclo": ciclo,
                "Fecha": fecha,
                "Operador": operador,
                "Barrenos": barrenos_ciclo,
                "Archivo": r.get(
                    "nombre_archivo",
                    rep.get("Archivo_ZDA", "-"),
                ),
            }
        )

    def _sort_key(x):
        fecha = pd.to_datetime(
            x.get("Fecha"),
            format="%d/%m/%Y",
            errors="coerce",
        )
        return (
            fecha if pd.notna(fecha) else pd.Timestamp.min,
            x.get("Jumbo", ""),
            str(x.get("Ciclo", "")),
        )

    return sorted(
        catalogo,
        key=_sort_key,
        reverse=True,
    )


def grafico_rop_por_barreno(
    df_rop: pd.DataFrame,
):
    """
    Curva ROP vs progresión/profundidad del barreno.
    """
    if (
        df_rop is None
        or df_rop.empty
        or "Profundidad_m" not in df_rop.columns
        or "ROP_m_min" not in df_rop.columns
    ):
        return None

    work = df_rop.copy()

    y_plot = (
        work["ROP_suave_m_min"]
        if "ROP_suave_m_min" in work.columns
        else work["ROP_m_min"]
    )

    fig = go.Figure()

    fig.add_trace(
        go.Scatter(
            x=work["Profundidad_m"],
            y=y_plot,
            mode="lines",
            name="ROP",
            line=dict(
                width=2.7,
                color="#2F7ED8",
                shape="spline",
                smoothing=0.55,
            ),
            fill="tozeroy",
            fillcolor="rgba(47,126,216,0.10)",
            customdata=np.column_stack(
                [
                    work["ROP_m_min"],
                ]
            ),
            hovertemplate=(
                "<b>Profundidad: %{x:.2f} m</b>"
                "<br>ROP: %{customdata[0]:.2f} m/min"
                "<extra></extra>"
            ),
        )
    )

    rop_prom = pd.to_numeric(
        work["ROP_m_min"],
        errors="coerce",
    ).mean()

    if pd.notna(rop_prom):
        fig.add_hline(
            y=float(rop_prom),
            line_width=1.2,
            line_dash="dot",
            line_color="#64748b",
            opacity=0.75,
            annotation_text=f"Promedio {rop_prom:.2f} m/min",
            annotation_position="top right",
        )

    ymax = pd.to_numeric(
        y_plot,
        errors="coerce",
    ).max()

    y_top = max(
        3.5,
        float(ymax) * 1.12
        if pd.notna(ymax)
        else 3.5,
    )

    fig.update_layout(
        **base_layout(
            450,
            margin=dict(
                l=75,
                r=40,
                t=38,
                b=65,
            ),
            xaxis=dict(
                title="Progresión del barreno (m)",
                rangemode="tozero",
                gridcolor="#e6edf5",
                zeroline=False,
            ),
            yaxis=dict(
                title="ROP (m/min)",
                range=[0, y_top],
                gridcolor="#e6edf5",
                zeroline=False,
            ),
            hovermode="x",
            showlegend=False,
        )
    )

    return fig


# Barrenos cuyo archivo MWD no calza con ningún registro de boom.dat (mismo Brazo+Secuencia):
# no se conoce su collar ejecutado ni si el sensor de profundidad arrancó en 0. Se muestran
# aparte y no entran en las curvas, el promedio ni las métricas del round (ver justo abajo).
TIPO_SOSPECHOSO = "Sin identificar (posible rehecho)"

TIPO_COLOR_ROP = {
    "Bottom": "#8B5CF6", "Easer": "#F59E0B", "Cut": "#EF4444",
    "Contour": "#0EA5E9", "Reaming": "#10B981", "Reference": "#6B7280",
}


def _color_tipo_rop(tipo) -> str:
    return TIPO_COLOR_ROP.get(str(tipo), "#94A3B8")


@st.cache_data(show_spinner=False, max_entries=8)
def _datos_conjunto_rop(source_path_str, mtime_ns, size_bytes, items):
    """Curva ROP y resumen de TODOS los barrenos de un round, para la vista de conjunto.

    `items`: tupla de (archivo_interno, barreno_id, tipo, x, y_ejecutado, z) — uno por barreno
    con MWD y coordenada ejecutada conocida (join Boom/Brazo + Secuencia ya resuelto por el
    llamador). Cacheado por archivo + su lista de barrenos: leer las ~50 curvas de un round
    toma alrededor de un segundo; sin caché se repetiría en cada interacción del selector de barreno.
    Devuelve (curvas largas para el overlay, resumen una fila por barreno para el mapa).
    """
    source_path = Path(source_path_str)
    curvas, resumen = [], []
    for archivo_interno, barreno_id, tipo, x, y_exec, z in items:
        df = _extraer_rop_mwd_desde_zda(source_path, archivo_interno)
        if df.empty:
            continue
        c = df[["Profundidad_m", "ROP_m_min", "ROP_suave_m_min"]].copy()
        c["Barreno_ID"] = barreno_id
        c["Tipo"] = tipo
        curvas.append(c)
        rop = pd.to_numeric(df["ROP_m_min"], errors="coerce")
        resumen.append({
            "Barreno_ID": barreno_id, "Tipo": tipo, "X": x, "Y": y_exec, "Z": z,
            "ROP_prom": rop.mean(), "ROP_mediana": rop.median(),
            "Profundidad_max": pd.to_numeric(df["Profundidad_m"], errors="coerce").max(),
            "N_muestras": len(df),
        })
    df_curvas = pd.concat(curvas, ignore_index=True) if curvas else pd.DataFrame(
        columns=["Profundidad_m", "ROP_m_min", "ROP_suave_m_min", "Barreno_ID", "Tipo"])
    df_resumen = pd.DataFrame(resumen)
    return df_curvas, df_resumen


def grafico_rop_round_overlay(df_curvas: pd.DataFrame):
    """Todas las curvas ROP del round superpuestas (una línea fina por barreno, coloreada por
    tipo), más la curva promedio del round en negro. Es la forma habitual de reportar el
    conjunto de curvas MWD de un disparo: permite ver la dispersión y detectar barrenos
    atípicos de un vistazo, algo que una sola curva a la vez no muestra.

    No dibuja los barrenos "Sin identificar" (sin match en boom.dat): de esos no se sabe si el
    sensor de profundidad arrancó en 0, así que su curva podría no representar avance real y
    distorsionaría tanto el eje de profundidad como la curva promedio."""
    if df_curvas is None or df_curvas.empty:
        return None

    df_curvas = df_curvas[df_curvas["Tipo"] != TIPO_SOSPECHOSO]
    if df_curvas.empty:
        return None

    fig = go.Figure()
    vistos = set()
    for barreno_id, g in df_curvas.groupby("Barreno_ID", sort=False):
        tipo = g["Tipo"].iloc[0]
        color = _color_tipo_rop(tipo)
        fig.add_trace(go.Scatter(
            x=g["Profundidad_m"], y=g["ROP_suave_m_min"], mode="lines",
            line=dict(width=1.3, color=color), opacity=0.55,
            name=str(tipo), legendgroup=str(tipo), showlegend=tipo not in vistos,
            hovertemplate=f"Barreno {barreno_id} · {tipo}<br>Profundidad: %{{x:.2f}} m"
                          "<br>ROP: %{y:.2f} m/min<extra></extra>",
        ))
        vistos.add(tipo)

    # Curva promedio del round: mediana de todos los barrenos en bins de profundidad de 0.1 m.
    prof = pd.to_numeric(df_curvas["Profundidad_m"], errors="coerce")
    bins = (prof / 0.1).round() * 0.1
    prom = df_curvas.assign(_bin=bins).groupby("_bin")["ROP_m_min"].median().reset_index()
    fig.add_trace(go.Scatter(
        x=prom["_bin"], y=prom["ROP_m_min"], mode="lines",
        line=dict(width=3.2, color="#111827", shape="spline", smoothing=0.4),
        name="Promedio del round", legendgroup="_prom",
        hovertemplate="Promedio del round<br>Profundidad: %{x:.2f} m<br>ROP: %{y:.2f} m/min<extra></extra>",
    ))

    ymax = pd.to_numeric(df_curvas["ROP_suave_m_min"], errors="coerce").max()
    fig.update_layout(**base_layout(
        480, margin=dict(l=75, r=30, t=30, b=65),
        xaxis=dict(title="Progresión del barreno (m)", rangemode="tozero", gridcolor="#e6edf5"),
        yaxis=dict(title="ROP (m/min)", range=[0, max(3.5, float(ymax) * 1.1 if pd.notna(ymax) else 3.5)],
                   gridcolor="#e6edf5"),
        legend=dict(orientation="h", y=-0.18, title="Tipo de barreno"),
        hovermode="closest",
    ))
    return fig


@st.cache_data(show_spinner=False, max_entries=8)
def _perfil_nominal_rop(source_path_str, mtime_ns, size_bytes):
    """Perfil nominal del frente (contorno de diseño), leído directamente de round-*.dat.

    Es un archivo aparte y pequeño (~1.3 KB) del propio ZDA: no hace falta reprocesar todo el
    archivo con el otro pipeline (el de Eficiencia de Perforación) para obtener el contorno.
    Devuelve la lista de puntos (X, Z) del polígono, o [] si no se pudo decodificar.
    """
    try:
        with zipfile.ZipFile(source_path_str, "r") as zf:
            names = zf.namelist()
            round_dat = next(
                (n for n in names if re.match(r"round-[^/]*\.dat$", n, re.I)
                 and not re.search(r"boom|counters|mwd", n, re.I)),
                None,
            )
            if not round_dat:
                return []
            perfil = parse_round_dat_profile(zf.read(round_dat))
    except (OSError, zipfile.BadZipFile):
        return []
    return perfil["polygon"] if perfil.get("decoded") else []


def grafico_rop_round_mapa(df_resumen: pd.DataFrame, perfil_nominal=None):
    """Mapa del frente (plano de navegación): un punto por barreno en su posición ejecutada
    (X, Z), coloreado por su ROP promedio. Es la representación espacial que permite ver si
    algún sector del frente (o algún tipo de barreno) perfora sistemáticamente más lento,
    algo que las curvas por separado no muestran porque no tienen referencia de ubicación."""
    if df_resumen is None or df_resumen.empty or df_resumen["X"].isna().all():
        return None

    d = df_resumen.dropna(subset=["X", "Z", "ROP_prom"]).copy()
    if d.empty:
        return None

    fig = go.Figure()
    if perfil_nominal:
        px = [p[0] for p in perfil_nominal] + [perfil_nominal[0][0]]
        pz = [p[1] for p in perfil_nominal] + [perfil_nominal[0][1]]
        fig.add_trace(go.Scatter(
            x=px, y=pz, mode="lines", line=dict(color="#94A3B8", width=1.6, dash="dot"),
            name="Perfil nominal", hoverinfo="skip", showlegend=False,
        ))
    fig.add_trace(go.Scatter(
        x=d["X"], y=d["Z"], mode="markers+text",
        text=d["Barreno_ID"].astype(str), textposition="top center",
        textfont=dict(size=9, color="#475569"),
        marker=dict(
            size=16, color=d["ROP_prom"], colorscale="RdYlGn", cmid=float(d["ROP_prom"].mean()),
            showscale=True, colorbar=dict(title="ROP prom.<br>(m/min)"),
            line=dict(color="#ffffff", width=1.2),
        ),
        customdata=np.column_stack([d["Tipo"], d["ROP_prom"], d["Profundidad_max"]]),
        hovertemplate="<b>Barreno %{text}</b> · %{customdata[0]}"
                      "<br>ROP promedio: %{customdata[1]:.2f} m/min"
                      "<br>Profundidad: %{customdata[2]:.2f} m<extra></extra>",
    ))
    fig.update_layout(**base_layout(
        480, margin=dict(l=60, r=20, t=30, b=55),
        xaxis=dict(title="X (m)", gridcolor="#e6edf5", zeroline=True, scaleanchor="y"),
        yaxis=dict(title="Z (m)", gridcolor="#e6edf5", zeroline=True),
        hovermode="closest", showlegend=False,
    ))
    return fig


def render_rop_section(
    resultados_validos,
    df_reportes,
    sel_jumbos,
    sel_tipos,
    sel_rocas,
    sel_operadores,
    sel_turnos=None,
):
    """
    Viñeta ROP: selección de ciclo y barreno + perfil de penetración.
    """
    st.subheader("ROP por barreno")
    st.caption(
        "Visualiza la tasa de penetración (ROP) registrada en el MWD "
        "y su comportamiento a lo largo de la perforación del barreno."
    )

    catalogo = _catalogo_rop_resultados(
        resultados_validos,
        df_reportes,
        sel_jumbos,
        sel_tipos,
        sel_rocas,
        sel_operadores,
        sel_turnos,
    )

    if not catalogo:
        st.info(
            "No hay ciclos ZDA visibles con información MWD "
            "para los filtros seleccionados."
        )
        return

    labels_ciclo = []
    map_ciclo = {}

    for item in catalogo:
        label = (
            f"{item['Fecha']} · {item['Jumbo']} · "
            f"Ciclo {item['Ciclo']} · {item['Operador']} · "
            f"{item['Barrenos']} barrenos"
        )
        # Garantizar claves únicas incluso si el label coincide.
        key = f"{label} · {item['idx']}"
        labels_ciclo.append(key)
        map_ciclo[key] = item

    col_sel1, col_sel2 = st.columns(
        [1.35, 1.0]
    )

    with col_sel1:
        ciclo_sel_key = st.selectbox(
            "Ciclo / archivo",
            labels_ciclo,
            key="rop_ciclo_sel",
            format_func=lambda x: x.rsplit(" · ", 1)[0],
        )

    item_sel = map_ciclo[ciclo_sel_key]
    r = item_sel["resultado"]
    mwd = r.get("mwd_barrenos").copy()

    # ------------------------------------------------------
    # Identificador de barreno del plan ZDA
    # ------------------------------------------------------
    # El MWD se identifica técnicamente por Brazo + Secuencia.
    # En el detalle del ZDA (boom.dat), el campo ID corresponde
    # al número/identificador de barreno reconstruido desde el ZDA.
    #
    # Cruce:
    #   detalle.ID + detalle.Boom + detalle.Secuencia
    #                      ↕
    #          MWD.Brazo + MWD.Secuencia
    detalle_round_rop = r.get("detalle")

    mapa_barreno_id = {}

    if (
        isinstance(detalle_round_rop, pd.DataFrame)
        and not detalle_round_rop.empty
        and {
            "ID",
            "Boom",
            "Secuencia",
        }.issubset(detalle_round_rop.columns)
    ):
        det_map = detalle_round_rop[
            [
                "ID",
                "Boom",
                "Secuencia",
            ]
        ].copy()

        det_map["Boom"] = pd.to_numeric(
            det_map["Boom"],
            errors="coerce",
        )
        det_map["Secuencia"] = pd.to_numeric(
            det_map["Secuencia"],
            errors="coerce",
        )

        det_map = det_map.dropna(
            subset=[
                "Boom",
                "Secuencia",
            ]
        )

        for _, det_row in det_map.iterrows():
            key_bs = (
                int(det_row["Boom"]),
                int(det_row["Secuencia"]),
            )
            barreno_id = det_row.get("ID")

            if (
                barreno_id is not None
                and not pd.isna(barreno_id)
                and str(barreno_id).strip()
            ):
                mapa_barreno_id[key_bs] = str(
                    barreno_id
                ).strip()

    def _barreno_id_mwd(row):
        brazo = pd.to_numeric(
            pd.Series([row.get("Brazo")]),
            errors="coerce",
        ).iloc[0]
        secuencia = pd.to_numeric(
            pd.Series([row.get("Secuencia")]),
            errors="coerce",
        ).iloc[0]

        if pd.isna(brazo) or pd.isna(secuencia):
            return "-"

        return mapa_barreno_id.get(
            (
                int(brazo),
                int(secuencia),
            ),
            "-",
        )

    mwd["Barreno_ID"] = mwd.apply(
        _barreno_id_mwd,
        axis=1,
    )

    # Preferir barrenos con profundidad útil.
    if "Profundidad_Max_MWD_m" in mwd.columns:
        mwd["_Prof"] = pd.to_numeric(
            mwd["Profundidad_Max_MWD_m"],
            errors="coerce",
        )
        mwd = mwd[
            mwd["_Prof"].notna()
            & (mwd["_Prof"] > 0)
        ].copy()

    if mwd.empty:
        st.info(
            "El ciclo seleccionado no tiene barrenos MWD "
            "con profundidad válida."
        )
        return

    mwd = mwd.sort_values(
        [
            c
            for c in [
                "Brazo",
                "Secuencia",
            ]
            if c in mwd.columns
        ]
    ).reset_index(drop=True)

    opciones_barreno = list(mwd.index)

    def _label_barreno(i):
        row = mwd.loc[i]

        barreno_id = (
            str(row.get("Barreno_ID", "-")).strip()
            or "-"
        )
        brazo = row.get("Brazo", "-")
        seq = row.get("Secuencia", "-")

        prof = pd.to_numeric(
            pd.Series(
                [row.get("Profundidad_Max_MWD_m")]
            ),
            errors="coerce",
        ).iloc[0]

        prof_txt = (
            f"{prof:.2f} m"
            if pd.notna(prof)
            else "-"
        )

        estado = str(
            row.get("Estado_MWD", "")
            or ""
        ).strip()

        return (
            f"Brazo {brazo} · "
            f"Secuencia {seq} · "
            f"Barreno {barreno_id} · "
            f"{prof_txt}"
            + (
                f" · {estado}"
                if estado
                else ""
            )
        )

    with col_sel2:
        barreno_idx = st.selectbox(
            "Barreno",
            opciones_barreno,
            key="rop_barreno_sel",
            format_func=_label_barreno,
        )

    row = mwd.loc[barreno_idx]

    barreno_id_sel = (
        str(row.get("Barreno_ID", "-")).strip()
        or "-"
    )

    source_path = Path(
        r.get("_source_path")
        or ""
    )
    archivo_interno = str(
        row.get("Archivo_Interno")
        or ""
    )

    df_rop = _extraer_rop_mwd_desde_zda(
        source_path,
        archivo_interno,
    )

    if df_rop.empty:
        st.warning(
            "No se pudo recuperar la curva ROP del barreno seleccionado. "
            "Verifica que el archivo ZDA fuente siga disponible en la sesión."
        )
        return

    rop_promedio = pd.to_numeric(
        df_rop["ROP_m_min"],
        errors="coerce",
    ).mean()

    rop_mediana = pd.to_numeric(
        df_rop["ROP_m_min"],
        errors="coerce",
    ).median()

    profundidad = pd.to_numeric(
        pd.Series(
            [row.get("Profundidad_Max_MWD_m")]
        ),
        errors="coerce",
    ).iloc[0]

    rep = r.get("resumen_reporte") or {}
    meta = r.get("metadata") or {}

    operador = (
        rep.get("Operador_Filtro")
        or meta.get("Operador")
        or meta.get("Operador_ZDA")
        or "SIN DATO"
    )
    tipo_roca = (
        rep.get("Tipo_Roca")
        or meta.get("Tipo_Roca")
        or "SIN DATO"
    )

    # ------------------------------------------------------
    # Tarjetas superiores
    # ------------------------------------------------------
    st.markdown(
        """
        <style>
        /* Tarjetas del barreno seleccionado: alto fijo e igual entre ambas para que se vean
           como un solo bloque compacto, ajustado a la cantidad real de datos que muestran. */
        .rop-kpi-card, .rop-info-card {
            height: auto; min-height: 162px;
            border-radius: 14px;
            box-sizing: border-box;
        }
        .rop-kpi-card {
            border: 1px solid #dfe7f1;
            background: #ffffff;
            padding: 0.85rem 1.15rem;
            box-shadow: 0 3px 12px rgba(15,23,42,0.045);
        }
        .rop-kpi-label {
            font-size: 0.86rem;
            color: #344054;
            font-weight: 650;
            margin-bottom: 0.30rem;
        }
        .rop-kpi-value {
            font-size: 2.05rem;
            line-height: 1.0;
            color: #111827;
            font-weight: 780;
            margin-bottom: 0.40rem;
        }
        .rop-kpi-unit {
            font-size: 1.05rem;
            font-weight: 500;
            color: #475467;
        }
        .rop-kpi-foot {
            font-size: 0.82rem;
            color: #667085;
        }
        .rop-info-card {
            border: 1px solid #e4e7ec;
            background: #f8fafc;
            padding: 0.70rem 1.05rem;
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            align-content: center;
            gap: 0.55rem 1.0rem;
        }
        .rop-info-label {
            font-size: 0.66rem;
            color: #98a2b3;
            text-transform: uppercase;
            letter-spacing: 0.045em;
        }
        .rop-info-value {
            font-size: 0.87rem;
            color: #1d2939;
            font-weight: 650;
            margin-top: 0.05rem;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }
        .rop-info-wide {
            grid-column: 1 / -1;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    c_kpi, c_info = st.columns(
        [0.70, 1.50]
    )

    with c_kpi:
        prom_txt = (
            f"{rop_promedio:.2f}"
            if pd.notna(rop_promedio)
            else "-"
        )
        med_txt = (
            f"{rop_mediana:.2f}"
            if pd.notna(rop_mediana)
            else "-"
        )

        st.markdown(
            f"""
            <div class="rop-kpi-card">
                <div class="rop-kpi-label">ROP promedio del barreno</div>
                <div class="rop-kpi-value">
                    {prom_txt}
                    <span class="rop-kpi-unit">m/min</span>
                </div>
                <div class="rop-kpi-foot">
                    Mediana: {med_txt} m/min · {len(df_rop)} muestras válidas
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with c_info:
        prof_txt = (
            f"{profundidad:.2f} m"
            if pd.notna(profundidad)
            else "-"
        )

        st.markdown(
            f"""
            <div class="rop-info-card">
                <div>
                    <div class="rop-info-label">Fecha</div>
                    <div class="rop-info-value">{item_sel['Fecha']}</div>
                </div>
                <div>
                    <div class="rop-info-label">Jumbo</div>
                    <div class="rop-info-value">{item_sel['Jumbo']}</div>
                </div>
                <div>
                    <div class="rop-info-label">Ciclo</div>
                    <div class="rop-info-value">{item_sel['Ciclo']}</div>
                </div>
                <div>
                    <div class="rop-info-label">Barrenos</div>
                    <div class="rop-info-value">{item_sel['Barrenos']}</div>
                </div>
                <div>
                    <div class="rop-info-label">Profundidad</div>
                    <div class="rop-info-value">{prof_txt}</div>
                </div>
                <div>
                    <div class="rop-info-label">Barreno</div>
                    <div class="rop-info-value">{barreno_id_sel}</div>
                </div>
                <div>
                    <div class="rop-info-label">Tipo de roca</div>
                    <div class="rop-info-value">{tipo_roca}</div>
                </div>
                <div>
                    <div class="rop-info-label">Brazo / Secuencia</div>
                    <div class="rop-info-value">
                        Brazo {row.get('Brazo','-')} · Secuencia {row.get('Secuencia','-')}
                    </div>
                </div>
                <div>
                    <div class="rop-info-label">Operador</div>
                    <div class="rop-info-value">{operador}</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    # Expander (no container fijo): permite contraer la curva individual para tener el
    # selector más cerca de "Vista de conjunto del round", justo debajo.
    with st.expander("Tasa de penetración ROP (m/min)", expanded=True):
        st.caption(
            "La curva muestra la variación de la tasa de penetración "
            "a lo largo de la progresión del barreno. La línea se suaviza "
            "visualmente con una mediana móvil de 5 muestras; el tooltip "
            "mantiene el valor ROP registrado."
        )

        fig_rop = grafico_rop_por_barreno(
            df_rop,
        )

        if fig_rop is not None:
            st.plotly_chart(
                fig_rop,
                width="stretch",
                config={
                    "displaylogo": False,
                    "scrollZoom": False,
                },
            )

    st.caption(
        "ROP = Rate of Penetration · unidad mostrada: metros por minuto (m/min). "
        "Barreno corresponde al ID del plan ZDA; Brazo + Secuencia identifica "
        "el registro MWD asociado dentro del ZDA."
    )

    # ------------------------------------------------------
    # Vista de conjunto del round: todas las curvas del ciclo a la vez
    # ------------------------------------------------------
    st.divider()
    st.markdown("#### Vista de conjunto del round")
    st.caption(
        "Todos los barrenos del ciclo seleccionado a la vez, en dos formas complementarias: "
        "las curvas ROP superpuestas (para ver la dispersión entre barrenos) y su posición real "
        "en el frente, coloreada por ROP promedio (para ver si algún sector o tipo de barreno "
        "perfora más lento). Usa la posición ejecutada (X, Z) reconstruida del ZDA."
    )

    items = []
    for _, mrow in mwd.iterrows():
        brazo_c = pd.to_numeric(pd.Series([mrow.get("Brazo")]), errors="coerce").iloc[0]
        sec_c = pd.to_numeric(pd.Series([mrow.get("Secuencia")]), errors="coerce").iloc[0]
        if pd.isna(brazo_c) or pd.isna(sec_c):
            continue
        det_fila = None
        if (
            isinstance(detalle_round_rop, pd.DataFrame)
            and not detalle_round_rop.empty
            and {"Boom", "Secuencia", "X", "Y", "Z"}.issubset(detalle_round_rop.columns)
        ):
            match = detalle_round_rop[
                (pd.to_numeric(detalle_round_rop["Boom"], errors="coerce") == int(brazo_c))
                & (pd.to_numeric(detalle_round_rop["Secuencia"], errors="coerce") == int(sec_c))
            ]
            if not match.empty:
                det_fila = match.iloc[0]
        if det_fila is not None:
            barreno_id, tipo = mrow.get("Barreno_ID", "-"), str(det_fila.get("Tipo"))
            x = float(det_fila.get("X")) if pd.notna(det_fila.get("X")) else np.nan
            y_e = float(det_fila.get("Y")) if pd.notna(det_fila.get("Y")) else np.nan
            z = float(det_fila.get("Z")) if pd.notna(det_fila.get("Z")) else np.nan
        else:
            # Sin registro en boom.dat con el mismo Brazo+Secuencia: no es necesariamente un
            # arranque abortado (esos suelen tener <1 m). Puede ser un barreno rehecho, donde
            # boom.dat y los archivos MWD numeran las secuencias de forma distinta (ver el ID
            # del plan y el tipo real en el ZDA). Se identifica por Brazo-Secuencia del MWD y
            # queda fuera del mapa espacial, porque no se conoce su posición ejecutada.
            barreno_id, tipo = f"B{int(brazo_c)}-{int(sec_c):02d}", TIPO_SOSPECHOSO
            x = y_e = z = np.nan
        items.append((str(mrow.get("Archivo_Interno") or ""), str(barreno_id), tipo, x, y_e, z))

    try:
        stat_zda = source_path.stat()
        df_curvas_round, df_resumen_round = _datos_conjunto_rop(
            str(source_path), stat_zda.st_mtime_ns, stat_zda.st_size, tuple(items),
        )
    except OSError:
        df_curvas_round, df_resumen_round = pd.DataFrame(), pd.DataFrame()

    if df_curvas_round.empty:
        st.info(
            "No se pudo reconstruir la vista de conjunto (el ZDA fuente ya no está disponible "
            "en esta sesión, o ningún barreno tiene curva MWD válida)."
        )
    else:
        sospechosos = df_resumen_round[df_resumen_round["Tipo"] == TIPO_SOSPECHOSO]
        confiables = df_resumen_round[df_resumen_round["Tipo"] != TIPO_SOSPECHOSO]

        m1, m2, m3 = st.columns(3)
        m1.metric("Barrenos con curva ROP", int(confiables.shape[0]))
        rop_prom_round = pd.to_numeric(confiables["ROP_prom"], errors="coerce").mean()
        m2.metric("ROP promedio del round", fmt(rop_prom_round, 2, " m/min"))
        con_posicion = int(confiables[["X", "Z"]].dropna().shape[0])
        m3.metric(
            "Con posición ejecutada (X, Z)", f"{con_posicion}/{confiables.shape[0]}",
            help="Barrenos con curva ROP a los que además se les pudo asociar collar ejecutado "
                 "(X, Z) en boom.dat, cruzando por Brazo + Secuencia. Los que faltan suelen ser "
                 "arranques abortados: el MWD registró una curva muy corta (menos de 1 m) y "
                 "boom.dat no tiene un registro de barreno con ese mismo Brazo + Secuencia. Los "
                 "barrenos sospechosos (ver aviso debajo) no se cuentan aquí.",
        )

        if not sospechosos.empty:
            st.markdown(
                f"⚠️ **{len(sospechosos)} barreno(s) sospechoso(s)** excluido(s) de las curvas, "
                "la curva promedio y las métricas de arriba.",
                help=(
                    "Barrenos excluidos: " + ", ".join(sospechosos["Barreno_ID"].astype(str)) + ". "
                    "Su archivo MWD no calza con ningún registro de boom.dat con el mismo "
                    "Brazo + Secuencia, y su curva de profundidad no arranca cerca de 0 m: no hay "
                    "forma de saber si el sensor arrancó en 0 o si es la continuación de un "
                    "barreno rehecho sin reiniciar. Se prefiere no mostrarlos a mostrar una "
                    "profundidad que podría no ser real."
                ),
            )

        st.markdown("##### Curvas ROP superpuestas del round")
        fig_overlay = grafico_rop_round_overlay(df_curvas_round)
        if fig_overlay is not None:
            st.plotly_chart(fig_overlay, width="stretch", config={"displaylogo": False})

        try:
            stat_zda2 = source_path.stat()
            perfil_nominal = _perfil_nominal_rop(str(source_path), stat_zda2.st_mtime_ns, stat_zda2.st_size)
        except OSError:
            perfil_nominal = []

        st.markdown("##### Mapa de ROP en el frente")
        if not perfil_nominal:
            st.caption("No se pudo leer el perfil nominal (round-*.dat) de este ZDA; se muestra solo el frente sin contorno.")
        fig_mapa = grafico_rop_round_mapa(df_resumen_round, perfil_nominal)
        if fig_mapa is not None:
            st.plotly_chart(fig_mapa, width="stretch", config={"displaylogo": False})
        else:
            st.info(
                "No hay barrenos con posición ejecutada (X, Z) conocida para dibujar el mapa "
                "del frente de este round."
            )


def render_resultados_section(resultados_validos):
    st.caption(
        f"{len(resultados_validos)} archivo(s) procesado(s) acumulado(s)"
    )

    # Un expander cerrado igualmente ejecuta/renderiza su contenido en Streamlit.
    # Para no serializar decenas de imágenes/tablas a la vez, se muestran
    # 10 ciclos por vista. La navegación se hace por ciclo, no por tamaño de página.
    n_total = len(resultados_validos)
    page_size = 10
    n_pages = max(1, (n_total + page_size - 1) // page_size)

    # Mantener la página actual dentro del rango válido.
    page_key = "resultados_page_nav"
    if page_key not in st.session_state:
        st.session_state[page_key] = 1
    st.session_state[page_key] = max(
        1,
        min(int(st.session_state[page_key]), n_pages),
    )

    # Selector directo: permite escribir/buscar por jumbo, ciclo o fecha.
    cycle_indices = list(range(n_total))

    def _cycle_label(i):
        rr = resultados_validos[i].get("resumen_reporte", {})
        fuente_i = rr.get("Fuente") or resultados_validos[i].get("fuente") or ""
        return (
            f"{rr.get('Jumbo', '-')} · "
            f"Ciclo {rr.get('Ciclo', '-')} · "
            f"{rr.get('Fecha_Inicio', '-')}"
            + (f" · {fuente_i}" if fuente_i else "")
        )

    jump_key = "resultados_jump_cycle"

    def _jump_to_cycle():
        selected_idx = st.session_state.get(jump_key)
        if selected_idx is not None:
            st.session_state[page_key] = int(selected_idx) // page_size + 1

    st.selectbox(
        "Ir directamente a un ciclo",
        options=cycle_indices,
        index=None,
        placeholder="Buscar por jumbo, ciclo o fecha...",
        format_func=_cycle_label,
        key=jump_key,
        on_change=_jump_to_cycle,
    )

    page = int(st.session_state[page_key])

    nav_prev, nav_info, nav_next = st.columns([1.1, 3.0, 1.1])

    with nav_prev:
        if st.button(
            "← Anterior",
            key="resultados_prev",
            disabled=page <= 1,
            width="stretch",
        ):
            page = max(1, page - 1)
            st.session_state[page_key] = page

    with nav_next:
        if st.button(
            "Siguiente →",
            key="resultados_next",
            disabled=page >= n_pages,
            width="stretch",
        ):
            page = min(n_pages, page + 1)
            st.session_state[page_key] = page

    start = (page - 1) * page_size
    end = min(start + page_size, n_total)

    with nav_info:
        st.markdown(
            f"<div style='text-align:center; padding-top:0.45rem;'>"
            f"<strong>Ciclos {start + 1}–{end} de {n_total}</strong>"
            f"<br><span style='color:#667085; font-size:0.86rem;'>"
            f"Página {page} de {n_pages}</span></div>",
            unsafe_allow_html=True,
        )

    desglosar = st.checkbox(
        "Desglosar todos los ciclos de esta página",
        value=False,
        key="desglosar_todos",
    )

    for idx, r in enumerate(resultados_validos[start:end], start=start):
        rep = r["resumen_reporte"]
        fuente = (
            rep.get("Fuente")
            or r.get("fuente")
            or "ZDA"
        )
        titulo = (
            f"{rep.get('Jumbo','-')} · "
            f"Ciclo {rep.get('Ciclo','-')} · "
            f"{rep.get('Fecha_Inicio','-')} · "
            f"{fuente}"
        )

        with st.expander(
            titulo,
            expanded=desglosar,
        ):
            col_info, col_nav = st.columns([4.7, 1.3])

            with col_info:
                m1, m2, m3, m4, m5, m6, m7 = st.columns(7)

                m1.metric(
                    "Serie",
                    rep.get("Numero_Serie") or "-",
                )
                m2.metric(
                    "Operador",
                    rep.get("Operador_ZDA")
                    or rep.get("Operador")
                    or rep.get("Operario")
                    or "-",
                    help=(
                        rep.get("Fuente_Operador")
                    ),
                )
                m3.metric(
                    "Sección",
                    seccion_desde_plan_texto(
                        rep.get("Plan_Perforacion")
                    ),
                )
                m4.metric(
                    "Tipo de roca",
                    rep.get("Tipo_Roca")
                    or tipo_roca_desde_plan_texto(
                        rep.get("Plan_Perforacion")
                    ),
                )
                m5.metric(
                    "Tipo de disparo",
                    rep.get("Tipo_Disparo") or "-",
                )
                m6.metric(
                    "Barrenos",
                    int(rep["Barrenos_Realizados"])
                    if pd.notna(rep.get("Barrenos_Realizados"))
                    else "-",
                )
                m7.metric(
                    "Metros perforados",
                    fmt(
                        rep.get("Metros_Perforados"),
                        2,
                        " m",
                    ),
                )

                a0, a1, a2, a3, a4 = st.columns(5)
                a0.metric(
                    "Lectura",
                    rep.get("Lectura_Confiable")
                    or rep.get("Estado")
                    or "-",
                )
                a1.metric(
                    "Movimiento automático",
                    fmt(
                        rep.get(
                            "Pct_Movimiento_Automatico_Brazos"
                        ),
                        1,
                        "%",
                    ),
                )
                a2.metric(
                    "Movimiento manual",
                    fmt(
                        rep.get(
                            "Pct_Movimiento_Manual_Brazos"
                        ),
                        1,
                        "%",
                    ),
                )
                a3.metric(
                    "Brazo 1 automático",
                    fmt(
                        rep.get("Pct_Automatico_Brazo1"),
                        1,
                        "%",
                    ),
                )
                a4.metric(
                    "Brazo 2 automático",
                    fmt(
                        rep.get("Pct_Automatico_Brazo2"),
                        1,
                        "%",
                    ),
                )

                # Conteo individual por brazo para este ciclo.
                detalle_round = r.get("detalle")
                b1_count = b2_count = 0
                if (
                    isinstance(detalle_round, pd.DataFrame)
                    and not detalle_round.empty
                    and "Boom" in detalle_round.columns
                ):
                    boom_round = pd.to_numeric(
                        detalle_round["Boom"],
                        errors="coerce",
                    )
                    b1_count = int(boom_round.eq(1).sum())
                    b2_count = int(boom_round.eq(2).sum())

                b1m, b2m, btm = st.columns(3)
                b1m.metric("Barrenos Brazo 1", b1_count)
                b2m.metric("Barrenos Brazo 2", b2_count)
                btm.metric("Total por brazos", b1_count + b2_count)

                z1, z2, z3 = st.columns(3)
                z1.metric(
                    "Inicio perforación real",
                    rep.get("Inicio_Perforacion") or "-",
                )
                z2.metric(
                    "Fin perforación real",
                    rep.get("Fin_Perforacion") or "-",
                )
                z3.metric(
                    "Tiempo de perforación",
                    rep.get("Tiempo_Perforacion_hms") or "-",
                )

            detalle_key = f"detalle_ciclo_{r.get('_cache_key', idx)}"
            mostrar_detalle = st.toggle(
                "Cargar gráficos y plano de este ciclo",
                value=False,
                key=detalle_key,
                help="Los visuales se generan bajo demanda para mantener bajo el uso de memoria.",
            )

            box_path = nav_path = None
            if mostrar_detalle:
                with st.spinner("Generando/cargando visuales del ciclo..."):
                    box_path, nav_path = asegurar_visuales_resultado(r)

            with col_nav:
                if nav_path:
                    st.caption(
                        "Plano reconstruido desde ZDA · "
                        f"sección {seccion_desde_plan_texto(rep.get('Plan_Perforacion'))}"
                    )
                    st.image(str(nav_path), width="stretch")
                elif not mostrar_detalle:
                    st.caption("Plano disponible bajo demanda")

            if box_path:
                st.image(str(box_path), width="stretch")
                with open(box_path, "rb") as fh:
                    png_bytes = fh.read()
                st.download_button(
                    "Descargar gráfico PNG",
                    png_bytes,
                    file_name=(
                        f"{rep.get('Jumbo','JUMBO')}_"
                        f"Ciclo_{rep.get('Ciclo','-')}.png"
                    ),
                    mime="image/png",
                    key=f"png_{idx}",
                )
                del png_bytes

            val = r.get("validacion")
            if (
                isinstance(val, pd.DataFrame)
                and not val.empty
            ):
                with st.expander(
                    "Validación",
                    expanded=False,
                ):
                    cols = [
                        c
                        for c in [
                            "Tipo",
                            "Esperado",
                            "Encontrado",
                            "Diferencia",
                            "Estado",
                        ]
                        if c in val.columns
                    ]
                    st.dataframe(
                        val[cols],
                        width="stretch",
                        hide_index=True,
                        height=220,
                    )

            vm = r.get("validacion_metros")
            if (
                isinstance(vm, pd.DataFrame)
                and not vm.empty
            ):
                with st.expander(
                    "Metros por tipo",
                    expanded=False,
                ):
                    cols = [
                        c
                        for c in [
                            "Tipo",
                            "N",
                            "Metros_Reporte_m",
                            "Metros_Extraidos_m",
                            "Diferencia_m",
                            "Estado",
                        ]
                        if c in vm.columns
                    ]
                    st.dataframe(
                        vm[cols],
                        width="stretch",
                        hide_index=True,
                        height=220,
                    )

            extras = r.get("extras")
            if (
                isinstance(extras, pd.DataFrame)
                and not extras.empty
            ):
                st.markdown(
                    f"#### Barrenos extra ({len(extras)})"
                )
                cols = [
                    c
                    for c in [
                        "ID",
                        "Tipo",
                        "Longitud_roca_m",
                        "Beta_grados",
                    ]
                    if c in extras.columns
                ]
                st.dataframe(
                    extras[cols],
                    width="stretch",
                    hide_index=True,
                    height=180,
                )

            mwd = r.get("mwd_barrenos")
            if (
                isinstance(mwd, pd.DataFrame)
                and not mwd.empty
            ):
                with st.expander(
                    f"MWD por brazo y secuencia ({len(mwd)})",
                    expanded=False,
                ):
                    cols = [
                        c
                        for c in [
                            "Brazo",
                            "Secuencia",
                            "Estado_MWD",
                            "Profundidad_Max_MWD_m",
                            "Muestras_MWD",
                            "Inicio_MWD",
                            "Fin_MWD",
                            "Duracion_MWD_s",
                        ]
                        if c in mwd.columns
                    ]
                    st.dataframe(
                        mwd[cols],
                        width="stretch",
                        hide_index=True,
                        height=260,
                    )


# ==========================================================
# EFICIENCIA DE PERFORACIÓN · GEOMETRÍA ZDA
# ==========================================================


@st.cache_data(show_spinner=False, max_entries=12)
def _cargar_eficiencia_desde_path(path_str, mtime_ns, size_bytes):
    # mtime/size forman parte de la clave del cache para invalidar si cambia el ZDA.
    raw = Path(path_str).read_bytes()
    return procesar_eficiencia_zda(raw, filename=Path(path_str).name)


def _volumen_unico_ciclo(resultado):
    """Única fuente: integral de las mismas secciones que muestran el 2D, 3D y PDF."""
    return volumen_unico_ciclo(resultado)


def _resumenes_eficiencia_sesion(paths):
    """Resumen liviano (DGT, fuera, no cubierto) por ZDA, calculado UNA vez por sesión.

    Antes el selector llamaba al resultado completo de cada ciclo mediante st.cache_data con
    max_entries=12: con más de 12 ciclos el caché expulsaba entradas en cada rerun y se
    reprocesaban todos los ZDA ante cualquier interacción. Aquí se guarda solo el resumen
    (sin secciones ni DataFrames) en session_state, con clave (ruta, fecha, tamaño, versión).
    Los ciclos que faltan se calculan en un solo lote (en paralelo si hay núcleos).
    """
    cache = st.session_state.setdefault("_eficiencia_resumenes", {})
    claves = {}
    for p in paths:
        try:
            stt = Path(p).stat()
            claves[str(p)] = (str(p), stt.st_mtime_ns, stt.st_size, PROCESADOR_EFICIENCIA_VERSION)
        except OSError:
            continue
    faltan = [p for p, k in claves.items() if k not in cache]
    if faltan:
        with st.spinner(f"Calculando eficiencia de {len(faltan)} ciclo(s)..."):
            nuevos = resumenes_eficiencia_lote(faltan)
        for p, resumen in nuevos.items():
            cache[claves[p]] = resumen
    return {p: cache.get(k) for p, k in claves.items()}


@st.cache_data(show_spinner=False, max_entries=8)
def _cached_mask_3d(mask_sections):
    fig_3d = go.Figure()
    valid_3d = True
    for mask_key_3d, color_3d, label_3d in (
        ("programado", "#315FCB", "Programado"),
        ("real", "#00A878", "Real"),
    ):
        rings_3d = []
        for sec_3d in mask_sections:
            boundary_3d = sec_3d[mask_key_3d].get("boundary", [])
            pts_3d = np.asarray(boundary_3d, dtype=float)
            if pts_3d.ndim != 2 or pts_3d.shape[0] < 3 or pts_3d.shape[1] < 2:
                valid_3d = False
                break
            pts_3d = pts_3d[:, :2]
            if np.allclose(pts_3d[0], pts_3d[-1]):
                pts_3d = pts_3d[:-1]
            seg_3d = np.linalg.norm(np.roll(pts_3d, -1, axis=0) - pts_3d, axis=1)
            perimeter_3d = float(seg_3d.sum())
            if perimeter_3d <= 0:
                valid_3d = False
                break
            arc_3d = np.r_[0.0, np.cumsum(seg_3d)]
            closed_3d = np.vstack([pts_3d, pts_3d[0]])
            samples_3d = np.linspace(0.0, perimeter_3d, 80, endpoint=False)
            ring_3d = np.column_stack([
                np.interp(samples_3d, arc_3d, closed_3d[:, axis_3d])
                for axis_3d in (0, 1)
            ])
            # Anclar el inicio en la misma dirección polar para reducir
            # torsiones visuales entre anillos de distinto número de vértices.
            start_3d = int(np.argmax(ring_3d[:, 0]))
            rings_3d.append(np.roll(ring_3d, -start_3d, axis=0))
        if not valid_3d:
            break
        n_ring_3d = len(rings_3d[0])
        xyz_3d = np.asarray([
            (pt[0], float(sec_3d["depth_m"]), pt[1])
            for sec_3d, ring_3d in zip(mask_sections, rings_3d)
            for pt in ring_3d
        ])
        tri_3d = []
        for idx_3d in range(len(rings_3d) - 1):
            for j_3d in range(n_ring_3d):
                a_3d = idx_3d * n_ring_3d + j_3d
                b_3d = idx_3d * n_ring_3d + (j_3d + 1) % n_ring_3d
                c_3d = (idx_3d + 1) * n_ring_3d + j_3d
                d_3d = (idx_3d + 1) * n_ring_3d + (j_3d + 1) % n_ring_3d
                tri_3d.extend([(a_3d, b_3d, c_3d), (b_3d, d_3d, c_3d)])
        tri_3d = np.asarray(tri_3d, dtype=int)
        fig_3d.add_trace(go.Mesh3d(
            x=xyz_3d[:, 0], y=xyz_3d[:, 1], z=xyz_3d[:, 2],
            i=tri_3d[:, 0], j=tri_3d[:, 1], k=tri_3d[:, 2],
            color=color_3d, opacity=0.25, flatshading=True,
            name=f"Modelo {label_3d}", showlegend=True,
            hovertemplate="X: %{x:.2f} m<br>Profundidad: %{y:.2f} m<br>Z: %{z:.2f} m<extra></extra>",
        ))
        for edge_idx_3d, edge_label_3d in (
            (0, f"collar ({float(mask_sections[0]['depth_m']):.2f} m)"),
            (-1, f"sección final ({float(mask_sections[-1]['depth_m']):.2f} m)"),
        ):
            ring_3d = rings_3d[edge_idx_3d]
            closed_ring_3d = np.vstack([ring_3d, ring_3d[0]])
            fig_3d.add_trace(go.Scatter3d(
                x=closed_ring_3d[:, 0],
                y=[float(mask_sections[edge_idx_3d]["depth_m"])] * len(closed_ring_3d),
                z=closed_ring_3d[:, 1], mode="lines",
                line=dict(color=color_3d, width=5 if edge_idx_3d == 0 else 4,
                          dash="solid" if edge_idx_3d == 0 else "dash"),
                name=f"{label_3d} · {edge_label_3d}",
            ))
    return fig_3d, valid_3d


@fragment
def _render_mask_depth_fragment(mask_volume, masks_result, round_id, ciclo_info):
    # Nuevo método en paralelo: máscaras trianguladas por profundidad e integración.
    with st.expander("Integración volumétrica de secciones por profundidad", expanded=True):
        if not mask_volume.get("ok"):
            st.info(mask_volume.get("reason", "No se pudo reconstruir la serie de máscaras."))
        else:
            st.caption("Fuente única de volumetría: triangulación de barrenos disponibles en cada "
                       "profundidad e integración trapezoidal. Los barrenos cortos no se extrapolan. "
                       "La reconstrucción geométrica no mide sobrerotura post-voladura.")
            mask_sections = mask_volume.get("sections", [])
            if mask_sections:
                selected_depth = st.select_slider(
                    "Profundidad de las secciones (m)",
                    options=list(range(len(mask_sections))),
                    value=len(mask_sections) - 1,
                    format_func=lambda idx: f"{mask_sections[idx]['depth_m']:.2f} m",
                    key=f"mask_depth_{round_id}",
                )
                section = mask_sections[selected_depth]
                # Escalas comunes para comparar ambos paneles: el cero de Z debe
                # ocupar exactamente la misma altura en programado y real.
                coords_depth = []
                for mask_key in ("programado", "real"):
                    mask_entry = section[mask_key]
                    coords_depth.extend(mask_entry.get("points", []))
                    coords_depth.extend(mask_entry.get("boundary", []))
                coords_depth.extend(masks_result.get("nominal", []))
                coords_depth = [pt for pt in coords_depth if len(pt) >= 2
                                and np.isfinite(pt[0]) and np.isfinite(pt[1])]
                if coords_depth:
                    shared_x = [min(pt[0] for pt in coords_depth) - 0.4,
                                max(pt[0] for pt in coords_depth) + 0.4]
                    shared_z = [min(pt[1] for pt in coords_depth) - 0.4,
                                max(pt[1] for pt in coords_depth) + 0.4]
                else:
                    shared_x, shared_z = [-3.0, 3.0], [-1.0, 5.5]
                mask_cols_depth = st.columns(2)
                for col_depth, key_depth, color_depth in zip(
                    mask_cols_depth, ("programado", "real"), ("#315FCB", "#00A878")
                ):
                    entry_depth = section[key_depth]
                    fig_depth = go.Figure()
                    mesh_x, mesh_z = [], []
                    for triangle in entry_depth.get("triangles", []):
                        ring = triangle + [triangle[0]]
                        mesh_x.extend([pt[0] for pt in ring] + [None])
                        mesh_z.extend([pt[1] for pt in ring] + [None])
                    if mesh_x:
                        fig_depth.add_trace(go.Scatter(
                            x=mesh_x, y=mesh_z, mode="lines",
                            line=dict(color=color_depth, width=0.7),
                            opacity=0.35, showlegend=False, hoverinfo="skip"))
                    outline_depth = entry_depth.get("boundary", [])
                    if outline_depth:
                        fig_depth.add_trace(go.Scatter(
                            x=[pt[0] for pt in outline_depth], y=[pt[1] for pt in outline_depth],
                            mode="lines", line=dict(color=color_depth, width=3),
                            name="Contorno reconstruido"))
                    # Referencia fija de diseño: solo visual, no recorta máscaras ni altera volúmenes.
                    nominal_depth = masks_result.get("nominal", [])
                    if nominal_depth:
                        fig_depth.add_trace(go.Scatter(
                            x=[pt[0] for pt in nominal_depth],
                            y=[pt[1] for pt in nominal_depth],
                            mode="lines",
                            line=dict(color="#64748B", dash="dash", width=1.5),
                            name="Perfil nominal", opacity=0.8,
                            hovertemplate="Perfil nominal<extra></extra>",
                        ))
                    pts_depth = entry_depth.get("points", [])
                    if pts_depth:
                        fig_depth.add_trace(go.Scatter(
                            x=[pt[0] for pt in pts_depth], y=[pt[1] for pt in pts_depth],
                            mode="markers", marker=dict(color=color_depth, size=5),
                            text=entry_depth.get("ids", []),
                            hovertemplate="Hole ID: %{text}<extra></extra>", name="Barrenos"))
                    fig_depth.update_layout(
                        template="plotly_white", height=460,
                        title=f"Sección {'programada' if key_depth == 'programado' else 'ejecutada'} a {section['depth_m']:.2f} m",
                        xaxis=dict(title="X (m)", range=shared_x, constrain="domain"),
                        yaxis=dict(title="Z (m)", range=shared_z, scaleanchor="x",
                                   scaleratio=1, constrain="domain"),
                        margin=dict(l=25, r=15, t=48, b=30),
                        legend=dict(orientation="h", y=1.01, yanchor="bottom"))
                    with col_depth:
                        st.plotly_chart(fig_depth, width="stretch",
                                        key=f"mask_section_{key_depth}_{round_id}")
                        area_depth = entry_depth.get("area_m2")
                        n_bar_depth = entry_depth.get("n_barrenos")
                        n_tot_depth = entry_depth.get("n_total")
                        if area_depth is None:
                            st.caption("Sección sin geometría válida: no se extrapola.")
                        else:
                            if n_bar_depth is None:
                                detalle_depth = f"{len(pts_depth)} puntos"
                            elif key_depth == "programado":
                                detalle_depth = (f"{n_bar_depth} barrenos · {len(pts_depth)} posiciones únicas "
                                                 "(IDs coincidentes juntos en el hover)")
                            else:
                                detalle_depth = (f"{n_bar_depth} de {n_tot_depth} barrenos alcanzan esta profundidad")
                                if n_tot_depth and n_bar_depth < n_tot_depth:
                                    detalle_depth += " · los más cortos no se extrapolan"
                                detalle_depth += f" · {len(pts_depth)} posiciones únicas"
                            st.caption(f"{detalle_depth} · Área: {area_depth:.2f} m²")
            if mask_volume.get("complete"):
                vol_base = float(mask_volume["programado_m3"])
                vol_fuera = float(mask_volume["outside_m3"])
                vol_no_cubierto = float(mask_volume["not_covered_m3"])
                pct_fuera = 100 * vol_fuera / vol_base if vol_base > 0 else None
                pct_no_cubierto = 100 * vol_no_cubierto / vol_base if vol_base > 0 else None
                pct_dgt = pct_fuera + pct_no_cubierto if pct_fuera is not None else None
                cols_mask_metrics = st.columns(4)
                cols_mask_metrics[0].metric("Volumen programado", f"{mask_volume['programado_m3']:,.2f} m³")
                cols_mask_metrics[1].metric("Volumen ejecutado", f"{mask_volume['real_m3']:,.2f} m³")
                cols_mask_metrics[2].metric(
                    "Fuera del programado",
                    f"{vol_fuera:,.2f} m³",
                    f"{pct_fuera:.2f}%" if pct_fuera is not None else None,
                    delta_color="off",
                )
                cols_mask_metrics[3].metric(
                    "No cubierto",
                    f"{vol_no_cubierto:,.2f} m³",
                    f"{pct_no_cubierto:.2f}%" if pct_no_cubierto is not None else None,
                    delta_color="off",
                )
                dgt_help = (
                    "FÓRMULAS\n"
                    "Fuera (%) = Volumen fuera del programado / Volumen programado × 100\n"
                    "No cubierto (%) = Volumen no cubierto / Volumen programado × 100\n"
                    "DGT (m³) = Volumen fuera del programado + Volumen no cubierto\n"
                    "DGT (%) = DGT (m³) / Volumen programado × 100\n\n"
                    "VALORES DEL CICLO SELECCIONADO\n"
                    f"Volumen programado: {vol_base:,.2f} m³\n"
                    f"Fuera del programado: {vol_fuera:,.2f} m³"
                    + (f" ({pct_fuera:.2f}%)\n" if pct_fuera is not None else "\n")
                    + f"No cubierto: {vol_no_cubierto:,.2f} m³"
                    + (f" ({pct_no_cubierto:.2f}%)\n" if pct_no_cubierto is not None else "\n")
                    + f"DGT = {vol_fuera:,.2f} + {vol_no_cubierto:,.2f} = {vol_fuera + vol_no_cubierto:,.2f} m³\n"
                    + (f"DGT (%) = ({vol_fuera + vol_no_cubierto:,.2f} / {vol_base:,.2f}) × 100 = {pct_dgt:.2f}%\n\n"
                       if pct_dgt is not None else "DGT (%): no disponible; volumen programado no válido.\n\n")
                    + "El cálculo usa los valores originales sin redondear. "
                    "Es una comparación geométrica estimada de la perforación; "
                    "no mide sobrerotura posterior a la voladura."
                )
                st.metric("Desviación geométrica total",
                          f"{vol_fuera + vol_no_cubierto:,.2f} m³",
                          f"{pct_dgt:.2f}% del programado" if pct_dgt is not None else None,
                          delta_color="off", help=dgt_help)
                # Variación porcentual respecto al volumen programado (mismo método).
                v_plan = float(mask_volume["programado_m3"])
                v_real = float(mask_volume["real_m3"])
                variation_pct = 100.0 * (v_real - v_plan) / v_plan if v_plan > 0 else None
                # El tooltip usa valores originales para el resultado y valores visibles
                # redondeados solo para mostrar la sustitución de la fórmula.
                variation_help = (
                    "FÓRMULA\n"
                    "Variación (%) = (Volumen ejecutado − Volumen programado) "
                    "/ Volumen programado × 100\n\n"
                    "VALORES DEL CICLO SELECCIONADO\n"
                    f"= ({v_real:,.2f} − {v_plan:,.2f}) / {v_plan:,.2f} × 100\n"
                )
                if variation_pct is not None:
                    variation_help += (
                        f"= {variation_pct:+.2f}%\n\n"
                        "El resultado se calcula con los volúmenes originales, sin redondear; "
                        "por eso puede diferir ligeramente del cálculo con los valores visibles."
                    )
                else:
                    variation_help += "Resultado no disponible: el volumen programado debe ser mayor que cero."
                variation_help += (
                    "\n\nComparación geométrica exploratoria de perforación; "
                    "no es una medición de sobrerotura posterior a la voladura."
                )
                st.metric("Variación de volumen (ejecutado vs. programado)",
                          f"{variation_pct:+.2f}%" if variation_pct is not None else "N/D",
                          help=variation_help)

                # Vista 3D construida con las MISMAS secciones utilizadas en la integral.
                # Cada superficie lateral interpola la frontera de secciones consecutivas;
                # no modifica el algoritmo ni los resultados volumétricos.
                with st.expander("Modelo volumétrico 3D · secciones trianguladas y reporte", expanded=True):
                    left_3d, right_report = st.columns([1.65, 1], gap="large")
                    fig_3d, valid_3d = _cached_mask_3d(mask_sections)
                    with left_3d:
                        if valid_3d and len(mask_sections) >= 2:
                            fig_3d.update_layout(
                                template="plotly_white", height=620,
                                scene=dict(xaxis_title="X (m)", yaxis_title="Profundidad (m)",
                                           zaxis_title="Z (m)", aspectmode="data"),
                                margin=dict(l=0, r=0, t=15, b=0),
                                legend=dict(orientation="h", y=-0.06),
                            )
                            st.plotly_chart(fig_3d, width="stretch",
                                            key=f"mask_3d_{round_id}")
                        else:
                            st.info("No hay contornos válidos en todas las profundidades para la vista 3D.")
                    with right_report:
                        st.markdown("#### Área de sección (m²)")
                        sec_first, sec_last = mask_sections[0], mask_sections[-1]
                        rows_area = []
                        for sec_label, sec_item in (("Collar", sec_first), ("Fondo", sec_last)):
                            ap = sec_item["programado"].get("area_m2")
                            ar = sec_item["real"].get("area_m2")
                            rows_area.append({"Sección": f"{sec_label} ({sec_item['depth_m']:.2f} m)",
                                              "Programado": round(ap, 2) if ap is not None else None,
                                              "Real": round(ar, 2) if ar is not None else None,
                                              "Diferencia": round(ar-ap, 2) if ap is not None and ar is not None else None,
                                              "Variación %": f"{100*(ar-ap)/ap:+.1f}%" if ap and ar is not None else "N/D"})
                        st.dataframe(pd.DataFrame(rows_area), hide_index=True, width="stretch")
                        st.markdown("#### Diferencias espaciales (m³)")
                        spatial_rows = [
                            {"Indicador": "Fuera del programado", "Volumen (m³)": f"{vol_fuera:,.2f}",
                             "% del programado": f"{pct_fuera:.2f}%" if pct_fuera is not None else "N/D"},
                            {"Indicador": "No cubierto", "Volumen (m³)": f"{vol_no_cubierto:,.2f}",
                             "% del programado": f"{pct_no_cubierto:.2f}%" if pct_no_cubierto is not None else "N/D"},
                            {"Indicador": "Desviación geométrica total", "Volumen (m³)": f"{vol_fuera + vol_no_cubierto:,.2f}",
                             "% del programado": f"{pct_dgt:.2f}%" if pct_dgt is not None else "N/D"},
                        ]
                        st.dataframe(
                            pd.DataFrame(spatial_rows).style.apply(
                                lambda row: ["font-weight: bold" if row.name == 2 else "" for _ in row], axis=1
                            ), hide_index=True, width="stretch",
                            column_config={
                                "Indicador": st.column_config.TextColumn(width=220),
                                "Volumen (m³)": st.column_config.TextColumn(width=125),
                                "% del programado": st.column_config.TextColumn(width=155),
                            },
                        )
                        st.markdown("#### Volumen integrado (m³)")
                        st.dataframe(pd.DataFrame([{
                            "Sección": f"0–{mask_sections[-1]['depth_m']:.2f} m",
                            "Programado": round(v_plan, 2), "Real": round(v_real, 2),
                            "Diferencia": round(v_real-v_plan, 2),
                            "Variación %": f"{variation_pct:+.1f}%" if variation_pct is not None else "N/D",
                        }]), hide_index=True, width="stretch")
                    # Generación bajo demanda: no se exportan imágenes al mover la profundidad.
                    st.markdown("#### Reporte técnico del ciclo")
                    report_key = f"pdf_perforacion_{round_id}"
                    if st.button("Generar reporte PDF", key=f"generar_pdf_{round_id}"):
                        if not (valid_3d and len(mask_sections) >= 2):
                            st.error("No se puede generar el reporte: falta un modelo 3D válido.")
                        else:
                            with st.spinner("Exportando los gráficos del ciclo y componiendo el PDF..."):
                                try:
                                    st.session_state[report_key] = generar_reporte_pdf(
                                        mask_volume, masks_result, fig_3d, round_id,
                                        fecha=ciclo_info.get("fecha"),
                                        equipo=ciclo_info.get("equipo"),
                                        operador=ciclo_info.get("operador"),
                                    )
                                except Exception as exc:
                                    st.session_state.pop(report_key, None)
                                    st.error(f"No se pudo generar el PDF: {exc}")
                    if st.session_state.get(report_key):
                        st.download_button(
                            "Descargar reporte PDF",
                            data=st.session_state[report_key],
                            file_name=f"Reporte_Eficiencia_Perforacion_Ciclo_{round_id}.pdf",
                            mime="application/pdf", key=f"descargar_pdf_{round_id}",
                        )
                        st.info("Vista 3D exploratoria: superficies interpoladas entre secciones. "
                                "Las áreas y volúmenes provienen de la integración de secciones, "
                                "no de la malla visual ni de una medición post-voladura.")
            else:
                st.warning("No se calculan volúmenes: existen secciones sin geometría válida. "
                           "No se interpolan áreas a través de vacíos de datos.")
            st.warning(mask_volume.get("warning", "Resultados exploratorios: no sustituyen los indicadores vigentes."))


_CSS_SELECTBOX_MARCA = """
<style>
/* ---- Campo del selector ------------------------------------------------------------
   Streamlit reciente dibuja st.selectbox con React Aria (ya NO con BaseWeb): el campo
   es un div con role="group" y fondo del tema; el texto es un input con role="combobox"
   y la flecha un button con aria-label="Open". Los selectores de abajo usan esos
   atributos (estables) y no las clases emotion generadas. */
__C__ [data-testid="stSelectbox"] [role="group"] {
    background-color: __FONDO__ !important;
    border: 1px solid __BORDE__ !important;
    border-radius: 8px !important;
    box-shadow: 0 1px 3px rgba(15, 23, 42, 0.25);
    transition: background-color 120ms ease, box-shadow 120ms ease;
}
__C__ [data-testid="stSelectbox"] [role="group"]:hover {
    background-color: __FONDO_HOVER__ !important;
}
__C__ [data-testid="stSelectbox"] [role="group"][data-focus-within] {
    border-color: __BORDE_FOCO__ !important;
    box-shadow: 0 0 0 3px __ANILLO__ !important;
}
__C__ [data-testid="stSelectbox"] input {
    background: transparent !important;
    color: __TEXTO__ !important;
    -webkit-text-fill-color: __TEXTO__ !important;
    caret-color: __TEXTO__ !important;
}
__C__ [data-testid="stSelectbox"] [role="group"] button,
__C__ [data-testid="stSelectbox"] [role="group"] svg {
    color: __TEXTO__ !important;
    fill: __TEXTO__ !important;
}
/* Streamlit anterior (BaseWeb): mismo aspecto. */
__C__ [data-testid="stSelectbox"] [data-baseweb="select"] > div {
    background-color: __FONDO__ !important;
    border: 1px solid __BORDE__ !important;
    border-radius: 8px !important;
}
__C__ [data-testid="stSelectbox"] [data-baseweb="select"] * {
    color: __TEXTO__ !important;
    -webkit-text-fill-color: __TEXTO__ !important;
}
/* ---- Lista desplegable ---------------------------------------------------------------
   Se monta fuera del contenedor (portal en <body>): se acota con :has() para que solo
   cambie mientras ESTE selector está abierto y no afecte a los demás. */
body:has(__C__ [role="combobox"][aria-expanded="true"]) [data-testid="stSelectboxVirtualDropdown"] {
    border: 1px solid __BORDE__ !important;
}
body:has(__C__ [role="combobox"][aria-expanded="true"]) [role="option"][data-selected="true"] [data-item-hl],
body:has(__C__ [role="combobox"][aria-expanded="true"]) [data-baseweb="popover"] [role="option"][aria-selected="true"] {
    background-color: __FONDO__ !important;
}
body:has(__C__ [role="combobox"][aria-expanded="true"]) [role="option"][data-hovered] [data-item-hl],
body:has(__C__ [role="combobox"][aria-expanded="true"]) [role="option"][data-focused] [data-item-hl],
body:has(__C__ [role="combobox"][aria-expanded="true"]) [data-baseweb="popover"] [role="option"]:hover {
    background-color: __RESALTE__ !important;
}
body:has(__C__ [role="combobox"][aria-expanded="true"]) [role="option"][data-selected="true"],
body:has(__C__ [role="combobox"][aria-expanded="true"]) [role="option"][data-hovered],
body:has(__C__ [role="combobox"][aria-expanded="true"]) [role="option"][data-focused],
body:has(__C__ [role="combobox"][aria-expanded="true"]) [data-baseweb="popover"] [role="option"][aria-selected="true"] *,
body:has(__C__ [role="combobox"][aria-expanded="true"]) [data-baseweb="popover"] [role="option"]:hover * {
    color: __TEXTO__ !important;
    -webkit-text-fill-color: __TEXTO__ !important;
}
</style>
"""


def _css_selectbox_marca(
    clase: str,
    fondo: str = "#125B37",
    fondo_hover: str = "#0F4A2D",
    texto: str = "#FFFFFF",
    borde: str = "#008F49",
    borde_foco: str = "#4CC38A",
    anillo: str = "rgba(0, 143, 73, 0.35)",
    resalte: str = "#008F49",
) -> str:
    """CSS con los colores de marca para un st.selectbox dentro del contenedor `clase`.

    `clase` es la clase que Streamlit genera para un widget/contenedor con key, p. ej.
    ".st-key-eficiencia_selector_destacado". Sirve para cualquier otro selector: basta
    con envolverlo en st.container(key="...") y llamar a esta función con esa clase.
    """
    css = _CSS_SELECTBOX_MARCA
    for marcador, valor in (
        ("__C__", clase), ("__FONDO_HOVER__", fondo_hover), ("__FONDO__", fondo),
        ("__TEXTO__", texto), ("__BORDE_FOCO__", borde_foco), ("__BORDE__", borde),
        ("__ANILLO__", anillo), ("__RESALTE__", resalte),
    ):
        css = css.replace(marcador, valor)
    return css


def render_eficiencia_perforacion_section(resultados, sel_jumbos, sel_tipos, sel_rocas, sel_operadores, sel_turnos=None):
    st.subheader("Eficiencia de Perforación")
    st.markdown("**Evaluación del cumplimiento y desviación geométrica de la perforación**")
    st.markdown(
        "**Objetivo:** Comparar la geometría programada y ejecutada de los barrenos "
        "mediante la reconstrucción de secciones transversales y modelos volumétricos, "
        "cuantificando las desviaciones respecto al diseño mediante integración numérica."
    )
    st.markdown(
        "- **DGT (Desviación Geométrica Total):** (volumen fuera del programado + volumen no cubierto) "
        "÷ volumen programado × 100. Compara geometrías de perforación; no mide la sobrerotura después "
        "de la voladura. Su detalle está en el cuadro resumen, en **Diferencias espaciales (m³)**.\n"
        "- **Variación de volumen:** (volumen ejecutado − volumen programado) ÷ volumen programado × 100. "
        "Se calcula más abajo integrando las secciones y aparece en el cuadro resumen, en **Volumen integrado (m³)**."
    )
    st.caption(
        "Estas desviaciones podrían contribuir a la sobrerotura o subexcavación; "
        "los resultados no constituyen una medición directa del perfil excavado después de la voladura."
    )

    # Calcular una sola vez por ZDA (cacheado por ruta, fecha y tamaño).
    # El porcentaje replica exactamente la fórmula de la tarjeta del ciclo.
    disponibles = []
    resumenes_eficiencia = _resumenes_eficiencia_sesion([
        r.get("_source_path") for r in resultados
        if not r.get("error") and r.get("_source_path") and Path(r["_source_path"]).exists()
    ])
    for r in resultados:
        path_str = r.get("_source_path")
        if r.get("error") or not path_str or not Path(path_str).exists():
            continue
        rep = r.get("resumen_reporte") or {}
        # Aplicar los mismos filtros del sidebar ANTES de construir el selector.
        # La clasificación se basa en el mismo conteo de barrenos realizados
        # que utiliza el reporte consolidado.
        n_clasificacion = rep.get("Barrenos_Realizados")
        if n_clasificacion is None or pd.isna(n_clasificacion):
            continue
        if clasificar_tipo_disparo_v33(n_clasificacion) not in sel_tipos:
            continue
        if str(rep.get("Jumbo")) not in {str(x) for x in sel_jumbos}:
            continue
        if rep.get("Tipo_Roca") not in sel_rocas or rep.get("Operador_Filtro") not in sel_operadores:
            continue
        if not turno_permitido(rep.get("Hora_Inicio"), sel_turnos):
            continue
        try:
            fecha_ciclo = datetime.strptime(str(rep.get("Fecha_Inicio")), "%d/%m/%Y").date()
            desde = st.session_state.get("fecha_inicio_zda_global")
            hasta = st.session_state.get("fecha_fin_zda_global")
            if (desde is not None and fecha_ciclo < desde) or (hasta is not None and fecha_ciclo > hasta):
                continue
        except (TypeError, ValueError):
            pass
        det_lista = r.get("detalle")
        n_barrenos_lista = rep.get("N_Barrenos") or rep.get("Barrenos") or rep.get("Barrenos_Realizados")
        if n_barrenos_lista is None and isinstance(det_lista, pd.DataFrame) and not det_lista.empty:
            n_barrenos_lista = len(det_lista)
        barrenos_txt = (
            f" · {int(n_barrenos_lista)} barrenos"
            if n_barrenos_lista is not None and pd.notna(n_barrenos_lista) else ""
        )
        # El selector usa la MISMA integral de secciones del análisis y PDF
        # (resumen calculado una vez por sesión; un ZDA defectuoso queda "No calculada").
        volumen_ciclo = resumenes_eficiencia.get(str(path_str))
        desviacion_pct = volumen_ciclo["dgt_pct"] if volumen_ciclo else np.nan
        fuera_pct = volumen_ciclo["fuera_pct"] if volumen_ciclo else np.nan
        no_cubierto_pct = volumen_ciclo["no_cubierto_pct"] if volumen_ciclo else np.nan
        desviacion_txt = f"{desviacion_pct:.1f}%" if np.isfinite(desviacion_pct) else "No calculada"
        fuera_txt = f"{fuera_pct:.1f}%" if np.isfinite(fuera_pct) else "N/D"
        no_cubierto_txt = f"{no_cubierto_pct:.1f}%" if np.isfinite(no_cubierto_pct) else "N/D"
        etiqueta = (
            f"{rep.get('Jumbo') or '-'} · Ciclo {rep.get('Ciclo') or '-'} · "
            f"{rep.get('Fecha_Inicio') or '-'}{barrenos_txt} · "
            f"{r.get('nombre_archivo') or Path(path_str).name} · DGT: {desviacion_txt} · Fuera: {fuera_txt} · No cubierto: {no_cubierto_txt}"
        )
        disponibles.append({
            "id": str(Path(path_str).resolve()), "etiqueta": etiqueta,
            "resultado": r, "desviacion": desviacion_pct,
            "fuera_pct": fuera_pct, "no_cubierto_pct": no_cubierto_pct,
            "fecha": str(rep.get("Fecha_Inicio") or ""),
            "ciclo": str(rep.get("Ciclo") or ""),
        })

    if not disponibles:
        st.info("No hay ciclos ZDA que cumplan los filtros seleccionados.")
        return

    # Encabezado del selector. Los <style> son elementos vacíos que Streamlit separa con su hueco
    # de 16 px: se ocultan solo aquí (:has) para que "Ciclo a analizar" y "Ordenar ciclos por"
    # queden juntos.
    with st.container(key="eficiencia_cabecera"):
        st.markdown("**Ciclo a analizar**")
        # Resaltar únicamente el selector de ciclos.
        st.markdown("""
        <style>
        .st-key-eficiencia_cabecera [data-testid="stElementContainer"]:has(style) { display: none; }
        .st-key-eficiencia_cabecera { gap: 0.5rem; }
        .st-key-eficiencia_selector_destacado {
            background: #E6F4EC;
            border: 1px solid #008F49;
            border-left: 5px solid #008F49;
            border-radius: 10px;
            padding: 12px 16px 14px;
            margin: 8px 0 16px;
        }
        .st-key-eficiencia_selector_destacado [data-testid="stSelectbox"] label p {
            color: #125B37;
            font-weight: 700;
        }
        </style>
        """, unsafe_allow_html=True)
        st.markdown(
            _css_selectbox_marca(".st-key-eficiencia_selector_destacado"),
            unsafe_allow_html=True,
        )
        orden = st.radio(
            "Ordenar ciclos por",
            ["Mayor desviación geométrica total", "Mayor fuera del programado",
             "Mayor no cubierto", "Fecha / ciclo"],
            index=0,
            key="eficiencia_orden_ciclos_dgt",
        )
    def clave_fecha(item):
        try:
            fecha = datetime.strptime(item["fecha"], "%d/%m/%Y")
        except (ValueError, TypeError):
            fecha = datetime.min
        try:
            ciclo = int(item["ciclo"])
        except (ValueError, TypeError):
            ciclo = -1
        return fecha, ciclo

    # Las métricas no disponibles se colocan al final sin atribuirles cero.
    if orden == "Fecha / ciclo":
        disponibles.sort(key=clave_fecha, reverse=True)
    else:
        campo = {
            "Mayor desviación geométrica total": "desviacion",
            "Mayor fuera del programado": "fuera_pct",
            "Mayor no cubierto": "no_cubierto_pct",
        }[orden]
        validos = [item for item in disponibles if np.isfinite(item[campo])]
        invalidos = [item for item in disponibles if not np.isfinite(item[campo])]
        validos.sort(key=lambda item: item[campo], reverse=True)
        invalidos.sort(key=clave_fecha, reverse=True)
        disponibles = validos + invalidos

    por_id = {item["id"]: item for item in disponibles}
    ids = [item["id"] for item in disponibles]
    if st.session_state.get("eficiencia_ciclo_id") not in por_id:
        st.session_state["eficiencia_ciclo_id"] = ids[0]
    with st.container(key="eficiencia_selector_destacado"):
        seleccion_id = st.selectbox(
            "Seleccionar ciclo", ids,
            format_func=lambda identificador: por_id[identificador]["etiqueta"],
            key="eficiencia_ciclo_id",
        )
    rsel = por_id[seleccion_id]["resultado"]
    path = Path(rsel["_source_path"])

    try:
        stat = path.stat()
        with st.spinner("Reconstruyendo geometría programada y real desde el ZDA..."):
            rr = _cargar_eficiencia_desde_path(
                str(path), stat.st_mtime_ns, stat.st_size
            )
    except Exception as exc:
        st.error(f"No se pudo reconstruir la geometría del ciclo: {exc}")
        return

    v = rr.get("volumetry") or {}
    df = rr.get("holes")
    meta = rr.get("metadata") or {}

    if not isinstance(df, pd.DataFrame) or df.empty:
        st.warning("El ZDA no contiene barrenos válidos para reconstruir la geometría.")
        return

    # Conciliación documentada: todos los registros y sus coordenadas originales.
    aud = rr.get("auditoria_conteos") or {}
    with st.expander("Auditoría ZDA · barrenos programados y ejecutados", expanded=False):
        if aud.get("conciliado"):
            st.success(
                f"Conciliado: {aud['programados_extraidos']} programados y "
                f"{aud['ejecutados_extraidos']} ejecutados recuperados del ZDA."
            )
        else:
            st.warning("El número de registros recuperados no coincide con lo declarado en round.txt.")
        columnas_aud = ["ID", "Tipo", "Estado_Barreno", "Plan_X", "Plan_Y", "Plan_Z",
                        "Plan_X2", "Plan_Y2", "Plan_Z2", "X", "Y", "Z",
                        "X2", "Y2", "Z2", "Offset_Boom", "Tamano_Registro_Boom"]
        aud_df = df[[c for c in columnas_aud if c in df.columns]].copy()
        st.dataframe(aud_df, hide_index=True, width="stretch")
        st.download_button(
            "Descargar auditoría de coordenadas CSV",
            data=aud_df.to_csv(index=False).encode("utf-8-sig"),
            file_name=f"Auditoria_ZDA_Ciclo_{meta.get('round', 'x')}.csv",
            mime="text/csv", key=f"descargar_auditoria_zda_{meta.get('round', 'x')}",
        )
        st.caption("Los barrenos no perforados conservan su collar y fondo programados; "
                   "sus coordenadas ejecutadas permanecen vacías. Los extras no tienen plan.")

    # --------------------------------------------------------------
    # Conteos ZDA
    # --------------------------------------------------------------
    n_prog = meta.get("planned_face_holes")
    n_real = meta.get("drilled_holes")
    if n_prog is None:
        n_prog = int(df["Plan_X"].notna().sum()) if "Plan_X" in df.columns else 0
    if n_real is None:
        n_real = int(len(df))

    # ÚNICA volumetría: secciones transversales variables e integración trapezoidal.
    # No se mezclan volúmenes del método de perímetro con máscaras trianguladas.
    m = rr.get("mask_volumetry") or {}
    unico = _volumen_unico_ciclo(rr)
    L = float(m.get("depth_m", np.nan)) if m.get("ok") else np.nan
    Vprog = unico["programado_m3"] if unico else np.nan
    Vreal = unico["real_m3"] if unico else np.nan
    per = v.get("perimeter_rows")
    if not isinstance(per, pd.DataFrame):
        per = pd.DataFrame()
    volumetria_ok = bool(unico)
    if not unico:
        st.error("Volumetría única no disponible: faltan secciones válidas o falla la identidad "
                 "volumen ejecutado − programado = fuera − no cubierto. "
                 "No se mostrarán volúmenes obtenidos con otro método.")

    # --------------------------------------------------------------
    # Datos del ciclo seleccionado
    # --------------------------------------------------------------
    rep_sel = rsel.get("resumen_reporte") or {}
    operador_sel = rep_sel.get("Operador_ZDA") or rep_sel.get("Operador") or rep_sel.get("Operario")
    seccion_sel = seccion_desde_plan_texto(rep_sel.get("Plan_Perforacion") or meta.get("drill_plan")).replace(" ", "")
    datos_ciclo = [
        ("Fecha", rep_sel.get("Fecha_Inicio") or meta.get("Fecha_Inicio")),
        ("Equipo", rep_sel.get("Jumbo") or meta.get("Jumbo")),
        ("Operador", operador_sel if not asig.es_sin_operador(operador_sel) else "Sin registrar"),
        ("Sección", seccion_sel if seccion_sel not in ("", "-") else None),     # solo si existe
    ]
    st.markdown(" | ".join(f"**{k}:** {v}" for k, v in datos_ciclo if v))

    # --------------------------------------------------------------
    # Tarjetas
    # --------------------------------------------------------------
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Barrenos programados", f"{int(n_prog)}" if n_prog is not None else "-")
    c2.metric("Barrenos realizados", f"{int(n_real)}" if n_real is not None else "-")
    c3.metric("Longitud de perforación alcanzada", f"{L:.2f} m" if np.isfinite(L) else "-")
    c4.metric("Volumen programado", f"{Vprog:.2f} m³" if np.isfinite(Vprog) else "-")
    c5.metric("Volumen real barrenado", f"{Vreal:.2f} m³" if np.isfinite(Vreal) else "-")

    fuera_prog = unico["outside_m3"] if unico else np.nan
    no_cubierto = unico["not_covered_m3"] if unico else np.nan
    d1, d2, d3 = st.columns(3)
    espacial_ok = bool(unico)
    fuera_pct = unico["fuera_pct"] if unico else np.nan
    no_cubierto_pct = unico["no_cubierto_pct"] if unico else np.nan
    dgt_m3 = unico["dgt_m3"] if unico else np.nan
    dgt_pct = unico["dgt_pct"] if unico else np.nan

    d1.metric(
        "Fuera del programado",
        f"{fuera_prog:.2f} m³" if np.isfinite(fuera_prog) else "No calculado",
        f"{fuera_pct:.2f}% del programado" if np.isfinite(fuera_pct) else None,
        delta_color="off",
    )
    d2.metric(
        "No cubierto respecto al programado",
        f"{no_cubierto:.2f} m³" if np.isfinite(no_cubierto) else "No calculado",
        f"{no_cubierto_pct:.2f}% del programado" if np.isfinite(no_cubierto_pct) else None,
        delta_color="off",
    )
    d3.metric(
        "Desviación geométrica total",
        f"{dgt_m3:.2f} m³" if np.isfinite(dgt_m3) else "No calculada",
        f"{dgt_pct:.2f}% del programado" if np.isfinite(dgt_pct) else None,
        delta_color="off",
        help="DGT = fuera del programado + no cubierto. Los porcentajes usan el volumen programado como denominador.",
    )

    if not espacial_ok:
        st.warning(
            "La comparación espacial Fuera del programado / No cubierto no pudo calcularse. "
            "Verifique que Shapely esté instalado (se incluye en requirements.txt)."
        )

    st.caption(
        "Los volúmenes provienen de integrar las áreas de las mismas secciones transversales "
        "a lo largo de la profundidad de referencia."
    )
    if not volumetria_ok or per.empty or not np.isfinite(L):
        motivo = v.get("reason") or "No se pudo construir un perímetro volumétrico confiable para este ciclo."
        st.info(
            f"Comparación volumétrica no disponible para este ciclo: {motivo} "
            "El modelo 3D de los barrenos disponibles se muestra igualmente."
        )

    # Vista de collar; la integración oficial utiliza las secciones por profundidad.
    with st.expander("Secciones transversales trianguladas (plan vs. ejecutado)", expanded=True):
        masks_result = rr.get("experimental_masks") or {}
        if not masks_result.get("ok"):
            st.info(masks_result.get("reason", "Máscaras no disponibles."))
        else:
            st.caption("Reconstrucción exploratoria con todos los collares disponibles. "
                       "La sección programada conserva el contorno de sus puntos; la ejecutada descarta "
                       "triángulos con lados mayores de 2 m. Las áreas mostradas son solo "
                       "diagnósticas: NO se usan para calcular volúmenes ni certifican el contorno.")
            mask_cols = st.columns(2)
            for mask_col, mask_key, mask_color, mask_title in zip(
                mask_cols, ("programado", "real"), ("#315FCB", "#00A878"),
                ("Sección programada · collar", "Sección ejecutada · collar")
            ):
                entry = masks_result.get("masks", {}).get(mask_key, {})
                mesh_fig = go.Figure()
                for triangle in entry.get("triangles", []):
                    ring = triangle + [triangle[0]]
                    mesh_fig.add_trace(go.Scatter(
                        x=[pt[0] for pt in ring], y=[pt[1] for pt in ring],
                        mode="lines", line=dict(color=mask_color, width=0.8),
                        opacity=0.38, hoverinfo="skip", showlegend=False,
                    ))
                outline = entry.get("boundary", [])
                if outline:
                    mesh_fig.add_trace(go.Scatter(
                        x=[pt[0] for pt in outline], y=[pt[1] for pt in outline],
                        mode="lines", line=dict(color=mask_color, width=3),
                        name="Contorno reconstruido",
                    ))
                nominal_line = masks_result.get("nominal", [])
                if nominal_line:
                    mesh_fig.add_trace(go.Scatter(
                        x=[pt[0] for pt in nominal_line], y=[pt[1] for pt in nominal_line],
                        mode="lines", line=dict(color="#64748B", dash="dash", width=1.5),
                        name="Perfil nominal", opacity=0.8,
                    ))
                coords = entry.get("points", [])
                if coords:
                    mesh_fig.add_trace(go.Scatter(
                        x=[pt[0] for pt in coords], y=[pt[1] for pt in coords],
                        mode="markers", marker=dict(color=mask_color, size=6),
                        text=entry.get("ids", []), hovertemplate="Hole ID: %{text}<br>X: %{x:.2f}<br>Z: %{y:.2f}<extra></extra>",
                        name="Collares",
                    ))
                mesh_fig.update_layout(
                    template="plotly_white", height=460,
                    title=mask_title, xaxis_title="X (m)", yaxis_title="Z (m)",
                    yaxis=dict(scaleanchor="x", scaleratio=1),
                    margin=dict(l=25, r=15, t=48, b=30),
                    legend=dict(orientation="h", y=1.01, yanchor="bottom"),
                )
                with mask_col:
                    st.plotly_chart(mesh_fig, width="stretch", key=f"mascara_{mask_key}_{meta.get('round', 'x')}")
                    st.caption(
                        f"{entry.get('n_barrenos', len(coords))} barrenos · "
                        f"{len(coords)} posiciones de collar únicas. "
                        "IDs superpuestos visibles juntos al pasar el cursor."
                    )
                    if entry.get("warning"):
                        st.warning(entry["warning"])

    # La función fragmentada no tiene acceso a `rep` ni a `rsel` del ámbito padre.
    # Se pasan explícitamente los datos de la selección para la cabecera del PDF.
    rep_seleccionado = rsel.get("resumen_reporte") or {}
    ciclo_info = {
        "fecha": rep_seleccionado.get("Fecha_Inicio") or meta.get("Fecha_Inicio") or meta.get("fecha"),
        "equipo": rep_seleccionado.get("Jumbo") or meta.get("Jumbo") or meta.get("jumbo"),
        "operador": (rep_seleccionado.get("Operador_ZDA") or rep_seleccionado.get("Operador")
                     or rep_seleccionado.get("Operario") or rep_seleccionado.get("Operador_Filtro")
                     or meta.get("Operador_ZDA") or meta.get("Operador") or meta.get("operador")),
    }
    _render_mask_depth_fragment(rr.get("mask_volumetry") or {}, masks_result, meta.get("round", "x"), ciclo_info)

    # El análisis principal termina con el modelo de secciones y su PDF. Las visualizaciones
    # históricas basadas en el perímetro (otro modelo de volumen) se retiraron para no presentar
    # volúmenes o perfiles contradictorios.

# ==========================================================
# PRESENTACIÓN POR SECCIONES
# ==========================================================

SECCIONES_ANALISIS = [
    "Uso Automático",
    "Longitud de Perforación",
    "Primer Golpe",
    "Eficiencia de Perforación",
    "ROP por barreno",
    "Clasificación",
    "Resultados por archivo",
    "Asignar operadores",
]

if "seccion_analisis_principal" not in st.session_state:
    st.session_state["seccion_analisis_principal"] = SECCIONES_ANALISIS[0]


def _cambiar_seccion_analisis(seccion):
    """
    Callback de navegación.

    Streamlit ejecuta el callback antes del rerun completo, por lo que
    la nueva sección ya está guardada cuando se vuelven a dibujar los
    botones. Así el resaltado cambia con el primer clic.
    """
    st.session_state["seccion_analisis_principal"] = seccion


st.markdown(
    """
    <style>
    /* Menú de secciones: todos los botones con la MISMA altura, texto centrado.
       Se apunta por la key del widget (st-key-btn_seccion_N), que Streamlit sí genera. */
    [class*="st-key-btn_seccion_"] button {
        height: 76px !important;
        min-height: 76px !important;
        max-height: 76px !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        text-align: center !important;
        padding: 0.4rem 0.6rem !important;
        overflow: hidden;
    }
    [class*="st-key-btn_seccion_"] button p {
        margin: 0 !important;
        line-height: 1.15 !important;
        text-align: center !important;
        white-space: normal;
    }
    @media (max-width: 1100px) {
        [class*="st-key-btn_seccion_"] button { height: 68px !important; min-height: 68px !important; max-height: 68px !important; }
    }
    /* Ventana angosta (p. ej. con el panel lateral abierto): 4 botones por fila en vez de
       apretar 8 y partir las palabras. */
    @media (max-width: 1350px) {
        [data-testid="stHorizontalBlock"]:has([class*="st-key-btn_seccion_"]) { flex-wrap: wrap !important; row-gap: 0.5rem; }
        [data-testid="stHorizontalBlock"]:has([class*="st-key-btn_seccion_"]) > [data-testid="stColumn"] {
            flex: 1 1 calc(25% - 1rem) !important; min-width: calc(25% - 1rem) !important;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

cols_sec = st.columns(len(SECCIONES_ANALISIS))
for i, seccion in enumerate(SECCIONES_ANALISIS):
    activa = st.session_state.get("seccion_analisis_principal") == seccion
    cols_sec[i].button(
        seccion,
        key=f"btn_seccion_{i}",
        width="stretch",
        type="primary" if activa else "secondary",
        help=f"Ir a la sección: {seccion}",
        on_click=_cambiar_seccion_analisis,
        args=(seccion,),
    )

seccion_activa = st.session_state.get(
    "seccion_analisis_principal",
    SECCIONES_ANALISIS[0],
)

if seccion_activa == "Uso Automático":
    with st.container(border=True):
        render_automation_section(
            df_automatico,
            global_jumbos,
            global_tipos,
            global_rocas,
            global_operadores,
            global_turnos,
        )

elif seccion_activa == "Longitud de Perforación":
    with st.container(border=True):
        render_cut_section(
            df_resumen,
            df_reportes,
            global_jumbos,
            global_tipos,
            global_rocas,
            global_operadores,
            global_turnos,
        )

elif seccion_activa == "Primer Golpe":
    with st.container(border=True):
        render_zda_section(
            df_zda,
            global_jumbos,
            global_tipos,
            global_rocas,
            global_operadores,
            global_turnos,
        )

elif seccion_activa == "Eficiencia de Perforación":
    with st.container(border=True):
        render_eficiencia_perforacion_section(
            resultados_validos, global_jumbos, global_tipos, global_rocas, global_operadores, global_turnos
        )

elif seccion_activa == "Clasificación":
    with st.container(border=True):
        st.subheader("Clasificación")
        render_classification_section(
            df_reportes,
            df_automatico,
            df_atipicos,
            global_jumbos,
            global_tipos,
            global_rocas,
            global_operadores,
            global_turnos,
        )

elif seccion_activa == "ROP por barreno":
    with st.container(border=True):
        render_rop_section(
            resultados_validos,
            df_reportes,
            global_jumbos,
            global_tipos,
            global_rocas,
            global_operadores,
            global_turnos,
        )

elif seccion_activa == "Asignar operadores":
    with st.container(border=True):
        render_asignar_operadores_section(df_reportes)

elif seccion_activa == "Resultados por archivo":
    with st.container(border=True):
        st.subheader("Resultados por archivo")
        render_resultados_section(
            resultados_validos,
        )

        if errores:
            st.divider()
            st.subheader("Archivos con error")
            st.dataframe(
                pd.DataFrame([
                    {
                        "Archivo": r.get("nombre_archivo"),
                        "Error": r.get("error"),
                    }
                    for r in errores
                ]),
                width="stretch",
                hide_index=True,
            )
