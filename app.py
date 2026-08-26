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

import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import plotly.graph_objects as go
from plotly.colors import qualitative
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

from procesador import (
    procesar_archivo,
    clasificar_tipo_disparo_v33,
    generar_grafico,
    generar_plano_zda_png,
    extraer_plano_navegacion_png,
)


# ==========================================================
# CONFIGURACIÓN
# ==========================================================

APP_VERSION_INTERNAL = "V34.45-PYTHON-NAVEGACION-CICLOS"
PUBLIC_VERSION = "v1.0"
CACHE_SCHEMA_VERSION = "v34_44_python_masivo_150_zda_20260826"
TIPOS_DISPARO = ["FRENTE", "SELLADA", "ESTOCADA Y/O CORRECCIONES"]
COLORES = qualitative.Plotly

st.set_page_config(page_title=f"EBR Drill Analytics · Piloto {PUBLIC_VERSION}", page_icon="⛏️", layout="wide")
st.title("EBR Drill Analytics")
st.caption(f"Piloto {PUBLIC_VERSION} · Análisis de reportes de perforación de equipos Jumbo Sandvik")
st.info(
    "Consolida y analiza información de perforación de reportes PDF y ZDA, mostrando automatización por jumbo y brazo, "
    "longitud perforada en barrenos Cut, tiempos de ciclo, clasificación de disparos y exportación de datos a Excel."
)

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
    h2 { font-size: 1.22rem !important; margin-top: 0.7rem !important; }
    h3 { font-size: 1.00rem !important; margin-top: 0.55rem !important; margin-bottom: 0.25rem !important; }
    [data-testid="stMetricLabel"] { font-size: 0.75rem !important; }
    [data-testid="stMetricValue"] { font-size: 1.20rem !important; }
    [data-testid="stDataFrame"] div[role="grid"] { font-size: 0.76rem !important; }
    [data-testid="stExpander"] summary { font-size: 0.88rem !important; font-weight: 600 !important; }
    .stDownloadButton button, .stButton button { min-height: 2.15rem !important; font-size: 0.80rem !important; }
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
        )):
            del st.session_state[key]

# ==========================================================
# FILTROS Y PARÁMETROS GLOBALES
# ==========================================================

def _valores_detectados_desde_cache():
    """
    Obtiene jumbos y tipos de disparo desde los resultados ya procesados.
    No presupone cuántos equipos existen.
    """
    jumbos = []
    tipos = []

    for r in st.session_state.procesados.values():
        if r.get("error"):
            continue

        rep = r.get("resumen_reporte") or {}
        jumbo = rep.get("Jumbo")
        tipo = rep.get("Tipo_Disparo")

        if jumbo and str(jumbo).strip():
            jumbos.append(str(jumbo).strip())
        if tipo and str(tipo).strip():
            tipos.append(str(tipo).strip())

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

    return jumbos, tipos


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


jumbos_detectados, tipos_detectados = _valores_detectados_desde_cache()

# Variables siempre definidas aunque aún no existan datos.
global_jumbos = []
global_tipos = []
global_lbl_auto = False
global_lbl_arm = False
global_lbl_cut = False
global_lbl_zda = False

with st.sidebar:
    st.header("Filtros y parámetros")

    if not jumbos_detectados:
        st.info(
            "Los filtros se habilitarán automáticamente después de procesar "
            "los primeros archivos PDF/ZDA."
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

        global_jumbos = st.multiselect(
            "Jumbos",
            jumbos_detectados,
            key="global_jumbos",
        )

        global_tipos = st.multiselect(
            "Tipo de disparo",
            tipos_detectados,
            key="global_tipos",
        )

        st.caption(
            "Los equipos se detectan automáticamente a partir de los archivos procesados."
        )
        st.caption(
            "Estos filtros se aplican a los gráficos y resúmenes consolidados."
        )
        st.caption(
            "**Clasificación:** FRENTE > 45 barrenos · SELLADA 25–45 barrenos · "
            "ESTOCADA Y/O CORRECCIONES < 25 barrenos. Conteo sobre barrenos de frente "
            "(Bottom + Easer + Cut + Contour); Reaming y Casing no se consideran."
        )

        st.divider()
        st.subheader("Opciones de gráficos")

        global_lbl_auto = st.checkbox(
            "Etiquetas · movimiento automático",
            value=False,
            key="opt_lbl_auto",
        )
        global_lbl_arm = st.checkbox(
            "Etiquetas · uso por brazo",
            value=False,
            key="opt_lbl_arm",
        )
        global_lbl_cut = st.checkbox(
            "Etiquetas · barrenos Cut",
            value=False,
            key="opt_lbl_cut",
        )
        global_lbl_zda = st.checkbox(
            "Etiquetas · tiempos de ciclo",
            value=False,
            key="opt_lbl_zda",
        )

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




def hash_archivo(uploaded_file) -> str:
    # Evita uploaded_file.getvalue(), que crea una copia completa del archivo
    # en RAM. getbuffer() entrega una vista de memoria sin duplicar el ZDA/PDF.
    h = hashlib.sha256()
    h.update(uploaded_file.getbuffer())
    h.update(CACHE_SCHEMA_VERSION.encode("utf-8"))
    return h.hexdigest()


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
        suffix = Path(archivo.name).suffix.lower()
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


def fmt(valor, dec=1, sufijo=""):
    if valor is None or pd.isna(valor):
        return "-"
    return f"{float(valor):.{dec}f}{sufijo}"


def figura_a_png(fig) -> bytes:
    buffer = BytesIO()
    fig.savefig(buffer, format="png", dpi=250, bbox_inches="tight")
    buffer.seek(0)
    return buffer.getvalue()


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


def _bd_operator_map(df_reportes):
    out = {}
    if df_reportes.empty or "Fuente" not in df_reportes.columns:
        return out

    pdfs = df_reportes[
        df_reportes["Fuente"].astype(str).str.upper().eq("PDF")
    ].copy()

    for _, r in pdfs.iterrows():
        op = r.get("Operario")
        if op is None or pd.isna(op) or str(op).strip() == "":
            continue
        op = str(op).strip()
        serie_base = re.sub(
            r"-(?:\d+|L)$",
            "",
            str(r.get("Numero_Serie") or ""),
            flags=re.IGNORECASE,
        )
        keys = [
            f"{r.get('Jumbo','')}|{r.get('Ciclo','')}",
            f"{serie_base}|{r.get('Ciclo','')}",
        ]
        for key in keys:
            if key.startswith("|") or key.endswith("|"):
                continue
            out.setdefault(key, [])
            if op not in out[key]:
                out[key].append(op)

    return {k: " / ".join(v) for k, v in out.items()}


def _bd_operador_para_zda(row, op_map):
    serie_base = re.sub(
        r"-(?:\d+|L)$",
        "",
        str(row.get("Numero_Serie") or ""),
        flags=re.IGNORECASE,
    )
    keys = [
        f"{row.get('Jumbo','')}|{row.get('Ciclo','')}",
        f"{serie_base}|{row.get('Ciclo','')}",
    ]
    for key in keys:
        if key in op_map:
            return op_map[key]
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

    op_map = _bd_operator_map(df_reportes)
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
            "OPERADOR": _bd_operador_para_zda(r, op_map),
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


def grafico_auto(df_auto: pd.DataFrame, mostrar_etiquetas: bool):
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

    fig = go.Figure()
    annotations = []
    points = []
    jumbos = list(df["Jumbo"].dropna().astype(str).unique())
    for idx, jumbo in enumerate(jumbos):
        g = df[df["Jumbo"].astype(str) == jumbo].sort_values("FechaHora")
        fig.add_trace(go.Scatter(
            x=g["FechaHora"], y=g["Pct_Movimiento_Automatico_Brazos"],
            mode="lines+markers", name=jumbo,
            line=dict(width=3, color=COLORES[idx % len(COLORES)], shape="spline"),
            marker=dict(size=9),
            customdata=g[["Barrenos_Realizados"]].to_numpy(),
            hovertemplate=(f"{jumbo}<br>%{{x|%d/%m %H:%M}}<br>Automático: %{{y:.1f}}%"
                           "<br>Barrenos: %{customdata[0]} tal.<extra></extra>"),
        ))
        auto = pd.to_numeric(g["Auto_Total_Brazos_min"], errors="coerce").sum(min_count=1)
        manual = pd.to_numeric(g["Manual_Total_Brazos_min"], errors="coerce").sum(min_count=1)
        kpi = auto / (auto + manual) * 100 if pd.notna(auto) and pd.notna(manual) and (auto + manual) > 0 else None
        annotations.append(dict(
            xref="paper", yref="paper", x=0.01, y=1.18 - idx*0.08, xanchor="left",
            text=f"<b>{jumbo}</b> · Automático global: <b>{fmt(kpi,1,'%')}</b>",
            showarrow=False, bgcolor="rgba(255,255,255,0.92)", bordercolor="#dbe3ea", borderpad=5,
        ))
        if mostrar_etiquetas:
            for i, (_, r) in enumerate(g.iterrows()):
                points.append(dict(
                    x=r["FechaHora"], y=r["Pct_Movimiento_Automatico_Brazos"],
                    text=f"{r['Pct_Movimiento_Automatico_Brazos']:.1f}%<br>{int(r['Barrenos_Realizados']) if pd.notna(r.get('Barrenos_Realizados')) else '-'} tal.",
                    rank=i+idx,
                ))
    if mostrar_etiquetas:
        annotations.extend(smart_annotations(points, x_window_hours=18, y_window=4, font_size=11))
    fig.update_layout(**base_layout(
        470, annotations=annotations, margin=dict(l=78,r=35,t=120,b=72),
        yaxis=dict(title="Movimiento automático (%)", rangemode="tozero", gridcolor="#eef2f7"),
        xaxis=dict(title="Fecha", tickformat="%d/%m", gridcolor="#eef2f7"),
    ))
    return fig


def grafico_brazos(df_auto: pd.DataFrame, jumbo: str, mostrar_etiquetas: bool):
    df = df_auto[df_auto["Jumbo"].astype(str) == str(jumbo)].copy()
    # La población ya llega filtrada desde el bloque de automatización.
    df = df[df["Pct_Automatico_Brazo1"].notna() | df["Pct_Automatico_Brazo2"].notna()].copy()
    df = asegurar_fechahora(df)
    if df.empty:
        return None

    fig = go.Figure()
    annotations = []
    points = []
    series = [
        ("Brazo 1", "Pct_Automatico_Brazo1", "Auto_Brazo1_min", "Manual_Brazo1_min", 0),
        ("Brazo 2", "Pct_Automatico_Brazo2", "Auto_Brazo2_min", "Manual_Brazo2_min", 1),
    ]
    for nombre, col, auto_col, man_col, idx in series:
        g = df[df[col].notna()].sort_values("FechaHora")
        if g.empty:
            continue
        auto = pd.to_numeric(g[auto_col], errors="coerce").sum(min_count=1)
        manual = pd.to_numeric(g[man_col], errors="coerce").sum(min_count=1)
        global_pct = auto / (auto + manual) * 100 if pd.notna(auto) and pd.notna(manual) and (auto + manual) > 0 else None
        fig.add_trace(go.Scatter(
            x=g["FechaHora"], y=g[col], mode="lines+markers",
            name=f"{nombre} · Global {fmt(global_pct,1,'%')}",
            line=dict(width=2.5, color=COLORES[idx], shape="spline"), marker=dict(size=8),
            hovertemplate=f"{nombre}<br>%{{x|%d/%m %H:%M}}<br>Automático: %{{y:.1f}}%<extra></extra>",
        ))
        annotations.append(dict(
            xref="paper", yref="paper", x=0.01 + idx*0.25, y=1.16, xanchor="left",
            text=f"<b>{nombre} global: {fmt(global_pct,1,'%')}</b>", showarrow=False,
            bgcolor="rgba(255,255,255,0.92)", bordercolor="#dbe3ea", borderpad=5,
        ))
        if mostrar_etiquetas:
            for i, (_, r) in enumerate(g.iterrows()):
                points.append(dict(x=r["FechaHora"], y=r[col], text=f"{r[col]:.1f}%", rank=i+idx*100))
    if mostrar_etiquetas:
        annotations.extend(smart_annotations(points, x_window_hours=18, y_window=4, font_size=10))
    fig.update_layout(**base_layout(
        430, title=dict(text=f"<b>{jumbo}</b>", x=0.01), annotations=annotations,
        margin=dict(l=72,r=30,t=92,b=72),
        yaxis=dict(title="Movimiento automático (%)", rangemode="tozero", gridcolor="#eef2f7"),
        xaxis=dict(title="Fecha", tickformat="%d/%m", gridcolor="#eef2f7"),
    ))
    return fig


def preparar_cut(df_resumen: pd.DataFrame, df_reportes: pd.DataFrame) -> pd.DataFrame:
    if df_resumen.empty:
        return pd.DataFrame()
    cut = df_resumen[df_resumen["Tipo"].astype(str).str.lower() == "cut"].copy()
    cut = cut[cut["Mediana"].notna()].copy()
    if cut.empty:
        return cut
    cols = [c for c in ["Jumbo","Ciclo","Fecha_Inicio","Tipo_Disparo"] if c in df_reportes.columns]
    tipos = df_reportes[cols].drop_duplicates(subset=[c for c in ["Jumbo","Ciclo","Fecha_Inicio"] if c in cols])
    cut = cut.merge(tipos, on=["Jumbo","Ciclo","Fecha_Inicio"], how="left")
    cut["Tipo_Disparo"] = cut["Tipo_Disparo"].fillna("SIN CLASIFICAR")
    return asegurar_fechahora(cut)



def grafico_cut(df_cut: pd.DataFrame, jumbos_visibles, tipos_visibles, mostrar_etiquetas: bool):
    if df_cut.empty:
        return None

    df = df_cut[
        df_cut["Jumbo"].astype(str).isin([str(x) for x in jumbos_visibles])
        & df_cut["Tipo_Disparo"].isin(tipos_visibles)
    ].copy()

    if df.empty:
        return None

    all_jumbos = sorted(df_cut["Jumbo"].dropna().astype(str).unique())
    visibles = sorted(df["Jumbo"].dropna().astype(str).unique())

    fig = go.Figure()
    points = []
    annotations = []

    for pos_visible, jumbo in enumerate(visibles):
        idx = all_jumbos.index(jumbo)
        g = df[df["Jumbo"].astype(str) == jumbo].sort_values("FechaHora")
        custom = g[["Ciclo", "Tipo_Disparo"]].to_numpy()

        fig.add_trace(go.Scatter(
            x=g["FechaHora"],
            y=g["Mediana"],
            mode="lines+markers",
            name=jumbo,
            customdata=custom,
            line=dict(
                width=3,
                color=COLORES[idx % len(COLORES)],
                shape="spline",
            ),
            marker=dict(size=8),
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
                    color=COLORES[idx % len(COLORES)],
                ),
                bgcolor="rgba(255,255,255,0.95)",
                bordercolor="#dbe3ea",
                borderpad=5,
            ))

        if mostrar_etiquetas:
            for i, (_, r) in enumerate(g.iterrows()):
                points.append(dict(
                    x=r["FechaHora"],
                    y=r["Mediana"],
                    text=f"{r['Mediana']:.2f} m",
                    rank=i + idx * 50,
                ))

    if mostrar_etiquetas:
        annotations.extend(
            smart_annotations(
                points,
                x_window_hours=18,
                y_window=0.18,
                font_size=11,
            )
        )

    fig.update_layout(**base_layout(
        450,
        annotations=annotations,
        margin=dict(l=80, r=30, t=88, b=70),
        yaxis=dict(
            title="Mediana de longitud perforada Cut (m)",
            gridcolor="#eef2f7",
        ),
        xaxis=dict(
            title="Fecha",
            tickformat="%d/%m",
            gridcolor="#eef2f7",
        ),
    ))
    return fig

# ==========================================================
# ZDA: RESÚMENES Y TIMELINE
# ==========================================================


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


def grafico_zda_timeline(rows: pd.DataFrame, mostrar_etiquetas: bool):
    if rows.empty:
        return None
    rows = rows.sort_values("Inicio_Perforacion_TS").copy()
    all_jumbos = sorted(rows["Jumbo"].dropna().astype(str).unique())
    fig = go.Figure()
    annotations = []
    legend_done = set()

    rows["_opDate"] = rows["Inicio_Perforacion_TS"].apply(zda_operational_date)
    for op_date, grupo in rows.groupby("_opDate", sort=True):
        grupo = grupo.sort_values("Inicio_Perforacion_TS")
        n = len(grupo)
        for i, (_, r) in enumerate(grupo.iterrows()):
            offset = (i - (n-1)/2) * 0.11
            x = op_date + timedelta(days=offset)
            y1 = zda_operational_hour(r["Inicio_Perforacion_TS"], op_date)
            y2 = zda_operational_hour(r["Fin_Perforacion_TS"], op_date)
            jumbo = str(r["Jumbo"])
            idx = all_jumbos.index(jumbo)
            n_b = r.get("Barrenos_ZDA") if pd.notna(r.get("Barrenos_ZDA")) else r.get("Barrenos_Realizados")
            custom = [[r.get("Ciclo"), n_b, r.get("Inicio_Perforacion"), r.get("Fin_Perforacion"), r.get("Tiempo_Perforacion_hms"), r.get("Labor") or "-", r.get("Tipo_Disparo")]] * 2
            fig.add_trace(go.Scatter(
                x=[x,x], y=[y1,y2], mode="lines+markers", name=jumbo,
                legendgroup=jumbo, showlegend=jumbo not in legend_done,
                line=dict(width=12,color=COLORES[idx%len(COLORES)]), marker=dict(size=8), customdata=custom,
                hovertemplate=(f"{jumbo} · Ciclo %{{customdata[0]}}<br>Tipo: %{{customdata[6]}}"
                               "<br>Inicio: %{customdata[2]}<br>Fin: %{customdata[3]}<br>Tiempo: %{customdata[4]}"
                               "<br>Barrenos: %{customdata[1]} B<br>Labor: %{customdata[5]}<extra></extra>"),
            ))
            legend_done.add(jumbo)
            if mostrar_etiquetas:
                annotations.append(dict(
                    x=x, y=(y1+y2)/2, xref="x", yref="y", showarrow=False, xshift=18,
                    text=f"C{r.get('Ciclo')} · {r.get('Tiempo_Perforacion_hms') or '-'} | {int(n_b) if pd.notna(n_b) else '-'}B",
                    font=dict(size=10,color="#334155"), bgcolor="rgba(255,255,255,.9)",
                    bordercolor="#dbe3ea", borderpad=3, xanchor="left",
                ))

    tickvals = [7,9,11,13,15,17,19,21,23,25,27,29,31]
    ticktext = [f"{v%24:02d}:00" for v in tickvals]
    op_dates = sorted(rows["_opDate"].unique())
    fig.update_layout(**base_layout(
        520, annotations=annotations, margin=dict(l=80,r=150,t=55,b=72),
        xaxis=dict(title="Fecha operativa", tickvals=op_dates,
                   ticktext=[pd.Timestamp(x).strftime("%d/%m") for x in op_dates], gridcolor="#eef2f7"),
        yaxis=dict(title="Hora", range=[31.2,6.8], tickmode="array", tickvals=tickvals, ticktext=ticktext, gridcolor="#eef2f7"),
        shapes=[
            dict(type="rect",xref="paper",x0=0,x1=1,yref="y",y0=7,y1=19,fillcolor="rgba(37,99,235,.035)",line=dict(width=0),layer="below"),
            dict(type="rect",xref="paper",x0=0,x1=1,yref="y",y0=19,y1=31,fillcolor="rgba(15,23,42,.035)",line=dict(width=0),layer="below"),
            dict(type="line",xref="paper",x0=0,x1=1,yref="y",y0=19,y1=19,line=dict(color="#94a3b8",width=1,dash="dot")),
        ],
    ))
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
    """Genera boxplot/plano solo cuando el usuario solicita ese ciclo."""
    cache_key = r.get("_cache_key") or hashlib.sha1(
        str(r.get("nombre_archivo", "")).encode("utf-8")
    ).hexdigest()
    box_path, nav_path = _visual_paths(cache_key)

    detalle = r.get("detalle")
    rep = r.get("resumen_reporte") or {}
    metadata = r.get("metadata") or rep
    fuente = str(rep.get("Fuente") or r.get("fuente") or "PDF").upper()
    source_path = Path(r.get("_source_path") or "")

    if not box_path.exists() and isinstance(detalle, pd.DataFrame) and not detalle.empty:
        fig = generar_grafico(detalle, metadata)
        fig.savefig(box_path, format="png", dpi=170, bbox_inches="tight", facecolor="white")
        plt.close(fig)
        del fig
        gc.collect()

    if not nav_path.exists():
        nav_bytes = None
        if fuente == "ZDA" and isinstance(detalle, pd.DataFrame) and not detalle.empty:
            nav_bytes = generar_plano_zda_png(detalle, metadata, resolution=150)
        elif fuente == "PDF" and source_path.exists():
            nav_bytes = extraer_plano_navegacion_png(source_path, resolution=150)

        if nav_bytes:
            nav_path.write_bytes(nav_bytes)
            del nav_bytes
            gc.collect()

    return box_path if box_path.exists() else None, nav_path if nav_path.exists() else None


# ==========================================================
# CARGA MASIVA EN DOS FASES
# ==========================================================

st.subheader("Archivos")

archivos = st.file_uploader(
    "Seleccionar archivos PDF / ZDA",
    type=["pdf", "zda"],
    accept_multiple_files=True,
    key=f"uploader_{st.session_state.uploader_version}",
    help=(
        "Modo masivo: al presionar Procesar, primero se copian los archivos a "
        "disco temporal y se libera el uploader. Luego se procesan uno por uno."
    ),
)

seleccion = archivos or []
total_upload_mb = sum(int(getattr(a, "size", 0) or 0) for a in seleccion) / (1024 * 1024)

col_process, col_clear = st.columns(2)
with col_process:
    procesar = st.button(
        "Procesar lote",
        type="primary",
        use_container_width=True,
        disabled=not seleccion,
    )
with col_clear:
    st.button(
        "Limpiar",
        use_container_width=True,
        on_click=limpiar_analisis,
    )

st.caption(
    f"Seleccionados: {len(seleccion)} · Tamaño del lote: {total_upload_mb:,.1f} MB · "
    f"Procesados acumulados: {len(st.session_state.procesados)}"
)

if seleccion and total_upload_mb >= 700:
    st.warning(
        "El lote cargado supera aproximadamente 700 MB. El modo masivo libera los "
        "archivos antes del parsing, pero Streamlit Community Cloud debe recibir primero "
        "todo el lote en el uploader. Si la carga por sí sola supera la memoria disponible, "
        "será necesario dividir únicamente la etapa de carga."
    )

if procesar and seleccion:
    estado_stage = st.empty()
    estado_stage.write(
        f"Preparando {len(seleccion)} archivo(s) en disco temporal..."
    )
    nuevos = preparar_archivos_en_disco(seleccion)
    estado_stage.empty()

    # CRÍTICO: borrar el widget ANTES de procesar. En el siguiente rerun los
    # 150 UploadedFile ya no permanecen en memoria; sólo quedan paths en disco.
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
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.write(
            "Selecciona uno o varios archivos PDF/ZDA y presiona "
            "**Procesar / recalcular**."
        )
    st.stop()


# ==========================================================
# CONSOLIDACIÓN
# ==========================================================

report_rows = [dict(r["resumen_reporte"]) for r in resultados_validos]
df_reportes = pd.DataFrame(report_rows)

# Fuerza la clasificación V33 para PDF y ZDA con el mismo criterio.
df_reportes["Tipo_Disparo"] = df_reportes["Barrenos_Realizados"].apply(clasificar_tipo_disparo_v33)
df_reportes["Considerado_KPI_Automatizacion"] = df_reportes["Tipo_Disparo"].eq("FRENTE")
for r in resultados_validos:
    rr = r["resumen_reporte"]
    rr["Tipo_Disparo"] = clasificar_tipo_disparo_v33(rr.get("Barrenos_Realizados"))
    rr["Considerado_KPI_Automatizacion"] = rr["Tipo_Disparo"] == "FRENTE"

# HTML V33 solo agrega Resumen_Ciclos de reportes cuyo conteo está OK.
df_resumen = concatenar_dataframes(resultados_validos, "resumen_ciclo", solo_ok=True)
df_detalle = concatenar_dataframes(resultados_validos, "detalle")
df_atipicos = concatenar_dataframes(resultados_validos, "atipicos")
df_automatico = df_reportes.copy()
df_zda = df_reportes[df_reportes["Fuente"].eq("ZDA")].copy() if "Fuente" in df_reportes.columns else pd.DataFrame()


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


# ==========================================================
# BLOQUE 1 - AUTOMATIZACIÓN
# ==========================================================

@fragment
def render_automation_section(
    df_automatico: pd.DataFrame,
    sel_jumbos,
    sel_tipos,
    mostrar_auto: bool,
    mostrar_arm: bool,
):
    if df_automatico.empty:
        st.info("Sin datos suficientes de automatización.")
        return

    df_visible = df_automatico[
        df_automatico["Jumbo"].astype(str).isin(
            [str(x) for x in sel_jumbos]
        )
        & df_automatico["Tipo_Disparo"].isin(sel_tipos)
    ].copy()

    st.subheader("Evolución del movimiento automático")

    fig_auto = grafico_auto(
        df_visible,
        mostrar_auto,
    )

    if fig_auto is not None:
        st.plotly_chart(
            fig_auto,
            use_container_width=True,
            config={"displaylogo": False},
        )
    else:
        st.info(
            "No hay ciclos visibles con datos de movimiento automático "
            "para los filtros globales seleccionados."
        )

    st.subheader("Uso automático por brazo")

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
            mostrar_arm,
        )
        if fig_arm is not None:
            hubo_grafico = True
            st.plotly_chart(
                fig_arm,
                use_container_width=True,
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
    mostrar_cut: bool,
):
    st.subheader("Evolución de la longitud perforada en barrenos Cut")

    df_cut = preparar_cut(
        df_resumen,
        df_reportes,
    )

    if df_cut.empty:
        st.info("Sin datos suficientes de barrenos Cut.")
        return

    st.caption(
        "Cada punto representa la mediana de la longitud perforada "
        "de los barrenos Cut de cada ciclo. Se aplican los filtros "
        "globales del panel lateral."
    )

    fig_cut = grafico_cut(
        df_cut,
        sel_jumbos,
        sel_tipos,
        mostrar_cut,
    )

    if fig_cut is not None:
        st.plotly_chart(
            fig_cut,
            use_container_width=True,
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
def render_zda_section(
    df_zda: pd.DataFrame,
    sel_jumbos,
    sel_tipos,
    mostrar_zda: bool,
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
        "Eje X: fecha operativa (07:00–07:00). Eje Y: hora. "
        "Cada barra va desde el primer registro MWD hasta el último. "
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
    ].copy()

    st.caption(
        "Los filtros globales del panel lateral actualizan la gráfica "
        "y los resúmenes de esta sección."
    )

    st.markdown("#### Hora promedio de inicio por jumbo y turno")

    shift = resumen_turnos_zda(
        zda_rows
    )

    if not shift.empty:
        st.dataframe(
            shift,
            use_container_width=True,
            hide_index=True,
        )
        st.caption(
            "“Ciclos día/noche” cuenta todos los rounds del turno. "
            "Para promedio, inicio más temprano e inicio más tarde "
            "se considera solo el primer round de cada jumbo por "
            "fecha operativa y turno."
        )
    else:
        st.info(
            "No hay ciclos visibles para calcular los indicadores de inicio."
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
                use_container_width=True,
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

    st.markdown(
        "#### Gráfica de tiempos x fecha operativa"
    )

    fig_zda = grafico_zda_timeline(
        zda_rows,
        mostrar_zda,
    )

    if fig_zda is not None:
        st.plotly_chart(
            fig_zda,
            use_container_width=True,
            config={"displaylogo": False},
        )
    else:
        st.info(
            "No hay ciclos visibles con los filtros globales seleccionados."
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
):
    filtrados = df_reportes[
        df_reportes["Jumbo"].astype(str).isin(
            [str(x) for x in sel_jumbos]
        )
        & df_reportes["Tipo_Disparo"].isin(sel_tipos)
    ].copy()

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
                use_container_width=True,
                hide_index=True,
            )

    with col_read:
        st.markdown("#### Resumen de lectura")

        pdf_rows = (
            filtrados[
                filtrados["Fuente"].eq("PDF")
            ]
            if "Fuente" in filtrados.columns
            else pd.DataFrame()
        )

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

        pdf_ok = (
            int(
                pdf_rows.get(
                    "Lectura_Confiable",
                    pd.Series(dtype=object),
                ).eq("OK").sum()
            )
            if not pdf_rows.empty
            else 0
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

        r1, r2, r3, r4, r5 = st.columns(5)

        r1.metric(
            "Archivos",
            len(filtrados),
        )
        r2.metric(
            "PDF",
            len(pdf_rows),
        )
        r3.metric(
            "ZDA",
            len(zda_rows_all),
        )
        r4.metric(
            "Lectura PDF OK",
            f"{pdf_ok}/{len(pdf_rows)}",
        )
        r5.metric(
            "Lectura ZDA OK",
            f"{zda_ok}/{len(zda_rows_all)}",
        )

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
    ].copy()

    cols_auto = [
        c
        for c in [
            "Fecha_Inicio",
            "Jumbo",
            "Ciclo",
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
        use_container_width=True,
        hide_index=True,
        height=360,
    )


# ==========================================================
# BLOQUE 5 - RESULTADOS POR ARCHIVO
# ==========================================================

@fragment
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
            use_container_width=True,
        ):
            page = max(1, page - 1)
            st.session_state[page_key] = page

    with nav_next:
        if st.button(
            "Siguiente →",
            key="resultados_next",
            disabled=page >= n_pages,
            use_container_width=True,
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
            or "PDF"
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
                m1, m2, m3, m4, m5, m6 = st.columns(6)

                m1.metric(
                    "Serie",
                    rep.get("Numero_Serie") or "-",
                )
                m2.metric(
                    "Sección",
                    seccion_desde_plan_texto(
                        rep.get("Plan_Perforacion")
                    ),
                )
                m3.metric(
                    "Tipo de disparo",
                    rep.get("Tipo_Disparo") or "-",
                )
                m4.metric(
                    "Barrenos",
                    int(rep["Barrenos_Realizados"])
                    if pd.notna(rep.get("Barrenos_Realizados"))
                    else "-",
                )
                m5.metric(
                    "Metros perforados",
                    fmt(
                        rep.get("Metros_Perforados"),
                        2,
                        " m",
                    ),
                )
                m6.metric(
                    "Lectura",
                    rep.get("Lectura_Confiable")
                    or rep.get("Estado")
                    or "-",
                )

                a1, a2, a3, a4 = st.columns(4)
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

                if fuente == "ZDA":
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
                    if fuente == "ZDA":
                        st.caption(
                            "Plano reconstruido desde ZDA · "
                            f"sección {seccion_desde_plan_texto(rep.get('Plan_Perforacion'))}"
                        )
                    else:
                        st.caption("Plano de navegación del PDF")
                    st.image(str(nav_path), use_container_width=True)
                elif not mostrar_detalle:
                    st.caption("Plano disponible bajo demanda")

            if box_path:
                st.image(str(box_path), use_container_width=True)
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
                        use_container_width=True,
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
                        use_container_width=True,
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
                    use_container_width=True,
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
                        use_container_width=True,
                        hide_index=True,
                        height=260,
                    )


# ==========================================================
# PRESENTACIÓN EN BLOQUES VISUALES
# ==========================================================

st.divider()
st.header("Consolidado")

with st.container(border=True):
    render_automation_section(
        df_automatico,
        global_jumbos,
        global_tipos,
        global_lbl_auto,
        global_lbl_arm,
    )

with st.container(border=True):
    render_cut_section(
        df_resumen,
        df_reportes,
        global_jumbos,
        global_tipos,
        global_lbl_cut,
    )

with st.container(border=True):
    render_zda_section(
        df_zda,
        global_jumbos,
        global_tipos,
        global_lbl_zda,
    )

with st.expander(
    "Clasificación de disparos",
    expanded=False,
):
    render_classification_section(
        df_reportes,
        df_automatico,
        df_atipicos,
        global_jumbos,
        global_tipos,
    )

with st.expander(
    f"Resultados por archivo · {len(resultados_validos)} archivos procesados",
    expanded=False,
):
    render_resultados_section(
        resultados_validos,
    )

if errores:
    with st.container(border=True):
        st.subheader("Archivos con error")
        st.dataframe(
            pd.DataFrame([
                {
                    "Archivo": r.get("nombre_archivo"),
                    "Error": r.get("error"),
                }
                for r in errores
            ]),
            use_container_width=True,
            hide_index=True,
        )
