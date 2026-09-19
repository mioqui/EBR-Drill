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

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
from matplotlib.patches import Patch

ORDEN_TIPOS = ["Bottom", "Easer", "Cut", "Contour", "Reaming", "Casing"]
JUMBOS = {"125D114796": "JUMB001", "125D98943": "JUMB002"}

VERSION_PROCESADOR = "V34.49-PLAN-COMPLETO"


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


def pct(parte: Optional[float], total: Optional[float]) -> Optional[float]:
    if parte is None or total is None or total <= 0:
        return None
    return (parte / total) * 100.0


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


# ==========================================================
# SALIDA ESTÁNDAR + LECTOR ZDA
# ==========================================================

ZDA_TIPO_CODES = {
    0: "Reaming",
    1: "Contour",
    4: "Cut",
    5: "Easer",
    8: "Bottom",
    9: "Casing",
}


def clasificar_tipo_disparo_v33(barrenos_realizados):
    """Clasificación: FRENTE >=45; SELLADA 25-44; ESTOCADA <25."""
    if barrenos_realizados is None or pd.isna(barrenos_realizados):
        return "SIN CLASIFICAR"
    n = int(barrenos_realizados)
    if n >= 45:
        return "FRENTE"
    if n >= 25:
        return "SELLADA"
    return "ESTOCADA Y/O CORRECCIONES"


def _enriquecer_resultado_estandar(resultado: Dict, fuente: str) -> Dict:
    """Normaliza la salida ZDA al esquema consumido por la aplicación."""
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
    metadata["Archivo_Fuente"] = metadata.get("Archivo_ZDA")

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
            if not ts_ok:
                if data[start + 17] == 0x8A and data[start + 21] == 3:
                    # Terminó el bloque de ejecutados de 297 bytes; a continuación
                    # hay barrenos solo programados de 154 bytes (marca kind=3).
                    # Nunca interpretarlos como perforaciones.
                    raw_records -= 1
                    break
                # Registro ejecutado con marca de tiempo dañada: se cuenta como
                # inválido y se continúa; NO se descartan los barrenos siguientes.
                invalid += 1
                continue
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
                ident = f"E{sum(1 for r in rows if str(r['ID']).startswith('E')) + 1}"
            length2 = round(length, 2)
            depth2 = round(y2, 2)
            rows.append({
                "Fuente": "ZDA",
                "Archivo_Fuente": nombre_archivo,
                "Archivo_ZDA": nombre_archivo,
               
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
                "Archivo_ZDA": metadata.get("Archivo_ZDA"),
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

    # Sin numeración visible: estilo del plano de navegación.
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

    # Flechas del sistema de referencia, del plano de navegación.
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

    # Calibración visual del plano reconstruido:
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

    # Calibración visual del borde superior:
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
    # Con esto se ajusta mejor la geometría de referencia:
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

    # Corona del frente:
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

    # Título del plano.
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
            "Fuente": "ZDA", "Archivo_Fuente": nombre, "Archivo_ZDA": nombre,
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
    """Procesa exclusivamente archivos ZDA."""
    if path.suffix.lower() != ".zda":
        raise ValueError("Formato no soportado. Use archivos .ZDA.")
    return procesar_zda(
        path,
        nombre_archivo=nombre_archivo,
        generar_visuales=generar_visuales,
    )


# ==========================================================
# EFICIENCIA DE PERFORACIÓN · MOTOR ZDA INTEGRADO
# ==========================================================


# (los imports generales ya están al inicio del módulo)
try:
    from shapely.geometry import Polygon
except Exception:
    Polygon = None


TIPO_CODES = {
    0: "Reaming",
    1: "Contour",
    3: "Reference",
    4: "Cut",
    5: "Easer",
    8: "Bottom",
    9: "Casing",
}

PERIMETER_TYPES = {"Contour", "Bottom"}
BOOM_RECORD_SIZE = 297
BOOM_FIRST_RECORD = 4

VERSION = "V1.2-ZDA-PLAN-COMPLETO"


def _ascii(data: bytes) -> str:
    return data.split(b"\x00", 1)[0].decode("utf-8", "ignore").strip()


def _f64(data: bytes, offset: int) -> float:
    return struct.unpack_from("<d", data, offset)[0]


def _u32(data: bytes, offset: int) -> int:
    return struct.unpack_from("<I", data, offset)[0]


def _u64(data: bytes, offset: int) -> int:
    return struct.unpack_from("<Q", data, offset)[0]


# np.trapezoid solo existe desde NumPy 2.0; requirements permite >=1.26.
_trapz = getattr(np, "trapezoid", None) or getattr(np, "trapz")


def _polygon_area(points: List[Tuple[float, float]]) -> float:
    if len(points) < 3:
        return float("nan")
    arr = np.asarray(points, dtype=float)
    x = arr[:, 0]
    z = arr[:, 1]
    return float(abs(np.dot(x, np.roll(z, -1)) - np.dot(z, np.roll(x, -1))) / 2.0)


def _segment_key(p: Tuple[float, float], nd: int = 5):
    return (round(float(p[0]), nd), round(float(p[1]), nd))


def _arc_center(p1, p2, curvature):
    """
    El 7mo double observado en round.dat se comporta como curvatura k=1/R.
    Signo positivo: arco horario de p1 a p2.
    """
    x1, z1 = p1
    x2, z2 = p2
    k = float(curvature)
    if abs(k) < 1e-12:
        return None, None, None
    r = 1.0 / abs(k)
    dx, dz = x2 - x1, z2 - z1
    chord = math.hypot(dx, dz)
    if chord <= 1e-12 or chord > 2 * r + 1e-9:
        return None, None, None
    mx, mz = (x1 + x2) / 2, (z1 + z2) / 2
    h = math.sqrt(max(r * r - (chord / 2) ** 2, 0.0))
    ux, uz = -dz / chord, dx / chord
    candidates = [(mx + ux * h, mz + uz * h), (mx - ux * h, mz - uz * h)]

    def signed_minor(center):
        cx, cz = center
        a1 = math.atan2(z1 - cz, x1 - cx)
        a2 = math.atan2(z2 - cz, x2 - cx)
        d = (a2 - a1 + math.pi) % (2 * math.pi) - math.pi
        return d

    want_clockwise = k > 0
    for c in candidates:
        d = signed_minor(c)
        if (want_clockwise and d < 0) or ((not want_clockwise) and d > 0):
            return c, r, d
    c = candidates[0]
    return c, r, signed_minor(c)


def _sample_segment(seg: Dict, n_arc: int = 18) -> List[Tuple[float, float]]:
    p1 = (seg["x1"], seg["z1"])
    p2 = (seg["x2"], seg["z2"])
    k = float(seg.get("curvature", 0.0) or 0.0)
    if abs(k) < 1e-12:
        return [p1, p2]

    center, r, delta = _arc_center(p1, p2, k)
    if center is None:
        return [p1, p2]
    cx, cz = center
    a1 = math.atan2(p1[1] - cz, p1[0] - cx)
    ts = np.linspace(0.0, 1.0, n_arc)
    return [(cx + r * math.cos(a1 + delta * t), cz + r * math.sin(a1 + delta * t)) for t in ts]


def _chain_segments(segments: List[Dict]) -> List[Dict]:
    if not segments:
        return []
    unused = list(range(len(segments)))
    chain = [segments[unused.pop(0)].copy()]

    while unused:
        end = (chain[-1]["x2"], chain[-1]["z2"])
        endk = _segment_key(end)
        found = None
        reverse = False
        for j, idx in enumerate(unused):
            s = segments[idx]
            if _segment_key((s["x1"], s["z1"])) == endk:
                found = j
                reverse = False
                break
            if _segment_key((s["x2"], s["z2"])) == endk:
                found = j
                reverse = True
                break
        if found is None:
            break
        idx = unused.pop(found)
        s = segments[idx].copy()
        if reverse:
            s["x1"], s["x2"] = s["x2"], s["x1"]
            s["y1"], s["y2"] = s["y2"], s["y1"]
            s["z1"], s["z2"] = s["z2"], s["z1"]
            s["curvature"] = -float(s.get("curvature", 0.0) or 0.0)
        chain.append(s)
    return chain


def parse_round_txt(text: str) -> Dict:
    kv = {}
    for line in text.splitlines():
        if ":" not in line:
            continue
        k, v = line.split(":", 1)
        kv[k.strip()] = v.strip()
    return {
        "rig": kv.get("rig"),
        "boom_count": int(kv["boom_count"]) if str(kv.get("boom_count", "")).isdigit() else None,
        "round": int(kv["round"]) if str(kv.get("round", "")).isdigit() else None,
        "drill_plan": kv.get("drill_plan"),
        "planned_face_holes": int(kv["planned_face_holes"]) if str(kv.get("planned_face_holes", "")).isdigit() else None,
        "navigation": kv.get("navigation"),
        "start": kv.get("start"),
        "end": kv.get("end"),
        "drilled_holes": int(kv["drilled_holes"]) if str(kv.get("drilled_holes", "")).isdigit() else None,
        "tunnel_id": kv.get("tunnel_id"),
        "curve_table": kv.get("curve_table"),
        "operator": _extract_operator(kv.get("tunnel_id", "")),
    }


def _extract_operator(tunnel_id: str) -> Optional[str]:
    txt = str(tunnel_id or "")
    m = re.search(r"\bOP\s*[:.=]\s*([^:]+?)(?=\s+\b[A-Z]{1,3}\s*[:.=]|$)", txt, re.I)
    if m:
        v = m.group(1).strip()
        return v or None
    return None


def parse_boom_dat(data: bytes) -> pd.DataFrame:
    """Lee los DOS formatos de registros del boom.dat DD322i.

    - 297 bytes: barreno ejecutado, con plan opcional y coordenadas reales.
    - 154 bytes (+ 5 bytes finales compartidos en el último registro):
      barreno SOLO programado; conserva collar y fondo aunque no se perforó.

    Las coordenadas planificadas ocupan offsets 111..158 en ambos formatos.
    La longitud de los registros se detecta a partir de las marcas de tiempo y
    coordenadas de ejecución, no de un número de barrenos fijo para el ciclo.
    """
    rows = []
    cursor = BOOM_FIRST_RECORD
    extra_index = 0
    while cursor + 159 <= len(data):
        start = cursor
        try:
            # El código del tipo se ubica a +175 en registro ejecutado
            # y a +90 en registro exclusivamente programado.
            tipo = None
            hole_id = _ascii(data[start + 26:start + 56])
            boom0 = data[start + 5]
            seq = data[start + 6]
            p = [_f64(data, start + o) for o in (111, 119, 127)]
            p2 = [_f64(data, start + o) for o in (135, 143, 151)]
            plan_len = math.dist(p, p2) if all(math.isfinite(v) for v in p + p2) else float('nan')
            long_record = False
            if start + BOOM_RECORD_SIZE <= len(data):
                ts = _u32(data, start + 163)
                end = _u32(data, start + 240)
                actual = [_f64(data, start + o) for o in (183, 191, 199, 257, 265, 273)]
                actual_ok = (all(math.isfinite(v) and abs(v) < 100 for v in actual)
                             and 0.1 < math.dist(actual[:3], actual[3:]) < 20)
                long_record = (1577836800 < ts <= end < 2051222400 and actual_ok
                               and boom0 in (0, 1) and 0 < seq < 100
                               )
            type_code = data[start + (175 if long_record else 90)]
            tipo = TIPO_CODES.get(type_code, f"Desconocido (código {type_code})")
            if long_record:
                record_size = BOOM_RECORD_SIZE
                a = actual[:3]
                a2 = actual[3:]
                start_ts, end_ts = ts, end
            else:
                record_size = 154
                # No tomar valores del registro siguiente como ejecución.
                a = a2 = [float('nan')] * 3
                start_ts = end_ts = None
                if not (tipo is not None and 0.1 < plan_len < 20.0
                        and all(math.isfinite(v) and abs(v) < 100 for v in p + p2)):
                    raise ValueError(f"Registro de diseño no reconocido en offset {start}.")

            if not hole_id:
                if long_record:
                    extra_index += 1
                    hole_id = f"E{extra_index}"
                else:
                    raise ValueError(f"Barreno programado sin ID en offset {start}.")
            extra = (not (0.1 < plan_len < 20.0)
                     or (hole_id.upper().startswith('E') and long_record))
            if (not long_record) and hole_id.upper().startswith('E'):
                raise ValueError(f"Barreno sin ejecución y marcado extra: {hole_id}.")
            planned_exists = (not extra and 0.1 < plan_len < 20.0)
            actual_len = math.dist(a, a2) if long_record else float('nan')
            rows.append({
                # En registros solo programados, Brazo es el brazo PLANIFICADO (offset +5).
                "ID": hole_id, "Brazo": boom0 + 1 if boom0 in (0, 1) else np.nan,
                "Secuencia": int(seq) if long_record and 0 < seq < 100 else np.nan,
                "Tipo": tipo, "Extra": extra,
                "Plan_X": p[0] if planned_exists else np.nan,
                "Plan_Y": p[1] if planned_exists else np.nan,
                "Plan_Z": p[2] if planned_exists else np.nan,
                "Plan_X2": p2[0] if planned_exists else np.nan,
                "Plan_Y2": p2[1] if planned_exists else np.nan,
                "Plan_Z2": p2[2] if planned_exists else np.nan,
                "Plan_Longitud_m": plan_len if planned_exists else np.nan,
                "X": a[0], "Y": a[1], "Z": a[2],
                "X2": a2[0], "Y2": a2[1], "Z2": a2[2],
                "Longitud_m": actual_len,
                "Avance_Y_m": a2[1] - a[1] if long_record else np.nan,
                "Inicio_TS": start_ts, "Fin_TS": end_ts,
                "Estado_Barreno": ("Extra ejecutado" if extra else
                                   "Plan y ejecución" if long_record else "Programado no perforado"),
                "Fuente_Plan": "ZDA boom.dat" if planned_exists else None,
                "Fuente_Real": "ZDA boom.dat" if long_record else None,
                "Offset_Boom": start, "Tamano_Registro_Boom": record_size,
            })
            cursor += record_size
        except (struct.error, IndexError, OverflowError) as exc:
            raise ValueError(f"Registro binario dañado en boom.dat offset {start}: {exc}") from exc
    if cursor + 5 != len(data):
        raise ValueError(f"Fin de boom.dat inesperado: cursor={cursor}, tamaño={len(data)}")
    df = pd.DataFrame(rows)
    if not df.empty:
        df["Desv_Collar_m"] = np.sqrt(
            (df["X"] - df["Plan_X"]) ** 2
            + (df["Y"] - df["Plan_Y"]) ** 2
            + (df["Z"] - df["Plan_Z"]) ** 2
        )
        df["Desv_Toe_m"] = np.sqrt(
            (df["X2"] - df["Plan_X2"]) ** 2
            + (df["Y2"] - df["Plan_Y2"]) ** 2
            + (df["Z2"] - df["Plan_Z2"]) ** 2
        )
    return df


def parse_round_dat_profile(data: bytes) -> Dict:
    """
    Recupera la geometría nominal del frente desde round-*.dat.

    En los ZDA DD322i analizados, el perfil aparece como registros:
      tipo=3, timestamp uint64, payload_len=56, 7 doubles
    con:
      x1,y1,z1,x2,y2,z2,curvature
    donde curvature=0 es línea y |curvature|=1/R para arcos.
    """
    segments = []
    for pos in range(0, max(0, len(data) - 69)):
        if data[pos] != 3:
            continue
        try:
            ts = _u64(data, pos + 1)
            length = _u32(data, pos + 9)
            if length != 56 or pos + 13 + 56 > len(data):
                continue
            vals = struct.unpack_from("<7d", data, pos + 13)
            if not all(math.isfinite(v) for v in vals):
                continue
            x1, y1, z1, x2, y2, z2, k = vals
            if max(abs(x1), abs(y1), abs(z1), abs(x2), abs(y2), abs(z2)) > 50:
                continue
            if math.dist((x1, y1, z1), (x2, y2, z2)) < 0.05:
                continue
            segments.append({
                "offset": pos,
                "timestamp": ts,
                "x1": x1, "y1": y1, "z1": z1,
                "x2": x2, "y2": y2, "z2": z2,
                "curvature": k,
            })
        except Exception:
            continue

    # Deduplicar por endpoints+curvatura
    uniq = []
    seen = set()
    for s in segments:
        key = tuple(round(s[k], 6) for k in ("x1","y1","z1","x2","y2","z2","curvature"))
        if key not in seen:
            seen.add(key)
            uniq.append(s)

    chain = _chain_segments(uniq)
    sampled = []
    for i, seg in enumerate(chain):
        pts = _sample_segment(seg)
        if i > 0:
            pts = pts[1:]
        sampled.extend(pts)

    area = _polygon_area(sampled) if len(sampled) >= 3 else float("nan")
    return {
        "segments": uniq,
        "chain": chain,
        "polygon": sampled,
        "area_m2": area,
        "decoded": len(sampled) >= 3 and math.isfinite(area),
    }


def _build_ordered_profiles(df: pd.DataFrame, nominal_profile: Dict | None = None) -> Dict:
    """Reconstrucción conservadora: el diseño manda; nunca completar con reales."""
    nominal_pts = (nominal_profile or {}).get("polygon", [])
    # LEGACY perimeter model requires pairs with both geometries; the official
    # masks use separate full plan (54) and executed (50) sets below.
    candidates = df[(~df["Extra"].fillna(False)) &
                    df[["Plan_X", "Plan_Z", "Plan_X2", "Plan_Z2",
                        "X", "Z", "X2", "Z2"]].notna().all(axis=1)].copy()
    if candidates.empty or len(nominal_pts) < 3:
        return {"ok": False, "reason": "Faltan collares programados o perfil nominal.", "audit_rows": candidates}
    boundary = np.asarray(nominal_pts, dtype=float)
    if not np.isfinite(boundary).all():
        return {"ok": False, "reason": "Perfil nominal con coordenadas inválidas.", "audit_rows": candidates}
    # Proyección a la polilínea nominal: posición longitudinal y distancia.
    a = boundary[:-1] if np.linalg.norm(boundary[0]-boundary[-1]) < 1e-6 else boundary
    b = np.roll(a, -1, axis=0)
    ab = b-a; lengths = np.linalg.norm(ab, axis=1)
    cumulative = np.r_[0., np.cumsum(lengths)]
    distances=[]; stations=[]
    for _, hole in candidates.iterrows():
        point=np.array([float(hole["Plan_X"]),float(hole["Plan_Z"])])
        u=np.clip(np.divide(np.sum((point-a)*ab,axis=1),np.sum(ab*ab,axis=1),out=np.zeros(len(a)),where=lengths>1e-9),0,1)
        d=np.linalg.norm(point-(a+u[:,None]*ab),axis=1)
        j=int(np.argmin(d));distances.append(float(d[j]));stations.append(float(cumulative[j]+u[j]*lengths[j]))
    candidates["Distancia_nominal_m"]=distances
    candidates["Posicion_perimetro_m"]=stations
    # Tipo es evidencia auxiliar; la proximidad geométrica es obligatoria.
    near=candidates["Distancia_nominal_m"] <= 0.65
    base = candidates["Tipo"].isin(PERIMETER_TYPES) & near
    per=candidates.loc[base].copy()
    # Otros tipos solo como candidatos si la selección principal deja huecos;
    # nunca se promueven indiscriminadamente todos los Easer cercanos.
    def _largest_gap(rows):
        if rows.empty: return float("inf")
        st=np.sort(rows["Posicion_perimetro_m"].to_numpy())
        return float(np.diff(np.r_[st,st[0]+cumulative[-1]]).max())
    if _largest_gap(per)>max(2.25,0.16*float(cumulative[-1])):
        extra=candidates.loc[near & ~candidates["Tipo"].isin(PERIMETER_TYPES)].copy()
        # Incorporar candidatos solo cuando reducen el mayor tramo sin puntos.
        for _, candidate in extra.sort_values("Distancia_nominal_m").iterrows():
            trial=pd.concat([per,candidate.to_frame().T],ignore_index=True)
            if _largest_gap(trial)<_largest_gap(per)-0.05:
                per=trial
    per=per.sort_values(["Posicion_perimetro_m","Distancia_nominal_m"]).drop_duplicates(subset=["Plan_X","Plan_Z"]).reset_index(drop=True)
    excluded=candidates.loc[~near,"ID"].tolist()
    if len(per)<8:
        return {"ok":False,"reason":"Menos de ocho collares programados próximos al perímetro nominal.","audit_rows":candidates,"excluded_interior_ids":excluded}
    # No publicar geometrías si faltan tramos enteros del perímetro.
    perimeter_length=float(cumulative[-1]);st=np.sort(per["Posicion_perimetro_m"].to_numpy())
    gaps=np.diff(np.r_[st,st[0]+perimeter_length]);max_gap=float(np.max(gaps))
    if max_gap>max(2.25,0.16*perimeter_length):
        return {"ok":False,"reason":f"Diseño perimetral incompleto: tramo sin collar programado de {max_gap:.2f} m. No se interpolan barrenos faltantes desde coordenadas reales.","audit_rows":candidates,"excluded_interior_ids":excluded,"max_gap_m":max_gap}
    needed=["Plan_X2","Plan_Z2","X","Z","X2","Z2","Plan_Y","Plan_Y2","Y","Y2"]
    if per[needed].isna().any().any():
        return {"ok":False,"reason":"Hay barrenos del perímetro sin coordenadas programadas o reales completas.","audit_rows":candidates,"excluded_interior_ids":excluded}
    planned=list(zip(per["Plan_X"],per["Plan_Z"]))
    actual_start=list(zip(per["X"],per["Z"]))
    actual_end=list(zip(per["X2"],per["Z2"]))
    if Polygon is None:
        return {"ok":False,"reason":"Shapely no está disponible para validar polígonos.","audit_rows":candidates}
    polys=[Polygon(points) for points in (planned,actual_start,actual_end)]
    if any((not q.is_valid or q.area<3) for q in polys):
        return {"ok":False,"reason":"El contorno programado o real se cruza, degenera o contiene coordenadas anómalas.","audit_rows":candidates,"excluded_interior_ids":excluded}
    return {"ok":True,"rows":per,"audit_rows":candidates,"excluded_interior_ids":excluded,
            "planned":planned,"actual_start":actual_start,"actual_end":actual_end,
            "area_planned_m2":_polygon_area(planned),"area_start_m2":_polygon_area(actual_start),
            "area_end_m2":_polygon_area(actual_end)}


def _safe_polygon(points):
    if Polygon is None or len(points) < 3:
        return None
    try:
        p = Polygon(points)
        if not p.is_valid:
            p = p.buffer(0)
        return p if not p.is_empty else None
    except Exception:
        return None


def calculate_volumetry(df: pd.DataFrame, nominal_profile: Dict) -> Dict:
    """Valida el contorno perimetral y fija la profundidad común de referencia.

    La profundidad común es la mediana del avance (eje Y) de los barrenos Cut; si no hay Cut se
    usan los Easer y, en último caso, todos los barrenos de frente. La integración de volúmenes
    ya NO se hace aquí: la realiza `calculate_mask_volumetry` con esta profundidad.
    Devuelve `ok`, `perimeter_rows`, `cut_median_advance_y_m` o, si falla, `ok=False` y `reason`.
    """
    prof = _build_ordered_profiles(df, nominal_profile)
    if not prof.get("ok"):
        return {"ok": False, "reason": prof.get("reason", "Contorno no validado.")}

    referencia = df[(df["Tipo"] == "Cut") & (df["Longitud_m"] > 0.1)].copy()

    if referencia.empty:
        referencia = df[
            (df["Tipo"] == "Easer")
            & (pd.to_numeric(df["Longitud_m"], errors="coerce") > 0.1)
        ].copy()

    if referencia.empty:
        referencia = df[
            df["Tipo"].isin(["Bottom", "Easer", "Cut", "Contour"])
            & (pd.to_numeric(df["Longitud_m"], errors="coerce") > 0.1)
        ].copy()

    if referencia.empty:
        return {
            "ok": False,
            "reason": "No existen barrenos de frente válidos para determinar la profundidad común.",
        }

    Ly = float(pd.to_numeric(referencia["Avance_Y_m"], errors="coerce").abs().median())

    per = prof["rows"].copy()
    per["Avance_Y_abs_m"] = (per["Y2"] - per["Y"]).abs()

    return {
        "ok": True,
        "perimeter_rows": per,
        "cut_median_advance_y_m": Ly,
    }


def _mask_triangles(pts: np.ndarray, planned: bool, nominal_buf, max_edge_m: float):
    """Triangula una nube de puntos y devuelve (triángulos válidos (k,3,2), unión Shapely o None).

    Versión vectorizada: todo el filtrado (validez, área mínima, arista máxima y centroide
    dentro del nominal ampliado) se hace en bloque con Shapely 2, sin crear un polígono ni un
    buffer por triángulo. El resultado es idéntico al del bucle anterior.
    """
    import shapely
    from matplotlib.tri import Triangulation

    tri = Triangulation(pts[:, 0], pts[:, 1]).triangles
    T = pts[tri]                                    # (n, 3, 2)
    polys = shapely.polygons(T)
    keep = shapely.is_valid(polys) & (shapely.area(polys) >= 1e-7)
    if not planned:
        # El programado no se recorta contra el nominal; el real descarta aristas largas
        # y triángulos cuyo centroide queda fuera del nominal + 1.5 m.
        edges_max = np.linalg.norm(T - np.roll(T, 1, axis=1), axis=2).max(axis=1)
        keep &= (edges_max <= max_edge_m) & shapely.covers(nominal_buf, shapely.centroid(polys))
    if not keep.any():
        return T[:0], None
    good = polys[keep]
    # Los triángulos de una Delaunay forman una cobertura (sin solapes y con aristas
    # compartidas), así que la unión de cobertura es ~10x más rápida que union_all.
    # Se valida que el área unida iguale la suma de áreas; si no, se usa union_all.
    union = None
    if hasattr(shapely, "coverage_union_all"):
        try:
            union = shapely.coverage_union_all(good)
            if union.is_empty or abs(union.area - float(shapely.area(good).sum())) > 1e-9:
                union = None
        except Exception:
            union = None
    if union is None:
        union = shapely.union_all(good)
    # normalize deja anillo con orientación y vértice inicial deterministas.
    return T[keep], shapely.normalize(union)


def build_experimental_masks(df: pd.DataFrame, nominal_profile: Dict, max_edge_m: float = 2.0) -> Dict:
    """Triangulación exploratoria de collares; NO valida perímetros ni calcula volúmenes.

    El plano programado conserva su triangulación sin recorte nominal. La nube real
    se triangula por separado; se rechazan triángulos largos y los alejados del
    perfil nominal. Nunca se sustituyen coordenadas programadas por reales.
    """
    nominal = _safe_polygon((nominal_profile or {}).get("polygon", []))
    if nominal is None or nominal.area < 1:
        return {"ok": False, "reason": "No existe perfil nominal válido en el ZDA."}
    nominal_buf = nominal.buffer(1.5)               # una sola vez (antes: uno por triángulo)
    result = {"ok": True, "nominal": list(nominal.exterior.coords), "masks": {}}
    for key, xc, zc, planned in (("programado", "Plan_X", "Plan_Z", True),
                                  ("real", "X", "Z", False)):
        rows = df[df[[xc, zc]].notna().all(axis=1)].copy()
        if planned:
            rows = rows[~rows["Extra"].fillna(False)]
        rows = rows[np.isfinite(rows[xc].to_numpy(dtype=float)) & np.isfinite(rows[zc].to_numpy(dtype=float))]
        rows = rows[(rows[xc].abs() < 100) & (rows[zc].abs() < 100)]
        # La triangulación usa ubicaciones únicas; el gráfico y la auditoría
        # conservan TODOS los ID coincidentes (p. ej. Cut / Reaming colocalizados).
        n_barrenos = len(rows)
        grouped = rows.groupby([xc, zc], sort=False, dropna=False)["ID"].agg(
            lambda values: ", ".join(str(v) for v in values)
        ).reset_index()
        pts = grouped[[xc, zc]].to_numpy(dtype=float)
        entry = {"points": [(float(x), float(z)) for x, z in pts],
                 "ids": grouped["ID"].tolist(), "n_barrenos": n_barrenos,
                 "triangles": [], "boundary": []}
        if len(pts) >= 3:
            try:
                # La envolvente planificada incluye sus puntos exteriores (sin recorte nominal).
                kept, union = _mask_triangles(pts, planned, nominal_buf, max_edge_m)
                entry["triangles"] = [[(float(x), float(z)) for x, z in tri] for tri in kept]
                if union is not None:
                    if union.geom_type == "Polygon":
                        entry["boundary"] = [(float(x), float(z)) for x, z in union.exterior.coords]
                    entry["mesh_area_m2"] = float(union.area)
            except (ValueError, RuntimeError) as exc:
                entry["warning"] = f"No se pudo triangular: {exc}"
        result["masks"][key] = entry
    return result


def calculate_mask_volumetry(df: pd.DataFrame, nominal_profile: Dict,
                              reference: Dict, n_sections: int = 41,
                              max_edge_m: float = 2.0) -> Dict:
    """Comparación EXPERIMENTAL por máscaras a profundidad longitudinal común.

    No extrapola barrenos reales cortos. Re-triangula cada sección y compara
    uniones poligonales completas, no solamente sus envolventes exteriores.
    Los resultados se mantienen separados de la volumetría vigente.
    """
    if not reference.get("ok"):
        return {"ok": False, "reason": "Sin profundidad de referencia validada."}
    depth = float(reference.get("cut_median_advance_y_m", 0))
    if not np.isfinite(depth) or depth <= 0.1:
        return {"ok": False, "reason": "Profundidad longitudinal no válida."}
    nominal = _safe_polygon((nominal_profile or {}).get("polygon", []))
    if nominal is None:
        return {"ok": False, "reason": "No hay perfil nominal para controlar los triángulos reales."}

    nominal_buf = nominal.buffer(1.5)               # una sola vez (antes: uno por triángulo)
    prepared = {}
    for key, planned in (("programado", True), ("real", False)):
        cols = (["Plan_X", "Plan_Y", "Plan_Z", "Plan_X2", "Plan_Y2", "Plan_Z2"]
                if planned else ["X", "Y", "Z", "X2", "Y2", "Z2"])
        rows = df.copy()
        if planned:
            rows = rows[~rows["Extra"].fillna(False)]
        vals = rows[cols].apply(pd.to_numeric, errors="coerce").to_numpy(dtype=float)
        valid = np.isfinite(vals).all(axis=1) & (np.abs(vals) < 100).all(axis=1)
        rows, vals = rows.loc[valid].copy(), vals[valid]
        if len(vals) < 3:
            return {"ok": False, "reason": f"Insuficientes coordenadas {key}."}
        prepared[key] = (rows, vals)
    # Cantidades independientes de la profundidad: se calculan una vez, no en cada sección.
    adv_all = {k: np.abs(v[1][:, 4] - v[1][:, 1]) for k, v in prepared.items()}
    len_real = pd.to_numeric(prepared["real"][0]["Longitud_m"], errors="coerce").to_numpy(dtype=float)

    def section(key, distance):
        rows, vals = prepared[key]
        advance = adv_all[key]
        valid = advance >= distance - 1e-6
        # Un barreno corto no puede proyectarse artificialmente hasta el fondo.
        if key == "real":
            valid &= np.isfinite(len_real) & (len_real >= distance - 1e-6)
        valid &= advance > 1e-5
        vals = vals[valid]
        ids = rows.loc[valid, "ID"].astype(str).tolist()
        n_total, n_barrenos = int(len(rows)), int(len(vals))
        if len(vals) < 3:
            return None, {"points": [], "ids": [], "triangles": [], "boundary": [], "area_m2": None,
                          "n_barrenos": n_barrenos, "n_total": n_total}
        fraction = np.clip(distance / np.abs(vals[:, 4] - vals[:, 1]), 0, 1)
        pts = np.column_stack((vals[:, 0] + fraction * (vals[:, 3] - vals[:, 0]),
                               vals[:, 2] + fraction * (vals[:, 5] - vals[:, 2])))
        # Deduplicación estable: posiciones coincidentes (p. ej. Reaming sobre un Cut)
        # se funden en un punto, pero conservando TODOS los IDs ("50, 89").
        _, first, inv = np.unique(np.round(pts, 5), axis=0, return_index=True, return_inverse=True)
        inv = np.asarray(inv).ravel()
        orden = np.argsort(first)
        ids = [", ".join(ids[j] for j in np.where(inv == g)[0]) for g in orden]
        pts = pts[first[orden]]
        entry = {"points": pts.tolist(), "ids": ids, "triangles": [], "boundary": [], "area_m2": None,
                 "n_barrenos": n_barrenos, "n_total": n_total}
        if len(pts) < 3:
            return None, entry
        try:
            # Programado: sin recorte nominal (se conserva la envolvente planificada).
            kept, geom = _mask_triangles(pts, key == "programado", nominal_buf, max_edge_m)
            entry["triangles"] = kept.tolist()
            if geom is None:
                return None, entry
            entry["area_m2"] = float(geom.area)
            if geom.geom_type == "Polygon":
                entry["boundary"] = list(geom.exterior.coords)
            return geom, entry
        except (ValueError, RuntimeError, FloatingPointError) as exc:
            entry["warning"] = str(exc)
            return None, entry

    distances = np.linspace(0, depth, max(5, int(n_sections)))
    sections = []
    for distance in distances:
        gp, ep = section("programado", float(distance))
        gr, er = section("real", float(distance))
        record = {"depth_m": float(distance), "programado": ep, "real": er,
                  "outside_m2": None, "not_covered_m2": None}
        if gp is not None and gr is not None:
            record["outside_m2"] = float(gr.difference(gp).area)
            record["not_covered_m2"] = float(gp.difference(gr).area)
        sections.append(record)
    # No integrar a través de huecos de información: exigir todas las secciones.
    keys = {"programado_m3": ("programado", "area_m2"),
            "real_m3": ("real", "area_m2"),
            "outside_m3": (None, "outside_m2"),
            "not_covered_m3": (None, "not_covered_m2")}
    result = {"ok": True, "experimental": True, "depth_m": depth,
              "sections": sections, "n_sections": len(sections), "max_edge_m": max_edge_m}
    for target, (group, field) in keys.items():
        values = [row[group][field] if group else row[field] for row in sections]
        result[target] = (float(_trapz(values, distances))
                          if all(v is not None and np.isfinite(v) for v in values) else None)
    if result["programado_m3"] is not None and result["real_m3"] is not None:
        result["balance_m3"] = result["real_m3"] - result["programado_m3"]
    else:
        result["balance_m3"] = None
    result["complete"] = all(result[k] is not None for k in keys)
    result["warning"] = ("Cálculo exploratorio; verificar la conectividad, los barrenos cortos "
                         "y el límite exterior antes de usar estos volúmenes como indicadores.")
    return result


def process_zda_bytes(raw: bytes, filename: str = "archivo.zda") -> Dict:
    with zipfile.ZipFile(BytesIO(raw)) as z:
        names = z.namelist()
        txt_name = next((n for n in names if re.match(r"round-.*\.txt$", n, re.I) and "hole_comment" not in n.lower()), None)
        boom_name = next((n for n in names if n.lower().endswith("-boom.dat")), None)
        round_dat_name = next((
            n for n in names
            if re.match(r"round-.*\.dat$", n, re.I)
            and not any(k in n.lower() for k in ("-boom", "-counters", "-mwd-"))
        ), None)

        if boom_name is None:
            raise ValueError("El ZDA no contiene boom.dat.")

        meta = parse_round_txt(z.read(txt_name).decode("utf-8", "ignore")) if txt_name else {}
        holes = parse_boom_dat(z.read(boom_name))
        nominal = parse_round_dat_profile(z.read(round_dat_name)) if round_dat_name else {
            "segments": [], "chain": [], "polygon": [], "area_m2": float("nan"), "decoded": False
        }

    plan_n = int(holes["Plan_X"].notna().sum())
    drilled_n = int(holes["X"].notna().sum())
    expected_plan = meta.get("planned_face_holes")
    expected_drilled = meta.get("drilled_holes")
    plan_ok = expected_plan is None or plan_n == int(expected_plan)
    drilled_ok = expected_drilled is None or drilled_n == int(expected_drilled)
    volumetry = calculate_volumetry(holes, nominal) if plan_ok and drilled_ok else {
        "ok": False, "reason": "Conciliación de barrenos incompleta: "
        f"plan {plan_n}/{expected_plan}, ejecutado {drilled_n}/{expected_drilled}."
    }
    masks = build_experimental_masks(holes, nominal)
    mask_vol = calculate_mask_volumetry(holes, nominal, volumetry) if plan_ok and drilled_ok else {
        "ok": False, "complete": False, "reason": volumetry["reason"]
    }

    return {
        "filename": filename,
        "metadata": meta,
        "holes": holes,
        "nominal_profile": nominal,
        "volumetry": volumetry,
        "experimental_masks": masks,
        "mask_volumetry": mask_vol,
        "auditoria_conteos": {"programados_extraidos": plan_n, "ejecutados_extraidos": drilled_n,
                              "programados_declarados": expected_plan, "ejecutados_declarados": expected_drilled,
                              "conciliado": bool(plan_ok and drilled_ok)},
        "internal_files": names,
        "version": VERSION,
    }


def process_zda_file(path: str | Path) -> Dict:
    p = Path(path)
    return process_zda_bytes(p.read_bytes(), p.name)


# ==========================================================
# RESÚMENES POR LOTE (selector de ciclos)
# ==========================================================

def volumen_unico_ciclo(resultado: Dict) -> Optional[Dict]:
    """Única fuente: integral de las mismas secciones que muestran el 2D, 3D y PDF.

    Devuelve None si la volumetría no es completa o falla la identidad geométrica
    (ejecutado − programado = fuera − no cubierto).
    """
    m = (resultado or {}).get("mask_volumetry") or {}
    claves = ("programado_m3", "real_m3", "outside_m3", "not_covered_m3")
    if not m.get("ok") or not m.get("complete"):
        return None
    try:
        valores = {k: float(m[k]) for k in claves}
        if not all(np.isfinite(x) and x >= 0 for x in valores.values()):
            return None
        vp, vr, vf, vn = (valores[k] for k in claves)
        tolerancia = max(0.02, 1e-4 * max(vp, vr, 1.0))
        if abs((vr - vp) - (vf - vn)) > tolerancia:
            return None
        if vp <= 0:
            return None
        return {**valores, "dgt_m3": vf + vn,
                "fuera_pct": 100 * vf / vp,
                "no_cubierto_pct": 100 * vn / vp,
                "dgt_pct": 100 * (vf + vn) / vp}
    except (ValueError, TypeError, KeyError):
        return None


def resumen_eficiencia_desde_path(path_str: str) -> Optional[Dict]:
    """Procesa un ZDA y devuelve solo el resumen liviano (sin secciones ni DataFrames)."""
    try:
        return volumen_unico_ciclo(process_zda_file(path_str))
    except Exception:
        return None          # ZDA defectuoso: el selector lo muestra como "No calculada"


def _resumen_worker(path_str: str):
    return path_str, resumen_eficiencia_desde_path(path_str)


def resumenes_eficiencia_lote(paths, max_workers: Optional[int] = None) -> Dict[str, Optional[Dict]]:
    """Resúmenes de varios ZDA. Usa procesos en paralelo si hay núcleos y archivos suficientes;
    ante cualquier problema con el paralelismo cae al modo secuencial (mismo resultado)."""
    import os
    lista = list(dict.fromkeys(str(p) for p in paths))
    if max_workers is None:
        max_workers = min(4, (os.cpu_count() or 1) - 1)
    if max_workers >= 2 and len(lista) >= 4:
        try:
            import multiprocessing as mp
            from concurrent.futures import ProcessPoolExecutor
            with ProcessPoolExecutor(max_workers=max_workers, mp_context=mp.get_context("spawn")) as ex:
                return dict(ex.map(_resumen_worker, lista, chunksize=1))
        except Exception:
            pass
    return dict(_resumen_worker(p) for p in lista)
