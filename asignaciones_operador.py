"""Asignación manual de operadores por ciclo (round), con persistencia en disco.

El ZDA guarda el operador como texto libre (`tunnel_id`); cuando no viene, el ciclo queda
"SIN DATO". Este módulo permite asignarlo a mano y que el dato se aplique en toda la app
(filtros, gráficos, Excel, "Resultados por archivo") sin tocar los archivos ZDA.

No depende de Streamlit: toda la lógica es pura y se puede probar por separado.

Esquema del archivo JSON (`operadores_asignados.json`):
    {"version": 1,
     "operadores": ["Nombre extra", ...],
     "asignaciones": {"JUMB002|371|26/08/2026|10:14:31":
                        {"operador": "Nilton Celis", "jumbo": "JUMB002", "ciclo": 371,
                         "fecha": "26/08/2026", "hora": "10:14:31", "asignado": "2026-09-19T02:30:00"}}}
"""
from __future__ import annotations

import csv
import io
import json
import os
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import numpy as np
import pandas as pd

ESQUEMA = 1
NOMBRE_ARCHIVO = "operadores_asignados.json"
_SIN_DATO = {"", "SIN DATO", "NONE", "NAN", "NULL"}
COLUMNAS_CSV = ["Jumbo", "Ciclo", "Fecha_Inicio", "Hora_Inicio", "Operador"]


# ----------------------------------------------------------------------------------------
# Utilidades básicas
# ----------------------------------------------------------------------------------------
def es_sin_operador(valor) -> bool:
    """True si el valor equivale a "no hay operador"."""
    if valor is None:
        return True
    try:
        if pd.isna(valor):
            return True
    except (TypeError, ValueError):
        pass
    return str(valor).strip().upper() in _SIN_DATO


def normalizar_nombre(valor) -> Optional[str]:
    """Nombre limpio (espacios colapsados) o None si está vacío."""
    if es_sin_operador(valor):
        return None
    return " ".join(str(valor).split())


def clave_ciclo(rep: Dict) -> str:
    """Identificador estable de un ciclo: Jumbo | Ciclo | Fecha | Hora de inicio."""
    def texto(campo):
        v = rep.get(campo)
        if v is None or (isinstance(v, float) and np.isnan(v)):
            return ""
        if campo == "Ciclo":
            try:
                return str(int(v))
            except (TypeError, ValueError):
                pass
        return str(v).strip()

    return "|".join(texto(c) for c in ("Jumbo", "Ciclo", "Fecha_Inicio", "Hora_Inicio"))


# ----------------------------------------------------------------------------------------
# Persistencia
# ----------------------------------------------------------------------------------------
def _vacio() -> Dict:
    return {"version": ESQUEMA, "operadores": [], "asignaciones": {}}


def directorio_datos() -> Path:
    """Carpeta donde se guarda el archivo: $EBR_DATA_DIR o ~/.ebr_drill_analytics."""
    base = os.environ.get("EBR_DATA_DIR")
    return Path(base).expanduser() if base else Path.home() / ".ebr_drill_analytics"


def ruta_asignaciones() -> Path:
    return directorio_datos() / NOMBRE_ARCHIVO


def cargar(ruta: Optional[Path] = None) -> Dict:
    """Lee el JSON. Si no existe o está dañado devuelve una estructura vacía
    (un archivo dañado se conserva como .bak para no perder nada)."""
    ruta = Path(ruta) if ruta else ruta_asignaciones()
    if not ruta.exists():
        return _vacio()
    try:
        data = json.loads(ruta.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            raise ValueError("formato inesperado")
    except (OSError, ValueError):
        try:
            ruta.replace(ruta.with_suffix(".json.bak"))
        except OSError:
            pass
        return _vacio()
    data.setdefault("version", ESQUEMA)
    data.setdefault("operadores", [])
    data.setdefault("asignaciones", {})
    return data


def guardar(data: Dict, ruta: Optional[Path] = None) -> Path:
    """Guarda de forma atómica. Si la carpeta no es escribible usa la carpeta temporal.
    Devuelve la ruta realmente usada."""
    destino = Path(ruta) if ruta else ruta_asignaciones()
    texto = json.dumps(data, ensure_ascii=False, indent=1, sort_keys=True)
    for candidato in (destino, Path(tempfile.gettempdir()) / "ebr_drill_analytics" / NOMBRE_ARCHIVO):
        try:
            candidato.parent.mkdir(parents=True, exist_ok=True)
            tmp = candidato.with_suffix(".json.tmp")
            tmp.write_text(texto, encoding="utf-8")
            tmp.replace(candidato)
            return candidato
        except OSError:
            continue
    raise OSError("No se pudo guardar el archivo de operadores asignados.")


# ----------------------------------------------------------------------------------------
# Edición del diccionario de asignaciones
# ----------------------------------------------------------------------------------------
def poner(data: Dict, rep: Dict, operador: str) -> bool:
    """Asigna (o reemplaza) el operador de un ciclo. Devuelve True si hubo cambio."""
    nombre = normalizar_nombre(operador)
    if not nombre:
        return quitar(data, clave_ciclo(rep))
    clave = clave_ciclo(rep)
    previo = data["asignaciones"].get(clave, {}).get("operador")
    if previo == nombre:
        return False
    data["asignaciones"][clave] = {
        "operador": nombre,
        "jumbo": str(rep.get("Jumbo") or ""),
        "ciclo": rep.get("Ciclo"),
        "fecha": str(rep.get("Fecha_Inicio") or ""),
        "hora": str(rep.get("Hora_Inicio") or ""),
        "asignado": datetime.now().isoformat(timespec="seconds"),
    }
    agregar_operador(data, nombre)
    return True


def quitar(data: Dict, clave: str) -> bool:
    return data["asignaciones"].pop(clave, None) is not None


def agregar_operador(data: Dict, nombre: str) -> bool:
    """Agrega un nombre a la lista de operadores conocidos (sin duplicar, sin distinguir mayúsculas)."""
    nombre = normalizar_nombre(nombre)
    if not nombre:
        return False
    existentes = {n.upper() for n in data["operadores"]}
    if nombre.upper() in existentes:
        return False
    data["operadores"].append(nombre)
    data["operadores"].sort(key=str.upper)
    return True


def lista_operadores(data: Dict, extras: Iterable[str] = ()) -> List[str]:
    """Operadores conocidos: los guardados + `extras` (catálogo del ZDA, detectados...)."""
    vistos: Dict[str, str] = {}
    for n in list(extras) + list(data.get("operadores", [])) + [
        a.get("operador") for a in data.get("asignaciones", {}).values()
    ]:
        nombre = normalizar_nombre(n)
        if nombre and nombre.upper() not in vistos:
            vistos[nombre.upper()] = nombre
    return sorted(vistos.values(), key=str.upper)


# ----------------------------------------------------------------------------------------
# Aplicación sobre los resultados ya procesados
# ----------------------------------------------------------------------------------------
def aplicar_a_resultados(resultados: Iterable[Dict], data: Dict) -> Dict[str, int]:
    """Aplica las asignaciones a `resumen_reporte` de cada resultado (modifica en sitio).

    Es idempotente: guarda el operador detectado por el ZDA la primera vez y lo restaura si
    la asignación se elimina. Actualiza Operador_ZDA, Operador y Fuente_Operador, que son los
    campos que leen filtros, gráficos, Excel y "Resultados por archivo".
    """
    n_manual = n_restaurados = 0
    asignaciones = data.get("asignaciones", {})
    for r in resultados:
        rep = r.get("resumen_reporte") if isinstance(r, dict) else None
        if not isinstance(rep, dict):
            continue
        if "_Operador_ZDA_Detectado" not in rep:
            rep["_Operador_ZDA_Detectado"] = rep.get("Operador_ZDA")
            rep["_Fuente_Operador_Original"] = rep.get("Fuente_Operador")
        asignado = asignaciones.get(clave_ciclo(rep))
        if asignado and normalizar_nombre(asignado.get("operador")):
            nombre = normalizar_nombre(asignado["operador"])
            rep["Operador_ZDA"] = nombre
            rep["Operador"] = nombre
            rep["Fuente_Operador"] = "Asignado manualmente"
            rep["Operador_Asignado_Manual"] = True
            n_manual += 1
        else:
            if rep.get("Operador_Asignado_Manual"):
                n_restaurados += 1
            detectado = rep["_Operador_ZDA_Detectado"]
            rep["Operador_ZDA"] = detectado
            rep["Operador"] = detectado
            rep["Fuente_Operador"] = rep["_Fuente_Operador_Original"]
            rep["Operador_Asignado_Manual"] = False
    return {"manuales": n_manual, "restaurados": n_restaurados}


# ----------------------------------------------------------------------------------------
# Sugerencias y edición en tabla
# ----------------------------------------------------------------------------------------
def sugerir(df: pd.DataFrame, col_operador: str = "operador", col_fecha: str = "FechaHora",
            col_jumbo: str = "Jumbo", horas_max: float = 6.0) -> pd.Series:
    """Sugiere operador para las filas sin operador: el del ciclo más cercano en el tiempo
    del MISMO jumbo, si está a `horas_max` horas o menos. Devuelve una serie alineada con
    `df` ("" si no hay sugerencia). Es solo una ayuda: nunca se aplica sin confirmar."""
    sugerencia = pd.Series("", index=df.index, dtype=object)
    if df.empty:
        return sugerencia
    fechas = pd.to_datetime(df[col_fecha], errors="coerce")
    conocido = ~df[col_operador].map(es_sin_operador) & fechas.notna()
    pendiente = df[col_operador].map(es_sin_operador) & fechas.notna()
    segundos = fechas.map(lambda x: x.timestamp() if pd.notna(x) else np.nan)   # independiente de la unidad de pandas
    for jumbo, idx in df.groupby(col_jumbo).groups.items():
        idx = pd.Index(idx)
        k_idx = idx[conocido.loc[idx].to_numpy()]
        p_idx = idx[pendiente.loc[idx].to_numpy()]
        if len(k_idx) == 0 or len(p_idx) == 0:
            continue
        k_t = segundos.loc[k_idx].to_numpy(dtype=float)
        orden = np.argsort(k_t)
        k_t, k_nombres = k_t[orden], df.loc[k_idx, col_operador].to_numpy()[orden]
        for i in p_idx:
            t = float(segundos.loc[i])
            pos = int(np.searchsorted(k_t, t))
            candidatos = [j for j in (pos - 1, pos) if 0 <= j < len(k_t)]
            mejor = min(candidatos, key=lambda j: abs(k_t[j] - t))
            if abs(k_t[mejor] - t) / 3600.0 <= horas_max:
                sugerencia.loc[i] = normalizar_nombre(k_nombres[mejor]) or ""
    return sugerencia


def cambios_desde_editor(original: pd.DataFrame, editado: pd.DataFrame,
                         col_clave: str = "clave", col_op: str = "Operador") -> Dict[str, Optional[str]]:
    """Compara la tabla mostrada con la editada y devuelve {clave: operador_nuevo | None}.
    None significa "quitar la asignación manual". Solo incluye las filas modificadas."""
    # Diccionarios (no Series): así los vacíos son siempre None y no NaN (NaN != NaN).
    base = {c: normalizar_nombre(v) for c, v in zip(original[col_clave], original[col_op])}
    cambios: Dict[str, Optional[str]] = {}
    for clave, valor in zip(editado[col_clave], editado[col_op]):
        nuevo = normalizar_nombre(valor)
        if clave in base and nuevo != base[clave]:
            cambios[clave] = nuevo
    return cambios


# ----------------------------------------------------------------------------------------
# CSV (copia de seguridad / compartir entre equipos)
# ----------------------------------------------------------------------------------------
def a_csv(data: Dict) -> bytes:
    """CSV (UTF-8 con BOM, separador coma) con las asignaciones, ordenadas por jumbo y fecha."""
    filas = []
    for a in data.get("asignaciones", {}).values():
        filas.append({"Jumbo": a.get("jumbo", ""), "Ciclo": a.get("ciclo", ""),
                      "Fecha_Inicio": a.get("fecha", ""), "Hora_Inicio": a.get("hora", ""),
                      "Operador": a.get("operador", "")})
    def orden(f):
        try:
            t = datetime.strptime(f"{f['Fecha_Inicio']} {f['Hora_Inicio']}", "%d/%m/%Y %H:%M:%S")
        except ValueError:
            t = datetime.min
        return (f["Jumbo"], t)
    filas.sort(key=orden)
    buf = io.StringIO()
    w = csv.DictWriter(buf, fieldnames=COLUMNAS_CSV)
    w.writeheader()
    w.writerows(filas)
    return buf.getvalue().encode("utf-8-sig")


def importar_csv(data: Dict, contenido: bytes) -> Tuple[int, List[str]]:
    """Importa asignaciones desde un CSV (coma o punto y coma). Devuelve (n_importadas, avisos).
    Las filas ya existentes se reemplazan; las filas incompletas se omiten con un aviso."""
    texto = contenido.decode("utf-8-sig", errors="replace")
    primera = texto.splitlines()[0] if texto.strip() else ""
    delimitador = ";" if primera.count(";") > primera.count(",") else ","
    lector = csv.DictReader(io.StringIO(texto), delimiter=delimitador)
    faltan = [c for c in COLUMNAS_CSV if c not in (lector.fieldnames or [])]
    if faltan:
        return 0, [f"Faltan columnas en el CSV: {', '.join(faltan)}."]
    n, avisos = 0, []
    for i, fila in enumerate(lector, start=2):
        rep = {"Jumbo": (fila.get("Jumbo") or "").strip(), "Ciclo": (fila.get("Ciclo") or "").strip(),
               "Fecha_Inicio": (fila.get("Fecha_Inicio") or "").strip(),
               "Hora_Inicio": (fila.get("Hora_Inicio") or "").strip()}
        nombre = normalizar_nombre(fila.get("Operador"))
        if not (rep["Jumbo"] and rep["Ciclo"] and rep["Fecha_Inicio"] and nombre):
            avisos.append(f"Fila {i} omitida: datos incompletos.")
            continue
        try:
            rep["Ciclo"] = int(rep["Ciclo"])
        except ValueError:
            avisos.append(f"Fila {i} omitida: ciclo no numérico ({rep['Ciclo']}).")
            continue
        poner(data, rep, nombre)
        n += 1
    return n, avisos
