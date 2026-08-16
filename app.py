from __future__ import annotations

from io import BytesIO
from pathlib import Path
import hashlib
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
st.caption("Sandvik iSURE® Round Report Analysis · El Brocal")

st.info(
    "Carga uno o varios reportes PDF de iSURE®. "
    "La aplicación procesa únicamente los archivos nuevos, muestra un gráfico "
    "por ciclo y genera un Excel consolidado."
)


# ==========================================================
# ESTADO DE SESIÓN
# ==========================================================

if "uploader_version" not in st.session_state:
    st.session_state.uploader_version = 0

if "procesados" not in st.session_state:
    st.session_state.procesados = {}


def limpiar_analisis():
    """
    Limpia resultados almacenados y fuerza la creación
    de un file_uploader nuevo y vacío.
    """
    st.session_state.procesados = {}
    st.session_state.uploader_version += 1


# ==========================================================
# ESTILO DEL FILE UPLOADER
# ==========================================================
# Aumenta la visibilidad de los controles del cargador,
# especialmente el botón "+" que aparece al cargar varios PDF.

st.markdown(
    """
    <style>
    [data-testid="stFileUploader"] {
        font-size: 1rem;
    }

    [data-testid="stFileUploader"] button {
        min-width: 48px !important;
        min-height: 48px !important;
        font-size: 1.2rem !important;
        border-radius: 8px !important;
    }

    [data-testid="stFileUploader"] button svg {
        width: 26px !important;
        height: 26px !important;
    }

    [data-testid="stFileUploaderFile"] {
        margin-right: 6px !important;
        margin-bottom: 6px !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ==========================================================
# FUNCIONES AUXILIARES
# ==========================================================

def hash_archivo(uploaded_file) -> str:
    """
    Identificador único por contenido.
    Permite evitar reprocesar un PDF ya analizado.
    """
    contenido = uploaded_file.getvalue()
    return hashlib.sha256(contenido).hexdigest()


def crear_excel(
    df_reportes: pd.DataFrame,
    df_resumen: pd.DataFrame,
    df_detalle: pd.DataFrame,
    df_validacion: pd.DataFrame,
    df_extras: pd.DataFrame,
) -> bytes:
    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_reportes.to_excel(writer, sheet_name="Resumen_Reportes", index=False)
        df_resumen.to_excel(writer, sheet_name="Resumen_Ciclos", index=False)
        df_detalle.to_excel(writer, sheet_name="Detalle_Barrenos", index=False)
        df_validacion.to_excel(writer, sheet_name="Validacion", index=False)
        df_extras.to_excel(writer, sheet_name="Barrenos_Extra", index=False)

    buffer.seek(0)
    wb = load_workbook(buffer)

    fill_header = PatternFill(fill_type="solid", fgColor="1F4E78")
    font_header = Font(color="FFFFFF", bold=True)

    for ws in wb.worksheets:
        ws.freeze_panes = "A2"

        if ws.max_row >= 1 and ws.max_column >= 1:
            ws.auto_filter.ref = ws.dimensions

        for cell in ws[1]:
            cell.fill = fill_header
            cell.font = font_header
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for columna in ws.columns:
            letra = columna[0].column_letter
            max_length = 0

            for cell in columna:
                if cell.value is None:
                    continue
                max_length = max(max_length, len(str(cell.value)))

            ws.column_dimensions[letra].width = min(max_length + 2, 35)

    salida = BytesIO()
    wb.save(salida)
    salida.seek(0)
    return salida.getvalue()


def figura_a_png(fig) -> bytes:
    buffer = BytesIO()
    fig.savefig(buffer, format="png", dpi=250, bbox_inches="tight")
    buffer.seek(0)
    return buffer.getvalue()


def crear_zip_graficos(graficos) -> bytes:
    buffer = BytesIO()

    with zipfile.ZipFile(buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zip_file:
        for nombre, contenido in graficos:
            zip_file.writestr(nombre, contenido)

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
        texto = texto.replace(origen, destino)

    return texto


def guardar_resultado_en_cache(archivo, clave_archivo: str):
    """
    Procesa un PDF una sola vez y guarda sus resultados
    en st.session_state.
    """
    temp_path = None

    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp:
            temp.write(archivo.getbuffer())
            temp_path = Path(temp.name)

        resultado = procesar_pdf(temp_path, nombre_archivo=archivo.name)

        fig = resultado["fig"]
        png_bytes = figura_a_png(fig)
        plt.close(fig)

        metadata = resultado["metadata"]

        fecha_segura = nombre_seguro(metadata.get("Fecha_Inicio") or "sin_fecha")
        jumbo_seguro = nombre_seguro(metadata.get("Jumbo") or "JUMBO")
        ciclo_seguro = nombre_seguro(metadata.get("Ciclo") or "sin_ciclo")

        nombre_png = (
            f"{jumbo_seguro}_"
            f"Ciclo_{ciclo_seguro}_"
            f"{fecha_segura}.png"
        )

        st.session_state.procesados[clave_archivo] = {
            "nombre_archivo": archivo.name,
            "metadata": metadata,
            "detalle": resultado["detalle"],
            "validacion": resultado["validacion"],
            "resumen_ciclo": resultado["resumen_ciclo"],
            "resumen_reporte": resultado["resumen_reporte"],
            "extras": resultado["extras"],
            "png_bytes": png_bytes,
            "nombre_png": nombre_png,
            "error": None,
        }

    except Exception as exc:
        st.session_state.procesados[clave_archivo] = {
            "nombre_archivo": archivo.name,
            "error": str(exc),
        }

    finally:
        if temp_path and temp_path.exists():
            temp_path.unlink(missing_ok=True)


# ==========================================================
# CONTROLES DE CARGA / LIMPIEZA
# ==========================================================

st.subheader("Reportes")

col_uploader, col_limpiar = st.columns([5, 1])

with col_uploader:
    archivos = st.file_uploader(
        "Seleccionar o agregar reportes PDF",
        type=["pdf"],
        accept_multiple_files=True,
        key=f"uploader_pdfs_{st.session_state.uploader_version}",
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
    st.write(f"**Archivos seleccionados:** {len(archivos)}")

    archivos_actuales = []
    for archivo in archivos:
        clave = hash_archivo(archivo)
        archivos_actuales.append((clave, archivo))

    nuevos = [
        (clave, archivo)
        for clave, archivo in archivos_actuales
        if clave not in st.session_state.procesados
    ]

    if nuevos:
        progreso = st.progress(0)

        for i, (clave, archivo) in enumerate(nuevos, start=1):
            with st.spinner(f"Procesando archivo nuevo: {archivo.name}"):
                guardar_resultado_en_cache(archivo, clave)

            progreso.progress(i / len(nuevos))

        progreso.empty()

    # ======================================================
    # MOSTRAR UNA SOLA LÍNEA PLEGABLE POR PDF
    # ======================================================

    resultados_validos = []
    errores_actuales = []

    for clave, archivo in archivos_actuales:
        resultado = st.session_state.procesados.get(clave)

        if not resultado:
            continue

        if resultado.get("error"):
            errores_actuales.append(
                {
                    "Archivo_PDF": archivo.name,
                    "Error": resultado["error"],
                }
            )
            continue

        resultados_validos.append(resultado)

        metadata = resultado["metadata"]
        validacion = resultado["validacion"]
        extras = resultado["extras"]
        png_bytes = resultado["png_bytes"]
        nombre_png = resultado["nombre_png"]

        titulo_expander = (
            f"{metadata['Jumbo']} | "
            f"Ciclo {metadata['Ciclo']} | "
            f"{metadata['Fecha_Inicio']} | "
            f"{archivo.name}"
        )

        with st.expander(titulo_expander, expanded=False):
            c1, c2, c3, c4 = st.columns(4)

            c1.metric("Jumbo", metadata["Jumbo"] or "-")
            c2.metric("Nº de serie", metadata["Numero_Serie"] or "-")
            c3.metric("Ciclo", metadata["Ciclo"] or "-")
            c4.metric(
                "Metros perforados",
                (
                    f"{metadata['Metros_Perforados']:.2f} m"
                    if metadata["Metros_Perforados"] is not None
                    else "-"
                ),
            )

            st.image(png_bytes, use_container_width=True)

            st.download_button(
                label="Descargar gráfico PNG",
                data=png_bytes,
                file_name=nombre_png,
                mime="image/png",
                key=f"grafico_{clave}",
                on_click="ignore",
            )

            st.markdown("#### Validación")

            st.dataframe(
                validacion[
                    ["Tipo", "Esperado", "Encontrado", "Diferencia", "Estado"]
                ],
                use_container_width=True,
                hide_index=True,
            )

            if (validacion["Estado"] == "REVISAR").any():
                st.warning(
                    "Este ciclo presenta diferencias entre el resumen "
                    "'TIPOS DE BARRENO' y los barrenos extraídos. "
                    "Revísalo antes de usar sus resultados."
                )
            else:
                st.success(
                    "La extracción reconcilia con el resumen "
                    "'TIPOS DE BARRENO'."
                )

            if not extras.empty:
                st.markdown("#### Barrenos extra")

                columnas_extras = [
                    col
                    for col in ["ID", "Tipo", "Longitud_roca_m", "Beta_grados"]
                    if col in extras.columns
                ]

                st.dataframe(
                    extras[columnas_extras],
                    use_container_width=True,
                    hide_index=True,
                )

    # ======================================================
    # CONSOLIDADO SOLO CON LOS PDF ACTUALMENTE SELECCIONADOS
    # ======================================================

    if resultados_validos:
        df_detalle = pd.concat(
            [r["detalle"] for r in resultados_validos],
            ignore_index=True,
        )

        df_resumen = pd.concat(
            [r["resumen_ciclo"] for r in resultados_validos],
            ignore_index=True,
        )

        df_reportes = pd.DataFrame(
            [r["resumen_reporte"] for r in resultados_validos]
        )

        df_validacion = pd.concat(
            [r["validacion"] for r in resultados_validos],
            ignore_index=True,
        )

        extras_validos = [
            r["extras"]
            for r in resultados_validos
            if not r["extras"].empty
        ]

        df_extras = (
            pd.concat(extras_validos, ignore_index=True)
            if extras_validos
            else pd.DataFrame(columns=df_detalle.columns)
        )

        st.divider()
        st.header("Consolidado")

        st.markdown("#### Resumen de reportes")
        st.dataframe(
            df_reportes,
            use_container_width=True,
            hide_index=True,
        )

        st.markdown("#### Resumen por ciclo y tipo")
        st.dataframe(
            df_resumen,
            use_container_width=True,
            hide_index=True,
        )

        excel_bytes = crear_excel(
            df_reportes,
            df_resumen,
            df_detalle,
            df_validacion,
            df_extras,
        )

        graficos_zip = [
            (r["nombre_png"], r["png_bytes"])
            for r in resultados_validos
        ]

        col_excel, col_graficos = st.columns(2)

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
            zip_bytes = crear_zip_graficos(graficos_zip)

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
        st.subheader("Archivos con error")

        st.dataframe(
            pd.DataFrame(errores_actuales),
            use_container_width=True,
            hide_index=True,
        )

else:
    st.write("Selecciona uno o varios PDF para iniciar el análisis.")
