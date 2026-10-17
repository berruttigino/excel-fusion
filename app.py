"""
ExcelFusion - versión web (Streamlit)
=====================================
Subí varios Excel/CSV, elegí cómo agrupar el resumen y descargá un reporte listo.

Para correrla en tu PC:
    pip install -r requirements.txt
    streamlit run app.py
"""
import io
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from consolidar_excels import EXTENSIONES, armar_resumen, formatear_hoja, leer_archivo


def cargar(archivos):
    """Lee los archivos subidos. Devuelve (consolidado | None, registro)."""
    registro, frames = [], []
    with tempfile.TemporaryDirectory() as tmp:
        for i, archivo in enumerate(archivos):
            ruta = Path(tmp) / Path(archivo.name).name
            if ruta.exists():  # dos archivos con el mismo nombre
                ruta = Path(tmp) / f"{ruta.stem}_{i}{ruta.suffix}"
            ruta.write_bytes(archivo.getvalue())
            frames += leer_archivo(ruta, registro)

    if not frames:
        return None, pd.DataFrame(registro, columns=["archivo", "hoja", "estado", "detalle"])

    consolidado = pd.concat(frames, ignore_index=True, sort=False)
    for col in consolidado.columns:
        if col in ("archivo_origen", "hoja_origen"):
            continue
        faltan = sorted({f["archivo_origen"].iat[0] for f in frames if col not in f.columns})
        if faltan:
            registro.append(("-", "-", "AVISO", f"Columna '{col}' falta en: {', '.join(faltan)}"))
    return consolidado, pd.DataFrame(registro, columns=["archivo", "hoja", "estado", "detalle"])


def exportar_excel(consolidado, resumen, log) -> bytes:
    """Arma el reporte (3 hojas) en memoria y devuelve los bytes del .xlsx."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl", datetime_format="DD/MM/YYYY") as xl:
        consolidado.to_excel(xl, sheet_name="Consolidado", index=False)
        resumen.to_excel(xl, sheet_name="Resumen", index=False)
        log.to_excel(xl, sheet_name="Registro", index=False)
        for hoja in xl.book.worksheets:
            formatear_hoja(hoja)
    return buffer.getvalue()


def main():
    st.set_page_config(page_title="ExcelFusion", page_icon="📊", layout="wide")
    st.title("📊 ExcelFusion")
    st.write(
        "Merge messy Excel and CSV files into one clean report. "
        "*Uní tus Excel y CSV desordenados en un solo reporte.*"
    )

    subidos = st.file_uploader(
        "Upload your files / Subí tus archivos",
        type=[e.lstrip(".") for e in sorted(EXTENSIONES)],
        accept_multiple_files=True,
    )
    if not subidos:
        st.info("Upload two or more files to start. / Subí dos o más archivos para empezar.")
        return

    consolidado, log = cargar(subidos)
    if consolidado is None:
        st.error("No file could be read. / No se pudo leer ningún archivo.")
        st.dataframe(log)
        return

    # Columnas de texto disponibles para agrupar el resumen
    opciones = ["archivo_origen"] + [
        c for c in consolidado.select_dtypes(exclude=["number", "datetime"]).columns
        if c not in ("archivo_origen", "hoja_origen")
    ]
    agrupar = st.selectbox("Group summary by / Agrupar resumen por", opciones)
    resumen = armar_resumen(consolidado, agrupar)

    c1, c2, c3 = st.columns(3)
    c1.metric("Files / Archivos", len(subidos))
    c2.metric("Rows / Filas", len(consolidado))
    c3.metric("Columns / Columnas", len(consolidado.columns))

    st.download_button(
        "⬇️ Download report (.xlsx) / Descargar reporte",
        data=exportar_excel(consolidado, resumen, log),
        file_name="reporte_consolidado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    tab1, tab2, tab3 = st.tabs(["Resumen", "Consolidado", "Registro"])
    with tab1:
        st.dataframe(resumen)
    with tab2:
        st.dataframe(consolidado)
    with tab3:
        st.dataframe(log)


if __name__ == "__main__":
    main()
