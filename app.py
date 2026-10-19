"""
ExcelFusion - versión web (Streamlit)
=====================================
Subí varios Excel/CSV, mirá el dashboard con KPIs y descargá un reporte listo.

Para correrla en tu PC:
    pip install -r requirements.txt
    streamlit run app.py
"""
import io
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

import marca
from consolidar_excels import EXTENSIONES, armar_resumen, formatear_hoja, leer_archivo
from dashboard import (PERIODOS, calcular_kpis, detectar_columnas, elegir_defaults, filtrar,
                       ranking, serie_tiempo, tabla_kpis)


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


def exportar_excel(consolidado, resumen, log, kpis=None, ranking_df=None) -> bytes:
    """Arma el reporte en memoria y devuelve los bytes del .xlsx.
    Hojas: Consolidado, Resumen, KPIs (si hay datos numéricos) y Registro."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl", datetime_format="DD/MM/YYYY") as xl:
        consolidado.to_excel(xl, sheet_name="Consolidado", index=False)
        resumen.to_excel(xl, sheet_name="Resumen", index=False)
        if kpis is not None:
            kpis.to_excel(xl, sheet_name="KPIs", index=False)
            if ranking_df is not None and len(ranking_df):
                ranking_df.to_excel(xl, sheet_name="KPIs", index=False, startrow=len(kpis) + 3)
        log.to_excel(xl, sheet_name="Registro", index=False)
        for hoja in xl.book.worksheets:
            formatear_hoja(hoja)
            if hoja.title == "KPIs":
                hoja.auto_filter.ref = None  # la hoja tiene dos tablas: sin filtro global
                if ranking_df is not None and len(ranking_df):
                    from openpyxl.styles import Font, PatternFill
                    for celda in hoja[len(kpis) + 4]:  # encabezado de la tabla de ranking
                        if celda.value is not None:
                            celda.font = Font(bold=True)
                            celda.fill = PatternFill("solid", fgColor="D9E1F2")
    return buffer.getvalue()


def numero(valor) -> str:
    """12345.678 -> '12,345.68'."""
    return f"{valor:,.2f}"


def mostrar_dashboard(df, defaults):
    """Pestaña Dashboard: elegís métrica, categoría y fechas, y ves KPIs y gráficos."""
    cols = detectar_columnas(df)
    if not cols["numericas"]:
        st.info("No numeric columns found to build KPIs. / No hay columnas numéricas para calcular KPIs.")
        return

    c1, c2, c3, c4 = st.columns(4)
    metrica = c1.selectbox("Metric / Métrica", cols["numericas"])
    categoria = c2.selectbox("Category / Categoría", cols["categoricas"]) if cols["categoricas"] else None
    fecha = c3.selectbox("Date / Fecha", cols["fechas"]) if cols["fechas"] else None
    freq = defaults["freq"]
    if fecha:
        opciones = list(PERIODOS)
        freq = c4.selectbox("Period / Período", opciones, index=opciones.index(freq),
                            format_func=lambda x: PERIODOS[x])

    valores, desde, hasta = None, None, None
    with st.expander("Filters / Filtros"):
        top = st.slider("Top N", min_value=3, max_value=20, value=10)
        if categoria:
            todos = sorted(df[categoria].dropna().astype(str).unique())
            valores = st.multiselect(f"{categoria}", todos, default=todos)
        if fecha and df[fecha].notna().any():
            minimo, maximo = df[fecha].min().date(), df[fecha].max().date()
            rango = st.date_input("Date range / Rango de fechas", value=(minimo, maximo),
                                  min_value=minimo, max_value=maximo)
            if isinstance(rango, (tuple, list)) and len(rango) == 2 and tuple(rango) != (minimo, maximo):
                desde, hasta = rango

    datos = filtrar(df, categoria, valores, fecha, desde, hasta)
    if datos.empty:
        st.warning("No data for the selected filters. / No hay datos con los filtros elegidos.")
        return

    k = calcular_kpis(datos, metrica, categoria, fecha, freq)
    m = st.columns(5)
    m[0].metric(f"Total {metrica}", numero(k["total"]))
    m[1].metric("Average / Promedio", numero(k["promedio"]))
    m[2].metric("Records / Registros", f"{k['registros']:,}")
    if k["lider"] is not None:
        m[3].metric(f"Top {categoria} ({k['lider_pct']:.1f}%)", k["lider"])
    if k["mejor_periodo"] is not None:
        m[4].metric("Best period / Mejor período", k["mejor_periodo"], help=f"Total: {numero(k['mejor_periodo_valor'])}")

    g1, g2 = st.columns(2)
    if categoria:
        rank = ranking(datos, categoria, metrica, top)
        with g1:
            st.subheader(f"{metrica} / {categoria}")
            st.bar_chart(rank, x=categoria, y=metrica, color=marca.VERDE)
            st.dataframe(rank)
    if fecha:
        serie = serie_tiempo(datos, fecha, metrica, freq)
        if len(serie):
            with g2:
                st.subheader(f"{metrica} / {PERIODOS[freq].lower()}")
                st.line_chart(serie, color=marca.TINTA)
        else:
            with g2:
                st.info("No valid dates to build the time chart. / No hay fechas válidas para el gráfico.")


def contenido():
    """Cargador, KPIs, dashboard y descarga (todo lo que va debajo de la cabecera)."""
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

    # KPIs automáticos para la hoja 'KPIs' del Excel (con las columnas más probables)
    defaults = elegir_defaults(consolidado)
    kpis_df = ranking_df = None
    if defaults["metrica"]:
        kpis_df = tabla_kpis(consolidado, defaults["metrica"], defaults["categoria"],
                             defaults["fecha"], defaults["freq"])
        if defaults["categoria"]:
            ranking_df = ranking(consolidado, defaults["categoria"], defaults["metrica"], top=10)

    c1, c2, c3 = st.columns(3)
    c1.metric("Files / Archivos", len(subidos))
    c2.metric("Rows / Filas", len(consolidado))
    c3.metric("Columns / Columnas", len(consolidado.columns))

    st.download_button(
        "⬇️ Download report (.xlsx) / Descargar reporte",
        data=exportar_excel(consolidado, resumen, log, kpis_df, ranking_df),
        file_name="reporte_consolidado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    tab0, tab1, tab2, tab3 = st.tabs(["Dashboard", "Resumen", "Consolidado", "Registro"])
    with tab0:
        mostrar_dashboard(consolidado, defaults)
    with tab1:
        st.dataframe(resumen)
    with tab2:
        st.dataframe(consolidado)
    with tab3:
        st.dataframe(log)


def main():
    st.set_page_config(page_title="ExcelFusion: many files, one report", page_icon=marca.icono_pagina(), layout="wide")
    marca.aplicar_estilo(st)
    marca.cabecera(st)
    contenido()
    marca.pie(st)


if __name__ == "__main__":
    main()
