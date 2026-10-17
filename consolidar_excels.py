#!/usr/bin/env python3
"""
ExcelFusion - Consolidador de archivos Excel/CSV
==================================
Lee todos los archivos .xlsx/.xls/.csv de una carpeta, unifica los nombres de
columnas, los junta en una sola tabla y genera un reporte en Excel con:

  - Consolidado : todos los datos juntos (con archivo y hoja de origen)
  - Resumen     : totales por grupo (por defecto, por archivo)
  - Registro    : qué se leyó, qué falló y qué columnas no coinciden

Uso:
    python consolidar_excels.py carpeta_con_archivos
    python consolidar_excels.py carpeta_con_archivos -o reporte.xlsx --agrupar producto
"""
import argparse
import sys
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd

EXTENSIONES = {".xlsx", ".xlsm", ".xls", ".csv"}

# Nombres distintos que significan lo mismo. Agregá los que veas en tus archivos.
ALIAS = {
    "importe": "monto",
    "total": "monto",
    "cant": "cantidad",
    "qty": "cantidad",
    "articulo": "producto",
    "item": "producto",
    "fecha_venta": "fecha",
}


def normalizar(nombre: str) -> str:
    """'Fecha de Venta ' -> 'fecha_de_venta' (sin tildes, minúsculas, con guiones bajos)."""
    texto = unicodedata.normalize("NFKD", str(nombre)).encode("ascii", "ignore").decode()
    texto = "_".join(texto.strip().lower().split())
    return ALIAS.get(texto, texto)


def leer_archivo(ruta: Path, registro: list) -> list:
    """Devuelve una lista de DataFrames (uno por hoja) o [] si el archivo falla."""
    try:
        if ruta.suffix.lower() == ".csv":
            hojas = {"csv": pd.read_csv(ruta, sep=None, engine="python")}
        else:
            hojas = pd.read_excel(ruta, sheet_name=None)
    except Exception as error:  # archivo corrupto, protegido, etc.
        registro.append((ruta.name, "-", "ERROR", str(error)[:150]))
        return []

    frames = []
    for hoja, df in hojas.items():
        df = df.dropna(how="all").dropna(axis=1, how="all")
        if df.empty:
            registro.append((ruta.name, hoja, "OMITIDA", "Hoja vacía"))
            continue
        df.columns = [normalizar(c) for c in df.columns]
        df = df.loc[:, ~df.columns.duplicated()]  # columnas repetidas tras normalizar
        for col in df.columns:
            if col == "fecha" or col.startswith("fecha_"):
                df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=True)
        df.insert(0, "archivo_origen", ruta.name)
        df.insert(1, "hoja_origen", hoja)
        frames.append(df)
        registro.append((ruta.name, hoja, "OK", f"{len(df)} filas"))
    return frames


def armar_resumen(df: pd.DataFrame, agrupar: str | None) -> pd.DataFrame:
    numericas = df.select_dtypes("number").columns.tolist()
    clave = agrupar if agrupar in df.columns else "archivo_origen"
    resumen = df.groupby(clave, dropna=False).agg(
        filas=(clave, "size"), **{c: (c, "sum") for c in numericas}
    )
    resumen.loc["TOTAL"] = resumen.sum(numeric_only=True)
    return resumen.reset_index()


def formatear_hoja(hoja) -> None:
    """Encabezado en negrita, filtros, fila fija y ancho de columnas razonable."""
    from openpyxl.styles import Font, PatternFill

    relleno = PatternFill("solid", fgColor="D9E1F2")
    for celda in hoja[1]:
        celda.font = Font(bold=True)
        celda.fill = relleno
    hoja.freeze_panes = "A2"
    hoja.auto_filter.ref = hoja.dimensions
    for columna in hoja.columns:
        ancho = max(len(str(c.value)) if c.value is not None else 0 for c in columna[:200])
        hoja.column_dimensions[columna[0].column_letter].width = min(ancho + 2, 45)


def main() -> int:
    ap = argparse.ArgumentParser(description="Consolida archivos Excel/CSV de una carpeta.")
    ap.add_argument("carpeta", type=Path, help="Carpeta con los archivos a unir")
    ap.add_argument("-o", "--salida", type=Path, help="Archivo de salida (.xlsx)")
    ap.add_argument("--agrupar", help="Columna para el resumen (ej: producto). Por defecto: archivo")
    args = ap.parse_args()

    if not args.carpeta.is_dir():
        print(f"No existe la carpeta: {args.carpeta}")
        return 1

    salida = args.salida or args.carpeta / f"reporte_consolidado_{datetime.now():%Y%m%d_%H%M}.xlsx"
    archivos = sorted(
        p for p in args.carpeta.iterdir()
        if p.suffix.lower() in EXTENSIONES
        and not p.name.startswith(("~$", "reporte_consolidado"))
        and p != salida
    )
    if not archivos:
        print("No encontré archivos Excel/CSV en esa carpeta.")
        return 1

    registro, frames = [], []
    for ruta in archivos:
        frames += leer_archivo(ruta, registro)
    if not frames:
        print("Ningún archivo se pudo leer. Revisá el registro de errores.")
        return 1

    consolidado = pd.concat(frames, ignore_index=True, sort=False)

    # Avisar de columnas que no están en todos los archivos
    for col in consolidado.columns:
        if col in ("archivo_origen", "hoja_origen"):
            continue
        faltan = sorted({f["archivo_origen"].iat[0] for f in frames if col not in f.columns})
        if faltan:
            registro.append(("-", "-", "AVISO", f"Columna '{col}' falta en: {', '.join(faltan)}"))

    resumen = armar_resumen(consolidado, args.agrupar)
    log = pd.DataFrame(registro, columns=["archivo", "hoja", "estado", "detalle"])

    with pd.ExcelWriter(salida, engine="openpyxl", datetime_format="DD/MM/YYYY") as xl:
        consolidado.to_excel(xl, sheet_name="Consolidado", index=False)
        resumen.to_excel(xl, sheet_name="Resumen", index=False)
        log.to_excel(xl, sheet_name="Registro", index=False)
        for hoja in xl.book.worksheets:
            formatear_hoja(hoja)

    print(f"Listo: {len(archivos)} archivos, {len(consolidado)} filas -> {salida}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
