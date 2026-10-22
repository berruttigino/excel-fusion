#!/usr/bin/env python3
"""
ExcelFusion - Consolidador de archivos Excel/CSV
================================================
Lee todos los archivos .xlsx/.xls/.csv de una carpeta, encuentra solo la fila de títulos,
unifica los nombres de columnas, los junta en una sola tabla y genera un reporte en Excel con:

  - Consolidado : todos los datos juntos (con archivo y hoja de origen)
  - Resumen     : totales por grupo (por defecto, por archivo)
  - Registro    : qué se leyó, qué falló y qué columnas no coinciden

Uso:
    python consolidar_excels.py carpeta_con_archivos
    python consolidar_excels.py carpeta_con_archivos -o reporte.xlsx --agrupar producto
"""
import argparse
import csv
import io
import re
import sys
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd

EXTENSIONES = {".xlsx", ".xlsm", ".xls", ".csv"}

# Nombres distintos que significan lo mismo. Agregá los que veas en tus archivos.
ALIAS = {
    "importe": "monto", "total": "monto", "amount": "monto",
    "cant": "cantidad", "qty": "cantidad", "quantity": "cantidad", "unidades": "cantidad",
    "articulo": "producto", "item": "producto", "product": "producto",
    "customer": "cliente",
    "fecha_venta": "fecha", "fecha_de_venta": "fecha", "date": "fecha",
}

# Filas como "Total", "Subtotal", "Total general" o "Grand total" (se omiten para no duplicar sumas)
PATRON_TOTAL = re.compile(
    r"^\s*(sub\s*)?total(es)?(\s+general(es)?)?\s*:?\s*$|^\s*(gran|grand)\s+total\s*:?\s*$", re.IGNORECASE)


def limpiar_nombre(nombre: str) -> str:
    """'Fecha de Venta ' -> 'fecha_de_venta' (sin tildes, minúsculas, con guiones bajos)."""
    texto = unicodedata.normalize("NFKD", str(nombre)).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]+", "_", texto).strip("_")


def normalizar(nombre: str) -> str:
    """Igual que limpiar_nombre, y además unifica sinónimos conocidos ('importe' -> 'monto')."""
    texto = limpiar_nombre(nombre)
    return ALIAS.get(texto, texto)


# ---------- lectura "cruda" (sin suponer dónde están los títulos) ----------
def _leer_csv_crudo(ruta: Path) -> pd.DataFrame:
    """Lee un CSV como tabla de texto: detecta el separador, tolera filas de distinto largo y codificaciones."""
    texto = None
    for codificacion in ("utf-8-sig", "latin-1"):
        try:
            texto = ruta.read_text(encoding=codificacion)
            break
        except UnicodeDecodeError:
            continue
    try:
        dialecto = csv.Sniffer().sniff(texto[:4096], delimiters=",;\t|")
    except csv.Error:
        dialecto = csv.excel
    filas = list(csv.reader(io.StringIO(texto), dialecto))
    ancho = max((len(f) for f in filas), default=0)
    df = pd.DataFrame([f + [""] * (ancho - len(f)) for f in filas])
    df = df.apply(lambda col: col.str.strip())
    return df.mask(df == "")


def _leer_crudo(ruta: Path) -> dict:
    if ruta.suffix.lower() == ".csv":
        return {"csv": _leer_csv_crudo(ruta)}
    return pd.read_excel(ruta, sheet_name=None, header=None)


# ---------- detección de títulos y tipos ----------
def _parece_numero(valor) -> bool:
    texto = str(valor).strip().replace(" ", "").lstrip("$€").rstrip("%")
    if "," in texto and "." in texto:
        texto = texto.replace(".", "").replace(",", ".")
    else:
        texto = texto.replace(",", ".")
    try:
        float(texto)
        return True
    except ValueError:
        return False


def _es_texto(valor) -> bool:
    return isinstance(valor, str) and valor.strip() != "" and not _parece_numero(valor)


def detectar_fila_titulos(raw: pd.DataFrame, max_filas: int = 30) -> int:
    """Posición de la fila de títulos: la primera casi completa y casi toda de texto.
    Salta títulos del reporte, filas de una sola celda y filas vacías de arriba."""
    muestra = raw.iloc[:max_filas]
    llenos = muestra.notna().sum(axis=1)
    if llenos.empty or llenos.max() == 0:
        return 0
    minimo = 1 if llenos.max() == 1 else max(2, 0.6 * llenos.max())
    for pos in range(len(muestra)):
        celdas = muestra.iloc[pos].dropna()
        if len(celdas) >= minimo and sum(_es_texto(v) for v in celdas) >= 0.8 * len(celdas):
            return pos
    return 0


def _nombres_unicos(titulos) -> list:
    vistos, nombres = {}, []
    for i, valor in enumerate(titulos):
        base = (normalizar(valor) if pd.notna(valor) else "") or f"columna_{i + 1}"
        vistos[base] = vistos.get(base, 0) + 1
        nombres.append(base if vistos[base] == 1 else f"{base}_{vistos[base]}")
    return nombres


def _afinar(serie: pd.Series) -> pd.Series:
    """Enteros si no hay decimales ni vacíos; si no, decimales."""
    serie = serie.astype("float64")
    if serie.notna().all() and (serie % 1 == 0).all():
        return serie.astype("int64")
    return serie


def _a_numero(col: pd.Series):
    """Convierte texto a número si TODOS los valores lo son.
    Entiende 1234.56, 1234,56, 1.234,56 (formato latino) y 1,234.56 (formato inglés)."""
    s = col.dropna().astype(str).str.strip()
    if s.empty or s.str.match(r"^-?0\d").any():  # códigos como 007 se dejan como texto
        return None
    texto = col.astype("string").str.strip()
    latino = r"-?\d{1,3}(?:\.\d{3})*(?:,\d+)?|-?\d+(?:,\d+)?"
    ingles = r"-?\d{1,3}(?:,\d{3})*(?:\.\d+)?|-?\d+(?:\.\d+)?"
    if s.str.contains(",").any() and s.str.fullmatch(latino).all():
        return _afinar(pd.to_numeric(texto.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)))
    if s.str.fullmatch(ingles).all():
        return _afinar(pd.to_numeric(texto.str.replace(",", "", regex=False)))
    return None


def _convertir_tipos(df: pd.DataFrame) -> pd.DataFrame:
    for c in df.columns:
        col = df[c].infer_objects()
        if not (pd.api.types.is_numeric_dtype(col) or pd.api.types.is_datetime64_any_dtype(col)):
            convertida = _a_numero(col)
            if convertida is not None:
                col = convertida
        df[c] = col
    return df


def quitar_filas_total(df: pd.DataFrame):
    """Quita filas cuyo texto es 'Total', 'Subtotal', 'Total general'... Devuelve (tabla, cantidad quitada)."""
    texto = df.select_dtypes(exclude=["number", "datetime"])
    if texto.empty:
        return df, 0
    mascara = texto.apply(lambda col: col.map(lambda v: isinstance(v, str) and bool(PATRON_TOTAL.match(v)))).any(axis=1)
    return df[~mascara].reset_index(drop=True), int(mascara.sum())


def _procesar_hoja(raw, archivo, hoja, registro, quitar_totales):
    raw = raw.dropna(how="all").dropna(axis=1, how="all")
    if raw.empty:
        registro.append((archivo, hoja, "OMITIDA", "Hoja vacía"))
        return None
    pos = detectar_fila_titulos(raw)
    df = raw.iloc[pos + 1:].copy()
    df.columns = _nombres_unicos(raw.iloc[pos])
    df = df.dropna(how="all").reset_index(drop=True)
    if df.empty:
        registro.append((archivo, hoja, "OMITIDA", "Tiene títulos pero no datos"))
        return None
    quitadas = 0
    if quitar_totales:  # antes de convertir tipos: una fila de totales con basura no debe arruinar la columna
        df, quitadas = quitar_filas_total(df)
    df = _convertir_tipos(df)
    for col in df.columns:
        if col == "fecha" or col.startswith("fecha_"):
            df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=True)
    df.insert(0, "archivo_origen", archivo)
    df.insert(1, "hoja_origen", hoja)
    registro.append((archivo, hoja, "OK", f"{len(df)} filas"))
    if pos:
        registro.append((archivo, hoja, "INFO", f"Se omitieron {pos} fila(s) anteriores a los títulos"))
    if quitadas:
        registro.append((archivo, hoja, "INFO", f"{quitadas} fila(s) de totales omitidas"))
    return df


def leer_archivo(ruta: Path, registro: list, quitar_totales: bool = True) -> list:
    """Devuelve una lista de DataFrames (uno por hoja) o [] si el archivo falla."""
    try:
        hojas = _leer_crudo(ruta)
    except Exception as error:  # archivo corrupto, protegido, etc.
        registro.append((ruta.name, "-", "ERROR", str(error)[:150]))
        return []
    frames = []
    for hoja, raw in hojas.items():
        df = _procesar_hoja(raw, ruta.name, hoja, registro, quitar_totales)
        if df is not None:
            frames.append(df)
    return frames


# ---------- resumen y formato ----------
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
    ap.add_argument("--conservar-totales", action="store_true", help="No omitir filas llamadas 'Total'")
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
        frames += leer_archivo(ruta, registro, quitar_totales=not args.conservar_totales)
    if not frames:
        print("Ningún archivo se pudo leer. Revisá el registro de errores.")
        return 1

    from mapeo import aplicar_mapeo, sugerir_mapeo  # import local: mapeo.py usa este mismo archivo

    sugerido = sugerir_mapeo(frames)
    for original, nuevo in sugerido.items():
        if original != nuevo:
            registro.append(("-", "-", "INFO", f"Columna '{original}' unificada como '{nuevo}'"))
    frames = aplicar_mapeo(frames, sugerido)
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
