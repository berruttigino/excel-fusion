"""
ExcelFusion - limpieza de datos
===============================
Cuatro acciones, todas con registro de lo que cambiaron:

  1. Quitar espacios sobrantes (al inicio, al final y repetidos adentro).
  2. Unificar variantes de escritura: "Café", "CAFÉ" y "cafe" pasan a ser lo mismo, y se conserva la
     escritura más común. Así los totales por categoría no se parten en pedazos.
  3. Eliminar filas duplicadas (dentro de cada archivo o entre todos). Las filas eliminadas no se
     pierden: salen aparte para poder revisarlas.
  4. Marcar filas con datos faltantes en una columna nueva, sin borrar nada.

Solo usa pandas, así que se puede probar sin Streamlit.
"""
import unicodedata

import numpy as np
import pandas as pd

PROTEGIDAS = ("archivo_origen", "hoja_origen")
COL_FALTANTES = "datos_faltantes"
MODOS_DUPLICADOS = ("no", "archivo", "todos")


def columnas_texto(df: pd.DataFrame) -> list:
    """Columnas de texto sobre las que se puede limpiar (no las de origen ni las numéricas o de fecha)."""
    return [c for c in df.columns
            if c not in PROTEGIDAS and c != COL_FALTANTES
            and not pd.api.types.is_numeric_dtype(df[c])
            and not pd.api.types.is_datetime64_any_dtype(df[c])]


def _es_texto(valor) -> bool:
    return isinstance(valor, str)


def _recortar(valor):
    if not _es_texto(valor):
        return valor
    limpio = " ".join(valor.split())      # quita espacios (también los de ancho fijo) y repetidos
    return limpio if limpio else np.nan   # un texto vacío cuenta como dato faltante


def recortar_espacios(df: pd.DataFrame, columnas: list):
    """Devuelve (tabla, celdas cambiadas, columnas afectadas)."""
    total, afectadas = 0, []
    for c in columnas:
        antes = df[c]
        despues = antes.map(_recortar)
        cambiadas = int((antes.notna() & (antes != despues)).sum())
        if cambiadas:
            df[c] = despues
            total += cambiadas
            afectadas.append(c)
    return df, total, afectadas


def _clave(valor: str) -> str:
    sin_tildes = unicodedata.normalize("NFKD", valor).encode("ascii", "ignore").decode()
    return " ".join(sin_tildes.casefold().split())


def unificar_variantes(df: pd.DataFrame, columnas: list):
    """Une variantes de mayúsculas, tildes y espacios, conservando la escritura más común.
    Devuelve (tabla, celdas cambiadas, ejemplos 'cafe -> Café')."""
    total, ejemplos = 0, []
    for c in columnas:
        validos = df[c][df[c].map(_es_texto)]
        if validos.nunique() < 2:
            continue
        claves = validos.map(_clave)
        if claves.nunique() == validos.nunique():
            continue  # no hay variantes de una misma cosa
        conteo = pd.DataFrame({"k": claves, "v": validos}).value_counts().reset_index(name="n")
        # La escritura más frecuente de cada grupo; si empatan, la que tiene tildes y mayúscula inicial
        conteo["tildes"] = conteo["v"].map(lambda t: sum(ord(ch) > 127 for ch in t))
        conteo["capital"] = conteo["v"].map(lambda t: t[:1].isupper() and t[1:].islower())
        conteo = conteo.sort_values(["k", "n", "tildes", "capital"], ascending=[True, False, False, False])
        mejor = conteo.drop_duplicates("k").set_index("k")["v"]
        nuevo = claves.map(mejor)
        cambiadas = nuevo != validos
        if cambiadas.any():
            df.loc[validos.index, c] = nuevo
            total += int(cambiadas.sum())
            pares = pd.DataFrame({"antes": validos[cambiadas], "despues": nuevo[cambiadas]}).drop_duplicates()
            ejemplos += [f"{c}: '{a}' → '{d}'" for a, d in pares.head(3).itertuples(index=False)]
    return df, total, ejemplos


def eliminar_duplicados(df: pd.DataFrame, modo: str):
    """modo 'archivo': duplicado = misma fila dentro del mismo archivo y hoja.
    modo 'todos': igual valores en todas las columnas de datos, aunque vengan de archivos distintos.
    Devuelve (tabla sin duplicados, filas eliminadas, cantidad)."""
    if modo == "archivo":
        subset = list(df.columns)
    elif modo == "todos":
        subset = [c for c in df.columns if c not in PROTEGIDAS]
    else:
        return df, df.iloc[0:0].copy(), 0
    if not subset:
        return df, df.iloc[0:0].copy(), 0
    marca = df.duplicated(subset=subset, keep="first")
    return df[~marca].reset_index(drop=True), df[marca].reset_index(drop=True), int(marca.sum())


def marcar_faltantes(df: pd.DataFrame):
    """Agrega la columna 'datos_faltantes' con los nombres de las columnas vacías de cada fila.
    Devuelve (tabla, filas con faltantes, {columna: cantidad de vacíos})."""
    columnas = [c for c in df.columns if c not in PROTEGIDAS and c != COL_FALTANTES]
    vacios = df[columnas].isna()
    nombres = np.array([f"{c}, " for c in columnas], dtype=object)
    texto = np.where(vacios.to_numpy(), nombres, "")
    unido = pd.Series(["".join(fila) for fila in texto], index=df.index).str.rstrip(", ")
    df[COL_FALTANTES] = unido.replace("", np.nan)
    por_columna = {c: int(n) for c, n in vacios.sum().items() if n}
    return df, int(vacios.any(axis=1).sum()), por_columna


def limpiar(df: pd.DataFrame, recortar: bool = True, variantes: bool = True,
            duplicados: str = "no", faltantes: bool = False):
    """Aplica las acciones elegidas. Devuelve (tabla limpia, filas eliminadas, resumen).
    El resumen es una tabla Acción / Cantidad / Detalle, con una fila por acción activada."""
    df = df.copy()
    texto = columnas_texto(df)
    filas = []

    if recortar:
        df, n, columnas = recortar_espacios(df, texto)
        filas.append(("Espacios sobrantes quitados (celdas)", n, "Columnas: " + ", ".join(columnas) if columnas else "-"))
    if variantes:
        df, n, ejemplos = unificar_variantes(df, texto)
        filas.append(("Variantes de escritura unificadas (celdas)", n, "; ".join(ejemplos) if ejemplos else "-"))

    eliminados = df.iloc[0:0].copy()
    if duplicados in ("archivo", "todos"):
        df, eliminados, n = eliminar_duplicados(df, duplicados)
        donde = "dentro de cada archivo" if duplicados == "archivo" else "entre todos los archivos"
        filas.append(("Filas duplicadas eliminadas", n, f"Criterio: filas idénticas {donde}"))
    if faltantes:
        df, n, por_columna = marcar_faltantes(df)
        detalle = ", ".join(f"{c}: {q}" for c, q in por_columna.items()) if por_columna else "-"
        filas.append(("Filas con datos faltantes marcadas", n, f"Vacíos por columna: {detalle}"))

    return df, eliminados, pd.DataFrame(filas, columns=["Acción", "Cantidad", "Detalle"])
