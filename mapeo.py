"""
ExcelFusion - mapeo de columnas entre archivos
==============================================
Cuando cada archivo llama distinto a lo mismo ("Importe neto", "Total", "Monto"), esta parte
sugiere qué columnas unir y aplica los cambios que el usuario confirma o corrige.

Regla de seguridad: dos columnas que aparecen JUNTAS en un mismo archivo nunca se sugieren
como equivalentes, porque en ese archivo son datos distintos.
"""
from difflib import SequenceMatcher

import pandas as pd

from consolidar_excels import ALIAS

COL_ORIGINAL = "Column / Columna"
COL_EN = "In files / En archivos"
COL_NUEVO = "Rename to / Renombrar a"
PROTEGIDAS = ("archivo_origen", "hoja_origen")


def _presencia(frames) -> dict:
    """columna -> conjunto de tablas (archivo/hoja) donde aparece."""
    presencia = {}
    for i, f in enumerate(frames):
        for c in f.columns:
            if c not in PROTEGIDAS:
                presencia.setdefault(c, set()).add(i)
    return presencia


def _parecidos(a: str, b: str) -> bool:
    """Nombres casi iguales ('precio_unit' / 'precio_unitario') o una abreviatura cortada ('cant' / 'cantidad')."""
    if SequenceMatcher(None, a, b).ratio() >= 0.84:
        return True
    corto, largo = sorted((a, b), key=len)
    return len(corto) >= 4 and len(largo) > len(corto) and largo.startswith(corto) and largo[len(corto)] != "_"


def _por_alias(nombre: str):
    """'importe_neto' -> 'monto' porque empieza con 'importe', que es alias de 'monto'."""
    for clave, destino in ALIAS.items():
        if nombre.startswith(clave + "_"):
            return destino
    return None


def sugerir_mapeo(frames) -> dict:
    """Devuelve {columna: nombre sugerido}. Si no hay nada que unir, cada columna conserva su nombre."""
    presencia = _presencia(frames)
    acumulado = {c: set(p) for c, p in presencia.items()}
    canonicas, destino = [], {}
    for c in sorted(presencia, key=lambda c: (-len(presencia[c]), len(c), c)):   # primero las más frecuentes
        libres = [k for k in canonicas if not (acumulado[k] & presencia[c])]      # nunca juntas en una tabla
        elegido = next((k for k in libres if _parecidos(c, k)), None)
        if elegido is None:
            alias = _por_alias(c)
            if alias in libres:
                elegido = alias
        destino[c] = elegido or c
        if elegido:
            acumulado[elegido] |= presencia[c]
        else:
            canonicas.append(c)
    return destino


def tabla_columnas(frames) -> pd.DataFrame:
    """Una fila por columna distinta: en cuántos archivos está y a qué nombre se unificaría."""
    presencia, sugerido = _presencia(frames), sugerir_mapeo(frames)
    total = len(frames)
    filas = [{COL_ORIGINAL: c, COL_EN: f"{len(presencia[c])} de {total}", COL_NUEVO: sugerido[c]}
             for c in sorted(presencia, key=lambda c: (-len(presencia[c]), c))]
    return pd.DataFrame(filas, columns=[COL_ORIGINAL, COL_EN, COL_NUEVO])


def hay_columnas_desparejas(frames) -> bool:
    """True si alguna columna no está en todas las tablas (conviene revisar el mapeo)."""
    return any(len(p) < len(frames) for p in _presencia(frames).values())


def _unir_duplicadas(df: pd.DataFrame) -> pd.DataFrame:
    """Si dos columnas terminaron con el mismo nombre, se combinan tomando el primer valor no vacío."""
    if not df.columns.duplicated().any():
        return df
    unidas = {}
    for nombre in dict.fromkeys(df.columns):
        grupo = df.loc[:, df.columns == nombre]
        unidas[nombre] = grupo.bfill(axis=1).iloc[:, 0] if grupo.shape[1] > 1 else grupo.iloc[:, 0]
    return pd.DataFrame(unidas)


def aplicar_mapeo(frames, mapeo: dict) -> list:
    """Renombra las columnas de cada tabla según el mapeo {original: nuevo}."""
    return [_unir_duplicadas(f.rename(columns=mapeo)) for f in frames]
