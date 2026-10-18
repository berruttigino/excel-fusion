"""
ExcelFusion - cálculos del dashboard y de los KPIs
==================================================
Solo usa pandas (sin Streamlit), así se puede probar y reutilizar por separado.
"""
import pandas as pd

# Nombres de columna que se prefieren al elegir valores por defecto
PREF_METRICA = ["monto", "total", "importe", "ventas", "venta", "valor", "precio", "cantidad"]
PREF_CATEGORIA = ["producto", "categoria", "cliente", "sucursal", "vendedor", "region", "rubro"]

PERIODOS = {"D": "Día", "W": "Semana", "M": "Mes"}


def _ordenar(columnas, preferidas, al_final=()):
    """Ordena poniendo primero las columnas preferidas y al final las indicadas."""
    def clave(c):
        if c in al_final:
            return len(preferidas) + 1
        return preferidas.index(c) if c in preferidas else len(preferidas)
    return sorted(columnas, key=clave)


def detectar_columnas(df: pd.DataFrame) -> dict:
    """Clasifica las columnas en numéricas (para medir), fechas y categorías (para agrupar)."""
    numericas = df.select_dtypes("number").columns.tolist()
    fechas = df.select_dtypes("datetime").columns.tolist()
    categoricas = [c for c in df.columns if c not in numericas + fechas and c != "hoja_origen"]
    return {
        "numericas": _ordenar(numericas, PREF_METRICA),
        "fechas": fechas,
        "categoricas": _ordenar(categoricas, PREF_CATEGORIA, al_final=("archivo_origen",)),
    }


def frecuencia_sugerida(df: pd.DataFrame, fecha: str | None) -> str:
    """Día, semana o mes según el rango de fechas que cubren los datos."""
    if not fecha:
        return "D"
    f = df[fecha].dropna()
    if f.empty:
        return "D"
    dias = (f.max() - f.min()).days
    return "M" if dias > 90 else "W" if dias > 21 else "D"


def elegir_defaults(df: pd.DataFrame) -> dict:
    """Métrica, categoría y fecha que se usan por defecto (y para la hoja KPIs del Excel)."""
    cols = detectar_columnas(df)
    metrica = cols["numericas"][0] if cols["numericas"] else None
    categoria = cols["categoricas"][0] if cols["categoricas"] else None
    fecha = cols["fechas"][0] if cols["fechas"] else None
    return {"metrica": metrica, "categoria": categoria, "fecha": fecha,
            "freq": frecuencia_sugerida(df, fecha)}


def filtrar(df, categoria=None, valores=None, fecha=None, desde=None, hasta=None) -> pd.DataFrame:
    """Filtra por valores de una categoría y/o por rango de fechas (ambos extremos incluidos)."""
    out = df
    if categoria and valores is not None:
        out = out[out[categoria].astype(str).isin([str(v) for v in valores])]
    if fecha and desde is not None:
        out = out[out[fecha] >= pd.Timestamp(desde)]
    if fecha and hasta is not None:
        out = out[out[fecha] < pd.Timestamp(hasta) + pd.Timedelta(days=1)]
    return out


def ranking(df, categoria, metrica, top=10) -> pd.DataFrame:
    """Total de la métrica por categoría, de mayor a menor, con % de participación."""
    total_por_cat = df.groupby(categoria)[metrica].sum().sort_values(ascending=False)
    total = total_por_cat.sum()
    participacion = (total_por_cat / total * 100).round(1) if total else total_por_cat * 0
    out = pd.DataFrame({metrica: total_por_cat, "participacion_%": participacion})
    return out.head(top).reset_index()


def serie_tiempo(df, fecha, metrica, freq="D") -> pd.Series:
    """Total de la métrica por día (D), semana (W) o mes (M)."""
    d = df.dropna(subset=[fecha])
    if d.empty:
        return pd.Series(dtype=float)
    periodo = d[fecha].dt.to_period(freq)
    serie = d.groupby(periodo)[metrica].sum()
    serie.index = serie.index.to_timestamp()
    return serie


def etiqueta_periodo(fecha: pd.Timestamp, freq: str) -> str:
    if freq == "M":
        return fecha.strftime("%m/%Y")
    if freq == "W":
        return "Semana del " + fecha.strftime("%d/%m/%Y")
    return fecha.strftime("%d/%m/%Y")


def calcular_kpis(df, metrica, categoria=None, fecha=None, freq="D") -> dict:
    """Indicadores clave de la métrica elegida."""
    s = df[metrica].dropna()
    k = {
        "total": float(s.sum()) if len(s) else 0.0,
        "promedio": float(s.mean()) if len(s) else 0.0,
        "maximo": float(s.max()) if len(s) else 0.0,
        "minimo": float(s.min()) if len(s) else 0.0,
        "registros": int(len(df)),
        "lider": None, "lider_pct": None,
        "mejor_periodo": None, "mejor_periodo_valor": None,
    }
    if categoria and len(df):
        r = ranking(df, categoria, metrica, top=1)
        if len(r):
            k["lider"] = str(r.iloc[0][categoria])
            k["lider_pct"] = float(r.iloc[0]["participacion_%"])
    if fecha and len(df):
        t = serie_tiempo(df, fecha, metrica, freq)
        if len(t):
            k["mejor_periodo"] = etiqueta_periodo(t.idxmax(), freq)
            k["mejor_periodo_valor"] = float(t.max())
    return k


def tabla_kpis(df, metrica, categoria=None, fecha=None, freq="D") -> pd.DataFrame:
    """Los KPIs en formato tabla (Indicador / Valor), lista para la hoja 'KPIs' del Excel."""
    k = calcular_kpis(df, metrica, categoria, fecha, freq)
    filas = [
        ("Métrica analizada", metrica),
        (f"Total de {metrica}", k["total"]),
        ("Promedio por registro", k["promedio"]),
        ("Valor máximo", k["maximo"]),
        ("Valor mínimo", k["minimo"]),
        ("Registros", k["registros"]),
    ]
    if k["lider"] is not None:
        filas += [(f"Líder en {categoria}", k["lider"]),
                  ("Participación del líder (%)", k["lider_pct"])]
    if k["mejor_periodo"] is not None:
        filas += [(f"Mejor período ({PERIODOS[freq].lower()})", k["mejor_periodo"]),
                  ("Total en el mejor período", k["mejor_periodo_valor"])]
    return pd.DataFrame(filas, columns=["Indicador", "Valor"])
