#!/usr/bin/env python3
"""Crea una carpeta 'datos_demo' con archivos de ventas desordenados a propósito
(columnas con nombres distintos, un CSV y una hoja vacía) para probar el consolidador."""
import random
from pathlib import Path

import pandas as pd

random.seed(7)
carpeta = Path("datos_demo")
carpeta.mkdir(exist_ok=True)
productos = ["Café", "Té", "Yerba", "Azúcar", "Galletitas"]


def ventas(n, columnas):
    fechas = pd.date_range("2026-09-01", periods=30).strftime("%d/%m/%Y")
    return pd.DataFrame({
        columnas[0]: [random.choice(fechas) for _ in range(n)],
        columnas[1]: [random.choice(productos) for _ in range(n)],
        columnas[2]: [random.randint(1, 20) for _ in range(n)],
        columnas[3]: [round(random.uniform(500, 9000), 2) for _ in range(n)],
    })


ventas(40, ["Fecha", "Producto", "Cantidad", "Monto"]).to_excel(
    carpeta / "ventas_norte.xlsx", index=False)

with pd.ExcelWriter(carpeta / "ventas_sur.xlsx") as xl:  # nombres distintos + hoja vacía
    ventas(35, ["Fecha Venta", "Artículo", "Cant", "Importe"]).to_excel(xl, sheet_name="Septiembre", index=False)
    pd.DataFrame().to_excel(xl, sheet_name="Vacia", index=False)

ventas(25, ["fecha", "producto", "cantidad", "total"]).to_csv(
    carpeta / "ventas_online.csv", index=False)

print(f"Archivos de prueba creados en: {carpeta.resolve()}")
