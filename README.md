# ExcelFusion

Merge dozens of messy Excel and CSV files into one clean report, in seconds.

Stop copy-pasting between spreadsheets. ExcelFusion reads every `.xlsx`, `.xls` and `.csv` file in a folder, standardizes the column names, combines everything into one table and builds a ready-to-share Excel report.

## What you get

The output workbook has three sheets:

| Sheet | Content |
|---|---|
| **Consolidado** | All rows from all files, with the source file and sheet of each row |
| **Resumen** | Totals per group (per file by default, or per any column you choose) |
| **Registro** | What was read, what failed, empty sheets skipped and columns missing in some files |

## Features

- Reads all sheets of every Excel file, plus CSVs
- Unifies different column names for the same thing (`Importe`, `Total` and `Monto` become one column)
- Removes accents and extra spaces from headers
- Parses date columns (day-first format)
- Skips empty sheets and reports corrupt files instead of crashing
- Formatted output: bold headers, filters, frozen top row, auto-sized columns

## Quick start

```bash
pip install -r requirements.txt

# 1. Generate sample messy files (optional)
python generar_datos_demo.py

# 2. Consolidate a folder
python consolidar_excels.py datos_demo

# Group the summary by a column and choose the output name
python consolidar_excels.py datos_demo --agrupar producto -o ventas_septiembre.xlsx
```

## Customize

Different column names in your files? Add them to the `ALIAS` dictionary at the top of `consolidar_excels.py`:

```python
ALIAS = {
    "importe": "monto",
    "cant": "cantidad",
    # "your_column": "standard_name",
}
```

## Need this adapted to your files?

I build custom Excel automation, reports and dashboards (Excel, SQL, Power BI, Python). Contact: _add your email or profile link here_.

---

# ExcelFusion (Español)

Une decenas de archivos Excel y CSV desordenados en un único reporte limpio.

**Uso rápido:** `pip install -r requirements.txt` y luego `python consolidar_excels.py carpeta_con_archivos`. El resultado es un Excel con las hojas *Consolidado*, *Resumen* y *Registro*.
