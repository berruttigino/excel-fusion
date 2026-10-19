"""
ExcelFusion - identidad visual de la app
========================================
Concepto: el isotipo es una celda combinada de planilla (tres celdas que se funden en una).
La interfaz repite esa idea: el cargador de archivos es una "celda seleccionada" (borde verde
y cuadradito de relleno en la esquina), las esquinas son casi rectas como una grilla y los
números usan cifras tabulares para que las columnas queden alineadas.

Los colores base también están en .streamlit/config.toml (tema de Streamlit).
"""
import re
from pathlib import Path

ASSETS = Path(__file__).parent / "assets"

TINTA = "#14213D"      # texto y títulos
VERDE = "#217346"      # color principal (verde de selección de Excel)
MENTA = "#E3F1E8"      # fondo de celda seleccionada
PAPEL = "#F5F7F4"      # fondo de página
LINEA = "#D8DED8"      # líneas de cuadrícula
AMBAR = "#F2B134"      # controlador de relleno: acento mínimo
APAGADO = "#4A5668"    # texto secundario

CSS = f"""
<style>
:root {{ --ink:{TINTA}; --green:{VERDE}; --mint:{MENTA}; --paper:{PAPEL}; --line:{LINEA}; --amber:{AMBAR}; --muted:{APAGADO}; }}

/* Tipografía: la familia de Excel (Calibri; Carlito es su equivalente libre), con alternativas del sistema */
.stApp, .stApp p, .stApp li, .stApp label, .stApp h1, .stApp h2, .stApp h3, .stApp h4,
.stApp button, .stApp input, .stApp textarea {{
  font-family: "Carlito", "Calibri", "Segoe UI", system-ui, -apple-system, "Helvetica Neue", sans-serif;
}}
.stApp p, .stApp li, .stApp label {{ font-size: 1.05rem; }}
.stApp h2, .stApp h3 {{ color: var(--ink); font-weight: 700; letter-spacing: -0.005em; }}
.stApp .block-container {{ max-width: 1200px; padding-top: 2.4rem; padding-bottom: 3rem; }}
header[data-testid="stHeader"] {{ background: transparent; }}

/* Cabecera */
.stApp .ef-hero {{ margin: 0 0 1.6rem 0; }}
.stApp .ef-hero svg {{ height: 54px; width: auto; display: block; margin-bottom: 1.7rem; }}
.stApp .ef-hero h1 {{
  font-size: 2.7rem; line-height: 1.05; font-weight: 700; letter-spacing: -0.012em;
  color: var(--ink); margin: 0 0 .6rem 0; padding: 0;
}}
.stApp .ef-hero p {{ color: var(--muted); font-size: 1.12rem; line-height: 1.45; margin: 0; max-width: 44rem; }}

/* Cargador de archivos = celda seleccionada de Excel */
[data-testid="stFileUploaderDropzone"] {{
  background: #FFFFFF; border: 2px solid var(--green); border-radius: 3px;
  position: relative; overflow: visible; padding: 1.5rem 1.3rem;
}}
[data-testid="stFileUploaderDropzone"]::after {{
  content: ""; position: absolute; right: -8px; bottom: -8px; width: 11px; height: 11px;
  background: var(--amber); border: 2px solid #FFFFFF; border-radius: 2px;
}}
[data-testid="stFileUploaderDropzone"] button {{
  background: var(--green); color: #FFFFFF; border: 0; border-radius: 3px; font-weight: 700;
}}
[data-testid="stFileUploaderDropzone"] button:hover {{ background: #1B5F3A; color: #FFFFFF; }}

/* Indicadores: cifras tabulares y un borde de columna en vez de tarjetas */
[data-testid="stMetric"] {{ border-left: 3px solid var(--green); padding: .15rem 0 .15rem .95rem; }}
[data-testid="stMetricLabel"] p {{ color: var(--muted); font-size: .98rem; }}
[data-testid="stMetricValue"] {{ color: var(--ink); font-weight: 700; font-variant-numeric: tabular-nums; }}
[data-testid="stMetricValue"] > div {{
  font-size: 1.75rem; line-height: 1.15; white-space: normal; overflow: visible; text-overflow: clip;
}}

/* Pestañas, avisos, filtros y botón de descarga */
button[data-baseweb="tab"] p {{ font-weight: 700; font-size: 1.05rem; }}
[data-testid="stAlert"] {{ border-radius: 3px; }}
[data-testid="stExpander"] {{ border-radius: 3px; border-color: var(--line); }}
[data-testid="stDownloadButton"] button {{
  background: var(--green); color: #FFFFFF; border: 0; border-radius: 3px;
  font-weight: 700; padding: .6rem 1.3rem;
}}
[data-testid="stDownloadButton"] button p {{ color: #FFFFFF; }}
[data-testid="stDownloadButton"] button:hover {{ background: #1B5F3A; color: #FFFFFF; }}

/* Pie */
.stApp .ef-foot {{
  margin-top: 3.2rem; padding-top: 1rem; border-top: 1px solid var(--line);
  color: var(--muted); font-size: .95rem;
}}
.stApp .ef-foot a {{ color: var(--green); font-weight: 700; text-decoration: none; }}
.stApp .ef-foot a:hover {{ text-decoration: underline; }}
</style>
"""


def _una_linea(texto: str) -> str:
    """Streamlit interpreta las líneas en blanco como fin del bloque HTML: todo en una línea."""
    return re.sub(r"\s*\n\s*", " ", texto).strip()


def _logo_svg() -> str:
    ruta = ASSETS / "logo.svg"
    if ruta.exists():
        return _una_linea(ruta.read_text(encoding="utf-8"))
    return "<strong>ExcelFusion</strong>"  # por si falta la carpeta assets


def icono_pagina():
    """Ícono de la pestaña del navegador: el PNG de la marca, o un emoji si falta."""
    ruta = ASSETS / "icon_192.png"
    return str(ruta) if ruta.exists() else "📊"


def aplicar_estilo(st) -> None:
    st.markdown(_una_linea(CSS), unsafe_allow_html=True)


def cabecera(st) -> None:
    st.markdown(
        _una_linea(
            f"""<div class="ef-hero">{_logo_svg()}
            <h1>Many files. One report.</h1>
            <p>Merge messy Excel and CSV files into one clean report, with KPIs and charts.<br>
            <em>Muchos archivos, un solo reporte: uní tus Excel y CSV desordenados, con KPIs y gráficos.</em></p>
            </div>"""
        ),
        unsafe_allow_html=True,
    )


def pie(st) -> None:
    st.markdown(
        _una_linea(
            """<div class="ef-foot">ExcelFusion, hecho por Gino Berrutti.
            <a href="https://github.com/berruttigino/excel-fusion" target="_blank">Código en GitHub</a>.
            ¿Querés una versión adaptada a tus archivos? Escribime por ahí.</div>"""
        ),
        unsafe_allow_html=True,
    )
