# ============================================================
# MAKROPETROL NEXUS v7.3 — Inventario · Ventas · CxC
# Plataforma de inteligencia logística multisede
#
# Incluye:
# - Recalculo dinámico en tiempo real desde datos base
# - Persistencia local diaria con SQLite (capa de acceso corta, sin
#   conexión global compartida, pensada para concurrencia futura)
# - Inventario + ventas por corte, con lectura Excel acelerada (calamine)
# - Cortes dedicados por proveedor/marca sin sobrescribir el corte maestro
# - Centro de productos sin rotación con trazabilidad histórica
# - Ventas acumuladas / por período / diarias
# - Compras, retiros y redistribución automática entre sedes (con Sankey)
# - ABC / Pareto, rotación y cobertura
# - Punto de reorden + lead time + stock de seguridad
# - Alertas y Health Score
# - Historial de cortes
# - Modo Gerencia / Modo Operativo
# - Exportación a 6 formatos: Excel operativo, Excel gerencial, HTML,
#   PDF ejecutivo, CSV, JSON + paquete ZIP completo
# - Hub de automatización mediante webhook (CRM / n8n / Make / WhatsApp provider)
#
# Requisitos: streamlit, pandas, numpy, openpyxl, plotly, requests
# Recomendado (mejora de rendimiento de lectura Excel): python-calamine
# Opcional (PDF ejecutivo): fpdf2
# INSTALAR (PowerShell / terminal de VS Code, Python 3.10 a 3.13):
# py -m pip install "streamlit>=1.63,<2" "pandas>=2.2,<3" "numpy>=1.26,<3" "plotly>=5.24,<7" "openpyxl>=3.1,<4" "requests>=2.31,<3" python-calamine "xlrd>=2.0.1,<3" "fpdf2>=2.8,<3"
# EJECUTAR: py -3.12 -m streamlit run makropetrol_nexus_v7_3.py
# Guardar este script junto al nexus_data.db anterior para conservar el historial.
# Los esquemas v6/v7 se añaden sin borrar el historial y crean un respaldo previo si hay datos.
# NUEVO V7: Carga única A2, 60 departamentos, 27 proveedores, reglas confirmables,
# mínimos/máximos manuales y reportes separados por gestión.
# CxC: factura, cliente, emisión, vencimiento y saldo; importe original opcional.
# El formato A2 específico de CxC requiere validación con un archivo real.
# Las plantillas y la guía de los datos están dentro de la aplicación.
# ============================================================

from __future__ import annotations

import io
import json
import re
import sqlite3
import unicodedata
import zipfile
import hashlib
import tempfile
import uuid
from html import escape
from concurrent.futures import ThreadPoolExecutor

from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd
import requests
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio

# Tema global NEXUS: evita plantillas heredadas que puedan serializar colores
# negros al exportar HTML y mantiene una identidad visual consistente.
NEXUS_BLUE = "#187C78"
NEXUS_BLUE_2 = "#125C59"
NEXUS_CYAN = "#3298AE"
NEXUS_GREEN = "#14A36A"
NEXUS_ORANGE = "#EF982D"
NEXUS_RED = "#DF5555"
NEXUS_GRAY = "#7B8799"
NEXUS_PALETTE = [NEXUS_BLUE, NEXUS_CYAN, NEXUS_GREEN, NEXUS_ORANGE, NEXUS_RED, NEXUS_BLUE_2, NEXUS_GRAY]
pio.templates.default = "plotly_white"

# Motor de lectura Excel acelerado. Si python-calamine no está instalado,
# NEXUS sigue funcionando normalmente con el motor por defecto de pandas.
try:
    import python_calamine
    EXCEL_ENGINE = "calamine"
except ImportError:
    EXCEL_ENGINE = None

# Generación de PDF ejecutivo (opcional). Si fpdf2 no está instalado, el
# botón de PDF se oculta automáticamente y todo lo demás sigue igual.
try:
    from fpdf import FPDF
    HAS_PDF = True
except Exception:
    HAS_PDF = False


# ============================================================
# CONFIGURACIÓN
# ============================================================

APP_VERSION = "7.3.0"
DB_PATH = Path(__file__).with_name("nexus_data.db")
CONTACTS_PATH = Path(__file__).with_name("nexus_contacts.json")

SEDES = [
    "Centurión",
    "Perrokeros",
    "Puerto La Cruz",
    "Barcelona",
    "Matriz Makropetrol",
]

SEDE_ICON = {
    "Centurión": "◉",
    "Perrokeros": "◈",
    "Puerto La Cruz": "◇",
    "Barcelona": "▣",
    "Matriz Makropetrol": "◆",
}

SALES_MODES = {
    "Ventas del período": "periodo",
    "Ventas acumuladas": "acumuladas",
    "Ventas diarias": "diarias",
}

DEFAULT_CONTACTS = {
    site: {
        "supervisor": "",
        "whatsapp": "",
        "webhook": "",
    }
    for site in SEDES
}

# ============================================================
# DOCUMENTOS OPERATIVOS — CONFIGURACIÓN DE ENCABEZADO
# ============================================================
COMPANY_INFO = {
    "name": "LUBRIGAMA ORIENTE, C.A.",
    "address": "AV FUERZAS ARMADAS, CASA S/N, BARRIO LA ADUANA, BARCELONA, ANZOATEGUI",
    "phone": "0412-033.40.40",
    "rif": "J-50555333-0",
}

DOCUMENT_ORDER_PREFIX = {
    "purchase": "OC",
    "withdrawal": "RET",
    "redistribution": "TR",
    "alerts": "ALT",
}

SUPPLIER_DEFAULTS = {
    "name": "PROVEEDOR / SEGÚN INVENTARIO",
    "address": "",
    "phone": "",
}

DOC_MIME_PDF = "application/pdf"
DOC_MIME_HTML = "text/html"
DOC_MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

WEBHOOK_EXECUTOR = ThreadPoolExecutor(max_workers=2, thread_name_prefix="nexus-webhook")


# ============================================================
# ESTILO
# ============================================================

st.set_page_config(
    page_title="Makropetrol NEXUS | Inventario, Ventas y CxC",
    page_icon="◈",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
<style>
:root{
 --bg:#f5f7fb; --surface:rgba(255,255,255,.92); --surface-2:#ffffff;
 --line:#e5eaf2; --text:#152038; --muted:#768399; --blue:#3867f4;
 --blue2:#2448c9; --green:#11996a; --orange:#ed8b24; --red:#dc5157;
 --cyan:#16a7b7; --violet:#7656e8; --navy:#10192c;
 --shadow-sm:0 7px 24px rgba(29,45,75,.055);
 --shadow:0 18px 46px rgba(29,45,75,.085);
 --radius:20px;
}
*{box-sizing:border-box}
html,body,[class*="css"]{font-family:Inter,ui-sans-serif,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif}
.stApp{
 background:
  radial-gradient(circle at 8% -8%,rgba(56,103,244,.12),transparent 27%),
  radial-gradient(circle at 103% 4%,rgba(22,167,183,.08),transparent 25%),
  linear-gradient(150deg,#fafcff 0%,#f2f6fb 52%,#eef3f8 100%);
 color:var(--text)
}
.block-container{max-width:1580px;padding-top:1.2rem;padding-bottom:3rem}
header[data-testid="stHeader"]{background:transparent}
section[data-testid="stSidebar"]{
 background:linear-gradient(180deg,rgba(16,25,44,.985),rgba(22,35,61,.985));
 border-right:1px solid rgba(255,255,255,.06);
}
section[data-testid="stSidebar"] *{color:#edf3ff}
section[data-testid="stSidebar"] .stCaptionContainer p,
section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p{color:#aebbd0}
section[data-testid="stSidebar"] hr{border-color:rgba(255,255,255,.09)}
section[data-testid="stSidebar"] .nav-group{color:#7f91ad!important}
section[data-testid="stSidebar"] button[kind="secondary"]{
 background:transparent;border:1px solid transparent;color:#c9d5e8;box-shadow:none
}
section[data-testid="stSidebar"] button[kind="secondary"]:hover{
 background:rgba(255,255,255,.06);border-color:rgba(255,255,255,.08)
}
section[data-testid="stSidebar"] button[kind="primary"]{
 background:linear-gradient(135deg,#446ff5,#2e55d9);border:1px solid rgba(255,255,255,.12);
 box-shadow:0 10px 24px rgba(31,70,190,.30)
}
section[data-testid="stSidebar"] [data-baseweb="input"],
section[data-testid="stSidebar"] [data-baseweb="select"]>div{background:rgba(255,255,255,.075)!important;border-color:rgba(255,255,255,.10)!important}
section[data-testid="stSidebar"] input{color:#f3f7ff!important}
section[data-testid="stSidebar"] input::placeholder{color:#8293ad!important}
h1,h2,h3,h4{font-family:Inter,ui-sans-serif,sans-serif!important;color:var(--text)!important;letter-spacing:-.035em}
.hero{
 background:linear-gradient(135deg,rgba(255,255,255,.96),rgba(247,250,255,.86));
 border:1px solid rgba(255,255,255,.96);box-shadow:var(--shadow);border-radius:30px;
 padding:28px 31px 24px;position:relative;overflow:hidden;isolation:isolate
}
.hero:before{content:"";position:absolute;inset:0;background:linear-gradient(90deg,rgba(56,103,244,.035),transparent 48%);z-index:-2}
.hero:after{content:"";position:absolute;width:390px;height:390px;right:-145px;top:-210px;border-radius:50%;background:radial-gradient(circle,rgba(56,103,244,.22),transparent 67%);z-index:-1}
.brand{font-size:2.05rem;font-weight:850;letter-spacing:-.055em}.brand span{color:var(--blue)}
.hero-sub{color:var(--muted);margin-top:5px;font-size:.93rem;max-width:900px}
.status-row{display:flex;gap:9px;align-items:center;flex-wrap:wrap;margin-top:15px;color:#66758c;font-size:.76rem}
.dot{width:8px;height:8px;border-radius:50%;background:#18b879;box-shadow:0 0 0 5px rgba(24,184,121,.11);animation:pulse 2s ease-in-out infinite}
@keyframes pulse{0%,100%{box-shadow:0 0 0 5px rgba(24,184,121,.10)}50%{box-shadow:0 0 0 9px rgba(24,184,121,.025)}}
.section-title{font-size:1.06rem;font-weight:850;margin:23px 0 10px;letter-spacing:-.025em}
.section-kicker{font-size:.69rem;text-transform:uppercase;letter-spacing:.11em;font-weight:850;color:var(--blue);margin-bottom:3px}
.kpi-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-top:17px}
.kpi{
 background:linear-gradient(145deg,rgba(255,255,255,.97),rgba(250,252,255,.90));
 border:1px solid rgba(255,255,255,.98);box-shadow:var(--shadow-sm);border-radius:21px;
 padding:18px 19px;min-height:121px;position:relative;overflow:hidden;transition:transform .18s ease,box-shadow .18s ease
}
.kpi:hover{transform:translateY(-2px);box-shadow:0 16px 38px rgba(29,45,75,.10)}
.kpi:before{content:"";position:absolute;left:0;top:0;bottom:0;width:4px;background:var(--blue)}
.kpi.green:before{background:var(--green)}.kpi.orange:before{background:var(--orange)}.kpi.red:before{background:var(--red)}.kpi.cyan:before{background:var(--cyan)}
.kpi-label{font-size:.71rem;color:var(--muted);font-weight:850;letter-spacing:.055em;text-transform:uppercase}
.kpi-value{font-size:1.72rem;font-weight:850;letter-spacing:-.045em;margin-top:9px}.kpi-note{font-size:.71rem;color:#98a4b5;margin-top:4px}
.panel{background:var(--surface);border:1px solid rgba(255,255,255,.97);box-shadow:var(--shadow-sm);border-radius:22px;padding:18px}
.glass-strip{background:rgba(255,255,255,.72);backdrop-filter:blur(12px);border:1px solid rgba(255,255,255,.88);border-radius:18px;padding:12px 14px;box-shadow:var(--shadow-sm)}
.badge{display:inline-flex;align-items:center;border-radius:999px;padding:5px 10px;font-size:.69rem;font-weight:850}
.badge-blue{background:#edf2ff;color:#355bd7}.badge-green{background:#e9f8f2;color:#137a52}.badge-orange{background:#fff2df;color:#b86a0c}.badge-red{background:#ffeded;color:#c43c43}.badge-gray{background:#eef1f5;color:#68758a}.badge-cyan{background:#e9f9fb;color:#147a86}
.stButton>button,.stDownloadButton>button{border-radius:13px;font-weight:780;min-height:43px;transition:transform .15s ease,box-shadow .15s ease}
.stButton>button:hover,.stDownloadButton>button:hover{transform:translateY(-1px)}
.stDownloadButton>button{border:0;background:linear-gradient(135deg,var(--blue),var(--blue2));color:#fff;box-shadow:0 9px 22px rgba(53,104,245,.22)}
[data-testid="stFileUploaderDropzone"]{border:1.5px dashed #c7d3e4!important;border-radius:17px!important;background:rgba(255,255,255,.75)!important;transition:border-color .2s ease,background .2s ease}
[data-testid="stFileUploaderDropzone"]:hover{border-color:#7d9bf5!important;background:#f8faff!important}
.stTabs [data-baseweb="tab-list"]{gap:6px;background:rgba(255,255,255,.72);border:1px solid var(--line);padding:6px;border-radius:16px;box-shadow:var(--shadow-sm)}
.stTabs [data-baseweb="tab"]{border-radius:11px;color:#7a8799;font-weight:760;padding:9px 15px}
.stTabs [aria-selected="true"]{background:#edf2ff!important;color:var(--blue)!important}
div[data-testid="stDataFrame"]{border:1px solid var(--line);border-radius:17px;overflow:hidden;box-shadow:0 7px 20px rgba(29,45,75,.035)}
[data-baseweb="input"],[data-baseweb="select"]>div{border-radius:12px!important}
.footer{text-align:center;color:#9aa6b6;font-size:.69rem;margin-top:34px}
.empty{background:rgba(255,255,255,.74);border:1px dashed #cbd5e4;border-radius:22px;padding:36px;text-align:center;color:#758195;box-shadow:var(--shadow-sm)}
.empty-title{font-weight:850;font-size:1.05rem;color:#273550}.empty-icon{font-size:2rem;margin-bottom:7px}
.nav-group{font-size:.64rem;font-weight:850;letter-spacing:.11em;text-transform:uppercase;margin:15px 0 3px 2px}
.gov-banner{display:flex;align-items:center;gap:8px;background:linear-gradient(135deg,#121d33,#22365b);color:#fff;border-radius:15px;padding:10px 14px;font-size:.78rem;font-weight:760;margin-bottom:10px;box-shadow:0 12px 30px rgba(16,25,44,.18)}
.health-big{display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;padding:20px 10px}
.health-big .num{font-size:3rem;font-weight:850;letter-spacing:-.055em;color:var(--blue)}
.health-big .lbl{font-size:.72rem;color:var(--muted);font-weight:850;letter-spacing:.075em;text-transform:uppercase;margin-top:2px}
.op-row{display:flex;justify-content:space-between;align-items:center;padding:9px 4px;border-bottom:1px solid var(--line);font-size:.85rem;gap:8px}.op-row:last-child{border-bottom:none}
.bar-track{background:#edf1f6;border-radius:8px;height:10px;overflow:hidden;flex:1;margin:0 10px}.bar-fill{height:100%;border-radius:8px;background:linear-gradient(90deg,var(--blue),var(--cyan))}
.opp-banner{display:flex;gap:22px;flex-wrap:wrap;background:linear-gradient(135deg,#e9f9f1,#edf5ff);border:1px solid #d8e9eb;border-radius:17px;padding:14px 18px;font-size:.85rem;font-weight:760;color:#245a4b;box-shadow:var(--shadow-sm)}
.scope-banner{display:flex;align-items:center;justify-content:space-between;gap:14px;flex-wrap:wrap;background:linear-gradient(135deg,#f2f5ff,#eefbfb);border:1px solid #dfe7f4;border-radius:18px;padding:13px 16px;margin:10px 0 16px}
.scope-title{font-weight:850;color:#24324a}.scope-sub{font-size:.75rem;color:#748198;margin-top:2px}
.metric-chip{display:inline-flex;align-items:center;gap:6px;padding:7px 10px;border-radius:999px;background:white;border:1px solid var(--line);font-size:.72rem;font-weight:780;color:#536177}
@media(max-width:1000px){.kpi-grid{grid-template-columns:repeat(2,1fr)}}
@media(max-width:650px){.kpi-grid{grid-template-columns:1fr}.hero{padding:21px}.opp-banner{flex-direction:column;gap:8px}.brand{font-size:1.65rem}}

/* NEXUS 6: jerarquía visual, legibilidad y navegación compacta. */
:root{--blue:#187c78;--blue2:#125c59;--cyan:#3298ae;--muted:#5f7282;--text:#182e3b;--line:#dfe7eb;--shadow-sm:0 3px 16px rgba(20,48,63,.045)}
.stApp{background:#f3f6f8;color:#182e3b}
.block-container{max-width:1660px;padding-top:1.4rem}
.hero{background:linear-gradient(115deg,#142b38,#20434c);border:0;border-radius:20px;padding:25px 30px;box-shadow:0 10px 30px #142b3810;margin-bottom:14px}
.hero .brand{color:#fff;font-size:1.8rem;font-weight:750}.hero .brand span{color:#79dbca}
.hero-sub{color:#c2d4dc}.status-row{color:#bcd0d8}.hero:before,.hero:after{display:none}
.section-title{font-size:1.08rem;font-weight:750;margin:22px 0 13px}
.kpi{background:#fff;border:1px solid #dfe7eb;border-radius:16px;padding:17px 19px;min-height:127px;box-shadow:var(--shadow-sm)}
.kpi:before{left:18px;top:0;bottom:auto;width:34px;height:3px;border-radius:0 0 4px 4px}
.kpi-label{font-size:.68rem;color:#59717d;letter-spacing:.055em}.kpi-value{font-size:clamp(1.15rem,1.6vw,1.75rem);color:#193643;overflow-wrap:anywhere}.kpi-note{color:#627783;line-height:1.5}
.panel{background:#fff;border:1px solid #dfe7eb;border-radius:16px}.health-big{flex-direction:row;gap:20px;padding:14px}.health-big .num{font-size:2rem}.health-big .lbl{font-size:.76rem}
section[data-testid="stSidebar"]{background:#142b38;border:0}
section[data-testid="stSidebar"] .stButton button{min-height:35px;font-size:.83rem;border-radius:9px;justify-content:flex-start}
section[data-testid="stSidebar"] button[kind="primary"]{background:#207e78;border:1px solid #40968e;box-shadow:none}
section[data-testid="stSidebar"] [data-testid="stVerticalBlock"]{gap:.3rem}
section[data-testid="stSidebar"] .nav-group{margin-top:13px;letter-spacing:.1em;color:#9ab3bf!important}
section[data-testid="stSidebar"] h1,section[data-testid="stSidebar"] h2,section[data-testid="stSidebar"] h3{color:#eef6fa!important}
section[data-testid="stSidebar"] label{color:#e1edf4!important}
section[data-testid="stSidebar"] .badge{color:#e3f6f1;background:#264c50}
.stTabs [data-baseweb="tab-list"]{background:#e9eff2;border:0;border-radius:12px;padding:5px;gap:3px;box-shadow:none}
.stTabs [data-baseweb="tab"]{color:#4e6877;font-weight:650;white-space:nowrap}
.stTabs [aria-selected="true"]{background:#fff!important;color:#146e68!important;box-shadow:0 2px 6px #1936430b}
.stButton>button,.stDownloadButton>button{font-weight:650;border-radius:10px}
.stButton>button[kind="primary"]{background:#187c78;border-color:#187c78;color:white}
.stDownloadButton>button{background:#187c78;box-shadow:none}
main [data-testid="stWidgetLabel"] p{color:#34515e}
main [data-baseweb="input"],main [data-baseweb="select"]>div{background:white;color:#182e3b}
.board-heading{display:flex;justify-content:space-between;font-weight:750;border-top:3px solid;padding:13px 6px;margin:12px 0;color:#274653}
.board-heading span{border-radius:12px;background:#e1e9ee;padding:0 9px;color:#466371}
.task-card{background:#fff;border:1px solid #dce5ea;border-radius:14px;padding:16px;margin-bottom:12px;overflow-wrap:anywhere;min-height:154px;box-shadow:0 3px 10px #19364305}
.task-card strong{display:block;font-size:.91rem;line-height:1.5;color:#223f4d}.task-site{font-size:.67rem;color:#187c78;margin-bottom:9px;text-transform:uppercase;letter-spacing:.04em}.task-owner{font-size:.78rem;color:#667b87;margin:12px 0 3px}.task-card small{font-size:.72rem}
@media(prefers-reduced-motion:reduce){*{animation:none!important;transition:none!important}}
@media(max-width:900px){.kpi-value{font-size:1.3rem}.block-container{padding-left:1rem;padding-right:1rem}.hero{padding:21px}.stTabs [data-baseweb="tab-list"]{overflow-x:auto}.hero .brand{font-size:1.55rem}}
</style>
""",
    unsafe_allow_html=True,
)


# ============================================================
# UTILIDADES
# ============================================================

def money(v: Any) -> str:
    try:
        value = float(v)
        return f"${value:,.2f}" if np.isfinite(value) else "N/D"
    except (TypeError, ValueError, OverflowError):
        return "N/D"


def integer(v: Any) -> str:
    try:
        return f"{int(round(float(v))):,}"
    except Exception:
        return "0"


def safe_float(v: Any, default: float = 0.0) -> float:
    try:
        x = float(v)
        return x if np.isfinite(x) else default
    except Exception:
        return default


def normalize_text(v: Any) -> str:
    x = str(v).strip().lower()
    x = re.sub(r"\s+", " ", x)
    return x



def normalize_key(v: Any) -> str:
    """Clave estable para búsquedas de proveedor/marca y comparaciones humanas."""
    x = normalize_text(v)
    x = unicodedata.normalize("NFKD", x)
    x = "".join(ch for ch in x if not unicodedata.combining(ch))
    x = re.sub(r"[^a-z0-9]+", " ", x)
    return re.sub(r"\s+", " ", x).strip()


def safe_slug(v: Any, fallback: str = "reporte") -> str:
    x = normalize_key(v).replace(" ", "_")
    return x[:80] or fallback


def supplier_matches(query: str, suppliers: list[str], limit: int = 12) -> list[str]:
    q = normalize_key(query)
    if not q:
        return suppliers[:limit]
    starts, contains = [], []
    for supplier in suppliers:
        k = normalize_key(supplier)
        if k.startswith(q):
            starts.append(supplier)
        elif q in k:
            contains.append(supplier)
    return (starts + contains)[:limit]


def parse_number(v: Any) -> float:
    if pd.isna(v):
        return 0.0
    if isinstance(v, (int, float, np.integer, np.floating)):
        return float(v)
    s = str(v).strip().replace("$", "").replace("€", "").replace(" ", "")
    if not s:
        return 0.0
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        parts = s.split(",")
        if len(parts[-1]) in (1, 2):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    else:
        pass
    try:
        return float(s)
    except Exception:
        return 0.0


def empty_state(title: str, message: str, icon: str = "◌") -> None:
    st.markdown(
        f"""
        <div class="empty">
            <div class="empty-icon">{icon}</div>
            <div class="empty-title">{title}</div>
            <div>{message}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def kpi(label: str, value: str, note: str = "", tone: str = "blue") -> None:
    st.markdown(
        f"""
        <div class="kpi {tone}">
            <div class="kpi-label">{escape(str(label))}</div>
            <div class="kpi-value">{escape(str(value))}</div>
            <div class="kpi-note">{escape(str(note))}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def chart_layout(fig: go.Figure, height: int = 330) -> go.Figure:
    fig.update_layout(
        template="plotly_white",
        height=height,
        margin=dict(l=10, r=10, t=42, b=10),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="DM Sans", color="#59677b"),
        legend=dict(orientation="h", y=1.03, x=0),
        hoverlabel=dict(bgcolor="white", font_size=13, font_family="DM Sans"),
    )
    return fig


def present_columns(df: pd.DataFrame, preferred: list[str]) -> list[str]:
    return [c for c in preferred if c in df.columns]


def build_executive_frame(all_df: pd.DataFrame) -> pd.DataFrame:
    if all_df.empty:
        return pd.DataFrame()
    preferred = [
        "Sede", "Código", "Descripción", "ABC", "Estado", "Prioridad", "Acción",
        "Ventas", "Demanda Mensual", "Existencia", "Stock Mínimo", "Stock Máximo",
        "Punto de Reorden", "Compra Ajustada", "Retiro Almacén", "Cobertura (meses)",
        "Valor Inventario ($)", "Capital Inmovilizado ($)", "Costo",
    ]
    return all_df[present_columns(all_df, preferred)].copy()


def export_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")


def export_json_bytes(df: pd.DataFrame) -> bytes:
    return df.to_json(orient="records", force_ascii=False, indent=2).encode("utf-8")


@st.cache_data(show_spinner=False)
def export_xlsx_bytes(df: pd.DataFrame, sheet_name: str = "Datos") -> bytes:
    from openpyxl.styles import Font, PatternFill

    out = io.BytesIO()
    safe_sheet = re.sub(r"[\\/*?:\[\]]", "_", str(sheet_name))[:31] or "Datos"
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=safe_sheet)
        ws = writer.book[safe_sheet]
        ws.freeze_panes = "A2"
        ws.sheet_view.showGridLines = False
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="3568F5")
        for col_cells in ws.columns:
            values = [str(c.value or "") for c in list(col_cells)[:120]]
            width = min(max(max((len(v) for v in values), default=8) + 2, 11), 42)
            ws.column_dimensions[col_cells[0].column_letter].width = width
    return out.getvalue()


@st.cache_data(show_spinner=False)
def build_interactive_html_report(
    all_df: pd.DataFrame,
    transfers: pd.DataFrame,
    alerts: pd.DataFrame,
    snapshot_label: str,
    score: int,
) -> bytes:
    if all_df.empty:
        body = "<h2>Sin datos disponibles</h2><p>Procese al menos un corte diario.</p>"
        return f"<!doctype html><html lang='es'><meta charset='utf-8'><body>{body}</body></html>".encode("utf-8")

    purchase_col = "Compra Ajustada" if "Compra Ajustada" in all_df.columns else "Compra Sugerida"
    site_summary = all_df.groupby("Sede").agg(
        Inventario=("Valor Inventario ($)", "sum"),
        Críticos=("Estado", lambda x: int(x.astype(str).str.startswith("CRÍTICO").sum())),
        Compras=(purchase_col, lambda x: int((x > 0).sum())),
        Retiros=("Retiro Almacén", lambda x: int((x > 0).sum())),
    ).reset_index()

    fig1 = px.bar(site_summary, x="Sede", y="Inventario", title="Valor de inventario por sede", color="Sede", color_discrete_sequence=NEXUS_PALETTE)
    fig2 = px.bar(site_summary, x="Sede", y=["Compras", "Retiros"], barmode="group", title="Acciones por sede", color_discrete_sequence=[NEXUS_ORANGE, NEXUS_BLUE])
    state = all_df["Estado"].value_counts().reset_index()
    state.columns = ["Estado", "Cantidad"]
    fig3 = px.pie(state, names="Estado", values="Cantidad", hole=.62, title="Distribución del estado logístico", color="Estado", color_discrete_map={"ÓPTIMO":NEXUS_GREEN,"CRÍTICO — SALDO NEGATIVO":NEXUS_RED,"CRÍTICO — COMPRAR":NEXUS_ORANGE,"SOBRESTOCK — RETIRAR":NEXUS_BLUE})

    top = all_df[all_df["Capital Inmovilizado ($)"] > 0].sort_values("Capital Inmovilizado ($)", ascending=False).head(12)
    fig4 = px.bar(top.sort_values("Capital Inmovilizado ($)"), x="Capital Inmovilizado ($)", y="Descripción", orientation="h", title="Top capital inmovilizado", color_discrete_sequence=[NEXUS_RED]) if not top.empty else None

    plotly_bundle_included = False

    def fig_html(fig):
        # Incluir plotly.js una sola vez reduce el HTML de ~20 MB a ~5 MB
        # cuando el dashboard contiene varios gráficos.
        nonlocal plotly_bundle_included
        include_js = "inline" if not plotly_bundle_included else False
        plotly_bundle_included = True
        return pio.to_html(
            fig, full_html=False, include_plotlyjs=include_js,
            config={"responsive": True, "displaylogo": False, "modeBarButtonsToRemove": ["lasso2d", "select2d"]},
        )

    exec_df = build_executive_frame(all_df).head(40)
    alert_html = alerts.head(30).to_html(index=False, classes="dataframe") if not alerts.empty else "<p>Sin alertas.</p>"
    transfer_html = transfers.head(30).to_html(index=False, classes="dataframe") if not transfers.empty else "<p>Sin transferencias sugeridas.</p>"
    table_html = exec_df.to_html(index=False, classes="dataframe")

    html = f"""<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Makropetrol NEXUS · Reporte Ejecutivo</title>
<style>
body{{margin:0;background:#f4f7fb;color:#172137;font-family:Inter,Segoe UI,Arial,sans-serif}}
.wrap{{max-width:1400px;margin:auto;padding:38px}}
.hero{{background:linear-gradient(135deg,#fff,#f0f5ff);border:1px solid #e2e8f2;border-radius:28px;padding:30px;box-shadow:0 18px 50px rgba(30,50,80,.10)}}
.brand{{font-size:28px;font-weight:800}} .brand span{{color:#3568f5}}
.muted{{color:#748198}} .score{{font-size:42px;font-weight:800;color:#3568f5}}
.kpis{{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-top:18px}}
.kpi{{background:#fff;border:1px solid #e3e9f1;border-radius:20px;padding:18px;box-shadow:0 8px 24px rgba(30,50,80,.07)}}
.kpi small{{color:#7d899b;text-transform:uppercase;font-weight:800;font-size:11px}} .kpi strong{{display:block;font-size:27px;margin-top:6px}}
.grid{{display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-top:18px}} .panel{{background:#fff;border:1px solid #e3e9f1;border-radius:20px;padding:18px;box-shadow:0 8px 24px rgba(30,50,80,.06)}}
.dataframe{{width:100%;border-collapse:collapse;font-size:12px}} .dataframe th{{background:#3568f5;color:#fff;padding:9px;text-align:left}} .dataframe td{{padding:8px;border-bottom:1px solid #e5eaf1}} .dataframe tr:nth-child(even){{background:#f9fbfd}}
@media(max-width:900px){{.kpis,.grid{{grid-template-columns:1fr 1fr}}}} @media(max-width:600px){{.kpis,.grid{{grid-template-columns:1fr}} .wrap{{padding:18px}}}}
</style>

</head>
<body><div class="wrap">
<div class="hero"><div class="brand">MAKROPETROL <span>NEXUS</span></div>
<div class="muted">Reporte ejecutivo interactivo · Corte {snapshot_label}</div>
<div style="margin-top:16px"><span class="muted">Salud logística</span><div class="score">{score}/100</div></div></div>
<div class="kpis">
<div class="kpi"><small>SKUs</small><strong>{len(all_df):,}</strong></div>
<div class="kpi"><small>Críticos</small><strong>{int(all_df['Estado'].astype(str).str.startswith('CRÍTICO').sum()):,}</strong></div>
<div class="kpi"><small>Compra neta</small><strong>{int(all_df[purchase_col].sum()):,}</strong></div>
<div class="kpi"><small>Capital inmovilizado</small><strong>${float(all_df['Capital Inmovilizado ($)'].sum()):,.2f}</strong></div>
</div>
<div class="grid">
<div class="panel">{fig_html(fig1)}</div><div class="panel">{fig_html(fig2)}</div>
<div class="panel">{fig_html(fig3)}</div>{f'<div class="panel">{fig_html(fig4)}</div>' if fig4 is not None else '<div class="panel"><p class="muted">Sin capital inmovilizado detectado.</p></div>'}
</div>
<div class="panel" style="margin-top:18px"><h2>Vista ejecutiva de productos</h2>{table_html}</div>
<div class="panel" style="margin-top:18px"><h2>Alertas</h2>{alert_html}</div>
<div class="panel" style="margin-top:18px"><h2>Redistribución</h2>{transfer_html}</div>
</div></body></html>"""
    return html.encode("utf-8")


@st.cache_data(show_spinner=False)
def build_executive_excel(all_df: pd.DataFrame, snapshot_label: str, score: int) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.utils import get_column_letter

    out = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumen Gerencial"

    navy, blue, white, gray = "18243A", "3568F5", "FFFFFF", "7B8799"
    ws.merge_cells("A1:F1")
    ws["A1"] = "MAKROPETROL NEXUS · RESUMEN GERENCIAL"
    ws["A1"].font = Font(size=18, bold=True, color=white)
    ws["A1"].fill = PatternFill("solid", fgColor=navy)
    ws.row_dimensions[1].height = 32
    ws.merge_cells("A2:F2")
    ws["A2"] = f"Corte {snapshot_label} · Salud logística {score}/100 · {health_label(score)}"
    ws["A2"].font = Font(size=10, italic=True, color=gray)

    if not all_df.empty:
        purchase_col = "Compra Ajustada" if "Compra Ajustada" in all_df.columns else "Compra Sugerida"
        by_site = all_df.groupby("Sede").agg(
            Inventario=("Valor Inventario ($)", "sum"),
            Críticos=("Estado", lambda x: int(x.astype(str).str.startswith("CRÍTICO").sum())),
            Compras=(purchase_col, lambda x: int((x > 0).sum())),
            Inmovilizado=("Capital Inmovilizado ($)", "sum"),
        ).reset_index()
        start = 4
        for j, col in enumerate(by_site.columns, 1):
            c = ws.cell(start, j, col)
            c.font = Font(bold=True, color=white)
            c.fill = PatternFill("solid", fgColor=blue)
        for i, row in enumerate(by_site.itertuples(index=False), start + 1):
            for j, val in enumerate(row, 1):
                ws.cell(i, j, val)
        ref = f"A{start}:{get_column_letter(len(by_site.columns))}{start + len(by_site)}"
        table = Table(displayName="ResumenGerencial", ref=ref)
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(table)
        for j in range(1, len(by_site.columns) + 1):
            ws.column_dimensions[get_column_letter(j)].width = 20

        exec_df = build_executive_frame(all_df).head(500)
        keep_cols = [c for c in ["Sede", "Código", "Descripción", "ABC", "Estado", "Acción", "Cobertura (meses)"] if c in exec_df.columns]
        sh2 = wb.create_sheet("Acciones recomendadas")
        for j, col in enumerate(keep_cols, 1):
            c = sh2.cell(1, j, col)
            c.font = Font(bold=True, color=white)
            c.fill = PatternFill("solid", fgColor=blue)
        for i, row in enumerate(exec_df[keep_cols].itertuples(index=False), 2):
            for j, val in enumerate(row, 1):
                sh2.cell(i, j, val)
        for j in range(1, len(keep_cols) + 1):
            sh2.column_dimensions[get_column_letter(j)].width = 22

    wb.save(out)
    return out.getvalue()


@st.cache_data(show_spinner=False)
def build_executive_pdf(all_df: pd.DataFrame, transfers: pd.DataFrame, snapshot_label: str, score: int) -> bytes | None:
    if not HAS_PDF:
        return None
    purchase_col = "Compra Ajustada" if "Compra Ajustada" in all_df.columns else "Compra Sugerida"
    critical = int(all_df["Estado"].astype(str).str.startswith("CRÍTICO").sum()) if not all_df.empty else 0
    purchases = int((all_df[purchase_col] > 0).sum()) if not all_df.empty else 0
    inv_value = float(monetary_sum_v71(all_df["Valor Inventario ($)"])) if not all_df.empty else 0.0
    immobilized = float(monetary_sum_v71(all_df["Capital Inmovilizado ($)"])) if not all_df.empty else 0.0
    avoided = float(monetary_sum_v71(transfers["Compra Evitada Estimada ($)"])) if not transfers.empty else 0.0

    pdf = FPDF()
    pdf.add_page()
    pdf.set_fill_color(24, 36, 58)
    pdf.rect(0, 0, 210, 28, "F")
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 18)
    pdf.set_xy(10, 8)
    pdf.cell(0, 10, "MAKROPETROL NEXUS - Reporte Ejecutivo")
    pdf.set_font("Helvetica", "", 10)
    pdf.set_xy(10, 18)
    pdf.cell(0, 8, f"Corte {snapshot_label}  ·  Salud logistica {score}/100 ({health_label(score)})")

    pdf.set_text_color(30, 30, 30)
    pdf.set_xy(10, 36)
    pdf.set_font("Helvetica", "B", 13)
    pdf.cell(0, 8, "Indicadores clave", ln=True)
    pdf.set_font("Helvetica", "", 11)
    rows = [
        ("Valor de inventario", money(inv_value)),
        ("SKUs criticos", f"{critical:,}"),
        ("SKUs a comprar", f"{purchases:,}"),
        ("Capital inmovilizado", money(immobilized)),
        ("Compra evitada por redistribucion", money(avoided)),
    ]
    for label, val in rows:
        pdf.set_x(12)
        pdf.cell(90, 8, label)
        pdf.cell(0, 8, val, ln=True)

    if not all_df.empty:
        pdf.ln(4)
        pdf.set_font("Helvetica", "B", 13)
        pdf.cell(0, 8, "Inventario por sede", ln=True)
        pdf.set_font("Helvetica", "", 10)
        by_site = all_df.groupby("Sede")["Valor Inventario ($)"].sum().sort_values(ascending=False)
        for sede, val in by_site.items():
            pdf.set_x(12)
            pdf.cell(90, 7, str(sede))
            pdf.cell(0, 7, money(val), ln=True)

    pdf.ln(4)
    pdf.set_font("Helvetica", "I", 9)
    pdf.set_text_color(120, 120, 120)
    pdf.multi_cell(0, 6, f"Generado automaticamente por NEXUS v{APP_VERSION} el {datetime.now().strftime('%d/%m/%Y %H:%M')}.")

    return bytes(pdf.output())


def build_readme_text(snapshot_label: str) -> bytes:
    text = f"""MAKROPETROL NEXUS · Paquete de reporte
Corte: {snapshot_label}
Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}

Contenido del paquete:
- Dashboard_Gerencial.html  → dashboard interactivo con gráficos Plotly, se abre en cualquier navegador.
- Reporte_Gerencial.xlsx    → resumen ejecutivo por sede + acciones recomendadas.
- Reporte_Operativo.xlsx    → detalle completo (inventario, compras, retiros, redistribución, alertas).
- Datos.csv / Datos.json    → datos consolidados para Power BI, APIs o automatizaciones.
- Reporte_Ejecutivo.pdf     → una página lista para enviar por WhatsApp o correo (si fpdf2 está instalado).
- Documentos Operativos      → pedidos de compra por sede, retiros por sede, alertas por sede y redistribución multisede.
- ZIP Documentos Operativos → paquete separado por tienda + redistribución global.

NEXUS v{APP_VERSION} · datos base persistidos en SQLite · recálculo dinámico.
"""
    return text.encode("utf-8")


@st.cache_data(show_spinner=False)
def export_bundle(
    excel_bytes: bytes,
    all_df: pd.DataFrame,
    transfers: pd.DataFrame,
    alerts: pd.DataFrame,
    snapshot_label: str,
    score: int,
) -> bytes:
    html = build_interactive_html_report(all_df, transfers, alerts, snapshot_label, score)
    gerencial = build_executive_excel(all_df, snapshot_label, score)
    pdf_bytes = build_executive_pdf(all_df, transfers, snapshot_label, score)
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(f"Reporte_Operativo_{snapshot_label}.xlsx", excel_bytes)
        z.writestr(f"Reporte_Gerencial_{snapshot_label}.xlsx", gerencial)
        z.writestr(f"Datos_{snapshot_label}.csv", export_csv_bytes(all_df))
        z.writestr(f"Datos_{snapshot_label}.json", export_json_bytes(all_df))
        z.writestr(f"Dashboard_Gerencial_{snapshot_label}.html", html)
        if pdf_bytes:
            z.writestr(f"Reporte_Ejecutivo_{snapshot_label}.pdf", pdf_bytes)
        z.writestr("README.txt", build_readme_text(snapshot_label))
    return out.getvalue()


# ============================================================
# SQLite — PERSISTENCIA DIARIA
# ============================================================

class ClosingConnection(sqlite3.Connection):
    """Confirma/revierte y cierra cada conexión, también ante errores."""
    def __exit__(self, exc_type, exc_value, traceback):
        try:
            return super().__exit__(exc_type, exc_value, traceback)
        finally:
            self.close()


def insert_sql_frame(con: sqlite3.Connection, table: str, frame: pd.DataFrame) -> None:
    allowed = {"inventory_snapshots", "sales_snapshots", "supplier_inventory_snapshots", "supplier_sales_snapshots"}
    if table not in allowed:
        raise ValueError("Tabla no admitida")
    if frame.empty:
        return
    columns = ",".join('"' + c + '"' for c in frame.columns)
    markers = ",".join("?" for _ in frame.columns)
    # executemany evita límites de variables de SQLite y conserva la transacción completa.
    con.executemany(f'INSERT INTO {table} ({columns}) VALUES ({markers})', frame.itertuples(index=False, name=None))


def db() -> sqlite3.Connection:
    con = sqlite3.connect(DB_PATH, timeout=15, factory=ClosingConnection)
    con.execute("PRAGMA journal_mode=WAL")
    return con


def init_db() -> None:
    with db() as con:
        con.execute(
            """
            CREATE TABLE IF NOT EXISTS inventory_snapshots (
                snapshot_date TEXT NOT NULL,
                site TEXT NOT NULL,
                code TEXT NOT NULL,
                description TEXT,
                existence REAL NOT NULL DEFAULT 0,
                cost REAL NOT NULL DEFAULT 0,
                supplier TEXT,
                category TEXT,
                PRIMARY KEY(snapshot_date, site, code)
            )
            """
        )
        con.execute(
            """
            CREATE TABLE IF NOT EXISTS sales_snapshots (
                snapshot_date TEXT NOT NULL,
                site TEXT NOT NULL,
                code TEXT NOT NULL,
                description TEXT,
                sales REAL NOT NULL DEFAULT 0,
                PRIMARY KEY(snapshot_date, site, code)
            )
            """
        )
        con.execute(
            """
            CREATE TABLE IF NOT EXISTS processing_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                processed_at TEXT NOT NULL,
                snapshot_date TEXT NOT NULL,
                site TEXT NOT NULL,
                inventory_rows INTEGER,
                sales_rows INTEGER,
                user_label TEXT DEFAULT ''
            )
            """
        )
        con.execute(
            """
            CREATE TABLE IF NOT EXISTS supplier_inventory_snapshots (
                snapshot_date TEXT NOT NULL,
                site TEXT NOT NULL,
                supplier_key TEXT NOT NULL,
                supplier_name TEXT NOT NULL,
                code TEXT NOT NULL,
                description TEXT,
                existence REAL NOT NULL DEFAULT 0,
                cost REAL NOT NULL DEFAULT 0,
                category TEXT,
                PRIMARY KEY(snapshot_date, site, supplier_key, code)
            )
            """
        )
        con.execute(
            """
            CREATE TABLE IF NOT EXISTS supplier_sales_snapshots (
                snapshot_date TEXT NOT NULL,
                site TEXT NOT NULL,
                supplier_key TEXT NOT NULL,
                supplier_name TEXT NOT NULL,
                code TEXT NOT NULL,
                description TEXT,
                sales REAL NOT NULL DEFAULT 0,
                PRIMARY KEY(snapshot_date, site, supplier_key, code)
            )
            """
        )
        con.execute(
            """
            CREATE TABLE IF NOT EXISTS supplier_processing_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                processed_at TEXT NOT NULL,
                snapshot_date TEXT NOT NULL,
                site TEXT NOT NULL,
                supplier_key TEXT NOT NULL,
                supplier_name TEXT NOT NULL,
                inventory_rows INTEGER,
                sales_rows INTEGER
            )
            """
        )
        # Índices para MAX(fecha), filtros por sede/proveedor y cargas históricas.
        con.execute("CREATE INDEX IF NOT EXISTS idx_inv_site_date ON inventory_snapshots(site, snapshot_date)")
        con.execute("CREATE INDEX IF NOT EXISTS idx_sales_site_date ON sales_snapshots(site, snapshot_date)")
        con.execute("CREATE INDEX IF NOT EXISTS idx_sup_inv_key_site_date ON supplier_inventory_snapshots(supplier_key, site, snapshot_date)")
        con.execute("CREATE INDEX IF NOT EXISTS idx_sup_sales_key_site_date ON supplier_sales_snapshots(supplier_key, site, snapshot_date)")
        con.execute("CREATE INDEX IF NOT EXISTS idx_sup_log_key_date ON supplier_processing_log(supplier_key, snapshot_date)")
        con.commit()


def reset_database() -> None:
    with db() as con:
        con.execute("DELETE FROM inventory_snapshots")
        con.execute("DELETE FROM sales_snapshots")
        con.execute("DELETE FROM processing_log")
        con.execute("DELETE FROM supplier_inventory_snapshots")
        con.execute("DELETE FROM supplier_sales_snapshots")
        con.execute("DELETE FROM supplier_processing_log")
        for table in ["financial_sales", "financial_ar", "financial_ar_batches", "financial_settings", "financial_log", "action_tasks"]:
            con.execute(f"DELETE FROM {table}")
        for table in ["v73_stock_basis", "v71_wh_events", "v71_wh_barcodes", "v71_wh_items", "v71_wh_sessions", "v7_sales_items", "v7_sales_batches", "v7_import_log", "v7_import_archive", "v7_site_config", "v7_stock_levels", "v7_product_dimensions"]:
            con.execute(f"DELETE FROM {table}")
        con.execute("UPDATE v7_department_rules SET supplier='',brand='',category='',match_text='',confirmed=0")
        touch_v7(con)
        con.commit()
    st.cache_data.clear()


def db_version() -> str:
    with db() as con:
        row = con.execute(
            "SELECT COALESCE(MAX(processed_at),'') FROM processing_log"
        ).fetchone()
    return str(row[0] or "") + ":" + v7_version()


def latest_snapshot_date(site: str) -> str | None:
    with db() as con:
        row = con.execute(
            "SELECT MAX(snapshot_date) FROM inventory_snapshots WHERE site=?",
            (site,),
        ).fetchone()
    return row[0] if row and row[0] else None


def previous_snapshot_date(site: str, current_date: str) -> str | None:
    with db() as con:
        row = con.execute(
            "SELECT MAX(snapshot_date) FROM sales_snapshots WHERE site=? AND snapshot_date<?",
            (site, current_date),
        ).fetchone()
    return row[0] if row and row[0] else None


def list_snapshot_dates(site: str | None = None) -> list[str]:
    with db() as con:
        if site:
            rows = con.execute(
                "SELECT DISTINCT snapshot_date FROM inventory_snapshots WHERE site=? ORDER BY snapshot_date DESC",
                (site,),
            ).fetchall()
        else:
            rows = con.execute(
                "SELECT DISTINCT snapshot_date FROM inventory_snapshots ORDER BY snapshot_date DESC"
            ).fetchall()
    return [r[0] for r in rows]


def save_snapshot(
    site: str,
    snapshot_date: str,
    inventory: pd.DataFrame,
    sales: pd.DataFrame,
) -> None:
    inv = inventory
    sal = sales

    inv_shaped = pd.DataFrame({
        "snapshot_date": snapshot_date,
        "site": site,
        "code": inv.get("Código", pd.Series("", index=inv.index)).astype(str),
        "description": inv.get("Descripción", pd.Series("", index=inv.index)).fillna("").astype(str)
        if "Descripción" in inv.columns else "",
        "existence": pd.to_numeric(inv.get("Existencia", pd.Series(0, index=inv.index)), errors="coerce").fillna(0.0),
        "cost": pd.to_numeric(inv.get("Costo", pd.Series(np.nan, index=inv.index)), errors="coerce").fillna(0.0),
        "cost_known": pd.to_numeric(inv.get("Costo", pd.Series(np.nan, index=inv.index)), errors="coerce").notna().astype(int),
        "supplier": inv.get("Proveedor", pd.Series("", index=inv.index)).fillna("").astype(str)
        if "Proveedor" in inv.columns else "",
        "category": inv.get("Categoría", pd.Series("", index=inv.index)).fillna("").astype(str)
        if "Categoría" in inv.columns else "",
    })

    sal_shaped = pd.DataFrame({
        "snapshot_date": snapshot_date,
        "site": site,
        "code": sal.get("Código", pd.Series("", index=sal.index)).astype(str),
        "description": sal.get("Descripción", pd.Series("", index=sal.index)).fillna("").astype(str)
        if "Descripción" in sal.columns else "",
        "sales": pd.to_numeric(sal.get("Ventas", pd.Series(0, index=sal.index)), errors="coerce").fillna(0.0),
    })

    with db() as con:
        con.execute(
            "DELETE FROM inventory_snapshots WHERE site=? AND snapshot_date=?",
            (site, snapshot_date),
        )
        con.execute(
            "DELETE FROM sales_snapshots WHERE site=? AND snapshot_date=?",
            (site, snapshot_date),
        )

        insert_sql_frame(con, "inventory_snapshots", inv_shaped)

        insert_sql_frame(con, "sales_snapshots", sal_shaped)

        con.execute(
            "INSERT INTO processing_log(processed_at,snapshot_date,site,inventory_rows,sales_rows) VALUES(?,?,?,?,?)",
            (
                datetime.now().isoformat(timespec="microseconds"),
                snapshot_date,
                site,
                len(inv_shaped),
                len(sal_shaped),
            ),
        )
        con.commit()

    st.cache_data.clear()


@st.cache_data(show_spinner=False)
def load_snapshot(site: str, snapshot_date: str) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    with db() as con:
        inv = pd.read_sql_query(
            "SELECT code AS Código, description AS Descripción, existence AS Existencia, CASE WHEN cost_known=1 THEN cost ELSE NULL END AS Costo, supplier AS Proveedor, category AS Categoría, department AS Departamento, brand AS Marca, source_value AS 'Valor inventario A2' FROM inventory_snapshots WHERE site=? AND snapshot_date=?",
            con,
            params=(site, snapshot_date),
        )
        sales = pd.read_sql_query(
            "SELECT code AS Código, description AS Descripción, sales AS Ventas FROM sales_snapshots WHERE site=? AND snapshot_date=?",
            con,
            params=(site, snapshot_date),
        )

        prev_date = previous_snapshot_date(site, snapshot_date)
        if prev_date:
            prev_sales = pd.read_sql_query(
                "SELECT code AS Código, sales AS Ventas FROM sales_snapshots WHERE site=? AND snapshot_date=?",
                con,
                params=(site, prev_date),
            )
        else:
            prev_sales = pd.DataFrame(columns=["Código", "Ventas"])
    return inv, sales, prev_sales



def supplier_db_version(supplier_name: str = "") -> str:
    key = normalize_key(supplier_name)
    with db() as con:
        if key:
            row = con.execute(
                "SELECT COALESCE(MAX(processed_at),'') FROM supplier_processing_log WHERE supplier_key=?",
                (key,),
            ).fetchone()
        else:
            row = con.execute(
                "SELECT COALESCE(MAX(processed_at),'') FROM supplier_processing_log"
            ).fetchone()
    return str(row[0] or "") + ":" + v7_version()


@st.cache_data(show_spinner=False)
def list_known_suppliers(data_version: str, supplier_version: str) -> list[str]:
    with db() as con:
        rows = con.execute(
            """
            SELECT supplier_name FROM supplier_inventory_snapshots
            WHERE TRIM(COALESCE(supplier_name,''))<>''
            UNION
            SELECT supplier FROM inventory_snapshots
            WHERE TRIM(COALESCE(supplier,''))<>''
            UNION SELECT name FROM v7_catalog WHERE kind='proveedores'
            """
        ).fetchall()
    names = sorted({str(r[0]).strip() for r in rows if r and str(r[0]).strip()}, key=lambda x: normalize_key(x))
    return names


def latest_supplier_snapshot_date(site: str, supplier_name: str) -> str | None:
    key = normalize_key(supplier_name)
    if not key:
        return None
    with db() as con:
        row = con.execute(
            "SELECT MAX(snapshot_date) FROM supplier_inventory_snapshots WHERE site=? AND supplier_key=?",
            (site, key),
        ).fetchone()
    return row[0] if row and row[0] else None


def previous_supplier_snapshot_date(site: str, supplier_name: str, current_date: str) -> str | None:
    key = normalize_key(supplier_name)
    with db() as con:
        row = con.execute(
            "SELECT MAX(snapshot_date) FROM supplier_sales_snapshots WHERE site=? AND supplier_key=? AND snapshot_date<?",
            (site, key, current_date),
        ).fetchone()
    return row[0] if row and row[0] else None


def save_supplier_snapshot(
    site: str,
    snapshot_date: str,
    supplier_name: str,
    inventory: pd.DataFrame,
    sales: pd.DataFrame,
) -> None:
    supplier_name = str(supplier_name).strip()
    supplier_key = normalize_key(supplier_name)
    if not supplier_key:
        raise ValueError("Debes indicar un proveedor o marca.")

    inv = inventory.copy()
    sal = sales.copy()
    inv["Proveedor"] = supplier_name

    inv_shaped = pd.DataFrame({
        "snapshot_date": snapshot_date,
        "site": site,
        "supplier_key": supplier_key,
        "supplier_name": supplier_name,
        "code": inv.get("Código", pd.Series("", index=inv.index)).astype(str),
        "description": inv.get("Descripción", pd.Series("", index=inv.index)).fillna("").astype(str),
        "existence": pd.to_numeric(inv.get("Existencia", pd.Series(0, index=inv.index)), errors="coerce").fillna(0.0),
        "cost": pd.to_numeric(inv.get("Costo", pd.Series(np.nan, index=inv.index)), errors="coerce").fillna(0.0),
        "cost_known": pd.to_numeric(inv.get("Costo", pd.Series(np.nan, index=inv.index)), errors="coerce").notna().astype(int),
        "category": inv.get("Categoría", pd.Series("", index=inv.index)).fillna("").astype(str),
    })
    sal_shaped = pd.DataFrame({
        "snapshot_date": snapshot_date,
        "site": site,
        "supplier_key": supplier_key,
        "supplier_name": supplier_name,
        "code": sal.get("Código", pd.Series("", index=sal.index)).astype(str),
        "description": sal.get("Descripción", pd.Series("", index=sal.index)).fillna("").astype(str),
        "sales": pd.to_numeric(sal.get("Ventas", pd.Series(0, index=sal.index)), errors="coerce").fillna(0.0),
    })

    with db() as con:
        con.execute(
            "DELETE FROM supplier_inventory_snapshots WHERE site=? AND snapshot_date=? AND supplier_key=?",
            (site, snapshot_date, supplier_key),
        )
        con.execute(
            "DELETE FROM supplier_sales_snapshots WHERE site=? AND snapshot_date=? AND supplier_key=?",
            (site, snapshot_date, supplier_key),
        )
        insert_sql_frame(con, "supplier_inventory_snapshots", inv_shaped)
        insert_sql_frame(con, "supplier_sales_snapshots", sal_shaped)
        con.execute(
            """
            INSERT INTO supplier_processing_log(
                processed_at,snapshot_date,site,supplier_key,supplier_name,inventory_rows,sales_rows
            ) VALUES(?,?,?,?,?,?,?)
            """,
            (
                datetime.now().isoformat(timespec="microseconds"), snapshot_date, site, supplier_key,
                supplier_name, len(inv_shaped), len(sal_shaped),
            ),
        )
        con.commit()
    st.cache_data.clear()


@st.cache_data(show_spinner=False)
def load_supplier_snapshot(site: str, supplier_name: str, snapshot_date: str) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    key = normalize_key(supplier_name)
    with db() as con:
        inv = pd.read_sql_query(
            """
            SELECT code AS Código, description AS Descripción, existence AS Existencia,
                   CASE WHEN cost_known=1 THEN cost ELSE NULL END AS Costo, supplier_name AS Proveedor, category AS Categoría,
                   department AS Departamento, brand AS Marca, source_value AS "Valor inventario A2"
            FROM supplier_inventory_snapshots
            WHERE site=? AND supplier_key=? AND snapshot_date=?
            """,
            con, params=(site, key, snapshot_date),
        )
        sales = pd.read_sql_query(
            """
            SELECT code AS Código, description AS Descripción, sales AS Ventas
            FROM supplier_sales_snapshots
            WHERE site=? AND supplier_key=? AND snapshot_date=?
            """,
            con, params=(site, key, snapshot_date),
        )
        prev_date = previous_supplier_snapshot_date(site, supplier_name, snapshot_date)
        if prev_date:
            prev_sales = pd.read_sql_query(
                """
                SELECT code AS Código, sales AS Ventas
                FROM supplier_sales_snapshots
                WHERE site=? AND supplier_key=? AND snapshot_date=?
                """,
                con, params=(site, key, prev_date),
            )
        else:
            prev_sales = pd.DataFrame(columns=["Código", "Ventas"])
    return inv, sales, prev_sales


@st.cache_data(show_spinner=False)
def build_rotation_history(data_version: str, sales_mode: str) -> pd.DataFrame:
    """Última fecha con movimiento positivo por sede/SKU usando el historial guardado."""
    with db() as con:
        hist = pd.read_sql_query(
            "SELECT snapshot_date, site AS Sede, code AS Código, sales AS Ventas FROM sales_snapshots ORDER BY site, code, snapshot_date",
            con,
        )
    if hist.empty:
        return pd.DataFrame(columns=["Sede", "Código", "Primera fecha", "Última fecha", "Última venta positiva"])
    hist["snapshot_date"] = pd.to_datetime(hist["snapshot_date"], errors="coerce")
    hist["Ventas"] = pd.to_numeric(hist["Ventas"], errors="coerce").fillna(0.0)
    hist = hist.dropna(subset=["snapshot_date"])
    if sales_mode == "acumuladas":
        hist["Movimiento"] = hist.groupby(["Sede", "Código"], observed=True)["Ventas"].diff().fillna(0).clip(lower=0)
    else:
        hist["Movimiento"] = hist["Ventas"].clip(lower=0)
    base = hist.groupby(["Sede", "Código"], observed=True).agg(
        **{"Primera fecha": ("snapshot_date", "min"), "Última fecha": ("snapshot_date", "max")}
    ).reset_index()
    positive = hist[hist["Movimiento"] > 0].groupby(["Sede", "Código"], observed=True)["snapshot_date"].max().rename("Última venta positiva").reset_index()
    return base.merge(positive, on=["Sede", "Código"], how="left")


# ============================================================
# EXCEL INPUT — NORMALIZACIÓN ROBUSTA
# ============================================================

def find_header_row(raw: pd.DataFrame) -> int:
    keywords = [
        "código", "codigo", "cod", "artículo", "articulo", "sku",
        "descripción", "descripcion", "existencia", "stock", "saldo",
        "cantidad", "ventas", "costo", "proveedor", "categoria", "categoría",
    ]
    best_row, best_score = 0, 0
    for idx, row in raw.iterrows():
        vals = [normalize_text(v) for v in row.tolist() if pd.notna(v)]
        score = sum(1 for kw in keywords if any(kw == v or kw in v for v in vals))
        if score > best_score:
            best_row, best_score = idx, score
    return best_row


def standardize_columns(df: pd.DataFrame, kind: str) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    aliases = {
        "Código": ["código", "codigo", "cod", "cod_articulo", "cod. artículo", "artículo", "articulo", "sku", "codigo producto"],
        "Descripción": ["descripción", "descripcion", "desc", "producto", "nombre", "detalle"],
        "Ventas": ["cantidad", "ventas", "salidas", "venta", "unidades vendidas", "cantidad vendida", "movimiento"],
        "Existencia": ["existencia", "stock", "saldo", "inventario", "disponible", "existencias"],
        "Costo": ["costo", "precio", "costo unitario", "coste", "precio costo", "precio de costo"],
        "Proveedor": ["proveedor", "marca/proveedor", "vendor", "supplier"],
        "Categoría": ["categoría", "categoria", "familia", "rubro", "grupo", "linea", "línea"],
    }
    rename = {}
    for col in df.columns:
        c = normalize_text(col)
        for target, variants in aliases.items():
            if c in variants or any(c.startswith(v + " ") for v in variants):
                if target not in rename.values():
                    rename[col] = target
                break
    df = df.rename(columns=rename)
    required = ["Código", "Existencia"] if kind == "inventario" else ["Código", "Ventas"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError("Faltan columnas de " + kind + ": " + ", ".join(missing))

    for col in ["Ventas", "Existencia", "Costo"]:
        if col in df.columns:
            df[col] = df[col].map(parse_number)

    if "Código" in df.columns:
        df["Código"] = (
            df["Código"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
        )
        df = df[(df["Código"] != "") & (df["Código"].str.lower() != "nan")]

    if "Descripción" not in df.columns:
        df["Descripción"] = ""
    df["Descripción"] = df["Descripción"].fillna("").astype(str).str.strip()

    if kind == "inventario":
        if "Existencia" not in df.columns:
            df["Existencia"] = 0.0
        if "Costo" not in df.columns:
            df["Costo"] = 0.0
    else:
        if "Ventas" not in df.columns:
            df["Ventas"] = 0.0

    if "Código" in df.columns:
        agg: dict[str, str] = {}
        for col in df.columns:
            if col == "Código":
                continue
            if col in ["Ventas", "Existencia"]:
                agg[col] = "sum"
            else:
                agg[col] = "first"
        df = df.groupby("Código", as_index=False).agg(agg)

    return df.reset_index(drop=True)


def _read_excel_fast(bio: io.BytesIO, **kwargs) -> pd.DataFrame:
    if EXCEL_ENGINE:
        try:
            bio.seek(0)
            return pd.read_excel(bio, engine=EXCEL_ENGINE, **kwargs)
        except Exception:
            pass
    bio.seek(0)
    return pd.read_excel(bio, **kwargs)


@st.cache_data(show_spinner=False)
def parse_uploaded_excel(raw_bytes: bytes, kind: str) -> tuple[pd.DataFrame, int]:
    bio = io.BytesIO(raw_bytes)
    # El encabezado suele estar al inicio: leer solo una muestra evita parsear
    # el Excel completo dos veces y acelera notablemente archivos grandes.
    raw = _read_excel_fast(bio, header=None, nrows=80)
    header_row = find_header_row(raw)
    df = _read_excel_fast(bio, skiprows=header_row)
    df = standardize_columns(df, kind)
    df = downcast_dataframe(df)
    return df, header_row


def downcast_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    df = df.copy()
    category_like = ["Estado", "Prioridad", "ABC", "Acción", "Sede", "Categoría",
                      "Proveedor", "Tiene costo", "Stock negativo", "Sin movimiento"]
    for col in category_like:
        if col in df.columns and df[col].dtype == object:
            if df[col].nunique(dropna=False) < max(50, len(df) // 2):
                df[col] = df[col].astype("category")

    integer_like = [
        "Ventas", "Demanda Mensual", "Existencia", "Stock Mínimo", "Stock Máximo",
        "Demanda Lead Time", "Stock Seguridad", "Punto de Reorden",
        "Compra Sugerida", "Retiro Almacén", "Compra Ajustada",
        "Transferencias Recibidas", "Transferencias Enviadas",
    ]
    for col in integer_like:
        if col in df.columns:
            s = pd.to_numeric(df[col], errors="coerce")
            if s.notna().all() and np.isfinite(s).all() and np.allclose(s, s.round(), rtol=0, atol=1e-9):
                df[col] = pd.to_numeric(s, downcast="integer")
    return df


# ============================================================
# MOTOR LOGÍSTICO
# ============================================================

@st.cache_data(show_spinner=False)
def calculate_site(
    inv: pd.DataFrame,
    sales: pd.DataFrame,
    prev_sales: pd.DataFrame,
    months_history: int,
    min_coverage: float,
    max_coverage: float,
    lead_time_days: int,
    safety_days: int,
    abc_basis: str,
    sales_mode: str,
    rolling_days: int,
    period_days: int | None = None,
    average_months: float | None = None,
) -> pd.DataFrame:
    inv = inv.copy()
    sales = sales.copy()
    prev_sales = prev_sales.copy()

    if "Existencia" not in inv.columns:
        inv["Existencia"] = 0.0
    if "Costo" not in inv.columns:
        inv["Costo"] = 0.0
    if "Descripción" not in inv.columns:
        inv["Descripción"] = ""
    if "Proveedor" not in inv.columns:
        inv["Proveedor"] = ""
    if "Categoría" not in inv.columns:
        inv["Categoría"] = ""
    if "Ventas" not in sales.columns:
        sales["Ventas"] = 0.0

    current = sales[present_columns(sales, ["Código", "Descripción", "Ventas"])].copy()
    if sales_mode == "acumuladas":
        prev = prev_sales.rename(columns={"Ventas": "Ventas Anteriores"})
        current = current.merge(prev, on="Código", how="left")
        current["Ventas del Corte"] = np.maximum(
            current["Ventas"].fillna(0) - current["Ventas Anteriores"].fillna(0), 0
        )
        current = current[["Código", "Ventas del Corte"]]
        current = current.rename(columns={"Ventas del Corte": "Ventas"})

    df = inv.merge(current, on="Código", how="left" if period_days is not None else "outer", suffixes=("_inv", "_ventas"))

    for col in ["Existencia", "Costo"]:
        if col not in df.columns:
            df[col] = 0.0
    if "Ventas" not in df.columns:
        df["Ventas"] = 0.0
    if "Descripción_inv" in df.columns or "Descripción_ventas" in df.columns:
        d1 = df.get("Descripción_inv", pd.Series("", index=df.index)).fillna("")
        d2 = df.get("Descripción_ventas", pd.Series("", index=df.index)).fillna("")
        df["Descripción"] = d1.where(d1.astype(str).str.strip() != "", d2)
    elif "Descripción" not in df.columns:
        df["Descripción"] = ""

    df["Ventas"] = pd.to_numeric(df["Ventas"], errors="coerce").fillna(0)
    df["Existencia"] = pd.to_numeric(df["Existencia"], errors="coerce").fillna(0)
    df["Costo"] = pd.to_numeric(df["Costo"], errors="coerce").fillna(0)

    months_history = max(1, int(months_history))
    min_coverage = max(0.1, float(min_coverage))
    max_coverage = max(min_coverage, float(max_coverage))
    lead_time_days = max(0, int(lead_time_days))
    safety_days = max(0, int(safety_days))
    rolling_days = max(1, int(rolling_days))

    if average_months is not None:
        if not np.isfinite(average_months) or average_months <= 0:
            raise ValueError("Los meses confirmados deben ser mayores que cero.")
        divisor, method = float(average_months), "Meses confirmados"
    elif period_days is not None and period_days > 0:
        divisor, method = float(period_days) / 30, "Días reales / 30"
    elif sales_mode == "diarias":
        divisor, method = 1 / 30, "Un día de ventas"
    else:
        divisor, method = float(months_history), "Meses del historial anterior"
    monthly_demand = df["Ventas"].clip(lower=0) / divisor
    df["Ventas utilizadas"] = df["Ventas"]
    df["Meses utilizados"] = divisor
    df["Método de promedio"] = method
    df["Cobertura mínima aplicada"] = min_coverage
    df["Cobertura máxima aplicada"] = max_coverage
    df["Plazo proveedor (días)"] = lead_time_days
    df["Seguridad (días)"] = safety_days
    df["Demanda Mensual"] = monthly_demand
    df["Demanda Diaria"] = monthly_demand / 30
    df["Stock Mínimo"] = np.ceil(monthly_demand * min_coverage).astype(int)
    df["Stock Máximo"] = np.ceil(monthly_demand * max_coverage).astype(int)
    df["Demanda Lead Time"] = np.ceil(df["Demanda Diaria"] * lead_time_days).astype(int)
    df["Stock Seguridad"] = np.ceil(df["Demanda Diaria"] * safety_days).astype(int)
    df["Reorden por plazo"] = np.ceil(df["Demanda Diaria"] * (lead_time_days + safety_days)).astype(int)
    df["Punto de Reorden"] = np.minimum(df["Stock Máximo"], np.maximum(df["Stock Mínimo"], df["Reorden por plazo"]))
    df["Advertencia de plazo"] = np.where(df["Reorden por plazo"] > df["Stock Máximo"],
        "El plazo y la seguridad requieren más unidades que el máximo. Revisa la cobertura máxima o el plazo; el máximo no se aumenta automáticamente.", "")
    df["Mínimo automático"] = df["Stock Mínimo"]
    df["Máximo automático"] = df["Stock Máximo"]
    manual_min = pd.to_numeric(df.get("Mínimo manual", pd.Series(np.nan, index=df.index)), errors="coerce")
    manual_max = pd.to_numeric(df.get("Máximo manual", pd.Series(np.nan, index=df.index)), errors="coerce")
    manual = manual_min.notna() & manual_max.notna() & manual_min.ge(0) & manual_max.ge(manual_min)
    df["Stock Mínimo"] = df["Stock Mínimo"].astype(float).where(~manual, manual_min)
    df["Stock Máximo"] = df["Stock Máximo"].astype(float).where(~manual, manual_max)
    df["Punto de Reorden"] = df["Punto de Reorden"].astype(float).where(~manual, manual_min)
    df["Política de stock"] = np.where(manual, "Manual", "Automática")

    df["Compra Sugerida"] = np.where(
        df["Existencia"] < df["Punto de Reorden"],
        np.ceil(np.maximum(df["Stock Máximo"] - df["Existencia"], 0)),
        0,
    ).astype(int)

    df["Retiro Almacén"] = np.where(
        df["Existencia"] > df["Stock Máximo"],
        np.maximum(df["Existencia"] - df["Stock Máximo"], 0),
        0,
    ).astype(float)

    positive_stock = np.maximum(df["Existencia"], 0)
    df["Cobertura (meses)"] = np.where(
        df["Demanda Mensual"] > 0,
        (positive_stock / df["Demanda Mensual"]).round(1),
        np.nan,
    )
    df["Rotación estimada (x/mes)"] = np.where(
        positive_stock > 0,
        (df["Demanda Mensual"] / positive_stock).round(2),
        0,
    )

    source_value = pd.to_numeric(df.get("Valor inventario A2", pd.Series(np.nan, index=df.index)), errors="coerce")
    df["Valor Inventario ($)"] = source_value.fillna(df["Existencia"] * df["Costo"]).round(2)
    df["Capital Inmovilizado ($)"] = (df["Retiro Almacén"] * df["Costo"]).round(2)
    df["Costo Compra Estimada ($)"] = (df["Compra Sugerida"] * df["Costo"]).round(2)
    df["Valor Movimiento ($)"] = (df["Demanda Mensual"] * df["Costo"]).round(2)

    if abc_basis == "Inventario":
        abc_value = np.maximum(df["Valor Inventario ($)"], 0)
    elif abc_basis == "Capital inmovilizado":
        abc_value = np.maximum(df["Capital Inmovilizado ($)"], 0)
    else:
        abc_value = np.maximum(df["Valor Movimiento ($)"], 0)

    total_abc = float(abc_value.sum())
    if total_abc > 0:
        order = np.argsort(-abc_value.to_numpy())
        cumulative = np.zeros(len(df))
        ordered = abc_value.to_numpy()[order]
        cumulative[order] = (np.cumsum(ordered) - ordered) / total_abc
        # Incluye en A/B el artículo que cruza el umbral (un único SKU dominante sigue siendo A).
        df["ABC"] = np.select(
            [cumulative <= 0.80, cumulative <= 0.95],
            ["A", "B"],
            default="C",
        )
    else:
        df["ABC"] = "C"

    df["Tiene costo"] = np.where(df["Costo"] > 0, "Sí", "No")
    df["Stock negativo"] = np.where(df["Existencia"] < 0, "Sí", "No")
    df["Sin movimiento"] = np.where((df["Ventas"] <= 0) & (df["Existencia"] > 0), "Sí", "No")

    df["Estado"] = np.select(
        [
            df["Existencia"] < 0,
            df["Existencia"] < df["Punto de Reorden"],
            df["Retiro Almacén"] > 0,
        ],
        [
            "CRÍTICO — SALDO NEGATIVO",
            "CRÍTICO — COMPRAR",
            "SOBRESTOCK — RETIRAR",
        ],
        default="ÓPTIMO",
    )

    df["Prioridad"] = np.select(
        [
            df["Existencia"] < 0,
            (df["Compra Sugerida"] > 0) & (df["ABC"] == "A"),
            df["Compra Sugerida"] > 0,
            (df["Retiro Almacén"] > 0) & (df["ABC"] == "A"),
            df["Retiro Almacén"] > 0,
        ],
        ["CRÍTICA", "ALTA", "MEDIA", "ALTA", "MEDIA"],
        default="BAJA",
    )

    df["Acción"] = np.select(
        [
            df["Existencia"] < 0,
            df["Compra Sugerida"] > 0,
            df["Retiro Almacén"] > 0,
        ],
        ["REVISAR SALDO / ABASTECER", "COMPRAR", "RETIRAR / REDISTRIBUIR"],
        default="NO HACER NADA",
    )

    keep = [
        "Código", "Descripción", "Proveedor", "Marca", "Departamento", "Categoría", "Ventas", "Demanda Mensual",
        "Mínimo automático", "Máximo automático", "Política de stock", "Valor inventario A2",
        "Demanda Diaria", "Existencia", "Costo", "Stock Mínimo", "Stock Máximo",
        "Demanda Lead Time", "Stock Seguridad", "Punto de Reorden", "Compra Sugerida",
        "Retiro Almacén", "Cobertura (meses)", "Rotación estimada (x/mes)", "Valor Inventario ($)",
        "Capital Inmovilizado ($)", "Costo Compra Estimada ($)", "Valor Movimiento ($)",
        "ABC", "Tiene costo", "Stock negativo", "Sin movimiento", "Estado", "Prioridad", "Acción",
    ]
    keep += [c for c in STOCK_AUDIT_COLUMNS_V73 if c in df and c not in keep]
    for col in keep:
        if col not in df.columns:
            df[col] = ""
    result = df[keep].copy()
    integer_cols = [
        "Ventas", "Demanda Mensual", "Existencia", "Stock Mínimo", "Stock Máximo",
        "Demanda Lead Time", "Stock Seguridad", "Punto de Reorden", "Compra Sugerida",
        "Retiro Almacén", "Transferencias Recibidas", "Transferencias Enviadas", "Compra Ajustada"
    ]
    for col in integer_cols:
        if col in result.columns:
            values = pd.to_numeric(result[col], errors="coerce").fillna(0)
            result[col] = values if col in {"Ventas", "Existencia", "Demanda Mensual", "Retiro Almacén", "Stock Mínimo", "Stock Máximo", "Punto de Reorden"} else values.round().astype(int)
    money_cols = [
        "Costo", "Valor Inventario ($)", "Capital Inmovilizado ($)",
        "Costo Compra Estimada ($)", "Valor Movimiento ($)"
    ]
    for col in money_cols:
        if col in result.columns:
            result[col] = pd.to_numeric(result[col], errors="coerce").fillna(0).round(2)
    if "Demanda Diaria" in result.columns:
        result["Demanda Diaria"] = pd.to_numeric(result["Demanda Diaria"], errors="coerce").fillna(0).round(2)
    if "Cobertura (meses)" in result.columns:
        result["Cobertura (meses)"] = pd.to_numeric(result["Cobertura (meses)"], errors="coerce").round(1)
    if "Rotación estimada (x/mes)" in result.columns:
        result["Rotación estimada (x/mes)"] = pd.to_numeric(result["Rotación estimada (x/mes)"], errors="coerce").fillna(0).round(2)
    result["_orden"] = result["Prioridad"].map({"CRÍTICA": 0, "ALTA": 1, "MEDIA": 2, "BAJA": 3})
    result = result.sort_values(
        by=["_orden", "Capital Inmovilizado ($)", "Demanda Mensual"],
        ascending=[True, False, False],
    ).drop(columns="_orden").reset_index(drop=True)
    return downcast_dataframe(result)


# ============================================================
# REDISTRIBUCIÓN MULTISEDE
# ============================================================

def build_redistribution(site_results: dict[str, pd.DataFrame]) -> pd.DataFrame:
    frames = []
    for site, df in site_results.items():
        if df is None or df.empty:
            continue
        x = df[[
            "Código", "Descripción", "Existencia", "Stock Máximo",
            "Stock Mínimo", "Costo", "Demanda Mensual"
        ]].copy()
        x.insert(0, "Sede", site)
        frames.append(x)
    if not frames:
        return pd.DataFrame()

    all_df = pd.concat(frames, ignore_index=True)
    suggestions: list[dict[str, Any]] = []

    for code, group in all_df.groupby("Código"):
        donors = group.copy()
        receivers = group.copy()
        donors["Exceso"] = np.maximum(donors["Existencia"] - donors["Stock Máximo"], 0)
        receivers["Déficit"] = np.maximum(receivers["Stock Mínimo"] - receivers["Existencia"], 0)
        donors = donors[donors["Exceso"] > 0].sort_values("Exceso", ascending=False)
        receivers = receivers[receivers["Déficit"] > 0].sort_values("Déficit", ascending=False)

        for r in receivers.itertuples(index=False):
            need = float(r.Déficit)
            if need <= 0:
                continue
            for d in donors.itertuples(index=False):
                if d.Sede == r.Sede:
                    continue
                available = float(d.Exceso)
                if available <= 0:
                    continue
                qty = int(np.floor(min(need, available)))
                if qty <= 0:
                    continue
                # Replacement cost belongs to the receiving site; another site is not a price quote.
                cost = float(r.Costo) if pd.notna(r.Costo) and np.isfinite(r.Costo) else np.nan
                suggestions.append(
                    {
                        "Código": code,
                        "Descripción": r.Descripción or d.Descripción,
                        "Origen": d.Sede,
                        "Destino": r.Sede,
                        "Existencia Origen": int(round(d.Existencia)),
                        "Necesidad Destino": int(round(r.Déficit)),
                        "Unidades Sugeridas": qty,
                        "Costo Unitario ($)": cost,
                        "Compra Evitada Estimada ($)": round(qty * cost, 2),
                        "Estado": "RECOMENDADA",
                    }
                )
                need -= qty
                idx = donors.index[donors["Sede"] == d.Sede][0]
                donors.loc[idx, "Exceso"] -= qty
                if need <= 0:
                    break

    if not suggestions:
        return pd.DataFrame(
            columns=[
                "Código", "Descripción", "Origen", "Destino", "Existencia Origen",
                "Necesidad Destino", "Unidades Sugeridas", "Costo Unitario ($)",
                "Compra Evitada Estimada ($)", "Estado"
            ]
        )

    return pd.DataFrame(suggestions).sort_values(
        "Compra Evitada Estimada ($)", ascending=False
    ).reset_index(drop=True)


def apply_redistribution_to_results(
    site_results: dict[str, pd.DataFrame],
    transfers: pd.DataFrame,
) -> dict[str, pd.DataFrame]:
    out = {site: df.copy() for site, df in site_results.items()}
    if transfers.empty:
        for site, df in out.items():
            df["Transferencias Recibidas"] = 0
            df["Transferencias Enviadas"] = 0
            df["Compra Ajustada"] = df["Compra Sugerida"]
        return out

    received = (
        transfers.groupby(["Destino", "Código"])["Unidades Sugeridas"]
        .sum().rename("Transferencias Recibidas").reset_index()
        .rename(columns={"Destino": "Sede"})
    )
    sent = (
        transfers.groupby(["Origen", "Código"])["Unidades Sugeridas"]
        .sum().rename("Transferencias Enviadas").reset_index()
        .rename(columns={"Origen": "Sede"})
    )

    for site, df in out.items():
        df.drop(columns=["Transferencias Recibidas", "Transferencias Enviadas"], errors="ignore", inplace=True)
        df = df.merge(received[received["Sede"] == site][["Código", "Transferencias Recibidas"]], on="Código", how="left")
        df = df.merge(sent[sent["Sede"] == site][["Código", "Transferencias Enviadas"]], on="Código", how="left")
        df["Transferencias Recibidas"] = df["Transferencias Recibidas"].fillna(0).astype(int)
        df["Transferencias Enviadas"] = df["Transferencias Enviadas"].fillna(0).astype(int)
        df["Compra Ajustada"] = (df["Compra Sugerida"] - df["Transferencias Recibidas"]).clip(lower=0)
        df["Costo Compra Estimada ($)"] = (df["Compra Ajustada"] * df["Costo"]).where(df["Compra Ajustada"].ne(0), 0.0).round(2)

        recibe_y_compraba = (df["Transferencias Recibidas"] > 0) & (df["Compra Sugerida"] > 0)
        envia_y_retiraba = (df["Transferencias Enviadas"] > 0) & (df["Retiro Almacén"] > 0)
        transferir_y_comprar = recibe_y_compraba & df["Compra Ajustada"].gt(0)
        df["Acción"] = np.select(
            [transferir_y_comprar, recibe_y_compraba, envia_y_retiraba],
            ["TRANSFERIR + COMPRAR", "TRANSFERIR DESDE OTRA SEDE", "TRANSFERIR A OTRA SEDE"],
            default=df["Acción"],
        )
        out[site] = df
    return out


# ============================================================
# MÉTRICAS GLOBALES
# ============================================================

def combine_sites(site_results: dict[str, pd.DataFrame]) -> pd.DataFrame:
    frames = []
    for site, df in site_results.items():
        if df is None or df.empty:
            continue
        x = df.copy()
        x.insert(0, "Sede", site)
        frames.append(x)
    if not frames:
        return pd.DataFrame()
    return downcast_dataframe(pd.concat(frames, ignore_index=True))


def health_score(df: pd.DataFrame) -> int:
    if df.empty:
        return 0
    n = len(df)
    critical = float((df["Estado"].str.startswith("CRÍTICO")).sum()) / n
    overstock = float((df["Estado"] == "SOBRESTOCK — RETIRAR").sum()) / n
    dead = float((df["Sin movimiento"] == "Sí").sum()) / n
    no_cost = float((df["Tiene costo"] == "No").sum()) / n
    negative = float((df["Stock negativo"] == "Sí").sum()) / n
    score = 100 - (critical * 45 + overstock * 25 + dead * 15 + no_cost * 10 + negative * 5)
    return int(np.clip(round(score), 0, 100))


def health_label(score: int) -> str:
    if score >= 85:
        return "SALUDABLE"
    if score >= 70:
        return "VIGILAR"
    if score >= 50:
        return "RIESGO"
    return "CRÍTICA"



def _quantity_text_v73(value) -> str:
    """Display physical quantities without forcing fractions to whole units."""
    try:
        number = float(value)
    except (TypeError, ValueError):
        return "N/D"
    if not np.isfinite(number):
        return "N/D"
    if number.is_integer():
        return f"{int(number):,}"
    return f"{number:,.12g}"


def build_alerts(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Tipo", "Severidad", "Sede", "Código", "Descripción", "Detalle"])

    base_cols = ["Sede", "Código", "Descripción"]
    pieces = []

    m = df["Estado"] == "CRÍTICO — SALDO NEGATIVO"
    if m.any():
        sub = df.loc[m, base_cols].copy()
        sub["Tipo"] = "Saldo negativo"
        sub["Severidad"] = "CRÍTICA"
        sub["Detalle"] = "Existencia " + df.loc[m, "Existencia"].map(_quantity_text_v73) + " · revisar origen y abastecer."
        pieces.append(sub)

    m = df["Estado"] == "CRÍTICO — COMPRAR"
    if m.any():
        sub = df.loc[m, base_cols].copy()
        sub["Tipo"] = "Stock crítico"
        sub["Severidad"] = "ALTA"
        sub["Detalle"] = (
            "Existencia " + df.loc[m, "Existencia"].map(_quantity_text_v73)
            + " · punto de reorden " + df.loc[m, "Punto de Reorden"].map(_quantity_text_v73) + "."
        )
        pieces.append(sub)

    m = df["Estado"] == "SOBRESTOCK — RETIRAR"
    if m.any():
        sub = df.loc[m, base_cols].copy()
        sub["Tipo"] = "Sobrestock"
        sub["Severidad"] = "MEDIA"
        sub["Detalle"] = "Retirar " + df.loc[m, "Retiro Almacén"].map(_quantity_text_v73) + " unidades."
        pieces.append(sub)

    m = (df["Sin movimiento"] == "Sí") & (df["Existencia"] > 0)
    if m.any():
        sub = df.loc[m, base_cols].copy()
        sub["Tipo"] = "Sin movimiento"
        sub["Severidad"] = "MEDIA"
        sub["Detalle"] = "Hay " + df.loc[m, "Existencia"].map(_quantity_text_v73) + " unidades sin venta en el corte."
        pieces.append(sub)

    m = df["Tiene costo"] == "No"
    if m.any():
        sub = df.loc[m, base_cols].copy()
        sub["Tipo"] = "Costo faltante"
        sub["Severidad"] = "BAJA"
        sub["Detalle"] = "Costo unitario no disponible; el análisis financiero puede estar subestimado."
        pieces.append(sub)

    if not pieces:
        return pd.DataFrame(columns=["Tipo", "Severidad", "Sede", "Código", "Descripción", "Detalle"])

    result = pd.concat(pieces, ignore_index=True)
    return result[["Tipo", "Severidad", "Sede", "Código", "Descripción", "Detalle"]]


# ============================================================
# CONTACTOS / AUTOMATIZACIÓN
# ============================================================

def load_contacts() -> dict[str, dict[str, str]]:
    if not CONTACTS_PATH.exists():
        return DEFAULT_CONTACTS.copy()
    try:
        raw = json.loads(CONTACTS_PATH.read_text(encoding="utf-8"))
        result = DEFAULT_CONTACTS.copy()
        for site in SEDES:
            result[site] = {
                "supervisor": str(raw.get(site, {}).get("supervisor", "")),
                "whatsapp": str(raw.get(site, {}).get("whatsapp", "")),
                "webhook": str(raw.get(site, {}).get("webhook", "")),
            }
        return result
    except Exception:
        return DEFAULT_CONTACTS.copy()


def save_contacts(contacts: dict[str, dict[str, str]]) -> None:
    CONTACTS_PATH.write_text(
        json.dumps(contacts, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def compose_digest(site: str, df: pd.DataFrame, transfers: pd.DataFrame) -> str:
    if df.empty:
        return f"MAKROPETROL NEXUS · {site}\nSin datos procesados."
    score = health_score(df)
    critical = int(df["Estado"].str.startswith("CRÍTICO").sum())
    purchases = int((df["Compra Ajustada"] > 0).sum()) if "Compra Ajustada" in df.columns else int((df["Compra Sugerida"] > 0).sum())
    withdrawals = int((df["Retiro Almacén"] > 0).sum())
    transfers_out = int(df.get("Transferencias Enviadas", pd.Series(dtype=float)).sum()) if "Transferencias Enviadas" in df.columns else 0
    transfers_in = int(df.get("Transferencias Recibidas", pd.Series(dtype=float)).sum()) if "Transferencias Recibidas" in df.columns else 0
    capital = float(monetary_sum_v71(df["Capital Inmovilizado ($)"]))
    parts = [
        f"*MAKROPETROL NEXUS* · {site}",
        f"Salud logística: {score}/100 · {health_label(score)}",
        f"🔴 Críticos: {critical}",
        f"🛒 Compras pendientes: {purchases}",
        f"📤 Retiros: {withdrawals}",
        f"♻️ Transferencias recibidas/enviadas: {transfers_in}/{transfers_out}",
        f"💰 Capital inmovilizado: {money(capital)}",
    ]
    if not transfers.empty:
        own = transfers[(transfers["Origen"] == site) | (transfers["Destino"] == site)].head(3)
        if not own.empty:
            parts.append("\n*Movimientos prioritarios:*")
            for r in own.to_dict("records"):
                parts.append(f"• {r['Código']} · {r['Origen']} → {r['Destino']} · {int(r['Unidades Sugeridas'])} uds")
    return "\n".join(parts)


def post_webhook(url: str, payload: dict[str, Any]) -> tuple[bool, str]:
    if not url.strip():
        return False, "Webhook vacío"
    try:
        r = requests.post(url.strip(), json=payload, timeout=8)
        if 200 <= r.status_code < 300:
            return True, f"HTTP {r.status_code}"
        return False, f"HTTP {r.status_code}: {r.text[:180]}"
    except Exception as exc:
        return False, str(exc)


def post_webhook_async(url: str, payload: dict[str, Any]):
    if not url.strip():
        return None
    return WEBHOOK_EXECUTOR.submit(post_webhook, url, payload)


# ============================================================
# DOCUMENTOS OPERATIVOS
# ============================================================

def _doc_safe_text(value: Any, fallback: str = "") -> str:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return fallback
    return str(value).strip() or fallback


def _doc_order_number(doc_type: str, site: str, snapshot_label: str, index: int = 1) -> str:
    prefix = DOCUMENT_ORDER_PREFIX.get(doc_type, "DOC")
    compact_date = re.sub(r"[^0-9]", "", str(snapshot_label))
    site_code = re.sub(r"[^A-Za-z0-9]", "", site)[:4].upper() or "SEDE"
    return f"{prefix}-{compact_date}-{site_code}-{index:04d}"


def _doc_prepare_frame(
    df: pd.DataFrame,
    document_type: str,
    site: str | None = None,
) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(
            columns=["Código", "Descripción", "Cantidad", "Precio Unitario", "Total"]
        )

    x = df.copy()
    if site and "Sede" in x.columns:
        x = x[x["Sede"].astype(str) == str(site)].copy()

    if document_type == "purchase":
        qty_col = "Compra Ajustada" if "Compra Ajustada" in x.columns else "Compra Sugerida"
        x = x[pd.to_numeric(x.get(qty_col, 0), errors="coerce").fillna(0) > 0].copy()
        if x.empty:
            return pd.DataFrame(columns=["Código", "Descripción", "Cantidad", "Precio Unitario", "Total"])
        result = pd.DataFrame({
            "Código": x["Código"].astype(str),
            "Descripción": x["Descripción"].astype(str),
            "Cantidad": pd.to_numeric(x[qty_col], errors="coerce").fillna(0).round().astype(int),
            "Precio Unitario": pd.to_numeric(x["Costo"], errors="coerce").round(2),
        })

    elif document_type == "withdrawal":
        x = x[pd.to_numeric(x.get("Retiro Almacén", 0), errors="coerce").fillna(0) > 0].copy()
        if x.empty:
            return pd.DataFrame(columns=["Código", "Descripción", "Cantidad", "Precio Unitario", "Total"])
        result = pd.DataFrame({
            "Código": x["Código"].astype(str),
            "Descripción": x["Descripción"].astype(str),
            "Cantidad": pd.to_numeric(x["Retiro Almacén"], errors="coerce").fillna(0),
            "Precio Unitario": pd.to_numeric(x["Costo"], errors="coerce").round(2),
        })

    elif document_type == "alerts":
        if "Tipo" not in x.columns:
            return pd.DataFrame(columns=["Código", "Descripción", "Alerta", "Severidad", "Detalle"])
        result = x[["Código", "Descripción", "Tipo", "Severidad", "Detalle"]].copy()
        result.columns = ["Código", "Descripción", "Alerta", "Severidad", "Detalle"]
        return result.reset_index(drop=True)

    elif document_type == "redistribution":
        result = x[[
            "Código", "Descripción", "Origen", "Destino",
            "Unidades Sugeridas", "Costo Unitario ($)",
            "Compra Evitada Estimada ($)", "Estado"
        ]].copy()
        return result.reset_index(drop=True)

    else:
        return pd.DataFrame()

    result["Total"] = (result["Cantidad"] * result["Precio Unitario"]).round(2)
    return result.reset_index(drop=True)


def _document_purchase_meta(
    site: str,
    snapshot_label: str,
    supplier: str = "",
    supplier_address: str = "",
    supplier_phone: str = "",
    sequence: int = 1,
) -> dict[str, str]:
    now = datetime.now()
    return {
        "company": COMPANY_INFO["name"],
        "company_address": COMPANY_INFO["address"],
        "company_phone": COMPANY_INFO["phone"],
        "company_rif": COMPANY_INFO["rif"],
        "document_title": "ORDEN DE COMPRA",
        "date": str(snapshot_label),
        "time": now.strftime("%H:%M:%S"),
        "site": site,
        "order_no": _doc_order_number("purchase", site, snapshot_label, sequence),
        "supplier": supplier or SUPPLIER_DEFAULTS["name"],
        "supplier_address": supplier_address or SUPPLIER_DEFAULTS["address"],
        "supplier_phone": supplier_phone or SUPPLIER_DEFAULTS["phone"],
        "operator": "NEXUS / ADMINISTRACION",
    }


def build_purchase_pdf(
    data: pd.DataFrame,
    site: str,
    snapshot_label: str,
    sequence: int = 1,
    supplier: str = "",
    supplier_address: str = "",
    supplier_phone: str = "",
) -> bytes | None:
    if not HAS_PDF:
        return None

    x = _doc_prepare_frame(data, "purchase", site)
    meta = _document_purchase_meta(
        site, snapshot_label, supplier, supplier_address, supplier_phone, sequence
    )
    total = float(monetary_sum_v71(x["Total"])) if not x.empty else 0.0

    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=14)
    pdf.add_page()
    pdf.set_title(f"{meta['document_title']} {meta['order_no']}")

    def header(page_no: int):
        pdf.set_fill_color(24, 36, 58)
        pdf.rect(0, 0, 210, 14, "F")
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Helvetica", "B", 11)
        pdf.set_xy(8, 3)
        pdf.cell(130, 6, _doc_safe_text(meta["company"]))
        pdf.set_font("Helvetica", "", 7)
        pdf.cell(0, 6, _doc_safe_text(meta["document_title"]), align="R")
        pdf.set_text_color(30, 30, 30)

        pdf.set_xy(8, 19)
        pdf.set_font("Helvetica", "B", 11)
        pdf.cell(0, 6, _doc_safe_text(meta["company"]))
        pdf.set_font("Helvetica", "", 7)
        pdf.set_xy(8, 25)
        pdf.multi_cell(115, 4, _doc_safe_text(meta["company_address"]))
        pdf.set_xy(132, 19)
        pdf.set_font("Helvetica", "B", 8)
        pdf.cell(68, 5, "ORDEN DE COMPRA", align="R")
        pdf.set_font("Helvetica", "", 7)
        pdf.set_xy(132, 25)
        pdf.cell(68, 4, f"Fecha: {meta['date']}", align="R")
        pdf.set_xy(132, 30)
        pdf.cell(68, 4, f"Hora: {meta['time']}", align="R")
        pdf.set_xy(132, 35)
        pdf.cell(68, 4, f"Pag: {page_no}", align="R")

        y = 43
        pdf.set_fill_color(238, 242, 249)
        pdf.rect(8, y, 194, 29, "F")
        pdf.set_draw_color(214, 222, 234)
        pdf.rect(8, y, 194, 29)
        pdf.set_font("Helvetica", "B", 7)
        pdf.set_xy(11, y + 3)
        pdf.cell(24, 4, "Telefono:")
        pdf.set_font("Helvetica", "", 7)
        pdf.cell(40, 4, meta["company_phone"])
        pdf.set_font("Helvetica", "B", 7)
        pdf.cell(16, 4, "Rif:")
        pdf.set_font("Helvetica", "", 7)
        pdf.cell(35, 4, meta["company_rif"])
        pdf.set_font("Helvetica", "B", 7)
        pdf.cell(27, 4, "Orden:")
        pdf.set_font("Helvetica", "", 7)
        pdf.cell(42, 4, meta["order_no"])
        pdf.ln(6)
        pdf.set_x(11)
        pdf.set_font("Helvetica", "B", 7)
        pdf.cell(24, 4, "Proveedor:")
        pdf.set_font("Helvetica", "", 7)
        pdf.cell(117, 4, _doc_safe_text(meta["supplier"]))
        pdf.set_font("Helvetica", "B", 7)
        pdf.cell(18, 4, "Sede:")
        pdf.set_font("Helvetica", "", 7)
        pdf.cell(30, 4, _doc_safe_text(site), align="R")
        pdf.ln(5)
        pdf.set_x(11)
        pdf.set_font("Helvetica", "B", 7)
        pdf.cell(24, 4, "Direccion:")
        pdf.set_font("Helvetica", "", 7)
        pdf.cell(117, 4, _doc_safe_text(meta["supplier_address"]))
        pdf.set_font("Helvetica", "B", 7)
        pdf.cell(18, 4, "Telefono:")
        pdf.set_font("Helvetica", "", 7)
        pdf.cell(30, 4, _doc_safe_text(meta["supplier_phone"]), align="R")
        pdf.set_xy(8, y + 34)

        widths = [25, 91, 18, 27, 33]
        headers = ["Código", "Descripción", "Cantidad", "Precio Unitario", "Total"]
        pdf.set_fill_color(53, 104, 245)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Helvetica", "B", 7)
        for w, h in zip(widths, headers):
            pdf.cell(w, 7, h, border=1, fill=True, align="C")
        pdf.ln()
        pdf.set_text_color(30, 30, 30)

    header(1)

    widths = [25, 91, 18, 27, 33]
    if x.empty:
        pdf.set_font("Helvetica", "", 8)
        pdf.cell(194, 8, "Sin productos pendientes de compra para esta sede.", border=1)
        pdf.ln(8)
    else:
        pdf.set_font("Helvetica", "", 7)
        for row_idx, row in enumerate(x.to_dict("records"), start=1):
            if pdf.get_y() > 270:
                pdf.add_page()
                header(pdf.page_no())
            desc = _doc_safe_text(row["Descripción"])
            pdf.cell(25, 6, _doc_safe_text(row["Código"]), border="LR")
            pdf.cell(91, 6, desc[:70], border="LR")
            pdf.cell(18, 6, _quantity_text_v73(row["Cantidad"]), border="LR", align="R")
            pdf.cell(27, 6, f"{float(row['Precio Unitario']):,.2f}", border="LR", align="R")
            pdf.cell(33, 6, f"{float(row['Total']):,.2f}", border="LR", align="R")
            pdf.ln()
        pdf.cell(sum(widths), 0, "", border="T")

    pdf.ln(7)
    pdf.set_font("Helvetica", "", 8)
    summary = [
        ("Total:", total),
        ("Flete:", 0.0),
        ("Descuento:", 0.0),
        ("Otro descuento:", 0.0),
        ("Total Neto:", total),
    ]
    x0 = 126
    for label, value in summary:
        pdf.set_x(x0)
        if label == "Total Neto:":
            pdf.set_font("Helvetica", "B", 9)
        pdf.cell(42, 6, label, align="R")
        pdf.cell(36, 6, f"{value:,.2f}", align="R")
        pdf.ln()
        pdf.set_font("Helvetica", "", 8)

    pdf.set_y(min(pdf.get_y() + 5, 278))
    pdf.set_draw_color(210, 218, 230)
    pdf.line(8, pdf.get_y(), 202, pdf.get_y())
    pdf.ln(3)
    pdf.set_font("Helvetica", "B", 7)
    pdf.cell(45, 5, "Operador:")
    pdf.set_font("Helvetica", "", 7)
    pdf.cell(55, 5, meta["operator"])
    pdf.set_font("Helvetica", "B", 7)
    pdf.cell(42, 5, "Total Items:")
    pdf.set_font("Helvetica", "", 7)
    pdf.cell(30, 5, f"{len(x):,}", align="R")
    if len(x) > 0 and pdf.page_no() > 1:
        pdf.ln(8)
        pdf.set_font("Helvetica", "I", 7)
        pdf.cell(0, 5, "Continua...", align="R")

    pdf.set_y(287)
    pdf.set_font("Helvetica", "", 6)
    pdf.set_text_color(120, 120, 120)
    pdf.cell(0, 4, f"MAKROPETROL NEXUS v{APP_VERSION} · Documento generado automaticamente", align="C")

    return bytes(pdf.output())


def build_withdrawal_pdf(
    data: pd.DataFrame,
    site: str,
    snapshot_label: str,
    sequence: int = 1,
) -> bytes | None:
    if not HAS_PDF:
        return None

    x = _doc_prepare_frame(data, "withdrawal", site)
    total = float(monetary_sum_v71(x["Total"])) if not x.empty else 0.0
    meta = _document_purchase_meta(site, snapshot_label, sequence=sequence)
    meta["document_title"] = "RETIRO DE ALMACÉN"
    meta["order_no"] = _doc_order_number("withdrawal", site, snapshot_label, sequence)

    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=14)
    pdf.add_page()

    def hdr(page_no: int):
        pdf.set_fill_color(24, 36, 58)
        pdf.rect(0, 0, 210, 14, "F")
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Helvetica", "B", 11)
        pdf.set_xy(8, 3)
        pdf.cell(135, 6, _doc_safe_text(COMPANY_INFO["name"]))
        pdf.set_font("Helvetica", "", 7)
        pdf.cell(0, 6, "RETIRO DE ALMACÉN", align="R")
        pdf.set_text_color(30, 30, 30)
        pdf.set_xy(8, 20)
        pdf.set_font("Helvetica", "B", 12)
        pdf.cell(125, 6, "RETIRO DE ALMACÉN")
        pdf.set_font("Helvetica", "", 7)
        pdf.cell(0, 6, f"Pag: {page_no}", align="R")
        pdf.set_xy(8, 28)
        pdf.cell(194, 5, f"Sede: {_doc_safe_text(site)}")
        pdf.set_xy(8, 34)
        pdf.cell(194, 5, f"Fecha: {snapshot_label}    Hora: {meta['time']}    Documento: {meta['order_no']}")
        pdf.set_xy(8, 44)
        widths=[25,91,18,27,33]
        heads=["Código","Descripción","Cantidad","Costo Unitario","Valor Retiro"]
        pdf.set_fill_color(53,104,245)
        pdf.set_text_color(255,255,255)
        pdf.set_font("Helvetica","B",7)
        for w,h in zip(widths,heads):
            pdf.cell(w,7,h,border=1,fill=True,align="C")
        pdf.ln()
        pdf.set_text_color(30,30,30)

    hdr(1)
    if x.empty:
        pdf.set_font("Helvetica","",8)
        pdf.cell(194,8,"Sin productos pendientes de retiro para esta sede.",border=1)
    else:
        for row in x.to_dict("records"):
            if pdf.get_y() > 270:
                pdf.add_page()
                hdr(pdf.page_no())
            pdf.set_font("Helvetica","",7)
            pdf.cell(25,6,_doc_safe_text(row["Código"]),border="LR")
            pdf.cell(91,6,_doc_safe_text(row["Descripción"])[:70],border="LR")
            pdf.cell(18,6,_quantity_text_v73(row["Cantidad"]),border="LR",align="R")
            pdf.cell(27,6,f"{float(row['Precio Unitario']):,.2f}",border="LR",align="R")
            pdf.cell(33,6,f"{float(row['Total']):,.2f}",border="LR",align="R")
            pdf.ln()
    pdf.ln(6)
    pdf.set_font("Helvetica","B",8)
    pdf.cell(145,6,"Valor total del retiro:",align="R")
    pdf.cell(45,6,f"{total:,.2f}",align="R")
    pdf.ln(8)
    pdf.set_font("Helvetica","",7)
    pdf.cell(45,5,"Motivo:")
    pdf.cell(149,5,"SOBRESTOCK / REDISTRIBUCIÓN / AJUSTE OPERATIVO")
    pdf.ln(8)
    pdf.cell(45,5,"Responsable de salida:")
    pdf.cell(70,5,"")
    pdf.cell(38,5,"Total Items:")
    pdf.cell(41,5,f"{len(x):,}",align="R")
    pdf.set_y(287)
    pdf.set_font("Helvetica","",6)
    pdf.set_text_color(120,120,120)
    pdf.cell(0,4,f"MAKROPETROL NEXUS v{APP_VERSION} · Documento generado automaticamente",align="C")
    return bytes(pdf.output())


def build_redistribution_pdf(
    transfers: pd.DataFrame,
    snapshot_label: str,
    sequence: int = 1,
) -> bytes | None:
    if not HAS_PDF:
        return None

    x = _doc_prepare_frame(transfers, "redistribution")
    units = int(pd.to_numeric(x.get("Unidades Sugeridas", 0), errors="coerce").fillna(0).sum()) if not x.empty else 0
    avoided = float(pd.to_numeric(x.get("Compra Evitada Estimada ($)", 0), errors="coerce").sum(min_count=len(x))) if not x.empty else 0.0
    doc_no = _doc_order_number("redistribution", "GLOBAL", snapshot_label, sequence)

    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=14)
    pdf.add_page()

    def hdr(page_no: int):
        pdf.set_fill_color(24,36,58)
        pdf.rect(0,0,210,15,"F")
        pdf.set_text_color(255,255,255)
        pdf.set_font("Helvetica","B",11)
        pdf.set_xy(8,4)
        pdf.cell(135,6,_doc_safe_text(COMPANY_INFO["name"]))
        pdf.set_font("Helvetica","",7)
        pdf.cell(0,6,"REDISTRIBUCIÓN MULTISEDE",align="R")
        pdf.set_text_color(30,30,30)
        pdf.set_xy(8,21)
        pdf.set_font("Helvetica","B",12)
        pdf.cell(125,6,"PLAN DE REDISTRIBUCIÓN")
        pdf.set_font("Helvetica","",7)
        pdf.cell(0,6,f"Pag: {page_no}",align="R")
        pdf.set_xy(8,30)
        pdf.cell(194,5,f"Fecha: {snapshot_label}    Hora: {datetime.now().strftime('%H:%M:%S')}    Documento: {doc_no}")
        pdf.set_xy(8,40)
        widths=[23,68,31,31,18,23]
        heads=["Código","Descripción","Origen","Destino","Cantidad","Ahorro Est."]
        pdf.set_fill_color(23,170,190)
        pdf.set_text_color(255,255,255)
        pdf.set_font("Helvetica","B",7)
        for w,h in zip(widths,heads):
            pdf.cell(w,7,h,border=1,fill=True,align="C")
        pdf.ln()
        pdf.set_text_color(30,30,30)

    hdr(1)
    if x.empty:
        pdf.set_font("Helvetica","",8)
        pdf.cell(194,8,"No existen transferencias sugeridas.",border=1)
    else:
        for row in x.to_dict("records"):
            if pdf.get_y() > 270:
                pdf.add_page()
                hdr(pdf.page_no())
            pdf.set_font("Helvetica","",7)
            pdf.cell(23,6,_doc_safe_text(row["Código"]),border="LR")
            pdf.cell(68,6,_doc_safe_text(row["Descripción"])[:50],border="LR")
            pdf.cell(31,6,_doc_safe_text(row["Origen"])[:20],border="LR")
            pdf.cell(31,6,_doc_safe_text(row["Destino"])[:20],border="LR")
            pdf.cell(18,6,f"{int(row['Unidades Sugeridas']):,}",border="LR",align="R")
            pdf.cell(23,6,f"{float(row['Compra Evitada Estimada ($)']):,.0f}",border="LR",align="R")
            pdf.ln()

    pdf.ln(7)
    pdf.set_font("Helvetica","B",8)
    pdf.cell(90,6,"Unidades a redistribuir:",align="R")
    pdf.cell(50,6,f"{units:,}",align="R")
    pdf.ln(6)
    pdf.cell(90,6,"Compra potencial evitada:",align="R")
    pdf.cell(50,6,f"${avoided:,.2f}",align="R")
    pdf.ln(10)
    pdf.set_font("Helvetica","I",7)
    pdf.multi_cell(194,5,"Criterio: mover excedentes existentes hacia sedes con déficit antes de generar compras nuevas.")
    pdf.set_y(287)
    pdf.set_font("Helvetica","",6)
    pdf.set_text_color(120,120,120)
    pdf.cell(0,4,f"MAKROPETROL NEXUS v{APP_VERSION} · Documento multisede generado automaticamente",align="C")
    return bytes(pdf.output())


def build_alerts_pdf(
    alerts: pd.DataFrame,
    site: str,
    snapshot_label: str,
    sequence: int = 1,
) -> bytes | None:
    if not HAS_PDF:
        return None
    x = _doc_prepare_frame(alerts, "alerts", site)
    doc_no = _doc_order_number("alerts", site, snapshot_label, sequence)
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=14)
    pdf.add_page()

    def hdr(page_no: int):
        pdf.set_fill_color(24,36,58)
        pdf.rect(0,0,210,15,"F")
        pdf.set_text_color(255,255,255)
        pdf.set_font("Helvetica","B",11)
        pdf.set_xy(8,4)
        pdf.cell(135,6,_doc_safe_text(COMPANY_INFO["name"]))
        pdf.set_font("Helvetica","",7)
        pdf.cell(0,6,"REPORTE DE ALERTAS",align="R")
        pdf.set_text_color(30,30,30)
        pdf.set_xy(8,21)
        pdf.set_font("Helvetica","B",12)
        pdf.cell(125,6,"REPORTE DE ALERTAS OPERATIVAS")
        pdf.set_font("Helvetica","",7)
        pdf.cell(0,6,f"Pag: {page_no}",align="R")
        pdf.set_xy(8,30)
        pdf.cell(194,5,f"Sede: {site}    Fecha: {snapshot_label}    Documento: {doc_no}")
        pdf.set_xy(8,40)
        widths=[24,48,33,22,67]
        heads=["Código","Descripción","Alerta","Severidad","Detalle"]
        pdf.set_fill_color(223,85,85)
        pdf.set_text_color(255,255,255)
        pdf.set_font("Helvetica","B",7)
        for w,h in zip(widths,heads):
            pdf.cell(w,7,h,border=1,fill=True,align="C")
        pdf.ln()
        pdf.set_text_color(30,30,30)
    hdr(1)
    if x.empty:
        pdf.set_font("Helvetica","",8)
        pdf.cell(194,8,"Sin alertas para esta sede.",border=1)
    else:
        for row in x.to_dict("records"):
            if pdf.get_y() > 268:
                pdf.add_page()
                hdr(pdf.page_no())
            pdf.set_font("Helvetica","",6.5)
            pdf.cell(24,7,_doc_safe_text(row["Código"]),border=1)
            pdf.cell(48,7,_doc_safe_text(row["Descripción"])[:34],border=1)
            pdf.cell(33,7,_doc_safe_text(row["Alerta"])[:25],border=1)
            pdf.cell(22,7,_doc_safe_text(row["Severidad"]),border=1,align="C")
            pdf.cell(67,7,_doc_safe_text(row["Detalle"])[:54],border=1)
            pdf.ln()
    pdf.set_y(287)
    pdf.set_font("Helvetica","",6)
    pdf.set_text_color(120,120,120)
    pdf.cell(0,4,f"MAKROPETROL NEXUS v{APP_VERSION} · Documento generado automaticamente",align="C")
    return bytes(pdf.output())


def build_operational_html(
    document_type: str,
    data: pd.DataFrame,
    site: str | None,
    snapshot_label: str,
    title: str,
) -> bytes:
    if document_type == "redistribution":
        x = _doc_prepare_frame(data, document_type)
    else:
        x = _doc_prepare_frame(data, document_type, site)

    css = """
    :root{--navy:#18243a;--blue:#3568f5;--cyan:#17aabe;--line:#dfe6ef;--bg:#f4f7fb;--muted:#748198}
    *{box-sizing:border-box} body{margin:0;background:var(--bg);color:#172137;font-family:Arial,Segoe UI,sans-serif}
    .page{max-width:1100px;margin:24px auto;padding:28px;background:#fff;border:1px solid #e1e7ef;box-shadow:0 18px 50px rgba(25,45,75,.08);border-radius:20px}
    .top{background:linear-gradient(135deg,#18243a,#263b65);color:#fff;border-radius:15px;padding:18px 20px;display:flex;justify-content:space-between;gap:20px}
    .brand{font-size:21px;font-weight:800}.doc{font-weight:800;text-align:right}.muted{color:var(--muted);font-size:12px}.meta{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin:16px 0}
    .box{border:1px solid var(--line);border-radius:12px;padding:12px;background:#fbfcfe}.label{font-size:10px;font-weight:800;text-transform:uppercase;color:var(--muted)}.value{font-size:13px;font-weight:700;margin-top:3px}
    .kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin:16px 0}.kpi{border:1px solid var(--line);border-radius:12px;padding:14px}.kpi strong{display:block;font-size:22px;margin-top:4px}
    table{width:100%;border-collapse:collapse;font-size:12px;margin-top:16px} th{background:var(--blue);color:#fff;text-align:left;padding:9px} td{padding:8px;border-bottom:1px solid var(--line)} tr:nth-child(even){background:#fafcff}
    .foot{margin-top:20px;padding-top:12px;border-top:1px solid var(--line);font-size:11px;color:var(--muted);display:flex;justify-content:space-between}
    .pill{display:inline-block;padding:5px 8px;border-radius:999px;background:#edf2ff;color:#355bd7;font-weight:800;font-size:10px}
    @media print{body{background:#fff}.page{box-shadow:none;border:0;margin:0;max-width:none;border-radius:0}.no-print{display:none!important}}
    @media(max-width:700px){.meta,.kpis{grid-template-columns:1fr 1fr}.top{flex-direction:column}.doc{text-align:left}}
    """
    headers = list(x.columns)
    rows = []
    for rec in x.to_dict("records"):
        row = "<tr>" + "".join(f"<td>{(_quantity_text_v73(rec.get(c)) if c == "Cantidad" else _doc_safe_text(rec.get(c,'')))}</td>" for c in headers) + "</tr>"
        rows.append(row)

    if document_type == "purchase":
        qty = float(pd.to_numeric(x.get("Cantidad", 0), errors="coerce").fillna(0).sum()) if not x.empty else 0
        total = float(pd.to_numeric(x.get("Total", 0), errors="coerce").sum(min_count=len(x))) if not x.empty else 0.0
        kpi_html = f"""
        <div class="kpi"><div class="label">Items</div><strong>{len(x):,}</strong></div>
        <div class="kpi"><div class="label">Unidades</div><strong>{_quantity_text_v73(qty)}</strong></div>
        <div class="kpi"><div class="label">Total</div><strong>${total:,.2f}</strong></div>
        <div class="kpi"><div class="label">Sede</div><strong>{_doc_safe_text(site)}</strong></div>
        """
    elif document_type == "withdrawal":
        qty = float(pd.to_numeric(x.get("Cantidad", 0), errors="coerce").fillna(0).sum()) if not x.empty else 0
        total = float(pd.to_numeric(x.get("Total", 0), errors="coerce").sum(min_count=len(x))) if not x.empty else 0.0
        kpi_html = f"""
        <div class="kpi"><div class="label">Items a retirar</div><strong>{len(x):,}</strong></div>
        <div class="kpi"><div class="label">Unidades</div><strong>{_quantity_text_v73(qty)}</strong></div>
        <div class="kpi"><div class="label">Capital liberable</div><strong>${total:,.2f}</strong></div>
        <div class="kpi"><div class="label">Sede</div><strong>{_doc_safe_text(site)}</strong></div>
        """
    elif document_type == "redistribution":
        units = int(pd.to_numeric(x.get("Unidades Sugeridas", 0), errors="coerce").fillna(0).sum()) if not x.empty else 0
        avoided = float(pd.to_numeric(x.get("Compra Evitada Estimada ($)", 0), errors="coerce").sum(min_count=len(x))) if not x.empty else 0.0
        routes = int(len(x)) if not x.empty else 0
        kpi_html = f"""
        <div class="kpi"><div class="label">Movimientos</div><strong>{routes:,}</strong></div>
        <div class="kpi"><div class="label">Unidades</div><strong>{units:,}</strong></div>
        <div class="kpi"><div class="label">Compra evitada</div><strong>${avoided:,.2f}</strong></div>
        <div class="kpi"><div class="label">Ámbito</div><strong>Multisede</strong></div>
        """
    else:
        critical = int((x["Severidad"].astype(str) == "CRÍTICA").sum()) if "Severidad" in x.columns else 0
        kpi_html = f"""
        <div class="kpi"><div class="label">Alertas</div><strong>{len(x):,}</strong></div>
        <div class="kpi"><div class="label">Críticas</div><strong>{critical:,}</strong></div>
        <div class="kpi"><div class="label">Sede</div><strong>{_doc_safe_text(site)}</strong></div>
        <div class="kpi"><div class="label">Corte</div><strong>{snapshot_label}</strong></div>
        """

    meta_left = f"""
    <div class="box"><div class="label">Empresa</div><div class="value">{_doc_safe_text(COMPANY_INFO["name"])}</div>
    <div class="muted">{_doc_safe_text(COMPANY_INFO["address"])}</div></div>
    <div class="box"><div class="label">Documento</div><div class="value">{_doc_safe_text(title)}</div>
    <div class="muted">Fecha: {snapshot_label} · Hora: {datetime.now().strftime('%H:%M:%S')}</div></div>
    """
    table = "<table><thead><tr>" + "".join(f"<th>{c}</th>" for c in headers) + "</tr></thead><tbody>" + "".join(rows) + "</tbody></table>"
    html = f"""<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
    <title>{title}</title><style>{css}</style></head><body><main class="page">
    <div class="top"><div><div class="brand">{_doc_safe_text(COMPANY_INFO["name"])}</div><div style="font-size:11px;margin-top:4px">{_doc_safe_text(COMPANY_INFO["address"])}</div></div>
    <div class="doc">{_doc_safe_text(title)}<div style="font-size:11px;margin-top:4px">Corte {snapshot_label}</div></div></div>
    <div class="meta">{meta_left}</div><div class="kpis">{kpi_html}</div>{table}
    <div class="foot"><span>Makropetrol NEXUS v{APP_VERSION}</span><span>Generado automáticamente · listo para imprimir / PDF</span></div>
    </main></body></html>"""
    return html.encode("utf-8")


@st.cache_data(show_spinner=False)
def build_purchase_excel(data: pd.DataFrame, site: str, snapshot_label: str, sequence: int = 1) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    x = _doc_prepare_frame(data, "purchase", site)
    qty = float(x["Cantidad"].sum()) if not x.empty else 0
    total = float(monetary_sum_v71(x["Total"])) if not x.empty else 0.0
    out = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Orden de Compra"
    navy, blue, line, white = "18243A", "3568F5", "DDE5EF", "FFFFFF"
    ws.merge_cells("A1:E1")
    ws["A1"] = COMPANY_INFO["name"]
    ws["A1"].font = Font(size=16, bold=True, color=white)
    ws["A1"].fill = PatternFill("solid", fgColor=navy)
    ws.row_dimensions[1].height = 28
    ws.merge_cells("A2:E2")
    ws["A2"] = COMPANY_INFO["address"]
    ws["A2"].font = Font(size=9, color="667287")
    ws["A3"] = "ORDEN DE COMPRA"
    ws["A3"].font = Font(size=14, bold=True, color=navy)
    ws["D3"] = f"Fecha: {snapshot_label}"
    ws["E3"] = f"Pag: 1"
    ws["A4"] = f"Telefono: {COMPANY_INFO['phone']}"
    ws["B4"] = f"Rif: {COMPANY_INFO['rif']}"
    ws["C4"] = f"Sede: {site}"
    ws["D4"] = "Operador:"
    ws["E4"] = "NEXUS / ADMINISTRACION"

    headers = ["Código","Descripción","Cantidad","Precio Unitario","Total"]
    for j,h in enumerate(headers,1):
        c=ws.cell(6,j,h)
        c.fill=PatternFill("solid",fgColor=blue)
        c.font=Font(bold=True,color=white)
        c.alignment=Alignment(horizontal="center")
    for i, row in enumerate(x.to_dict("records"), 7):
        vals = [
            row["Código"],
            row["Descripción"],
            float(row["Cantidad"]),
            float(row["Precio Unitario"]),
            float(row["Total"]),
        ]
        for j, val in enumerate(vals, 1):
            ws.cell(i, j, val)
    total_row=7+len(x)
    ws.cell(total_row+1,4,"Total Neto:")
    ws.cell(total_row+1,4).font=Font(bold=True)
    ws.cell(total_row+1,5,total)
    ws.cell(total_row+1,5).font=Font(bold=True)
    ws.cell(total_row+2,1,"Total Items:")
    ws.cell(total_row+2,2,len(x))
    ws.cell(total_row+2,3,"Unidades:")
    ws.cell(total_row+2,4,qty)
    for row in ws.iter_rows():
        for cell in row:
            cell.border=Border(bottom=Side(style="hair",color=line))
            cell.alignment=Alignment(vertical="center",wrap_text=True)
    ws.freeze_panes="A7"
    widths=[18,62,12,18,18]
    for idx,w in enumerate(widths,1):
        ws.column_dimensions[get_column_letter(idx)].width=w
    for r in range(7,total_row+1):
        ws.cell(r,4).number_format='$#,##0.00'
        ws.cell(r,5).number_format='$#,##0.00'
    ws.cell(total_row+1,5).number_format='$#,##0.00'
    wb.save(out)
    return out.getvalue()


# ============================================================
# EXCEL PREMIUM
# ============================================================

@st.cache_data(show_spinner=False)
def export_excel(
    site_results: dict[str, pd.DataFrame],
    transfers: pd.DataFrame,
    months_history: int,
    min_coverage: float,
    max_coverage: float,
    lead_time_days: int,
    snapshot_label: str,
) -> bytes:
    from openpyxl import Workbook
    from openpyxl.chart import BarChart, Reference
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.utils import get_column_letter

    all_df = combine_sites(site_results)
    out = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "00 Dashboard"

    navy = "18243A"
    blue = "3568F5"
    gray = "7B8799"
    line = "DDE5EF"
    white = "FFFFFF"

    thin = Side(style="thin", color=line)

    def title(ws_, title, subtitle, end_col=10):
        ws_.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_col)
        ws_.cell(1, 1).value = title
        ws_.cell(1, 1).font = Font(size=20, bold=True, color=white)
        ws_.cell(1, 1).fill = PatternFill("solid", fgColor=navy)
        ws_.cell(1, 1).alignment = Alignment(vertical="center")
        ws_.row_dimensions[1].height = 34
        ws_.merge_cells(start_row=2, start_column=1, end_row=2, end_column=end_col)
        ws_.cell(2, 1).value = subtitle
        ws_.cell(2, 1).font = Font(size=10, color=gray, italic=True)
        ws_.row_dimensions[2].height = 23

    def write_table(ws_, df_: pd.DataFrame, start_row=4, start_col=1, table_name="Table1"):
        if df_.empty:
            ws_.cell(start_row, start_col).value = "Sin registros"
            return
        for j, col in enumerate(df_.columns, start_col):
            c = ws_.cell(start_row, j, col)
            c.fill = PatternFill("solid", fgColor=blue)
            c.font = Font(bold=True, color=white)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = Border(bottom=thin)
        for i, row in enumerate(df_.itertuples(index=False), start_row + 1):
            for j, val in enumerate(row, start_col):
                cell = ws_.cell(i, j, val)
                cell.border = Border(bottom=Side(style="hair", color=line))
                cell.alignment = Alignment(vertical="center")
        end_row = start_row + len(df_)
        end_col = start_col + len(df_.columns) - 1
        ref = f"{get_column_letter(start_col)}{start_row}:{get_column_letter(end_col)}{end_row}"
        table = Table(displayName=table_name, ref=ref)
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True, showColumnStripes=False)
        ws_.add_table(table)
        ws_.freeze_panes = f"{get_column_letter(start_col)}{start_row+1}"
        ws_.auto_filter.ref = ref
        for j in range(start_col, end_col + 1):
            max_len = max(len(str(ws_.cell(r, j).value or "")) for r in range(start_row, min(end_row, start_row + 100) + 1))
            ws_.column_dimensions[get_column_letter(j)].width = min(max(max_len + 2, 11), 42)

    score = health_score(all_df)
    kpis = [
        ("Valor inventario", float(monetary_sum_v71(all_df["Valor Inventario ($)"])) if not all_df.empty else 0),
        ("Compras", int((all_df.get("Compra Ajustada", all_df.get("Compra Sugerida", pd.Series(dtype=float))) > 0).sum()) if not all_df.empty else 0),
        ("Retiros", int((all_df.get("Retiro Almacén", pd.Series(dtype=float)) > 0).sum()) if not all_df.empty else 0),
        ("Capital inmovilizado", float(monetary_sum_v71(all_df["Capital Inmovilizado ($)"])) if not all_df.empty else 0),
    ]
    title(ws, "MAKROPETROL NEXUS · REPORTE EJECUTIVO", f"Corte: {snapshot_label} · Historial: {months_history} meses · Lead time: {lead_time_days} días", 10)
    ws["A4"] = "SALUD LOGÍSTICA"
    ws["A4"].font = Font(bold=True, color=navy, size=12)
    ws["B4"] = f"{score}/100 · {health_label(score)}"
    ws["B4"].font = Font(bold=True, color=blue, size=16)
    for idx, (label, value) in enumerate(kpis, 1):
        col = 1 + (idx - 1) * 2
        ws.cell(6, col, label).font = Font(bold=True, color=gray, size=9)
        ws.cell(7, col, value)
        ws.cell(7, col).font = Font(bold=True, color=navy, size=15)
        if "Valor" in label or "Capital" in label:
            ws.cell(7, col).number_format = '$#,##0.00'
        else:
            ws.cell(7, col).number_format = '#,##0'

    if not all_df.empty:
        by_site = all_df.groupby("Sede").agg(
            SKUs=("Código", "count"),
            Compras=("Compra Sugerida", lambda x: int((x > 0).sum())),
            Retiros=("Retiro Almacén", lambda x: int((x > 0).sum())),
            Inventario=("Valor Inventario ($)", "sum"),
            Inmovilizado=("Capital Inmovilizado ($)", "sum"),
        ).reset_index()
        write_table(ws, by_site, start_row=10, table_name="ResumenSedes")

        chart = BarChart()
        chart.title = "Valor de inventario por sede"
        data_ref = Reference(ws, min_col=6, min_row=10, max_row=10 + len(by_site))
        cat_ref = Reference(ws, min_col=1, min_row=11, max_row=10 + len(by_site))
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cat_ref)
        chart.height = 7
        chart.width = 12
        ws.add_chart(chart, "H10")

    sheets = [
        ("01 Inventario", all_df, "InventarioGeneral"),
        ("02 Compras", all_df[all_df.get("Compra Ajustada", all_df.get("Compra Sugerida", pd.Series(dtype=float))) > 0].copy() if not all_df.empty else all_df, "PlanCompras"),
        ("03 Retiros", all_df[all_df.get("Retiro Almacén", pd.Series(dtype=float)) > 0].copy() if not all_df.empty else all_df, "PlanRetiros"),
        ("04 Redistribución", transfers.copy(), "PlanTransferencias"),
        ("05 Alertas", build_alerts(all_df), "Alertas"),
    ]
    for sheet_name, data, table_name in sheets:
        sh = wb.create_sheet(sheet_name)
        title(sh, sheet_name.replace("01 ", "").replace("02 ", "").replace("03 ", "").replace("04 ", "").replace("05 ", ""), f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}", min(max(len(data.columns), 8), 16) if not data.empty else 8)
        write_table(sh, data, start_row=4, table_name=table_name)

    crit = wb.create_sheet("06 Criterios")
    title(crit, "CRITERIOS DEL MOTOR", "Parámetros y definiciones usados en el corte", 5)
    criteria = pd.DataFrame(
        [
            ["Meses de historial", months_history, "Divide el total de ventas del período para obtener demanda mensual."],
            ["Cobertura mínima", min_coverage, "Meses de cobertura usados como umbral operativo."],
            ["Cobertura máxima", max_coverage, "Meses de cobertura objetivo antes de sugerir retiro."],
            ["Lead time", lead_time_days, "Días estimados de entrega del proveedor."],
            ["Stock seguridad", "Demanda diaria × días de seguridad", "Protección adicional ante variabilidad/entrega."],
            ["Punto de reorden", "máximo(stock mínimo, demanda lead time + stock seguridad)", "Umbral para iniciar abastecimiento."],
            ["ABC", "80% / 95% / resto", "Clasificación A/B/C según base seleccionada."],
            ["Redistribución", "Exceso de una sede → déficit de otra", "Se prioriza mover inventario existente antes de comprar."],
        ],
        columns=["Parámetro", "Valor", "Interpretación"],
    )
    write_table(crit, criteria, start_row=4, table_name="CriteriosMotor")

    for sh in wb.worksheets:
        sh.sheet_view.showGridLines = False
        for row in sh.iter_rows():
            for cell in row:
                cell.alignment = Alignment(vertical="center", wrap_text=False)
        headers = {cell.value: cell.column for cell in sh[4]}
        integer_headers = {"Ventas", "Demanda Mensual", "Existencia", "Stock Mínimo", "Stock Máximo", "Demanda Lead Time", "Stock Seguridad", "Punto de Reorden", "Compra Sugerida", "Compra Ajustada", "Retiro Almacén", "Unidades Sugeridas", "Transferencias Recibidas", "Transferencias Enviadas"}
        money_headers = {"Costo", "Valor Inventario ($)", "Capital Inmovilizado ($)", "Costo Compra Estimada ($)", "Valor Movimiento ($)", "Costo Unitario ($)", "Compra Evitada Estimada ($)"}
        for h, col in headers.items():
            if h in integer_headers:
                for c in sh.iter_cols(min_col=col, max_col=col, min_row=5, max_row=sh.max_row):
                    for cell in c:
                        cell.number_format = '#,##0'
            if h in money_headers:
                for c in sh.iter_cols(min_col=col, max_col=col, min_row=5, max_row=sh.max_row):
                    for cell in c:
                        cell.number_format = '$#,##0.00'

    wb.save(out)
    return out.getvalue()


# ============================================================
# NEXUS 6 — DATOS FINANCIEROS Y DESEMPEÑO INTEGRAL
# ============================================================

FIN_CURRENCIES = ["USD", "VES", "EUR"]
FIN_SALES_FIELDS = {
    "Fecha": ["fecha", "dia", "fecha venta"],
    "Ventas netas": ["ventas netas", "venta neta", "ventas sin iva", "net sales"],
    "Ventas crédito": ["ventas credito", "ventas a credito", "ventas credito base cxc", "credit sales"],
    "Costo de ventas": ["costo de ventas", "coste de ventas", "costo vendido", "cogs"],
    "Cobros": ["cobros", "cobros reales", "cobranza", "collections"],
}
FIN_AR_FIELDS = {
    "Factura": ["factura", "numero factura", "documento", "nro factura", "invoice"],
    "Cliente": ["cliente", "nombre cliente", "razon social", "customer"],
    "Emisión": ["emision", "fecha emision", "fecha factura", "issued"],
    "Vencimiento": ["vencimiento", "fecha vencimiento", "vence", "due"],
    "Saldo": ["saldo", "saldo pendiente", "por cobrar", "balance"],
    "Importe factura": ["importe factura", "total factura", "monto original", "importe"],
}
FIN_AGING = ["Al día", "1–30 días", "31–60 días", "61–90 días", "Más de 90 días"]
TASK_STATES = ["Pendiente", "En curso", "Bloqueada", "Resuelta"]


def backup_database_bytes() -> bytes:
    """SQLite.backup incluye las operaciones confirmadas presentes en WAL."""
    with tempfile.TemporaryDirectory(prefix="nexus_backup_") as folder:
        path = Path(folder) / "nexus_data.db"
        target = sqlite3.connect(path)
        try:
            with db() as source:
                source.backup(target)
        finally:
            target.close()
        return path.read_bytes()


def init_finance_db() -> None:
    with db() as con:
        exists = con.execute("SELECT 1 FROM sqlite_master WHERE name='financial_sales'").fetchone()
        has_data = con.execute("SELECT 1 FROM inventory_snapshots UNION ALL SELECT 1 FROM supplier_inventory_snapshots LIMIT 1").fetchone()
    if not exists and has_data:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        DB_PATH.with_name(f"nexus_respaldo_antes_v6_{stamp}.db").write_bytes(backup_database_bytes())
    with db() as con:
        con.executescript("""
            CREATE TABLE IF NOT EXISTS financial_sales (
                day TEXT NOT NULL, site TEXT NOT NULL, currency TEXT NOT NULL,
                net_sales REAL NOT NULL, credit_sales REAL NOT NULL,
                cogs REAL, collections REAL, PRIMARY KEY(day,site,currency));
            CREATE TABLE IF NOT EXISTS financial_ar_batches (
                as_of TEXT NOT NULL, site TEXT NOT NULL, currency TEXT NOT NULL,
                row_count INTEGER NOT NULL, PRIMARY KEY(as_of,site,currency));
            CREATE TABLE IF NOT EXISTS financial_ar (
                as_of TEXT NOT NULL, site TEXT NOT NULL, currency TEXT NOT NULL,
                invoice TEXT NOT NULL, client TEXT NOT NULL, issued TEXT NOT NULL,
                due TEXT NOT NULL, balance REAL NOT NULL, amount REAL,
                PRIMARY KEY(as_of,site,currency,invoice));
            CREATE TABLE IF NOT EXISTS financial_settings (
                site TEXT PRIMARY KEY, inventory_currency TEXT NOT NULL DEFAULT '');
            CREATE TABLE IF NOT EXISTS financial_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT, saved_at TEXT NOT NULL,
                kind TEXT NOT NULL, site TEXT NOT NULL, currency TEXT NOT NULL,
                start_date TEXT NOT NULL, end_date TEXT NOT NULL, rows INTEGER NOT NULL,
                file_name TEXT NOT NULL, sha256 TEXT NOT NULL);
            CREATE TABLE IF NOT EXISTS action_tasks (
                id TEXT PRIMARY KEY, source_key TEXT UNIQUE, title TEXT NOT NULL,
                site TEXT NOT NULL, owner TEXT NOT NULL DEFAULT '', due TEXT,
                status TEXT NOT NULL DEFAULT 'Pendiente', priority TEXT NOT NULL DEFAULT 'Media',
                notes TEXT NOT NULL DEFAULT '', created_at TEXT NOT NULL, updated_at TEXT NOT NULL);
            CREATE INDEX IF NOT EXISTS idx_fin_sales ON financial_sales(currency,site,day);
            CREATE INDEX IF NOT EXISTS idx_fin_ar ON financial_ar(currency,site,as_of);
            CREATE INDEX IF NOT EXISTS idx_fin_batch ON financial_ar_batches(currency,site,as_of);
        """)


def finance_version() -> str:
    with db() as con:
        version = str(con.execute("SELECT COALESCE(MAX(id),0) FROM financial_log").fetchone()[0])
    return version + ":" + v7_version()


def finance_settings() -> dict[str, str]:
    with db() as con:
        return dict(con.execute("SELECT site,inventory_currency FROM financial_settings").fetchall())


def financial_number(value: Any, optional: bool = False, decimal: str = ".") -> float:
    """Importes estrictos: jamás convierte un texto inválido en cero."""
    if pd.isna(value) or str(value).strip() == "":
        if optional:
            return np.nan
        raise ValueError("Hay un importe obligatorio vacío.")
    if isinstance(value, (int, float, np.integer, np.floating)):
        number = float(value)
    else:
        text = str(value).strip().replace("\u00a0", "").replace(" ", "")
        text = re.sub(r"(?i)(USD|VES|EUR|BS\.?|US\$|\$|€)", "", text)
        if text.startswith("(") and text.endswith(")"):
            text = "-" + text[1:-1]
        if decimal == ",":
            if not re.fullmatch(r"[+-]?(?:\d{1,3}(?:\.\d{3})+|\d+)(?:,\d+)?", text):
                raise ValueError(f"Importe inválido para coma decimal: {str(value)[:60]}")
            text = text.replace(".", "").replace(",", ".")
        else:
            if not re.fullmatch(r"[+-]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?", text):
                raise ValueError(f"Importe inválido para punto decimal: {str(value)[:60]}")
            text = text.replace(",", "")
        number = float(text)
    if not np.isfinite(number):
        raise ValueError("Hay importes infinitos o no numéricos.")
    return round(number, 2)


def financial_date(value: Any, day_first: bool = True) -> str:
    if pd.isna(value) or str(value).strip() == "":
        raise ValueError("Hay una fecha obligatoria vacía.")
    if isinstance(value, (int, float, np.integer, np.floating)):
        if not 1 <= float(value) <= 100000:
            raise ValueError(f"Fecha Excel inválida: {value}")
        parsed = pd.Timestamp("1899-12-30") + pd.to_timedelta(float(value), unit="D")
    elif isinstance(value, (date, datetime, pd.Timestamp)):
        parsed = pd.Timestamp(value)
    elif re.match(r"^\d{4}-\d{2}-\d{2}", str(value).strip()):
        parsed = pd.to_datetime(value, errors="coerce", yearfirst=True)
    else:
        parsed = pd.to_datetime(value, errors="coerce", dayfirst=day_first)
    if pd.isna(parsed):
        raise ValueError(f"Fecha inválida: {str(value)[:60]}")
    return parsed.date().isoformat()


@st.cache_data(show_spinner=False, max_entries=12)
def financial_sheets(raw: bytes) -> list[str]:
    with pd.ExcelFile(io.BytesIO(raw), engine=EXCEL_ENGINE) as book:
        return book.sheet_names


@st.cache_data(show_spinner=False, max_entries=12)
def read_financial_file(raw: bytes, name: str, sheet: str, header: int, delimiter: str) -> pd.DataFrame:
    if name.lower().endswith(".csv"):
        try:
            text = raw.decode("utf-8-sig")
        except UnicodeDecodeError:
            text = raw.decode("cp1252")
        frame = pd.read_csv(io.StringIO(text), sep=delimiter, header=header, dtype=object, keep_default_na=False)
    else:
        frame = _read_excel_fast(io.BytesIO(raw), sheet_name=sheet, header=header, dtype=object, keep_default_na=False)
    if len(frame) > 300000:
        raise ValueError("Divide el archivo en lotes de hasta 300.000 filas.")
    frame.columns = [str(c).strip() for c in frame.columns]
    if frame.columns.duplicated().any():
        raise ValueError("Hay columnas repetidas. Asigna un nombre distinto a cada columna.")
    return frame.replace(r"^\s*$", np.nan, regex=True).dropna(how="all").reset_index(drop=True)


def normalize_financial_frame(frame: pd.DataFrame, kind: str, mapping: dict[str, str],
                              site: str, currency: str, decimal: str, day_first: bool,
                              as_of: str | None = None) -> pd.DataFrame:
    schema = FIN_SALES_FIELDS if kind == "ventas" else FIN_AR_FIELDS
    optional = {"Costo de ventas", "Cobros"} if kind == "ventas" else {"Importe factura"}
    used = [mapping.get(c) for c in schema if mapping.get(c)]
    if len(used) != len(set(used)):
        raise ValueError("Una columna de origen está asignada a más de un campo.")
    for field in set(schema) - optional:
        if not mapping.get(field) or mapping[field] not in frame:
            raise ValueError(f"Falta asignar la columna obligatoria: {field}.")
    for col in frame:
        key = normalize_key(col)
        if key in {"sede", "sucursal", "site"}:
            values = frame[col].dropna().map(normalize_key)
            if (values != normalize_key(site)).any():
                raise ValueError("El archivo tiene otra sede. Separa las sedes antes de cargarlo.")
        if key in {"moneda", "currency"}:
            values = frame[col].dropna().astype(str).str.upper().str.strip()
            if (values != currency).any():
                raise ValueError("La moneda del archivo no coincide con la seleccionada.")
    result = pd.DataFrame(index=frame.index)
    for field in schema:
        source = mapping.get(field)
        result[field] = frame[source] if source else np.nan
    numbers = ["Ventas netas", "Ventas crédito", "Costo de ventas", "Cobros"] if kind == "ventas" else ["Saldo", "Importe factura"]
    for field in numbers:
        values = []
        for i, val in enumerate(result[field]):
            try:
                values.append(financial_number(val, field in optional, decimal))
            except ValueError as exc:
                raise ValueError(f"Fila de datos {i + 1}, {field}: {exc}") from exc
        result[field] = pd.Series(values, index=result.index, dtype=float)
    for field in (["Fecha"] if kind == "ventas" else ["Emisión", "Vencimiento"]):
        result[field] = result[field].map(lambda v: financial_date(v, day_first))
    if kind == "ventas":
        if result["Fecha"].duplicated().any():
            raise ValueError("Debe haber una sola fila de totales por día y sede. Hay fechas repetidas.")
        if (result["Fecha"] > date.today().isoformat()).any():
            raise ValueError("Las ventas reales no pueden tener fechas futuras.")
        # Ventas/costos/cobros netos pueden ser negativos por devoluciones y ajustes reales.
        return result.sort_values("Fecha").reset_index(drop=True)
    for field in ["Factura", "Cliente"]:
        result[field] = result[field].fillna("").astype(str).str.strip()
        if (result[field] == "").any():
            raise ValueError(f"Hay facturas sin {field.lower()}.")
    if result["Factura"].str.casefold().duplicated().any():
        raise ValueError("Hay facturas repetidas en esta sede. Usa serie y número completos; una fila por factura.")
    if (result["Saldo"] < 0).any():
        raise ValueError("CxC admite saldos pendientes desde cero. Aplica las notas de crédito a sus facturas antes de importar.")
    if (result["Vencimiento"] < result["Emisión"]).any():
        raise ValueError("Hay vencimientos anteriores a la emisión.")
    if as_of and (result["Emisión"] > as_of).any():
        raise ValueError("Hay facturas emitidas después del corte de CxC.")
    bad_amount = result["Importe factura"].notna() & ((result["Importe factura"] < 0) | (result["Saldo"] > result["Importe factura"] + .01))
    if bad_amount.any():
        raise ValueError("Hay saldos superiores al importe original o importes de factura negativos.")
    return result.reset_index(drop=True)


def save_financial_data(frame: pd.DataFrame, kind: str, site: str, currency: str,
                        file_name: str, raw: bytes, as_of: str | None = None) -> None:
    if site not in SEDES or currency not in FIN_CURRENCIES or kind not in {"ventas", "cxc"}:
        raise ValueError("Sede, moneda o tipo de dato inválidos.")
    if kind == "ventas" and frame.empty:
        raise ValueError("No hay ventas para guardar.")
    if kind == "cxc" and (not as_of or as_of > date.today().isoformat()):
        raise ValueError("Selecciona un corte de CxC válido, hasta hoy.")
    def nullable(v):
        return None if pd.isna(v) else float(v)
    with db() as con:
        if kind == "ventas":
            rows = [(r[0], site, currency, float(r[1]), float(r[2]), nullable(r[3]), nullable(r[4])) for r in frame[list(FIN_SALES_FIELDS)].itertuples(index=False, name=None)]
            con.executemany("""INSERT INTO financial_sales VALUES (?,?,?,?,?,?,?)
                ON CONFLICT(day,site,currency) DO UPDATE SET net_sales=excluded.net_sales,
                credit_sales=excluded.credit_sales,cogs=excluded.cogs,collections=excluded.collections""", rows)
            start, end = str(frame["Fecha"].min()), str(frame["Fecha"].max())
        else:
            con.execute("DELETE FROM financial_ar WHERE as_of=? AND site=? AND currency=?", (as_of, site, currency))
            rows = [(as_of, site, currency, r[0], r[1], r[2], r[3], float(r[4]), nullable(r[5])) for r in frame[list(FIN_AR_FIELDS)].itertuples(index=False, name=None)]
            con.executemany("INSERT INTO financial_ar VALUES (?,?,?,?,?,?,?,?,?)", rows)
            con.execute("INSERT INTO financial_ar_batches VALUES (?,?,?,?) ON CONFLICT(as_of,site,currency) DO UPDATE SET row_count=excluded.row_count", (as_of, site, currency, len(frame)))
            start = end = as_of
        con.execute("INSERT INTO financial_log(saved_at,kind,site,currency,start_date,end_date,rows,file_name,sha256) VALUES(?,?,?,?,?,?,?,?,?)",
                    (datetime.now().isoformat(timespec="microseconds"), kind, site, currency, start, end, len(frame), file_name, hashlib.sha256(raw).hexdigest()))
    st.cache_data.clear()


def age_receivables(frame: pd.DataFrame, as_of: str) -> pd.DataFrame:
    result = frame.copy()
    overdue = (pd.Timestamp(as_of) - pd.to_datetime(result["due"])).dt.days.clip(lower=0)
    result["Días vencidos"] = overdue.astype(int)
    result["Tramo"] = pd.cut(overdue, [-1, 0, 30, 60, 90, np.inf], labels=FIN_AGING).astype(str)
    result["Corte CxC"] = as_of
    return result


def ratio(numerator: Any, denominator: Any, factor: float = 1.0) -> float:
    if pd.isna(numerator) or pd.isna(denominator) or denominator <= 0:
        return np.nan
    return float(numerator / denominator * factor)


def complete_sum(values: pd.Series) -> float:
    return float(values.sum()) if len(values) and values.notna().all() else np.nan


def performance_ratios(row: dict[str, Any], days: int) -> dict[str, Any]:
    """Ratios del período. No confunde venta, saldo de cartera y cobro."""
    net, cogs = row["Ventas netas"], row["Costo de ventas"]
    complete = bool(row["Ventas completas"])
    credit = row["Ventas crédito"] if complete else np.nan
    cogs = cogs if complete else np.nan
    row["Margen bruto"] = net - cogs if complete else np.nan
    row["Margen bruto %"] = ratio(row["Margen bruto"], net, 100)
    row["Rotación período"] = ratio(cogs, row["Base inventario"]) if pd.notna(cogs) and cogs > 0 else np.nan
    row["Días inventario"] = ratio(row["Base inventario"], cogs, days)
    row["Días CxC"] = ratio(row["Base CxC"], credit, days)
    row["Ciclo operativo"] = row["Días inventario"] + row["Días CxC"]
    row["Cartera vencida %"] = ratio(row["CxC vencida"], row["CxC"], 100)
    if row["CxC"] == 0:
        row["Cartera vencida %"] = 0.0
    row["Capital en inventario y CxC"] = row["Inventario al cierre"] + row["CxC al cierre"]
    return row


@st.cache_data(show_spinner=False, max_entries=24)
def build_financial_report(fin_version: str, inv_version: str, settings_token: tuple,
                           start: str, end: str, currency: str, sites: tuple[str, ...]) -> dict[str, Any]:
    days = (date.fromisoformat(end) - date.fromisoformat(start)).days + 1
    if days < 1 or not sites:
        raise ValueError("Selecciona sedes y un período válido.")
    settings = dict(settings_token)
    opening = (date.fromisoformat(start) - timedelta(days=1)).isoformat()
    result, details = [], []
    with db() as con:
        sales = pd.read_sql_query("SELECT * FROM financial_sales WHERE currency=? AND day BETWEEN ? AND ? ORDER BY day", con, params=(currency, start, end))
        sales = sales[sales["site"].isin(sites)].copy()
        for site in sites:
            warnings = []
            sd = sales[sales["site"] == site]
            complete = sd["day"].nunique() == days
            if not complete:
                warnings.append(f"Ventas: {sd['day'].nunique()}/{days} días cargados")
            inv = pd.read_sql_query("""SELECT snapshot_date AS fecha, SUM(existence*cost) AS valor,
                SUM(CASE WHEN existence>0 AND cost<=0 THEN 1 ELSE 0 END) AS sin_costo,
                SUM(CASE WHEN existence<0 OR cost<0 THEN 1 ELSE 0 END) AS negativos
                FROM inventory_snapshots WHERE site=? AND snapshot_date<=? GROUP BY snapshot_date ORDER BY snapshot_date""", con, params=(site, end))
            inv_date = str(inv.iloc[-1]["fecha"]) if not inv.empty else ""
            matching = settings.get(site, "") == currency
            if not matching:
                warnings.append("Moneda de inventario sin confirmar o diferente")
            inv_value = float(inv.iloc[-1]["valor"]) if not inv.empty and matching else np.nan
            inv_clean = not inv.empty and int(inv.iloc[-1]["sin_costo"] + inv.iloc[-1]["negativos"]) == 0
            if not inv.empty and not inv_clean:
                warnings.append("Inventario con costos faltantes o valores negativos")
            if inv_date != end:
                warnings.append(f"Inventario: corte {inv_date or 'no cargado'}")
            inv_close = inv_value if inv_date == end and inv_clean else np.nan
            inv_open = inv[inv["fecha"] == opening]
            inv_base, inv_method = inv_close, "Saldo final · aproximación"
            if not inv_open.empty and matching and int(inv_open.iloc[0]["sin_costo"] + inv_open.iloc[0]["negativos"]) == 0:
                inv_base = (float(inv_open.iloc[0]["valor"]) + inv_close) / 2
                inv_method = "Promedio inicial/final"
            if pd.isna(inv_base):
                inv_method = "No disponible"
            batch = con.execute("SELECT MAX(as_of) FROM financial_ar_batches WHERE site=? AND currency=? AND as_of<=?", (site, currency, end)).fetchone()[0]
            ar_value = ar_over = ar_base = ar_close = np.nan
            ar_method = "No disponible"
            if batch:
                ar = pd.read_sql_query("SELECT * FROM financial_ar WHERE site=? AND currency=? AND as_of=?", con, params=(site, currency, batch))
                aged = age_receivables(ar, batch)
                details.append(aged)
                ar_value = float(aged["balance"].sum())
                ar_over = float(aged.loc[aged["Días vencidos"] > 0, "balance"].sum())
                if batch == end:
                    ar_close = ar_value
                    ar_base, ar_method = ar_value, "Saldo final · aproximación"
                    opening_batch = con.execute("SELECT 1 FROM financial_ar_batches WHERE site=? AND currency=? AND as_of=?", (site, currency, opening)).fetchone()
                    if opening_batch:
                        ar_initial = con.execute("SELECT COALESCE(SUM(balance),0) FROM financial_ar WHERE site=? AND currency=? AND as_of=?", (site, currency, opening)).fetchone()[0]
                        ar_base, ar_method = (float(ar_initial) + ar_value) / 2, "Promedio inicial/final"
            if batch != end:
                warnings.append(f"CxC: corte {batch or 'no cargado'}")
            cogs = complete_sum(sd["cogs"])
            if pd.isna(cogs):
                warnings.append("Falta costo real de ventas")
            row = {"Sede": site, "Moneda": currency, "Días cargados": sd["day"].nunique(), "Ventas completas": complete,
                   "Ventas netas": complete_sum(sd["net_sales"]), "Ventas crédito": complete_sum(sd["credit_sales"]),
                   "Costo de ventas": cogs, "Cobros registrados": complete_sum(sd["collections"]),
                   "Inventario": inv_value, "Inventario al cierre": inv_close, "Base inventario": inv_base,
                   "CxC": ar_value, "CxC al cierre": ar_close, "Base CxC": ar_base, "CxC vencida": ar_over,
                   "Corte inventario": inv_date or "Sin datos", "Corte CxC": batch or "Sin datos",
                   "Método inventario": inv_method, "Método CxC": ar_method, "Calidad de datos": " · ".join(warnings) or "Completa"}
            result.append(performance_ratios(row, days))
    summary = pd.DataFrame(result)
    total = {"Ventas completas": bool(summary["Ventas completas"].all())}
    for col in ["Ventas netas", "Ventas crédito", "Costo de ventas", "Cobros registrados", "Inventario", "Inventario al cierre", "Base inventario", "CxC", "CxC al cierre", "Base CxC", "CxC vencida"]:
        total[col] = complete_sum(summary[col])
    total = performance_ratios(total, days)
    ar_details = pd.concat(details, ignore_index=True) if details else pd.DataFrame()
    return {"summary": summary, "total": total, "sales": sales, "ar": ar_details, "days": days, "start": start, "end": end, "currency": currency}


def fin_value(value: Any, currency: str = "", suffix: str = "") -> str:
    if pd.isna(value) or not np.isfinite(float(value)):
        return "N/D"
    return f"{currency + ' ' if currency else ''}{float(value):,.2f}{suffix}"


def suggest_finance_action(row: pd.Series, max_dio: int, max_dso: int, max_late: int) -> tuple[str, str]:
    if row["Calidad de datos"] != "Completa":
        return "Completar datos", row["Calidad de datos"]
    if row["Cartera vencida %"] > max_late or row["Días CxC"] > max_dso:
        return "Priorizar cobranza", "Revisar facturas vencidas y acuerdos de pago antes de ampliar crédito."
    if row["Días inventario"] > max_dio:
        return "Revisar existencias", "Validar productos lentos, redistribución y compras pendientes."
    if pd.notna(row["Margen bruto"]) and row["Margen bruto"] < 0:
        return "Revisar margen", "Contrastar costos reales, devoluciones y descuentos."
    if pd.isna(row["Días inventario"]) or pd.isna(row["Días CxC"]):
        return "Revisar actividad", "No hay base positiva suficiente para calcular todos los ratios."
    return "Dentro de metas", "Mantener seguimiento del inventario y la cartera."


def save_action_task(title: str, site: str, priority: str = "Media", source_key: str | None = None,
                     owner: str = "", due: str | None = None, notes: str = "") -> bool:
    if not title.strip():
        raise ValueError("Escribe una acción concreta.")
    stamp = datetime.now().isoformat(timespec="microseconds")
    with db() as con:
        cursor = con.execute("""INSERT OR IGNORE INTO action_tasks
            (id,source_key,title,site,owner,due,status,priority,notes,created_at,updated_at)
            VALUES (?,?,?,?,?,?,'Pendiente',?,?,?,?)""",
            (uuid.uuid4().hex[:12], source_key, title.strip(), site, owner, due, priority, notes, stamp, stamp))
        return cursor.rowcount == 1


def safe_export_frame(frame: pd.DataFrame) -> pd.DataFrame:
    result = frame.copy()
    for col in result.select_dtypes(include=["object", "string", "category"]).columns:
        result[col] = result[col].astype(object).map(lambda v: "'" + v if isinstance(v, str) and v.lstrip().startswith(("=", "+", "-", "@")) else v)
    return result


@st.cache_data(show_spinner=False, max_entries=6)
def finance_export(report: dict) -> bytes:
    """Paquete liviano CSV + HTML; sólo se genera cuando se solicita."""
    summary = report["summary"]
    title = f"NEXUS · Desempeño · {report['start']} a {report['end']} · {report['currency']}"
    body = f"<!doctype html><html lang='es'><meta charset='utf-8'><title>{escape(title)}</title><style>body{{font:14px Arial;margin:32px;color:#152b37}}table{{border-collapse:collapse;white-space:nowrap;font-size:12px}}td,th{{padding:9px;border:1px solid #dee7e9;text-align:left}}th{{background:#e5f3f1}}.scroll{{overflow:auto}}</style><h1>{escape(title)}</h1><p>Ventas y cobros son flujos del período; inventario y CxC son saldos a sus cortes. N/D indica datos faltantes o denominador no positivo.</p><div class='scroll'>{summary.to_html(index=False,escape=True,na_rep='N/D',float_format=lambda x: f'{x:,.2f}')}</div><p>Los días de inventario y CxC usan promedios inicial/final cuando existen ambos cortes; de lo contrario, el saldo final como aproximación. El ciclo operativo no descuenta plazos de pago a proveedores. Los cobros registrados pueden corresponder a ventas de otros períodos.</p></html>"
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("Desempeno.html", body.encode("utf-8"))
        for name, frame in [("Indicadores", summary), ("Ventas", report["sales"]), ("Cartera", report["ar"])]:
            z.writestr(f"{name}.csv", safe_export_frame(frame).to_csv(index=False).encode("utf-8-sig"))
        z.writestr("Metodologia.txt", FIN_METHODOLOGY.encode("utf-8"))
    return out.getvalue()


FIN_METHODOLOGY = """CÓMO LEER EL ANÁLISIS
• Ventas netas: ventas sin IVA, netas de devoluciones y descuentos. Una fila por día/sede.
• Ventas crédito: importe a crédito del día, con la MISMA base de impuestos que los saldos de CxC. Si CxC incluye IVA, este campo también. No se resta de ventas netas para inferir contado.
• Costo de ventas: costo real de la mercancía vendida. No se sustituye por costo unitario actual × unidades.
• Cobros: dinero efectivamente recibido según el reporte del período; puede incluir facturas antiguas. Vacío significa no informado.
• CxC: fotografía COMPLETA de facturas y saldos pendientes al cierre de una fecha. Nunca se suman cortes de diferentes fechas. Un corte vacío confirmado equivale a cartera cero; ausencia de archivo NO equivale a cero.
• Antigüedad: días calendario después del vencimiento al corte de cartera. Una factura que vence en la fecha del corte está al día. Se clasifica por saldo pendiente.
• Margen bruto = ventas netas − costo de ventas. No es utilidad neta: no descuenta gastos operativos ni impuestos sobre la renta.
• Base inventario = (inventario anterior al inicio + inventario al cierre) / 2. Base CxC análoga. Sin saldo inicial exacto se usa el final, etiquetado como aproximación.
• Rotación del período = costo de ventas / base inventario. No se anualiza.
• Días inventario = base inventario / costo de ventas × días del período.
• Días CxC (DSO) = base CxC / ventas crédito × días del período. Es una estimación, no la media de días reales por factura.
• Ciclo operativo = días inventario + días CxC. NO es ciclo de caja: falta descontar días de cuentas por pagar.
• Capital en inventario y CxC = ambos saldos válidos al mismo cierre. No es capital de trabajo neto ni dinero disponible.
• Los ratios requieren ventas completas de cada día calendario y cortes válidos en el cierre. Los días sin actividad deben cargarse como cero. Si el informe omite días cerrados, se pueden completar con cero sólo tras confirmarlo en la carga.
• Si hay costos faltantes, stock negativo, moneda no confirmada, períodos incompletos o cortes antiguos, se muestran advertencias y se bloquean los ratios afectados. Los totales con sedes faltantes muestran N/D. Nunca se promedian porcentajes de sedes.
• Cada moneda se analiza por separado. No se aplican tipos de cambio automáticamente. La moneda de inventario se declara por sede y debe ser consistente en su historial.
• Los límites del semáforo son metas internas configurables, no estándares universales.
Referencia de fórmulas: ACCA, Working capital management:
https://www.accaglobal.com/uk/en/student/exam-support-resources/fundamentals-exams-study-resources/f9/technical-articles/wcm.html

INICIO RÁPIDO
1. Guarda el script junto a tu nexus_data.db anterior; cierra la app antigua y ejecuta la nueva.
2. Sigue cargando inventario y ventas EN UNIDADES en Cortes Diarios.
3. En Datos Financieros, declara la moneda en que están los costos del inventario de cada sede.
4. Carga el resumen diario de ventas EN DINERO de cada sede y moneda. No sirve un acumulado repetido diariamente.
5. Carga la cartera completa de cada sede y moneda, con la fecha real de cierre. Un nuevo corte conserva los anteriores; recargar la misma fecha reemplaza esa fotografía.
6. Abre Desempeño Integral y elige período, sedes y moneda. Consulta Cuentas por Cobrar para priorizar la cartera vencida.
7. Usa Gestión Visual para crear tareas, asignar responsables y actualizar estados. No envía mensajes automáticamente.
8. Descarga un respaldo desde el menú lateral cuando necesites copiar toda la base de datos.
"""


def render_upload_mapping(frame: pd.DataFrame, kind: str, token: str) -> dict[str, str]:
    schema = FIN_SALES_FIELDS if kind == "ventas" else FIN_AR_FIELDS
    optional = {"Costo de ventas", "Cobros", "Importe factura"}
    mapping = {}
    with st.expander("Relacionar las columnas de mi archivo", expanded=True):
        columns = st.columns(2)
        for index, (target, aliases) in enumerate(schema.items()):
            choices = ["Sin asignar"] + list(frame.columns)
            match = next((c for c in frame if normalize_key(c) in aliases), None)
            with columns[index % 2]:
                chosen = st.selectbox(target + (" · opcional" if target in optional else " *"), choices,
                                      index=choices.index(match) if match else 0, key=f"map_{token}_{target}")
            mapping[target] = "" if chosen == "Sin asignar" else chosen
    return mapping


def render_financial_import(kind: str) -> None:
    is_sales = kind == "ventas"
    schema = FIN_SALES_FIELDS if is_sales else FIN_AR_FIELDS
    st.info("Carga un resumen diario: una fila por fecha, con importes totales de la sede. Los acumulados deben convertirse en movimientos diarios." if is_sales else
            "Carga la cartera completa de una sede y moneda al cierre elegido. Una fila por factura, incluyendo serie. Se conserva el historial de cortes.")
    st.download_button("Descargar plantilla CSV vacía", pd.DataFrame(columns=list(schema)).to_csv(index=False).encode("utf-8-sig"),
                       f"Plantilla_{kind}_NEXUS.csv", "text/csv", key=f"template_{kind}")
    cols = st.columns(3)
    site = cols[0].selectbox("Sede", SEDES, key=f"import_site_{kind}")
    currency = cols[1].selectbox("Moneda del archivo", FIN_CURRENCIES, key=f"import_currency_{kind}")
    cutoff = cols[2].date_input("Fecha de corte de cartera", date.today(), max_value=date.today(), key="ar_cutoff") if not is_sales else None
    if is_sales:
        cols[2].caption("Ventas netas sin IVA. Ventas crédito con la misma base de impuestos que CxC.")
    no_ar = not is_sales and st.checkbox("Confirmo que esta sede tiene CERO saldo de CxC en esta moneda y fecha", key=f"zero_ar_{site}_{currency}_{cutoff}")
    if no_ar:
        st.warning("Guardar este corte reemplazará la cartera de esta sede, moneda y fecha por cero.")
        if st.button("Guardar corte de cartera cero", key="save_zero_ar"):
            save_financial_data(pd.DataFrame(columns=FIN_AR_FIELDS), "cxc", site, currency, "Cartera cero confirmada", b"zero-confirmed", cutoff.isoformat())
            st.success("Corte de cartera cero guardado.")
        return
    upload = st.file_uploader("Archivo Excel o CSV", type=["xlsx", "xls", "csv"], key=f"upload_fin_{kind}")
    if not upload:
        return
    raw = upload.getvalue()
    token = f"{kind}_{hashlib.sha256(raw).hexdigest()[:10]}"
    with st.expander("Opciones de lectura", expanded=False):
        opt = st.columns(2)
        header = int(opt[0].number_input("Fila de encabezados", 1, 200, 1, key=f"header_{token}")) - 1
        decimal_label = opt[1].selectbox("Formato de importes en texto", ["Punto decimal · 1,234.56", "Coma decimal · 1.234,56"], key=f"decimal_{token}")
        decimal = "," if decimal_label.startswith("Coma") else "."
        day_first = st.selectbox("Fechas en texto", ["Día / mes / año", "Mes / día / año"], key=f"dates_{token}").startswith("Día")
        delimiter = st.selectbox("Separador CSV", [",", ";", "\t"], format_func=lambda x: {",": "Coma", ";": "Punto y coma", "\t": "Tabulación"}[x], key=f"delim_{token}") if upload.name.lower().endswith(".csv") else ","
        try:
            sheet = st.selectbox("Hoja del libro", financial_sheets(raw), key=f"sheet_{token}") if not upload.name.lower().endswith(".csv") else ""
        except Exception as exc:
            st.error(f"No se pudo leer el libro: {exc}")
            return
    try:
        frame = read_financial_file(raw, upload.name, sheet, header, delimiter)
    except Exception as exc:
        st.error(f"No se pudo leer el archivo: {exc}")
        return
    if frame.empty:
        st.warning("El archivo no contiene filas de datos.")
        return
    st.caption(f"{len(frame):,} filas leídas. Vista previa de las primeras 20.")
    st.dataframe(frame.head(20), hide_index=True, width="stretch")
    mapping = render_upload_mapping(frame, kind, token + f"_{header}_{sheet}")
    try:
        prepared = normalize_financial_frame(frame, kind, mapping, site, currency, decimal, day_first, cutoff.isoformat() if cutoff else None)
    except ValueError as exc:
        st.warning(str(exc))
        return
    if is_sales:
        lo, hi = prepared["Fecha"].min(), prepared["Fecha"].max()
        missing = sorted(set(pd.date_range(lo, hi).strftime("%Y-%m-%d")) - set(prepared["Fecha"]))
        if missing:
            st.warning(f"Faltan {len(missing)} días entre {lo} y {hi}. Permanecerán sin datos salvo que confirmes que no hubo actividad.")
            if st.checkbox("Confirmo que TODOS los días omitidos dentro de ese intervalo tuvieron actividad cero", key=f"fill_{token}"):
                zeros = pd.DataFrame({"Fecha": missing, "Ventas netas": 0.0, "Ventas crédito": 0.0,
                                      "Costo de ventas": 0.0 if prepared["Costo de ventas"].notna().all() else np.nan,
                                      "Cobros": 0.0 if prepared["Cobros"].notna().all() else np.nan})
                prepared = pd.concat([prepared, zeros], ignore_index=True).sort_values("Fecha").reset_index(drop=True)
        if prepared[["Ventas netas", "Ventas crédito", "Costo de ventas", "Cobros"]].lt(0).any().any():
            st.info("Hay importes negativos: se conservarán como devoluciones o ajustes netos. Revisa que sean correctos.")
        st.caption("Recargar una fecha actualiza su resumen; no lo suma. Los días ajenos al archivo permanecen guardados.")
    else:
        st.caption("Este archivo sustituirá TODAS las facturas del mismo corte, sede y moneda. No se suma a cortes anteriores.")
    st.markdown("**Datos preparados**")
    st.dataframe(prepared.head(20), hide_index=True, width="stretch")
    review_token = hashlib.sha256(repr((site, currency, cutoff, decimal, day_first, mapping, len(prepared))).encode()).hexdigest()[:12]
    confirm = st.checkbox("Verifiqué sede, moneda, fechas e importes; el archivo contiene el reporte completo indicado", key=f"confirm_{token}_{review_token}")
    if st.button("Guardar datos validados", type="primary", key=f"save_{token}", disabled=not confirm):
        try:
            save_financial_data(prepared, kind, site, currency, upload.name, raw, cutoff.isoformat() if cutoff else None)
            st.success(f"Guardadas {len(prepared):,} filas para {site} en {currency}.")
        except Exception as exc:
            st.error(f"No se guardaron los datos: {exc}")


def render_financial_data() -> None:
    st.subheader("Datos financieros")
    st.caption("Conecta los reportes de administración con el inventario que ya gestionas.")
    sales_tab, ar_tab, currency_tab, guide_tab = st.tabs(["Ventas en dinero", "Cartera CxC", "Moneda del inventario", "Guía e historial"])
    with sales_tab:
        render_financial_import("ventas")
    with ar_tab:
        render_financial_import("cxc")
    with currency_tab:
        st.info("Indica en qué moneda están expresados los costos de tus cortes de inventario. Esta declaración aplica al historial de la sede; no convierte importes ni cambia tus costos.")
        settings = finance_settings()
        with st.form("currency_settings"):
            values = {}
            for site in SEDES:
                options = ["Sin confirmar"] + FIN_CURRENCIES
                current = settings.get(site) or "Sin confirmar"
                values[site] = st.selectbox(site, options, index=options.index(current), key=f"inv_currency_{site}")
            if st.form_submit_button("Guardar monedas del inventario"):
                with db() as con:
                    con.executemany("INSERT INTO financial_settings VALUES (?,?) ON CONFLICT(site) DO UPDATE SET inventory_currency=excluded.inventory_currency",
                                    [(s, "" if c == "Sin confirmar" else c) for s, c in values.items()])
                st.cache_data.clear()
                st.success("Monedas actualizadas.")
    with guide_tab:
        st.markdown(FIN_METHODOLOGY)
        with db() as con:
            log = pd.read_sql_query("SELECT saved_at AS Guardado,kind AS Tipo,site AS Sede,currency AS Moneda,start_date AS Desde,end_date AS Hasta,rows AS Filas,file_name AS Archivo FROM financial_log ORDER BY id DESC LIMIT 200", con)
        st.dataframe(log, hide_index=True, width="stretch")


def demo_financial_report(start: str, end: str, currency: str, sites: tuple[str, ...]) -> dict:
    days = (date.fromisoformat(end) - date.fromisoformat(start)).days + 1
    rows, daily, ar_rows = [], [], []
    for index, site in enumerate(sites):
        net = (1500 + index * 350) * days
        cogs, credit = net * .68, net * .72
        inv = cogs / days * (32 + index * 19)
        ar = credit / days * (19 + index * 9)
        late = ar * (.06 + index * .08)
        row = {"Sede": site, "Moneda": currency, "Días cargados": days, "Ventas completas": True,
               "Ventas netas": net, "Ventas crédito": credit, "Costo de ventas": cogs, "Cobros registrados": net * .82,
               "Inventario": inv, "Inventario al cierre": inv, "Base inventario": inv,
               "CxC": ar, "CxC al cierre": ar, "Base CxC": ar, "CxC vencida": late,
               "Corte inventario": end, "Corte CxC": end, "Método inventario": "Promedio inicial/final",
               "Método CxC": "Promedio inicial/final", "Calidad de datos": "Completa"}
        rows.append(performance_ratios(row, days))
        for day in pd.date_range(start, end):
            daily.append({"day": day.date().isoformat(), "site": site, "currency": currency, "net_sales": net / days,
                          "credit_sales": credit / days, "cogs": cogs / days, "collections": net * .82 / days})
        for overdue, balance in [(0, ar - late), (45, late)]:
            ar_rows.append({"as_of": end, "site": site, "currency": currency, "invoice": f"DEMO-{index}-{overdue}",
                            "client": f"Cliente de ejemplo {index + 1}", "issued": (pd.Timestamp(end) - pd.Timedelta(days=90)).date().isoformat(),
                            "due": (pd.Timestamp(end) - pd.Timedelta(days=overdue)).date().isoformat(), "balance": balance, "amount": balance})
    summary = pd.DataFrame(rows)
    total = {"Ventas completas": True}
    for col in ["Ventas netas", "Ventas crédito", "Costo de ventas", "Cobros registrados", "Inventario", "Inventario al cierre", "Base inventario", "CxC", "CxC al cierre", "Base CxC", "CxC vencida"]:
        total[col] = complete_sum(summary[col])
    return {"summary": summary, "total": performance_ratios(total, days), "sales": pd.DataFrame(daily),
            "ar": age_receivables(pd.DataFrame(ar_rows), end), "days": days, "start": start, "end": end, "currency": currency}


def financial_filters(prefix: str):
    result = financial_filters_v7(prefix)
    return result[:4] if result is not None else None


def render_receivables(report: dict) -> None:
    ar = report["ar"].copy()
    if ar.empty:
        if report["summary"]["CxC"].notna().any():
            st.success("Los cortes de cartera disponibles tienen saldo cero.")
        else:
            empty_state("Todavía no hay cartera cargada", "Importa el reporte completo en Carga de datos → Cuentas por cobrar.")
        return
    st.caption("La antigüedad se calcula a la fecha del corte de cada sede. Un corte antiguo conserva su saldo original; pagos posteriores requieren cargar un nuevo corte.")
    late_only = st.toggle("Mostrar sólo facturas vencidas", key="ar_late_only")
    query = st.text_input("Buscar cliente o factura", key="ar_search").strip()
    if late_only:
        ar = ar[ar["Días vencidos"] > 0]
    if query:
        ar = ar[ar["client"].str.contains(query, case=False, regex=False) | ar["invoice"].str.contains(query, case=False, regex=False)]
    ar = ar[ar["balance"] > 0].copy()
    if ar.empty:
        st.info("No hay facturas pendientes para estos filtros.")
        return
    a, b = st.columns(2)
    aging = ar.groupby("Tramo")["balance"].sum().reindex(FIN_AGING, fill_value=0).reset_index()
    with a:
        fig = px.bar(aging, x="Tramo", y="balance", color="Tramo", color_discrete_sequence=[NEXUS_GREEN, NEXUS_BLUE, NEXUS_ORANGE, NEXUS_RED, "#8D273B"], labels={"balance": report["currency"]}, title="Antigüedad del saldo pendiente")
        fig.update_layout(showlegend=False)
        st.plotly_chart(chart_layout(fig, 340), width="stretch")
    with b:
        clients = ar.groupby(["site", "client"], as_index=False)["balance"].sum().nlargest(12, "balance")
        clients["Cliente / sede"] = clients["client"] + " · " + clients["site"]
        fig = px.bar(clients.sort_values("balance"), y="Cliente / sede", x="balance", orientation="h", color_discrete_sequence=[NEXUS_BLUE], labels={"balance": report["currency"]}, title="Mayor saldo por cliente y sede")
        st.plotly_chart(chart_layout(fig, 340), width="stretch")
    renamed = ar.rename(columns={"site": "Sede", "client": "Cliente", "invoice": "Factura", "issued": "Emisión", "due": "Vencimiento", "balance": "Saldo", "amount": "Importe factura", "currency": "Moneda"})
    columns = ["Sede", "Cliente", "Factura", "Emisión", "Vencimiento", "Días vencidos", "Tramo", "Saldo", "Moneda", "Corte CxC"]
    renamed = renamed.sort_values(["Días vencidos", "Saldo"], ascending=False)
    st.dataframe(renamed[columns], width="stretch", hide_index=True, height=450,
                 column_config={"Saldo": st.column_config.NumberColumn(format="%.2f"), "Días vencidos": st.column_config.NumberColumn(format="%d días")})
    st.download_button("Descargar cartera filtrada", safe_export_frame(renamed[columns]).to_csv(index=False).encode("utf-8-sig"),
                       f"CxC_{report['end']}_{report['currency']}.csv", "text/csv", key="ar_csv")


def render_performance(receivables_only: bool = False) -> None:
    st.subheader("Cuentas por cobrar" if receivables_only else "Desempeño integral")
    st.caption("Inventario, ventas y cartera: una lectura conjunta para decidir dónde actuar.")
    params = financial_filters("performance")
    if not params:
        return
    start, end, currency, sites = params
    demo = st.toggle("Ver una demostración con datos ficticios", key="finance_demo")
    if demo:
        st.warning("DEMOSTRACIÓN · Datos ficticios. No se guardan en tu base de datos.")
        report = demo_financial_report(start, end, currency, sites)
    else:
        report = build_financial_report(finance_version(), db_version(), tuple(sorted(finance_settings().items())), start, end, currency, sites)
    summary, total = report["summary"].copy(), report["total"]
    partial = not total["Ventas completas"]
    with st.expander("Fuentes, fechas y calidad de datos", expanded=(summary["Calidad de datos"] != "Completa").any()):
        st.dataframe(summary[["Sede", "Corte inventario", "Corte CxC", "Días cargados", "Método inventario", "Método CxC", "Calidad de datos"]], width="stretch", hide_index=True)
        st.caption(f"Período de {report['days']} días. N/D = no disponible. Los importes de cortes antiguos se muestran con su fecha; no activan los ratios del cierre solicitado.")
        if st.button("Cargar o corregir datos financieros", key="go_fin_upload"):
            st.session_state["nexus_view"] = "Datos Financieros"
            st.rerun()
    cols = st.columns(4)
    for col, label, field, note, tone in [
        (cols[0], "Inventario", "Inventario", "A costo · ver fecha por sede", "blue"),
        (cols[1], "Ventas registradas", "Ventas netas", "Período incompleto" if partial else f"{report['days']} días · sin IVA", "green"),
        (cols[2], "Cuentas por cobrar", "CxC", "Saldo del último corte disponible", "cyan"),
        (cols[3], "Cartera vencida", "CxC vencida", fin_value(total["Cartera vencida %"], suffix="%") + " del saldo", "red")]:
        with col:
            kpi(label, fin_value(total[field], currency), note, tone)
    if receivables_only:
        render_receivables(report)
        return
    cols = st.columns(4)
    for col, label, field, suffix, tone in [
        (cols[0], "Días de inventario", "Días inventario", " días", "blue"),
        (cols[1], "Días de CxC", "Días CxC", " días", "orange"),
        (cols[2], "Ciclo operativo", "Ciclo operativo", " días", "cyan"),
        (cols[3], "Margen bruto", "Margen bruto %", "%", "green")]:
        with col:
            kpi(label, fin_value(total[field], suffix=suffix), "Promedio o aproximación según las fuentes", tone)
    st.caption("El ciclo operativo suma inventario y CxC. Para calcular el ciclo de caja completo faltan las cuentas por pagar a proveedores.")
    overview, seats_tab, collections_tab, method_tab = st.tabs(["Visión del negocio", "Comparar sedes", "Cartera y cobros", "Cómo se calcula"])
    with overview:
        a, b = st.columns(2)
        with a:
            plot = summary.melt(id_vars="Sede", value_vars=["Inventario", "CxC"], var_name="Concepto", value_name="Importe").dropna()
            if not plot.empty:
                fig = px.bar(plot, x="Sede", y="Importe", color="Concepto", barmode="group", color_discrete_map={"Inventario": NEXUS_BLUE, "CxC": NEXUS_CYAN}, title=f"Capital por sede · {currency}")
                st.plotly_chart(chart_layout(fig, 360), width="stretch")
            else:
                empty_state("Capital pendiente de completar", "Carga CxC y declara la moneda de tus inventarios.")
        with b:
            trend = report["sales"].copy()
            if not trend.empty:
                grouped = trend.groupby("day").agg(Ventas=("net_sales", "sum"), Cobros=("collections", lambda v: v.sum() if v.notna().all() else np.nan), Sedes=("site", "nunique"))
                grouped = grouped.reindex(pd.date_range(start, end).strftime("%Y-%m-%d"))
                grouped.loc[grouped["Sedes"] != len(sites), ["Ventas", "Cobros"]] = np.nan
                grouped.index.name = "Fecha"
                fig = px.line(grouped.reset_index(), x="Fecha", y=["Ventas", "Cobros"], color_discrete_sequence=[NEXUS_BLUE, NEXUS_GREEN], title="Ventas netas y cobros registrados")
                fig.update_traces(connectgaps=False)
                st.plotly_chart(chart_layout(fig, 360), width="stretch")
                st.caption("Cobros puede incluir ventas de meses anteriores y otra base de impuestos. Su diferencia con ventas no representa la deuda pendiente.")
            else:
                empty_state("Falta el reporte de ventas", "Importa ventas en dinero para ver la evolución.")
        c1, c2, c3 = st.columns(3)
        with c1:
            kpi("Margen bruto", fin_value(total["Margen bruto"], currency), "Antes de gastos operativos e impuestos", "green")
        with c2:
            kpi("Rotación del período", fin_value(total["Rotación período"], suffix=" veces"), "Costo de ventas / inventario base", "blue")
        with c3:
            kpi("Inventario + CxC al cierre", fin_value(total["Capital en inventario y CxC"], currency), "Requiere ambos saldos al mismo cierre", "cyan")
    with seats_tab:
        st.markdown("**Metas internas para el seguimiento**")
        caps = st.columns(3)
        max_dio = caps[0].number_input("Máximo de días de inventario", 1, 730, 60, key="target_dio")
        max_dso = caps[1].number_input("Máximo de días de CxC", 1, 730, 45, key="target_dso")
        max_late = caps[2].number_input("Máximo de cartera vencida (%)", 0, 100, 10, key="target_overdue")
        decisions = summary.apply(lambda r: suggest_finance_action(r, max_dio, max_dso, max_late), axis=1)
        summary["Estado"] = [d[0] for d in decisions]
        summary["Acción sugerida"] = [d[1] for d in decisions]
        scatter = summary.dropna(subset=["Días inventario", "Días CxC", "Inventario al cierre"]).copy()
        if not scatter.empty:
            scatter["Tamaño"] = scatter["Inventario al cierre"].clip(lower=1)
            fig = px.scatter(scatter, x="Días inventario", y="Días CxC", size="Tamaño", color="Estado", hover_name="Sede", size_max=52, color_discrete_sequence=NEXUS_PALETTE, title="Inventario lento y cobranza lenta requieren atención conjunta")
            fig.add_vline(x=max_dio, line_dash="dot", line_color=NEXUS_GRAY)
            fig.add_hline(y=max_dso, line_dash="dot", line_color=NEXUS_GRAY)
            st.plotly_chart(chart_layout(fig, 420), width="stretch")
        cols = ["Sede", "Estado", "Ventas netas", "Inventario", "CxC", "Cartera vencida %", "Días inventario", "Días CxC", "Margen bruto %", "Acción sugerida"]
        st.dataframe(summary[cols], hide_index=True, width="stretch")
        if st.button("Crear tareas de seguimiento para estas sedes", disabled=demo, key="finance_tasks"):
            created = 0
            for _, row in summary.iterrows():
                if row["Estado"] == "Dentro de metas":
                    continue
                created += save_action_task(f"{row['Estado']} · {row['Sede']}", row["Sede"], "Alta",
                                            f"fin:{row['Sede']}:{currency}:{start}:{end}:{row['Estado']}", notes=row["Acción sugerida"])
            st.success(f"{created} tareas nuevas en Gestión Visual; se evitan duplicados del mismo período.")
    with collections_tab:
        kpi("Cobros registrados", fin_value(total["Cobros registrados"], currency), "Movimientos reales reportados en las fechas seleccionadas", "green")
        render_receivables(report)
    with method_tab:
        st.markdown(FIN_METHODOLOGY)
    if not demo:
        report["summary"] = summary
        if st.button("Preparar informe de desempeño", key="prepare_fin_report"):
            st.session_state["finance_export_payload"] = (hashlib.sha256(repr((finance_version(), db_version(), start, end, currency, sites, tuple(finance_settings().items()), max_dio, max_dso, max_late)).encode()).hexdigest(), finance_export(report))
        expected = hashlib.sha256(repr((finance_version(), db_version(), start, end, currency, sites, tuple(finance_settings().items()), max_dio, max_dso, max_late)).encode()).hexdigest()
        payload = st.session_state.get("finance_export_payload")
        if payload and payload[0] == expected:
            st.download_button("Descargar informe HTML + CSV", payload[1], f"NEXUS_Desempeno_{end}_{currency}.zip", "application/zip", key="download_fin_report")


def render_action_board(all_df: pd.DataFrame) -> None:
    st.subheader("Gestión visual")
    st.caption("Transforma alertas en tareas, asigna responsables y registra lo que ya se resolvió.")
    with st.expander("Crear una tarea", expanded=False):
        with st.form("new_action_task", clear_on_submit=True):
            title = st.text_input("Acción a realizar")
            a, b, c = st.columns(3)
            site = a.selectbox("Sede de la tarea", ["General"] + SEDES)
            owner = b.text_input("Responsable")
            priority = c.selectbox("Prioridad", ["Alta", "Media", "Baja"])
            due = st.date_input("Fecha objetivo", date.today())
            notes = st.text_area("Notas")
            if st.form_submit_button("Crear tarea"):
                if title.strip():
                    save_action_task(title, site, priority, owner=owner, due=due.isoformat(), notes=notes)
                    st.success("Tarea creada.")
                else:
                    st.warning("Escribe la acción antes de guardar.")
    if not all_df.empty and st.button("Crear tareas de los 20 mayores problemas de inventario", key="tasks_from_inventory"):
        selected = all_df[all_df["Estado"] != "ÓPTIMO"].copy()
        selected["_negative"] = (selected["Existencia"] < 0).astype(int)
        selected = selected.sort_values(["_negative", "Capital Inmovilizado ($)"], ascending=False).head(20)
        count = 0
        for _, row in selected.iterrows():
            site = str(row["Sede"])
            cut = latest_snapshot_date(site) or ""
            count += save_action_task(f"{row['Acción']} · {row['Código']} · {row['Descripción']}", site, "Alta",
                                      f"inv:{site}:{cut}:{row['Código']}:{row['Acción']}", notes=f"Existencia: {row['Existencia']} · Estado: {row['Estado']}")
        st.success(f"{count} tareas nuevas; no se duplican las del mismo corte.")
    with db() as con:
        tasks = pd.read_sql_query("SELECT id AS ID,title AS Tarea,site AS Sede,owner AS Responsable,due AS Fecha,status AS Estado,priority AS Prioridad,notes AS Notas FROM action_tasks ORDER BY created_at DESC", con)
    if tasks.empty:
        empty_state("Tu plan de acción empieza aquí", "Crea una tarea o genera tareas desde inventario y Desempeño Integral.", "✓")
        return
    a, b = st.columns(2)
    selected_sites = site_selector_v712("Filtrar tareas por sede", "task_sites", options=["General"], container=a)
    include_done = b.toggle("Incluir resueltas", value=False, key="task_include_done")
    data = tasks[tasks["Sede"].isin(selected_sites)].copy()
    if not include_done:
        data = data[data["Estado"] != "Resuelta"].copy()
    status_colors = [NEXUS_BLUE, NEXUS_CYAN, NEXUS_ORANGE, NEXUS_GREEN]
    for column, state, color in zip(st.columns(4), TASK_STATES, status_colors):
        with column:
            subset = data[data["Estado"] == state]
            st.markdown(f'<div class="board-heading" style="border-color:{color}">{escape(state)} <span>{len(subset)}</span></div>', unsafe_allow_html=True)
            for _, row in subset.head(8).iterrows():
                overdue = bool(row["Fecha"] and str(row["Fecha"]) < date.today().isoformat() and state != "Resuelta")
                stamp = f"{'Vencida · ' if overdue else ''}{row['Fecha'] or 'Sin fecha'}"
                st.markdown(f'<div class="task-card"><div class="task-site">{escape(str(row["Sede"]))} · {escape(str(row["Prioridad"]))}</div><strong>{escape(str(row["Tarea"]))}</strong><div class="task-owner">{escape(str(row["Responsable"] or "Sin responsable"))}</div><small style="color:{NEXUS_RED if overdue else NEXUS_GRAY}">{escape(stamp)}</small></div>', unsafe_allow_html=True)
            if len(subset) > 8:
                st.caption(f"{len(subset) - 8} tareas más en la tabla.")
    if data.empty:
        st.info("No hay tareas para los filtros seleccionados.")
        return
    st.markdown("**Actualizar tareas**")
    st.caption("Edita estado, responsable, fecha o notas y guarda los cambios.")
    data["Fecha"] = pd.to_datetime(data["Fecha"], errors="coerce").dt.date
    with st.form("edit_action_tasks"):
        edited = st.data_editor(data, hide_index=True, width="stretch", disabled=["ID", "Tarea", "Sede"],
                                column_config={"Estado": st.column_config.SelectboxColumn(options=TASK_STATES, required=True),
                                               "Prioridad": st.column_config.SelectboxColumn(options=["Alta", "Media", "Baja"], required=True),
                                               "Fecha": st.column_config.DateColumn(format="DD/MM/YYYY")}, key="task_editor")
        if st.form_submit_button("Guardar cambios en tareas"):
            if not edited["Estado"].isin(TASK_STATES).all() or not edited["Prioridad"].isin(["Alta", "Media", "Baja"]).all():
                st.error("Hay estados o prioridades inválidos.")
                return
            with db() as con:
                for _, row in edited.iterrows():
                    con.execute("UPDATE action_tasks SET owner=?,due=?,status=?,priority=?,notes=?,updated_at=? WHERE id=?",
                                (str(row["Responsable"] or ""), None if pd.isna(row["Fecha"]) else str(row["Fecha"]), row["Estado"], row["Prioridad"], str(row["Notas"] or ""), datetime.now().isoformat(), row["ID"]))
            st.rerun()


# ============================================================
# INICIALIZACIÓN Y SIDEBAR
# ============================================================

# """Lectura de reportes impresos de A2, sin modificar los archivos originales.

# Contrato público: parse_a2_report(raw, filename, kind='auto') -> dict con
# kind, data (pandas.DataFrame), metadata, warnings y catalog. No escribe en DB.
# Los identificadores ausentes se conservan vacíos y bloquean ready_to_save.
# Los importes del período nunca se distribuyen artificialmente entre fechas.
# """

import io
import math
import re
import struct
import unicodedata
from datetime import datetime
from decimal import Decimal, InvalidOperation

import pandas as pd


def _a2_text(value):
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _a2_norm(value):
    return re.sub(r"[^a-z0-9]+", " ", unicodedata.normalize("NFKD", _a2_text(value)).encode("ascii", "ignore").decode().lower()).strip()


def _a2_number(value):
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value) if math.isfinite(float(value)) else None
    text = _a2_text(value).replace("\u00a0", "").replace(" ", "")
    if not text:
        return None
    negative = text.startswith("(") and text.endswith(")")
    text = text.strip("()")
    text = re.sub(r"^(USD|US\$|Bs\.?|VES|\$)", "", text, flags=re.I)
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".") if text.rfind(",") > text.rfind(".") else text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")
    try:
        number = float(Decimal(text))
        return (-number if negative else number) if math.isfinite(number) else None
    except (InvalidOperation, ValueError):
        return None


def _a2_sum(values):
    return float(sum((Decimal(str(v)) for v in values if v is not None), Decimal("0")))


def _a2_legacy_grid(raw):
    """Fallback limitado al flujo BIFF antiguo que exporta A2 (NUMBER/LABEL).

    No se usa para otros libros BIFF: un registro desconocido con datos o un
    archivo truncado genera error en vez de omitir contenido silenciosamente.
    """
    cells = {}
    offset = 0
    eof = False
    while offset + 4 <= len(raw):
        record, size = struct.unpack_from("<HH", raw, offset)
        payload = raw[offset + 4:offset + 4 + size]
        if len(payload) != size:
            raise ValueError("El archivo A2 está incompleto: vuelve a exportarlo.")
        if record in (3, 4):
            if len(payload) < (15 if record == 3 else 8):
                raise ValueError("Una celda A2 está incompleta.")
            row, col = struct.unpack_from("<HH", payload)
            if row > 65535 or col > 255:
                raise ValueError("Dimensiones A2 no admitidas.")
            if record == 3:
                value = struct.unpack_from("<d", payload, 7)[0]
            else:
                length = payload[7]
                if 8 + length > len(payload):
                    raise ValueError("Una etiqueta A2 está incompleta.")
                value = payload[8:8 + length].decode("cp1252", errors="replace")
            if (row, col) in cells:
                raise ValueError("El archivo contiene celdas superpuestas: revisa la exportación.")
            cells[row, col] = value
        elif record == 10:
            eof = True
            break
        elif record not in (0x0809, 0x0200):
            raise ValueError("Este formato XLS necesita el lector xlrd; no corresponde al exportador A2 validado.")
        offset += 4 + size
    if not cells or not eof:
        raise ValueError("No se pudo leer un reporte A2 completo.")
    height = max(r for r, c in cells) + 1
    width = max(c for r, c in cells) + 1
    grid = [[""] * width for _ in range(height)]
    for (row, col), value in cells.items():
        grid[row][col] = value
    return [("Reporte A2", grid)], "A2 BIFF antiguo"


def _a2_read_grids(raw, filename):
    if not raw:
        raise ValueError("El archivo está vacío.")
    legacy = raw[:4] == b"\x09\x08\x06\x00"
    if legacy:
        # La estructura real de los cuatro adjuntos es conocida y se valida
        # completa. Evita que calamine interprete como XLS moderno un BIFF viejo.
        return _a2_legacy_grid(raw)
    suffix = filename.lower().rsplit(".", 1)[-1]
    if suffix == "csv":
        last_error = None
        for encoding in ("utf-8-sig", "cp1252"):
            try:
                frame = pd.read_csv(io.BytesIO(raw), header=None, dtype=object, sep=None, engine="python", encoding=encoding, keep_default_na=False)
                return [("CSV", frame.values.tolist())], "CSV"
            except (UnicodeError, pd.errors.ParserError) as exc:
                last_error = exc
        raise ValueError("No se pudo leer el CSV; comprueba su separador.") from last_error
    try:
        engine = "openpyxl" if raw[:2] == b"PK" else "xlrd"
        sheets = pd.read_excel(io.BytesIO(raw), sheet_name=None, header=None, dtype=object, engine=engine, keep_default_na=False)
        return [(name, frame.values.tolist()) for name, frame in sheets.items()], engine
    except Exception as exc:
        raise ValueError("No se pudo leer el libro. Admite XLS de A2, XLSX y CSV; para XLS estándar instala xlrd.") from exc


def read_a2_grid(raw: bytes, filename: str = "reporte.xls") -> pd.DataFrame:
    """Lee celdas sin asumir columnas de CxC u otro reporte todavía no validado.

    Devuelve un DataFrame sin encabezado, con índices de fila/columna desde 0.
    La capa que conozca el reporte debe detectar, validar y mapear sus campos.
    Rechaza libros de varias hojas para evitar una elección silenciosa.
    """
    sheets, _ = _a2_read_grids(raw, filename)
    if len(sheets) != 1:
        raise ValueError("El archivo tiene varias hojas; selecciona o exporta una sola para esta carga.")
    return pd.DataFrame(sheets[0][1])


_A2_ALIASES = {
    "code": {"codigo", "cod", "codigo producto", "sku"},
    "description": {"descripcion", "producto", "descripcion producto"},
    "existence": {"existencia", "stock", "existencias"},
    "cost": {"costo", "costo unitario"},
    "inventory_value": {"valor inventario", "valor de inventario"},
    "quantity": {"cantidad", "ventas", "unidades vendidas"},
    "gross_sales": {"monto bruto", "importe bruto", "venta bruta"},
    "discounts": {"descuentos", "descuento"},
    "tax": {"i v a", "iva", "impuesto", "impuestos"},
    "profit": {"utilidad"},
    "margin": {"porcentaje utilidad"},
    "items": {"numero de items", "numero items", "items"},
    "active": {"activo"},
    "classification": {"clasificacion"},
    "address": {"direccion"},
    "supplier": {"proveedor"},
    "brand": {"marca"},
    "category": {"categoria"},
    "department": {"departamento"},
}


def _a2_header(grid, requested):
    candidates = []
    for row_index, row in enumerate(grid[:100]):
        fields = {}
        for col, value in enumerate(row):
            normalized = _a2_norm(value)
            if _a2_text(value).strip().startswith("%") and normalized == "utilidad":
                fields["margin"] = col
                continue
            for key, aliases in _A2_ALIASES.items():
                if normalized in aliases:
                    fields[key] = col
                    break
        if "description" not in fields:
            continue
        detected = None
        if {"existence", "cost"} <= fields.keys():
            detected = "inventario"
        elif {"quantity", "gross_sales"} <= fields.keys():
            detected = "ventas"
        elif {"active", "classification"} <= fields.keys():
            detected = "departamentos"
        elif {"active", "address"} <= fields.keys():
            detected = "proveedores"
        if detected and requested in ("auto", detected):
            candidates.append((len(fields), row_index, detected, fields))
    return max(candidates, default=None, key=lambda candidate: candidate[0])


def _a2_range_value(row, field, anchors, numeric=False):
    start = anchors.get(field)
    if start is None:
        return None if numeric else ""
    end = min((col for col in anchors.values() if col > start), default=len(row))
    choices = [value for value in row[start:end] if _a2_text(value) != ""]
    if numeric:
        numbers = [_a2_number(v) for v in choices]
        numbers = [v for v in numbers if v is not None]
        return numbers[-1] if numbers else None
    return _a2_text(choices[0]) if choices else ""


def _a2_date(text):
    match = re.search(r"\b(\d{2}/\d{2}/\d{4})\b", text)
    if not match:
        return None
    try:
        return datetime.strptime(match.group(1), "%d/%m/%Y").date().isoformat()
    except ValueError:
        return None


def _a2_metadata(grid, header_row, source, sheet, reader):
    metadata = {"filename": source, "sheet": sheet, "reader": reader,
                "header_row": header_row + 1, "currency": None, "site": None,
                "snapshot_date": None, "report_date": None,
                "period_start": None, "period_end": None,
                "unresolved_rows": [], "totals_printed": {},
                "reconciliation": [], "reconciled": None, "ready_to_save": True}
    for row in grid[:header_row]:
        text = " | ".join(_a2_text(v) for v in row if _a2_text(v))
        normal = _a2_norm(text)
        if "fecha" in normal and _a2_date(text):
            metadata["report_date"] = _a2_date(text)
        if normal.startswith("desde"):
            metadata["period_start"] = _a2_date(text)
        if normal.startswith("hasta"):
            metadata["period_end"] = _a2_date(text)
        if "moneda" in normal:
            if re.search(r"\b(dolares|dolar|usd)\b", normal):
                metadata["currency"] = "USD"
            elif re.search(r"\b(bolivares|bolivar|ves)\b", normal):
                metadata["currency"] = "VES"
        deposit = re.search(r"dep[oó]sitos?\s*:?[\s]+(.+)", text, re.I)
        if deposit:
            metadata["site"] = deposit.group(1).split("|")[0].strip().title()
    return metadata


def _a2_compare(metadata, field, observed, printed, warnings):
    if printed is None:
        return
    observed = float(observed)
    difference = round(observed - float(printed), 6)
    matches = abs(difference) <= 0.011
    metadata["totals_printed"][field] = float(printed)
    metadata["reconciliation"].append({"field": field, "extracted": observed,
        "printed": float(printed), "difference": difference, "matches": matches})
    if not matches:
        metadata["ready_to_save"] = False
        warnings.append(f"El total de {field} no coincide: extraído {observed:,.2f}; impreso {printed:,.2f}. Revisa el archivo antes de guardar.")


def parse_a2_report(raw: bytes, filename: str, kind: str = "auto") -> dict:
    """Extrae filas reales y su procedencia; nunca inventa SKU ni relaciones.

    metadata.ready_to_save=False requiere resolver identificadores duplicados/
    ausentes o diferencias con el total impreso. data conserva TODAS las filas,
    incluidas las pendientes; los consumidores no deben descartarlas en silencio.
    currency/site pueden ser None: la aplicación debe pedir el contexto faltante.
    """
    aliases = {"inventory": "inventario", "sales": "ventas", "departments": "departamentos", "suppliers": "proveedores"}
    requested = aliases.get(kind, kind)
    if requested not in ("auto", "inventario", "ventas", "departamentos", "proveedores"):
        raise ValueError("Tipo A2 no admitido: usa inventario, ventas, departamentos o proveedores.")
    sheets, reader = _a2_read_grids(raw, filename)
    matches = [(name, grid, _a2_header(grid, requested)) for name, grid in sheets]
    matches = [(name, grid, found) for name, grid, found in matches if found]
    if not matches:
        raise ValueError("No se reconoce este reporte A2. Selecciona el tipo correcto o usa la carga de tabla estándar.")
    if len(matches) > 1:
        raise ValueError("El libro contiene varios reportes reconocibles. Exporta cada reporte por separado para evitar omitir hojas.")
    sheet, grid, (_, header_row, detected, anchors) = matches[0]
    anchors = dict(anchors)
    if "code" not in anchors and detected == "ventas":
        anchors["code"] = 0
    metadata = _a2_metadata(grid, header_row, filename, sheet, reader)
    metadata["column_anchors"] = dict(anchors)
    metadata["snapshot_date"] = metadata["report_date"] if detected == "inventario" else metadata["period_end"]
    warnings = []
    records = []
    summary_rows = []
    printed_count = None
    numeric_keys = ("existence", "cost", "inventory_value") if detected == "inventario" else ("quantity", "gross_sales", "discounts", "tax", "cost", "profit")
    for source_index, row in enumerate(grid[header_row + 1:], start=header_row + 2):
        description = _a2_range_value(row, "description", anchors)
        code = _a2_range_value(row, "code", anchors)
        normalized = _a2_norm(description)
        # No descartar por un rótulo 'Total' en OTRA columna: el A2 real
        # imprime ese rótulo sobre una fila que también contiene un producto.
        is_detail = bool(description) and normalized not in {"descripcion", "total", "totales", "total general"}
        if detected in ("inventario", "ventas"):
            numbers = {key: _a2_range_value(row, key, anchors, numeric=True) for key in numeric_keys}
            required = ("existence", "cost") if detected == "inventario" else ("quantity", "gross_sales")
            is_detail = is_detail and all(numbers.get(key) is not None for key in required)
            if is_detail:
                record = {"Código": code, "Descripción": description, "Fila origen": source_index}
                if detected == "inventario":
                    record.update({"Existencia": numbers["existence"], "Costo": numbers["cost"],
                        "Valor inventario A2": numbers["inventory_value"],
                        "Proveedor": _a2_range_value(row, "supplier", anchors),
                        "Marca": _a2_range_value(row, "brand", anchors),
                        "Categoría": _a2_range_value(row, "category", anchors),
                        "Departamento": _a2_range_value(row, "department", anchors)})
                else:
                    record.update({"Ventas": numbers["quantity"], "Monto bruto": numbers["gross_sales"],
                        "Descuentos": numbers["discounts"], "IVA reportado": numbers["tax"],
                        "Costo de ventas": numbers["cost"], "Utilidad A2": numbers["profit"]})
                    record["Importe neto"] = float(Decimal(str(numbers["gross_sales"])) - Decimal(str(numbers["discounts"]))) if numbers["discounts"] is not None else None
                records.append(record)
            elif not description and sum(v is not None for v in numbers.values()) >= 2:
                summary_rows.append(numbers)
        elif detected == "departamentos":
            is_detail = is_detail and (_a2_norm(_a2_range_value(row, "active", anchors)) in ("si", "no") or _a2_range_value(row, "items", anchors, numeric=True) is not None)
            if is_detail:
                active = _a2_norm(_a2_range_value(row, "active", anchors))
                records.append({"Código": code, "Departamento": description, "Activo": True if active == "si" else False if active == "no" else None,
                    "Clasificación": _a2_range_value(row, "classification", anchors),
                    "Número de ítems": _a2_range_value(row, "items", anchors, numeric=True), "Fila origen": source_index})
        else:
            is_detail = is_detail and bool(code) and _a2_norm(_a2_range_value(row, "active", anchors)) in ("si", "no")
            if is_detail:
                records.append({"Código": code, "Proveedor": description,
                    "Activo": _a2_norm(_a2_range_value(row, "active", anchors)) == "si", "Fila origen": source_index})
        if not is_detail:
            for column, value in enumerate(row):
                label = _a2_norm(value)
                if label in {"total registros", "total items"}:
                    right_numbers = [_a2_number(v) for v in row[column + 1:]]
                    right_numbers = [v for v in right_numbers if v is not None]
                    if right_numbers:
                        if label == "total registros" or detected == "ventas":
                            printed_count = right_numbers[-1]
                        elif detected == "departamentos":
                            metadata["totals_printed"]["catalog_items"] = right_numbers[-1]
    if not records:
        raise ValueError("El encabezado se reconoce, pero no hay filas válidas de detalle.")
    frame = pd.DataFrame(records)
    metadata["row_count"] = len(frame)
    metadata["unresolved_rows"] = frame.loc[frame["Código"].eq("")].to_dict("records")
    if metadata["unresolved_rows"]:
        missing = len(metadata["unresolved_rows"])
        warnings.append(f"{missing} fila(s) no tienen código en el archivo original. Se conservan pendientes; asigna el código real antes de guardar por producto.")
        metadata["ready_to_save"] = False
    duplicate_codes = frame.loc[frame["Código"].ne("") & frame["Código"].duplicated(keep=False), "Código"].unique().tolist()
    metadata["duplicate_codes"] = duplicate_codes
    if duplicate_codes:
        metadata["ready_to_save"] = False
        warnings.append(f"Hay {len(duplicate_codes)} códigos repetidos; revisa si corresponden a depósitos/presentaciones diferentes antes de consolidar.")
    _a2_compare(metadata, "row_count", len(frame), printed_count, warnings)
    if detected in ("inventario", "ventas"):
        fields = {"existence": "Existencia", "inventory_value": "Valor inventario A2"} if detected == "inventario" else {
            "quantity": "Ventas", "gross_sales": "Monto bruto", "discounts": "Descuentos",
            "tax": "IVA reportado", "cogs": "Costo de ventas", "reported_profit": "Utilidad A2"}
        for field, column in fields.items():
            values = [_a2_number(value) for value in frame[column].tolist()]
            observed = _a2_sum(values)
            metadata[field] = observed if all(value is not None for value in values) else None
            raw_key = {"cogs": "cost", "reported_profit": "profit"}.get(field, field)
            printed_values = [row[raw_key] for row in summary_rows if row.get(raw_key) is not None]
            if len(set(printed_values)) == 1 and metadata[field] is not None:
                _a2_compare(metadata, field, observed, printed_values[0], warnings)
            elif len(set(printed_values)) > 1:
                warnings.append(f"Hay varios subtotales de {field}; no se tomó uno como total general automáticamente.")
        if detected == "ventas":
            metadata["net_sales"] = _a2_sum([_a2_number(v) for v in frame["Importe neto"]]) if frame["Importe neto"].notna().all() else None
            metadata["sales_granularity"] = "period"
            if metadata["period_start"] and metadata["period_end"]:
                start = datetime.fromisoformat(metadata["period_start"])
                end = datetime.fromisoformat(metadata["period_end"])
                if end < start:
                    metadata["ready_to_save"] = False
                    warnings.append("La fecha final del período es anterior a la inicial.")
                else:
                    metadata["period_days"] = (end - start).days + 1
            warnings.append("Ventas agrupadas por producto para todo el período. Este archivo no identifica ventas a crédito, cobros ni fechas individuales.")
            if metadata["currency"] is None:
                warnings.append("El reporte de ventas no indica moneda: selecciónala antes de usar los importes.")
            if metadata["net_sales"] is not None and metadata.get("cogs") is not None and metadata.get("reported_profit") is not None:
                metadata["profit_identity_difference"] = round(metadata["net_sales"] - metadata["cogs"] - metadata["reported_profit"], 6)
                if abs(metadata["profit_identity_difference"]) > 0.02:
                    warnings.append("La utilidad impresa por A2 difiere de neto menos costo. Se conservan ambos valores sin corregir el origen.")
        else:
            metadata["inventory_value_calculated"] = _a2_sum([float(Decimal(str(r["Existencia"])) * Decimal(str(r["Costo"]))) for r in records])
            differences = [r for r in records if r["Valor inventario A2"] is not None and abs(r["Existencia"] * r["Costo"] - r["Valor inventario A2"]) > 0.015]
            metadata["valuation_differences"] = len(differences)
            if differences:
                warnings.append(f"{len(differences)} producto(s) presentan diferencia entre costo × existencia y valor impreso por A2; se conserva el valor original para auditoría.")
            for label in ("Proveedor", "Marca", "Categoría", "Departamento"):
                if frame[label].eq("").all():
                    metadata.setdefault("missing_dimensions", []).append(label)
            if metadata.get("missing_dimensions"):
                warnings.append("El inventario no incluye " + ", ".join(metadata["missing_dimensions"]) + ". Se completan con relaciones confirmadas del catálogo.")
    if metadata["reconciliation"]:
        metadata["reconciled"] = all(item["matches"] for item in metadata["reconciliation"])
    if detected == "departamentos":
        warnings.append("El reporte lista departamentos, pero no contiene una relación explícita con proveedores o marcas. Clasificación no equivale a categoría de producto.")
    catalog = {"departments": frame.to_dict("records") if detected == "departamentos" else [],
               "suppliers": frame.to_dict("records") if detected == "proveedores" else [], "relationships": []}
    return {"kind": detected, "data": frame, "metadata": metadata, "warnings": warnings, "catalog": catalog}


# Catálogos precargados, extraídos de los archivos A2 suministrados el 04/09/2026.
# Se conservan códigos/estados del origen; sin direcciones, teléfonos ni correos.
A2_DEFAULT_DEPARTMENTS = [{'Código': '', 'Departamento': 'VOLTEX', 'Activo': None, 'Clasificación': ''},
 {'Código': '2', 'Departamento': 'ACCESORIOS SAFARI', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '3', 'Departamento': 'ACDELCO', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '4', 'Departamento': 'ATLANTIC', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '6', 'Departamento': 'BOSS', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '7', 'Departamento': 'CARSPRAY', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '8', 'Departamento': 'CASTROL', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '10', 'Departamento': 'DENSO', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '11', 'Departamento': 'DR MARCUS', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '12', 'Departamento': 'DUNCAN', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '13', 'Departamento': 'FERRETERIA', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '14', 'Departamento': 'FILTROS', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '15', 'Departamento': 'FILTROS K&N', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '16', 'Departamento': 'FILTROS RTC', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '17', 'Departamento': 'FILTROS WEB', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '18', 'Departamento': 'FILTROS WINNER', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '20', 'Departamento': 'FLUVECA', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '21', 'Departamento': 'FREEZING Y PASSARO', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '22', 'Departamento': 'GAT', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '23', 'Departamento': 'GRANEL', 'Activo': True, 'Clasificación': 'Ensamblaje'},
 {'Código': '24', 'Departamento': 'GULF', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '25', 'Departamento': 'HAMMER', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '26', 'Departamento': 'INCA', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '27', 'Departamento': 'LIQUI MOLY', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '28', 'Departamento': 'LUBRIMAX', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '30', 'Departamento': 'MOTUL', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '31', 'Departamento': 'MOTORCRAFT', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '32', 'Departamento': 'MOPAR', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '33', 'Departamento': 'OTROS', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '34', 'Departamento': 'PDV', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '35', 'Departamento': 'PEGATANKE Y SILICON', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '36', 'Departamento': 'PRODUCTOS POR CAJAS', 'Activo': True, 'Clasificación': 'Compuesto'},
 {'Código': '37', 'Departamento': 'PROMOCIONES', 'Activo': True, 'Clasificación': 'Compuesto'},
 {'Código': '38', 'Departamento': 'RODANOL', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '39', 'Departamento': 'RINOX', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '40', 'Departamento': 'VENEZOLANA DE SPRAY', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '41', 'Departamento': 'SHELL', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '42', 'Departamento': 'SKY', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '44', 'Departamento': 'SQ', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '45', 'Departamento': 'TERMINALES Y TIRRAP', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '46', 'Departamento': 'IVICA', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '48', 'Departamento': 'KEYSTONE', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '49', 'Departamento': 'OILVEN', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '50', 'Departamento': 'ROSHFRANS', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '51', 'Departamento': 'TOYOTA', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '52', 'Departamento': 'VALVOLINE', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '53', 'Departamento': 'ULTRALUB', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '54', 'Departamento': 'PROLED', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '55', 'Departamento': 'ULTRA OIL', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '56', 'Departamento': 'VENEMAX', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '57', 'Departamento': 'LAVADO Y MANTENIMIENTO', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '58', 'Departamento': 'DAUER', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '59', 'Departamento': 'IPONE', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '60', 'Departamento': 'OILSTONE', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '61', 'Departamento': 'F1', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '62', 'Departamento': 'MAGNUM-INNOVA', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '63', 'Departamento': 'MOBIL', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '64', 'Departamento': 'FILTROS MILLARD', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': '65', 'Departamento': 'LARSSON', 'Activo': True, 'Clasificación': 'Producto'},
 {'Código': 'BORRA', 'Departamento': 'BORRA', 'Activo': True, 'Clasificación': 'Producto'}]

A2_DEFAULT_SUPPLIERS = [{'Código': 'J000003505', 'Proveedor': 'ACUMULADORES DUNCAN, C.A', 'Activo': True},
 {'Código': 'J000565732', 'Proveedor': 'EL PINOSO FABRIL, C.A', 'Activo': True},
 {'Código': 'J000593221', 'Proveedor': 'JETFILTER, C.A', 'Activo': True},
 {'Código': 'J095161609', 'Proveedor': 'DIMAX, C.A', 'Activo': True},
 {'Código': 'J293548551', 'Proveedor': 'GRUPO AMIGO, C.A', 'Activo': True},
 {'Código': 'J301692349', 'Proveedor': 'EVERGREEN SERVICE, C.A', 'Activo': True},
 {'Código': 'J302006236', 'Proveedor': 'H. MOTORES VALENCIA, C.A', 'Activo': True},
 {'Código': 'J309052357', 'Proveedor': 'EMPRODES, C.A', 'Activo': True},
 {'Código': 'J310904693', 'Proveedor': 'IMPRESOS DI SOL, C.A', 'Activo': True},
 {'Código': 'J311153659', 'Proveedor': 'KEYSTONE, C.A', 'Activo': True},
 {'Código': 'J316015505', 'Proveedor': 'DISTRIBUIDORA BATOR, C.A', 'Activo': True},
 {'Código': 'J316228568', 'Proveedor': 'H MOTORES NAGUANAGUA, C.A', 'Activo': True},
 {'Código': 'J402865430', 'Proveedor': 'LUBRIGAMA SERVICE C.A', 'Activo': True},
 {'Código': 'J404825088', 'Proveedor': 'BOTLS AND TOOLS, C.A', 'Activo': True},
 {'Código': 'J406252719', 'Proveedor': 'MAKROPETROL, C.A', 'Activo': True},
 {'Código': 'J407816713', 'Proveedor': 'ESTRADIN, C.A', 'Activo': True},
 {'Código': 'J409979776', 'Proveedor': 'FILTLUBVEN, C.A', 'Activo': True},
 {'Código': 'J411252581', 'Proveedor': 'LUBRIGAMA, C.A', 'Activo': True},
 {'Código': 'J411257753', 'Proveedor': 'TRANSPORTE H&R 2018, C.A', 'Activo': True},
 {'Código': 'J411575410', 'Proveedor': 'REPRESENTACIONES GH 2018, C.A', 'Activo': True},
 {'Código': 'J413105438', 'Proveedor': 'DISBATTERY LUBRICANTES, S.A', 'Activo': True},
 {'Código': 'J500407858', 'Proveedor': 'CORPORACION SYGAR DE VENEZUELA, C.A', 'Activo': True},
 {'Código': 'J502489479', 'Proveedor': 'GRUPO FEDERAL CENTURY, C.A', 'Activo': False},
 {'Código': 'J502658246', 'Proveedor': 'CARS PETROLY, C.A', 'Activo': True},
 {'Código': 'J503092912', 'Proveedor': 'MAXIOIL GROUP, C.A', 'Activo': True},
 {'Código': 'J504512435', 'Proveedor': 'VOLTEX01, C.A.', 'Activo': True},
 {'Código': 'J508452240', 'Proveedor': 'REPRESENTACIONES ASGARD LED, C.A', 'Activo': True}]

A2_DEFAULT_RELATIONSHIPS = []  # El origen no contiene relaciones explícitas.


# ============================================================
# NEXUS 7 — CARGA UNIFICADA, CATÁLOGOS Y NIVELES DE INVENTARIO
# ============================================================

V7_DIMENSIONS = ["Departamento", "Proveedor", "Marca", "Categoría"]


def init_v7_db() -> None:
    with db() as con:
        exists = con.execute("SELECT 1 FROM sqlite_master WHERE name='v7_meta'").fetchone()
        has_data = con.execute("SELECT 1 FROM inventory_snapshots UNION ALL SELECT 1 FROM supplier_inventory_snapshots LIMIT 1").fetchone()
    if not exists and has_data:
        DB_PATH.with_name(f"nexus_respaldo_antes_v7_{datetime.now():%Y%m%d_%H%M%S_%f}.db").write_bytes(backup_database_bytes())
    with db() as con:
        for table in ["inventory_snapshots", "supplier_inventory_snapshots"]:
            present = {r[1] for r in con.execute(f"PRAGMA table_info({table})")}
            for column, sql_type in [("department", "TEXT DEFAULT ''"), ("brand", "TEXT DEFAULT ''"), ("source_value", "REAL")]:
                if column not in present:
                    con.execute(f"ALTER TABLE {table} ADD COLUMN {column} {sql_type}")
        con.executescript("""
            CREATE TABLE IF NOT EXISTS v7_meta (key TEXT PRIMARY KEY,value TEXT NOT NULL);
            CREATE TABLE IF NOT EXISTS v7_catalog (
                kind TEXT NOT NULL,code TEXT NOT NULL,name TEXT NOT NULL,active TEXT,classification TEXT,
                PRIMARY KEY(kind,name));
            CREATE TABLE IF NOT EXISTS v7_department_rules (
                department TEXT PRIMARY KEY,supplier TEXT NOT NULL DEFAULT '',brand TEXT NOT NULL DEFAULT '',
                category TEXT NOT NULL DEFAULT '',match_text TEXT NOT NULL DEFAULT '',confirmed INTEGER NOT NULL DEFAULT 0);
            CREATE TABLE IF NOT EXISTS v7_product_dimensions (
                code TEXT PRIMARY KEY,description TEXT,department TEXT NOT NULL DEFAULT '',supplier TEXT NOT NULL DEFAULT '',
                brand TEXT NOT NULL DEFAULT '',category TEXT NOT NULL DEFAULT '');
            CREATE TABLE IF NOT EXISTS v7_stock_levels (
                site TEXT NOT NULL,code TEXT NOT NULL,min_stock REAL NOT NULL,max_stock REAL NOT NULL,
                updated_at TEXT NOT NULL,PRIMARY KEY(site,code),CHECK(min_stock>=0 AND max_stock>=min_stock));
            CREATE TABLE IF NOT EXISTS v7_site_config (
                site TEXT NOT NULL,scope TEXT NOT NULL DEFAULT '',sales_start TEXT,sales_end TEXT,PRIMARY KEY(site,scope));
            CREATE TABLE IF NOT EXISTS v7_sales_batches (
                id INTEGER PRIMARY KEY AUTOINCREMENT,site TEXT NOT NULL,scope TEXT NOT NULL DEFAULT '',currency TEXT NOT NULL,
                start_date TEXT NOT NULL,end_date TEXT NOT NULL,file_name TEXT NOT NULL,sha256 TEXT NOT NULL,saved_at TEXT NOT NULL,
                total_units REAL NOT NULL,net_sales REAL,gross_sales REAL,discounts REAL,tax REAL,cogs REAL,credit_sales REAL,collections REAL,
                UNIQUE(site,scope,start_date,end_date));
            CREATE TABLE IF NOT EXISTS v7_sales_items (
                batch_id INTEGER NOT NULL,row_no INTEGER NOT NULL,code TEXT NOT NULL,description TEXT NOT NULL,
                units REAL NOT NULL,net_sales REAL,cogs REAL,gross_sales REAL,discounts REAL,tax REAL,
                PRIMARY KEY(batch_id,row_no));
            CREATE TABLE IF NOT EXISTS v7_import_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,saved_at TEXT NOT NULL,site TEXT NOT NULL,kind TEXT NOT NULL,
                filename TEXT NOT NULL,sha256 TEXT NOT NULL,metadata TEXT NOT NULL);
            CREATE TABLE IF NOT EXISTS v7_import_archive (
                id INTEGER PRIMARY KEY AUTOINCREMENT,archived_at TEXT NOT NULL,batch_metadata TEXT NOT NULL,items_json TEXT NOT NULL);
            CREATE INDEX IF NOT EXISTS idx_v7_sales_scope ON v7_sales_batches(site,scope,start_date,end_date);
            CREATE INDEX IF NOT EXISTS idx_v7_items_code ON v7_sales_items(batch_id,code);
        """)
        if not con.execute("SELECT 1 FROM v7_meta WHERE key='catalog_seeded'").fetchone():
            seed_catalogs_v7(con)
            con.execute("INSERT INTO v7_meta VALUES ('catalog_seeded','1')")
            touch_v7(con)


def touch_v7(con: sqlite3.Connection) -> None:
    con.execute("INSERT OR REPLACE INTO v7_meta VALUES ('version',?)", (datetime.now().isoformat(timespec="microseconds"),))


def v7_version() -> str:
    with db() as con:
        present = con.execute("SELECT 1 FROM sqlite_master WHERE name='v7_meta'").fetchone()
        if not present:
            return ""
        row = con.execute("SELECT value FROM v7_meta WHERE key='version'").fetchone()
    return row[0] if row else ""


def seed_catalogs_v7(con: sqlite3.Connection) -> None:
    for kind, rows in [("departamentos", A2_DEFAULT_DEPARTMENTS), ("proveedores", A2_DEFAULT_SUPPLIERS)]:
        for item in rows:
            code = str(item.get("Código", ""))
            name = str(item.get("Departamento" if kind == "departamentos" else "Proveedor", ""))
            if not name:
                continue
            con.execute("INSERT OR IGNORE INTO v7_catalog VALUES (?,?,?,?,?)", (kind, code, name, str(item.get("Activo", "")), str(item.get("Clasificación", ""))))
            if kind == "departamentos":
                con.execute("INSERT OR IGNORE INTO v7_department_rules(department) VALUES (?)", (name,))


def catalog_frame_v7(kind: str) -> pd.DataFrame:
    with db() as con:
        return pd.read_sql_query("SELECT code AS Código,name AS Descripción,active AS Activo,classification AS Clasificación FROM v7_catalog WHERE kind=? ORDER BY name", con, params=(kind,))


def save_catalog_v7(parsed: dict) -> None:
    kind = parsed["kind"]
    if kind not in {"departamentos", "proveedores"}:
        raise ValueError("Selecciona un reporte de departamentos o proveedores.")
    if parsed["metadata"].get("reconciled") is False:
        raise ValueError("Los registros del catálogo no coinciden con el total de A2.")
    with db() as con:
        for _, row in parsed["data"].iterrows():
            name = str(row["Departamento" if kind == "departamentos" else "Proveedor"]).strip()
            if not name:
                continue
            con.execute("INSERT INTO v7_catalog VALUES (?,?,?,?,?) ON CONFLICT(kind,name) DO UPDATE SET code=excluded.code,active=excluded.active,classification=excluded.classification",
                        (kind, str(row.get("Código", "")), name, str(row.get("Activo", "")), str(row.get("Clasificación", ""))))
            if kind == "departamentos":
                con.execute("INSERT OR IGNORE INTO v7_department_rules(department) VALUES (?)", (name,))
        touch_v7(con)
    st.cache_data.clear()


def enrich_dimensions_v7(frame: pd.DataFrame) -> pd.DataFrame:
    """Sólo aplica asignaciones explícitas o reglas previamente confirmadas."""
    if frame.empty:
        return frame.copy()
    data = frame.copy()
    for col in V7_DIMENSIONS:
        data[col] = data.get(col, pd.Series("", index=data.index)).astype(object).fillna("").astype(str)
    with db() as con:
        items = pd.read_sql_query("SELECT code AS Código,department AS Departamento,supplier AS Proveedor,brand AS Marca,category AS Categoría FROM v7_product_dimensions", con)
        rules = pd.read_sql_query("SELECT * FROM v7_department_rules WHERE confirmed=1", con)
    if not items.empty:
        by_code = items.set_index("Código")
        for col in V7_DIMENSIONS:
            assigned = data["Código"].astype(str).map(by_code[col]).fillna("")
            data[col] = data[col].where(assigned.eq(""), assigned)
    if not rules.empty:
        descriptions = data["Descripción"].astype(str).map(normalize_key)
        candidates = {i: [] for i in data.index}
        for _, rule in rules.iterrows():
            terms = [normalize_key(t) for t in str(rule["match_text"]).split(";") if normalize_key(t)]
            if terms:
                pattern = r"(?:^| )(?:" + "|".join(re.escape(t) for t in terms) + r")(?: |$)"
                for i in data.index[descriptions.str.contains(pattern, regex=True)]:
                    candidates[i].append(rule)
        for i, matches in candidates.items():
            if not data.at[i, "Departamento"] and len(matches) == 1:
                data.at[i, "Departamento"] = matches[0]["department"]
        by_dept = rules.set_index("department")
        for col, source in [("Proveedor", "supplier"), ("Marca", "brand"), ("Categoría", "category")]:
            linked = data["Departamento"].map(by_dept[source]).fillna("")
            data[col] = data[col].where(data[col].ne(""), linked)
    return data


def catalog_rule_suggestions_v7() -> pd.DataFrame:
    with db() as con:
        frame = pd.read_sql_query("SELECT department AS Departamento,supplier AS Proveedor,brand AS Marca,category AS Categoría,match_text AS 'Texto en descripción',confirmed AS Confirmada FROM v7_department_rules ORDER BY department", con)
    frame["Confirmada"] = frame["Confirmada"].astype(bool)
    generic = {"FERRETERIA", "FILTROS", "GRANEL", "OTROS", "PRODUCTOS POR CAJAS", "PROMOCIONES", "TERMINALES Y TIRRAP", "LAVADO Y MANTENIMIENTO", "BORRA", "ACCESORIOS SAFARI", "FREEZING Y PASSARO", "PEGATANKE Y SILICON", "VENEZOLANA DE SPRAY", "MAGNUM-INNOVA"}
    for i, row in frame.iterrows():
        dept = str(row["Departamento"])
        if row["Confirmada"] or row["Texto en descripción"] or dept in generic:
            continue
        brand = dept.removeprefix("FILTROS ")
        frame.at[i, "Texto en descripción"] = brand
        frame.at[i, "Marca"] = brand
        if dept.startswith("FILTROS "):
            frame.at[i, "Categoría"] = "Filtros"
    return frame


def _null_v7(value):
    return None if value is None or pd.isna(value) else float(value)


def sales_overlaps_v7(site: str, scope: str, start: str, end: str) -> pd.DataFrame:
    with db() as con:
        return pd.read_sql_query("SELECT * FROM v7_sales_batches WHERE site=? AND scope=? AND start_date<=? AND end_date>=?", con, params=(site, scope, end, start))


def resolve_sales_codes_v7(sales: pd.DataFrame, inv: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    data = sales.copy()
    if inv.empty:
        return data, 0
    lookup = inv[["Código", "Descripción"]].copy()
    lookup["match"] = lookup["Descripción"].astype(str).str.casefold().str.replace(r"\s+", " ", regex=True).str.strip()
    lookup = lookup[lookup["match"].ne("") & ~lookup["match"].duplicated(keep=False)]
    keymap = lookup.set_index("match")["Código"]
    mask = data["Código"].astype(str).str.strip().eq("")
    keys = data.loc[mask, "Descripción"].astype(str).str.casefold().str.replace(r"\s+", " ", regex=True).str.strip()
    found = keys.map(keymap).fillna("")
    data.loc[mask, "Código"] = found
    return data, int(found.ne("").sum())



def init_optional_costs_v71():
    """Distinguish absent unit costs from explicit zero without replacing history."""
    tables = ("inventory_snapshots", "supplier_inventory_snapshots")
    with db() as con:
        missing = [table for table in tables
                   if "cost_known" not in {r[1] for r in con.execute(f"PRAGMA table_info({table})")}]
        existing = bool(missing) and any(con.execute(f"SELECT 1 FROM {table} LIMIT 1").fetchone() for table in missing)
    if existing:
        backup = DB_PATH.with_name(f"nexus_respaldo_pre_costos_v71_{datetime.now():%Y%m%d_%H%M%S_%f}.db")
        with sqlite3.connect(DB_PATH) as origin, sqlite3.connect(backup) as destination:
            origin.backup(destination)
    with db() as con:
        for table in missing:
            con.execute(f"ALTER TABLE {table} ADD COLUMN cost_known INTEGER NOT NULL DEFAULT 0")
            # Historical zero could mean missing: do not retroactively declare it supplied.
            con.execute(f"UPDATE {table} SET cost_known=CASE WHEN cost<>0 THEN 1 ELSE 0 END")


def _restore_optional_costs_v71(result, inventory, abc_basis="Movimiento"):
    out = result.copy()
    costs = pd.to_numeric(inventory.get("Costo", pd.Series(np.nan, index=inventory.index)), errors="coerce")
    costs = costs.where(np.isfinite(costs))
    lookup = pd.Series(costs.to_numpy(), index=inventory["Código"].astype(str))
    supplied = out["Código"].astype(str).map(lookup)
    out["Costo"] = supplied
    out["Costo informado"] = np.where(supplied.notna(), "Sí", "No")
    out["Tiene costo"] = out["Costo informado"].astype(object)
    reported = pd.to_numeric(out.get("Valor inventario A2", pd.Series(np.nan, index=out.index)), errors="coerce")
    reported = reported.where(np.isfinite(reported))
    value = out["Existencia"] * supplied
    value = value.where(out["Existencia"].ne(0), 0.0)
    out["Valor Inventario ($)"] = reported.where(reported.notna(), value).round(2)
    for field, units in [("Capital Inmovilizado ($)", "Retiro Almacén"),
                         ("Costo Compra Estimada ($)", "Compra Sugerida"),
                         ("Valor Movimiento ($)", "Demanda Mensual")]:
        quantity = pd.to_numeric(out[units], errors="coerce")
        out[field] = (quantity * supplied).where(quantity.ne(0), 0.0).round(2)
    metric = ("Valor Inventario ($)" if abc_basis == "Inventario" else
              "Capital Inmovilizado ($)" if abc_basis == "Capital inmovilizado" else "Valor Movimiento ($)")
    out["ABC"] = out["ABC"].astype(object)
    out.loc[out[metric].isna(), "ABC"] = "N/D"
    out["Cobertura valoración"] = np.where(out["Valor Inventario ($)"].notna(), "Disponible", "Costo o valor pendiente")
    return out


def monetary_sum_v71(values):
    series = pd.to_numeric(pd.Series(values, dtype=object), errors="coerce")
    if series.empty:
        return 0.0
    return float(series.sum()) if series.notna().all() and np.isfinite(series).all() else np.nan

def _store_inventory_v7(con, site, cutoff, inventory, scope):
    data = inventory.copy()
    for col in ["Proveedor", "Marca", "Categoría", "Departamento"]:
        if col not in data:
            data[col] = ""
    if scope:
        data["Proveedor"] = scope
    if data["Código"].astype(str).str.strip().eq("").any() or data["Código"].duplicated().any():
        raise ValueError("El inventario tiene códigos vacíos o repetidos.")
    base = []
    for _, r in data.iterrows():
        raw_cost = pd.to_numeric(pd.Series([r.get("Costo")]), errors="coerce").iloc[0]
        known = bool(pd.notna(raw_cost) and np.isfinite(raw_cost))
        existence = float(r["Existencia"])
        if not np.isfinite(existence):
            raise ValueError("Hay existencias no numéricas o no finitas.")
        if known and raw_cost < 0:
            raise ValueError("El costo unitario no puede ser negativo.")
        tail = (str(r["Código"]), str(r["Descripción"]), existence, float(raw_cost) if known else 0.0)
        value = _null_v7(r.get("Valor inventario A2"))
        if value is not None and not np.isfinite(float(value)):
            raise ValueError("El valor de inventario debe ser finito o quedar vacío.")
        if scope:
            base.append((cutoff, site, normalize_key(scope), scope) + tail + (str(r["Categoría"]), str(r["Departamento"]), str(r["Marca"]), value, int(known)))
        else:
            base.append((cutoff, site) + tail + (str(r["Proveedor"]), str(r["Categoría"]), str(r["Departamento"]), str(r["Marca"]), value, int(known)))
    if scope:
        con.execute("DELETE FROM supplier_inventory_snapshots WHERE site=? AND snapshot_date=? AND supplier_key=?", (site, cutoff, normalize_key(scope)))
        con.executemany("INSERT INTO supplier_inventory_snapshots(snapshot_date,site,supplier_key,supplier_name,code,description,existence,cost,category,department,brand,source_value,cost_known) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", base)
        con.execute("INSERT INTO supplier_processing_log(processed_at,snapshot_date,site,supplier_key,supplier_name,inventory_rows,sales_rows) VALUES(?,?,?,?,?,?,0)", (datetime.now().isoformat(), cutoff, site, normalize_key(scope), scope, len(data)))
    else:
        con.execute("DELETE FROM inventory_snapshots WHERE site=? AND snapshot_date=?", (site, cutoff))
        con.executemany("INSERT INTO inventory_snapshots(snapshot_date,site,code,description,existence,cost,supplier,category,department,brand,source_value,cost_known) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)", base)
        con.execute("INSERT INTO processing_log(processed_at,snapshot_date,site,inventory_rows,sales_rows) VALUES(?,?,?,?,0)", (datetime.now().isoformat(timespec="microseconds"), cutoff, site, len(data)))


def save_upload_bundle_v7(site: str, currency: str, cutoff: str, inventory: dict | None,
                           sales_reports: list[dict], ar: pd.DataFrame | None, scope: str = "",
                           replace_overlaps: bool = False) -> dict:
    if site not in SEDES or currency not in FIN_CURRENCIES:
        raise ValueError("Selecciona sede y moneda válidas.")
    if cutoff > date.today().isoformat():
        raise ValueError("El corte no puede ser futuro.")
    if scope and ar is not None:
        raise ValueError("La cartera se carga para la sede completa, no para un proveedor.")
    spans = sorted((p["metadata"]["period_start"], p["metadata"]["period_end"]) for p in sales_reports)
    for i, (start, end) in enumerate(spans):
        if not start or not end or start > end or end > cutoff:
            raise ValueError("Revisa las fechas de ventas: deben ser válidas y terminar como máximo en la fecha del corte.")
        if i and start <= spans[i-1][1]:
            raise ValueError("Dos archivos seleccionados cubren días repetidos. Carga períodos que no se solapen.")
    if inventory is None and not sales_reports and ar is None:
        raise ValueError("Carga al menos un reporte.")
    counts = {"Inventario": 0, "Ventas": 0, "CxC": 0, "Sin código": 0}
    with db() as con:
        if inventory is not None:
            if inventory["metadata"].get("reconciled") is False:
                raise ValueError("El inventario no coincide con sus totales impresos.")
            _store_inventory_v7(con, site, cutoff, inventory["data"], scope)
            counts["Inventario"] = len(inventory["data"])
            if not scope:
                con.execute("INSERT INTO financial_settings VALUES (?,?) ON CONFLICT(site) DO UPDATE SET inventory_currency=excluded.inventory_currency", (site, currency))
        for parsed in sales_reports:
            meta, data = parsed["metadata"], parsed["data"]
            if meta.get("reconciled") is False:
                raise ValueError("Un reporte de ventas no concilia con los totales de A2.")
            start, end = meta["period_start"], meta["period_end"]
            old = pd.read_sql_query("SELECT * FROM v7_sales_batches WHERE site=? AND scope=? AND start_date<=? AND end_date>=?", con, params=(site, scope, end, start))
            for _, prior in old.iterrows():
                exact = prior["start_date"] == start and prior["end_date"] == end
                covered = start <= prior["start_date"] and prior["end_date"] <= end
                if not exact and (not covered or not replace_overlaps):
                    raise ValueError(f"El período {start}–{end} se solapa con {prior['start_date']}–{prior['end_date']}. No se puede sumar ni recortar automáticamente.")
                stored = pd.read_sql_query("SELECT * FROM v7_sales_items WHERE batch_id=?", con, params=(int(prior["id"]),))
                con.execute("INSERT INTO v7_import_archive(archived_at,batch_metadata,items_json) VALUES(?,?,?)", (datetime.now().isoformat(), prior.to_json(force_ascii=False), stored.to_json(orient="records", force_ascii=False)))
                con.execute("DELETE FROM v7_sales_items WHERE batch_id=?", (int(prior["id"]),))
                con.execute("DELETE FROM v7_sales_batches WHERE id=?", (int(prior["id"]),))
            vals = (site, scope, currency, start, end, meta.get("filename", "A2"), meta.get("sha256", ""), datetime.now().isoformat(timespec="microseconds"), float(data["Ventas"].sum()),
                    _null_v7(meta.get("net_sales")), _null_v7(meta.get("gross_sales")), _null_v7(meta.get("discounts")), _null_v7(meta.get("tax")), _null_v7(meta.get("cogs")), _null_v7(meta.get("credit_sales")), _null_v7(meta.get("collections")))
            cursor = con.execute("INSERT INTO v7_sales_batches(site,scope,currency,start_date,end_date,file_name,sha256,saved_at,total_units,net_sales,gross_sales,discounts,tax,cogs,credit_sales,collections) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", vals)
            rows = []
            for i, r in data.reset_index(drop=True).iterrows():
                rows.append((cursor.lastrowid, i, str(r["Código"]).strip(), str(r["Descripción"]), float(r["Ventas"]), _null_v7(r.get("Importe neto")), _null_v7(r.get("Costo de ventas")), _null_v7(r.get("Monto bruto")), _null_v7(r.get("Descuentos")), _null_v7(r.get("IVA reportado"))))
            con.executemany("INSERT INTO v7_sales_items VALUES(?,?,?,?,?,?,?,?,?,?)", rows)
            counts["Ventas"] += len(data)
            counts["Sin código"] += int(data["Código"].astype(str).str.strip().eq("").sum())
        if spans:
            con.execute("INSERT INTO v7_site_config VALUES (?,?,?,?) ON CONFLICT(site,scope) DO UPDATE SET sales_start=excluded.sales_start,sales_end=excluded.sales_end", (site, scope, spans[0][0], max(s[1] for s in spans)))
        if ar is not None:
            con.execute("DELETE FROM financial_ar WHERE as_of=? AND site=? AND currency=?", (cutoff, site, currency))
            rows = [(cutoff, site, currency, r[0], r[1], r[2], r[3], float(r[4]), _null_v7(r[5])) for r in ar[list(FIN_AR_FIELDS)].itertuples(index=False, name=None)]
            con.executemany("INSERT INTO financial_ar VALUES (?,?,?,?,?,?,?,?,?)", rows)
            con.execute("INSERT INTO financial_ar_batches VALUES(?,?,?,?) ON CONFLICT(as_of,site,currency) DO UPDATE SET row_count=excluded.row_count", (cutoff, site, currency, len(ar)))
            con.execute("INSERT INTO financial_log(saved_at,kind,site,currency,start_date,end_date,rows,file_name,sha256) VALUES(?,?,?,?,?,?,?,?,?)", (datetime.now().isoformat(), "cxc", site, currency, cutoff, cutoff, len(ar), "Carga unificada", ""))
            counts["CxC"] = len(ar)
        for p in ([inventory] if inventory else []) + sales_reports:
            metadata = p["metadata"].copy()
            metadata.pop("unresolved_rows", None)
            con.execute("INSERT INTO v7_import_log(saved_at,site,kind,filename,sha256,metadata) VALUES (?,?,?,?,?,?)", (datetime.now().isoformat(), site, p["kind"], metadata.get("filename", ""), metadata.get("sha256", ""), json.dumps(metadata, ensure_ascii=False, default=str)))
        touch_v7(con)
    st.cache_data.clear()
    return counts


def _selected_sales_batches_v73(con, site: str, cutoff: str, scope: str = ""):
    """Choose whole stored reports; never prorate a report to fit an inventory cut."""
    config = con.execute("SELECT sales_start,sales_end FROM v7_site_config WHERE site=? AND scope=?", (site, scope)).fetchone()
    if not config:
        return pd.DataFrame(), {}
    meta = {"start": str(config[0]), "end": str(config[1]), "days": 0,
            "pending": 0, "pending_units": 0.0, "currencies": [],
            "source_signature": "", "blocked": False, "excluded": [], "reason": ""}
    try:
        start_window, end_window = (date.fromisoformat(str(value)) for value in config)
        cut = min(date.fromisoformat(str(cutoff)), date.today())
        if start_window > end_window:
            raise ValueError("intervalo invertido")
    except (TypeError, ValueError):
        meta.update(blocked=True, reason="El corte o las fechas del historial guardado no son válidos. Revisa el historial antes de calcular.")
        return pd.DataFrame(), meta
    stored = pd.read_sql_query("SELECT * FROM v7_sales_batches WHERE site=? AND scope=? ORDER BY start_date,id", con, params=(site, scope))
    eligible = []
    invalid_selected = False
    for index, row in stored.iterrows():
        info = {"id": int(row["id"]), "file_name": str(row["file_name"]),
                "start": str(row["start_date"]), "end": str(row["end_date"])}
        try:
            begin = date.fromisoformat(info["start"])
        except (TypeError, ValueError):
            begin = None
        try:
            end = date.fromisoformat(info["end"])
        except (TypeError, ValueError):
            end = None
        if begin is None or end is None:
            if (end is not None and end < start_window) or (begin is not None and begin > end_window):
                continue
            info["reason"] = "Fechas del reporte no válidas; no se utiliza."
            meta["excluded"].append(info)
            invalid_selected = True
            continue
        if begin > end:
            if min(begin, end) <= end_window and max(begin, end) >= start_window:
                info["reason"] = "El inicio del reporte es posterior a su final."
                meta["excluded"].append(info)
                invalid_selected = True
            continue
        if begin < start_window or end > end_window:
            continue
        if end > cut:
            info["reason"] = ("Las ventas terminan después del corte de inventario."
                              if end > date.fromisoformat(str(cutoff))
                              else "Las ventas terminan después de la fecha actual.")
            meta["excluded"].append(info)
            continue
        eligible.append(index)
    selected = stored.loc[eligible].copy()
    if invalid_selected:
        meta.update(blocked=True, reason="Hay fechas inválidas en el historial de esta sede. Corrige ese historial antes de calcular propuestas automáticas.")
        return selected.iloc[0:0], meta
    if not selected.empty:
        previous_end = None
        for row in selected.itertuples():
            begin, end = date.fromisoformat(row.start_date), date.fromisoformat(row.end_date)
            if previous_end is not None and begin <= previous_end:
                meta.update(blocked=True, reason="El historial guardado contiene reportes con días repetidos. Revisa sus períodos para evitar sumar ventas dos veces.")
                return selected.iloc[0:0], meta
            previous_end = end
    if selected.empty:
        meta.update(blocked=True, reason=("Ningún reporte completo del historial elegido termina dentro del corte de inventario y de la fecha actual. Revisa el corte o el historial; no se recortan ventas acumuladas automáticamente."))
        return selected, meta
    signature_rows = [[int(row["id"]), str(row["start_date"]), str(row["end_date"]),
                       str(row["sha256"]), str(row["saved_at"])] for _, row in selected.iterrows()]
    meta.update(start=str(selected["start_date"].min()), end=str(selected["end_date"].max()),
                days=sum((date.fromisoformat(row.end_date) - date.fromisoformat(row.start_date)).days + 1 for row in selected.itertuples()),
                currencies=sorted(selected["currency"].astype(str).unique()),
                source_signature=hashlib.sha256(json.dumps(signature_rows, ensure_ascii=False, separators=(",", ":")).encode("utf-8")).hexdigest())
    return selected, meta


def operating_sales_v7(site: str, cutoff: str, scope: str = "") -> tuple[pd.DataFrame, dict]:
    with db() as con:
        batches, meta = _selected_sales_batches_v73(con, site, cutoff, scope)
        if batches.empty:
            return pd.DataFrame(), meta
        ids = [int(v) for v in batches["id"]]
        items = pd.read_sql_query(f"SELECT code AS Código,description AS Descripción,units AS Ventas,net_sales AS 'Importe neto',cogs AS 'Costo de ventas',gross_sales AS 'Monto bruto',discounts AS Descuentos,tax AS 'IVA reportado' FROM v7_sales_items WHERE batch_id IN ({','.join('?' for _ in ids)})", con, params=ids)
    pending = items["Código"].fillna("").astype(str).str.strip().eq("")
    meta.update(pending=int(pending.sum()), pending_units=float(items.loc[pending, "Ventas"].sum()))
    resolved = items[~pending].copy()
    agg = {"Descripción": "first", "Ventas": "sum"}
    for col in ["Importe neto", "Costo de ventas", "Monto bruto", "Descuentos", "IVA reportado"]:
        agg[col] = lambda s: float(s.sum()) if s.notna().all() else np.nan
    grouped = resolved.groupby("Código", as_index=False, sort=False).agg(agg)
    if len(meta["currencies"]) != 1:
        grouped[["Importe neto", "Costo de ventas", "Monto bruto", "Descuentos", "IVA reportado"]] = np.nan
    return grouped, meta


@st.cache_data(show_spinner=False, max_entries=12)
def _sales_rows_for_reports_v7(version: str, cuts: tuple, scope: str = "") -> list[dict]:
    frames = []
    for site, cutoff in cuts:
        with db() as con:
            batches, meta = _selected_sales_batches_v73(con, site, cutoff, scope)
            if batches.empty:
                continue
            ids = [int(value) for value in batches["id"]]
            placeholders = ','.join('?' for _ in ids)
            frame = pd.read_sql_query(f"""
                SELECT b.site AS Sede,i.code AS Código,i.description AS Descripción,i.units AS Ventas,
                       i.net_sales AS 'Importe neto',i.cogs AS 'Costo de ventas',i.gross_sales AS 'Monto bruto',
                       i.discounts AS Descuentos,i.tax AS 'IVA reportado',b.currency AS 'Moneda ventas',
                       b.start_date||' a '||b.end_date AS 'Período ventas',i.row_no+1 AS 'Fila de ventas',
                       b.file_name AS 'Archivo de ventas'
                FROM v7_sales_items i JOIN v7_sales_batches b ON b.id=i.batch_id
                WHERE b.id IN ({placeholders}) ORDER BY b.start_date,i.row_no""", con, params=ids)
        if not frame.empty:
            frame = enrich_dimensions_v7(frame)
            if scope:
                frame["Proveedor"] = scope
            frames.append(frame)
    if not frames:
        return []
    return json.loads(pd.concat(frames, ignore_index=True).to_json(orient="records", force_ascii=False))


def sales_rows_for_reports_v7(dates_used: dict) -> list[dict]:
    return _sales_rows_for_reports_v7(v7_version(), tuple(sorted(dates_used.items())))


@st.cache_data(show_spinner=False, max_entries=6)
def build_rotation_history_v7(data_version: str, sales_mode: str) -> pd.DataFrame:
    legacy = build_rotation_history(data_version, sales_mode)
    with db() as con:
        periods = pd.read_sql_query("""SELECT b.site AS Sede,i.code AS Código,b.start_date,b.end_date
            FROM v7_sales_items i JOIN v7_sales_batches b ON b.id=i.batch_id
            WHERE b.scope='' AND i.code<>'' AND i.units>0 ORDER BY b.end_date""", con)
    if periods.empty:
        return legacy
    periods = periods.drop_duplicates(["Sede", "Código"], keep="last")
    periods["Último período con ventas"] = periods["start_date"] + " a " + periods["end_date"]
    return legacy.merge(periods[["Sede", "Código", "Último período con ventas"]], on=["Sede", "Código"], how="outer")


def apply_stock_settings_v7(inv: pd.DataFrame, site: str) -> pd.DataFrame:
    result = enrich_dimensions_v7(inv)
    with db() as con:
        manual = pd.read_sql_query("SELECT code AS Código,min_stock AS 'Mínimo manual',max_stock AS 'Máximo manual' FROM v7_stock_levels WHERE site=?", con, params=(site,))
    if not manual.empty:
        result = result.merge(manual, on="Código", how="left", validate="one_to_one")
    return result


def calculate_loaded_site_v7(site: str, cutoff: str, scope: str, months_history: int, min_coverage: float,
                             max_coverage: float, lead_time_days: int, safety_days: int,
                             abc_basis: str, sales_mode: str, rolling_days: int) -> pd.DataFrame:
    inv, legacy_sales, previous = load_supplier_snapshot(site, scope, cutoff) if scope else load_snapshot(site, cutoff)
    inv = apply_stock_settings_v7(inv, site)
    sales, metadata = operating_sales_v7(site, cutoff, scope)
    if metadata.get("blocked"):
        legacy_sales = legacy_sales.iloc[0:0].copy()
    has_history = (bool(metadata and metadata.get("days", 0) > 0) or not legacy_sales.empty) and not metadata.get("blocked", False)
    average_months, basis_pending = stock_basis_v73(site, scope, metadata)
    if metadata and metadata.get("days", 0) > 0:
        period_days = metadata["days"]
        mode = "periodo"
    else:
        sales, period_days, mode = legacy_sales, None, sales_mode
    if sales.empty:
        sales = pd.DataFrame({"Código": inv["Código"].astype(str), "Ventas": 0.0})
    result = calculate_site(inv, sales, previous, months_history, min_coverage, max_coverage,
                            lead_time_days, safety_days, abc_basis, mode, rolling_days, period_days=period_days, average_months=average_months)
    if metadata and metadata.get("days", 0) > 0:
        extras = [c for c in ["Código", "Importe neto", "Monto bruto", "Descuentos", "IVA reportado", "Costo de ventas"] if c in sales]
        if len(extras) > 1:
            result = result.merge(sales[extras], on="Código", how="left", validate="one_to_one")
        result["Período ventas"] = f"{metadata['start']} a {metadata['end']}"
        result["Días del historial"] = period_days
        result["Moneda ventas"] = metadata["currencies"][0] if len(metadata["currencies"]) == 1 else "Varias · importes no sumables"
    else:
        result["Período ventas"] = "Historial anterior · configuración del motor"
    result = enrich_dimensions_v7(result)
    result["Fecha del inventario"] = cutoff
    result["Moneda inventario"] = finance_settings().get(site, "Sin moneda") if not scope else "Sin moneda"
    result["Corte desactualizado"] = "Sí" if metadata.get("end", cutoff) > cutoff else "No"
    result["Historial de ventas"] = "Disponible" if has_history else "Pendiente de cargar"
    if not has_history:
        automatic = result["Política de stock"].eq("Automática")
        for col in ["Stock Mínimo", "Stock Máximo", "Punto de Reorden", "Mínimo automático", "Máximo automático"]:
            result[col] = result[col].astype(float)
            result.loc[automatic, col] = np.nan
        for col in ["Compra Sugerida", "Retiro Almacén", "Capital Inmovilizado ($)", "Costo Compra Estimada ($)"]:
            result.loc[automatic, col] = 0
        for col in ["Estado", "Acción", "Sin movimiento"]:
            result[col] = result[col].astype(object)
        result.loc[automatic & result["Existencia"].ge(0), "Estado"] = "PENDIENTE — CARGAR VENTAS"
        result.loc[automatic, "Acción"] = "CARGAR HISTORIAL DE VENTAS"
        result["Sin movimiento"] = "Sin historial"
        result["Período ventas"] = "Pendiente de cargar"
    result["Alcance del historial"] = scope
    result["Revisión del cálculo"] = ""
    result["Ventas sin código del historial"] = metadata.get("pending_units", 0)
    automatic = result["Política de stock"].eq("Automática")
    blocked = pd.Series(False, index=result.index)
    if not has_history:
        blocked |= automatic
        result.loc[automatic, "Revisión del cálculo"] = metadata.get("reason") or "No hay un período de ventas válido para este corte. Carga o revisa el historial."
    if basis_pending:
        blocked |= automatic
        result.loc[automatic, "Revisión del cálculo"] = "Cambió el historial: confirma los meses de esta nueva carga o vuelve a días reales."
    if metadata.get("pending", 0):
        unmatched = automatic & result["Ventas"].le(0)
        blocked |= unmatched
        result.loc[unmatched, "Revisión del cálculo"] = "Hay ventas sin código. Confirma si corresponden a este producto antes de tratarlo como sin ventas."
    if metadata.get("excluded"):
        result["Fuentes excluidas del cálculo"] = str(metadata["excluded"])
    for column in ["Stock Mínimo","Stock Máximo","Punto de Reorden","Mínimo automático","Máximo automático"]:
        result[column] = pd.to_numeric(result[column],errors="coerce").astype(float)
        result.loc[blocked,column] = np.nan
    for column in ["Compra Sugerida","Retiro Almacén","Capital Inmovilizado ($)","Costo Compra Estimada ($)"]:
        result.loc[blocked,column] = 0.0
    for column in ["Estado","Acción","Sin movimiento"]:
        result[column] = result[column].astype(object)
    result.loc[blocked & result["Existencia"].ge(0),"Estado"] = "PENDIENTE — REVISAR VENTAS"
    result.loc[blocked,"Acción"] = "REVISAR BASE DEL CÁLCULO"
    result.loc[blocked,"Sin movimiento"] = "Pendiente de verificar"
    result.loc[result["Ventas"].lt(0),"Revisión del cálculo"] = "Ventas netas negativas: se conserva el saldo, pero se usa demanda cero. Revisa devoluciones y período."
    return _restore_optional_costs_v71(result, inv, abc_basis)


@st.cache_data(show_spinner=False, max_entries=12)
def load_live_results_v7(data_version, months_history, min_coverage, max_coverage, lead_time_days,
                          safety_days, abc_basis, sales_mode, rolling_days):
    results, dates_used = {}, {}
    for site in SEDES:
        cutoff = latest_snapshot_date(site)
        if not cutoff:
            continue
        results[site] = calculate_loaded_site_v7(site, cutoff, "", months_history, min_coverage, max_coverage,
                                                lead_time_days, safety_days, abc_basis, sales_mode, rolling_days)
        dates_used[site] = cutoff
    return results, dates_used


@st.cache_data(show_spinner=False, max_entries=12)
def load_supplier_live_results_v7(data_version, supplier_version, supplier_name, selected_sites,
                                   months_history, min_coverage, max_coverage, lead_time_days,
                                   safety_days, abc_basis, sales_mode, rolling_days):
    results, dates_used, sources = {}, {}, {}
    for site in selected_sites:
        dedicated = latest_supplier_snapshot_date(site, supplier_name)
        cutoff = dedicated or latest_snapshot_date(site)
        if not cutoff:
            continue
        data = calculate_loaded_site_v7(site, cutoff, supplier_name if dedicated else "", months_history,
                                        min_coverage, max_coverage, lead_time_days, safety_days, abc_basis, sales_mode, rolling_days)
        if not dedicated:
            data = data[data["Proveedor"].astype(str).map(normalize_key).eq(normalize_key(supplier_name))].copy()
        if data.empty:
            continue
        results[site], dates_used[site] = data, cutoff
        sources[site] = "Corte dedicado" if dedicated else "Catálogo confirmado del corte maestro"
    return results, dates_used, sources


# NEXUS 7 · Ventas reales por día o período, sin distribuir importes ficticios.
# Integrar después de las funciones financieras v6 y de la migración v7.


def _financial_sales_selection_v7(daily, periods, start, end):
    """Elige fuentes sin duplicar días; los períodos recortados nunca se prorratean."""
    expected = set(pd.date_range(start, end).strftime("%Y-%m-%d"))
    daily = daily.copy()
    periods = periods.copy()
    notes, rejected = [], []
    # Una serie diaria completa conserva su detalle y tiene prioridad.
    if set(daily["day"]) == expected and not daily["day"].duplicated().any():
        return daily, periods.iloc[:0].copy(), len(expected), notes, rejected
    candidates = periods[(periods["start_date"] >= start) & (periods["end_date"] <= end)].copy()
    clipped = periods[~periods.index.isin(candidates.index)]
    if not clipped.empty:
        notes.append(f"{len(clipped)} reporte(s) de período atraviesan los filtros: excluidos sin prorratear")
        rejected.extend(dict(row, motivo="Período fuera de los límites seleccionados") for row in clipped.to_dict("records"))
    invalid = set()
    intervals = []
    for idx, row in candidates.sort_values(["start_date", "end_date"]).iterrows():
        try:
            first, last = date.fromisoformat(row["start_date"]), date.fromisoformat(row["end_date"])
            if first > last:
                raise ValueError("Rango invertido")
        except (TypeError, ValueError):
            invalid.add(idx)
            continue
        # Ante datos históricos inconsistentes, excluye TODOS los bloques en
        # conflicto en lugar de escoger arbitrariamente y duplicar sus valores.
        for other_idx, other_first, other_last in intervals:
            if first <= other_last and last >= other_first:
                invalid.update([idx, other_idx])
        intervals.append((idx, first, last))
    if invalid:
        notes.append(f"{len(invalid)} período(s) inválidos o solapados excluidos; vuelve a cargar los reportes")
        rejected.extend(dict(row, motivo="Período inválido o solapado") for row in candidates.loc[list(invalid)].to_dict("records"))
        candidates = candidates.drop(index=list(invalid))
    covered = set()
    for row in candidates.itertuples():
        covered.update(pd.date_range(row.start_date, row.end_date).strftime("%Y-%m-%d"))
    selected_daily = daily[~daily["day"].isin(covered)].copy()
    covered.update(selected_daily["day"].tolist())
    return selected_daily, candidates, len(covered & expected), notes, rejected


@st.cache_data(show_spinner=False, max_entries=24)
def build_financial_report_v7(fin_version: str, inv_version: str, settings_token: tuple,
                              start: str, end: str, currency: str, sites: tuple[str, ...],
                              analysis_date: str | None = None) -> dict[str, Any]:
    """Contrato v6 + period_sales, analysis_date, notas y exclusiones auditables.

    ``sales`` contiene sólo días reales elegidos. ``period_sales`` conserva el
    rango y los importes originales. Ausencia de crédito, cobros o costos => N/D.
    Los campos históricos «al cierre» significan aquí cierre de analysis_date.
    """
    analysis_date = analysis_date or end
    start_day, end_day, cut_day = map(date.fromisoformat, (start, end, analysis_date))
    days = (end_day - start_day).days + 1
    sites = tuple(dict.fromkeys(sites))
    if days < 1 or not sites:
        raise ValueError("Selecciona sedes y un período válido.")
    if cut_day < end_day:
        raise ValueError("La fecha de análisis debe ser igual o posterior al fin del período de ventas.")
    if cut_day > date.today():
        raise ValueError("La fecha de análisis no puede ser futura.")
    settings = dict(settings_token)
    opening = (start_day - timedelta(days=1)).isoformat()
    result, details, chosen_daily, chosen_periods, excluded = [], [], [], [], []
    with db() as con:
        daily = pd.read_sql_query("SELECT * FROM financial_sales WHERE currency=? AND day BETWEEN ? AND ? ORDER BY day", con, params=(currency, start, end))
        periods = pd.read_sql_query("""SELECT * FROM v7_sales_batches
            WHERE currency=? AND scope='' AND start_date<=? AND end_date>=?
            ORDER BY site,start_date,end_date""", con, params=(currency, end, start))
        for site in sites:
            warnings, missing, notes = [], [], []
            sd, sp, covered_days, source_notes, rejected = _financial_sales_selection_v7(
                daily[daily["site"] == site], periods[periods["site"] == site], start, end)
            chosen_daily.append(sd)
            chosen_periods.append(sp)
            excluded.extend(rejected)
            notes.extend(source_notes)
            complete = covered_days == days
            if not complete:
                warnings.append(f"Ventas registradas: {covered_days}/{days} días cubiertos")
            # Opcionales se suman sólo si todas las fuentes seleccionadas los
            # reportaron; un cero declarado sigue siendo un cero válido.
            amounts = {field: complete_sum(pd.to_numeric(
                           pd.Series(sd[field].tolist() + sp[field].tolist(), dtype=object),
                           errors="coerce").replace([np.inf, -np.inf], np.nan))
                       for field in ["net_sales", "credit_sales", "cogs", "collections"]}
            for field, label in [("credit_sales", "Ventas a crédito no informadas: días de CxC N/D"),
                                 ("cogs", "Costo real de ventas no informado: margen y rotación N/D"),
                                 ("collections", "Cobros no informados")]:
                if pd.isna(amounts[field]):
                    missing.append(label)
            if complete and pd.isna(amounts["net_sales"]):
                complete = False
                warnings.append("Ventas netas incompletas en las fuentes seleccionadas")
            inv = pd.read_sql_query("""SELECT snapshot_date AS fecha,
                CASE WHEN COUNT(*)=COUNT(CASE WHEN source_value IS NOT NULL THEN source_value WHEN existence=0 THEN 0 WHEN cost_known=1 THEN existence*cost ELSE NULL END) THEN SUM(CASE WHEN source_value IS NOT NULL THEN source_value WHEN existence=0 THEN 0 WHEN cost_known=1 THEN existence*cost ELSE NULL END) ELSE NULL END AS valor,
                SUM(CASE WHEN source_value IS NULL AND existence<>0 AND cost_known=0 THEN 1 ELSE 0 END) AS sin_costo,
                SUM(CASE WHEN existence<0 OR COALESCE(source_value,existence*cost)<0
                    OR (source_value IS NULL AND cost<0) THEN 1 ELSE 0 END) AS negativos
                FROM inventory_snapshots WHERE site=? AND snapshot_date<=?
                GROUP BY snapshot_date ORDER BY snapshot_date""", con, params=(site, analysis_date))
            inv_date = str(inv.iloc[-1]["fecha"]) if not inv.empty else ""
            matching = settings.get(site, "") == currency
            inv_value = float(inv.iloc[-1]["valor"]) if not inv.empty and matching and pd.notna(inv.iloc[-1]["valor"]) else np.nan
            inv_clean = not inv.empty and int(inv.iloc[-1]["sin_costo"] + inv.iloc[-1]["negativos"]) == 0
            if inv.empty:
                missing.append("Inventario pendiente de cargar")
            elif not matching:
                missing.append("Confirma la moneda del inventario para compararlo con las ventas")
            elif not inv_clean:
                warnings.append("Inventario con costos faltantes o valores negativos")
            if inv_date and inv_date != analysis_date:
                notes.append(f"Inventario al {inv_date}; se necesita corte {analysis_date} para sus ratios")
            inv_close = inv_value if inv_date == analysis_date and inv_clean else np.nan
            inv_base, inv_method = inv_close, "No disponible"
            if pd.notna(inv_close):
                inv_method = ("Saldo final · aproximación" if analysis_date == end
                              else f"Saldo actual al {analysis_date} · aproximación sobre ventas históricas")
                inv_open = inv[inv["fecha"] == opening]
                if analysis_date == end and not inv_open.empty and matching and int(inv_open.iloc[0]["sin_costo"] + inv_open.iloc[0]["negativos"]) == 0:
                    inv_base = (float(inv_open.iloc[0]["valor"]) + inv_close) / 2
                    inv_method = "Promedio inicial/final"
            batch = con.execute("SELECT MAX(as_of) FROM financial_ar_batches WHERE site=? AND currency=? AND as_of<=?", (site, currency, analysis_date)).fetchone()[0]
            ar_value = ar_over = ar_base = ar_close = np.nan
            ar_method = "No disponible"
            if batch:
                ar = pd.read_sql_query("SELECT * FROM financial_ar WHERE site=? AND currency=? AND as_of=?", con, params=(site, currency, batch))
                aged = age_receivables(ar, batch)
                details.append(aged)
                ar_value = float(aged["balance"].sum())
                ar_over = float(aged.loc[aged["Días vencidos"] > 0, "balance"].sum())
                if batch == analysis_date:
                    ar_close = ar_value
                    ar_base = ar_value
                    ar_method = ("Saldo final · aproximación" if analysis_date == end
                                 else f"Saldo actual al {analysis_date} · aproximación sobre ventas históricas")
                    opening_batch = con.execute("SELECT 1 FROM financial_ar_batches WHERE site=? AND currency=? AND as_of=?", (site, currency, opening)).fetchone()
                    if analysis_date == end and opening_batch:
                        ar_initial = con.execute("SELECT COALESCE(SUM(balance),0) FROM financial_ar WHERE site=? AND currency=? AND as_of=?", (site, currency, opening)).fetchone()[0]
                        ar_base = (float(ar_initial) + ar_value) / 2
                        ar_method = "Promedio inicial/final"
                else:
                    notes.append(f"CxC al {batch}; se necesita corte {analysis_date} para sus ratios")
            else:
                missing.append("Cartera CxC pendiente de cargar")
            source = "Días reales y períodos A2" if not sd.empty and not sp.empty else "Períodos A2" if not sp.empty else "Días reales" if not sd.empty else "Sin datos"
            row = {"Sede": site, "Moneda": currency, "Días cargados": covered_days, "Ventas completas": complete,
                   "Fuente de ventas": source, "Días con detalle diario": len(sd), "Períodos utilizados": len(sp),
                   "Ventas netas": amounts["net_sales"], "Ventas crédito": amounts["credit_sales"],
                   "Costo de ventas": amounts["cogs"], "Cobros registrados": amounts["collections"],
                   "Inventario": inv_value, "Inventario al cierre": inv_close, "Base inventario": inv_base,
                   "CxC": ar_value, "CxC al cierre": ar_close, "Base CxC": ar_base, "CxC vencida": ar_over,
                   "Corte inventario": inv_date or "Sin datos", "Corte CxC": batch or "Sin datos",
                   "Fecha de análisis": analysis_date, "Método inventario": inv_method, "Método CxC": ar_method,
                   "Calidad de datos": " · ".join(warnings) or "Completa", "Datos pendientes": " · ".join(missing),
                   "Notas de cálculo": " · ".join(notes)}
            result.append(performance_ratios(row, days))
    summary = pd.DataFrame(result)
    total = {"Ventas completas": bool(summary["Ventas completas"].all())}
    for col in ["Ventas netas", "Ventas crédito", "Costo de ventas", "Cobros registrados", "Inventario", "Inventario al cierre", "Base inventario", "CxC", "CxC al cierre", "Base CxC", "CxC vencida"]:
        total[col] = complete_sum(summary[col])
    total = performance_ratios(total, days)
    return {"summary": summary, "total": total,
            "sales": pd.concat(chosen_daily, ignore_index=True),
            "period_sales": pd.concat(chosen_periods, ignore_index=True),
            "excluded_period_sales": pd.DataFrame(excluded),
            "ar": pd.concat(details, ignore_index=True) if details else pd.DataFrame(),
            "days": days, "start": start, "end": end, "currency": currency, "analysis_date": analysis_date}





# ============================================================
# NEXUS 7 · Reportes por gestión. Integrar antes de la interfaz.
# Utiliza exclusivamente las dependencias Python de la aplicación.
# ============================================================

REPORT_ID_COLUMNS_V7 = ["Sede", "Código", "Descripción", "Marca", "Proveedor", "Departamento", "Categoría"]
REPORT_META_COLUMNS_V7 = ["Marca", "Proveedor", "Departamento", "Categoría"]
REPORT_NOTES_V7 = {
    "Inventario": "Existencia del corte, límites vigentes y situación del almacén. N/D significa dato no informado.",
    "Ventas": "Incluye todas las filas del reporte de ventas, también productos fuera del inventario y pendientes sin código. Importes y costos informados se totalizan por moneda; no equivalen a cartera por cobrar.",
    "Compras": "Compra neta después de transferencias propuestas. El presupuesto es estimado, sin impuestos ni flete. N/D indica costo no disponible.",
    "Retiros": "Retiro neto = exceso menos transferencias de salida propuestas. Son recomendaciones; no movimientos confirmados.",
    "Redistribución": "Movimientos propuestos entre sedes. Su aprobación y ejecución deben registrarse en el sistema de origen.",
    "Alertas": "Problemas detectados en el corte. Un producto puede presentar varias alertas.",
}


def _report_num_v7(df, name, default=np.nan):
    if name not in df.columns:
        return pd.Series(default, index=df.index, dtype=float)
    return pd.to_numeric(df[name], errors="coerce").replace([np.inf, -np.inf], np.nan)


def _report_text_v7(value):
    if value is None or pd.isna(value):
        return "N/D"
    text = str(value).strip()
    return text if text else "N/D"


def _report_currency_v7(frame, column):
    values = frame.get(column, pd.Series("Sin moneda", index=frame.index)).astype(object).map(_report_text_v7)
    missing = values.str.casefold().isin(["n/d", "sin moneda", "sinmoneda", "sin asignar", "no definido", "no definida", "no especificada"])
    return values.mask(missing, "Sin moneda")


def _report_safe_cell_v7(value):
    """Excel/CSV treats an untrusted formula-looking identifier as text."""
    if isinstance(value, str):
        value = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", value)
        if value.lstrip().startswith(("=", "+", "-", "@")):
            return "'" + value
    return value


def _report_safe_frame_v7(frame):
    clean = frame.copy()
    for col in clean.select_dtypes(include=["object", "string"]).columns:
        clean[col] = clean[col].map(_report_safe_cell_v7)
    return clean


def _report_metadata_v7(frame, products, site_column="Sede"):
    """Enrich only exact, unambiguous product/site metadata matches."""
    result = frame.copy()
    fields = ["Descripción"] + REPORT_META_COLUMNS_V7
    for field in fields:
        if field not in result.columns:
            result[field] = ""
    if not result.empty and not products.empty and "Código" in result and "Código" in products:
        keys = ["Código"]
        if site_column in result and "Sede" in products:
            keys = [site_column, "Código"]
        source = products.copy()
        if site_column != "Sede" and "Sede" in source and len(keys) == 2:
            source = source.rename(columns={"Sede": site_column})
        # A missing code is a pending identity, never a key for joining products.
        source = source[~source["Código"].map(_report_text_v7).isin(["N/D", "Sin código", "Sin asignar"])]
        for field in fields:
            if field not in source:
                continue
            subset = source[keys + [field]].copy()
            subset[field] = subset[field].astype(object).map(_report_text_v7)
            subset = subset[~subset[field].isin(["N/D", "Sin asignar"])].drop_duplicates()
            # Conflicting catalogs stay unresolved instead of taking the first match.
            subset = subset[~subset.duplicated(keys, keep=False)]
            if subset.empty:
                continue
            lookup = subset.set_index(keys)[field]
            target_keys = pd.MultiIndex.from_frame(result[keys]) if len(keys) > 1 else result[keys[0]]
            found = pd.Series(lookup.reindex(target_keys).to_numpy(), index=result.index)
            missing = result[field].map(_report_text_v7).isin(["N/D", "Sin asignar"])
            result.loc[missing, field] = found.loc[missing]
    for field in fields:
        result[field] = result[field].astype(object).map(_report_text_v7)
        if field in REPORT_META_COLUMNS_V7:
            result[field] = result[field].replace("N/D", "Sin asignar")
    return result


def project_sales_report_v7(all_df):
    """Use the complete sales ledger when supplied, independently of inventory.

    all_df.attrs['v7_sales_rows'] must be a list of dicts, one per sales row/site/
    period. An explicit empty list represents zero sales; it does not fall back
    to inventory. Missing-code rows remain separate and retain their amounts.
    Filtering must happen on the returned frame, not merely on all_df, because
    pandas propagates attrs unchanged when its inventory rows are filtered.
    """
    products = all_df.copy() if all_df is not None else pd.DataFrame()
    has_ledger = all_df is not None and "v7_sales_rows" in all_df.attrs
    if has_ledger:
        ledger = all_df.attrs["v7_sales_rows"]
        if not isinstance(ledger, list) or any(not isinstance(row, dict) for row in ledger):
            raise ValueError("v7_sales_rows debe ser una lista de filas de ventas.")
        sales = pd.DataFrame(ledger)
    else:
        sales = products.copy()
    for column in ["Sede", "Código"]:
        if column not in sales:
            sales[column] = "N/D"
        sales[column] = sales[column].astype(object).map(_report_text_v7)
        if column in products:
            products[column] = products[column].astype(object).map(_report_text_v7)
    missing_code = sales["Código"].isin(["N/D", "Sin código", "Sin asignar"])
    sales.loc[missing_code, "Código"] = "Sin código"
    sales = _report_metadata_v7(sales, products)
    sales["Ventas del período"] = _report_num_v7(sales, "Ventas")
    sales["Moneda ventas"] = _report_currency_v7(sales, "Moneda ventas")
    columns = REPORT_ID_COLUMNS_V7 + ["Ventas del período"]
    columns += [col for col in ["Demanda Mensual", "Rotación estimada (x/mes)", "ABC", "Sin movimiento"] if col in sales]
    money_columns = ["Importe neto", "Monto bruto", "Descuentos", "IVA reportado", "Costo de ventas"]
    for column in money_columns:
        if column in sales:
            sales[column] = _report_num_v7(sales, column)
            columns.append(column)
    columns += ["Moneda ventas"]
    if "Período ventas" in sales:
        columns.append("Período ventas")
    if missing_code.any():
        sales["Identificación"] = np.where(missing_code, "Pendiente: sin código en origen", "Con código")
        columns.append("Identificación")
        columns += [col for col in ["Fila de ventas", "Archivo de ventas"] if col in sales]
    return sales.reindex(columns=columns).reset_index(drop=True)


def project_report_frames(all_df, transfers, alerts):
    """Public contract: ordered, complete area projections; never mutate inputs."""
    x = all_df.copy() if all_df is not None else pd.DataFrame()
    for col in REPORT_ID_COLUMNS_V7:
        if col not in x:
            x[col] = "N/D"
        x[col] = x[col].astype(object).map(_report_text_v7)
        if col in REPORT_META_COLUMNS_V7:
            x[col] = x[col].replace("N/D", "Sin asignar")
    x["Moneda inventario"] = _report_currency_v7(x, "Moneda inventario")
    if "Fecha inventario" not in x:
        x["Fecha inventario"] = x.get("Fecha del inventario", pd.Series("N/D", index=x.index))
    x["Fecha inventario"] = x["Fecha inventario"].astype(object).map(_report_text_v7)
    if "Política de stock" not in x:
        x["Política de stock"] = x.get("Política stock inventario", pd.Series("N/D", index=x.index))
    x["Política de stock"] = x["Política de stock"].astype(object).map(_report_text_v7)
    x["Déficit hasta mínimo"] = (_report_num_v7(x, "Stock Mínimo") - _report_num_v7(x, "Existencia")).clip(lower=0)
    purchase_source = "Compra Ajustada" if "Compra Ajustada" in x else "Compra Sugerida"
    x["Compra neta"] = _report_num_v7(x, purchase_source)
    x["Transferencias recibidas"] = _report_num_v7(x, "Transferencias Recibidas", 0).fillna(0)
    x["Transferencias de salida"] = _report_num_v7(x, "Transferencias Enviadas", 0).fillna(0)
    x["Exceso"] = _report_num_v7(x, "Retiro Almacén")
    x["Retiro neto"] = (x["Exceso"] - x["Transferencias de salida"]).clip(lower=0)
    cost = _report_num_v7(x, "Costo")
    if "Tiene costo" in x:
        cost = cost.mask(x["Tiene costo"].astype(str).str.casefold().eq("no"))
    cost = cost.mask(cost.lt(0))
    x["Costo unitario"] = cost
    x["Presupuesto estimado"] = (x["Compra neta"] * cost).round(2)
    x["Valor del retiro"] = (x["Retiro neto"] * cost).round(2)
    calculated_value = (_report_num_v7(x, "Existencia") * cost).round(2)
    # A2's reported amount retains precision that a displayed two-decimal cost loses.
    reported_value = _report_num_v7(x, "Valor inventario A2")
    reported_value = reported_value.where(reported_value.notna(), _report_num_v7(x, "Valor Inventario ($)"))
    x["Valor inventario"] = reported_value.where(reported_value.notna(), calculated_value)
    x["Ventas del período"] = _report_num_v7(x, "Ventas")

    def select(data, columns):
        out = data.reindex(columns=columns).copy()
        for col in REPORT_ID_COLUMNS_V7:
            if col in out:
                out[col] = out[col].map(_report_text_v7)
        return out.reset_index(drop=True)

    frames = {
        "Inventario": select(x, REPORT_ID_COLUMNS_V7 + [
            "Fecha inventario", "Existencia", "Ventas utilizadas", "Período ventas", "Días del historial",
            "Método de promedio", "Meses utilizados", "Demanda Mensual", "Cobertura mínima aplicada", "Cobertura máxima aplicada",
            "Mínimo automático", "Máximo automático", "Stock Mínimo", "Stock Máximo", "Política de stock",
            "Revisión del cálculo", "Advertencia de plazo", "Ventas sin código del historial", "Punto de Reorden", "Cobertura (meses)", "Valor inventario", "Moneda inventario", "Estado"]),
        "Ventas": project_sales_report_v7(all_df),
        "Compras": select(x[x["Compra neta"].gt(0)], REPORT_ID_COLUMNS_V7 + [
            "Fecha inventario", "Existencia", "Déficit hasta mínimo", "Transferencias recibidas", "Compra neta", "Costo unitario", "Presupuesto estimado", "Moneda inventario", "Política de stock", "Prioridad"]),
        "Retiros": select(x[x["Retiro neto"].gt(0)], REPORT_ID_COLUMNS_V7 + [
            "Fecha inventario", "Existencia", "Stock Máximo", "Exceso", "Transferencias de salida", "Retiro neto", "Costo unitario", "Valor del retiro", "Moneda inventario", "Política de stock"]),
    }
    movement = _report_metadata_v7(transfers if transfers is not None else pd.DataFrame(), x, "Origen")
    frames["Redistribución"] = select(movement, ["Código", "Descripción"] + REPORT_META_COLUMNS_V7 + [
        "Origen", "Destino", "Unidades Sugeridas", "Costo Unitario ($)", "Compra Evitada Estimada ($)"])
    problems = _report_metadata_v7(alerts if alerts is not None else pd.DataFrame(), x)
    frames["Alertas"] = select(problems, REPORT_ID_COLUMNS_V7 + ["Tipo", "Severidad", "Detalle"])
    return frames


def filter_report_frames_v7(frames, selections):
    result = {}
    for section, frame in frames.items():
        mask = pd.Series(True, index=frame.index)
        for field, chosen in selections.items():
            if not chosen:
                continue
            if field == "Sede" and section == "Redistribución":
                mask &= frame["Origen"].isin(chosen) | frame["Destino"].isin(chosen)
            elif field in frame:
                mask &= frame[field].isin(chosen)
        result[section] = frame.loc[mask].reset_index(drop=True)
    return result


def _report_summary_v7(frames):
    rows = []
    metrics = {
        "Inventario": [("Existencia", "Unidades en almacén"), ("Valor inventario", "Valor de inventario estimado")],
        "Ventas": [("Ventas del período", "Unidades vendidas del período"), ("Importe neto", "Ventas netas informadas"), ("Costo de ventas", "Costo de ventas informado")],
        "Compras": [("Compra neta", "Unidades por comprar"), ("Presupuesto estimado", "Presupuesto estimado")],
        "Retiros": [("Retiro neto", "Unidades por retirar"), ("Valor del retiro", "Valor estimado del retiro")],
        "Redistribución": [("Unidades Sugeridas", "Unidades por transferir")],
        "Alertas": [],
    }
    for section, df in frames.items():
        rows.append([section, "Registros", len(df), "Productos/sede" if section == "Inventario" else "Filas por período/sede" if section == "Ventas" else "Filas"])
        for col, label in metrics.get(section, []):
            if col in df:
                money_columns = {"Importe neto", "Costo de ventas", "Valor inventario", "Presupuesto estimado", "Valor del retiro"}
                if col in money_columns:
                    currencies = _report_currency_v7(df, "Moneda ventas" if section == "Ventas" else "Moneda inventario")
                    for currency in sorted(currencies.unique()):
                        values = pd.to_numeric(df.loc[currencies.eq(currency), col], errors="coerce")
                        missing = int(values.isna().sum())
                        if currency in {"Sin moneda", "Sin asignar"}:
                            rows.append([section, f"{label} · {currency}", np.nan, "Moneda pendiente: importes sin consolidar"])
                        else:
                            rows.append([section, f"{label} · {currency}", values.sum(min_count=1), f"Parcial: {missing} sin dato" if missing else "Completo"])
                    continue
                values = pd.to_numeric(df[col], errors="coerce")
                value = values.sum(min_count=1) if len(df) else 0
                missing = int(values.isna().sum())
                rows.append([section, label, value, f"Parcial: {missing} sin dato" if missing else "Completo"])
    return pd.DataFrame(rows, columns=["Gestión", "Indicador", "Valor", "Cobertura del dato"])


def _report_scope_v7(frames):
    values = {}
    for field in ["Sede", "Marca", "Proveedor", "Departamento"]:
        found = set()
        for df in frames.values():
            if field in df:
                found.update(df[field].dropna().astype(str))
            if field == "Sede" and "Origen" in df:
                found.update(df["Origen"].dropna().astype(str))
                found.update(df["Destino"].dropna().astype(str))
        ordered = sorted(found)
        values[field] = ", ".join(ordered[:4]) + (f" y {len(ordered)-4} más" if len(ordered) > 4 else "") if ordered else "Sin registros"
    return " · ".join(f"{key}: {value}" for key, value in values.items())


def _report_management_groups_v7(frames, field):
    inventory = frames.get("Inventario", pd.DataFrame())
    columns = [field, "Moneda inventario", "Productos/sede", "Existencia", "Valor inventario", "Compra neta", "Presupuesto estimado", "Retiro neto", "Alertas", "Cobertura monetaria"]
    if inventory.empty or field not in inventory:
        return pd.DataFrame(columns=columns)
    inventory = inventory.copy()
    inventory["Moneda inventario"] = _report_currency_v7(inventory, "Moneda inventario")
    keys = [field, "Moneda inventario"]
    grouped = inventory.groupby(keys, dropna=False, observed=True).agg(**{
        "Productos/sede": ("Código", "size"), "Existencia": ("Existencia", "sum"),
        "Valor inventario": ("Valor inventario", lambda values: values.sum(min_count=1)),
    })
    unknown = grouped.index.get_level_values("Moneda inventario") == "Sin moneda"
    grouped.loc[unknown, "Valor inventario"] = np.nan
    for section, columns in [("Compras", ["Compra neta", "Presupuesto estimado"]), ("Retiros", ["Retiro neto"])]:
        detail = frames.get(section, pd.DataFrame()).copy()
        detail["Moneda inventario"] = _report_currency_v7(detail, "Moneda inventario")
        for column in columns:
            if not detail.empty:
                amounts = detail.groupby(keys, dropna=False)[column].sum(min_count=1)
                grouped[column] = amounts.reindex(grouped.index)
                no_actions = ~grouped.index.isin(amounts.index)
                grouped.loc[no_actions, column] = 0
            else:
                grouped[column] = 0
            if column == "Presupuesto estimado":
                grouped.loc[unknown, column] = np.nan
    # Currency does not belong to an alert; join it by exact inventory identity
    # to avoid copying the same department's alert count into every currency.
    alerts = frames.get("Alertas", pd.DataFrame()).copy()
    if not alerts.empty:
        lookup_keys = ["Sede", "Código"]
        lookup = inventory[lookup_keys + ["Moneda inventario"]].drop_duplicates()
        lookup = lookup[~lookup.duplicated(lookup_keys, keep=False)]
        alerts = alerts.merge(lookup, on=lookup_keys, how="left", validate="many_to_one")
        alerts["Moneda inventario"] = _report_currency_v7(alerts, "Moneda inventario")
        grouped["Alertas"] = alerts.groupby(keys, dropna=False, observed=True).size().reindex(grouped.index).fillna(0).astype(int)
    else:
        grouped["Alertas"] = 0
    grouped["Cobertura monetaria"] = np.where(unknown, "Moneda pendiente: importes no sumables", "Importes de la moneda indicada; N/D = sin dato")
    return grouped.reset_index()


def build_report_excel_v7(frames, snapshot_label, management=False):
    """One sheet per area, full row counts, typed values and useful print setup."""
    from openpyxl import Workbook
    from openpyxl.chart import BarChart, Reference
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo

    wb = Workbook()
    wb.remove(wb.active)
    summary = _report_summary_v7(frames)
    managerial = {"Resumen por sede": _report_management_groups_v7(frames, "Sede"), "Resumen por departamento": _report_management_groups_v7(frames, "Departamento")} if management else {}
    sheets = {"Resumen gerencial" if management else "Resumen": summary, **managerial, **frames}
    for index, (section, frame) in enumerate(sheets.items()):
        ws = wb.create_sheet(section[:31])
        count = max(len(frame.columns), 6)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=count)
        ws.cell(1, 1, f"NEXUS · {section.upper()}").font = Font(name="Calibri", bold=True, size=19, color="FFFFFF")
        ws.cell(1, 1).fill = PatternFill("solid", fgColor="163743")
        ws.row_dimensions[1].height = 38
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=count)
        ws.cell(2, 1, _report_safe_cell_v7(f"Corte: {snapshot_label} | {len(frame):,} registros | N/D = no informado")).font = Font(size=10, color="617584")
        ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=count)
        note = REPORT_NOTES_V7.get(section, "Cada gestión conserva su detalle completo en una hoja independiente. Los totales monetarios parciales indican costos faltantes.")
        ws.cell(3, 1, note).alignment = Alignment(wrap_text=True, vertical="center")
        ws.cell(3, 1).font = Font(size=10, color="617584")
        ws.row_dimensions[3].height = 30
        ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=count)
        ws.cell(4, 1, _report_safe_cell_v7(_report_scope_v7(frames))).alignment = Alignment(wrap_text=True, vertical="center")
        ws.cell(4, 1).font = Font(size=9, color="617584")
        ws.row_dimensions[4].height = 28
        safe = _report_safe_frame_v7(frame)
        for j, col in enumerate(safe.columns, 1):
            cell = ws.cell(6, j, col)
            cell.font = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
            cell.fill = PatternFill("solid", fgColor="187C78")
            cell.alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[6].height = 31
        for i, row in enumerate(safe.itertuples(index=False, name=None), 7):
            for j, value in enumerate(row, 1):
                if pd.isna(value):
                    value = "N/D"
                elif isinstance(value, (np.integer, np.floating)):
                    value = value.item()
                cell = ws.cell(i, j, value)
                cell.font = Font(name="Calibri", size=10, color="243B49")
                cell.alignment = Alignment(vertical="center", wrap_text=True)
                if isinstance(value, (int, float)):
                    cell.number_format = '#,##0.000000' if safe.columns[j-1] in {'Meses utilizados','Demanda Mensual','Cobertura mínima aplicada','Cobertura máxima aplicada'} else '#,##0.00;[Red](#,##0.00);"–"'
            ws.row_dimensions[i].height = 30
        for j, col in enumerate(safe.columns, 1):
            width = 52 if col in {"Descripción", "Detalle"} else 25 if col in REPORT_META_COLUMNS_V7 or col in {"Estado", "Indicador", "Cobertura del dato"} else 19
            ws.column_dimensions[get_column_letter(j)].width = width
        if len(safe):
            ref = f"A6:{get_column_letter(len(safe.columns))}{6+len(safe)}"
            table = Table(displayName=f"NexusArea{index}", ref=ref)
            table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
            ws.add_table(table)
        else:
            ws.cell(7, 1, "Sin registros para esta gestión y filtros.")
        ws.freeze_panes = "D7"
        ws.sheet_view.showGridLines = False
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A3
        ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
        ws.print_title_rows = "1:6"
        ws.print_options.horizontalCentered = True
        ws.oddFooter.center.text = "NEXUS · Página &P de &N"
        ws.print_area = f"A1:{get_column_letter(count)}{max(7, 6 + len(safe))}"
    # A compact chart is placed beside (not on top of) the summary data.
    inv = frames.get("Inventario", pd.DataFrame())
    if not inv.empty:
        chart_data = inv.groupby("Sede", dropna=False)["Existencia"].sum(min_count=1).reset_index()
        ws = wb.worksheets[0]
        ws.cell(6, 7, "Sede"); ws.cell(6, 8, "Unidades en almacén")
        for i, row in enumerate(chart_data.itertuples(index=False, name=None), 7):
            ws.cell(i, 7, _report_safe_cell_v7(row[0])); ws.cell(i, 8, float(row[1]) if pd.notna(row[1]) else None)
        chart = BarChart()
        chart.title = "Existencia por sede"
        chart.y_axis.title = "Unidades"
        chart.add_data(Reference(ws, min_col=8, min_row=6, max_row=6+len(chart_data)), titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=7, min_row=7, max_row=6+len(chart_data)))
        chart.height, chart.width = 8, 17
        ws.add_chart(chart, "G12")
        ws.print_area = f"A1:P{max(29, 6+len(summary))}"
    out = io.BytesIO(); wb.save(out)
    return out.getvalue()


def _report_figures_v7(frames):
    figures = []
    inventory = frames.get("Inventario", pd.DataFrame())
    if not inventory.empty:
        inv = inventory.copy()
        for col in ["Sede", "Estado"]:
            inv[col] = inv[col].map(lambda value: escape(_report_text_v7(value)))
        site = inv.groupby("Sede", dropna=False)["Existencia"].sum(min_count=1).reset_index()
        figures.append(px.bar(site, x="Sede", y="Existencia", title="Existencia por sede", color_discrete_sequence=["#187C78"]))
        states = inv["Estado"].value_counts().rename_axis("Estado").reset_index(name="Productos")
        figures.append(px.pie(states, names="Estado", values="Productos", hole=.62, title="Estado del inventario", color_discrete_sequence=["#187C78", "#EF982D", "#DF5555", "#3298AE"]))
    actions = []
    for name, column in [("Compras", "Compra neta"), ("Retiros", "Retiro neto")]:
        frame = frames.get(name, pd.DataFrame())
        if not frame.empty:
            grouped = frame.groupby("Sede", dropna=False)[column].sum(min_count=1)
            actions.extend({"Sede": escape(str(site)), "Gestión": name, "Unidades": value} for site, value in grouped.items())
    if actions:
        figures.append(px.bar(pd.DataFrame(actions), x="Sede", y="Unidades", color="Gestión", barmode="group", title="Unidades recomendadas por gestión", color_discrete_sequence=["#EF982D", "#187C78"]))
    sales = frames.get("Ventas", pd.DataFrame())
    if not sales.empty:
        history = sales.groupby("Sede", dropna=False)["Ventas del período"].sum(min_count=1).reset_index()
        history["Sede"] = history["Sede"].astype(str).map(escape)
        figures.append(px.bar(history, x="Sede", y="Ventas del período", title="Unidades vendidas por sede · período cargado", color_discrete_sequence=["#3298AE"]))
        if "Importe neto" in sales and "Moneda ventas" in sales:
            for currency, entries in sales.groupby("Moneda ventas", dropna=False, sort=True, observed=True):
                if _report_text_v7(currency) in {"N/D", "Sin moneda", "Sin asignar"}:
                    continue
                values = entries.groupby("Sede", dropna=False)["Importe neto"].sum(min_count=1).reset_index().dropna(subset=["Importe neto"])
                if not values.empty:
                    values["Sede"] = values["Sede"].astype(str).map(escape)
                    figures.append(px.bar(values, x="Sede", y="Importe neto", title=f"Ventas netas por sede · {escape(str(currency))}", labels={"Importe neto": escape(str(currency))}, color_discrete_sequence=["#187C78"]))
    movement = frames.get("Redistribución", pd.DataFrame())
    if not movement.empty:
        routes = movement.dropna(subset=["Origen", "Destino", "Unidades Sugeridas"])
        routes = routes[pd.to_numeric(routes["Unidades Sugeridas"], errors="coerce").gt(0)]
        if not routes.empty:
            routes = routes.groupby(["Origen", "Destino"], dropna=False)["Unidades Sugeridas"].sum().reset_index()
            nodes = sorted(set(routes["Origen"]).union(routes["Destino"]))
            indices = {name: i for i, name in enumerate(nodes)}
            sankey = go.Figure(go.Sankey(node=dict(label=[escape(str(name)) for name in nodes], color="#187C78", pad=24), link=dict(source=routes["Origen"].map(indices).tolist(), target=routes["Destino"].map(indices).tolist(), value=routes["Unidades Sugeridas"].tolist(), color="rgba(50,152,174,.3)")))
            sankey.update_layout(title="Redistribución propuesta entre sedes")
            figures.append(sankey)
    retirement = frames.get("Retiros", pd.DataFrame())
    if not retirement.empty:
        retirement = retirement.copy()
        retirement["Moneda inventario"] = _report_currency_v7(retirement, "Moneda inventario")
        for currency, entries in retirement.groupby("Moneda inventario", observed=True, dropna=False):
            if currency == "Sin moneda":
                continue
            top = entries.dropna(subset=["Valor del retiro"]).nlargest(10, "Valor del retiro").sort_values("Valor del retiro").copy()
            if not top.empty:
                top["Producto"] = top["Código"].astype(str) + " · " + top["Descripción"].astype(str).str.slice(0, 42)
                top["Producto"] = top["Producto"].map(escape)
                figures.append(px.bar(top, x="Valor del retiro", y="Producto", orientation="h", title=f"Mayor valor en retiros propuestos · {escape(str(currency))}", labels={"Valor del retiro": escape(str(currency))}, color_discrete_sequence=["#3298AE"]))
    problems = frames.get("Alertas", pd.DataFrame())
    if not problems.empty:
        counts = problems["Severidad"].value_counts().rename_axis("Severidad").reset_index(name="Alertas")
        counts["Severidad"] = counts["Severidad"].astype(str).map(escape)
        figures.append(px.bar(counts, x="Severidad", y="Alertas", title="Alertas por severidad", color_discrete_sequence=["#DF5555"]))
    for fig in figures:
        fig.update_layout(template="plotly_white", height=340, margin=dict(l=15, r=15, t=50, b=35), font=dict(family="Arial", color="#243B49"))
    return figures


def build_report_html_v7(frames, snapshot_label):
    plots = []
    for i, fig in enumerate(_report_figures_v7(frames)):
        plots.append('<div class="chart">' + pio.to_html(fig, full_html=False, include_plotlyjs=True if i == 0 else False, config={"responsive": True, "displaylogo": False}) + '</div>')
    buttons, panels = [], []
    for i, (name, df) in enumerate(frames.items()):
        buttons.append(f'<button class="tab{" active" if i == 0 else ""}" data-panel="area{i}" onclick="openArea(this)">{escape(name)} <span>{len(df):,}</span></button>')
        table = df.to_html(index=False, escape=True, na_rep="N/D", border=0, classes="data", float_format=lambda v: f"{v:,.2f}") if not df.empty else '<p class="empty">Sin registros para esta gestión y filtros.</p>'
        panels.append(f'<section id="area{i}" class="area" {"" if i == 0 else "hidden"}><h2>{escape(name)}</h2><p class="muted">{escape(REPORT_NOTES_V7.get(name, ""))}</p><label>Buscar en esta gestión <input type="search" placeholder="Código, descripción, marca…" oninput="filterRows(this)"></label><div class="table-wrap">{table}</div><p class="muted count">{len(df):,} registros completos</p></section>')
    summary = _report_summary_v7(frames)
    cards = []
    for _, row in summary[summary["Indicador"].ne("Registros")].head(6).iterrows():
        value = "N/D" if pd.isna(row["Valor"]) else f'{row["Valor"]:,.2f}'
        cards.append(f'<article class="metric"><small>{escape(row["Indicador"])}</small><strong>{value}</strong><span>{escape(row["Cobertura del dato"])}</span></article>')
    html = '''<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>NEXUS · Reporte por gestión</title><style>
    *{box-sizing:border-box}body{margin:0;background:#f3f6f7;color:#243b49;font:14px/1.5 Arial,sans-serif}.wrap{max-width:1500px;margin:auto;padding:28px}header{background:#163743;color:white;padding:30px;border-radius:18px}header p{margin:5px 0;color:#cfe5e5}h1{margin:6px 0;font-size:29px}.eyebrow{text-transform:uppercase;letter-spacing:2px;font-size:11px}.metrics{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin:20px 0}.metric,.chart,.area{background:white;border:1px solid #dfebeb;border-radius:14px;padding:20px}.metric small,.metric span{display:block;color:#617584}.metric strong{display:block;font-size:29px;color:#187c78}.plots{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px;margin-bottom:22px}.chart{overflow:hidden;padding:8px}.tabs{display:flex;gap:8px;flex-wrap:wrap;margin:18px 0}.tab{background:#fff;border:1px solid #cddede;padding:12px 18px;border-radius:9px;font-weight:700;color:#24404a;cursor:pointer}.tab.active{background:#187c78;color:white}.tab span{opacity:.72;margin-left:8px}.table-wrap{overflow:auto;max-height:680px;margin-top:18px}.data{border-collapse:collapse;width:100%;font-size:12px}.data th{position:sticky;top:0;background:#187c78;color:white;z-index:1;text-align:left;white-space:normal;min-width:110px}.data td,.data th{padding:10px;border-bottom:1px solid #e7eeee}.data td{min-width:100px;max-width:350px}.data tr:nth-child(even){background:#f5f9f9}.muted,footer{color:#617584}.area h2{margin-top:0}input{padding:10px;border:1px solid #bed4d4;border-radius:7px;margin-left:12px;min-width:260px}footer{margin:24px 0;font-size:12px}.empty{padding:28px;background:#f4f8f8;border-radius:10px}[hidden]{display:none!important}@media(max-width:800px){.metrics,.plots{grid-template-columns:1fr}.wrap{padding:12px}input{display:block;margin:8px 0;min-width:0;width:100%}}@media print{body{background:white}.wrap{max-width:none}.tabs,input,label{display:none}.area[hidden]{display:block!important}.area{break-before:page}.table-wrap{max-height:none;overflow:visible}.data{font-size:8px}.data th,.data td{padding:4px;min-width:0}.plots{display:none}}
    </style></head><body><main class="wrap">'''
    html += f'<header><span class="eyebrow">Makropetrol · Centro de reportes</span><h1>La información de cada gestión</h1><p>Corte {escape(str(snapshot_label))}</p><p>{escape(_report_scope_v7(frames))}</p></header><div class="metrics">{"".join(cards)}</div><div class="plots">{"".join(plots)}</div><nav class="tabs">{"".join(buttons)}</nav>{"".join(panels)}<footer>Datos completos en cada sección. N/D = no informado. El inventario conserva su moneda de carga. Los totales de inventario, compras, retiros y ventas se separan por moneda; los importes sin moneda no se consolidan. No se realizan conversiones. Las propuestas no equivalen a órdenes ejecutadas.</footer>'
    html += '''</main><script>
    function openArea(button){document.querySelectorAll('.area').forEach(p=>p.hidden=true);document.getElementById(button.dataset.panel).hidden=false;document.querySelectorAll('.tab').forEach(b=>b.classList.remove('active'));button.classList.add('active')}
    function filterRows(input){let panel=input.closest('.area'),term=input.value.toLocaleLowerCase(),shown=0;panel.querySelectorAll('tbody tr').forEach(row=>{row.hidden=!row.textContent.toLocaleLowerCase().includes(term);if(!row.hidden)shown++});panel.querySelector('.count').textContent=shown.toLocaleString()+' registros visibles'}
    </script></body></html>'''
    return html.encode("utf-8")


def build_report_pdf_v7(frames, snapshot_label, detailed=False):
    if not HAS_PDF:
        return None
    # Per-product cards wrap every field; avoid tiny 14-column landscape tables.
    # Core Helvetica supports Spanish. Unsupported glyphs are transliterated.
    def text(value):
        raw = _report_text_v7(value).replace("→", " > ").replace("—", "-").replace("–", "-")
        return raw.encode("latin-1", "replace").decode("latin-1")

    class ReportPDF(FPDF):
        def footer(self):
            self.set_y(-12); self.set_font("Helvetica", "", 8); self.set_text_color(97, 117, 132)
            self.cell(0, 5, f"NEXUS | {text(snapshot_label)} | Pagina {self.page_no()}/{{nb}}", align="R")

    pdf = ReportPDF(orientation="P", unit="mm", format="A4")
    pdf.alias_nb_pages(); pdf.set_auto_page_break(auto=False); pdf.set_margins(12, 12, 12)

    def wrapped(value, width, font_size=9):
        pdf.set_font("Helvetica", "", font_size)
        words = text(value).split()
        lines, line = [], ""
        for word in words:
            if pdf.get_string_width(word) > width:
                # Break unusually long identifiers without losing characters.
                chunks, chunk = [], ""
                for letter in word:
                    if chunk and pdf.get_string_width(chunk + letter) > width:
                        chunks.append(chunk); chunk = ""
                    chunk += letter
                words_part = chunks + [chunk]
            else:
                words_part = [word]
            for part in words_part:
                candidate = f"{line} {part}".strip()
                if line and pdf.get_string_width(candidate) > width:
                    lines.append(line); line = part
                else:
                    line = candidate
        if line: lines.append(line)
        return lines or [""]

    current_section = "Resumen"
    def page(section, note=""):
        pdf.add_page(); pdf.set_fill_color(22, 55, 67); pdf.rect(0, 0, 210, 32, "F")
        pdf.set_xy(12, 8); pdf.set_font("Helvetica", "B", 17); pdf.set_text_color(255, 255, 255)
        pdf.cell(186, 9, text(f"NEXUS | {section}"))
        pdf.set_xy(12, 20); pdf.set_font("Helvetica", "", 9); pdf.cell(186, 5, text(f"Corte: {snapshot_label}"))
        pdf.set_xy(12, 38); pdf.set_text_color(97, 117, 132)
        for line in wrapped(note, 183, 8):
            pdf.cell(186, 4, line, new_x="LMARGIN", new_y="NEXT")
        pdf.ln(3)

    def draw_lines(lines, fill=False, bold=False, size=9):
        for line in lines:
            if pdf.get_y() + 4.5 > 278:
                page(current_section, "Continuación. N/D = no informado.")
            pdf.set_font("Helvetica", "B" if bold else "", size)
            pdf.set_text_color(24, 61, 71)
            pdf.cell(186, 4.5, line, new_x="LMARGIN", new_y="NEXT", fill=fill)

    page("Resumen por gestión", _report_scope_v7(frames))
    summary = _report_summary_v7(frames)
    for _, row in summary.iterrows():
        value = "N/D" if pd.isna(row["Valor"]) else f'{row["Valor"]:,.2f}'
        draw_lines(wrapped(f'{row["Gestión"]} | {row["Indicador"]}: {value} | {row["Cobertura del dato"]}', 184))
    pdf.ln(5)
    draw_lines(wrapped("El inventario conserva su moneda de carga. Inventario, compras, retiros y ventas se totalizan por moneda; los importes sin moneda no se consolidan. Presupuestos sin impuestos ni flete. Las recomendaciones todavía no son movimientos ejecutados.", 184, 8), size=8)
    if not detailed:
        pdf.ln(4)
        draw_lines(wrapped("Informe ejecutivo: hasta 20 registros por gestión, ordenados por relevancia. Los totales anteriores incluyen todas las filas de la selección. Excel, HTML, CSV y JSON conservan el detalle completo. El PDF detallado está disponible desde la aplicación.", 184, 9))
    for section, df in frames.items():
        current_section = section
        note = REPORT_NOTES_V7.get(section, "")
        preview = df
        if not detailed:
            sort_columns = {"Inventario": "Valor inventario", "Ventas": "Ventas del período", "Compras": "Presupuesto estimado", "Retiros": "Valor del retiro", "Redistribución": "Unidades Sugeridas"}
            sort_by = sort_columns.get(section)
            if section in {"Inventario", "Compras", "Retiros"}:
                currencies = _report_currency_v7(df, "Moneda inventario")
                if currencies.nunique() > 1 or currencies.eq("Sin moneda").any():
                    sort_by = {"Inventario": "Existencia", "Compras": "Compra neta", "Retiros": "Retiro neto"}[section]
            if sort_by in df:
                preview = df.sort_values(sort_by, ascending=False, na_position="last", kind="stable")
            elif section == "Alertas" and not df.empty:
                rank = df["Severidad"].map({"CRÍTICA": 0, "ALTA": 1, "MEDIA": 2, "BAJA": 3}).fillna(4)
                preview = df.loc[rank.sort_values(kind="stable").index]
            preview = preview.head(20)
            note += f" Mostrando {len(preview)} de {len(df):,} registros. Detalle completo en Excel/HTML/CSV/JSON."
        else:
            note += f" Detalle completo: {len(df):,} registros."
        page(section, note)
        if df.empty:
            draw_lines(["Sin registros para esta gestión y filtros."])
            continue
        for number, record in enumerate(preview.to_dict("records"), 1):
            headline = f'{number:04d} | {record.get("Código", "N/D")} | {record.get("Descripción", "N/D")}'
            meta = " | ".join(f"{field}: {record.get(field, 'N/D')}" for field in REPORT_META_COLUMNS_V7)
            metrics = []
            for key, value in record.items():
                if key in REPORT_META_COLUMNS_V7 or key in {"Código", "Descripción"}:
                    continue
                if isinstance(value, (float, np.floating)) and pd.notna(value):
                    value = f"{value:,.2f}"
                metrics.append(f"{key}: {_report_text_v7(value)}")
            hlines, mlines, dlines = wrapped(headline, 183, 9), wrapped(meta, 183, 8), wrapped(" | ".join(metrics), 183, 9)
            height = (len(hlines)+len(mlines)+len(dlines))*4.5 + 5
            if pdf.get_y() + min(height, 225) > 278:
                page(section, note)
            pdf.set_fill_color(233, 244, 243)
            draw_lines(hlines, fill=True, bold=True)
            draw_lines(mlines, size=8)
            draw_lines(dlines)
            pdf.ln(4)
    content = pdf.output()
    return content.encode("latin-1") if isinstance(content, str) else bytes(content)


def build_report_package_v7(frames, snapshot_label):
    slug = safe_slug(snapshot_label)
    operational = build_report_excel_v7(frames, snapshot_label)
    management = build_report_excel_v7(frames, snapshot_label, management=True)
    html = build_report_html_v7(frames, snapshot_label)
    pdf = build_report_pdf_v7(frames, snapshot_label)
    json_bytes = json.dumps({"corte": str(snapshot_label), "alcance": _report_scope_v7(frames), "secciones": {name: json.loads(df.to_json(orient="records", force_ascii=False)) for name, df in frames.items()}}, ensure_ascii=False, indent=2).encode("utf-8")
    csv_out = io.BytesIO()
    with zipfile.ZipFile(csv_out, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, df in frames.items():
            archive.writestr(f"{safe_slug(name)}_{slug}.csv", _report_safe_frame_v7(df).to_csv(index=False).encode("utf-8-sig"))
    bundle_out = io.BytesIO()
    with zipfile.ZipFile(bundle_out, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(f"NEXUS_Operativo_{slug}.xlsx", operational)
        archive.writestr(f"NEXUS_Gerencial_{slug}.xlsx", management)
        archive.writestr(f"NEXUS_Visual_{slug}.html", html)
        archive.writestr(f"NEXUS_Secciones_{slug}.json", json_bytes)
        for name, df in frames.items():
            archive.writestr(f"CSV/{safe_slug(name)}_{slug}.csv", _report_safe_frame_v7(df).to_csv(index=False).encode("utf-8-sig"))
        if pdf:
            archive.writestr(f"NEXUS_Ejecutivo_{slug}.pdf", pdf)
        archive.writestr("LEEME.txt", "NEXUS · Reportes por gestión\nCorte: " + str(snapshot_label) + "\n" + _report_scope_v7(frames) + "\n\n" + "\n".join(f"{name}: {REPORT_NOTES_V7.get(name, '')}" for name in frames) + "\n\nExcel, HTML, CSV y JSON conservan todos los registros del filtro, incluidas las ventas pendientes sin código. PDF ejecutivo: totales completos y hasta 20 registros por gestión, con el límite señalado en cada sección. N/D = no informado; Sin asignar = clasificación pendiente. Las cantidades de compra y retiro son netas de las transferencias propuestas. Las propuestas no equivalen a órdenes aprobadas. El inventario conserva su moneda de carga. Inventario, compras, retiros y ventas se totalizan por moneda; los importes sin moneda permanecen sin consolidar. Los CSV neutralizan fórmulas en textos; JSON conserva los valores originales.")
    return {"operativo": operational, "gerencial": management, "html": html, "pdf": pdf, "csv": csv_out.getvalue(), "json": json_bytes, "zip": bundle_out.getvalue()}


def _render_report_downloads_v7(frames, snapshot_label, key):
    digest = hashlib.sha256(str(snapshot_label).encode())
    for name, frame in frames.items():
        digest.update(name.encode()); digest.update(frame.to_json(orient="split", force_ascii=False).encode())
    token = digest.hexdigest()[:24]
    state_key = f"reports_v7_{key}"
    if st.button("Preparar archivos de esta selección", type="primary", key=f"prepare_{key}"):
        with st.spinner("Preparando los reportes por gestión…"):
            st.session_state[state_key] = (token, build_report_package_v7(frames, snapshot_label))
    ready = st.session_state.get(state_key)
    if not ready or ready[0] != token:
        st.caption("La descarga incluirá todas las filas de las gestiones y filtros seleccionados.")
        return
    package = ready[1]
    formats = [("Excel operativo", "operativo", "xlsx", DOC_MIME_XLSX), ("Excel gerencial", "gerencial", "xlsx", DOC_MIME_XLSX), ("PDF ejecutivo", "pdf", "pdf", DOC_MIME_PDF), ("Reporte visual HTML", "html", "html", DOC_MIME_HTML), ("CSV por gestión (ZIP)", "csv", "zip", "application/zip"), ("JSON por gestión", "json", "json", "application/json"), ("Paquete completo", "zip", "zip", "application/zip")]
    columns = st.columns(3)
    for index, (label, identifier, extension, mime) in enumerate(formats):
        if package.get(identifier):
            columns[index % 3].download_button(label, package[identifier], f"NEXUS_{safe_slug(key)}_{identifier}_{safe_slug(snapshot_label)}.{extension}", mime, key=f"download_{key}_{identifier}", width="stretch")
        elif identifier == "pdf":
            columns[index % 3].info("PDF disponible al instalar fpdf2.")
    st.caption("El PDF ejecutivo presenta totales completos y hasta 20 registros por gestión. Los otros formatos contienen todas las filas.")
    if HAS_PDF:
        with st.expander("PDF detallado con todos los registros"):
            st.caption("Puede generar muchas páginas. Usa los filtros para seleccionar la sede y gestión necesarias.")
            detail_key = state_key + "_detail"
            if st.button("Preparar PDF detallado", key=f"prepare_detail_{key}"):
                with st.spinner("Generando todas las páginas del detalle…"):
                    st.session_state[detail_key] = (token, build_report_pdf_v7(frames, snapshot_label, detailed=True))
            detail = st.session_state.get(detail_key)
            if detail and detail[0] == token:
                st.download_button("Descargar PDF detallado", detail[1], f"NEXUS_Detallado_{safe_slug(snapshot_label)}.pdf", DOC_MIME_PDF, key=f"download_detail_{key}")


def render_reports_v7(all_df, transfers, alerts, snapshot_label):
    st.header("Reportes por gestión")
    st.caption("Elige el área, aplica tus filtros y descarga únicamente la información que necesitas.")
    frames = project_report_frames(all_df, transfers, alerts)
    widgets = st.columns(4)
    filters = {}
    for widget, field in zip(widgets, ["Sede", "Proveedor", "Marca", "Departamento"]):
        options = set()
        for frame in frames.values():
            if field in frame:
                options.update(frame[field].dropna().astype(str))
            elif field == "Sede" and "Origen" in frame:
                options.update(frame["Origen"].dropna().astype(str)); options.update(frame["Destino"].dropna().astype(str))
        if field == "Sede":
            filters[field] = site_selector_v712("Sedes del informe", f"report_filter_v7_{field}", options=sorted(options), container=widget)
        else:
            filters[field] = widget.multiselect(field, sorted(options), key=f"report_filter_v7_{field}", placeholder="Todos")
    if not filters["Sede"]:
        st.info("Selecciona al menos una sede para preparar el informe. Limpiar la selección no incluye todas automáticamente.")
        return
    scoped = filter_report_frames_v7(frames, filters)
    section = st.selectbox("¿Qué gestión necesitas?", ["Reporte general"] + list(frames), key="report_section_v7")
    chosen = scoped if section == "Reporte general" else {section: scoped[section]}
    summary = _report_summary_v7(chosen)
    kpis = summary[summary["Indicador"].ne("Registros")].head(4)
    if not kpis.empty:
        for widget, (_, row) in zip(st.columns(len(kpis)), kpis.iterrows()):
            widget.metric(row["Indicador"], "N/D" if pd.isna(row["Valor"]) else f'{row["Valor"]:,.2f}')
            if row["Cobertura del dato"] != "Completo":
                widget.caption(row["Cobertura del dato"])
    figures = _report_figures_v7(chosen)
    if figures:
        with st.expander("Ver gráficas", expanded=True):
            for i, fig in enumerate(figures):
                st.plotly_chart(fig, width="stretch", key=f"report_plot_v7_{i}")
    if section == "Reporte general":
        tabs = st.tabs(list(chosen))
        for tab, (name, frame) in zip(tabs, chosen.items()):
            with tab:
                st.caption(REPORT_NOTES_V7.get(name, ""))
                st.dataframe(frame, hide_index=True, width="stretch", height=380)
                st.caption(f"{len(frame):,} filas completas")
    else:
        st.caption(REPORT_NOTES_V7.get(section, ""))
        st.dataframe(chosen[section], hide_index=True, width="stretch", height=430)
        st.caption(f"{len(chosen[section]):,} filas completas")
    _render_report_downloads_v7(chosen, snapshot_label, "centro")


def download_section_v7(section, all_df, transfers, alerts, snapshot_label, filters=None):
    """Pass explicit filters for sales ledgers; filtering all_df leaves attrs intact.

    filters has the same {field: [values]} contract as filter_report_frames_v7.
    Callers can alternatively provide all_df.attrs['v7_report_filters'].
    """
    frames = project_report_frames(all_df, transfers, alerts)
    selected = filters if filters is not None else all_df.attrs.get("v7_report_filters", {})
    if selected:
        frames = filter_report_frames_v7(frames, selected)
    if section not in frames:
        return
    with st.expander(f"Descargar reporte de {section.lower()}"):
        st.caption(REPORT_NOTES_V7.get(section, ""))
        _render_report_downloads_v7({section: frames[section]}, snapshot_label, safe_slug(section))





@st.cache_data(show_spinner=False, max_entries=10)
def cached_a2_v7(raw: bytes, name: str, kind: str) -> dict:
    parsed = parse_a2_report(raw, name, kind)
    parsed["metadata"]["sha256"] = hashlib.sha256(raw).hexdigest()
    return parsed


def detect_cxc_v7(raw: bytes, name: str) -> tuple[pd.DataFrame, dict, pd.DataFrame]:
    """Detecta encabezados reales de una cartera tabular; no infiere saldos de ventas."""
    grid = read_a2_grid(raw, name)
    variants = {field: {normalize_key(a) for a in aliases + [field]} for field, aliases in FIN_AR_FIELDS.items()}
    variants["Factura"] |= {"n documento", "no factura", "numero de factura", "nro documento"}
    variants["Cliente"] |= {"nombre del cliente", "razon social del cliente"}
    best, anchors, score = 0, {}, 0
    for idx, row in grid.head(100).iterrows():
        found = {}
        for col, value in row.items():
            key = normalize_key(value)
            for field, names in variants.items():
                if key in names and field not in found:
                    found[field] = int(col)
        if len(found) > score:
            best, anchors, score = int(idx), found, len(found)
    if not {"Factura", "Saldo"}.issubset(anchors):
        raise ValueError("No pude reconocer factura y saldo en este archivo. Abre 'Otros formatos' para asignar sus columnas. El formato exacto de CxC de A2 aún requiere una muestra.")
    positions = sorted(int(col) for col, value in grid.iloc[best].items() if not pd.isna(value) and str(value).strip())
    records = []
    for _, row in grid.iloc[best + 1:].iterrows():
        record = {}
        for field, start in anchors.items():
            end = next((x for x in positions if x > start), len(row))
            values = [v for v in row.iloc[start:end] if not pd.isna(v) and str(v).strip()]
            record[field] = values[0] if values else np.nan
        invoice = _a2_text(record.get("Factura"))
        if not invoice and not any(_a2_text(record.get(field)) for field in ["Cliente", "Emisión", "Vencimiento", "Saldo"]):
            continue
        if normalize_key(invoice).startswith("total") and not _a2_text(record.get("Cliente")):
            continue
        record["Factura"] = invoice
        records.append(record)
    frame = pd.DataFrame(records)
    mapping = {c: c if c in frame else "" for c in FIN_AR_FIELDS}
    return frame, mapping, grid


def _parse_upload_card_v7(upload, kind):
    if upload is None:
        return None, None
    try:
        parsed = cached_a2_v7(upload.getvalue(), upload.name, kind)
        meta = parsed["metadata"]
        if meta.get("duplicate_codes"):
            return None, "Hay códigos repetidos en el reporte; revisa la exportación antes de guardarla."
        if meta.get("reconciled") is False:
            return None, "Los totales extraídos no coinciden con los totales impresos; el archivo requiere revisión."
        return parsed, None
    except Exception as exc:
        return None, str(exc)


def _navigation_v7(target):
    st.session_state["nexus_view"] = target
    st.rerun()





def render_import_status_v7():
    st.markdown("**Datos que ya están guardados**")
    rows = []
    with db() as con:
        for site in SEDES:
            inv = con.execute("SELECT MAX(snapshot_date),COUNT(DISTINCT snapshot_date) FROM inventory_snapshots WHERE site=?", (site,)).fetchone()
            sales = con.execute("SELECT MAX(end_date),COUNT(*) FROM v7_sales_batches WHERE site=? AND scope=''", (site,)).fetchone()
            ar = con.execute("SELECT MAX(as_of) FROM financial_ar_batches WHERE site=?", (site,)).fetchone()[0]
            rows.append({"Sede": site, "Inventario": inv[0] or "Pendiente", "Ventas hasta": sales[0] or "Sin períodos A2", "Períodos guardados": sales[1], "CxC": ar or "Pendiente"})
    st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch")
    settings = finance_settings()
    missing_currency = [row["Sede"] for row in rows if row["Inventario"] != "Pendiente" and not settings.get(row["Sede"])]
    if missing_currency:
        with st.expander("Completar la moneda de inventarios anteriores", expanded=False):
            st.caption("Declara la moneda de los costos ya guardados. Esto no convierte precios ni altera las existencias.")
            with st.form("v7_old_inventory_currency"):
                chosen = {site: st.selectbox(f"Moneda del inventario · {site}", ["Por confirmar"] + FIN_CURRENCIES, key=f"v7_old_currency_{site}") for site in missing_currency}
                if st.form_submit_button("Guardar monedas confirmadas"):
                    with db() as con:
                        for site, currency in chosen.items():
                            if currency != "Por confirmar":
                                con.execute("INSERT OR REPLACE INTO financial_settings VALUES (?,?)", (site, currency))
                        touch_v7(con)
                    st.cache_data.clear()
                    st.rerun()


def render_saved_history_v7():
    with db() as con:
        batches = pd.read_sql_query("SELECT * FROM v7_sales_batches WHERE scope='' ORDER BY site,start_date", con)
    if batches.empty:
        return
    with st.expander("Cambiar el historial usado para mínimos y compras", expanded=False):
        selected_site = st.selectbox("Sede del historial", sorted(batches["site"].unique()), key="v7_history_site")
        available = batches[batches["site"].eq(selected_site)]
        with db() as con:
            config = con.execute("SELECT sales_start,sales_end FROM v7_site_config WHERE site=? AND scope=''", (selected_site,)).fetchone()
        initial = tuple(date.fromisoformat(d) for d in config) if config else (date.fromisoformat(available["start_date"].min()), date.fromisoformat(available["end_date"].max()))
        selected = st.date_input("Fechas del historial guardado", initial, key=f"v7_saved_period_{selected_site}_{v7_version()}")
        st.dataframe(available[["start_date", "end_date", "currency", "file_name", "total_units"]].rename(columns={"start_date":"Desde","end_date":"Hasta","currency":"Moneda","file_name":"Archivo","total_units":"Unidades"}), hide_index=True, width="stretch")
        st.caption("Se utilizan únicamente los reportes enteros incluidos entre las fechas elegidas. No se dividen ventas acumuladas para inventar movimientos mensuales.")
        if st.button("Usar este historial", key="v7_use_history"):
            chosen = available[available["start_date"].ge(selected[0].isoformat()) & available["end_date"].le(selected[1].isoformat())] if len(selected) == 2 else pd.DataFrame()
            if chosen.empty:
                st.error("Selecciona un intervalo que incluya por completo al menos un reporte guardado.")
            else:
                with db() as con:
                    con.execute("UPDATE v7_site_config SET sales_start=?,sales_end=? WHERE site=? AND scope=''", (selected[0].isoformat(), selected[1].isoformat(), selected_site))
                    touch_v7(con)
                st.cache_data.clear()
                st.rerun()
    with st.expander("Completar códigos de ventas guardadas", expanded=False):
        with db() as con:
            pending = pd.read_sql_query("SELECT b.site AS Sede,b.start_date||' a '||b.end_date AS Período,i.description AS Descripción,i.units AS Unidades,i.code AS Código,i.batch_id,i.row_no FROM v7_sales_items i JOIN v7_sales_batches b ON b.id=i.batch_id WHERE TRIM(i.code)='' ORDER BY b.site,b.start_date,i.row_no", con)
            known = {r[0] for r in con.execute("SELECT DISTINCT code FROM inventory_snapshots UNION SELECT DISTINCT code FROM supplier_inventory_snapshots")}
        if pending.empty:
            st.success("No quedan ventas sin código.")
        else:
            st.caption("Introduce el código real tal como aparece en inventario. El importe original de la venta se conserva.")
            edited = st.data_editor(pending.drop(columns=["batch_id","row_no"]), disabled=["Sede","Período","Descripción","Unidades"], hide_index=True, width="stretch", key="v7_saved_pending")
            if st.button("Guardar códigos completados", key="v7_resolve_saved"):
                new_codes = edited["Código"].fillna("").astype(str).str.strip()
                invalid = [code for code in new_codes if code and code not in known]
                if invalid:
                    st.error("Estos códigos no están en un inventario cargado: " + ", ".join(invalid[:10]))
                else:
                    with db() as con:
                        for i, code in enumerate(new_codes):
                            if code:
                                con.execute("UPDATE v7_sales_items SET code=? WHERE batch_id=? AND row_no=?", (code, int(pending.iloc[i]["batch_id"]), int(pending.iloc[i]["row_no"])))
                        touch_v7(con)
                    st.cache_data.clear()
                    st.rerun()


def render_catalogs_v7():
    departments, suppliers = catalog_frame_v7("departamentos"), catalog_frame_v7("proveedores")
    a, b = st.columns(2)
    with a:
        kpi("Departamentos guardados", str(len(departments)), "Precargados desde tu reporte A2", "blue")
    with b:
        kpi("Proveedores guardados", str(len(suppliers)), "Con sus códigos y estados originales", "cyan")
    st.info("Tus archivos listan departamentos y proveedores, pero no indican qué proveedor corresponde a cada departamento ni la clasificación de cada artículo. Completa esas relaciones una vez; la app las reutiliza automáticamente.")
    with st.expander("Ver o actualizar los catálogos de A2", expanded=False):
        left, right = st.columns(2)
        with left:
            st.dataframe(departments, hide_index=True, width="stretch")
        with right:
            st.dataframe(suppliers, hide_index=True, width="stretch")
        source = st.file_uploader("Nuevo reporte de departamentos o proveedores", type=["xls", "xlsx", "csv"], key="v7_catalog_upload")
        if source:
            try:
                parsed = cached_a2_v7(source.getvalue(), source.name, "auto")
                if parsed["kind"] not in {"departamentos", "proveedores"}:
                    st.warning("Este espacio recibe únicamente los catálogos.")
                else:
                    st.dataframe(parsed["data"], hide_index=True, width="stretch")
                    if st.button("Actualizar catálogo", key="v7_update_catalog"):
                        save_catalog_v7(parsed)
                        st.success("Catálogo actualizado; las relaciones existentes se conservaron.")
            except Exception as exc:
                st.error(str(exc))
    st.markdown("**Relaciona departamento, proveedor y marca**")
    st.caption("Las sugerencias de texto y marca son editables. Sólo se aplican las filas que marques como Confirmada. Usa ; para varias expresiones de búsqueda. Si un producto coincide con varias reglas, queda sin asignar hasta que lo revises.")
    suggestions = catalog_rule_suggestions_v7()
    with st.form("v7_rules_form"):
        rules = st.data_editor(suggestions, hide_index=True, width="stretch", disabled=["Departamento"],
            column_config={"Proveedor": st.column_config.SelectboxColumn(options=[""] + suppliers["Descripción"].tolist()),
                           "Confirmada": st.column_config.CheckboxColumn("Confirmada")}, key="v7_rules_editor")
        if st.form_submit_button("Guardar relaciones y actualizar productos"):
            with db() as con:
                for _, row in rules.iterrows():
                    texts = ["" if pd.isna(row[c]) else str(row[c]).strip() for c in ["Proveedor", "Marca", "Categoría", "Texto en descripción"]]
                    if texts[0] and texts[0] not in suppliers["Descripción"].tolist():
                        raise ValueError("El proveedor debe estar en tu catálogo.")
                    con.execute("UPDATE v7_department_rules SET supplier=?,brand=?,category=?,match_text=?,confirmed=? WHERE department=?", tuple(texts) + (int(bool(row["Confirmada"])), row["Departamento"]))
                touch_v7(con)
            st.cache_data.clear()
            st.success("Relaciones guardadas. Las siguientes vistas ya usan esta clasificación.")
    with st.expander("Asignar un artículo específico o corregir una excepción", expanded=False):
        with db() as con:
            products = pd.read_sql_query("SELECT code AS Código,MAX(description) AS Descripción FROM inventory_snapshots GROUP BY code ORDER BY code", con)
        if products.empty:
            st.caption("Carga primero el inventario para buscar productos por su código.")
        else:
            search = st.text_input("Buscar artículo", key="v7_catalog_search").strip()
            if search:
                choices = products[products["Código"].str.contains(search, case=False, regex=False) | products["Descripción"].str.contains(search, case=False, regex=False)].head(100)
                if choices.empty:
                    st.info("No se encontraron artículos.")
                else:
                    by_code = choices.set_index("Código")["Descripción"].to_dict()
                    code = st.selectbox("Artículo", choices["Código"].tolist(), format_func=lambda x: f"{x} · {by_code[x]}", key="v7_catalog_product")
                    with st.form("v7_product_dimensions"):
                        dept = st.selectbox("Departamento del artículo", [""] + departments["Descripción"].tolist())
                        supplier = st.selectbox("Proveedor del artículo · opcional", [""] + suppliers["Descripción"].tolist())
                        brand = st.text_input("Marca del artículo · opcional")
                        category = st.text_input("Categoría del artículo · opcional")
                        if st.form_submit_button("Guardar clasificación del artículo"):
                            with db() as con:
                                con.execute("INSERT OR REPLACE INTO v7_product_dimensions VALUES (?,?,?,?,?,?)", (code, by_code[code], dept, supplier, brand, category))
                                touch_v7(con)
                            st.cache_data.clear()
                            st.success("Clasificación guardada para las próximas cargas de ese código.")


def filter_products_v7(data: pd.DataFrame, key: str) -> pd.DataFrame:
    result = data.copy()
    cols = st.columns(4)
    for col, field in zip(cols, ["Sede", "Departamento", "Proveedor", "Marca"]):
        if field not in result:
            continue
        values = result[field].astype(object).fillna("").astype(str).replace("", "Sin asignar")
        if field == "Sede":
            selected = site_selector_v712("Sedes", f"{key}_Sede_v712", options=sorted(values.unique()), container=col)
            result = result[values.isin(selected)]
        else:
            choices = ["Todos"] + sorted(values.unique())
            picked = col.selectbox(field, choices, key=f"{key}_{field}")
            if picked != "Todos":
                result = result[values.eq(picked)]
    search = st.text_input("Buscar código o descripción", key=f"{key}_search").strip()
    if search:
        result = result[result["Código"].astype(str).str.contains(search, case=False, regex=False) | result["Descripción"].astype(str).str.contains(search, case=False, regex=False)]
    return result


def render_stock_levels_v7(all_df):
    st.subheader("Mínimos y máximos")
    st.caption("La app propone niveles según tus ventas. Puedes fijar valores propios por producto y sede; quedarán guardados para las próximas cargas.")
    if all_df.empty:
        empty_state("Primero carga un inventario", "En Carga de datos encontrarás el espacio para subirlo.")
        return
    data = filter_products_v7(all_df, "v7_levels")
    if data.empty:
        st.info("No hay productos para esos filtros.")
        return
    render_stock_calculation_v73(data)
    columns = ["Sede", "Código", "Descripción", "Ventas utilizadas", "Meses utilizados", "Demanda Mensual", "Existencia", "Mínimo automático", "Máximo automático", "Stock Mínimo", "Stock Máximo", "Política de stock", "Revisión del cálculo"]
    st.dataframe(data[[c for c in columns if c in data]].head(500), hide_index=True, width="stretch")
    with st.expander("Límites manuales · valores fijos que sustituyen las ventas", expanded=False):
        with st.form("v7_stock_bulk"):
            a, b = st.columns(2)
            low = a.number_input("Mínimo en unidades", min_value=0, value=1, step=1)
            high = b.number_input("Máximo en unidades", min_value=0, value=2, step=1)
            mode = st.radio("Modo", ["Volver al cálculo automático", "Fijar estos límites"], horizontal=True)
            if st.form_submit_button(f"Aplicar a los {len(data):,} productos filtrados"):
                if high < low:
                    st.error("El máximo debe ser igual o mayor que el mínimo.")
                else:
                    save_levels_v7(data, low, high, automatic=mode.startswith("Volver"))
                    st.rerun()
    st.markdown("**Editar límites por producto**")
    st.caption("Los límites manuales usan el mínimo como punto de reorden y el máximo como objetivo. Desmarca 'Manual' para recuperar las propuestas del motor.")
    editable = data[["Sede", "Código", "Descripción", "Stock Mínimo", "Stock Máximo"]].head(300).copy()
    editable[["Stock Mínimo", "Stock Máximo"]] = editable[["Stock Mínimo", "Stock Máximo"]].fillna(0)
    editable["Manual"] = data.head(300)["Política de stock"].eq("Manual").values
    with st.form("v7_stock_editor_form"):
        changed = st.data_editor(editable, hide_index=True, width="stretch", disabled=["Sede", "Código", "Descripción"],
            column_config={"Stock Mínimo": st.column_config.NumberColumn(min_value=0,step=1,required=True),
                           "Stock Máximo": st.column_config.NumberColumn(min_value=0,step=1,required=True)}, key="v7_stock_editor")
        if st.form_submit_button("Guardar mínimos y máximos"):
            invalid = changed["Manual"] & (changed["Stock Mínimo"].isna() | changed["Stock Máximo"].isna() | (changed["Stock Máximo"] < changed["Stock Mínimo"]) | (changed["Stock Mínimo"] < 0))
            if invalid.any():
                st.error("Revisa los límites: mínimo desde cero y máximo igual o mayor que el mínimo.")
            else:
                with db() as con:
                    for _, row in changed.iterrows():
                        if row["Manual"]:
                            con.execute("INSERT OR REPLACE INTO v7_stock_levels VALUES(?,?,?,?,?)", (row["Sede"], row["Código"], float(row["Stock Mínimo"]), float(row["Stock Máximo"]), datetime.now().isoformat()))
                        else:
                            con.execute("DELETE FROM v7_stock_levels WHERE site=? AND code=?", (row["Sede"], row["Código"]))
                    touch_v7(con)
                st.cache_data.clear()
                st.rerun()
    if len(data) > 300:
        st.caption("El editor muestra 300 productos; usa los filtros para editar otros. La aplicación por grupo y los reportes conservan todos.")
    download_section_v7("Inventario", data, pd.DataFrame(), pd.DataFrame(), "Mínimos y máximos")


def save_levels_v7(data, low, high, automatic=False):
    if low < 0 or high < low:
        raise ValueError("Límites inválidos.")
    with db() as con:
        for _, row in data.iterrows():
            if automatic:
                con.execute("DELETE FROM v7_stock_levels WHERE site=? AND code=?", (row["Sede"], row["Código"]))
            else:
                con.execute("INSERT OR REPLACE INTO v7_stock_levels VALUES(?,?,?,?,?)", (row["Sede"], row["Código"], float(low), float(high), datetime.now().isoformat()))
        touch_v7(con)
    st.cache_data.clear()


def render_inventory_simple_v7(all_df):
    st.subheader("Inventario")
    st.caption("Existencias y niveles de almacén, con clasificación de cada producto.")
    if all_df.empty:
        empty_state("Carga tu primer inventario", "Ve a Carga de datos y sube el Excel de A2.")
        return
    data = filter_products_v7(all_df, "v7_inv")
    cols = st.columns(4)
    with cols[0]:
        kpi("Productos", integer(len(data)), "En el filtro actual", "blue")
    with cols[1]:
        kpi("Existencia", integer(data["Existencia"].sum()), "Unidades en almacén", "cyan")
    with cols[2]:
        kpi("Bajo mínimo", integer((data["Existencia"] < data["Stock Mínimo"]).sum()), "Productos para revisar", "orange")
    with cols[3]:
        kpi("Sobre máximo", integer((data["Existencia"] > data["Stock Máximo"]).sum()), "Productos con exceso", "red")
    frame = project_report_frames(data, pd.DataFrame(), pd.DataFrame())["Inventario"]
    st.dataframe(frame.head(750), hide_index=True, width="stretch", height=540)
    if not data.empty:
        by_dept = data.assign(Departamento=data["Departamento"].astype(object).fillna("").replace("", "Sin asignar")).groupby("Departamento", observed=True)["Existencia"].sum().nlargest(15).reset_index()
        fig = px.bar(by_dept, x="Departamento", y="Existencia", color_discrete_sequence=[NEXUS_BLUE], title="Existencias por departamento")
        st.plotly_chart(chart_layout(fig, 310), width="stretch")
    if st.button("Configurar mínimos y máximos", key="v7_go_levels"):
        _navigation_v7("Mínimos y Máximos")
    download_section_v7("Inventario", data, pd.DataFrame(), pd.DataFrame(), latest_display)


def render_purchases_simple_v7(all_df, transfers, alerts):
    st.subheader("Compras")
    st.caption("Necesidades de compra después de considerar las redistribuciones propuestas.")
    if all_df.empty:
        empty_state("Primero carga inventario y ventas", "La app calculará la reposición con los días reales del período.")
        return
    data = filter_products_v7(all_df, "v7_buy")
    frame = project_report_frames(data, transfers, alerts)["Compras"]
    if frame.empty:
        st.success("No hay compras pendientes para los filtros actuales.")
    else:
        cols = st.columns(3)
        with cols[0]:
            kpi("Productos a comprar", integer(len(frame)), "Compra neta", "orange")
        with cols[1]:
            kpi("Unidades a comprar", integer(frame["Compra neta"].sum()), "Después de redistribuciones", "blue")
        with cols[2]:
            kpi("Proveedores", integer(frame["Proveedor"].nunique()), "Incluye pendientes de asignación", "cyan")
        st.dataframe(frame.head(750), hide_index=True, width="stretch", height=540)
        grouped = frame.groupby("Proveedor", observed=True)["Compra neta"].sum().nlargest(12).reset_index()
        fig = px.bar(grouped, x="Proveedor", y="Compra neta", color_discrete_sequence=[NEXUS_ORANGE], title="Unidades por proveedor")
        st.plotly_chart(chart_layout(fig, 310), width="stretch")
    download_section_v7("Compras", data, transfers, alerts, latest_display)


# NEXUS 7 · Interfaz y documentos financieros con períodos reales.
# Integrar después de period_finance.py y de los helpers de interfaz v6.

FIN_METHODOLOGY_V7 = """CÓMO LEER EL ANÁLISIS
• Ventas netas: importe sin IVA, después de descuentos y devoluciones, del período que declara el reporte.
• Un archivo A2 de varios meses conserva un único rango real. No se reparte artificialmente entre días o meses. Para una evolución mensual se necesitan reportes separados por mes.
• Cuando existe una serie diaria completa se usa esa serie. En otro caso se utilizan períodos completos que caben en los filtros y días reales fuera de esos períodos, sin duplicar importes. Los rangos recortados o solapados se excluyen y se informan.
• La cobertura completa exige que las fuentes elegidas cubran todos los días del período. Un día ausente no se convierte en cero. Con cobertura parcial se muestran las ventas registradas y se bloquean los ratios que requieren todo el período.
• Ventas a crédito y cobros son datos distintos. Si A2 no los informa permanecen N/D; no se supone venta de contado ni se calcula cartera restando cobros a ventas.
• El crédito utilizado para días de CxC debe tener la misma base de impuestos que los saldos de cartera. Cada moneda se analiza por separado, sin conversiones automáticas.
• Costo de ventas: costo real reportado de la mercancía vendida. No se sustituye por el costo unitario actual multiplicado por unidades.
• Margen bruto = ventas netas − costo real de ventas. No descuenta gastos operativos ni impuestos sobre la renta.
• Inventario: valor reportado por A2 cuando existe; en otro caso, existencia × costo. La moneda se confirma por sede. Los costos faltantes y valores negativos impiden los ratios afectados.
• La fecha de análisis de inventario y CxC es independiente del período de ventas y debe ser igual o posterior a su fin. Un corte anterior se muestra con su fecha, pero no activa los ratios de la fecha solicitada.
• Si la fecha de análisis coincide con el fin de ventas y existe el corte exacto del día anterior al inicio, la base es el promedio de esos dos saldos. Sin corte inicial exacto, se usa el saldo final como aproximación.
• Si la fecha de análisis es posterior al fin de ventas, se usa únicamente el saldo actual exacto como aproximación sobre ventas históricas. No se promedian saldos de extremos de períodos distintos.
• Rotación del período = costo de ventas / base de inventario. Días de inventario = base de inventario / costo de ventas × días del período. La rotación no se anualiza.
• Días de CxC = base de cartera / ventas a crédito × días del período. Es una estimación; no mide el tiempo real medio de pago de cada factura. Sin ventas a crédito reportadas, permanece N/D.
• CxC es una fotografía completa de facturas y saldos pendientes. Un corte vacío confirmado equivale a cartera cero; ausencia de archivo no equivale a cero. Nunca se suman fotografías de distintas fechas.
• La antigüedad se calcula al corte real de CxC, no a una fecha posterior sin nuevos datos. Vencer en la fecha del corte sigue siendo estar al día. Los cobros pueden corresponder a ventas de otros períodos.
• Ciclo operativo = días de inventario + días de CxC. Para obtener ciclo de caja faltan los plazos de pago a proveedores.
• Inventario + CxC al corte de análisis requiere ambos saldos válidos en esa fecha. No representa dinero disponible ni capital de trabajo neto.
• N/D identifica datos ausentes o denominadores no positivos. Los totales con sedes o campos faltantes conservan N/D; los porcentajes globales se recalculan a partir de las bases, no promediando porcentajes de sedes.
• Integridad de fuentes indica si los datos elegidos son consistentes. Datos pendientes explica qué indicadores adicionales se activan al completar campos opcionales. Las metas de seguimiento son configurables.

QUÉ CARGAR PARA CxC
En Carga de datos, sube la cartera completa de la sede e indica su moneda y fecha de corte. Se requiere una fila por factura con: factura, cliente, emisión, vencimiento y saldo pendiente. El importe original es opcional. Para calcular días de CxC se necesitan además ventas a crédito reales del mismo período y con la misma base de impuestos.
"""


def _finance_frames_v7(report: dict) -> dict[str, pd.DataFrame]:
    """Separa indicadores, fuentes y movimientos conservando datos verificables."""
    summary = report["summary"].copy()
    if "Fecha de análisis" not in summary:
        summary["Fecha de análisis"] = report.get("analysis_date", report["end"])
    indicator_cols = ["Sede", "Moneda", "Ventas netas", "Ventas crédito", "Costo de ventas", "Margen bruto", "Margen bruto %",
                      "Cobros registrados", "Inventario", "CxC", "CxC vencida", "Cartera vencida %", "Días inventario",
                      "Días CxC", "Rotación período", "Ciclo operativo", "Capital en inventario y CxC", "Estado", "Acción sugerida"]
    source_cols = ["Sede", "Fecha de análisis", "Corte inventario", "Corte CxC", "Fuente de ventas", "Días cargados", "Días con detalle diario",
                   "Períodos utilizados", "Ventas completas", "Método inventario", "Método CxC", "Calidad de datos", "Datos pendientes", "Notas de cálculo"]
    daily = report.get("sales", pd.DataFrame()).copy()
    periods = report.get("period_sales", pd.DataFrame()).copy()
    ar = report.get("ar", pd.DataFrame()).copy()
    daily_cols = ["day", "site", "currency", "net_sales", "credit_sales", "cogs", "collections"]
    period_cols = ["site", "currency", "start_date", "end_date", "total_units", "net_sales", "gross_sales", "discounts", "tax", "cogs", "credit_sales", "collections", "file_name"]
    names = {"day": "Fecha", "site": "Sede", "currency": "Moneda", "net_sales": "Ventas netas", "credit_sales": "Ventas crédito",
             "cogs": "Costo de ventas", "collections": "Cobros", "start_date": "Desde", "end_date": "Hasta", "total_units": "Unidades",
             "gross_sales": "Ventas brutas", "discounts": "Descuentos", "tax": "IVA", "file_name": "Archivo de origen",
             "client": "Cliente", "invoice": "Factura", "issued": "Emisión", "due": "Vencimiento", "balance": "Saldo", "amount": "Importe factura"}
    frames = {
        "Indicadores": summary[[c for c in indicator_cols if c in summary]],
        "Fuentes y cortes": summary[[c for c in source_cols if c in summary]].rename(columns={"Días cargados": "Días cubiertos", "Calidad de datos": "Integridad de fuentes"}),
        "Ventas por período": periods.reindex(columns=period_cols).rename(columns=names),
        "Ventas diarias reales": daily.reindex(columns=daily_cols).rename(columns=names),
        "Cartera CxC": ar.reindex(columns=["site", "currency", "client", "invoice", "issued", "due", "balance", "amount", "Días vencidos", "Tramo", "Corte CxC"]).rename(columns=names),
    }
    rejected = report.get("excluded_period_sales", pd.DataFrame())
    if not rejected.empty:
        frames["Períodos excluidos"] = rejected.reindex(columns=period_cols + ["motivo"]).rename(columns={**names, "motivo": "Motivo de exclusión"})
    return frames


def _finance_daily_plot_v7(report: dict):
    sales = report.get("sales", pd.DataFrame())
    if sales.empty:
        return None
    pieces = []
    for site, rows in sales.groupby("site"):
        daily = rows.set_index("day").reindex(pd.date_range(report["start"], report["end"]).strftime("%Y-%m-%d"))
        daily.index.name = "Fecha"
        daily["Sede"] = site
        pieces.append(daily.reset_index()[["Fecha", "Sede", "net_sales", "collections"]])
    plot = pd.concat(pieces, ignore_index=True).rename(columns={"net_sales": "Ventas netas", "collections": "Cobros"})
    plot = plot.melt(id_vars=["Fecha", "Sede"], value_vars=["Ventas netas", "Cobros"], var_name="Concepto", value_name="Importe")
    fig = px.line(plot, x="Fecha", y="Importe", color="Sede", line_dash="Concepto", markers=True,
                  color_discrete_sequence=NEXUS_PALETTE, title="Detalle diario realmente reportado", labels={"Importe": report["currency"]})
    fig.update_traces(connectgaps=False)
    return chart_layout(fig, 340)


def _finance_period_plot_v7(report: dict):
    periods = report.get("period_sales", pd.DataFrame())
    if periods.empty:
        return None
    plot = periods.copy()
    plot["Reporte"] = plot["site"] + " · " + plot["start_date"] + " → " + plot["end_date"]
    fig = px.bar(plot, y="Reporte", x="net_sales", color="site", orientation="h", text_auto=",.2f",
                 color_discrete_sequence=NEXUS_PALETTE, labels={"net_sales": report["currency"], "site": "Sede"},
                 title="Ventas netas por período reportado", hover_data={"start_date": True, "end_date": True, "site": False})
    fig.update_layout(showlegend=False)
    return chart_layout(fig, max(300, min(600, len(plot) * 55 + 130)))


def _finance_excel_v7(report: dict, frames: dict) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    wb.remove(wb.active)
    cut = report.get("analysis_date", report["end"])
    for name, frame in frames.items():
        ws = wb.create_sheet(name[:31])
        size = max(len(frame.columns), 4)
        for row_number, title in [(1, f"NEXUS · {name}"), (2, f"Ventas: {report['start']} a {report['end']} · Corte de análisis: {cut} · {report['currency']}"),
                                  (3, "N/D = dato no informado. Los períodos conservan sus fechas e importes originales.")]:
            ws.merge_cells(start_row=row_number, start_column=1, end_row=row_number, end_column=size)
            cell = ws.cell(row_number, 1, title)
            cell.font = Font(name="Calibri", size=18 if row_number == 1 else 10, bold=row_number == 1, color="FFFFFF" if row_number == 1 else "617584")
            cell.alignment = Alignment(wrap_text=True, vertical="center")
            if row_number == 1:
                cell.fill = PatternFill("solid", fgColor="163743")
            ws.row_dimensions[row_number].height = 34 if row_number == 1 else 29
        safe = safe_export_frame(frame)
        for index, label in enumerate(safe.columns, 1):
            cell = ws.cell(5, index, str(label))
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="177D79")
            cell.alignment = Alignment(wrap_text=True)
            ws.column_dimensions[get_column_letter(index)].width = min(48, max(17, len(str(label)) + 3))
        ws.row_dimensions[5].height = 31
        for row_number, row in enumerate(safe.itertuples(index=False, name=None), 6):
            for index, value in enumerate(row, 1):
                cell = ws.cell(row_number, index, "N/D" if pd.isna(value) else value)
                cell.font = Font(name="Calibri", size=10, color="1B3543")
                cell.alignment = Alignment(vertical="top", wrap_text=isinstance(value, str) and len(value) > 45)
                if isinstance(value, (float, np.floating)):
                    cell.number_format = '#,##0.00;[Red](#,##0.00);0.00'
                if row_number % 2 == 0:
                    cell.fill = PatternFill("solid", fgColor="F0F6F5")
        ws.freeze_panes = "C6"
        if len(frame.columns):
            ws.auto_filter.ref = f"A5:{get_column_letter(len(frame.columns))}{max(5, len(frame) + 5)}"
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A3
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.print_title_rows = "1:5"
    method = wb.create_sheet("Cómo se calcula")
    method.column_dimensions["A"].width = 120
    for index, line in enumerate(FIN_METHODOLOGY_V7.splitlines(), 1):
        method.cell(index, 1, line).alignment = Alignment(wrap_text=True, vertical="top")
        method.row_dimensions[index].height = max(20, 15 * ((len(line) // 115) + 1))
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


@st.cache_data(show_spinner=False, max_entries=6)
def finance_export_v7(report: dict) -> bytes:
    """ZIP compatible: informe visual autónomo, Excel por sección y CSV reales."""
    frames = _finance_frames_v7(report)
    cut = report.get("analysis_date", report["end"])
    title = f"NEXUS · Desempeño · ventas {report['start']} a {report['end']} · análisis {cut} · {report['currency']}"
    total = report["total"]
    cards = []
    for label, field in [("Inventario · último corte", "Inventario"), ("Ventas registradas", "Ventas netas"), ("Cartera · último corte", "CxC"), ("Margen bruto", "Margen bruto")]:
        cards.append(f'<article><small>{escape(label)}</small><strong>{escape(fin_value(total[field], report["currency"]))}</strong></article>')
    charts = []
    for fig in [_finance_period_plot_v7(report), _finance_daily_plot_v7(report)]:
        if fig is not None:
            charts.append(pio.to_html(fig, full_html=False, include_plotlyjs=True if not charts else False, config={"displaylogo": False, "responsive": True}))
    sections = []
    for name, frame in frames.items():
        if frame.empty and name not in {"Indicadores", "Fuentes y cortes"}:
            body = '<p class="muted">Sin datos reportados para esta sección.</p>'
        else:
            body = f'<div class="scroll">{frame.to_html(index=False, escape=True, na_rep="N/D", float_format=lambda v: f"{v:,.2f}")}</div>'
        sections.append(f'<section><h2>{escape(name)}</h2>{body}</section>')
    methodology = "".join(f"<p>{escape(line)}</p>" for line in FIN_METHODOLOGY_V7.splitlines() if line)
    html = f'''<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{escape(title)}</title>
    <style>body{{margin:0;background:#f2f6f7;color:#163743;font:14px Arial,sans-serif}}main{{max-width:1500px;margin:auto;padding:32px}}header{{background:#163743;color:white;padding:28px;border-radius:16px}}h1{{font-size:27px;margin:8px 0}}h2{{font-size:19px}}header p,.muted{{color:#7b9099}}header p{{color:#d7e7e8}}.cards{{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin:20px 0}}article,section,.chart{{background:white;border:1px solid #dde7e8;border-radius:14px;padding:20px;margin-bottom:18px}}article strong{{display:block;font-size:24px;margin-top:9px}}small{{color:#617584}}table{{border-collapse:collapse;white-space:nowrap;font-size:12px;min-width:100%}}td,th{{padding:10px 12px;text-align:left;border-bottom:1px solid #e3ebed}}th{{background:#e7f3f1;color:#176b67}}tr:nth-child(even){{background:#f6f9f9}}.scroll{{overflow:auto}}section p{{line-height:1.6}}@media(max-width:760px){{main{{padding:12px}}.cards{{grid-template-columns:repeat(2,1fr)}}article strong{{font-size:18px}}}}@media print{{body{{background:white}}main{{padding:0}}.scroll{{overflow:visible}}.cards{{grid-template-columns:repeat(4,1fr)}}section{{break-inside:avoid}}}}</style></head><body><main>
    <header><small style="color:#8dddd6">MAKROPETROL · NEXUS</small><h1>Ventas, inventario y cuentas por cobrar</h1><p>Ventas: {escape(report['start'])} a {escape(report['end'])} · {report['days']} días · {escape(report['currency'])}<br>Corte de análisis: {escape(cut)}</p></header>
    <div class="cards">{''.join(cards)}</div><p class="muted">{'Cobertura completa del período de ventas.' if total['Ventas completas'] else 'Ventas registradas con cobertura parcial: consulta Fuentes y cortes.'} Cada saldo conserva la fecha que aparece por sede.</p>
    {''.join('<div class="chart">' + chart + '</div>' for chart in charts)}{''.join(sections)}<section><h2>Cómo se calcula y qué falta cargar</h2>{methodology}</section></main></body></html>'''
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as package:
        package.writestr("Desempeno.html", html.encode("utf-8"))
        package.writestr("Desempeno.xlsx", _finance_excel_v7(report, frames))
        for name, frame in frames.items():
            slug = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode().replace(" ", "_")
            package.writestr(f"{slug}.csv", safe_export_frame(frame).to_csv(index=False, na_rep="N/D").encode("utf-8-sig"))
        package.writestr("Metodologia.txt", FIN_METHODOLOGY_V7.encode("utf-8"))
    return out.getvalue()


def financial_report_exports_v7(report: dict) -> bytes:
    """Nombre de integración alternativo; v6 usaba finance_export(report)."""
    return finance_export_v7(report)


def _suggest_finance_action_v7(row, max_dio, max_dso, max_late):
    # Las alertas disponibles siguen siendo útiles aunque falte otro dato opcional.
    if row["Cartera vencida %"] > max_late or row["Días CxC"] > max_dso:
        return "Priorizar cobranza", "Revisar saldos vencidos y acuerdos de pago según el corte de cartera."
    if row["Días inventario"] > max_dio:
        return "Revisar existencias", "Revisar productos lentos, redistribución y compras pendientes."
    if pd.notna(row["Margen bruto"]) and row["Margen bruto"] < 0:
        return "Revisar margen", "Contrastar costo real, devoluciones y descuentos."
    if row["Calidad de datos"] != "Completa":
        return "Completar fuentes", row["Calidad de datos"]
    if row.get("Notas de cálculo", ""):
        return "Revisar cortes", row["Notas de cálculo"]
    if pd.isna(row["Días inventario"]) or pd.isna(row["Días CxC"]):
        return "Seguimiento parcial", row.get("Datos pendientes", "") or "Completa las bases necesarias para calcular los indicadores pendientes."
    return "Dentro de metas", "Mantener seguimiento de los indicadores disponibles."


def render_performance_v7(receivables_only: bool = False) -> None:
    st.subheader("Cuentas por cobrar" if receivables_only else "Desempeño integral")
    st.caption("Ventas de su período real, inventario y cartera a la fecha que necesitas analizar.")
    params = financial_filters_v7("performance")
    if not params:
        return
    start, end, currency, sites, analysis_date = params
    demo = st.toggle("Ver una demostración con datos ficticios", key="finance_demo_v7")
    if demo:
        st.warning("DEMOSTRACIÓN · Datos ficticios. No se guardan en tu base de datos. Los saldos del ejemplo corresponden al fin del período de ventas.")
        report = demo_financial_report(start, end, currency, sites)
        report["analysis_date"] = end
        report["period_sales"] = pd.DataFrame()
        analysis_date = end
    else:
        try:
            report = build_financial_report_v7(finance_version(), db_version(), tuple(sorted(finance_settings().items())), start, end, currency, sites, analysis_date)
        except ValueError as exc:
            st.error(str(exc))
            return
    summary, total = report["summary"].copy(), report["total"]
    for field in ["Datos pendientes", "Notas de cálculo"]:
        if field not in summary:
            summary[field] = ""
    partial = not total["Ventas completas"]
    st.caption(f"Ventas del {start} al {end} · {report['days']} días. Inventario y CxC: análisis al {analysis_date}.")
    if analysis_date != end:
        st.info("Los indicadores que relacionan saldos con ventas usan el saldo actual como aproximación sobre ese período histórico. Revisa los métodos por sede.")
    has_notes = (summary["Calidad de datos"] != "Completa").any() or summary[["Datos pendientes", "Notas de cálculo"]].ne("").any().any()
    with st.expander("Fuentes, fechas y datos que faltan", expanded=bool(has_notes)):
        sources = _finance_frames_v7({**report, "summary": summary})["Fuentes y cortes"]
        st.dataframe(sources, width="stretch", hide_index=True)
        st.caption("Días cubiertos incluye períodos completos y días reales sin duplicarlos. N/D indica información pendiente; los demás resultados siguen disponibles.")
        if st.button("Ir a cargar o corregir reportes", key="go_fin_upload_v7"):
            st.session_state["nexus_view"] = "Carga de datos"
            st.rerun()
    cols = st.columns(4)
    for col, label, field, note, tone in [
        (cols[0], "Inventario", "Inventario", "Valor del último corte · ver fecha por sede", "blue"),
        (cols[1], "Ventas registradas", "Ventas netas", "Cobertura parcial del período" if partial else f"{report['days']} días cubiertos · sin IVA", "green"),
        (cols[2], "Cuentas por cobrar", "CxC", "Saldo del último corte de cartera", "cyan"),
        (cols[3], "Cartera vencida", "CxC vencida", fin_value(total["Cartera vencida %"], suffix="%") + " del saldo", "red")]:
        with col:
            kpi(label, fin_value(total[field], currency), note, tone)
    if receivables_only:
        st.caption("Para activar días de CxC hacen falta ventas a crédito reales. El reporte de ventas A2 aportado no identifica crédito ni cobros.")
        render_receivables({**report, "end": analysis_date})
        return
    cols = st.columns(4)
    for col, label, field, suffix, tone in [
        (cols[0], "Días de inventario", "Días inventario", " días", "blue"),
        (cols[1], "Días de CxC", "Días CxC", " días", "orange"),
        (cols[2], "Ciclo operativo", "Ciclo operativo", " días", "cyan"),
        (cols[3], "Margen bruto", "Margen bruto %", "%", "green")]:
        with col:
            note = "Requiere ventas a crédito reportadas" if field == "Días CxC" and pd.isna(total[field]) else "Base y método disponibles por sede"
            kpi(label, fin_value(total[field], suffix=suffix), note, tone)
    st.caption("El ciclo operativo suma días de inventario y CxC. Para obtener el ciclo de caja faltan los plazos de pago a proveedores.")
    overview, seats_tab, collections_tab, method_tab = st.tabs(["Visión del negocio", "Comparar sedes", "Cartera y cobros", "Cómo se calcula"])
    with overview:
        a, b = st.columns(2)
        with a:
            plot = summary.melt(id_vars="Sede", value_vars=["Inventario", "CxC"], var_name="Concepto", value_name="Importe").dropna()
            if not plot.empty:
                fig = px.bar(plot, x="Sede", y="Importe", color="Concepto", barmode="group", color_discrete_map={"Inventario": NEXUS_BLUE, "CxC": NEXUS_CYAN}, title=f"Saldos por sede · {currency}")
                st.plotly_chart(chart_layout(fig, 360), width="stretch")
                st.caption("Último corte disponible de cada concepto; consulta sus fechas arriba.")
            else:
                empty_state("Saldos pendientes", "Carga el inventario y la cartera, y confirma su moneda en Carga de datos.")
        with b:
            periods_plot = _finance_period_plot_v7(report)
            daily_plot = _finance_daily_plot_v7(report)
            if periods_plot is not None:
                st.plotly_chart(periods_plot, width="stretch")
                st.caption("Cada barra conserva el rango completo de su archivo. Rangos de distinta duración no son directamente comparables; no se inventa una evolución mensual.")
            if daily_plot is not None:
                st.plotly_chart(daily_plot, width="stretch")
                st.caption("Se muestran únicamente días reales. Los espacios sin datos quedan abiertos. Cobros puede incluir facturas de otros períodos.")
            if periods_plot is None and daily_plot is None:
                empty_state("Ventas pendientes", "Carga el reporte A2 o selecciona un período que contenga completo su rango.")
        if not report.get("period_sales", pd.DataFrame()).empty:
            with st.expander("Ver los importes y fechas de los reportes A2"):
                st.dataframe(_finance_frames_v7(report)["Ventas por período"], width="stretch", hide_index=True)
        c1, c2, c3 = st.columns(3)
        with c1:
            kpi("Margen bruto", fin_value(total["Margen bruto"], currency), "Antes de gastos operativos e impuestos", "green")
        with c2:
            kpi("Rotación del período", fin_value(total["Rotación período"], suffix=" veces"), "Costo real de ventas / inventario base", "blue")
        with c3:
            kpi("Inventario + CxC al corte", fin_value(total["Capital en inventario y CxC"], currency), f"Requiere ambos saldos al {analysis_date}", "cyan")
    with seats_tab:
        st.markdown("**Metas internas para el seguimiento**")
        caps = st.columns(3)
        max_dio = caps[0].number_input("Máximo de días de inventario", 1, 730, 60, key="target_dio")
        max_dso = caps[1].number_input("Máximo de días de CxC", 1, 730, 45, key="target_dso")
        max_late = caps[2].number_input("Máximo de cartera vencida (%)", 0, 100, 10, key="target_overdue")
        decisions = summary.apply(lambda row: _suggest_finance_action_v7(row, max_dio, max_dso, max_late), axis=1)
        summary["Estado"] = [item[0] for item in decisions]
        summary["Acción sugerida"] = [item[1] for item in decisions]
        scatter = summary.dropna(subset=["Días inventario", "Días CxC", "Inventario al cierre"]).copy()
        if not scatter.empty:
            scatter["Tamaño"] = scatter["Inventario al cierre"].clip(lower=1)
            fig = px.scatter(scatter, x="Días inventario", y="Días CxC", size="Tamaño", color="Estado", hover_name="Sede", size_max=52, color_discrete_sequence=NEXUS_PALETTE, title="Inventario y cobranza por sede")
            fig.add_vline(x=max_dio, line_dash="dot", line_color=NEXUS_GRAY)
            fig.add_hline(y=max_dso, line_dash="dot", line_color=NEXUS_GRAY)
            st.plotly_chart(chart_layout(fig, 420), width="stretch")
        columns = ["Sede", "Estado", "Ventas netas", "Inventario", "CxC", "Cartera vencida %", "Días inventario", "Días CxC", "Margen bruto %", "Acción sugerida"]
        st.dataframe(summary[columns], hide_index=True, width="stretch")
        if st.button("Crear tareas de seguimiento para estas sedes", disabled=demo, key="finance_tasks_v7"):
            created = 0
            for _, row in summary.iterrows():
                if row["Estado"] == "Dentro de metas":
                    continue
                created += save_action_task(f"{row['Estado']} · {row['Sede']}", row["Sede"], "Alta" if row["Estado"] in {"Priorizar cobranza", "Revisar margen"} else "Media",
                                            f"fin:{row['Sede']}:{currency}:{start}:{end}:{analysis_date}:{row['Estado']}", notes=row["Acción sugerida"])
            st.success(f"{created} tareas nuevas; se evitan duplicados del mismo período y corte.")
    with collections_tab:
        kpi("Cobros registrados", fin_value(total["Cobros registrados"], currency), "Cobros reales informados para el período seleccionado", "green")
        if pd.isna(total["Cobros registrados"]):
            st.info("Este reporte no informa todos los cobros. Carga el resumen financiero que incluya los importes efectivamente cobrados para completar el indicador.")
        render_receivables({**report, "end": analysis_date})
    with method_tab:
        st.markdown(FIN_METHODOLOGY_V7)
    if not demo:
        report["summary"] = summary
        token = (finance_version(), db_version(), start, end, currency, sites, analysis_date,
                 tuple(sorted(finance_settings().items())), max_dio, max_dso, max_late, "finance_v7")
        expected = hashlib.sha256(repr(token).encode()).hexdigest()
        if st.button("Preparar informe de desempeño", key="prepare_fin_report_v7"):
            with st.spinner("Preparando indicadores, fuentes y reportes del período…"):
                st.session_state["finance_export_payload_v7"] = (expected, finance_export_v7(report))
        payload = st.session_state.get("finance_export_payload_v7")
        if payload and payload[0] == expected:
            st.download_button("Descargar informe · Excel + HTML + CSV", payload[1],
                               f"NEXUS_Desempeno_{start}_{end}_corte_{analysis_date}_{currency}.zip", "application/zip", key="download_fin_report_v7")

def render_operational_documents_v7():
    st.markdown('<div class="section-title">Centro de documentos y exportación</div>', unsafe_allow_html=True)
    st.caption("Genera documentos listos para operación y administración. Elige el documento y la sede para preparar una orden o un movimiento.")

    if all_df.empty:
        empty_state("Nada que documentar", "Carga y procesa al menos un corte diario.", "📭")
    else:
        f1, f2, f3 = st.columns(3)
        with f1:
            doc_type_label = st.selectbox(
                "Tipo de documento",
                [
                    "🛒 Pedido de compra",
                    "📤 Retiro de almacén",
                    "🚨 Alertas por sede",
                    "♻️ Redistribución multisede",
                ],
                key="doc_type_selector",
            )
        with f2:
            if "Redistribución" in doc_type_label:
                selected_doc_site = "GLOBAL / TODAS LAS SEDES"
            else:
                doc_sites = sorted(all_df["Sede"].unique())
                selected_doc_site = st.selectbox(
                    "Sede",
                    doc_sites,
                    key="doc_site_selector",
                )
        with f3:
            output_format = st.selectbox(
                "Formato",
                ["PDF", "HTML imprimible", "Excel", "CSV"],
                key="doc_format_selector",
            )

        if "Pedido de compra" in doc_type_label:
            doc_kind = "purchase"
            doc_title = "ORDEN DE COMPRA"
            doc_df = all_df if selected_doc_site == "GLOBAL / TODAS LAS SEDES" else all_df[all_df["Sede"] == selected_doc_site]
            prepared = _doc_prepare_frame(doc_df, "purchase", selected_doc_site)
        elif "Retiro" in doc_type_label:
            doc_kind = "withdrawal"
            doc_title = "RETIRO DE ALMACÉN"
            doc_df = all_df if selected_doc_site == "GLOBAL / TODAS LAS SEDES" else all_df[all_df["Sede"] == selected_doc_site]
            prepared = _doc_prepare_frame(doc_df, "withdrawal", selected_doc_site)
        elif "Alertas" in doc_type_label:
            doc_kind = "alerts"
            doc_title = "REPORTE DE ALERTAS OPERATIVAS"
            doc_df = alerts if selected_doc_site == "GLOBAL / TODAS LAS SEDES" else alerts[alerts["Sede"] == selected_doc_site]
            prepared = _doc_prepare_frame(doc_df, "alerts", selected_doc_site)
        else:
            doc_kind = "redistribution"
            doc_title = "PLAN DE REDISTRIBUCIÓN"
            doc_df = transfers
            prepared = _doc_prepare_frame(doc_df, "redistribution")

        if doc_kind in {"purchase", "withdrawal"}:
            units = float(pd.to_numeric(prepared.get("Cantidad", 0), errors="coerce").fillna(0).sum()) if not prepared.empty else 0
            value = float(pd.to_numeric(prepared.get("Total", 0), errors="coerce").sum(min_count=len(prepared))) if not prepared.empty else 0.0
            k1, k2, k3, k4 = st.columns(4)
            with k1: kpi("Items", integer(len(prepared)), "Documento", "blue")
            with k2: kpi("Unidades", _quantity_text_v73(units), "Cantidad física", "cyan")
            with k3: kpi("Valor", money(value), "Estimado", "orange")
            with k4: kpi("Ámbito", "Por sede", selected_doc_site, "green")
        elif doc_kind == "alerts":
            crit = int((prepared["Severidad"].astype(str) == "CRÍTICA").sum()) if not prepared.empty else 0
            high = int((prepared["Severidad"].astype(str) == "ALTA").sum()) if not prepared.empty else 0
            k1, k2, k3, k4 = st.columns(4)
            with k1: kpi("Alertas", integer(len(prepared)), "Registros", "red")
            with k2: kpi("Críticas", integer(crit), "Atención inmediata", "red")
            with k3: kpi("Altas", integer(high), "Prioridad", "orange")
            with k4: kpi("Ámbito", "Por sede", selected_doc_site, "blue")
        else:
            units = int(pd.to_numeric(prepared.get("Unidades Sugeridas", 0), errors="coerce").fillna(0).sum()) if not prepared.empty else 0
            avoided = float(pd.to_numeric(prepared.get("Compra Evitada Estimada ($)", 0), errors="coerce").sum(min_count=len(prepared))) if not prepared.empty else 0.0
            routes = int(len(prepared)) if not prepared.empty else 0
            k1, k2, k3, k4 = st.columns(4)
            with k1: kpi("Movimientos", integer(routes), "Cruces origen → destino", "green")
            with k2: kpi("Unidades", integer(units), "Redistribución", "blue")
            with k3: kpi("Compra evitada", money(avoided), "Estimación", "orange")
            with k4: kpi("Ámbito", "Multisede", "Cruce global", "cyan")

        st.markdown('<div class="section-title">Vista previa del documento</div>', unsafe_allow_html=True)
        preview_cols = prepared.head(60)
        if preview_cols.empty:
            empty_state("Sin registros para este documento", "No hay productos que cumplan las reglas actuales.", "✓")
        else:
            st.dataframe(
                preview_cols,
                width="stretch",
                hide_index=True,
                height=430,
            )
            if len(prepared) > 60:
                st.caption(f"Vista previa limitada a 60 filas. El archivo conserva las {len(prepared):,} filas completas.")

        st.markdown('<div class="section-title">Descargar documento</div>', unsafe_allow_html=True)
        site_slug = safe_slug(selected_doc_site)
        doc_slug = safe_slug(doc_title)

        if output_format == "PDF":
            if doc_kind == "purchase":
                file_bytes = build_purchase_pdf(doc_df, selected_doc_site, latest_display, sequence=1)
            elif doc_kind == "withdrawal":
                file_bytes = build_withdrawal_pdf(doc_df, selected_doc_site, latest_display)
            elif doc_kind == "alerts":
                file_bytes = build_alerts_pdf(alerts, selected_doc_site, latest_display)
            else:
                file_bytes = build_redistribution_pdf(transfers, latest_display)
            if file_bytes:
                st.download_button(
                    "📄 Descargar PDF", file_bytes,
                    f"NEXUS_{doc_slug}_{site_slug}_{latest_display}.pdf",
                    DOC_MIME_PDF, width="stretch",
                )
            else:
                st.info("Instala fpdf2 para habilitar PDF.")

        elif output_format == "HTML imprimible":
            html_site = None if doc_kind == "redistribution" else selected_doc_site
            file_bytes = build_operational_html(doc_kind, doc_df, html_site, latest_display, f"{doc_title} · {selected_doc_site}")
            st.download_button(
                "🌐 Descargar HTML imprimible", file_bytes,
                f"NEXUS_{doc_slug}_{site_slug}_{latest_display}.html",
                DOC_MIME_HTML, width="stretch",
            )

        elif output_format == "Excel":
            if doc_kind == "purchase":
                file_bytes = build_purchase_excel(doc_df, selected_doc_site, latest_display)
            else:
                file_bytes = export_xlsx_bytes(prepared, doc_title)
            st.download_button(
                "📊 Descargar Excel", file_bytes,
                f"NEXUS_{doc_slug}_{site_slug}_{latest_display}.xlsx",
                DOC_MIME_XLSX, width="stretch",
            )

        else:
            file_bytes = export_csv_bytes(prepared)
            st.download_button(
                "🧾 Descargar CSV", file_bytes,
                f"NEXUS_{doc_slug}_{site_slug}_{latest_display}.csv",
                "text/csv", width="stretch",
            )



render_performance = render_performance_v7
finance_export = finance_export_v7
FIN_METHODOLOGY = FIN_METHODOLOGY_V7


# """Lectura de reportes impresos de A2, sin modificar los archivos originales.

# Contrato público: parse_a2_report(raw, filename, kind='auto') -> dict con
# kind, data (pandas.DataFrame), metadata, warnings y catalog. No escribe en DB.
# Los identificadores ausentes se conservan vacíos y bloquean ready_to_save.
# Los importes del período nunca se distribuyen artificialmente entre fechas.
# """

import csv
import io
import zipfile
import math
import re
import struct
import unicodedata
from datetime import datetime
from decimal import Decimal, InvalidOperation

import pandas as pd


def _a2_text(value):
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _a2_norm(value):
    return re.sub(r"[^a-z0-9]+", " ", unicodedata.normalize("NFKD", _a2_text(value)).encode("ascii", "ignore").decode().lower()).strip()


def _a2_number(value):
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value) if math.isfinite(float(value)) else None
    text = _a2_text(value).replace("\u00a0", "").replace(" ", "")
    if not text:
        return None
    negative = text.startswith("(") and text.endswith(")")
    text = text.strip("()")
    text = re.sub(r"^(USD|US\$|Bs\.?|VES|\$)", "", text, flags=re.I)
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".") if text.rfind(",") > text.rfind(".") else text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")
    try:
        number = float(Decimal(text))
        return (-number if negative else number) if math.isfinite(number) else None
    except (InvalidOperation, ValueError):
        return None


def _a2_sum(values):
    return float(sum((Decimal(str(v)) for v in values if v is not None), Decimal("0")))


def _a2_legacy_grid(raw):
    """Fallback limitado al flujo BIFF antiguo que exporta A2 (NUMBER/LABEL).

    No se usa para otros libros BIFF: un registro desconocido con datos o un
    archivo truncado genera error en vez de omitir contenido silenciosamente.
    """
    cells = {}
    offset = 0
    eof = False
    while offset + 4 <= len(raw):
        record, size = struct.unpack_from("<HH", raw, offset)
        payload = raw[offset + 4:offset + 4 + size]
        if len(payload) != size:
            raise ValueError("El archivo A2 está incompleto: vuelve a exportarlo.")
        if record in (3, 4):
            if len(payload) < (15 if record == 3 else 8):
                raise ValueError("Una celda A2 está incompleta.")
            row, col = struct.unpack_from("<HH", payload)
            if row > 65535 or col > 255:
                raise ValueError("Dimensiones A2 no admitidas.")
            if record == 3:
                value = struct.unpack_from("<d", payload, 7)[0]
            else:
                length = payload[7]
                if 8 + length > len(payload):
                    raise ValueError("Una etiqueta A2 está incompleta.")
                value = payload[8:8 + length].decode("cp1252", errors="replace")
            if (row, col) in cells:
                raise ValueError("El archivo contiene celdas superpuestas: revisa la exportación.")
            cells[row, col] = value
        elif record == 10:
            eof = True
            break
        elif record not in (0x0809, 0x0200):
            raise ValueError("Este formato XLS necesita el lector xlrd; no corresponde al exportador A2 validado.")
        offset += 4 + size
    if not cells or not eof:
        raise ValueError("No se pudo leer un reporte A2 completo.")
    height = max(r for r, c in cells) + 1
    width = max(c for r, c in cells) + 1
    _a2_check_grid_bounds(height, width)
    grid = [[""] * width for _ in range(height)]
    for (row, col), value in cells.items():
        grid[row][col] = value
    return [("Reporte A2", grid)], "A2 BIFF antiguo"


_A2_MAX_BYTES = 32 * 1024 * 1024
_A2_MAX_ROWS = 100000
_A2_MAX_COLUMNS = 256
_A2_MAX_CELLS = 2000000


def _a2_check_grid_bounds(height, width):
    if height > _A2_MAX_ROWS or width > _A2_MAX_COLUMNS or height * width > _A2_MAX_CELLS:
        raise ValueError("El reporte supera 100.000 filas, 256 columnas o 2 millones de celdas. Exporta por sede o período.")


def _a2_read_grids(raw, filename):
    """Lectura acotada y sin ejecución de macros; conserva ceros iniciales de texto."""
    if not raw:
        raise ValueError("El archivo está vacío.")
    if len(raw) > _A2_MAX_BYTES:
        raise ValueError("Máximo 32 MB por reporte. Exporta una sede o período por archivo.")
    if raw[:4] == b"\x09\x08\x06\x00":
        return _a2_legacy_grid(raw)
    suffix = filename.lower().rsplit(".", 1)[-1]
    if suffix == "csv":
        text = None
        for encoding in ("utf-8-sig", "cp1252"):
            try:
                text = raw.decode(encoding)
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            raise ValueError("El CSV tiene una codificación no admitida; guárdalo como UTF-8.")
        try:
            dialect = csv.Sniffer().sniff(text[:65536], delimiters=";,\t|")
            delimiter = dialect.delimiter
        except csv.Error:
            # Los títulos de A2 pueden tener menos columnas que la tabla.
            counts = {sep: sum(line.count(sep) for line in text.splitlines()[:100]) for sep in (";", "\t", ",", "|")}
            delimiter = max(counts, key=counts.get)
        grid, width = [], 0
        for row in csv.reader(io.StringIO(text), delimiter=delimiter):
            width = max(width, len(row))
            _a2_check_grid_bounds(len(grid) + 1, width)
            grid.append(row)
        return [("CSV", [row + [""] * (width - len(row)) for row in grid])], "CSV"
    try:
        sheets, cells = [], 0
        if raw[:2] == b"PK":
            from openpyxl import load_workbook
            with zipfile.ZipFile(io.BytesIO(raw)) as archive:
                if sum(item.file_size for item in archive.infolist()) > 128 * 1024 * 1024:
                    raise ValueError("El Excel expandido supera 128 MB; divide el reporte.")
            workbook = load_workbook(io.BytesIO(raw), read_only=True, data_only=True, keep_links=False)
            try:
                if len(workbook.sheetnames) > 32:
                    raise ValueError("Máximo 32 hojas por libro; exporta únicamente los reportes necesarios.")
                for sheet in workbook.worksheets:
                    _a2_check_grid_bounds(sheet.max_row or 0, sheet.max_column or 0)
                    grid, width = [], 0
                    for row in sheet.iter_rows(values_only=True):
                        values = ["" if value is None else value for value in row]
                        width = max(width, len(values))
                        _a2_check_grid_bounds(len(grid) + 1, width)
                        grid.append(values)
                    cells += len(grid) * width
                    if cells > _A2_MAX_CELLS:
                        raise ValueError("El libro supera 2 millones de celdas entre todas sus hojas.")
                    sheets.append((sheet.title, grid))
            finally:
                workbook.close()
            return sheets, "openpyxl"
        import xlrd
        workbook = xlrd.open_workbook(file_contents=raw, on_demand=True)
        try:
            if workbook.nsheets > 32:
                raise ValueError("Máximo 32 hojas por libro.")
            for name in workbook.sheet_names():
                sheet = workbook.sheet_by_name(name)
                _a2_check_grid_bounds(sheet.nrows, sheet.ncols)
                cells += sheet.nrows * sheet.ncols
                if cells > _A2_MAX_CELLS:
                    raise ValueError("El libro supera 2 millones de celdas entre todas sus hojas.")
                sheets.append((name, [sheet.row_values(i) for i in range(sheet.nrows)]))
                workbook.unload_sheet(name)
        finally:
            workbook.release_resources()
        return sheets, "xlrd"
    except (ValueError, ImportError):
        raise
    except Exception as exc:
        raise ValueError("No se pudo leer el libro. Admite XLS de A2, XLSX y CSV; para XLS estándar instala xlrd.") from exc


def read_a2_grid(raw: bytes, filename: str = "reporte.xls") -> pd.DataFrame:
    sheets, _ = _a2_read_grids(raw, filename)
    if len(sheets) != 1:
        raise ValueError("El archivo tiene varias hojas; selecciona o exporta una sola para esta carga.")
    return pd.DataFrame(sheets[0][1])


_A2_ALIASES = {
    "code": {"codigo", "cod", "codigo producto", "codigo de producto", "codigo del producto", "codigo articulo", "codigo de articulo", "codigo del articulo", "sku", "referencia", "ref", "codigo interno"},
    "description": {"descripcion", "producto", "articulo", "descripcion producto", "descripcion del producto", "descripcion de producto", "descripcion articulo", "nombre producto", "nombre del producto", "nombre del articulo"},
    "existence": {"existencia", "exist", "stock", "existencias", "saldo unidades", "stock actual", "existencia actual", "cantidad disponible", "disponible", "cantidad en existencia"},
    "cost": {"costo", "costo unitario", "costo promedio", "costo promedio unitario", "costo actual", "coste", "coste unitario", "costo de ventas"},
    "inventory_value": {"valor inventario", "valor de inventario", "valor del inventario", "valor existencia", "valor total inventario", "valorizacion", "valorizacion inventario"},
    "quantity": {"cantidad", "ventas", "unidades", "unidades vendidas", "cantidad vendida", "cant", "cant vendida", "venta unidades", "ventas unidades"},
    "gross_sales": {"monto bruto", "importe bruto", "venta bruta", "ventas brutas", "importe de ventas", "monto ventas", "monto de ventas"},
    "net_sales": {"importe neto", "monto neto", "venta neta", "ventas netas"},
    "discounts": {"descuentos", "descuento"},
    "tax": {"i v a", "iva", "impuesto", "impuestos"},
    "profit": {"utilidad", "beneficio"},
    "margin": {"porcentaje utilidad", "margen", "margen utilidad"},
    "items": {"numero de items", "numero items", "items"},
    "active": {"activo"},
    "classification": {"clasificacion"},
    "address": {"direccion"},
    "supplier": {"proveedor"},
    "brand": {"marca"},
    "category": {"categoria"},
    "department": {"departamento"},
}
_A2_LABELS = {"code": "Código / SKU", "description": "Descripción", "existence": "Existencia", "cost": "Costo", "inventory_value": "Valor inventario", "quantity": "Cantidad vendida", "gross_sales": "Monto bruto", "net_sales": "Importe neto", "discounts": "Descuentos", "tax": "IVA reportado", "profit": "Utilidad", "supplier": "Proveedor", "brand": "Marca", "category": "Categoría", "department": "Departamento"}
_A2_ALIAS_INDEX = {alias: field for field, aliases in _A2_ALIASES.items() for alias in aliases}


def _a2_header_field(value):
    text, normalized = _a2_text(value), _a2_norm(value)
    if text.startswith("%") and normalized == "utilidad":
        return "margin"
    return _A2_ALIAS_INDEX.get(normalized)


def _a2_detect_kind(fields, requested):
    if "description" not in fields:
        return None
    if "existence" in fields:
        detected = "inventario"
    elif "quantity" in fields:
        detected = "ventas"
    elif {"active", "classification"} <= fields.keys():
        detected = "departamentos"
    elif {"active", "address"} <= fields.keys():
        detected = "proveedores"
    else:
        return None
    return detected if requested in ("auto", detected) else None


def _a2_header_at(grid, row_index, requested):
    row = grid[row_index]
    fields, duplicates, bounds = {}, set(), set()
    height = 1
    for col, value in enumerate(row):
        if _a2_text(value):
            bounds.add(col)
        field = _a2_header_field(value)
        consumed = 1
        if field is None and _a2_text(value):
            parts = [_a2_text(value)]
            for offset in (1, 2):
                if row_index + offset >= len(grid):
                    break
                other = grid[row_index + offset]
                item = other[col] if col < len(other) else ""
                if _a2_number(item) is not None:
                    break
                if _a2_text(item):
                    parts.append(_a2_text(item))
                joined = _a2_header_field(" ".join(parts))
                if joined:
                    field, consumed = joined, offset + 1
                    break
        if field:
            if field in fields:
                duplicates.add(field)
            fields[field] = col
            height = max(height, consumed)
    for field in duplicates:
        del fields[field]
    detected = _a2_detect_kind(fields, requested)
    if detected:
        fields["_bounds"] = sorted(bounds)
        fields["_height"] = height
        fields["_ambiguous"] = sorted(duplicates)
        return (sum(not key.startswith("_") for key in fields), row_index, detected, fields)
    return None


def _a2_header(grid, requested):
    candidates = [found for i in range(min(len(grid), 250)) if (found := _a2_header_at(grid, i, requested))]
    # La primera tabla reconocida evita perder productos antes de un encabezado
    # repetido que tenga más campos. Las filas dudosas quedan en revisión.
    return min(candidates, default=None, key=lambda value: value[1])


def _a2_column_letter(index):
    result = ""
    while index >= 0:
        index, remainder = divmod(index, 26)
        result = chr(65 + remainder) + result
        index -= 1
    return result


def _a2_preview_data_mapping(grid, header_index, anchors):
    """Sugiere columnas sólo cuando las primeras filas válidas coinciden."""
    output = {}
    columns = [value for key, value in anchors.items() if not key.startswith("_")]
    boundaries = columns + anchors.get("_bounds", [])
    numeric = {"existence", "cost", "inventory_value", "quantity", "gross_sales", "net_sales", "discounts", "tax", "profit", "items"}
    required = "existence" if "existence" in anchors else "quantity" if "quantity" in anchors else None
    for field, start in anchors.items():
        if field.startswith("_"):
            continue
        end = min((col for col in boundaries if col > start), default=max(map(len, grid), default=0))
        positions = []
        for row in grid[header_index + anchors.get("_height", 1):header_index + anchors.get("_height", 1) + 30]:
            if required and _a2_range_value(row, required, anchors, numeric=True) is None:
                continue
            matches = [col for col in range(start, min(len(row), end)) if _a2_text(row[col]) and (field not in numeric or _a2_number(row[col]) is not None)]
            if len(matches) == 1:
                positions.append(matches[0])
        output[field] = positions[0] if positions and len(set(positions)) == 1 else None
    return output


def inspect_a2_report(raw: bytes, filename: str, kind: str = "auto") -> dict:
    """Vista acotada para elegir hoja, fila (1-based) y columnas (0-based)."""
    requested = {"inventory": "inventario", "sales": "ventas", "departments": "departamentos", "suppliers": "proveedores"}.get(kind, kind)
    sheets, reader = _a2_read_grids(raw, filename)
    output = {"reader": reader, "sheets": [], "fields": dict(_A2_LABELS)}
    for name, grid in sheets:
        candidates = []
        for i in range(min(len(grid), 250)):
            found = _a2_header_at(grid, i, requested)
            if found:
                score, index, detected, anchors = found
                candidates.append({"kind": detected, "header_row": index + 1, "header_height": anchors["_height"], "mapping": {k: v for k, v in anchors.items() if not k.startswith("_")}, "data_mapping": _a2_preview_data_mapping(grid, index, anchors), "score": score})
        preview = [{"Fila": i + 1, **{_a2_column_letter(j): _a2_text(v) for j, v in enumerate(row)}} for i, row in enumerate(grid[:40])]
        output["sheets"].append({"name": name, "rows": len(grid), "columns": max(map(len, grid), default=0), "preview": preview, "candidates": candidates[:20]})
    return output


def _a2_range_value(row, field, anchors, numeric=False):
    start = anchors.get(field)
    if start is None:
        return None if numeric else ""
    other_cols = [col for key, col in anchors.items() if not key.startswith("_") and isinstance(col, int)]
    boundaries = other_cols + anchors.get("_bounds", [])
    end = start + 1 if anchors.get("_exact") else min((col for col in boundaries if col > start), default=len(row))
    choices = [value for value in row[start:end] if _a2_text(value) != ""]
    if numeric:
        numbers = [_a2_number(value) for value in choices]
        numbers = [value for value in numbers if value is not None]
        # Más de un número bajo el mismo encabezado requiere mapeo explícito.
        return numbers[0] if len(numbers) == 1 else None
    return _a2_text(choices[0]) if choices else ""


def _a2_date(text):
    match = re.search(r"\b(\d{2}/\d{2}/\d{4})\b", text)
    if not match:
        return None
    try:
        return datetime.strptime(match.group(1), "%d/%m/%Y").date().isoformat()
    except ValueError:
        return None


def _a2_metadata(grid, header_row, source, sheet, reader):
    metadata = {"filename": source, "sheet": sheet, "reader": reader,
                "header_row": header_row + 1, "currency": None, "site": None,
                "snapshot_date": None, "report_date": None,
                "period_start": None, "period_end": None,
                "unresolved_rows": [], "totals_printed": {},
                "reconciliation": [], "reconciled": None, "ready_to_save": True}
    for row in grid[:header_row]:
        text = " | ".join(_a2_text(v) for v in row if _a2_text(v))
        normal = _a2_norm(text)
        if "fecha" in normal and _a2_date(text):
            metadata["report_date"] = _a2_date(text)
        if normal.startswith("desde"):
            metadata["period_start"] = _a2_date(text)
        if normal.startswith("hasta"):
            metadata["period_end"] = _a2_date(text)
        if "moneda" in normal:
            if re.search(r"\b(dolares|dolar|usd)\b", normal):
                metadata["currency"] = "USD"
            elif re.search(r"\b(bolivares|bolivar|ves)\b", normal):
                metadata["currency"] = "VES"
        deposit = re.search(r"dep[oó]sitos?\s*:?[\s]+(.+)", text, re.I)
        if deposit:
            metadata["site"] = deposit.group(1).split("|")[0].strip().title()
    return metadata


def _a2_compare(metadata, field, observed, printed, warnings):
    if printed is None:
        return
    observed = float(observed)
    difference = round(observed - float(printed), 6)
    # Sólo sumas monetarias: el origen puede redondear cada línea a centavos.
    # Hasta cinco centavos se conserva como diferencia visible, no se corrige.
    tolerance = min(0.05, 0.005 * max(int(metadata.get("row_count", 1)), 1)) if field not in {"row_count", "quantity", "existence"} else 0.000001
    exact = abs(difference) <= 0.000001
    matches = abs(difference) <= tolerance + 0.00000001
    metadata["totals_printed"][field] = float(printed)
    metadata["reconciliation"].append({"field": field, "extracted": observed, "printed": float(printed), "difference": difference, "matches": matches, "exact": exact, "tolerance": tolerance})
    if not matches:
        metadata["ready_to_save"] = False
        metadata.setdefault("blocking_reasons", []).append("total_mismatch")
        warnings.append(f"El total de {field} no coincide: extraído {observed:,.2f}; impreso {printed:,.2f}. Revisa el archivo antes de guardar.")
    elif not exact:
        warnings.append(f"Diferencia de redondeo en {field}: {difference:+.2f}. Dentro de tolerancia monetaria de {tolerance:.3f}; se conserva el valor extraído y el total original por separado.")


def parse_a2_report(raw: bytes, filename: str, kind: str = "auto", sheet_name=None, header_row=None, mapping=None) -> dict:
    """Lee reportes sin inventar códigos, costos, moneda ni fechas.

    header_row usa números de Excel desde 1. mapping={campo: columna_desde_0}
    usa columnas de DATOS exactas; None mantiene detección por anclas impresas.
    El resultado puede contener filas pendientes: ready_to_save las bloquea.
    """
    requested = {"inventory": "inventario", "sales": "ventas", "departments": "departamentos", "suppliers": "proveedores"}.get(kind, kind)
    if requested not in ("auto", "inventario", "ventas", "departamentos", "proveedores"):
        raise ValueError("Tipo A2 no admitido: usa inventario, ventas, departamentos o proveedores.")
    sheets, reader = _a2_read_grids(raw, filename)
    if sheet_name is not None:
        sheets = [(name, grid) for name, grid in sheets if name == sheet_name]
        if not sheets:
            raise ValueError("La hoja seleccionada no existe en este archivo.")
    matches = []
    for name, grid in sheets:
        if header_row is not None:
            if not isinstance(header_row, int) or header_row < 1 or header_row > len(grid):
                raise ValueError("La fila de encabezado debe existir en la hoja (numeración desde 1).")
            index = header_row - 1
            found = _a2_header_at(grid, index, requested)
            if mapping is not None:
                width = max(map(len, grid), default=0)
                selected = {key: value for key, value in mapping.items() if value is not None}
                if any(key not in _A2_ALIASES or not isinstance(value, int) or value < 0 or value >= width for key, value in selected.items()):
                    raise ValueError("El mapeo contiene un campo o una columna inválida.")
                if len(set(selected.values())) != len(selected):
                    raise ValueError("Una columna no puede representar dos campos diferentes.")
                detected = _a2_detect_kind(selected, requested)
                if detected:
                    selected.update({"_exact": True, "_height": 1, "_bounds": [], "_ambiguous": []})
                    found = (len(mapping), index, detected, selected)
            if found:
                matches.append((name, grid, found))
        else:
            if mapping is not None:
                raise ValueError("Selecciona la fila de encabezado cuando asignes columnas manualmente.")
            found = _a2_header(grid, requested)
            if found:
                matches.append((name, grid, found))
    if not matches:
        raise ValueError("No se pudo identificar la tabla. Abre 'Ajustar lectura': elige hoja, fila de encabezado y columnas. Inventario necesita descripción y existencia; ventas, descripción y cantidad. El código real es necesario para guardar por producto.")
    if len(matches) > 1:
        raise ValueError("Hay varias hojas con reportes. Selecciona una hoja o carga cada hoja como una entrada del lote.")
    sheet, grid, (_, header_index, detected, anchors) = matches[0]
    anchors = dict(anchors)
    metadata = _a2_metadata(grid, header_index, filename, sheet, reader)
    metadata.update({"column_anchors": {key: value for key, value in anchors.items() if not key.startswith("_")}, "header_height": anchors.get("_height", 1), "manual_mapping": mapping is not None, "blocking_reasons": [], "rejected_rows": [], "skipped_header_rows": [], "availability": {}})
    metadata["snapshot_date"] = metadata["report_date"] if detected == "inventario" else metadata["period_end"]
    warnings, records, summary_rows, printed_counts = [], [], [], []
    numeric_keys = ("existence", "cost", "inventory_value") if detected == "inventario" else ("quantity", "gross_sales", "net_sales", "discounts", "tax", "cost", "profit")
    first_data_index = header_index + anchors.get("_height", 1)
    skip_until = first_data_index
    labels = {"descripcion", "total", "totales", "total general", "subtotal", "sub total", "total registros", "total items"}
    for index in range(first_data_index, len(grid)):
        if index < skip_until:
            continue
        row, source_index = grid[index], index + 1
        repeated = _a2_header_at(grid, index, detected)
        if repeated:
            if mapping is None:
                anchors = dict(repeated[3])
            skip_until = index + repeated[3].get("_height", 1)
            metadata["skipped_header_rows"].append(source_index)
            continue
        description = _a2_range_value(row, "description", anchors)
        code = _a2_range_value(row, "code", anchors)
        normalized = _a2_norm(description)
        row_labels = {_a2_norm(value) for value in row if _a2_text(value)}
        is_summary_label = normalized in labels or bool(re.match(r"^(?:sub ?total|total general)(?: |$)", normalized))
        is_detail = bool(description) and not is_summary_label
        if detected in ("inventario", "ventas"):
            numbers = {key: _a2_range_value(row, key, anchors, numeric=True) for key in numeric_keys}
            required = "existence" if detected == "inventario" else "quantity"
            has_number = any(value is not None for value in numbers.values())
            is_detail = is_detail and numbers.get(required) is not None
            if is_detail:
                record = {"Código": code, "Descripción": description, "Fila origen": source_index}
                if detected == "inventario":
                    record.update({"Existencia": numbers["existence"], "Costo": numbers["cost"], "Valor inventario A2": numbers["inventory_value"], "Costo disponible": numbers["cost"] is not None, "Proveedor": _a2_range_value(row, "supplier", anchors), "Marca": _a2_range_value(row, "brand", anchors), "Categoría": _a2_range_value(row, "category", anchors), "Departamento": _a2_range_value(row, "department", anchors)})
                else:
                    record.update({"Ventas": numbers["quantity"], "Monto bruto": numbers["gross_sales"], "Descuentos": numbers["discounts"], "IVA reportado": numbers["tax"], "Costo de ventas": numbers["cost"], "Utilidad A2": numbers["profit"]})
                    net = numbers["net_sales"]
                    if net is None and numbers["gross_sales"] is not None and numbers["discounts"] is not None:
                        net = float(Decimal(str(numbers["gross_sales"])) - Decimal(str(numbers["discounts"])))
                    record["Importe neto"] = net
                records.append(record)
            elif (not description or is_summary_label) and has_number:
                # Total Items describe cantidad de filas; nunca se suma como unidades.
                if not row_labels.intersection({"total registros", "total items"}):
                    summary_rows.append({**numbers, "_general": bool(row_labels.intersection({"total general", "totales generales"})), "_subtotal": any(label.startswith(("subtotal", "sub total")) for label in row_labels), "_row": source_index})
            elif description and (has_number or (code and _a2_norm(code) not in labels)):
                # Mantener evidencia de una fila sospechosa, sin perderla silenciosamente.
                metadata["rejected_rows"].append({"Fila origen": source_index, "Código": code, "Descripción": description, "motivo": f"No se pudo leer {required} en una sola celda numérica"})
        elif detected == "departamentos":
            is_detail = is_detail and (_a2_norm(_a2_range_value(row, "active", anchors)) in ("si", "no") or _a2_range_value(row, "items", anchors, numeric=True) is not None)
            if is_detail:
                active = _a2_norm(_a2_range_value(row, "active", anchors))
                records.append({"Código": code, "Departamento": description, "Activo": True if active == "si" else False if active == "no" else None, "Clasificación": _a2_range_value(row, "classification", anchors), "Número de ítems": _a2_range_value(row, "items", anchors, numeric=True), "Fila origen": source_index})
        else:
            is_detail = is_detail and bool(code) and _a2_norm(_a2_range_value(row, "active", anchors)) in ("si", "no")
            if is_detail:
                records.append({"Código": code, "Proveedor": description, "Activo": _a2_norm(_a2_range_value(row, "active", anchors)) == "si", "Fila origen": source_index})
        if not is_detail:
            for column, value in enumerate(row):
                label = _a2_norm(value)
                if label in {"total registros", "total items"}:
                    right_numbers = [_a2_number(value) for value in row[column + 1:]]
                    right_numbers = [value for value in right_numbers if value is not None]
                    if right_numbers:
                        if label == "total registros" or detected == "ventas":
                            printed_counts.append(right_numbers[-1])
                        elif detected == "departamentos":
                            metadata["totals_printed"]["catalog_items"] = right_numbers[-1]
    if not records:
        raise ValueError("El encabezado se reconoce, pero no hay filas con descripción y cantidad válida. Revisa las columnas en Ajustar lectura.")
    frame = pd.DataFrame(records)
    metadata["row_count"] = len(frame)
    metadata["unresolved_rows"] = frame.loc[frame["Código"].eq("")].to_dict("records")
    if metadata["unresolved_rows"]:
        metadata["ready_to_save"] = False
        metadata["blocking_reasons"].append("missing_codes")
        warnings.append(f"{len(metadata['unresolved_rows'])} fila(s) no tienen código en el archivo original. Se conservan pendientes; asigna el código real antes de guardar por producto.")
    duplicates = frame.loc[frame["Código"].ne("") & frame["Código"].duplicated(keep=False), "Código"].unique().tolist()
    metadata["duplicate_codes"] = duplicates
    if duplicates:
        metadata["ready_to_save"] = False
        metadata["blocking_reasons"].append("duplicate_codes")
        warnings.append(f"Hay {len(duplicates)} códigos repetidos; revisa si corresponden a depósitos/presentaciones diferentes antes de consolidar.")
    if metadata["rejected_rows"]:
        metadata["ready_to_save"] = False
        metadata["blocking_reasons"].append("unreadable_rows")
        warnings.append(f"{len(metadata['rejected_rows'])} fila(s) parecen productos pero tienen cantidades ambiguas o ausentes. Se muestran para revisión; no se omitieron sin aviso.")
    if len(set(printed_counts)) == 1:
        _a2_compare(metadata, "row_count", len(frame), printed_counts[0], warnings)
    elif len(set(printed_counts)) > 1:
        warnings.append("Hay varios contadores de filas; no se interpretó uno como total general.")
    if detected in ("inventario", "ventas"):
        fields = {"existence": "Existencia", "inventory_value": "Valor inventario A2"} if detected == "inventario" else {"quantity": "Ventas", "gross_sales": "Monto bruto", "discounts": "Descuentos", "tax": "IVA reportado", "cogs": "Costo de ventas", "reported_profit": "Utilidad A2"}
        for field, column in fields.items():
            values = [_a2_number(value) for value in frame[column]]
            observed = _a2_sum(values)
            metadata[field] = observed if all(value is not None for value in values) else None
            raw_key = {"cogs": "cost", "reported_profit": "profit"}.get(field, field)
            candidates = [row for row in summary_rows if not row["_subtotal"] and row.get(raw_key) is not None]
            general = [row for row in candidates if row["_general"]]
            candidates = general or candidates
            printed_values = [row[raw_key] for row in candidates]
            if len(set(printed_values)) == 1 and metadata[field] is not None:
                _a2_compare(metadata, field, observed, printed_values[0], warnings)
            elif len(set(printed_values)) > 1:
                warnings.append(f"Hay varios subtotales de {field}; no se tomó uno como total general automáticamente.")
        if detected == "ventas":
            metadata["net_sales"] = _a2_sum([_a2_number(value) for value in frame["Importe neto"]]) if frame["Importe neto"].notna().all() else None
            metadata["sales_granularity"] = "period"
            metadata["availability"].update({"sales_units": True, "sales_amounts": metadata["net_sales"] is not None, "costs_complete": bool(frame["Costo de ventas"].notna().all())})
            if metadata["period_start"] and metadata["period_end"]:
                start, end = datetime.fromisoformat(metadata["period_start"]), datetime.fromisoformat(metadata["period_end"])
                if end < start:
                    metadata["ready_to_save"] = False
                    metadata["blocking_reasons"].append("invalid_period")
                    warnings.append("La fecha final del período es anterior a la inicial.")
                else:
                    metadata["period_days"] = (end - start).days + 1
                if metadata["report_date"] and metadata["period_end"] > metadata["report_date"]:
                    metadata["date_requires_confirmation"] = True
                    warnings.append(f"El período termina el {metadata['period_end']}, después de la emisión {metadata['report_date']}. Confirma o corrige el período real antes de calcular demanda; no se corrigió la fecha automáticamente.")
            warnings.append("Ventas agrupadas por producto para todo el período. Este archivo no identifica ventas a crédito, cobros ni fechas individuales.")
            if metadata["currency"] is None and frame["Monto bruto"].notna().any():
                warnings.append("El reporte de ventas no indica moneda: selecciónala antes de usar los importes.")
            if metadata["net_sales"] is None:
                warnings.append("Se pueden analizar unidades vendidas. Falta un importe neto completo: los indicadores monetarios de ventas quedan no disponibles.")
            if metadata["net_sales"] is not None and metadata.get("cogs") is not None and metadata.get("reported_profit") is not None:
                metadata["profit_identity_difference"] = round(metadata["net_sales"] - metadata["cogs"] - metadata["reported_profit"], 6)
                if abs(metadata["profit_identity_difference"]) > 0.02:
                    warnings.append("La utilidad impresa por A2 difiere de neto menos costo. Se conservan ambos valores sin corregir el origen.")
        else:
            costs_complete = bool(frame["Costo"].notna().all())
            metadata["availability"].update({"inventory_units": True, "costs_complete": costs_complete, "inventory_value": metadata.get("inventory_value") is not None or costs_complete})
            metadata["missing_cost_rows"] = int(frame["Costo"].isna().sum())
            metadata["cost_missing"] = not costs_complete
            metadata["inventory_value_calculated"] = _a2_sum([float(Decimal(str(record["Existencia"])) * Decimal(str(record["Costo"]))) for record in records]) if costs_complete else None
            differences = [record for record in records if record["Costo"] is not None and record["Valor inventario A2"] is not None and abs(record["Existencia"] * record["Costo"] - record["Valor inventario A2"]) > 0.015]
            metadata["valuation_differences"] = len(differences)
            if differences:
                warnings.append(f"{len(differences)} producto(s) presentan diferencia entre costo × existencia y valor impreso por A2; se conserva el valor original para auditoría.")
            if not costs_complete:
                warnings.append(f"{metadata['missing_cost_rows']} producto(s) sin costo: cantidades y reposición pueden calcularse; su valoración queda no disponible. No se interpretó costo ausente como cero.")
            for label in ("Proveedor", "Marca", "Categoría", "Departamento"):
                if frame[label].eq("").all():
                    metadata.setdefault("missing_dimensions", []).append(label)
            if metadata.get("missing_dimensions"):
                warnings.append("El inventario no incluye " + ", ".join(metadata["missing_dimensions"]) + ". Se completan con relaciones confirmadas del catálogo.")
    if metadata["reconciliation"]:
        metadata["reconciled"] = all(item["matches"] for item in metadata["reconciliation"])
    if detected == "departamentos":
        warnings.append("El reporte lista departamentos, pero no contiene una relación explícita con proveedores o marcas. Clasificación no equivale a categoría de producto.")
    catalog = {"departments": frame.to_dict("records") if detected == "departamentos" else [], "suppliers": frame.to_dict("records") if detected == "proveedores" else [], "relationships": []}
    return {"kind": detected, "data": frame, "metadata": metadata, "warnings": warnings, "catalog": catalog}




# Carga unificada de una o varias sedes. Este bloque se integra en el único .py.
V71_UPLOAD_MAX_FILES = 30
V71_UPLOAD_MAX_FILE = 20 * 1024 * 1024
V71_UPLOAD_MAX_TOTAL = 100 * 1024 * 1024
V71_UPLOAD_MAX_ROWS = 300000
V71_UPLOAD_KINDS = {"Inventario": "inventario", "Ventas": "ventas", "CxC (opcional)": "cxc"}


def _upload_date_v71(value):
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return ""
    if isinstance(value, (datetime, date)):
        return value.date().isoformat() if isinstance(value, datetime) else value.isoformat()
    try:
        return date.fromisoformat(str(value)[:10]).isoformat()
    except (TypeError, ValueError):
        return ""


def _known_site_v71(value):
    return next((site for site in SEDES if normalize_key(site) == normalize_key(value or "")), "")


def _upload_budget_v71(files):
    if len(files) > V71_UPLOAD_MAX_FILES:
        raise ValueError(f"Carga hasta {V71_UPLOAD_MAX_FILES} archivos por lote.")
    total = 0
    for file in files:
        size = getattr(file, "size", None)
        size = len(file.getbuffer()) if size is None else size
        if not size:
            raise ValueError(f"{file.name}: el archivo está vacío.")
        if size > V71_UPLOAD_MAX_FILE:
            raise ValueError(f"{file.name}: supera los 20 MB permitidos por archivo.")
        total += size
    if total > V71_UPLOAD_MAX_TOTAL:
        raise ValueError("Este lote supera 100 MB. Divídelo en varias cargas.")


@st.cache_data(show_spinner=False, max_entries=32, ttl=900)
def _cached_upload_parse_v71(digest, name, kind, options_json, _raw):
    # El digest forma parte de la clave: cambiar el contenido invalida la lectura.
    options = json.loads(options_json or "{}")
    result = parse_a2_report(_raw, name, kind, **options)
    if len(result["data"]) > 100000:
        raise ValueError("Divide este reporte en archivos de hasta 100.000 filas.")
    result["metadata"]["sha256"] = digest
    result["metadata"].setdefault("filename", name)
    return result


def _parse_file_v71(file, kind="auto", options=None):
    raw = file.getvalue()
    digest = hashlib.sha256(raw).hexdigest()
    return _cached_upload_parse_v71(digest, file.name, kind, json.dumps(options or {}, sort_keys=True), raw)


def _probe_upload_v71(file, forced_kind=""):
    if forced_kind == "cxc":
        return {"kind": "cxc", "metadata": {}, "error": ""}
    try:
        parsed = _parse_file_v71(file, forced_kind or "auto")
        return {"kind": parsed["kind"], "metadata": parsed["metadata"], "parsed": parsed, "error": ""}
    except Exception as exc:
        hint = normalize_key(file.name)
        guess = "cxc" if any(word in hint for word in ("cxc", "por cobrar", "cartera")) else ""
        return {"kind": forced_kind or guess, "metadata": {}, "error": str(exc)}


def _excel_column_label_v71(index):
    result, number = "", int(index) + 1
    while number:
        number, rest = divmod(number - 1, 26)
        result = chr(65 + rest) + result
    return result


def _manual_read_v71(file, kind, key):
    inspection = inspect_a2_report(file.getvalue(), file.name, kind)
    sheets = inspection.get("sheets", [])
    if not sheets:
        raise ValueError("El archivo no contiene hojas con celdas legibles.")
    names = [str(sheet["name"]) for sheet in sheets]
    selected = st.selectbox("Hoja que contiene este reporte", names, key=key + "_sheet")
    sheet = sheets[names.index(selected)]
    candidates = [candidate for candidate in sheet.get("candidates", []) if candidate.get("kind") == kind]
    best = max(candidates, key=lambda row: row.get("score", 0)) if candidates else {}
    preview = pd.DataFrame(sheet.get("preview", []))
    if not preview.empty:
        st.dataframe(preview, hide_index=True, width="stretch", height=260)
    header = st.number_input("Fila donde empieza el encabezado (la primera fila es 1)",
                             min_value=1, max_value=max(int(sheet.get("rows") or 1), 1),
                             value=max(int(best.get("header_row") or 1), 1), step=1, key=key + "_header")
    count = int(sheet.get("columns") or 0)
    options = [-1] + list(range(count))
    labels = inspection.get("fields", {})
    wanted = (["code", "description", "existence", "cost", "inventory_value", "supplier", "brand", "category", "department"]
              if kind == "inventario" else ["code", "description", "quantity", "gross_sales", "net_sales", "discounts", "tax", "cost", "profit"])
    friendly = {"code": "Código / SKU", "description": "Descripción", "existence": "Existencia", "cost": "Costo", "quantity": "Unidades vendidas",
                "inventory_value": "Valor de inventario", "gross_sales": "Monto bruto", "net_sales": "Importe neto", "discounts": "Descuentos", "tax": "IVA", "profit": "Utilidad",
                "supplier": "Proveedor", "brand": "Marca", "category": "Categoría", "department": "Departamento"}
    mapping = {}
    controls = st.columns(2)
    for index, field in enumerate(wanted):
        suggested = best.get("data_mapping", {}).get(field)
        initial = int(suggested) if suggested is not None else -1
        chosen = controls[index % 2].selectbox(str(labels.get(field, friendly.get(field, field))), options,
            index=options.index(initial) if initial in options else 0,
            format_func=lambda value: "No está en el archivo" if value < 0 else _excel_column_label_v71(value),
            key=key + "_map_" + field)
        if chosen >= 0:
            mapping[field] = int(chosen)
    st.caption("Selecciona la columna donde aparecen los DATOS de cada campo. En los reportes impresos de A2 el título puede estar desplazado. Los valores ausentes siguen sin dato; no se convierten en cero.")
    return _parse_file_v71(file, kind, {"sheet_name": selected, "header_row": int(header), "mapping": mapping})



def _a2_diagnostic_frames_v72(parsed, enriched=None):
    """Return evidence already supplied by the reader; never recompute acceptance."""
    metadata = parsed.get("metadata", {})
    labels = {"row_count": "Productos / filas", "existence": "Existencias (unidades)",
              "inventory_value": "Valor de inventario", "quantity": "Unidades vendidas",
              "gross_sales": "Ventas brutas", "net_sales": "Ventas netas", "discounts": "Descuentos",
              "tax": "IVA", "cogs": "Costo de ventas", "reported_profit": "Utilidad impresa"}
    comparison = []
    for item in metadata.get("reconciliation", []):
        comparison.append({"Total revisado": labels.get(item.get("field"), item.get("field", "Total")),
            "Impreso en A2": item.get("printed"), "Leído de las filas": item.get("extracted"),
            "Diferencia": item.get("difference"), "Tolerancia": item.get("tolerance"),
            "Resultado": "Coincide" if item.get("exact") else "Redondeo admitido" if item.get("matches") else "Revisar"})
    rejected = pd.DataFrame(metadata.get("rejected_rows", []))
    if not rejected.empty:
        rejected = rejected.rename(columns={"motivo": "Motivo"})
    frame = parsed.get("data", pd.DataFrame())
    coverage = []
    if parsed.get("kind") == "inventario":
        missing = {"", "nan", "none", "<na>", "n/d", "-", "sin clasificar", "sin departamento", "sin proveedor", "sin marca"}
        def informed(data, field):
            values = data.get(field, pd.Series("", index=data.index)).fillna("").astype(str).str.strip().str.casefold()
            return int((~values.isin(missing)).sum())
        for field in ("Departamento", "Proveedor", "Marca"):
            raw_count = informed(frame, field)
            confirmed_count = informed(enriched, field) if isinstance(enriched, pd.DataFrame) else raw_count
            coverage.append({"Clasificación": field, "Filas en el archivo": raw_count,
                             "Con relaciones confirmadas": confirmed_count,
                             "Pendientes": max(0, len(frame) - confirmed_count), "Total de productos": len(frame)})
    return pd.DataFrame(comparison), rejected, pd.DataFrame(coverage)


def _a2_rejected_cells_v72(file, parsed, limit=30):
    """Show original cells from the chosen sheet, only when requested for review."""
    metadata = parsed.get("metadata", {})
    numbers = []
    for item in metadata.get("rejected_rows", [])[:limit]:
        value = item.get("Fila origen")
        if isinstance(value, (int, np.integer)) and value > 0:
            numbers.append(int(value))
    if not numbers:
        return pd.DataFrame()
    sheets, _ = _a2_read_grids(file.getvalue(), file.name)
    selected = [(name, rows) for name, rows in sheets if name == metadata.get("sheet")]
    if not selected and len(sheets) == 1:
        selected = sheets
    if len(selected) != 1:
        raise ValueError("Selecciona de nuevo la hoja del reporte para revisar sus celdas originales.")
    name, grid = selected[0]
    evidence = []
    for number in numbers:
        if number <= len(grid):
            values = [f"Columna {index + 1}: {_a2_text(value)}" for index, value in enumerate(grid[number - 1]) if _a2_text(value)]
            evidence.append({"Hoja": name, "Fila origen": number, "Celdas originales": " | ".join(values)})
    return pd.DataFrame(evidence)


def _render_a2_diagnostics_v72(file, parsed, key):
    metadata = parsed.get("metadata", {})
    enriched = None
    if parsed.get("kind") == "inventario":
        try:
            enriched = enrich_dimensions_v7(parsed["data"])
        except (sqlite3.Error, KeyError, ValueError):
            st.caption("La cobertura muestra los datos del archivo; no se pudo consultar la relación guardada del catálogo.")
    comparison, rejected, coverage = _a2_diagnostic_frames_v72(parsed, enriched)
    st.markdown("**Comprobación de totales**")
    if comparison.empty:
        st.caption("No se identificó un total general impreso comparable. Esto no confirma que la lectura esté completa: revisa las filas detectadas y las columnas del archivo.")
    else:
        st.dataframe(comparison, hide_index=True, width="stretch",
            column_config={name: st.column_config.NumberColumn(format="%.6f")
                           for name in ("Impreso en A2", "Leído de las filas", "Diferencia", "Tolerancia")})
        st.caption("Diferencia = leído menos impreso. La tolerancia sólo admite el redondeo que indica cada fila; no cambia los datos ni permite omitir productos.")
        if metadata.get("reconciled") is False:
            st.info("Revisa la fila marcada «Revisar». En «Elegir hoja, encabezado o columnas», confirma la hoja y las columnas donde están las cantidades y los importes. Si A2 imprime un total incorrecto, vuelve a exportar el reporte corregido.")
    if not rejected.empty:
        st.markdown(f"**Filas pendientes de lectura: {len(rejected):,}**")
        st.dataframe(rejected.head(30), hide_index=True, width="stretch")
        st.caption("Fila origen es el número de fila en la hoja original. Se muestran hasta 30 casos; la carga queda pendiente hasta resolverlos.")
        if st.checkbox("Ver los valores originales de estas filas", key=key + "_rejected_cells"):
            try:
                cells = _a2_rejected_cells_v72(file, parsed)
                if not cells.empty:
                    st.dataframe(cells, hide_index=True, width="stretch")
                else:
                    st.caption("El lector no conservó una referencia de fila para estos casos.")
            except (ValueError, OSError, ImportError) as exc:
                st.warning(f"No se pudo abrir la vista de celdas originales: {exc}")
    if not coverage.empty:
        st.markdown("**Áreas y clasificación de los productos**")
        st.dataframe(coverage, hide_index=True, width="stretch")
        st.caption("Sede = almacén o sucursal. Departamento = área o familia del producto. Esta revisión cuenta los campos del archivo y las relaciones que ya confirmaste en el catálogo.")
        pending = coverage.loc[coverage["Pendientes"] > 0, "Clasificación"].tolist()
        if pending:
            st.info("Faltan datos de " + ", ".join(pending) + ". Puedes cargar inventario y ventas; esos productos quedan sin clasificar hasta completar su relación por Código en «Proveedores y departamentos». Si el archivo sí trae esos campos, ajusta sus columnas en la lectura manual.")
            st.caption("Un listado de nombres de departamentos no indica a qué área pertenece cada producto. Se necesita Código → Departamento / Proveedor / Marca, o una regla que tú confirmes; no se asignan áreas por intuición.")


def _review_a2_file_v71(file, kind, key, initial=None):
    parsed, error = None, ""
    try:
        parsed = initial if initial is not None and initial.get("kind") == kind else _parse_file_v71(file, kind)
    except Exception as exc:
        error = str(exc)
    with st.expander(f"Revisar lectura · {file.name}", expanded=bool(error or (parsed and (parsed.get("metadata", {}).get("reconciled") is False or parsed.get("metadata", {}).get("rejected_rows"))))):
        if error:
            st.warning(error)
        manual = st.checkbox("Elegir hoja, encabezado o columnas", value=bool(error), key=key + "_manual")
        if manual:
            try:
                parsed = _manual_read_v71(file, kind, key)
                error = ""
            except Exception as exc:
                parsed, error = None, str(exc)
                st.error(error)
        if parsed:
            st.caption(f"{len(parsed['data']):,} filas detectadas · hoja {parsed['metadata'].get('sheet', 'seleccionada')}")
            st.dataframe(parsed["data"].head(12), hide_index=True, width="stretch")
            _render_a2_diagnostics_v72(file, parsed, key)
            for warning in parsed.get("warnings", [])[:8]:
                st.caption(str(warning))
            if parsed["metadata"].get("reconciled") is False:
                error = "Los datos extraídos no concilian con los totales impresos. Revisa la lectura o el archivo antes de guardar."
                st.error(error)
            # Códigos faltantes de ventas y fechas se resuelven en la misma carga.
            resolvable = {"invalid_period"} | ({"missing_codes"} if kind == "ventas" else set())
            blocking = [reason for reason in parsed["metadata"].get("blocking_reasons", []) if reason not in resolvable]
            if blocking:
                messages = {"missing_codes": "Hay productos sin código", "duplicate_codes": "Hay códigos repetidos que requieren revisión",
                    "unreadable_rows": "Hay filas de productos con cantidades ilegibles", "total_mismatch": "Los totales no concilian con el reporte"}
                error = "; ".join(messages.get(reason, str(reason)) for reason in blocking)
                st.error(error)
            if kind == "inventario":
                bad = parsed["data"]["Código"].astype(str).str.strip().eq("") | parsed["data"]["Código"].duplicated(keep=False)
                if bad.any():
                    error = "El inventario tiene códigos vacíos o repetidos. Corrige el código en A2 o separa los depósitos antes de guardarlo."
                    st.error(error)
                    st.dataframe(parsed["data"].loc[bad].head(30), hide_index=True, width="stretch")
    return parsed, error


def _review_ar_file_v71(file, assignment, key):
    with st.expander(f"Revisar CxC · {file.name}", expanded=False):
        try:
            frame, mapping, grid = detect_cxc_v7(file.getvalue(), file.name)
            controls = st.columns(2)
            decimal = "," if controls[0].selectbox("Decimales de CxC en texto", ["Punto · 1,234.56", "Coma · 1.234,56"], key=key + "_decimal").startswith("Coma") else "."
            day_first = controls[1].selectbox("Fechas de CxC", ["Día/mes/año", "Mes/día/año"], key=key + "_dates").startswith("Día")
            normalized = normalize_financial_frame(frame, "cxc", mapping, assignment["site"], assignment["currency"], decimal, day_first, assignment["cutoff"])
            st.dataframe(normalized.head(20), hide_index=True, width="stretch")
            st.caption(f"{len(normalized):,} facturas · saldo {fin_value(normalized['Saldo'].sum(), assignment['currency'])}")
            return normalized, ""
        except Exception as exc:
            st.error(str(exc))
            st.caption("Para una cartera con otras columnas utiliza 'Opciones avanzadas → Otros reportes financieros'.")
            return None, str(exc)


def validate_multisite_assignments_v71(assignments, scope="", today=None):
    """Valida TODO el lote antes de abrir la primera transacción por sede."""
    current_day = _upload_date_v71(today or date.today())
    errors, groups, seen = [], {}, set()
    for position, item in enumerate(assignments, start=1):
        label = item.get("filename") or f"Archivo {position}"
        site, currency, kind = item.get("site", ""), item.get("currency", ""), item.get("kind", "")
        cutoff = _upload_date_v71(item.get("cutoff"))
        if site not in SEDES or currency not in FIN_CURRENCIES:
            errors.append(f"{label}: confirma una sede y una moneda válidas.")
            continue
        if not cutoff or cutoff > current_day:
            errors.append(f"{label}: el corte debe ser una fecha válida y no futura.")
            continue
        if kind not in {"inventario", "ventas", "cxc"}:
            errors.append(f"{label}: selecciona el tipo de reporte.")
            continue
        if item.get("error"):
            errors.append(f"{label}: {item['error']}")
            continue
        parsed = item.get("parsed")
        if kind != "cxc" and parsed is None:
            errors.append(f"{label}: la lectura todavía no es válida.")
            continue
        if kind == "cxc" and item.get("ar") is None:
            errors.append(f"{label}: la cartera todavía no es válida.")
            continue
        meta = parsed["metadata"] if parsed else item.get("metadata", {})
        detected_site = _known_site_v71(meta.get("site"))
        detected_currency = str(meta.get("currency") or "").upper()
        if detected_site and detected_site != site:
            errors.append(f"{label}: el archivo identifica {detected_site}, pero seleccionaste {site}.")
        if detected_currency in FIN_CURRENCIES and detected_currency != currency:
            errors.append(f"{label}: sus importes están en {detected_currency}. Seleccionar {currency} no convierte el dinero.")
        if meta.get("reconciled") is False:
            errors.append(f"{label}: los totales no concilian con A2.")
        resolvable = {"invalid_period"} | ({"missing_codes"} if kind == "ventas" else set())
        if any(reason not in resolvable for reason in meta.get("blocking_reasons", [])):
            errors.append(f"{label}: hay filas ambiguas, identificadores o totales pendientes de corregir en la lectura.")
        digest = item.get("digest") or meta.get("sha256", "")
        if digest:
            # Un archivo idéntico no puede convertirse en dos sedes por asignación.
            identity = (digest, meta.get("sheet", ""))
            if identity in seen:
                errors.append(f"{label}: el mismo archivo aparece más de una vez en este lote.")
            seen.add(identity)
        group = groups.setdefault(site, {"site": site, "currency": currency, "cutoff": cutoff,
            "inventory": None, "sales_reports": [], "ar": None, "scope": scope,
            "replace_overlaps": False, "files": []})
        if group["currency"] != currency:
            errors.append(f"{site}: los archivos del lote deben usar una misma moneda. Carga monedas distintas por separado.")
        if group["cutoff"] != cutoff:
            errors.append(f"{site}: asigna el mismo corte a los archivos de esta sede o sepáralos en dos cargas.")
        group["files"].append(label)
        if kind == "inventario":
            if group["inventory"] is not None:
                errors.append(f"{site}: hay más de un inventario. Usa un corte completo por sede; no se suman archivos potencialmente repetidos.")
            data = parsed["data"]
            if data.empty or data["Código"].astype(str).str.strip().eq("").any() or data["Código"].duplicated().any():
                errors.append(f"{label}: el inventario necesita productos con códigos únicos y no vacíos.")
            group["inventory"] = parsed
        elif kind == "ventas":
            start, end = _upload_date_v71(item.get("start")), _upload_date_v71(item.get("end"))
            if not start or not end:
                errors.append(f"{label}: indica la fecha inicial y final del reporte de ventas.")
            elif start > end:
                errors.append(f"{label}: el inicio {start} es posterior al final {end}; revisa el período.")
            elif end > current_day:
                errors.append(f"{label}: el final {end} es posterior a hoy ({current_day}); confirma el año y las fechas del archivo.")
            elif end > cutoff:
                errors.append(f"{label}: las ventas terminan el {end}, después del corte de inventario {cutoff}. Usa un inventario posterior o exporta ventas hasta ese corte.")
            metadata = dict(meta)
            if metadata.get("period_start") != start or metadata.get("period_end") != end:
                metadata.setdefault("source_period_start", metadata.get("period_start"))
                metadata.setdefault("source_period_end", metadata.get("period_end"))
                metadata["period_confirmed_by_user"] = True
            metadata["period_start"], metadata["period_end"] = start, end
            if start and end and start <= end:
                metadata["period_days"] = (date.fromisoformat(end) - date.fromisoformat(start)).days + 1
            group["sales_reports"].append(dict(parsed, metadata=metadata))
        else:
            if scope:
                errors.append(f"{site}: CxC se carga para la sede completa, no para un proveedor.")
            if group["ar"] is not None:
                errors.append(f"{site}: hay más de un archivo de CxC. Usa una cartera completa por sede.")
            group["ar"] = item["ar"]
    for site, group in groups.items():
        spans = sorted((p["metadata"]["period_start"], p["metadata"]["period_end"]) for p in group["sales_reports"])
        for previous, following in zip(spans, spans[1:]):
            if following[0] and previous[1] and following[0] <= previous[1]:
                errors.append(f"{site}: dos reportes de ventas comparten días. Selecciona períodos sin solapamiento.")
    return list(dict.fromkeys(errors)), list(groups.values())


def _bundle_fingerprint_v71(bundle):
    pieces = [{key: value for key, value in bundle.items() if key not in {"inventory", "sales_reports", "ar", "files"}}]
    for parsed in ([bundle["inventory"]] if bundle.get("inventory") is not None else []) + bundle.get("sales_reports", []):
        pieces.append({"metadata": parsed["metadata"], "data": parsed["data"].to_json(orient="split", date_format="iso", force_ascii=False)})
    if bundle.get("ar") is not None:
        pieces.append({"cxc": bundle["ar"].to_json(orient="split", date_format="iso", force_ascii=False)})
    return hashlib.sha256(json.dumps(pieces, sort_keys=True, ensure_ascii=False, default=str).encode()).hexdigest()


def save_multisite_bundles_v71(bundles, completed=None):
    """Cada sede es atómica. Un reintento omite sólo las sedes confirmadas de ESTE lote."""
    completed = completed if completed is not None else {}
    result = {"saved": [], "skipped": [], "errors": [], "completed": completed}
    for bundle in bundles:
        fingerprint = _bundle_fingerprint_v71(bundle)
        if fingerprint in completed:
            result["skipped"].append({"Sede": bundle["site"], **completed[fingerprint]})
            continue
        try:
            counts = save_upload_bundle_v7(bundle["site"], bundle["currency"], bundle["cutoff"], bundle.get("inventory"),
                bundle.get("sales_reports", []), bundle.get("ar"), bundle.get("scope", ""), bool(bundle.get("replace_overlaps")))
            completed[fingerprint] = counts
            result["saved"].append({"Sede": bundle["site"], **counts})
        except Exception as exc:
            result["errors"].append({"Sede": bundle["site"], "Error": str(exc)})
            # Evita continuar un lote después de un fallo de disco/base; lo confirmado se conserva.
            break
    return result


def _resolve_bundle_codes_v71(assignments, key):
    inventories = {item["site"]: item["parsed"]["data"] for item in assignments if item.get("kind") == "inventario" and item.get("parsed") is not None}
    valid = True
    for index, item in enumerate(assignments):
        if item.get("kind") != "ventas" or item.get("parsed") is None:
            continue
        parsed = item["parsed"]
        data, matched = resolve_sales_codes_v7(parsed["data"], inventories.get(item["site"], pd.DataFrame()))
        if matched:
            st.caption(f"{item['filename']}: {matched} códigos recuperados por descripción exacta y única en el inventario de {item['site']}.")
        missing = data["Código"].astype(str).str.strip().eq("")
        if missing.any():
            with st.expander(f"Completar {int(missing.sum())} códigos · {item['site']} · {item['filename']}", expanded=True):
                st.caption("Escribe los códigos reales o conserva estas filas pendientes. Sus importes siguen presentes; la reposición por producto requiere un código identificado.")
                known = inventories.get(item["site"], pd.DataFrame())
                suggestions = []
                if not known.empty:
                    descriptions = known["Descripción"].astype(str).map(normalize_key)
                    for missing_row in data.loc[missing].head(20).to_dict("records"):
                        description = normalize_key(missing_row["Descripción"])
                        if len(description) < 15:
                            continue
                        candidates = known.loc[descriptions.str.startswith(description, na=False)].head(5)
                        for candidate in candidates.to_dict("records"):
                            suggestions.append({"Descripción A2": missing_row["Descripción"], "Código candidato": candidate["Código"], "Descripción en inventario": candidate["Descripción"]})
                if suggestions:
                    st.caption("Posibles coincidencias por descripción truncada. Comprueba la presentación y escribe el código sólo si corresponde; no se asignan automáticamente.")
                    st.dataframe(pd.DataFrame(suggestions), hide_index=True, width="stretch")
                cols = [col for col in ["Código", "Descripción", "Ventas", "Fila origen"] if col in data]
                corrected = st.data_editor(data.loc[missing, cols], disabled=[col for col in cols if col != "Código"], hide_index=True, width="stretch", key=f"{key}_codes_{index}")
                data.loc[missing, "Código"] = corrected["Código"].fillna("").astype(str).str.strip().values
                if data["Código"].eq("").any():
                    accepted = st.checkbox("Guardar los códigos faltantes como pendientes", key=f"{key}_pending_{index}")
                    valid = valid and accepted
        parsed["data"] = data
    return valid


def _preflight_stored_periods_v71(bundles):
    problems, replace_sites = [], set()
    for bundle in bundles:
        for parsed in bundle["sales_reports"]:
            meta = parsed["metadata"]
            if not meta.get("period_start") or not meta.get("period_end"):
                continue
            old = sales_overlaps_v7(bundle["site"], bundle["scope"], meta["period_start"], meta["period_end"])
            for prior in old.to_dict("records"):
                exact = prior["start_date"] == meta["period_start"] and prior["end_date"] == meta["period_end"]
                covered = meta["period_start"] <= prior["start_date"] and prior["end_date"] <= meta["period_end"]
                if not exact and not covered:
                    problems.append(f"{bundle['site']}: {meta['period_start']}–{meta['period_end']} se cruza parcialmente con {prior['start_date']}–{prior['end_date']}. Elige períodos completos sin días repetidos.")
                elif not exact:
                    replace_sites.add(bundle["site"])
    return problems, replace_sites


def render_load_center_v7():
    st.subheader("Carga de datos")
    st.caption("Carga una sede o varias en el mismo lote. Inventario y ventas permiten trabajar sin CxC; la cartera se incorpora cuando la tengas.")
    load_tab, catalogs_tab = st.tabs(["Cargar reportes", "Catálogos y relaciones"])
    with load_tab:
        notice = st.session_state.get("v7_import_success")
        if notice:
            st.success(notice)
        mode = st.radio("¿Qué vas a cargar?", ["Una sede", "Varias sedes"], horizontal=True, key="v71_load_mode")
        files, forced = [], {}
        batch_sites = list(SEDES)
        if mode == "Una sede":
            cards = st.columns(2)
            inv_file = cards[0].file_uploader("1 · Inventario", type=["xls", "xlsx", "csv"], max_upload_size=20, key="v71_single_inventory")
            sales_files = cards[1].file_uploader("2 · Ventas (uno o varios períodos)", type=["xls", "xlsx", "csv"], accept_multiple_files=True, max_upload_size=20, key="v71_single_sales")
            ar_file = None
            if st.checkbox("Añadir cuentas por cobrar (opcional)", key="v71_include_ar"):
                ar_file = st.file_uploader("Reporte de CxC", type=["xls", "xlsx", "csv"], max_upload_size=20, key="v71_single_ar")
                st.caption("Dejarlo vacío mantiene CxC como no cargada; permite generar el reporte de inventario y ventas.")
            for file, kind in [(inv_file, "inventario")] + [(file, "ventas") for file in sales_files] + [(ar_file, "cxc")]:
                if file is not None:
                    forced[len(files)] = kind
                    files.append(file)
        else:
            batch_sites = site_selector_v712("Sedes que vas a actualizar en este lote", "v712_load_sites")
            st.caption("Seleccionar todas habilita las cinco sedes para este lote. Cada archivo debe pertenecer a una sola sede; la aplicación no copia un mismo inventario a todas.")
            files = st.file_uploader("Inventarios y ventas de las sedes (CxC opcional)", type=["xls", "xlsx", "csv"], accept_multiple_files=True, max_upload_size=20, key="v71_many_files")
            st.caption("Hasta 30 archivos · 20 MB por archivo · 100 MB por lote. Asigna cada archivo a su sede; no se mezclan almacenes.")
        try:
            _upload_budget_v71(files)
        except ValueError as exc:
            st.error(str(exc))
            return
        probes, parsed_rows = [], 0
        for index, file in enumerate(files):
            probe = _probe_upload_v71(file, forced.get(index, ""))
            probes.append(probe)
            parsed_rows += len(probe["parsed"]["data"]) if probe.get("parsed") is not None else 0
            if parsed_rows > V71_UPLOAD_MAX_ROWS:
                st.error("El lote supera 300.000 filas útiles. Divídelo en varias cargas para no saturar la memoria.")
                return
        fingerprint = hashlib.sha256(repr([(file.name, hashlib.sha256(file.getvalue()).hexdigest()) for file in files]).encode()).hexdigest()[:16]
        key = "v71_batch_" + mode.replace(" ", "_") + "_" + fingerprint
        scope = ""
        dedicated, zero_ar = False, False
        with st.expander("Opciones avanzadas", expanded=False):
            dedicated = st.checkbox("Los inventarios y ventas de este lote son de un único proveedor", key=key + "_dedicated")
            if dedicated:
                names = catalog_frame_v7("proveedores")["Descripción"].tolist()
                scope = st.selectbox("Proveedor de estos reportes", [""] + names, format_func=lambda text: text or "Seleccionar proveedor", key=key + "_scope")
                st.caption("Se guarda un corte separado del inventario general de cada sede.")
            st.caption("La opción cartera cero sólo se ofrece en una sede y requiere confirmación explícita. No tener el archivo de CxC no significa saldo cero.")
        assignments = []
        if files:
            labels = {value: label for label, value in V71_UPLOAD_KINDS.items()}
            inv_meta = next((probe["metadata"] for probe in probes if probe["kind"] == "inventario"), {})
            suggested_site = _known_site_v71(inv_meta.get("site"))
            suggested_currency = inv_meta.get("currency") if inv_meta.get("currency") in FIN_CURRENCIES else ""
            suggested_cutoff = _upload_date_v71(inv_meta.get("snapshot_date")) or date.today().isoformat()
            site, currency, cutoff = suggested_site, suggested_currency, suggested_cutoff
            if mode == "Una sede":
                controls = st.columns(3)
                site = controls[0].selectbox("Sede de los archivos", [""] + SEDES, index=([""] + SEDES).index(suggested_site), format_func=lambda value: value or "Seleccionar sede", key=key + "_site")
                currency = controls[1].selectbox("Moneda de los importes", [""] + FIN_CURRENCIES, index=([""] + FIN_CURRENCIES).index(suggested_currency), format_func=lambda value: value or "Seleccionar moneda", key=key + "_currency")
                cutoff = controls[2].date_input("Fecha de corte", date.fromisoformat(suggested_cutoff), key=key + "_cutoff").isoformat()
            defaults = []
            for index, (file, probe) in enumerate(zip(files, probes)):
                meta = probe["metadata"]
                row = {"Archivo": file.name, "Tipo": labels.get(probe["kind"], "Por asignar"),
                    "Sede": site if mode == "Una sede" else _known_site_v71(meta.get("site")),
                    "Moneda": currency if mode == "Una sede" else (meta.get("currency") if meta.get("currency") in FIN_CURRENCIES else ""),
                    "Corte": date.fromisoformat(cutoff if mode == "Una sede" else (_upload_date_v71(meta.get("snapshot_date")) if probe["kind"] == "inventario" else "") or date.today().isoformat()),
                    "Desde": date.fromisoformat(meta["period_start"]) if _upload_date_v71(meta.get("period_start")) else None,
                    "Hasta": date.fromisoformat(meta["period_end"]) if _upload_date_v71(meta.get("period_end")) else None}
                defaults.append(row)
            st.markdown("**Revisa la asignación y las fechas de cada archivo**")
            st.caption("Las sugerencias provienen de las celdas del archivo. Confirma la sede y moneda en A2. Desde/Hasta sólo se usa para ventas; corrige cualquier año equivocado en el reporte de origen.")
            disabled = ["Archivo"] + (["Sede", "Moneda", "Corte", "Tipo"] if mode == "Una sede" else [])
            edited = st.data_editor(pd.DataFrame(defaults), hide_index=True, width="stretch", num_rows="fixed", disabled=disabled,
                column_config={"Tipo": st.column_config.SelectboxColumn(options=["Por asignar"] + list(V71_UPLOAD_KINDS)),
                    "Sede": st.column_config.SelectboxColumn(options=[""] + SEDES),
                    "Moneda": st.column_config.SelectboxColumn(options=[""] + FIN_CURRENCIES),
                    "Corte": st.column_config.DateColumn(format="DD/MM/YYYY"),
                    "Desde": st.column_config.DateColumn(format="DD/MM/YYYY"), "Hasta": st.column_config.DateColumn(format="DD/MM/YYYY")}, key=key + "_assignments")
            for index, (file, row) in enumerate(zip(files, edited.to_dict("records"))):
                item = {"filename": file.name, "digest": hashlib.sha256(file.getvalue()).hexdigest(), "kind": V71_UPLOAD_KINDS.get(row["Tipo"], ""),
                    "site": site if mode == "Una sede" else (row.get("Sede") or ""), "currency": currency if mode == "Una sede" else (row.get("Moneda") or ""),
                    "cutoff": cutoff if mode == "Una sede" else _upload_date_v71(row.get("Corte")),
                    "start": _upload_date_v71(row.get("Desde")), "end": _upload_date_v71(row.get("Hasta")), "parsed": None, "ar": None, "error": ""}
                if item["kind"] in {"inventario", "ventas"}:
                    item["parsed"], item["error"] = _review_a2_file_v71(file, item["kind"], key + "_file_" + str(index), probes[index].get("parsed"))
                elif item["kind"] == "cxc" and item["site"] in SEDES and item["currency"] in FIN_CURRENCIES and item["cutoff"]:
                    item["ar"], item["error"] = _review_ar_file_v71(file, item, key + "_file_" + str(index))
                assignments.append(item)
            unresolved_ok = _resolve_bundle_codes_v71(assignments, key)
            errors, bundles = validate_multisite_assignments_v71(assignments, scope)
            if mode == "Varias sedes":
                if not batch_sites:
                    errors.append("Selecciona al menos una sede para el lote o pulsa «Seleccionar todas las sedes».")
                outside = sorted({item["site"] for item in assignments if item["site"] in SEDES and item["site"] not in batch_sites})
                if outside:
                    errors.append("Hay archivos asignados a sedes fuera del alcance seleccionado: " + ", ".join(outside) + ". Inclúyelas en la selección o corrige la sede de esos archivos.")
                represented = {item["site"] for item in assignments if item["site"] in SEDES}
                unchanged = [name for name in batch_sites if name not in represented]
                if unchanged:
                    st.info("Sedes seleccionadas sin archivos en este lote: " + ", ".join(unchanged) + ". Sus datos guardados no se actualizarán.")
            if dedicated and not scope:
                errors.append("Selecciona el proveedor o desactiva el alcance de un único proveedor.")
            extra_errors, replace_sites = _preflight_stored_periods_v71(bundles)
            errors.extend(extra_errors)
            if replace_sites:
                st.warning("Los nuevos archivos cubren períodos ya guardados de: " + ", ".join(sorted(replace_sites)) + ". Los anteriores se archivarán; las ventas no se suman dos veces.")
                replacing = st.checkbox("Sustituir los períodos anteriores completamente cubiertos", key=key + "_replace")
                for bundle in bundles:
                    bundle["replace_overlaps"] = replacing
                if not replacing:
                    errors.append("Confirma la sustitución de períodos cubiertos o elige otro lote.")
            if mode == "Una sede" and len(bundles) == 1 and not scope and not any(item["kind"] == "cxc" for item in assignments):
                with st.expander("Confirmar que no hay facturas pendientes · avanzado", expanded=False):
                    zero_ar = st.checkbox(f"Confirmo CxC CERO para {site} en {currency} al {cutoff}", key=key + "_zero_ar")
                    st.caption("Marca únicamente si verificaste el saldo cero. Sin marcar, el reporte de almacén funciona y no se registra una cartera nueva.")
                if zero_ar:
                    bundles[0]["ar"] = pd.DataFrame(columns=FIN_AR_FIELDS)
            summary = []
            for bundle in bundles:
                inventory = bundle["inventory"]
                has_sales = bool(bundle["sales_reports"])
                summary.append({"Sede": bundle["site"], "Moneda": bundle["currency"], "Corte": bundle["cutoff"],
                    "Productos": len(inventory["data"]) if inventory is not None else "Sin inventario nuevo",
                    "Filas ventas": sum(len(report["data"]) for report in bundle["sales_reports"]),
                    "CxC": f"{len(bundle['ar']):,} facturas" if bundle["ar"] is not None else "No incluida (opcional)",
                    "Disponible con esta carga": "Stock, rotación y reposición" if inventory is not None and has_sales else ("Stock y valoración disponible" if inventory is not None else "Actualizar ventas/cartera guardada")})
            if summary:
                st.markdown("**Resumen del lote**")
                st.dataframe(pd.DataFrame(summary), hide_index=True, width="stretch")
            for error in list(dict.fromkeys(errors)):
                st.error(error)
            if not unresolved_ok:
                st.info("Completa los códigos pendientes o confirma que deseas conservar esas filas para identificarlas después.")
            reviewed = st.checkbox("Revisé la sede, moneda, fechas y datos detectados de cada archivo", key=key + "_reviewed")
            can_save = bool(bundles) and not errors and unresolved_ok and reviewed
            if st.button("Guardar lote y actualizar resultados", type="primary", width="stretch", disabled=not can_save, key=key + "_save"):
                active_id = hashlib.sha256("|".join(_bundle_fingerprint_v71(bundle) for bundle in bundles).encode()).hexdigest()
                progress = st.session_state.get("v71_active_import", {})
                if progress.get("id") != active_id:
                    progress = {"id": active_id, "completed": {}}
                    st.session_state["v71_active_import"] = progress
                with st.spinner("Guardando cada sede por separado…"):
                    result = save_multisite_bundles_v71(bundles, progress["completed"])
                if result["errors"]:
                    done = result["saved"] + result["skipped"]
                    if done:
                        st.success("Quedaron guardadas: " + ", ".join(row["Sede"] for row in done) + ".")
                    for failure in result["errors"]:
                        st.error(f"{failure['Sede']}: no se guardó esta sede. {failure['Error']}")
                    st.info("El lote quedó incompleto. Al reintentar sin cambiarlo se omiten las sedes ya confirmadas y se continúa con las pendientes.")
                else:
                    total = len(result["saved"]) + len(result["skipped"])
                    st.session_state["v7_import_success"] = f"Lote guardado: {total} sede(s). Inventario y ventas se pueden evaluar sin cargar CxC."
                    st.session_state.pop("v71_active_import", None)
                    _navigation_v7("Reportes")
        else:
            st.info("Empieza con inventario y ventas de una o varias sedes. Con inventario solamente puedes revisar existencias; la rotación y reposición necesitan ventas conocidas.")
        with st.expander("Historial y cargas anteriores", expanded=False):
            render_import_status_v7()
            render_saved_history_v7()
        with st.expander("Otros reportes financieros · avanzado", expanded=False):
            if st.checkbox("Abrir carga financiera adicional", key="v71_open_extra_finance"):
                extra_kind = st.radio("Reporte adicional", ["Ventas diarias / crédito / cobros", "Cartera con otras columnas"], key="v71_extra_kind")
                render_financial_import("ventas" if extra_kind.startswith("Ventas") else "cxc")
    with catalogs_tab:
        render_catalogs_v7()


# Herramientas de almacén v7.1: conteos separados del inventario administrativo.
WH_MAX_ROWS_V71 = 20000


def init_warehouse_v71():
    """Migración aditiva. Nunca modifica los cortes originales de A2."""
    with db() as con:
        con.executescript("""
        CREATE TABLE IF NOT EXISTS v71_wh_meta (key TEXT PRIMARY KEY, value TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS v71_wh_sessions (
            id TEXT PRIMARY KEY, site TEXT NOT NULL, cutoff TEXT NOT NULL,
            created_at TEXT NOT NULL, closed_at TEXT, title TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS v71_wh_items (
            session_id TEXT NOT NULL, code TEXT NOT NULL, description TEXT NOT NULL,
            expected REAL NOT NULL, counted REAL, updated_at TEXT,
            PRIMARY KEY(session_id,code), CHECK(counted IS NULL OR counted>=0));
        CREATE TABLE IF NOT EXISTS v71_wh_barcodes (
            session_id TEXT NOT NULL, barcode TEXT NOT NULL, code TEXT NOT NULL,
            PRIMARY KEY(session_id,barcode));
        CREATE TABLE IF NOT EXISTS v71_wh_events (
            id INTEGER PRIMARY KEY AUTOINCREMENT, session_id TEXT NOT NULL,
            code TEXT NOT NULL, previous REAL, quantity REAL NOT NULL,
            recorded_at TEXT NOT NULL, source TEXT NOT NULL);
        CREATE INDEX IF NOT EXISTS v71_wh_sessions_site ON v71_wh_sessions(site,created_at);
        CREATE INDEX IF NOT EXISTS v71_wh_events_session ON v71_wh_events(session_id,recorded_at);
        INSERT OR IGNORE INTO v71_wh_meta VALUES ('schema_version','1');
        """)


def _wh_text_v71(value):
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return ""
    return str(value).strip()


def _wh_norm_v71(value):
    return re.sub(r"[^a-z0-9]+", " ", unicodedata.normalize("NFKD", _wh_text_v71(value)).encode("ascii", "ignore").decode().lower()).strip()


def wh_quantity_v71(value):
    """Decimal explícito; nunca adivina separadores de miles."""
    if isinstance(value, bool):
        raise ValueError("La cantidad debe ser numérica.")
    text = _wh_text_v71(value)
    if not re.fullmatch(r"\+?\d+(?:[.,]\d+)?", text):
        raise ValueError("Usa cantidades >= 0, sin separadores de miles; por ejemplo 1200 o 12,5.")
    quantity = float(text.replace(",", "."))
    if not np.isfinite(quantity) or quantity > 1e12:
        raise ValueError("Cantidad fuera del intervalo permitido.")
    return quantity


def wh_begin_count_v71(frame, site, cutoff, title=""):
    required = {"Sede", "Código", "Descripción", "Existencia"}
    if not required.issubset(frame.columns):
        raise ValueError("El conteo necesita Sede, Código, Descripción y Existencia.")
    date.fromisoformat(str(cutoff))
    data = frame.loc[frame["Sede"].astype(str).eq(str(site)), list(required)].copy().reset_index(drop=True)
    if not 0 < len(data) <= WH_MAX_ROWS_V71:
        raise ValueError(f"Selecciona entre 1 y {WH_MAX_ROWS_V71:,} productos de una sede.")
    data["Código"] = data["Código"].map(_wh_text_v71)
    if data["Código"].eq("").any() or data["Código"].duplicated().any():
        raise ValueError("Hay códigos vacíos o repetidos en esa sede. Revisa el catálogo antes de congelar el conteo.")
    stock = pd.to_numeric(data["Existencia"], errors="coerce")
    if stock.isna().any() or not np.isfinite(stock.to_numpy(dtype=float)).all():
        raise ValueError("Hay existencias no numéricas; corrige la carga antes del conteo.")
    session_id, now = uuid.uuid4().hex, datetime.now().isoformat(timespec="seconds")
    with db() as con:
        con.execute("INSERT INTO v71_wh_sessions VALUES (?,?,?,?,NULL,?)", (session_id, str(site), str(cutoff), now, str(title).strip()[:160] or f"Conteo {site} {cutoff}"))
        con.executemany("INSERT INTO v71_wh_items(session_id,code,description,expected) VALUES (?,?,?,?)", [
            (session_id, row["Código"], _wh_text_v71(row["Descripción"]), float(stock.loc[idx])) for idx, row in data.iterrows()
        ])
    return session_id


def wh_count_sessions_v71(site=None):
    with db() as con:
        sql = "SELECT id,site,cutoff,created_at,closed_at,title FROM v71_wh_sessions"
        params = ()
        if site is not None:
            sql += " WHERE site=?"
            params = (site,)
        return pd.read_sql_query(sql + " ORDER BY created_at DESC LIMIT 100", con, params=params)


def _wh_session_v71(con, session_id, editable=False):
    row = con.execute("SELECT site,cutoff,closed_at FROM v71_wh_sessions WHERE id=?", (session_id,)).fetchone()
    if row is None:
        raise ValueError("Conteo no encontrado.")
    if editable and row[2] is not None:
        raise ValueError("El conteo está cerrado. Reábrelo antes de registrar cantidades.")
    return row


def _wh_resolve_v71(con, session_id, token):
    token = _wh_text_v71(token)
    direct = con.execute("SELECT code FROM v71_wh_items WHERE session_id=? AND code=?", (session_id, token)).fetchone()
    barcode = con.execute("SELECT code FROM v71_wh_barcodes WHERE session_id=? AND barcode=?", (session_id, token)).fetchone()
    if direct and barcode and direct[0] != barcode[0]:
        raise ValueError("Código ambiguo; revisa la correspondencia del código de barras.")
    if not direct and not barcode:
        raise ValueError(f"Código no encontrado en este conteo: {token[:90]}")
    return (direct or barcode)[0]


def wh_resolve_code_v71(session_id, token):
    with db() as con:
        _wh_session_v71(con, session_id)
        return _wh_resolve_v71(con, session_id, token)


def _wh_record_count_v71(con, session_id, code, quantity, source):
    now = datetime.now().isoformat(timespec="microseconds")
    old = con.execute("SELECT counted FROM v71_wh_items WHERE session_id=? AND code=?", (session_id, code)).fetchone()
    if old is None:
        raise ValueError("Producto ajeno al conteo.")
    con.execute("UPDATE v71_wh_items SET counted=?,updated_at=? WHERE session_id=? AND code=?", (quantity, now, session_id, code))
    con.execute("INSERT INTO v71_wh_events(session_id,code,previous,quantity,recorded_at,source) VALUES (?,?,?,?,?,?)", (session_id, code, old[0], quantity, now, source))


def wh_save_count_v71(session_id, token, quantity):
    quantity = wh_quantity_v71(quantity)
    with db() as con:
        con.execute("BEGIN IMMEDIATE")
        _wh_session_v71(con, session_id, editable=True)
        code = _wh_resolve_v71(con, session_id, token)
        _wh_record_count_v71(con, session_id, code, quantity, "Captura manual / lector")
    return code


def wh_close_count_v71(session_id, closed=True):
    with db() as con:
        _wh_session_v71(con, session_id)
        con.execute("UPDATE v71_wh_sessions SET closed_at=? WHERE id=?", (datetime.now().isoformat(timespec="seconds") if closed else None, session_id))


def wh_count_report_v71(session_id):
    with db() as con:
        site, cutoff, closed_at = _wh_session_v71(con, session_id)
        data = pd.read_sql_query("SELECT code AS 'Código', description AS 'Descripción', expected AS 'Existencia al congelar', counted AS 'Conteo físico', updated_at AS 'Última captura' FROM v71_wh_items WHERE session_id=? ORDER BY code LIMIT ?", con, params=(session_id, WH_MAX_ROWS_V71 + 1))
    if len(data) > WH_MAX_ROWS_V71:
        raise ValueError("El conteo excede el límite de exportación.")
    data.insert(0, "Sede", site)
    data.insert(1, "Corte del inventario", cutoff)
    data["Diferencia"] = data["Conteo físico"] - data["Existencia al congelar"]
    data["Estado"] = np.select([data["Conteo físico"].isna(), data["Diferencia"].abs().le(1e-9)], ["Sin contar", "Coincide"], default="Diferencia por revisar")
    data["Conteo cerrado"] = "Sí" if closed_at else "No"
    return data


def wh_csv_v71(raw):
    import csv
    if len(raw) > 8 * 1024 * 1024:
        raise ValueError("El CSV supera 8 MB.")
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = raw.decode("cp1252")
    try:
        dialect = csv.Sniffer().sniff(text[:8192], delimiters=",;\t")
    except csv.Error:
        dialect = csv.excel
    reader = csv.reader(io.StringIO(text), dialect)
    headers = next(reader, [])
    norm = [_wh_norm_v71(h) for h in headers]
    if not headers or any(not h for h in norm) or len(set(norm)) != len(norm):
        raise ValueError("El CSV necesita una fila de encabezados únicos.")
    rows = []
    for i, row in enumerate(reader, 2):
        if not any(cell.strip() for cell in row):
            continue
        if len(row) != len(headers):
            raise ValueError(f"Fila {i}: cantidad de columnas distinta a la cabecera.")
        rows.append(row)
        if len(rows) > WH_MAX_ROWS_V71:
            raise ValueError(f"Máximo {WH_MAX_ROWS_V71:,} filas por CSV.")
    return pd.DataFrame(rows, columns=headers, dtype=str)


def _wh_column_v71(frame, names):
    matches = [c for c in frame.columns if _wh_norm_v71(c) in names]
    if len(matches) != 1:
        raise ValueError("Elige una sola columna de " + ", ".join(names) + ".")
    return matches[0]


def wh_save_barcodes_v71(session_id, frame):
    code_col = _wh_column_v71(frame, {"codigo", "sku"})
    barcode_col = _wh_column_v71(frame, {"codigo de barras", "codigo de barra", "barcode"})
    if not 0 < len(frame) <= WH_MAX_ROWS_V71:
        raise ValueError("Cantidad de correspondencias inválida.")
    records = [(str(row[barcode_col]).strip(), str(row[code_col]).strip()) for _, row in frame.iterrows()]
    if any(not b or not c for b, c in records) or len({b for b, _ in records}) != len(records):
        raise ValueError("Hay códigos de barras vacíos o repetidos. Corrige el CSV.")
    with db() as con:
        con.execute("BEGIN IMMEDIATE")
        _wh_session_v71(con, session_id, editable=True)
        codes = {r[0] for r in con.execute("SELECT code FROM v71_wh_items WHERE session_id=?", (session_id,))}
        existing = dict(con.execute("SELECT barcode,code FROM v71_wh_barcodes WHERE session_id=?", (session_id,)))
        for barcode, code in records:
            if code not in codes:
                raise ValueError(f"Código fuera del conteo: {code[:90]}")
            if (barcode in codes and barcode != code) or (barcode in existing and existing[barcode] != code):
                raise ValueError(f"Código de barras ambiguo o ya asignado: {barcode[:90]}")
        con.executemany("INSERT OR IGNORE INTO v71_wh_barcodes VALUES (?,?,?)", [(session_id, barcode, code) for barcode, code in records])
    return len(records)


def wh_validate_bulk_v71(session_id, frame):
    code_col = _wh_column_v71(frame, {"codigo", "sku", "codigo de barras", "barcode"})
    quantity_col = _wh_column_v71(frame, {"conteo fisico", "cantidad contada", "cantidad", "conteo"})
    if not 0 < len(frame) <= WH_MAX_ROWS_V71:
        raise ValueError(f"El CSV necesita entre 1 y {WH_MAX_ROWS_V71:,} filas.")
    records, issues, seen = [], [], set()
    with db() as con:
        _wh_session_v71(con, session_id)
        for i, (_, row) in enumerate(frame.iterrows(), 2):
            token = _wh_text_v71(row[code_col])
            try:
                code = _wh_resolve_v71(con, session_id, token)
                if code in seen:
                    raise ValueError("Producto repetido en el CSV, incluso si usa otro código de barras.")
                seen.add(code)
                quantity = wh_quantity_v71(row[quantity_col])
                records.append({"Código": code, "Conteo físico": quantity})
            except (ValueError, TypeError) as exc:
                issues.append({"Fila CSV": i, "Código recibido": token, "Problema": str(exc)})
    return pd.DataFrame(records, columns=["Código", "Conteo físico"]), pd.DataFrame(issues, columns=["Fila CSV", "Código recibido", "Problema"])


def wh_apply_bulk_v71(session_id, frame):
    valid, issues = wh_validate_bulk_v71(session_id, frame)
    if not issues.empty:
        raise ValueError("Corrige todas las filas observadas. No se guardó ninguna cantidad del CSV.")
    with db() as con:
        con.execute("BEGIN IMMEDIATE")
        _wh_session_v71(con, session_id, editable=True)
        for row in valid.to_dict("records"):
            _wh_record_count_v71(con, session_id, row["Código"], float(row["Conteo físico"]), "CSV de conteo")
    return len(valid)


def wh_catalog_review_v71(frame):
    data = frame.copy()
    for column in ["Sede", "Código", "Descripción"]:
        if column not in data:
            data[column] = ""
    site = data["Sede"].map(_wh_text_v71)
    code = data["Código"].map(_wh_text_v71)
    desc = data["Descripción"].map(_wh_norm_v71)
    same_code = pd.DataFrame({"site": site, "code": code}).duplicated(keep=False) & code.ne("")
    candidate = pd.DataFrame({"site": site, "description": desc, "code": code}).groupby(["site", "description"], dropna=False)["code"].transform("nunique").gt(1) & desc.ne("")
    findings = []
    for mask, reason in [(code.eq(""), "Código faltante"), (same_code, "Código repetido en la misma sede"), (desc.eq(""), "Descripción faltante"), (candidate, "Descripción similar con códigos distintos: revisar, no fusionar automáticamente")]:
        if mask.any():
            part = data.loc[mask, ["Sede", "Código", "Descripción"]].copy()
            part["Revisión"] = reason
            findings.append(part)
    photo_columns = [c for c in data.columns if _wh_norm_v71(c) in {"foto", "imagen", "url foto", "url imagen", "foto url", "imagen url"}]
    if photo_columns:
        informed = pd.Series(False, index=data.index)
        for c in photo_columns:
            informed |= data[c].map(_wh_text_v71).ne("")
        if (~informed).any():
            part = data.loc[~informed, ["Sede", "Código", "Descripción"]].copy()
            part["Revisión"] = "Campo de foto vacío en la tabla cargada; comprobar en el catálogo"
            findings.append(part)
    return (pd.concat(findings, ignore_index=True) if findings else pd.DataFrame(columns=["Sede", "Código", "Descripción", "Revisión"])), bool(photo_columns)


def wh_customer_catalog_v71(frame, include_category=False, include_brand=False):
    columns = [c for c in ["Código", "Descripción", "Existencia", "Sede"] if c in frame]
    if include_category and "Categoría" in frame:
        columns.append("Categoría")
    if include_brand and "Marca" in frame:
        columns.append("Marca")
    # Lista cerrada: el catálogo de clientes nunca incluye costo, margen ni cartera.
    return frame.loc[:, columns].copy()


def _wh_download_csv_v71(label, frame, filename, key):
    safe = frame.copy()
    for column in safe.columns:
        safe[column] = safe[column].map(lambda value: "'" + value if isinstance(value, str) and value.lstrip().startswith(("=", "+", "-", "@", "\t", "\r")) else value)
    st.download_button(label, safe.to_csv(index=False).encode("utf-8-sig"), filename, "text/csv", key=key)


def _render_count_v71(all_df):
    st.caption("Congela una sede, registra cantidades reales y revisa diferencias. Cada recuento reemplaza la cantidad anterior; nunca modifica existencias en A2.")
    sessions = wh_count_sessions_v71()
    with st.expander("Iniciar un conteo", expanded=sessions.empty):
        if all_df.empty or "Sede" not in all_df:
            st.info("Carga un inventario para iniciar un conteo.")
        else:
            sites = sorted(all_df["Sede"].dropna().astype(str).unique())
            site = st.selectbox("Sede del conteo", sites, key="wh_new_site")
            known_cutoff = globals().get("dates_used", {}).get(site) or latest_snapshot_date(site)
            st.write(f"Corte que se congelará: {known_cutoff or 'fecha no disponible'} · {int(all_df['Sede'].astype(str).eq(site).sum()):,} productos")
            title = st.text_input("Nombre del conteo", value="", placeholder="Ej.: Pasillo filtros · turno mañana", key="wh_new_title")
            if st.button("Congelar existencias e iniciar", disabled=not bool(known_cutoff), key="wh_new"):
                try:
                    new_id = wh_begin_count_v71(all_df, site, str(known_cutoff), title)
                    st.session_state["wh_selected_session"] = new_id
                    st.rerun()
                except (ValueError, sqlite3.Error) as exc:
                    st.error(str(exc))
    if sessions.empty:
        return
    labels = {r["id"]: f"{r['title']} · {r['created_at']} · {'Cerrado' if r['closed_at'] else 'Abierto'}" for r in sessions.to_dict("records")}
    chosen = st.session_state.get("wh_selected_session")
    if chosen not in labels:
        st.session_state["wh_selected_session"] = sessions.iloc[0]["id"]
    session_id = st.selectbox("Continuar un conteo guardado", list(labels), format_func=labels.get, key="wh_selected_session")
    report = wh_count_report_v71(session_id)
    done = int(report["Conteo físico"].notna().sum())
    left, right = st.columns(2)
    left.metric("Productos contados", f"{done:,} / {len(report):,}")
    right.metric("Cobertura del conteo", f"{100 * done / len(report):.1f}%" if len(report) else "0%")
    st.caption("Sin contar es distinto de cero. Las diferencias solo se calculan para cantidades registradas.")
    closed = bool(report["Conteo cerrado"].eq("Sí").any())
    if not closed:
        with st.form("wh_capture_form", clear_on_submit=True):
            token = st.text_input("Código del producto o lectura del código de barras", help="Un lector USB actúa como teclado. Para códigos de barras distintos al SKU, carga primero su correspondencia.")
            quantity = st.text_input("Cantidad física total", placeholder="Ej.: 12 o 12,5")
            if st.form_submit_button("Guardar cantidad total"):
                try:
                    code = wh_save_count_v71(session_id, token, quantity)
                    st.session_state["wh_last_saved"] = f"Cantidad guardada para {code}."
                    st.rerun()
                except (ValueError, sqlite3.Error) as exc:
                    st.error(str(exc))
        if st.session_state.get("wh_last_saved"):
            st.success(st.session_state.pop("wh_last_saved"))
        with st.expander("Cargar códigos de barras o un conteo por CSV"):
            st.caption("CSV UTF-8, hasta 20.000 filas. Conserva códigos como texto. Cantidades sin separadores de miles. Corrige códigos desconocidos o repetidos antes de guardar.")
            st.download_button("Plantilla de códigos de barras", "Código;Código de barras\n".encode("utf-8-sig"), "plantilla_codigos_barras.csv", "text/csv", key="wh_map_template")
            barcode_upload = st.file_uploader("Correspondencia: Código y Código de barras", type=["csv"], key="wh_barcodes_file")
            if barcode_upload is not None and st.button("Guardar correspondencia", key="wh_save_barcodes"):
                try:
                    count = wh_save_barcodes_v71(session_id, wh_csv_v71(barcode_upload.getvalue()))
                    st.success(f"{count:,} correspondencias comprobadas y guardadas.")
                except (ValueError, UnicodeError, sqlite3.Error) as exc:
                    st.error(str(exc))
            template = report[["Código", "Descripción"]].copy()
            template["Conteo físico"] = ""
            _wh_download_csv_v71("Descargar plantilla de este conteo", template, "plantilla_conteo.csv", "wh_count_template")
            count_upload = st.file_uploader("Conteo: Código y Conteo físico", type=["csv"], key="wh_counts_file")
            if count_upload is not None:
                try:
                    imported = wh_csv_v71(count_upload.getvalue())
                    valid, issues = wh_validate_bulk_v71(session_id, imported)
                    if not issues.empty:
                        st.warning("No se guardará este archivo hasta corregir todas las observaciones.")
                        st.dataframe(issues.head(100), hide_index=True, width="stretch")
                    else:
                        st.caption(f"{len(valid):,} productos válidos. Se reemplazarán sus cantidades anteriores.")
                        if st.button("Guardar cantidades del CSV", key="wh_apply_counts"):
                            wh_apply_bulk_v71(session_id, imported)
                            st.rerun()
                except (ValueError, UnicodeError, sqlite3.Error) as exc:
                    st.error(str(exc))
    st.dataframe(report.head(500), hide_index=True, width="stretch")
    if len(report) > 500:
        st.caption("Vista previa: 500 filas. El CSV contiene el conteo completo.")
    _wh_download_csv_v71("Descargar comparación completa", report, f"conteo_{session_id[:8]}.csv", "wh_count_export")
    if st.button("Reabrir conteo" if closed else "Cerrar conteo y conservar revisión", key="wh_toggle_close"):
        wh_close_count_v71(session_id, closed=not closed)
        st.rerun()


def render_warehouse_tools_v71(all_df):
    with st.expander("Herramientas de almacén", expanded=False):
        tool = st.selectbox("Qué necesitas hacer", ["Conteo físico", "Revisar catálogo", "Catálogo con existencias"], key="wh_tool")
        if tool == "Conteo físico":
            _render_count_v71(all_df)
        elif tool == "Revisar catálogo":
            st.caption("Identifica códigos repetidos y descripciones que merecen revisión. No fusiona productos ni modifica el sistema administrativo.")
            data = all_df
            uploaded = st.file_uploader("Opcional: catálogo CSV con Código, Descripción, Sede y Foto / URL foto", type=["csv"], key="wh_audit_catalog")
            if uploaded is not None:
                try:
                    data = wh_csv_v71(uploaded.getvalue())
                    rename = {}
                    for col in data.columns:
                        norm = _wh_norm_v71(col)
                        if norm in {"codigo", "sku", "descripcion", "sede"}:
                            rename[col] = {"codigo": "Código", "sku": "Código", "descripcion": "Descripción", "sede": "Sede"}[norm]
                    if len(set(rename.values())) != len(rename):
                        raise ValueError("El catálogo tiene columnas equivalentes repetidas.")
                    data = data.rename(columns=rename)
                    if not {"Código", "Descripción"}.issubset(data.columns):
                        raise ValueError("El catálogo debe tener Código y Descripción.")
                    if "Sede" not in data:
                        data["Sede"] = "Catálogo sin sede"
                except (ValueError, UnicodeError) as exc:
                    st.error(str(exc))
                    return
            findings, has_photo = wh_catalog_review_v71(data)
            if not has_photo:
                st.info("Sin un campo de foto en la fuente no se puede determinar qué productos carecen de imagen. Puedes adjuntar un catálogo que incluya ese campo.")
            if findings.empty:
                st.success("No se encontraron observaciones con los campos disponibles.")
            else:
                st.write(f"{len(findings):,} observaciones para revisar.")
                st.dataframe(findings.head(500), hide_index=True, width="stretch")
                _wh_download_csv_v71("Descargar observaciones", findings, "revision_catalogo.csv", "wh_audit_export")
        else:
            st.caption("Genera una lista de productos y existencias para compartir. La tarifa P12 requiere una lista de precios confirmada; aquí no se calculan precios.")
            if all_df.empty:
                st.info("Carga un inventario para preparar el catálogo.")
                return
            sites = sorted(all_df["Sede"].dropna().astype(str).unique()) if "Sede" in all_df else []
            selected = site_selector_v712("Sedes del catálogo", "wh_catalog_sites", options=sites)
            data = all_df[all_df["Sede"].astype(str).isin(selected)] if "Sede" in all_df else all_df
            if st.checkbox("Solo productos con existencia positiva", value=True, key="wh_catalog_positive"):
                data = data[pd.to_numeric(data["Existencia"], errors="coerce").gt(0)]
            category = st.checkbox("Incluir categoría", key="wh_catalog_category")
            brand = st.checkbox("Incluir marca", key="wh_catalog_brand")
            catalog = wh_customer_catalog_v71(data, category, brand)
            st.dataframe(catalog.head(300), hide_index=True, width="stretch")
            st.caption(f"{len(catalog):,} productos/sede en la exportación. Existencias del último corte cargado; no es una reserva de mercancía.")
            _wh_download_csv_v71("Descargar catálogo con existencias", catalog, f"catalogo_existencias_{date.today().isoformat()}.csv", "wh_catalog_export")


# Guía del puesto y formatos admitidos, integrada en Reportes.
ROLE_GUIDE_V71 = [
    ('Mínimos y máximos', 'Disponible', 'Mínimos y Máximos: políticas manuales o demanda del período. Inventario y ventas; sin CxC.'),
    ('Sobrestock', 'Disponible', 'Inventario y Compras: excedentes, retiros y posibles redistribuciones.'),
    ('Productos sin rotación', 'Disponible', 'Más herramientas → Sin Rotación. Identifica ausencia de ventas en el historial cargado; no supone inactividad fuera de ese período.'),
    ('Productos sin fotos', 'Revisión con datos adicionales', 'Necesita una columna de foto/URL o catálogo exportado. No puede deducir qué imágenes tiene A2 desde inventario y ventas.'),
    ('RX y desempeño de asesores', 'Requiere definir RX y aportar reporte', 'Estos archivos no identifican asesores ni el significado de RX; no se inventan métricas individuales.'),
    ('Reposición de mercancía', 'Disponible', 'Compras: propuestas por sede, proveedor y departamento. Las transferencias sugeridas reducen la compra neta.'),
    ('Clientes y productos duplicados', 'Productos: revisión; clientes: otro insumo', 'Inventario → revisión de catálogo detecta candidatos. Para clientes hace falta el maestro de clientes; ninguna fila se fusiona automáticamente.'),
    ('Gráficas por departamento', 'Disponible', 'Reportes y Gestión Visual, con departamentos y proveedores confirmados.'),
    ('Mejoras e implementación de IA', 'Apoyo al análisis', 'Lectura de datos, validaciones y reglas explicables. No requiere servicios de IA de pago.'),
    ('CRM: venta, cobro y postventa', 'Integración aparte', 'Automatización conserva tareas y webhooks manuales existentes. Este cambio no envía mensajes ni conecta A2 o un CRM por sí solo.'),
    ('Pistoleo / código de barras', 'Conteo físico', 'Escáner en modo teclado: SKU directo o equivalencia de código de barras confirmada. Guarda conteos y diferencias.'),
    ('Inventario vs ventas vs CxC', 'Disponible por componentes', 'Inventario + ventas generan el informe de almacén. CxC amplía análisis de saldos y vencimientos; ausencia de archivo no significa cartera cero.'),
    ('Cargas masivas de cantidades', 'Preparación y revisión', 'Carga multisede de reportes y CSV de conteos. Exporta propuestas; no escribe movimientos ni cantidades en A2.'),
    ('Meses anteriores vs mes actual', 'Disponible con períodos comparables', 'Carga informes separados con fechas reales. Un acumulado de varios meses no permite reconstruir ventas de cada mes.'),
    ('Revisión física del inventario', 'Conteo físico', 'Compara el conteo contra una fotografía de existencias guardada; distingue pendiente de contar de cero contado.'),
    ('Catálogo semanal P12 con existencias', 'Catálogo con existencias; P12 por definir', 'Exporta código, descripción, sede y disponibilidad sin costos internos. Para precios P12 hace falta confirmar qué lista o campo representa P12.'),
]


def render_role_guide_v71():
    st.caption('Alcance basado en tu descripción de cargo. Los cálculos y documentos son propuestas para revisión; las acciones físicas y los cambios en A2 requieren ejecución en sus procesos correspondientes.')
    st.dataframe(pd.DataFrame(ROLE_GUIDE_V71,columns=['Función del puesto','Estado','Dónde / datos necesarios']),hide_index=True,width='stretch')
    st.markdown('**Si prefieres una tabla sencilla antes de cargar:**')
    st.code('Inventario: Código | Descripción | Existencia | Costo (opcional)\nVentas: Código | Descripción | Cantidad | Monto bruto (opcional)\nContexto por archivo: sede, moneda, fecha de inventario y período real de ventas.\nCxC, opcional: Factura | Cliente | Emisión | Vencimiento | Saldo',language=None)
    st.caption('Conserva los códigos como texto y sus ceros iniciales. Evita mezclar sedes o períodos sin identificarlos. Si falta costo, se conserva el cálculo en unidades y queda pendiente la valoración que lo necesita. Las columnas se pueden reasignar en la revisión de carga.')


# """Fechas de consulta seguras sin cambiar los reportes guardados."""


def date_value_v72(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if value is None:
        return None
    try:
        return date.fromisoformat(str(value).strip())
    except (ValueError, TypeError):
        return None


def inspect_finance_dates_v72(stored, today=None):
    today = today or date.today()
    keep, problems = [], []
    for index, row in stored.iterrows():
        first, last = date_value_v72(row.get("start_date")), date_value_v72(row.get("end_date"))
        reason = ""
        if first is None or last is None:
            reason = "Fecha ausente o no válida"
        elif first > last:
            reason = "La fecha inicial es posterior a la final"
        elif first > today or last > today:
            reason = "El período contiene una fecha futura"
        elif first < date(1900, 1, 1):
            reason = "La fecha es anterior a 1900"
        elif (last - first).days > 1095:
            reason = "El período supera tres años; revisa sus fechas"
        if reason:
            problems.append({"Sede": row.get("site", ""), "Desde": str(row.get("start_date", "")),
                             "Hasta": str(row.get("end_date", "")), "Motivo": reason})
        else:
            keep.append(index)
    return stored.loc[keep].copy(), pd.DataFrame(problems, columns=["Sede", "Desde", "Hasta", "Motivo"])


def safe_finance_period_v72(value, fallback, today=None):
    today = today or date.today()
    if not isinstance(value, (tuple, list)) or len(value) > 2:
        return fallback, True
    dates = tuple(date_value_v72(v) for v in value)
    if any(v is None or v < date(1900, 1, 1) or v > today for v in dates):
        return fallback, True
    if len(dates) == 2 and (dates[1] < dates[0] or (dates[1] - dates[0]).days > 1095):
        return fallback, True
    return dates, False


def financial_filters_v7(prefix: str) -> tuple | None:
    """Valida fuentes y estado antes de construir los calendarios de Streamlit."""
    today = date.today()
    with db() as con:
        stored = pd.read_sql_query("""SELECT start_date,end_date,currency,site FROM v7_sales_batches
            WHERE scope='' UNION ALL SELECT day AS start_date,day AS end_date,currency,site
            FROM financial_sales""", con)
        ar_sources = pd.read_sql_query("SELECT as_of,currency,site FROM financial_ar_batches", con)
        inv_sources = pd.read_sql_query("""SELECT DISTINCT i.site,f.inventory_currency AS currency
            FROM inventory_snapshots i LEFT JOIN financial_settings f ON f.site=i.site""", con)
    cols = st.columns([2, 1, 2])
    observed = set(stored["currency"].dropna()) | set(ar_sources["currency"].dropna()) | set(inv_sources["currency"].dropna())
    currencies = [c for c in FIN_CURRENCIES if c in observed] + [c for c in FIN_CURRENCIES if c not in observed]
    currency = cols[1].selectbox("Moneda", currencies, key=f"{prefix}_currency")
    available = stored[stored["currency"] == currency]
    ar_available = ar_sources[ar_sources["currency"] == currency]
    active = set(available["site"]) | set(ar_available["site"]) | set(inv_sources.loc[inv_sources["currency"] == currency, "site"])
    site_options = list(dict.fromkeys(list(SEDES) + sorted(active)))
    sites_key, sites_currency_key = f"{prefix}_sites_v7", f"{prefix}_sites_currency_v7"
    if st.session_state.get(sites_currency_key) != currency:
        st.session_state[sites_key] = [site for site in site_options if site in active]
        st.session_state[sites_currency_key] = currency
    sites = site_selector_v712("Sedes", sites_key, options=site_options, container=cols[2])
    selected, issues = inspect_finance_dates_v72(available[available["site"].isin(sites)], today)
    if not issues.empty:
        st.warning(f"Hay {len(issues)} períodos guardados con fechas por revisar. No se usan como fechas iniciales de esta consulta y tus archivos guardados no se modifican.")
        with st.expander("Ver fechas guardadas que requieren revisión", expanded=True):
            st.dataframe(issues, hide_index=True, width="stretch")
            st.caption("Confirma las fechas reales en A2 y vuelve a cargar el reporte. Un período futuro no acredita ventas históricas de esta consulta; revisa también la cobertura y los períodos excluidos del informe.")
            if st.button("Ir a Carga de datos", key=f"{prefix}_fix_source_dates"):
                _navigation_v7("Carga de datos")
    a2_ranges = selected[selected["start_date"] != selected["end_date"]]
    if not a2_ranges.empty:
        newest = a2_ranges.sort_values(["end_date", "start_date"], ascending=[False, False]).iloc[0]
        default_start, default_end = date_value_v72(newest["start_date"]), date_value_v72(newest["end_date"])
    elif not selected.empty:
        default_end = date_value_v72(selected["end_date"].max())
        default_start = max(default_end.replace(day=1), date_value_v72(selected["start_date"].min()))
    else:
        default_end, default_start = today, today.replace(day=1)
    fallback = (default_start, default_end)
    filter_token = (currency, tuple(sites), default_start.isoformat(), default_end.isoformat())
    token_key, period_key = f"{prefix}_defaults_token_v7", f"{prefix}_period_v7"
    changed_filter = st.session_state.get(token_key) != filter_token
    if changed_filter:
        st.session_state[period_key] = fallback
        st.session_state[token_key] = filter_token
    safe_period, repaired = safe_finance_period_v72(st.session_state.get(period_key), fallback, today)
    if repaired:
        st.session_state[period_key] = safe_period
        st.info("El calendario tenía fechas fuera del rango admitido. Se restableció el filtro de consulta; los reportes originales conservan sus fechas.")
    period = cols[0].date_input("Período de ventas", min_value=date(1900, 1, 1), max_value=today, key=period_key)
    if not isinstance(period, (tuple, list)) or len(period) != 2 or not sites:
        st.info("Selecciona ambas fechas y al menos una sede. Las sedes sin datos se mostrarán como no informadas.")
        return None
    first, last = period
    if last < first or last > today or (last - first).days > 1095:
        st.warning("Selecciona un período válido de hasta tres años y sin fechas futuras.")
        return None
    with db() as con:
        marks = ",".join("?" for _ in sites)
        inv_cut = con.execute(f"SELECT MAX(snapshot_date) FROM inventory_snapshots WHERE site IN ({marks}) AND snapshot_date<=?", (*sites, today.isoformat())).fetchone()[0]
    candidates = [date_value_v72(inv_cut)] + [date_value_v72(c) for c in ar_available.loc[ar_available["site"].isin(sites), "as_of"]]
    saved_cuts = [cut for cut in candidates if cut is not None and date(1900, 1, 1) <= cut <= today]
    default_cut = max([last] + saved_cuts)
    cut_key = f"{prefix}_analysis_date_v7"
    previous_cut = date_value_v72(st.session_state.get(cut_key))
    if changed_filter or previous_cut is None or previous_cut < last or previous_cut > today:
        st.session_state[cut_key] = default_cut
    analysis = st.date_input("Fecha de análisis de inventario y CxC", min_value=last, max_value=today, key=cut_key)
    st.caption("Desempeño integral compara inventario, ventas y, cuando está disponible, CxC. El período filtra la consulta; no cambia las fechas de los reportes guardados.")
    return first.isoformat(), last.isoformat(), currency, tuple(sites), analysis.isoformat()


# Ayuda integrada: cada afirmación corresponde a una pantalla o salida existente.
# No ejecuta importaciones de datos, envíos, movimientos A2 ni cambios en la base.
HELP_MODULES_V72 = [
    ("Centro de Control", "Principal", "Revisar el estado general y decidir qué atender primero.", "Inventarios guardados; ventas para demanda, rotación y reposición.", "Indicadores, prioridades y accesos a las otras pantallas."),
    ("Carga de datos", "Principal", "Importar archivos de una o varias sedes y revisar su lectura antes de guardar.", "XLS, XLSX o CSV; sede, moneda, corte y período real. CxC es opcional.", "Datos guardados para los cálculos; historial de cargas y clasificación de productos."),
    ("Inventario", "Principal", "Consultar productos, existencias y niveles; hacer conteos y revisar el catálogo.", "Inventario por sede. Ventas y políticas de stock para interpretar los niveles.", "Reporte de inventario; CSV de conteos, revisión de catálogo o catálogo con existencias."),
    ("Compras", "Principal", "Revisar la reposición neta después de las transferencias sugeridas.", "Inventario, ventas con su período o límites manuales; costos para presupuesto.", "Productos y unidades a comprar, proveedor y presupuesto cuando existe costo."),
    ("Mínimos y Máximos", "Principal", "Definir mínimos y máximos manuales o revisar las propuestas por demanda.", "Productos y sede; ventas para propuesta automática; cobertura, plazo y seguridad.", "Políticas guardadas que actualizan las recomendaciones del motor."),
    ("Reportes", "Principal", "Elegir la gestión, filtrar y descargar el informe final o un documento.", "Inventario y ventas para almacén. CxC y otros campos sólo para sus indicadores.", "Informe de almacén, análisis financiero u órdenes y movimientos propuestos."),
    ("Desempeño Integral", "Más herramientas", "Comparar inventario, ventas y, cuando existe, cartera por sede y moneda.", "Ventas con fechas e importes; inventario valorado. CxC, ventas a crédito y costo real de ventas activan sus indicadores.", "Indicadores, fuentes y cortes; informe ZIP con Excel, HTML y CSV."),
    ("Cuentas por Cobrar", "Más herramientas", "Revisar saldos pendientes, vencimientos y clientes que requieren seguimiento.", "Cartera completa: factura, cliente, emisión, vencimiento y saldo por sede y moneda.", "Cartera filtrada descargable. Días de CxC requiere además ventas a crédito reales."),
    ("Gestión Visual", "Más herramientas", "Organizar tareas, responsables, fechas y estados de resolución.", "Una tarea escrita o problemas detectados en inventario o desempeño.", "Tablero local de seguimiento; guardar una tarea no ejecuta el trabajo en A2."),
    ("Proveedores", "Más herramientas", "Ver stock, compras y productos sin rotación de un proveedor.", "Relaciones de productos confirmadas o cortes dedicados al proveedor.", "Análisis por proveedor y sus reportes; propuesta de orden de compra."),
    ("Sin Rotación", "Más herramientas", "Encontrar productos con existencias y sin ventas detectadas en el historial.", "Inventario positivo y ventas de un período conocido.", "Tabla y Excel de productos sin rotación; el historial limita lo que se puede concluir."),
    ("Redistribución", "Más herramientas", "Ver excedentes de una sede que pueden cubrir necesidades de otra.", "Inventario, demanda o límites de al menos dos sedes; códigos compatibles.", "Plan de transferencias y Excel. Una propuesta no confirma el traslado físico."),
    ("Centro Visual", "Más herramientas", "Explorar gráficas del inventario y los cálculos logísticos.", "Datos guardados y clasificación disponible por sede, proveedor o departamento.", "Gráficas y comparaciones de los datos del motor."),
    ("Finanzas", "Más herramientas", "Revisar valor de inventario, sobrestock y presupuesto estimado de compra.", "Existencias y costos o valores informados. Los importes incompletos quedan N/D.", "Vista financiera del almacén; no equivale a contabilidad, caja o utilidad neta."),
    ("Alertas", "Más herramientas", "Priorizar problemas de inventario por tipo y severidad.", "Inventario y resultados del motor; ventas y políticas donde correspondan.", "Alertas para revisar; también se incluyen en Reportes → Alertas."),
    ("Historial", "Más herramientas", "Consultar qué sede y corte se procesaron y cuántas filas se guardaron.", "Al menos una carga guardada.", "Registro de procesamiento. Las cargas de ventas también se revisan en Carga de datos."),
    ("Automatización", "Más herramientas", "Preparar un resumen logístico para una integración ya configurada.", "Datos guardados, sede y webhook de destino configurado.", "Vista previa del resumen y envío al pulsar el botón; no incluye programación semanal ni campañas CRM completas."),
]

HELP_REPORTS_V72 = [
    ("Inventario", "Existencia, mínimo, máximo, punto de reorden, cobertura, estado, fecha y valor disponible.", "Inventario. Ventas o política manual para niveles; costo/valor para dinero."),
    ("Ventas", "Unidades del período y sus importes informados; incluye productos fuera del inventario y filas pendientes sin código.", "Reporte de ventas. No reconstruye meses desde un acumulado ni deduce cobros."),
    ("Compras", "Déficit, transferencias recibidas, compra neta, costo y presupuesto estimado.", "Inventario y demanda/política de stock. Sin costo conserva unidades y deja importe N/D."),
    ("Retiros", "Exceso, transferencias de salida, retiro neto y valor disponible.", "Inventario y política de stock. Retiro neto = exceso menos salidas propuestas."),
    ("Redistribución", "Producto, sede de origen, destino, unidades sugeridas y compra evitada estimada.", "Sedes con exceso y déficit compatibles. Puede quedar vacío si no hay propuestas."),
    ("Alertas", "Sede, producto, tipo de problema, severidad y detalle.", "Problemas detectados en el corte; un producto puede tener varias alertas."),
]

HELP_FORMATS_V72 = [
    ("Excel operativo", "Trabajar las tablas y sus filas en Excel.", "Detalle de todas las filas del filtro."),
    ("Excel gerencial", "Revisar el resumen y presentar los resultados.", "Resumen y detalle de la selección."),
    ("PDF ejecutivo", "Entregar una lectura breve para dirección.", "Totales completos y hasta 20 registros por gestión; requiere fpdf2."),
    ("PDF detallado", "Imprimir todas las filas seleccionadas.", "Se prepara aparte en su desplegable; puede tener muchas páginas; requiere fpdf2."),
    ("Reporte visual HTML", "Abrir un informe visual en el navegador.", "Todas las filas del filtro."),
    ("CSV por gestión (ZIP)", "Usar las tablas en Excel u otro proceso.", "Un CSV por gestión seleccionada con sus filas completas."),
    ("JSON por gestión", "Entregar datos estructurados a otra herramienta.", "Datos completos de la selección."),
    ("Paquete completo", "Guardar juntas las salidas preparadas.", "ZIP con los formatos del paquete; el PDF detallado se prepara y descarga aparte."),
]

HELP_GLOSSARY_V72 = [
    ("Sede", "Lugar cuyo inventario o ventas estás cargando: por ejemplo, Centurión. Cada archivo debe corresponder a su sede real."),
    ("Departamento", "Clasificación de productos dentro de una sede. Para usarla hace falta que venga en el archivo o que confirmes su relación en Catálogos y relaciones."),
    ("Área / gestión del reporte", "La sección que quieres descargar: Inventario, Ventas, Compras, Retiros, Redistribución o Alertas. No es una sede ni un departamento de A2."),
    ("Corte", "Fecha de la fotografía de existencias o saldos. No necesariamente es el día en que subes el archivo."),
    ("Período de ventas", "Fecha inicial y final de las ventas incluidas. Un acumulado conserva todo su rango; no es un detalle mensual."),
    ("Desempeño integral", "Comparación del negocio por sede: inventario, ventas y cartera cuando está disponible. No exige que todos los indicadores tengan dato para mostrar los demás."),
    ("Unidades / importe", "Unidades son cantidades de productos. Importe es dinero en la moneda indicada. Tener 100 unidades no informa cuánto se vendió en dinero."),
    ("Costo / costo de ventas", "Costo unitario del inventario y costo real de lo vendido son datos diferentes. El margen necesita costo real de ventas."),
    ("CxC", "Cuentas por cobrar: lo que los clientes deben según facturas pendientes. No se obtiene restando números del reporte de productos vendidos."),
    ("N/D", "Dato no disponible o indicador que no se puede calcular con sus fuentes actuales. No significa cero."),
    ("Sin asignar", "Falta clasificar departamento, proveedor, marca o categoría. El producto puede existir y tener cantidades correctas."),
    ("Cobertura", "Tiempo estimado que alcanza el stock según la demanda utilizada. Revisa siempre las fechas y parámetros del cálculo."),
    ("Lead time / plazo del proveedor", "Días esperados desde el pedido hasta recibir mercancía; se configura en el motor."),
    ("Stock de seguridad", "Reserva adicional para cubrir variaciones o retrasos, según los días configurados."),
    ("ABC", "Clasificación por importancia relativa según la base elegida en el motor: movimiento, inventario o capital inmovilizado."),
    ("Compra neta / retiro neto", "Recomendación después de considerar transferencias propuestas. No acredita que ya se compró, retiró o trasladó mercancía."),
    ("Health Score / salud", "Puntuación interna de las condiciones del stock calculadas por el motor. Sirve para priorizar revisión; no es una auditoría de la empresa."),
    ("Rastreo de pila", "Detalle técnico de un error del programa. La última línea identifica el fallo; el resto indica en qué parte ocurrió."),
]


def render_help_v72():
    """Guía de consulta gradual; usa sólo Streamlit, pandas y nombres de la app."""
    st.header("Guía de uso")
    st.caption("Aprende el recorrido que vas a repetir: cargar → revisar → guardar → analizar → descargar. Las otras herramientas se consultan cuando las necesitas.")
    topic = st.radio(
        "¿Qué quieres aprender?",
        ["Empezar y cargar", "Sedes y clasificación", "Reportes y descargas", "Mapa de la aplicación", "Glosario", "Resolver avisos"],
        horizontal=True,
        key="v72_help_topic",
    )
    if topic == "Empezar y cargar":
        st.info("Para tu informe de almacén empieza con inventario y ventas. Añade CxC sólo cuando necesites analizar la cartera.")
        st.markdown("**Tu primera carga**")
        st.markdown("1. Abre **Carga de datos → Cargar reportes** y elige **Una sede**.\n2. Sube el inventario y las ventas de esa misma sede. Puedes cargar varios períodos de ventas.\n3. Revisa la sede, la moneda y el corte del inventario; comprueba el inicio y el fin de cada período de ventas.\n4. Mira la vista previa: códigos, descripciones y cantidades deben corresponder al archivo. Si se detectó mal una columna, corrígela en la revisión de lectura.\n5. Resuelve los avisos bloqueantes y confirma que revisaste los datos. Pulsa **Guardar lote y actualizar resultados**.\n6. En **Reportes**, elige **Informe final de almacén** y prepara los archivos de tu selección.")
        with st.expander("Cargar varias sedes en el mismo trabajo"):
            st.markdown("1. Elige **Varias sedes** y sube los archivos de las sucursales que vas a procesar.\n2. Asigna a cada archivo su **tipo y sede reales**, además de moneda y fechas. Un inventario de Centurión y sus ventas se asignan a Centurión; los archivos de otra sede se asignan a esa otra sede.\n3. Revisa el resumen por sede. Evita mezclar inventarios completos, reportes parciales y acumulados repetidos.\n4. Guarda el lote. Después selecciona todas las sedes que quieras comparar o incluir en el informe.")
            st.caption("Límite de carga: 30 archivos, 20 MB por archivo y 100 MB por lote. Para una sede con varios archivos parciales del mismo inventario, revisa el alcance antes de guardar; no sumes fotografías completas del mismo stock.")
        with st.expander("Rutina diaria o semanal, sin recorrer todos los menús"):
            st.markdown("**Al cargar:** comprueba sede, fechas, filas y totales.\n\n**Al revisar:** abre Inventario, Compras y Alertas; ajusta mínimos y máximos únicamente cuando cambie tu política.\n\n**Al entregar:** ve a Reportes, revisa el filtro y descarga el archivo que necesita cada persona.\n\n**Al terminar:** prepara un respaldo desde el menú lateral si necesitas conservar una copia de tus datos.")
        if st.button("Ir a Carga de datos", key="v72_help_go_load"):
            st.session_state["nexus_view"] = "Carga de datos"
            st.rerun()
    elif topic == "Sedes y clasificación":
        st.markdown("**Tres decisiones distintas**")
        st.dataframe(pd.DataFrame([
            ("Sede", "¿De qué sucursal son estos datos?", "Centurión, Barcelona u otra sede configurada."),
            ("Departamento", "¿En qué grupo de productos se clasifican?", "El departamento que figure en tu catálogo A2."),
            ("Gestión del reporte", "¿Qué trabajo quiero informar?", "Inventario, Ventas, Compras, Retiros, Redistribución o Alertas."),
        ], columns=["Campo", "Pregunta que responde", "Ejemplo"]), hide_index=True, width="stretch")
        st.markdown("**Seleccionar todas las sedes** sirve para consultar o preparar un lote con varias sedes. Cada archivo conserva su asignación real: seleccionar todas no copia el mismo inventario a todas las sucursales. Una sede sin archivos guardados no adquiere datos al seleccionarla.")
        with st.expander("El producto aparece, pero el departamento o proveedor dice «Sin asignar»", expanded=True):
            st.markdown("1. Abre **Carga de datos → Catálogos y relaciones**.\n2. Revisa o carga los catálogos de departamentos y proveedores.\n3. Confirma las relaciones que sí conoces entre departamento, proveedor, marca y texto del producto.\n4. Usa la asignación de un artículo específico para corregir una excepción por código.\n5. Guarda las relaciones y vuelve a Inventario o Reportes.")
            st.caption("Una lista de nombres de departamentos no indica por sí sola a cuál pertenece cada producto. Cuando varias reglas coinciden, revisa la clasificación; no se debe elegir una relación al azar.")
        st.markdown("**Inventario de un proveedor:** utiliza la opción avanzada de corte dedicado cuando el archivo sólo representa ese proveedor. Así identificas su alcance antes de tratarlo como inventario general.")
    elif topic == "Reportes y descargas":
        st.markdown("**Informe final de almacén**")
        st.markdown("En **Reportes → Informe final de almacén**, selecciona sedes y, si lo necesitas, proveedor, marca o departamento. Elige **Reporte general** para reunir las seis gestiones, o una sola gestión para una entrega específica.")
        st.dataframe(pd.DataFrame(HELP_REPORTS_V72, columns=["Gestión", "Qué contiene", "Qué necesita"]), hide_index=True, width="stretch")
        st.markdown("Pulsa **Preparar archivos de esta selección** y después el botón de descarga. Si cambias filtros, prepara de nuevo para que la descarga corresponda a lo que estás viendo.")
        with st.expander("Qué formato descargar", expanded=True):
            st.dataframe(pd.DataFrame(HELP_FORMATS_V72, columns=["Formato", "Para qué usarlo", "Alcance"]), hide_index=True, width="stretch")
        with st.expander("Revisar el informe antes de entregarlo"):
            st.markdown("1. Confirma sedes, fechas y monedas de los datos.\n2. Comprueba códigos y unidades contra la fuente; verifica los totales que estén disponibles.\n3. Revisa qué costos o clasificaciones figuran N/D o Sin asignar.\n4. Comprueba las filas pendientes sin código en Ventas: sus importes se conservan, pero su producto aún debe identificarse.\n5. Revisa Compras, Retiros y Redistribución como propuestas pendientes de aprobación y ejecución.")
            st.caption("Los importes se presentan por moneda. No se convierten monedas automáticamente. Los CSV de la selección son una exportación de NEXUS, no una plantilla de importación A2 validada.")
        with st.expander("Análisis financiero y documentos operativos"):
            st.markdown("**Análisis financiero · opcional** abre Desempeño Integral. El ZIP incluye Indicadores, Fuentes y cortes, Ventas por período, Ventas diarias reales y Cartera CxC; cuando corresponde, añade Períodos excluidos. Los apartados sin su fuente conservan datos pendientes.\n\n**Órdenes y movimientos** prepara pedido de compra, retiro de almacén, alertas o redistribución. Permite PDF, HTML imprimible, Excel o CSV. La vista previa puede limitar filas; el archivo conserva el detalle preparado.")
        if st.button("Ir a Reportes", key="v72_help_go_reports"):
            st.session_state["nexus_view"] = "Reportes"
            st.rerun()
    elif topic == "Mapa de la aplicación":
        st.caption("Los seis accesos principales cubren el trabajo habitual. En Más herramientas están las consultas especializadas.")
        selected = st.selectbox("Conocer un módulo", [row[0] for row in HELP_MODULES_V72], key="v72_help_module")
        row = next(row for row in HELP_MODULES_V72 if row[0] == selected)
        st.markdown(f"**{row[0]} · {row[1]}**\n\n{row[2]}\n\n**Datos necesarios:** {row[3]}\n\n**Resultado:** {row[4]}")
        if st.button(f"Abrir {selected}", key="v72_help_open_module"):
            st.session_state["nexus_view"] = selected
            st.rerun()
        with st.expander("Ver el mapa completo de los 17 módulos"):
            st.dataframe(pd.DataFrame(HELP_MODULES_V72, columns=["Módulo", "Acceso", "Para qué sirve", "Datos necesarios", "Resultado"]), hide_index=True, width="stretch")
        with st.expander("Tus tareas de almacén: lo automatizado y lo que necesita otro reporte"):
            render_role_guide_v71()
            st.caption("La revisión de fotos necesita un catálogo que informe foto o URL; RX y P12 necesitan su definición y fuente. Un conteo físico se guarda separado del inventario administrativo y no modifica A2.")
    elif topic == "Glosario":
        selected = st.selectbox("Término que quiero entender", [row[0] for row in HELP_GLOSSARY_V72], key="v72_help_term")
        st.info(next(row[1] for row in HELP_GLOSSARY_V72 if row[0] == selected))
        with st.expander("Ver todos los términos"):
            st.dataframe(pd.DataFrame(HELP_GLOSSARY_V72, columns=["Término", "Qué significa"]), hide_index=True, width="stretch")
        st.markdown("**Ejemplo:** cargas el inventario de Centurión al 6 de septiembre y ventas del 1 de julio al 5 de septiembre. Centurión es la sede; el 6 de septiembre es el corte; julio a septiembre es el período. Elegir Compras como gestión prepara la propuesta de reposición. No necesitas CxC para ese informe.")
        st.caption("Las fechas del ejemplo son ilustrativas. En tu carga utiliza las fechas reales de cada archivo.")
    else:
        with st.expander("Fecha futura o fuera del rango permitido", expanded=True):
            st.markdown("Revisa el **año, mes y día** detectados. Tu archivo WEBBvendidos declara como fin el **05/09/2027**; comprueba el rango real en A2 y corrígelo durante la carga si corresponde. La fecha no debe cambiarse sólo para pasar una validación: de ella depende el cálculo de demanda. Conserva el archivo original para comparar.")
        with st.expander("Los totales no coinciden o el inventario no se reconoce"):
            st.markdown("Revisa la hoja elegida, la fila del encabezado y las columnas asignadas. Comprueba que existencia no se haya confundido con costo o importe; verifica separadores decimales y filas de subtotales. Compara el total leído con el total impreso. Una diferencia importante debe resolverse antes de guardar; una tolerancia de redondeo no debe ocultar productos faltantes.")
            st.caption("Si persiste, conserva el XLS/XLSX original y comparte el mensaje de revisión con sus totales. Una captura no permite reconstruir todas las filas del archivo.")
        with st.expander("Veo N/D en Desempeño Integral"):
            st.markdown("Abre **Fuentes, fechas y datos que faltan**. Comprueba sede, moneda, período y fecha de análisis. Con ventas sólo en unidades no hay importe de venta; sin costo real de ventas no hay margen; sin cartera no hay CxC; sin ventas a crédito reales no hay días de CxC. Los demás resultados disponibles siguen siendo utilizables.")
        with st.expander("No veo una sede o me aparece Sin asignar"):
            st.markdown("Revisa qué sedes están configuradas y cuáles tienen datos guardados. Para departamentos o proveedores, ve a **Catálogos y relaciones** y confirma la clasificación. Una sede, un departamento y una gestión del informe son campos diferentes.")
        with st.expander("Pantalla en blanco, rastreo de pila o estoy abriendo otra versión"):
            st.markdown("Comprueba la versión y ruta de abajo. En VS Code, el archivo abierto y el que ejecutas en la terminal pueden ser distintos. Detén el servidor anterior con Ctrl+C, ejecuta el archivo completo actualizado desde su carpeta y abre la dirección que imprime la terminal. Conserva nexus_data.db junto al script para usar tu historial.\n\nSi hay un rastreo de pila, copia la última línea y el nombre de la función; adjunta la captura de la versión y la pantalla en la que ocurrió.")
        with st.expander("Datos de esta ejecución para soporte"):
            import sys
            script_path = str(Path(__file__).resolve())
            st.code(f"NEXUS: {APP_VERSION}\nArchivo ejecutado: {script_path}\nBase de datos: {Path(DB_PATH).resolve()}\nPython: {sys.version.split()[0]}\nIntérprete: {sys.executable}\nStreamlit: {st.__version__}", language=None)
            st.caption("Estos datos identifican el programa que está abierto. No se envían automáticamente a ningún servicio.")



def _set_site_selection_v712(widget_key, values):
    st.session_state[widget_key] = list(values)


def site_selector_v712(label, key, default=None, options=None, container=None):
    """A scope selector; selecting all never assigns one source file to many branches."""
    target = container if container is not None else st
    available = list(dict.fromkeys(list(SEDES) + [str(value) for value in (options or []) if str(value).strip()]))
    if key not in st.session_state:
        initial = available if default is None else list(default)
        st.session_state[key] = [value for value in initial if value in available]
    else:
        current = st.session_state.get(key, [])
        current = [current] if isinstance(current, str) else list(current or [])
        cleaned = [value for value in current if value in available]
        if current != cleaned:
            st.session_state[key] = cleaned
    controls = target.columns([3, 2])
    controls[0].button("Seleccionar todas las sedes", key=key + "__all", width="stretch",
        on_click=_set_site_selection_v712, args=(key, available))
    controls[1].button("Limpiar selección", key=key + "__clear", width="stretch",
        on_click=_set_site_selection_v712, args=(key, []))
    selected = target.multiselect(label, available, key=key, placeholder="Selecciona una o varias sedes")
    target.caption(f"{len(selected)} de {len(available)} sedes seleccionadas. Las sedes sin datos quedan pendientes de carga.")
    return selected


def init_stock_policy_v73():
    with db() as con:
        con.execute("""CREATE TABLE IF NOT EXISTS v73_stock_basis (
            site TEXT NOT NULL,scope TEXT NOT NULL DEFAULT '',method TEXT NOT NULL,
            months REAL,source_signature TEXT NOT NULL,updated_at TEXT NOT NULL,
            PRIMARY KEY(site,scope),CHECK(method IN ('days','months')),
            CHECK(months IS NULL OR months>0))""")


def stock_basis_v73(site, scope, metadata):
    with db() as con:
        row = con.execute("SELECT method,months,source_signature FROM v73_stock_basis WHERE site=? AND scope=?", (site,scope)).fetchone()
    if not row or row[0] == 'days':
        return None, False
    valid = bool(metadata.get('source_signature')) and row[2] == metadata['source_signature']
    return (float(row[1]) if valid else None), not valid


def save_stock_basis_v73(targets, method, months=None):
    if method not in {'days','months'}:
        raise ValueError('Selecciona días reales o meses confirmados.')
    if method == 'months' and (months is None or not np.isfinite(months) or not 0 < months <= 120):
        raise ValueError('Indica una cantidad de meses mayor que cero y hasta 120.')
    rows = []
    for site, scope, cutoff in targets:
        if site not in SEDES:
            raise ValueError('La sede no es válida.')
        _, metadata = operating_sales_v7(site, cutoff, scope)
        if method == 'months' and not metadata.get('source_signature'):
            raise ValueError(f'{site}: carga un historial de ventas válido antes de confirmar sus meses.')
        rows.append((site,scope,method,float(months) if method=='months' else None,
                     metadata.get('source_signature',''),datetime.now().isoformat()))
    with db() as con:
        con.executemany('INSERT OR REPLACE INTO v73_stock_basis VALUES(?,?,?,?,?,?)', rows)
        touch_v7(con)
    st.cache_data.clear()


STOCK_AUDIT_COLUMNS_V73 = ['Ventas utilizadas','Período ventas','Días del historial','Método de promedio',
    'Meses utilizados','Demanda Mensual','Cobertura mínima aplicada','Cobertura máxima aplicada',
    'Mínimo automático','Máximo automático','Plazo proveedor (días)','Seguridad (días)',
    'Reorden por plazo','Advertencia de plazo','Revisión del cálculo','Alcance del historial']


def attach_sales_ledger_v73(data, cuts, supplier=''):
    """A configured but excluded history is empty, never inventory-derived sales."""
    ledger = []
    version = v7_version()
    for site, cutoff in sorted(cuts.items()):
        products = data.loc[data['Sede'].eq(site)]
        scopes = products['Alcance del historial'].dropna().unique() if 'Alcance del historial' in products else []
        scope = str(scopes[0]) if len(scopes) == 1 else supplier
        with db() as con:
            configured = con.execute('SELECT 1 FROM v7_site_config WHERE site=? AND scope=?',(site,scope)).fetchone()
        if configured:
            rows = _sales_rows_for_reports_v7(version,((site,cutoff),),scope)
            if supplier and scope != supplier:
                rows = [r for r in rows if normalize_key(r.get('Proveedor','')) == normalize_key(supplier)]
        else:
            if 'Historial de ventas' in products:
                products = products.loc[products['Historial de ventas'].eq('Disponible')]
            columns = present_columns(products,['Sede','Código','Descripción','Departamento','Proveedor','Marca','Categoría','Ventas','Período ventas','Moneda ventas'])
            rows = json.loads(products[columns].to_json(orient='records',force_ascii=False))
        ledger.extend(rows)
    data.attrs['v7_sales_rows'] = ledger
    return data


def render_stock_calculation_v73(data):
    st.markdown('**Cómo se calculan tus niveles**')
    st.caption('Promedio mensual = unidades vendidas ÷ meses utilizados. Mínimo = promedio × cobertura mínima; máximo = promedio × cobertura máxima. Se redondea hacia arriba sólo el nivel final. Con 100 unidades / 2 meses y coberturas 0,5 / 1: mínimo 25 y máximo 50.')
    if data.empty:
        return
    manual = data['Política de stock'].eq('Manual')
    if manual.any():
        st.warning(f'{int(manual.sum())} productos tienen límites manuales: esos límites prevalecen sobre la fórmula. Puedes volver al cálculo automático más abajo.')
    if 'Revisión del cálculo' in data:
        issues = data.loc[data['Revisión del cálculo'].fillna('').ne(''),['Sede','Código','Revisión del cálculo']]
        if not issues.empty:
            st.warning(f'{len(issues)} productos requieren revisar su historial o los códigos pendientes antes de usar propuestas automáticas.')
    contexts = data[['Sede','Alcance del historial','Fecha del inventario']].drop_duplicates()
    overview = data[[c for c in ['Sede','Alcance del historial','Período ventas','Días del historial','Método de promedio','Meses utilizados','Cobertura mínima aplicada','Cobertura máxima aplicada'] if c in data]].drop_duplicates()
    st.dataframe(overview,hide_index=True,width='stretch')
    with st.expander('Elegir días reales o confirmar los meses del historial'):
        st.caption('Esta base se aplica al historial completo de cada sede/proveedor mostrado, incluso si filtraste algunos productos. Cambia el divisor del promedio; no cambia las fechas ni las ventas guardadas.')
        existing_months = data['Método de promedio'].eq('Meses confirmados').all()
        divisors = pd.to_numeric(data['Meses utilizados'],errors='coerce').dropna().unique()
        current_months = float(divisors[0]) if existing_months and len(divisors)==1 else 2.0
        context_token = repr((contexts.to_dict('records'),overview.to_dict('records')))
        if st.session_state.get('v73_basis_context') != context_token:
            st.session_state['v73_basis_method'] = 'Meses que confirmo' if existing_months else 'Días reales del reporte'
            st.session_state['v73_confirmed_months'] = current_months
            st.session_state['v73_basis_context'] = context_token
        method = st.radio('Base del promedio mensual',['Días reales del reporte','Meses que confirmo'],key='v73_basis_method')
        months = None
        if method == 'Meses que confirmo':
            months = st.number_input('Meses representados por estas ventas',min_value=0.01,max_value=120.0,step=0.25,key='v73_confirmed_months')
            st.caption('Si confirmas 2, se divide por 2 aunque las fechas sumen 67 días. Si cargas otro historial, tendrás que confirmar sus meses de nuevo.')
        st.dataframe(contexts,hide_index=True,width='stretch')
        if st.button('Aplicar base de cálculo a estas sedes',type='primary',key='v73_save_basis'):
            try:
                save_stock_basis_v73(list(contexts.itertuples(index=False,name=None)), 'months' if months is not None else 'days', months)
                st.rerun()
            except ValueError as exc:
                st.error(str(exc))
    with st.expander('Comprobar un producto paso a paso'):
        choices = list(range(len(data)))
        chosen = st.selectbox('Producto para comprobar',choices,format_func=lambda i:f"{data.iloc[i]['Sede']} · {data.iloc[i]['Código']} · {data.iloc[i]['Descripción']}",key='v73_audit_product')
        row = data.iloc[chosen]
        for field in ['Revisión del cálculo','Advertencia de plazo']:
            if str(row.get(field,'') or '').strip():
                st.info(str(row[field]))
        def n(value):
            return 'N/D' if pd.isna(value) else f'{float(value):,.4f}'.rstrip('0').rstrip('.')
        st.write(f"Ventas usadas: {n(row.get('Ventas utilizadas'))} unidades · divisor: {n(row.get('Meses utilizados'))} meses · promedio: {n(row.get('Demanda Mensual'))} unidades/mes.")
        st.write(f"Mínimo automático: redondear hacia arriba({n(row.get('Demanda Mensual'))} × {n(row.get('Cobertura mínima aplicada'))}) = {n(row.get('Mínimo automático'))}.")
        st.write(f"Máximo automático: redondear hacia arriba({n(row.get('Demanda Mensual'))} × {n(row.get('Cobertura máxima aplicada'))}) = {n(row.get('Máximo automático'))}.")
        st.write(f"Límites usados: mínimo {n(row.get('Stock Mínimo'))}, máximo {n(row.get('Stock Máximo'))}. Política: {row.get('Política de stock')}.")
        st.caption('La compra se activa al bajar del punto de reorden y busca completar el máximo. Las transferencias propuestas reducen la compra neta. El retiro compara existencias con máximo; no confirma una salida física.')

init_db()
init_finance_db()
init_v7_db()
init_warehouse_v71()
init_optional_costs_v71()
init_stock_policy_v73()
contacts = load_contacts()
st.caption(f"MAKROPETROL NEXUS {APP_VERSION} · Carga multisede · CxC opcional")

NAV_GROUPS: dict[str, list[str]] = {
    "▣ DIRECCIÓN": ["Centro de Control", "Desempeño Integral", "Cuentas por Cobrar", "Gestión Visual"],
    "▣ OPERACIÓN": ["Inventario", "Compras", "Mínimos y Máximos", "Proveedores", "Sin Rotación", "Redistribución"],
    "▣ INTELIGENCIA": ["Centro Visual", "Finanzas", "Alertas"],
    "▣ DATOS": ["Carga de datos", "Historial"],
    "▣ AUTOMATIZACIÓN": ["Automatización"],
    "▣ REPORTES": ["Reportes"],
}

with st.sidebar:
    st.markdown(
        """
        <div style="padding:8px 6px 14px">
            <div style="font-family:Manrope;font-size:1.2rem;font-weight:800">◈ MAKROPETROL <span style="color:#79dbca">NEXUS</span></div>
            <div style="font-size:.72rem;color:#7b8799;margin-top:3px">OPERACIÓN CONECTADA · V7</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    modo_gerencia = st.toggle("👔 Modo Gerencia", value=st.session_state.get("modo_gerencia", False), key="modo_gerencia",
                               help="Oculta columnas técnicas y parámetros; muestra KPIs, tendencias, alertas y recomendaciones.")
    st.caption("🔧 Modo Operativo" if not modo_gerencia else "Vista simplificada para dirección")
    st.divider()

    if st.session_state.get("nexus_view") in {"Cortes Diarios", "Datos Financieros"}:
        st.session_state["nexus_view"] = "Carga de datos"
    flat_views = [v for group in NAV_GROUPS.values() for v in group] + ["Guía de uso"]
    if st.session_state.get("nexus_view") not in flat_views:
        st.session_state["nexus_view"] = flat_views[0]

    primary_views = ["Centro de Control", "Carga de datos", "Inventario", "Compras", "Mínimos y Máximos", "Reportes"]
    for item in primary_views:
        active = st.session_state["nexus_view"] == item
        if st.button(item, key=f"nav_{item}", width="stretch", type="primary" if active else "secondary"):
            st.session_state["nexus_view"] = item
            st.rerun()
    with st.expander("Más herramientas", expanded=st.session_state["nexus_view"] not in primary_views):
        extra = st.selectbox("Abrir una herramienta", ["Selecciona…"] + [v for v in flat_views if v not in primary_views], key="v71_extra_navigation")
        if st.button("Abrir herramienta", key="v71_open_extra", disabled=extra == "Selecciona…"):
            _navigation_v7(extra)
    if st.button("Guía de uso · empezar aquí", key="v72_help_navigation", width="stretch"):
        _navigation_v7("Guía de uso")
    st.caption(f"Versión activa: NEXUS {APP_VERSION}")
    with st.expander("Comprobar archivo abierto"):
        st.code(str(Path(__file__).resolve()), language=None)
        st.caption("Esta es la ruta del código que está funcionando. Debe terminar en makropetrol_nexus_v7_3.py.")
    view = st.session_state["nexus_view"]

    st.divider()
    with st.popover("⚙️ Configuración del motor", width="stretch"):
        st.markdown("**Historial**")
        months_history = st.number_input("Meses del historial anterior sin fechas", min_value=1, max_value=36, value=3, step=1)
        st.caption("Reportes A2 con fechas: cambia su divisor en Mínimos y Máximos → Elegir días reales o confirmar los meses. Este control sólo se usa para datos anteriores sin período A2.")
        rolling_days = st.number_input("Ventana diaria (días)", min_value=1, max_value=365, value=30, step=1)
        st.markdown("**Cobertura**")
        min_coverage = st.number_input("Cobertura mínima (meses)", min_value=0.1, max_value=12.0, value=0.5, step=0.5)
        max_coverage = st.number_input("Cobertura máxima (meses)", min_value=0.5, max_value=24.0, value=1.0, step=0.5)
        st.markdown("**Proveedor**")
        lead_time_days = st.number_input("Lead time proveedor (días)", min_value=0, max_value=180, value=7, step=1)
        st.markdown("**Seguridad**")
        safety_days = st.number_input("Stock de seguridad (días)", min_value=0, max_value=90, value=3, step=1)
        st.markdown("**Demanda**")
        sales_mode_label = st.selectbox("Tipo de ventas", list(SALES_MODES.keys()))
        sales_mode = SALES_MODES[sales_mode_label]
        st.markdown("**ABC**")
        abc_basis = st.selectbox("Base ABC", ["Movimiento", "Inventario", "Capital inmovilizado"])

    dates = list_snapshot_dates()
    latest_global = dates[0] if dates else None
    status = f"● {latest_global}" if latest_global else "● Sin cortes"
    st.markdown(f'<span class="badge badge-green">{status}</span>', unsafe_allow_html=True)
    if st.button("🔄 Forzar recálculo", width="stretch", help="Limpia la caché interna y recalcula todo desde cero."):
        st.cache_data.clear()
        st.toast("Caché limpiada · recalculando motor", icon="🔄")
        st.rerun()

    with st.expander("Respaldo y mantenimiento"):
        if st.button("Preparar respaldo de mis datos", key="prepare_db_backup"):
            st.session_state["db_backup_bytes"] = backup_database_bytes()
            st.session_state["db_backup_name"] = f"nexus_respaldo_{datetime.now():%Y%m%d_%H%M%S}.db"
        if st.session_state.get("db_backup_bytes"):
            st.download_button("Descargar respaldo", st.session_state["db_backup_bytes"], st.session_state["db_backup_name"], "application/octet-stream", key="download_db_backup")
        st.warning("Reiniciar borra inventarios, ventas, cartera, tareas, conteos físicos e historial. Se creará un respaldo antes del borrado.")
        reset_text = st.text_input("Escribe BORRAR para habilitar el reinicio", key="reset_confirmation")
        if st.button("Reiniciar base de datos", width="stretch", disabled=reset_text != "BORRAR"):
            DB_PATH.with_name(f"nexus_respaldo_{datetime.now():%Y%m%d_%H%M%S_%f}.db").write_bytes(backup_database_bytes())
            reset_database()
            st.session_state.clear()
            st.toast("Sistema reiniciado a cero exitosamente.", icon="✅")
            st.rerun()

    st.caption(f"NEXUS v{APP_VERSION} · SQLite local")


@st.cache_data(show_spinner=False)
def load_live_results(
    data_version: str,
    months_history: int,
    min_coverage: float,
    max_coverage: float,
    lead_time_days: int,
    safety_days: int,
    abc_basis: str,
    sales_mode: str,
    rolling_days: int,
) -> tuple[dict[str, pd.DataFrame], dict[str, str]]:
    results: dict[str, pd.DataFrame] = {}
    dates_used: dict[str, str] = {}
    for site in SEDES:
        d = latest_snapshot_date(site)
        if not d:
            continue
        inv, sales, prev = load_snapshot(site, d)
        if "Código" not in inv.columns or "Código" not in sales.columns:
            continue
        results[site] = calculate_site(
            inv,
            sales,
            prev,
            months_history=int(months_history),
            min_coverage=float(min_coverage),
            max_coverage=float(max_coverage),
            lead_time_days=int(lead_time_days),
            safety_days=int(safety_days),
            abc_basis=abc_basis,
            sales_mode=sales_mode,
            rolling_days=int(rolling_days),
        )
        dates_used[site] = d
    return results, dates_used



@st.cache_data(show_spinner=False)
def load_supplier_live_results(
    data_version: str,
    supplier_version: str,
    supplier_name: str,
    selected_sites: tuple[str, ...],
    months_history: int,
    min_coverage: float,
    max_coverage: float,
    lead_time_days: int,
    safety_days: int,
    abc_basis: str,
    sales_mode: str,
    rolling_days: int,
) -> tuple[dict[str, pd.DataFrame], dict[str, str], dict[str, str]]:
    """
    Carga primero el corte dedicado del proveedor. Si no existe, reutiliza el
    corte maestro filtrando el inventario por proveedor y recalcula desde cero.
    """
    results: dict[str, pd.DataFrame] = {}
    dates_used: dict[str, str] = {}
    sources: dict[str, str] = {}
    key = normalize_key(supplier_name)
    if not key:
        return results, dates_used, sources

    for site in selected_sites:
        dedicated_date = latest_supplier_snapshot_date(site, supplier_name)
        if dedicated_date:
            inv, sales, prev = load_supplier_snapshot(site, supplier_name, dedicated_date)
            d = dedicated_date
            sources[site] = "Corte dedicado del proveedor"
        else:
            d = latest_snapshot_date(site)
            if not d:
                continue
            inv, sales, prev = load_snapshot(site, d)
            if "Proveedor" not in inv.columns:
                continue
            mask = inv["Proveedor"].fillna("").astype(str).map(normalize_key) == key
            inv = inv.loc[mask].copy()
            if inv.empty:
                continue
            codes = set(inv["Código"].astype(str))
            sales = sales[sales["Código"].astype(str).isin(codes)].copy()
            prev = prev[prev["Código"].astype(str).isin(codes)].copy()
            sources[site] = "Corte maestro filtrado por proveedor"

        if inv.empty or "Código" not in inv.columns:
            continue
        if sales.empty:
            sales = pd.DataFrame({"Código": inv["Código"].astype(str), "Ventas": 0})
        results[site] = calculate_site(
            inv, sales, prev,
            months_history=int(months_history), min_coverage=float(min_coverage),
            max_coverage=float(max_coverage), lead_time_days=int(lead_time_days),
            safety_days=int(safety_days), abc_basis=abc_basis, sales_mode=sales_mode,
            rolling_days=int(rolling_days),
        )
        dates_used[site] = d
    return results, dates_used, sources


@st.cache_data(show_spinner=False)
def build_dashboard_state(
    site_results: dict[str, pd.DataFrame],
) -> tuple[pd.DataFrame, dict[str, pd.DataFrame], pd.DataFrame, pd.DataFrame]:
    transfers = build_redistribution(site_results)
    site_results_adjusted = apply_redistribution_to_results(site_results, transfers)
    all_df = combine_sites(site_results_adjusted)
    alerts = build_alerts(all_df)
    return transfers, site_results_adjusted, all_df, alerts



@st.cache_data(show_spinner=False)
def load_full_dashboard_state(
    data_version: str,
    months_history: int,
    min_coverage: float,
    max_coverage: float,
    lead_time_days: int,
    safety_days: int,
    abc_basis: str,
    sales_mode: str,
    rolling_days: int,
):
    """Cachea el estado completo con argumentos pequeños para evitar re-hashear DataFrames en cada interacción."""
    site_results, dates_used = load_live_results(
        data_version, months_history, min_coverage, max_coverage,
        lead_time_days, safety_days, abc_basis, sales_mode, rolling_days,
    )
    transfers, adjusted, all_df, alerts = build_dashboard_state(site_results)
    return site_results, dates_used, transfers, adjusted, all_df, alerts


load_live_results = load_live_results_v7
load_supplier_live_results = load_supplier_live_results_v7

current_db_version = db_version()
current_supplier_version = supplier_db_version()
site_results, dates_used, transfers, site_results_adjusted, all_df, alerts = load_full_dashboard_state(
    current_db_version,
    int(months_history),
    float(min_coverage),
    float(max_coverage),
    int(lead_time_days),
    int(safety_days),
    abc_basis,
    sales_mode,
    int(rolling_days),
)

all_df = attach_sales_ledger_v73(all_df, dates_used)
latest_display = max(dates_used.values()) if dates_used else "sin cortes"
if not all_df.empty and "Costo informado" in all_df and all_df["Costo informado"].eq("No").any():
    st.warning("Hay productos sin costo informado. Las propuestas en unidades siguen disponibles; las valoraciones que dependen de esos costos quedan pendientes. Los gráficos monetarios pueden mostrar sólo los importes conocidos; revisa la cobertura en los informes.")
score = health_score(all_df)

st.markdown(
    f"""
    <div class="hero">
        <div class="brand">Makropetrol <span>NEXUS</span></div>
        <div class="hero-sub">Inventario que rota. Ventas que rinden. Cobranza bajo control.</div>
        <div class="status-row"><span class="dot"></span> {len(dates_used)} sedes con inventario <span>•</span> Último corte: {latest_display} <span>•</span> {escape(view)}</div>
    </div>
    """,
    unsafe_allow_html=True,
)


def _inventory_column_config(data: pd.DataFrame) -> dict[str, Any]:
    cfg: dict[str, Any] = {}
    if "Cobertura (meses)" in data.columns:
        cfg["Cobertura (meses)"] = st.column_config.NumberColumn(
            "Cobertura", format="%.1f meses", width="small"
        )
    if "Retiro Almacén" in data.columns and not data.empty:
        cfg["Retiro Almacén"] = st.column_config.ProgressColumn(
            "Retiro", min_value=0,
            max_value=max(1, int(pd.to_numeric(data["Retiro Almacén"], errors="coerce").max())),
            format="%d uds",
        )
    if "Compra Ajustada" in data.columns and not data.empty:
        cfg["Compra Ajustada"] = st.column_config.ProgressColumn(
            "Compra neta", min_value=0,
            max_value=max(1, int(pd.to_numeric(data["Compra Ajustada"], errors="coerce").max())),
            format="%d uds",
        )
    return cfg


@st.fragment
def render_inventario_fragment(all_df: pd.DataFrame, modo_gerencia: bool) -> None:
    if all_df.empty:
        empty_state("Sin inventario", "Procesa al menos una sede desde Carga de datos.", "📦")
        return

    fsite, fsupplier = st.columns(2)
    with fsite:
        selected_inventory_sites = site_selector_v712("Sedes del inventario", "inv_site_filter_v712", options=sorted(all_df["Sede"].unique()))
    data = all_df[all_df["Sede"].isin(selected_inventory_sites)].copy()
    with fsupplier:
        supplier_values = sorted([x for x in data.get("Proveedor", pd.Series(dtype=str)).dropna().astype(str).unique() if x.strip()], key=normalize_key)
        supplier_pick = st.selectbox("Proveedor", ["Todos"] + supplier_values, key="inv_supplier_filter") if supplier_values else "Todos"
    if supplier_pick != "Todos" and "Proveedor" in data.columns:
        data = data[data["Proveedor"].astype(str) == supplier_pick]

    search = st.text_input(
        "Buscar código o descripción",
        placeholder="Ej. TMB03 / ACEITE / FILTRO",
        key="inv_search",
    )
    if search:
        q = normalize_text(search)
        data = data[
            data["Código"].astype(str).str.lower().str.contains(q, na=False)
            | data["Descripción"].astype(str).str.lower().str.contains(q, na=False)
        ]

    state_options = sorted(data["Estado"].astype(str).unique())
    states = st.multiselect("Estado", state_options, default=state_options, key="inv_states")
    data = data[data["Estado"].astype(str).isin(states)]

    purchase_col = "Compra Ajustada" if "Compra Ajustada" in data.columns else "Compra Sugerida"
    c1, c2, c3, c4 = st.columns(4)
    with c1: kpi("Productos", integer(len(data)), "Filtrados", "blue")
    with c2: kpi("Críticos", integer(data["Estado"].astype(str).str.startswith("CRÍTICO").sum()), "Atención", "red")
    with c3: kpi("Compras", integer(data[purchase_col].sum()), "Unidades", "orange")
    with c4: kpi("Sobrestock", integer(data["Retiro Almacén"].sum()), "Unidades", "green")

    if modo_gerencia:
        cols = ["Sede", "Código", "Descripción", "ABC", "Estado", "Prioridad", "Acción",
                purchase_col, "Retiro Almacén", "Cobertura (meses)",
                "Valor Inventario ($)", "Capital Inmovilizado ($)"]
    else:
        cols = ["Sede", "Código", "Descripción", "Proveedor", "Categoría", "ABC", "Estado",
                "Prioridad", "Acción", "Ventas", "Demanda Mensual", "Existencia",
                "Stock Mínimo", "Stock Máximo", "Punto de Reorden", purchase_col,
                "Retiro Almacén", "Cobertura (meses)", "Rotación estimada (x/mes)",
                "Valor Inventario ($)", "Capital Inmovilizado ($)"]

    cols = [c for c in cols if c in data.columns]
    shown = data[cols].head(750)
    cfg = _inventory_column_config(shown)
    for c in ["Valor Inventario ($)", "Capital Inmovilizado ($)"]:
        if c in shown.columns:
            cfg[c] = st.column_config.NumberColumn(c, format="$%.2f")
    if "Prioridad" in shown.columns:
        cfg["Prioridad"] = st.column_config.TextColumn("Prioridad", width="small")
    if "Descripción" in shown.columns:
        cfg["Descripción"] = st.column_config.TextColumn("Descripción", width="large")

    st.dataframe(
        shown,
        width="stretch",
        hide_index=True,
        height=620,
        column_config=cfg,
    )
    if len(data) > len(shown):
        st.caption(f"Mostrando {len(shown):,} de {len(data):,} registros. La exportación conserva el 100% de los datos.")


@st.fragment
def render_compras_fragment(all_df: pd.DataFrame, modo_gerencia: bool) -> None:
    if all_df.empty:
        empty_state("Sin plan de compras", "Carga un corte diario para calcular necesidades.", "🛒")
        return
    purchase_col = "Compra Ajustada" if "Compra Ajustada" in all_df.columns else "Compra Sugerida"
    data = all_df[all_df[purchase_col] > 0]
    if data.empty:
        st.success("No existen compras pendientes después de considerar redistribuciones.")
        return

    c1, c2, c3 = st.columns(3)
    with c1: kpi("SKUs a comprar", integer(len(data)), "Compra neta", "orange")
    with c2: kpi("Unidades", integer(data[purchase_col].sum()), "Necesidad neta", "blue")
    with c3: kpi("Costo estimado", money((data[purchase_col] * data["Costo"]).sum()), "Costo disponible", "green")

    a, b, c = st.columns(3)
    with a:
        sites_filter = site_selector_v712("Sedes de la compra", "purchase_sites", options=sorted(data["Sede"].unique()))
    with b:
        priorities = st.multiselect("Prioridad", sorted(data["Prioridad"].unique()), default=sorted(data["Prioridad"].unique()), key="purchase_priorities")
    with c:
        supplier_values = sorted([x for x in data.get("Proveedor", pd.Series(dtype=str)).dropna().astype(str).unique() if x.strip()], key=normalize_key)
        purchase_supplier = st.selectbox("Proveedor", ["Todos"] + supplier_values, key="purchase_supplier") if supplier_values else "Todos"
    data = data[data["Sede"].isin(sites_filter) & data["Prioridad"].isin(priorities)]
    if purchase_supplier != "Todos" and "Proveedor" in data.columns:
        data = data[data["Proveedor"].astype(str) == purchase_supplier]

    if modo_gerencia:
        cols = ["Sede", "Código", "Descripción", "Proveedor", purchase_col, "Costo Compra Estimada ($)", "Prioridad", "Acción"]
    else:
        cols = ["Sede", "Código", "Descripción", "Proveedor", "ABC", "Ventas", "Existencia",
                "Punto de Reorden", purchase_col, "Costo", "Costo Compra Estimada ($)", "Prioridad", "Acción"]
    cols = [c for c in cols if c in data.columns]

    cfg = {
        "Costo": st.column_config.NumberColumn("Costo", format="$%.2f"),
        "Costo Compra Estimada ($)": st.column_config.NumberColumn("Costo estimado", format="$%.2f"),
    }
    if purchase_col in data.columns and not data.empty:
        cfg[purchase_col] = st.column_config.ProgressColumn(
            "Compra neta", min_value=0,
            max_value=max(1, int(data[purchase_col].max())),
            format="%d uds",
        )
    st.dataframe(data[cols].head(750), width="stretch", hide_index=True, height=580, column_config=cfg)
    if len(data) > 750:
        st.caption(f"Mostrando 750 de {len(data):,} registros. La exportación conserva todos.")

    supplier = data.groupby("Proveedor", dropna=False)[purchase_col].sum().reset_index().sort_values(purchase_col, ascending=False).head(10)
    fig = px.bar(
        supplier, x=purchase_col, y="Proveedor", orientation="h",
        title="Top proveedores por unidades a comprar",
        color_discrete_sequence=[NEXUS_ORANGE],
    )
    st.plotly_chart(chart_layout(fig, 320), width="stretch")


@st.fragment
def render_alertas_fragment(alerts: pd.DataFrame) -> None:
    if alerts.empty:
        empty_state("Todo tranquilo", "No se detectaron anomalías con los parámetros actuales.", "✓")
        return
    sev = st.multiselect("Severidad", sorted(alerts["Severidad"].unique()), default=sorted(alerts["Severidad"].unique()), key="alert_sev")
    typ = st.multiselect("Tipo", sorted(alerts["Tipo"].unique()), default=sorted(alerts["Tipo"].unique()), key="alert_type")
    data = alerts[alerts["Severidad"].isin(sev) & alerts["Tipo"].isin(typ)]
    counts = data["Severidad"].value_counts().reset_index()
    counts.columns = ["Severidad", "Cantidad"]
    fig = px.bar(
        counts, x="Severidad", y="Cantidad",
        title="Alertas por severidad",
        color="Severidad",
        color_discrete_map={"CRÍTICA": NEXUS_RED, "ALTA": NEXUS_ORANGE, "MEDIA": NEXUS_BLUE, "BAJA": NEXUS_GRAY},
    )
    st.plotly_chart(chart_layout(fig, 280), width="stretch")
    st.dataframe(data.head(750), width="stretch", hide_index=True, height=580)
    if len(data) > 750:
        st.caption(f"Mostrando 750 de {len(data):,} alertas.")


@st.fragment
def render_automation_fragment(
    site_results_adjusted: dict[str, pd.DataFrame],
    transfers: pd.DataFrame,
    all_df: pd.DataFrame,
    contacts: dict[str, dict[str, str]],
) -> None:
    contact_rows = []
    for site in SEDES:
        c = contacts.get(site, {})
        contact_rows.append({
            "Sede": site,
            "Supervisor": c.get("supervisor", ""),
            "WhatsApp": c.get("whatsapp", ""),
            "Webhook": c.get("webhook", ""),
        })
    edited = st.data_editor(
        pd.DataFrame(contact_rows),
        width="stretch",
        hide_index=True,
        num_rows="fixed",
        key="contacts_editor",
    )
    if st.button("💾 Guardar contactos y webhooks", width="stretch"):
        new_contacts = {}
        for row in edited.to_dict("records"):
            new_contacts[str(row["Sede"])] = {
                "supervisor": str(row.get("Supervisor", "")),
                "whatsapp": str(row.get("WhatsApp", "")),
                "webhook": str(row.get("Webhook", "")),
            }
        save_contacts(new_contacts)
        contacts.clear()
        contacts.update(new_contacts)
        st.toast("Contactos y webhooks guardados", icon="💾")

    if all_df.empty:
        empty_state("Automatización lista para conectar", "Configura los webhooks y procesa un corte diario.", "⚙")
        return

    site = st.selectbox("Sede a notificar", sorted(site_results_adjusted.keys()), key="automation_site")
    message = compose_digest(site, site_results_adjusted[site], transfers)
    st.code(message, language="text")

    supervisor = contacts.get(site, {})
    df_site = site_results_adjusted[site]
    purchase_col = "Compra Ajustada" if "Compra Ajustada" in df_site.columns else "Compra Sugerida"
    payload = {
        "event": "nexus.daily.logistics",
        "version": APP_VERSION,
        "timestamp": datetime.now().isoformat(timespec="microseconds"),
        "site": site,
        "supervisor": supervisor.get("supervisor", ""),
        "whatsapp": supervisor.get("whatsapp", ""),
        "health_score": health_score(df_site),
        "message": message,
        "metrics": {
            "critical": int(df_site["Estado"].astype(str).str.startswith("CRÍTICO").sum()),
            "purchase_skus": int((df_site[purchase_col] > 0).sum()),
            "purchase_units": int(df_site[purchase_col].sum()),
            "withdrawal_skus": int((df_site["Retiro Almacén"] > 0).sum()),
            "inventory_value": float(df_site["Valor Inventario ($)"].sum()),
            "capital_immobilized": float(df_site["Capital Inmovilizado ($)"].sum()),
        },
    }

    with st.expander("Vista técnica del payload", expanded=False):
        st.json(payload)

    if st.button("🚀 Enviar al webhook de esta sede", width="stretch"):
        future = post_webhook_async(supervisor.get("webhook", ""), payload)
        if future is None:
            st.error("No se envió: webhook vacío.")
        else:
            st.toast("Webhook disparado en segundo plano", icon="🚀")
            st.session_state["last_webhook_site"] = site
    st.info("Arquitectura: NEXUS → webhook → n8n/Make/API → CRM + proveedor de WhatsApp.")


# ============================================================
# VISTAS PRINCIPALES
# ============================================================

if view == "Centro de Control":
    if st.session_state.get("v7_import_success"):
        st.success(st.session_state.pop("v7_import_success"))
    sales_rows = all_df.attrs.get("v7_sales_rows", [])
    pending_rows = [r for r in sales_rows if not str(r.get("Código", "")).strip()]
    if pending_rows:
        st.warning(f"Hay {len(pending_rows)} filas vendidas sin código. Sus importes están incluidos en Finanzas y en Reportes → Ventas; completa sus códigos en Carga de datos para vincularlas al inventario.")
    if not all_df.empty and all_df.get("Corte desactualizado", pd.Series(dtype=str)).eq("Sí").any():
        st.warning("Hay ventas posteriores al último inventario. Las existencias siguen siendo las de su corte: carga un inventario reciente antes de ejecutar compras o retiros.")
    if not all_df.empty and all_df.get("Historial de ventas", pd.Series(dtype=str)).eq("Pendiente de cargar").any():
        st.info("Falta historial de ventas en alguna sede. Sus propuestas automáticas de mínimos, máximos, compras y retiros quedan pendientes; puedes fijar límites manuales.")
    if modo_gerencia:
        st.markdown('<div class="gov-banner">👔 Modo Gerencia activo · vista simplificada de indicadores y recomendaciones</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Centro de Control Ejecutivo</div>', unsafe_allow_html=True)
    shortcuts = st.columns(3)
    for column, label, target in [(shortcuts[0], "Inventario × ventas × CxC", "Desempeño Integral"),
                                   (shortcuts[1], "Mi plan de acción", "Gestión Visual"),
                                   (shortcuts[2], "Cargar datos de A2", "Datos Financieros")]:
        if column.button(label, width="stretch", key=f"shortcut_{target}"):
            st.session_state["nexus_view"] = target
            st.rerun()
    if len(set(dates_used.values())) > 1:
        st.caption("Las sedes tienen cortes de inventario de fechas distintas. El análisis integral comprueba las fechas del período seleccionado.")
    if all_df.empty:
        empty_state("NEXUS está listo", "Carga los reportes de Inventario y Ventas desde 'Carga de datos' para activar el tablero.", "◈")
    else:
        critical = int(all_df["Estado"].str.startswith("CRÍTICO").sum())
        purchase_col = "Compra Ajustada" if "Compra Ajustada" in all_df.columns else "Compra Sugerida"
        purchases = int((all_df[purchase_col] > 0).sum())
        purchase_units = int(all_df[purchase_col].sum())
        withdrawals = int((all_df["Retiro Almacén"] > 0).sum())
        inventory_value = float(monetary_sum_v71(all_df["Valor Inventario ($)"]))
        immobilized = float(monetary_sum_v71(all_df["Capital Inmovilizado ($)"]))
        avoided = float(monetary_sum_v71(transfers["Compra Evitada Estimada ($)"])) if not transfers.empty else 0.0
        optimal = int((all_df["Estado"] == "ÓPTIMO").sum())

        c1, c2, c3, c4 = st.columns(4)
        with c1: kpi("Inventario", money(inventory_value), f"{len(all_df):,} SKUs", "blue")
        with c2: kpi("Críticos", integer(critical), "Saldo negativo + bajo reorden", "red")
        with c3: kpi("Compra", f"{integer(purchase_units)} uds", f"{purchases} SKUs a comprar", "orange")
        with c4: kpi("Capital", money(immobilized), "Inmovilizado por sobrestock", "green")

        st.markdown('<div class="section-title">Salud logística</div>', unsafe_allow_html=True)
        st.markdown(
            f"""<div class="panel health-big">
                <div class="num">{score} / 100</div>
                <div class="lbl">{health_label(score)}</div>
            </div>""",
            unsafe_allow_html=True,
        )

        left, right = st.columns(2)
        with left:
            st.markdown('<div class="section-title">Inventario por sede</div>', unsafe_allow_html=True)
            by_site = all_df.groupby("Sede")["Valor Inventario ($)"].sum().sort_values(ascending=False)
            max_val = float(by_site.max()) if len(by_site) else 1.0
            rows_html = ""
            for sede, val in by_site.items():
                pct = max(4, round(val / max_val * 100)) if max_val else 4
                rows_html += (
                    f'<div class="op-row"><span>{SEDE_ICON.get(sede, "◆")} {sede}</span>'
                    f'<div class="bar-track"><div class="bar-fill" style="width:{pct}%"></div></div>'
                    f'<span>{money(val)}</span></div>'
                )
            st.markdown(f'<div class="panel">{rows_html}</div>', unsafe_allow_html=True)
        with right:
            st.markdown('<div class="section-title">Estado operativo</div>', unsafe_allow_html=True)
            state_rows = [
                ("Óptimo", optimal, "green"),
                ("Comprar", purchases, "orange"),
                ("Retirar", withdrawals, "blue"),
                ("Crítico", critical, "red"),
            ]
            rows_html = "".join(
                f'<div class="op-row"><span>{label}</span><span class="badge badge-{tone}">{integer(count)}</span></div>'
                for label, count, tone in state_rows
            )
            st.markdown(f'<div class="panel">{rows_html}</div>', unsafe_allow_html=True)

        n_transfers = int(len(transfers)) if not transfers.empty else 0
        st.markdown(
            f"""<div class="opp-banner">
                <span>♻️ {n_transfers} oportunidad(es) de redistribución</span>
                <span>💰 {money(avoided)} de compra potencial evitada</span>
            </div>""",
            unsafe_allow_html=True,
        )

        if not modo_gerencia:
            st.markdown('<div class="section-title">Prioridades inmediatas</div>', unsafe_allow_html=True)
            top_alerts = alerts.head(15) if not alerts.empty else pd.DataFrame()
            if top_alerts.empty:
                st.success("No hay alertas activas con los parámetros actuales.")
            else:
                st.dataframe(top_alerts, width="stretch", hide_index=True, height=380)

            a, b = st.columns(2)
            with a:
                action_counts = pd.DataFrame({
                    "Acción": ["Óptimo", "Bajo reorden", "Sobrestock", "Saldo negativo"],
                    "Cantidad": [optimal, int((all_df["Estado"] == "CRÍTICO — COMPRAR").sum()),
                                 int((all_df["Estado"] == "SOBRESTOCK — RETIRAR").sum()),
                                 int((all_df["Estado"] == "CRÍTICO — SALDO NEGATIVO").sum())],
                })
                fig = px.pie(action_counts, names="Acción", values="Cantidad", hole=.68, color="Acción", color_discrete_map={"Óptimo":NEXUS_GREEN,"Bajo reorden":NEXUS_ORANGE,"Sobrestock":NEXUS_BLUE,"Saldo negativo":NEXUS_RED})
                fig.update_traces(textposition="inside", textinfo="percent+label")
                st.plotly_chart(chart_layout(fig, 300), width="stretch")
            with b:
                top = all_df[all_df["Capital Inmovilizado ($)"] > 0].sort_values("Capital Inmovilizado ($)", ascending=False).head(12)
                if top.empty:
                    st.success("No existe capital inmovilizado por sobrestock.")
                else:
                    fig = px.bar(top.sort_values("Capital Inmovilizado ($)"), x="Capital Inmovilizado ($)", y="Descripción", orientation="h", hover_data=["Sede", "Código", "Retiro Almacén"], color_discrete_sequence=[NEXUS_RED])
                    st.plotly_chart(chart_layout(fig, 300), width="stretch")
        else:
            st.markdown('<div class="section-title">Recomendaciones para dirección</div>', unsafe_allow_html=True)
            recs = []
            if critical > 0:
                recs.append(f"🔴 Atender **{critical} SKUs críticos** (saldo negativo o bajo el punto de reorden) antes de nueva compra.")
            if n_transfers > 0:
                recs.append(f"♻️ Ejecutar las **{n_transfers} redistribuciones** sugeridas: evitan {money(avoided)} en compras nuevas.")
            if withdrawals > 0:
                recs.append(f"📤 Hay **{withdrawals} SKUs en sobrestock**; considerar retiro o promoción para liberar capital.")
            if not recs:
                recs.append("✅ La operación está dentro de parámetros saludables. No se requieren acciones urgentes.")
            st.markdown('<div class="panel">' + "".join(f"<p>{r}</p>" for r in recs) + "</div>", unsafe_allow_html=True)

elif view == "Carga de datos":
    render_load_center_v7()

elif view == "Centro Visual":
    st.markdown('<div class="section-title">Centro Visual · lectura rápida para gerencia</div>', unsafe_allow_html=True)
    if all_df.empty:
        empty_state("Sin datos visuales", "Carga cortes diarios primero.", "◌")
    else:
        purchase_col = "Compra Ajustada" if "Compra Ajustada" in all_df.columns else "Compra Sugerida"
        critical_df = all_df[all_df["Estado"].astype(str).str.startswith("CRÍTICO")].copy()
        purchase_df = all_df[all_df[purchase_col] > 0].copy()
        over_df = all_df[all_df["Retiro Almacén"] > 0].copy()
        k1,k2,k3,k4 = st.columns(4)
        with k1: kpi("Salud logística", f"{score}/100", health_label(score), "blue")
        with k2: kpi("Críticos", integer(len(critical_df)), "Atención inmediata", "red")
        with k3: kpi("Compra neta", integer(purchase_df[purchase_col].sum()), "Unidades", "orange")
        with k4: kpi("Sobrestock", integer(over_df["Retiro Almacén"].sum()), "Unidades", "green")

        v1,v2 = st.columns(2)
        with v1:
            site = all_df.groupby("Sede")["Valor Inventario ($)"].sum().reset_index()
            fig=px.bar(site,x="Sede",y="Valor Inventario ($)",title="Valor de inventario por sede", color="Sede", color_discrete_sequence=NEXUS_PALETTE)
            st.plotly_chart(chart_layout(fig,360),width="stretch")
        with v2:
            act=pd.DataFrame({"Tipo":["Crítico","Comprar","Retirar","Óptimo"],"Cantidad":[len(critical_df),len(purchase_df),len(over_df),int((all_df["Estado"]=="ÓPTIMO").sum())]})
            fig=px.pie(act,names="Tipo",values="Cantidad",hole=.62,title="Estado operativo")
            st.plotly_chart(chart_layout(fig,360),width="stretch")

        v3,v4 = st.columns(2)
        with v3:
            top=all_df[all_df["Capital Inmovilizado ($)"]>0].sort_values("Capital Inmovilizado ($)",ascending=False).head(10)
            if not top.empty:
                fig=px.bar(top.sort_values("Capital Inmovilizado ($)"),x="Capital Inmovilizado ($)",y="Descripción",orientation="h",title="Capital inmovilizado")
                st.plotly_chart(chart_layout(fig,360),width="stretch")
            else:
                empty_state("Sin capital inmovilizado", "No hay sobrestock valorizado.", "💰")
        with v4:
            if not transfers.empty:
                routes=transfers.groupby(["Origen","Destino"])["Unidades Sugeridas"].sum().reset_index()
                fig=px.bar(routes,x="Unidades Sugeridas",y="Origen",color="Destino",orientation="h",title="Redistribución sugerida")
                st.plotly_chart(chart_layout(fig,360),width="stretch")
            else:
                empty_state("Sin redistribuciones", "No se detectaron cruces de excedente/déficit.", "♻")


        v5, v6 = st.columns(2)
        with v5:
            if "Proveedor" in all_df.columns:
                top_sup = all_df.groupby("Proveedor", dropna=False)["Valor Inventario ($)"].sum().reset_index()
                top_sup = top_sup[top_sup["Proveedor"].astype(str).str.strip() != ""].sort_values("Valor Inventario ($)", ascending=False).head(12)
                if not top_sup.empty:
                    fig = px.bar(top_sup.sort_values("Valor Inventario ($)"), x="Valor Inventario ($)", y="Proveedor", orientation="h", title="Top proveedores por valor de inventario")
                    st.plotly_chart(chart_layout(fig, 360), width="stretch")
        with v6:
            dead = all_df[(all_df["Ventas"] <= 0) & (all_df["Existencia"] > 0)].copy()
            if not dead.empty:
                dead_site = dead.groupby("Sede")["Valor Inventario ($)"].sum().reset_index()
                fig = px.bar(dead_site, x="Sede", y="Valor Inventario ($)", color="Sede", title="Capital sin rotación por sede", color_discrete_sequence=NEXUS_PALETTE)
                st.plotly_chart(chart_layout(fig, 360), width="stretch")
            else:
                empty_state("Sin inventario detenido", "Todos los productos con stock presentan movimiento.", "✓")

elif view == "Inventario":
    render_warehouse_tools_v71(all_df)
    render_inventory_simple_v7(all_df)

elif view == "Compras":
    render_purchases_simple_v7(all_df, transfers, alerts)

elif view == "Mínimos y Máximos":
    render_stock_levels_v7(all_df)

elif view == "Proveedores":
    st.markdown('<div class="section-title">Inteligencia por proveedor / marca</div>', unsafe_allow_html=True)
    st.caption("Busca cualquiera de tus líneas, recalcula únicamente esa marca y genera reportes aislados por proveedor.")

    known_suppliers = list_known_suppliers(current_db_version, current_supplier_version)
    if not known_suppliers and all_df.empty:
        empty_state("Aún no hay proveedores", "Carga un corte general o un corte dedicado desde Carga de datos.", "🏷️")
    else:
        p1, p2 = st.columns([1.2, 1])
        with p1:
            supplier_query = st.text_input(
                "Buscar proveedor",
                value=st.session_state.get("supplier_focus", ""),
                placeholder="Escribe nombre o marca...",
                key="supplier_analysis_query",
            ).strip()
        matches = supplier_matches(supplier_query, known_suppliers, 20) if supplier_query else known_suppliers[:20]
        with p2:
            if matches:
                default_index = 0
                selected_supplier = st.selectbox("Proveedor detectado", matches, index=default_index, key="supplier_analysis_pick")
            else:
                selected_supplier = supplier_query
                st.text_input("Proveedor detectado", value=selected_supplier, disabled=True, key="supplier_analysis_no_match")

        supplier_name = selected_supplier or supplier_query
        if not supplier_name:
            empty_state("Selecciona una marca", "Escribe el nombre del proveedor para abrir su tablero dedicado.", "⌕")
        else:
            st.session_state["supplier_focus"] = supplier_name
            available_sites = SEDES.copy()
            selected_supplier_sites = site_selector_v712(
                "Sedes incluidas en el análisis",
                f"supplier_analysis_sites_{safe_slug(supplier_name)}",
                options=available_sites,
            )
            if not selected_supplier_sites:
                empty_state("Sin sedes", "Selecciona al menos una sede para calcular el proveedor.", "🏢")
            else:
                supplier_results, supplier_dates, supplier_sources = load_supplier_live_results(
                    current_db_version, current_supplier_version, supplier_name,
                    tuple(selected_supplier_sites), int(months_history), float(min_coverage), float(max_coverage),
                    int(lead_time_days), int(safety_days), abc_basis, sales_mode, int(rolling_days),
                )
                sup_transfers, supplier_adjusted, supplier_df, supplier_alerts = build_dashboard_state(supplier_results)

                if supplier_df.empty:
                    empty_state(
                        f"Sin datos para {supplier_name}",
                        "No encontré esta marca en el corte maestro ni en cortes dedicados. Ve a Carga de datos → Opciones del período y del proveedor y carga sus archivos.",
                        "🏷️",
                    )
                else:
                    purchase_col = "Compra Ajustada" if "Compra Ajustada" in supplier_df.columns else "Compra Sugerida"
                    inventory_value = float(monetary_sum_v71(supplier_df["Valor Inventario ($)"]))
                    purchase_units = int(supplier_df[purchase_col].sum())
                    purchase_value = float(monetary_sum_v71((supplier_df[purchase_col] * supplier_df["Costo"])))
                    zero_df = supplier_df[(supplier_df["Ventas"] <= 0) & (supplier_df["Existencia"] > 0)].copy()
                    zero_value = float(monetary_sum_v71(zero_df["Valor Inventario ($)"])) if not zero_df.empty else 0.0
                    sup_score = health_score(supplier_df)
                    source_summary = " · ".join(f"{site}: {supplier_sources.get(site,'—')}" for site in supplier_dates)

                    st.markdown(
                        f"""<div class="scope-banner"><div><div class="scope-title">🏷️ {supplier_name}</div>
                        <div class="scope-sub">{source_summary or 'Proveedor cargado'} · Salud {sup_score}/100</div></div>
                        <span class="badge badge-cyan">{len(supplier_df):,} SKU</span></div>""",
                        unsafe_allow_html=True,
                    )
                    k1, k2, k3, k4 = st.columns(4)
                    with k1: kpi("Inventario marca", money(inventory_value), f"{len(supplier_df):,} SKU", "blue")
                    with k2: kpi("Compra neta", f"{integer(purchase_units)} uds", money(purchase_value), "orange")
                    with k3: kpi("Sin rotación", integer(len(zero_df)), money(zero_value), "red")
                    with k4: kpi("Salud proveedor", f"{sup_score}/100", health_label(sup_score), "green")

                    tab_res, tab_prod, tab_buy, tab_zero, tab_reports = st.tabs([
                        "📊 Resumen", "📦 Productos", "🛒 Compras", "🧊 Sin rotación", "📄 Reportes"
                    ])
                    with tab_res:
                        a, b = st.columns(2)
                        with a:
                            by_site = supplier_df.groupby("Sede")["Valor Inventario ($)"].sum().reset_index()
                            fig = px.bar(by_site, x="Sede", y="Valor Inventario ($)", color="Sede", title="Valor de inventario de la marca por sede", color_discrete_sequence=NEXUS_PALETTE)
                            st.plotly_chart(chart_layout(fig, 340), width="stretch")
                        with b:
                            state = supplier_df["Estado"].astype(str).value_counts().reset_index()
                            state.columns = ["Estado", "Cantidad"]
                            fig = px.pie(state, names="Estado", values="Cantidad", hole=.66, title="Estado logístico de la marca")
                            st.plotly_chart(chart_layout(fig, 340), width="stretch")
                        c, d = st.columns(2)
                        with c:
                            buy_site = supplier_df.groupby("Sede")[purchase_col].sum().reset_index()
                            fig = px.bar(buy_site, x="Sede", y=purchase_col, title="Unidades a comprar por sede", color="Sede", color_discrete_sequence=NEXUS_PALETTE)
                            st.plotly_chart(chart_layout(fig, 330), width="stretch")
                        with d:
                            top = supplier_df.sort_values("Valor Inventario ($)", ascending=False).head(12)
                            fig = px.bar(top.sort_values("Valor Inventario ($)"), x="Valor Inventario ($)", y="Descripción", orientation="h", title="Productos con mayor valor de inventario")
                            st.plotly_chart(chart_layout(fig, 330), width="stretch")

                    with tab_prod:
                        prod_cols = ["Sede","Código","Descripción","Categoría","ABC","Estado","Ventas","Existencia","Punto de Reorden",purchase_col,"Retiro Almacén","Cobertura (meses)","Valor Inventario ($)"]
                        prod_cols = [c for c in prod_cols if c in supplier_df.columns]
                        st.dataframe(supplier_df[prod_cols], width="stretch", hide_index=True, height=620)

                    with tab_buy:
                        buy_df = supplier_df[supplier_df[purchase_col] > 0].copy()
                        if buy_df.empty:
                            st.success("No hay compras pendientes para esta marca con los parámetros actuales.")
                        else:
                            cols = ["Sede","Código","Descripción","ABC","Existencia","Punto de Reorden",purchase_col,"Costo","Prioridad","Acción"]
                            cols = [c for c in cols if c in buy_df.columns]
                            st.dataframe(buy_df[cols], width="stretch", hide_index=True, height=540)

                    with tab_zero:
                        if zero_df.empty:
                            st.success("Todos los productos con stock registran movimiento en el corte actual.")
                        else:
                            top_zero = zero_df.sort_values("Valor Inventario ($)", ascending=False).head(20)
                            fig = px.bar(top_zero.sort_values("Valor Inventario ($)"), x="Valor Inventario ($)", y="Descripción", color="Sede", orientation="h", title="Capital de la marca sin rotación")
                            st.plotly_chart(chart_layout(fig, 380), width="stretch")
                            st.dataframe(zero_df[[c for c in ["Sede","Código","Descripción","Existencia","Costo","Valor Inventario ($)","ABC"] if c in zero_df.columns]], width="stretch", hide_index=True, height=450)

                    with tab_reports:
                        st.markdown("**Reportes por gestión de este proveedor**")
                        slug = safe_slug(supplier_name)
                        sup_latest = max(supplier_dates.values()) if supplier_dates else latest_display
                        report_token = hashlib.sha256(repr((supplier_name, supplier_dates, current_db_version, current_supplier_version, months_history, min_coverage, max_coverage, lead_time_days)).encode()).hexdigest()[:20]
                        supplier_df = attach_sales_ledger_v73(supplier_df, supplier_dates, supplier_name)
                        render_reports_v7(supplier_df, sup_transfers, supplier_alerts, f"{sup_latest} · {supplier_name}")
                        purchase_sites = sorted(supplier_df["Sede"].unique())
                        purchase_site = st.selectbox("Orden de compra para la sede", purchase_sites, key=f"supplier_purchase_site_{slug}")
                        purchase_state_key = f"supplier_purchase_pdf_{report_token}_{safe_slug(purchase_site)}"
                        if st.button("🛒 Preparar orden de compra de esta marca", width="stretch", key=f"prepare_oc_{report_token}_{safe_slug(purchase_site)}"):
                            site_df = supplier_df[supplier_df["Sede"] == purchase_site].copy()
                            st.session_state[purchase_state_key] = build_purchase_pdf(
                                site_df, purchase_site, sup_latest, supplier=supplier_name
                            )
                        purchase_pdf = st.session_state.get(purchase_state_key)
                        if purchase_pdf:
                            st.download_button(
                                "⬇ Descargar orden de compra",
                                purchase_pdf,
                                f"OC_{slug}_{safe_slug(purchase_site)}_{sup_latest}.pdf",
                                DOC_MIME_PDF,
                                width="stretch",
                            )


elif view == "Sin Rotación":
    st.markdown('<div class="section-title">Productos sin rotación</div>', unsafe_allow_html=True)
    st.caption("Detecta inventario con existencia positiva y cero venta en el corte actual; Cruza el historial diario y los períodos A2. Un período acumulado no informa la fecha exacta de la última venta.")
    if all_df.empty:
        empty_state("Sin datos de rotación", "Carga al menos un corte diario para activar el análisis.", "🧊")
    else:
        zero = all_df[(pd.to_numeric(all_df["Ventas"], errors="coerce").fillna(0) <= 0) & (pd.to_numeric(all_df["Existencia"], errors="coerce").fillna(0) > 0)].copy()
        zero = zero[zero.get("Historial de ventas", pd.Series("Disponible", index=zero.index)).eq("Disponible")]
        if zero.empty:
            st.success("No hay productos con stock y cero ventas en el corte actual.")
        else:
            rotation_hist = build_rotation_history_v7(current_db_version, sales_mode)
            if not rotation_hist.empty:
                zero = zero.merge(rotation_hist, on=["Sede", "Código"], how="left")
                ref_date = pd.to_datetime(latest_display, errors="coerce")
                base_date = pd.to_datetime(zero["Última venta positiva"].fillna(zero["Primera fecha"]), errors="coerce")
                row_ref = pd.to_datetime(zero["Última fecha"], errors="coerce").fillna(ref_date)
                zero["Días sin venta detectada"] = (row_ref - base_date).dt.days.clip(lower=0)

            f1, f2, f3 = st.columns(3)
            with f1:
                site_filter = site_selector_v712("Sedes sin rotación", "zero_sites", options=sorted(zero["Sede"].unique()))
            supplier_opts = sorted([x for x in zero["Proveedor"].dropna().astype(str).unique() if x.strip()], key=normalize_key) if "Proveedor" in zero.columns else []
            with f2:
                supplier_filter = st.multiselect("Proveedores", supplier_opts, default=supplier_opts, key="zero_suppliers") if supplier_opts else []
            with f3:
                zero_search = st.text_input("Buscar producto", placeholder="Código o descripción", key="zero_search")

            data = zero[zero["Sede"].isin(site_filter)].copy()
            if supplier_opts:
                data = data[data["Proveedor"].astype(str).isin(supplier_filter)]
            if zero_search:
                q = normalize_text(zero_search)
                data = data[
                    data["Código"].astype(str).str.lower().str.contains(q, na=False) |
                    data["Descripción"].astype(str).str.lower().str.contains(q, na=False)
                ]

            value = float(monetary_sum_v71(data["Valor Inventario ($)"])) if not data.empty else 0.0
            units = int(pd.to_numeric(data["Existencia"], errors="coerce").fillna(0).sum()) if not data.empty else 0
            inv_total = float(monetary_sum_v71(all_df["Valor Inventario ($)"])) if not all_df.empty else 0.0
            pct = (value / inv_total * 100) if inv_total else 0.0
            longest = int(pd.to_numeric(data.get("Días sin venta detectada", pd.Series(dtype=float)), errors="coerce").max()) if "Días sin venta detectada" in data.columns and data["Días sin venta detectada"].notna().any() else 0
            k1,k2,k3,k4 = st.columns(4)
            with k1: kpi("SKUs sin rotación", integer(len(data)), "Stock > 0 · ventas = 0", "red")
            with k2: kpi("Unidades detenidas", integer(units), "Existencia física", "orange")
            with k3: kpi("Capital detenido", money(value), f"{pct:.1f}% del inventario", "red")
            with k4: kpi("Mayor inactividad", f"{longest} días" if longest else "—", "Según historial cargado", "blue")

            z1, z2 = st.columns(2)
            with z1:
                if "Proveedor" in data.columns:
                    by_sup = data.groupby("Proveedor", dropna=False)["Valor Inventario ($)"].sum().reset_index().sort_values("Valor Inventario ($)", ascending=False).head(15)
                    fig = px.bar(by_sup.sort_values("Valor Inventario ($)"), x="Valor Inventario ($)", y="Proveedor", orientation="h", title="Capital sin rotación por proveedor")
                    st.plotly_chart(chart_layout(fig, 390), width="stretch")
            with z2:
                by_site = data.groupby("Sede")["Valor Inventario ($)"].sum().reset_index()
                fig = px.pie(by_site, names="Sede", values="Valor Inventario ($)", hole=.62, title="Distribución del capital detenido")
                st.plotly_chart(chart_layout(fig, 390), width="stretch")

            if not data.empty:
                tree = data.copy()
                tree["Proveedor"] = tree.get("Proveedor", pd.Series("Sin proveedor", index=tree.index)).astype(str).replace({"nan": "Sin proveedor", "": "Sin proveedor"})
                tree["Producto"] = tree["Descripción"].where(tree["Descripción"].astype(str).str.strip() != "", tree["Código"].astype(str))
                tree = tree[tree["Valor Inventario ($)"] > 0].copy()
                if not tree.empty:
                    fig = px.treemap(tree.head(500), path=["Sede", "Proveedor", "Producto"], values="Valor Inventario ($)", title="Mapa de capital sin rotación")
                    st.plotly_chart(chart_layout(fig, 480), width="stretch")

            cols = ["Sede","Proveedor","Código","Descripción","ABC","Existencia","Costo","Valor Inventario ($)","Última venta positiva","Último período con ventas","Días sin venta detectada","Marca","Departamento","Categoría","Acción"]
            cols = [c for c in cols if c in data.columns]
            st.dataframe(data.sort_values("Valor Inventario ($)", ascending=False)[cols].head(1000), width="stretch", hide_index=True, height=620)
            xls = io.BytesIO()
            data[cols].to_excel(xls, index=False, sheet_name="Sin Rotacion")
            st.download_button("⬇ Exportar productos sin rotación", xls.getvalue(), f"NEXUS_Sin_Rotacion_{latest_display}.xlsx", DOC_MIME_XLSX)

elif view == "Redistribución":
    st.markdown('<div class="section-title">Motor de redistribución entre sedes</div>', unsafe_allow_html=True)
    if transfers.empty:
        empty_state("No hay transferencias sugeridas", "No se detectaron al mismo tiempo excedentes y déficits compatibles.", "♻️")
    else:
        avoided = float(monetary_sum_v71(transfers["Compra Evitada Estimada ($)"]))
        units = int(transfers["Unidades Sugeridas"].sum())
        c1, c2, c3 = st.columns(3)
        with c1: kpi("Transferencias", integer(len(transfers)), "Movimientos sugeridos", "green")
        with c2: kpi("Unidades", integer(units), "Redistribución total", "blue")
        with c3: kpi("Compra evitada", money(avoided), "Ahorro estimado", "orange")
        st.dataframe(
            transfers.style.format({"Costo Unitario ($)": "${:,.2f}", "Compra Evitada Estimada ($)": "${:,.2f}"}),
            width="stretch",
            hide_index=True,
            height=600,
        )
        route = transfers.groupby(["Origen", "Destino"])["Unidades Sugeridas"].sum().reset_index()
        if not route.empty:
            st.markdown('<div class="section-title">Flujo entre sedes</div>', unsafe_allow_html=True)
            nodes = sorted(set(route["Origen"]) | set(route["Destino"]))
            node_index = {n: i for i, n in enumerate(nodes)}
            palette = ["#3568f5", "#17aabe", "#14a36a", "#ef982d", "#df5555", "#244bd7", "#7b8799"]
            node_colors = [palette[i % len(palette)] for i in range(len(nodes))]
            link_colors = ["rgba(" + ",".join(str(int(node_colors[node_index[o]].lstrip("#")[i:i+2], 16)) for i in (0, 2, 4)) + ",0.333)" for o in route["Origen"]]
            sankey = go.Figure(go.Sankey(
                arrangement="snap",
                node=dict(
                    label=[f"{SEDE_ICON.get(n,'◆')} {n}" for n in nodes],
                    color=node_colors,
                    pad=18,
                    thickness=16,
                    line=dict(color="rgba(0,0,0,0)", width=0),
                ),
                link=dict(
                    source=[node_index[o] for o in route["Origen"]],
                    target=[node_index[d] for d in route["Destino"]],
                    value=route["Unidades Sugeridas"],
                    color=link_colors,
                    label=[f"{int(v):,} uds" for v in route["Unidades Sugeridas"]],
                ),
            ))
            st.plotly_chart(chart_layout(sankey, 380), width="stretch")
        wb = io.BytesIO()
        transfers.to_excel(wb, index=False, sheet_name="Redistribucion")
        st.download_button("⬇ Exportar plan de redistribución", wb.getvalue(), "Plan_Redistribucion_Makropetrol.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif view == "Desempeño Integral":
    render_performance()

elif view == "Cuentas por Cobrar":
    render_performance(receivables_only=True)

elif view == "Datos Financieros":
    render_load_center_v7()

elif view == "Gestión Visual":
    render_action_board(all_df)

elif view == "Finanzas":
    st.markdown('<div class="section-title">Control financiero del inventario</div>', unsafe_allow_html=True)
    if all_df.empty:
        empty_state("Sin datos financieros", "Procesa cortes para calcular inventario, compras y capital inmovilizado.", "💰")
    else:
        inv_value = float(monetary_sum_v71(all_df["Valor Inventario ($)"]))
        immobilized = float(monetary_sum_v71(all_df["Capital Inmovilizado ($)"]))
        purchase_col = "Compra Ajustada" if "Compra Ajustada" in all_df.columns else "Compra Sugerida"
        purchase_value = float(monetary_sum_v71((all_df[purchase_col] * all_df["Costo"])))
        no_cost = int((all_df["Costo"] <= 0).sum())
        c1, c2, c3, c4 = st.columns(4)
        with c1: kpi("Valor inventario", money(inv_value), "Existencia × costo", "blue")
        with c2: kpi("Capital inmovilizado", money(immobilized), "Sobrestock", "red")
        with c3: kpi("Compra estimada", money(purchase_value), "Compra neta", "orange")
        with c4: kpi("Sin costo", integer(no_cost), "Revisar fuente", "green")
        a, b = st.columns(2)
        by_site = all_df.groupby("Sede")["Valor Inventario ($)"].sum().reset_index()
        with a:
            fig = px.bar(by_site, x="Sede", y="Valor Inventario ($)", title="Valor de inventario por sede", color="Sede", color_discrete_sequence=NEXUS_PALETTE)
            st.plotly_chart(chart_layout(fig), width="stretch")
        with b:
            by_site2 = all_df.groupby("Sede")["Capital Inmovilizado ($)"].sum().reset_index()
            fig = px.bar(by_site2, x="Sede", y="Capital Inmovilizado ($)", title="Capital inmovilizado por sede", color="Sede", color_discrete_sequence=NEXUS_PALETTE)
            st.plotly_chart(chart_layout(fig), width="stretch")
        st.markdown('<div class="section-title">Capital inmovilizado por producto</div>', unsafe_allow_html=True)
        financial = all_df[all_df["Capital Inmovilizado ($)"] > 0].sort_values("Capital Inmovilizado ($)", ascending=False).head(100)
        fin_cols = ["Sede", "Código", "Descripción", "Costo", "Existencia", "Retiro Almacén", "Capital Inmovilizado ($)"] if not modo_gerencia else ["Sede", "Descripción", "Retiro Almacén", "Capital Inmovilizado ($)"]
        max_cap = float(financial["Capital Inmovilizado ($)"].max()) if not financial.empty else 1.0
        st.dataframe(
            financial[fin_cols],
            width="stretch",
            hide_index=True,
            height=540,
            column_config={
                "Costo": st.column_config.NumberColumn(format="$%.2f"),
                "Capital Inmovilizado ($)": st.column_config.ProgressColumn(
                    "Capital inmovilizado", min_value=0, max_value=max(1.0, max_cap), format="$%.0f"
                ),
            },
        )

elif view == "Alertas":
    st.markdown('<div class="section-title">Centro de alertas</div>', unsafe_allow_html=True)
    render_alertas_fragment(alerts)

elif view == "Automatización":
    st.markdown('<div class="section-title">Hub de automatización</div>', unsafe_allow_html=True)
    render_automation_fragment(site_results_adjusted, transfers, all_df, contacts)

elif view == "Historial":
    st.markdown('<div class="section-title">Historial de cortes y trazabilidad</div>', unsafe_allow_html=True)
    with db() as con:
        log = pd.read_sql_query(
            "SELECT processed_at, snapshot_date, site, inventory_rows, sales_rows FROM processing_log ORDER BY processed_at DESC LIMIT 300",
            con,
        )
    if log.empty:
        empty_state("Sin historial", "Los cortes que guardes aparecerán aquí.", "◷")
    else:
        st.dataframe(log, width="stretch", hide_index=True, height=620)

elif view == "Guía de uso":
    render_help_v72()

elif view == "Reportes":
    report_route = st.radio("¿Qué informe necesitas?", ["Informe final de almacén", "Análisis financiero · opcional", "Órdenes y movimientos"], horizontal=True, key="v71_report_route")
    if report_route == "Informe final de almacén":
        st.info("Con inventario y ventas puedes preparar el informe final de almacén: existencias, mínimos, máximos, compras, sobrestock y redistribución. CxC es opcional y no impide esta descarga.")
        render_reports_v7(all_df, transfers, alerts, latest_display)
    elif report_route == "Análisis financiero · opcional":
        st.caption("Los indicadores de inventario y ventas se calculan con sus fuentes disponibles. La cartera queda como no informada cuando no se carga CxC.")
        render_performance()
    else:
        render_operational_documents_v7()
    with st.expander("Mis funciones: qué automatiza NEXUS y qué datos necesita"):
        render_role_guide_v71()

if not all_df.empty and view not in {"Reportes", "Carga de datos"}:
    if st.button("Ver reportes por área", key="go_exports", width="content"):
        st.session_state["nexus_view"] = "Reportes"
        st.rerun()


st.markdown(
    f"<div class='footer'>Makropetrol NEXUS v{APP_VERSION} · Inventario · Ventas · Cuentas por cobrar · Gestión visual</div>",
    unsafe_allow_html=True,
)
