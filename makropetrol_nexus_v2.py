# ============================================================
# MAKROPETROL ERP — NEXUS LOGISTICS v2 (CORREGIDO Y ESTABLE)
# Dashboard ejecutivo + inventario + compras + retiros + logística
# Requiere: streamlit, pandas, numpy, openpyxl, plotly
# ============================================================

import io
import re
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go


# ============================================================
# CONFIGURACIÓN
# ============================================================

st.set_page_config(
    page_title="Makropetrol Nexus | Control Logístico",
    page_icon="◈",
    layout="wide",
    initial_sidebar_state="expanded",
)

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

# ============================================================
# ESTILO — interfaz premium / glassmorphism / dashboard
# ============================================================

st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@500;600;700;800&display=swap');

:root{
    --bg:#f4f7fb;
    --surface:rgba(255,255,255,.82);
    --surface-solid:#ffffff;
    --line:#e5eaf2;
    --text:#162033;
    --muted:#7b8799;
    --blue:#3568f5;
    --blue2:#244bd7;
    --cyan:#1ab7c9;
    --green:#16a66a;
    --orange:#f39a2f;
    --red:#e55353;
    --shadow:0 12px 34px rgba(38,55,82,.09);
}

html, body, [class*="css"]{
    font-family:'DM Sans',sans-serif;
}

.stApp{
    background:
        radial-gradient(circle at 8% 2%, rgba(53,104,245,.09), transparent 28%),
        radial-gradient(circle at 95% 8%, rgba(26,183,201,.08), transparent 26%),
        linear-gradient(135deg,#f7f9fc 0%,#eef3f9 100%);
    color:var(--text);
}

header[data-testid="stHeader"]{
    background:rgba(244,247,251,.70);
}

.block-container{
    padding-top:1.6rem;
    padding-bottom:3rem;
    max-width:1500px;
}

section[data-testid="stSidebar"]{
    background:rgba(250,252,255,.94);
    border-right:1px solid var(--line);
}

section[data-testid="stSidebar"] > div{
    padding-top:1.3rem;
}

h1,h2,h3,h4{
    font-family:'Manrope',sans-serif !important;
    color:var(--text) !important;
    letter-spacing:-.025em;
}

.hero{
    background:linear-gradient(135deg,rgba(255,255,255,.92),rgba(255,255,255,.63));
    border:1px solid rgba(255,255,255,.95);
    box-shadow:var(--shadow);
    border-radius:26px;
    padding:26px 30px 22px;
    position:relative;
    overflow:hidden;
}

.hero:after{
    content:"";
    position:absolute;
    width:270px;height:270px;
    right:-90px;top:-120px;
    border-radius:50%;
    background:radial-gradient(circle,rgba(53,104,245,.16),transparent 68%);
}

.brand{
    font-family:'Manrope',sans-serif;
    font-size:1.9rem;
    font-weight:800;
    letter-spacing:-.04em;
}

.brand span{color:var(--blue);}

.hero-sub{
    color:var(--muted);
    margin-top:4px;
    font-size:.94rem;
}

.status-row{
    display:flex;
    align-items:center;
    gap:8px;
    margin-top:16px;
    color:#657188;
    font-size:.78rem;
}

.dot{
    width:8px;height:8px;border-radius:50%;
    background:#1abf73;
    box-shadow:0 0 0 5px rgba(26,191,115,.10);
}

.section-title{
    font-family:'Manrope',sans-serif;
    font-size:1.08rem;
    font-weight:800;
    margin:24px 0 10px;
}

.kpi-grid{
    display:grid;
    grid-template-columns:repeat(4,1fr);
    gap:14px;
    margin-top:18px;
}

.kpi{
    background:var(--surface);
    border:1px solid rgba(255,255,255,.95);
    box-shadow:var(--shadow);
    border-radius:20px;
    padding:18px 19px;
    min-height:122px;
    position:relative;
    overflow:hidden;
}

.kpi:before{
    content:"";
    position:absolute;
    left:0;top:0;bottom:0;
    width:4px;
    background:var(--blue);
}

.kpi.green:before{background:var(--green)}
.kpi.orange:before{background:var(--orange)}
.kpi.red:before{background:var(--red)}

.kpi-label{
    color:var(--muted);
    font-size:.76rem;
    font-weight:700;
    text-transform:uppercase;
    letter-spacing:.045em;
}

.kpi-value{
    font-family:'Manrope',sans-serif;
    font-size:1.72rem;
    font-weight:800;
    margin-top:9px;
    letter-spacing:-.04em;
}

.kpi-note{
    color:#9aa5b5;
    font-size:.73rem;
    margin-top:3px;
}

.panel{
    background:var(--surface);
    border:1px solid rgba(255,255,255,.95);
    box-shadow:var(--shadow);
    border-radius:22px;
    padding:18px;
}

.badge{
    display:inline-block;
    border-radius:999px;
    padding:5px 10px;
    font-size:.72rem;
    font-weight:800;
}

.badge-blue{background:#edf2ff;color:#355bd7}
.badge-green{background:#eafaf3;color:#158456}
.badge-orange{background:#fff5e7;color:#bc6d12}
.badge-red{background:#fff0f0;color:#c53c3c}

.sidebar-logo{
    padding:8px 8px 18px;
    font-family:'Manrope',sans-serif;
    font-weight:800;
    font-size:1.18rem;
}

.sidebar-logo small{
    display:block;
    color:var(--muted);
    font-family:'DM Sans',sans-serif;
    font-size:.72rem;
    font-weight:500;
    margin-top:3px;
}

div[data-baseweb="select"] > div,
.stTextInput > div > div,
.stNumberInput > div > div{
    border-radius:12px !important;
    border-color:var(--line) !important;
    background:white !important;
}

.stButton > button{
    border-radius:12px;
    border:1px solid #dce4f2;
    background:#fff;
    color:#2d3b55;
    font-weight:700;
    min-height:42px;
    transition:.18s ease;
}

.stButton > button:hover{
    border-color:#b8c9f8;
    color:var(--blue);
    transform:translateY(-1px);
    box-shadow:0 8px 18px rgba(53,104,245,.10);
}

.stDownloadButton > button{
    border-radius:12px;
    border:0;
    background:linear-gradient(135deg,var(--blue),var(--blue2));
    color:white;
    font-weight:800;
    min-height:44px;
    box-shadow:0 8px 20px rgba(53,104,245,.20);
}

.stDownloadButton > button:hover{
    color:white;
    transform:translateY(-1px);
}

[data-testid="stFileUploaderDropzone"]{
    border:1.5px dashed #cfd8e7 !important;
    border-radius:16px !important;
    background:rgba(255,255,255,.65) !important;
}

[data-testid="stFileUploaderDropzone"]:hover{
    border-color:var(--blue) !important;
    background:#fff !important;
}

.stTabs [data-baseweb="tab-list"]{
    gap:7px;
    background:rgba(255,255,255,.70);
    border:1px solid var(--line);
    padding:6px;
    border-radius:16px;
}

.stTabs [data-baseweb="tab"]{
    border-radius:11px;
    color:#7b8799;
    font-weight:700;
    padding:9px 16px;
}

.stTabs [aria-selected="true"]{
    background:#edf2ff !important;
    color:var(--blue) !important;
}

div[data-testid="stDataFrame"]{
    border-radius:16px;
    overflow:hidden;
    border:1px solid var(--line);
}

hr{border-color:var(--line);}

.footer{
    text-align:center;
    color:#9aa5b5;
    font-size:.72rem;
    margin-top:34px;
}

@media(max-width:1000px){
    .kpi-grid{grid-template-columns:repeat(2,1fr)}
}

@media(max-width:650px){
    .kpi-grid{grid-template-columns:1fr}
    .hero{padding:20px}
}
</style>
""",
    unsafe_allow_html=True,
)


# ============================================================
# HELPERS
# ============================================================

def money(value):
    try:
        return f"${float(value):,.2f}"
    except Exception:
        return "$0.00"

def integer(value):
    try:
        return f"{int(round(float(value))):,}"
    except Exception:
        return "0"

def normalize_text(value):
    value = str(value).strip().lower()
    value = re.sub(r"\s+", " ", value)
    return value

# ---> FUNCIÓN AÑADIDA PARA CORREGIR EL ERROR <---
def preview_table(df, view_type="general"):
    """Renderiza la tabla en pantalla con los formatos financieros de dólares"""
    format_dict = {}
    for col in ["Costo", "Capital Inmovilizado ($)", "Valor Inventario ($)"]:
        if col in df.columns:
            format_dict[col] = "${:,.2f}"
    if "Cobertura (meses)" in df.columns:
        format_dict["Cobertura (meses)"] = "{:,.1f}"
        
    st.dataframe(
        df.style.format(format_dict),
        use_container_width=True,
        hide_index=True,
        height=540
    )


def find_header_row(raw):
    keywords = [
        "código", "codigo", "cod", "artículo", "articulo",
        "descripción", "descripcion", "existencia",
        "stock", "cantidad", "ventas"
    ]
    best_row, best_score = 0, 0
    for idx, row in raw.iterrows():
        values = [normalize_text(v) for v in row.tolist() if pd.notna(v)]
        score = sum(
            1 for k in keywords
            if any(k == v or k in v for v in values)
        )
        if score > best_score:
            best_row, best_score = idx, score
    return best_row if best_score >= 1 else 0


def standardize_columns(df, kind):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    aliases = {
        "Código": ["código", "codigo", "cod", "cod_articulo", "cod. artículo", "artículo", "articulo", "sku"],
        "Descripción": ["descripción", "descripcion", "desc", "producto", "nombre"],
        "Ventas": ["cantidad", "ventas", "salidas", "venta", "unidades vendidas"],
        "Existencia": ["existencia", "stock", "saldo", "inventario", "disponible"],
        "Costo": ["costo", "precio", "costo unitario", "coste", "precio costo"],
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

    for col in ["Ventas", "Existencia", "Costo"]:
        if col in df.columns:
            if df[col].dtype == object:
                s = (
                    df[col].astype(str)
                    .str.replace("$", "", regex=False)
                    .str.replace(" ", "", regex=False)
                    .str.replace(",", "", regex=False)
                )
                df[col] = pd.to_numeric(s, errors="coerce")
            else:
                df[col] = pd.to_numeric(df[col], errors="coerce")
            df[col] = df[col].fillna(0)

    if "Código" in df.columns:
        df["Código"] = (
            df["Código"]
            .astype(str)
            .str.strip()
            .str.replace(r"\.0$", "", regex=True)
        )
        df = df[df["Código"].notna() & (df["Código"] != "") & (df["Código"].str.lower() != "nan")]

    if "Descripción" in df.columns:
        df["Descripción"] = df["Descripción"].fillna("").astype(str).str.strip()

    if "Código" in df.columns:
        agg = {}
        for col in df.columns:
            if col == "Código":
                continue
            if col in ["Ventas", "Existencia", "Costo"]:
                agg[col] = "sum"
            else:
                agg[col] = "first"
        df = df.groupby("Código", as_index=False).agg(agg)

    return df


def read_excel_smart(uploaded_file, kind):
    raw = pd.read_excel(uploaded_file, header=None, engine="openpyxl")
    header_row = find_header_row(raw)
    df = pd.read_excel(uploaded_file, skiprows=header_row, engine="openpyxl")
    return standardize_columns(df, kind), header_row


def process_site(df_sales, df_inventory, months=3, min_months=0.5, max_months=1.0):
    sales = df_sales.copy()
    inv = df_inventory.copy()

    if "Ventas" not in sales.columns:
        sales["Ventas"] = 0
    if "Existencia" not in inv.columns:
        inv["Existencia"] = 0
    if "Costo" not in inv.columns:
        inv["Costo"] = 0

    df = sales.merge(inv, on="Código", how="outer", suffixes=("_ventas", "_inv"))

    if "Descripción_ventas" in df.columns or "Descripción_inv" in df.columns:
        a = df.get("Descripción_ventas", pd.Series("", index=df.index))
        b = df.get("Descripción_inv", pd.Series("", index=df.index))
        df["Descripción"] = a.replace("", np.nan).fillna(b).fillna("")
    elif "Descripción" not in df.columns:
        df["Descripción"] = ""

    if "Ventas" not in df.columns:
        df["Ventas"] = 0
    if "Ventas_inv" in df.columns:
        df["Ventas"] = df["Ventas"].fillna(df["Ventas_inv"]).fillna(0)
    if "Existencia" not in df.columns:
        df["Existencia"] = 0
    if "Existencia_ventas" in df.columns:
        df["Existencia"] = df["Existencia"].fillna(df["Existencia_ventas"]).fillna(0)
    if "Costo" not in df.columns:
        df["Costo"] = 0

    df["Ventas"] = pd.to_numeric(df["Ventas"], errors="coerce").fillna(0)
    df["Existencia"] = pd.to_numeric(df["Existencia"], errors="coerce").fillna(0)
    df["Costo"] = pd.to_numeric(df["Costo"], errors="coerce").fillna(0)

    months = max(1, int(months))
    min_months = max(0.1, float(min_months))
    max_months = max(min_months, float(max_months))

    df["Prom. Mensual"] = np.ceil(df["Ventas"] / months).astype(int)
    df["Prom. Quincenal"] = np.ceil(df["Prom. Mensual"] / 2).astype(int)

    df["Stock Mínimo"] = np.where(
        df["Ventas"] <= 0,
        1,
        np.ceil(df["Prom. Mensual"] * min_months),
    ).astype(int)
    
    df["Stock Máximo"] = np.where(
        df["Ventas"] <= 0,
        2,
        np.ceil(df["Prom. Mensual"] * max_months),
    ).astype(int)

    df["Compra Sugerida"] = np.where(
        df["Existencia"] < df["Stock Mínimo"],
        np.maximum(df["Stock Máximo"] - df["Existencia"], 0),
        0,
    ).astype(int)

    df["Retiro Almacén"] = np.where(
        df["Existencia"] > df["Stock Máximo"],
        df["Existencia"] - df["Stock Máximo"],
        0,
    ).astype(int)

    df["Capital Inmovilizado ($)"] = (
        df["Retiro Almacén"] * df["Costo"]
    ).round(2)

    df["Valor Inventario ($)"] = (
        df["Existencia"] * df["Costo"]
    ).round(2)

    df["Cobertura (meses)"] = np.where(
        df["Prom. Mensual"] > 0,
        np.maximum(df["Existencia"], 0) / df["Prom. Mensual"],
        np.nan,
    ).round(1)

    df["Estado"] = np.select(
        [
            df["Existencia"] < df["Stock Mínimo"],
            df["Retiro Almacén"] > 0,
        ],
        [
            "CRÍTICO — COMPRAR",
            "SOBRESTOCK — RETIRAR",
        ],
        default="ÓPTIMO",
    )

    df["Prioridad"] = np.select(
        [
            (df["Compra Sugerida"] > 0) & (df["Ventas"] > 0),
            df["Compra Sugerida"] > 0,
            df["Retiro Almacén"] > 0,
        ],
        ["ALTA", "MEDIA", "MEDIA"],
        default="BAJA",
    )

    drop_cols = [
        c for c in df.columns
        if c.endswith("_ventas") or c.endswith("_inv")
    ]
    df = df.drop(columns=drop_cols, errors="ignore")

    return df.sort_values(
        by=["Prioridad", "Capital Inmovilizado ($)", "Ventas"],
        ascending=[True, False, False],
    ).reset_index(drop=True)


def make_excel_report(site_data, site_name, months, executive_note=""):
    df = site_data.copy()
    output = io.BytesIO()

    integer_cols = ["Ventas", "Existencia", "Prom. Mensual", "Prom. Quincenal", "Stock Mínimo", "Stock Máximo", "Compra Sugerida", "Retiro Almacén"]
    for col in integer_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).round().astype(int)
    if "Cobertura (meses)" in df.columns:
        df["Cobertura (meses)"] = pd.to_numeric(df["Cobertura (meses)"], errors="coerce").round(1)

    general_cols = [
        "Código", "Descripción", "Ventas", "Costo", "Existencia",
        "Prom. Quincenal", "Prom. Mensual", "Stock Mínimo", "Stock Máximo",
        "Compra Sugerida", "Retiro Almacén", "Capital Inmovilizado ($)",
        "Valor Inventario ($)", "Cobertura (meses)", "Estado", "Prioridad"
    ]
    general = df[[c for c in general_cols if c in df.columns]].copy()
    compras = df[df["Compra Sugerida"] > 0][[c for c in ["Código","Descripción","Ventas","Existencia","Stock Mínimo","Stock Máximo","Compra Sugerida","Costo","Prioridad","Estado"] if c in df.columns]].copy()
    retiros = df[df["Retiro Almacén"] > 0][[c for c in ["Código","Descripción","Ventas","Existencia","Stock Máximo","Retiro Almacén","Costo","Capital Inmovilizado ($)","Cobertura (meses)","Estado"] if c in df.columns]].copy()
    financiero = df[df["Capital Inmovilizado ($)"] > 0][[c for c in ["Código","Descripción","Costo","Existencia","Retiro Almacén","Capital Inmovilizado ($)","Valor Inventario ($)","Cobertura (meses)"] if c in df.columns]].sort_values("Capital Inmovilizado ($)", ascending=False)
    criticos = df[df["Estado"] == "CRÍTICO — COMPRAR"][general.columns].copy()

    resumen = pd.DataFrame({
        "Indicador": [
            "Sede", "Fecha de generación", "Meses de historial", "SKUs analizados",
            "Productos críticos", "Órdenes de compra", "Retiros sugeridos",
            "Valor total del inventario", "Capital inmovilizado por sobrestock",
            "Unidades a comprar", "Unidades a retirar"
        ],
        "Valor": [
            site_name, datetime.now().strftime("%d/%m/%Y %H:%M"), months, len(df),
            int((df["Estado"] == "CRÍTICO — COMPRAR").sum()), int((df["Compra Sugerida"] > 0).sum()),
            int((df["Retiro Almacén"] > 0).sum()), round(float(df["Valor Inventario ($)"].sum()),2),
            round(float(df["Capital Inmovilizado ($)"].sum()),2), int(df["Compra Sugerida"].sum()),
            int(df["Retiro Almacén"].sum())
        ]
    })

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        resumen.to_excel(writer, sheet_name="00 Ejecutivo", index=False)
        general.to_excel(writer, sheet_name="01 Inventario General", index=False)
        compras.to_excel(writer, sheet_name="02 Plan de Compras", index=False)
        retiros.to_excel(writer, sheet_name="03 Plan de Retiros", index=False)
        financiero.to_excel(writer, sheet_name="04 Control Financiero", index=False)
        criticos.to_excel(writer, sheet_name="05 Alertas Críticas", index=False)

        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
        navy = "18243A"; blue = "3568F5"; white = "FFFFFF"; line = "DCE4F0"; green="E9F8F0"; red="FDEEEE"; orange="FFF4E2"
        thin = Side(style="thin", color=line)
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            ws.sheet_view.showGridLines = False
            for cell in ws[1]:
                cell.font = Font(bold=True, color=white, size=10)
                cell.fill = PatternFill("solid", fgColor=blue)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = Border(bottom=thin)
            ws.row_dimensions[1].height = 28
            for row in ws.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(vertical="center")
                    cell.border = Border(bottom=Side(style="hair", color=line))
            for idx, col in enumerate(ws.columns, 1):
                max_len = max([len(str(c.value or "")) for c in list(col)[:160]], default=10)
                width = min(max(max_len + 2, 11), 42)
                if idx == 2 and ws.title != "00 Ejecutivo":
                    width = 44
                ws.column_dimensions[get_column_letter(idx)].width = width

            headers_row = [c.value for c in ws[1]]
            for idx, name in enumerate(headers_row, 1):
                col_letter = get_column_letter(idx)
                if name in integer_cols or name in ["Ventas", "Existencia", "Compra Sugerida", "Retiro Almacén", "Prom. Mensual", "Prom. Quincenal", "Stock Mínimo", "Stock Máximo"]:
                    for cell in ws[col_letter][1:]: cell.number_format = "#,##0"
                elif name in ["Costo", "Capital Inmovilizado ($)", "Valor Inventario ($)", "Valor total del inventario", "Capital inmovilizado por sobrestock"]:
                    for cell in ws[col_letter][1:]: cell.number_format = '$#,##0.00'
                elif name == "Cobertura (meses)":
                    for cell in ws[col_letter][1:]: cell.number_format = '0.0'

    return output.getvalue()

def combine_sites(site_results):
    frames = []
    for site, df in site_results.items():
        if df is None or df.empty:
            continue
        x = df.copy()
        x.insert(0, "Sede", site)
        frames.append(x)
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


def kpi_card(label, value, note="", tone="blue"):
    st.markdown(
        f"""
        <div class="kpi {tone}">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def chart_layout(fig, height=320):
    fig.update_layout(
        height=height,
        margin=dict(l=10, r=10, t=35, b=10),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="DM Sans", color="#5d6a7e"),
        legend=dict(orientation="h", y=1.02, x=0),
    )
    return fig


# ============================================================
# ESTADO
# ============================================================

if "site_results" not in st.session_state:
    st.session_state.site_results = {}

if "active_view" not in st.session_state:
    st.session_state.active_view = "Centro de Control"


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:
    st.markdown(
        """
        <div class="sidebar-logo">
            ◈ MAKROPETROL <span style="color:#3568f5">NEXUS</span>
            <small>Control logístico inteligente</small>
        </div>
        """,
        unsafe_allow_html=True,
    )

    view = st.radio(
        "MÓDULOS",
        [
            "Centro de Control",
            "Carga de Datos",
            "Compras",
            "Retiros",
            "Inventario",
            "Finanzas",
        ],
        label_visibility="visible",
    )
    st.session_state.active_view = view

    st.divider()
    st.markdown("**Parámetros del motor**")

    months = st.number_input(
        "Meses representados por ventas",
        min_value=1,
        max_value=24,
        value=3,
        step=1,
    )

    min_months = st.number_input(
        "Stock mínimo — cobertura",
        min_value=0.1,
        max_value=6.0,
        value=0.5,
        step=0.5,
    )

    max_months = st.number_input(
        "Stock máximo — cobertura",
        min_value=0.5,
        max_value=12.0,
        value=1.0,
        step=0.5,
    )

    st.divider()
    loaded = len(st.session_state.site_results)
    st.markdown(
        f'<span class="badge badge-green">● {loaded}/5 sedes procesadas</span>',
        unsafe_allow_html=True,
    )

    if loaded:
        if st.button("↺ Limpiar todos los resultados", use_container_width=True):
            st.session_state.site_results = {}
            st.rerun()


# ============================================================
# HEADER
# ============================================================

st.markdown(
    f"""
    <div class="hero">
        <div class="brand">Makropetrol <span>NEXUS</span></div>
        <div class="hero-sub">
            Centro de inteligencia logística · inventario · abastecimiento · redistribución · control financiero
        </div>
        <div class="status-row">
            <span class="dot"></span>
            Motor operativo disponible
            <span>•</span>
            Última sesión: {datetime.now().strftime("%d/%m/%Y %H:%M")}
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# CENTRO DE CONTROL
# ============================================================

all_data = combine_sites(st.session_state.site_results)

if view == "Centro de Control":
    st.markdown('<div class="section-title">Resumen ejecutivo</div>', unsafe_allow_html=True)

    if all_data.empty:
        st.info("Aún no hay sedes procesadas. Ve a **Carga de Datos**, sube Inventario + Ventas y procesa la sede.")
    else:
        total_skus = len(all_data)
        critical = int((all_data["Compra Sugerida"] > 0).sum())
        withdrawals = int((all_data["Retiro Almacén"] > 0).sum())
        inventory_value = float(all_data["Valor Inventario ($)"].sum())
        immobilized = float(all_data["Capital Inmovilizado ($)"].sum())

        st.markdown('<div class="kpi-grid">', unsafe_allow_html=True)
        c1, c2, c3, c4 = st.columns(4)
        with c1: kpi_card("SKUs analizados", integer(total_skus), "Todas las sedes", "blue")
        with c2: kpi_card("Compras sugeridas", integer(critical), "Productos bajo mínimo", "orange")
        with c3: kpi_card("Retiros sugeridos", integer(withdrawals), "Productos sobre máximo", "green")
        with c4: kpi_card("Capital inmovilizado", money(immobilized), "Valor del sobrestock", "red")
        st.markdown("</div>", unsafe_allow_html=True)

        left, right = st.columns([1.25, 1])

        with left:
            st.markdown('<div class="section-title">Situación por sede</div>', unsafe_allow_html=True)
            summary = (
                all_data.groupby("Sede")
                .agg(
                    SKUs=("Código", "count"),
                    Compras=("Compra Sugerida", lambda x: int((x > 0).sum())),
                    Retiros=("Retiro Almacén", lambda x: int((x > 0).sum())),
                    Inventario=(
                        "Valor Inventario ($)",
                        "sum",
                    ),
                    Inmovilizado=(
                        "Capital Inmovilizado ($)",
                        "sum",
                    ),
                )
                .reset_index()
            )
            st.dataframe(
                summary.style.format(
                    {
                        "Inventario": "${:,.2f}",
                        "Inmovilizado": "${:,.2f}",
                    }
                ),
                use_container_width=True,
                hide_index=True,
            )

        with right:
            st.markdown('<div class="section-title">Distribución de acciones</div>', unsafe_allow_html=True)
            actions = pd.DataFrame({
                "Acción": ["Óptimo", "Comprar", "Retirar"],
                "Cantidad": [
                    int((all_data["Estado"] == "ÓPTIMO").sum()),
                    int((all_data["Compra Sugerida"] > 0).sum()),
                    int((all_data["Retiro Almacén"] > 0).sum()),
                ],
            })
            fig = px.pie(actions, names="Acción", values="Cantidad", hole=.68)
            fig.update_traces(textposition="inside", textinfo="percent+label")
            st.plotly_chart(chart_layout(fig, 300), use_container_width=True)

        st.markdown('<div class="section-title">Top 15 productos con mayor capital inmovilizado</div>', unsafe_allow_html=True)
        top = (
            all_data[all_data["Capital Inmovilizado ($)"] > 0]
            .sort_values("Capital Inmovilizado ($)", ascending=False)
            .head(15)
        )
        if top.empty:
            st.success("No hay capital inmovilizado por sobrestock con los parámetros actuales.")
        else:
            fig = px.bar(
                top.sort_values("Capital Inmovilizado ($)"),
                x="Capital Inmovilizado ($)",
                y="Descripción",
                orientation="h",
                hover_data=["Sede", "Código", "Existencia", "Retiro Almacén"],
            )
            st.plotly_chart(chart_layout(fig, 470), use_container_width=True)


# ============================================================
# CARGA DE DATOS
# ============================================================

elif view == "Carga de Datos":
    st.markdown('<div class="section-title">Procesamiento multisede</div>', unsafe_allow_html=True)
    st.caption("Sube los dos libros de cada sede. El sistema detecta automáticamente la fila de encabezados, normaliza columnas y genera el plan de acción.")

    selected_sites = st.multiselect(
        "Sedes a procesar",
        SEDES,
        default=[s for s in SEDES if s in st.session_state.site_results] or ["Centurión"],
    )

    for site in selected_sites:
        with st.expander(f"{SEDE_ICON[site]}  {site}", expanded=True):
            a, b = st.columns(2)

            with a:
                inventory_file = st.file_uploader(
                    "Inventario + costos",
                    type=["xlsx"],
                    key=f"inv_{site}",
                    help="Libro con Código, Existencia y, si está disponible, Costo.",
                )

            with b:
                sales_file = st.file_uploader(
                    "Reporte de ventas",
                    type=["xlsx"],
                    key=f"sales_{site}",
                    help="Libro con Código y Cantidad/Ventas.",
                )

            if inventory_file and sales_file:
                if st.button(f"⚡ Procesar {site}", key=f"process_{site}", use_container_width=True):
                    try:
                        inv, inv_header = read_excel_smart(inventory_file, "inventario")
                        sales, sales_header = read_excel_smart(sales_file, "ventas")

                        missing_i = [c for c in ["Código", "Existencia"] if c not in inv.columns]
                        missing_v = [c for c in ["Código", "Ventas"] if c not in sales.columns]

                        if missing_i or missing_v:
                            if missing_i:
                                st.error(f"Inventario: faltan columnas detectables: {', '.join(missing_i)}")
                            if missing_v:
                                st.error(f"Ventas: faltan columnas detectables: {', '.join(missing_v)}")
                        else:
                            result = process_site(
                                sales,
                                inv,
                                months=months,
                                min_months=min_months,
                                max_months=max_months,
                            )
                            st.session_state.site_results[site] = result
                            st.success(
                                f"{site} procesada · {len(result):,} SKUs · "
                                f"{int((result['Compra Sugerida'] > 0).sum()):,} compras · "
                                f"{int((result['Retiro Almacén'] > 0).sum()):,} retiros."
                            )
                    except Exception as e:
                        st.error(f"No se pudo procesar {site}: {e}")

            if site in st.session_state.site_results:
                df = st.session_state.site_results[site]
                c1, c2, c3 = st.columns(3)
                with c1: kpi_card("SKUs", integer(len(df)), site, "blue")
                with c2: kpi_card("Comprar", integer((df["Compra Sugerida"] > 0).sum()), "Bajo mínimo", "orange")
                with c3: kpi_card("Retirar", integer((df["Retiro Almacén"] > 0).sum()), "Sobre máximo", "green")

                excel = make_excel_report(
                    df,
                    site,
                    months,
                    "Reporte generado por Makropetrol NEXUS.",
                )
                st.download_button(
                    "⬇ Descargar reporte ejecutivo Excel",
                    data=excel,
                    file_name=f"Makropetrol_Nexus_{site.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"download_{site}",
                )


# ============================================================
# COMPRAS
# ============================================================

elif view == "Compras":
    st.markdown('<div class="section-title">Centro de abastecimiento</div>', unsafe_allow_html=True)

    if all_data.empty:
        st.info("Procesa al menos una sede para generar órdenes de compra.")
    else:
        compras = all_data[all_data["Compra Sugerida"] > 0].copy()
        if compras.empty:
            st.success("No existen compras sugeridas con los parámetros actuales.")
        else:
            c1, c2, c3 = st.columns(3)
            with c1: kpi_card("SKUs a comprar", integer(len(compras)), "Productos bajo mínimo", "orange")
            with c2: kpi_card("Unidades", integer(compras["Compra Sugerida"].sum()), "Cantidad sugerida", "blue")
            with c3: kpi_card("Costo estimado", money((compras["Compra Sugerida"] * compras["Costo"]).sum()), "Si el costo está cargado", "green")

            f1, f2 = st.columns(2)
            with f1:
                site_filter = st.multiselect("Filtrar por sede", sorted(compras["Sede"].unique()), default=sorted(compras["Sede"].unique()))
            with f2:
                priority = st.multiselect("Prioridad", sorted(compras["Prioridad"].unique()), default=sorted(compras["Prioridad"].unique()))

            view_df = compras[
                compras["Sede"].isin(site_filter) &
                compras["Prioridad"].isin(priority)
            ].copy()

            cols = [
                "Sede", "Código", "Descripción", "Ventas", "Existencia",
                "Stock Mínimo", "Stock Máximo", "Compra Sugerida",
                "Costo", "Prioridad", "Estado"
            ]
            
            # -> FIX APLICADO AQUÍ: SE PASAN SOLO LAS COLUMNAS NECESARIAS
            preview_table(view_df[cols], "compras")

            export = io.BytesIO()
            view_df.to_excel(export, index=False, sheet_name="Ordenes de Compra")
            st.download_button(
                "⬇ Exportar lista de compras",
                data=export.getvalue(),
                file_name="Ordenes_Compra_Makropetrol.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


# ============================================================
# RETIROS
# ============================================================

elif view == "Retiros":
    st.markdown('<div class="section-title">Redistribución y retiros de almacén</div>', unsafe_allow_html=True)

    if all_data.empty:
        st.info("Procesa al menos una sede para generar el plan de retiros.")
    else:
        retiros = all_data[all_data["Retiro Almacén"] > 0].copy()
        if retiros.empty:
            st.success("No existen retiros sugeridos con los parámetros actuales.")
        else:
            c1, c2, c3 = st.columns(3)
            with c1: kpi_card("SKUs con sobrestock", integer(len(retiros)), "Para redistribución", "green")
            with c2: kpi_card("Unidades a retirar", integer(retiros["Retiro Almacén"].sum()), "Exceso sobre máximo", "orange")
            with c3: kpi_card("Capital liberable", money(retiros["Capital Inmovilizado ($)"].sum()), "Valor del exceso", "red")

            site_filter = st.multiselect(
                "Sedes",
                sorted(retiros["Sede"].unique()),
                default=sorted(retiros["Sede"].unique()),
            )
            data = retiros[retiros["Sede"].isin(site_filter)].copy()

            cols = [
                "Sede", "Código", "Descripción", "Ventas", "Existencia",
                "Stock Máximo", "Retiro Almacén", "Costo",
                "Capital Inmovilizado ($)", "Cobertura (meses)"
            ]
            
            # -> FIX APLICADO AQUÍ: SE PASAN SOLO LAS COLUMNAS NECESARIAS
            preview_table(data[cols], "retiros")

            export = io.BytesIO()
            data.to_excel(export, index=False, sheet_name="Retiros")
            st.download_button(
                "⬇ Exportar plan de retiros",
                data=export.getvalue(),
                file_name="Plan_Retiros_Makropetrol.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


# ============================================================
# INVENTARIO
# ============================================================

elif view == "Inventario":
    st.markdown('<div class="section-title">Explorador de inventario</div>', unsafe_allow_html=True)

    if all_data.empty:
        st.info("Procesa al menos una sede para explorar el inventario.")
    else:
        site_filter = st.selectbox("Sede", ["Todas"] + sorted(all_data["Sede"].unique()))
        data = all_data.copy() if site_filter == "Todas" else all_data[all_data["Sede"] == site_filter].copy()

        search = st.text_input("Buscar por código o descripción", placeholder="Ej. 40011 / SHELL / LIQUI MOLY")
        if search:
            q = normalize_text(search)
            data = data[
                data["Código"].astype(str).str.lower().str.contains(q, na=False) |
                data["Descripción"].astype(str).str.lower().str.contains(q, na=False)
            ]

        state = st.multiselect(
            "Estado",
            sorted(data["Estado"].unique()),
            default=sorted(data["Estado"].unique()),
        )
        data = data[data["Estado"].isin(state)]

        cols = [
            "Sede", "Código", "Descripción", "Costo", "Ventas",
            "Existencia", "Prom. Mensual", "Stock Mínimo",
            "Stock Máximo", "Compra Sugerida", "Retiro Almacén",
            "Valor Inventario ($)", "Estado"
        ]
        
        # -> FIX APLICADO AQUÍ
        preview_table(data[cols], "inventario")


# ============================================================
# FINANZAS
# ============================================================

elif view == "Finanzas":
    st.markdown('<div class="section-title">Inteligencia financiera del inventario</div>', unsafe_allow_html=True)

    if all_data.empty:
        st.info("Procesa al menos una sede para visualizar los indicadores financieros.")
    else:
        inventory_value = float(all_data["Valor Inventario ($)"].sum())
        immobilized = float(all_data["Capital Inmovilizado ($)"].sum())
        purchase_value = float((all_data["Compra Sugerida"] * all_data["Costo"]).sum())
        zero_cost = int((all_data["Costo"] <= 0).sum())

        c1, c2, c3, c4 = st.columns(4)
        with c1: kpi_card("Valor inventario", money(inventory_value), "Existencia × costo", "blue")
        with c2: kpi_card("Capital inmovilizado", money(immobilized), "Sobrestock", "red")
        with c3: kpi_card("Compra estimada", money(purchase_value), "Unidades sugeridas × costo", "orange")
        with c4: kpi_card("SKUs sin costo", integer(zero_cost), "Revisar fuente", "green")

        a, b = st.columns(2)

        with a:
            by_site = all_data.groupby("Sede")["Valor Inventario ($)"].sum().reset_index()
            fig = px.bar(by_site, x="Sede", y="Valor Inventario ($)", title="Valor del inventario por sede")
            st.plotly_chart(chart_layout(fig, 350), use_container_width=True)

        with b:
            by_site = all_data.groupby("Sede")["Capital Inmovilizado ($)"].sum().reset_index()
            fig = px.bar(by_site, x="Sede", y="Capital Inmovilizado ($)", title="Capital inmovilizado por sede")
            st.plotly_chart(chart_layout(fig, 350), use_container_width=True)

        st.markdown('<div class="section-title">Mayor capital inmovilizado</div>', unsafe_allow_html=True)
        financial = (
            all_data[all_data["Capital Inmovilizado ($)"] > 0]
            .sort_values("Capital Inmovilizado ($)", ascending=False)
            .head(50)
        )
        
        cols = [
            "Sede", "Código", "Descripción", "Costo",
            "Existencia", "Stock Máximo", "Retiro Almacén",
            "Capital Inmovilizado ($)"
        ]
        
        # -> FIX APLICADO AQUÍ
        preview_table(financial[cols], "finanzas")


# ============================================================
# FOOTER
# ============================================================

st.markdown(
    """
    <div class="footer">
        Makropetrol NEXUS · Plataforma de control logístico multisede · v2.1
        <br>
        Diseñada para convertir reportes operativos en decisiones accionables.
    </div>
    """,
    unsafe_allow_html=True,
)