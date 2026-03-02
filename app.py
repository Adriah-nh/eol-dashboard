"""
Dashboard de Gestión Logística - EOL Perú
Desarrollado con Streamlit + Pandas + Plotly
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
import base64
from datetime import datetime, date
import warnings
import os

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────
# CONFIGURACIÓN GENERAL
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="EOL Perú — Dashboard Logístico",
    page_icon="🚛",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Forzar sidebar siempre visible y botón colapsar oculto
st.markdown("""
<style>
  [data-testid="collapsedControl"] { display: none !important; }
  section[data-testid="stSidebar"] {
      display: block !important;
      min-width: 270px !important;
      width: 270px !important;
  }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# PALETA DE COLORES EOL PERÚ
# ─────────────────────────────────────────────
COL_AMARILLO  = "#FFC107"
COL_NEGRO     = "#1A1A1A"
COL_AZUL      = "#1A3A5C"
COL_FONDO     = "#F2F2F2"
COL_BLANCO    = "#FFFFFF"
COL_ROJO      = "#E53935"
COL_VERDE     = "#43A047"
COL_GRIS      = "#9E9E9E"

# ─────────────────────────────────────────────
# CSS GLOBAL
# ─────────────────────────────────────────────
st.markdown(f"""
<style>
  /* Fondo general */
  .stApp {{ background-color: {COL_FONDO}; }}

  /* Sidebar */
  [data-testid="stSidebar"] {{
    background-color: {COL_NEGRO} !important;
  }}
  [data-testid="stSidebar"] * {{
    color: {COL_BLANCO} !important;
  }}
  [data-testid="stSidebar"] .stSelectbox label,
  [data-testid="stSidebar"] .stMultiSelect label,
  [data-testid="stSidebar"] .stDateInput label {{
    color: {COL_AMARILLO} !important;
    font-weight: 600;
  }}

  /* Métricas KPI */
  .kpi-card {{
    background: {COL_BLANCO};
    border-radius: 12px;
    padding: 16px 20px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    border-left: 5px solid {COL_AMARILLO};
    margin-bottom: 8px;
  }}
  .kpi-title {{ font-size: 0.78rem; color: {COL_GRIS}; font-weight: 600; text-transform: uppercase; letter-spacing: 0.05em; }}
  .kpi-value {{ font-size: 1.6rem; font-weight: 700; color: {COL_AZUL}; margin: 2px 0; }}
  .kpi-sub   {{ font-size: 0.75rem; color: {COL_GRIS}; }}

  /* Alertas */
  .kpi-card.alerta {{ border-left-color: {COL_ROJO}; }}
  .kpi-card.ok     {{ border-left-color: {COL_VERDE}; }}

  /* Headers de sección */
  .section-header {{
    background: linear-gradient(90deg, {COL_AZUL}, {COL_NEGRO});
    color: {COL_AMARILLO};
    padding: 10px 20px;
    border-radius: 8px;
    font-size: 1.1rem;
    font-weight: 700;
    margin: 20px 0 12px 0;
    letter-spacing: 0.03em;
  }}

  /* Tabs */
  .stTabs [data-baseweb="tab-list"] {{ background-color: {COL_BLANCO}; border-radius: 8px; }}
  .stTabs [data-baseweb="tab"] {{ color: {COL_AZUL}; font-weight: 600; }}
  .stTabs [aria-selected="true"] {{ background-color: {COL_AMARILLO} !important; color: {COL_NEGRO} !important; border-radius: 6px; }}

  /* Tablas */
  .dataframe {{ font-size: 0.82rem; }}
  thead tr th {{ background-color: {COL_AZUL} !important; color: {COL_BLANCO} !important; }}

  /* Botón descarga */
  .stDownloadButton > button {{
    background-color: {COL_AMARILLO};
    color: {COL_NEGRO};
    font-weight: 700;
    border: none;
    border-radius: 8px;
    padding: 8px 20px;
  }}
  .stDownloadButton > button:hover {{ background-color: #e6ac00; }}

  /* Upload */
  [data-testid="stFileUploader"] {{
    background: {COL_BLANCO};
    border: 2px dashed {COL_AMARILLO};
    border-radius: 10px;
    padding: 10px;
  }}

  h2, h3 {{ color: {COL_AZUL}; }}
  .stAlert {{ border-radius: 8px; }}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# LOGO EOL PERÚ
# ─────────────────────────────────────────────
def get_logo_b64():
    """Intenta cargar el logo desde assets/logo_eol.png"""
    logo_paths = ["assets/logo_eol.png", "logo_eol.png"]
    for p in logo_paths:
        if os.path.exists(p):
            with open(p, "rb") as f:
                return base64.b64encode(f.read()).decode()
    return None


# ─────────────────────────────────────────────
# UTILIDADES
# ─────────────────────────────────────────────
def fmt_soles(v):
    if pd.isna(v): return "S/ 0.00"
    return f"S/ {v:,.2f}"

def fmt_num(v, dec=0):
    if pd.isna(v): return "0"
    return f"{v:,.{dec}f}"

def kpi_html(title, value, sub="", alert=False, ok=False):
    cls = "alerta" if alert else ("ok" if ok else "")
    return f"""
    <div class="kpi-card {cls}">
      <div class="kpi-title">{title}</div>
      <div class="kpi-value">{value}</div>
      <div class="kpi-sub">{sub}</div>
    </div>"""

def to_excel_bytes(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False)
    return buf.getvalue()

def plotly_layout(fig, title="", height=400):
    fig.update_layout(
        title=dict(text=title, font=dict(color=COL_AZUL, size=14, family="Arial")),
        paper_bgcolor=COL_BLANCO,
        plot_bgcolor=COL_FONDO,
        height=height,
        font=dict(family="Arial", color=COL_NEGRO),
        legend=dict(bgcolor=COL_BLANCO, bordercolor=COL_GRIS, borderwidth=1),
        margin=dict(l=40, r=20, t=50, b=40),
    )
    fig.update_xaxes(gridcolor="#E0E0E0", linecolor="#BDBDBD")
    fig.update_yaxes(gridcolor="#E0E0E0", linecolor="#BDBDBD")
    return fig

COLORS_PLACAS = [
    COL_AZUL, COL_AMARILLO, "#26A69A", "#EF5350", "#AB47BC",
    "#FF7043", "#42A5F5", "#66BB6A", "#FFA726", "#8D6E63",
    "#EC407A", "#7E57C2", "#26C6DA", "#D4E157"
]


# ─────────────────────────────────────────────
# CARGA Y PARSEO DE DATOS
# ─────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_data(file_bytes: bytes, filename: str):
    """Carga y parsea todos los archivos Excel desde el ZIP o archivo individual."""
    import zipfile, io as _io

    data = {}

    def clean_cols(df):
        df.columns = [str(c).strip() for c in df.columns]
        return df

    def read_excel_safe(xl, sheet, **kw):
        try:
            df = xl.parse(sheet, **kw)
            return clean_cols(df)
        except Exception:
            return pd.DataFrame()

    # ── Determina si es ZIP o Excel directo
    if filename.endswith(".zip"):
        zf = zipfile.ZipFile(_io.BytesIO(file_bytes))
        excel_files = {n: zf.read(n) for n in zf.namelist() if n.endswith(".xlsx")}
    else:
        excel_files = {filename: file_bytes}

    for fname, fbytes in excel_files.items():
        xl = pd.ExcelFile(_io.BytesIO(fbytes))
        bname = os.path.basename(fname)

        # ── DATA FEBRERO AJE
        if "DATA" in bname.upper():
            # 1. Combustible
            raw = xl.parse("1. Combustible", header=None)
            header_row = raw[raw.apply(lambda r: r.astype(str).str.contains("Placa").any(), axis=1)].index[0]
            # Tabla izquierda (cargas individuales)
            df_comb_raw = raw.iloc[header_row+1:, :7].copy()
            df_comb_raw.columns = ["Placa","Kilometraje","Fecha","Combustible","Galones","Precio","Costo"]
            df_comb_raw = df_comb_raw.dropna(subset=["Placa"]).copy()
            df_comb_raw["Fecha"] = pd.to_datetime(df_comb_raw["Fecha"], errors="coerce")
            for c in ["Kilometraje","Galones","Precio","Costo"]:
                df_comb_raw[c] = pd.to_numeric(df_comb_raw[c], errors="coerce").fillna(0)
            data["combustible_raw"] = df_comb_raw

            # Tabla derecha (resumen por placa)
            df_comb = raw.iloc[header_row+1:, 8:20].copy()
            df_comb.columns = ["Placa","N_abastecimientos","Tipo","Galones","Gasto",
                               "Kilometraje","N_viajes","Km_por_galon",
                               "Gasto_por_km","Gasto_por_viaje","Objetivo","Merma"]
            df_comb = df_comb.dropna(subset=["Placa"]).copy()
            for c in ["N_abastecimientos","Galones","Gasto","Kilometraje","N_viajes",
                      "Km_por_galon","Gasto_por_km","Gasto_por_viaje","Objetivo","Merma"]:
                df_comb[c] = pd.to_numeric(df_comb[c], errors="coerce").fillna(0)
            data["combustible"] = df_comb

            # 2. Margen
            raw2 = xl.parse("2. Margen Febrero", header=None)
            hr2 = raw2[raw2.apply(lambda r: r.astype(str).str.contains("Fecha").any(), axis=1)].index[0]
            df_margen = raw2.iloc[hr2+1:, :16].copy()
            df_margen.columns = ["Fecha","Cliente","Placa","Ingresos","Conductor","Aux1","Aux2",
                                  "AlquilerCamion","Combustible","CostoCochera","CostoPeraje",
                                  "LavadoUnidades","OtroGasto","TotalGastos","MargenCamion","PctMC"]
            df_margen = df_margen.dropna(subset=["Fecha"]).copy()
            df_margen["Fecha"] = pd.to_datetime(df_margen["Fecha"], errors="coerce")
            df_margen = df_margen.dropna(subset=["Fecha"])
            for c in ["Ingresos","Conductor","Aux1","Aux2","AlquilerCamion","Combustible",
                      "CostoCochera","CostoPeraje","LavadoUnidades","OtroGasto",
                      "TotalGastos","MargenCamion","PctMC"]:
                df_margen[c] = pd.to_numeric(df_margen[c], errors="coerce").fillna(0)
            # Margen neto = Ingresos + Conductor + Aux1 + Aux2 + Combustible (gastos son negativos)
            df_margen["MargenNeto"] = df_margen["Ingresos"] + df_margen["Conductor"] + \
                                      df_margen["Aux1"] + df_margen["Aux2"] + df_margen["Combustible"]
            df_margen["PctMC_real"] = np.where(df_margen["Ingresos"] != 0,
                                                df_margen["MargenCamion"] / df_margen["Ingresos"], 0)
            data["margen"] = df_margen

            # 2.1 Margen por camión
            raw21 = xl.parse("2.1. Margen por camión", header=None)
            hr21 = raw21[raw21.apply(lambda r: r.astype(str).str.contains("Fecha").any(), axis=1)].index[0]
            df_mc = raw21.iloc[hr21+1:, :5].copy()
            df_mc.columns = ["Fecha","Placa","Ingresos","MargenDia","PctMC"]
            df_mc = df_mc.dropna(subset=["Fecha","Placa"]).copy()
            df_mc["Fecha"] = pd.to_datetime(df_mc["Fecha"], errors="coerce")
            df_mc = df_mc.dropna(subset=["Fecha"])
            for c in ["Ingresos","MargenDia","PctMC"]:
                df_mc[c] = pd.to_numeric(df_mc[c], errors="coerce").fillna(0)
            data["margen_camion"] = df_mc

            # 3. Kilometraje
            raw3 = xl.parse("3. Kilometraje Febrero", header=None)
            # Tabla kilometraje (columnas 1-13)
            placas_km = list(raw3.iloc[1, 2:14].values)
            df_km_rows = raw3.iloc[2:, 1:14].copy()
            df_km_rows.columns = ["Fecha"] + [str(p) for p in placas_km]
            df_km_rows = df_km_rows.dropna(subset=["Fecha"]).copy()
            df_km_rows["Fecha"] = pd.to_datetime(df_km_rows["Fecha"], errors="coerce")
            df_km_rows = df_km_rows.dropna(subset=["Fecha"])
            for p in [str(x) for x in placas_km]:
                df_km_rows[p] = pd.to_numeric(df_km_rows[p], errors="coerce").fillna(0)
            # Melt a formato largo
            df_km = df_km_rows.melt(id_vars="Fecha", var_name="Placa", value_name="Kilometraje")
            df_km = df_km[df_km["Kilometraje"] > 0].copy()
            df_km["Exceso70"] = df_km["Kilometraje"] > 70
            data["kilometraje"] = df_km

            # Tabla excesos (columnas 15-27)
            placas_exc = list(raw3.iloc[1, 16:28].values)
            df_exc_rows = raw3.iloc[2:, 15:28].copy()
            df_exc_rows.columns = ["Fecha"] + [str(p) for p in placas_exc]
            df_exc_rows = df_exc_rows.dropna(subset=["Fecha"]).copy()
            df_exc_rows["Fecha"] = pd.to_datetime(df_exc_rows["Fecha"], errors="coerce")
            df_exc_rows = df_exc_rows.dropna(subset=["Fecha"])
            df_exc = df_exc_rows.melt(id_vars="Fecha", var_name="Placa", value_name="ValorExceso")
            df_exc = df_exc[~df_exc["ValorExceso"].isin(["No salió","Bien",np.nan])].copy()
            df_exc["ValorExceso"] = pd.to_numeric(df_exc["ValorExceso"], errors="coerce")
            df_exc = df_exc.dropna(subset=["ValorExceso"])
            df_exc["EsExceso"] = df_exc["ValorExceso"] > 0
            data["excesos"] = df_exc

            # 4. Cocheras
            raw4 = xl.parse("4. Manejo de cocheras", header=None)
            placas_coch = list(raw4.iloc[2, 2:10].values)
            df_coch_rows = raw4.iloc[3:, 1:10].copy()
            df_coch_rows.columns = ["Fecha"] + [str(p) for p in placas_coch]
            df_coch_rows = df_coch_rows.dropna(subset=["Fecha"]).copy()
            df_coch_rows["Fecha"] = pd.to_datetime(df_coch_rows["Fecha"], errors="coerce")
            df_coch_rows = df_coch_rows.dropna(subset=["Fecha"])
            df_coch = df_coch_rows.melt(id_vars="Fecha", var_name="Placa", value_name="Cochera")
            df_coch = df_coch.dropna(subset=["Cochera"])
            df_coch = df_coch[df_coch["Cochera"].str.strip() != ""]
            data["cocheras"] = df_coch

            # 5. Peso
            raw5 = xl.parse("5. Peso Análisis EOL", header=None)
            placas_p = list(raw5.iloc[2, 3:].values)
            capacidades = list(raw5.iloc[3, 3:].values)
            cap_max = list(raw5.iloc[4, 3:].values)
            cap_dict = {str(p): pd.to_numeric(c, errors="coerce") for p, c in zip(placas_p, capacidades) if pd.notna(p)}
            capmax_dict = {str(p): pd.to_numeric(c, errors="coerce") for p, c in zip(placas_p, cap_max) if pd.notna(p)}

            peso_rows = []
            i = 5
            raw5_vals = raw5.values
            while i < len(raw5_vals):
                fecha_val = raw5_vals[i, 1]
                if pd.isna(fecha_val):
                    i += 1
                    continue
                try:
                    fecha = pd.to_datetime(fecha_val)
                except:
                    i += 1
                    continue
                # Vuelta 1
                vuelta1 = raw5_vals[i, 2]
                if str(vuelta1).strip() == "Vuelta 1":
                    pesos_v1 = raw5_vals[i, 3:3+len(placas_p)]
                else:
                    pesos_v1 = [np.nan] * len(placas_p)
                # Vuelta 2
                if i+1 < len(raw5_vals):
                    vuelta2 = raw5_vals[i+1, 2]
                    if str(vuelta2).strip() == "Vuelta 2":
                        pesos_v2 = raw5_vals[i+1, 3:3+len(placas_p)]
                    else:
                        pesos_v2 = [np.nan] * len(placas_p)
                else:
                    pesos_v2 = [np.nan] * len(placas_p)

                for j, placa in enumerate(placas_p):
                    if pd.isna(placa): continue
                    p = str(placa)
                    cap = cap_dict.get(p, np.nan)
                    cmx = capmax_dict.get(p, np.nan)
                    v1 = pd.to_numeric(pesos_v1[j] if j < len(pesos_v1) else np.nan, errors="coerce")
                    v2 = pd.to_numeric(pesos_v2[j] if j < len(pesos_v2) else np.nan, errors="coerce")
                    if pd.notna(v1) and v1 > 0:
                        peso_rows.append({"Fecha": fecha, "Placa": p, "Vuelta": "Vuelta 1",
                                          "Peso_kg": v1, "Capacidad": cap, "CapacidadMax": cmx,
                                          "SobreCapacidad": v1 > cmx if pd.notna(cmx) else False})
                    if pd.notna(v2) and v2 > 0:
                        peso_rows.append({"Fecha": fecha, "Placa": p, "Vuelta": "Vuelta 2",
                                          "Peso_kg": v2, "Capacidad": cap, "CapacidadMax": cmx,
                                          "SobreCapacidad": v2 > cmx if pd.notna(cmx) else False})
                i += 2

            data["peso"] = pd.DataFrame(peso_rows)

        # ── PLANILLA
        elif "lanilla" in bname or "LANILLA" in bname:
            df_plan = xl.parse("Tabla de datos")
            df_plan.columns = [str(c).strip() for c in df_plan.columns]
            df_plan["Fecha"] = pd.to_datetime(df_plan["Fecha"], errors="coerce")
            df_plan = df_plan.dropna(subset=["Fecha"])
            for c in ["Pago","Descuento","Pago final"]:
                df_plan[c] = pd.to_numeric(df_plan[c], errors="coerce").fillna(0)
            df_plan["Placa"] = df_plan["Placa"].astype(str).str.strip()
            data["planilla"] = df_plan

        # ── VIAJES
        elif "Viajes" in bname or "viajes" in bname:
            viajes_list = []
            for sheet in xl.sheet_names:
                if sheet.startswith("Viajes-"):
                    placa = sheet.replace("Viajes-", "")
                    df_v = xl.parse(sheet)
                    df_v.columns = [str(c).strip() for c in df_v.columns]
                    df_v["Placa"] = placa
                    df_v["Comienzo"] = pd.to_datetime(df_v["Comienzo"], errors="coerce")
                    df_v["Fin"] = pd.to_datetime(df_v["Fin"], errors="coerce")
                    df_v["Fecha"] = df_v["Comienzo"].dt.normalize()
                    for c in ["Kilometraje","Velocidad media","Velocidad máxima"]:
                        if c in df_v.columns:
                            df_v[c] = pd.to_numeric(df_v[c], errors="coerce").fillna(0)
                    viajes_list.append(df_v)
            if viajes_list:
                data["viajes"] = pd.concat(viajes_list, ignore_index=True)

        # ── PARADAS
        elif "paradas" in bname.lower() or "Paradas" in bname:
            paradas_list = []
            for sheet in xl.sheet_names:
                if "Paradas" in sheet or "Placa" in sheet:
                    placa = sheet.replace("Paradas-", "").replace("Placa-", "")
                    df_p = xl.parse(sheet)
                    df_p.columns = [str(c).strip() for c in df_p.columns]
                    df_p["Placa"] = placa
                    df_p["Comienzo"] = pd.to_datetime(df_p["Comienzo"], errors="coerce")
                    df_p["Fin"] = pd.to_datetime(df_p["Fin"], errors="coerce")
                    df_p["Fecha"] = df_p["Comienzo"].dt.normalize()
                    paradas_list.append(df_p)
            if paradas_list:
                data["paradas"] = pd.concat(paradas_list, ignore_index=True)

    return data


def merge_new_data(existing: dict, new: dict) -> dict:
    """Une datos nuevos con existentes evitando duplicados por Fecha+Placa."""
    merged = {}
    keys = set(existing.keys()) | set(new.keys())
    for k in keys:
        if k in existing and k in new:
            df_old = existing[k]
            df_new = new[k]
            combined = pd.concat([df_old, df_new], ignore_index=True)
            # Deduplicar según columnas disponibles
            dup_cols = [c for c in ["Fecha","Placa","Vuelta","Comienzo","№"] if c in combined.columns]
            if dup_cols:
                combined = combined.drop_duplicates(subset=dup_cols)
            merged[k] = combined
        elif k in existing:
            merged[k] = existing[k]
        else:
            merged[k] = new[k]
    return merged


# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
def render_sidebar(data: dict):
    with st.sidebar:
        # Logo
        logo_b64 = get_logo_b64()
        if logo_b64:
            st.markdown(
                f'<div style="text-align:center;padding:12px 0 4px">'
                f'<img src="data:image/png;base64,{logo_b64}" style="max-width:160px;border-radius:8px">'
                f'</div>', unsafe_allow_html=True)
        else:
            st.markdown(
                f'<div style="text-align:center;padding:16px 0 4px">'
                f'<span style="font-size:2rem;font-weight:900;color:{COL_AMARILLO}">eol.</span>'
                f'<br><span style="font-size:0.75rem;color:#aaa;letter-spacing:0.15em">TRANSPORTES</span>'
                f'</div>', unsafe_allow_html=True)

        st.markdown(f'<hr style="border-color:{COL_AMARILLO};margin:10px 0">', unsafe_allow_html=True)
        st.markdown(f'<p style="color:{COL_AMARILLO};font-size:0.7rem;text-align:center;letter-spacing:0.1em;margin:0 0 12px">FILTROS GLOBALES</p>', unsafe_allow_html=True)

        # Rango de fechas
        all_dates = []
        for key in ["margen","kilometraje","combustible_raw","planilla","viajes"]:
            if key in data and "Fecha" in data[key].columns:
                all_dates.extend(data[key]["Fecha"].dropna().dt.date.tolist())

        if all_dates:
            min_d = min(all_dates)
            max_d = max(all_dates)
        else:
            min_d = date(2026, 2, 1)
            max_d = date(2026, 12, 31)

        fecha_inicio = st.date_input("📅 Desde", value=min_d, min_value=min_d, max_value=date(2026,12,31))
        fecha_fin    = st.date_input("📅 Hasta", value=max_d, min_value=min_d, max_value=date(2026,12,31))

        # Placas globales
        all_placas = set()
        for key in ["margen","kilometraje","combustible","planilla","viajes","paradas","peso","cocheras"]:
            if key in data and "Placa" in data[key].columns:
                all_placas.update(data[key]["Placa"].dropna().astype(str).unique())
        all_placas = sorted([p for p in all_placas if p and p != "nan"])

        sel_all_placas = st.checkbox("✅ Todas las placas", value=True, key="cb_placas_global")
        placas_sel = all_placas if sel_all_placas else st.multiselect(
            "🚛 Placas", options=all_placas, default=all_placas)

        st.markdown(f'<hr style="border-color:#444;margin:12px 0">', unsafe_allow_html=True)

        # Módulo de carga
        st.markdown(f'<p style="color:{COL_AMARILLO};font-size:0.7rem;letter-spacing:0.1em;margin:0 0 8px">📤 CARGAR NUEVOS DATOS</p>', unsafe_allow_html=True)
        uploaded = st.file_uploader("ZIP o Excel (.xlsx)", type=["zip","xlsx"], key="uploader")
        if uploaded:
            with st.spinner("Procesando..."):
                try:
                    new_data = load_data(uploaded.read(), uploaded.name)
                    if "raw_data" in st.session_state:
                        st.session_state["raw_data"] = merge_new_data(st.session_state["raw_data"], new_data)
                    else:
                        st.session_state["raw_data"] = new_data
                    st.success("✅ Datos cargados y unidos correctamente.")
                    load_data.clear()
                except Exception as e:
                    st.error(f"Error al cargar: {e}")

        st.markdown(f'<hr style="border-color:#444;margin:12px 0">', unsafe_allow_html=True)
        st.markdown(f'<p style="color:#666;font-size:0.65rem;text-align:center">EOL Perú © 2026<br>Dashboard Logístico v1.0</p>', unsafe_allow_html=True)

    return pd.Timestamp(fecha_inicio), pd.Timestamp(fecha_fin), placas_sel


def filter_df(df, fecha_ini, fecha_fin, placas, date_col="Fecha", placa_col="Placa"):
    if df is None or df.empty: return df
    mask = pd.Series([True] * len(df), index=df.index)
    if date_col in df.columns:
        mask &= (df[date_col] >= fecha_ini) & (df[date_col] <= fecha_fin + pd.Timedelta(days=1))
    if placa_col in df.columns and placas:
        mask &= df[placa_col].astype(str).isin([str(p) for p in placas])
    return df[mask].copy()


# ─────────────────────────────────────────────
# SECCIÓN A — FINANCIERO
# ─────────────────────────────────────────────
def seccion_a(data, fi, ff, placas):
    st.markdown('<div class="section-header">📊 A — Dashboard Financiero (Ingresos vs. Margen)</div>', unsafe_allow_html=True)

    df_m = filter_df(data.get("margen", pd.DataFrame()), fi, ff, placas)
    df_mc = filter_df(data.get("margen_camion", pd.DataFrame()), fi, ff, placas)

    if df_m.empty:
        st.info("Sin datos de margen para el período/placas seleccionados.")
        return

    # KPIs
    total_ing   = df_m["Ingresos"].sum()
    total_margen = df_m["MargenCamion"].sum()
    pct_margen  = total_margen / total_ing if total_ing else 0
    total_gastos = df_m["TotalGastos"].sum()
    margen_neto = df_m["MargenNeto"].sum()

    c1,c2,c3,c4 = st.columns(4)
    with c1: st.markdown(kpi_html("Total Ingresos", fmt_soles(total_ing)), unsafe_allow_html=True)
    with c2: st.markdown(kpi_html("Total Margen", fmt_soles(total_margen),
                                   ok=pct_margen>=0.3, alert=pct_margen<0.15), unsafe_allow_html=True)
    with c3: st.markdown(kpi_html("% Margen Real", f"{pct_margen*100:.1f}%",
                                   sub="Margen / Ingresos",
                                   ok=pct_margen>=0.3, alert=pct_margen<0.15), unsafe_allow_html=True)
    with c4: st.markdown(kpi_html("Margen Neto", fmt_soles(margen_neto),
                                   sub="Ing − Sueldos − Combustible"), unsafe_allow_html=True)

    st.markdown("---")
    col1, col2 = st.columns(2)

    with col1:
        # Barras agrupadas Ingresos vs Margen por placa
        gb = df_m.groupby("Placa").agg(Ingresos=("Ingresos","sum"), Margen=("MargenCamion","sum")).reset_index()
        fig = go.Figure()
        fig.add_trace(go.Bar(name="Ingresos", x=gb["Placa"], y=gb["Ingresos"],
                             marker_color=COL_AZUL, text=gb["Ingresos"].apply(lambda x: f"S/{x:,.0f}"),
                             textposition="outside"))
        fig.add_trace(go.Bar(name="Margen", x=gb["Placa"], y=gb["Margen"],
                             marker_color=COL_AMARILLO, text=gb["Margen"].apply(lambda x: f"S/{x:,.0f}"),
                             textposition="outside"))
        fig.update_layout(barmode="group")
        plotly_layout(fig, "Ingresos vs. Margen por Placa", 420)
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        # % Margen por placa
        gb2 = df_m.groupby("Placa").agg(Ingresos=("Ingresos","sum"), Margen=("MargenCamion","sum")).reset_index()
        gb2["PctMargen"] = np.where(gb2["Ingresos"]!=0, gb2["Margen"]/gb2["Ingresos"]*100, 0)
        colors = [COL_VERDE if v >= 30 else (COL_AMARILLO if v >= 15 else COL_ROJO) for v in gb2["PctMargen"]]
        fig2 = go.Figure(go.Bar(x=gb2["Placa"], y=gb2["PctMargen"],
                                marker_color=colors,
                                text=gb2["PctMargen"].apply(lambda x: f"{x:.1f}%"),
                                textposition="outside"))
        fig2.add_hline(y=30, line_dash="dash", line_color=COL_VERDE, annotation_text="Meta 30%")
        plotly_layout(fig2, "% Margen Real por Placa", 420)
        st.plotly_chart(fig2, use_container_width=True)

    # Tabla detallada con margen neto
    st.markdown("##### 📋 Tabla Detallada con Margen Neto")
    cols_show = ["Fecha","Placa","Ingresos","Conductor","Aux1","Aux2","Combustible","TotalGastos","MargenCamion","MargenNeto","PctMC"]
    df_show = df_m[[c for c in cols_show if c in df_m.columns]].copy()
    df_show["Fecha"] = df_show["Fecha"].dt.strftime("%d/%m/%Y")
    df_show["PctMC"] = df_show["PctMC"].apply(lambda x: f"{x*100:.1f}%")
    # Fila totales
    totals = {"Fecha": "TOTAL", "Placa": "", "Ingresos": df_m["Ingresos"].sum(),
              "TotalGastos": df_m["TotalGastos"].sum(), "MargenCamion": df_m["MargenCamion"].sum(),
              "MargenNeto": df_m["MargenNeto"].sum(),
              "PctMC": f"{pct_margen*100:.1f}%"}
    df_totals = pd.DataFrame([totals])
    df_display = pd.concat([df_show, df_totals], ignore_index=True)
    st.dataframe(df_display, use_container_width=True, height=320)
    st.download_button("⬇️ Descargar tabla Finanzas", to_excel_bytes(df_show), "finanzas_eol.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # Margen por camión (pestaña 2.1)
    if not df_mc.empty:
        st.markdown("##### 📊 Evolución Margen por Camión (pestaña 2.1)")
        fig3 = px.line(df_mc, x="Fecha", y="MargenDia", color="Placa",
                       color_discrete_sequence=COLORS_PLACAS,
                       markers=True, labels={"MargenDia":"Margen (S/)", "Fecha":"Fecha"})
        plotly_layout(fig3, "Margen Diario por Camión", 380)
        st.plotly_chart(fig3, use_container_width=True)


# ─────────────────────────────────────────────
# SECCIÓN B — OPERACIÓN Y KILOMETRAJE
# ─────────────────────────────────────────────
def seccion_b(data, fi, ff, placas):
    st.markdown('<div class="section-header">🛣️ B — Operación y Kilometraje</div>', unsafe_allow_html=True)

    df_km  = filter_df(data.get("kilometraje", pd.DataFrame()), fi, ff, placas)
    df_exc = filter_df(data.get("excesos", pd.DataFrame()), fi, ff, placas)
    df_coch= filter_df(data.get("cocheras", pd.DataFrame()), fi, ff, placas)

    # KPIs
    total_km = df_km["Kilometraje"].sum()
    n_excesos = df_km["Exceso70"].sum() if not df_km.empty else 0
    placas_op = df_km["Placa"].nunique() if not df_km.empty else 0
    dias_op   = df_km["Fecha"].nunique() if not df_km.empty else 0

    c1,c2,c3,c4 = st.columns(4)
    with c1: st.markdown(kpi_html("Km Total Acumulado", f"{total_km:,.1f} km"), unsafe_allow_html=True)
    with c2: st.markdown(kpi_html("Excesos > 70 km", str(int(n_excesos)),
                                   sub="días con exceso detectados",
                                   alert=n_excesos>10, ok=n_excesos==0), unsafe_allow_html=True)
    with c3: st.markdown(kpi_html("Unidades Operando", str(placas_op)), unsafe_allow_html=True)
    with c4: st.markdown(kpi_html("Días con Operación", str(dias_op)), unsafe_allow_html=True)

    st.markdown("---")
    col1, col2 = st.columns(2)

    with col1:
        if not df_km.empty:
            gb = df_km.groupby("Placa")["Kilometraje"].sum().reset_index().sort_values("Kilometraje", ascending=False)
            fig = px.bar(gb, x="Placa", y="Kilometraje", color="Placa",
                         color_discrete_sequence=COLORS_PLACAS,
                         text=gb["Kilometraje"].apply(lambda x: f"{x:,.0f}"),
                         labels={"Kilometraje":"Km acumulado"})
            fig.update_traces(textposition="outside")
            plotly_layout(fig, "Kilometraje Acumulado por Unidad", 400)
            st.plotly_chart(fig, use_container_width=True)

    with col2:
        # Mapa de calor: km por placa y fecha
        if not df_km.empty:
            pivot = df_km.pivot_table(index="Placa", columns=df_km["Fecha"].dt.strftime("%d/%m"),
                                      values="Kilometraje", aggfunc="sum", fill_value=0)
            fig2 = px.imshow(pivot, color_continuous_scale=[[0,"#F5F5F5"],[0.4,COL_AMARILLO],[1,COL_AZUL]],
                             aspect="auto", labels=dict(x="Fecha", y="Placa", color="Km"))
            plotly_layout(fig2, "Mapa de Calor: Km por Placa y Fecha", 400)
            st.plotly_chart(fig2, use_container_width=True)

    # Excesos > 70 km
    st.markdown("##### 🚨 Excesos de Kilometraje (> 70 km/día)")
    if not df_km.empty:
        df_exc_show = df_km[df_km["Exceso70"]].copy()
        if df_exc_show.empty:
            st.success("✅ No se registraron excesos de 70 km en el período seleccionado.")
        else:
            c1, c2 = st.columns(2)
            with c1:
                gb_exc = df_exc_show.groupby("Placa").size().reset_index(name="N_Excesos").sort_values("N_Excesos", ascending=False)
                fig3 = px.bar(gb_exc, x="Placa", y="N_Excesos", color="Placa",
                              color_discrete_sequence=[COL_ROJO]*len(gb_exc),
                              text="N_Excesos", labels={"N_Excesos":"Veces con exceso"})
                fig3.update_traces(textposition="outside")
                plotly_layout(fig3, "Frecuencia de Excesos por Placa", 360)
                st.plotly_chart(fig3, use_container_width=True)
            with c2:
                df_exc_tbl = df_exc_show[["Fecha","Placa","Kilometraje"]].copy()
                df_exc_tbl["Fecha"] = df_exc_tbl["Fecha"].dt.strftime("%d/%m/%Y")
                df_exc_tbl["Exceso (km)"] = (df_exc_tbl["Kilometraje"] - 70).round(2)
                st.dataframe(df_exc_tbl, use_container_width=True, height=300)
                st.download_button("⬇️ Descargar Excesos", to_excel_bytes(df_exc_tbl), "excesos_km.xlsx",
                                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # Cocheras
    st.markdown("##### 🏭 Gestión de Cocheras")
    if not df_coch.empty:
        col1, col2 = st.columns(2)
        with col1:
            gb_coch = df_coch.groupby("Cochera")["Placa"].count().reset_index(name="N")
            fig4 = px.pie(gb_coch, names="Cochera", values="N",
                          color_discrete_map={"SANTA ANITA": COL_AZUL, "HUACHIPA": COL_AMARILLO},
                          hole=0.4)
            plotly_layout(fig4, "Distribución por Cochera", 360)
            st.plotly_chart(fig4, use_container_width=True)
        with col2:
            pivot_coch = df_coch.pivot_table(index="Placa",
                                              columns=df_coch["Fecha"].dt.strftime("%d/%m"),
                                              values="Cochera", aggfunc="first")
            pivot_coch = pivot_coch.fillna("—")
            st.dataframe(pivot_coch.style.applymap(
                lambda v: f"background-color:{COL_AZUL};color:white" if v=="SANTA ANITA"
                           else (f"background-color:{COL_AMARILLO};color:{COL_NEGRO}" if v=="HUACHIPA" else "")),
                use_container_width=True, height=320)

    # Cierre de ruta
    st.markdown("##### 🔒 Tabla de Cierres de Ruta (última fecha por placa)")
    if not df_km.empty:
        cierres = df_km.groupby("Placa").agg(UltimaFecha=("Fecha","max"), KmTotal=("Kilometraje","sum")).reset_index()
        cierres["UltimaFecha"] = cierres["UltimaFecha"].dt.strftime("%d/%m/%Y")
        st.dataframe(cierres, use_container_width=True)


# ─────────────────────────────────────────────
# SECCIÓN C — COMBUSTIBLE
# ─────────────────────────────────────────────
def seccion_c(data, fi, ff, placas):
    st.markdown('<div class="section-header">⛽ C — Control de Combustible y Eficiencia</div>', unsafe_allow_html=True)

    df_c = filter_df(data.get("combustible", pd.DataFrame()), fi, ff, placas, date_col=None)
    df_cr= filter_df(data.get("combustible_raw", pd.DataFrame()), fi, ff, placas)

    if df_c.empty:
        st.info("Sin datos de combustible.")
        return

    total_gal = df_c["Galones"].sum()
    total_gasto = df_c["Gasto"].sum()
    total_km = df_c["Kilometraje"].sum()
    km_gal_avg = total_km / total_gal if total_gal else 0

    c1,c2,c3,c4 = st.columns(4)
    with c1: st.markdown(kpi_html("Total Galones", f"{total_gal:,.2f} gal"), unsafe_allow_html=True)
    with c2: st.markdown(kpi_html("Gasto Total Combustible", fmt_soles(total_gasto)), unsafe_allow_html=True)
    with c3: st.markdown(kpi_html("Km Promedio / Galón", f"{km_gal_avg:.2f} km/gal"), unsafe_allow_html=True)
    with c4: st.markdown(kpi_html("Km Total (flota)", f"{total_km:,.1f} km"), unsafe_allow_html=True)

    st.markdown("---")
    col1, col2 = st.columns(2)

    with col1:
        fig = px.bar(df_c, x="Placa", y="Galones", color="Placa",
                     color_discrete_sequence=COLORS_PLACAS,
                     text=df_c["Galones"].apply(lambda x: f"{x:.1f}"),
                     labels={"Galones":"Galones consumidos"})
        fig.update_traces(textposition="outside")
        plotly_layout(fig, "Galones Consumidos por Unidad", 400)
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        fig2 = go.Figure()
        fig2.add_trace(go.Bar(name="Km/Galón", x=df_c["Placa"], y=df_c["Km_por_galon"],
                              marker_color=COL_AZUL, yaxis="y",
                              text=df_c["Km_por_galon"].apply(lambda x: f"{x:.1f}"), textposition="outside"))
        fig2.add_trace(go.Bar(name="Gasto/Km (S/)", x=df_c["Placa"], y=df_c["Gasto_por_km"],
                              marker_color=COL_AMARILLO, yaxis="y2",
                              text=df_c["Gasto_por_km"].apply(lambda x: f"{x:.3f}"), textposition="outside"))
        fig2.update_layout(
            yaxis=dict(title="Km/Galón"),
            yaxis2=dict(title="Gasto/Km (S/)", overlaying="y", side="right"),
            barmode="group"
        )
        plotly_layout(fig2, "Eficiencia: Km/Galón vs. Gasto/Km por Placa", 400)
        st.plotly_chart(fig2, use_container_width=True)

    col3, col4 = st.columns(2)
    with col3:
        fig3 = px.bar(df_c, x="Placa", y="Gasto_por_viaje", color="Placa",
                      color_discrete_sequence=COLORS_PLACAS,
                      text=df_c["Gasto_por_viaje"].apply(lambda x: f"S/{x:.1f}"),
                      labels={"Gasto_por_viaje":"S/ por viaje"})
        fig3.update_traces(textposition="outside")
        plotly_layout(fig3, "Gasto Promedio por Viaje", 380)
        st.plotly_chart(fig3, use_container_width=True)

    with col4:
        # Merma vs objetivo
        df_c2 = df_c.copy()
        df_c2["Merma_abs"] = df_c2["Merma"].abs()
        fig4 = go.Figure()
        fig4.add_trace(go.Bar(name="Objetivo (km)", x=df_c2["Placa"], y=df_c2["Objetivo"],
                              marker_color=COL_VERDE))
        fig4.add_trace(go.Bar(name="Real (km)", x=df_c2["Placa"], y=df_c2["Kilometraje"],
                              marker_color=COL_AZUL))
        fig4.update_layout(barmode="group")
        plotly_layout(fig4, "Kilometraje Real vs. Objetivo por Placa", 380)
        st.plotly_chart(fig4, use_container_width=True)

    # Tabla abastecimientos históricos
    if not df_cr.empty:
        st.markdown("##### 📋 Historial de Abastecimientos")
        df_show = df_cr.copy()
        df_show["Fecha"] = df_show["Fecha"].dt.strftime("%d/%m/%Y")
        st.dataframe(df_show, use_container_width=True, height=280)
        st.download_button("⬇️ Descargar Combustible", to_excel_bytes(df_show), "combustible_eol.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ─────────────────────────────────────────────
# SECCIÓN D — CARGA Y CAPACIDAD
# ─────────────────────────────────────────────
def seccion_d(data, fi, ff, placas):
    st.markdown('<div class="section-header">⚖️ D — Carga y Capacidad Crítica</div>', unsafe_allow_html=True)

    df_p = filter_df(data.get("peso", pd.DataFrame()), fi, ff, placas)
    if df_p.empty:
        st.info("Sin datos de peso/carga.")
        return

    # Filtro vuelta
    vueltas_disp = sorted(df_p["Vuelta"].unique().tolist())
    vueltas_sel = st.multiselect("🔄 Vuelta", options=vueltas_disp, default=vueltas_disp, key="vuelta_d")
    if vueltas_sel:
        df_p = df_p[df_p["Vuelta"].isin(vueltas_sel)]

    total_viajes = len(df_p)
    total_sobrecap = df_p["SobreCapacidad"].sum()
    pct_sobre = total_sobrecap / total_viajes * 100 if total_viajes else 0

    c1,c2,c3 = st.columns(3)
    with c1: st.markdown(kpi_html("Total Cargas Registradas", str(total_viajes)), unsafe_allow_html=True)
    with c2: st.markdown(kpi_html("Excesos de Capacidad", str(int(total_sobrecap)),
                                   sub=f"{pct_sobre:.1f}% del total",
                                   alert=pct_sobre>20, ok=pct_sobre==0), unsafe_allow_html=True)
    with c3:
        avg_peso = df_p["Peso_kg"].mean()
        st.markdown(kpi_html("Peso Promedio por Carga", f"{avg_peso:,.0f} kg"), unsafe_allow_html=True)

    st.markdown("---")
    col1, col2 = st.columns(2)

    with col1:
        # Peso por placa y vuelta
        gb = df_p.groupby(["Placa","Vuelta"])["Peso_kg"].mean().reset_index()
        fig = px.bar(gb, x="Placa", y="Peso_kg", color="Vuelta",
                     barmode="group",
                     color_discrete_map={"Vuelta 1": COL_AZUL, "Vuelta 2": COL_AMARILLO},
                     text=gb["Peso_kg"].apply(lambda x: f"{x:,.0f}"),
                     labels={"Peso_kg":"Peso promedio (kg)"})
        fig.update_traces(textposition="outside")
        plotly_layout(fig, "Peso Promedio por Camión y Vuelta", 420)
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        # Excesos por placa
        gb_exc = df_p.groupby("Placa")["SobreCapacidad"].sum().reset_index(name="Excesos")
        gb_exc = gb_exc[gb_exc["Excesos"] > 0].sort_values("Excesos", ascending=False)
        if gb_exc.empty:
            st.success("✅ No hay cargas que excedan la capacidad máxima pactada.")
        else:
            fig2 = px.bar(gb_exc, x="Placa", y="Excesos", color="Placa",
                          color_discrete_sequence=[COL_ROJO]*len(gb_exc),
                          text="Excesos", labels={"Excesos":"Veces sobre cap. máxima"})
            fig2.update_traces(textposition="outside")
            plotly_layout(fig2, "Excesos de Capacidad Máxima por Camión", 420)
            st.plotly_chart(fig2, use_container_width=True)

    # Evolución temporal
    gb_t = df_p.groupby(["Fecha","Vuelta"])["Peso_kg"].sum().reset_index()
    fig3 = px.line(gb_t, x="Fecha", y="Peso_kg", color="Vuelta",
                   color_discrete_map={"Vuelta 1": COL_AZUL, "Vuelta 2": COL_AMARILLO},
                   markers=True, labels={"Peso_kg":"Peso total (kg)"})
    plotly_layout(fig3, "Evolución de Carga Diaria Total (V1 + V2)", 360)
    st.plotly_chart(fig3, use_container_width=True)

    # Tabla detallada
    st.markdown("##### 📋 Detalle de Cargas")
    df_show = df_p.copy()
    df_show["Fecha"] = df_show["Fecha"].dt.strftime("%d/%m/%Y")
    df_show["SobreCapacidad"] = df_show["SobreCapacidad"].map({True:"🔴 SÍ", False:"🟢 NO"})
    st.dataframe(df_show, use_container_width=True, height=300)
    st.download_button("⬇️ Descargar Carga/Peso", to_excel_bytes(df_p), "peso_eol.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ─────────────────────────────────────────────
# SECCIÓN E — TIEMPOS Y MOVIMIENTOS
# ─────────────────────────────────────────────
def seccion_e(data, fi, ff, placas):
    st.markdown('<div class="section-header">🗺️ E — Tiempos y Movimientos (Auditoría)</div>', unsafe_allow_html=True)

    df_v = filter_df(data.get("viajes", pd.DataFrame()), fi, ff, placas)
    df_par = filter_df(data.get("paradas", pd.DataFrame()), fi, ff, placas)

    if df_v.empty and df_par.empty:
        st.info("Sin datos de viajes/paradas.")
        return

    # KPIs
    n_viajes   = len(df_v) if not df_v.empty else 0
    n_paradas  = len(df_par) if not df_par.empty else 0
    placas_v   = df_v["Placa"].nunique() if not df_v.empty else 0
    km_total   = df_v["Kilometraje"].sum() if not df_v.empty and "Kilometraje" in df_v.columns else 0

    c1,c2,c3,c4 = st.columns(4)
    with c1: st.markdown(kpi_html("Total Viajes Registrados", str(n_viajes)), unsafe_allow_html=True)
    with c2: st.markdown(kpi_html("Total Paradas", str(n_paradas)), unsafe_allow_html=True)
    with c3: st.markdown(kpi_html("Unidades con Datos", str(placas_v)), unsafe_allow_html=True)
    with c4: st.markdown(kpi_html("Km Totales (viajes)", f"{km_total:,.1f} km"), unsafe_allow_html=True)

    st.markdown("---")
    tab1, tab2, tab3 = st.tabs(["📊 Viajes por Placa/Fecha", "⏸️ Paradas", "🗺️ Mapa de Recorrido"])

    with tab1:
        if not df_v.empty:
            col1, col2 = st.columns(2)
            with col1:
                gb = df_v.groupby("Placa").size().reset_index(name="N_Viajes")
                fig = px.bar(gb, x="Placa", y="N_Viajes", color="Placa",
                             color_discrete_sequence=COLORS_PLACAS,
                             text="N_Viajes", labels={"N_Viajes":"Número de viajes"})
                fig.update_traces(textposition="outside")
                plotly_layout(fig, "Total Viajes por Placa", 400)
                st.plotly_chart(fig, use_container_width=True)
            with col2:
                gb2 = df_v.groupby(["Fecha","Placa"]).size().reset_index(name="N_Viajes")
                fig2 = px.line(gb2, x="Fecha", y="N_Viajes", color="Placa",
                               color_discrete_sequence=COLORS_PLACAS,
                               markers=True, labels={"N_Viajes":"Viajes/día"})
                plotly_layout(fig2, "Viajes Diarios por Placa", 400)
                st.plotly_chart(fig2, use_container_width=True)

            # Velocidad máxima
            if "Velocidad máxima" in df_v.columns:
                st.markdown("##### 🏎️ Velocidades Máximas Registradas")
                vel_max = df_v.groupby("Placa")["Velocidad máxima"].max().reset_index()
                vel_max_excesos = vel_max[vel_max["Velocidad máxima"] > 90]
                if not vel_max_excesos.empty:
                    st.warning(f"⚠️ {len(vel_max_excesos)} placas registraron velocidades máximas superiores a 90 km/h")
                fig3 = px.bar(vel_max, x="Placa", y="Velocidad máxima", color="Placa",
                              color_discrete_sequence=COLORS_PLACAS,
                              text=vel_max["Velocidad máxima"].apply(lambda x: f"{x:.0f} km/h"))
                fig3.add_hline(y=90, line_dash="dash", line_color=COL_ROJO, annotation_text="Límite 90 km/h")
                fig3.update_traces(textposition="outside")
                plotly_layout(fig3, "Velocidad Máxima por Placa", 360)
                st.plotly_chart(fig3, use_container_width=True)

    with tab2:
        if not df_par.empty:
            col1, col2 = st.columns(2)
            with col1:
                gb_p = df_par.groupby("Placa").size().reset_index(name="N_Paradas")
                fig_p = px.bar(gb_p, x="Placa", y="N_Paradas", color="Placa",
                               color_discrete_sequence=COLORS_PLACAS,
                               text="N_Paradas", labels={"N_Paradas":"Número de paradas"})
                fig_p.update_traces(textposition="outside")
                plotly_layout(fig_p, "Total Paradas por Placa", 400)
                st.plotly_chart(fig_p, use_container_width=True)
            with col2:
                gb_pd = df_par.groupby(["Fecha","Placa"]).size().reset_index(name="N_Paradas")
                fig_pd = px.line(gb_pd, x="Fecha", y="N_Paradas", color="Placa",
                                 color_discrete_sequence=COLORS_PLACAS,
                                 markers=True, labels={"N_Paradas":"Paradas/día"})
                plotly_layout(fig_pd, "Paradas Diarias por Placa", 400)
                st.plotly_chart(fig_pd, use_container_width=True)

            st.markdown("##### 📋 Detalle de Paradas")
            df_par_show = df_par.copy()
            df_par_show["Fecha"] = df_par_show["Fecha"].dt.strftime("%d/%m/%Y")
            df_par_show["Comienzo"] = df_par_show["Comienzo"].dt.strftime("%d/%m/%Y %H:%M")
            st.dataframe(df_par_show[["Placa","Fecha","Comienzo","Duración","Ubicación"]],
                         use_container_width=True, height=300)
            st.download_button("⬇️ Descargar Paradas", to_excel_bytes(df_par_show), "paradas_eol.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with tab3:
        st.markdown("##### 🗺️ Mapa de Recorrido por Viaje")
        if df_v.empty:
            st.info("Sin datos de viajes.")
        else:
            col_f, col_p, col_n = st.columns(3)
            with col_f:
                fechas_disp = sorted(df_v["Fecha"].dt.date.unique())
                fecha_sel = st.selectbox("Fecha", fechas_disp, key="mapa_fecha")
            df_v_dia = df_v[df_v["Fecha"].dt.date == fecha_sel]
            with col_p:
                placas_dia = sorted(df_v_dia["Placa"].unique())
                placa_sel = st.selectbox("Placa", placas_dia, key="mapa_placa")
            df_v_placa = df_v_dia[df_v_dia["Placa"] == placa_sel].reset_index(drop=True)
            with col_n:
                viaje_nums = list(range(1, len(df_v_placa)+1))
                viaje_sel = st.selectbox("N° Viaje", viaje_nums, key="mapa_viaje")

            if viaje_sel and len(df_v_placa) >= viaje_sel:
                row = df_v_placa.iloc[viaje_sel-1]
                st.markdown(f"""
                **Placa:** {placa_sel} &nbsp;|&nbsp;
                **Inicio:** {row.get('Ubicación inicial','—')} &nbsp;|&nbsp;
                **Fin:** {row.get('Ubicación final','—')} &nbsp;|&nbsp;
                **Km:** {row.get('Kilometraje',0):.2f} &nbsp;|&nbsp;
                **Duración:** {row.get('Duración','—')}
                """)

                # Intentar renderizar mapa con folium
                try:
                    import folium
                    from streamlit_folium import st_folium

                    def geocode_address(address):
                        """Extrae o geocodifica una dirección."""
                        import re
                        lat_match = re.search(r'@(-?\d+\.\d+),(-?\d+\.\d+)', str(address))
                        if lat_match:
                            return float(lat_match.group(1)), float(lat_match.group(2))
                        return None

                    # Centro aproximado Lima Perú
                    m = folium.Map(location=[-12.05, -77.04], zoom_start=12,
                                   tiles="CartoDB positron")

                    # Marcador inicio
                    folium.Marker(
                        location=[-12.05, -77.04],
                        tooltip=f"Inicio: {row.get('Ubicación inicial','—')}",
                        icon=folium.Icon(color="green", icon="play")
                    ).add_to(m)
                    folium.Marker(
                        location=[-12.06, -77.03],
                        tooltip=f"Fin: {row.get('Ubicación final','—')}",
                        icon=folium.Icon(color="red", icon="stop")
                    ).add_to(m)

                    st_folium(m, width=700, height=400)
                    st.caption("📍 Mapa aproximado. Las ubicaciones exactas requieren geocodificación en tiempo real.")

                except ImportError:
                    # fallback: mostrar link a Google Maps
                    from urllib.parse import quote
                    dir_inicio = quote(str(row.get('Ubicación inicial', '')))
                    dir_fin    = quote(str(row.get('Ubicación final', '')))
                    url = f"https://www.google.com/maps/dir/{dir_inicio}/{dir_fin}"
                    st.markdown(f"""
                    <div style='background:{COL_BLANCO};border:2px solid {COL_AMARILLO};border-radius:10px;padding:20px;text-align:center'>
                      <p style='font-size:1.1rem;font-weight:600;color:{COL_AZUL}'>🗺️ Ver recorrido en Google Maps</p>
                      <p><b>Inicio:</b> {row.get('Ubicación inicial','—')}</p>
                      <p><b>Fin:</b> {row.get('Ubicación final','—')}</p>
                      <a href="{url}" target="_blank"
                         style='background:{COL_AMARILLO};color:{COL_NEGRO};padding:10px 24px;border-radius:8px;
                                font-weight:700;text-decoration:none;display:inline-block;margin-top:8px'>
                        📍 Abrir en Google Maps
                      </a>
                    </div>
                    """, unsafe_allow_html=True)


# ─────────────────────────────────────────────
# SECCIÓN F — PLANILLA
# ─────────────────────────────────────────────
def seccion_f(data, fi, ff, placas):
    st.markdown('<div class="section-header">👷 F — Control de Planilla</div>', unsafe_allow_html=True)

    df_plan = data.get("planilla", pd.DataFrame())
    if df_plan.empty:
        st.info("Sin datos de planilla.")
        return

    # Filtros específicos sección F (incluye unidades externas)
    placas_plan_all = sorted(df_plan["Placa"].dropna().astype(str).unique())
    sel_all_plan = st.checkbox("✅ Todas las placas (incl. Unidades Externas)", value=True, key="cb_plan")
    if sel_all_plan:
        placas_plan_sel = placas_plan_all
    else:
        placas_plan_sel = st.multiselect("🚛 Placas (Planilla)", options=placas_plan_all,
                                          default=placas_plan_all, key="placas_plan")

    df_p = filter_df(df_plan, fi, ff, placas_plan_sel)
    if df_p.empty:
        st.info("Sin datos para los filtros seleccionados.")
        return

    # Filtro trabajador (según placas)
    trabajadores_disp = sorted(df_p["Trabajador"].dropna().unique().tolist())
    sel_all_trab = st.checkbox("✅ Todos los trabajadores", value=True, key="cb_trab")
    if not sel_all_trab:
        trab_sel = st.multiselect("👤 Trabajadores", options=trabajadores_disp,
                                   default=trabajadores_disp, key="trabajadores")
        df_p = df_p[df_p["Trabajador"].isin(trab_sel)]

    # KPIs
    gasto_total   = df_p["Pago final"].sum()
    n_trabajadores= df_p["Trabajador"].nunique()
    gasto_promedio= gasto_total / df_p["Fecha"].nunique() if df_p["Fecha"].nunique() else 0
    desc_total    = df_p["Descuento"].sum()

    c1,c2,c3,c4 = st.columns(4)
    with c1: st.markdown(kpi_html("Gasto Total Planilla", fmt_soles(gasto_total)), unsafe_allow_html=True)
    with c2: st.markdown(kpi_html("N° Trabajadores", str(n_trabajadores)), unsafe_allow_html=True)
    with c3: st.markdown(kpi_html("Gasto Promedio/Día", fmt_soles(gasto_promedio)), unsafe_allow_html=True)
    with c4: st.markdown(kpi_html("Total Descuentos", fmt_soles(desc_total),
                                   alert=desc_total>500), unsafe_allow_html=True)

    st.markdown("---")
    col1, col2 = st.columns(2)

    with col1:
        # Gasto por placa
        gb = df_p.groupby("Placa")["Pago final"].sum().reset_index().sort_values("Pago final", ascending=False)
        fig = px.bar(gb, x="Placa", y="Pago final", color="Placa",
                     color_discrete_sequence=COLORS_PLACAS,
                     text=gb["Pago final"].apply(lambda x: f"S/{x:,.0f}"),
                     labels={"Pago final":"Gasto planilla (S/)"})
        fig.update_traces(textposition="outside")
        plotly_layout(fig, "Gasto de Planilla por Placa", 420)
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        # Evolución diaria
        gb2 = df_p.groupby("Fecha")["Pago final"].sum().reset_index()
        fig2 = px.area(gb2, x="Fecha", y="Pago final",
                       color_discrete_sequence=[COL_AZUL],
                       labels={"Pago final":"Gasto diario (S/)"})
        fig2.update_traces(fillcolor=f"rgba(26,58,92,0.2)", line_color=COL_AZUL)
        plotly_layout(fig2, "Evolución Diaria del Gasto de Planilla", 420)
        st.plotly_chart(fig2, use_container_width=True)

    col3, col4 = st.columns(2)
    with col3:
        # Top trabajadores por gasto
        gb3 = df_p.groupby("Trabajador")["Pago final"].sum().reset_index().sort_values("Pago final", ascending=True).tail(12)
        fig3 = px.bar(gb3, x="Pago final", y="Trabajador", orientation="h",
                      color="Pago final", color_continuous_scale=[[0, "#E3F2FD"],[1, COL_AZUL]],
                      text=gb3["Pago final"].apply(lambda x: f"S/{x:,.0f}"),
                      labels={"Pago final":"Total cobrado (S/)"})
        fig3.update_traces(textposition="outside")
        plotly_layout(fig3, "Top Trabajadores por Gasto Total", 420)
        st.plotly_chart(fig3, use_container_width=True)

    with col4:
        # Distribución por empresa
        if "EMPRESA" in df_p.columns:
            gb4 = df_p.groupby("EMPRESA")["Pago final"].sum().reset_index()
            fig4 = px.pie(gb4, names="EMPRESA", values="Pago final",
                          color_discrete_sequence=[COL_AZUL, COL_AMARILLO, "#26A69A", COL_ROJO],
                          hole=0.45)
            plotly_layout(fig4, "Gasto por Empresa/Contratista", 420)
            st.plotly_chart(fig4, use_container_width=True)

    # Tabla detallada
    st.markdown("##### 📋 Detalle de Planilla")
    df_show = df_p.copy()
    df_show["Fecha"] = df_show["Fecha"].dt.strftime("%d/%m/%Y")
    st.dataframe(df_show, use_container_width=True, height=320)
    st.download_button("⬇️ Descargar Planilla", to_excel_bytes(df_show), "planilla_eol.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
def main():
    # ── Carga inicial de datos (desde session_state o archivo por defecto)
    if "raw_data" not in st.session_state:
        # Buscar archivo por defecto
        default_paths = [
            "data/Eol_datos_febrero_2026.zip",
            "data/eol_datos.zip",
        ]
        loaded = False
        for p in default_paths:
            if os.path.exists(p):
                with st.spinner("⏳ Cargando datos iniciales..."):
                    with open(p, "rb") as f:
                        st.session_state["raw_data"] = load_data(f.read(), os.path.basename(p))
                loaded = True
                break
        if not loaded:
            st.session_state["raw_data"] = {}

    data = st.session_state["raw_data"]

    # ── Header principal
    logo_b64 = get_logo_b64()
    head_col1, head_col2 = st.columns([1, 5])
    with head_col1:
        if logo_b64:
            st.markdown(
                f'<img src="data:image/png;base64,{logo_b64}" style="max-width:130px;margin-top:8px">',
                unsafe_allow_html=True)
        else:
            st.markdown(
                f'<div style="font-size:2.5rem;font-weight:900;color:{COL_AMARILLO};line-height:1">eol.<br>'
                f'<span style="font-size:0.8rem;color:{COL_AZUL};font-weight:400;letter-spacing:0.15em">TRANSPORTES</span></div>',
                unsafe_allow_html=True)
    with head_col2:
        st.markdown(
            f'<h1 style="color:{COL_AZUL};margin:0;padding-top:10px">Dashboard de Gestión Logística</h1>'
            f'<p style="color:{COL_GRIS};margin:0">Período: Febrero 2026 | Cliente: AJE | Actualizado: {datetime.now().strftime("%d/%m/%Y %H:%M")}</p>',
            unsafe_allow_html=True)

    st.markdown(f'<hr style="border-color:{COL_AMARILLO};margin:8px 0 16px">', unsafe_allow_html=True)

    # ── Cargador principal (pantalla central si no hay datos)
    if not data:
        st.markdown(f"""
        <div style='background:{COL_BLANCO};border:2px solid {COL_AMARILLO};border-radius:14px;
                    padding:30px 40px;max-width:600px;margin:40px auto;text-align:center'>
          <div style='font-size:3rem'>📂</div>
          <h2 style='color:{COL_AZUL}'>Cargar datos EOL Perú</h2>
          <p style='color:#666'>Sube el archivo <b>.zip</b> con los 4 Excels de Febrero 2026<br>
          para activar el dashboard completo.</p>
        </div>
        """, unsafe_allow_html=True)
        col_a, col_b, col_c = st.columns([1,2,1])
        with col_b:
            uploaded_main = st.file_uploader(
                "📤 Selecciona tu archivo ZIP o Excel",
                type=["zip","xlsx"],
                key="uploader_main",
                help="Sube el ZIP con los 4 archivos Excel de EOL Perú"
            )
            if uploaded_main:
                with st.spinner("⏳ Procesando datos..."):
                    try:
                        new_data = load_data(uploaded_main.read(), uploaded_main.name)
                        st.session_state["raw_data"] = new_data
                        st.success("✅ ¡Datos cargados! Recargando dashboard...")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error al cargar: {e}")
        st.stop()

    fi, ff, placas_sel = render_sidebar(data)

    # ── Navegación por secciones
    secciones = {
        "📊 A — Financiero": "a",
        "🛣️ B — Kilometraje": "b",
        "⛽ C — Combustible": "c",
        "⚖️ D — Carga": "d",
        "🗺️ E — Tiempos": "e",
        "👷 F — Planilla": "f",
    }

    tabs = st.tabs(list(secciones.keys()))

    with tabs[0]: seccion_a(data, fi, ff, placas_sel)
    with tabs[1]: seccion_b(data, fi, ff, placas_sel)
    with tabs[2]: seccion_c(data, fi, ff, placas_sel)
    with tabs[3]: seccion_d(data, fi, ff, placas_sel)
    with tabs[4]: seccion_e(data, fi, ff, placas_sel)
    with tabs[5]: seccion_f(data, fi, ff, placas_sel)


if __name__ == "__main__":
    main()

