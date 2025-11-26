# ======================= ABEX AI — PETRONAS Glassy UI (Streamlit) =======================
# Square UI (no rounded corners), white/black text only, PETRONAS teal accents.
# Background image loaded via robust resolver (Option 3). Works whether image is at:
#   - assets/teal-bg.jpg (next to Home.py)
#   - pages/assets/teal-bg.jpg (inside pages/)
#   - <this_file>/assets/teal-bg.jpg (next to this page file)
# Tabs are square with teal hover. Floating back-arrow toggles the sidebar.
# ----------------------------------------------------------------------------------------
import os
import base64
import pathlib
from pathlib import Path

import streamlit as st
import pandas as pd
import numpy as np
import io
import requests
from PIL import Image

# ML/Stats
from sklearn.impute import KNNImputer
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import MinMaxScaler
from sklearn.ensemble import RandomForestRegressor
from sklearn.metrics import mean_squared_error, r2_score
from scipy.stats import linregress

# Viz
import plotly.express as px
import plotly.graph_objects as go

# ---------------------------
# Page Setup (sidebar collapsed)
# ---------------------------
st.set_page_config(
    page_title="ABEX AI RT2025",
    page_icon="💠",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ---------------------------
# PETRONAS Theme / Tokens
# ---------------------------
PETRONAS = {
    "bg": "#FFFFFF",                       # base canvas (fallback if image missing)
    "panel": "rgba(255,255,255,0.70)",     # translucent glass panel
    "text_black": "#000000",
    "text_white": "#FFFFFF",
    "border": "rgba(0,0,0,0.08)",
    "primary": "#00A19C",                  # PETRONAS teal
    "primary_hover": "#008B86",            # darker teal
    "shadow": "0 8px 28px rgba(0,0,0,0.12)",
}
RADIUS = 0  # strictly square (no rounded UI)

# ---------------------------
# Logo loader (optional)
# ---------------------------
def _load_logo_base64():
    """Return (mime, base64str) for logo from assets folder (svg or png)."""
    for fname in ("petronas_logo.svg", "petronas_logo.png"):
        fpath = Path("assets") / fname
        if fpath.exists():
            data = fpath.read_bytes()
            if fname.endswith(".svg"):
                return "image/svg+xml", base64.b64encode(data).decode("utf-8")
            else:
                return "image/png", base64.b64encode(data).decode("utf-8")
    return None, None

# ---------------------------
# Option 3: Robust background resolver
# ---------------------------
def resolve_bg_path():
    """Try common locations to find teal-bg.jpg and return a Path or None."""
    candidates = [
        Path("assets/teal-bg.jpg"),                   # next to Home.py
        Path("pages/assets/teal-bg.jpg"),            # under pages/
        Path(__file__).parent / "assets" / "teal-bg.jpg",  # next to this page file
    ]
    for p in candidates:
        try:
            if p.exists():
                return p
        except Exception:
            continue
    return None

BG_PATH = resolve_bg_path()

# Sidebar diagnostics
with st.sidebar:
    st.caption("Background image status")
    if not BG_PATH or not BG_PATH.exists():
        st.error("❌ Image not found in common locations")
        st.code("assets/teal-bg.jpg\npages/assets/teal-bg.jpg\n<this_page>/assets/teal-bg.jpg")
        st.caption("Run the app from the project root (same folder as Home.py).\nPlace teal-bg.jpg in one of the paths above.")
    else:
        try:
            img = Image.open(BG_PATH)
            st.success(f"✅ Image found at: {BG_PATH}")
            st.image(img, caption="Preview: teal-bg.jpg", use_column_width=True)
        except Exception as e:
            st.error(f"⚠️ Image exists but cannot be opened: {e}")
            st.caption("Check file permissions/format or re-save as JPG.")

# ---------------------------
# Global CSS (glassy + square tabs + background image via resolved path)
# ---------------------------
def inject_global_css():
    # Build a CSS-friendly relative path from CWD to BG_PATH
    if BG_PATH and BG_PATH.exists():
        try:
            rel = BG_PATH.resolve().relative_to(Path.cwd().resolve())
            bg_url = str(rel).replace(os.sep, "/")
        except Exception:
            # Fallback to relpath if not under CWD
            rel = os.path.relpath(str(BG_PATH), start=str(Path.cwd()))
            bg_url = rel.replace(os.sep, "/")
        bg_css = f"background: url('{bg_url}') center/cover no-repeat fixed !important;"
    else:
        bg_css = f"background: {PETRONAS['bg']} !important;"

    css = f"""
    <style>
      :root {{
        --pet-bg: {PETRONAS["bg"]};
        --pet-panel: {PETRONAS["panel"]};
        --pet-text: {PETRONAS["text_black"]};
        --pet-text-invert: {PETRONAS["text_white"]};
        --pet-primary: {PETRONAS["primary"]};
        --pet-primary-hover: {PETRONAS["primary_hover"]};
        --pet-border: {PETRONAS["border"]};
        --pet-shadow: {PETRONAS["shadow"]};
      }}

      /* App background via resolved file path (Option 3) */
      .stApp {{
        position: relative;
        {bg_css}
        color: var(--pet-text) !important;
      }}

      /* Optional readability overlay (tune the alpha if needed) */
      .stApp::before {{
        content: "";
        position: fixed; inset: 0;
        background: rgba(255,255,255,0.60);   /* try 0.35–0.50 if you want more image */
        backdrop-filter: blur(2px);
        pointer-events: none;
        z-index: 0;
      }}
      .stApp > div:first-child {{
        position: relative;
        z-index: 1;
      }}

      /* Square corners globally */
      * {{ border-radius: {RADIUS}px !important; }}

      /* Glass panels */
      [data-testid="stVerticalBlock"] > div,
      .st-expander, .stDataFrame, .stMarkdown, .stAlert {{
        background: var(--pet-panel) !important;
        backdrop-filter: blur(10px);
        -webkit-backdrop-filter: blur(10px);
        border: 1px solid var(--pet-border) !important;
        box-shadow: var(--pet-shadow);
      }}

      /* Top bar */
      .pet-topbar {{
        position: sticky; top: 0; z-index: 50;
        background: var(--pet-panel);
        backdrop-filter: blur(10px);
        border-bottom: 1px solid var(--pet-border);
        box-shadow: var(--pet-shadow);
        padding: 10px 16px;
        display: grid; grid-template-columns: 40px 1fr auto; gap: 12px; align-items: center;
      }}
      .pet-brand {{ font-weight: 700; letter-spacing: .2px; }}
      .pet-sub   {{ font-size: 12px; color: #000; opacity: .8; }}

      /* Buttons (square, glassy, lift on hover) */
      .pet-btn a, .pet-btn button {{
        display: inline-block;
        background: var(--pet-panel);
        color: var(--pet-text) !important;
        border: 1px solid var(--pet-border);
        box-shadow: var(--pet-shadow);
        text-decoration: none;
        padding: 10px 14px;
        font-weight: 600;
      }}
      .pet-btn a:hover, .pet-btn button:hover {{
        background: var(--pet-primary);
        color: var(--pet-text-invert) !important;
        transform: translateY(-1px);
      }}
      .pet-btn a:focus-visible, .pet-btn button:focus-visible {{
        outline: 3px solid var(--pet-primary);
        outline-offset: 2px;
      }}

      /* Metrics */
      [data-testid="stMetric"] {{
        background: var(--pet-panel) !important;
        backdrop-filter: blur(10px);
        border: 1px solid var(--pet-border);
        box-shadow: var(--pet-shadow);
        padding: 8px;
      }}
      [data-testid="stMetric"] * {{ color: var(--pet-text) !important; }}

      /* Tabs: square + hover effect (teal on hover/active) */
      .stTabs [role="tablist"] {{
        border-bottom: none !important;
        gap: 8px;
      }}
      .stTabs [role="tab"] {{
        background: rgba(255,255,255,0.85) !important; /* neutral glass */
        color: var(--pet-text) !important;
        border: 1px solid var(--pet-border) !important;
        border-radius: 0 !important;                    /* square */
        padding: 10px 14px !important;
        box-shadow: none !important;
        transition: background .15s ease, box-shadow .15s ease, border-color .15s ease;
      }}
      .stTabs [role="tab"]:hover {{
        background: var(--pet-primary) !important;
        color: var(--pet-text-invert) !important;
        border-color: var(--pet-primary-hover) !important;
        box-shadow: 0 6px 18px rgba(0,0,0,0.16) !important;
      }}
      .stTabs [aria-selected="true"] {{
        background: var(--pet-primary) !important;
        color: var(--pet-text-invert) !important;
        border: 2px solid var(--pet-primary-hover) !important;
        box-shadow: 0 6px 18px rgba(0,0,0,0.18) !important;
      }}
      .stTabs [role="tab"]:focus-visible {{
        outline: 3px solid var(--pet-primary) !important;
        outline-offset: 2px !important;
      }}

      /* Inputs */
      input, textarea, select, .stTextInput, .stNumberInput, .stSlider, .stSelectbox {{
        background: rgba(255,255,255,0.95) !important;
        border: 1px solid var(--pet-border) !important;
        box-shadow: none !important;
        color: var(--pet-text) !important;
      }}

      /* Plotly */
      .plotly, .js-plotly-plot {{ background: #FFFFFF !important; }}

      /* Sidebar glass */
      [data-testid="stSidebar"] > div:first-child {{
        background: var(--pet-panel) !important;
        backdrop-filter: blur(10px);
        border-right: 1px solid var(--pet-border);
        box-shadow: var(--pet-shadow);
      }}

      /* Floating back-arrow */
      .pet-toggle {{
        position: fixed;
        left: 8px; top: 68px; z-index: 1000;
        width: 40px; height: 40px;
        display: grid; place-items: center;
        background: var(--pet-panel);
        border: 1px solid var(--pet-border);
        box-shadow: var(--pet-shadow);
        cursor: pointer;
        user-select: none;
      }}
      .pet-toggle:hover {{
        background: var(--pet-primary);
        color: var(--pet-text-invert);
      }}
    </style>
    """
    st.markdown(css, unsafe_allow_html=True)

# ---------------------------
# Top bar
# ---------------------------
def render_topbar():
    mime, b64 = _load_logo_base64()
    if b64:
        logo_html = f'data:{mime};base64,{b64}'
    else:
        # Fallback teal square
        logo_html = '<div style="width:32px;height:32px;background:#00A19C;"></div>'

    bar = f"""
    <div class="pet-topbar">
      <div>{logo_html}</div>
      <div>
        <div class="pet-brand">ABEX AI RT2025</div>
        <div class="pet-sub">Data‑driven cost prediction</div>
      </div>
      <div class="pet-btn"></div>
    </div>
    """
    st.markdown(bar, unsafe_allow_html=True)

# ---------------------------
# Sidebar toggle (floating arrow)
# ---------------------------
def render_sidebar_toggle():
    html = """
    <div class="pet-toggle" id="pet-toggle" title="Toggle sidebar">
      &#x25C0;
    </div>
    <script>
      const btn = document.getElementById('pet-toggle');
      function toggleSidebar() {
        const root = window.parent.document.querySelector('[data-testid="stSidebar"]');
        if (!root) return;
        const ctrl = window.parent.document.querySelector('[data-testid="collapsedControl"]');
        if (ctrl) { ctrl.click(); return; }
        const expanded = root.getAttribute('aria-expanded');
        if (expanded === 'true') {
          root.setAttribute('aria-expanded', 'false');
          root.style.transform = 'translateX(-100%)';
        } else {
          root.setAttribute('aria-expanded', 'true');
          root.style.transform = 'translateX(0)';
        }
      }
      btn.addEventListener('click', toggleSidebar);
    </script>
    """
    st.markdown(html, unsafe_allow_html=True)

# ---------------------------
# Glassy link button helper (fixed anchor)
# ---------------------------
def link_button(label: str, url: str, new_tab: bool = True):
    target = ' target="_blank" rel="noopener noreferrer"' if new_tab else ""
    html = f'<div class="pet-btn"><a href="{url}"{target}>{label}</a></div>'
    st.markdown(html, unsafe_allow_html=True)

# ---------------------------
# Apply UI
# ---------------------------
inject_global_css()
render_topbar()
render_sidebar_toggle()

# ---------------------------
# Utility helpers
# ---------------------------
def human_format(num, pos=None):
    try:
        num = float(num)
    except Exception:
        return str(num)
    if num >= 1e9: return f'{num/1e9:.1f}B'
    if num >= 1e6: return f'{num/1e6:.1f}M'
    if num >= 1e3: return f'{num/1e3:.1f}K'
    return f'{num:.0f}'

def format_with_commas(num):
    try:
        return f"{float(num):,.2f}"
    except Exception:
        return str(num)

def get_currency_symbol(df: pd.DataFrame):
    for col in df.columns:
        uc = col.upper()
        if "RM" in uc: return "RM"
        if "USD" in uc or "$" in col: return "USD"
        if "€" in col: return "€"
        if "£" in col: return "£"
    try:
        sample_vals = df.iloc[:20].astype(str).values.flatten().tolist()
        if any("RM" in v.upper() for v in sample_vals): return "RM"
        if any("€" in v for v in sample_vals): return "€"
        if any("£" in v for v in sample_vals): return "£"
        if any("$" in v for v in sample_vals): return "USD"
    except Exception:
        pass
    return ""

def toast(msg, icon="✅"):
    try:
        st.toast(f"{icon} {msg}")
    except:
        st.success(msg if icon == "✅" else msg)

# ---------------------------
# Remote data manifest (GitHub)
# ---------------------------
GITHUB_USER = "Hafizuddin-Abd-Rahman-Dev-Upstream"
REPO_NAME   = "Cost-Predictor"
BRANCH      = "main"
DATA_FOLDER = "pages/data_ABEX"

@st.cache_data(ttl=600)
def list_csvs_from_manifest(folder_path: str):
    manifest_url = f"https://raw.githubusercontent.com/{GITHUB_USER}/{REPO_NAME}/{BRANCH}/{folder_path}/files.json"
    try:
        res = requests.get(manifest_url, timeout=10)
        res.raise_for_status()
        return res.json()
    except Exception as e:
        st.error(f"Failed to load CSV manifest: {e}")
        return []

# ---------------------------
# Modeling
# ---------------------------
def build_model(df: pd.DataFrame):
    imputed = pd.DataFrame(KNNImputer(n_neighbors=5).fit_transform(df), columns=df.columns)
    X = imputed.iloc[:, :-1]
    y = imputed.iloc[:, -1]
    scaler = MinMaxScaler().fit(X)
    model = RandomForestRegressor(random_state=42).fit(scaler.transform(X), y)
    return dict(imputed=imputed, X=X, y=y, scaler=scaler, model=model, target=y.name)

def evaluate_model(X, y, test_size=0.2):
    Xtr, Xte, ytr, yte = train_test_split(X, y, test_size=test_size, random_state=42)
    scaler = MinMaxScaler().fit(Xtr)
    model = RandomForestRegressor(random_state=42).fit(scaler.transform(Xtr), ytr)
    yhat = model.predict(scaler.transform(Xte))
    rmse = float(np.sqrt(mean_squared_error(yte, yhat)))
    r2 = float(r2_score(yte, yhat))
    return dict(model=model, scaler=scaler, rmse=rmse, r2=r2)

def single_prediction(X, y, payload: dict):
    scaler = MinMaxScaler().fit(X)
    model = RandomForestRegressor(random_state=42).fit(scaler.transform(X), y)
    cols = list(X.columns)
    row = {c: np.nan for c in cols}
    for c, v in payload.items():
        try:
            row[c] = float(v) if (v is not None and str(v).strip() != "") else np.nan
        except:
            row[c] = np.nan
    df_in = pd.DataFrame([row], columns=cols)
    pred = float(model.predict(scaler.transform(df_in))[0])
    return pred

def cost_breakdown(base_pred: float, eprr: dict, sst_pct: float, owners_pct: float, cont_pct: float, esc_pct: float):
    owners_cost = round(base_pred * (owners_pct / 100.0), 2)
    sst_cost = round(base_pred * (sst_pct / 100.0), 2)
    contingency_cost = round((base_pred + owners_cost) * (cont_pct / 100.0), 2)
    escalation_cost  = round((base_pred + owners_cost) * (esc_pct / 100.0), 2)
    eprr_costs = {k: round(base_pred * (v / 100.0), 2) for k, v in (eprr or {}).items()}
    grand_total = round(base_pred + owners_cost + contingency_cost + escalation_cost, 2)
    return owners_cost, sst_cost, contingency_cost, escalation_cost, eprr_costs, grand_total

# ---------------------------
# Session state
# ---------------------------
if "authenticated" not in st.session_state: st.session_state.authenticated = False
if "datasets"     not in st.session_state: st.session_state.datasets     = {}
if "predictions"  not in st.session_state: st.session_state.predictions  = {}
if "processed_excel_files" not in st.session_state: st.session_state.processed_excel_files = set()

# ---------------------------
# Minimal Auth (wire to st.secrets in deployment)
# ---------------------------
APPROVED_EMAILS = st.secrets.get("emails", [])
correct_password = st.secrets.get("password", None)

if not st.session_state.authenticated:
    with st.container():
        with st.form("login"):
            st.markdown('<div class="pet-brand">Sign in</div><div class="pet-sub">Secure</div>', unsafe_allow_html=True)
            c1, c2 = st.columns([2, 1])
            with c1:
                email = st.text_input("Corporate Email", placeholder="name@company.com")
            with c2:
                password = st.text_input("Access Password", type="password", placeholder="••••••••")
            submitted = st.form_submit_button("Sign in")
            st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
        if submitted:
            if email in APPROVED_EMAILS and password == correct_password:
                st.session_state.authenticated = True
                toast("Access granted.")
                st.rerun()
            else:
                st.error("Invalid email or password. Contact Cost Engineering Focal.")
    st.stop()

# ---------------------------
# Tabs
# ---------------------------
tab_data, tab_model, tab_viz, tab_predict, tab_results = st.tabs(
    ["📁 Data", "⚙️ Model", "📈 Visualization", "🎯 Predict", "📄 Results"]
)

# ---------------------------
# DATA TAB
# ---------------------------
with tab_data:
    with st.container():
        st.markdown('<div class="pet-brand">Data Sources</div><div class="pet-sub">Step 1</div>', unsafe_allow_html=True)
        c1, c2 = st.columns([1.2, 1])
        with c1:
            data_source = st.radio("Choose data source", ["Upload CSV", "Load from Server"], horizontal=True)
        with c2:
            st.caption("Enterprise Storage")
            st.write("")
        data_link = "https://petronas.sharepoint.com/sites/ecm_ups_coe/confidential/DFE%20Cost%20Engineering/Forms/AllItems.aspx?id=%2Fsites%2Fecm%5Fups%5Fcoe%2Fconfidential%2FDFE%20Cost%20Engineering%2F2%2ETemplate%20Tools%2FCost%20Predictor%2FDatabase%2FABEX%20%28DDRR%29%20%2D%20RT%20Q1%202025&viewid=25092e6d%2D373d%2D41fe%2D8f6f%2D486cd8cdd5b8"
        link_button("Open Enterprise Storage", data_link, new_tab=True)

        uploaded_files = []
        if data_source == "Upload CSV":
            uploaded_files = st.file_uploader("Upload CSV files (max 200MB)", type="csv", accept_multiple_files=True)
        else:
            github_csvs = list_csvs_from_manifest(DATA_FOLDER)
            if github_csvs:
                selected_file = st.selectbox("Choose CSV from GitHub", github_csvs)
                if st.button("Load selected CSV"):
                    raw_url = f"https://raw.githubusercontent.com/{GITHUB_USER}/{REPO_NAME}/{BRANCH}/{DATA_FOLDER}/{selected_file}"
                    try:
                        df = pd.read_csv(raw_url)
                        fake = type("FakeUpload", (), {"name": selected_file})
                        uploaded_files = [fake]
                        st.session_state.datasets[selected_file] = df
                        st.session_state.predictions.setdefault(selected_file, [])
                        toast(f"Loaded from GitHub: {selected_file}")
                    except Exception as e:
                        st.error(f"Error loading CSV: {e}")
            else:
                st.info("No CSV files found in GitHub folder.")

        # Ingest uploads
        if uploaded_files:
            for up in uploaded_files:
                if up.name not in st.session_state.datasets:
                    if hasattr(up, "read"):
                        df = pd.read_csv(up)
                    else:
                        df = st.session_state.datasets.get(up.name, None)
                    if df is not None:
                        st.session_state.datasets[up.name] = df
                        st.session_state.predictions.setdefault(up.name, [])
            toast("Dataset(s) added.")

        # Actions
        st.divider()
        c1, c2, c3 = st.columns([1, 1, 2])
        with c1:
            if st.button("🧹 Clear all predictions"):
                st.session_state.predictions = {}
                toast("All predictions cleared.", "🧹")
        with c2:
            if st.button("🧼 Clear processed files history"):
                st.session_state.processed_excel_files = set()
                toast("Processed files history cleared.", "🧼")
        with c3:
            if st.button("🔁 Refresh server manifest"):
                list_csvs_from_manifest.clear()
                toast("Server manifest refreshed.", "🔁")

        st.divider()

        # Dataset select
        if st.session_state.datasets:
            ds_name = st.selectbox("Active dataset", list(st.session_state.datasets.keys()))
            df = st.session_state.datasets[ds_name]
            currency = get_currency_symbol(df)
            colA, colB, colC = st.columns([1, 1, 1])
            with colA: st.metric("Rows", f"{df.shape[0]:,}")
            with colB: st.metric("Columns", f"{df.shape[1]:,}")
            with colC: st.metric("Currency", f"{currency or '—'}")

            with st.expander("Preview (first 10 rows)", expanded=False):
                st.dataframe(df.head(10), use_container_width=True)
        else:
            st.info("Upload or load a dataset to proceed.")
        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

# ---------------------------
# MODEL TAB
# ---------------------------
with tab_model:
    if not st.session_state.datasets:
        st.info("No dataset. Go to **Data** tab to upload or load.")
    else:
        ds_name = st.selectbox("Dataset for model training", list(st.session_state.datasets.keys()), key="ds_model")
        df = st.session_state.datasets[ds_name]
        with st.spinner("Imputing & preparing..."):
            imputed = pd.DataFrame(KNNImputer(n_neighbors=5).fit_transform(df), columns=df.columns)
            X = imputed.iloc[:, :-1]
            y = imputed.iloc[:, -1]
            target_column = y.name
        with st.container():
            st.markdown('<div class="pet-brand">Train & Evaluate</div><div class="pet-sub">Step 2</div>', unsafe_allow_html=True)
            c1, c2 = st.columns([1, 3])
            with c1:
                test_size = st.slider("Test size", 0.1, 0.5, 0.2, 0.05, help="Fraction of data used for testing")
                run = st.button("Run training")
            with c2:
                st.caption("Random Forest on min‑max scaled features; reproducible with random_state=42.")
                st.write("")
            if run:
                with st.spinner("Training model..."):
                    metrics = evaluate_model(X, y, test_size=test_size)
                    c1, c2 = st.columns(2)
                    with c1: st.metric("RMSE", f"{metrics['rmse']:,.2f}")
                    with c2: st.metric("R²", f"{metrics['r2']:.3f}")
                    st.session_state["_last_metrics"] = metrics
                    toast("Training complete.")
        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

# ---------------------------
# VISUALIZATION TAB
# ---------------------------
with tab_viz:
    if not st.session_state.datasets:
        st.info("No dataset. Go to **Data** tab to upload or load.")
    else:
        ds_name = st.selectbox("Dataset for visualization", list(st.session_state.datasets.keys()), key="ds_viz")
        df = st.session_state.datasets[ds_name]
        imputed = pd.DataFrame(KNNImputer(n_neighbors=5).fit_transform(df), columns=df.columns)
        X = imputed.iloc[:, :-1]
        y = imputed.iloc[:, -1]
        target_column = y.name

        # Correlation Matrix
        with st.container():
            st.markdown('<div class="pet-brand">Correlation Matrix</div><div class="pet-sub">Exploration</div>', unsafe_allow_html=True)
            corr = imputed.corr(numeric_only=True)
            fig = px.imshow(corr, text_auto=".2f", aspect="auto", color_continuous_scale="RdBu_r", zmin=-1, zmax=1)
            fig.update_layout(
                margin=dict(l=0, r=0, t=10, b=0),
                paper_bgcolor="#FFFFFF",
                plot_bgcolor="#FFFFFF",
                font=dict(color="#000000"),
                xaxis=dict(color="#000000"),
                yaxis=dict(color="#000000"),
                legend=dict(font=dict(color="#000000"))
            )
            st.plotly_chart(fig, use_container_width=True)
        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

        # Feature Importance
        with st.container():
            st.markdown('<div class="pet-brand">Feature Importance</div><div class="pet-sub">Model</div>', unsafe_allow_html=True)
            scaler = MinMaxScaler().fit(X)
            model = RandomForestRegressor(random_state=42).fit(scaler.transform(X), y)
            importances = model.feature_importances_
            fi = pd.DataFrame({"feature": X.columns, "importance": importances}).sort_values("importance", ascending=True)
            fig = go.Figure(go.Bar(x=fi["importance"], y=fi["feature"], orientation='h', marker_color=PETRONAS["primary"]))
            fig.update_layout(
                xaxis_title="Importance", yaxis_title="Feature",
                margin=dict(l=0, r=0, t=10, b=0),
                paper_bgcolor="#FFFFFF", plot_bgcolor="#FFFFFF",
                font=dict(color="#000000"),
                xaxis=dict(color="#000000"), yaxis=dict(color="#000000"),
                legend=dict(font=dict(color="#000000"))
            )
            st.plotly_chart(fig, use_container_width=True)
        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

        # Cost Curve (robust linregress)
        with st.container():
            st.markdown('<div class="pet-brand">Cost Curve</div><div class="pet-sub">Trend</div>', unsafe_allow_html=True)
            feat = st.selectbox("Select feature for cost curve", X.columns)
            x_vals = imputed[feat].values
            y_vals = y.values
            mask = (~np.isnan(x_vals)) & (~np.isnan(y_vals))
            scatter_df = pd.DataFrame({feat: x_vals[mask], target_column: y_vals[mask]})
            fig = px.scatter(scatter_df, x=feat, y=target_column, opacity=0.65)
            fig.update_traces(marker=dict(color=PETRONAS["primary"]))
            if mask.sum() >= 2 and np.unique(x_vals[mask]).size >= 2:
                xv = scatter_df[feat].to_numpy(dtype=float)
                yv = scatter_df[target_column].to_numpy(dtype=float)
                slope, intercept, r_value, p_value, std_err = linregress(xv, yv)
                x_line = np.linspace(xv.min(), xv.max(), 100)
                y_line = slope * x_line + intercept
                fig.add_trace(go.Scatter(
                    x=x_line, y=y_line, mode="lines",
                    name=f"Fit: y={slope:.2f}x+{intercept:.2f} (R²={r_value**2:.3f})",
                    line=dict(color=PETRONAS["primary"]) 
                ))
            else:
                st.warning("Not enough valid/variable data to compute regression.")
            fig.update_layout(
                margin=dict(l=0, r=0, t=10, b=0),
                paper_bgcolor="#FFFFFF", plot_bgcolor="#FFFFFF",
                font=dict(color="#000000"),
                xaxis=dict(color="#000000"), yaxis=dict(color="#000000"),
                legend=dict(font=dict(color="#000000"))
            )
            st.plotly_chart(fig, use_container_width=True)
        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

# ---------------------------
# PREDICT TAB
# ---------------------------
with tab_predict:
    if not st.session_state.datasets:
        st.info("No dataset. Go to **Data** tab to upload or load.")
    else:
        ds_name = st.selectbox("Dataset for prediction", list(st.session_state.datasets.keys()), key="ds_pred")
        df = st.session_state.datasets[ds_name]
        currency = get_currency_symbol(df)

        # Build model on full imputed data for on-demand prediction
        built = build_model(df)
        X, y, scaler, model, target_column = built["X"], built["y"], built["scaler"], built["model"], built["target"]

        # Config card
        with st.container():
            st.markdown('<div class="pet-brand">Configuration (EPRR • Taxes • Owner • Risk)</div><div class="pet-sub">Step 3</div>', unsafe_allow_html=True)
            c1, c2 = st.columns([1, 1])
            with c1:
                st.markdown("**EPRR Breakdown (%)**")
                eng = st.slider("Engineering", 0, 100, 12)
                prep = st.slider("Preparation", 0, 100, 7)
                remv = st.slider("Removal", 0, 100, 54)
                remd = st.slider("Remediation", 0, 100, 27)
            with c2:
                st.markdown("**Financial (%)**")
                sst_pct    = st.slider("SST", 0, 100, 0)
                owners_pct = st.slider("Owner's Cost", 0, 100, 0)
                cont_pct   = st.slider("Contingency", 0, 100, 0)
                esc_pct    = st.slider("Escalation & Inflation", 0, 100, 0)

            eprr = {"Engineering": eng, "Preparation": prep, "Removal": remv, "Remediation": remd}
            eprr_total = sum(eprr.values())
            if abs(eprr_total - 100) > 1e-6 and eprr_total > 0:
                st.warning(f"EPRR total is {eprr_total}%. Consider normalizing to 100% for reporting consistency.")
        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

        # Single prediction
        with st.container():
            st.markdown('<div class="pet-brand">Predict (Single)</div><div class="pet-sub">Step 4</div>', unsafe_allow_html=True)
            project_name = st.text_input("Project Name", placeholder="e.g., Offshore Pipeline Replacement 2025")
            st.caption("Provide feature values (leave blank for NaN).")
            cols_per_row = 3
            new_data = {}
            cols = list(X.columns)
            rows = (len(cols) + cols_per_row - 1) // cols_per_row
            for r in range(rows):
                row_cols = st.columns(cols_per_row)
                for i in range(cols_per_row):
                    idx = r * cols_per_row + i
                    if idx < len(cols):
                        col_name = cols[idx]
                        with row_cols[i]:
                            val = st.text_input(col_name, key=f"in_{col_name}")
                            new_data[col_name] = val

            if st.button("Run Prediction"):
                pred = single_prediction(X, y, new_data)
                owners_cost, sst_cost, contingency_cost, escalation_cost, eprr_costs, grand_total = cost_breakdown(
                    pred, eprr, sst_pct, owners_pct, cont_pct, esc_pct
                )
                # Store
                result = {"Project Name": project_name, **{c: new_data[c] for c in cols}, target_column: round(pred, 2)}
                for k, v in eprr_costs.items(): result[f"{k} Cost"] = v
                result["SST Cost"] = sst_cost
                result["Owner's Cost"] = owners_cost
                result["Cost Contingency"] = contingency_cost
                result["Escalation & Inflation"] = escalation_cost
                result["Grand Total"] = grand_total
                st.session_state.predictions.setdefault(ds_name, []).append(result)
                toast("Prediction added to Results.")

                # Summary metrics
                cA, cB, cC, cD, cE = st.columns(5)
                with cA: st.metric("Predicted",         f"{currency} {pred:,.2f}")
                with cB: st.metric("Owner's",           f"{currency} {owners_cost:,.2f}")
                with cC: st.metric("Contingency",       f"{currency} {contingency_cost:,.2f}")
                with cD: st.metric("Escalation",        f"{currency} {escalation_cost:,.2f}")
                with cE: st.metric("Grand Total",       f"{currency} {grand_total:,.2f}")
        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

        # Batch (Excel)
        with st.container():
            st.markdown('<div class="pet-brand">Batch (Excel)</div>', unsafe_allow_html=True)
            xls = st.file_uploader("Upload Excel for batch prediction", type=["xlsx"])
            if xls:
                file_id = f"{xls.name}_{xls.size}_{ds_name}"
                if file_id not in st.session_state.processed_excel_files:
                    batch_df = pd.read_excel(xls)
                    missing = [c for c in X.columns if c not in batch_df.columns]
                    if missing:
                        st.error(f"Missing required columns in Excel: {missing}")
                    else:
                        scaler_b = MinMaxScaler().fit(X)
                        model_b  = RandomForestRegressor(random_state=42).fit(scaler_b.transform(X), y)
                        preds = model_b.predict(scaler_b.transform(batch_df[X.columns]))
                        batch_df[target_column] = preds
                        for i, row in batch_df.iterrows():
                            name = row.get("Project Name", f"Project {i+1}")
                            entry = {"Project Name": name}
                            entry.update(row[X.columns].to_dict())
                            entry[target_column] = round(float(preds[i]), 2)
                            owners_cost, sst_cost, contingency_cost, escalation_cost, eprr_costs, grand_total = cost_breakdown(
                                float(preds[i]), eprr, sst_pct, owners_pct, cont_pct, esc_pct
                            )
                            for k, v in eprr_costs.items(): entry[f"{k} Cost"] = v
                            entry["SST Cost"] = sst_cost
                            entry["Owner's Cost"] = owners_cost
                            entry["Cost Contingency"] = contingency_cost
                            entry["Escalation & Inflation"] = escalation_cost
                            entry["Grand Total"] = grand_total
                            st.session_state.predictions.setdefault(ds_name, []).append(entry)
                        st.session_state.processed_excel_files.add(file_id)
                        toast("Batch prediction complete.")
        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

# ---------------------------
# RESULTS TAB
# ---------------------------
with tab_results:
    if not st.session_state.datasets:
        st.info("No dataset. Go to **Data** tab to upload or load.")
    else:
        ds_name = st.selectbox("Dataset", list(st.session_state.datasets.keys()), key="ds_results")
        preds = st.session_state.predictions.get(ds_name, [])
        with st.container():
            st.markdown(f'<div class="pet-brand">Project Entries</div><div class="pet-sub">{len(preds)} saved</div>', unsafe_allow_html=True)
            if preds:
                if st.button("🗑️ Delete all entries"):
                    st.session_state.predictions[ds_name] = []
                    to_remove = {fid for fid in st.session_state.processed_excel_files if fid.endswith(ds_name)}
                    for fid in to_remove: st.session_state.processed_excel_files.remove(fid)
                    toast("All entries removed.", "🗑️")
                    st.rerun()
                names = [p.get("Project Name", f"Project {i+1}") for i, p in enumerate(preds)]
                st.write(", ".join(names) if names else "—")
            else:
                st.info("No predictions yet — add one from the **Predict** tab.")
        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

        with st.container():
            st.markdown('<div class="pet-brand">Summary Table & Export</div><div class="pet-sub">Download</div>', unsafe_allow_html=True)
            if preds:
                df_preds = pd.DataFrame(preds)
                df_disp = df_preds.copy()
                num_cols = df_disp.select_dtypes(include=[np.number]).columns
                for col in num_cols:
                    df_disp[col] = df_disp[col].apply(lambda x: format_with_commas(x))
                st.dataframe(df_disp, use_container_width=True, height=420)

                bio = io.BytesIO()
                df_preds.to_excel(bio, index=False, engine="openpyxl")
                bio.seek(0)
                st.download_button(
                    "⬇️ Download as Excel",
                    data=bio,
                    file_name=f"{ds_name}_predictions.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.info("No data to export yet.")
        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

# ---------------------------
# FOOTER
# ---------------------------
st.markdown(
    """
    <div style="padding:8px 0; color:#000;">
      © 2025 PETRONAS • Cost Engineering • ABEX AI RT2025
    </div>
    """,
    unsafe_allow_html=True
)
