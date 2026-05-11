# ======================================================================================
# CAPEX AI RT2026
# ======================================================================================

import io
import json
import zipfile
import re
import requests
import numpy as np
import pandas as pd
import streamlit as st

try:
    from sklearn.impute import KNNImputer, SimpleImputer
    from sklearn.model_selection import train_test_split
    from sklearn.preprocessing import StandardScaler
    from sklearn.ensemble import RandomForestRegressor, GradientBoostingRegressor
    from sklearn.pipeline import Pipeline
    from sklearn.metrics import mean_squared_error, r2_score, mean_absolute_error
except Exception as e:
    st.error(
        "Missing dependency: scikit-learn.\n\n"
        "Add scikit-learn to requirements.txt and redeploy.\n\n"
        f"Details: {e}"
    )
    st.stop()

try:
    import torch
    import torch.nn as nn
    from torch.utils.data import DataLoader, TensorDataset
    TORCH_AVAILABLE = True
except ImportError:
    TORCH_AVAILABLE = False

import plotly.express as px
import plotly.graph_objects as go

# ——————————————————————————————
# PAGE CONFIG
# ——————————————————————————————

st.set_page_config(
    page_title="CAPEX AI RT2026",
    page_icon="💠",
    layout="wide",
    initial_sidebar_state="expanded",
)

PETRONAS = {
    "teal":      "#00A19B",
    "teal_dark": "#008C87",
    "purple":    "#6C4DD3",
    "white":     "#FFFFFF",
    "black":     "#0E1116",
    "border":    "rgba(0,0,0,0.10)",
}

SHAREPOINT_LINKS = {
    "Shallow Water": "https://petronas.sharepoint.com/sites/your-site/shallow-water",
    "Deep Water":    "https://petronas.sharepoint.com/sites/your-site/deep-water",
    "Onshore":       "https://petronas.sharepoint.com/sites/your-site/onshore",
    "Uncon":         "https://petronas.sharepoint.com/sites/your-site/uncon",
    "CCS":           "https://petronas.sharepoint.com/sites/your-site/ccs",
}

st.markdown(
    f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html, body {{ font-family: 'Inter', sans-serif; }}
[data-testid="stAppViewContainer"] {{
  background: {PETRONAS["white"]}; color: {PETRONAS["black"]}; padding-top: 0.5rem;
}}
#MainMenu, footer {{ visibility: hidden; }}
[data-testid="stSidebar"] {{
  background: linear-gradient(180deg, {PETRONAS["teal"]} 0%, {PETRONAS["teal_dark"]} 100%) !important;
  color: #fff !important; border-top-right-radius: 16px; border-bottom-right-radius: 16px;
  box-shadow: 0 6px 20px rgba(0,0,0,0.15);
}}
[data-testid="stSidebar"] * {{ color: #fff !important; }}
[data-testid="collapsedControl"] {{
  position: fixed !important; top: 50% !important; left: 10px !important;
  transform: translateY(-50%) !important; z-index: 9999 !important;
}}
.petronas-hero {{
  border-radius: 20px; padding: 28px 32px; margin: 6px 0 18px 0; color: #fff;
  background: linear-gradient(135deg, {PETRONAS["teal"]}, {PETRONAS["purple"]}, {PETRONAS["black"]});
  background-size: 200% 200%;
  animation: heroGradient 8s ease-in-out infinite, fadeIn .8s ease-in-out, heroPulse 5s ease-in-out infinite;
  box-shadow: 0 10px 24px rgba(0,0,0,.12);
}}
@keyframes heroGradient {{
  0% {{ background-position: 0% 50%; }} 50% {{ background-position: 100% 50%; }} 100% {{ background-position: 0% 50%; }}
}}
@keyframes fadeIn {{ from {{ opacity: 0; transform: translateY(10px); }} to {{ opacity: 1; transform: translateY(0); }} }}
@keyframes heroPulse {{
  0%   {{ box-shadow: 0 0 16px rgba(0,161,155,0.45); }} 25%  {{ box-shadow: 0 0 26px rgba(108,77,211,0.55); }}
  50%  {{ box-shadow: 0 0 36px rgba(0,161,155,0.55); }} 75%  {{ box-shadow: 0 0 26px rgba(108,77,211,0.55); }}
  100% {{ box-shadow: 0 0 16px rgba(0,161,155,0.45); }}
}}
.petronas-hero h1 {{ margin: 0 0 5px; font-weight: 800; letter-spacing: 0.3px; }}
.petronas-hero p {{ margin: 0; opacity: .9; font-weight: 500; }}
.stButton > button, .stDownloadButton > button, .petronas-button {{
  border-radius: 10px; padding: .6rem 1.1rem; font-weight: 600; color: #fff !important; border: none;
  background: linear-gradient(to right, {PETRONAS["teal"]}, {PETRONAS["purple"]}); background-size: 200% auto;
  transition: background-position .85s ease, transform .2s ease, box-shadow .25s ease;
  text-decoration: none; display: inline-block;
}}
.stButton > button:hover, .stDownloadButton > button:hover, .petronas-button:hover {{
  background-position: right center; transform: translateY(-1px); box-shadow: 0 6px 16px rgba(0,0,0,0.18);
}}
.stTabs [role="tablist"] {{ display: flex; gap: 8px; border-bottom: none; padding-bottom: 6px; }}
.stTabs [role="tab"] {{
  background: #fff; color: {PETRONAS["black"]}; border-radius: 8px; padding: 10px 18px;
  border: 1px solid {PETRONAS["border"]}; font-weight: 600; transition: all .3s ease; position: relative;
}}
.stTabs [role="tab"]:hover {{ background: linear-gradient(to right, {PETRONAS["teal"]}, {PETRONAS["purple"]}); color: #fff; }}
.stTabs [role="tab"][aria-selected="true"] {{
  background: linear-gradient(to right, {PETRONAS["teal"]}, {PETRONAS["purple"]});
  color: #fff; border-color: transparent; box-shadow: 0 4px 16px rgba(0,0,0,0.15);
}}
.stTabs [role="tab"][aria-selected="true"]::after {{
  content: ''; position: absolute; left: 10%; bottom: -3px; width: 80%; height: 3px;
  background: linear-gradient(90deg, {PETRONAS["teal"]}, {PETRONAS["purple"]}, {PETRONAS["teal"]});
  background-size: 200% 100%; border-radius: 2px; animation: glowSlide 2.5s linear infinite;
}}
@keyframes glowSlide {{
  0% {{ background-position: 0% 50%; }} 50% {{ background-position: 100% 50%; }} 100% {{ background-position: 0% 50%; }}
}}
</style>
""",
    unsafe_allow_html=True,
)

st.markdown(
    """
<div class="petronas-hero">
  <h1>CAPEX AI RT2026</h1>
  <p>Data-driven CAPEX prediction &middot; RF &middot; GB &middot; MLP</p>
</div>
""",
    unsafe_allow_html=True,
)

# ——————————————————————————————
# AUTH
# ——————————————————————————————

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

APPROVED_EMAILS  = [str(e).strip().lower() for e in st.secrets.get("emails", [])]
correct_password = st.secrets.get("password", None)

if not st.session_state.authenticated:
    with st.form("login_form"):
        st.markdown("#### Access Required")
        email     = (st.text_input("Email Address", key="login_email") or "").strip().lower()
        password  = st.text_input("Access Password", type="password", key="login_pwd")
        submitted = st.form_submit_button("Login")
    if submitted:
        if (email in APPROVED_EMAILS) and (password == correct_password):
            st.session_state.authenticated = True
            st.success("Access granted.")
            st.rerun()
        else:
            st.error("Invalid credentials.")
    st.stop()

# ——————————————————————————————
# SESSION STATE
# ——————————————————————————————

for _k, _v in [
    ("datasets", {}), ("predictions", {}), ("processed_excel_files", set()),
    ("_last_metrics", None), ("projects", {}), ("component_labels", {}),
    ("uploader_nonce", 0), ("widget_nonce", 0),
]:
    if _k not in st.session_state:
        st.session_state[_k] = _v

# ——————————————————————————————
# HELPERS
# ——————————————————————————————

def toast(msg, icon="ok"):
    try:    st.toast(f"{msg}")
    except: st.success(msg)

def is_junk_col(c):
    h = str(c).strip().upper()
    return (not h) or h.startswith("UNNAMED") or h in {"INDEX", "IDX"}

def currency_from_header(h):
    h = (h or "").strip().upper()
    if "€" in h: return "€"
    if "£" in h: return "£"
    if "$" in h: return "$"
    if re.search(r"\bUSD\b", h): return "USD"
    if re.search(r"\b(MYR|RM)\b", h): return "RM"
    return ""

def get_currency_symbol(df, target_col=None):
    if df is None or df.empty: return ""
    if target_col and target_col in df.columns:
        return currency_from_header(str(target_col))
    for c in reversed(df.columns):
        if not is_junk_col(c):
            return currency_from_header(str(c))
    return ""

def cost_breakdown(base_pred, sst_pct, owners_pct, cont_pct, esc_pct):
    base_pred        = float(base_pred)
    owners_cost      = round(base_pred * (owners_pct / 100.0), 2)
    sst_cost         = round(base_pred * (sst_pct   / 100.0), 2)
    contingency_cost = round((base_pred + owners_cost) * (cont_pct / 100.0), 2)
    escalation_cost  = round((base_pred + owners_cost) * (esc_pct  / 100.0), 2)
    grand_total      = round(base_pred + owners_cost + sst_cost + contingency_cost + escalation_cost, 2)
    return owners_cost, sst_cost, contingency_cost, escalation_cost, grand_total

def project_components_df(proj):
    rows = []
    for c in proj.get("components", []):
        rows.append({
            "Component":    c["component_type"],
            "Dataset":      c["dataset"],
            "Model":        c.get("model_used", ""),
            "Base CAPEX":   float(c["prediction"]),
            "Owner's Cost": float(c["breakdown"]["owners_cost"]),
            "Contingency":  float(c["breakdown"]["contingency_cost"]),
            "Escalation":   float(c["breakdown"]["escalation_cost"]),
            "SST":          float(c["breakdown"]["sst_cost"]),
            "Grand Total":  float(c["breakdown"]["grand_total"]),
        })
    return pd.DataFrame(rows)

def project_totals(proj):
    dfc = project_components_df(proj)
    if dfc.empty:
        return {"capex_sum": 0.0, "owners": 0.0, "cont": 0.0, "esc": 0.0, "sst": 0.0, "grand_total": 0.0}
    return {
        "capex_sum":   float(dfc["Base CAPEX"].sum()),
        "owners":      float(dfc["Owner's Cost"].sum()),
        "cont":        float(dfc["Contingency"].sum()),
        "esc":         float(dfc["Escalation"].sum()),
        "sst":         float(dfc["SST"].sum()),
        "grand_total": float(dfc["Grand Total"].sum()),
    }

# ——————————————————————————————
# DATA PREPROCESSOR
# ——————————————————————————————

class DataPreprocessor:
    @staticmethod
    def clean_dataframe(df):
        df  = df.copy()
        bad = [c for c in df.columns if is_junk_col(c)]
        if bad:
            df = df.drop(columns=bad)
        return df

    @staticmethod
    def extract_features_target(df):
        if df is None or df.empty:
            raise ValueError("Empty dataset")
        target_col   = df.columns[-1]
        feature_cols = [c for c in df.columns if c != target_col]
        if not feature_cols:
            raise ValueError("No feature columns found")
        X = df[feature_cols].copy()
        y = pd.to_numeric(df[target_col], errors="coerce")
        if y.isna().sum() / len(y) > 0.8:
            raise ValueError(f"Target '{target_col}' has too many missing values")
        return X, y, target_col

    @staticmethod
    def validate_feature_columns(X, required_cols=None):
        X       = X.copy()
        non_num = [c for c in X.columns if X[c].dtype == object]
        if non_num:
            st.warning(f"Non-numeric columns will be converted: {non_num}")
            for c in non_num:
                X[c] = pd.to_numeric(X[c], errors="coerce")
        if required_cols:
            missing = [c for c in required_cols if c not in X.columns]
            if missing:
                raise ValueError(f"Missing columns: {missing}")
        return X

# ——————————————————————————————
# MLP MODEL
# ——————————————————————————————

class CapexMLP(nn.Module if TORCH_AVAILABLE else object):
    def __init__(self, n_features):
        if not TORCH_AVAILABLE:
            raise ImportError("PyTorch not installed")
        super().__init__()
        self.net = nn.Sequential(
            nn.Linear(n_features, 128), nn.BatchNorm1d(128), nn.ReLU(), nn.Dropout(0.3),
            nn.Linear(128, 64),         nn.BatchNorm1d(64),  nn.ReLU(), nn.Dropout(0.2),
            nn.Linear(64, 32),          nn.ReLU(),
            nn.Linear(32, 1),
        )

    def forward(self, x):
        return self.net(x).squeeze(1)


class MLPWrapper:
    def __init__(self, n_features, epochs=200, lr=0.001, batch_size=32,
                 patience=20, random_state=42):
        self.n_features   = n_features
        self.epochs       = epochs
        self.lr           = lr
        self.batch_size   = batch_size
        self.patience     = patience
        self.random_state = random_state
        self.model        = None
        self.scaler_X     = StandardScaler()
        self.scaler_y     = StandardScaler()
        self.train_losses = []
        self.val_losses   = []
        self.imputer      = SimpleImputer(strategy="median")

    def fit(self, X, y):
        torch.manual_seed(self.random_state)
        np.random.seed(self.random_state)
        X_imp = self.imputer.fit_transform(X)
        X_sc  = self.scaler_X.fit_transform(X_imp).astype(np.float32)
        y_sc  = self.scaler_y.fit_transform(y.reshape(-1, 1)).ravel().astype(np.float32)
        n_val = max(1, int(len(X_sc) * 0.15))
        X_tr, X_val = X_sc[:-n_val], X_sc[-n_val:]
        y_tr, y_val = y_sc[:-n_val], y_sc[-n_val:]
        loader = DataLoader(
            TensorDataset(torch.from_numpy(X_tr), torch.from_numpy(y_tr)),
            batch_size=self.batch_size, shuffle=True
        )
        self.model = CapexMLP(self.n_features)
        opt   = torch.optim.Adam(self.model.parameters(), lr=self.lr)
        sched = torch.optim.lr_scheduler.ReduceLROnPlateau(opt, patience=10, factor=0.5)
        crit  = nn.MSELoss()
        best_val, best_state, no_imp = float("inf"), None, 0
        self.train_losses, self.val_losses = [], []
        for _ in range(self.epochs):
            self.model.train()
            ep = 0.0
            for Xb, yb in loader:
                opt.zero_grad()
                loss = crit(self.model(Xb), yb)
                loss.backward()
                opt.step()
                ep += loss.item() * len(Xb)
            ep /= len(X_tr)
            self.model.eval()
            with torch.no_grad():
                vl = crit(self.model(torch.from_numpy(X_val)), torch.from_numpy(y_val)).item()
            self.train_losses.append(ep)
            self.val_losses.append(vl)
            sched.step(vl)
            if vl < best_val:
                best_val  = vl
                best_state = {k: v.clone() for k, v in self.model.state_dict().items()}
                no_imp    = 0
            else:
                no_imp += 1
                if no_imp >= self.patience:
                    break
        if best_state:
            self.model.load_state_dict(best_state)
        self.model.eval()
        return self

    def predict(self, X):
        X_imp = self.imputer.transform(X)
        X_sc  = self.scaler_X.transform(X_imp).astype(np.float32)
        with torch.no_grad():
            y_sc = self.model(torch.from_numpy(X_sc)).numpy()
        return self.scaler_y.inverse_transform(y_sc.reshape(-1, 1)).ravel()

# ——————————————————————————————
# MODEL PIPELINE — RF + GB + MLP
# ——————————————————————————————

class ModelPipeline:
    MODEL_CANDIDATES = {
        "RandomForest": lambda rs=42: RandomForestRegressor(
            n_estimators=200, max_depth=None, min_samples_split=2,
            min_samples_leaf=1, random_state=rs, n_jobs=-1),
        "GradientBoosting": lambda rs=42: GradientBoostingRegressor(
            n_estimators=200, learning_rate=0.05, max_depth=4,
            subsample=0.8, random_state=rs),
    }

    @classmethod
    def create_pipeline(cls, name, random_state=42):
        if name not in cls.MODEL_CANDIDATES:
            name = "RandomForest"
        ctor = cls.MODEL_CANDIDATES[name]
        try:    model = ctor(random_state)
        except: model = ctor()
        return Pipeline([("imputer", SimpleImputer(strategy="median")), ("model", model)])

    @classmethod
    @st.cache_resource(show_spinner=False)
    def train_all_cached(_cls, X, y, test_size=0.20, random_state=42,
                         mlp_epochs=200, mlp_lr=0.001, mlp_batch=32, mlp_patience=20):
        X_arr = X.values.astype(np.float32)
        y_arr = y.values.astype(np.float32)
        X_train, X_test, y_train, y_test = train_test_split(
            X_arr, y_arr, test_size=test_size, random_state=random_state)
        results = {}
        for name in ("RandomForest", "GradientBoosting"):
            pipe = _cls.create_pipeline(name, random_state)
            pipe.fit(X_train, y_train)
            yp = pipe.predict(X_test)
            results[name] = {
                "pipeline": pipe,
                "r2":       round(float(r2_score(y_test, yp)), 4),
                "rmse":     round(float(np.sqrt(mean_squared_error(y_test, yp))), 4),
                "mae":      round(float(mean_absolute_error(y_test, yp)), 4),
                "y_test":   y_test,
                "y_pred":   yp,
            }
        if TORCH_AVAILABLE:
            mlp = MLPWrapper(X_train.shape[1], mlp_epochs, mlp_lr,
                             mlp_batch, mlp_patience, random_state)
            mlp.fit(X_train, y_train)
            yp_mlp = mlp.predict(X_test)
            results["MLP"] = {
                "pipeline":     mlp,
                "r2":           round(float(r2_score(y_test, yp_mlp)), 4),
                "rmse":         round(float(np.sqrt(mean_squared_error(y_test, yp_mlp))), 4),
                "mae":          round(float(mean_absolute_error(y_test, yp_mlp)), 4),
                "y_test":       y_test,
                "y_pred":       yp_mlp,
                "train_losses": mlp.train_losses,
                "val_losses":   mlp.val_losses,
            }
        else:
            results["MLP"] = {
                "pipeline": None, "r2": None, "rmse": None, "mae": None,
                "y_test":   y_test, "y_pred": np.zeros_like(y_test),
                "train_losses": [], "val_losses": [],
            }
        valid = {k: v for k, v in results.items() if v["r2"] is not None}
        best  = max(valid, key=lambda k: valid[k]["r2"])
        bm    = valid[best]
        return {
            "rf":  results["RandomForest"],
            "gb":  results["GradientBoosting"],
            "mlp": results["MLP"],
            "best": best,
            "pipeline":     bm["pipeline"],
            "feature_cols": list(X.columns),
            "model": best,
            "r2":   bm["r2"],
            "rmse": bm["rmse"],
            "mae":  bm["mae"],
        }

    @staticmethod
    def prepare_prediction_input(feature_cols, payload):
        row = {}
        for col in feature_cols:
            val = payload.get(col, np.nan)
            if val is None or (isinstance(val, str) and val.strip() == ""):
                row[col] = np.nan
            elif isinstance(val, (int, float, np.number)):
                row[col] = float(val)
            else:
                try:    row[col] = float(val)
                except: row[col] = np.nan
        return pd.DataFrame([row], columns=feature_cols)

# ——————————————————————————————
# MONTE CARLO
# ——————————————————————————————

def monte_carlo_simulation(model_pipeline, feature_cols, base_values,
                           n_simulations=1000, feature_uncertainty=0.05,
                           cost_uncertainty=None):
    np.random.seed(42)
    base_array = np.array([float(base_values.get(c, np.nan)) for c in feature_cols])
    preds = []
    for _ in range(n_simulations):
        noise  = np.random.normal(0, feature_uncertainty, len(base_array))
        sim_df = pd.DataFrame([base_array * (1 + noise)], columns=feature_cols)
        try:    preds.append(float(model_pipeline.predict(sim_df)[0]))
        except: preds.append(0.0)
    return pd.DataFrame({"prediction": preds})

# ——————————————————————————————
# GITHUB
# ——————————————————————————————

GITHUB_USER = "Hafizuddin-Abd-Rahman-Dev-Upstream"
REPO_NAME   = "Cost-Predictor"
BRANCH      = "main"
DATA_FOLDER = "pages/data_CAPEX"

@st.cache_data(ttl=600, show_spinner=False)
def fetch_json(url):
    r = requests.get(url, timeout=15)
    r.raise_for_status()
    return r.json()

@st.cache_data(ttl=600, show_spinner=False)
def list_csvs_from_manifest(folder_path):
    url = (
        f"https://raw.githubusercontent.com/{GITHUB_USER}/{REPO_NAME}"
        f"/{BRANCH}/{folder_path}/files.json"
    )
    try:
        data = fetch_json(url)
        return [str(x) for x in data] if isinstance(data, list) else []
    except Exception as e:
        st.error(f"Failed to load CSV manifest: {e}")
        return []

# ——————————————————————————————
# NAV
# ——————————————————————————————

nav_labels = ["SHALLOW WATER", "DEEP WATER", "ONSHORE", "UNCON", "CCS"]
nav_cols   = st.columns(len(nav_labels))
for col, label in zip(nav_cols, nav_labels):
    with col:
        url = SHAREPOINT_LINKS.get(label.title(), "#")
        st.markdown(
            f'<a href="{url}" target="_blank" rel="noopener" class="petronas-button"'
            f' style="width:100%;text-align:center;display:inline-block;">{label}</a>',
            unsafe_allow_html=True,
        )

tab_data, tab_pb, tab_mc, tab_compare = st.tabs(
    ["📊 Data", "🏗️ Project Builder", "🎲 Monte Carlo", "🔀 Compare Projects"]
)

# ======================================================================================
# TAB 1 — DATA
# ======================================================================================

with tab_data:
    st.markdown('<h3 style="margin-top:0;color:#000;">📁 Data</h3>', unsafe_allow_html=True)

    st.markdown('<h4 style="margin:0;color:#000;">Data Sources</h4><p></p>', unsafe_allow_html=True)
    c1, c2 = st.columns([1.2, 1])
    with c1:
        data_source = st.radio("Choose data source", ["Upload CSV", "Load from Server"],
                               horizontal=True, key="data_source")
    with c2:
        st.caption("Enterprise Storage (SharePoint)")
        data_link = (
            "https://petronas.sharepoint.com/sites/ecm_ups_coe/confidential/"
            "DFE%20Cost%20Engineering/Forms/AllItems.aspx?"
            "id=%2Fsites%2Fecm%5Fups%5Fcoe%2Fconfidential%2FDFE%20Cost%20Engineering"
            "%2F2%2ETemplate%20Tools%2FCost%20Predictor%2FDatabase%2FCAPEX%20%2D%20RT%20Q1%202025"
        )
        st.markdown(
            f'<a href="{data_link}" target="_blank" rel="noopener" class="petronas-button">'
            f'Open Enterprise Storage</a>',
            unsafe_allow_html=True,
        )

    uploaded_files = []
    if data_source == "Upload CSV":
        uploaded_files = st.file_uploader(
            "Upload CSV files (max 200MB)", type="csv", accept_multiple_files=True,
            key=f"csv_uploader_{st.session_state.uploader_nonce}",
        )
    else:
        github_csvs = list_csvs_from_manifest(DATA_FOLDER)
        if github_csvs:
            selected_file = st.selectbox("Choose CSV from GitHub", github_csvs,
                                         key="github_csv_select")
            if st.button("Load selected CSV", key="load_github_csv_btn"):
                raw_url = (
                    f"https://raw.githubusercontent.com/{GITHUB_USER}/{REPO_NAME}"
                    f"/{BRANCH}/{DATA_FOLDER}/{selected_file}"
                )
                try:
                    df = DataPreprocessor.clean_dataframe(pd.read_csv(raw_url))
                    st.session_state.datasets[selected_file] = df
                    st.session_state.predictions.setdefault(selected_file, [])
                    toast(f"Loaded: {selected_file}")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error loading CSV: {e}")
        else:
            st.info("No CSV files found in GitHub folder.")

    if uploaded_files:
        for up in uploaded_files:
            if up.name not in st.session_state.datasets:
                try:
                    df = DataPreprocessor.clean_dataframe(pd.read_csv(up))
                    st.session_state.datasets[up.name] = df
                    st.session_state.predictions.setdefault(up.name, [])
                except Exception as e:
                    st.error(f"Failed to read {up.name}: {e}")
        toast("Dataset(s) added.")

    st.divider()

    cA, cB, cC, cD = st.columns([1, 1, 1, 2])
    with cA:
        if st.button("🧹 Clear predictions", key="clear_preds_btn"):
            st.session_state.predictions = {k: [] for k in st.session_state.predictions}
            toast("Predictions cleared.")
            st.rerun()
    with cB:
        if st.button("🧺 Clear history", key="clear_processed_btn"):
            st.session_state.processed_excel_files = set()
            toast("History cleared.")
            st.rerun()
    with cC:
        if st.button("🔁 Refresh", key="refresh_manifest_btn"):
            list_csvs_from_manifest.clear()
            fetch_json.clear()
            toast("Refreshed.")
            st.rerun()
    with cD:
        if st.button("🗂️ Clear all data", key="clear_datasets_btn"):
            st.session_state.datasets                = {}
            st.session_state.predictions             = {}
            st.session_state.processed_excel_files   = set()
            st.session_state._last_metrics           = None
            st.session_state.uploader_nonce         += 1
            st.session_state.widget_nonce           += 1
            toast("All data cleared.")
            st.rerun()

    st.divider()

    if st.session_state.datasets:
        ds_name_data      = st.selectbox("Active dataset",
                                         list(st.session_state.datasets.keys()),
                                         key="active_dataset_data")
        df_active         = st.session_state.datasets[ds_name_data]
        target_col_active = df_active.columns[-1]
        currency_active   = get_currency_symbol(df_active, target_col_active)

        colA, colB, colC, colD2 = st.columns([1, 1, 1, 2])
        colA.metric("Rows",     f"{df_active.shape[0]:,}")
        colB.metric("Columns",  f"{df_active.shape[1]:,}")
        colC.metric("Currency", currency_active or "—")
        colD2.caption(f"Target column: **{target_col_active}**")

        with st.expander("Preview (first 10 rows)", expanded=False):
            st.dataframe(df_active.head(10), use_container_width=True)
    else:
        st.info("Upload or load a dataset to proceed.")

    # ── MODEL TRAINING ────────────────────────────────────────────────────────
    st.divider()
    st.markdown('<h3 style="margin-top:0;color:#000;">⚙️ Model Training</h3>',
                unsafe_allow_html=True)

    if st.session_state.datasets:
        ds_name_model = st.selectbox("Dataset for training",
                                     list(st.session_state.datasets.keys()), key="ds_model")
        df_model = st.session_state.datasets[ds_name_model]

        try:
            X, y, target_col = DataPreprocessor.extract_features_target(df_model)
            st.success(f"Data prepared: {X.shape[1]} features, target: {target_col}")
            c1, c2, c3 = st.columns(3)
            c1.metric("Features", X.shape[1])
            c2.metric("Samples",  X.shape[0])
            vn = int(y.notna().sum())
            c3.metric("Valid targets", f"{vn} ({vn/len(y)*100:.1f}%)")
            _x_ok = True
        except Exception as e:
            st.error(f"Data preparation failed: {e}")
            _x_ok = False

        if _x_ok:
            st.markdown("##### Train / Test Split")
            _sc, _bc = st.columns([3, 1])
            with _sc:
                test_size = st.slider("Test set size", 0.10, 0.40, 0.20, 0.05,
                                      key="train_test_size",
                                      help="Proportion of data held out for evaluation")
                _tp = round((1 - test_size) * 100)
                _ep = round(test_size * 100)
                _nt = int(len(X) * (1 - test_size))
                _ne = len(X) - _nt
                st.markdown(
                    f'<div style="display:flex;height:14px;border-radius:7px;overflow:hidden;margin-top:4px;">'
                    f'<div style="width:{_tp}%;background:#00A19B;"></div>'
                    f'<div style="width:{_ep}%;background:#6C4DD3;"></div></div>'
                    f'<div style="display:flex;justify-content:space-between;font-size:12px;margin-top:4px;">'
                    f'<span style="color:#00A19B;font-weight:600;">Train {_tp}% ({_nt} rows)</span>'
                    f'<span style="color:#6C4DD3;font-weight:600;">Test {_ep}% ({_ne} rows)</span></div>',
                    unsafe_allow_html=True,
                )
            with _bc:
                st.write("")
                st.write("")
                run_train = st.button("🚀 Train RF, GB & MLP", key="run_training_btn",
                                      type="primary")

            with st.expander("🧠 MLP Hyperparameters", expanded=False):
                if not TORCH_AVAILABLE:
                    st.warning("PyTorch not installed — MLP will be skipped. "
                               "Add torch to requirements.txt.")
                _m1, _m2, _m3, _m4 = st.columns(4)
                with _m1:
                    mlp_epochs = st.number_input("Max Epochs", 50, 500, 200, 50, key="mlp_epochs")
                with _m2:
                    mlp_lr = st.select_slider("Learning Rate",
                                              [0.0001, 0.0005, 0.001, 0.005, 0.01],
                                              value=0.001, key="mlp_lr")
                with _m3:
                    mlp_batch = st.selectbox("Batch Size", [16, 32, 64, 128],
                                             index=1, key="mlp_batch")
                with _m4:
                    mlp_patience = st.number_input("Early Stop Patience", 5, 50, 20, 5,
                                                   key="mlp_patience")
                st.caption("Architecture: Input -> 128 (ReLU,BN,Drop 0.3) "
                           "-> 64 (ReLU,BN,Drop 0.2) -> 32 (ReLU) -> 1")

            if run_train:
                try:
                    with st.spinner("Training RF, GB and MLP..."):
                        metrics = ModelPipeline.train_all_cached(
                            X, y,
                            test_size    = float(test_size),
                            random_state = 42,
                            mlp_epochs   = int(mlp_epochs),
                            mlp_lr       = float(mlp_lr),
                            mlp_batch    = int(mlp_batch),
                            mlp_patience = int(mlp_patience),
                        )
                    st.session_state._last_metrics = metrics
                    st.session_state[f"trained_model__{ds_name_model}"]    = metrics
                    st.session_state[f"current_pipeline__{ds_name_model}"] = metrics["pipeline"]
                    st.session_state[f"feature_cols__{ds_name_model}"]     = metrics["feature_cols"]
                    knn = KNNImputer(n_neighbors=5)
                    knn.fit(X)
                    st.session_state[f"knn_imputer_{ds_name_model}"] = knn
                    toast("Training complete!")

                    if not TORCH_AVAILABLE:
                        st.warning("MLP skipped — add torch to requirements.txt.")

                    # 3-way comparison table
                    st.markdown("##### Model Comparison — RF vs GB vs MLP")
                    rf  = metrics["rf"]
                    gb  = metrics["gb"]
                    mlp = metrics["mlp"]
                    cdf = pd.DataFrame({
                        "Metric": [
                            "R2 Score (higher is better)",
                            "RMSE (lower is better)",
                            "MAE (lower is better)",
                        ],
                        "Random Forest":       [rf["r2"],  rf["rmse"],  rf["mae"]],
                        "Gradient Boosting":   [gb["r2"],  gb["rmse"],  gb["mae"]],
                        "MLP (Deep Learning)": [
                            mlp["r2"]   if mlp["r2"]   is not None else float("nan"),
                            mlp["rmse"] if mlp["rmse"] is not None else float("nan"),
                            mlp["mae"]  if mlp["mae"]  is not None else float("nan"),
                        ],
                    })

                    def _hl3(row):
                        vals = {
                            "Random Forest":       row["Random Forest"],
                            "Gradient Boosting":   row["Gradient Boosting"],
                            "MLP (Deep Learning)": row["MLP (Deep Learning)"],
                        }
                        valid = {k: v for k, v in vals.items()
                                 if not (isinstance(v, float) and np.isnan(v))}
                        if not valid:
                            return [""] * len(row)
                        best_k = (max if "higher" in row["Metric"] else min)(
                            valid, key=valid.get)
                        return [
                            "background-color:#d4f5f3;font-weight:700"
                            if c == best_k else ""
                            for c in row.index
                        ]

                    st.dataframe(
                        cdf.style.apply(_hl3, axis=1).format({
                            "Random Forest":       "{:.4f}",
                            "Gradient Boosting":   "{:.4f}",
                            "MLP (Deep Learning)": lambda v: f"{v:.4f}" if not np.isnan(v) else "N/A",
                        }),
                        use_container_width=True,
                        hide_index=True,
                    )

                    winner = metrics["best"]
                    wlabel = {
                        "RandomForest":     "Random Forest",
                        "GradientBoosting": "Gradient Boosting",
                        "MLP":              "MLP (Deep Learning)",
                    }.get(winner, winner)
                    st.success(f"{wlabel} selected as active model (highest R2)")

                    _k1, _k2, _k3, _k4 = st.columns(4)
                    _k1.metric("Model", wlabel)
                    _k2.metric("R2",    f"{metrics['r2']:.4f}")
                    _k3.metric("RMSE",  f"{metrics['rmse']:,.2f}")
                    _k4.metric("MAE",   f"{metrics['mae']:,.2f}")

                    # Actual vs Predicted — all 3 overlaid
                    st.markdown("##### Actual vs Predicted — All Models")
                    fig_sc = go.Figure()
                    all_v  = []
                    for _key, _label, _colour in [
                        ("rf",  "Random Forest",    "#00A19B"),
                        ("gb",  "Gradient Boosting","#6C4DD3"),
                        ("mlp", "MLP",              "#F4801A"),
                    ]:
                        _m = metrics[_key]
                        if _m["r2"] is None:
                            continue
                        fig_sc.add_trace(go.Scatter(
                            x=_m["y_test"], y=_m["y_pred"], mode="markers",
                            marker=dict(color=_colour, opacity=0.55, size=6),
                            name=_label,
                        ))
                        all_v.extend(_m["y_test"].tolist())
                        all_v.extend(_m["y_pred"].tolist())
                    _lo, _hi = float(min(all_v)), float(max(all_v))
                    fig_sc.add_trace(go.Scatter(
                        x=[_lo, _hi], y=[_lo, _hi], mode="lines",
                        line=dict(color="#888", dash="dash", width=1.5),
                        name="Perfect fit",
                    ))
                    fig_sc.update_layout(
                        xaxis_title="Actual CAPEX (MM USD)",
                        yaxis_title="Predicted CAPEX (MM USD)",
                        height=420, margin=dict(l=0, r=0, t=10, b=0),
                        paper_bgcolor="white", plot_bgcolor="white",
                        legend=dict(orientation="h", y=-0.18),
                    )
                    st.plotly_chart(fig_sc, use_container_width=True)

                    # MLP loss curve
                    if TORCH_AVAILABLE and metrics["mlp"]["train_losses"]:
                        st.markdown("##### MLP Training & Validation Loss")
                        _tl = metrics["mlp"]["train_losses"]
                        _vl = metrics["mlp"]["val_losses"]
                        fig_l = go.Figure()
                        fig_l.add_trace(go.Scatter(
                            x=list(range(1, len(_tl)+1)), y=_tl,
                            mode="lines", name="Train Loss",
                            line=dict(color="#00A19B", width=2),
                        ))
                        fig_l.add_trace(go.Scatter(
                            x=list(range(1, len(_vl)+1)), y=_vl,
                            mode="lines", name="Val Loss",
                            line=dict(color="#F4801A", width=2, dash="dot"),
                        ))
                        fig_l.update_layout(
                            xaxis_title="Epoch",
                            yaxis_title="MSE Loss (scaled)",
                            height=300, margin=dict(l=0, r=0, t=10, b=0),
                            paper_bgcolor="white", plot_bgcolor="white",
                            legend=dict(orientation="h", y=-0.22),
                        )
                        st.plotly_chart(fig_l, use_container_width=True)
                        st.caption(
                            f"Trained for {len(_tl)} epochs "
                            f"(early stopping patience={mlp_patience})"
                        )

                    # Feature importance RF vs GB
                    st.markdown("##### Feature Importance — RF vs GB")
                    st.caption("MLP does not produce feature importances natively.")
                    _fl, _fr = st.columns(2)
                    for _con, (_lbl, _bk) in zip(
                        [_fl, _fr],
                        [("Random Forest", "rf"), ("Gradient Boosting", "gb")],
                    ):
                        _pipe = metrics[_bk]["pipeline"]
                        _imps = _pipe.named_steps["model"].feature_importances_
                        _fi   = pd.DataFrame({
                            "Feature":    metrics["feature_cols"],
                            "Importance": _imps,
                        }).sort_values("Importance", ascending=True)
                        _fig_fi = go.Figure(go.Bar(
                            x=_fi["Importance"], y=_fi["Feature"], orientation="h",
                            marker_color="#00A19B" if _bk == "rf" else "#6C4DD3",
                        ))
                        _fig_fi.update_layout(
                            title=_lbl, xaxis_title="Importance",
                            height=max(260, 32 * len(_fi)),
                            margin=dict(l=0, r=0, t=35, b=0),
                            paper_bgcolor="white", plot_bgcolor="white",
                        )
                        with _con:
                            st.plotly_chart(_fig_fi, use_container_width=True)

                except Exception as e:
                    st.error(f"Training failed: {e}")

    # ── VISUALIZATION (4 tabs) ────────────────────────────────────────────────
    st.divider()
    st.markdown('<h3 style="margin-top:0;color:#000;">📈 Visualization</h3>',
                unsafe_allow_html=True)

    if st.session_state.datasets:
        ds_name_viz    = st.selectbox("Dataset for visualization",
                                      list(st.session_state.datasets.keys()), key="ds_viz")
        df_viz         = st.session_state.datasets[ds_name_viz]
        target_col_viz = df_viz.columns[-1]
        currency_viz   = get_currency_symbol(df_viz, target_col_viz) or "USD"
        num_viz        = df_viz.select_dtypes(include=[np.number])

        vt1, vt2, vt3, vt4 = st.tabs([
            "🔥 Correlation Matrix",
            "🌐 3D Cost Surface",
            "📊 Distribution",
            "🔗 Scatter Matrix",
        ])

        with vt1:
            try:
                if len(num_viz.columns) > 1:
                    corr     = num_viz.corr()
                    fig_corr = px.imshow(corr, text_auto=".2f", aspect="auto",
                                         color_continuous_scale="RdBu_r", zmin=-1, zmax=1)
                    fig_corr.update_layout(margin=dict(l=0, r=0, t=10, b=0),
                                           paper_bgcolor="white")
                    st.plotly_chart(fig_corr, use_container_width=True)
                    st.caption("+1 = strong positive correlation  |  -1 = inverse  |  0 = weak")

                    if f"current_pipeline__{ds_name_viz}" in st.session_state:
                        _pipe_v = st.session_state[f"current_pipeline__{ds_name_viz}"]
                        if hasattr(_pipe_v.named_steps.get("model"), "feature_importances_"):
                            _imps_v = _pipe_v.named_steps["model"].feature_importances_
                            _fn_v   = st.session_state[f"feature_cols__{ds_name_viz}"]
                            _fi_v   = pd.DataFrame({
                                "feature":    _fn_v,
                                "importance": _imps_v,
                            }).sort_values("importance", ascending=True)
                            st.markdown("##### Feature Importance (active model)")
                            _fig_fv = go.Figure(go.Bar(
                                x=_fi_v["importance"], y=_fi_v["feature"],
                                orientation="h", marker_color="#00A19B",
                            ))
                            _fig_fv.update_layout(xaxis_title="Importance",
                                                  margin=dict(l=0, r=0, t=10, b=0))
                            st.plotly_chart(_fig_fv, use_container_width=True)
                else:
                    st.info("Need at least 2 numeric columns for correlation matrix.")
            except Exception as e:
                st.error(f"Correlation error: {e}")

        with vt2:
            st.caption("Drag to rotate  |  Scroll to zoom  |  Hover for exact values")
            _cl    = {c.lower(): c for c in df_viz.columns}
            _x_col = _cl.get("water_depth_m") or _cl.get("length_km")
            _y_col = _cl.get("topsides_weight_t") or _cl.get("diameter_inch")
            _af    = [c for c in num_viz.columns if c != target_col_viz]
            if not _x_col and len(_af) >= 1: _x_col = _af[0]
            if not _y_col and len(_af) >= 2: _y_col = _af[1]
            if _x_col and _y_col and len(_af) >= 2:
                _c3a, _c3b = st.columns(2)
                with _c3a:
                    _x_col = st.selectbox("X axis", _af,
                                          index=_af.index(_x_col) if _x_col in _af else 0,
                                          key="viz3d_x")
                with _c3b:
                    _yi    = _af.index(_y_col) if _y_col in _af else 0
                    _y_col = st.selectbox("Y axis", _af, index=_yi, key="viz3d_y")
                _p3d = df_viz[[_x_col, _y_col, target_col_viz]].dropna()
                if len(_p3d) >= 3:
                    _fig3d = go.Figure(data=[go.Scatter3d(
                        x=_p3d[_x_col], y=_p3d[_y_col], z=_p3d[target_col_viz],
                        mode="markers",
                        marker=dict(
                            size=5, color=_p3d[target_col_viz],
                            colorscale="Teal",
                            colorbar=dict(title=dict(text=f"CAPEX ({currency_viz}M)", side="right"),
                                          thickness=14, len=0.6),
                            opacity=0.85,
                            line=dict(width=0.4, color="rgba(0,0,0,0.25)"),
                        ),
                        hovertemplate=(
                            f"<b>{_x_col}:</b> %{{x:,.1f}}<br>"
                            f"<b>{_y_col}:</b> %{{y:,.1f}}<br>"
                            f"<b>CAPEX:</b> {currency_viz} %{{z:,.2f}}M<extra></extra>"
                        ),
                    )])
                    _fig3d.update_layout(
                        scene=dict(
                            xaxis=dict(title=_x_col,           backgroundcolor="rgba(0,0,0,0)", gridcolor="#e0e0e0", showbackground=True),
                            yaxis=dict(title=_y_col,           backgroundcolor="rgba(0,0,0,0)", gridcolor="#e0e0e0", showbackground=True),
                            zaxis=dict(title=f"CAPEX ({currency_viz}M)", backgroundcolor="rgba(0,0,0,0)", gridcolor="#e0e0e0", showbackground=True),
                            camera=dict(eye=dict(x=1.6, y=1.6, z=0.8)),
                        ),
                        margin=dict(l=0, r=0, t=10, b=0),
                        height=540,
                        paper_bgcolor="white",
                    )
                    st.plotly_chart(_fig3d, use_container_width=True)
                    _s1, _s2, _s3, _s4 = st.columns(4)
                    _s1.metric("Min CAPEX",    f"{currency_viz} {_p3d[target_col_viz].min():,.1f}M")
                    _s2.metric("Max CAPEX",    f"{currency_viz} {_p3d[target_col_viz].max():,.1f}M")
                    _s3.metric("Mean CAPEX",   f"{currency_viz} {_p3d[target_col_viz].mean():,.1f}M")
                    _s4.metric("Data points",  f"{len(_p3d):,}")
                else:
                    st.warning("Not enough data after removing nulls.")
            else:
                st.info(f"Need at least 2 numeric feature columns. Found: {list(num_viz.columns)}")

        with vt3:
            _dist_col = st.selectbox("Select column", num_viz.columns.tolist(),
                                     index=len(num_viz.columns)-1, key="viz_dist_col")
            _d1, _d2  = st.columns(2)
            with _d1:
                _fh = px.histogram(df_viz, x=_dist_col, nbins=30,
                                   title=f"Distribution — {_dist_col}",
                                   color_discrete_sequence=["#00A19B"],
                                   template="plotly_white")
                _fh.update_layout(height=320, margin=dict(l=0, r=0, t=35, b=0), showlegend=False)
                st.plotly_chart(_fh, use_container_width=True)
            with _d2:
                _fb = px.box(df_viz, y=_dist_col, title=f"Box Plot — {_dist_col}",
                             color_discrete_sequence=["#6C4DD3"], template="plotly_white")
                _fb.update_layout(height=320, margin=dict(l=0, r=0, t=35, b=0), showlegend=False)
                st.plotly_chart(_fb, use_container_width=True)
            st.markdown("##### Descriptive Statistics")
            st.dataframe(
                df_viz[_dist_col].describe().to_frame().T.style.format("{:,.2f}"),
                use_container_width=True,
            )

        with vt4:
            _nc = num_viz.columns.tolist()
            if len(_nc) >= 3:
                _sel_sc = st.multiselect(
                    "Choose columns (max 6 recommended)", _nc,
                    default=_nc[:min(6, len(_nc))], key="viz_scatter_cols",
                )
                if len(_sel_sc) >= 2:
                    _fsm = px.scatter_matrix(
                        df_viz[_sel_sc].dropna(), dimensions=_sel_sc,
                        color=df_viz[target_col_viz],
                        color_continuous_scale="Teal",
                        title="Scatter Matrix", template="plotly_white", opacity=0.6,
                    )
                    _fsm.update_traces(marker=dict(size=3))
                    _fsm.update_layout(
                        height=600, margin=dict(l=0, r=0, t=35, b=0),
                        coloraxis_colorbar=dict(title=f"CAPEX ({currency_viz}M)", thickness=12),
                    )
                    st.plotly_chart(_fsm, use_container_width=True)
                    st.caption(
                        "Each cell = one feature pair  |  "
                        "Colour = CAPEX  |  "
                        "Diagonal = that feature's distribution"
                    )
                else:
                    st.info("Select at least 2 columns.")
            else:
                st.info("Need at least 3 numeric columns for scatter matrix.")

    # ── PREDICT ───────────────────────────────────────────────────────────────
    st.divider()
    st.markdown('<h3 style="margin-top:0;color:#000;">🎯 Predict</h3>',
                unsafe_allow_html=True)

    if st.session_state.datasets:
        ds_name_pred = st.selectbox("Dataset for prediction",
                                    list(st.session_state.datasets.keys()), key="ds_pred")
        df_pred = st.session_state.datasets[ds_name_pred]

        if f"current_pipeline__{ds_name_pred}" not in st.session_state:
            st.warning("Please train a model first using the Model Training section above.")
        else:
            pipeline      = st.session_state[f"current_pipeline__{ds_name_pred}"]
            feature_cols  = st.session_state[f"feature_cols__{ds_name_pred}"]
            target_col    = df_pred.columns[-1]
            currency_pred = get_currency_symbol(df_pred, target_col)
            meta          = st.session_state.get(f"trained_model__{ds_name_pred}", {})
            active_model  = meta.get("best", "—")
            active_label  = {
                "RandomForest":     "Random Forest",
                "GradientBoosting": "Gradient Boosting",
                "MLP":              "MLP (Deep Learning)",
            }.get(active_model, active_model)
            st.info(f"Active model: **{active_label}**")

            st.markdown('<h4 style="margin:0;color:#000;">Cost Factors</h4>'
                        '<p>Adjust cost percentages</p>', unsafe_allow_html=True)
            c1, c2 = st.columns(2)
            with c1:
                sst_pct    = st.number_input("SST (%)",          0.0, 100.0, 0.0, 0.5, key="pred_sst")
                owners_pct = st.number_input("Owner's Cost (%)", 0.0, 100.0, 0.0, 0.5, key="pred_owner")
            with c2:
                cont_pct   = st.number_input("Contingency (%)",  0.0, 100.0, 0.0, 0.5, key="pred_cont")
                esc_pct    = st.number_input("Escalation & Inflation (%)",
                                             0.0, 100.0, 0.0, 0.5, key="pred_esc")

            st.markdown('<h4 style="margin:0;color:#000;">Project Details</h4>',
                        unsafe_allow_html=True)
            project_name = st.text_input(
                "Project Name",
                placeholder="e.g., Offshore Pipeline Replacement 2026",
                key="pred_project_name",
            )

            st.markdown("##### Feature Values")
            st.caption(f"Enter values for {len(feature_cols)} features. Leave blank for NaN.")
            input_values = {}
            for i in range(0, len(feature_cols), 3):
                _cols = st.columns(3)
                for j, feat in enumerate(feature_cols[i:i+3]):
                    with _cols[j]:
                        val = st.text_input(feat, value="",
                                            key=f"input_{feat}_{ds_name_pred}",
                                            help=f"Enter value for {feat}")
                        if val.strip() in ("", "nan"):
                            input_values[feat] = np.nan
                        else:
                            try:    input_values[feat] = float(val)
                            except: input_values[feat] = np.nan

            use_knn = st.checkbox(
                "Use intelligent imputation for missing features (KNN)", value=False,
                key=f"use_knn_{ds_name_pred}",
                help="Estimates missing values from similar rows in the training data.",
            )

            if st.button("Run Prediction", key="run_pred_btn", type="primary"):
                try:
                    pred_input      = ModelPipeline.prepare_prediction_input(feature_cols, input_values)
                    original_inputs = input_values.copy()
                    imp_method      = "pipeline median"

                    if use_knn:
                        knn_key = f"knn_imputer_{ds_name_pred}"
                        if knn_key in st.session_state:
                            arr        = st.session_state[knn_key].transform(pred_input)
                            pred_input = pd.DataFrame(arr, columns=feature_cols)
                            imp_method = "KNN (intelligent)"
                        else:
                            st.warning("KNN imputer not available. Using median fallback.")

                    base_pred = float(pipeline.predict(pred_input)[0])

                    st.markdown("##### Feature Values Used for Prediction")
                    st.caption(f"Imputation method: **{imp_method}**")
                    _crows = []
                    for col in feature_cols:
                        u = original_inputs.get(col)
                        v = pred_input[col].iloc[0]
                        _crows.append({
                            "Feature":    col,
                            "Your Input": f"{u:,.2f}" if (u is not None and not pd.isna(u)) else "—",
                            "Value Used": f"{v:,.2f}" if isinstance(v, (int, float)) else str(v),
                            "Source":     "User provided"
                                          if (u is not None and not pd.isna(u))
                                          else f"Imputed ({imp_method})",
                        })
                    st.dataframe(
                        pd.DataFrame(_crows),
                        use_container_width=True,
                        height=min(400, 35 * len(feature_cols)),
                    )

                    owners_cost, sst_cost, contingency_cost, escalation_cost, grand_total = \
                        cost_breakdown(base_pred, sst_pct, owners_pct, cont_pct, esc_pct)

                    result = {
                        "Project Name": project_name,
                        "Base CAPEX":   round(base_pred, 2),
                        "Owner's Cost": owners_cost,
                        "SST Cost":     sst_cost,
                        "Contingency":  contingency_cost,
                        "Escalation":   escalation_cost,
                        "Grand Total":  grand_total,
                        "Target":       round(base_pred, 2),
                        "_imputation_method": imp_method,
                    }
                    for col in feature_cols:
                        result[col] = pred_input[col].iloc[0]

                    st.session_state.predictions.setdefault(ds_name_pred, []).append(result)
                    toast("Prediction added!")

                    st.markdown("##### Prediction Results")
                    r1, r2, r3, r4, r5 = st.columns(5)
                    r1.metric("Base CAPEX",   f"{currency_pred} {base_pred:,.2f}")
                    r2.metric("Owner's Cost", f"{currency_pred} {owners_cost:,.2f}")
                    r3.metric("SST",          f"{currency_pred} {sst_cost:,.2f}")
                    r4.metric("Contingency",  f"{currency_pred} {contingency_cost:,.2f}")
                    r5.metric("Grand Total",  f"{currency_pred} {grand_total:,.2f}")

                except Exception as e:
                    st.error(f"Prediction failed: {str(e)}")

            # Batch prediction
            st.markdown("---")
            st.markdown('<h4 style="margin:0;color:#000;">Batch Prediction (Excel)</h4>',
                        unsafe_allow_html=True)
            excel_file = st.file_uploader(
                "Upload Excel file for batch prediction", type=["xlsx"],
                key=f"batch_excel_{st.session_state.widget_nonce}",
            )
            if excel_file:
                file_id = f"{excel_file.name}_{excel_file.size}_{ds_name_pred}"
                if file_id not in st.session_state.processed_excel_files:
                    try:
                        batch_df     = pd.read_excel(excel_file)
                        missing_cols = [c for c in feature_cols if c not in batch_df.columns]
                        if missing_cols:
                            st.error(f"Missing required columns in Excel: {missing_cols}")
                        else:
                            X_batch           = DataPreprocessor.validate_feature_columns(
                                batch_df[feature_cols])
                            predictions_batch = pipeline.predict(X_batch)
                            for i, (_, row) in enumerate(batch_df.iterrows()):
                                bp       = float(predictions_batch[i])
                                oc, sc, cc, ec, gt = cost_breakdown(
                                    bp, sst_pct, owners_pct, cont_pct, esc_pct)
                                r = {
                                    "Project Name": str(row.get("Project Name", f"Project {i+1}")),
                                    "Base CAPEX":   round(bp, 2),
                                    "Owner's Cost": oc,
                                    "SST Cost":     sc,
                                    "Contingency":  cc,
                                    "Escalation":   ec,
                                    "Grand Total":  gt,
                                    "Target":       round(bp, 2),
                                }
                                for feat in feature_cols:
                                    r[feat] = row.get(feat, np.nan)
                                st.session_state.predictions.setdefault(
                                    ds_name_pred, []).append(r)
                            st.session_state.processed_excel_files.add(file_id)
                            st.success(f"Processed {len(batch_df)} rows")
                    except Exception as e:
                        st.error(f"Batch processing failed: {str(e)}")
                else:
                    st.info("This file has already been processed.")

    # Results / export
    st.divider()
    st.markdown('<h3 style="margin-top:0;color:#000;">📄 Results</h3>',
                unsafe_allow_html=True)

    if st.session_state.datasets:
        ds_name_res = st.selectbox("Dataset for results",
                                   list(st.session_state.datasets.keys()), key="ds_results")
        preds = st.session_state.predictions.get(ds_name_res, [])
        if preds:
            df_preds  = pd.DataFrame(preds)
            disp_cols = ["Project Name", "Base CAPEX", "Owner's Cost", "SST Cost",
                         "Contingency", "Escalation", "Grand Total"]
            disp_cols = [c for c in disp_cols if c in df_preds.columns]
            df_disp   = df_preds[disp_cols].copy()
            for col in disp_cols[1:]:
                df_disp[col] = df_disp[col].apply(
                    lambda x: f"{x:,.2f}" if not pd.isna(x) else "")
            st.dataframe(df_disp, use_container_width=True, height=300)

            st.markdown("##### Export Results")
            _c1, _c2 = st.columns(2)
            with _c1:
                bio = io.BytesIO()
                df_preds.to_excel(bio, index=False, engine="openpyxl")
                bio.seek(0)
                st.download_button(
                    "Download Excel (with all features)", data=bio,
                    file_name=f"{ds_name_res}_predictions.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_excel_btn",
                )
            with _c2:
                st.download_button(
                    "Download CSV (with all features)",
                    data=df_preds.to_csv(index=False),
                    file_name=f"{ds_name_res}_predictions.csv",
                    mime="text/csv",
                    key="download_csv_btn",
                )
            if st.button("Clear all predictions", key="clear_predictions_btn"):
                st.session_state.predictions[ds_name_res] = []
                st.rerun()
        else:
            st.info("No predictions yet. Make some predictions above.")

# ======================================================================================
# TAB 2 — PROJECT BUILDER
# ======================================================================================

with tab_pb:
    st.markdown('<h4 style="margin-top:0;color:#000;">🏗️ Project Builder</h4>',
                unsafe_allow_html=True)
    st.caption("Assemble multi-component CAPEX projects")

    if not st.session_state.datasets:
        st.info("No datasets loaded. Please load data in the Data tab first.")
    else:
        colA, colB = st.columns([2, 1])
        with colA:
            new_project_name = st.text_input(
                "New Project Name", placeholder="e.g., CAPEX 2026",
                key="pb_new_project_name",
            )
        with colB:
            if new_project_name and new_project_name not in st.session_state.projects:
                if st.button("Create Project", key="pb_create_project_btn"):
                    st.session_state.projects[new_project_name] = {
                        "components": [], "totals": {}, "currency": "",
                        "cost_factors": {
                            "sst_pct": 0.0, "owners_pct": 0.0,
                            "cont_pct": 0.0, "esc_pct": 0.0,
                        },
                    }
                    toast(f"Project '{new_project_name}' created.")
                    st.rerun()

        if not st.session_state.projects:
            st.info("Create a project above, then add components.")
        else:
            proj_sel = st.selectbox("Select project",
                                    list(st.session_state.projects.keys()),
                                    key="pb_project_select")
            proj = st.session_state.projects[proj_sel]

            st.markdown("##### Project Cost Factors")
            cf1, cf2 = st.columns(2)
            with cf1:
                proj["cost_factors"]["sst_pct"] = st.number_input(
                    "SST (%)", 0.0, 100.0,
                    value=proj["cost_factors"].get("sst_pct", 0.0),
                    step=0.5, key=f"pb_sst_{proj_sel}",
                )
                proj["cost_factors"]["owners_pct"] = st.number_input(
                    "Owner's Cost (%)", 0.0, 100.0,
                    value=proj["cost_factors"].get("owners_pct", 0.0),
                    step=0.5, key=f"pb_owners_{proj_sel}",
                )
            with cf2:
                proj["cost_factors"]["cont_pct"] = st.number_input(
                    "Contingency (%)", 0.0, 100.0,
                    value=proj["cost_factors"].get("cont_pct", 0.0),
                    step=0.5, key=f"pb_cont_{proj_sel}",
                )
                proj["cost_factors"]["esc_pct"] = st.number_input(
                    "Escalation & Inflation (%)", 0.0, 100.0,
                    value=proj["cost_factors"].get("esc_pct", 0.0),
                    step=0.5, key=f"pb_esc_{proj_sel}",
                )

            st.markdown("##### Add Component")
            ds_names         = sorted(st.session_state.datasets.keys())
            dataset_for_comp = st.selectbox("Dataset for component", ds_names,
                                            key="pb_dataset_for_component")
            df_comp = st.session_state.datasets[dataset_for_comp]

            if f"current_pipeline__{dataset_for_comp}" not in st.session_state:
                st.warning(
                    f"Please train a model for '{dataset_for_comp}' in the Data tab first.")
            else:
                pipeline     = st.session_state[f"current_pipeline__{dataset_for_comp}"]
                feature_cols = st.session_state[f"feature_cols__{dataset_for_comp}"]
                target_col   = df_comp.columns[-1]
                curr_ds      = get_currency_symbol(df_comp, target_col)

                component_type = st.text_input(
                    "Component type",
                    placeholder="e.g., Pipeline, Platform, Subsea",
                    key=f"pb_component_type_{proj_sel}",
                )

                st.markdown("###### Component Features")
                comp_inputs = {}
                for i in range(0, len(feature_cols), 2):
                    _cols = st.columns(2)
                    for j, feat in enumerate(feature_cols[i:i+2]):
                        with _cols[j]:
                            val = st.text_input(
                                feat, placeholder="Enter value",
                                key=f"pb_{feat}_{proj_sel}_{dataset_for_comp}",
                            )
                            if val.strip() in ("", "nan"):
                                comp_inputs[feat] = np.nan
                            else:
                                try:    comp_inputs[feat] = float(val)
                                except: comp_inputs[feat] = np.nan

                if st.button("➕ Add Component to Project",
                             key=f"pb_add_comp_{proj_sel}"):
                    if not component_type:
                        st.error("Please enter a component type.")
                    else:
                        try:
                            pi = ModelPipeline.prepare_prediction_input(feature_cols, comp_inputs)
                            bp = float(pipeline.predict(pi)[0])
                            cf = proj["cost_factors"]
                            oc, sc, cc, ec, gt = cost_breakdown(
                                bp, cf["sst_pct"], cf["owners_pct"],
                                cf["cont_pct"], cf["esc_pct"],
                            )
                            proj["components"].append({
                                "component_type": component_type,
                                "dataset":        dataset_for_comp,
                                "prediction":     bp,
                                "breakdown": {
                                    "owners_cost":      oc,
                                    "sst_cost":         sc,
                                    "contingency_cost": cc,
                                    "escalation_cost":  ec,
                                    "grand_total":      gt,
                                },
                            })
                            proj["currency"] = curr_ds
                            toast(f"Component '{component_type}' added.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Failed to add component: {str(e)}")

            st.markdown("---")
            st.markdown("##### Project Overview")
            comps = proj.get("components", [])
            if not comps:
                st.info("No components yet. Add components above.")
            else:
                st.dataframe(pd.DataFrame([{
                    "Component":  c["component_type"],
                    "Dataset":    c["dataset"],
                    "Base CAPEX": f"{curr_ds} {c['prediction']:,.2f}",
                    "Grand Total":f"{curr_ds} {c['breakdown']['grand_total']:,.2f}",
                } for c in comps]), use_container_width=True)

                totals = project_totals(proj)
                t1, t2, t3 = st.columns(3)
                t1.metric("Total Base CAPEX",  f"{curr_ds} {totals['capex_sum']:,.2f}")
                t2.metric("Total SST",         f"{curr_ds} {totals['sst']:,.2f}")
                t3.metric("Total Grand Total", f"{curr_ds} {totals['grand_total']:,.2f}")

                st.markdown("##### Component Management")
                for idx, comp in enumerate(comps):
                    _c1, _c2, _c3 = st.columns([3, 2, 1])
                    _c1.write(f"**{comp['component_type']}**")
                    _c1.caption(f"{comp['dataset']} | Base: {curr_ds} {comp['prediction']:,.2f}")
                    _c2.write(f"Grand Total: {curr_ds} {comp['breakdown']['grand_total']:,.2f}")
                    with _c3:
                        if st.button("🗑️", key=f"del_comp_{proj_sel}_{idx}"):
                            comps.pop(idx)
                            st.rerun()

                st.markdown("---")
                proj_json = json.dumps(proj, indent=2, default=float)
                st.download_button(
                    "Download Project (JSON)", data=proj_json,
                    file_name=f"{proj_sel}.json", mime="application/json",
                    key=f"dl_json_{proj_sel}",
                )

            st.markdown("##### Import Project")
            up_json = st.file_uploader("Upload project JSON", type=["json"],
                                       key=f"import_{proj_sel}")
            if up_json:
                try:
                    st.session_state.projects[proj_sel] = json.load(up_json)
                    toast("Project imported successfully.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Failed to import project: {str(e)}")

# ======================================================================================
# TAB 3 — MONTE CARLO
# ======================================================================================

with tab_mc:
    st.markdown('<h3 style="margin-top:0;color:#000;">🎲 Monte Carlo Analysis</h3>',
                unsafe_allow_html=True)
    st.caption("Simple uncertainty analysis for project costs")

    if not st.session_state.projects:
        st.info("No projects found. Create a project in the Project Builder tab first.")
    else:
        proj_names  = list(st.session_state.projects.keys())
        proj_sel_mc = st.selectbox("Select project", proj_names, key="mc_project_select")
        proj_mc     = st.session_state.projects[proj_sel_mc]
        comps_mc    = proj_mc.get("components", [])

        if not comps_mc:
            st.warning("This project has no components. Add components in the Project Builder.")
        else:
            st.markdown("##### Simulation Settings")
            mc1, mc2 = st.columns(2)
            with mc1:
                n_simulations       = st.number_input("Number of simulations",
                                                       100, 10000, 1000, 100, key="mc_n_sims")
                feature_uncertainty = st.slider("Feature uncertainty (%)",
                                                0.0, 50.0, 10.0, 1.0, key="mc_feat_unc")
            with mc2:
                cost_uncertainty = st.slider("Cost uncertainty (%)",
                                             0.0, 30.0, 5.0, 1.0, key="mc_cost_unc")
                budget           = st.number_input("Budget threshold (MM USD)",
                                                    min_value=0.0, value=1000.0,
                                                    step=10.0, key="mc_budget")

            if st.button("Run Monte Carlo", type="primary", key="mc_run"):
                try:
                    with st.spinner("Running simulations..."):
                        all_sims = []
                        for comp in comps_mc:
                            ds_name = comp["dataset"]
                            if f"current_pipeline__{ds_name}" not in st.session_state:
                                st.warning(f"No trained model for {ds_name}")
                                continue
                            _pipe  = st.session_state[f"current_pipeline__{ds_name}"]
                            _fcols = st.session_state[f"feature_cols__{ds_name}"]
                            sims   = monte_carlo_simulation(
                                _pipe, _fcols, {}, int(n_simulations),
                                feature_uncertainty / 100,
                            )
                            all_sims.append(sims["prediction"].values)

                    if all_sims:
                        total = np.sum(all_sims, axis=0)
                        p50   = np.percentile(total, 50)
                        p80   = np.percentile(total, 80)
                        p90   = np.percentile(total, 90)
                        ep    = (total > budget).mean() * 100

                        rc1, rc2, rc3, rc4 = st.columns(4)
                        rc1.metric("P50", f"${p50:,.0f}M")
                        rc2.metric("P80", f"${p80:,.0f}M")
                        rc3.metric("P90", f"${p90:,.0f}M")
                        rc4.metric(f"P(>${budget:,.0f}M)", f"{ep:.1f}%")

                        fig_mc = px.histogram(
                            x=total, nbins=50,
                            title="Total Cost Distribution",
                            labels={"x": "Total Cost (MM USD)", "y": "Frequency"},
                            color_discrete_sequence=["#00A19B"],
                        )
                        fig_mc.add_vline(
                            x=budget, line_dash="dash", line_color="red",
                            annotation_text=f"Budget: ${budget:,.0f}M",
                        )
                        st.plotly_chart(fig_mc, use_container_width=True)
                    else:
                        st.warning("No valid simulations were generated.")
                except Exception as e:
                    st.error(f"Monte Carlo failed: {str(e)}")

# ======================================================================================
# TAB 4 — COMPARE PROJECTS
# ======================================================================================

with tab_compare:
    st.markdown('<h3 style="margin-top:0;color:#000;">🔀 Compare Projects</h3>',
                unsafe_allow_html=True)

    if len(st.session_state.projects) < 2:
        st.info("Create at least 2 projects in the Project Builder to compare.")
    else:
        project_names     = list(st.session_state.projects.keys())
        selected_projects = st.multiselect(
            "Select projects to compare", project_names,
            default=project_names[:2] if len(project_names) >= 2 else project_names,
        )

        if len(selected_projects) < 2:
            st.warning("Please select at least 2 projects.")
        else:
            comparison_data = []
            for pn in selected_projects:
                _p = st.session_state.projects[pn]
                _t = project_totals(_p)
                comparison_data.append({
                    "Project":      pn,
                    "Components":   len(_p.get("components", [])),
                    "Base CAPEX":   _t["capex_sum"],
                    "SST":          _t["sst"],
                    "Owner's Cost": _t["owners"],
                    "Contingency":  _t["cont"],
                    "Escalation":   _t["esc"],
                    "Grand Total":  _t["grand_total"],
                })
            df_cmp = pd.DataFrame(comparison_data)
            st.markdown("##### Project Comparison")
            st.dataframe(
                df_cmp.style.format(
                    {c: "{:,.2f}" for c in df_cmp.columns
                     if c not in ("Project", "Components")}
                ),
                use_container_width=True,
            )

            viz_type = st.selectbox("Chart type", ["Bar Chart", "Stacked Bar"],
                                    key="viz_type")
            if viz_type == "Bar Chart":
                fig_cmp = px.bar(
                    df_cmp, x="Project", y="Grand Total",
                    title="Grand Total by Project", text="Grand Total",
                    color_discrete_sequence=["#00A19B"],
                )
                fig_cmp.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
            else:
                melt_df = df_cmp.melt(
                    id_vars=["Project"],
                    value_vars=["Base CAPEX", "SST", "Owner's Cost", "Contingency", "Escalation"],
                    var_name="Cost Type", value_name="Amount",
                )
                fig_cmp = px.bar(
                    melt_df, x="Project", y="Amount", color="Cost Type",
                    title="Cost Breakdown by Project", barmode="stack",
                )
            st.plotly_chart(fig_cmp, use_container_width=True)
