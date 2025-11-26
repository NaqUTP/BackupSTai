import streamlit as st
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
from sklearn.model_selection import train_test_split, cross_val_score
from sklearn.preprocessing import MinMaxScaler
from sklearn.ensemble import RandomForestRegressor, GradientBoostingRegressor
from sklearn.linear_model import Ridge, Lasso
from sklearn.svm import SVR
from sklearn.tree import DecisionTreeRegressor
from sklearn.metrics import mean_squared_error, r2_score, mean_absolute_error
import numpy as np
from scipy.stats import linregress
from sklearn.impute import KNNImputer
import io
import requests
from matplotlib.ticker import FuncFormatter
import json
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from datetime import datetime

# Page config
st.set_page_config(
    page_title="CAPEX AI RT2025",
    page_icon="💲",
    initial_sidebar_state="expanded"
)

# Simple auth placeholder
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

APPROVED_EMAILS = st.secrets.get("emails", [])
correct_password = st.secrets.get("password", None)

if not st.session_state.authenticated:
    with st.form("login_form"):
        st.markdown("#### 🔐 Access Required", unsafe_allow_html=True)
        email = st.text_input("Email Address")
        password = st.text_input("Enter Access Password", type="password")
        submitted = st.form_submit_button("Login")

        if submitted:
            if (not APPROVED_EMAILS or email in APPROVED_EMAILS) and (correct_password is None or password == correct_password):
                st.session_state.authenticated = True
                st.success("✅ Access granted.")
                st.rerun()
            else:
                st.error("❌ Invalid email or password. Please contact Cost Engineering Focal for access")
    st.stop()

# Repo config
GITHUB_USER = "Hafizuddin-Abd-Rahman-Dev-Upstream"
REPO_NAME = "Cost-Predictor"
BRANCH = "main"
DATA_FOLDER = "pages/data_CAPEX"

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

# Helpers
def human_format(num, pos=None):
    try:
        num = float(num)
    except Exception:
        return str(num)
    if num >= 1e9:
        return f'{num/1e9:.1f}B'
    elif num >= 1e6:
        return f'{num/1e6:.1f}M'
    elif num >= 1e3:
        return f'{num/1e3:.1f}K'
    else:
        return f'{num:.0f}'

def format_with_commas(num):
    try:
        return f"{float(num):,.2f}"
    except Exception:
        return str(num)

def get_currency_symbol(df: pd.DataFrame) -> str:
    symbols = ["$", "€", "£", "RM", "IDR", "SGD"]
    for col in df.columns:
        u = str(col).upper()
        for sym in symbols:
            if sym in u:
                return sym
    return ""

def format_currency(amount, currency=''):
    try:
        return f"{currency} {float(amount):,.2f}" if currency else f"{float(amount):,.2f}"
    except Exception:
        return str(amount)

def download_all_predictions():
    preds_all = st.session_state.get("predictions", {})
    if not preds_all or all(len(v) == 0 for v in preds_all.values()):
        st.sidebar.error("No predictions available to download")
        return

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        summary_data = []
        for dataset_name, preds in preds_all.items():
            if not preds:
                continue
            for pred in preds:
                row = pred.copy()
                row["Dataset"] = dataset_name.replace(".csv", "")
                summary_data.append(row)
        if summary_data:
            pd.DataFrame(summary_data).to_excel(writer, sheet_name="All Predictions", index=False)
        for dataset_name, preds in preds_all.items():
            if preds:
                sheet = dataset_name.replace(".csv", "")[:31]
                pd.DataFrame(preds).to_excel(writer, sheet_name=sheet, index=False)
    output.seek(0)
    st.sidebar.download_button(
        "📥 Download All Predictions",
        data=output,
        file_name="All_Predictions.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# AUTO MODEL SELECTION - Train multiple models and select best
def train_best_model(df: pd.DataFrame, test_size=0.2, random_state=42):
    """
    Trains multiple regression models and automatically selects the best one based on R² score.
    Returns dict with best model, scaler, imputer, features, metrics, and comparison results.
    """
    if df.shape[1] < 2:
        raise ValueError("Dataset must have at least one feature and one target column.")

    target_col = df.columns[-1]
    X = df.iloc[:, :-1].select_dtypes(include=[np.number]).copy()
    y = df[target_col].astype(float).copy()

    if X.shape[1] == 0:
        raise ValueError("No numeric features found in the dataset.")

    # Impute and scale
    imputer = KNNImputer(n_neighbors=5)
    X_imputed = pd.DataFrame(imputer.fit_transform(X), columns=X.columns)
    
    scaler = MinMaxScaler()
    X_scaled = scaler.fit_transform(X_imputed)

    X_train, X_test, y_train, y_test = train_test_split(X_scaled, y, test_size=test_size, random_state=random_state)

    # Define models to test
    models = {
        'Random Forest': RandomForestRegressor(n_estimators=100, random_state=random_state, n_jobs=-1),
        'Gradient Boosting': GradientBoostingRegressor(n_estimators=100, random_state=random_state),
        'Ridge Regression': Ridge(alpha=1.0, random_state=random_state),
        'Lasso Regression': Lasso(alpha=1.0, random_state=random_state),
        'SVR': SVR(kernel='rbf'),
        'Decision Tree': DecisionTreeRegressor(random_state=random_state)
    }

    results = {}
    best_score = -np.inf
    best_model_name = None
    best_model = None

    # Train and evaluate each model
    for name, model in models.items():
        try:
            model.fit(X_train, y_train)
            y_pred = model.predict(X_test)
            
            rmse = np.sqrt(mean_squared_error(y_test, y_pred))
            r2 = r2_score(y_test, y_pred)
            mae = mean_absolute_error(y_test, y_pred)
            
            # Cross-validation score
            cv_scores = cross_val_score(model, X_train, y_train, cv=5, scoring='r2')
            cv_mean = cv_scores.mean()
            
            results[name] = {
                'rmse': float(rmse),
                'r2': float(r2),
                'mae': float(mae),
                'cv_r2_mean': float(cv_mean),
                'cv_r2_std': float(cv_scores.std())
            }
            
            # Select best model based on R² score
            if r2 > best_score:
                best_score = r2
                best_model_name = name
                best_model = model
        except Exception as e:
            st.warning(f"Failed to train {name}: {e}")
            results[name] = {'rmse': np.nan, 'r2': np.nan, 'mae': np.nan, 'cv_r2_mean': np.nan, 'cv_r2_std': np.nan}

    if best_model is None:
        raise ValueError("No model could be trained successfully.")

    return {
        "model": best_model,
        "model_name": best_model_name,
        "scaler": scaler,
        "imputer": imputer,
        "features": X.columns.tolist(),
        "target": target_col,
        "metrics": results[best_model_name],
        "all_results": results,
        "X_train_shape": X_train.shape,
        "X_test_shape": X_test.shape
    }

# REPORT GENERATION FUNCTIONS

def create_project_excel_report(project_name, project_data, currency=""):
    """Generate detailed Excel report for a project"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Summary sheet
        summary_data = []
        total_capex = 0
        total_grand = 0
        
        for comp in project_data.get("components", []):
            capex = float(comp["prediction"])
            grand = float(comp["breakdown"]["grand_total"])
            total_capex += capex
            total_grand += grand
            
            summary_data.append({
                "Component": comp["component_type"],
                "Dataset": comp["dataset"],
                f"Predicted CAPEX {currency}": capex,
                f"Owner's Cost {currency}": comp["breakdown"]["owners_cost"],
                f"Contingency {currency}": comp["breakdown"]["contingency_cost"],
                f"Escalation {currency}": comp["breakdown"]["escalation_cost"],
                f"Grand Total {currency}": grand
            })
        
        df_summary = pd.DataFrame(summary_data)
        df_summary.loc[len(df_summary)] = ["", "TOTAL", total_capex, "", "", "", total_grand]
        df_summary.to_excel(writer, sheet_name="Summary", index=False)
        
        # EPCIC Breakdown sheet
        epcic_data = []
        for comp in project_data.get("components", []):
            row = {"Component": comp["component_type"]}
            for phase, details in comp["breakdown"]["epcic"].items():
                row[f"{phase} {currency}"] = details["cost"]
                row[f"{phase} %"] = details["percentage"]
            epcic_data.append(row)
        
        if epcic_data:
            pd.DataFrame(epcic_data).to_excel(writer, sheet_name="EPCIC Breakdown", index=False)
        
        # Detailed inputs sheet
        detailed_data = []
        for comp in project_data.get("components", []):
            row = {
                "Component": comp["component_type"],
                "Dataset": comp["dataset"],
            }
            row.update(comp["inputs"])
            row[comp["breakdown"]["target_col"]] = comp["prediction"]
            detailed_data.append(row)
        
        pd.DataFrame(detailed_data).to_excel(writer, sheet_name="Detailed Inputs", index=False)
    
    output.seek(0)
    return output

def create_project_pptx_report(project_name, project_data, currency="", comparison_data=None):
    """Generate comprehensive PowerPoint report"""
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    
    # Title slide
    title_slide = prs.slides.add_slide(prs.slide_layouts[6])
    left = Inches(1)
    top = Inches(2.5)
    width = Inches(8)
    height = Inches(1.5)
    
    txBox = title_slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.text = f"CAPEX Analysis Report\n{project_name}"
    
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = RGBColor(46, 134, 171)
    
    # Date
    date_box = title_slide.shapes.add_textbox(Inches(1), Inches(6), Inches(8), Inches(0.5))
    date_tf = date_box.text_frame
    date_tf.text = f"Generated: {datetime.now().strftime('%B %d, %Y')}"
    date_tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    date_tf.paragraphs[0].font.size = Pt(14)
    
    # Executive Summary slide
    summary_slide = prs.slides.add_slide(prs.slide_layouts[5])
    title = summary_slide.shapes.title
    title.text = "Executive Summary"
    
    components = project_data.get("components", [])
    total_capex = sum(float(c["prediction"]) for c in components)
    total_grand = sum(float(c["breakdown"]["grand_total"]) for c in components)
    
    left = Inches(1)
    top = Inches(1.5)
    width = Inches(8)
    height = Inches(5)
    
    txBox = summary_slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    
    summary_text = f"""Project Overview:
    
Total Components: {len(components)}
Total CAPEX: {currency} {total_capex:,.2f}
Total Grand Total: {currency} {total_grand:,.2f}

Components Breakdown:"""
    
    for comp in components:
        summary_text += f"\n  • {comp['component_type']}: {currency} {comp['breakdown']['grand_total']:,.2f}"
    
    tf.text = summary_text
    for paragraph in tf.paragraphs:
        paragraph.font.size = Pt(16)
    
    # Component Details slide
    for idx, comp in enumerate(components):
        detail_slide = prs.slides.add_slide(prs.slide_layouts[5])
        title = detail_slide.shapes.title
        title.text = f"Component {idx+1}: {comp['component_type']}"
        
        left = Inches(1)
        top = Inches(1.5)
        width = Inches(8)
        height = Inches(5.5)
        
        txBox = detail_slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        tf.word_wrap = True
        
        detail_text = f"""Dataset: {comp['dataset']}
Base CAPEX: {currency} {comp['prediction']:,.2f}

Cost Breakdown:
  • Owner's Cost: {currency} {comp['breakdown']['owners_cost']:,.2f}
  • Contingency: {currency} {comp['breakdown']['contingency_cost']:,.2f}
  • Escalation & Inflation: {currency} {comp['breakdown']['escalation_cost']:,.2f}
  • Pre-Development: {currency} {comp['breakdown']['predev_cost']:,.2f}

Grand Total: {currency} {comp['breakdown']['grand_total']:,.2f}

EPCIC Breakdown:"""
        
        for phase, details in comp['breakdown']['epcic'].items():
            if details['percentage'] > 0:
                detail_text += f"\n  • {phase}: {currency} {details['cost']:,.2f} ({details['percentage']}%)"
        
        tf.text = detail_text
        for paragraph in tf.paragraphs:
            paragraph.font.size = Pt(14)
    
    # Comparison slide (if provided)
    if comparison_data:
        comp_slide = prs.slides.add_slide(prs.slide_layouts[5])
        title = comp_slide.shapes.title
        title.text = "Project Comparison"
        
        left = Inches(1)
        top = Inches(1.5)
        width = Inches(8)
        height = Inches(5.5)
        
        txBox = comp_slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        tf.word_wrap = True
        
        comp_text = "Comparative Analysis:\n\n"
        for proj_name, proj_info in comparison_data.items():
            comp_text += f"{proj_name}:\n"
            comp_text += f"  • CAPEX: {currency} {proj_info['capex']:,.2f}\n"
            comp_text += f"  • Grand Total: {currency} {proj_info['grand_total']:,.2f}\n\n"
        
        tf.text = comp_text
        for paragraph in tf.paragraphs:
            paragraph.font.size = Pt(16)
    
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output

def create_comparison_excel_report(projects_dict, currency=""):
    """Generate Excel report comparing multiple projects"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Summary comparison
        summary_data = []
        for proj_name, proj_data in projects_dict.items():
            components = proj_data.get("components", [])
            total_capex = sum(float(c["prediction"]) for c in components)
            total_grand = sum(float(c["breakdown"]["grand_total"]) for c in components)
            
            summary_data.append({
                "Project": proj_name,
                "Number of Components": len(components),
                f"Total CAPEX {currency}": total_capex,
                f"Total Grand Total {currency}": total_grand
            })
        
        pd.DataFrame(summary_data).to_excel(writer, sheet_name="Projects Summary", index=False)
        
        # Detailed breakdown for each project
        for proj_name, proj_data in projects_dict.items():
            comp_data = []
            for comp in proj_data.get("components", []):
                comp_data.append({
                    "Component": comp["component_type"],
                    "Dataset": comp["dataset"],
                    f"CAPEX {currency}": comp["prediction"],
                    f"Owner's Cost {currency}": comp["breakdown"]["owners_cost"],
                    f"Contingency {currency}": comp["breakdown"]["contingency_cost"],
                    f"Escalation {currency}": comp["breakdown"]["escalation_cost"],
                    f"Grand Total {currency}": comp["breakdown"]["grand_total"]
                })
            
            sheet_name = proj_name[:31]  # Excel sheet name limit
            pd.DataFrame(comp_data).to_excel(writer, sheet_name=sheet_name, index=False)
    
    output.seek(0)
    return output

# Initialize session state
def init_state():
    ss = st.session_state
    ss.setdefault("datasets", {})
    ss.setdefault("predictions", {})
    ss.setdefault("processed_excel_files", set())
    ss.setdefault("models", {})
    ss.setdefault("projects", {})
    ss.setdefault("component_labels", {})
init_state()

# Sidebar controls
st.sidebar.header('Data Controls')
if st.sidebar.button("Clear all predictions"):
    st.session_state['predictions'] = {}
    st.sidebar.success("All predictions cleared!")
if st.sidebar.button("Clear processed files history"):
    st.session_state['processed_excel_files'] = set()
    st.sidebar.success("Processed files history cleared!")
if st.sidebar.button("📥 Download All Predictions"):
    if st.session_state['predictions']:
        download_all_predictions()
        st.sidebar.success("All predictions compiled successfully!")
    else:
        st.sidebar.warning("No predictions to download.")

st.sidebar.markdown('---')
st.sidebar.header('System Controls')
if st.sidebar.button("🔄 Refresh System"):
    list_csvs_from_manifest.clear()
st.sidebar.markdown('---')

st.sidebar.subheader("📁 Choose Data Source")
data_source = st.sidebar.radio("Data Source", ["Upload CSV", "Load from Server"], index=0)
uploaded_files = []
if data_source == "Upload CSV":
    uploaded_files = st.sidebar.file_uploader(
        "Upload CSV files (max 200MB)", type="csv", accept_multiple_files=True
    )
    st.sidebar.markdown("### 📁 Or access data from external link")
    data_link = "https://petronas.sharepoint.com/sites/ecm_ups_coe/confidential/DFE%20Cost%20Engineering/Forms/AllItems.aspx?id=%2Fsites%2Fecm%5Fups%5Fcoe%2Fconfidential%2FDFE%20Cost%20Engineering%2F2%2ETemplate%20Tools%2FCost%20Predictor%2FDatabase%2FCAPEX%20%2D%20RT%20Q1%202025&viewid=25092e6d%2D373d%2D41fe%2D8f6f%2D486cd8cdd5b8"
    st.sidebar.markdown(
        f'<a href="{data_link}" target="_blank"><button style="background-color:#0099ff;color:white;padding:8px 16px;border:none;border-radius:4px;">Open Data Storage</button></a>',
        unsafe_allow_html=True
    )
else:
    github_csvs = list_csvs_from_manifest(DATA_FOLDER)
    if github_csvs:
        selected_file = st.sidebar.selectbox("Choose CSV from GitHub", github_csvs)
        if selected_file:
            raw_url = f"https://raw.githubusercontent.com/{GITHUB_USER}/{REPO_NAME}/{BRANCH}/{DATA_FOLDER}/{selected_file}"
            try:
                df = pd.read_csv(raw_url)
                fake_file = type('FakeUpload', (), {'name': selected_file})
                uploaded_files.append(fake_file)
                st.session_state['datasets'][selected_file] = df
                st.session_state['predictions'].setdefault(selected_file, [])
                st.success(f"✅ Loaded from GitHub: {selected_file}")
            except Exception as e:
                st.error(f"Error loading CSV: {e}")
    else:
        st.warning("No CSV files found in GitHub folder.")

# Persist uploaded CSVs
for uploaded_file in uploaded_files:
    if uploaded_file.name not in st.session_state['datasets']:
        try:
            df = pd.read_csv(uploaded_file)
            st.session_state['datasets'][uploaded_file.name] = df
            st.session_state['predictions'].setdefault(uploaded_file.name, [])
        except Exception as e:
            st.error(f"Failed to read uploaded file {uploaded_file.name}: {e}")

st.sidebar.markdown('---')
if st.sidebar.checkbox("🧹 Cleanup Current Session", value=False):
    uploaded_names = {f.name for f in uploaded_files}
    for name in list(st.session_state['datasets'].keys()):
        if name not in uploaded_names:
            del st.session_state['datasets'][name]
            st.session_state['predictions'].pop(name, None)

# Main layout
st.title('💲CAPEX AI RT2025💲')
st.caption("🤖 Automatic Best Model Selection | 📊 Comprehensive Reports")

if not st.session_state['datasets']:
    st.info("Please upload one or more CSV files to begin.")
else:
    selected_dataset_name = st.sidebar.selectbox(
        "Select a dataset for prediction",
        list(st.session_state['datasets'].keys())
    )
    df = st.session_state['datasets'][selected_dataset_name]
    clean_name = selected_dataset_name.replace('.csv', '')
    st.subheader(f"📊 Dataset: {clean_name}")

    currency = get_currency_symbol(df)

    try:
        imputer_preview = KNNImputer(n_neighbors=5)
        numeric_preview_cols = df.select_dtypes(include=[np.number]).columns
        if len(numeric_preview_cols) > 0:
            df_imputed_preview = pd.DataFrame(
                imputer_preview.fit_transform(df[numeric_preview_cols]),
                columns=numeric_preview_cols
            )
        else:
            df_imputed_preview = pd.DataFrame()
    except Exception:
        df_imputed_preview = df.select_dtypes(include=[np.number]).copy()

    target_column = df.columns[-1]
    st.caption(f"🎯 Target column: '{target_column}'")

    # Model Training with Auto Selection
    with st.expander('🤖 Automatic Model Training & Selection', expanded=False):
        st.header('Automatic Best Model Selection')
        st.write("The system will train 6 different models and automatically select the best performer:")
        st.write("• Random Forest • Gradient Boosting • Ridge Regression • Lasso Regression • SVR • Decision Tree")
        
        test_size = st.slider('Test size (holdout fraction)', 0.05, 0.5, 0.2, step=0.01, key=f"ts_{selected_dataset_name}")
        
        if st.button("⚙️ Train & Auto-Select Best Model", key=f"train_btn_{selected_dataset_name}"):
            with st.spinner("Training multiple models and selecting best..."):
                try:
                    model_bundle = train_best_model(df, test_size=float(test_size))
                    st.session_state['models'][selected_dataset_name] = model_bundle
                    
                    st.success(f"✅ Best Model Selected: **{model_bundle['model_name']}**")
                    
                    # Display metrics
                    col1, col2, col3, col4 = st.columns(4)
                    metrics = model_bundle["metrics"]
                    col1.metric("RMSE", f"{metrics['rmse']:,.2f}")
                    col2.metric("R² Score", f"{metrics['r2']:.3f}")
                    col3.metric("MAE", f"{metrics['mae']:,.2f}")
                    col4.metric("CV R² Mean", f"{metrics['cv_r2_mean']:.3f}")
                    
                    # Show all models comparison
                    st.subheader("Model Comparison Results")
                    results_df = pd.DataFrame(model_bundle['all_results']).T
                    results_df = results_df.sort_values('r2', ascending=False)
                    st.dataframe(results_df.style.format({
                        'rmse': '{:.2f}',
                        'r2': '{:.3f}',
                        'mae': '{:.2f}',
                        'cv_r2_mean': '{:.3f}',
                        'cv_r2_std': '{:.3f}'
                    }).background_gradient(subset=['r2'], cmap='RdYlGn'))
                    
                except Exception as e:
                    st.error(f"Training failed: {e}")
        else:
            mb = st.session_state['models'].get(selected_dataset_name)
            if mb:
                st.info(f"✅ Current Model: **{mb['model_name']}**")
                m = mb['metrics']
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("RMSE", f"{m['rmse']:,.2f}")
                col2.metric("R² Score", f"{m['r2']:.3f}")
                col3.metric("MAE", f"{m['mae']:,.2f}")
                col4.metric("CV R² Mean", f"{m['cv_r2_mean']:.3f}")
                
                # Show comparison table
                if 'all_results' in mb:
                    with st.expander("View All Models Comparison"):
                        results_df = pd.DataFrame(mb['all_results']).T
                        results_df = results_df.sort_values('r2', ascending=False)
                        st.dataframe(results_df.style.format({
                            'rmse': '{:.2f}',
                            'r2': '{:.3f}',
                            'mae': '{:.2f}',
                            'cv_r2_mean': '{:.3f}',
                            'cv_r2_std': '{:.3f}'
                        }))
            else:
                st.info("No trained model yet. Click 'Train & Auto-Select Best Model' above.")

    # Data Overview
    with st.expander('Data Overview', expanded=False):
        st.write('Dataset Shape:', df.shape)
        st.dataframe(df.head())

    # Data Visualization
    with st.expander('Data Visualization', expanded=False):
        st.subheader('Correlation Matrix')
        num_df = df.select_dtypes(include=[np.number]).copy()
        feature_count = len(num_df.columns)
        corr_height = min(9, max(6, feature_count * 0.5))
        fig, ax = plt.subplots(figsize=(8, corr_height))
        if feature_count >= 1:
            sns.heatmap(num_df.corr(), annot=True, cmap='coolwarm', fmt='.2f', annot_kws={"size": 10}, ax=ax)
            plt.tight_layout()
            st.pyplot(fig)
        else:
            st.write("No numeric features to display correlation.")
        plt.close()

        st.subheader('Feature Importance')
        mb = st.session_state['models'].get(selected_dataset_name)
        if mb and hasattr(mb['model'], 'feature_importances_'):
            feat_cols = mb['features']
            importances = mb['model'].feature_importances_
            fi_df = pd.DataFrame({"feature": feat_cols, "importance": importances}).sort_values("importance", ascending=False)
            fi_height = min(8, max(3, len(fi_df) * 0.3))
            fig, ax = plt.subplots(figsize=(8, fi_height))
            sns.barplot(data=fi_df, x="importance", y="feature", ax=ax)
            plt.title(f"Feature Importance ({mb['model_name']})")
            plt.tight_layout()
            st.pyplot(fig)
            plt.close()
        else:
            st.info("Feature importance not available for the selected model.")

        st.subheader('Cost Curve')
        other_cols = [c for c in df.columns if c != target_column]
        if other_cols:
            feature = st.selectbox('Select feature for cost curve', other_cols, key=f'cc_{selected_dataset_name}')
            fig, ax = plt.subplots(figsize=(7, 6))
            x_vals = df[feature].values
            y_vals = df[target_column].values
            mask = (~pd.isna(x_vals)) & (~pd.isna(y_vals))
            if mask.sum() >= 2:
                slope, intercept, r_val, _, _ = linregress(x_vals[mask], y_vals[mask])
                sns.scatterplot(x=x_vals, y=y_vals, label='Original Data', ax=ax)
                x_line = np.linspace(min(x_vals[mask]), max(x_vals[mask]), 100)
                y_line = slope * x_line + intercept
                ax.plot(x_line, y_line, color='red', label=f'Fit: y = {slope:.2f}x + {intercept:.2f}')
                ax.text(0.05, 0.95, f'$R^2$ = {r_val**2:.3f}', transform=ax.transAxes,
                        verticalalignment='top', bbox=dict(facecolor='white', alpha=0.5))
            else:
                sns.scatterplot(x=x_vals, y=y_vals, label='Original Data', ax=ax)
                st.warning("Not enough data for regression.")
            ax.set_xlabel(feature)
            ax.set_ylabel(target_column)
            ax.set_title(f'Cost Curve: {feature} vs {target_column}')
            ax.legend()
            ax.xaxis.set_major_formatter(FuncFormatter(human_format))
            ax.yaxis.set_major_formatter(FuncFormatter(human_format))
            plt.tight_layout()
            st.pyplot(fig)
            plt.close()
        else:
            st.write("No suitable feature for cost curve visualisation.")

    # Cost Breakdown Configuration
    with st.expander('Cost Breakdown Configuration', expanded=False):
        st.subheader("🔧 Cost Breakdown Percentage Input")
        st.markdown("Enter the percentage breakdown for EPCIC phases. Leave as 0 if not applicable.")
        epcic_percentages = {}
        col_ep1, col_ep2, col_ep3, col_ep4, col_ep5 = st.columns(5)
        epcic_percentages["Engineering"] = col_ep1.number_input("Engineering (%)", min_value=0.0, max_value=100.0, value=0.0, key=f"ep_eng_{selected_dataset_name}")
        epcic_percentages["Procurement"] = col_ep2.number_input("Procurement (%)", min_value=0.0, max_value=100.0, value=0.0, key=f"ep_proc_{selected_dataset_name}")
        epcic_percentages["Construction"] = col_ep3.number_input("Construction (%)", min_value=0.0, max_value=100.0, value=0.0, key=f"ep_const_{selected_dataset_name}")
        epcic_percentages["Installation"] = col_ep4.number_input("Installation (%)", min_value=0.0, max_value=100.0, value=0.0, key=f"ep_inst_{selected_dataset_name}")
        epcic_percentages["Commissioning"] = col_ep5.number_input("Commissioning (%)", min_value=0.0, max_value=100.0, value=0.0, key=f"ep_comm_{selected_dataset_name}")
        epcic_total = sum(epcic_percentages.values())
        if abs(epcic_total - 100.0) > 1e-3 and epcic_total > 0:
            st.warning(f"⚠️ EPCIC total is {epcic_total:.2f}%. Consider summing to 100% if applicable.")

        st.subheader("💼 Pre-Dev and Owner's Cost Percentage Input")
        col_pd1, col_pd2 = st.columns(2)
        predev_percentage = col_pd1.number_input("Pre-Development (%)", min_value=0.0, max_value=100.0, value=0.0, key=f"pd_{selected_dataset_name}")
        owners_percentage = col_pd2.number_input("Owner's Cost (%)", min_value=0.0, max_value=100.0, value=0.0, key=f"owners_{selected_dataset_name}")

        col_ct1, col_ct2 = st.columns(2)
        contingency_percentage = col_ct1.number_input("Contingency (%)", min_value=0.0, max_value=100.0, value=0.0, key=f"cont_{selected_dataset_name}")
        escalation_percentage = col_ct2.number_input("Escalation & Inflation (%)", min_value=0.0, max_value=100.0, value=0.0, key=f"escal_{selected_dataset_name}")

    # Make New Predictions
    st.header('Make New Predictions')
    project_name = st.text_input('Enter Project Name', key=f"pn_{selected_dataset_name}")

    if selected_dataset_name not in st.session_state['models']:
        st.info("No trained model for this dataset. Train it in 'Automatic Model Training & Selection' expander.")
    else:
        mb = st.session_state['models'][selected_dataset_name]
        feat_cols = mb["features"]

        num_features = len(feat_cols)
        if num_features == 0:
            st.error("No feature columns available.")
        else:
            if num_features <= 2:
                cols = st.columns(num_features)
            else:
                cols = []
                for i in range(0, num_features, 2):
                    row_cols = st.columns(min(2, num_features - i))
                    cols.extend(row_cols)

            new_data = {}
            for i, col in enumerate(feat_cols):
                col_idx = i % len(cols)
                user_val = cols[col_idx].text_input(f'{col}', key=f'input_{selected_dataset_name}_{col}')
                if str(user_val).strip().lower() in ("", "nan"):
                    new_data[col] = np.nan
                else:
                    try:
                        new_data[col] = float(user_val)
                    except Exception:
                        new_data[col] = np.nan

            if st.button('Predict', key=f'predict_btn_{selected_dataset_name}'):
                try:
                    df_input_raw = pd.DataFrame([new_data], columns=feat_cols)
                    X_imputed = pd.DataFrame(mb['imputer'].transform(df_input_raw), columns=feat_cols)
                    X_scaled = mb['scaler'].transform(X_imputed)
                    pred = float(mb['model'].predict(X_scaled)[0])

                    result = {'Project Name': project_name, **{k: (v if not pd.isna(v) else None) for k, v in new_data.items()}, mb['target']: round(pred, 2)}

                    for phase, pct in epcic_percentages.items():
                        cost = round(pred * (pct / 100.0), 2)
                        result[f"{phase} Cost"] = cost

                    predev_cost = round(pred * (predev_percentage / 100.0), 2)
                    owners_cost = round(pred * (owners_percentage / 100.0), 2)
                    contingency_base = pred + owners_cost
                    contingency_cost = round(contingency_base * (contingency_percentage / 100.0), 2)
                    escalation_base = pred + owners_cost
                    escalation_cost = round(escalation_base * (escalation_percentage / 100.0), 2)
                    grand_total = round(pred + owners_cost + contingency_cost + escalation_cost, 2)

                    result.update({
                        "Pre-Development Cost": predev_cost,
                        "Owner's Cost": owners_cost,
                        "Cost Contingency": contingency_cost,
                        "Escalation & Inflation": escalation_cost,
                        "Grand Total": grand_total
                    })

                    st.session_state['predictions'].setdefault(selected_dataset_name, [])
                    st.session_state['predictions'][selected_dataset_name].append(result)

                    display_text = (
                        f"### ✅ Cost Summary of project {project_name}\n\n"
                        f"**Model Used:** {mb['model_name']}\n\n"
                        f"**{mb['target']}:** {format_currency(pred, currency)}\n\n"
                        + (f"**Pre-Development:** {format_currency(predev_cost, currency)}\n\n" if predev_percentage > 0 else "")
                        + (f"**Owner's Cost:** {format_currency(owners_cost, currency)}\n\n" if owners_percentage > 0 else "")
                        + (f"**Contingency:** {format_currency(contingency_cost, currency)}\n\n" if contingency_percentage > 0 else "")
                        + (f"**Escalation & Inflation:** {format_currency(escalation_cost, currency)}\n\n" if escalation_percentage > 0 else "")
                        + f"**Grand Total:** {format_currency(grand_total, currency)}"
                    )
                    st.success(display_text)
                except Exception as e:
                    st.error(f"Prediction failed: {e}")

        st.write("Or upload an Excel file for batch prediction:")
        excel_file = st.file_uploader("Upload Excel file", type=["xlsx"], key=f"excel_{selected_dataset_name}")
        if excel_file:
            file_id = f"{excel_file.name}_{excel_file.size}_{selected_dataset_name}"
            if file_id not in st.session_state['processed_excel_files']:
                try:
                    batch_df = pd.read_excel(excel_file)
                    mb = st.session_state['models'][selected_dataset_name]
                    if set(mb['features']).issubset(batch_df.columns):
                        X_batch = batch_df[mb['features']]
                        X_imputed = pd.DataFrame(mb['imputer'].transform(X_batch), columns=mb['features'])
                        X_scaled = mb['scaler'].transform(X_imputed)
                        preds = mb['model'].predict(X_scaled)
                        batch_df[mb['target']] = preds
                        st.session_state['predictions'].setdefault(selected_dataset_name, [])
                        for i, row in batch_df.iterrows():
                            name = row.get("Project Name", f"Project {i+1}")
                            entry = {"Project Name": name}
                            entry.update(row[mb['features']].to_dict())
                            p = float(preds[i])
                            entry[mb['target']] = round(p, 2)
                            for phase, percent in epcic_percentages.items():
                                entry[f"{phase} Cost"] = round(p * (percent / 100.0), 2)
                            predev_cost = round(p * (predev_percentage / 100.0), 2)
                            owners_cost = round(p * (owners_percentage / 100.0), 2)
                            contingency_base = p + owners_cost
                            contingency_cost = round(contingency_base * (contingency_percentage / 100.0), 2)
                            escalation_base = p + owners_cost
                            escalation_cost = round(escalation_base * (escalation_percentage / 100.0), 2)
                            grand_total = round(p + owners_cost + contingency_cost + escalation_cost, 2)
                            entry["Pre-Development Cost"] = predev_cost
                            entry["Owner's Cost"] = owners_cost
                            entry["Cost Contingency"] = contingency_cost
                            entry["Escalation & Inflation"] = escalation_cost
                            entry["Grand Total"] = grand_total
                            st.session_state['predictions'][selected_dataset_name].append(entry)
                        st.session_state['processed_excel_files'].add(file_id)
                        st.success(f"Batch prediction successful using {mb['model_name']}!")
                    else:
                        st.error("Excel missing required feature columns.")
                except Exception as e:
                    st.error(f"Batch prediction failed: {e}")

    # Simplified Project List
    with st.expander('Simplified Project List', expanded=True):
        preds = st.session_state['predictions'].get(selected_dataset_name, [])
        if preds:
            if st.button('Delete All', key=f'delete_all_{selected_dataset_name}'):
                st.session_state['predictions'][selected_dataset_name] = []
                to_remove = {fid for fid in st.session_state['processed_excel_files'] if fid.endswith(selected_dataset_name)}
                for fid in to_remove:
                    st.session_state['processed_excel_files'].remove(fid)
                st.rerun()
            for i, p in enumerate(preds):
                c1, c2 = st.columns([3, 1])
                c1.write(p.get('Project Name', f"Project {i+1}"))
                if c2.button('Delete', key=f'del_{selected_dataset_name}_{i}'):
                    preds.pop(i)
                    st.rerun()
        else:
            st.write("No predictions yet.")

    # Prediction summary table
    st.header(f"Prediction Summary: {clean_name}")
    preds = st.session_state['predictions'].get(selected_dataset_name, [])
    if preds:
        df_preds = pd.DataFrame(preds)
        num_cols = df_preds.select_dtypes(include=[np.number]).columns
        df_preds_display = df_preds.copy()
        for col in num_cols:
            df_preds_display[col] = df_preds_display[col].apply(lambda x: format_with_commas(x))
        st.dataframe(df_preds_display, use_container_width=True)

        towrite = io.BytesIO()
        df_preds.to_excel(towrite, index=False, engine='openpyxl')
        towrite.seek(0)
        st.download_button(
            "Download Predictions as Excel",
            data=towrite,
            file_name=f"{selected_dataset_name}_predictions.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.write("No predictions available.")

# Tabbed interface
tab1, tab2, tab3 = st.tabs(["📁 Datasets & Models", "🏗️ Project Builder", "🔀 Compare Projects"])

with tab1:
    st.write("Use the sidebar to upload datasets and train models with automatic best model selection.")
    st.write("The system will evaluate 6 different models and automatically choose the best performer.")

with tab2:
    st.header("🏗️ Project Builder")
    colA, colB = st.columns([3, 2])
    with colA:
        new_project_name = st.text_input(
            "Project Name", placeholder="e.g., Project A",
            key="pb_project_name_input"
        )
    with colB:
        if new_project_name and new_project_name not in st.session_state["projects"]:
            if st.button("Create Project", key="pb_create_project_btn"):
                st.session_state["projects"][new_project_name] = {"components": [], "totals": {}, "currency": ""}
                st.success(f"Project '{new_project_name}' created.")

    if not st.session_state["datasets"]:
        st.info("Upload datasets in the sidebar first.")
    else:
        existing_projects = list(st.session_state["projects"].keys())
        proj_sel = st.selectbox(
            "Choose project",
            ([new_project_name] + existing_projects) if new_project_name else existing_projects,
            key="pb_project_select"
        )

        ds_names = sorted(st.session_state["datasets"].keys())
        dataset_for_comp = st.selectbox(
            "Dataset for this component", ds_names,
            key="pb_dataset_for_component"
        )

        default_label = st.session_state["component_labels"].get(dataset_for_comp, "")
        component_type = st.text_input(
            "Component type (Oil & Gas term)",
            value=(default_label or "FPSO / Pipeline / Wellhead / Subsea"),
            key=f"pb_component_type_{proj_sel or new_project_name or 'NoProject'}"
        )

        if dataset_for_comp and dataset_for_comp not in st.session_state["models"]:
            with st.expander("Train model for selected dataset", expanded=False):
                st.write("Quick train with automatic best model selection.")
                quick_test_size = st.slider("Test size", 0.05, 0.5, 0.2, step=0.01, key=f"quick_ts_{dataset_for_comp}")
                if st.button("⚙️ Quick Train", key=f"quick_train_btn_{dataset_for_comp}"):
                    with st.spinner("Training and selecting best model..."):
                        try:
                            bundle_q = train_best_model(st.session_state["datasets"][dataset_for_comp], test_size=float(quick_test_size))
                            st.session_state["models"][dataset_for_comp] = bundle_q
                            st.success(f"✅ Best Model: {bundle_q['model_name']} (R²: {bundle_q['metrics']['r2']:.3f})")
                        except Exception as e:
                            st.error(f"Training failed: {e}")

        if dataset_for_comp and dataset_for_comp in st.session_state["models"]:
            mb_comp = st.session_state["models"][dataset_for_comp]
            st.info(f"Using model: **{mb_comp['model_name']}** (R²: {mb_comp['metrics']['r2']:.3f})")
            
            st.markdown("**Component feature inputs**")
            feat_cols = mb_comp["features"]
            cols = st.columns(2)
            comp_inputs = {}
            for i, c in enumerate(feat_cols):
                key = f"pb_{proj_sel}_{dataset_for_comp}_feature_{i}_{c}"
                comp_inputs[c] = cols[i % 2].text_input(c, key=key)

            st.markdown("---")
            st.markdown("**Cost Percentage Inputs**")
            cp1, cp2, cp3, cp4, cp5 = st.columns(5)
            epcic_percentages_pb = {
                "Engineering": cp1.number_input("Engineering (%)", 0.0, 100.0, 0.0, key=f"pb_eng_{proj_sel}"),
                "Procurement": cp2.number_input("Procurement (%)", 0.0, 100.0, 0.0, key=f"pb_proc_{proj_sel}"),
                "Construction": cp3.number_input("Construction (%)", 0.0, 100.0, 0.0, key=f"pb_const_{proj_sel}"),
                "Installation": cp4.number_input("Installation (%)", 0.0, 100.0, 0.0, key=f"pb_inst_{proj_sel}"),
                "Commissioning": cp5.number_input("Commissioning (%)", 0.0, 100.0, 0.0, key=f"pb_comm_{proj_sel}"),
            }
            pd1, pd2 = st.columns(2)
            predev_pct = pd1.number_input("Pre-Development (%)", 0.0, 100.0, 0.0, key=f"pb_predev_{proj_sel}")
            owners_pct = pd2.number_input("Owner's Cost (%)", 0.0, 100.0, 0.0, key=f"pb_owners_{proj_sel}")
            ct1, ct2 = st.columns(2)
            contingency_pct = ct1.number_input("Contingency (%)", 0.0, 100.0, 0.0, key=f"pb_cont_{proj_sel}")
            escalation_pct = ct2.number_input("Escalation & Inflation (%)", 0.0, 100.0, 0.0, key=f"pb_escal_{proj_sel}")

            if proj_sel:
                if st.button("➕ Predict & Add Component", key=f"pb_add_comp_{proj_sel}_{dataset_for_comp}"):
                    if proj_sel not in st.session_state["projects"]:
                        st.error("Please create/select a project first.")
                    else:
                        row = {}
                        for f in mb_comp["features"]:
                            v = comp_inputs.get(f, "")
                            if v is None or str(v).strip() == "":
                                row[f] = np.nan
                            else:
                                try:
                                    row[f] = float(v)
                                except Exception:
                                    row[f] = np.nan
                        try:
                            df_input_raw = pd.DataFrame([row], columns=mb_comp["features"])
                            X_imputed = pd.DataFrame(mb_comp['imputer'].transform(df_input_raw), columns=mb_comp['features'])
                            X_scaled = mb_comp['scaler'].transform(X_imputed)
                            p = float(mb_comp['model'].predict(X_scaled)[0])
                            
                            epcic_breakdown = {}
                            for phase, pct in epcic_percentages_pb.items():
                                cost = round(p * (pct / 100.0), 2)
                                epcic_breakdown[phase] = {"cost": cost, "percentage": pct}
                            
                            predev_cost = round(p * (predev_pct / 100.0), 2)
                            owners_cost = round(p * (owners_pct / 100.0), 2)
                            contingency_base = p + owners_cost
                            contingency_cost = round(contingency_base * (contingency_pct / 100.0), 2)
                            escalation_base = p + owners_cost
                            escalation_cost = round(escalation_base * (escalation_pct / 100.0), 2)
                            grand_total = round(p + owners_cost + contingency_cost + escalation_cost, 2)

                            comp_entry = {
                                "component_type": component_type or default_label or "Component",
                                "dataset": dataset_for_comp,
                                "model_used": mb_comp['model_name'],
                                "inputs": {k: v for k, v in row.items()},
                                "prediction": p,
                                "breakdown": {
                                    "epcic": epcic_breakdown,
                                    "predev_cost": predev_cost,
                                    "owners_cost": owners_cost,
                                    "contingency_cost": contingency_cost,
                                    "escalation_cost": escalation_cost,
                                    "grand_total": grand_total,
                                    "target_col": mb_comp['target']
                                }
                            }
                            st.session_state["projects"][proj_sel]["components"].append(comp_entry)
                            st.session_state["component_labels"][dataset_for_comp] = component_type or default_label
                            if not st.session_state["projects"][proj_sel]["currency"]:
                                st.session_state["projects"][proj_sel]["currency"] = currency
                            st.success(f"✅ Added {comp_entry['component_type']} using {mb_comp['model_name']}")
                        except Exception as e:
                            st.error(f"Failed to predict: {e}")

        st.markdown("---")
        st.subheader("Current Project Overview")
        if st.session_state["projects"]:
            sel = st.selectbox("View/Edit project", list(st.session_state["projects"].keys()), key="pb_view_project_select")
            proj = st.session_state["projects"][sel]
            comps = proj.get("components", [])
            if not comps:
                st.info("No components yet. Add one above.")
            else:
                rows = []
                for c in comps:
                    rows.append({
                        "Component": f"{c['component_type']}",
                        "Dataset": c["dataset"],
                        "Model": c.get("model_used", "N/A"),
                        "Predicted CAPEX": c["prediction"],
                        "Grand Total": c["breakdown"]["grand_total"]
                    })
                dfc = pd.DataFrame(rows)
                st.dataframe(dfc, use_container_width=True)
                total_capex = float(sum(r["Predicted CAPEX"] for r in rows))
                total_grand = float(sum(r["Grand Total"] for r in rows))
                proj["totals"] = {"capex_sum": total_capex, "grand_total": total_grand}
                col_t1, col_t2 = st.columns(2)
                col_t1.metric("Project CAPEX", f"{total_capex:,.2f}")
                col_t2.metric("Project Grand Total", f"{total_grand:,.2f}")

                st.subheader("📊 Component Composition")
                stack_rows = []
                for c in comps:
                    capex = float(c["prediction"])
                    owners = float(c["breakdown"]["owners_cost"])
                    cont = float(c["breakdown"]["contingency_cost"])
                    escal = float(c["breakdown"]["escalation_cost"])
                    predev = float(c["breakdown"]["predev_cost"])
                    stack_rows.append({"Component": c["component_type"], "CAPEX": capex, "Owners": owners,
                                       "Contingency": cont, "Escalation": escal, "PreDev": predev})
                dfc_stack = pd.DataFrame(stack_rows).groupby("Component", as_index=False).sum()
                if not dfc_stack.empty:
                    cats = ["CAPEX", "Owners", "Contingency", "Escalation", "PreDev"]
                    colors = ["#2E86AB","#6C757D","#C0392B","#17A589","#9B59B6"]
                    x = np.arange(len(dfc_stack["Component"]))
                    fig, ax = plt.subplots(figsize=(min(12, 1.2*len(x)+6), 6))
                    bottom = np.zeros(len(x))
                    for ccol, colname in zip(colors, cats):
                        vals = dfc_stack[colname].values
                        ax.bar(x, vals, bottom=bottom, color=ccol, label=colname)
                        bottom += vals
                    ax.set_xticks(x)
                    ax.set_xticklabels(dfc_stack["Component"], rotation=30, ha="right")
                    ax.set_ylabel(f"Cost {proj.get('currency','')}".strip())
                    ax.set_title("Component Composition (Grand Total)")
                    ax.legend(ncols=5, fontsize=9, loc="upper center", bbox_to_anchor=(0.5, 1.12))
                    ax.grid(axis="y", linestyle="--", alpha=0.4)
                    fig.tight_layout()
                    st.pyplot(fig, use_container_width=True)
                    plt.close()

                for idx, c in enumerate(comps):
                    cc1, cc2, cc3 = st.columns([6, 3, 1])
                    cc1.write(f"**{c['component_type']}** — *{c['dataset']}* — Model: {c.get('model_used', 'N/A')}")
                    cc2.write(f"Grand Total: {c['breakdown']['grand_total']:,.2f}")
                    if cc3.button("🗑️", key=f"pb_del_comp_{sel}_{idx}"):
                        comps.pop(idx)
                        st.rerun()

                st.markdown("---")
                st.subheader("📥 Download Project Reports")
                
                col_dl1, col_dl2 = st.columns(2)
                
                with col_dl1:
                    excel_report = create_project_excel_report(sel, proj, proj.get("currency", ""))
                    st.download_button(
                        "⬇️ Download Excel Report",
                        data=excel_report,
                        file_name=f"{sel}_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                with col_dl2:
                    pptx_report = create_project_pptx_report(sel, proj, proj.get("currency", ""))
                    st.download_button(
                        "⬇️ Download PowerPoint Report",
                        data=pptx_report,
                        file_name=f"{sel}_Report.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
                
                st.download_button(
                    "⬇️ Download Project (JSON)",
                    data=json.dumps(st.session_state["projects"][sel], indent=2),
                    file_name=f"{sel}.json",
                    mime="application/json"
                )
                
                up = st.file_uploader("Import project JSON", type=["json"], key=f"pb_import_json_{sel}")
                if up is not None:
                    try:
                        data = json.load(up)
                        st.session_state["projects"][sel] = data
                        st.success("Project imported successfully.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Failed to import: {e}")

with tab3:
    st.header("🔀 Compare Projects")
    proj_names = list(st.session_state["projects"].keys())
    if len(proj_names) < 2:
        st.info("Create at least two projects to compare.")
    else:
        compare = st.multiselect("Pick projects to compare", proj_names, default=proj_names[:2], key="compare_select")
        if len(compare) >= 2:
            comp_rows = []
            comparison_dict = {}
            
            for p in compare:
                proj = st.session_state["projects"][p]
                grand_total = proj.get("totals", {}).get("grand_total", 0.0)
                capex_sum = proj.get("totals", {}).get("capex_sum", 0.0)
                num_components = len(proj.get("components", []))
                
                comp_rows.append({
                    "Project": p,
                    "Components": num_components,
                    "CAPEX Sum": capex_sum,
                    "Grand Total": grand_total
                })
                
                comparison_dict[p] = {
                    "capex": capex_sum,
                    "grand_total": grand_total,
                    "components": num_components
                }
            
            dft = pd.DataFrame(comp_rows).set_index("Project")
            st.dataframe(dft.style.format({
                "CAPEX Sum": "{:,.2f}",
                "Grand Total": "{:,.2f}"
            }), use_container_width=True)

            # Grand totals bar chart
            st.subheader("📊 Project Grand Totals")
            fig, ax = plt.subplots(figsize=(min(12, 1.2*len(dft)+6), 5))
            bars = ax.bar(dft.index, dft["Grand Total"], color="#2E86AB")
            ax.set_title("Project Grand Totals Comparison")
            ax.set_ylabel("Cost")
            ax.grid(axis="y", linestyle="--", alpha=0.4)
            ax.tick_params(axis='x', rotation=25)
            for b in bars:
                ax.annotate(f"{b.get_height():,.0f}", (b.get_x()+b.get_width()/2, b.get_height()),
                            ha='center', va='bottom', fontsize=9)
            fig.tight_layout()
            st.pyplot(fig, use_container_width=True)
            plt.close()

            # Project composition stacked bar
            comp_rows = []
            for p in compare:
                proj = st.session_state["projects"][p]
                comps = proj.get("components", [])
                capex = owners = cont = escal = predev = 0.0
                for c in comps:
                    capex += float(c["prediction"])
                    owners += float(c["breakdown"]["owners_cost"])
                    cont += float(c["breakdown"]["contingency_cost"])
                    escal += float(c["breakdown"]["escalation_cost"])
                    predev += float(c["breakdown"]["predev_cost"])
                comp_rows.append({"Project": p, "CAPEX": capex, "Owners": owners,
                                  "Contingency": cont, "Escalation": escal, "PreDev": predev})
            
            comp_df = pd.DataFrame(comp_rows).set_index("Project").reset_index()
            st.subheader("🏗️ Project Composition (Stacked)")
            cats = ["CAPEX","Owners","Contingency","Escalation","PreDev"]
            colors = ["#2E86AB","#6C757D","#C0392B","#17A589","#9B59B6"]
            x = np.arange(len(comp_df["Project"]))
            fig, ax = plt.subplots(figsize=(min(12, 1.2*len(x)+6), 6))
            bottom = np.zeros(len(x))
            for ccol, colname in zip(colors, cats):
                vals = comp_df[colname].values
                ax.bar(x, vals, bottom=bottom, color=ccol, label=colname)
                bottom += vals
            ax.set_xticks(x)
            ax.set_xticklabels(comp_df["Project"], rotation=25, ha="right")
            ax.set_ylabel("Cost")
            ax.set_title("Project Composition (Grand Total)")
            ax.legend(ncols=5, fontsize=9, loc="upper center", bbox_to_anchor=(0.5, 1.12))
            ax.grid(axis="y", linestyle="--", alpha=0.4)
            fig.tight_layout()
            st.pyplot(fig, use_container_width=True)
            plt.close()

            # Difference waterfall (if exactly 2 projects)
            if len(compare) == 2:
                p1, p2 = compare
                comp2 = comp_df[comp_df["Project"] == p2].iloc[0]
                comp1 = comp_df[comp_df["Project"] == p1].iloc[0]
                delta_dict = {
                    "CAPEX": float(comp2["CAPEX"] - comp1["CAPEX"]),
                    "Owners": float(comp2["Owners"] - comp1["Owners"]),
                    "Contingency": float(comp2["Contingency"] - comp1["Contingency"]),
                    "Escalation": float(comp2["Escalation"] - comp1["Escalation"]),
                    "PreDev": float(comp2["PreDev"] - comp1["PreDev"]),
                }
                st.subheader(f"🔀 Difference Breakdown: {p2} vs {p1}")
                cats = ["CAPEX","Owners","Contingency","Escalation","PreDev"]
                vals = [delta_dict.get(k,0.0) for k in cats]
                cum = np.cumsum([0]+vals[:-1])
                colors_diff = ["#2E86AB" if v>=0 else "#C0392B" for v in vals]
                fig, ax = plt.subplots(figsize=(8, 5))
                for i,(v, base, ccol) in enumerate(zip(vals, cum, colors_diff)):
                    ax.bar(i, v, bottom=base, color=ccol)
                    ax.annotate(f"{v:+,.0f}", (i, base + v/2), ha="center", va="center", color="white", fontsize=9)
                ax.set_xticks(range(len(cats)))
                ax.set_xticklabels(cats, rotation=0)
                total = sum(vals)
                ax.axhline(0, color="black", linewidth=0.8)
                ax.set_title(f"Δ Grand Total Waterfall (Total Δ = {total:+,.0f})")
                ax.set_ylabel("Δ Cost")
                ax.grid(axis="y", linestyle="--", alpha=0.4)
                fig.tight_layout()
                st.pyplot(fig, use_container_width=True)
                plt.close()

            # Component-level comparison
            st.subheader("📋 Component-Level Details")
            for p in compare:
                with st.expander(f"Project: {p}"):
                    proj = st.session_state["projects"][p]
                    comps = proj.get("components", [])
                    if comps:
                        comp_detail_rows = []
                        for c in comps:
                            epcic_str = ", ".join([f"{phase}: {details['cost']:,.0f}" 
                                                  for phase, details in c["breakdown"]["epcic"].items() 
                                                  if details['percentage'] > 0])
                            comp_detail_rows.append({
                                "Component": c["component_type"],
                                "Dataset": c["dataset"],
                                "Model": c.get("model_used", "N/A"),
                                "CAPEX": c["prediction"],
                                "Owner's": c["breakdown"]["owners_cost"],
                                "Contingency": c["breakdown"]["contingency_cost"],
                                "Escalation": c["breakdown"]["escalation_cost"],
                                "Grand Total": c["breakdown"]["grand_total"],
                                "EPCIC": epcic_str
                            })
                        comp_detail_df = pd.DataFrame(comp_detail_rows)
                        st.dataframe(comp_detail_df, use_container_width=True)
                    else:
                        st.info("No components in this project.")

            # Download comparison reports
            st.markdown("---")
            st.subheader("📥 Download Comparison Reports")
            
            col_comp1, col_comp2 = st.columns(2)
            
            with col_comp1:
                projects_to_compare = {p: st.session_state["projects"][p] for p in compare}
                currency_comp = st.session_state["projects"][compare[0]].get("currency", "")
                excel_comp = create_comparison_excel_report(projects_to_compare, currency_comp)
                st.download_button(
                    "⬇️ Download Comparison Excel",
                    data=excel_comp,
                    file_name=f"Project_Comparison_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            with col_comp2:
                # Create combined PowerPoint with comparison data
                pptx_comp = create_project_pptx_report(
                    f"Comparison: {', '.join(compare)}",
                    st.session_state["projects"][compare[0]],
                    currency_comp,
                    comparison_dict
                )
                st.download_button(
                    "⬇️ Download Comparison PowerPoint",
                    data=pptx_comp,
                    file_name=f"Project_Comparison_{datetime.now().strftime('%Y%m%d')}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )

    # --- PRESCRIPTIVE ADVISOR (ENHANCED CHAT STYLE, BUTTON-BASED QUESTIONS) ---
    st.markdown("---")
    st.header("🤖 Prescriptive Advisor")

    # Simple chat-style bubbles using HTML/CSS
    st.markdown(
        """
        <style>
        .chat-row {
            display: flex;
            margin-bottom: 0.5rem;
        }
        .chat-bubble {
            padding: 0.75rem 1rem;
            border-radius: 1rem;
            max-width: 80%;
            font-size: 0.9rem;
            line-height: 1.4;
        }
        .chat-bubble-user {
            margin-left: auto;
            background-color: #0d6efd;
            color: white;
        }
        .chat-bubble-bot {
            margin-right: auto;
            background-color: #f1f3f5;
            color: #212529;
        }
        .chat-name {
            font-size: 0.7rem;
            font-weight: 600;
            opacity: 0.7;
            margin-bottom: 0.25rem;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    def chat_user(message: str):
        st.markdown(
            f"""
            <div class="chat-row">
                <div class="chat-bubble chat-bubble-user">
                    <div class="chat-name">You</div>
                    <div>{message}</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    def chat_bot(message: str):
        st.markdown(
            f"""
            <div class="chat-row">
                <div class="chat-bubble chat-bubble-bot">
                    <div class="chat-name">Advisor</div>
                    <div>{message}</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    proj_names = list(st.session_state["projects"].keys())
    if not proj_names:
        st.info("Create at least one project in the Project Builder to get prescriptive advice.")
    else:
        # Choose which projects this advisor should reason over
        selected_for_advice = st.multiselect(
            "Select projects for advice",
            proj_names,
            default=proj_names,
            key="advisor_projects_select"
        )

        if not selected_for_advice:
            st.warning("Select at least one project to continue.")
        else:
            # Build a stats dict for the selected projects
            advisor_stats = {}
            for p in selected_for_advice:
                proj = st.session_state["projects"][p]
                comps = proj.get("components", [])
                capex_sum = 0.0
                grand_total = 0.0
                owners = 0.0
                contingency = 0.0
                escalation = 0.0
                predev = 0.0

                for c in comps:
                    capex_sum += float(c["prediction"])
                    grand_total += float(c["breakdown"]["grand_total"])
                    owners += float(c["breakdown"]["owners_cost"])
                    contingency += float(c["breakdown"]["contingency_cost"])
                    escalation += float(c["breakdown"]["escalation_cost"])
                    predev += float(c["breakdown"]["predev_cost"])

                advisor_stats[p] = {
                    "components": len(comps),
                    "capex_sum": capex_sum,
                    "grand_total": grand_total,
                    "owners": owners,
                    "contingency": contingency,
                    "escalation": escalation,
                    "predev": predev,
                    "currency": proj.get("currency", "")
                }

            st.subheader("💬 Question Library")

            # You can add / remove questions here (no NLP, all pre-defined)
            question_options = [
                "Which project is most cost-effective (lowest Grand Total)?",
                "Which project requires the largest investment (highest Grand Total)?",
                "Which project has the highest average cost per component?",
                "Within a project, which components are the main cost drivers?",
                "If we must cut 10% of budget, which components are the best candidates?",
                "Which project has the highest proportion of indirect costs (Owner's + Pre-Dev + Contingency + Escalation)?",
                "Across selected projects, which components are the top 10 most expensive?",
                "Suggest a primary project and backup option based on cost and component count.",
            ]

            # A small "button-like" feeling using radio + selectbox
            question_group = st.radio(
                "Question type",
                ["Project-level decisions", "Component-level insights"],
                horizontal=True,
                key="advisor_question_group"
            )

            if question_group == "Project-level decisions":
                subset = [
                    q for q in question_options
                    if "project" in q.lower() or "option" in q.lower()
                ]
            else:
                subset = [
                    q for q in question_options
                    if "component" in q.lower() or "budget" in q.lower()
                ]

            question = st.selectbox(
                "Pick a question for the advisor to answer",
                subset,
                key="advisor_question_select"
            )

            # Chat-style display
            chat_user(question)

            # ---------- Q1: Most cost-effective project ----------
            if question == "Which project is most cost-effective (lowest Grand Total)?":
                best_proj = min(advisor_stats.items(), key=lambda x: x[1]["grand_total"])
                name, stats = best_proj
                cur = stats["currency"]
                msg = (
                    f"Based on <b>Grand Total</b>, the most cost-effective project is "
                    f"<b>{name}</b>.<br><br>"
                    f"- Grand Total: <b>{cur} {stats['grand_total']:,.2f}</b><br>"
                    f"- Components: {stats['components']}<br>"
                    f"- Total base CAPEX: {cur} {stats['capex_sum']:,.2f}"
                )
                chat_bot(msg)

            # ---------- Q2: Largest investment ----------
            elif question == "Which project requires the largest investment (highest Grand Total)?":
                worst_proj = max(advisor_stats.items(), key=lambda x: x[1]["grand_total"])
                name, stats = worst_proj
                cur = stats["currency"]
                msg = (
                    f"<b>{name}</b> currently requires the <b>largest overall investment</b>.<br><br>"
                    f"- Grand Total: <b>{cur} {stats['grand_total']:,.2f}</b><br>"
                    f"- Components: {stats['components']}<br>"
                    f"- Total base CAPEX: {cur} {stats['capex_sum']:,.2f}"
                )
                chat_bot(msg)

            # ---------- Q3: Highest average cost per component ----------
            elif question == "Which project has the highest average cost per component?":
                avg_costs = []
                for name, stats in advisor_stats.items():
                    if stats["components"] > 0:
                        avg = stats["grand_total"] / stats["components"]
                    else:
                        avg = 0.0
                    avg_costs.append((name, avg, stats["currency"], stats["components"], stats["grand_total"]))

                name, avg, cur, comp_count, total_gt = max(avg_costs, key=lambda x: x[1])
                msg = (
                    f"<b>{name}</b> has the <b>highest average Grand Total per component</b>.<br><br>"
                    f"- Components: {comp_count}<br>"
                    f"- Grand Total: {cur} {total_gt:,.2f}<br>"
                    f"- <b>Average per component</b>: {cur} {avg:,.2f}<br><br>"
                    f"This typically indicates more complex or expensive components on average."
                )
                chat_bot(msg)

                # Small comparison table
                rows = []
                for n, stats in advisor_stats.items():
                    if stats["components"] > 0:
                        rows.append({
                            "Project": n,
                            "Components": stats["components"],
                            "Grand Total": stats["grand_total"],
                            "Avg per Component": stats["grand_total"] / stats["components"]
                        })
                if rows:
                    df_avg = pd.DataFrame(rows)
                    st.dataframe(
                        df_avg.style.format({
                            "Grand Total": "{:,.2f}",
                            "Avg per Component": "{:,.2f}"
                        }),
                        use_container_width=True
                    )

            # ---------- Q4: Main cost drivers within a project ----------
            elif question == "Within a project, which components are the main cost drivers?":
                proj_choice = st.selectbox(
                    "Select project to analyse",
                    selected_for_advice,
                    key="advisor_cost_driver_project"
                )
                proj = st.session_state["projects"][proj_choice]
                comps = proj.get("components", [])
                if not comps:
                    chat_bot("This project has no components yet. Add components in the Project Builder first.")
                else:
                    rows = []
                    for c in comps:
                        rows.append({
                            "Component": c["component_type"],
                            "Dataset": c["dataset"],
                            "Model": c.get("model_used", "N/A"),
                            "CAPEX": float(c["prediction"]),
                            "Owners": float(c["breakdown"]["owners_cost"]),
                            "Contingency": float(c["breakdown"]["contingency_cost"]),
                            "Escalation": float(c["breakdown"]["escalation_cost"]),
                            "PreDev": float(c["breakdown"]["predev_cost"]),
                            "Grand Total": float(c["breakdown"]["grand_total"])
                        })
                    df_comp = pd.DataFrame(rows).sort_values("Grand Total", ascending=False)
                    cur = proj.get("currency", "")

                    top = df_comp.iloc[0]
                    msg = (
                        f"In <b>{proj_choice}</b>, the main cost driver (by Grand Total) is:<br><br>"
                        f"- Component: <b>{top['Component']}</b><br>"
                        f"- Grand Total: {cur} {top['Grand Total']:,.2f}<br>"
                        f"- Base CAPEX: {cur} {top['CAPEX']:,.2f}<br>"
                        f"- Owner's: {cur} {top['Owners']:,.2f}<br>"
                        f"- Contingency: {cur} {top['Contingency']:,.2f}<br>"
                        f"- Escalation: {cur} {top['Escalation']:,.2f}<br>"
                        f"- Pre-Dev: {cur} {top['PreDev']:,.2f}"
                    )
                    chat_bot(msg)

                    st.write("Top cost drivers (sorted by Grand Total):")
                    st.dataframe(
                        df_comp.style.format({
                            "CAPEX": "{:,.2f}",
                            "Owners": "{:,.2f}",
                            "Contingency": "{:,.2f}",
                            "Escalation": "{:,.2f}",
                            "PreDev": "{:,.2f}",
                            "Grand Total": "{:,.2f}",
                        }),
                        use_container_width=True,
                    )

            # ---------- Q5: 10% budget cut recommendation ----------
            elif question == "If we must cut 10% of budget, which components are the best candidates?":
                proj_choice = st.selectbox(
                    "Select project to stress-test",
                    selected_for_advice,
                    key="advisor_budget_cut_project"
                )
                proj = st.session_state["projects"][proj_choice]
                comps = proj.get("components", [])
                if not comps:
                    chat_bot("This project has no components yet, so no budget cut analysis is possible.")
                else:
                    rows = []
                    total_gt = 0.0
                    for c in comps:
                        gt = float(c["breakdown"]["grand_total"])
                        total_gt += gt
                        rows.append({
                            "Component": c["component_type"],
                            "Grand Total": gt,
                            "Contingency": float(c["breakdown"]["contingency_cost"]),
                            "Escalation": float(c["breakdown"]["escalation_cost"])
                        })
                    df_comp = pd.DataFrame(rows).sort_values("Grand Total", ascending=False)
                    target_cut = 0.10 * total_gt
                    cur = proj.get("currency", "")

                    msg = (
                        f"We target a <b>10% reduction</b> of <b>{proj_choice}</b> Grand Total.<br><br>"
                        f"- Current Grand Total: {cur} {total_gt:,.2f}<br>"
                        f"- Required reduction (10%): <b>{cur} {target_cut:,.2f}</b><br><br>"
                        f"A practical starting point is to review components with large "
                        f"contingency + escalation (\"flexible\" cost pool) before touching base CAPEX."
                    )
                    chat_bot(msg)

                    df_comp["Flex_Pool"] = df_comp["Contingency"] + df_comp["Escalation"]
                    df_comp = df_comp.sort_values("Flex_Pool", ascending=False)

                    st.write("Recommended components to review first (high contingency + escalation):")
                    st.dataframe(
                        df_comp.head(5).style.format({
                            "Grand Total": "{:,.2f}",
                            "Contingency": "{:,.2f}",
                            "Escalation": "{:,.2f}",
                            "Flex_Pool": "{:,.2f}",
                        }),
                        use_container_width=True,
                    )

                    possible_saving = df_comp["Flex_Pool"].sum()
                    st.info(
                        f"Total contingency + escalation across all components: "
                        f"{cur} {possible_saving:,.2f}. "
                        f"This pool can be optimised before reducing base scope."
                    )

            # ---------- Q6: Highest proportion of indirect costs ----------
            elif question == "Which project has the highest proportion of indirect costs (Owner's + Pre-Dev + Contingency + Escalation)?":
                ratios = []
                for name, stats in advisor_stats.items():
                    indirect = stats["owners"] + stats["predev"] + stats["contingency"] + stats["escalation"]
                    gt = stats["grand_total"]
                    if gt > 0:
                        ratio = indirect / gt
                    else:
                        ratio = 0.0
                    ratios.append((name, ratio, indirect, gt, stats["currency"]))

                name, ratio, indirect, gt, cur = max(ratios, key=lambda x: x[1])

                msg = (
                    f"<b>{name}</b> has the <b>highest proportion of indirect costs</b> "
                    f"(Owner's + Pre-Dev + Contingency + Escalation).<br><br>"
                    f"- Indirect costs: {cur} {indirect:,.2f}<br>"
                    f"- Grand Total: {cur} {gt:,.2f}<br>"
                    f"- Indirect / Grand Total: <b>{ratio*100:,.1f}%</b><br><br>"
                    f"This project is a strong candidate for optimising indirect cost assumptions."
                )
                chat_bot(msg)

                rows = []
                for n, stats in advisor_stats.items():
                    indirect_n = stats["owners"] + stats["predev"] + stats["contingency"] + stats["escalation"]
                    gt_n = stats["grand_total"]
                    ratio_n = indirect_n / gt_n if gt_n > 0 else 0.0
                    rows.append({
                        "Project": n,
                        "Indirect Costs": indirect_n,
                        "Grand Total": gt_n,
                        "Indirect %": ratio_n * 100,
                    })
                df_indirect = pd.DataFrame(rows).sort_values("Indirect %", ascending=False)
                st.dataframe(
                    df_indirect.style.format({
                        "Indirect Costs": "{:,.2f}",
                        "Grand Total": "{:,.2f}",
                        "Indirect %": "{:,.1f}",
                    }),
                    use_container_width=True,
                )

            # ---------- Q7: Top 10 most expensive components across projects ----------
            elif question == "Across selected projects, which components are the top 10 most expensive?":
                comp_rows = []
                for p in selected_for_advice:
                    proj = st.session_state["projects"][p]
                    cur = proj.get("currency", "")
                    for c in proj.get("components", []):
                        comp_rows.append({
                            "Project": p,
                            "Component": c["component_type"],
                            "Dataset": c["dataset"],
                            "Grand Total": float(c["breakdown"]["grand_total"]),
                            "CAPEX": float(c["prediction"]),
                            "Owners": float(c["breakdown"]["owners_cost"]),
                            "Contingency": float(c["breakdown"]["contingency_cost"]),
                            "Escalation": float(c["breakdown"]["escalation_cost"]),
                            "PreDev": float(c["breakdown"]["predev_cost"]),
                            "Currency": cur,
                        })
                if not comp_rows:
                    chat_bot("No components found in the selected projects.")
                else:
                    df_top = pd.DataFrame(comp_rows).sort_values("Grand Total", ascending=False)
                    msg = (
                        "Here are the <b>top 10 most expensive components</b> across the selected projects, "
                        "ranked by Grand Total. These are your primary levers if you need major cost impact."
                    )
                    chat_bot(msg)

                    st.dataframe(
                        df_top.head(10).style.format({
                            "Grand Total": "{:,.2f}",
                            "CAPEX": "{:,.2f}",
                            "Owners": "{:,.2f}",
                            "Contingency": "{:,.2f}",
                            "Escalation": "{:,.2f}",
                            "PreDev": "{:,.2f}",
                        }),
                        use_container_width=True,
                    )

            # ---------- Q8: Recommend primary & backup project ----------
            elif question == "Suggest a primary project and backup option based on cost and component count.":
                # Heuristic: low Grand Total, but not too few components
                scored = []
                for name, stats in advisor_stats.items():
                    gt = stats["grand_total"]
                    comps = stats["components"]
                    score = gt / (comps + 1)  # simple ratio to balance both
                    scored.append((name, score, stats))

                scored = sorted(scored, key=lambda x: x[1])
                if not scored:
                    chat_bot("No suitable projects to compare.")
                else:
                    primary_name, primary_score, primary_stats = scored[0]
                    backup_msg = "No backup (only one project selected)."
                    backup_name = None
                    if len(scored) > 1:
                        backup_name, backup_score, backup_stats = scored[1]
                        backup_msg = (
                            f"Backup option: <b>{backup_name}</b> – "
                            f"Grand Total {backup_stats['currency']} {backup_stats['grand_total']:,.2f}, "
                            f"{backup_stats['components']} components."
                        )

                    cur = primary_stats["currency"]
                    msg = (
                        f"My recommendation based on cost and component count:<br><br>"
                        f"🔹 <b>Primary candidate:</b> <b>{primary_name}</b><br>"
                        f"&nbsp;&nbsp;&nbsp;- Grand Total: {cur} {primary_stats['grand_total']:,.2f}<br>"
                        f"&nbsp;&nbsp;&nbsp;- Components: {primary_stats['components']}<br>"
                        f"&nbsp;&nbsp;&nbsp;- Base CAPEX: {cur} {primary_stats['capex_sum']:,.2f}<br><br>"
                        f"🔸 {backup_msg}<br><br>"
                        f"This rule-based suggestion prefers projects with lower total cost "
                        f"while still having enough components to represent a realistic scope."
                    )
                    chat_bot(msg)

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #6c757d; padding: 20px;'>
    <p><strong>CAPEX AI RT2025</strong> | Powered by Automatic ML Model Selection & Prescriptive Advisor</p>
    <p>📊 6 Models Evaluated | 🎯 Best Model Auto-Selected | 💬 Rule-Based Cost Engineering Insights</p>
</div>
""", unsafe_allow_html=True)
