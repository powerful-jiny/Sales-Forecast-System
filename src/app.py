import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from pathlib import Path
import io

# ──────────────────────────────────────────────────────────────
# Page configuration
# ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Sales Forecast System",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ──────────────────────────────────────────────────────────────
# Custom CSS
# ──────────────────────────────────────────────────────────────
st.markdown(
    """
    <style>
    /* ── Global ── */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    .main { background-color: #f0f4f8; }

    /* ── Sidebar ── */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #1a2e4a 0%, #0d1b2e 100%);
        color: white;
    }
    section[data-testid="stSidebar"] .stRadio label,
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] span {
        color: #c8d8e8 !important;
    }
    section[data-testid="stSidebar"] .stRadio [data-testid="stMarkdownContainer"] p {
        font-size: 15px;
        font-weight: 500;
    }

    /* ── Header banner ── */
    .header-banner {
        background: linear-gradient(135deg, #1a2e4a 0%, #2563eb 100%);
        padding: 28px 36px;
        border-radius: 14px;
        margin-bottom: 28px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.15);
    }
    .header-title {
        color: white;
        font-size: 30px;
        font-weight: 700;
        margin: 0;
    }
    .header-subtitle {
        color: #93c5fd;
        font-size: 14px;
        margin-top: 6px;
    }

    /* ── KPI card ── */
    .kpi-card {
        background: white;
        border-radius: 12px;
        padding: 22px 24px;
        box-shadow: 0 2px 12px rgba(0,0,0,0.08);
        border-left: 5px solid #2563eb;
        height: 100%;
    }
    .kpi-label {
        color: #64748b;
        font-size: 13px;
        font-weight: 500;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
    .kpi-value {
        color: #1e293b;
        font-size: 28px;
        font-weight: 700;
        margin: 6px 0 4px;
    }
    .kpi-delta-pos { color: #22c55e; font-size: 13px; }
    .kpi-delta-neg { color: #ef4444; font-size: 13px; }
    .kpi-blue   { border-left-color: #2563eb; }
    .kpi-green  { border-left-color: #22c55e; }
    .kpi-orange { border-left-color: #f97316; }
    .kpi-purple { border-left-color: #8b5cf6; }

    /* ── Content card ── */
    .content-card {
        background: white;
        border-radius: 12px;
        padding: 24px;
        box-shadow: 0 2px 12px rgba(0,0,0,0.08);
        margin-bottom: 20px;
    }
    .section-title {
        color: #1e293b;
        font-size: 18px;
        font-weight: 600;
        margin-bottom: 16px;
        padding-bottom: 10px;
        border-bottom: 2px solid #e2e8f0;
    }

    /* ── Upload zone ── */
    .upload-zone {
        border: 2px dashed #93c5fd;
        border-radius: 10px;
        padding: 28px;
        text-align: center;
        background: #eff6ff;
    }

    /* ── Progress bar label ── */
    .progress-label {
        display: flex;
        justify-content: space-between;
        font-size: 13px;
        color: #475569;
        margin-bottom: 4px;
    }

    /* ── Tab headers ── */
    .stTabs [data-baseweb="tab-list"] {
        gap: 6px;
        background: #f1f5f9;
        border-radius: 10px;
        padding: 4px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        padding: 8px 18px;
        font-weight: 500;
    }
    .stTabs [aria-selected="true"] {
        background: white;
        box-shadow: 0 1px 4px rgba(0,0,0,0.1);
    }

    /* ── Metric override ── */
    [data-testid="stMetricValue"] { font-size: 24px !important; font-weight: 700; }
    [data-testid="stMetricLabel"] { font-size: 13px !important; color: #64748b; }

    /* ── Hide default streamlit chrome ── */
    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ──────────────────────────────────────────────────────────────
# Helpers
# ──────────────────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent.parent
TEMPLATES_DIR = BASE_DIR / "templates"
DATA_DIR = BASE_DIR / "data"

MONTH_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def load_excel(path: Path) -> pd.DataFrame:
    """Load first sheet from an Excel file, return empty DataFrame on error."""
    try:
        return pd.read_excel(path, header=0)
    except (FileNotFoundError, ValueError, Exception) as e:
        st.warning(f"Could not load {path.name}: {e}")
        return pd.DataFrame()


def load_uploaded(file, file_type: str = "xlsx") -> pd.DataFrame:
    """Load uploaded file (xlsx or csv) into a DataFrame."""
    if file is None:
        return pd.DataFrame()
    try:
        if file_type == "xlsx" or file.name.endswith((".xlsx", ".xls")):
            return pd.read_excel(file, header=0)
        return pd.read_csv(file)
    except Exception as e:
        st.error(f"Error reading file: {e}")
        return pd.DataFrame()


def fmt_currency(v) -> str:
    """Format a number as a compact currency string."""
    try:
        v = float(v)
    except (TypeError, ValueError):
        return "–"
    if abs(v) >= 1_000_000:
        return f"${v/1_000_000:.1f}M"
    if abs(v) >= 1_000:
        return f"${v/1_000:.0f}K"
    return f"${v:,.0f}"


def fmt_number(v) -> str:
    try:
        return f"{float(v):,.0f}"
    except (TypeError, ValueError):
        return "–"


def kpi_card(label: str, value: str, delta: str = "", color: str = "blue") -> str:
    delta_html = ""
    if delta:
        css = "kpi-delta-pos" if delta.startswith("+") else "kpi-delta-neg"
        delta_html = f'<div class="{css}">{delta}</div>'
    return (
        f'<div class="kpi-card kpi-{color}">'
        f'  <div class="kpi-label">{label}</div>'
        f'  <div class="kpi-value">{value}</div>'
        f'  {delta_html}'
        f'</div>'
    )


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return buf.getvalue()


def section(title: str) -> None:
    st.markdown(f'<div class="section-title">{title}</div>', unsafe_allow_html=True)


def empty_state(message: str) -> None:
    st.markdown(
        f'<div style="text-align:center;padding:48px 0;color:#94a3b8;font-size:15px;">'
        f'📂 {message}</div>',
        unsafe_allow_html=True,
    )


# ──────────────────────────────────────────────────────────────
# Sidebar
# ──────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(
        '<div style="text-align:center;padding:20px 0 10px;">'
        '<span style="font-size:42px;">📊</span>'
        '<p style="color:white;font-size:18px;font-weight:700;margin:8px 0 0;">Sales Forecast</p>'
        '<p style="color:#93c5fd;font-size:12px;margin:2px 0 0;">Management System</p>'
        '</div>',
        unsafe_allow_html=True,
    )
    st.markdown("---")

    menu = st.radio(
        "Navigation",
        options=[
            "📈  Annual Plan",
            "📊  6M Forecast",
            "📅  Weekly Actual",
            "🔧  Master",
            "📉  Summary",
        ],
        label_visibility="collapsed",
    )

    st.markdown("---")
    st.markdown(
        '<p style="color:#64748b;font-size:11px;text-align:center;">v1.0 · Sales Forecast System</p>',
        unsafe_allow_html=True,
    )

# ──────────────────────────────────────────────────────────────
# Header banner (shared)
# ──────────────────────────────────────────────────────────────
PAGE_META = {
    "📈  Annual Plan":   ("📈 Annual Plan",   "Full-year revenue & quantity plan by account and product"),
    "📊  6M Forecast":   ("📊 6M Forecast",   "6-month rolling revenue forecast"),
    "📅  Weekly Actual": ("📅 Weekly Actual", "Weekly actual revenue vs. monthly plan — balance tracking"),
    "🔧  Master":        ("🔧 Master Data",   "Account, Price and Tester master records"),
    "📉  Summary":       ("📉 Summary",       "Consolidated KPI dashboard across all modules"),
}

title, subtitle = PAGE_META[menu]
st.markdown(
    f'<div class="header-banner">'
    f'<p class="header-title">{title}</p>'
    f'<p class="header-subtitle">{subtitle}</p>'
    f'</div>',
    unsafe_allow_html=True,
)

# ══════════════════════════════════════════════════════════════
# PAGE: Annual Plan
# ══════════════════════════════════════════════════════════════
if menu == "📈  Annual Plan":
    upload_tab, view_tab, chart_tab = st.tabs(["📤 Upload", "📋 Data View", "📊 Charts"])

    with upload_tab:
        st.markdown('<div class="content-card">', unsafe_allow_html=True)
        section("Upload Annual Plan File")
        st.markdown(
            '<div class="upload-zone">'
            '<p style="font-size:32px;margin:0;">📁</p>'
            '<p style="color:#1d4ed8;font-weight:600;">Drop your Excel file here</p>'
            '<p style="color:#64748b;font-size:13px;">Supports .xlsx, .xls, .csv</p>'
            '</div>',
            unsafe_allow_html=True,
        )
        uploaded = st.file_uploader(
            "Choose Annual Plan file",
            type=["xlsx", "xls", "csv"],
            key="annual_upload",
            label_visibility="collapsed",
        )
        if uploaded:
            st.session_state["annual_df"] = load_uploaded(uploaded)
            st.success(f"✅ Loaded **{uploaded.name}** — {len(st.session_state['annual_df'])} rows")
        st.markdown("</div>", unsafe_allow_html=True)

        # Also show template download
        tpl_path = TEMPLATES_DIR / "Annual_Plan_Template.xlsx"
        if tpl_path.exists():
            with open(tpl_path, "rb") as f:
                st.download_button(
                    "⬇️  Download Template",
                    data=f,
                    file_name="Annual_Plan_Template.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

    # Resolve dataframe
    if "annual_df" not in st.session_state:
        tpl = TEMPLATES_DIR / "Annual_Plan_Template.xlsx"
        st.session_state["annual_df"] = load_excel(tpl) if tpl.exists() else pd.DataFrame()

    df_annual = st.session_state["annual_df"]

    with view_tab:
        st.markdown('<div class="content-card">', unsafe_allow_html=True)
        section("Annual Plan Data")
        if df_annual.empty:
            empty_state("No data loaded. Upload a file on the Upload tab.")
        else:
            # Detect month target columns (must contain a month abbr AND a value keyword)
            VALUE_KEYWORDS = ("_Qty", "_Sales", "_Revenue", "_Target", "Qty", "Sales", "Revenue", "Target")
            month_cols = [
                c for c in df_annual.columns
                if any(m in str(c) for m in MONTH_ABBR) and any(k in str(c) for k in VALUE_KEYWORDS)
            ]
            info_cols = [c for c in df_annual.columns if c not in month_cols]

            # KPIs
            numeric_cols = df_annual.select_dtypes(include="number").columns.tolist()
            total_rev = df_annual[numeric_cols].sum().sum() if numeric_cols else 0
            num_accounts = df_annual["Account_Name"].nunique() if "Account_Name" in df_annual.columns else len(df_annual)

            c1, c2, c3, c4 = st.columns(4)
            c1.markdown(kpi_card("Total Annual Target", fmt_currency(total_rev), color="blue"), unsafe_allow_html=True)
            c2.markdown(kpi_card("Accounts", str(num_accounts), color="green"), unsafe_allow_html=True)
            if "Account_Manager" in df_annual.columns:
                managers = df_annual["Account_Manager"].nunique()
                c3.markdown(kpi_card("Account Managers", str(managers), color="orange"), unsafe_allow_html=True)
            c4.markdown(kpi_card("Data Rows", str(len(df_annual)), color="purple"), unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)

            # Filter
            if "Account_Manager" in df_annual.columns:
                mgr_list = ["All"] + sorted(df_annual["Account_Manager"].dropna().unique().tolist())
                mgr = st.selectbox("Filter by Account Manager", mgr_list)
                view_df = df_annual if mgr == "All" else df_annual[df_annual["Account_Manager"] == mgr]
            else:
                view_df = df_annual

            st.dataframe(view_df, use_container_width=True, height=380)

            # Export
            st.download_button(
                "⬇️  Export to Excel",
                data=to_excel_bytes(view_df),
                file_name="Annual_Plan_Export.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        st.markdown("</div>", unsafe_allow_html=True)

    with chart_tab:
        st.markdown('<div class="content-card">', unsafe_allow_html=True)
        section("Monthly Revenue Overview")
        if df_annual.empty:
            empty_state("No data loaded. Upload a file on the Upload tab.")
        else:
            # Build monthly totals from columns containing month names
            month_data = {}
            for m in MONTH_ABBR:
                cols = [c for c in df_annual.columns if m in str(c) and ("Target" in str(c) or "Sales" in str(c) or "Revenue" in str(c))]
                if cols:
                    month_data[m] = df_annual[cols].sum().sum()
            # Fallback: use all numeric columns in order
            if not month_data:
                numeric_cols = df_annual.select_dtypes(include="number").columns.tolist()
                for i, col in enumerate(numeric_cols[:12]):
                    month_data[MONTH_ABBR[i]] = df_annual[col].sum()

            if month_data:
                months = list(month_data.keys())
                values = list(month_data.values())

                col1, col2 = st.columns(2)
                with col1:
                    fig_bar = px.bar(
                        x=months, y=values,
                        labels={"x": "Month", "y": "Revenue"},
                        color=values,
                        color_continuous_scale="Blues",
                        title="Monthly Revenue Target (Bar)",
                    )
                    fig_bar.update_layout(
                        plot_bgcolor="white", paper_bgcolor="white",
                        showlegend=False, coloraxis_showscale=False,
                        margin=dict(l=0, r=0, t=40, b=0),
                    )
                    st.plotly_chart(fig_bar, use_container_width=True)

                with col2:
                    fig_line = go.Figure()
                    fig_line.add_trace(go.Scatter(
                        x=months, y=values,
                        mode="lines+markers",
                        line=dict(color="#2563eb", width=3),
                        marker=dict(size=8, color="#2563eb"),
                        fill="tozeroy",
                        fillcolor="rgba(37,99,235,0.1)",
                    ))
                    fig_line.update_layout(
                        title="Monthly Revenue Trend (Line)",
                        plot_bgcolor="white", paper_bgcolor="white",
                        xaxis_title="Month", yaxis_title="Revenue",
                        margin=dict(l=0, r=0, t=40, b=0),
                    )
                    st.plotly_chart(fig_line, use_container_width=True)

                # By account pie (if account col exists)
                if "Account_Name" in df_annual.columns:
                    numeric_cols = df_annual.select_dtypes(include="number").columns.tolist()
                    if numeric_cols:
                        by_acct = df_annual.groupby("Account_Name")[numeric_cols].sum().sum(axis=1).nlargest(10)
                        fig_pie = px.pie(
                            values=by_acct.values,
                            names=by_acct.index,
                            title="Revenue Share by Account (Top 10)",
                            color_discrete_sequence=px.colors.sequential.Blues_r,
                        )
                        fig_pie.update_layout(paper_bgcolor="white", margin=dict(l=0, r=0, t=40, b=0))
                        st.plotly_chart(fig_pie, use_container_width=True)
            else:
                empty_state("No monthly data columns found in the uploaded file.")
        st.markdown("</div>", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════
# PAGE: 6M Forecast
# ══════════════════════════════════════════════════════════════
elif menu == "📊  6M Forecast":
    upload_tab, view_tab, chart_tab = st.tabs(["📤 Upload", "📋 Data View", "📊 Comparison"])

    with upload_tab:
        st.markdown('<div class="content-card">', unsafe_allow_html=True)
        section("Upload 6M Forecast File")
        uploaded = st.file_uploader(
            "Choose 6M Forecast file",
            type=["xlsx", "xls", "csv"],
            key="fcst_upload",
            label_visibility="collapsed",
        )
        if uploaded:
            st.session_state["fcst_df"] = load_uploaded(uploaded)
            st.success(f"✅ Loaded **{uploaded.name}** — {len(st.session_state['fcst_df'])} rows")

        tpl_path = TEMPLATES_DIR / "6M_Forecast_Template.xlsx"
        if tpl_path.exists():
            with open(tpl_path, "rb") as f:
                st.download_button(
                    "⬇️  Download Template",
                    data=f,
                    file_name="6M_Forecast_Template.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        st.markdown("</div>", unsafe_allow_html=True)

    if "fcst_df" not in st.session_state:
        tpl = TEMPLATES_DIR / "6M_Forecast_Template.xlsx"
        st.session_state["fcst_df"] = load_excel(tpl) if tpl.exists() else pd.DataFrame()

    df_fcst = st.session_state["fcst_df"]

    with view_tab:
        st.markdown('<div class="content-card">', unsafe_allow_html=True)
        section("6M Forecast Data")
        if df_fcst.empty:
            empty_state("No data loaded. Upload a file on the Upload tab.")
        else:
            numeric_cols = df_fcst.select_dtypes(include="number").columns.tolist()
            total_fcst = df_fcst[numeric_cols].sum().sum() if numeric_cols else 0
            num_acc = df_fcst["Account_Name"].nunique() if "Account_Name" in df_fcst.columns else len(df_fcst)

            c1, c2, c3 = st.columns(3)
            c1.markdown(kpi_card("Total 6M Forecast", fmt_currency(total_fcst), color="blue"), unsafe_allow_html=True)
            c2.markdown(kpi_card("Accounts", str(num_acc), color="green"), unsafe_allow_html=True)
            c3.markdown(kpi_card("Forecast Rows", str(len(df_fcst)), color="purple"), unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)
            st.dataframe(df_fcst, use_container_width=True, height=380)

            st.download_button(
                "⬇️  Export to Excel",
                data=to_excel_bytes(df_fcst),
                file_name="6M_Forecast_Export.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        st.markdown("</div>", unsafe_allow_html=True)

    with chart_tab:
        st.markdown('<div class="content-card">', unsafe_allow_html=True)
        section("Forecast vs. Plan Comparison")
        if df_fcst.empty:
            empty_state("No data loaded. Upload a file on the Upload tab.")
        else:
            fcst_cols = [c for c in df_fcst.columns if "Fcst" in str(c) or "Forecast" in str(c) or "Revenue" in str(c) or "Month" in str(c)]
            if not fcst_cols:
                fcst_cols = df_fcst.select_dtypes(include="number").columns.tolist()

            month_totals = {}
            for i, col in enumerate(fcst_cols[:6]):
                label = col if len(col) <= 15 else f"Month {i+1}"
                month_totals[label] = df_fcst[col].sum()

            if month_totals:
                labels = list(month_totals.keys())
                vals = list(month_totals.values())

                fig = make_subplots(rows=1, cols=2, subplot_titles=("Monthly Forecast", "Cumulative Forecast"))

                fig.add_trace(go.Bar(
                    x=labels, y=vals, name="Forecast",
                    marker_color=["#2563eb", "#3b82f6", "#60a5fa", "#93c5fd", "#bfdbfe", "#dbeafe"][:len(vals)],
                ), row=1, col=1)

                cumulative = list(np.cumsum(vals))
                fig.add_trace(go.Scatter(
                    x=labels, y=cumulative, mode="lines+markers",
                    name="Cumulative",
                    line=dict(color="#22c55e", width=3),
                    marker=dict(size=8),
                ), row=1, col=2)

                fig.update_layout(
                    plot_bgcolor="white", paper_bgcolor="white",
                    showlegend=False, height=380,
                    margin=dict(l=0, r=0, t=50, b=0),
                )
                st.plotly_chart(fig, use_container_width=True)

                # Progress bars per month
                st.markdown("#### Month-by-Month Progress")
                max_val = max(vals) if vals else 1
                for lbl, v in zip(labels, vals):
                    pct = int(v / max_val * 100) if max_val else 0
                    st.markdown(
                        f'<div class="progress-label"><span>{lbl}</span><span>{fmt_currency(v)}</span></div>',
                        unsafe_allow_html=True,
                    )
                    st.progress(pct / 100)
        st.markdown("</div>", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════
# PAGE: Weekly Actual
# ══════════════════════════════════════════════════════════════
elif menu == "📅  Weekly Actual":
    upload_tab, view_tab, chart_tab = st.tabs(["📤 Upload", "📋 Data View", "📊 Balance Tracking"])

    with upload_tab:
        st.markdown('<div class="content-card">', unsafe_allow_html=True)
        section("Upload Weekly Actual File")
        uploaded = st.file_uploader(
            "Choose Weekly Actual file",
            type=["xlsx", "xls", "csv"],
            key="weekly_upload",
            label_visibility="collapsed",
        )
        if uploaded:
            st.session_state["weekly_df"] = load_uploaded(uploaded)
            st.success(f"✅ Loaded **{uploaded.name}** — {len(st.session_state['weekly_df'])} rows")

        tpl_path = TEMPLATES_DIR / "Weekly_Actual_Template.xlsx"
        if tpl_path.exists():
            with open(tpl_path, "rb") as f:
                st.download_button(
                    "⬇️  Download Template",
                    data=f,
                    file_name="Weekly_Actual_Template.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        st.markdown("</div>", unsafe_allow_html=True)

    if "weekly_df" not in st.session_state:
        tpl = TEMPLATES_DIR / "Weekly_Actual_Template.xlsx"
        st.session_state["weekly_df"] = load_excel(tpl) if tpl.exists() else pd.DataFrame()

    df_weekly = st.session_state["weekly_df"]

    with view_tab:
        st.markdown('<div class="content-card">', unsafe_allow_html=True)
        section("Weekly Actual Data")
        if df_weekly.empty:
            empty_state("No data loaded. Upload a file on the Upload tab.")
        else:
            plan_cols = [c for c in df_weekly.columns if "Plan" in str(c) and ("Revenue" in str(c) or "Sales" in str(c))]
            actual_cols = [c for c in df_weekly.columns if ("WW" in str(c) or "Actual" in str(c)) and ("Revenue" in str(c) or "Sales" in str(c)) and "Balance" not in str(c)]

            plan_total = df_weekly[plan_cols].sum().sum() if plan_cols else 0
            actual_total = df_weekly[actual_cols].sum().sum() if actual_cols else 0
            balance = plan_total - actual_total
            ach_rate = (actual_total / plan_total * 100) if plan_total else 0

            c1, c2, c3, c4 = st.columns(4)
            c1.markdown(kpi_card("Monthly Plan", fmt_currency(plan_total), color="blue"), unsafe_allow_html=True)
            c2.markdown(kpi_card("Actual Revenue", fmt_currency(actual_total), color="green"), unsafe_allow_html=True)
            c3.markdown(kpi_card("Balance", fmt_currency(balance), color="orange"), unsafe_allow_html=True)
            c4.markdown(kpi_card("Achievement", f"{ach_rate:.1f}%", color="purple"), unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)
            st.dataframe(df_weekly, use_container_width=True, height=380)

            st.download_button(
                "⬇️  Export to Excel",
                data=to_excel_bytes(df_weekly),
                file_name="Weekly_Actual_Export.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        st.markdown("</div>", unsafe_allow_html=True)

    with chart_tab:
        st.markdown('<div class="content-card">', unsafe_allow_html=True)
        section("Weekly Balance Tracking")
        if df_weekly.empty:
            empty_state("No data loaded. Upload a file on the Upload tab.")
        else:
            # Identify weekly actual columns (WW prefix)
            ww_cols = sorted([c for c in df_weekly.columns if str(c).startswith("WW") and "Balance" not in str(c)])
            bal_cols = sorted([c for c in df_weekly.columns if "Balance" in str(c)])

            if ww_cols:
                ww_totals = {c: df_weekly[c].sum() for c in ww_cols}
                wks = list(ww_totals.keys())
                actuals = list(ww_totals.values())
                cum_actual = list(np.cumsum(actuals))

                fig = go.Figure()
                fig.add_trace(go.Bar(x=wks, y=actuals, name="Weekly Revenue", marker_color="#2563eb"))
                fig.add_trace(go.Scatter(
                    x=wks, y=cum_actual, name="Cumulative",
                    mode="lines+markers", line=dict(color="#f97316", width=3),
                    marker=dict(size=10), yaxis="y2",
                ))
                fig.update_layout(
                    title="Weekly Revenue & Cumulative",
                    yaxis=dict(title="Weekly Revenue"),
                    yaxis2=dict(title="Cumulative Revenue", overlaying="y", side="right"),
                    plot_bgcolor="white", paper_bgcolor="white",
                    legend=dict(orientation="h", y=-0.15),
                    margin=dict(l=0, r=0, t=50, b=0),
                    height=380,
                )
                st.plotly_chart(fig, use_container_width=True)
            else:
                # Generic numeric columns
                numeric_cols = df_weekly.select_dtypes(include="number").columns.tolist()
                if numeric_cols:
                    col_totals = {c: df_weekly[c].sum() for c in numeric_cols[:8]}
                    fig = px.bar(
                        x=list(col_totals.keys()),
                        y=list(col_totals.values()),
                        labels={"x": "Column", "y": "Total"},
                        title="Column Totals",
                        color_discrete_sequence=["#2563eb"],
                    )
                    fig.update_layout(plot_bgcolor="white", paper_bgcolor="white", margin=dict(l=0, r=0, t=50, b=0))
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    empty_state("No weekly data columns (WW**) found in the uploaded file.")
        st.markdown("</div>", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════
# PAGE: Master
# ══════════════════════════════════════════════════════════════
elif menu == "🔧  Master":
    acc_tab, price_tab, tester_tab = st.tabs(["👥 Account Master", "💰 Price Master", "🔬 Tester Master"])

    # ── Account Master ──
    with acc_tab:
        st.markdown('<div class="content-card">', unsafe_allow_html=True)
        section("Account Master Data")

        acc_upload = st.file_uploader(
            "Upload Account Master (xlsx/csv)",
            type=["xlsx", "xls", "csv"],
            key="acc_upload",
        )
        if acc_upload:
            st.session_state["acc_df"] = load_uploaded(acc_upload)
            st.success(f"✅ Loaded {len(st.session_state['acc_df'])} records")

        if "acc_df" not in st.session_state:
            tpl = TEMPLATES_DIR / "Account_Master_Template.xlsx"
            st.session_state["acc_df"] = load_excel(tpl) if tpl.exists() else pd.DataFrame()

        df_acc = st.session_state["acc_df"]
        if df_acc.empty:
            empty_state("No Account Master data. Upload a file above.")
        else:
            c1, c2 = st.columns(2)
            c1.markdown(kpi_card("Total Accounts", str(len(df_acc)), color="blue"), unsafe_allow_html=True)
            if "Account_Manager" in df_acc.columns:
                c2.markdown(kpi_card("Account Managers", str(df_acc["Account_Manager"].nunique()), color="green"), unsafe_allow_html=True)
            st.markdown("<br>", unsafe_allow_html=True)

            search = st.text_input("🔍 Search accounts", placeholder="Type account name or manager…")
            if search:
                text_cols = [c for c in df_acc.columns if df_acc[c].dtype == object]
                if text_cols:
                    mask = df_acc[text_cols].apply(
                        lambda col: col.astype(str).str.contains(search, case=False, na=False)
                    ).any(axis=1)
                    df_acc = df_acc[mask]

            st.dataframe(df_acc, use_container_width=True, height=380)
            st.download_button(
                "⬇️  Export to Excel",
                data=to_excel_bytes(df_acc),
                file_name="Account_Master_Export.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            # Bar chart by manager
            if "Account_Manager" in df_acc.columns:
                by_mgr = df_acc.groupby("Account_Manager").size().reset_index(name="Accounts")
                fig = px.bar(
                    by_mgr, x="Account_Manager", y="Accounts",
                    title="Accounts per Manager",
                    color="Accounts",
                    color_continuous_scale="Blues",
                )
                fig.update_layout(
                    plot_bgcolor="white", paper_bgcolor="white",
                    showlegend=False, coloraxis_showscale=False,
                    margin=dict(l=0, r=0, t=50, b=0),
                )
                st.plotly_chart(fig, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    # ── Price Master ──
    with price_tab:
        st.markdown('<div class="content-card">', unsafe_allow_html=True)
        section("Price Master Data")

        price_upload = st.file_uploader(
            "Upload Price Master (xlsx/csv)",
            type=["xlsx", "xls", "csv"],
            key="price_upload",
        )
        if price_upload:
            st.session_state["price_df"] = load_uploaded(price_upload)
            st.success(f"✅ Loaded {len(st.session_state['price_df'])} records")

        if "price_df" not in st.session_state:
            st.session_state["price_df"] = pd.DataFrame()

        df_price = st.session_state["price_df"]
        if df_price.empty:
            st.info("💡 Price Master columns: Device | PKG_Type | PKG_Body | PKG_Pin | Price")
            empty_state("No Price Master data. Upload a file above.")
        else:
            c1, c2 = st.columns(2)
            c1.markdown(kpi_card("Price Records", str(len(df_price)), color="blue"), unsafe_allow_html=True)
            if "Price" in df_price.columns:
                avg_price = df_price["Price"].mean()
                c2.markdown(kpi_card("Avg. Price", fmt_currency(avg_price), color="green"), unsafe_allow_html=True)
            st.markdown("<br>", unsafe_allow_html=True)
            st.dataframe(df_price, use_container_width=True, height=380)
            st.download_button(
                "⬇️  Export to Excel",
                data=to_excel_bytes(df_price),
                file_name="Price_Master_Export.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        st.markdown("</div>", unsafe_allow_html=True)

    # ── Tester Master ──
    with tester_tab:
        st.markdown('<div class="content-card">', unsafe_allow_html=True)
        section("Tester Master Data")

        tester_upload = st.file_uploader(
            "Upload Tester Master (xlsx/csv)",
            type=["xlsx", "xls", "csv"],
            key="tester_upload",
        )
        if tester_upload:
            st.session_state["tester_df"] = load_uploaded(tester_upload)
            st.success(f"✅ Loaded {len(st.session_state['tester_df'])} records")

        if "tester_df" not in st.session_state:
            st.session_state["tester_df"] = pd.DataFrame()

        df_tester = st.session_state["tester_df"]
        if df_tester.empty:
            st.info("💡 Tester Master columns: Tester_Name | Division | Device | Specification")
            empty_state("No Tester Master data. Upload a file above.")
        else:
            st.markdown(kpi_card("Tester Records", str(len(df_tester)), color="blue"), unsafe_allow_html=True)
            st.markdown("<br>", unsafe_allow_html=True)
            st.dataframe(df_tester, use_container_width=True, height=380)
            st.download_button(
                "⬇️  Export to Excel",
                data=to_excel_bytes(df_tester),
                file_name="Tester_Master_Export.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        st.markdown("</div>", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════
# PAGE: Summary
# ══════════════════════════════════════════════════════════════
elif menu == "📉  Summary":
    st.markdown('<div class="content-card">', unsafe_allow_html=True)
    section("📊 Key Performance Indicators")

    # Gather data from session state
    df_annual = st.session_state.get("annual_df", pd.DataFrame())
    df_fcst = st.session_state.get("fcst_df", pd.DataFrame())
    df_weekly = st.session_state.get("weekly_df", pd.DataFrame())
    df_acc = st.session_state.get("acc_df", pd.DataFrame())

    # If all empty, load templates
    if df_annual.empty:
        tpl = TEMPLATES_DIR / "Annual_Plan_Template.xlsx"
        df_annual = load_excel(tpl) if tpl.exists() else pd.DataFrame()
    if df_acc.empty:
        tpl = TEMPLATES_DIR / "Account_Master_Template.xlsx"
        df_acc = load_excel(tpl) if tpl.exists() else pd.DataFrame()

    # KPIs
    annual_total = df_annual.select_dtypes(include="number").sum().sum() if not df_annual.empty else 0
    fcst_total = df_fcst.select_dtypes(include="number").sum().sum() if not df_fcst.empty else 0
    num_accounts = df_acc["Account_Name"].nunique() if "Account_Name" in df_acc.columns else 0
    weekly_plan_cols = [c for c in df_weekly.columns if "Plan" in str(c) and ("Revenue" in str(c) or "Sales" in str(c))] if not df_weekly.empty else []
    weekly_plan = df_weekly[weekly_plan_cols].sum().sum() if weekly_plan_cols else 0

    k1, k2, k3, k4 = st.columns(4)
    k1.markdown(kpi_card("Annual Plan Total", fmt_currency(annual_total), color="blue"), unsafe_allow_html=True)
    k2.markdown(kpi_card("6M Forecast Total", fmt_currency(fcst_total), color="green"), unsafe_allow_html=True)
    k3.markdown(kpi_card("Active Accounts", str(num_accounts), color="orange"), unsafe_allow_html=True)
    k4.markdown(kpi_card("Monthly Plan (WA)", fmt_currency(weekly_plan), color="purple"), unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)

    # Charts row
    col1, col2 = st.columns(2)

    with col1:
        st.markdown('<div class="content-card">', unsafe_allow_html=True)
        section("Annual Plan — Monthly Breakdown")
        if not df_annual.empty:
            month_data = {}
            for m in MONTH_ABBR:
                cols = [c for c in df_annual.columns if m in str(c)]
                if cols:
                    month_data[m] = df_annual[cols].select_dtypes(include="number").sum().sum()
            if not month_data:
                numeric_cols = df_annual.select_dtypes(include="number").columns.tolist()
                for i, col in enumerate(numeric_cols[:12]):
                    month_data[MONTH_ABBR[i]] = df_annual[col].sum()

            if month_data:
                fig = px.bar(
                    x=list(month_data.keys()),
                    y=list(month_data.values()),
                    labels={"x": "Month", "y": "Revenue"},
                    color=list(month_data.values()),
                    color_continuous_scale="Blues",
                )
                fig.update_layout(
                    plot_bgcolor="white", paper_bgcolor="white",
                    showlegend=False, coloraxis_showscale=False,
                    margin=dict(l=0, r=0, t=10, b=0), height=300,
                )
                st.plotly_chart(fig, use_container_width=True)
            else:
                empty_state("No annual data available.")
        else:
            empty_state("Load Annual Plan data first.")
        st.markdown("</div>", unsafe_allow_html=True)

    with col2:
        st.markdown('<div class="content-card">', unsafe_allow_html=True)
        section("6M Forecast Trend")
        if not df_fcst.empty:
            fcst_cols = [c for c in df_fcst.columns if "Fcst" in str(c) or "Month" in str(c) or "Revenue" in str(c)]
            if not fcst_cols:
                fcst_cols = df_fcst.select_dtypes(include="number").columns.tolist()
            fcst_totals = {f"M{i+1}": df_fcst[c].sum() for i, c in enumerate(fcst_cols[:6])}
            if fcst_totals:
                fig = go.Figure(go.Scatter(
                    x=list(fcst_totals.keys()),
                    y=list(fcst_totals.values()),
                    mode="lines+markers",
                    line=dict(color="#22c55e", width=3),
                    marker=dict(size=10, color="#22c55e"),
                    fill="tozeroy", fillcolor="rgba(34,197,94,0.1)",
                ))
                fig.update_layout(
                    plot_bgcolor="white", paper_bgcolor="white",
                    xaxis_title="Month", yaxis_title="Forecast Revenue",
                    margin=dict(l=0, r=0, t=10, b=0), height=300,
                )
                st.plotly_chart(fig, use_container_width=True)
            else:
                empty_state("No 6M forecast data available.")
        else:
            empty_state("Load 6M Forecast data first.")
        st.markdown("</div>", unsafe_allow_html=True)

    # Account distribution
    if not df_acc.empty and "Account_Manager" in df_acc.columns:
        st.markdown('<div class="content-card">', unsafe_allow_html=True)
        section("Account Distribution by Manager")
        by_mgr = df_acc.groupby("Account_Manager").size().reset_index(name="Accounts")
        fig = px.pie(
            by_mgr, names="Account_Manager", values="Accounts",
            color_discrete_sequence=px.colors.sequential.Blues_r,
        )
        fig.update_layout(paper_bgcolor="white", margin=dict(l=0, r=0, t=10, b=0), height=350)
        st.plotly_chart(fig, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    # Data status table
    st.markdown('<div class="content-card">', unsafe_allow_html=True)
    section("📁 Data Status")
    status_rows = [
        {
            "Module": "📈 Annual Plan",
            "Status": "✅ Loaded" if not df_annual.empty else "⚠️ Not loaded",
            "Rows": len(df_annual) if not df_annual.empty else 0,
            "Columns": len(df_annual.columns) if not df_annual.empty else 0,
        },
        {
            "Module": "📊 6M Forecast",
            "Status": "✅ Loaded" if not df_fcst.empty else "⚠️ Not loaded",
            "Rows": len(df_fcst) if not df_fcst.empty else 0,
            "Columns": len(df_fcst.columns) if not df_fcst.empty else 0,
        },
        {
            "Module": "📅 Weekly Actual",
            "Status": "✅ Loaded" if not df_weekly.empty else "⚠️ Not loaded",
            "Rows": len(df_weekly) if not df_weekly.empty else 0,
            "Columns": len(df_weekly.columns) if not df_weekly.empty else 0,
        },
        {
            "Module": "👥 Account Master",
            "Status": "✅ Loaded" if not df_acc.empty else "⚠️ Not loaded",
            "Rows": len(df_acc) if not df_acc.empty else 0,
            "Columns": len(df_acc.columns) if not df_acc.empty else 0,
        },
    ]
    st.dataframe(pd.DataFrame(status_rows), use_container_width=True, hide_index=True)
    st.markdown("</div>", unsafe_allow_html=True)
