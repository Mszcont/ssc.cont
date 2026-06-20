# -*- coding: utf-8 -*-
"""
SSC Project Management System — Enhanced Edition
==================================================
A full-featured build adding: login & role-based permissions, edit/delete for
all records, dashboard alerts, Excel/PDF report export, database backup &
restore, and a polished interface.
"""

import streamlit as st
import sqlite3
import pandas as pd
import plotly.express as px
import hashlib
import io
import os
import urllib.request
from datetime import datetime, date

# =====================================================================================
# General settings
# =====================================================================================
DB_PATH = "ssc_projects.db"
APP_TITLE = "SSC Project Management System"
CURRENCY = "SAR"
# An Arabic-capable font is used as a fallback so that any Arabic text the user
# enters (project/client names, notes, etc.) still renders correctly inside PDF
# reports. No manual setup is normally needed — the app tries to download it
# automatically on first PDF export if an internet connection is available.
ARABIC_FONT_PATH = "Amiri-Regular.ttf"
ARABIC_FONT_DOWNLOAD_URLS = [
    "https://raw.githubusercontent.com/google/fonts/main/ofl/amiri/Amiri-Regular.ttf",
    "https://cdn.jsdelivr.net/gh/notofonts/notofonts.github.io/fonts/NotoNaskhArabic/unhinted/ttf/NotoNaskhArabic-Regular.ttf",
]

st.set_page_config(page_title=APP_TITLE, page_icon="🏗️", layout="wide")

# --- Global UI styling (CSS) ---
st.markdown("""
<style>
    html, body, [class*="css"]  { font-family: 'Segoe UI', Tahoma, sans-serif; }
    .stApp { direction: ltr; }
    div[data-testid="stMetric"] {
        background-color: #ffffff;
        border: 1px solid #e6e6e6;
        border-radius: 10px;
        padding: 14px 16px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    }
    div[data-testid="stMetricLabel"] { direction: ltr; text-align: left; }
    section[data-testid="stSidebar"] {
        background-color: #0f2540;
    }
    section[data-testid="stSidebar"] * { color: #f5f5f5 !important; }
    .ssc-alert-box {
        background-color: #fff4e5;
        border-left: 5px solid #e67e22;
        padding: 10px 14px;
        border-radius: 6px;
        margin-bottom: 8px;
    }
    .ssc-alert-danger {
        background-color: #fdecea;
        border-left: 5px solid #c0392b;
        padding: 10px 14px;
        border-radius: 6px;
        margin-bottom: 8px;
    }
    .ssc-badge {
        display: inline-block;
        padding: 2px 10px;
        border-radius: 12px;
        font-size: 0.8em;
        color: white;
    }
</style>
""", unsafe_allow_html=True)


# =====================================================================================
# Font size / color settings (customizable from the sidebar)
# =====================================================================================
if "font_size" not in st.session_state:
    st.session_state.font_size = 16          # px
if "font_color" not in st.session_state:
    st.session_state.font_color = "#1a1a1a"  # default text color (main content area only)


def apply_dynamic_style():
    """Injects CSS with the user-selected font size/color, then re-asserts the
    sidebar's light text color last so it isn't affected by the general
    font-color choice (the sidebar background stays dark).

    IMPORTANT: call this exactly once per script run (either from login_page()
    while logged out, or once after the sidebar settings widgets while logged
    in) — calling it twice in the same run can make the two injected <style>
    tags fight over which one "wins" the cascade, which is what made earlier
    changes look like they weren't being applied."""
    size = st.session_state.font_size
    color = st.session_state.font_color
    st.markdown(f"""
    <style>
        html, body, [class*="css"], [data-testid="stAppViewContainer"],
        .stMarkdown, .stMarkdown p, .stText,
        p, span, label, li,
        div[data-testid="stMetricLabel"], div[data-testid="stMetricValue"],
        .stDataFrame, .stSelectbox label, .stTextInput label, .stNumberInput label,
        .stDateInput label, .stRadio label, .stCheckbox label, .stTextArea label,
        h1, h2, h3, h4, h5, h6 {{
            font-size: {size}px !important;
            color: {color} !important;
        }}
        h1 {{ font-size: {size + 14}px !important; }}
        h2 {{ font-size: {size + 8}px !important; }}
        h3 {{ font-size: {size + 4}px !important; }}
        h4, h5 {{ font-size: {size + 2}px !important; }}
        /* Re-assert the sidebar text color after the general override above */
        section[data-testid="stSidebar"] * {{
            color: #f5f5f5 !important;
        }}
    </style>
    """, unsafe_allow_html=True)


# =====================================================================================
# Database layer
# =====================================================================================
def get_db_connection():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def _safe_add_column(cursor, table, coldef):
    """Adds a new column to an existing table if it doesn't already exist
    (keeps compatibility with databases created by earlier versions)."""
    try:
        cursor.execute(f"ALTER TABLE {table} ADD COLUMN {coldef}")
    except sqlite3.OperationalError:
        pass  # column already exists


def hash_password(password: str, salt: str = "ssc_static_salt_v1") -> str:
    return hashlib.sha256((salt + password).encode("utf-8")).hexdigest()


def init_db():
    conn = get_db_connection()
    cur = conn.cursor()

    # Projects table
    cur.execute("""
        CREATE TABLE IF NOT EXISTS projects (
            project_id TEXT PRIMARY KEY,
            project_name TEXT NOT NULL,
            start_date TEXT,
            end_date TEXT,
            contract_value REAL
        )""")
    # Upgrade older tables with new columns
    _safe_add_column(cur, "projects", "client_name TEXT")
    _safe_add_column(cur, "projects", "status TEXT DEFAULT 'In Progress'")
    _safe_add_column(cur, "projects", "created_at TEXT")

    # Revenues table
    cur.execute("""
        CREATE TABLE IF NOT EXISTS revenues (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id TEXT,
            date TEXT,
            amount REAL,
            stage TEXT,
            FOREIGN KEY(project_id) REFERENCES projects(project_id)
        )""")
    _safe_add_column(cur, "revenues", "notes TEXT")

    # Expenses table
    cur.execute("""
        CREATE TABLE IF NOT EXISTS expenses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id TEXT,
            date TEXT,
            expense_type TEXT,
            amount REAL,
            notes TEXT,
            FOREIGN KEY(project_id) REFERENCES projects(project_id)
        )""")

    # Labor table
    cur.execute("""
        CREATE TABLE IF NOT EXISTS labor (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id TEXT,
            days REAL,
            wage REAL,
            total REAL,
            FOREIGN KEY(project_id) REFERENCES projects(project_id)
        )""")
    _safe_add_column(cur, "labor", "worker_name TEXT")
    _safe_add_column(cur, "labor", "date TEXT")

    # Users table
    cur.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            full_name TEXT,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'User',
            created_at TEXT
        )""")
    # Per-user display preferences (so font size/color persist across logins/devices)
    _safe_add_column(cur, "users", "font_size INTEGER DEFAULT 16")
    _safe_add_column(cur, "users", "font_color TEXT DEFAULT '#1a1a1a'")

    conn.commit()

    # Create the default admin account if no users exist yet
    cur.execute("SELECT COUNT(*) FROM users")
    if cur.fetchone()[0] == 0:
        cur.execute(
            "INSERT INTO users (username, full_name, password_hash, role, created_at) VALUES (?,?,?,?,?)",
            ("admin", "System Administrator", hash_password("admin123"), "Admin", str(datetime.now()))
        )
        conn.commit()

    conn.close()


init_db()


# =====================================================================================
# Authentication layer (login)
# =====================================================================================
def verify_user(username: str, password: str):
    conn = get_db_connection()
    row = conn.execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()
    conn.close()
    if row and row["password_hash"] == hash_password(password):
        return dict(row)
    return None


def login_page():
    apply_dynamic_style()
    st.title("🏗️ " + APP_TITLE)
    st.subheader("Login")
    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Sign In", use_container_width=True)
        if submitted:
            user = verify_user(username.strip(), password)
            if user:
                st.session_state.auth_user = user
                # Restore this user's saved display preferences (font size/color)
                st.session_state.font_size = user.get("font_size") or 16
                st.session_state.font_color = user.get("font_color") or "#1a1a1a"
                st.rerun()
            else:
                st.error("Incorrect username or password.")
    st.caption("🔑 Default account on first run: **admin** / **admin123** — please change it immediately from the (User Management) page.")


def require_login():
    if "auth_user" not in st.session_state:
        login_page()
        st.stop()


def current_user():
    return st.session_state.get("auth_user")


def is_admin():
    user = current_user()
    return bool(user and user.get("role") == "Admin")


# =====================================================================================
# Data-fetching helper functions
# =====================================================================================
def fetch_df(query, params=()):
    conn = get_db_connection()
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df


def get_projects_df():
    return fetch_df("SELECT * FROM projects ORDER BY created_at DESC")


def get_project_options():
    """Returns a {project_name: project_id} dict for use in dropdown menus."""
    df = get_projects_df()
    return {row["project_name"]: row["project_id"] for _, row in df.iterrows()}


def build_summary():
    """Builds a comprehensive financial summary table per project (revenue,
    expenses, labor, net profit, collection rate)."""
    df_p = get_projects_df()
    df_r = fetch_df("SELECT project_id, SUM(amount) as total_rev FROM revenues GROUP BY project_id")
    df_e = fetch_df("SELECT project_id, SUM(amount) as total_exp FROM expenses GROUP BY project_id")
    df_l = fetch_df("SELECT project_id, SUM(total) as total_labor FROM labor GROUP BY project_id")

    if df_p.empty:
        return df_p

    df = df_p.merge(df_r, on="project_id", how="left") \
             .merge(df_e, on="project_id", how="left") \
             .merge(df_l, on="project_id", how="left")
    for col in ["total_rev", "total_exp", "total_labor", "contract_value"]:
        if col not in df.columns:
            df[col] = 0.0
    df.fillna(0, inplace=True)

    df["total_costs"] = df["total_exp"] + df["total_labor"]
    df["net_profit"] = df["total_rev"] - df["total_costs"]
    df["collection_rate"] = df.apply(
        lambda r: round((r["total_rev"] / r["contract_value"]) * 100, 1) if r["contract_value"] else 0, axis=1
    )
    df["profit_margin"] = df.apply(
        lambda r: round((r["net_profit"] / r["total_rev"]) * 100, 1) if r["total_rev"] else 0, axis=1
    )
    return df


def build_alerts(df_summary):
    """Scans project data and generates alerts (budget overrun, approaching
    deadline, low collection rate on completed projects)."""
    alerts = []
    if df_summary.empty:
        return alerts

    today = date.today()
    for _, row in df_summary.iterrows():
        name = row["project_name"]
        status = row.get("status", "In Progress") or "In Progress"

        # Budget overrun
        if row["total_costs"] > row["contract_value"] and row["contract_value"] > 0:
            over = row["total_costs"] - row["contract_value"]
            alerts.append(("danger", f"⚠️ Project «{name}»: costs exceeded the contract value by {over:,.0f} {CURRENCY}"))

        # Approaching the expected end date (within 30 days) and still in progress
        if status == "In Progress" and row.get("end_date"):
            try:
                end_dt = datetime.strptime(str(row["end_date"]), "%Y-%m-%d").date()
                days_left = (end_dt - today).days
                if 0 <= days_left <= 30:
                    alerts.append(("warning", f"⏰ Project «{name}»: {days_left} day(s) remaining until the expected end date"))
                elif days_left < 0:
                    alerts.append(("danger", f"⌛ Project «{name}»: the expected end date has passed by {abs(days_left)} day(s) and it is still «In Progress»"))
            except ValueError:
                pass

        # Low collection rate on a completed project
        if status == "Completed" and row["collection_rate"] < 90:
            alerts.append(("warning", f"💰 Project «{name}» is completed, but the collection rate is only {row['collection_rate']}% of the contract value"))

    return alerts


# =====================================================================================
# Export: Excel and PDF
# =====================================================================================
def export_to_excel(df_summary, df_revenues, df_expenses, df_labor):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_summary.to_excel(writer, sheet_name="Project Summary", index=False)
        df_revenues.to_excel(writer, sheet_name="Revenues", index=False)
        df_expenses.to_excel(writer, sheet_name="Expenses", index=False)
        df_labor.to_excel(writer, sheet_name="Labor", index=False)
    buffer.seek(0)
    return buffer


def _ensure_arabic_font_file():
    """Makes sure an Arabic-capable font file exists locally. If it doesn't,
    tries to download one automatically (requires an internet connection in
    the runtime environment, e.g. the deployment server). Caches a failure
    in the current session to avoid retrying on every export."""
    if os.path.exists(ARABIC_FONT_PATH) and os.path.getsize(ARABIC_FONT_PATH) > 10_000:
        return True
    if st.session_state.get("_arabic_font_dl_failed"):
        return False

    for url in ARABIC_FONT_DOWNLOAD_URLS:
        try:
            tmp_path = ARABIC_FONT_PATH + ".part"
            urllib.request.urlretrieve(url, tmp_path)
            if os.path.exists(tmp_path) and os.path.getsize(tmp_path) > 10_000:
                os.replace(tmp_path, ARABIC_FONT_PATH)
                return True
        except Exception:
            continue

    st.session_state["_arabic_font_dl_failed"] = True
    return False


def _get_pdf_font():
    """Tries to register an Arabic-capable font (local or auto-downloaded).
    If that's not possible at all, falls back to Helvetica, which won't
    render Arabic characters correctly."""
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    if _ensure_arabic_font_file():
        try:
            pdfmetrics.registerFont(TTFont("ArabicFont", ARABIC_FONT_PATH))
            return "ArabicFont", True
        except Exception:
            pass
    return "Helvetica", False


def _reshaping_libs_available():
    """Checks whether the arabic_reshaper and python-bidi libraries are
    available — needed to correctly join Arabic letters and reverse the
    display direction inside PDF reports, in case any Arabic text was
    entered into project/client names or notes."""
    try:
        import arabic_reshaper  # noqa: F401
        from bidi.algorithm import get_display  # noqa: F401
        return True
    except ImportError:
        return False


def _ar(text):
    """Reshapes Arabic text for correct display (joining + direction) if the
    required libraries are available; otherwise returns the text unchanged
    (harmless no-op for plain English text)."""
    text = "" if text is None else str(text)
    try:
        import arabic_reshaper
        from bidi.algorithm import get_display
        return get_display(arabic_reshaper.reshape(text))
    except ImportError:
        return text


def export_to_pdf(df_summary, totals):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm

    font_name, font_ok = _get_pdf_font()
    reshape_ok = _reshaping_libs_available()
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, topMargin=1.5 * cm, bottomMargin=1.5 * cm)

    title_style = ParagraphStyle("title", fontName=font_name, fontSize=16, alignment=1, spaceAfter=12)
    normal_style = ParagraphStyle("normal", fontName=font_name, fontSize=9, alignment=1)

    elements = []
    elements.append(Paragraph(_ar(f"Financial Report — {APP_TITLE}"), title_style))
    elements.append(Paragraph(_ar(f"Report date: {date.today().isoformat()}"), normal_style))
    elements.append(Spacer(1, 10))

    # General summary table
    summary_rows = [[_ar("Value"), _ar("Item")]]
    for label, value in totals:
        summary_rows.append([f"{value:,.2f}", _ar(label)])
    t1 = Table(summary_rows, colWidths=[6 * cm, 6 * cm])
    t1.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), font_name),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0f2540")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
    ]))
    elements.append(t1)
    elements.append(Spacer(1, 16))

    # Detailed per-project table
    header = [_ar(h) for h in ["Net Profit", "Total Costs", "Revenue", "Contract Value", "Project Name"]]
    rows = [header]
    for _, r in df_summary.iterrows():
        rows.append([
            f"{r['net_profit']:,.0f}",
            f"{r['total_costs']:,.0f}",
            f"{r['total_rev']:,.0f}",
            f"{r['contract_value']:,.0f}",
            _ar(r["project_name"]),
        ])
    t2 = Table(rows, colWidths=[3.2 * cm, 3.6 * cm, 3.2 * cm, 3.2 * cm, 5 * cm])
    t2.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), font_name),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0f2540")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f5f7fa")]),
    ]))
    elements.append(t2)

    doc.build(elements)
    buffer.seek(0)
    return buffer, font_ok, reshape_ok


# =====================================================================================
# Page 1: Dashboard & Reports
# =====================================================================================
def page_dashboard():
    st.title("📊 Live Financial Dashboard & Analysis")
    st.markdown("---")

    df_summary = build_summary()

    if df_summary.empty:
        st.info("No projects registered yet. Please add projects from the sidebar.")
        return

    # --- Filters ---
    with st.expander("🔎 Filters", expanded=False):
        statuses = sorted(df_summary["status"].fillna("In Progress").unique().tolist())
        chosen_status = st.multiselect("Project status:", statuses, default=statuses)
    df_view = df_summary[df_summary["status"].fillna("In Progress").isin(chosen_status)] if chosen_status else df_summary

    # --- Alerts ---
    st.markdown("#### 🚨 Alerts")
    alerts = build_alerts(df_view)
    if alerts:
        for level, msg in alerts:
            css_class = "ssc-alert-danger" if level == "danger" else "ssc-alert-box"
            st.markdown(f'<div class="{css_class}">{msg}</div>', unsafe_allow_html=True)
    else:
        st.success("✅ No alerts at the moment — all projects are within normal limits.")
    st.markdown("---")

    # --- Key metrics ---
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Contract Value", f"{df_view['contract_value'].sum():,.0f} {CURRENCY}")
    col2.metric("Total Revenue Collected", f"{df_view['total_rev'].sum():,.0f} {CURRENCY}")
    col3.metric("Total Expenses & Labor", f"{df_view['total_costs'].sum():,.0f} {CURRENCY}")
    profit = df_view["net_profit"].sum()
    col4.metric("Total Net Profit", f"{profit:,.0f} {CURRENCY}")

    col5, col6 = st.columns(2)
    total_contract = df_view["contract_value"].sum()
    collection_rate = (df_view["total_rev"].sum() / total_contract * 100) if total_contract else 0
    col5.metric("Overall Collection Rate", f"{collection_rate:,.1f}%")
    total_rev_sum = df_view["total_rev"].sum()
    margin = (profit / total_rev_sum * 100) if total_rev_sum else 0
    col6.metric("Overall Profit Margin", f"{margin:,.1f}%")

    st.markdown("### 📋 Project Performance Summary")
    display_cols = ["project_id", "project_name", "status", "contract_value", "total_rev",
                     "total_costs", "net_profit", "collection_rate", "profit_margin"]
    display_cols = [c for c in display_cols if c in df_view.columns]
    st.dataframe(
        df_view[display_cols].rename(columns={
            "project_id": "Project ID", "project_name": "Project Name", "status": "Status",
            "contract_value": "Contract Value", "total_rev": "Revenue",
            "total_costs": "Total Costs", "net_profit": "Net Profit",
            "collection_rate": "Collection Rate %", "profit_margin": "Profit Margin %",
        }),
        use_container_width=True, hide_index=True
    )

    st.markdown("### 📈 Visual Analytics")
    c1, c2 = st.columns(2)
    with c1:
        fig = px.bar(
            df_view, x="project_name", y=["total_rev", "total_costs", "net_profit"],
            barmode="group", title="Revenue vs. Costs vs. Profit per Project",
            labels={"value": f"Amount ({CURRENCY})", "project_name": "Project", "variable": "Item"}
        )
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        df_exp_type = fetch_df("SELECT expense_type, SUM(amount) as amount FROM expenses GROUP BY expense_type")
        if not df_exp_type.empty:
            fig2 = px.pie(df_exp_type, names="expense_type", values="amount", title="Expense Breakdown by Type")
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.info("No expenses recorded yet to show a breakdown.")


# =====================================================================================
# Page 2: Projects (Add / Edit / Delete)
# =====================================================================================
def page_projects():
    st.title("🏗️ Project Management")
    st.markdown("---")
    tab_add, tab_manage = st.tabs(["➕ Add New Project", "📋 List & Edit"])

    with tab_add:
        with st.form("project_form", clear_on_submit=True):
            p_id = st.text_input("Project ID (unique code):")
            p_name = st.text_input("Project Name:")
            client_name = st.text_input("Client Name:")
            col1, col2 = st.columns(2)
            start_d = col1.date_input("Start Date")
            end_d = col2.date_input("Expected End Date")
            val = st.number_input(f"Total Contract Value ({CURRENCY}):", min_value=0.0, format="%.2f")
            status = st.selectbox("Project Status:", ["In Progress", "On Hold", "Completed"])

            submit = st.form_submit_button("Save Project", use_container_width=True)
            if submit:
                if p_id and p_name:
                    conn = get_db_connection()
                    try:
                        conn.execute(
                            "INSERT INTO projects (project_id, project_name, start_date, end_date, "
                            "contract_value, client_name, status, created_at) VALUES (?,?,?,?,?,?,?,?)",
                            (p_id, p_name, str(start_d), str(end_d), val, client_name, status, str(datetime.now()))
                        )
                        conn.commit()
                        st.success(f"Project ({p_name}) registered successfully!")
                    except sqlite3.IntegrityError:
                        st.error("This Project ID is already registered! Please use a unique ID.")
                    finally:
                        conn.close()
                else:
                    st.warning("Please fill in the required fields (ID and Name).")

    with tab_manage:
        df_p = get_projects_df()
        if df_p.empty:
            st.info("No projects registered yet.")
            return

        st.dataframe(df_p, use_container_width=True, hide_index=True)
        st.markdown("#### ✏️ Edit or Delete a Project")
        options = {f"{row['project_name']} ({row['project_id']})": row["project_id"] for _, row in df_p.iterrows()}
        choice = st.selectbox("Select project:", list(options.keys()), key="proj_edit_choice")
        pid = options[choice]
        record = df_p[df_p["project_id"] == pid].iloc[0]

        with st.form("project_edit_form"):
            new_name = st.text_input("Project Name:", value=record["project_name"])
            new_client = st.text_input("Client Name:", value=record.get("client_name") or "")
            col1, col2 = st.columns(2)
            try:
                sd = datetime.strptime(str(record["start_date"]), "%Y-%m-%d").date()
            except Exception:
                sd = date.today()
            try:
                ed = datetime.strptime(str(record["end_date"]), "%Y-%m-%d").date()
            except Exception:
                ed = date.today()
            new_start = col1.date_input("Start Date:", value=sd, key="edit_start")
            new_end = col2.date_input("End Date:", value=ed, key="edit_end")
            new_val = st.number_input("Contract Value:", min_value=0.0, value=float(record["contract_value"] or 0), format="%.2f")
            status_options = ["In Progress", "On Hold", "Completed"]
            current_status = record.get("status") or "In Progress"
            new_status = st.selectbox("Status:", status_options,
                                       index=status_options.index(current_status) if current_status in status_options else 0)

            colA, colB = st.columns(2)
            save_btn = colA.form_submit_button("💾 Save Changes", use_container_width=True)
            with colB:
                confirm_del = st.checkbox("I confirm I want to delete this project and all its related data")
                del_btn = st.form_submit_button("🗑️ Delete Project Permanently", use_container_width=True, disabled=not confirm_del)

            if save_btn:
                conn = get_db_connection()
                conn.execute(
                    "UPDATE projects SET project_name=?, client_name=?, start_date=?, end_date=?, "
                    "contract_value=?, status=? WHERE project_id=?",
                    (new_name, new_client, str(new_start), str(new_end), new_val, new_status, pid)
                )
                conn.commit()
                conn.close()
                st.success("Changes saved successfully.")
                st.rerun()

            if del_btn and confirm_del:
                conn = get_db_connection()
                conn.execute("DELETE FROM revenues WHERE project_id=?", (pid,))
                conn.execute("DELETE FROM expenses WHERE project_id=?", (pid,))
                conn.execute("DELETE FROM labor WHERE project_id=?", (pid,))
                conn.execute("DELETE FROM projects WHERE project_id=?", (pid,))
                conn.commit()
                conn.close()
                st.success("Project and all related data have been deleted.")
                st.rerun()


# =====================================================================================
# Page 3: Revenues (Add / Edit / Delete)
# =====================================================================================
def page_revenues():
    st.title("💰 Payments & Revenue Management")
    st.markdown("---")

    p_options = get_project_options()
    if not p_options:
        st.warning("Please add a project first so you can record revenue.")
        return

    tab_add, tab_manage = st.tabs(["➕ Record a Payment", "📋 List & Edit"])

    with tab_add:
        with st.form("revenue_form", clear_on_submit=True):
            selected_p = st.selectbox("Select project:", list(p_options.keys()))
            rev_date = st.date_input("Payment Received Date")
            stage = st.selectbox("Payment Type / Stage:", ["Payment 1", "Payment 2", "Payment 3", "Payment 4", "Final Payment"])
            amount = st.number_input(f"Payment Amount ({CURRENCY}):", min_value=0.0, format="%.2f")
            notes = st.text_area("Notes:")

            submit = st.form_submit_button("Record Payment", use_container_width=True)
            if submit:
                conn = get_db_connection()
                conn.execute(
                    "INSERT INTO revenues (project_id, date, amount, stage, notes) VALUES (?,?,?,?,?)",
                    (p_options[selected_p], str(rev_date), amount, stage, notes)
                )
                conn.commit()
                conn.close()
                st.success(f"{stage} of {amount:,.2f} {CURRENCY} recorded for the project successfully!")

    with tab_manage:
        filter_p = st.selectbox("Filter by project:", ["All"] + list(p_options.keys()), key="rev_filter")
        if filter_p == "All":
            df_rev = fetch_df("""
                SELECT r.id, p.project_name, r.date, r.amount, r.stage, r.notes
                FROM revenues r JOIN projects p ON r.project_id = p.project_id
                ORDER BY r.date DESC
            """)
        else:
            df_rev = fetch_df("""
                SELECT r.id, p.project_name, r.date, r.amount, r.stage, r.notes
                FROM revenues r JOIN projects p ON r.project_id = p.project_id
                WHERE r.project_id = ? ORDER BY r.date DESC
            """, params=(p_options[filter_p],))

        if df_rev.empty:
            st.info("No revenue recorded yet.")
            return

        st.dataframe(df_rev, use_container_width=True, hide_index=True)
        st.metric("Total Revenue Shown", f"{df_rev['amount'].sum():,.2f} {CURRENCY}")

        st.markdown("#### ✏️ Edit or Delete a Payment")
        rev_labels = {f"#{row['id']} — {row['project_name']} — {row['date']} — {row['amount']:,.0f} {CURRENCY}": row["id"]
                      for _, row in df_rev.iterrows()}
        chosen = st.selectbox("Select record:", list(rev_labels.keys()), key="rev_edit_choice")
        rid = rev_labels[chosen]
        record = df_rev[df_rev["id"] == rid].iloc[0]

        with st.form("revenue_edit_form"):
            try:
                rdate = datetime.strptime(str(record["date"]), "%Y-%m-%d").date()
            except Exception:
                rdate = date.today()
            new_date = st.date_input("Date:", value=rdate, key="rev_edit_date")
            stage_options = ["Payment 1", "Payment 2", "Payment 3", "Payment 4", "Final Payment"]
            new_stage = st.selectbox("Stage:", stage_options,
                                      index=stage_options.index(record["stage"]) if record["stage"] in stage_options else 0)
            new_amount = st.number_input("Amount:", min_value=0.0, value=float(record["amount"]), format="%.2f")
            new_notes = st.text_area("Notes:", value=record.get("notes") or "")

            colA, colB = st.columns(2)
            save_btn = colA.form_submit_button("💾 Save Changes", use_container_width=True)
            with colB:
                confirm_del = st.checkbox("I confirm I want to delete this record", key="rev_confirm_del")
                del_btn = st.form_submit_button("🗑️ Delete", use_container_width=True, disabled=not confirm_del)

            if save_btn:
                conn = get_db_connection()
                conn.execute("UPDATE revenues SET date=?, amount=?, stage=?, notes=? WHERE id=?",
                             (str(new_date), new_amount, new_stage, new_notes, rid))
                conn.commit()
                conn.close()
                st.success("Changes saved.")
                st.rerun()

            if del_btn and confirm_del:
                conn = get_db_connection()
                conn.execute("DELETE FROM revenues WHERE id=?", (rid,))
                conn.commit()
                conn.close()
                st.success("Record deleted.")
                st.rerun()


# =====================================================================================
# Page 4: Expenses (Add / Edit / Delete)
# =====================================================================================
def page_expenses():
    st.title("📉 Operating Expense Management & Analysis")
    st.markdown("---")

    p_options = get_project_options()
    if not p_options:
        st.warning("Please add a project first so you can record expenses.")
        return

    tab_add, tab_manage = st.tabs(["➕ Record an Expense", "📋 List & Edit"])
    expense_types = ["Materials", "Transport", "Equipment", "Site", "Other"]

    with tab_add:
        with st.form("expense_form", clear_on_submit=True):
            selected_p = st.selectbox("Select the project for this expense:", list(p_options.keys()))
            exp_date = st.date_input("Expense Date")
            exp_type = st.selectbox("Expense Type (for breakdown analysis):", expense_types)
            amount = st.number_input(f"Amount ({CURRENCY}):", min_value=0.0, format="%.2f")
            notes = st.text_area("Additional Notes:")

            submit = st.form_submit_button("Record Expense", use_container_width=True)
            if submit:
                conn = get_db_connection()
                conn.execute(
                    "INSERT INTO expenses (project_id, date, expense_type, amount, notes) VALUES (?,?,?,?,?)",
                    (p_options[selected_p], str(exp_date), exp_type, amount, notes)
                )
                conn.commit()
                conn.close()
                st.success("Expense recorded and the financial analysis updated!")

    with tab_manage:
        filter_p = st.selectbox("Filter by project:", ["All"] + list(p_options.keys()), key="exp_filter")
        if filter_p == "All":
            df_exp = fetch_df("""
                SELECT e.id, p.project_name, e.date, e.expense_type, e.amount, e.notes
                FROM expenses e JOIN projects p ON e.project_id = p.project_id
                ORDER BY e.date DESC
            """)
        else:
            df_exp = fetch_df("""
                SELECT e.id, p.project_name, e.date, e.expense_type, e.amount, e.notes
                FROM expenses e JOIN projects p ON e.project_id = p.project_id
                WHERE e.project_id = ? ORDER BY e.date DESC
            """, params=(p_options[filter_p],))

        if df_exp.empty:
            st.info("No expenses recorded yet.")
            return

        st.dataframe(df_exp, use_container_width=True, hide_index=True)
        st.metric("Total Expenses Shown", f"{df_exp['amount'].sum():,.2f} {CURRENCY}")

        st.markdown("#### ✏️ Edit or Delete an Expense")
        exp_labels = {f"#{row['id']} — {row['project_name']} — {row['date']} — {row['amount']:,.0f} {CURRENCY}": row["id"]
                      for _, row in df_exp.iterrows()}
        chosen = st.selectbox("Select record:", list(exp_labels.keys()), key="exp_edit_choice")
        eid = exp_labels[chosen]
        record = df_exp[df_exp["id"] == eid].iloc[0]

        with st.form("expense_edit_form"):
            try:
                edate = datetime.strptime(str(record["date"]), "%Y-%m-%d").date()
            except Exception:
                edate = date.today()
            new_date = st.date_input("Date:", value=edate, key="exp_edit_date")
            new_type = st.selectbox("Type:", expense_types,
                                     index=expense_types.index(record["expense_type"]) if record["expense_type"] in expense_types else 0)
            new_amount = st.number_input("Amount:", min_value=0.0, value=float(record["amount"]), format="%.2f")
            new_notes = st.text_area("Notes:", value=record.get("notes") or "")

            colA, colB = st.columns(2)
            save_btn = colA.form_submit_button("💾 Save Changes", use_container_width=True)
            with colB:
                confirm_del = st.checkbox("I confirm I want to delete this record", key="exp_confirm_del")
                del_btn = st.form_submit_button("🗑️ Delete", use_container_width=True, disabled=not confirm_del)

            if save_btn:
                conn = get_db_connection()
                conn.execute("UPDATE expenses SET date=?, expense_type=?, amount=?, notes=? WHERE id=?",
                             (str(new_date), new_type, new_amount, new_notes, eid))
                conn.commit()
                conn.close()
                st.success("Changes saved.")
                st.rerun()

            if del_btn and confirm_del:
                conn = get_db_connection()
                conn.execute("DELETE FROM expenses WHERE id=?", (eid,))
                conn.commit()
                conn.close()
                st.success("Record deleted.")
                st.rerun()


# =====================================================================================
# Page 5: Labor Costs (Add / Edit / Delete)
# =====================================================================================
def page_labor():
    st.title("👷 Labor Wage & Cost Calculation for Projects")
    st.markdown("---")

    p_options = get_project_options()
    if not p_options:
        st.warning("Please add a project first so you can allocate labor.")
        return

    tab_add, tab_manage = st.tabs(["➕ Record Labor Cost", "📋 List & Edit"])

    with tab_add:
        with st.form("labor_form", clear_on_submit=True):
            selected_p = st.selectbox("Select project:", list(p_options.keys()))
            worker_name = st.text_input("Worker / Team Name (optional):")
            labor_date = st.date_input("Date:")
            days = st.number_input("Total Days / Hours Allocated:", min_value=0.0, step=1.0)
            wage = st.number_input(f"Daily / Hourly Wage ({CURRENCY}):", min_value=0.0, format="%.2f")

            submit = st.form_submit_button("Calculate & Record Cost", use_container_width=True)
            if submit:
                total_labor_cost = days * wage
                conn = get_db_connection()
                conn.execute(
                    "INSERT INTO labor (project_id, worker_name, date, days, wage, total) VALUES (?,?,?,?,?,?)",
                    (p_options[selected_p], worker_name, str(labor_date), days, wage, total_labor_cost)
                )
                conn.commit()
                conn.close()
                st.success(f"Total labor cost of {total_labor_cost:,.2f} {CURRENCY} recorded successfully")

    with tab_manage:
        filter_p = st.selectbox("Filter by project:", ["All"] + list(p_options.keys()), key="labor_filter")
        if filter_p == "All":
            df_lab = fetch_df("""
                SELECT l.id, p.project_name, l.worker_name, l.date, l.days, l.wage, l.total
                FROM labor l JOIN projects p ON l.project_id = p.project_id
                ORDER BY l.date DESC
            """)
        else:
            df_lab = fetch_df("""
                SELECT l.id, p.project_name, l.worker_name, l.date, l.days, l.wage, l.total
                FROM labor l JOIN projects p ON l.project_id = p.project_id
                WHERE l.project_id = ? ORDER BY l.date DESC
            """, params=(p_options[filter_p],))

        if df_lab.empty:
            st.info("No labor records yet.")
            return

        st.dataframe(df_lab, use_container_width=True, hide_index=True)
        st.metric("Total Labor Cost Shown", f"{df_lab['total'].sum():,.2f} {CURRENCY}")

        st.markdown("#### ✏️ Edit or Delete a Labor Record")
        lab_labels = {f"#{row['id']} — {row['project_name']} — {row['date']} — {row['total']:,.0f} {CURRENCY}": row["id"]
                      for _, row in df_lab.iterrows()}
        chosen = st.selectbox("Select record:", list(lab_labels.keys()), key="lab_edit_choice")
        lid = lab_labels[chosen]
        record = df_lab[df_lab["id"] == lid].iloc[0]

        with st.form("labor_edit_form"):
            new_worker = st.text_input("Worker / Team Name:", value=record.get("worker_name") or "")
            try:
                ldate = datetime.strptime(str(record["date"]), "%Y-%m-%d").date()
            except Exception:
                ldate = date.today()
            new_date = st.date_input("Date:", value=ldate, key="lab_edit_date")
            new_days = st.number_input("Days/Hours:", min_value=0.0, value=float(record["days"] or 0), step=1.0)
            new_wage = st.number_input("Wage:", min_value=0.0, value=float(record["wage"] or 0), format="%.2f")

            colA, colB = st.columns(2)
            save_btn = colA.form_submit_button("💾 Save Changes", use_container_width=True)
            with colB:
                confirm_del = st.checkbox("I confirm I want to delete this record", key="lab_confirm_del")
                del_btn = st.form_submit_button("🗑️ Delete", use_container_width=True, disabled=not confirm_del)

            if save_btn:
                new_total = new_days * new_wage
                conn = get_db_connection()
                conn.execute("UPDATE labor SET worker_name=?, date=?, days=?, wage=?, total=? WHERE id=?",
                             (new_worker, str(new_date), new_days, new_wage, new_total, lid))
                conn.commit()
                conn.close()
                st.success("Changes saved.")
                st.rerun()

            if del_btn and confirm_del:
                conn = get_db_connection()
                conn.execute("DELETE FROM labor WHERE id=?", (lid,))
                conn.commit()
                conn.close()
                st.success("Record deleted.")
                st.rerun()


# =====================================================================================
# Page 6: Reports & Export
# =====================================================================================
def page_reports():
    st.title("📑 Reports & Export")
    st.markdown("---")

    df_summary = build_summary()
    if df_summary.empty:
        st.info("Not enough data to generate a report.")
        return

    p_options = get_project_options()
    chosen = st.multiselect("Select projects to include in the report (leave empty for all):", list(p_options.keys()))
    if chosen:
        ids = [p_options[c] for c in chosen]
        df_view = df_summary[df_summary["project_id"].isin(ids)]
    else:
        ids = list(p_options.values())
        df_view = df_summary

    st.dataframe(df_view, use_container_width=True, hide_index=True)

    df_revenues = fetch_df("""
        SELECT p.project_name, r.date, r.amount, r.stage, r.notes
        FROM revenues r JOIN projects p ON r.project_id = p.project_id
        WHERE r.project_id IN ({})
        ORDER BY r.date
    """.format(",".join("?" * len(ids))), params=tuple(ids)) if ids else pd.DataFrame()

    df_expenses = fetch_df("""
        SELECT p.project_name, e.date, e.expense_type, e.amount, e.notes
        FROM expenses e JOIN projects p ON e.project_id = p.project_id
        WHERE e.project_id IN ({})
        ORDER BY e.date
    """.format(",".join("?" * len(ids))), params=tuple(ids)) if ids else pd.DataFrame()

    df_labor = fetch_df("""
        SELECT p.project_name, l.worker_name, l.date, l.days, l.wage, l.total
        FROM labor l JOIN projects p ON l.project_id = p.project_id
        WHERE l.project_id IN ({})
        ORDER BY l.date
    """.format(",".join("?" * len(ids))), params=tuple(ids)) if ids else pd.DataFrame()

    st.markdown("### ⬇️ Download Report")
    col1, col2 = st.columns(2)

    with col1:
        excel_buffer = export_to_excel(df_view, df_revenues, df_expenses, df_labor)
        st.download_button(
            "📊 Download Excel Report", data=excel_buffer,
            file_name=f"SSC_Report_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with col2:
        totals = [
            ("Total Contract Value", df_view["contract_value"].sum()),
            ("Total Revenue", df_view["total_rev"].sum()),
            ("Total Expenses & Labor", df_view["total_costs"].sum()),
            ("Net Profit", df_view["net_profit"].sum()),
        ]
        pdf_buffer, font_ok, reshape_ok = export_to_pdf(df_view, totals)
        st.download_button(
            "📄 Download PDF Report", data=pdf_buffer,
            file_name=f"SSC_Report_{date.today().isoformat()}.pdf",
            mime="application/pdf",
            use_container_width=True
        )
        if not font_ok or not reshape_ok:
            issues = []
            if not font_ok:
                issues.append(
                    "Could not auto-download/register an Arabic-capable font (requires an internet "
                    f"connection on first run). You can place a font file manually (e.g. Amiri-Regular.ttf) "
                    f"next to the app, named: {ARABIC_FONT_PATH}"
                )
            if not reshape_ok:
                issues.append(
                    "The **arabic-reshaper** and **python-bidi** libraries are not installed; without them, "
                    "any Arabic text (e.g. in project/client names) will render disjointed and in reversed "
                    "order inside the PDF. Install them with:\n\n"
                    "`pip install arabic-reshaper python-bidi`"
                )
            st.warning("⚠️ Arabic text in the PDF report may not display correctly, for the following reason(s):\n\n- " + "\n- ".join(issues))
        else:
            st.caption("✅ Full Arabic-text support is enabled for the PDF (font + character shaping), in case any record contains Arabic.")


# =====================================================================================
# Page 7: Backup & Restore
# =====================================================================================
def page_backup():
    st.title("💾 Backup & Restore")
    st.markdown("---")

    st.markdown("#### ⬇️ Download a Backup")
    if os.path.exists(DB_PATH):
        with open(DB_PATH, "rb") as f:
            db_bytes = f.read()
        st.download_button(
            "Download Backup Now",
            data=db_bytes,
            file_name=f"ssc_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db",
            mime="application/octet-stream",
            use_container_width=True
        )
    else:
        st.warning("Database file not found.")

    st.markdown("---")
    st.markdown("#### ⬆️ Restore a Backup")
    st.warning("⚠️ Restoring a backup will delete all current data and replace it entirely with the uploaded file's contents. This action cannot be undone.")
    uploaded = st.file_uploader("Choose a backup file (.db):", type=["db"])
    confirm = st.checkbox("I confirm I want to completely replace the current data")
    if st.button("Restore Now", disabled=not (uploaded and confirm), use_container_width=True):
        with open(DB_PATH, "wb") as f:
            f.write(uploaded.getbuffer())
        st.success("Backup restored successfully. The system will now reload.")
        st.session_state.pop("auth_user", None)
        st.rerun()


# =====================================================================================
# Page 8: User Management (Admins only)
# =====================================================================================
def page_users():
    st.title("👤 User Management")
    st.markdown("---")

    tab_add, tab_manage = st.tabs(["➕ Add User", "📋 List & Edit"])

    with tab_add:
        with st.form("user_form", clear_on_submit=True):
            username = st.text_input("Username (for login):")
            full_name = st.text_input("Full Name:")
            password = st.text_input("Password:", type="password")
            role = st.selectbox("Role:", ["User", "Admin"])
            submit = st.form_submit_button("Add User", use_container_width=True)
            if submit:
                if username and password:
                    conn = get_db_connection()
                    try:
                        conn.execute(
                            "INSERT INTO users (username, full_name, password_hash, role, created_at) VALUES (?,?,?,?,?)",
                            (username.strip(), full_name, hash_password(password), role, str(datetime.now()))
                        )
                        conn.commit()
                        st.success(f"User ({username}) added successfully.")
                    except sqlite3.IntegrityError:
                        st.error("This username already exists.")
                    finally:
                        conn.close()
                else:
                    st.warning("Please fill in the username and password.")

    with tab_manage:
        df_u = fetch_df("SELECT id, username, full_name, role, created_at FROM users ORDER BY id")
        st.dataframe(df_u, use_container_width=True, hide_index=True)

        admin_count = int((df_u["role"] == "Admin").sum())
        labels = {f"{row['username']} ({row['full_name']})": row["id"] for _, row in df_u.iterrows()}
        chosen = st.selectbox("Select a user:", list(labels.keys()), key="user_edit_choice")
        uid = labels[chosen]
        record = df_u[df_u["id"] == uid].iloc[0]
        is_last_admin = record["role"] == "Admin" and admin_count <= 1

        with st.form("user_edit_form"):
            new_full_name = st.text_input("Full Name:", value=record["full_name"] or "")
            role_options = ["User", "Admin"]
            new_role = st.selectbox("Role:", role_options,
                                     index=role_options.index(record["role"]) if record["role"] in role_options else 0,
                                     disabled=is_last_admin)
            new_password = st.text_input("New Password (leave empty to keep unchanged):", type="password")

            colA, colB = st.columns(2)
            save_btn = colA.form_submit_button("💾 Save Changes", use_container_width=True)
            with colB:
                confirm_del = st.checkbox("I confirm I want to delete this user", key="user_confirm_del",
                                           disabled=is_last_admin)
                del_btn = st.form_submit_button("🗑️ Delete User", use_container_width=True,
                                                 disabled=not confirm_del or is_last_admin)

            if is_last_admin:
                st.caption("⚠️ The role of the last remaining Admin user in the system cannot be changed or deleted.")

            if save_btn:
                conn = get_db_connection()
                if new_password:
                    conn.execute("UPDATE users SET full_name=?, role=?, password_hash=? WHERE id=?",
                                 (new_full_name, new_role, hash_password(new_password), uid))
                else:
                    conn.execute("UPDATE users SET full_name=?, role=? WHERE id=?",
                                 (new_full_name, new_role, uid))
                conn.commit()
                conn.close()
                st.success("Changes saved.")
                st.rerun()

            if del_btn and confirm_del and not is_last_admin:
                conn = get_db_connection()
                conn.execute("DELETE FROM users WHERE id=?", (uid,))
                conn.commit()
                conn.close()
                st.success("User deleted.")
                st.rerun()


# =====================================================================================
# Main execution & navigation
# =====================================================================================
require_login()
user = current_user()

st.sidebar.title("🏗️ SSC Smart System")
st.sidebar.markdown(f"👋 Welcome, **{user.get('full_name') or user.get('username')}**  \n`{user.get('role')}`")
st.sidebar.markdown("---")

def _save_font_prefs(size, color):
    """Persists the current user's font preferences to the database so they
    are restored automatically the next time they log in (from any device)."""
    conn = get_db_connection()
    conn.execute("UPDATE users SET font_size=?, font_color=? WHERE id=?",
                 (size, color, user["id"]))
    conn.commit()
    conn.close()
    st.session_state.auth_user["font_size"] = size
    st.session_state.auth_user["font_color"] = color


with st.sidebar.expander("⚙️ Display Settings (font size & color)", expanded=False):
    chosen_size = st.slider(
        "Font size (px):", min_value=12, max_value=28,
        value=st.session_state.font_size, step=1, key="font_size_slider"
    )
    chosen_color = st.color_picker(
        "Font color:", value=st.session_state.font_color, key="font_color_picker"
    )
    # Persist only when something actually changed, to avoid hitting the DB on every rerun
    if chosen_size != st.session_state.font_size or chosen_color != st.session_state.font_color:
        st.session_state.font_size = chosen_size
        st.session_state.font_color = chosen_color
        _save_font_prefs(chosen_size, chosen_color)

    if st.button("↩️ Reset to Default", use_container_width=True, key="reset_font_btn"):
        st.session_state.font_size = 16
        st.session_state.font_color = "#1a1a1a"
        # Clear the widgets' own remembered state too — otherwise the slider/color
        # picker would keep showing their last position instead of resetting.
        for widget_key in ("font_size_slider", "font_color_picker"):
            st.session_state.pop(widget_key, None)
        _save_font_prefs(16, "#1a1a1a")
        st.rerun()

# Apply styling exactly once per run, using the now-current values
apply_dynamic_style()

st.sidebar.markdown("---")

menu_items = [
    "📊 Dashboard & Reports",
    "🏗️ Projects",
    "💰 Revenues",
    "📉 Expenses",
    "👷 Labor Costs",
    "📑 Reports & Export",
]
if is_admin():
    menu_items += ["💾 Backup", "👤 User Management"]

menu = st.sidebar.radio("Go to:", menu_items)

st.sidebar.markdown("---")
if st.sidebar.button("🚪 Log Out", use_container_width=True):
    st.session_state.pop("auth_user", None)
    st.rerun()

if menu == "📊 Dashboard & Reports":
    page_dashboard()
elif menu == "🏗️ Projects":
    page_projects()
elif menu == "💰 Revenues":
    page_revenues()
elif menu == "📉 Expenses":
    page_expenses()
elif menu == "👷 Labor Costs":
    page_labor()
elif menu == "📑 Reports & Export":
    page_reports()
elif menu == "💾 Backup" and is_admin():
    page_backup()
elif menu == "👤 User Management" and is_admin():
    page_users()
