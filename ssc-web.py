# -*- coding: utf-8 -*-
"""
نظام SSC لإدارة المشاريع — النسخة المطوّرة
=============================================
تطوير شامل يضيف: تسجيل دخول وصلاحيات، تعديل/حذف للسجلات، تنبيهات لوحة التحكم،
تصدير تقارير Excel وPDF، نسخ احتياطي/استرجاع لقاعدة البيانات، وواجهة محسّنة.
"""

import streamlit as st
import sqlite3
import pandas as pd
import plotly.express as px
import hashlib
import io
import os
from datetime import datetime, date

# =====================================================================================
# إعدادات عامة
# =====================================================================================
DB_PATH = "ssc_projects.db"
APP_TITLE = "نظام SSC لإدارة المشاريع"
# ضع ملف خط عربي (مثل Amiri-Regular.ttf أو NotoNaskhArabic-Regular.ttf) في نفس مجلد
# البرنامج وعدّل المسار أدناه لتفعيل عرض النصوص العربية بشكل صحيح داخل تقارير PDF.
ARABIC_FONT_PATH = "Amiri-Regular.ttf"

st.set_page_config(page_title=APP_TITLE, page_icon="🏗️", layout="wide")

# --- تنسيق عام للواجهة (CSS) ---
st.markdown("""
<style>
    html, body, [class*="css"]  { font-family: 'Tahoma', 'Segoe UI', sans-serif; }
    .stApp { direction: rtl; }
    div[data-testid="stMetric"] {
        background-color: #ffffff;
        border: 1px solid #e6e6e6;
        border-radius: 10px;
        padding: 14px 16px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    }
    div[data-testid="stMetricLabel"] { direction: rtl; text-align: right; }
    section[data-testid="stSidebar"] {
        background-color: #0f2540;
    }
    section[data-testid="stSidebar"] * { color: #f5f5f5 !important; }
    .ssc-alert-box {
        background-color: #fff4e5;
        border-right: 5px solid #e67e22;
        padding: 10px 14px;
        border-radius: 6px;
        margin-bottom: 8px;
    }
    .ssc-alert-danger {
        background-color: #fdecea;
        border-right: 5px solid #c0392b;
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
# إعدادات حجم ولون الخط (قابلة للتخصيص من القائمة الجانبية)
# =====================================================================================
if "font_size" not in st.session_state:
    st.session_state.font_size = 16          # px
if "font_color" not in st.session_state:
    st.session_state.font_color = "#1a1a1a"  # لون نص افتراضي (المنطقة الرئيسية فقط)


def apply_dynamic_style():
    """يحقن CSS بحجم ولون الخط المختارين من المستخدم، ويعيد فرض لون نص القائمة
    الجانبية في النهاية (لكي لا يتأثر بلون الخط العام المختار)."""
    size = st.session_state.font_size
    color = st.session_state.font_color
    st.markdown(f"""
    <style>
        html, body, [class*="css"], .stMarkdown, .stMarkdown p, .stText,
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
        /* إعادة فرض لون نص القائمة الجانبية بعد التخصيص العام أعلاه */
        section[data-testid="stSidebar"] * {{
            color: #f5f5f5 !important;
        }}
    </style>
    """, unsafe_allow_html=True)


apply_dynamic_style()


# =====================================================================================
# طبقة قاعدة البيانات
# =====================================================================================
def get_db_connection():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def _safe_add_column(cursor, table, coldef):
    """يضيف عموداً جديداً لجدول قديم إن لم يكن موجوداً (للحفاظ على التوافق مع قواعد بيانات سابقة)."""
    try:
        cursor.execute(f"ALTER TABLE {table} ADD COLUMN {coldef}")
    except sqlite3.OperationalError:
        pass  # العمود موجود مسبقاً


def hash_password(password: str, salt: str = "ssc_static_salt_v1") -> str:
    return hashlib.sha256((salt + password).encode("utf-8")).hexdigest()


def init_db():
    conn = get_db_connection()
    cur = conn.cursor()

    # جدول المشاريع
    cur.execute("""
        CREATE TABLE IF NOT EXISTS projects (
            project_id TEXT PRIMARY KEY,
            project_name TEXT NOT NULL,
            start_date TEXT,
            end_date TEXT,
            contract_value REAL
        )""")
    # ترقية الجدول القديم بأعمدة جديدة
    _safe_add_column(cur, "projects", "client_name TEXT")
    _safe_add_column(cur, "projects", "status TEXT DEFAULT 'جاري'")
    _safe_add_column(cur, "projects", "created_at TEXT")

    # جدول الإيرادات
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

    # جدول المصروفات
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

    # جدول العمالة
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

    # جدول المستخدمين
    cur.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            full_name TEXT,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'مستخدم',
            created_at TEXT
        )""")

    conn.commit()

    # إنشاء حساب المدير الافتراضي إن لم يوجد أي مستخدم
    cur.execute("SELECT COUNT(*) FROM users")
    if cur.fetchone()[0] == 0:
        cur.execute(
            "INSERT INTO users (username, full_name, password_hash, role, created_at) VALUES (?,?,?,?,?)",
            ("admin", "مدير النظام", hash_password("admin123"), "مدير", str(datetime.now()))
        )
        conn.commit()

    conn.close()


init_db()


# =====================================================================================
# طبقة المصادقة (تسجيل الدخول)
# =====================================================================================
def verify_user(username: str, password: str):
    conn = get_db_connection()
    row = conn.execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()
    conn.close()
    if row and row["password_hash"] == hash_password(password):
        return dict(row)
    return None


def login_page():
    st.title("🏗️ " + APP_TITLE)
    st.subheader("تسجيل الدخول")
    with st.form("login_form"):
        username = st.text_input("اسم المستخدم")
        password = st.text_input("كلمة المرور", type="password")
        submitted = st.form_submit_button("دخول", use_container_width=True)
        if submitted:
            user = verify_user(username.strip(), password)
            if user:
                st.session_state.auth_user = user
                st.rerun()
            else:
                st.error("اسم المستخدم أو كلمة المرور غير صحيحة.")
    st.caption("🔑 الحساب الافتراضي عند أول تشغيل: **admin** / **admin123** — يرجى تغييره فوراً من صفحة (إدارة المستخدمين).")


def require_login():
    if "auth_user" not in st.session_state:
        login_page()
        st.stop()


def current_user():
    return st.session_state.get("auth_user")


def is_admin():
    user = current_user()
    return bool(user and user.get("role") == "مدير")


# =====================================================================================
# دوال مساعدة لجلب البيانات
# =====================================================================================
def fetch_df(query, params=()):
    conn = get_db_connection()
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df


def get_projects_df():
    return fetch_df("SELECT * FROM projects ORDER BY created_at DESC")


def get_project_options():
    """يرجع قاموس {اسم المشروع: رقمه} لاستخدامه في القوائم المنسدلة."""
    df = get_projects_df()
    return {row["project_name"]: row["project_id"] for _, row in df.iterrows()}


def build_summary():
    """يبني جدول ملخص مالي شامل لكل مشروع (إيرادات، مصروفات، عمالة، صافي الربح، نسبة التحصيل)."""
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

    df["اجمالي_المصروفات"] = df["total_exp"] + df["total_labor"]
    df["صافي_الربح"] = df["total_rev"] - df["اجمالي_المصروفات"]
    df["نسبة_التحصيل"] = df.apply(
        lambda r: round((r["total_rev"] / r["contract_value"]) * 100, 1) if r["contract_value"] else 0, axis=1
    )
    df["هامش_الربح"] = df.apply(
        lambda r: round((r["صافي_الربح"] / r["total_rev"]) * 100, 1) if r["total_rev"] else 0, axis=1
    )
    return df


def build_alerts(df_summary):
    """يفحص بيانات المشاريع ويولّد تنبيهات (تجاوز التكلفة، اقتراب الانتهاء، عدم وجود إيرادات حديثة)."""
    alerts = []
    if df_summary.empty:
        return alerts

    today = date.today()
    for _, row in df_summary.iterrows():
        name = row["project_name"]
        status = row.get("status", "جاري") or "جاري"

        # تجاوز الميزانية
        if row["اجمالي_المصروفات"] > row["contract_value"] and row["contract_value"] > 0:
            over = row["اجمالي_المصروفات"] - row["contract_value"]
            alerts.append(("danger", f"⚠️ مشروع «{name}»: المصروفات تجاوزت قيمة العقد بمقدار {over:,.0f} ر.س"))

        # اقتراب موعد الانتهاء (خلال 30 يوماً) وما زال جارياً
        if status == "جاري" and row.get("end_date"):
            try:
                end_dt = datetime.strptime(str(row["end_date"]), "%Y-%m-%d").date()
                days_left = (end_dt - today).days
                if 0 <= days_left <= 30:
                    alerts.append(("warning", f"⏰ مشروع «{name}»: متبقٍ {days_left} يوماً على تاريخ الانتهاء المتوقع"))
                elif days_left < 0:
                    alerts.append(("danger", f"⌛ مشروع «{name}»: تجاوز تاريخ الانتهاء المتوقع بـ {abs(days_left)} يوماً وما زال «جاري»"))
            except ValueError:
                pass

        # نسبة تحصيل منخفضة لمشروع منتهٍ
        if status == "منتهي" and row["نسبة_التحصيل"] < 90:
            alerts.append(("warning", f"💰 مشروع «{name}» منتهٍ ولكن نسبة التحصيل {row['نسبة_التحصيل']}% فقط من قيمة العقد"))

    return alerts


# =====================================================================================
# التصدير: Excel و PDF
# =====================================================================================
def export_to_excel(df_summary, df_revenues, df_expenses, df_labor):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_summary.to_excel(writer, sheet_name="ملخص المشاريع", index=False)
        df_revenues.to_excel(writer, sheet_name="الإيرادات", index=False)
        df_expenses.to_excel(writer, sheet_name="المصروفات", index=False)
        df_labor.to_excel(writer, sheet_name="العمالة", index=False)
    buffer.seek(0)
    return buffer


def _get_pdf_font():
    """يحاول تسجيل خط عربي. إن لم يتوفر، يعود لخط Helvetica (لن يعرض الحروف العربية بشكل صحيح)."""
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    if os.path.exists(ARABIC_FONT_PATH):
        try:
            pdfmetrics.registerFont(TTFont("ArabicFont", ARABIC_FONT_PATH))
            return "ArabicFont", True
        except Exception:
            pass
    return "Helvetica", False


def _ar(text):
    """يهيئ النص العربي للعرض الصحيح (اتجاه ودمج الحروف) إذا توفرت المكتبات اللازمة."""
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
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, topMargin=1.5 * cm, bottomMargin=1.5 * cm)

    title_style = ParagraphStyle("title", fontName=font_name, fontSize=16, alignment=1, spaceAfter=12)
    normal_style = ParagraphStyle("normal", fontName=font_name, fontSize=9, alignment=1)

    elements = []
    elements.append(Paragraph(_ar(f"التقرير المالي — {APP_TITLE}"), title_style))
    elements.append(Paragraph(_ar(f"تاريخ التقرير: {date.today().isoformat()}"), normal_style))
    elements.append(Spacer(1, 10))

    # جدول ملخص عام
    summary_rows = [[_ar("القيمة"), _ar("البند")]]
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

    # جدول تفصيلي بالمشاريع
    header = [_ar(h) for h in ["صافي الربح", "إجمالي المصروفات", "الإيرادات", "قيمة العقد", "اسم المشروع"]]
    rows = [header]
    for _, r in df_summary.iterrows():
        rows.append([
            f"{r['صافي_الربح']:,.0f}",
            f"{r['اجمالي_المصروفات']:,.0f}",
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
    return buffer, font_ok


# =====================================================================================
# الصفحة 1: لوحة التحكم والتقارير
# =====================================================================================
def page_dashboard():
    st.title("📊 لوحة التحكم والتحليل المالي اللحظي")
    st.markdown("---")

    df_summary = build_summary()

    if df_summary.empty:
        st.info("لا توجد مشاريع مسجلة حالياً. يرجى إضافة مشاريع من القائمة الجانبية.")
        return

    # --- فلاتر ---
    with st.expander("🔎 الفلاتر", expanded=False):
        statuses = sorted(df_summary["status"].fillna("جاري").unique().tolist())
        chosen_status = st.multiselect("حالة المشروع:", statuses, default=statuses)
    df_view = df_summary[df_summary["status"].fillna("جاري").isin(chosen_status)] if chosen_status else df_summary

    # --- تنبيهات ---
    st.markdown("#### 🚨 تنبيهات")
    alerts = build_alerts(df_view)
    if alerts:
        for level, msg in alerts:
            css_class = "ssc-alert-danger" if level == "danger" else "ssc-alert-box"
            st.markdown(f'<div class="{css_class}">{msg}</div>', unsafe_allow_html=True)
    else:
        st.success("✅ لا توجد تنبيهات حالياً — كل المشاريع ضمن الحدود الطبيعية.")
    st.markdown("---")

    # --- المؤشرات الرئيسية ---
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("إجمالي قيمة العقود", f"{df_view['contract_value'].sum():,.0f} ر.س")
    col2.metric("إجمالي الإيرادات المحصلة", f"{df_view['total_rev'].sum():,.0f} ر.س")
    col3.metric("إجمالي المصاريف والعمالة", f"{df_view['اجمالي_المصروفات'].sum():,.0f} ر.س")
    profit = df_view["صافي_الربح"].sum()
    col4.metric("صافي الأرباح الإجمالية", f"{profit:,.0f} ر.س")

    col5, col6 = st.columns(2)
    total_contract = df_view["contract_value"].sum()
    collection_rate = (df_view["total_rev"].sum() / total_contract * 100) if total_contract else 0
    col5.metric("نسبة التحصيل العامة", f"{collection_rate:,.1f}%")
    total_rev_sum = df_view["total_rev"].sum()
    margin = (profit / total_rev_sum * 100) if total_rev_sum else 0
    col6.metric("هامش الربح العام", f"{margin:,.1f}%")

    st.markdown("### 📋 ملخص أداء المشاريع")
    display_cols = ["project_id", "project_name", "status", "contract_value", "total_rev",
                     "اجمالي_المصروفات", "صافي_الربح", "نسبة_التحصيل", "هامش_الربح"]
    display_cols = [c for c in display_cols if c in df_view.columns]
    st.dataframe(
        df_view[display_cols].rename(columns={
            "project_id": "رقم المشروع", "project_name": "اسم المشروع", "status": "الحالة",
            "contract_value": "قيمة العقد", "total_rev": "الإيرادات",
            "نسبة_التحصيل": "نسبة التحصيل %", "هامش_الربح": "هامش الربح %",
        }),
        use_container_width=True, hide_index=True
    )

    st.markdown("### 📈 تحليلات بيانية")
    c1, c2 = st.columns(2)
    with c1:
        fig = px.bar(
            df_view, x="project_name", y=["total_rev", "اجمالي_المصروفات", "صافي_الربح"],
            barmode="group", title="مقارنة الإيرادات والمصروفات والأرباح لكل مشروع",
            labels={"value": "المبلغ (ر.س)", "project_name": "المشروع", "variable": "البند"}
        )
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        df_exp_type = fetch_df("SELECT expense_type, SUM(amount) as amount FROM expenses GROUP BY expense_type")
        if not df_exp_type.empty:
            fig2 = px.pie(df_exp_type, names="expense_type", values="amount", title="توزيع المصروفات حسب النوع")
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.info("لا توجد مصروفات مسجلة لعرض توزيعها.")


# =====================================================================================
# الصفحة 2: المشاريع (إضافة / تعديل / حذف)
# =====================================================================================
def page_projects():
    st.title("🏗️ إدارة المشاريع")
    st.markdown("---")
    tab_add, tab_manage = st.tabs(["➕ إضافة مشروع جديد", "📋 القائمة والتعديل"])

    with tab_add:
        with st.form("project_form", clear_on_submit=True):
            p_id = st.text_input("رقم المشروع (كود فريد):")
            p_name = st.text_input("اسم المشروع:")
            client_name = st.text_input("اسم العميل:")
            col1, col2 = st.columns(2)
            start_d = col1.date_input("تاريخ البداية")
            end_d = col2.date_input("تاريخ النهاية المتوقع")
            val = st.number_input("قيمة العقد الإجمالية (ر.س):", min_value=0.0, format="%.2f")
            status = st.selectbox("حالة المشروع:", ["جاري", "متوقف", "منتهي"])

            submit = st.form_submit_button("حفظ المشروع", use_container_width=True)
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
                        st.success(f"تم تسجيل مشروع ({p_name}) بنجاح!")
                    except sqlite3.IntegrityError:
                        st.error("رقم المشروع مسجل مسبقاً! يرجى استخدام رقم فريد.")
                    finally:
                        conn.close()
                else:
                    st.warning("يرجى ملء الحقول الأساسية (الرقم والاسم).")

    with tab_manage:
        df_p = get_projects_df()
        if df_p.empty:
            st.info("لا توجد مشاريع مسجلة بعد.")
            return

        st.dataframe(df_p, use_container_width=True, hide_index=True)
        st.markdown("#### ✏️ تعديل أو حذف مشروع")
        options = {f"{row['project_name']} ({row['project_id']})": row["project_id"] for _, row in df_p.iterrows()}
        choice = st.selectbox("اختر المشروع:", list(options.keys()), key="proj_edit_choice")
        pid = options[choice]
        record = df_p[df_p["project_id"] == pid].iloc[0]

        with st.form("project_edit_form"):
            new_name = st.text_input("اسم المشروع:", value=record["project_name"])
            new_client = st.text_input("اسم العميل:", value=record.get("client_name") or "")
            col1, col2 = st.columns(2)
            try:
                sd = datetime.strptime(str(record["start_date"]), "%Y-%m-%d").date()
            except Exception:
                sd = date.today()
            try:
                ed = datetime.strptime(str(record["end_date"]), "%Y-%m-%d").date()
            except Exception:
                ed = date.today()
            new_start = col1.date_input("تاريخ البداية:", value=sd, key="edit_start")
            new_end = col2.date_input("تاريخ النهاية:", value=ed, key="edit_end")
            new_val = st.number_input("قيمة العقد:", min_value=0.0, value=float(record["contract_value"] or 0), format="%.2f")
            status_options = ["جاري", "متوقف", "منتهي"]
            current_status = record.get("status") or "جاري"
            new_status = st.selectbox("الحالة:", status_options,
                                       index=status_options.index(current_status) if current_status in status_options else 0)

            colA, colB = st.columns(2)
            save_btn = colA.form_submit_button("💾 حفظ التعديلات", use_container_width=True)
            with colB:
                confirm_del = st.checkbox("أؤكد رغبتي في حذف هذا المشروع وكل بياناته المرتبطة")
                del_btn = st.form_submit_button("🗑️ حذف المشروع نهائياً", use_container_width=True, disabled=not confirm_del)

            if save_btn:
                conn = get_db_connection()
                conn.execute(
                    "UPDATE projects SET project_name=?, client_name=?, start_date=?, end_date=?, "
                    "contract_value=?, status=? WHERE project_id=?",
                    (new_name, new_client, str(new_start), str(new_end), new_val, new_status, pid)
                )
                conn.commit()
                conn.close()
                st.success("تم حفظ التعديلات بنجاح.")
                st.rerun()

            if del_btn and confirm_del:
                conn = get_db_connection()
                conn.execute("DELETE FROM revenues WHERE project_id=?", (pid,))
                conn.execute("DELETE FROM expenses WHERE project_id=?", (pid,))
                conn.execute("DELETE FROM labor WHERE project_id=?", (pid,))
                conn.execute("DELETE FROM projects WHERE project_id=?", (pid,))
                conn.commit()
                conn.close()
                st.success("تم حذف المشروع وكل بياناته المرتبطة.")
                st.rerun()


# =====================================================================================
# الصفحة 3: الإيرادات (إضافة / تعديل / حذف)
# =====================================================================================
def page_revenues():
    st.title("💰 إدارة الدفعات والإيرادات")
    st.markdown("---")

    p_options = get_project_options()
    if not p_options:
        st.warning("يرجى إضافة مشروع أولاً لتتمكن من إضافة إيرادات.")
        return

    tab_add, tab_manage = st.tabs(["➕ تسجيل دفعة", "📋 القائمة والتعديل"])

    with tab_add:
        with st.form("revenue_form", clear_on_submit=True):
            selected_p = st.selectbox("اختر المشروع:", list(p_options.keys()))
            rev_date = st.date_input("تاريخ استلام الدفعة")
            stage = st.selectbox("نوع الدفعة / المرحلة:", ["دفعة 1", "دفعة 2", "دفعة 3", "دفعة 4", "دفعة ختامية"])
            amount = st.number_input("مبلغ الدفعة (ر.س):", min_value=0.0, format="%.2f")
            notes = st.text_area("ملاحظات:")

            submit = st.form_submit_button("تسجيل الدفعة", use_container_width=True)
            if submit:
                conn = get_db_connection()
                conn.execute(
                    "INSERT INTO revenues (project_id, date, amount, stage, notes) VALUES (?,?,?,?,?)",
                    (p_options[selected_p], str(rev_date), amount, stage, notes)
                )
                conn.commit()
                conn.close()
                st.success(f"تم تسجيل {stage} بمبلغ {amount:,.2f} ر.س للمشروع بنجاح!")

    with tab_manage:
        filter_p = st.selectbox("تصفية حسب المشروع:", ["الكل"] + list(p_options.keys()), key="rev_filter")
        if filter_p == "الكل":
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
            st.info("لا توجد إيرادات مسجلة.")
            return

        st.dataframe(df_rev, use_container_width=True, hide_index=True)
        st.metric("إجمالي الإيرادات المعروضة", f"{df_rev['amount'].sum():,.2f} ر.س")

        st.markdown("#### ✏️ تعديل أو حذف دفعة")
        rev_labels = {f"#{row['id']} — {row['project_name']} — {row['date']} — {row['amount']:,.0f} ر.س": row["id"]
                      for _, row in df_rev.iterrows()}
        chosen = st.selectbox("اختر السجل:", list(rev_labels.keys()), key="rev_edit_choice")
        rid = rev_labels[chosen]
        record = df_rev[df_rev["id"] == rid].iloc[0]

        with st.form("revenue_edit_form"):
            try:
                rdate = datetime.strptime(str(record["date"]), "%Y-%m-%d").date()
            except Exception:
                rdate = date.today()
            new_date = st.date_input("التاريخ:", value=rdate, key="rev_edit_date")
            stage_options = ["دفعة 1", "دفعة 2", "دفعة 3", "دفعة 4", "دفعة ختامية"]
            new_stage = st.selectbox("المرحلة:", stage_options,
                                      index=stage_options.index(record["stage"]) if record["stage"] in stage_options else 0)
            new_amount = st.number_input("المبلغ:", min_value=0.0, value=float(record["amount"]), format="%.2f")
            new_notes = st.text_area("ملاحظات:", value=record.get("notes") or "")

            colA, colB = st.columns(2)
            save_btn = colA.form_submit_button("💾 حفظ التعديلات", use_container_width=True)
            with colB:
                confirm_del = st.checkbox("أؤكد رغبتي في حذف هذا السجل", key="rev_confirm_del")
                del_btn = st.form_submit_button("🗑️ حذف", use_container_width=True, disabled=not confirm_del)

            if save_btn:
                conn = get_db_connection()
                conn.execute("UPDATE revenues SET date=?, amount=?, stage=?, notes=? WHERE id=?",
                             (str(new_date), new_amount, new_stage, new_notes, rid))
                conn.commit()
                conn.close()
                st.success("تم حفظ التعديلات.")
                st.rerun()

            if del_btn and confirm_del:
                conn = get_db_connection()
                conn.execute("DELETE FROM revenues WHERE id=?", (rid,))
                conn.commit()
                conn.close()
                st.success("تم حذف السجل.")
                st.rerun()


# =====================================================================================
# الصفحة 4: المصروفات (إضافة / تعديل / حذف)
# =====================================================================================
def page_expenses():
    st.title("📉 إدارة وتحليل المصروفات التشغيلية")
    st.markdown("---")

    p_options = get_project_options()
    if not p_options:
        st.warning("يرجى إضافة مشروع أولاً لتتمكن من إضافة مصروفات.")
        return

    tab_add, tab_manage = st.tabs(["➕ تسجيل مصروف", "📋 القائمة والتعديل"])
    expense_types = ["مواد", "نقل", "معدات", "موقع", "أخرى"]

    with tab_add:
        with st.form("expense_form", clear_on_submit=True):
            selected_p = st.selectbox("اختر المشروع الموجه له المصروف:", list(p_options.keys()))
            exp_date = st.date_input("تاريخ الصرف")
            exp_type = st.selectbox("نوع المصروف (تحليل المصروفات):", expense_types)
            amount = st.number_input("القيمة (ر.س):", min_value=0.0, format="%.2f")
            notes = st.text_area("ملاحظات إضافية:")

            submit = st.form_submit_button("تسجيل المصروف", use_container_width=True)
            if submit:
                conn = get_db_connection()
                conn.execute(
                    "INSERT INTO expenses (project_id, date, expense_type, amount, notes) VALUES (?,?,?,?,?)",
                    (p_options[selected_p], str(exp_date), exp_type, amount, notes)
                )
                conn.commit()
                conn.close()
                st.success("تم تسجيل المصروف وتحديث بند التحليل المالي!")

    with tab_manage:
        filter_p = st.selectbox("تصفية حسب المشروع:", ["الكل"] + list(p_options.keys()), key="exp_filter")
        if filter_p == "الكل":
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
            st.info("لا توجد مصروفات مسجلة.")
            return

        st.dataframe(df_exp, use_container_width=True, hide_index=True)
        st.metric("إجمالي المصروفات المعروضة", f"{df_exp['amount'].sum():,.2f} ر.س")

        st.markdown("#### ✏️ تعديل أو حذف مصروف")
        exp_labels = {f"#{row['id']} — {row['project_name']} — {row['date']} — {row['amount']:,.0f} ر.س": row["id"]
                      for _, row in df_exp.iterrows()}
        chosen = st.selectbox("اختر السجل:", list(exp_labels.keys()), key="exp_edit_choice")
        eid = exp_labels[chosen]
        record = df_exp[df_exp["id"] == eid].iloc[0]

        with st.form("expense_edit_form"):
            try:
                edate = datetime.strptime(str(record["date"]), "%Y-%m-%d").date()
            except Exception:
                edate = date.today()
            new_date = st.date_input("التاريخ:", value=edate, key="exp_edit_date")
            new_type = st.selectbox("النوع:", expense_types,
                                     index=expense_types.index(record["expense_type"]) if record["expense_type"] in expense_types else 0)
            new_amount = st.number_input("القيمة:", min_value=0.0, value=float(record["amount"]), format="%.2f")
            new_notes = st.text_area("ملاحظات:", value=record.get("notes") or "")

            colA, colB = st.columns(2)
            save_btn = colA.form_submit_button("💾 حفظ التعديلات", use_container_width=True)
            with colB:
                confirm_del = st.checkbox("أؤكد رغبتي في حذف هذا السجل", key="exp_confirm_del")
                del_btn = st.form_submit_button("🗑️ حذف", use_container_width=True, disabled=not confirm_del)

            if save_btn:
                conn = get_db_connection()
                conn.execute("UPDATE expenses SET date=?, expense_type=?, amount=?, notes=? WHERE id=?",
                             (str(new_date), new_type, new_amount, new_notes, eid))
                conn.commit()
                conn.close()
                st.success("تم حفظ التعديلات.")
                st.rerun()

            if del_btn and confirm_del:
                conn = get_db_connection()
                conn.execute("DELETE FROM expenses WHERE id=?", (eid,))
                conn.commit()
                conn.close()
                st.success("تم حذف السجل.")
                st.rerun()


# =====================================================================================
# الصفحة 5: تكاليف العمالة (إضافة / تعديل / حذف)
# =====================================================================================
def page_labor():
    st.title("👷 احتساب أجور وتكاليف العمالة للمشاريع")
    st.markdown("---")

    p_options = get_project_options()
    if not p_options:
        st.warning("يرجى إضافة مشروع أولاً لتتمكن من تخصيص عمالة.")
        return

    tab_add, tab_manage = st.tabs(["➕ تسجيل تكلفة عمالة", "📋 القائمة والتعديل"])

    with tab_add:
        with st.form("labor_form", clear_on_submit=True):
            selected_p = st.selectbox("اختر المشروع:", list(p_options.keys()))
            worker_name = st.text_input("اسم العامل / الفريق (اختياري):")
            labor_date = st.date_input("التاريخ:")
            days = st.number_input("عدد الأيام / الساعات الكلية المخصصة:", min_value=0.0, step=1.0)
            wage = st.number_input("أجر اليوم الواحد / الساعة (ر.س):", min_value=0.0, format="%.2f")

            submit = st.form_submit_button("احسب وسجل التكلفة", use_container_width=True)
            if submit:
                total_labor_cost = days * wage
                conn = get_db_connection()
                conn.execute(
                    "INSERT INTO labor (project_id, worker_name, date, days, wage, total) VALUES (?,?,?,?,?,?)",
                    (p_options[selected_p], worker_name, str(labor_date), days, wage, total_labor_cost)
                )
                conn.commit()
                conn.close()
                st.success(f"تم تسجيل التكلفة الإجمالية للعمالة بمبلغ {total_labor_cost:,.2f} ر.س")

    with tab_manage:
        filter_p = st.selectbox("تصفية حسب المشروع:", ["الكل"] + list(p_options.keys()), key="labor_filter")
        if filter_p == "الكل":
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
            st.info("لا توجد سجلات عمالة.")
            return

        st.dataframe(df_lab, use_container_width=True, hide_index=True)
        st.metric("إجمالي تكلفة العمالة المعروضة", f"{df_lab['total'].sum():,.2f} ر.س")

        st.markdown("#### ✏️ تعديل أو حذف سجل عمالة")
        lab_labels = {f"#{row['id']} — {row['project_name']} — {row['date']} — {row['total']:,.0f} ر.س": row["id"]
                      for _, row in df_lab.iterrows()}
        chosen = st.selectbox("اختر السجل:", list(lab_labels.keys()), key="lab_edit_choice")
        lid = lab_labels[chosen]
        record = df_lab[df_lab["id"] == lid].iloc[0]

        with st.form("labor_edit_form"):
            new_worker = st.text_input("اسم العامل / الفريق:", value=record.get("worker_name") or "")
            try:
                ldate = datetime.strptime(str(record["date"]), "%Y-%m-%d").date()
            except Exception:
                ldate = date.today()
            new_date = st.date_input("التاريخ:", value=ldate, key="lab_edit_date")
            new_days = st.number_input("عدد الأيام/الساعات:", min_value=0.0, value=float(record["days"] or 0), step=1.0)
            new_wage = st.number_input("الأجر:", min_value=0.0, value=float(record["wage"] or 0), format="%.2f")

            colA, colB = st.columns(2)
            save_btn = colA.form_submit_button("💾 حفظ التعديلات", use_container_width=True)
            with colB:
                confirm_del = st.checkbox("أؤكد رغبتي في حذف هذا السجل", key="lab_confirm_del")
                del_btn = st.form_submit_button("🗑️ حذف", use_container_width=True, disabled=not confirm_del)

            if save_btn:
                new_total = new_days * new_wage
                conn = get_db_connection()
                conn.execute("UPDATE labor SET worker_name=?, date=?, days=?, wage=?, total=? WHERE id=?",
                             (new_worker, str(new_date), new_days, new_wage, new_total, lid))
                conn.commit()
                conn.close()
                st.success("تم حفظ التعديلات.")
                st.rerun()

            if del_btn and confirm_del:
                conn = get_db_connection()
                conn.execute("DELETE FROM labor WHERE id=?", (lid,))
                conn.commit()
                conn.close()
                st.success("تم حذف السجل.")
                st.rerun()


# =====================================================================================
# الصفحة 6: التقارير والتصدير
# =====================================================================================
def page_reports():
    st.title("📑 التقارير والتصدير")
    st.markdown("---")

    df_summary = build_summary()
    if df_summary.empty:
        st.info("لا توجد بيانات كافية لإصدار تقرير.")
        return

    p_options = get_project_options()
    chosen = st.multiselect("اختر المشاريع المطلوبة في التقرير (اتركها فارغة لتشمل الكل):", list(p_options.keys()))
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

    st.markdown("### ⬇️ تحميل التقرير")
    col1, col2 = st.columns(2)

    with col1:
        excel_buffer = export_to_excel(df_view, df_revenues, df_expenses, df_labor)
        st.download_button(
            "📊 تحميل تقرير Excel", data=excel_buffer,
            file_name=f"تقرير_SSC_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with col2:
        totals = [
            ("إجمالي قيمة العقود", df_view["contract_value"].sum()),
            ("إجمالي الإيرادات", df_view["total_rev"].sum()),
            ("إجمالي المصروفات والعمالة", df_view["اجمالي_المصروفات"].sum()),
            ("صافي الربح", df_view["صافي_الربح"].sum()),
        ]
        pdf_buffer, font_ok = export_to_pdf(df_view, totals)
        st.download_button(
            "📄 تحميل تقرير PDF", data=pdf_buffer,
            file_name=f"تقرير_SSC_{date.today().isoformat()}.pdf",
            mime="application/pdf",
            use_container_width=True
        )
        if not font_ok:
            st.caption(
                f"⚠️ لتفعيل عرض النصوص العربية بشكل صحيح في PDF، ضع ملف خط عربي "
                f"(مثل Amiri-Regular.ttf) بجانب البرنامج باسم: {ARABIC_FONT_PATH}"
            )


# =====================================================================================
# الصفحة 7: النسخ الاحتياطي والاستعادة
# =====================================================================================
def page_backup():
    st.title("💾 النسخ الاحتياطي والاستعادة")
    st.markdown("---")

    st.markdown("#### ⬇️ تحميل نسخة احتياطية")
    if os.path.exists(DB_PATH):
        with open(DB_PATH, "rb") as f:
            db_bytes = f.read()
        st.download_button(
            "تحميل نسخة احتياطية الآن",
            data=db_bytes,
            file_name=f"ssc_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db",
            mime="application/octet-stream",
            use_container_width=True
        )
    else:
        st.warning("لم يتم العثور على ملف قاعدة البيانات.")

    st.markdown("---")
    st.markdown("#### ⬆️ استرجاع نسخة احتياطية")
    st.warning("⚠️ استرجاع نسخة احتياطية سيحذف كل البيانات الحالية ويستبدلها بالكامل بمحتوى الملف المرفوع. هذا الإجراء لا يمكن التراجع عنه.")
    uploaded = st.file_uploader("اختر ملف نسخة احتياطية (.db):", type=["db"])
    confirm = st.checkbox("أؤكد أنني أريد استبدال البيانات الحالية بالكامل")
    if st.button("استرجاع الآن", disabled=not (uploaded and confirm), use_container_width=True):
        with open(DB_PATH, "wb") as f:
            f.write(uploaded.getbuffer())
        st.success("تم استرجاع النسخة الاحتياطية بنجاح. سيتم إعادة تحميل النظام.")
        st.session_state.pop("auth_user", None)
        st.rerun()


# =====================================================================================
# الصفحة 8: إدارة المستخدمين (للمدير فقط)
# =====================================================================================
def page_users():
    st.title("👤 إدارة المستخدمين")
    st.markdown("---")

    tab_add, tab_manage = st.tabs(["➕ إضافة مستخدم", "📋 القائمة والتعديل"])

    with tab_add:
        with st.form("user_form", clear_on_submit=True):
            username = st.text_input("اسم المستخدم (للدخول):")
            full_name = st.text_input("الاسم الكامل:")
            password = st.text_input("كلمة المرور:", type="password")
            role = st.selectbox("الصلاحية:", ["مستخدم", "مدير"])
            submit = st.form_submit_button("إضافة المستخدم", use_container_width=True)
            if submit:
                if username and password:
                    conn = get_db_connection()
                    try:
                        conn.execute(
                            "INSERT INTO users (username, full_name, password_hash, role, created_at) VALUES (?,?,?,?,?)",
                            (username.strip(), full_name, hash_password(password), role, str(datetime.now()))
                        )
                        conn.commit()
                        st.success(f"تم إضافة المستخدم ({username}) بنجاح.")
                    except sqlite3.IntegrityError:
                        st.error("اسم المستخدم موجود مسبقاً.")
                    finally:
                        conn.close()
                else:
                    st.warning("يرجى تعبئة اسم المستخدم وكلمة المرور.")

    with tab_manage:
        df_u = fetch_df("SELECT id, username, full_name, role, created_at FROM users ORDER BY id")
        st.dataframe(df_u, use_container_width=True, hide_index=True)

        admin_count = int((df_u["role"] == "مدير").sum())
        labels = {f"{row['username']} ({row['full_name']})": row["id"] for _, row in df_u.iterrows()}
        chosen = st.selectbox("اختر مستخدماً:", list(labels.keys()), key="user_edit_choice")
        uid = labels[chosen]
        record = df_u[df_u["id"] == uid].iloc[0]
        is_last_admin = record["role"] == "مدير" and admin_count <= 1

        with st.form("user_edit_form"):
            new_full_name = st.text_input("الاسم الكامل:", value=record["full_name"] or "")
            role_options = ["مستخدم", "مدير"]
            new_role = st.selectbox("الصلاحية:", role_options,
                                     index=role_options.index(record["role"]) if record["role"] in role_options else 0,
                                     disabled=is_last_admin)
            new_password = st.text_input("كلمة مرور جديدة (اتركها فارغة لعدم التغيير):", type="password")

            colA, colB = st.columns(2)
            save_btn = colA.form_submit_button("💾 حفظ التعديلات", use_container_width=True)
            with colB:
                confirm_del = st.checkbox("أؤكد رغبتي في حذف هذا المستخدم", key="user_confirm_del",
                                           disabled=is_last_admin)
                del_btn = st.form_submit_button("🗑️ حذف المستخدم", use_container_width=True,
                                                 disabled=not confirm_del or is_last_admin)

            if is_last_admin:
                st.caption("⚠️ لا يمكن تعديل صلاحية أو حذف آخر مستخدم مدير في النظام.")

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
                st.success("تم حفظ التعديلات.")
                st.rerun()

            if del_btn and confirm_del and not is_last_admin:
                conn = get_db_connection()
                conn.execute("DELETE FROM users WHERE id=?", (uid,))
                conn.commit()
                conn.close()
                st.success("تم حذف المستخدم.")
                st.rerun()


# =====================================================================================
# التنفيذ الرئيسي والتنقل
# =====================================================================================
require_login()
user = current_user()

st.sidebar.title("🏗️ نظام SSC الذكي")
st.sidebar.markdown(f"👋 أهلاً، **{user.get('full_name') or user.get('username')}**  \n`{user.get('role')}`")
st.sidebar.markdown("---")

with st.sidebar.expander("⚙️ إعدادات العرض (حجم ولون الخط)", expanded=False):
    st.session_state.font_size = st.slider(
        "حجم الخط (px):", min_value=12, max_value=28,
        value=st.session_state.font_size, step=1, key="font_size_slider"
    )
    st.session_state.font_color = st.color_picker(
        "لون الخط:", value=st.session_state.font_color, key="font_color_picker"
    )
    if st.button("↩️ إعادة الضبط الافتراضي", use_container_width=True, key="reset_font_btn"):
        st.session_state.font_size = 16
        st.session_state.font_color = "#1a1a1a"
        st.rerun()

# إعادة تطبيق التنسيق بعد قراءة القيم الحالية (تشمل أي تغيير من المستخدم أعلاه)
apply_dynamic_style()

st.sidebar.markdown("---")

menu_items = [
    "📊 لوحة التحكم والتقارير",
    "🏗️ المشاريع",
    "💰 الإيرادات",
    "📉 المصروفات",
    "👷 تكاليف العمالة",
    "📑 التقارير والتصدير",
]
if is_admin():
    menu_items += ["💾 النسخ الاحتياطي", "👤 إدارة المستخدمين"]

menu = st.sidebar.radio("الانتقال إلى:", menu_items)

st.sidebar.markdown("---")
if st.sidebar.button("🚪 تسجيل الخروج", use_container_width=True):
    st.session_state.pop("auth_user", None)
    st.rerun()

if menu == "📊 لوحة التحكم والتقارير":
    page_dashboard()
elif menu == "🏗️ المشاريع":
    page_projects()
elif menu == "💰 الإيرادات":
    page_revenues()
elif menu == "📉 المصروفات":
    page_expenses()
elif menu == "👷 تكاليف العمالة":
    page_labor()
elif menu == "📑 التقارير والتصدير":
    page_reports()
elif menu == "💾 النسخ الاحتياطي" and is_admin():
    page_backup()
elif menu == "👤 إدارة المستخدمين" and is_admin():
    page_users()
