import streamlit as st
import sqlite3
import pandas as pd
import plotly.express as px
from datetime import datetime

# --- إعدادات الصفحة ---
st.set_page_config(page_title="نظام SSC لإدارة المشاريع", page_icon="🏗️", layout="wide")

# --- إدارة قاعدة البيانات ---
def get_db_connection():
    conn = sqlite3.connect('ssc_projects.db')
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()
    # جدول المشاريع
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS projects (
            project_id TEXT PRIMARY KEY,
            project_name TEXT NOT NULL,
            start_date TEXT,
            end_date TEXT,
            contract_value REAL
        )''')
    # جدول الإيرادات
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS revenues (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id TEXT,
            date TEXT,
            amount REAL,
            stage TEXT,
            FOREIGN KEY(project_id) REFERENCES projects(project_id) ON DELETE CASCADE
        )''')
    # جدول المصروفات
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS expenses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id TEXT,
            date TEXT,
            expense_type TEXT,
            amount REAL,
            notes TEXT,
            FOREIGN KEY(project_id) REFERENCES projects(project_id) ON DELETE CASCADE
        )''')
    # جدول العمالة
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS labor (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id TEXT,
            worker_count INTEGER DEFAULT 1,
            days INTEGER,
            wage REAL,
            total REAL,
            FOREIGN KEY(project_id) REFERENCES projects(project_id) ON DELETE CASCADE
        )''')
    
    # فحص عمود worker_count
    cursor.execute("PRAGMA table_info(labor)")
    columns = [column[1] for column in cursor.fetchall()]
    if 'worker_count' not in columns:
        try:
            cursor.execute("ALTER TABLE labor ADD COLUMN worker_count INTEGER DEFAULT 1")
        except sqlite3.OperationalError:
            pass

    conn.commit()
    conn.close()

init_db()

# --- القائمة الجانبية للتنقل ---
st.sidebar.title("🏗️ نظام SSC الذكي")
st.sidebar.markdown("---")
menu = st.sidebar.radio("الانتقال إلى:", [
    "📊 لوحة التحكم والتقارير", 
    "➕ إضافة وإدارة المشاريع", 
    "💰 تسجيل الإيرادات", 
    "📉 تسجيل المصروفات", 
    "👷 تكاليف العمالة"
])

# --- 1. لوحة التحكم والتقارير ---
if menu == "📊 لوحة التحكم والتقارير":
    st.title("📊 لوحة التحكم والتحليل المالي اللحظي")
    st.markdown("---")
    
    conn = get_db_connection()
    df_p = pd.read_sql_query("SELECT * FROM projects", conn)
    df_r = pd.read_sql_query("SELECT project_id, amount FROM revenues", conn)
    df_e = pd.read_sql_query("SELECT project_id, amount FROM expenses", conn)
    df_l = pd.read_sql_query("SELECT project_id, total FROM labor", conn)
    conn.close()
    
    if not df_p.empty:
        # حساب المجاميع لكل جدول على حدة بشكل آمن مع تجميعها حسب كود المشروع
        df_r_grouped = df_r.groupby('project_id')['amount'].sum().reset_index(name='total_rev') if not df_r.empty else pd.DataFrame(columns=['project_id', 'total_rev'])
        df_e_grouped = df_e.groupby('project_id')['amount'].sum().reset_index(name='total_exp') if not df_e.empty else pd.DataFrame(columns=['project_id', 'total_exp'])
        df_l_grouped = df_l.groupby('project_id')['total'].sum().reset_index(name='total_labor') if not df_l.empty else pd.DataFrame(columns=['project_id', 'total_labor'])
        
        # ربط الجداول بالجدول الرئيسي للمشاريع (Left Join) لضمان عدم اختفاء أي مشروع
        df_summary = df_p.copy()
        df_summary = df_summary.merge(df_r_grouped, on='project_id', how='left')
        df_summary = df_summary.merge(df_e_grouped, on='project_id', how='left')
        df_summary = df_summary.merge(df_l_grouped, on='project_id', how='left')
        
        # تحويل القيم الفارغة (NaN) إلى أصفار لإجراء العمليات الحسابية
        df_summary['total_rev'] = df_summary['total_rev'].fillna(0)
        df_summary['total_exp'] = df_summary['total_exp'].fillna(0)
        df_summary['total_labor'] = df_summary['total_labor'].fillna(0)
        
        # العمليات الحسابية النهائية
        df_summary['اجمالي_المصروفات'] = df_summary['total_exp'] + df_summary['total_labor']
        df_summary['صافي_الربح'] = df_summary['total_rev'] - df_summary['اجمالي_المصروفات']
        
        # كروت المؤشرات المالية العامة
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("إجمالي قيمة العقود", f"{df_summary['contract_value'].sum():,.2f} ر.س")
        col2.metric("إجمالي الإيرادات المحصلة", f"{df_summary['total_rev'].sum():,.2f} ر.س")
        col3.metric("إجمالي المصاريف والعمالة", f"{df_summary['اجمالي_المصروفات'].sum():,.2f} ر.س")
        
        profit = df_summary['صافي_الربح'].sum()
        col4.metric("صافي الأرباح الإجمالية", f"{profit:,.2f} ر.س", delta=f"{profit:,.2f}", delta_color="normal")
        
        st.markdown("### 📋 ملخص أداء المشاريع")
        
        # تنسيق الجدول للعرض المباشر
        df_display = df_summary.rename(columns={
            'project_id': 'كود المشروع',
            'project_name': 'اسم المشروع',
            'contract_value': 'قيمة العقد',
            'total_rev': 'إجمالي المحصل',
            'اجمالي_المصروفات': 'إجمالي المصاريف',
            'صافي_الربح': 'صافي الربح'
        })
        
        # عرض البيانات وتنسيق الأرقام بفاصلة الآلاف وعلامة العملة
        st.dataframe(
            df_display[['كود المشروع', 'اسم المشروع', 'قيمة العقد', 'إجمالي المحصل', 'إجمالي المصاريف', 'صافي_الربح']], 
            use_container_width=True
        )
        
        st.markdown("### 📈 تحليلات بيانية")
        fig = px.bar(df_summary, x='project_name', y=['total_rev', 'اجمالي_المصروفات', 'صافي_الربح'], 
                     barmode='group', title='مقارنة الإيرادات والمصروفات والأرباح لكل مشروع',
                     labels={'value': 'المبلغ (ر.س)', 'project_name': 'اسم المشروع', 'variable': 'التصنيف'})
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("لا توجد مشاريع مسجلة حالياً. يرجى إضافة مشاريع من القائمة الجانبية.")

# --- 2. إضافة وإدارة المشروع ---
elif menu == "➕ إضافة وإدارة المشاريع":
    st.title("➕ تسجيل مشروع جديد أو حذفه")
    st.markdown("---")
    
    tab1, tab2 = st.tabs(["إضافة مشروع جديد", "إدارة المشاريع الحالية"])
    
    with tab1:
        with st.form("project_form"):
            p_id = st.text_input("رقم المشروع (كود فريد):")
            p_name = st.text_input("اسم المشروع:")
            col1, col2 = st.columns(2)
            start_d = col1.date_input("تاريخ البداية")
            end_d = col2.date_input("تاريخ النهاية المتوقع")
            val = st.number_input("قيمة العقد الإجمالية (ر.س):", min_value=0.0, format="%.2f")
            
            submit = st.form_submit_button("حفظ المشروع")
            if submit:
                if p_id and p_name:
                    conn = get_db_connection()
                    try:
                        conn.execute("INSERT INTO projects VALUES (?, ?, ?, ?, ?)", (p_id, p_name, str(start_d), str(end_d), val))
                        conn.commit()
                        st.success(f"تم تسجيل مشروع ({p_name}) بنجاح!")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        st.error("رقم المشروع مسجل مسبقاً! يرجى استخدام رقم فريد.")
                    finally:
                        conn.close()
                else:
                    st.warning("يرجى ملء الحقول الأساسية (الرقم والاسم).")
                    
    with tab2:
        conn = get_db_connection()
        projects_df = pd.read_sql_query("SELECT * FROM projects", conn)
        conn.close()
        
        if not projects_df.empty:
            st.write("المشاريع المسجلة حالياً:")
            projects_display = projects_df.rename(columns={
                'project_id': 'كود المشروع',
                'project_name': 'اسم المشروع',
                'start_date': 'تاريخ البداية',
                'end_date': 'تاريخ النهاية المتوقع',
                'contract_value': 'إجمالي قيمة العقد (ر.س)'
            })
            st.dataframe(projects_display, use_container_width=True)
            
            project_to_delete = st.selectbox("اختر مشروعاً لحذفه نهائياً:", projects_df['project_name'].tolist(), key="del_proj")
            p_id_to_del = projects_df[projects_df['project_name'] == project_to_delete]['project_id'].values[0]
            
            if st.button("❌ حذف المشروع المختار بجميع بياناته", type="primary"):
                conn = get_db_connection()
                conn.execute("PRAGMA foreign_keys = ON")
                conn.execute("DELETE FROM projects WHERE project_id = ?", (p_id_to_del,))
                conn.commit()
                conn.close()
                st.success(f"تم حذف المشروع {project_to_delete} بنجاح!")
                st.rerun()
        else:
            st.info("لا توجد مشاريع مضافة بعد.")

# --- 3. تسجيل الإيرادات ---
elif menu == "💰 تسجيل الإيرادات":
    st.title("💰 إدارة الدفعات والإيرادات")
    st.markdown("---")
    
    conn = get_db_connection()
    projects = pd.read_sql_query("SELECT project_id, project_name FROM projects", conn)
    conn.close()
    
    if not projects.empty:
        p_options = {row['project_name']: row['project_id'] for _, row in projects.iterrows()}
        selected_p = st.selectbox("اختر المشروع:", list(p_options.keys()))
        
        with st.form("revenue_form"):
            rev_date = st.date_input("تاريخ استلام الدفعة")
            stage = st.selectbox("نوع الدفعة / المرحلة:", ["دفعة 1", "دفعة 2", "دفعة 3", "دفعة 4", "دفعة ختامية"])
            amount = st.number_input("مبلغ الدفعة (ر.س):", min_value=0.0, format="%.2f")
            
            submit = st.form_submit_button("تسجيل الدفعة")
            if submit:
                conn = get_db_connection()
                conn.execute("INSERT INTO revenues (project_id, date, amount, stage) VALUES (?, ?, ?, ?)", 
                             (p_options[selected_p], str(rev_date), amount, stage))
                conn.commit()
                conn.close()
                st.success(f"تم تسجيل {stage} بمبلغ {amount:,.2f} ر.س للمشروع بنجاح!")
    else:
        st.warning("يرجى إضافة مشروع أولاً لتتمكن من إضافة إيرادات.")

# --- 4. تسجيل المصروفات ---
elif menu == "📉 تسجيل المصروفات":
    st.title("📉 إدارة وتحليل المصروفات التشغيلية")
    st.markdown("---")
    
    conn = get_db_connection()
    projects = pd.read_sql_query("SELECT project_id, project_name FROM projects", conn)
    conn.close()
    
    if not projects.empty:
        p_options = {row['project_name']: row['project_id'] for _, row in projects.iterrows()}
        selected_p = st.selectbox("اختر المشروع الموجه له المصروف:", list(p_options.keys()))
        
        with st.form("expense_form"):
            exp_date = st.date_input("تاريخ الصرف")
            exp_type = st.selectbox("نوع المصروف (تحليل المصروفات):", ["مواد", "نقل", "معدات", "موقع", "أخرى"])
            amount = st.number_input("القيمة (ر.س):", min_value=0.0, format="%.2f")
            notes = st.text_area("ملاحظات إضافية:")
            
            submit = st.form_submit_button("تسجيل المصروف")
            if submit:
                conn = get_db_connection()
                conn.execute("INSERT INTO expenses (project_id, date, expense_type, amount, notes) VALUES (?, ?, ?, ?, ?)", 
                             (p_options[selected_p], str(exp_date), exp_type, amount, notes))
                conn.commit()
                conn.close()
                st.success("تم تسجيل المصروف وتحديث بند التحليل المالي!")
    else:
        st.warning("يرجى إضافة مشروع أولاً لتتمكن من إضافة مصروفات.")

# --- 5. تكاليف العمالة ---
elif menu == "👷 تكاليف العمالة":
    st.title("👷 احتساب أجور وتكاليف العمالة للمشاريع")
    st.markdown("---")
    
    conn = get_db_connection()
    projects = pd.read_sql_query("SELECT project_id, project_name FROM projects", conn)
    conn.close()
    
    if not projects.empty:
        p_options = {row['project_name']: row['project_id'] for _, row in projects.iterrows()}
        selected_p = st.selectbox("اختر المشروع:", list(p_options.keys()))
        
        with st.form("labor_form"):
            workers = st.number_input("عدد العمال المخصصين:", min_value=1, step=1, value=1)
            days = st.number_input("عدد الأيام / الساعات الكلية المخصصة لكل عامل:", min_value=1, step=1)
            wage = st.number_input("أجر العامل الواحد لليوم / الساعة (ر.س):", min_value=0.0, format="%.2f")
            
            submit = st.form_submit_button("احسب وسجل التكلفة")
            if submit:
                total_labor_cost = workers * days * wage
                conn = get_db_connection()
                conn.execute("INSERT INTO labor (project_id, worker_count, days, wage, total) VALUES (?, ?, ?, ?, ?)", 
                             (p_options[selected_p], workers, days, wage, total_labor_cost))
                conn.commit()
                conn.close()
                st.success(f"تم تسجيل التكلفة الإجمالية لعدد ({workers}) عمال بمبلغ {total_labor_cost:,.2f} ر.س")
    else:
        st.warning("يرجى إضافة مشروع أولاً لتتمكن من تخصيص عمالة.")
