import streamlit as st
import sqlite3
import pandas as pd
import plotly.express as px
from datetime import datetime

# --- إعدادات الصفحة ---
st.set_page_config(page_title="نظام SSC لإدارة المشاريع", page_icon="🏗️", layout="wide")

# --- إدارة قاعدة البيانات ---
def get_db_connection():
    """إنشاء اتصال آمن بقاعدة البيانات"""
    conn = sqlite3.connect('ssc_projects.db')
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    """تهيئة قاعدة البيانات وإنشاء الجداول"""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        # جدول المشاريع
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS projects (
                project_id TEXT PRIMARY KEY,
                project_name TEXT NOT NULL,
                start_date TEXT,
                end_date TEXT,
                contract_value REAL,
                created_date TEXT DEFAULT CURRENT_TIMESTAMP
            )''')
        # جدول الإيرادات
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS revenues (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id TEXT,
                date TEXT,
                amount REAL,
                stage TEXT,
                created_date TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY(project_id) REFERENCES projects(project_id)
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
                created_date TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY(project_id) REFERENCES projects(project_id)
            )''')
        # جدول العمالة
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS labor (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id TEXT,
                days INTEGER,
                wage REAL,
                total REAL,
                created_date TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY(project_id) REFERENCES projects(project_id)
            )''')
        conn.commit()
    except sqlite3.Error as e:
        st.error(f"خطأ في قاعدة البيانات: {e}")
    finally:
        conn.close()

init_db()

# --- دوال التحقق من صحة المدخلات ---
def validate_amount(amount, min_val=0, max_val=999999999):
    """التحقق من صحة قيمة المبلغ"""
    return min_val <= amount <= max_val

def validate_days(days, min_val=1, max_val=365):
    """التحقق من صحة عدد الأيام"""
    return min_val <= days <= max_val

def validate_project_id(project_id):
    """التحقق من صحة رقم المشروع"""
    return len(project_id.strip()) > 0 and len(project_id) <= 50

def validate_project_name(project_name):
    """التحقق من صحة اسم المشروع"""
    return len(project_name.strip()) > 0 and len(project_name) <= 100

# --- القائمة الجانبية للتنقل ---
st.sidebar.title("🏗️ نظام SSC الذكي")
st.sidebar.markdown("---")
menu = st.sidebar.radio("الانتقال إلى:", [
    "📊 لوحة التحكم والتقارير", 
    "➕ إضافة مشروع جديد", 
    "💰 تسجيل الإيرادات", 
    "📉 تسجيل المصروفات",
    "👷 تكاليف العمالة"
])

# --- 1. لوحة التحكم والتقارير ---
if menu == "📊 لوحة التحكم والتقارير":
    st.title("📊 لوحة التحكم والتحليل المالي اللحظي")
    st.markdown("---")
    
    conn = get_db_connection()
    try:
        # جلب البيانات لعمل الحسابات
        df_p = pd.read_sql_query("SELECT * FROM projects", conn)
        df_r = pd.read_sql_query("SELECT project_id, SUM(amount) as total_rev FROM revenues GROUP BY project_id", conn)
        df_e = pd.read_sql_query("SELECT project_id, SUM(amount) as total_exp FROM expenses GROUP BY project_id", conn)
        df_l = pd.read_sql_query("SELECT project_id, SUM(total) as total_labor FROM labor GROUP BY project_id", conn)
        
        if not df_p.empty:
            # دمج الجداول لحساب الصافي
            df_summary = df_p.merge(df_r, on='project_id', how='left').merge(df_e, on='project_id', how='left').merge(df_l, on='project_id', how='left')
            df_summary.fillna(0, inplace=True)
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
            st.dataframe(df_summary[['project_id', 'project_name', 'contract_value', 'total_rev', 'اجمالي_المصروفات', 'صافي_الربح']], use_container_width=True)
            
            # رسومات بيانية تفاعلية
            st.markdown("### 📈 تحليلات بيانية")
            fig = px.bar(df_summary, x='project_name', y=['total_rev', 'اجمالي_المصروفات', 'صافي_الربح'], 
                        barmode='group', title='مقارنة الإيرادات والمصروفات والأرباح للمشاريع')
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("لا توجد مشاريع مسجلة حالياً. يرجى إضافة مشاريع من القائمة الجانبية.")
    except Exception as e:
        st.error(f"حدث خطأ أثناء جلب البيانات: {e}")
    finally:
        conn.close()

# --- 2. إضافة مشروع جديد ---
elif menu == "➕ إضافة مشروع جديد":
    st.title("➕ تسجيل مشروع جديد في النظام")
    st.markdown("---")
    
    with st.form("project_form"):
        p_id = st.text_input("رقم المشروع (كود فريد):")
        p_name = st.text_input("اسم المشروع:")
        col1, col2 = st.columns(2)
        start_d = col1.date_input("تاريخ البداية")
        end_d = col2.date_input("تاريخ النهاية المتوقع")
        val = st.number_input("قيمة العقد الإجمالية (ر.س):", min_value=0.0, format="%.2f")
        
        submit = st.form_submit_button("حفظ المشروع")
        if submit:
            # التحقق من صحة المدخلات
            if not validate_project_id(p_id):
                st.error("⚠️ رقم المشروع يجب أن يكون غير فارغ ولا يزيد عن 50 حرف.")
            elif not validate_project_name(p_name):
                st.error("⚠️ اسم المشروع يجب أن يكون غير فارغ ولا يزيد عن 100 حرف.")
            elif not validate_amount(val):
                st.error("⚠️ قيمة العقد يجب أن تكون بين 0 و 999,999,999.")
            elif start_d >= end_d:
                st.error("⚠️ تاريخ النهاية يجب أن يكون بعد تاريخ البداية.")
            else:
                conn = get_db_connection()
                try:
                    conn.execute("INSERT INTO projects VALUES (?, ?, ?, ?, ?, ?)", 
                                (p_id, p_name, str(start_d), str(end_d), val, datetime.now().isoformat()))
                    conn.commit()
                    st.success(f"✅ تم تسجيل مشروع ({p_name}) بنجاح!")
                except sqlite3.IntegrityError:
                    st.error("❌ رقم المشروع مسجل مسبقاً! يرجى استخدام رقم فريد.")
                except Exception as e:
                    st.error(f"❌ خطأ أثناء حفظ المشروع: {e}")
                finally:
                    conn.close()

# --- 3. تسجيل الإيرادات ---
elif menu == "💰 تسجيل الإيرادات":
    st.title("💰 إدارة الدفعات والإيرادات")
    st.markdown("---")
    
    conn = get_db_connection()
    try:
        projects = pd.read_sql_query("SELECT project_id, project_name FROM projects ORDER BY project_name", conn)
        
        if not projects.empty:
            p_options = {row['project_name']: row['project_id'] for _, row in projects.iterrows()}
            selected_p = st.selectbox("اختر المشروع:", list(p_options.keys()))
            
            with st.form("revenue_form"):
                rev_date = st.date_input("تاريخ استلام الدفعة")
                stage = st.selectbox("نوع الدفعة / المرحلة:", ["دفعة 1", "دفعة 2", "دفعة 3", "دفعة 4", "دفعة ختامية"])
                amount = st.number_input("مبلغ الدفعة (ر.س):", min_value=0.0, format="%.2f")
                
                submit = st.form_submit_button("تسجيل الدفعة")
                if submit:
                    if not validate_amount(amount):
                        st.error("⚠️ المبلغ يجب أن يكون أكبر من صفر وأقل من 999,999,999.")
                    else:
                        conn_insert = get_db_connection()
                        try:
                            conn_insert.execute("INSERT INTO revenues (project_id, date, amount, stage, created_date) VALUES (?, ?, ?, ?, ?)", 
                                            (p_options[selected_p], str(rev_date), amount, stage, datetime.now().isoformat()))
                            conn_insert.commit()
                            st.success(f"✅ تم تسجيل {stage} بمبلغ {amount:,.2f} ر.س للمشروع بنجاح!")
                        except Exception as e:
                            st.error(f"❌ خطأ أثناء تسجيل الدفعة: {e}")
                        finally:
                            conn_insert.close()
        else:
            st.warning("⚠️ يرجى إضافة مشروع أولاً لتتمكن من إضافة إيرادات.")
    except Exception as e:
        st.error(f"خطأ أثناء جلب المشاريع: {e}")
    finally:
        conn.close()

# --- 4. تسجيل المصروفات ---
elif menu == "📉 تسجيل المصروفات":
    st.title("📉 إدارة وتحليل المصروفات التشغيلية")
    st.markdown("---")
    
    conn = get_db_connection()
    try:
        projects = pd.read_sql_query("SELECT project_id, project_name FROM projects ORDER BY project_name", conn)
        
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
                    if not validate_amount(amount):
                        st.error("⚠️ المبلغ يجب أن يكون أكبر من صفر وأقل من 999,999,999.")
                    else:
                        conn_insert = get_db_connection()
                        try:
                            conn_insert.execute("INSERT INTO expenses (project_id, date, expense_type, amount, notes, created_date) VALUES (?, ?, ?, ?, ?, ?)", 
                                            (p_options[selected_p], str(exp_date), exp_type, amount, notes, datetime.now().isoformat()))
                            conn_insert.commit()
                            st.success("✅ تم تسجيل المصروف وتحديث بند التحليل المالي!")
                        except Exception as e:
                            st.error(f"❌ خطأ أثناء تسجيل المصروف: {e}")
                        finally:
                            conn_insert.close()
        else:
            st.warning("⚠️ يرجى إضافة مشروع أولاً لتتمكن من إضافة مصروفات.")
    except Exception as e:
        st.error(f"خطأ أثناء جلب المشاريع: {e}")
    finally:
        conn.close()

# --- 5. تكاليف العمالة ---
elif menu == "👷 تكاليف العمالة":
    st.title("👷 احتساب أجور وتكاليف العمالة للمشاريع")
    st.markdown("---")
    
    conn = get_db_connection()
    try:
        projects = pd.read_sql_query("SELECT project_id, project_name FROM projects ORDER BY project_name", conn)
        
        if not projects.empty:
            p_options = {row['project_name']: row['project_id'] for _, row in projects.iterrows()}
            selected_p = st.selectbox("اختر المشروع:", list(p_options.keys()))
            
            with st.form("labor_form"):
                days = st.number_input("عدد الأيام / الساعات الكلية المخصصة:", min_value=1, step=1)
                wage = st.number_input("أجر اليوم الواحد / الساعة (ر.س):", min_value=0.0, format="%.2f")
                
                submit = st.form_submit_button("احسب وسجل التكلفة")
                if submit:
                    if not validate_days(days):
                        st.error("⚠️ عدد الأيام يجب أن يكون بين 1 و 365.")
                    elif not validate_amount(wage):
                        st.error("⚠️ قيمة الأجر يجب أن تكون موجبة وأقل من 999,999,999.")
                    else:
                        total_labor_cost = days * wage
                        conn_insert = get_db_connection()
                        try:
                            conn_insert.execute("INSERT INTO labor (project_id, days, wage, total, created_date) VALUES (?, ?, ?, ?, ?)", 
                                            (p_options[selected_p], days, wage, total_labor_cost, datetime.now().isoformat()))
                            conn_insert.commit()
                            st.success(f"✅ تم تسجيل التكلفة الإجمالية للعمالة بمبلغ {total_labor_cost:,.2f} ر.س")
                        except Exception as e:
                            st.error(f"❌ خطأ أثناء تسجيل تكلفة العمالة: {e}")
                        finally:
                            conn_insert.close()
        else:
            st.warning("⚠️ يرجى إضافة مشروع أولاً لتتمكن من تخصيص عمالة.")
    except Exception as e:
        st.error(f"خطأ أثناء جلب المشاريع: {e}")
    finally:
        conn.close()
