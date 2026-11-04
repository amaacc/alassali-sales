"""
app.py — Daily Showrooms Sales Cloud Web Application (DailyShowroomsSales Replacement)
Built with Streamlit + SQLAlchemy + Playwright + Odoo XML-RPC + Plotly
Designed for: Al-Assali Trading Co. (شركة العسلي للتجارة)
Developer: Abdullah Al-Khatib
"""

import os
import sys
import io
import time
from datetime import datetime, date
from calendar import monthrange
import pandas as pd
import streamlit as st

# Setup paths
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

from database import (
    SessionLocal, init_db, Branch, City, Showroom, OldClosing, ReportBatch,
    get_database_status_info
)
from processor import (
    SalesDataProcessor, parse_date_value, OldClosingsFileUploadedError,
    generate_monthly_excel, generate_periodic_excel, generate_comparative_excel
)
from pdf_generator import PDFReportGenerator
from odoo_client import OdooClient
import charts
from charts import PLOTLY_CONFIG
from security_gate import enforce_security_activation_gate, get_machine_id

# =============================================================================
# STREAMLIT CONFIG & RTL STYLING
# =============================================================================
st.set_page_config(
    page_title="نظام مبيعات وإقفالات المعارض اليومي | شركة العسلي",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&display=swap');

html, body, [class*="css"], .stMarkdown, .stText, .stButton, .stTextInput, .stSelectbox {
    font-family: 'Cairo', 'Segoe UI', Tahoma, sans-serif !important;
    direction: rtl;
    text-align: right;
}

/* Metric card styling */
div[data-testid="stMetricValue"] {
    font-size: 1.6rem !important;
    font-weight: 700;
}

.kpi-card {
    background: #FFFFFF;
    border-radius: 10px;
    padding: 16px;
    border: 1px solid #E2E8F0;
    box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    margin-bottom: 12px;
}

.alert-box {
    padding: 14px 18px;
    border-radius: 8px;
    margin-bottom: 12px;
    font-size: 11pt;
}

.alert-red {
    background-color: #FEF2F2;
    border: 1px solid #F87171;
    color: #991B1B;
}

.alert-blue {
    background-color: #EFF6FF;
    border: 1px solid #60A5FA;
    color: #1E40AF;
}

.alert-green {
    background-color: #ECFDF5;
    border: 1px solid #34D399;
    color: #065F46;
}

/* Tabs header RTL */
button[data-baseweb="tab"] {
    font-size: 1.05rem !important;
    font-weight: 600 !important;
}

/* Dataframe header styling */
.dataframe th {
    background-color: #F1F5F9 !important;
    color: #0F172A !important;
    text-align: center !important;
}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# =============================================================================
# 1. HARDWARE ID ACTIVATION GATE (FIRST LINE OF DEFENSE)
# =============================================================================
enforce_security_activation_gate()

# Initialize database schema
init_db()

# =============================================================================
# 2. USER AUTHENTICATION GATE (SECOND LINE OF DEFENSE)
# =============================================================================
APP_PASSWORD = os.environ.get("APP_PASSWORD")
if not APP_PASSWORD:
    try:
        if hasattr(st, "secrets") and "APP_PASSWORD" in st.secrets:
            APP_PASSWORD = st.secrets["APP_PASSWORD"]
    except Exception:
        pass
APP_PASSWORD = APP_PASSWORD or "assali2026"

def check_password():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if st.session_state.authenticated:
        return True

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.write("")
        st.write("")
        st.markdown(
            """
            <div style="text-align: center; padding: 25px; background: white; border-radius: 12px; border: 1px solid #CBD5E1;">
                <h2 style="color: #1E3A8A; margin-bottom: 5px;">🏢 شركة العسلي للتجارة</h2>
                <h4 style="color: #475569; margin-top: 0;">نظام مراقبة وإصدار تقارير مبيعات المعارض اليومي</h4>
                <p style="color: #64748B; font-size: 10pt;">النسخة السحابية المعتمدة (بديل Microsoft Access)</p>
            </div>
            """,
            unsafe_allow_html=True
        )
        st.write("")
        pwd = st.text_input("🔐 كلمة المرور للوصول إلى النظام:", type="password")
        if st.button("تسجيل الدخول", use_container_width=True, type="primary"):
            if pwd == APP_PASSWORD:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("❌ كلمة المرور غير صحيحة، يرجى المحاولة مرة أخرى.")
    return False

if not check_password():
    st.stop()

@st.cache_data(ttl=60)
def get_cached_system_stats():
    """Caches database counters for 60 seconds to eliminate query lag on every Streamlit rerun."""
    db = SessionLocal()
    try:
        old_count = db.query(OldClosing).count()
        batch_count = db.query(ReportBatch).count()
        showrooms_count = db.query(Showroom).count()
        return showrooms_count, old_count, batch_count
    finally:
        db.close()

# =============================================================================
# SIDEBAR
# =============================================================================
with st.sidebar:
    st.markdown("### 🏢 شركة العسلي للتجارة")
    st.caption("نظام مبيعات المعارض وإقفالات نقاط البيع (POS)")
    st.markdown("---")

    # Security & Machine Info
    dev_mid = get_machine_id()
    db_status = get_database_status_info()

    st.markdown("#### 🛡️ أمان وترخيص الجهاز:")
    st.write(f"• كود الجهاز: `{dev_mid}`")
    st.write("• حالة الترخيص: 🟢 **مفعل ومعتمد**")
    st.write(f"• قاعدة البيانات: **{db_status['type']}**")
    st.markdown("---")

    showrooms_count, old_count, batch_count = get_cached_system_stats()

    st.markdown("#### 📌 إحصائيات النظام:")
    st.write(f"• المعارض المسجلة: **{showrooms_count}** معرضاً")
    st.write(f"• الجلسات المعتمدة (tblOLD): **{old_count:,}** جلسة")
    st.write(f"• الدفعات المؤرشفة: **{batch_count}** دفعة")

    st.markdown("---")
    if st.button("🚪 تسجيل الخروج", use_container_width=True):
        st.session_state.authenticated = False
        st.rerun()

# =============================================================================
# MAIN APP TABS
# =============================================================================
st.title("📊 نظام مراقبة وإصدار تقارير مبيعات المعارض اليومي")

tab_radar, tab_data, tab_print, tab_monthly, tab_archive = st.tabs([
    "🛰️ الرادار الصباحي الذكي (Morning Radar)",
    "🛒 جلب بيانات المبيعات والمعاينة (Sales Data)",
    "📄 طباعة التقارير اليومية والاعتماد (Daily Reports)",
    "📊 التقارير الدورية والمقارنات (Periodic & Comparative Reports)",
    "🗄️ الأرشيف، الإعدادات والتراجع (Archive & Settings)"
])

# Initialize session state objects
if "raw_records" not in st.session_state:
    st.session_state.raw_records = []
if "processed_summary" not in st.session_state:
    st.session_state.processed_summary = None
if "active_date_str" not in st.session_state:
    st.session_state.active_date_str = date.today().strftime("%Y-%m-%d")

# =============================================================================
# TAB 1: SMART MORNING RADAR
# =============================================================================
with tab_radar:
    st.subheader("🛰️ الرادار الصباحي الذكي لفحص جلسات نقاط البيع (Odoo POS Radar)")
    st.info(
        "يقوم الرادار بفحص جلسات أودو لحظياً، وتحديد المعارض المتأخرة في الإقفال، والتمييز الذكي بين الجلسات "
        "الرئيسية الجاهزة للطباعة والجلسات الفرعية الشهرية المسموح ببقائها حتى نهاية الشهر."
    )

    col1, col2, col3 = st.columns([1.5, 1.5, 2])
    with col1:
        radar_date = st.date_input("تاريخ يوم التقرير المستهدف:", value=date.today())
        target_date_str = radar_date.strftime("%Y-%m-%d")
    with col2:
        allow_secondary = st.checkbox("السماح ببقاء الجلسات الفرعية الشهرية مفتوحة", value=True)
    with col3:
        st.write("")
        st.write("")
        btn_scan = st.button("🔍 مسح جلسات نقاط البيع الآن", type="primary", use_container_width=True)

    if btn_scan or "radar_cache" in st.session_state:
        if btn_scan:
            with st.spinner("جاري الاتصال بسيرفر أودو وفحص حالة جلسات نقاط البيع..."):
                try:
                    db = SessionLocal()
                    proc = SalesDataProcessor(db)
                    client = OdooClient()
                    radar_result = client.scan_morning_radar(target_date_str, allow_secondary_monthly=allow_secondary)
                    st.session_state.radar_cache = radar_result
                    st.session_state.active_date_str = target_date_str
                except Exception as e:
                    st.error(f"❌ خطأ أثناء الاتصال بأودو: {e}")
                finally:
                    db.close()

        radar_data = st.session_state.get("radar_cache")
        if radar_data:
            stats = radar_data["stats"]

            # KPI Summary Row
            r_col1, r_col2, r_col3, r_col4 = st.columns(4)
            with r_col1:
                st.metric("إجمالي نقاط البيع النشطة", stats["total_pos_scanned"])
            with r_col2:
                st.metric("الجلسات المقفلة الجاهزة", stats["closed_ready_count"])
            with r_col3:
                st.metric("جلسات فرعية مسموحة", stats["allowed_secondary_count"])
            with r_col4:
                st.metric("معارض متأخرة في الإقفال", stats["late_count"], delta=f"-{stats['late_count']}" if stats["late_count"] > 0 else "جاهز", delta_color="inverse")

            # Readiness Banner
            if radar_data["ready_for_daily_report"]:
                st.markdown(
                    """
                    <div class="alert-box alert-green">
                        ✅ <b>كافة المعارض مقفلة وجاهزة لإصدار التقرير اليومي!</b><br/>
                        يمكنك الانتقال فوراً إلى تبويب <b>(جلب بيانات المبيعات والمعاينة)</b> لتحميل البيانات والطباعة.
                    </div>
                    """,
                    unsafe_allow_html=True
                )
            else:
                st.markdown(
                    f"""
                    <div class="alert-box alert-red">
                        ⚠️ <b>تنبيه: يوجد عدد ({stats['late_count']}) معرضاً متأخراً في الإقفال لم تقفل جلساتهم حتى الآن:</b><br/>
                        يُرجى التواصل مع مدراء المعارض أدناه لإقفال الصندوق، أو المتابعة إذا كانت هناك استثناءات.
                    </div>
                    """,
                    unsafe_allow_html=True
                )

            # Details
            if radar_data["late_showrooms"]:
                st.markdown("#### ⏳ قائمة المعارض المتأخرة:")
                late_df = pd.DataFrame(radar_data["late_showrooms"])[["pos_name", "session_name", "opened_by", "opened_at", "status_text"]]
                late_df.columns = ["نقطة البيع", "رقم الجلسة", "الكاشير / المدير", "تاريخ ووقت الفتح", "الحالة"]
                st.dataframe(late_df, use_container_width=True)

            if radar_data["allowed_secondary_sessions"]:
                with st.expander(f"ℹ️ الجلسات الفرعية الشهرية المسموح ببقائها ({len(radar_data['allowed_secondary_sessions'])} جلسة)"):
                    sec_df = pd.DataFrame(radar_data["allowed_secondary_sessions"])[["pos_name", "session_name", "opened_at", "status_text"]]
                    sec_df.columns = ["نقطة البيع", "رقم الجلسة", "تاريخ الفتح", "البيان"]
                    st.dataframe(sec_df, use_container_width=True)

# =============================================================================
# TAB 2: SALES DATA FETCHING & REVIEW
# =============================================================================
with tab_data:
    st.subheader("🛒 جلب بيانات المبيعات والمعاينة والتدقيق (Data Intake & Audit)")
    
    radar_data = st.session_state.get("radar_cache")
    if radar_data and radar_data.get("late_showrooms"):
        st.markdown(
            f"""
            <div class="alert-box alert-red" style="margin-bottom: 15px;">
                <strong>⚠️ تنبيه تحذيري من الرادار الصباحي:</strong> هناك {len(radar_data['late_showrooms'])} معارض لم تغلق جلساتها الرئيسية، وقد لا تظهر مبيعاتها في التقرير إلا بعد إقفالها:<br/>
                {'<br/>'.join(f"• <b>{x['pos_name']}</b>: {x['status_text']}" for x in radar_data['late_showrooms'])}
            </div>
            """,
            unsafe_allow_html=True
        )

    # Fetch Controls
    c_btn1, c_btn2, c_date = st.columns([2, 2, 2])
    with c_date:
        fetch_date = st.date_input("تاريخ المبيعات المطلوبة:", value=date.today())
        fetch_date_str = fetch_date.strftime("%Y-%m-%d")

    db = SessionLocal()
    processor = SalesDataProcessor(db)

    with c_btn1:
        st.write("")
        st.write("")
        if st.button("⚡ جلب مباشر وسريع من أودو (Odoo XML-RPC)", type="primary", use_container_width=True):
            with st.spinner("جاري جلب إقفالات نقاط البيع من سيرفر أودو مباشرة..."):
                try:
                    client = OdooClient()
                    records = client.fetch_daily_pos_records(fetch_date_str, processor.resolver)
                    st.session_state.raw_records = records
                    st.session_state.active_date_str = fetch_date_str
                    st.session_state.pdf_cache = None
                    st.success(f"✅ تم بنجاح جلب {len(records)} جلسة من أودو.")
                except Exception as e:
                    st.error(f"❌ خطأ أثناء الاتصال بأودو: {e}")

    with c_btn2:
        st.write("")
        st.write("")
        excel_file = st.file_uploader("أو رفع ملف إكسيل أودو (xlsx)", type=["xlsx", "xls"], label_visibility="collapsed")
        if excel_file is not None and st.button("📥 قراءة ملف الإكسيل المرفوع", use_container_width=True):
            with st.spinner("جاري قراءة وتدقيق ملف الإكسيل المرفوع..."):
                try:
                    records = processor.parse_odoo_excel_file(excel_file)
                    st.session_state.raw_records = records
                    st.session_state.active_date_str = fetch_date_str
                    st.session_state.pdf_cache = None
                    st.success(f"✅ تم تحميل وتدقيق {len(records)} جلسة من ملف الإكسيل.")
                except Exception as e:
                    st.error(f"❌ خطأ في معالجة الملف: {e}")

    # Display and Edit Records Table
    records = st.session_state.get("raw_records", [])
    if records:
        st.markdown("---")
        st.markdown(f"#### 📋 معاينة وتدقيق الجلسات المجلوبة ليوم ({st.session_state.active_date_str}):")

        # Smart Audit Notice
        has_adj = any(abs(r.get("acc_adjustment", 0.0)) > 0.001 for r in records)
        if has_adj:
            adj_sessions = [f"{r['closing_no']} ({r['showroom']}): {r['acc_adjustment']:+,.2f} ر.س" for r in records if abs(r.get("acc_adjustment", 0.0)) > 0.001]
            st.markdown(
                f"""
                <div style="background-color: #FEF3C7; border: 1px solid #F59E0B; color: #92400E; padding: 12px 16px; border-radius: 8px; margin-bottom: 12px; font-size: 10.5pt;">
                    <strong>🔍 تدقيق قيود التسوية المحاسبية:</strong> تم رصد فروقات محاسبية لاحقة بعد إقفال مدراء المعارض، وتم توثيقها تلقائياً في عمود <b>(تعديل محاسبي)</b>:<br/>
                    {' • '.join(adj_sessions)}
                </div>
                """,
                unsafe_allow_html=True
            )

        df_edit = pd.DataFrame(records)

        # Reorder and format columns
        display_cols = [
            "closing_no", "showroom", "date", "closing_hour",
            "net_total", "returns", "tax", "net_sales",
            "open", "credit", "cash", "commissions", "expenses", "diff", "acc_adjustment"
        ]

        column_config = {
            "closing_no": st.column_config.TextColumn("رقم الإقفال", disabled=True),
            "showroom": st.column_config.TextColumn("المعرض", disabled=True),
            "date": st.column_config.TextColumn("التاريخ", disabled=True),
            "closing_hour": st.column_config.NumberColumn("ساعة الإقفال", disabled=True),
            "net_total": st.column_config.NumberColumn("الإجمالي", format="%.2f"),
            "returns": st.column_config.NumberColumn("ج. المردودات", format="%.2f"),
            "tax": st.column_config.NumberColumn("الضريبة", format="%.2f"),
            "net_sales": st.column_config.NumberColumn("ص. الدخل", format="%.2f"),
            "open": st.column_config.NumberColumn("الافتتاحي", format="%.2f"),
            "credit": st.column_config.NumberColumn("الشبكة", format="%.2f"),
            "cash": st.column_config.NumberColumn("النقد", format="%.2f"),
            "commissions": st.column_config.NumberColumn("عمولات", format="%.0f"),
            "expenses": st.column_config.NumberColumn("مصاريف", format="%.2f"),
            "diff": st.column_config.NumberColumn("الفارق", format="%.2f"),
            "acc_adjustment": st.column_config.NumberColumn("تعديل محاسبي", format="%.2f"),
        }

        edited_df = st.data_editor(
            df_edit[display_cols],
            column_config=column_config,
            use_container_width=True,
            num_rows="fixed",
            key="sales_editor"
        )

        # Update back to session state
        updated_records = edited_df.to_dict(orient="records")
        for i, up_r in enumerate(updated_records):
            for orig_k, orig_v in records[i].items():
                if orig_k not in up_r:
                    up_r[orig_k] = orig_v
            up_r["total"] = up_r["net_total"] + up_r["returns"]

        summary = processor.process_selected_records(updated_records)
        st.session_state.processed_summary = summary
        gt = summary["grand_totals"]

        # Real-time Executive Summary Ribbon
        st.markdown("#### 📈 إجماليات الدفعة الحالية المعتمدة:")
        kpi_col1, kpi_col2, kpi_col3, kpi_col4, kpi_col5, kpi_col6 = st.columns(6)
        with kpi_col1:
            st.metric("عدد المعارض", f"{gt['unique_showrooms_count']}")
        with kpi_col2:
            st.metric("المبيعات (ريال)", f"{gt['total']:,.2f}")
        with kpi_col3:
            st.metric("ص. الدخل (ريال)", f"{gt['net_sales']:,.2f}")
        with kpi_col4:
            st.metric("الشبكة (ريال)", f"{gt['credit']:,.2f}")
        with kpi_col5:
            st.metric("النقد (ريال)", f"{gt['cash']:,.2f}")
        with kpi_col6:
            st.metric("التحويل (ريال)", f"{gt['trans']:,.2f}")

        with st.expander("📊 التحليل البصري لطرق التحصيل اليومية (شبكة vs نقد vs تحويل)"):
            fig_daily_pay = charts.build_daily_payment_donut(gt)
            st.plotly_chart(fig_daily_pay, use_container_width=True, config=PLOTLY_CONFIG)
    else:
        st.info("👈 اضغط على أحد أزرار الجلب أعلاه لبدء مراجعة واعتماد المبيعات.")
    db.close()

# =============================================================================
# TAB 3: PDF REPORTS GENERATION & COMMIT
# =============================================================================
with tab_print:
    st.subheader("📄 محرك توليد تقارير الـ PDF الثلاثة والاعتماد (Exact Access Replicas)")
    summary = st.session_state.processed_summary

    if not summary or not summary.get("detailed_rows"):
        st.warning("⚠️ يرجى أولاً جلب ومراجعة بيانات المبيعات في التبويب السابق قبل توليد التقارير.")
    else:
        detailed_rows = summary["detailed_rows"]
        totals_rows = summary["totals_rows"]
        grand_totals = summary["grand_totals"]

        col_date, col_gen = st.columns([1.5, 2])
        with col_date:
            rep_date_input = st.date_input("تاريخ التقرير المطبوع:", value=date.today())
            report_date_str = rep_date_input.strftime("%d/%m/%Y")
            file_date_prefix = rep_date_input.strftime("%Y-%m-%d")

        with col_gen:
            st.write("")
            st.write("")
            gen_pdfs = st.button("🖨️ توليد ملفات الـ PDF الثلاثة الآن", type="primary", use_container_width=True)

        if gen_pdfs or "pdf_cache" in st.session_state:
            if gen_pdfs:
                with st.spinner("جاري تصميم وتوليد ملفات الـ PDF الثلاثة بدقة البكسل والسرعة الفائقة..."):
                    pdf_generator = PDFReportGenerator()
                    pdf_dict = pdf_generator.generate_all_three_pdfs(detailed_rows, totals_rows, grand_totals, report_date_str)
                    detailed_bytes = pdf_dict["detailed"]
                    totals_bytes = pdf_dict["totals"]
                    color_bytes = pdf_dict["color"]

                    st.session_state.pdf_cache = {
                        "detailed": detailed_bytes,
                        "totals": totals_bytes,
                        "color": color_bytes,
                        "date_prefix": file_date_prefix
                    }
                    st.success("✅ تم توليد ملفات الـ PDF الثلاثة بنجاح وجاهزة للتنزيل والطباعة!")

            pdf_cache = st.session_state.get("pdf_cache")
            if pdf_cache:
                d_prefix = pdf_cache["date_prefix"]
                st.markdown("---")
                st.markdown("#### 📥 تنزيل وطباعة التقارير الرسمية المعتمدة:")

                col_dl1, col_dl2, col_dl3 = st.columns(3)
                with col_dl1:
                    st.download_button(
                        label=f"📄 تحميل التقرير التفصيلي ({d_prefix}rep.pdf)",
                        data=pdf_cache["detailed"],
                        file_name=f"{d_prefix}rep.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                        type="primary"
                    )
                with col_dl2:
                    st.download_button(
                        label=f"📑 تحميل التقرير الإجمالي ({d_prefix}tot.pdf)",
                        data=pdf_cache["totals"],
                        file_name=f"{d_prefix}tot.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
                with col_dl3:
                    st.download_button(
                        label=f"🎨 تحميل التقرير الرقابي الملون ({d_prefix}color.pdf)",
                        data=pdf_cache["color"],
                        file_name=f"{d_prefix}color.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )

                # COMMIT BATCH SECTION
                st.markdown("---")
                st.markdown("#### 🔒 اعتماد الدفعة وترحيل الجلسات إلى الأرشيف السحابي (Commit to tblOLD):")
                st.write(
                    "عند الضغط على اعتماد، سيتم حفظ نسخة رقمية من ملفات الـ PDF الثلاثة في الأرشيف السحابي، "
                    "وترحيل أرقام الجلسات إلى جدول `tbl_old` لمنع تكرار طباعتها في الأيام القادمة."
                )

                confirm_commit = st.checkbox("أقر بصحة البيانات ومطابقة المبالغ وأرغب في اعتماد التقرير رسمياً")
                if st.button("💾 اعتماد الدفعة وترحيل الجلسات رسمياً", type="primary", disabled=not confirm_commit):
                    with st.spinner("جاري الترحيل والأرشفة..."):
                        db = SessionLocal()
                        batch_id = f"BATCH-{d_prefix}-{int(time.time())}"
                        try:
                            new_batch = ReportBatch(
                                batch_id=batch_id,
                                report_date=d_prefix,
                                sessions_count=grand_totals["sessions_count"],
                                unique_showrooms_count=grand_totals["unique_showrooms_count"],
                                total_sales=grand_totals["total"],
                                net_sales=grand_totals["net_sales"],
                                detailed_pdf_blob=pdf_cache["detailed"],
                                totals_pdf_blob=pdf_cache["totals"],
                                color_pdf_blob=pdf_cache["color"],
                                notes=f"تم الاعتماد بواسطة المدير بتاريخ {datetime.now().strftime('%Y-%m-%d %H:%M')}"
                            )
                            db.add(new_batch)

                            for row in detailed_rows:
                                c_no = row["closing_no"].strip()
                                existing = db.query(OldClosing).filter(OldClosing.closing_no == c_no).first()
                                if not existing:
                                    db.add(OldClosing(
                                        closing_no=c_no,
                                        showroom_short=row["showroom"],
                                        closing_date=row["date"],
                                        net_total=row["net_total"],
                                        batch_id=batch_id
                                    ))
                            db.commit()
                            get_cached_system_stats.clear()

                            st.success(f"🎉 تم اعتماد الدفعة رقم `{batch_id}` وترحيل {len(detailed_rows)} جلسة إلى جدول tblOLD بنجاح!")
                            st.session_state.raw_records = []
                            st.session_state.processed_summary = None
                            st.session_state.pdf_cache = None
                            time.sleep(2)
                            st.rerun()
                        except Exception as e:
                            db.rollback()
                            st.error(f"❌ فشل الاعتماد: {e}")
                        finally:
                            db.close()

# =============================================================================
# TAB 4: PERIODIC & COMPARATIVE EXECUTIVE SALES REPORTS
# =============================================================================
with tab_monthly:
    st.subheader("📊 مركز التقارير المالية الدورية والتحليلية المقارنة")
    st.info(
        "يتيح لك هذا المركز استخراج التقارير المالية التنفيذية لكافة الفترات (شهري، ربع سنوي، نصف سنوي، ثلاثة أرباع السنة، سنوي)، "
        "بالإضافة إلى تقارير المقارنات التحليلية وحساب نسب النمو والانخفاض مباشرة من قيود أودو مع التصدير الفوري لإكسيل وPDF رسمي."
    )

    sub_periodic, sub_comparative = st.tabs([
        "📅 تقارير الفترات الزمنية (Periodic Sales Reports)",
        "📈 تقارير المقارنات ونسب النمو (Comparative & Growth Analytics)"
    ])

    # -------------------------------------------------------------------------
    # SUB-TAB 1: PERIODIC SALES REPORTS
    # -------------------------------------------------------------------------
    with sub_periodic:
        st.markdown("#### 📅 استخراج تقارير المبيعات الدورية:")
        col_type, col_p1, col_p2, col_p_btn = st.columns([1.6, 1.2, 1.2, 1.5])

        with col_type:
            p_kind = st.selectbox(
                "نوع التقرير الدوري:",
                [
                    "شهري (Monthly)",
                    "ربع سنوي (Quarterly)",
                    "نصف سنوي (Semi-Annual)",
                    "ثلاثة أرباع السنة (9-Month)",
                    "سنوي (Annual)",
                    "فترة مخصصة (Custom Range)"
                ],
                key="per_kind"
            )

        now_dt = datetime.now()
        default_month = now_dt.month - 1 if now_dt.month > 1 else 12
        default_year = now_dt.year if now_dt.month > 1 else now_dt.year - 1

        p_start_date = ""
        p_end_date = ""
        p_title = ""
        p_label = ""
        p_excel_name = ""
        p_pdf_name = ""

        if p_kind == "شهري (Monthly)":
            with col_p1:
                p_m = st.selectbox("الشهر:", list(range(1, 13)), index=default_month - 1, format_func=lambda m: f"شهر {m} ({m:02d})", key="per_m")
            with col_p2:
                p_y = st.selectbox("السنة:", [2024, 2025, 2026, 2027], index=2, key="per_y")
            _, last_day = monthrange(p_y, p_m)
            p_start_date = f"{p_y}-{p_m:02d}-01"
            p_end_date = f"{p_y}-{p_m:02d}-{last_day:02d}"
            p_title = f"مبيعات معارض العسلي لشهر {p_m} - {p_y}"
            p_label = f"شهر {p_m}-{p_y}"
            p_excel_name = f"تقرير شهر {p_m}-{p_y}.xlsx"
            p_pdf_name = f"تقرير شهر {p_m}-{p_y}.pdf"

        elif p_kind == "ربع سنوي (Quarterly)":
            with col_p1:
                q_opts = {
                    "الربع الأول (Q1: يناير - مارس)": (1, "01-01", "03-31", "الأول"),
                    "الربع الثاني (Q2: أبريل - يونيو)": (2, "04-01", "06-30", "الثاني"),
                    "الربع الثالث (Q3: يوليو - سبتمبر)": (3, "07-01", "09-30", "الثالث"),
                    "الربع الرابع (Q4: أكتوبر - ديسمبر)": (4, "10-01", "12-31", "الرابع")
                }
                q_selected = st.selectbox("الربع السنوي:", list(q_opts.keys()), index=2, key="per_q")
                _, q_s, q_e, q_ar = q_opts[q_selected]
            with col_p2:
                p_y = st.selectbox("السنة:", [2024, 2025, 2026, 2027], index=2, key="per_q_y")
            p_start_date = f"{p_y}-{q_s}"
            p_end_date = f"{p_y}-{q_e}"
            p_title = f"مبيعات معارض العسلي للربع {q_ar} - {p_y}"
            p_label = f"الربع {q_ar} {p_y}"
            p_excel_name = f"تقرير الربع {q_ar} {p_y}.xlsx"
            p_pdf_name = f"تقرير الربع {q_ar} {p_y}.pdf"

        elif p_kind == "نصف سنوي (Semi-Annual)":
            with col_p1:
                h_opts = {
                    "النصف الأول (H1: يناير - يونيو)": ("01-01", "06-30", "الأول"),
                    "النصف الثاني (H2: يوليو - ديسمبر)": ("07-01", "12-31", "الثاني")
                }
                h_selected = st.selectbox("النصف السنوي:", list(h_opts.keys()), index=0, key="per_h")
                h_s, h_e, h_ar = h_opts[h_selected]
            with col_p2:
                p_y = st.selectbox("السنة:", [2024, 2025, 2026, 2027], index=2, key="per_h_y")
            p_start_date = f"{p_y}-{h_s}"
            p_end_date = f"{p_y}-{h_e}"
            p_title = f"مبيعات معارض العسلي للنصف {h_ar} - {p_y}"
            p_label = f"النصف {h_ar} {p_y}"
            p_excel_name = f"تقرير النصف {h_ar} {p_y}.xlsx"
            p_pdf_name = f"تقرير النصف {h_ar} {p_y}.pdf"

        elif p_kind == "ثلاثة أرباع السنة (9-Month)":
            with col_p1:
                st.caption("الفترة: 1 يناير وحتى 30 سبتمبر")
            with col_p2:
                p_y = st.selectbox("السنة:", [2024, 2025, 2026, 2027], index=2, key="per_9m_y")
            p_start_date = f"{p_y}-01-01"
            p_end_date = f"{p_y}-09-30"
            p_title = f"مبيعات معارض العسلي لفترة ثلاثة أرباع السنة (9M) - {p_y}"
            p_label = f"ثلاثة أرباع السنة {p_y}"
            p_excel_name = f"تقرير ثلاثة أرباع السنة {p_y}.xlsx"
            p_pdf_name = f"تقرير ثلاثة أرباع السنة {p_y}.pdf"

        elif p_kind == "سنوي (Annual)":
            with col_p1:
                st.caption("الفترة: كامل السنة المالية (12 شهراً)")
            with col_p2:
                p_y = st.selectbox("السنة المالية:", [2024, 2025, 2026, 2027], index=2, key="per_fy_y")
            p_start_date = f"{p_y}-01-01"
            p_end_date = f"{p_y}-12-31"
            p_title = f"مبيعات معارض العسلي السنوي - {p_y}"
            p_label = f"سنوي {p_y}"
            p_excel_name = f"تقرير سنوي {p_y}.xlsx"
            p_pdf_name = f"تقرير سنوي {p_y}.pdf"

        elif p_kind == "فترة مخصصة (Custom Range)":
            with col_p1:
                c_start = st.date_input("من تاريخ:", value=date(now_dt.year, 1, 1), key="per_cust_s")
            with col_p2:
                c_end = st.date_input("إلى تاريخ:", value=date.today(), key="per_cust_e")
            p_start_date = c_start.strftime("%Y-%m-%d")
            p_end_date = c_end.strftime("%Y-%m-%d")
            p_title = f"مبيعات معارض العسلي للفترة من {p_start_date} إلى {p_end_date}"
            p_label = f"{p_start_date} إلى {p_end_date}"
            p_excel_name = f"تقرير مبيعات المعارض {p_start_date} إلى {p_end_date}.xlsx"
            p_pdf_name = f"تقرير مبيعات المعارض {p_start_date} إلى {p_end_date}.pdf"

        with col_p_btn:
            st.write("")
            st.write("")
            btn_fetch_periodic = st.button("🔍 استخراج التقرير من أودو الآن", type="primary", use_container_width=True, key="btn_fetch_per")

        if btn_fetch_periodic:
            with st.spinner(f"جاري سحب وتجميع المبيعات للفترة ({p_label}) من أودو..."):
                try:
                    db = SessionLocal()
                    proc = SalesDataProcessor(db)
                    client = OdooClient()
                    res_p = client.fetch_period_sales(
                        start_date=p_start_date,
                        end_date=p_end_date,
                        title=p_title,
                        period_type=p_kind,
                        showroom_resolver=proc.resolver
                    )
                    res_p["period_label"] = p_label
                    res_p["excel_filename"] = p_excel_name
                    res_p["pdf_filename"] = p_pdf_name
                    st.session_state.periodic_data = res_p
                    st.success(f"✅ تم سحب مبيعات ({p_label}) بنجاح ({res_p['grand_totals']['showrooms_count']} معرضاً)!")
                except Exception as e:
                    st.error(f"❌ خطأ أثناء استخراج التقرير الدوري: {e}")
                finally:
                    db.close()

        # Render Periodic Data
        p_data = st.session_state.get("periodic_data")
        if p_data:
            st.markdown("---")
            st.markdown(
                f"""
                <div style="background: linear-gradient(135deg, #1F4E79 0%, #2F5597 100%); color: white; padding: 18px 24px; border-radius: 10px; margin-bottom: 20px; text-align: center;">
                    <h2 style="margin: 0; color: white; font-weight: 800;">{p_data['title']}</h2>
                    <p style="margin: 6px 0 0 0; opacity: 0.9; font-size: 11pt;">شركة العسلي للتجارة — تقرير الأداء المالي للمعارض مرتباً حسب الصافي من الأعلى إلى الأقل</p>
                </div>
                """,
                unsafe_allow_html=True
            )

            p_gt = p_data["grand_totals"]
            p_rows = p_data["rows"]
            p_top = p_rows[0] if p_rows else None

            # KPI Ribbon
            pk1, pk2, pk3, pk4 = st.columns(4)
            with pk1:
                st.metric("إجمالي المبيعات (شامل الضريبة)", f"{p_gt['total']:,.2f} ر.س")
            with pk2:
                st.metric("إجمالي ضريبة القيمة المضافة", f"{p_gt['tax']:,.2f} ر.س")
            with pk3:
                st.metric("صافي المبيعات (بدون الضريبة)", f"{p_gt['net']:,.2f} ر.س", delta="الصافي")
            with pk4:
                top_name = p_top["showroom"] if p_top else "-"
                top_val = f"{p_top['net']:,.2f} ر.س ({(p_top['net']/p_gt['net']*100):.1f}%)" if p_top and p_gt['net'] > 0 else "-"
                st.metric("المعرض الأول (Top 1)", top_name, delta=top_val)

            # Downloads
            per_ex_name = p_data.get("excel_filename", "تقرير_المبيعات.xlsx")
            per_pd_name = p_data.get("pdf_filename", "تقرير_المبيعات.pdf")

            per_excel_bytes = generate_periodic_excel(p_data)
            pdf_gen = PDFReportGenerator()
            per_pdf_bytes = pdf_gen.generate_periodic_pdf(p_data)

            st.markdown("##### 📥 خيارات التصدير والطباعة الرسمية:")
            col_pdl1, col_pdl2 = st.columns(2)
            with col_pdl1:
                st.download_button(
                    label=f"📥 تحميل ملف الإكسيل ({per_ex_name})",
                    data=per_excel_bytes,
                    file_name=per_ex_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary",
                    key="dl_per_excel"
                )
            with col_pdl2:
                st.download_button(
                    label=f"🖨️ طباعة وتحميل التقرير الرسمي ({per_pd_name})",
                    data=per_pdf_bytes,
                    file_name=per_pd_name,
                    mime="application/pdf",
                    use_container_width=True,
                    key="dl_per_pdf"
                )

            # Table
            st.markdown("#### 📋 جدول مبيعات المعارض للفترة المحددة:")
            p_table_rows = []
            for r in p_rows:
                p_table_rows.append({
                    "م": r["seq"],
                    "المعرض": r["showroom"],
                    "الإجمالي": r["total"],
                    "الضريبة": r["tax"],
                    "الصافي": r["net"],
                })

            p_table_rows.append({
                "م": "المجموع",
                "المعرض": f"عدد المعارض: {p_gt['showrooms_count']}",
                "الإجمالي": p_gt["total"],
                "الضريبة": p_gt["tax"],
                "الصافي": p_gt["net"],
            })

            df_p = pd.DataFrame(p_table_rows)
            p_col_config = {
                "م": st.column_config.TextColumn("م", width="small"),
                "المعرض": st.column_config.TextColumn("المعرض", width="large"),
                "الإجمالي": st.column_config.NumberColumn("الإجمالي", format="%.2f ر.س"),
                "الضريبة": st.column_config.NumberColumn("الضريبة", format="%.2f ر.س"),
                "الصافي": st.column_config.NumberColumn("الصافي", format="%.2f ر.س"),
            }
            st.dataframe(df_p, column_config=p_col_config, use_container_width=True, hide_index=True)

            # Interactive Multi-Pattern Charts Suite
            with st.expander("📊 التحليل البصري التفاعلي للمبيعات (أنماط ومخططات بيانية متعددة)", expanded=True):
                col_ch1, col_ch2 = st.columns([2.5, 1])
                with col_ch1:
                    chart_mode_p = st.radio(
                        "اختر نمط الرسم البياني المطلوب:",
                        [
                            "📊 أعمدة أفقية مرتبة (Ranked Bars)",
                            "🍩 الحصص السوقية (Market Share Donut)",
                            "🗂️ خريطة الشجرة (Treemap)",
                            "📈 منحنى باريتو التراكمي (Pareto 80/20)",
                            "🌊 منحنى التوزيع المالي (Area Curve)"
                        ],
                        horizontal=True,
                        key="chart_mode_per"
                    )
                with col_ch2:
                    if chart_mode_p in ["📊 أعمدة أفقية مرتبة (Ranked Bars)", "🌊 منحنى التوزيع المالي (Area Curve)"]:
                        top_n_p = st.slider("عدد المعارض المعروضة:", min_value=5, max_value=min(40, len(p_rows)), value=min(15, len(p_rows)), key="slider_per_bar")
                    else:
                        top_n_p = len(p_rows)

                if chart_mode_p == "📊 أعمدة أفقية مرتبة (Ranked Bars)":
                    fig_p = charts.build_periodic_horizontal_bar(p_rows, top_n=top_n_p, title=f"مبيعات أعلى {top_n_p} معرضاً للفترة (صافي بالريال)")
                    st.plotly_chart(fig_p, use_container_width=True, config=PLOTLY_CONFIG)

                elif chart_mode_p == "🍩 الحصص السوقية (Market Share Donut)":
                    fig_p = charts.build_periodic_donut_chart(p_rows, grand_total_net=p_gt["net"], title=f"توزيع الحصص السوقية لصافي المبيعات — {p_data['title']}")
                    st.plotly_chart(fig_p, use_container_width=True, config=PLOTLY_CONFIG)

                elif chart_mode_p == "🗂️ خريطة الشجرة (Treemap)":
                    fig_p = charts.build_periodic_treemap(p_rows, title=f"الوزن النسبي لمبيعات المعارض (Treemap) — {p_data['title']}")
                    st.plotly_chart(fig_p, use_container_width=True, config=PLOTLY_CONFIG)

                elif chart_mode_p == "📈 منحنى باريتو التراكمي (Pareto 80/20)":
                    st.caption("💡 **قاعدة باريتو (80/20):** يوضح الخط الأحمر التراكمي المعارض الرئيسية المسؤولة عن توليد 80% من إجمالي مبيعات الشركة.")
                    fig_p = charts.build_periodic_pareto(p_rows, title=f"تحليل باريتو التراكمي لمبيعات المعارض — {p_data['title']}")
                    st.plotly_chart(fig_p, use_container_width=True, config=PLOTLY_CONFIG)

                elif chart_mode_p == "🌊 منحنى التوزيع المالي (Area Curve)":
                    fig_p = charts.build_periodic_area_curve(p_rows[:top_n_p], title=f"منحنى انحدار حجم المبيعات لأعلى {top_n_p} معرضاً")
                    st.plotly_chart(fig_p, use_container_width=True, config=PLOTLY_CONFIG)

    # -------------------------------------------------------------------------
    # SUB-TAB 2: COMPARATIVE & GROWTH ANALYTICS REPORTS
    # -------------------------------------------------------------------------
    with sub_comparative:
        st.markdown("#### 📈 استخراج تقارير المقارنات ونسب النمو والانخفاض:")
        col_c_mode, col_c1, col_c2, col_c_btn = st.columns([1.8, 1.2, 1.2, 1.5])

        with col_c_mode:
            c_mode = st.selectbox(
                "نوع المقارنة التحليلية:",
                [
                    "مقارنة شهر مع الشهر السابق (Month-over-Month)",
                    "مقارنة شهر مع نفس الشهر من العام السابق (Year-over-Year MoM)",
                    "مقارنة ربع سنوي مع الربع السابق (Quarter-over-Quarter)",
                    "مقارنة ربع سنوي مع نفس الربع من العام السابق (Year-over-Year QoQ)",
                    "مقارنة سنوية (Year-over-Year Annual)",
                    "مقارنة فترتين مخصصتين (Custom Comparison)"
                ],
                key="cmp_mode"
            )

        c_curr_period = {}
        c_prev_period = {}
        c_title = ""
        c_excel_name = ""
        c_pdf_name = ""

        if c_mode == "مقارنة شهر مع الشهر السابق (Month-over-Month)":
            with col_c1:
                cm_m = st.selectbox("الشهر الحالي:", list(range(1, 13)), index=default_month - 1, format_func=lambda m: f"شهر {m} ({m:02d})", key="cmp_mom_m")
            with col_c2:
                cm_y = st.selectbox("السنة:", [2024, 2025, 2026, 2027], index=2, key="cmp_mom_y")

            _, d1 = monthrange(cm_y, cm_m)
            c_curr_period = {'start_date': f"{cm_y}-{cm_m:02d}-01", 'end_date': f"{cm_y}-{cm_m:02d}-{d1:02d}", 'label': f"شهر {cm_m}-{cm_y}"}
            prev_m = 12 if cm_m == 1 else cm_m - 1
            prev_y = cm_y - 1 if cm_m == 1 else cm_y
            _, d2 = monthrange(prev_y, prev_m)
            c_prev_period = {'start_date': f"{prev_y}-{prev_m:02d}-01", 'end_date': f"{prev_y}-{prev_m:02d}-{d2:02d}", 'label': f"شهر {prev_m}-{prev_y}"}
            c_title = f"مقارنة مبيعات شهر {cm_m}-{cm_y} مع الشهر السابق ({prev_m}-{prev_y})"
            c_excel_name = f"تقرير مقارنة شهر {cm_m}-{cm_y} مع الشهر السابق.xlsx"
            c_pdf_name = f"تقرير مقارنة شهر {cm_m}-{cm_y} مع الشهر السابق.pdf"

        elif c_mode == "مقارنة شهر مع نفس الشهر من العام السابق (Year-over-Year MoM)":
            with col_c1:
                cm_m = st.selectbox("الشهر الحالي:", list(range(1, 13)), index=default_month - 1, format_func=lambda m: f"شهر {m} ({m:02d})", key="cmp_yoy_m")
            with col_c2:
                cm_y = st.selectbox("السنة الحالية:", [2025, 2026, 2027], index=1, key="cmp_yoy_y")

            _, d1 = monthrange(cm_y, cm_m)
            c_curr_period = {'start_date': f"{cm_y}-{cm_m:02d}-01", 'end_date': f"{cm_y}-{cm_m:02d}-{d1:02d}", 'label': f"شهر {cm_m}-{cm_y}"}
            prev_y = cm_y - 1
            _, d2 = monthrange(prev_y, cm_m)
            c_prev_period = {'start_date': f"{prev_y}-{cm_m:02d}-01", 'end_date': f"{prev_y}-{cm_m:02d}-{d2:02d}", 'label': f"شهر {cm_m}-{prev_y}"}
            c_title = f"مقارنة مبيعات شهر {cm_m}-{cm_y} مع نفس الشهر من العام السابق ({cm_m}-{prev_y})"
            c_excel_name = f"تقرير مقارنة شهر {cm_m}-{cm_y} مع نفس الشهر {prev_y}.xlsx"
            c_pdf_name = f"تقرير مقارنة شهر {cm_m}-{cm_y} مع نفس الشهر {prev_y}.pdf"

        elif c_mode == "مقارنة ربع سنوي مع الربع السابق (Quarter-over-Quarter)":
            q_map = {
                1: ("01-01", "03-31", "الربع الأول"),
                2: ("04-01", "06-30", "الربع الثاني"),
                3: ("07-01", "09-30", "الربع الثالث"),
                4: ("10-01", "12-31", "الربع الرابع")
            }
            with col_c1:
                cq_num = st.selectbox("الربع الحالي:", [1, 2, 3, 4], index=2, format_func=lambda q: q_map[q][2], key="cmp_qoq_q")
            with col_c2:
                cq_y = st.selectbox("السنة:", [2024, 2025, 2026, 2027], index=2, key="cmp_qoq_y")

            s1, e1, l1 = q_map[cq_num]
            c_curr_period = {'start_date': f"{cq_y}-{s1}", 'end_date': f"{cq_y}-{e1}", 'label': f"{l1} {cq_y}"}
            prev_q = 4 if cq_num == 1 else cq_num - 1
            prev_y = cq_y - 1 if cq_num == 1 else cq_y
            s2, e2, l2 = q_map[prev_q]
            c_prev_period = {'start_date': f"{prev_y}-{s2}", 'end_date': f"{prev_y}-{e2}", 'label': f"{l2} {prev_y}"}
            c_title = f"مقارنة مبيعات {l1} {cq_y} مع {l2} {prev_y}"
            c_excel_name = f"تقرير مقارنة {l1} {cq_y} مع {l2} {prev_y}.xlsx"
            c_pdf_name = f"تقرير مقارنة {l1} {cq_y} مع {l2} {prev_y}.pdf"

        elif c_mode == "مقارنة ربع سنوي مع نفس الربع من العام السابق (Year-over-Year QoQ)":
            q_map = {
                1: ("01-01", "03-31", "الربع الأول"),
                2: ("04-01", "06-30", "الربع الثاني"),
                3: ("07-01", "09-30", "الربع الثالث"),
                4: ("10-01", "12-31", "الربع الرابع")
            }
            with col_c1:
                cq_num = st.selectbox("الربع الحالي:", [1, 2, 3, 4], index=2, format_func=lambda q: q_map[q][2], key="cmp_yoy_q")
            with col_c2:
                cq_y = st.selectbox("السنة الحالية:", [2025, 2026, 2027], index=1, key="cmp_yoy_qy")

            s1, e1, l1 = q_map[cq_num]
            prev_y = cq_y - 1
            c_curr_period = {'start_date': f"{cq_y}-{s1}", 'end_date': f"{cq_y}-{e1}", 'label': f"{l1} {cq_y}"}
            c_prev_period = {'start_date': f"{prev_y}-{s1}", 'end_date': f"{prev_y}-{e1}", 'label': f"{l1} {prev_y}"}
            c_title = f"مقارنة مبيعات {l1} {cq_y} مع نفس الربع من عام {prev_y}"
            c_excel_name = f"تقرير مقارنة {l1} {cq_y} مع {l1} {prev_y}.xlsx"
            c_pdf_name = f"تقرير مقارنة {l1} {cq_y} مع {l1} {prev_y}.pdf"

        elif c_mode == "مقارنة سنوية (Year-over-Year Annual)":
            with col_c1:
                cy_curr = st.selectbox("السنة الحالية:", [2025, 2026, 2027], index=1, key="cmp_ann_cy")
            with col_c2:
                cy_base = st.selectbox("سنة الأساس للمقارنة:", [2024, 2025, 2026], index=1, key="cmp_ann_by")

            c_curr_period = {'start_date': f"{cy_curr}-01-01", 'end_date': f"{cy_curr}-12-31", 'label': f"عام {cy_curr}"}
            c_prev_period = {'start_date': f"{cy_base}-01-01", 'end_date': f"{cy_base}-12-31", 'label': f"عام {cy_base}"}
            c_title = f"مقارنة مبيعات عام {cy_curr} مع عام {cy_base}"
            c_excel_name = f"تقرير مقارنة مبيعات عام {cy_curr} مع عام {cy_base}.xlsx"
            c_pdf_name = f"تقرير مقارنة مبيعات عام {cy_curr} مع عام {cy_base}.pdf"

        elif c_mode == "مقارنة فترتين مخصصتين (Custom Comparison)":
            with col_c1:
                cs_curr = st.date_input("بداية الفترة الحالية:", value=date(now_dt.year, 9, 1), key="cmp_cs_curr")
                ce_curr = st.date_input("نهاية الفترة الحالية:", value=date(now_dt.year, 9, 30), key="cmp_ce_curr")
            with col_c2:
                cs_prev = st.date_input("بداية فترة الأساس:", value=date(now_dt.year, 8, 1), key="cmp_cs_prev")
                ce_prev = st.date_input("نهاية فترة الأساس:", value=date(now_dt.year, 8, 31), key="cmp_ce_prev")

            c_curr_period = {'start_date': cs_curr.strftime("%Y-%m-%d"), 'end_date': ce_curr.strftime("%Y-%m-%d"), 'label': "الفترة الحالية"}
            c_prev_period = {'start_date': cs_prev.strftime("%Y-%m-%d"), 'end_date': ce_prev.strftime("%Y-%m-%d"), 'label': "فترة الأساس"}
            c_title = f"مقارنة مبيعات الفترة ({c_curr_period['start_date']}) مع ({c_prev_period['start_date']})"
            c_excel_name = "تقرير مقارنة مبيعات مخصصة.xlsx"
            c_pdf_name = "تقرير مقارنة مبيعات مخصصة.pdf"

        with col_c_btn:
            st.write("")
            st.write("")
            btn_fetch_comp = st.button("⚡ استخراج وتحليل المقارنة الآن", type="primary", use_container_width=True, key="btn_fetch_cmp")

        if btn_fetch_comp:
            with st.spinner(f"جاري سحب ومقارنة البيانات بين ({c_curr_period['label']}) و ({c_prev_period['label']}) من أودو..."):
                try:
                    db = SessionLocal()
                    proc = SalesDataProcessor(db)
                    client = OdooClient()
                    res_c = client.fetch_comparative_sales(
                        current_period=c_curr_period,
                        previous_period=c_prev_period,
                        title=c_title,
                        showroom_resolver=proc.resolver
                    )
                    res_c["excel_filename"] = c_excel_name
                    res_c["pdf_filename"] = c_pdf_name
                    st.session_state.comparative_data = res_c
                    st.success(f"✅ تم سحب وتحليل المقارنة بنجاح ({res_c['grand_totals']['showrooms_count']} معرضاً)!")
                except Exception as e:
                    st.error(f"❌ خطأ أثناء استخراج المقارنة: {e}")
                finally:
                    db.close()

        # Render Comparative Data
        c_data = st.session_state.get("comparative_data")
        if c_data:
            st.markdown("---")
            st.markdown(
                f"""
                <div style="background: linear-gradient(135deg, #1B4F72 0%, #2874A6 100%); color: white; padding: 18px 24px; border-radius: 10px; margin-bottom: 20px; text-align: center;">
                    <h2 style="margin: 0; color: white; font-weight: 800;">{c_data['title']}</h2>
                    <p style="margin: 6px 0 0 0; opacity: 0.9; font-size: 11pt;">نطاق المقارنة: {c_data['current_label']} ({c_data['current_range']}) مقابل {c_data['previous_label']} ({c_data['previous_range']})</p>
                </div>
                """,
                unsafe_allow_html=True
            )

            c_gt = c_data["grand_totals"]
            c_rows = c_data["rows"]

            # 4 Comparative KPI Cards
            ck1, ck2, ck3, ck4 = st.columns(4)
            with ck1:
                st.metric(f"صافي {c_data['current_label']}", f"{c_gt['current_net']:,.2f} ر.س")
            with ck2:
                st.metric(f"صافي {c_data['previous_label']}", f"{c_gt['previous_net']:,.2f} ر.س")
            with ck3:
                diff_sym = "+" if c_gt['diff'] > 0 else ""
                st.metric("قيمة الفارق المالي", f"{diff_sym}{c_gt['diff']:,.2f} ر.س", delta=f"{diff_sym}{c_gt['diff']:,.2f} ر.س")
            with ck4:
                growth_sym = "+" if c_gt['growth_pct'] > 0 else ""
                st.metric("معدل النمو الإجمالي للشركة", f"{growth_sym}{c_gt['growth_pct']:.2f}%", delta=f"{growth_sym}{c_gt['growth_pct']:.2f}%")

            # Downloads
            cmp_ex_name = c_data.get("excel_filename", "تقرير_المقارنة.xlsx")
            cmp_pd_name = c_data.get("pdf_filename", "تقرير_المقارنة.pdf")

            cmp_excel_bytes = generate_comparative_excel(c_data)
            pdf_gen = PDFReportGenerator()
            cmp_pdf_bytes = pdf_gen.generate_comparative_pdf(c_data)

            st.markdown("##### 📥 خيارات التصدير والطباعة الرسمية للمقارنة:")
            col_cdl1, col_cdl2 = st.columns(2)
            with col_cdl1:
                st.download_button(
                    label=f"📥 تحميل ملف إكسيل المقارنة ({cmp_ex_name})",
                    data=cmp_excel_bytes,
                    file_name=cmp_ex_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary",
                    key="dl_cmp_excel"
                )
            with col_cdl2:
                st.download_button(
                    label=f"🖨️ طباعة وتحميل تقرير المقارنة الرسمي ({cmp_pd_name})",
                    data=cmp_pdf_bytes,
                    file_name=cmp_pd_name,
                    mime="application/pdf",
                    use_container_width=True,
                    key="dl_cmp_pdf"
                )

            # Table
            st.markdown("#### 📋 جدول مقارنة مبيعات المعارض ونسب النمو:")
            c_table_rows = []
            for r in c_rows:
                c_table_rows.append({
                    "م": r["seq"],
                    "المعرض": r["showroom"],
                    f"صافي {c_data['previous_label']}": r["previous_net"],
                    f"صافي {c_data['current_label']}": r["current_net"],
                    "الفارق (ر.س)": r["diff"],
                    "نسبة النمو (%)": r["growth_pct"] / 100.0,
                })

            c_table_rows.append({
                "م": "المجموع",
                "المعرض": f"عدد المعارض: {c_gt['showrooms_count']}",
                f"صافي {c_data['previous_label']}": c_gt["previous_net"],
                f"صافي {c_data['current_label']}": c_gt["current_net"],
                "الفارق (ر.س)": c_gt["diff"],
                "نسبة النمو (%)": c_gt["growth_pct"] / 100.0,
            })

            df_c = pd.DataFrame(c_table_rows)
            c_col_config = {
                "م": st.column_config.TextColumn("م", width="small"),
                "المعرض": st.column_config.TextColumn("المعرض", width="large"),
                f"صافي {c_data['previous_label']}": st.column_config.NumberColumn(f"صافي {c_data['previous_label']}", format="%.2f ر.س"),
                f"صافي {c_data['current_label']}": st.column_config.NumberColumn(f"صافي {c_data['current_label']}", format="%.2f ر.س"),
                "الفارق (ر.س)": st.column_config.NumberColumn("الفارق (ر.س)", format="%.2f ر.س"),
                "نسبة النمو (%)": st.column_config.NumberColumn("نسبة النمو (%)", format="%.2f%%"),
            }
            st.dataframe(df_c, column_config=c_col_config, use_container_width=True, hide_index=True)

            # Comparative Multi-Pattern Charts Suite
            with st.expander("📊 التحليل البصري التفاعلي للمقارنات ونسب النمو (أنماط بيانية متعددة)", expanded=True):
                col_chc1, col_chc2 = st.columns([2.5, 1])
                with col_chc1:
                    chart_mode_c = st.radio(
                        "اختر نمط الرسم البياني المطلوب للمقارنة:",
                        [
                            "📊 أعمدة متجاورة للمقارنة المباشرة (Grouped Bars)",
                            "⚖️ شلال الفوارق المالية (Diverging Variance Green/Red)",
                            "🎯 مصفوفة النمو والأداء (Growth vs Volume Matrix)",
                            "📍 مسار قفزات المبيعات (Dumbbell Growth Chart)"
                        ],
                        horizontal=True,
                        key="chart_mode_cmp"
                    )
                with col_chc2:
                    if chart_mode_c in ["📊 أعمدة متجاورة للمقارنة المباشرة (Grouped Bars)", "📍 مسار قفزات المبيعات (Dumbbell Growth Chart)"]:
                        top_n_c = st.slider("عدد المعارض المعروضة:", min_value=5, max_value=min(40, len(c_rows)), value=min(15, len(c_rows)), key="slider_cmp_bar")
                    else:
                        top_n_c = len(c_rows)

                if chart_mode_c == "📊 أعمدة متجاورة للمقارنة المباشرة (Grouped Bars)":
                    fig_c = charts.build_comparative_grouped_bars(
                        c_rows,
                        current_label=c_data["current_label"],
                        previous_label=c_data["previous_label"],
                        top_n=top_n_c,
                        title=f"مقارنة مبيعات أعلى {top_n_c} معرضاً بين ({c_data['current_label']}) و ({c_data['previous_label']})"
                    )
                    st.plotly_chart(fig_c, use_container_width=True, config=PLOTLY_CONFIG)

                elif chart_mode_c == "⚖️ شلال الفوارق المالية (Diverging Variance Green/Red)":
                    st.caption("💡 **مخطط الفوارق المالية:** يوضح صافي الزيادة (الأخضر) أو الانخفاض (الأحمر) بالريال لكل معرض، مما يكشف المعارض الرابحة والمتراجعة فوراً.")
                    fig_c = charts.build_comparative_diverging_bars(
                        c_rows,
                        current_label=c_data["current_label"],
                        previous_label=c_data["previous_label"],
                        title=f"الفوارق المالية في صافي المبيعات (النمو والانخفاض بالريال)"
                    )
                    st.plotly_chart(fig_c, use_container_width=True, config=PLOTLY_CONFIG)

                elif chart_mode_c == "🎯 مصفوفة النمو والأداء (Growth vs Volume Matrix)":
                    st.caption("💡 **مصفوفة النمو:** تصنيف المعارض إلى 4 أرباع؛ المحور السيني يمثل حجم المبيعات الأساسي، المحور الصادي يمثل نسبة النمو %، وحجم الفقاعة يمثل المبيعات الحالية.")
                    fig_c = charts.build_comparative_growth_matrix(
                        c_rows,
                        current_label=c_data["current_label"],
                        previous_label=c_data["previous_label"],
                        title=f"مصفوفة كفاءة النمو مقابل حجم المبيعات — {c_data['title']}"
                    )
                    st.plotly_chart(fig_c, use_container_width=True, config=PLOTLY_CONFIG)

                elif chart_mode_c == "📍 مسار قفزات المبيعات (Dumbbell Growth Chart)":
                    st.caption("💡 **مخطط مسار القفزات (Dumbbell):** يربط بخط ملون بين نقطة المعرض في الفترة السابقة (دائرة رمادية) ونقطته في الفترة الحالية (ماسة ملونة).")
                    fig_c = charts.build_comparative_dumbbell(
                        c_rows,
                        current_label=c_data["current_label"],
                        previous_label=c_data["previous_label"],
                        top_n=top_n_c,
                        title=f"مسار قفزات وتغير مبيعات المعارض لأعلى {top_n_c} معرضاً"
                    )
                    st.plotly_chart(fig_c, use_container_width=True, config=PLOTLY_CONFIG)

# =============================================================================
# TAB 5: ARCHIVE, SETTINGS & ROLLBACK
# =============================================================================
with tab_archive:
    st.subheader("🗄️ سجل الدفعات السابقة، الأرشيف السحابي وإدارة النظام")

    sub_archive, sub_rollback, sub_showrooms, sub_seed, sub_security = st.tabs([
        "📁 سجل الدفعات السابقة",
        "⏮️ التراجع عن آخر اعتماد (Rollback)",
        "🏢 إدارة المعارض والمدن",
        "📥 استيراد أرشيف الجلسات (tblOLD)",
        "🛡️ أمان الأجهزة والربط السحابي"
    ])

    db = SessionLocal()

    # 1. BATCHES ARCHIVE
    with sub_archive:
        st.markdown("#### 📂 استعراض وتحميل تقارير الدفعات السابقة:")
        batches = db.query(ReportBatch).filter(ReportBatch.is_rolled_back == False).order_by(ReportBatch.created_at.desc()).all()
        if not batches:
            st.info("لا توجد دفعات معتمدة سابقة حتى الآن.")
        else:
            for b in batches:
                with st.expander(f"📅 تقرير يوم: {b.report_date} | الدفعة: `{b.batch_id}` | المبيعات: {b.total_sales:,.2f} ريال ({b.unique_showrooms_count} معرضاً)"):
                    st.write(f"• **تاريخ الاعتماد:** {b.created_at.strftime('%Y-%m-%d %H:%M:%S')}")
                    st.write(f"• **عدد الجلسات:** {b.sessions_count} | **صافي الدخل:** {b.net_sales:,.2f} ريال")
                    c1, c2, c3 = st.columns(3)
                    if b.detailed_pdf_blob:
                        with c1:
                            st.download_button(
                                label="📥 التقرير التفصيلي (rep.pdf)",
                                data=b.detailed_pdf_blob,
                                file_name=f"{b.report_date}rep.pdf",
                                mime="application/pdf",
                                key=f"dl_rep_{b.batch_id}"
                            )
                    if b.totals_pdf_blob:
                        with c2:
                            st.download_button(
                                label="📥 التقرير الإجمالي (tot.pdf)",
                                data=b.totals_pdf_blob,
                                file_name=f"{b.report_date}tot.pdf",
                                mime="application/pdf",
                                key=f"dl_tot_{b.batch_id}"
                            )
                    if b.color_pdf_blob:
                        with c3:
                            st.download_button(
                                label="📥 التقرير الرقابي (color.pdf)",
                                data=b.color_pdf_blob,
                                file_name=f"{b.report_date}color.pdf",
                                mime="application/pdf",
                                key=f"dl_col_{b.batch_id}"
                            )

    # 2. ROLLBACK LAST BATCH
    with sub_rollback:
        st.markdown("#### ⚠️ التراجع عن آخر دفعة معتمدة (Rollback Last Batch):")
        st.warning(
            "تتيح لك هذه الميزة التراجع عن آخر دفعة اعتمدتها، وحذف جلساتها من جدول `tbl_old`، "
            "لكي تتمكن من إعادة قراءة بيانات اليوم وإضافة أي معرض كان متأخراً وإعادة استخراج التقرير."
        )

        last_batch = db.query(ReportBatch).filter(ReportBatch.is_rolled_back == False).order_by(ReportBatch.created_at.desc()).first()
        if not last_batch:
            st.info("لا توجد دفعات نشطة متاحة للتراجع.")
        else:
            st.write(f"• **آخر دفعة معتمدة:** `{last_batch.batch_id}`")
            st.write(f"• **تاريخ التقرير:** {last_batch.report_date} | **عدد الجلسات:** {last_batch.sessions_count}")

            confirm_rb = st.checkbox("أؤكد رغبتي في التراجع عن هذه الدفعة وإلغاء ترحيل جلساتها من tblOLD")
            if st.button("🚨 إلغاء اعتماد الدفعة واسترجاع الجلسات", type="secondary", disabled=not confirm_rb):
                try:
                    db.query(OldClosing).filter(OldClosing.batch_id == last_batch.batch_id).delete()
                    last_batch.is_rolled_back = True
                    db.commit()
                    get_cached_system_stats.clear()
                    st.success(f"✅ تم التراجع عن الدفعة `{last_batch.batch_id}` بنجاح! يمكنك الآن إعادة إصدار التقرير.")
                    time.sleep(2)
                    st.rerun()
                except Exception as e:
                    db.rollback()
                    st.error(f"❌ خطأ أثناء التراجع: {e}")

    # 3. SHOWROOMS & CITIES MANAGEMENT
    with sub_showrooms:
        st.markdown("#### 🏢 قائمة المعارض ونقاط البيع المعتمدة (59 سجلاً):")
        showrooms_list = db.query(Showroom).order_by(Showroom.No.asc()).all()
        sh_data = []
        for s in showrooms_list:
            sh_data.append({
                "الكود": s.No,
                "اسم نقطة البيع في أودو (DESC)": s.DESC,
                "الاسم المختصر للطباعة (Short)": s.Short,
                "رقم المدينة": s.CityNo or "غير محدد",
                "اسم المدينة": s.city.CityName if s.city else "غير محدد",
                "المنطقة": s.city.branch.BranchName if s.city and s.city.branch else "عام",
            })
        st.dataframe(pd.DataFrame(sh_data), use_container_width=True)

        with st.expander("➕ إضافة معرض أو نقطة بيع جديدة:"):
            with st.form("new_showroom_form"):
                new_no = st.text_input("كود المعرض (No): e.g. 160")
                new_desc = st.text_input("اسم نقطة البيع في أودو (DESC)")
                new_short = st.text_input("الاسم المختصر للطباعة (Short)")
                cities = db.query(City).order_by(City.CityNo.asc()).all()
                city_options = {f"{c.CityNo} - {c.CityName}": c.CityNo for c in cities}
                new_city = st.selectbox("المدينة:", options=list(city_options.keys()))

                if st.form_submit_button("حفظ المعرض الجديد"):
                    if not (new_no and new_desc and new_short):
                        st.error("يرجى ملء كافة الحقول المطلوبة.")
                    else:
                        c_no = city_options[new_city]
                        new_sh = Showroom(No=new_no.strip(), DESC=new_desc.strip(), Short=new_short.strip(), CityNo=c_no)
                        try:
                            db.add(new_sh)
                            db.commit()
                            st.success(f"✅ تم إضافة المعرض `{new_short}` بنجاح!")
                            time.sleep(1)
                            st.rerun()
                        except Exception as e:
                            db.rollback()
                            st.error(f"❌ تعذر الحفظ: {e}")

    # 4. SEED / IMPORT OLD CLOSINGS
    with sub_seed:
        st.markdown("#### 📥 إدارة واستيراد أرشيف الإقفالات السابقة (tblOLD):")
        total_old = db.query(OldClosing).count()
        st.info(f"📊 إجمالي عدد الجلسات المؤرشفة والمعتمدة حالياً في قاعدة البيانات (`tbl_old`): **{total_old:,}** جلسة.")

        st.markdown("##### 1️⃣ رفع ملف إكسل أو CSV للأرشيف (مثل `qryExportOldDetailed.xlsx`):")
        old_file = st.file_uploader(
            "اختر ملف الإكسل المصدر من أكسيس أو أودو لتحديث جدول tblOLD:",
            type=["xlsx", "xls", "csv"],
            key="old_closings_uploader"
        )
        if old_file is not None:
            if st.button("🚀 بدء استيراد وتحديث جدول tblOLD من الملف المرفوع"):
                try:
                    with st.spinner("جاري استيراد الأرشيف وفحص السجلات الجديدة..."):
                        stats = processor.import_old_closings_file(old_file, batch_id="FILE_UPLOAD_OLD")
                        get_cached_system_stats.clear()
                        st.success(
                            f"✅ تم الانتهاء بنجاح! إجمالي السجلات في الملف: {stats['total_in_file']:,} | "
                            f"تمت إضافة: {stats['new_added']:,} سجل جديد | "
                            f"الموجود مسبقاً: {stats['already_existed']:,} | "
                            f"الإجمالي الكلي الآن: {stats['final_total']:,}"
                        )
                        time.sleep(1.5)
                        st.rerun()
                except Exception as e:
                    st.error(f"❌ خطأ أثناء الاستيراد: {e}")

        st.markdown("---")
        st.markdown("##### 2️⃣ استيراد أرقام إقفالات نصية يدوية:")
        pasted_text = st.text_area("أرقام الإقفالات (كل رقم في سطر مستقل مثل POS/57800):", height=120)
        if st.button("استيراد الأرقام الملصقة إلى tblOLD"):
            if pasted_text:
                lines_in = [x.strip() for x in pasted_text.splitlines() if x.strip()]
                added = 0
                for c_no in lines_in:
                    exists = db.query(OldClosing).filter(OldClosing.closing_no == c_no).first()
                    if not exists:
                        db.add(OldClosing(closing_no=c_no, batch_id="MANUAL_IMPORT"))
                        added += 1
                db.commit()
                get_cached_system_stats.clear()
                st.success(f"✅ تم استيراد {added} رقم إقفال جديد إلى tblOLD بنجاح!")
                time.sleep(1)
                st.rerun()

    # 5. DEVICE SECURITY & CLOUD RESILIENCE
    with sub_security:
        st.markdown("#### 🛡️ أمان الأجهزة المصرح لها والربط السحابي المقاوم لانقطاع الكهرباء:")

        # Section A: Device Security & Machine ID
        st.markdown("##### 1️⃣ بيانات ترخيص هذا الجهاز (Hardware Licensing):")
        cur_mid = get_machine_id()
        st.info(f"• **كود بصمة العتاد الصلبة (Machine ID):** `{cur_mid}`")
        st.write("• **حالة التفعيل:** 🟢 هذا الجهاز مرخص ومسجل في منظومة حماية شركة العسلي التجارية.")
        st.write("• **التحكم المركزي:** يمكنك إدارة هذا الجهاز وتوليد مفاتيحه أو حظره عن بعد عبر برنامج `Keygen_Admin.exe`.")

        col_sec1, col_sec2 = st.columns(2)
        with col_sec1:
            if st.button("🔄 فحص ومزامنة حالة الجهاز مع Odoo الآن", use_container_width=True):
                from security_gate import fetch_device_status_from_odoo
                st_stat = fetch_device_status_from_odoo(cur_mid)
                st.session_state.odoo_device_status = st_stat
                st.success(f"حالة الجهاز في سجل أودو المركزي: **{st_stat.upper()}**")
        with col_sec2:
            if st.button("⚠️ إلغاء تفعيل هذا الجهاز (Deactivate Device)", use_container_width=True, type="secondary"):
                from security_gate import deactivate_device
                deactivate_device()
                st.warning("تم إلغاء تفعيل هذا الجهاز. جاري إعادة القفل...")
                time.sleep(1)
                st.rerun()

        st.markdown("---")

        # Section B: Cloud Database Resilience
        st.markdown("##### 2️⃣ الربط السحابي وقاعدة البيانات المقاومة لانقطاع الكهرباء (Cloud Database):")
        st.write(
            "لحماية عملك من انقطاع الكهرباء أو تعطل أجهزة الشركة، يمكنك تشغيل البرنامج من أي جهاز (لابتوب أو منزل) "
            "مباشرة مع قاعدة بيانات سحابية مشفرة (مثل **Supabase** أو **Neon PostgreSQL**) دون الحاجة لوجود سيرفر محلي."
        )

        db_stat = get_database_status_info()
        if db_stat["is_cloud"]:
            st.success(f"🟢 النظام متصل حالياً بقاعدة بيانات سحابية مركزية مشفرة: **{db_stat['type']}**\n\n📌 المضيف السحابي: `{db_stat['url_masked']}`")
            st.info("✨ ميزة الاستقلالية مفعلة: يمكنك الآن تشغيل البرنامج من أي جهاز حاسوب أو لابتوب مصرح به وستظهر نفس البيانات وآخر إقفالات ومبيعات دون تكرار أو تعارض نهائياً.")
            if st.button("🔄 التبديل مؤقتاً إلى القاعدة المحلية (Offline Mode)"):
                env_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
                if os.path.exists(env_file):
                    with open(env_file, "r", encoding="utf-8") as f:
                        content = f.read()
                    content = content.replace('DATABASE_URL=', '# DATABASE_URL=')
                    with open(env_file, "w", encoding="utf-8") as f:
                        f.write(content)
                st.success("تم التبديل للقاعدة المحلية. جاري إعادة التحميل...")
                st.rerun()
        else:
            st.warning(f"🟡 النظام يعمل حالياً على قاعدة بيانات محلية مستقلة: **{db_stat['type']}**.")

        with st.expander("🚀 أداة ترحيل الأرشيف والربط السحابي (Cloud Database Connection & Migration):"):
            st.write("أدخل رابط قاعدة البيانات السحابية (PostgreSQL / Supabase) لترحيل الـ 22,867 جلسة أو تفعيل الربط السحابي لهذا الجهاز:")
            target_pg_url = st.text_input(
                "رابط اتصال PostgreSQL السحابي المشفر:",
                placeholder="postgresql://postgres:password@db.supabase.co:5432/postgres?sslmode=require",
                type="password"
            )
            col_m1, col_m2 = st.columns(2)
            with col_m1:
                if st.button("🚀 بدء ترحيل الأرشيف والجداول إلى السحابة", type="primary"):
                    if not target_pg_url:
                        st.error("يرجى إدخال رابط اتصال PostgreSQL.")
                    else:
                        with st.spinner("جاري ترحيل كافة الجداول والأرشيف السحابي... قد يستغرق ذلك دقيقة واحدة:"):
                            try:
                                from migrate_to_cloud_db import migrate_sqlite_to_cloud_postgresql, clean_pg_url
                                ok = migrate_sqlite_to_cloud_postgresql(target_pg_url)
                                if ok:
                                    cleaned = clean_pg_url(target_pg_url)
                                    env_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
                                    with open(env_file, "w", encoding="utf-8") as f:
                                        f.write(f'# Daily Sales Cloud App — Configuration\nAPP_PASSWORD=assali2026\nDATABASE_URL="{cleaned}"\n')
                                    st.success("✅ تم ترحيل كافة الجداول والأرشيف بنجاح تام، وتم ربط البرنامج بالسحابة فوراً!")
                                    st.rerun()
                                else:
                                    st.error("❌ فشل الترحيل، يرجى التأكد من صحة الرابط والاتصال بالإنترنت.")
                            except Exception as ex:
                                st.error(f"❌ خطأ أثناء الترحيل: {ex}")
            with col_m2:
                if st.button("🔗 حفظ الرابط والاتصال بالسحابة مباشرة"):
                    if not target_pg_url:
                        st.error("يرجى إدخال رابط اتصال PostgreSQL.")
                    else:
                        from database import clean_pg_url
                        cleaned = clean_pg_url(target_pg_url)
                        env_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
                        with open(env_file, "w", encoding="utf-8") as f:
                            f.write(f'# Daily Sales Cloud App — Configuration\nAPP_PASSWORD=assali2026\nDATABASE_URL="{cleaned}"\n')
                        st.success("✅ تم حفظ الرابط وربط البرنامج بالسحابة بنجاح!")
                        st.rerun()

    db.close()
