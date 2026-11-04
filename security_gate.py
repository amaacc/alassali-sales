"""
security_gate.py — Enterprise Cryptographic Hardware Licensing & Activation Gate
Company: Al-Assali Trading Co. (شركة العسلي للتجارة)
Architect: Abdullah Al-Khatib
"""

import os
import sys
import platform
import time
from typing import Tuple, Dict, Any, Optional
import streamlit as st

# Ensure root directory is accessible for licensing import
root_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if root_dir not in sys.path:
    sys.path.insert(0, root_dir)

try:
    from licensing import (
        get_machine_id,
        verify_license_key,
        is_activated,
        activate_device,
        deactivate_device,
        fetch_device_status_from_odoo,
        register_or_update_device_in_odoo,
        _MASTER_PASS
    )
except ImportError:
    # Standalone fallback if licensing.py is missing
    def get_machine_id():
        import hashlib, uuid
        return hashlib.sha256(str(uuid.getnode()).encode()).hexdigest()[:19].upper()
    def is_activated(): return True
    def activate_device(k, n=None): return True, "Activated"
    def deactivate_device(): return True
    def fetch_device_status_from_odoo(d=None): return "active"
    def register_or_update_device_in_odoo(**kw): return "active"
    _MASTER_PASS = "KHATIB@MASTER#2026"


APP_NAME = "DailySalesApp"
APP_VERSION = "v1.00"


def enforce_security_activation_gate() -> bool:
    """
    Renders corporate activation lock screen if the machine is unauthorized.
    Halts Streamlit application rendering until validated and activated.
    Returns True if machine is authorized and active.
    """
    # 0. Check if running in cloud deployment mode (Linux / Streamlit Community Cloud / Cloud Hosted)
    if sys.platform != "win32" or os.environ.get("IS_CLOUD_DEPLOYMENT") == "1" or os.environ.get("STREAMLIT_SHARING_HOST"):
        return True
    try:
        if hasattr(st, "secrets") and (st.secrets.get("IS_CLOUD_DEPLOYMENT", False) is True or "DATABASE_URL" in st.secrets):
            return True
    except Exception:
        pass

    machine_id = get_machine_id()

    # 1. Check local cryptographic license file (%APPDATA%\AlAssaliOdooExporter\license.dat)
    local_active = is_activated()

    # 2. Check remote kill-switch / blacklist status in Odoo device registry
    # Check remote status once per session to maintain high speed
    if "odoo_device_status" not in st.session_state:
        try:
            remote_status = fetch_device_status_from_odoo(machine_id)
            st.session_state.odoo_device_status = remote_status
        except Exception:
            st.session_state.odoo_device_status = "active"

    remote_status = st.session_state.get("odoo_device_status", "active")

    # If blocked remotely by Admin in Keygen_Admin / Odoo, revoke immediately
    if remote_status == "blocked":
        deactivate_device()
        st.markdown(
            """
            <div style="background-color: #7F1D1D; color: white; padding: 25px; border-radius: 12px; text-align: center; margin: 40px auto; max-width: 650px;">
                <h2 style="color: #FECACA; margin-bottom: 10px;">🚫 تم حظر هذا الجهاز من قبل الإدارة</h2>
                <p style="font-size: 13pt; line-height: 1.6;">
                    تم سحب ترخيص هذا الجهاز وإيقاف تشغيله عبر سجل الأجهزة المركزي في أودو.<br/>
                    يرجى مراجعة إدارة تقنية المعلومات بشركة العسلي للتجارة لإعادة التفعيل.
                </p>
                <div style="background: rgba(0,0,0,0.3); padding: 10px; border-radius: 8px; font-family: monospace; font-size: 14pt; margin-top: 15px;">
                    كود الجهاز المحظور: """ + machine_id + """
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )
        st.stop()
        return False

    # If already activated locally and not blocked
    if local_active:
        # Register / heartbeat in Odoo in background
        if "heartbeat_done" not in st.session_state:
            try:
                register_or_update_device_in_odoo(
                    device_id=machine_id,
                    device_name=os.environ.get("COMPUTERNAME", platform.node()),
                    app_name=APP_NAME,
                    app_version=APP_VERSION
                )
                st.session_state.heartbeat_done = True
            except Exception:
                pass
        return True

    # =========================================================================
    # RENDER CORPORATE ACTIVATION LOCK SCREEN
    # =========================================================================
    st.markdown(
        """
        <div style="background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%); color: white; padding: 28px; border-radius: 14px; text-align: center; margin: 20px auto; max-width: 680px; border: 1px solid #334155; box-shadow: 0 10px 25px rgba(0,0,0,0.3);">
            <div style="font-size: 42px; margin-bottom: 10px;">🔐</div>
            <h2 style="color: #38BDF8; font-weight: 800; margin: 0;">بوابة الأمان والترخيص المشفر</h2>
            <h4 style="color: #94A3B8; font-weight: 500; margin: 6px 0 18px 0;">شركة العسلي للتجارة — Al-Assali Trading Co.</h4>
            <div style="background-color: rgba(239, 68, 68, 0.15); border: 1px solid #EF4444; color: #FCA5A5; padding: 12px 18px; border-radius: 8px; font-size: 11pt; margin-bottom: 20px;">
                ⚠️ <b>تنبيه أمني:</b> هذا الجهاز غير مصرح له بتشغيل نظام مراقبة وتقارير مبيعات المعارض.<br/>
                لحماية البيانات المالية السرية للشركة، يجب تفعيل الجهاز أولاً عبر كود الترخيص المعتمد.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    with st.container():
        col_c1, col_center, col_c2 = st.columns([1, 2.5, 1])
        with col_center:
            st.markdown("##### 📌 كود بصمة العتاد الخاصة بهذا الجهاز (Machine ID):")
            st.code(machine_id, language="text")
            st.caption("قم بنسخ هذا الكود وإرساله للمسؤول لتوليد مفتاح التفعيل عبر برنامج `Keygen_Admin`.")

            with st.form("activation_form"):
                st.markdown("##### 🔑 إدخال بيانات التفعيل:")
                device_name_input = st.text_input(
                    "اسم الجهاز / المعرف الوصفي:",
                    value=os.environ.get("COMPUTERNAME", platform.node()),
                    placeholder="مثال: لابتوب الإدارة المالية"
                )
                license_key_input = st.text_input(
                    "مفتاح التفعيل (Activation Key):",
                    type="password",
                    placeholder="XXXX-XXXX-XXXX-XXXX"
                )

                submit_btn = st.form_submit_button("🚀 تفعيل هذا الجهاز والبدء الآن", type="primary", use_container_width=True)

                if submit_btn:
                    if not license_key_input:
                        st.error("يرجى إدخال مفتاح التفعيل.")
                    else:
                        with st.spinner("جاري التحقق من التوقيع الرقمي والاتصال بسجل أودو..."):
                            success, msg = activate_device(
                                input_key=license_key_input,
                                device_name=device_name_input,
                                app_name=APP_NAME,
                                app_version=APP_VERSION
                            )
                            if success:
                                st.success(f"✅ {msg}")
                                st.session_state.odoo_device_status = "active"
                                time.sleep(1.2)
                                st.rerun()
                            else:
                                st.error(f"❌ {msg}")

    # Stop rendering the rest of the application
    st.stop()
    return False
