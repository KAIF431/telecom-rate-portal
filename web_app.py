import os
import io
import re
import time
import json
import base64
import zipfile
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import importlib

import auth_manager as am
importlib.reload(am)
import secrets
from config_manager import load_config, add_member, remove_member, set_active_member
from core_engine import (
    detect_header_row_and_columns,
    load_qc_exception_codes,
    fast_process_dataframe,
    generate_target_excel_bytes,
    build_output_filename,
    TEMPLATE_DIR
)

# Output directory for reliable caching of generated Excel files
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "generated_output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Isolated Per-Browser Session Handlers (Prevents Chrome vs Incognito session collision)
def get_current_browser_token():
    try:
        tok = st.query_params.get("session_token")
        if isinstance(tok, list):
            tok = tok[0] if tok else None
        if tok:
            tok = str(tok).strip()
        return tok if tok else None
    except Exception:
        return None

def authenticate_current_session():
    # If already set in this browser's session state
    if st.session_state.get("user") is not None:
        return st.session_state.user
    
    # Check URL query param session token
    token = get_current_browser_token()
    if token:
        user_info = am.get_session_user(token)
        if user_info:
            st.session_state.user = user_info
            st.session_state.session_token = token
            return user_info
        else:
            try:
                st.query_params.clear()
            except Exception:
                pass
    return None

def set_login_session(session_data):
    token = secrets.token_hex(16)
    am.register_session(token, session_data)
    st.session_state.user = session_data
    st.session_state.session_token = token
    try:
        st.query_params["session_token"] = token
    except Exception:
        pass

def clear_current_session():
    token = st.session_state.get("session_token") or get_current_browser_token()
    if token:
        am.destroy_session(token)
    try:
        st.query_params.clear()
    except Exception:
        try:
            if "session_token" in st.query_params:
                del st.query_params["session_token"]
        except Exception:
            pass
    st.session_state.user = None
    st.session_state.session_token = None
    st.session_state.pending_totp_user = None
    st.session_state.processed_results = None
    st.session_state.elapsed_time = None

st.set_page_config(
    page_title="Vendor Rate Sheet Standardizer",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling (Clean, professional, completely hidden Streamlit default header/toolbar)
st.markdown("""
<style>
    .main-title {
        font-size: 2.1rem;
        font-weight: 700;
        color: #1E3A8A;
        margin-bottom: 0.2rem;
    }
    .sub-title {
        font-size: 1rem;
        color: #4B5563;
        margin-bottom: 1.2rem;
    }
    .badge-success {
        background-color: #DEF7EC;
        color: #03543F;
        padding: 8px 16px;
        border-radius: 6px;
        font-weight: 600;
        display: inline-block;
        margin: 8px 0;
        font-size: 1rem;
    }
    .auto-dl-banner {
        background-color: #EFF6FF;
        border-left: 4px solid #2563EB;
        padding: 12px 18px;
        color: #1E40AF;
        font-weight: 600;
        border-radius: 4px;
        margin: 12px 0;
        font-size: 1rem;
    }
    /* Hide Streamlit default Deploy button, 3-dots menu & footer while keeping sidebar toggle visible */
    #MainMenu {visibility: hidden !important; display: none !important;}
    footer {visibility: hidden !important; display: none !important;}
    .stDeployButton {display: none !important;}
    [data-testid="stToolbar"] {visibility: hidden !important; display: none !important;}
    [data-testid="stDecoration"] {visibility: hidden !important; display: none !important;}
    [data-testid="stStatusWidget"] {visibility: hidden !important; display: none !important;}
    header {background: transparent !important;}
    [data-testid="stSidebarCollapsedControl"] {
        display: block !important;
        visibility: visible !important;
        z-index: 999999 !important;
    }
</style>
""", unsafe_allow_html=True)

# ----------------- AUTHENTICATION CHECK -----------------
st.session_state.user = authenticate_current_session()

if "pending_totp_user" not in st.session_state:
    st.session_state.pending_totp_user = None

if "email_otp_sent_for" not in st.session_state:
    st.session_state.email_otp_sent_for = None

if st.session_state.user is None:
    # Render Login Page (Centered Box)
    login_c1, login_c2, login_c3 = st.columns([1, 1.8, 1])
    with login_c2:
        st.markdown("""
        <div style="background: white; padding: 26px 30px; border-radius: 12px; border: 1px solid #E2E8F0; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.06); margin-top: 30px; margin-bottom: 20px;">
            <div style="text-align: center; margin-bottom: 18px;">
                <div style="font-size: 2.3rem;">🔐</div>
                <div style="font-size: 1.55rem; font-weight: 700; color: #1E3A8A; margin-top: 4px;">Rate Standardizer Portal</div>
                <div style="font-size: 0.88rem; color: #64748B;">Authorized access with Password + Authy 2FA.</div>
            </div>
        """, unsafe_allow_html=True)

        # Check if user is in Step 2 (Authy / Authenticator Code)
        if st.session_state.pending_totp_user is not None:
            pending = st.session_state.pending_totp_user
            st.markdown(f"#### 📱 Step 2: Authy Verification")
            st.info(f"User: **{pending['name']}** (`{pending['username']}`)")
            st.caption("Enter the current 6-digit verification code from Admin's Authy app:")

            totp_input = st.text_input("6-Digit Authy Code", max_chars=6, key="login_totp", placeholder="e.g. 123456")
            st.write("")

            c_btn1, c_btn2 = st.columns([1.5, 1])
            if c_btn1.button("🚀 Verify & Log In", type="primary", use_container_width=True):
                if totp_input and len(totp_input.strip()) == 6:
                    ok_t, session_data = am.complete_totp_login(pending, totp_input)
                    if ok_t:
                        set_login_session(session_data)
                        st.session_state.pending_totp_user = None
                        st.success(f"Login successful! Welcome back, {session_data['name']}!")
                        st.rerun()
                    else:
                        st.error(session_data)
                else:
                    st.warning("Please enter a valid 6-digit code.")

            if c_btn2.button("⬅️ Back", use_container_width=True):
                st.session_state.pending_totp_user = None
                st.rerun()

        else:
            tab_login, tab_forgot = st.tabs(["🔑 Sign In", "❓ Forgot Password"])

            with tab_login:
                u_name = st.text_input("Username", key="login_u", placeholder="e.g. kaif")
                p_word = st.text_input("Password", type="password", key="login_p", placeholder="Enter your password")
                st.write("")
                if st.button("👉 Continue to Authy 2FA", type="primary", use_container_width=True):
                    if u_name and p_word:
                        ok_cred, res_cred = am.verify_credentials(u_name, p_word)
                        if ok_cred:
                            st.session_state.pending_totp_user = res_cred
                            st.rerun()
                        else:
                            st.error(res_cred)
                    else:
                        st.warning("Both Username and Password are required.")

            with tab_forgot:
                st.markdown("**Choose Password Reset Option:**")
                reset_opt = st.radio(
                    "Select Method:",
                    ["👤 Option 1: Request Admin to Reset Password", f"📧 Option 2: Reset via OTP to {am.ADMIN_EMAIL}"],
                    label_visibility="collapsed"
                )

                if "Option 1" in reset_opt:
                    st.caption("Your password reset request will be submitted directly to Admin for approval:")
                    req_u = st.text_input("Your Username", key="req_u_admin")
                    req_note = st.text_area("Note / Reason for Admin", placeholder="e.g. Forgot my password, please reset.", key="req_note_admin")
                    if st.button("📩 Send Request to Admin", use_container_width=True, type="primary"):
                        if req_u:
                            ok_req, msg_req = am.request_password_reset(req_u, req_note)
                            if ok_req:
                                st.success(msg_req)
                            else:
                                st.error(msg_req)
                        else:
                            st.warning("Please enter your username.")

                else:
                    st.caption(f"A 6-digit verification OTP will be sent to Admin email (`{am.ADMIN_EMAIL}`):")
                    otp_u = st.text_input("Your Username", key="otp_u_email")
                    
                    if st.button("📤 Send OTP to Email", use_container_width=True):
                        if otp_u:
                            ok_o, msg_o, code_val = am.generate_email_otp(otp_u)
                            if ok_o:
                                st.session_state.email_otp_sent_for = otp_u
                                st.success(msg_o)
                            else:
                                st.error(msg_o)
                        else:
                            st.warning("Please enter your username first.")

                    if st.session_state.email_otp_sent_for:
                        st.divider()
                        st.markdown(f"**Reset Password for `{st.session_state.email_otp_sent_for}`:**")
                        user_entered_otp = st.text_input("Enter 6-digit OTP received on email:", max_chars=6, key="entered_otp")
                        new_p_input = st.text_input("Enter New Password:", type="password", key="new_p_input")
                        
                        if st.button("✅ Verify OTP & Save New Password", type="primary", use_container_width=True):
                            if user_entered_otp and new_p_input:
                                ok_v, msg_v = am.verify_email_otp(st.session_state.email_otp_sent_for, user_entered_otp)
                                if ok_v:
                                    am.update_password(st.session_state.email_otp_sent_for, new_p_input)
                                    st.session_state.email_otp_sent_for = None
                                    st.success("🎉 Password changed successfully! You can now log in with your new password.")
                                    time.sleep(1.5)
                                    st.rerun()
                                else:
                                    st.error(msg_v)
                            else:
                                st.warning("Both OTP and New Password are required.")

        st.markdown("</div>", unsafe_allow_html=True)
    st.stop()

# User is authenticated
current_user = st.session_state.user or {}
is_admin = (current_user.get("role") == "admin")
active_member = current_user.get("name", "Kaif")

# ----------------- SIDEBAR: USER INFO & ADMIN CONSOLE -----------------
with st.sidebar:
    st.markdown(f"""
    <div style="padding: 10px 14px; background: #EEF2FF; border-left: 4px solid #4F46E5; border-radius: 6px; margin-bottom: 12px;">
        <div style="font-size: 0.72rem; color: #6B7280; text-transform: uppercase; font-weight: 600;">Logged In User</div>
        <div style="font-size: 1.15rem; color: #1E1B4B; font-weight: 700;">{current_user.get('name')}</div>
        <div style="font-size: 0.8rem; color: #4F46E5; font-weight: 600;">Role: {current_user.get('role', 'user').upper()}</div>
        <div style="font-size: 0.72rem; color: #9CA3AF; margin-top: 4px;">📅 Session active till Monday 00:00</div>
    </div>
    """, unsafe_allow_html=True)

    if st.button("🚪 Logout", key="sidebar_logout_btn", use_container_width=True):
        clear_current_session()
        st.rerun()

    st.divider()

    # Admin Management Section (Visible ONLY to Admin like Kaif)
    if is_admin:
        st.subheader("⚙️ Admin Management")
        
        # 1. Pending Reset Requests
        pending_reqs = am.get_pending_reset_requests()
        if pending_reqs:
            st.warning(f"🔔 **{len(pending_reqs)} Password Reset Request(s) Pending!**")
            with st.expander("📬 Review Reset Requests", expanded=True):
                for pr in pending_reqs:
                    st.write(f"**User:** `{pr['username']}` ({pr['name']})")
                    st.write(f"**Current Password:** `{pr.get('current_password', '')}`")
                    st.caption(f"Time: {pr['requested_at']} | Note: {pr.get('note', '')}")
                    new_pw_for_req = st.text_input(f"New Password for {pr['username']}:", key=f"pw_req_{pr['username']}")
                    rc1, rc2 = st.columns(2)
                    if rc1.button("Set PW", key=f"btn_set_{pr['username']}", type="primary", use_container_width=True):
                        if new_pw_for_req:
                            am.update_password(pr['username'], new_pw_for_req)
                            st.success(f"Password updated for {pr['username']}!")
                            st.rerun()
                        else:
                            st.warning("Enter a new password.")
                    if rc2.button("Dismiss", key=f"btn_dism_{pr['username']}", use_container_width=True):
                        am.dismiss_reset_request(pr['username'])
                        st.rerun()
                    st.divider()

        # 2. Latest Active Email OTP (for kaifofficial4@gmail.com)
        active_otp_info = am.get_last_active_otp()
        if active_otp_info:
            with st.expander("📧 Latest Email OTP Sent", expanded=False):
                st.info(f"**User:** `{active_otp_info['username']}`\n\n**OTP Code:** `{active_otp_info['otp']}`\n\nSent to: `{active_otp_info['target_email']}` (Time: {active_otp_info['created_at']})")

        # 3. Master Users Table (ID & Password explicitly visible to Admin!)
        with st.expander("👥 All Users List (IDs & Passwords)", expanded=False):
            st.caption(f"🔑 Shared Master Authy Key: `{am.get_master_totp_secret()}`")
            all_users = am.get_all_users()
            user_table_data = []
            for u_k, u_v in all_users.items():
                user_table_data.append({
                    "Username": u_k,
                    "Name": u_v.get("name"),
                    "Password": u_v.get("password"),
                    "Role": u_v.get("role", "user").upper(),
                    "Authy Key": u_v.get("totp_secret", am.get_master_totp_secret()),
                    "Created At": u_v.get("created_at")
                })
            st.dataframe(pd.DataFrame(user_table_data), use_container_width=True, hide_index=True)

        # 4. Create New User & Set Password Direct
        with st.expander("➕ Create New User", expanded=False):
            st.caption(f"ℹ️ New users will use Admin's Master Authy 2FA (`{am.get_master_totp_secret()}`).")
            new_u_username = st.text_input("Username", key="new_u_username", placeholder="e.g. rahul")
            new_u_name = st.text_input("Full Name", key="new_u_name", placeholder="e.g. Rahul Sharma")
            new_u_pw = st.text_input("Password (Set Directly)", key="new_u_pw", placeholder="e.g. rahul123")
            new_u_role = st.selectbox("Role", ["user", "admin"], key="new_u_role")
            new_u_email = st.text_input("Email (Optional)", key="new_u_email")
            if st.button("Create User Account", use_container_width=True, type="primary"):
                if new_u_username and new_u_pw:
                    ok_c, msg_c = am.create_user(new_u_username, new_u_name, new_u_pw, new_u_role, new_u_email)
                    if ok_c:
                        st.success(msg_c)
                        st.rerun()
                    else:
                        st.error(msg_c)
                else:
                    st.warning("Both Username and Password are required.")

        # 5. Reset Any User Password
        with st.expander("🔑 Reset Any User Password", expanded=False):
            all_users = am.get_all_users()
            target_user = st.selectbox("Select User", list(all_users.keys()), key="reset_target_u")
            st.caption(f"Current Password: `{all_users[target_user].get('password', '')}`")
            new_pw_admin = st.text_input("New Password to Set:", key=f"admin_new_pw_{target_user}")
            if st.button("Update Password Now", key=f"admin_btn_save_pw_{target_user}", use_container_width=True, type="primary"):
                if new_pw_admin:
                    ok_u, msg_u = am.update_password(target_user, new_pw_admin)
                    if ok_u:
                        st.success(msg_u)
                        st.rerun()
                    else:
                        st.error(msg_u)
                else:
                    st.warning("Password cannot be blank.")

        # 6. Edit User Name
        with st.expander("✏️ Edit User Display Name", expanded=False):
            sb_users = am.get_all_users()
            sb_u_sel = st.selectbox("Select User", list(sb_users.keys()), key="sidebar_edit_name_user")
            sb_cur_name = sb_users[sb_u_sel].get("name", "")
            sb_new_name = st.text_input(f"New Name for {sb_u_sel}:", value=sb_cur_name, key=f"sidebar_new_name_input_{sb_u_sel}")
            if st.button("Save Name", key=f"sidebar_btn_save_name_{sb_u_sel}", use_container_width=True, type="primary"):
                if sb_new_name and sb_new_name.strip():
                    ok_sn, msg_sn = am.update_user_name(sb_u_sel, sb_new_name.strip())
                    if ok_sn:
                        if sb_u_sel == current_user.get("username"):
                            current_user["name"] = sb_new_name.strip()
                            st.session_state.user["name"] = sb_new_name.strip()
                        st.success(msg_sn)
                        st.rerun()
                    else:
                        st.error(msg_sn)
                else:
                    st.warning("Name cannot be empty.")

        # 7. Delete User Account
        with st.expander("🗑️ Delete User Account", expanded=False):
            deletable_users = [u for u in am.get_all_users().keys() if u != "kaif"]
            if deletable_users:
                u_to_delete = st.selectbox("Select User to Remove", deletable_users, key="del_u_select")
                if st.button("Delete Account", use_container_width=True, type="primary"):
                    ok_d, msg_d = am.delete_user(u_to_delete)
                    if ok_d:
                        st.success(msg_d)
                        st.rerun()
                    else:
                        st.error(msg_d)
            else:
                st.caption("No non-admin users to delete.")

        # 7. Recent Processing Activity
        with st.expander("📜 Recent User Processing Log", expanded=False):
            sb_hist = am.get_processing_history()[:8]
            if sb_hist:
                for sh in sb_hist:
                    st.markdown(f"""
                    <div style="font-size: 0.85rem; padding: 6px 0;">
                        👤 <b>{sh.get('user_name')}</b> (<code>{sh.get('username')}</code>)<br>
                        🏷️ <code>{sh.get('vendor_name')}</code><br>
                        📊 {sh.get('clean_rows', 0):,} rows &nbsp;|&nbsp; ⏱️ {sh.get('elapsed_sec', 0)}s<br>
                        <small style="color: #64748B;">{sh.get('timestamp')}</small>
                    </div>
                    """, unsafe_allow_html=True)
                    st.divider()
            else:
                st.caption("No processing activity recorded yet.")

        st.divider()

# ----------------- MAIN PORTAL -----------------
top_head_col1, top_head_col2 = st.columns([3, 1.2])

with top_head_col1:
    st.markdown('<div class="main-title">⚡ Vendor Rate Sheet Standardizer</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">High-speed wholesale rate deck standardization, zero data loss, exact sample formatting.</div>', unsafe_allow_html=True)

with top_head_col2:
    st.markdown(f"""
    <div style="text-align: right; margin-top: 4px; margin-bottom: 6px;">
        <span style="background: #EEF2FF; color: #3730A3; font-weight: 600; padding: 5px 12px; border-radius: 20px; font-size: 0.85rem; border: 1px solid #C7D2FE;">
            👤 <b>{current_user.get('name')}</b> ({current_user.get('role', 'user').upper()})
        </span>
    </div>
    """, unsafe_allow_html=True)
    
    if st.button("🚪 Log Out", key="top_logout_btn", use_container_width=True, type="secondary"):
        clear_current_session()
        st.rerun()

# Template Download Buttons Row
t_col1, t_col2, t_col3 = st.columns([1.5, 1.5, 3])

if is_admin:
    st.markdown("---")
    admin_nav_mode = st.radio(
        "Navigation Mode:",
        ["⚡ Rate Deck Standardizer", "👑 Admin Management Panel"],
        horizontal=True,
        key="main_admin_nav_toggle"
    )
    st.markdown("---")
    
    if admin_nav_mode == "👑 Admin Management Panel":
        st.markdown("## 👑 Admin Control Panel")
        st.caption("Master controls: Manage all users, view IDs & passwords, reset credentials, review requests.")
        
        a_tab1, a_tab_hist, a_tab_edit, a_tab2, a_tab3, a_tab4 = st.tabs([
            "👥 All Users (IDs & Passwords)",
            "📜 User Activity & Processing History",
            "✏️ Edit User Names",
            "➕ Create New User",
            "🔑 Reset Passwords",
            "🔔 Pending Reset Requests"
        ])
        
        with a_tab1:
            st.subheader("👥 Registered Users Master Table")
            st.caption(f"🔑 **Shared Master Authy Key:** `{am.get_master_totp_secret()}` (All accounts authenticate using this key on Admin's device)")
            all_u = am.get_all_users()
            u_data = []
            for uk, uv in all_u.items():
                u_data.append({
                    "Username": uk,
                    "Name": uv.get("name"),
                    "Password": uv.get("password"),
                    "Role": uv.get("role", "user").upper(),
                    "Authy Key": uv.get("totp_secret", am.get_master_totp_secret()),
                    "Created At": uv.get("created_at")
                })
            st.dataframe(pd.DataFrame(u_data), use_container_width=True, hide_index=True)
            
            act_otp = getattr(am, 'get_last_active_otp', lambda: None)()
            if act_otp:
                st.info(f"📧 **Latest Active Email OTP:** Code: `{act_otp['otp']}` for User: `{act_otp['username']}` (Sent to: {act_otp['target_email']})")

        with a_tab_hist:
            st.subheader("📜 User Processing History & Audit Trail")
            st.caption("Complete tracking of which user ID processed which vendor sheet, clean row counts, and timestamps.")
            
            history = am.get_processing_history()
            if not history:
                st.info("ℹ️ No rate sheets processed yet. As users process files, complete activity records will appear here in real time.")
            else:
                tot_sheets = len(history)
                tot_rows = sum(h.get("clean_rows", 0) for h in history)
                active_users_count = len(set(h.get("username") for h in history if h.get("username")))
                
                m1, m2, m3, m4 = st.columns([1.5, 1.5, 1.5, 1.2])
                m1.metric("Total Sheets Standardized", f"{tot_sheets:,}")
                m2.metric("Total Rows Processed", f"{tot_rows:,}")
                m3.metric("Contributing Users", f"{active_users_count}")
                with m4:
                    st.write("")
                    if st.button("🗑️ Clear Audit Log", key="main_clear_hist_btn", use_container_width=True):
                        am.clear_processing_history()
                        st.success("Activity log cleared!")
                        time.sleep(0.8)
                        st.rerun()
                
                st.divider()
                
                fc1, fc2 = st.columns([1.5, 2.5])
                with fc1:
                    user_options = ["All Users"] + sorted(list(set(h.get("username", "") for h in history if h.get("username"))))
                    selected_user = st.selectbox("Filter by User ID:", user_options, key="audit_filter_user")
                with fc2:
                    filter_kw = st.text_input("Search by Vendor or File Name:", key="audit_search_kw", placeholder="Type vendor or filename to search...")
                
                filtered = history
                if selected_user != "All Users":
                    filtered = [h for h in filtered if h.get("username") == selected_user]
                if filter_kw:
                    kw_low = filter_kw.strip().lower()
                    filtered = [h for h in filtered if kw_low in h.get("vendor_name", "").lower() or kw_low in h.get("output_filename", "").lower()]
                
                table_rows = []
                for entry in filtered:
                    table_rows.append({
                        "Timestamp": entry.get("timestamp"),
                        "User ID": entry.get("username"),
                        "User Full Name": entry.get("user_name"),
                        "Vendor / Trunk Name": entry.get("vendor_name"),
                        "Standardized File": entry.get("output_filename"),
                        "Clean Rows": f"{entry.get('clean_rows', 0):,}",
                        "QC Excluded": f"{entry.get('qc_excluded', 0):,}",
                        "Speed": f"{entry.get('elapsed_sec', 0)}s"
                    })
                
                st.dataframe(pd.DataFrame(table_rows), use_container_width=True, hide_index=True)

        with a_tab_edit:
            st.subheader("✏️ Edit User Display / Full Name")
            st.caption("Change the full name of Admin (yourself) or any user. The updated name reflects immediately across the portal and generated filenames.")
            all_users_dict = am.get_all_users()
            selected_edit_u = st.selectbox("Select User Account to Edit:", list(all_users_dict.keys()), key="main_edit_u_select")
            u_info = all_users_dict[selected_edit_u]
            
            st.info(f"**Username:** `{selected_edit_u}` | **Current Full Name:** `{u_info.get('name')}` | **Role:** `{u_info.get('role', 'user').upper()}`")
            
            new_name_val = st.text_input("New Full Name / Display Name:", value=u_info.get("name", ""), key=f"main_input_edit_name_{selected_edit_u}")
            if st.button("💾 Save Updated Name", type="primary", use_container_width=True, key=f"main_btn_save_name_{selected_edit_u}"):
                if new_name_val and new_name_val.strip():
                    ok_n, msg_n = am.update_user_name(selected_edit_u, new_name_val.strip())
                    if ok_n:
                        if selected_edit_u == current_user.get("username"):
                            current_user["name"] = new_name_val.strip()
                            st.session_state.user["name"] = new_name_val.strip()
                        st.success(msg_n)
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.error(msg_n)
                else:
                    st.warning("Name cannot be empty.")

        with a_tab2:
            st.subheader("➕ Create New User")
            st.caption(f"ℹ️ New user account will automatically use Admin's Master Authy 2FA (`{am.get_master_totp_secret()}`).")
            cu_col1, cu_col2 = st.columns(2)
            with cu_col1:
                nu_uname = st.text_input("Username", key="main_nu_uname", placeholder="e.g. rahul")
                nu_name = st.text_input("Full Name", key="main_nu_name", placeholder="e.g. Rahul Sharma")
            with cu_col2:
                nu_pw = st.text_input("Password (Set Directly)", key="main_nu_pw", placeholder="e.g. rahul123")
                nu_role = st.selectbox("Role", ["user", "admin"], key="main_nu_role")
            nu_email = st.text_input("Email (Optional)", key="main_nu_email")
            
            if st.button("🚀 Create User Account Now", type="primary", use_container_width=True):
                if nu_uname and nu_pw:
                    ok_cu, msg_cu = am.create_user(nu_uname, nu_name, nu_pw, nu_role, nu_email)
                    if ok_cu:
                        st.success(msg_cu)
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.error(msg_cu)
                else:
                    st.warning("Both Username and Password are required.")

        with a_tab3:
            st.subheader("🔑 Reset Any User's Password")
            all_u = am.get_all_users()
            sel_u = st.selectbox("Select User Account:", list(all_u.keys()), key="main_sel_reset_u")
            st.info(f"Selected: **{sel_u}** | Current Password: `{all_u[sel_u].get('password', '')}` | Authy Key: `{all_u[sel_u].get('totp_secret', am.get_master_totp_secret())}`")
            new_pw_main = st.text_input("Enter New Password to Set:", key=f"main_new_pw_input_{sel_u}")
            if st.button("💾 Save New Password", type="primary", use_container_width=True, key=f"main_btn_save_pw_{sel_u}"):
                if new_pw_main:
                    ok_up, msg_up = am.update_password(sel_u, new_pw_main)
                    if ok_up:
                        st.success(msg_up)
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.error(msg_up)
                else:
                    st.warning("Password cannot be blank.")

        with a_tab4:
            st.subheader("🔔 Password Reset Requests from Users")
            p_reqs = am.get_pending_reset_requests()
            if not p_reqs:
                st.success("✅ No pending reset requests. Everything is up to date!")
            else:
                for rq in p_reqs:
                    st.markdown(f"""
                    <div style="background: #FFFBEB; border-left: 4px solid #F59E0B; padding: 12px 16px; border-radius: 6px; margin-bottom: 12px;">
                        <b>User:</b> {rq['username']} ({rq['name']})<br>
                        <b>Current Password:</b> <code>{rq.get('current_password', '')}</code><br>
                        <b>Note:</b> {rq.get('note', 'No note')}<br>
                        <small style="color: #78350F;">Requested: {rq.get('requested_at', '')}</small>
                    </div>
                    """, unsafe_allow_html=True)
                    new_pw_for_rq = st.text_input(f"New Password for {rq['username']}:", key=f"main_rq_pw_{rq['username']}")
                    rq_c1, rq_c2 = st.columns(2)
                    if rq_c1.button(f"Update Password for {rq['username']}", key=f"main_btn_set_{rq['username']}", type="primary", use_container_width=True):
                        if new_pw_for_rq:
                            am.update_password(rq['username'], new_pw_for_rq)
                            st.success(f"Password updated for {rq['username']}!")
                            time.sleep(1)
                            st.rerun()
                        else:
                            st.warning("Please enter a new password.")
                    if rq_c2.button(f"Dismiss Request", key=f"main_btn_dism_{rq['username']}", use_container_width=True):
                        am.dismiss_reset_request(rq['username'])
                        st.rerun()
                    st.divider()

        st.stop()

sample_path = os.path.join(TEMPLATE_DIR, "Vendor Sample sheet.xlsx")
qc_path = os.path.join(TEMPLATE_DIR, "QC exception Vlookup.xlsx")

if os.path.exists(sample_path):
    with open(sample_path, "rb") as f_samp:
        t_col1.download_button(
            label="📥 Download Sample Format Sheet",
            data=f_samp.read(),
            file_name="Vendor Sample sheet.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

if os.path.exists(qc_path):
    with open(qc_path, "rb") as f_qc:
        t_col2.download_button(
            label="📥 Download QC Exception Codes",
            data=f_qc.read(),
            file_name="QC exception Vlookup.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

st.write("")

qc_codes = load_qc_exception_codes()

# Initialize session state for persistent results (strictly per-browser session state)
if "processed_results" not in st.session_state:
    st.session_state.processed_results = None
    st.session_state.elapsed_time = None

if "auto_download_triggered" not in st.session_state:
    st.session_state.auto_download_triggered = False

# File Uploader
uploaded_files = st.file_uploader(
    "Upload Vendor Rate Sheet(s) (Max 5 files or ZIP Archive)",
    type=["xlsx", "xls", "csv", "zip"],
    accept_multiple_files=True,
    help="Supports .xlsx, .xls, .csv, or .zip archives with multiple sheets."
)

if uploaded_files:
    raw_files = []
    for uploaded in uploaded_files:
        if uploaded.name.lower().endswith(".zip"):
            try:
                with zipfile.ZipFile(uploaded) as z:
                    for filename in z.namelist():
                        if filename.lower().endswith(('.xlsx', '.xls', '.csv')) and not filename.startswith('__MACOSX'):
                            file_bytes = z.read(filename)
                            raw_files.append((os.path.basename(filename), io.BytesIO(file_bytes)))
            except Exception as e:
                st.error(f"Error reading ZIP file '{uploaded.name}': {e}")
        else:
            raw_files.append((uploaded.name, uploaded))

    if len(raw_files) > 5:
        st.warning(f"⚠️ Maximum 5 sheets allowed per batch. Processing first 5 files out of {len(raw_files)}.")
        raw_files = raw_files[:5]

    st.markdown("### 📝 Enter Vendor / Trunk Names")
    st.caption("Customize Vendor / Trunk name for each file. Live final filename preview is shown below:")

    vendor_names = {}
    for idx, (fname, f_content) in enumerate(raw_files):
        clean_stem = os.path.splitext(fname)[0]
        clean_stem = re.sub(r'_\d{8,}.*$', '', clean_stem).strip()
        default_vendor_name = f"Name - {clean_stem}"

        c1, c2 = st.columns([1.5, 2])
        v_name = c1.text_input(
            f"Vendor Name for File {idx+1} ({fname}):",
            value=default_vendor_name,
            key=f"v_name_input_{idx}"
        )
        vendor_names[idx] = v_name
        
        # Real-time live preview that updates instantly without session-state caching bugs
        clean_v_name = v_name.strip() if v_name else f"File {idx+1}"
        if 'a to z' in clean_v_name.lower():
            preview_filename = f"{clean_v_name} [Comments] {active_member}.xlsx"
        else:
            preview_filename = f"{clean_v_name} A to Z [Comments] {active_member}.xlsx"
            
        c2.markdown(f"""
        <div style="margin-top: 26px; padding: 8px 12px; background: #F8FAFC; border: 1px solid #CBD5E1; border-radius: 6px; font-family: monospace; font-size: 0.88rem; color: #1E293B; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">
            📄 <b>Preview:</b> <span style="color: #2563EB;">{preview_filename}</span>
        </div>
        """, unsafe_allow_html=True)

    st.write("")
    
    # User's explicit requested button name: Proceed for Download
    btn_label = f"⚡ Proceed for Download ({len(raw_files)} Sheets)" if len(raw_files) > 1 else "⚡ Proceed for Download"
    if st.button(btn_label, type="primary", use_container_width=True):
        t_start = time.time()
        processed_list = []
        progress_bar = st.progress(0)

        user_dl_dir = os.path.join(os.path.expanduser('~'), 'Downloads')
        kaif_dir = r"C:\Kaif\Standardized_Sheets"
        os.makedirs(kaif_dir, exist_ok=True)

        for f_idx, (fname, f_content) in enumerate(raw_files):
            f_content.seek(0)
            df_raw = None
            try:
                if fname.lower().endswith(('.xlsx', '.xlsm', '.xls')):
                    engine = 'xlrd' if fname.lower().endswith('.xls') else None
                    xl = pd.ExcelFile(f_content, engine=engine)
                    best_df = None
                    best_score = -1
                    for s_name in xl.sheet_names:
                        temp_df = xl.parse(s_name, header=None)
                        if len(temp_df) < 2:
                            continue
                        _, col_map = detect_header_row_and_columns(temp_df)
                        score = len(col_map)
                        if ('dial_code' in col_map or 'dial_code_split' in col_map) and 'rate' in col_map:
                            score += 15
                        if score > best_score and len(temp_df) >= 3:
                            best_score = score
                            best_df = temp_df
                    df_raw = best_df if best_df is not None else xl.parse(0, header=None)
                elif fname.lower().endswith('.csv'):
                    try:
                        df_raw = pd.read_csv(f_content, sep=None, engine='python', header=None)
                    except Exception:
                        f_content.seek(0)
                        try:
                            df_raw = pd.read_csv(f_content, sep=None, engine='python', header=None, encoding='latin-1')
                        except Exception:
                            f_content.seek(0)
                            df_raw = pd.read_csv(f_content, header=None, encoding='latin-1', low_memory=False)
            except Exception as e:
                st.error(f"Could not parse file '{fname}': {e}")
                continue

            if df_raw is None or len(df_raw) == 0:
                st.warning(f"File '{fname}' is empty.")
                continue

            df_clean, df_excluded, stats = fast_process_dataframe(df_raw, qc_codes, fname)

            user_base = vendor_names.get(f_idx, f"Name - {fname}")
            final_filename = build_output_filename(user_base, stats['comments_tag'], active_member)

            excel_bytes = generate_target_excel_bytes(df_clean)

            # Guaranteed Instant Auto-Save directly to user's Windows Downloads folder
            saved_paths = []
            try:
                dl_path = os.path.join(user_dl_dir, final_filename)
                with open(dl_path, "wb") as f_out:
                    f_out.write(excel_bytes)
                saved_paths.append(dl_path)
            except Exception:
                pass

            try:
                kaif_path = os.path.join(kaif_dir, final_filename)
                with open(kaif_path, "wb") as f_out:
                    f_out.write(excel_bytes)
                saved_paths.append(kaif_path)
            except Exception:
                pass

            # Guaranteed Save to OUTPUT_DIR
            cached_path = os.path.join(OUTPUT_DIR, final_filename)
            try:
                with open(cached_path, "wb") as f_out:
                    f_out.write(excel_bytes)
            except Exception:
                pass

            processed_list.append({
                "filename": final_filename,
                "vendor_name": user_base,
                "bytes": excel_bytes,
                "cached_path": cached_path,
                "stats": stats,
                "rows_count": len(df_clean),
                "saved_paths": saved_paths
            })

            progress_bar.progress((f_idx + 1) / len(raw_files))

        t_end = time.time()
        elapsed_sec = round(t_end - t_start, 2)

        # Log audit trail for Admin view
        for p_item in processed_list:
            am.log_processing_activity(
                username=current_user.get("username", "unknown"),
                user_name=current_user.get("name", "Unknown"),
                output_filename=p_item["filename"],
                vendor_name=p_item.get("vendor_name", ""),
                clean_rows=p_item["stats"].get("final_clean_rows", p_item.get("rows_count", 0)),
                qc_excluded=p_item["stats"].get("qc_excluded", 0),
                elapsed_sec=elapsed_sec
            )

        # Store in session state for this user's browser
        st.session_state.processed_results = processed_list
        st.session_state.elapsed_time = elapsed_sec
        st.rerun()

# ----------------- DISPLAY PERSISTENT RESULTS & DOWNLOAD BUTTONS -----------------
if st.session_state.processed_results:
    st.write("---")
    
    # Top Action Bar: Status Header + Clear All Button
    col_hdr, col_clr = st.columns([3, 1])
    col_hdr.markdown('<div class="badge-success">🎉 All Sheets Standardized Successfully! Click below to download:</div>', unsafe_allow_html=True)
    
    if col_clr.button("🗑️ Clear All / Reset", type="secondary", use_container_width=True):
        st.session_state.processed_results = None
        st.session_state.elapsed_time = None
        st.rerun()

    # Time Taken Display Banner (Exact requested feature)
    if st.session_state.get("elapsed_time") is not None:
        sec_val = st.session_state["elapsed_time"]
        st.markdown(f"""
        <div style="background: #F0FDF4; border: 1.5px solid #22C55E; border-radius: 8px; padding: 12px 18px; margin: 8px 0 16px 0; display: flex; align-items: center; justify-content: space-between;">
            <span style="font-size: 1.05rem; color: #15803D; font-weight: 600;">
                ⏱️ <b>Processed & Downloaded in:</b> <span style="font-size: 1.35rem; color: #166534; font-weight: 800;">{sec_val:.2f} seconds</span>
            </span>
            <span style="font-size: 0.85rem; color: #16A34A; background: #DCFCE7; padding: 5px 14px; border-radius: 20px; font-weight: 700;">
                ⚡ Ultra-Fast Execution
            </span>
        </div>
        """, unsafe_allow_html=True)

    results = st.session_state.processed_results

    # Render prominent download buttons and path info for each processed file
    for idx, item in enumerate(results):
        out_name = item["filename"]
        out_bytes = item.get("bytes")
        stats = item["stats"]
        saved_paths = item.get("saved_paths", [])

        main_path = saved_paths[0] if saved_paths else f"C:\\Users\\Windows 10 Pro\\Downloads\\{out_name}"

        # Multi-location fallback to guarantee bytes are never empty
        if not out_bytes:
            for possible_p in [
                item.get("cached_path"),
                main_path,
                os.path.join(OUTPUT_DIR, out_name),
                os.path.join(os.path.expanduser('~'), 'Downloads', out_name),
                os.path.join(kaif_dir, out_name)
            ]:
                if possible_p and os.path.exists(possible_p):
                    try:
                        with open(possible_p, "rb") as f_cached:
                            out_bytes = f_cached.read()
                        if out_bytes:
                            break
                    except Exception:
                        pass
        if not out_bytes:
            out_bytes = b""

        st.markdown(f"""
        <div class="auto-dl-banner">
            📁 <b>Standardized File:</b> <code style="color: #1E3A8A; font-size: 1rem;">{out_name}</code><br>
            📍 <b>Exact Saved Location:</b> <code style="color: #047857;">{main_path}</code><br>
            <span style="font-size: 0.9rem; color: #4B5563;">
                Clean Rows: <b>{stats['final_clean_rows']:,}</b> &nbsp;|&nbsp; 
                QC Excluded: <b>{stats['qc_excluded']:,}</b> &nbsp;|&nbsp; 
                USA/Canada/PR Excluded: <b>{stats['nanp_excluded']:,}</b> &nbsp;|&nbsp;
                Rates to -1: <b>{stats['rates_minus1_count']:,}</b>
            </span>
        </div>
        """, unsafe_allow_html=True)

        if hasattr(os, 'startfile'):
            col_dl, col_excel, col_folder = st.columns([1.5, 1, 1])
            col_dl.download_button(
                label=f"⬇️ Download",
                data=out_bytes,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"persistent_dl_{idx}",
                type="primary",
                use_container_width=True
            )

            if col_excel.button(f"📊 Open in Excel", key=f"open_excel_{idx}", use_container_width=True):
                try:
                    os.startfile(main_path)
                except Exception as e:
                    st.info(f"File path: {main_path}")

            if col_folder.button(f"📂 Open Folder", key=f"open_folder_{idx}", use_container_width=True):
                try:
                    os.startfile(os.path.dirname(main_path))
                except Exception as e:
                    st.info(f"Folder: {os.path.dirname(main_path)}")
        else:
            st.download_button(
                label=f"⬇️ Download {out_name}",
                data=out_bytes,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"persistent_dl_{idx}",
                type="primary",
                use_container_width=True
            )

    # If multiple files, provide a single ZIP download
    if len(results) > 1:
        st.write("")
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for item in results:
                f_name = item["filename"]
                f_bytes = item.get("bytes")
                if not f_bytes:
                    for possible_p in [
                        item.get("cached_path"),
                        item.get("saved_paths", [""])[0] if item.get("saved_paths") else "",
                        os.path.join(OUTPUT_DIR, f_name),
                        os.path.join(os.path.expanduser('~'), 'Downloads', f_name),
                        os.path.join(kaif_dir, f_name)
                    ]:
                        if possible_p and os.path.exists(possible_p):
                            try:
                                with open(possible_p, "rb") as f_b:
                                    f_bytes = f_b.read()
                                if f_bytes:
                                    break
                            except Exception:
                                pass
                if f_bytes:
                    zip_file.writestr(f_name, f_bytes)
        zip_buffer.seek(0)
        
        st.download_button(
            label=f"📦 Download All {len(results)} Standardized Files (ZIP)",
            data=zip_buffer.getvalue(),
            file_name=f"Standardized_Rate_Sheets_{active_member}.zip",
            mime="application/zip",
            key="persistent_zip_dl",
            type="primary",
            use_container_width=True
        )
