import os
import datetime
import io
import csv
import re
import gc
import time
import json
import base64
import zipfile
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import importlib
from streamlit_cookies_controller import CookieController

# Initialize controller early in script execution
controller = CookieController()

import auth_manager as am
import billing_manager as bm
try:
    importlib.reload(am)
    importlib.reload(bm)
except Exception:
    pass

# Bulletproof self-contained avatar resolution
def resolve_user_avatar(username):
    u = str(username or '').strip().lower()
    if "arhama" in u:
        return "avatar_niqab_classic.png"
    elif "kaif" in u:
        return "avatar_man_glasses.png"
    elif "aliba" in u:
        return "avatar_hijab_niqab_dark.png"
    # Assign pleasant deterministic avatar from pool
    h = sum(ord(c) for c in u) % len(SAFE_AVAILABLE_AVATARS)
    return SAFE_AVAILABLE_AVATARS[h]["id"]

SAFE_AVAILABLE_AVATARS = [
    # 🧕 15 Female & Hijab Avatars (Arhama, Aliba & Girls)
    {"id": "avatar_niqab_classic.png", "title": "1. Classic Black Niqab", "category": "Female & Hijab", "badge": ""},
    {"id": "avatar_hijab_niqab_dark.png", "title": "2. Dark Niqab Veil", "category": "Female & Hijab", "badge": ""},
    {"id": "avatar_hijab_bw_art.png", "title": "3. Floral Art Silhouette", "category": "Female & Hijab", "badge": ""},
    {"id": "avatar_hijab_black.png", "title": "4. Black Hijab", "category": "Female & Hijab", "badge": ""},
    {"id": "avatar_hijab_charcoal.png", "title": "5. Charcoal Hijab", "category": "Female & Hijab", "badge": ""},
    {"id": "avatar_hijab_gold.png", "title": "6. Gold Mustard Hijab", "category": "Female & Hijab", "badge": ""},
    {"id": "avatar_hijab_ruby_red.png", "title": "7. Ruby Red Hijab", "category": "Female & Hijab", "badge": ""},
    {"id": "avatar_hijab_shades_green.png", "title": "8. Green with Shades", "category": "Female & Hijab", "badge": ""},
    {"id": "avatar_hijab_coral_pink.png", "title": "9. Coral Pink Hijab", "category": "Female & Hijab", "badge": ""},
    {"id": "avatar_hijab_green_trim.png", "title": "10. Green Trim Hijab", "category": "Female & Hijab", "badge": ""},
    {"id": "avatar_hijab_silver_gray.png", "title": "11. Silver Gray Hijab", "category": "Female & Hijab", "badge": ""},
    {"id": "avatar_cartoon_curly_hair.png", "title": "12. Curly Hair Girl", "category": "Female & Hijab", "badge": ""},
    {"id": "avatar_cartoon_asian_woman.png", "title": "13. Yellow Top Specs", "category": "Female & Hijab", "badge": ""},
    {"id": "avatar_cartoon_short_hair_glasses.png", "title": "14. Short Hair Specs", "category": "Female & Hijab", "badge": ""},
    {"id": "avatar_cartoon_modern_woman.png", "title": "15. Modern Chic Girl", "category": "Female & Hijab", "badge": ""},

    # 🧔 6 Male & Executive Avatars (Kaif & Boys)
    {"id": "avatar_man_glasses.png", "title": "16. Bearded Specs (Kaif)", "category": "Male & Executive", "badge": ""},
    {"id": "avatar_cartoon_beard_suit.png", "title": "17. Suit & Beard", "category": "Male & Executive", "badge": ""},
    {"id": "avatar_cartoon_bearded_specs.png", "title": "18. Executive Specs", "category": "Male & Executive", "badge": ""},
    {"id": "avatar_cartoon_young_man.png", "title": "19. Young Casual", "category": "Male & Executive", "badge": ""},
    {"id": "avatar_cartoon_blonde_glasses.png", "title": "20. Blonde Specs", "category": "Male & Executive", "badge": ""},
    {"id": "avatar_cartoon_teen_brown_hair.png", "title": "21. Brown Hair Casual", "category": "Male & Executive", "badge": ""}
]

def save_user_avatar_choice(u_name, avatar_id):
    if hasattr(am, "update_user_avatar"):
        try:
            return am.update_user_avatar(u_name, avatar_id)
        except Exception:
            pass
    try:
        db_p = os.path.join(os.path.dirname(__file__), "users_db.json")
        if os.path.exists(db_p):
            with open(db_p, "r", encoding="utf-8") as f:
                d = json.load(f)
            u = str(u_name).strip().lower()
            if u in d.get("users", {}):
                d["users"][u]["avatar"] = avatar_id
                with open(db_p, "w", encoding="utf-8") as f:
                    json.dump(d, f, indent=2)
            return True, "Avatar updated."
    except Exception as e:
        return False, str(e)
    return True, "Avatar updated."
import core_engine
try:
    importlib.reload(core_engine)
except Exception:
    pass
import secrets
from config_manager import load_config, add_member, remove_member, set_active_member
from core_engine import (
    detect_header_row_and_columns,
    load_qc_exception_codes,
    fast_process_dataframe,
    generate_target_excel_bytes,
    generate_target_csv_bytes,
    build_output_filename,
    build_csv_output_filename,
    analyze_vendor_comments,
    TEMPLATE_DIR
)

# Output directory for reliable caching of generated Excel files
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "generated_output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

def read_disk_bytes(p):
    """Reads file bytes from disk on demand, preventing memory bloat in st.session_state."""
    if p and os.path.exists(p):
        try:
            with open(p, "rb") as f:
                return f.read()
        except Exception:
            pass
    return b""

def load_excel_safely(f_content, fname):
    """Safely loads Excel files (.xlsx, .xls, .xlsm, .xlsb) using calamine with xlrd/openpyxl fallback."""
    f_content.seek(0)
    try:
        return pd.ExcelFile(f_content, engine='calamine')
    except Exception:
        pass
    f_content.seek(0)
    try:
        engine = 'xlrd' if fname.lower().endswith('.xls') else None
        return pd.ExcelFile(f_content, engine=engine)
    except Exception:
        f_content.seek(0)
        return pd.ExcelFile(f_content)

def load_csv_safely(f_content, nrows=None):
    """
    Safely loads any CSV file, including ragged vendor sheets with multi-line metadata headers,
    inconsistent column counts, unusual delimiters (comma, semicolon, tab, pipe),
    multi-encoding fallbacks (utf-8, latin-1, cp1252), and European decimal commas.
    """
    f_content.seek(0)
    raw = f_content.read()
    text = ""
    if isinstance(raw, bytes):
        for enc in ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']:
            try:
                text = raw.decode(enc)
                break
            except Exception:
                continue
        if not text:
            text = raw.decode('latin-1', errors='replace')
    else:
        text = str(raw)

    f_content.seek(0)

    # Detect best delimiter by counting candidates across the first 10 non-empty lines
    lines = [l for l in text.splitlines() if l.strip()][:10]
    best_sep = ','
    if lines:
        counts = {sep: sum(l.count(sep) for l in lines) for sep in [';', '\t', '|', ',']}
        if counts[';'] > counts[','] and counts[';'] >= len(lines):
            best_sep = ';'
        elif counts['\t'] > counts[','] and counts['\t'] >= len(lines):
            best_sep = '\t'
        elif counts['|'] > counts[','] and counts['|'] >= len(lines):
            best_sep = '|'

    try:
        return pd.read_csv(io.StringIO(text), sep=best_sep, header=None, low_memory=False, nrows=nrows)
    except Exception:
        pass

    for sep in [best_sep, ',', ';', '\t', '|']:
        try:
            reader = csv.reader(io.StringIO(text), delimiter=sep)
            rows = []
            for i, r in enumerate(reader):
                if nrows and i >= nrows:
                    break
                rows.append(r)
            if rows:
                max_cols = max(len(r) for r in rows)
                if max_cols > 1:
                    padded = [r + [''] * (max_cols - len(r)) for r in rows]
                    return pd.DataFrame(padded)
        except Exception:
            pass

    reader = csv.reader(io.StringIO(text))
    rows = []
    for i, r in enumerate(reader):
        if nrows and i >= nrows:
            break
        rows.append(r)
    if not rows:
        return pd.DataFrame()
    max_cols = max(len(r) for r in rows)
    padded = [r + [''] * (max_cols - len(r)) for r in rows]
    return pd.DataFrame(padded)


# ----------------- SESSION STORE (TTL 8H, SERVER-SIDE + IN-MEMORY) -----------------
@st.cache_resource
def get_session_cache():
    """In-memory cache mapping sha256(token) -> {user, exp}. Fast lookups across runs."""
    return {}

# Strictly enforce authentication: session cookie + server store (no URL tokens)
def authenticate_current_session():
    if st.session_state.get("auth_logged_out"):
        return None

    # 1. Active memory user in current WebSocket session
    user = st.session_state.get("user")
    if user:
        return user

    # 2. Read tp_sid session cookie via WebSocket CookieController
    token = None
    try:
        token = controller.get("tp_sid")
    except Exception:
        pass

    # Fallback to HTTP context cookies if available
    if not token:
        try:
            if hasattr(st, "context") and hasattr(st.context, "cookies") and st.context.cookies:
                token = st.context.cookies.get("tp_sid") or st.context.cookies.get("kaif_auth_token")
        except Exception:
            pass

    if token:
        token_str = str(token).strip()
        cache = get_session_cache()
        th = am.hash_token(token_str)
        now = time.time()
        
        valid_user = None
        # Check fast in-memory cache
        if th in cache and cache[th].get("exp", 0) > now:
            valid_user = cache[th].get("user")
        else:
            # Fallback to server store
            valid_user = am.get_server_session_user(token_str)
            if valid_user:
                cache[th] = {"user": valid_user, "exp": now + 8 * 3600}

        if valid_user:
            st.session_state.user = valid_user
            st.session_state.active_token = token_str
            st.session_state.cookie_synced = True
            return valid_user

    # Never expose tokens in st.query_params - scrub any stray query params
    try:
        for k in ["tp_sid", "session_token", "_auth_sync"]:
            if k in st.query_params:
                del st.query_params[k]
    except Exception:
        pass

    return None

def set_login_session(session_data):
    st.session_state.user = session_data
    token = session_data.get("token")
    st.session_state.active_token = token
    st.session_state.cookie_synced = False
    st.session_state.auth_logged_out = False
    cache = get_session_cache()
    if token:
        cache[am.hash_token(token)] = {"user": session_data, "exp": time.time() + 8 * 3600}

def clear_current_session():
    tok = st.session_state.get("active_token")
    if not tok:
        try:
            tok = controller.get("tp_sid")
        except Exception:
            pass
    if tok:
        am.delete_server_session(tok)
        am.destroy_session(tok)
        cache = get_session_cache()
        th = am.hash_token(tok)
        if th in cache:
            del cache[th]
    try:
        controller.remove("tp_sid", path="/")
    except Exception:
        pass
    for z_key in ["batch_zip_xl_path", "batch_zip_csv_path"]:
        z_p = st.session_state.get(z_key)
        if z_p and os.path.exists(z_p):
            try:
                os.remove(z_p)
            except Exception:
                pass
        st.session_state[z_key] = None
    st.session_state.user = None
    st.session_state.active_token = None
    st.session_state.cookie_id = None
    st.session_state.cookie_synced = False
    st.session_state.pending_totp_user = None
    st.session_state.processed_results = None
    st.session_state.elapsed_time = None
    st.session_state.auth_logged_out = True
    gc.collect()
    try:
        for k in ["tp_sid", "session_token", "_auth_sync"]:
            if k in st.query_params:
                del st.query_params[k]
    except Exception:
        pass

def save_user_processed_cache(username, results, elapsed_time):
    try:
        cache_file = os.path.join(OUTPUT_DIR, f"last_results_{username}.json")
        data = {
            "elapsed_time": elapsed_time,
            "items": []
        }
        for r in results:
            data["items"].append({
                "filename": r["filename"],
                "original_filename": r.get("original_filename", r.get("vendor_name", "")),
                "csv_filename": r.get("csv_filename", ""),
                "vendor_name": r.get("vendor_name", ""),
                "cached_path": r.get("cached_path", ""),
                "cached_csv_path": r.get("cached_csv_path", ""),
                "stats": r.get("stats", {}),
                "rows_count": r.get("rows_count", 0),
                "saved_paths": r.get("saved_paths", [])
            })
        with open(cache_file, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except Exception:
        pass

def load_user_processed_cache(username):
    try:
        cache_file = os.path.join(OUTPUT_DIR, f"last_results_{username}.json")
        if os.path.exists(cache_file):
            with open(cache_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            items = []
            for item in data.get("items", []):
                p = item.get("cached_path")
                # Do NOT load file bytes into memory. Use read_disk_bytes on demand!
                if (p and os.path.exists(p)) or os.path.exists(os.path.join(OUTPUT_DIR, item.get("filename", ""))):
                    item["bytes"] = None
                    item["csv_bytes"] = None
                    items.append(item)
            if items:
                return items, data.get("elapsed_time", 0.0)
    except Exception:
        pass
    return None, None

def clear_user_processed_cache(username):
    try:
        cache_file = os.path.join(OUTPUT_DIR, f"last_results_{username}.json")
        if os.path.exists(cache_file):
            os.remove(cache_file)
        # Also clean any stale batch zip files for this user
        for fname in os.listdir(OUTPUT_DIR):
            if fname.startswith(f"batch_") and username in fname and fname.endswith(".zip"):
                try:
                    os.remove(os.path.join(OUTPUT_DIR, fname))
                except Exception:
                    pass
    except Exception:
        pass

st.set_page_config(
    page_title="Telecom Hub",
    page_icon="🌐",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling (Theme 4: Deep Ocean Teal & Frost - Cloud Terminal / Tech Precision)
st.markdown("""
<style>
    /* Complete True Full-Screen Layout */
    [data-testid="stAppViewContainer"] {
        width: 100% !important;
        max-width: 100vw !important;
        overflow-x: hidden !important;
    }
    [data-testid="stAppViewContainer"] > section.main {
        width: 100% !important;
        max-width: 100% !important;
        flex: 1 1 auto !important;
        padding-left: 0 !important;
    }
    .block-container,
    div[data-testid="stMainBlockContainer"],
    div.stMainBlockContainer {
        width: 100% !important;
        max-width: 100% !important;
        padding-top: 1.0rem !important;
        padding-bottom: 2rem !important;
        padding-left: 3.5rem !important;
        padding-right: 2rem !important;
    }
    .stApp {
        background-color: #F0F9FF !important;
    }
    .main-title {
        font-size: 1.85rem;
        font-weight: 700;
        color: #0C4A6E !important;
        margin-bottom: 0.15rem;
    }
    .sub-title {
        font-size: 0.92rem;
        color: #475569 !important;
        margin-bottom: 0.8rem;
    }
    .badge-success {
        background-color: #E0F2FE;
        color: #0369A1;
        padding: 8px 16px;
        border-radius: 6px;
        font-weight: 600;
        display: inline-block;
        margin: 8px 0;
        font-size: 1rem;
    }
    .auto-dl-banner {
        background-color: #E0F2FE;
        border-left: 4px solid #0E7490;
        padding: 12px 18px;
        color: #0C4A6E;
        font-weight: 600;
        border-radius: 4px;
        margin: 12px 0;
        font-size: 1rem;
    }
    /* Hide Streamlit default Deploy button, 3-dots menu, footer, and Manage App button */
    #MainMenu, [data-testid="stMainMenu"] {visibility: hidden !important; display: none !important;}
    footer {visibility: hidden !important; display: none !important;}
    .stDeployButton {display: none !important;}
    [data-testid="stToolbarActions"] {display: none !important; visibility: hidden !important;}
    [data-testid="stDecoration"] {display: none !important; visibility: hidden !important; height: 0 !important;}
    [data-testid="stStatusWidget"] {visibility: hidden !important; display: none !important;}
    [data-testid="manage-app-button"] {display: none !important; visibility: hidden !important;}
    [data-testid="stManageApp"] {display: none !important; visibility: hidden !important;}
    button:has([data-testid="stIconManageApp"]) {display: none !important; visibility: hidden !important;}
    div[class*="manageApp"], div[class*="ManageApp"] {display: none !important; visibility: hidden !important;}
    .viewerBadge_container__1QSob, div[class*="viewerBadge"] {display: none !important; visibility: hidden !important;}
    
    /* Eliminate Ghost Lines (Top Horizontal Line & Collapsed Sidebar Vertical Line) */
    header, header[data-testid="stHeader"] {
        background: transparent !important;
        border: none !important;
        border-bottom: none !important;
        box-shadow: none !important;
        outline: none !important;
        height: 48px !important;
        z-index: 999998 !important;
    }
    /* When Sidebar is Expanded */
    section[data-testid="stSidebar"][aria-expanded="true"],
    [data-testid="stSidebar"][aria-expanded="true"] {
        min-width: 260px !important;
        max-width: 280px !important;
        width: 270px !important;
        margin-left: 0px !important;
        transform: none !important;
        visibility: visible !important;
        display: block !important;
        background-color: #F8FAFC !important;
        border-right: 1.5px solid #E2E8F0 !important;
        z-index: 999999 !important;
    }

    /* When Sidebar is Collapsed - 100% OFF-SCREEN and HIDDEN */
    section[data-testid="stSidebar"][aria-expanded="false"],
    [data-testid="stSidebar"][aria-expanded="false"] {
        min-width: 0px !important;
        max-width: 0px !important;
        width: 0px !important;
        margin-left: -500px !important;
        transform: translateX(-500px) !important;
        border: none !important;
        border-right: none !important;
        box-shadow: none !important;
        visibility: hidden !important;
        display: none !important;
        overflow: hidden !important;
    }
    
    [data-testid="stToolbar"] {
        background: transparent !important;
        visibility: visible !important;
        display: flex !important;
    }

    /* Steady, Stylish Fixed Corner Sidebar Expand Button - Never Jumps */
    [data-testid="stExpandSidebarButton"],
    button[data-testid="stExpandSidebarButton"],
    header [data-testid="stExpandSidebarButton"] {
        display: inline-flex !important;
        visibility: visible !important;
        opacity: 1 !important;
        position: fixed !important;
        top: 10px !important;
        left: 12px !important;
        background: #0E7490 !important;
        color: #FFFFFF !important;
        border: 1.5px solid #0891B2 !important;
        border-radius: 8px !important;
        box-shadow: 0 2px 8px rgba(14, 116, 144, 0.3) !important;
        z-index: 99999999 !important;
        width: 36px !important;
        height: 36px !important;
        min-width: 36px !important;
        min-height: 36px !important;
        align-items: center !important;
        justify-content: center !important;
        cursor: pointer !important;
        pointer-events: auto !important;
        transition: background-color 0.15s ease, transform 0.15s ease !important;
    }
    [data-testid="stExpandSidebarButton"]:hover,
    button[data-testid="stExpandSidebarButton"]:hover {
        background: #0891B2 !important;
        transform: scale(1.04) !important;
    }
    [data-testid="stExpandSidebarButton"] svg,
    [data-testid="stExpandSidebarButton"] span {
        fill: #FFFFFF !important;
        stroke: #FFFFFF !important;
        color: #FFFFFF !important;
        width: 20px !important;
        height: 20px !important;
    }
    
    /* Sidebar Close Button Inside Expanded Sidebar */
    section[data-testid="stSidebar"] [data-testid="stSidebarCollapseButton"] {
        position: static !important;
        background: transparent !important;
        color: #0E7490 !important;
        box-shadow: none !important;
        border: none !important;
        cursor: pointer !important;
    }
    section[data-testid="stSidebar"] [data-testid="stSidebarCollapseButton"] svg {
        fill: #0E7490 !important;
        stroke: #0E7490 !important;
        color: #0E7490 !important;
    }

    /* Big, Bold, Premium Card-style Portal Buttons in Sidebar */
    section[data-testid="stSidebar"] div[role="radiogroup"] {
        display: flex !important;
        flex-direction: column !important;
        gap: 8px !important;
        background: transparent !important;
        border: none !important;
        padding: 0 !important;
        width: 100% !important;
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] label {
        display: flex !important;
        align-items: center !important;
        padding: 10px 14px !important;
        background: #FFFFFF !important;
        border: 1.5px solid #CBD5E1 !important;
        border-radius: 8px !important;
        cursor: pointer !important;
        font-weight: 700 !important;
        font-size: 0.90rem !important;
        color: #0F172A !important;
        transition: all 0.15s ease-in-out !important;
        box-shadow: 0 1px 3px rgba(0,0,0,0.03) !important;
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] label:hover {
        border-color: #0E7490 !important;
        background: #F0F9FF !important;
        color: #0E7490 !important;
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] label[data-checked="true"],
    section[data-testid="stSidebar"] div[role="radiogroup"] label:has(input:checked) {
        background: #0E7490 !important;
        border-color: #0891B2 !important;
        color: #FFFFFF !important;
        box-shadow: 0 2px 6px rgba(14, 116, 144, 0.25) !important;
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] label:has(input:checked) span,
    section[data-testid="stSidebar"] div[role="radiogroup"] label:has(input:checked) p {
        color: #FFFFFF !important;
        font-weight: 800 !important;
    }
    
    /* Segmented Pill Navigation Control */
    div[role="radiogroup"] {
        background: #E0F2FE !important;
        border: 1px solid #BAE6FD !important;
        border-radius: 8px !important;
        padding: 3px !important;
        display: inline-flex !important;
        gap: 3px !important;
        margin-bottom: 8px !important;
    }
    div[role="radiogroup"] input[type="radio"] {
        position: absolute !important;
        opacity: 0 !important;
        width: 1px !important;
        height: 1px !important;
        pointer-events: none !important;
    }
    div[role="radiogroup"] > label {
        padding: 6px 16px !important;
        border-radius: 6px !important;
        font-weight: 500 !important;
        font-size: 0.88rem !important;
        color: #475569 !important;
        cursor: pointer !important;
        transition: all 0.15s ease !important;
        margin: 0 !important;
        background: transparent !important;
    }
    div[role="radiogroup"] > label:has(input:focus-visible) {
        outline: 2px solid #0E7490 !important;
        outline-offset: 2px !important;
    }
    div[role="radiogroup"] > label:has(input:checked) {
        background: #FFFFFF !important;
        color: #0C4A6E !important;
        font-weight: 600 !important;
        box-shadow: 0 1px 3px rgba(14, 116, 144, 0.15) !important;
    }

    /* Clean SaaS Card Container Styling */
    [data-testid="stVerticalBlockBorderWrapper"] {
        background: #FFFFFF !important;
        border: 1px solid #BAE6FD !important;
        border-radius: 10px !important;
        box-shadow: 0 1px 4px rgba(12, 74, 110, 0.04) !important;
        padding: 16px 20px !important;
        margin-bottom: 12px !important;
    }

    /* Executive Deck Result Card & Structured KPI Strip */
    .deck-result-card {
        background: #FFFFFF !important;
        border: 1px solid #BAE6FD !important;
        border-radius: 10px !important;
        box-shadow: 0 1px 4px rgba(12, 74, 110, 0.04) !important;
        padding: 16px 20px !important;
        margin-bottom: 12px !important;
    }
    .deck-title-bar {
        display: flex !important;
        justify-content: space-between !important;
        align-items: center !important;
        flex-wrap: wrap !important;
        gap: 8px !important;
        margin-bottom: 12px !important;
        padding-bottom: 10px !important;
        border-bottom: 1px solid #E0F2FE !important;
    }
    .deck-name {
        font-size: 1.02rem !important;
        font-weight: 700 !important;
        color: #0C4A6E !important;
        word-break: break-all !important;
    }
    .deck-badge-eff {
        background: #FEF3C7 !important;
        color: #92400E !important;
        border: 1px solid #F59E0B !important;
        padding: 3px 10px !important;
        border-radius: 6px !important;
        font-weight: 700 !important;
        font-size: 0.82rem !important;
    }
    .kpi-strip {
        display: flex !important;
        flex-wrap: wrap !important;
        gap: 10px !important;
    }
    .kpi-chip {
        padding: 8px 14px !important;
        border-radius: 8px !important;
        display: flex !important;
        flex-direction: column !important;
        min-width: 110px !important;
        flex: 1 1 0px !important;
    }
    .kpi-label {
        font-size: 0.68rem !important;
        font-weight: 700 !important;
        letter-spacing: 0.05em !important;
        text-transform: uppercase !important;
        margin-bottom: 2px !important;
    }
    .kpi-val {
        font-size: 1.25rem !important;
        font-weight: 800 !important;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif !important;
        line-height: 1.2 !important;
    }
    .kpi-neutral {
        background: #F0F9FF !important;
        border: 1px solid #BAE6FD !important;
    }
    .kpi-neutral .kpi-label { color: #475569 !important; }
    .kpi-neutral .kpi-val { color: #0C4A6E !important; }

    .kpi-success {
        background: #F0FDF4 !important;
        border: 1px solid #BBF7D0 !important;
    }
    .kpi-success .kpi-label { color: #166534 !important; }
    .kpi-success .kpi-val { color: #15803D !important; }

    .kpi-danger {
        background: #FEF2F2 !important;
        border: 1px solid #FECACA !important;
    }
    .kpi-danger .kpi-label { color: #991B1B !important; }
    .kpi-danger .kpi-val { color: #DC2626 !important; }

    .kpi-warning {
        background: #FFFBEB !important;
        border: 1px solid #FDE68A !important;
    }
    .kpi-warning .kpi-label { color: #92400E !important; }
    .kpi-warning .kpi-val { color: #B45309 !important; }

    .kpi-subtle {
        background: #F0F9FF !important;
        border: 1px solid #BAE6FD !important;
    }
    .kpi-subtle .kpi-label { color: #475569 !important; }
    .kpi-subtle .kpi-val { color: #0C4A6E !important; }

    /* Subtle Compact Buttons for Top Profile: Logout & Avatar */
    div.st-key-top_logout_btn > button,
    div.st-key-btn_toggle_avatar_picker > button,
    button[key="top_logout_btn"],
    button[key="btn_toggle_avatar_picker"] {
        padding: 1px 8px !important;
        min-height: 22px !important;
        height: 22px !important;
        font-size: 0.68rem !important;
        font-weight: 600 !important;
        line-height: 1 !important;
        border-radius: 6px !important;
        border: 1px solid #CBD5E1 !important;
        background-color: #FFFFFF !important;
        color: #475569 !important;
        box-shadow: 0 1px 2px rgba(0,0,0,0.03) !important;
    }
    div.st-key-top_logout_btn > button:hover,
    button[key="top_logout_btn"]:hover {
        background-color: #FEF2F2 !important;
        border-color: #F87171 !important;
        color: #DC2626 !important;
    }
    div.st-key-btn_toggle_avatar_picker > button:hover,
    button[key="btn_toggle_avatar_picker"]:hover {
        background-color: #F0F9FF !important;
        border-color: #38BDF8 !important;
        color: #0284C7 !important;
    }

    /* Subtle professional micro-animations (pure animations, preserving exact site colors) */
    @keyframes fadeInUp {
        from {
            opacity: 0;
            transform: translateY(6px);
        }
        to {
            opacity: 1;
            transform: translateY(0);
        }
    }
    .auto-dl-banner, .badge-success, [data-testid="stDataFrame"], .stTabs {
        animation: fadeInUp 0.3s ease-out;
    }
    button[kind="primary"] {
        background-color: #0E7490 !important;
        color: #FFFFFF !important;
        border: 1px solid #0E7490 !important;
        transition: transform 0.15s ease, box-shadow 0.15s ease !important;
    }
    button[kind="primary"]:hover {
        background-color: #155E75 !important;
        color: #FFFFFF !important;
        transform: translateY(-1.5px) !important;
        box-shadow: 0 4px 12px rgba(14, 116, 144, 0.25) !important;
    }
    button[kind="secondary"] {
        border: 1px solid #7DD3FC !important;
        color: #0C4A6E !important;
        background: #FFFFFF !important;
        transition: transform 0.15s ease, box-shadow 0.15s ease !important;
    }
    button[kind="secondary"]:hover {
        background-color: #F0F9FF !important;
        border-color: #0E7490 !important;
        color: #083344 !important;
        transform: translateY(-1.5px) !important;
        box-shadow: 0 4px 10px rgba(14, 116, 144, 0.08) !important;
    }
    [data-testid="stFileUploaderDropzone"] {
        transition: transform 0.2s ease, border-color 0.2s ease, box-shadow 0.2s ease !important;
        background: #F0F9FF !important;
        border: 2px dashed #BAE6FD !important;
        border-radius: 8px !important;
    }
    [data-testid="stFileUploaderDropzone"]:hover {
        transform: translateY(-2px) !important;
        border-color: #0E7490 !important;
        box-shadow: 0 6px 14px rgba(14, 116, 144, 0.1) !important;
    }

    /* Responsive Adjustments for Mobile */
    @media (max-width: 768px) {
        .main-title {
            font-size: 1.45rem !important;
        }
        div[role="radiogroup"] {
            width: 100% !important;
            justify-content: center !important;
        }
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
    # Client-side session cookie management
    if st.session_state.get("auth_logged_out"):
        components.html("""
        <script>
        (function() {
            var sec = window.location.protocol === "https:" ? "; Secure" : "";
            try {
                parent.document.cookie = "tp_sid=; path=/" + sec + "; Max-Age=0; SameSite=Lax";
                parent.document.cookie = "kaif_auth_token=; path=/" + sec + "; Max-Age=0; SameSite=Lax";
            } catch(e) {
                document.cookie = "tp_sid=; path=/" + sec + "; Max-Age=0; SameSite=Lax";
                document.cookie = "kaif_auth_token=; path=/" + sec + "; Max-Age=0; SameSite=Lax";
            }
            try {
                parent.sessionStorage.removeItem("tp_sid");
                parent.sessionStorage.removeItem("kaif_auth_token");
            } catch(e) {}
            try {
                if (parent.location.search) {
                    parent.history.replaceState({}, '', parent.location.pathname);
                }
            } catch(e) {}
        })();
        </script>
        """, height=0)
        st.session_state.auth_logged_out = False

    # Render Login Page (Centered Box)
    login_c1, login_c2, login_c3 = st.columns([1, 1.8, 1])
    with login_c2:
        st.markdown("""
        <div style="background: white; padding: 26px 30px; border-radius: 12px; border: 1px solid #BAE6FD; box-shadow: 0 4px 6px -1px rgba(12, 74, 110, 0.06); margin-top: 30px; margin-bottom: 20px;">
            <div style="text-align: center; margin-bottom: 18px;">
                <div style="font-size: 2.3rem;">🔐</div>
                <div style="font-size: 1.55rem; font-weight: 700; color: #0C4A6E; margin-top: 4px;">Rate Standardizer Portal</div>
                <div style="font-size: 0.88rem; color: #475569;">Authorized access with Password + Authy 2FA.</div>
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
                        new_sid = session_data["token"]
                        set_login_session(session_data)
                        st.session_state.pending_totp_user = None
                        st.session_state.auth_logged_out = False
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
u_name = (current_user.get("username") or "kaif").strip().lower()
user_role = (current_user.get("role") or "user").strip().lower()

# 3-Role Hierarchy: Super Admin (Kaif), Admin, User
is_super_admin = (u_name == "kaif" or user_role == "super_admin")
is_admin = is_super_admin or (user_role == "admin")
is_user = True
active_member = current_user.get("name", "Kaif")

# Ensure persistent session cookie is saved in browser for F5 reload persistence (run ONCE per login)
active_tok = st.session_state.get("active_token")
if active_tok and not st.session_state.get("cookie_synced"):
    try:
        controller.set("tp_sid", active_tok, max_age=28800, path="/")
    except Exception:
        pass
    components.html(f"""
    <script>
    (function() {{
        var sec = window.location.protocol === "https:" ? "; Secure; SameSite=None" : "; SameSite=Lax";
        var cVal = "tp_sid={active_tok}; path=/; max-age=28800" + sec;
        try {{ parent.document.cookie = cVal; }} catch(e) {{ document.cookie = cVal; }}
        try {{ parent.sessionStorage.setItem("tp_sid", "{active_tok}"); }} catch(e) {{}}
    }})();
    </script>
    """, height=0)
    st.session_state["cookie_synced"] = True

def get_avatar_base64(filename):
    if not filename:
        filename = "avatar_man_glasses.png"
    p = os.path.join(os.path.dirname(__file__), "assets", "avatars", filename)
    if not os.path.exists(p):
        p = os.path.join(os.path.dirname(__file__), "assets", "avatars", "avatar_man_glasses.png")
    if os.path.exists(p):
        try:
            with open(p, "rb") as f:
                return base64.b64encode(f.read()).decode("utf-8")
        except Exception:
            return ""
    return ""


@st.cache_data(show_spinner=False)
def get_cached_template_file(filename):
    p = os.path.join(TEMPLATE_DIR, filename)
    if os.path.exists(p):
        try:
            with open(p, "rb") as f:
                return f.read()
        except Exception:
            pass
    return b""

sample_bytes = get_cached_template_file("Vendor Sample sheet.xlsx")
qc_bytes = get_cached_template_file("QC exception Vlookup.xlsx")
country_billing_bytes = get_cached_template_file("Country Billing.xlsx")
separate_codes_bytes = get_cached_template_file("Separate_codes_sample_ratesheet.xls")

def safe_get_master_excel_bytes():
    try:
        if hasattr(bm, "get_master_excel_bytes"):
            b = bm.get_master_excel_bytes()
            if b and len(b) > 1000:
                return b
    except Exception:
        pass
    for p in [
        os.path.join(os.path.dirname(__file__), "templates", "Country Billing.xlsx"),
        os.path.join(os.path.dirname(__file__), "billing_master.xlsx"),
        r"C:\vendor Rates\Country Billing.xlsx"
    ]:
        if os.path.exists(p):
            try:
                with open(p, "rb") as f:
                    c = f.read()
                    if c and len(c) > 1000:
                        return c
            except Exception:
                pass
    if country_billing_bytes and len(country_billing_bytes) > 1000:
        return country_billing_bytes
    try:
        if hasattr(bm, "load_master_billing"):
            data = bm.load_master_billing()
            raw_list = data.get("raw_list", []) if isinstance(data, dict) else []
            if raw_list:
                df = pd.DataFrame(raw_list)
                df = df.rename(columns={"dial_code": "Dial Code", "destination": "Summary", "minimum": "Minimum", "increment": "Increment"})
                cols = ["Dial Code", "Summary", "Minimum", "Increment"]
                df = df[[col for col in cols if col in df.columns]]
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                    df.to_excel(writer, index=False, sheet_name="Master Billing")
                return buf.getvalue()
    except Exception:
        pass
    return b""

def safe_get_cc_master_bytes():
    fn = "A-Z CC Rates.xlsx"
    for p in [
        os.path.join(os.path.dirname(__file__), "templates", "CC_Rates_Sheet_2026_02_01.xlsx"),
        os.path.join(os.path.dirname(__file__), "templates", "CC_Rates_Sheet.xlsx"),
        os.path.join(os.path.dirname(__file__), "templates", "A-Z CC Rates.xlsx"),
        r"C:\vendor Rates\CC_Rates_Sheet_2026_02_01.xlsx"
    ]:
        if os.path.exists(p):
            try:
                with open(p, "rb") as f:
                    c = f.read()
                    if c and len(c) > 500:
                        return c, fn
            except Exception:
                pass
    return None, fn

def safe_get_cli_master_bytes():
    fn = "A-Z CLI Rates.xlsx"
    for p in [
        os.path.join(os.path.dirname(__file__), "templates", "CLI_Rate_sheet_2026_02_01.xlsx"),
        os.path.join(os.path.dirname(__file__), "templates", "CLI_Rate_sheet.xlsx"),
        os.path.join(os.path.dirname(__file__), "templates", "A-Z CLI Rates.xlsx"),
        r"C:\vendor Rates\CLI_Rate_sheet_2026_02_01.xlsx"
    ]:
        if os.path.exists(p):
            try:
                with open(p, "rb") as f:
                    c = f.read()
                    if c and len(c) > 500:
                        return c, fn
            except Exception:
                pass
    return None, fn

# ----------------- SIDEBAR NAVIGATION (MOUNTED FIRST) -----------------
with st.sidebar:
    st.markdown("""
    <div style="padding: 4px 0 8px 0; border-bottom: 1.5px solid #E2E8F0; margin-bottom: 10px;">
        <div style="font-size: 0.68rem; font-weight: 800; text-transform: uppercase; letter-spacing: 0.08em; color: #64748B;">Navigation</div>
        <div style="font-size: 1.05rem; font-weight: 800; color: #0C4A6E;">🌐 Telecom Hub</div>
    </div>
    """, unsafe_allow_html=True)
    
    sidebar_portals = ["🏢 Vendor Portal", "👥 Customer Portal"]
    if is_super_admin:
        sidebar_portals.append("👑 Super Admin (Accounts)")
        
    if "current_portal_choice" not in st.session_state:
        st.session_state["current_portal_choice"] = "🏢 Vendor Portal"
        
    def sync_from_sidebar():
        st.session_state["current_portal_choice"] = st.session_state["app_active_portal_selector"]
        
    active_portal = st.radio(
        "Select Portal:",
        sidebar_portals,
        index=sidebar_portals.index(st.session_state["current_portal_choice"]) if st.session_state["current_portal_choice"] in sidebar_portals else 0,
        key="app_active_portal_selector",
        label_visibility="collapsed",
        on_change=sync_from_sidebar
    )
    st.session_state["current_portal_choice"] = active_portal
    
    st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)

    # 📁 Standard Formats / Standard A-Z Rates Expanders in Sidebar
    # Live synced whenever Admin replaces sheets!
    if active_portal in ["🏢 Vendor Portal", "👑 Super Admin (Accounts)"]:
        with st.expander("📁 Standard Formats ▾", expanded=False):
            st.caption("Standard Vendor Formats & Exceptions:")
            if sample_bytes:
                st.download_button(
                    label="📄 Vendor Sample Rate Sheet (.xlsx)",
                    data=sample_bytes,
                    file_name="Vendor Sample Format.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="sb_dl_vendor_sample",
                    use_container_width=True
                )
            if qc_bytes:
                st.download_button(
                    label="📑 QC Blocked Codes (.xlsx)",
                    data=qc_bytes,
                    file_name="QC Blocked.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="sb_dl_qc_blocked",
                    use_container_width=True
                )
            country_billing_live = safe_get_master_excel_bytes()
            if country_billing_live:
                st.download_button(
                    label="📊 Country Master Billing (.xlsx)",
                    data=country_billing_live,
                    file_name="Country Billing.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="sb_dl_country_billing",
                    use_container_width=True
                )
            if separate_codes_bytes:
                st.download_button(
                    label="📁 Separate Codes (.xls)",
                    data=separate_codes_bytes,
                    file_name="Separate_codes_sample_ratesheet.xls",
                    mime="application/vnd.ms-excel",
                    key="sb_dl_sep_codes",
                    use_container_width=True
                )

    if active_portal in ["👥 Customer Portal", "👑 Super Admin (Accounts)"]:
        with st.expander("📊 Standard A-Z Rates ▾", expanded=False):
            st.caption("Master Customer Rate Decks:")
            cc_b, cc_fn = safe_get_cc_master_bytes()
            if cc_b:
                st.download_button(
                    label="📊 CC Rates Sheet (.xlsx)",
                    data=cc_b,
                    file_name=cc_fn,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="sb_dl_cc_master",
                    use_container_width=True
                )
            cli_b, cli_fn = safe_get_cli_master_bytes()
            if cli_b:
                st.download_button(
                    label="📊 CLI Rates Sheet (.xlsx)",
                    data=cli_b,
                    file_name=cli_fn,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="sb_dl_cli_master",
                    use_container_width=True
                )
    
    st.markdown("---")
    
    # Bottom User Profile Card
    curr_u = (current_user.get("username") or "kaif").strip().lower()
    cur_u_av = (current_user.get("avatar") if isinstance(current_user, dict) else None) or resolve_user_avatar(curr_u)
    cur_av_b64 = get_avatar_base64(cur_u_av)
    role_display = "👑 Super Admin" if is_super_admin else ("🛡️ Admin" if is_admin else "👤 User")
    
    st.markdown(f"""
    <div style="background: #F8FAFC; border: 1.5px solid #E2E8F0; border-radius: 10px; padding: 10px 12px; margin-bottom: 8px;">
        <div style="display: flex; align-items: center; gap: 10px;">
            <img src="data:image/png;base64,{cur_av_b64}" style="width: 38px; height: 38px; border-radius: 8px; border: 1.5px solid #0E7490; object-fit: cover;" />
            <div style="overflow: hidden;">
                <div style="font-weight: 800; font-size: 0.90rem; color: #0F172A; white-space: nowrap; text-overflow: ellipsis; overflow: hidden;">{active_member}</div>
                <div style="font-size: 0.70rem; font-weight: 700; color: #0284C7;">{role_display}</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    col_sb_act1, col_sb_act2 = st.columns([1, 1])
    with col_sb_act1:
        if st.button("🎨 Avatar", key="sb_btn_toggle_avatar", use_container_width=True, help="Change your profile avatar"):
            st.session_state["show_avatar_picker"] = not st.session_state.get("show_avatar_picker", False)
            st.rerun()
    with col_sb_act2:
        if st.button("🚪 Logout", key="sb_btn_logout", use_container_width=True, help="Logout from portal"):
            clear_current_session()
            components.html("""
            <script>
            setTimeout(function() {
                try { parent.window.location.reload(); } catch(e) { window.location.reload(); }
            }, 150);
            </script>
            <div style="font-family: sans-serif; color: #0E7490; padding: 10px; text-align: center;">Logging out...</div>
            """, height=50)
            st.stop()

# ----------------- MAIN PORTAL HEADER (Clean, Full-Width, No Badges) -----------------
if active_portal == "👑 Super Admin (Accounts)":
    st.markdown('<div class="main-title">👑 Super Admin Control Panel</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">Master user management: View all IDs & passwords, edit display names, create/delete accounts, and security controls.</div>', unsafe_allow_html=True)
elif active_portal == "👥 Customer Portal":
    st.markdown('<div class="main-title">👥 Customer Rate Sheet Standardizer</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">Standardize customer rate decks against Master A-Z CC & CLI ratesheets with zero data loss.</div>', unsafe_allow_html=True)
else:
    st.markdown('<div class="main-title">⚡ Vendor Rate Sheet Standardizer</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">High-speed wholesale rate deck standardization, zero data loss, exact sample formatting.</div>', unsafe_allow_html=True)

# Expandable Profile Avatar Selector
if st.session_state.get("show_avatar_picker", False):
    with st.container():
        st.markdown("""
        <div style="background: #FFFFFF; border: 1.5px solid #BAE6FD; border-radius: 12px; padding: 14px 18px; margin: 10px 0 15px 0; box-shadow: 0 2px 8px rgba(14, 116, 144, 0.08);">
            <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #E0F2FE; padding-bottom: 8px; margin-bottom: 8px;">
                <div style="font-weight: 700; color: #0C4A6E; font-size: 0.98rem;">
                    🎨 Choose Your Profile Avatar (21 Curated Options)
                </div>
                <span style="font-size: 0.78rem; color: #0369A1; font-weight: 700; background: #E0F2FE; border: 1px solid #BAE6FD; padding: 4px 12px; border-radius: 14px;">
                    🧕 15 Female / Hijab &nbsp;|&nbsp; 🧔 6 Male
                </span>
            </div>
            <div style="font-size: 0.82rem; color: #475569;">
                Apne pasandida avatar ke niche <b>'Select'</b> click karein:
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        av_tab1, av_tab2, av_tab3 = st.tabs([
            "🧕 Female / Hijab & Girls (15 Avatars)",
            "🧔 Male & Executive / Boys (6 Avatars)",
            "🌟 All 21 Avatars (Single View)"
        ])
        
        with av_tab1:
            female_avatars = [a for a in SAFE_AVAILABLE_AVATARS if a["category"] == "Female & Hijab"]
            chunk_size = 5
            for i in range(0, len(female_avatars), chunk_size):
                chunk = female_avatars[i:i + chunk_size]
                cols = st.columns(len(chunk))
                for idx, av in enumerate(chunk):
                    with cols[idx]:
                        is_active = (av["id"] == cur_u_av)
                        border_col = "#0E7490" if is_active else "#BAE6FD"
                        bg_col = "#E0F2FE" if is_active else "#FFFFFF"
                        b64 = get_avatar_base64(av["id"])
                        badge_html = f"<div style='font-size: 0.65rem; color: #0E7490; font-weight: 700;'>★ Active</div>" if is_active else ""
                        st.markdown(f"""
                        <div style="text-align: center; background: {bg_col}; border: 1.5px solid {border_col}; border-radius: 10px; padding: 8px 4px; margin-bottom: 6px;">
                            <img src="data:image/png;base64,{b64}" style="width: 52px; height: 52px; border-radius: 8px; object-fit: cover;" />
                            <div style="font-size: 0.74rem; font-weight: 600; color: #0C4A6E; margin-top: 4px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="{av['title']}">
                                {av['title']}
                            </div>
                            {badge_html}
                        </div>
                        """, unsafe_allow_html=True)
                        if st.button("Select", key=f"sel_f_av_{av['id']}", use_container_width=True, type="primary" if is_active else "secondary"):
                            ok, msg = save_user_avatar_choice(curr_u, av["id"])
                            if ok:
                                if isinstance(current_user, dict):
                                    current_user["avatar"] = av["id"]
                                if "user" in st.session_state and isinstance(st.session_state.user, dict):
                                    st.session_state.user["avatar"] = av["id"]
                                st.toast(f"✅ Avatar updated to {av['title']}!")
                                st.session_state["show_avatar_picker"] = False
                                st.rerun()

        with av_tab2:
            male_avatars = [a for a in SAFE_AVAILABLE_AVATARS if a["category"] == "Male & Executive"]
            chunk_size = 6
            for i in range(0, len(male_avatars), chunk_size):
                chunk = male_avatars[i:i + chunk_size]
                cols = st.columns(len(chunk))
                for idx, av in enumerate(chunk):
                    with cols[idx]:
                        is_active = (av["id"] == cur_u_av)
                        border_col = "#0E7490" if is_active else "#BAE6FD"
                        bg_col = "#E0F2FE" if is_active else "#FFFFFF"
                        b64 = get_avatar_base64(av["id"])
                        badge_html = f"<div style='font-size: 0.65rem; color: #0E7490; font-weight: 700;'>★ Active</div>" if is_active else ""
                        st.markdown(f"""
                        <div style="text-align: center; background: {bg_col}; border: 1.5px solid {border_col}; border-radius: 10px; padding: 8px 4px; margin-bottom: 6px;">
                            <img src="data:image/png;base64,{b64}" style="width: 52px; height: 52px; border-radius: 8px; object-fit: cover;" />
                            <div style="font-size: 0.74rem; font-weight: 600; color: #0C4A6E; margin-top: 4px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="{av['title']}">
                                {av['title']}
                            </div>
                            {badge_html}
                        </div>
                        """, unsafe_allow_html=True)
                        if st.button("Select", key=f"sel_m_av_{av['id']}", use_container_width=True, type="primary" if is_active else "secondary"):
                            ok, msg = save_user_avatar_choice(curr_u, av["id"])
                            if ok:
                                if isinstance(current_user, dict):
                                    current_user["avatar"] = av["id"]
                                if "user" in st.session_state and isinstance(st.session_state.user, dict):
                                    st.session_state.user["avatar"] = av["id"]
                                st.toast(f"✅ Avatar updated to {av['title']}!")
                                st.session_state["show_avatar_picker"] = False
                                st.rerun()

        with av_tab3:
            st.caption("Showing all 21 avatars together:")
            chunk_size = 7
            for i in range(0, len(SAFE_AVAILABLE_AVATARS), chunk_size):
                chunk = SAFE_AVAILABLE_AVATARS[i:i + chunk_size]
                cols = st.columns(len(chunk))
                for idx, av in enumerate(chunk):
                    with cols[idx]:
                        is_active = (av["id"] == cur_u_av)
                        border_col = "#0E7490" if is_active else "#BAE6FD"
                        bg_col = "#E0F2FE" if is_active else "#FFFFFF"
                        b64 = get_avatar_base64(av["id"])
                        st.markdown(f"""
                        <div style="text-align: center; background: {bg_col}; border: 1.5px solid {border_col}; border-radius: 10px; padding: 6px 2px; margin-bottom: 6px;">
                            <img src="data:image/png;base64,{b64}" style="width: 48px; height: 48px; border-radius: 8px; object-fit: cover;" />
                            <div style="font-size: 0.70rem; font-weight: 600; color: #0C4A6E; margin-top: 2px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="{av['title']}">
                                {av['title']}
                            </div>
                        </div>
                        """, unsafe_allow_html=True)
                        if st.button("Select", key=f"sel_all_av_{av['id']}", use_container_width=True, type="primary" if is_active else "secondary"):
                            ok, msg = save_user_avatar_choice(curr_u, av["id"])
                            if ok:
                                if isinstance(current_user, dict):
                                    current_user["avatar"] = av["id"]
                                if "user" in st.session_state and isinstance(st.session_state.user, dict):
                                    st.session_state.user["avatar"] = av["id"]
                                st.toast(f"✅ Avatar updated to {av['title']}!")
                                st.session_state["show_avatar_picker"] = False
                                st.rerun()

        col_cls1, col_cls2 = st.columns([4, 1])
        with col_cls2:
            if st.button("✖ Close Selector", key="btn_close_av_picker", use_container_width=True):
                st.session_state["show_avatar_picker"] = False
                st.rerun()

def get_historical_file_bytes(filename):
    if not filename:
        return None
    p = os.path.join(OUTPUT_DIR, filename)
    if os.path.exists(p):
        try:
            with open(p, "rb") as f:
                return f.read()
        except Exception:
            pass
    dl_p = os.path.join(os.path.expanduser("~"), "Downloads", filename)
    if os.path.exists(dl_p):
        try:
            with open(dl_p, "rb") as f:
                return f.read()
        except Exception:
            pass
    return None

def get_historical_csv_bytes(filename, excel_bytes=None):
    if not filename:
        return None, ""
    base = filename.rsplit(".", 1)[0]
    csv_fname = base + ".csv"
    p_csv = os.path.join(OUTPUT_DIR, csv_fname)
    if os.path.exists(p_csv):
        try:
            with open(p_csv, "rb") as f:
                return f.read(), csv_fname
        except Exception:
            pass
    dl_csv = os.path.join(os.path.expanduser("~"), "Downloads", csv_fname)
    if os.path.exists(dl_csv):
        try:
            with open(dl_csv, "rb") as f:
                return f.read(), csv_fname
        except Exception:
            pass
    if excel_bytes:
        try:
            df = pd.read_excel(io.BytesIO(excel_bytes), engine="openpyxl")
            buf = io.StringIO()
            df.to_csv(buf, index=False)
            return buf.getvalue().encode("utf-8"), csv_fname
        except Exception:
            pass
    return None, csv_fname


def render_activity_history_tab(prefix="vend"):
    st.subheader("📜 User Processing History & Download Center")
    st.caption("Complete tracking of every rate sheet standardized by all team members with 1-click Excel and CSV downloads.")
    
    history = am.get_processing_history()
    if not history:
        st.info("ℹ️ No rate sheets processed yet. As users process files, complete activity records will appear here in real time.")
        return

    def extract_entry_date(ts):
        if not ts:
            return None
        try:
            s = str(ts).strip()[:10]
            return datetime.datetime.strptime(s, "%Y-%m-%d").date()
        except Exception:
            return None

    now_ist = datetime.datetime.utcnow() + datetime.timedelta(hours=5, minutes=30)
    today = now_ist.date()
    all_dates = [extract_entry_date(h.get("timestamp")) for h in history]
    valid_dates = [d for d in all_dates if d is not None]
    min_available_date = min(valid_dates) if valid_dates else today

    cutoff_7d = today - datetime.timedelta(days=7)
    sheets_7d = [h for h in history if extract_entry_date(h.get("timestamp")) and extract_entry_date(h.get("timestamp")) >= cutoff_7d]
    total_sheets = len(sheets_7d)
    arhama_total = sum(1 for h in sheets_7d if "arhama" in (str(h.get("username", "")) + str(h.get("user_name", "")) + str(h.get("output_filename", ""))).lower())
    kaif_total = sum(1 for h in sheets_7d if "kaif" in (str(h.get("username", "")) + str(h.get("user_name", "")) + str(h.get("output_filename", ""))).lower())
    aliba_total = sum(1 for h in sheets_7d if "aliba" in (str(h.get("username", "")) + str(h.get("user_name", "")) + str(h.get("output_filename", ""))).lower())

    m1, m2, m3, m4, m5 = st.columns([1.5, 1.2, 1.2, 1.2, 1])
    m1.metric("📁 Rolling 7-Day Sheets", f"{total_sheets:,}")
    m2.metric("👩 Arhama", f"{arhama_total}")
    m3.metric("👨 Kaif", f"{kaif_total}")
    m4.metric("👩 Aliba", f"{aliba_total}")
    with m5:
        st.write("")
        if st.button("🗑️ Clear Log", key=f"{prefix}_clear_hist_btn", use_container_width=True):
            am.clear_processing_history()
            st.success("Activity log cleared!")
            time.sleep(0.8)
            st.rerun()

    st.info("🕒 **Strict 7-Day Auto-Retention Active:** All rate sheets and audit records are saved for exactly 7 days (168 hours) from upload, after which they are automatically deleted.")
    st.divider()

    # Date Filter Mode Selector
    st.markdown("**📅 Filter History by Date:**")
    date_filter_mode = st.radio(
        "Choose Date View:",
        ["⚡ Last 7 Days (Default)", "🌐 All Time", "📆 Date Range"],
        horizontal=True,
        key=f"{prefix}_hist_date_filter_mode"
    )

    sel_start_date = None
    sel_end_date = None

    if date_filter_mode == "⚡ Last 7 Days (Default)":
        sel_start_date = today - datetime.timedelta(days=7)
        sel_end_date = today

    elif date_filter_mode == "📆 Date Range":
        col_dr1, col_dr2 = st.columns(2)
        with col_dr1:
            from_d = st.date_input("From Date:", value=(min_available_date if min_available_date else (today - datetime.timedelta(days=7))), key=f"{prefix}_hist_from_date_input")
        with col_dr2:
            to_d = st.date_input("To Date:", value=today, key=f"{prefix}_hist_to_date_input")
        sel_start_date = min(from_d, to_d)
        sel_end_date = max(from_d, to_d)

    elif date_filter_mode == "🌐 All Time":
        sel_start_date = None
        sel_end_date = None

    # Secondary Filters: User and Keyword Search
    fc1, fc2 = st.columns([1.5, 2.5])
    with fc1:
        user_options = ["All Users", "Arhama", "Kaif", "Aliba"]
        selected_user = st.selectbox("Filter by User:", user_options, key=f"{prefix}_audit_filter_user")
    with fc2:
        filter_kw = st.text_input("Search by Vendor or File Name:", key=f"{prefix}_audit_search_kw", placeholder="Type vendor or filename to search...")

    # Filtering logic
    filtered = []
    for entry in history:
        e_date = extract_entry_date(entry.get("timestamp"))
        
        if sel_start_date and sel_end_date:
            if not e_date or not (sel_start_date <= e_date <= sel_end_date):
                continue
        
        if selected_user != "All Users":
            u_match = (str(entry.get("user_name", "")) + " " + str(entry.get("username", ""))).lower()
            f_match = str(entry.get("output_filename", "")).lower()
            sel_low = selected_user.lower()
            if sel_low not in u_match and sel_low not in f_match:
                continue
        
        if filter_kw:
            kw_low = filter_kw.strip().lower()
            v_name = str(entry.get("vendor_name", "")).lower()
            out_f = str(entry.get("output_filename", "")).lower()
            u_f = (str(entry.get("user_name", "")) + " " + str(entry.get("username", ""))).lower()
            if kw_low not in v_name and kw_low not in out_f and kw_low not in u_f:
                continue
                
        filtered.append(entry)

    st.markdown(f"**Showing {len(filtered)} Rate Sheet(s) Matching Filter:**")

    if not filtered:
        st.warning("⚠️ No processing records found for the selected date / filter criteria.")
    else:
        # 1-Click Instant Sheet Download
        st.markdown("#### ⚡ 1-Click Instant Sheet Download")
        sheet_options = [h["output_filename"] for h in filtered if h.get("output_filename")]
        if sheet_options:
            sel_sheet = st.selectbox("Select any Rate Sheet to download immediately:", sheet_options, key=f"{prefix}_admin_quick_dl_select")
            if sel_sheet:
                xl_bytes = get_historical_file_bytes(sel_sheet)
                csv_b, csv_fn = get_historical_csv_bytes(sel_sheet, xl_bytes)
                
                c_dl1, c_dl2 = st.columns(2)
                with c_dl1:
                    if xl_bytes:
                        st.download_button(
                            label=f"📊 Download Excel (.xlsx)",
                            data=xl_bytes,
                            file_name=sel_sheet,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"{prefix}_admin_hist_quick_dl_xl_{sel_sheet}",
                            type="primary",
                            use_container_width=True
                        )
                    else:
                        st.button(f"Excel File Not in Cache", disabled=True, use_container_width=True, key=f"{prefix}_btn_no_xl_{sel_sheet}")
                with c_dl2:
                    if csv_b:
                        st.download_button(
                            label=f"📄 Download CSV (.csv)",
                            data=csv_b,
                            file_name=csv_fn,
                            mime="text/csv",
                            key=f"{prefix}_admin_hist_quick_dl_csv_{sel_sheet}",
                            use_container_width=True
                        )
                    else:
                        st.button(f"CSV Not in Cache", disabled=True, use_container_width=True, key=f"{prefix}_btn_no_csv_{sel_sheet}")

        st.markdown("---")
        st.markdown("#### 📜 Complete Activity & History Table")

        table_rows = []
        for entry in filtered:
            u_disp = entry.get("user_name") or entry.get("username") or "Unknown"
            ts = str(entry.get("timestamp", ""))
            elapsed = entry.get("elapsed_sec", 0)

            time_disp = ""
            if entry.get("time_range"):
                time_disp = entry.get("time_range")
            elif entry.get("start_time") and entry.get("end_time"):
                time_disp = f"{entry.get('start_time')} - {entry.get('end_time')} ({elapsed}s)"
            elif len(ts) >= 19:
                try:
                    end_dt = datetime.datetime.strptime(ts[:19], "%Y-%m-%d %H:%M:%S")
                    start_dt = end_dt - datetime.timedelta(seconds=float(elapsed))
                    time_disp = f"{start_dt.strftime('%H:%M:%S')} - {end_dt.strftime('%H:%M:%S')} ({elapsed}s)"
                except Exception:
                    time_disp = ts[11:19]
            else:
                time_disp = ts

            clean_fn = re.sub(r'\bupdated\b', 'Updated', str(entry.get("output_filename", "")), flags=re.IGNORECASE)
            table_rows.append({
                "Date": ts[:10] if len(ts) >= 10 else "",
                "Time (Kitne Baje Se Kitne Baje Tak)": time_disp,
                "User Name": u_disp,
                "Vendor / Trunk Name": entry.get("vendor_name"),
                "Standardized File": clean_fn,
                "Clean Rows": f"{entry.get('clean_rows', 0):,}",
                "Blocked Codes": f"{entry.get('qc_excluded', 0):,}",
                "Duration": f"{elapsed}s"
            })
        
        df_audit = pd.DataFrame(table_rows)
        st.dataframe(df_audit, use_container_width=True, hide_index=True)

        csv_data = df_audit.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Download This History (CSV)",
            data=csv_data,
            file_name=f"audit_history_{today.strftime('%Y%m%d')}.csv",
            mime="text/csv",
            key=f"{prefix}_btn_download_audit_csv"
        )

def render_billing_rules_tab():
    st.subheader("👑 Master Billing Rules Manager (`Country Billing.xlsx`)")
    st.caption("View, search, edit, and add official company billing rules for any destination. When market rules change (e.g. Indonesia 60/1 -> 60/60), update them here without touching code.")
    
    bm_data = bm.load_master_billing()
    raw_rules = bm_data.get("raw_list", [])
    
    b_sub1, b_sub2, b_sub3 = st.tabs(["🔍 Search & Edit Rule", "➕ Add New Destination Rule", "📁 Bulk Upload New Sheet"])
    
    with b_sub1:
        st.markdown(f"**Total Master Rules:** `{len(raw_rules):,}` dial codes & destinations")
        q_dest = st.text_input("Search Dial Code or Destination (e.g. 852, 9370, Hong Kong, Indonesia):", key="admin_search_billing_q", placeholder="Type dial code or destination name...")
        
        matches = []
        if q_dest and q_dest.strip():
            q_low = q_dest.strip().lower()
            matches = [r for r in raw_rules if q_low in str(r.get("destination", "")).lower() or q_low in str(r.get("dial_code", "")).lower()]
        else:
            matches = raw_rules[:25]
            
        if matches:
            st.markdown(f"Found **{len(matches):,}** destination(s):")
            match_display = [f"[{m.get('dial_code', '')}] {m.get('destination', '')}" if m.get('dial_code') else m.get('destination', '') for m in matches]
            sel_to_edit = st.selectbox("Select Rule to Edit:", match_display, key="admin_sel_billing_edit")
            sel_idx = match_display.index(sel_to_edit) if sel_to_edit in match_display else 0
            curr_rule = matches[sel_idx]
            
            if curr_rule:
                b_ed1, b_ed2, b_ed3 = st.columns([2, 1, 1])
                with b_ed1:
                    st.info(f"**Code:** `{curr_rule.get('dial_code', '')}` | **Destination:** `{curr_rule['destination']}` | **Current:** `{curr_rule['minimum']} / {curr_rule['increment']}`")
                with b_ed2:
                    new_min_val = st.number_input("Minimum (sec):", min_value=1, max_value=300, value=int(curr_rule['minimum']), key=f"b_input_min_{sel_to_edit}")
                with b_ed3:
                    new_inc_val = st.number_input("Increment (sec):", min_value=1, max_value=300, value=int(curr_rule['increment']), key=f"b_input_inc_{sel_to_edit}")
                    
                b_btn_c1, b_btn_c2 = st.columns([1, 1])
                with b_btn_c1:
                    if st.button(f"💾 Save Changes ({new_min_val}/{new_inc_val})", type="primary", use_container_width=True, key=f"b_btn_save_{sel_to_edit}"):
                        ok_u, msg_u = bm.update_or_add_master_rule(curr_rule.get('dial_code', ''), curr_rule['destination'], new_min_val, new_inc_val)
                        if ok_u:
                            st.success(msg_u)
                            time.sleep(1)
                            st.rerun()
                        else:
                            st.error(msg_u)
                with b_btn_c2:
                    if st.button("🗑️ Delete Rule", type="secondary", use_container_width=True, key=f"b_btn_del_{sel_to_edit}"):
                        ok_d, msg_d = bm.delete_master_rule(curr_rule.get('dial_code', '') or curr_rule['destination'])
                        st.success(msg_d)
                        time.sleep(1)
                        st.rerun()
    
    with b_sub2:
        st.markdown("#### ➕ Add New Dial Code / Destination Rule")
        with st.form("admin_add_billing_form", clear_on_submit=True):
            add_col_c1, add_col_c2 = st.columns([1, 2])
            with add_col_c1:
                add_code = st.text_input("Dial Code (Optional):", placeholder="e.g. 8529")
            with add_col_c2:
                add_d_name = st.text_input("Destination Name:", placeholder="e.g. HONG KONG MOBILE CSL")
            col_a1, col_a2 = st.columns(2)
            with col_a1:
                add_min = st.number_input("Minimum (sec):", min_value=1, max_value=300, value=1)
            with col_a2:
                add_inc = st.number_input("Increment (sec):", min_value=1, max_value=300, value=1)
            if st.form_submit_button("➕ Add to Master Rules", type="primary"):
                if (add_d_name and add_d_name.strip()) or (add_code and add_code.strip()):
                    ok_a, msg_a = bm.update_or_add_master_rule(add_code.strip(), add_d_name.strip(), add_min, add_inc)
                    if ok_a:
                        st.success(msg_a)
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.error(msg_a)
                else:
                    st.warning("Dial Code or Destination name cannot be empty.")

    with b_sub3:
        st.markdown("#### 📁 Bulk Upload Updated `Country Billing.xlsx`")
        st.caption("Upload your new Excel file with 4 columns (Dial Code, Summary, Minimum, Increment) or 3 columns (Destinations, Minimum, Increment) to replace all master rules in 1 second.")
        up_b_file = st.file_uploader("Choose Excel file (.xlsx):", type=["xlsx"], key="admin_billing_uploader")
        if up_b_file is not None:
            if st.button("⚡ Replace All Rules with Uploaded Sheet", type="primary", key="admin_btn_apply_billing_file"):
                ok_f, msg_f = bm.save_new_master_excel(up_b_file)
                if ok_f:
                    st.success(msg_f)
                    time.sleep(1)
                    st.rerun()
                else:
                    st.error(msg_f)
        
        st.divider()
        admin_master_bytes = safe_get_master_excel_bytes()
        if admin_master_bytes:
            st.download_button(
                label="📥 Download Current Master Rules (.xlsx)",
                data=admin_master_bytes,
                file_name="Country Billing.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="admin_btn_dl_master_billing"
            )

def render_customer_master_az_tab():
    st.subheader("👑 Master A-Z Sheet (CC & CLI Manager)")
    st.caption("Manage, download, and dynamically update Master A-Z CC and CLI rate sheets without modifying any application code.")
    
    cc_bytes, cc_filename = safe_get_cc_master_bytes()
    cli_bytes, cli_filename = safe_get_cli_master_bytes()
    
    c_m1, c_m2 = st.columns(2)
    with c_m1:
        st.markdown(f"""
        <div style="background: #F0FDF4; border: 1.5px solid #86EFAC; border-radius: 10px; padding: 16px; margin-bottom: 14px;">
            <div style="font-weight: 800; color: #166534; font-size: 1.0rem;">📊 Master CC Rates Sheet</div>
            <div style="color: #15803D; font-size: 0.82rem; margin: 4px 0 8px 0;">Current active file: <code>{cc_filename}</code> ({len(cc_bytes)/1024:.1f} KB)</div>
            <div style="color: #475569; font-size: 0.78rem;">Live reference deck used for Call Center (CC) destinations and pricing.</div>
        </div>
        """, unsafe_allow_html=True)
        if cc_bytes:
            st.download_button(
                label=f"📥 Download Current CC Rates Sheet (.xlsx)",
                data=cc_bytes,
                file_name=cc_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="cust_az_dl_cc_btn",
                type="primary",
                use_container_width=True
            )
        else:
            st.warning("CC Rates Sheet not found.")
            
        st.markdown("##### 🔄 Replace Master CC Rates Sheet")
        up_cc = st.file_uploader("Upload new CC Rates Sheet (.xlsx):", type=["xlsx"], key="cust_up_cc_sheet")
        if up_cc is not None:
            if st.button("⚡ Save & Apply New CC Rates Sheet", type="primary", key="cust_apply_cc_btn", use_container_width=True):
                target_p = os.path.join(os.path.dirname(__file__), "templates", "CC_Rates_Sheet_2026_02_01.xlsx")
                with open(target_p, "wb") as f:
                    f.write(up_cc.getvalue())
                st.success("✅ Master CC Rates Sheet updated successfully! New sheet is now active.")
                time.sleep(1)
                st.rerun()

    with c_m2:
        st.markdown(f"""
        <div style="background: #F0F9FF; border: 1.5px solid #BAE6FD; border-radius: 10px; padding: 16px; margin-bottom: 14px;">
            <div style="font-weight: 800; color: #0C4A6E; font-size: 1.0rem;">📊 Master CLI Rates Sheet</div>
            <div style="color: #0369A1; font-size: 0.82rem; margin: 4px 0 8px 0;">Current active file: <code>{cli_filename}</code> ({len(cli_bytes)/1024:.1f} KB)</div>
            <div style="color: #475569; font-size: 0.78rem;">Live reference deck used for CLI / Pure conversational routes.</div>
        </div>
        """, unsafe_allow_html=True)
        if cli_bytes:
            st.download_button(
                label=f"📥 Download Current CLI Rates Sheet (.xlsx)",
                data=cli_bytes,
                file_name=cli_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="cust_az_dl_cli_btn",
                type="primary",
                use_container_width=True
            )
        else:
            st.warning("CLI Rates Sheet not found.")

        st.markdown("##### 🔄 Replace Master CLI Rates Sheet")
        up_cli = st.file_uploader("Upload new CLI Rates Sheet (.xlsx):", type=["xlsx"], key="cust_up_cli_sheet")
        if up_cli is not None:
            if st.button("⚡ Save & Apply New CLI Rates Sheet", type="primary", key="cust_apply_cli_btn", use_container_width=True):
                target_p = os.path.join(os.path.dirname(__file__), "templates", "CLI_Rate_sheet_2026_02_01.xlsx")
                with open(target_p, "wb") as f:
                    f.write(up_cli.getvalue())
                st.success("✅ Master CLI Rates Sheet updated successfully! New sheet is now active.")
                time.sleep(1)
                st.rerun()

def render_super_admin_accounts_tabs():
    sa_tab1, sa_tab_edit, sa_tab2, sa_tab_del, sa_tab3, sa_tab4, sa_tab_email = st.tabs([
        "👥 All Users (IDs & Passwords)",
        "✏️ Edit User Names & Roles",
        "➕ Create New User",
        "🗑️ Delete User",
        "🔑 Reset Passwords",
        "🔔 Pending Reset Requests",
        "📧 Auto-Email Dispatcher (Setup)"
    ])
    
    with sa_tab1:
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
                "Avatar": uv.get("avatar", ""),
                "Authy Key": uv.get("totp_secret", am.get_master_totp_secret()),
                "Created At": uv.get("created_at")
            })
        st.dataframe(pd.DataFrame(u_data), use_container_width=True, hide_index=True)
        
        act_otp = getattr(am, 'get_last_active_otp', lambda: None)()
        if act_otp:
            st.info(f"📧 **Latest Active Email OTP:** Code: `{act_otp['otp']}` for User: `{act_otp['username']}` (Sent to: {act_otp['target_email']})")

        deletable_u = [u for u in all_u.keys() if u.lower() != "kaif"]
        if deletable_u:
            st.divider()
            st.markdown("**🗑️ Quick User Delete:**")
            q_del_c1, q_del_c2 = st.columns([3, 1])
            with q_del_c1:
                q_sel = st.selectbox("Select user account to delete:", deletable_u, key="sa_quick_del_user_sel")
            with q_del_c2:
                st.write("")
                if st.button("🗑️ Delete This User", type="secondary", use_container_width=True, key=f"sa_btn_quick_del_{q_sel}"):
                    ok_qd, msg_qd = am.delete_user(q_sel)
                    if ok_qd:
                        st.success(msg_qd)
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.error(msg_qd)

    with sa_tab_edit:
        st.subheader("✏️ Edit User Display Name & Avatar")
        st.caption("Change the full name and assign profile avatar for Admin (yourself) or any user. Changes reflect immediately across the portal.")
        all_users_dict = am.get_all_users()
        selected_edit_u = st.selectbox("Select User Account to Edit:", list(all_users_dict.keys()), key="sa_edit_u_select")
        u_info = all_users_dict[selected_edit_u]
        
        cur_u_av = u_info.get("avatar") or resolve_user_avatar(selected_edit_u)
        cur_av_b64 = get_avatar_base64(cur_u_av)
        
        e_col1, e_col2 = st.columns([1, 4])
        with e_col1:
            st.markdown(f"""
            <div style="text-align: center; padding: 10px; background: #F0F9FF; border: 1.5px solid #BAE6FD; border-radius: 10px;">
                <img src="data:image/png;base64,{cur_av_b64}" style="width: 55px; height: 55px; border-radius: 8px; object-fit: cover; border: 1px solid #0E7490;" />
                <div style="font-size: 0.70rem; color: #0C4A6E; font-weight: 600; margin-top: 4px;">Current Avatar</div>
            </div>
            """, unsafe_allow_html=True)
        with e_col2:
            st.info(f"**Username:** `{selected_edit_u}` | **Current Full Name:** `{u_info.get('name')}` | **Role:** `{u_info.get('role', 'user').upper()}`")
        
        ed_c1, ed_c2 = st.columns(2)
        with ed_c1:
            new_name_val = st.text_input("New Full Name / Display Name:", value=u_info.get("name", ""), key=f"sa_input_edit_name_{selected_edit_u}")
        with ed_c2:
            role_options = ["user", "admin", "super_admin"]
            role_labels = {"super_admin": "👑 Super Admin", "admin": "🛡️ Admin", "user": "👤 User"}
            current_r = u_info.get("role", "user")
            r_idx = role_options.index(current_r) if current_r in role_options else 0
            new_role_val = st.selectbox("Assign Role:", role_options, index=r_idx, format_func=lambda x: role_labels.get(x, x), key=f"sa_sel_edit_role_{selected_edit_u}")
        
        av_opts = [a["id"] for a in SAFE_AVAILABLE_AVATARS]
        av_dict = {a["id"]: f"{a['title']} ({a['category']})" for a in SAFE_AVAILABLE_AVATARS}
        cur_idx = av_opts.index(cur_u_av) if cur_u_av in av_opts else 0
        new_av_val = st.selectbox("Assign Profile Avatar:", av_opts, index=cur_idx, format_func=lambda x: av_dict.get(x, x), key=f"sa_sel_edit_av_{selected_edit_u}")
        
        if st.button("💾 Save Updated User Details", type="primary", use_container_width=True, key=f"sa_btn_save_name_{selected_edit_u}"):
            if new_name_val and new_name_val.strip():
                ok_n, msg_n = am.update_user_name(selected_edit_u, new_name_val.strip())
                ok_a, msg_a = save_user_avatar_choice(selected_edit_u, new_av_val)
                if hasattr(am, "update_user_role"):
                    am.update_user_role(selected_edit_u, new_role_val)
                if ok_n and ok_a:
                    if selected_edit_u == current_user.get("username"):
                        current_user["name"] = new_name_val.strip()
                        current_user["avatar"] = new_av_val
                        current_user["role"] = new_role_val
                        if "user" in st.session_state and isinstance(st.session_state.user, dict):
                            st.session_state.user["name"] = new_name_val.strip()
                            st.session_state.user["avatar"] = new_av_val
                            st.session_state.user["role"] = new_role_val
                    st.success(f"Details updated successfully for '{selected_edit_u}'!")
                    time.sleep(1)
                    st.rerun()
                else:
                    st.error(msg_n or msg_a)
            else:
                st.warning("Name cannot be empty.")

    with sa_tab2:
        st.subheader("➕ Create New User")
        st.caption(f"ℹ️ New user account will automatically use Admin's Master Authy 2FA (`{am.get_master_totp_secret()}`).")
        with st.form("sa_create_user_form", clear_on_submit=False):
            cu_col1, cu_col2 = st.columns(2)
            with cu_col1:
                nu_uname = st.text_input("Username", key="sa_nu_uname", placeholder="e.g. newuser")
                nu_name = st.text_input("Full Name", key="sa_nu_name", placeholder="e.g. User Name")
            with cu_col2:
                nu_pw = st.text_input("Password (Set Directly)", key="sa_nu_pw", placeholder="e.g. User@123")
                nu_role = st.selectbox("Role", ["user", "admin", "super_admin"], format_func=lambda x: {"super_admin": "👑 Super Admin", "admin": "🛡️ Admin", "user": "👤 User"}.get(x, x), key="sa_nu_role")
            nu_email = st.text_input("Email (Optional)", key="sa_nu_email")
            
            av_opts = [a["id"] for a in SAFE_AVAILABLE_AVATARS]
            av_dict = {a["id"]: f"{a['title']} ({a['category']})" for a in SAFE_AVAILABLE_AVATARS}
            nu_avatar = st.selectbox("Initial Profile Avatar:", av_opts, format_func=lambda x: av_dict.get(x, x), key="sa_nu_avatar")
            
            submitted = st.form_submit_button("🚀 Create User Account Now", type="primary", use_container_width=True)
            if submitted:
                clean_u = (nu_uname or "").strip().lower()
                clean_p = (nu_pw or "").strip()
                clean_n = (nu_name or "").strip()
                if clean_u and clean_p:
                    ok_cu, msg_cu = am.create_user(clean_u, clean_n, clean_p, nu_role, nu_email.strip())
                    if ok_cu:
                        save_user_avatar_choice(clean_u, nu_avatar)
                        st.success(f"✅ User '{clean_u}' created successfully with role '{nu_role}'!")
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.error(msg_cu)
                else:
                    st.warning("⚠️ Both Username and Password are required.")

    with sa_tab_del:
        st.subheader("🗑️ Delete User Account")
        st.caption("Permanently delete a registered user account. The Master Admin account ('kaif') is protected and cannot be deleted.")
        all_u_del = am.get_all_users()
        deletable_users = [u for u in all_u_del.keys() if u.lower() != "kaif"]
        if not deletable_users:
            st.info("ℹ️ No other user accounts available to delete. Only Master Admin ('kaif') exists.")
        else:
            del_c1, del_c2 = st.columns([2, 1])
            with del_c1:
                sel_del_u = st.selectbox("Select User Account to Delete:", deletable_users, key="sa_sel_del_user")
                target_info = all_u_del[sel_del_u]
                st.warning(f"⚠️ **Confirm Deletion:** You are about to permanently delete **{sel_del_u}** (Name: `{target_info.get('name')}`, Role: `{target_info.get('role', 'user').upper()}`). This cannot be undone.")
            with del_c2:
                st.write("")
                st.write("")
                if st.button(f"🗑️ Delete User '{sel_del_u}'", type="secondary", use_container_width=True, key=f"sa_btn_exec_del_{sel_del_u}"):
                    ok_del, msg_del = am.delete_user(sel_del_u)
                    if ok_del:
                        st.success(msg_del)
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.error(msg_del)

    with sa_tab3:
        st.subheader("🔑 Reset Any User's Password")
        all_u = am.get_all_users()
        sel_u = st.selectbox("Select User Account:", list(all_u.keys()), key="sa_sel_reset_u")
        st.info(f"Selected: **{sel_u}** | Current Password: `{all_u[sel_u].get('password', '')}` | Authy Key: `{all_u[sel_u].get('totp_secret', am.get_master_totp_secret())}`")
        new_pw_main = st.text_input("Enter New Password to Set:", key=f"sa_new_pw_input_{sel_u}")
        if st.button("💾 Save New Password", type="primary", use_container_width=True, key=f"sa_btn_save_pw_{sel_u}"):
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

    with sa_tab4:
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
                new_pw_for_rq = st.text_input(f"New Password for {rq['username']}:", key=f"sa_rq_pw_{rq['username']}")
                rq_c1, rq_c2 = st.columns(2)
                if rq_c1.button(f"Update Password for {rq['username']}", key=f"sa_btn_set_{rq['username']}", type="primary", use_container_width=True):
                    if new_pw_for_rq:
                        am.update_password(rq['username'], new_pw_for_rq)
                        st.success(f"Password updated for {rq['username']}!")
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.warning("Please enter a new password.")
                if rq_c2.button(f"Dismiss Request", key=f"sa_btn_dism_{rq['username']}", use_container_width=True):
                    am.dismiss_reset_request(rq['username'])
                    st.rerun()
                st.divider()

    with sa_tab_email:
        st.subheader("📧 Automated Rate Sheet Email Dispatcher")
        st.caption("Configure automated email dispatch so that every standardized rate sheet is instantly emailed upon successful processing.")
        
        st.markdown("""
        <div style="background: #FFFBEB; border: 1.5px solid #FDE68A; border-radius: 10px; padding: 12px 16px; margin-bottom: 14px;">
            <div style="font-weight: 700; color: #92400E; font-size: 0.88rem;">🟡 Status: Upper Setup UI Active (Backend Auto-Email Hook in Next Phase)</div>
            <div style="font-size: 0.78rem; color: #78350F; margin-top: 3px;">
                Jab bhi Kaif, Arhama ya Aliba koi bhi rate sheet successfully process karenge, system automated email ke sath finalized Excel file attach karke notification bhej dega. Local history save karne ka magajmari khatam!
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        with st.container(border=True):
            st.markdown("**⚙️ Email Dispatcher Configuration:**")
            em_c1, em_c2 = st.columns(2)
            with em_c1:
                st.text_input("SMTP Server Host:", value="smtp.gmail.com", disabled=True, key="sa_em_smtp_host")
                st.text_input("Sender Email Address:", value="kaifofficial4@gmail.com", key="sa_em_sender")
                st.checkbox("Auto-attach processed Excel file (.xlsx)", value=True, key="sa_em_attach_xlsx")
            with em_c2:
                st.text_input("SMTP Port:", value="587 (TLS)", disabled=True, key="sa_em_smtp_port")
                st.text_input("Default Notification Recipient(s):", value="kaifofficial4@gmail.com", help="Separate multiple emails with commas", key="sa_em_recipients")
                st.checkbox("Send instant alert upon processing errors / mismatches", value=True, key="sa_em_alert_err")
                
            st.info("💡 **Kaif's Workflow:** Processed rate sheets will dispatch directly to your inbox automatically.")
            if st.button("💾 Save Email Dispatcher Settings", type="primary", key="sa_btn_save_email_cfg"):
                st.success("✅ Email Dispatcher settings recorded! SMTP pipeline will activate on schedule.")


# ==============================================================================
# 👥 CUSTOMER RATE DECK STANDARDIZER ENGINE
# ==============================================================================

MASTER_TRUNK_ENTITIES = [
    {"code": "MCM", "name": "🏛️ My Country Mobile (MCM)"},
    {"code": "ACE", "name": "🏛️ Ace Peak Investment (ACE)"},
    {"code": "TECH", "name": "🏛️ Techknowledge Open Systems (TECH)"},
    {"code": "TWICH", "name": "🏛️ Twiching General Trading (TWICH)"},
    {"code": "SOFT", "name": "🏛️ Softtop Limited (SOFT)"},
    {"code": "LETS", "name": "🏛️ Lets Dial SG (LETS)"},
    {"code": "AJX", "name": "🏛️ Ajoxi Limited (AJX)"},
    {"code": "MERA", "name": "🏛️ MeraTalk LLC (MERA)"},
    {"code": "TEL", "name": "🏛️ Teloz Limited (TEL)"},
    {"code": "ROZ", "name": "🏛️ Rozper Limited (ROZ)"},
]

def format_telecom_rate(val):
    """
    Rounds telecom rates to at most 6 decimal places.
    Leaves rates with fewer decimals (e.g. 0.0054, 0.01) intact without unwanted padding,
    and strips IEEE floating point artifacts (e.g. 0.14650000000000002 -> 0.1465).
    """
    try:
        import pandas as pd
        if pd.isna(val) or str(val).strip() in ['', 'nan', 'none', 'null']:
            return ''
        v_clean = str(val).replace('$', '').replace(',', '.').strip()
        f_val = float(v_clean)
        if f_val == -1:
            return '-1'
        r = round(f_val, 6)
        s = f'{r:.6f}'.rstrip('0').rstrip('.')
        if s == '':
            return '0'
        return s
    except Exception:
        return str(val).strip()


_MASTER_SHEETS_CACHE = {}

def get_master_reference_df(source="CC"):
    global _MASTER_SHEETS_CACHE
    src_key = str(source).upper().strip()
    if src_key in _MASTER_SHEETS_CACHE:
        return _MASTER_SHEETS_CACHE[src_key]

    import os
    import pandas as pd
    if src_key == "CC":
        p = os.path.join(os.path.dirname(__file__), "templates", "CC_Rates_Sheet_2026_02_01.xlsx")
    else:
        p = os.path.join(os.path.dirname(__file__), "templates", "CLI_Rate_sheet_2026_02_01.xlsx")

    if not os.path.exists(p):
        p = r"C:/Vendor Rates/" + os.path.basename(p)

    if os.path.exists(p):
        try:
            df = pd.read_excel(p, dtype=str)
            _MASTER_SHEETS_CACHE[src_key] = df
            return df
        except Exception:
            pass
    return pd.DataFrame()

def import_destination_from_master(target_dest, scope="All", new_rate_str="0.0", source="CC", eff_date_str="2026-02-01", existing_dial_codes=None):
    import re
    import pandas as pd

    df_m = get_master_reference_df(source)
    if df_m.empty:
        return pd.DataFrame(), 0

    dest_clean = str(target_dest).strip()
    dest_escaped = re.escape(dest_clean)

    mask = df_m['Summary'].astype(str).str.contains(dest_escaped, case=False, na=False)
    if mask.sum() == 0:
        words = [w for w in re.split(r'[\s\-_]+', dest_clean) if len(w) > 1]
        if len(words) > 1:
            mask = df_m['Summary'].astype(str).apply(lambda s: all(w.lower() in s.lower() for w in words))
    if mask.sum() == 0:
        mask = df_m['Dial Code'].astype(str).str.startswith(dest_clean, na=False)

    if mask.sum() == 0:
        return pd.DataFrame(), 0

    df_sub = df_m[mask].copy()

    mobile_pattern = r'\b(?:mobile|cell|mob|cellular|wireless)\b'
    if scope in ["Fixed", "Fixed Only", "☎️ Fixed Only"]:
        df_sub = df_sub[~df_sub['Summary'].astype(str).str.contains(mobile_pattern, case=False, na=False)]
    elif scope in ["Mobile", "Mobile Only", "📱 Mobile Only"]:
        df_sub = df_sub[df_sub['Summary'].astype(str).str.contains(mobile_pattern, case=False, na=False)]

    nanp_pattern = r'\b(?:USA|UNITED STATES|CANADA|PUERTO RICO)\b'
    df_sub = df_sub[~df_sub['Summary'].astype(str).str.contains(nanp_pattern, case=False, na=False)]

    if df_sub.empty:
        return pd.DataFrame(), 0

    res_df = pd.DataFrame()
    res_df['Dial Code'] = df_sub['Dial Code'].astype(str).str.strip()
    res_df['Summary'] = df_sub['Summary'].astype(str).str.strip()
    res_df['New Rate'] = format_telecom_rate(new_rate_str)
    res_df['Minimum'] = df_sub['Minimum'].astype(str).str.strip() if 'Minimum' in df_sub.columns else '1'
    res_df['Increment'] = df_sub['Increment'].astype(str).str.strip() if 'Increment' in df_sub.columns else '1'
    res_df['Effective Date'] = str(eff_date_str)

    # STRICT DEDUPLICATION: Exclude any dial code already present in customer sheet!
    if existing_dial_codes:
        existing_set = set(str(c).strip() for c in existing_dial_codes)
        res_df = res_df[~res_df['Dial Code'].isin(existing_set)]

    # Deduplicate newly added rows themselves
    res_df = res_df.drop_duplicates(subset=['Dial Code'], keep='first').reset_index(drop=True)
    return res_df, len(res_df)


_master_dest_cache = None
def get_master_destinations_data():
    global _master_dest_cache
    if _master_dest_cache is not None:
        return _master_dest_cache
    json_path = os.path.join(os.path.dirname(__file__), "templates", "master_destinations.json")
    if os.path.exists(json_path):
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                _master_dest_cache = json.load(f)
                return _master_dest_cache
        except Exception:
            pass
    return {"destinations": [], "details": {}}

def smart_spell_check_destination(query, top_n=3):
    import difflib
    q = str(query or "").strip().lower()
    if not q or len(q) < 2:
        return []
    data = get_master_destinations_data()
    all_dests = data.get("destinations", [])
    if not all_dests:
        return []

    words = [w for w in re.split(r'[\s\-_]+', q) if len(w) > 1]
    
    # Check if exact token match already works
    if len(words) > 1:
        exact_tok = [d for d in all_dests if all(w in d.lower() for w in words)]
        if exact_tok:
            return exact_tok[:top_n]
    elif len(words) == 1:
        exact_sub = [d for d in all_dests if words[0] in d.lower()]
        if exact_sub:
            return exact_sub[:top_n]

    all_words = {}
    for d in all_dests:
        for w in re.split(r'[\s\-_]+', d):
            wl = w.lower()
            if len(wl) > 2:
                if wl not in all_words:
                    all_words[wl] = []
                all_words[wl].append(d)

    token_candidates = []
    for qw in words:
        close_w = difflib.get_close_matches(qw, all_words.keys(), n=3, cutoff=0.6)
        for cw in close_w:
            token_candidates.extend(all_words[cw])

    if token_candidates:
        scored = {}
        for d in token_candidates:
            d_lower = d.lower()
            score = 0
            for qw in words:
                if qw in d_lower:
                    score += 2
                else:
                    d_tokens = re.split(r'[\s\-_]+', d_lower)
                    if difflib.get_close_matches(qw, d_tokens, cutoff=0.6):
                        score += 1.5
            scored[d] = score
        sorted_dests = sorted(scored.keys(), key=lambda x: scored[x], reverse=True)
        return sorted_dests[:top_n]
    return []

def generate_customer_excel_bytes(df_clean, sheet_name="Rate_Amendment"):
    """
    Generates target Excel file matching 'Vendor Sample sheet.xlsx' exactly:
    - Sheet Name: 'Rate_Amendment'
    - Headers: ['Dial Code', 'Summary', 'New Rate', 'Eff. Date', 'Minimum', 'Increment']
    - Header fill: Silver/Grey (indexed 22 / D9D9D9)
    - Header font: Calibri, 11pt, Bold, black
    - Alignment: Center, Center
    - Eff. Date: Native datetime.date with number_format = 'mm-dd-yy'
    - AutoFilter: Enabled across A1:F...
    - Column widths matching Vendor Sample Sheet exactly
    """
    import io, gc, datetime
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.styles.colors import Color

    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name

    headers = ['Dial Code', 'Summary', 'New Rate', 'Eff. Date', 'Minimum', 'Increment']
    ws.append(headers)

    fill_22 = PatternFill(fill_type='solid', start_color='D9D9D9', end_color='D9D9D9')
    font_hdr = Font(name='Calibri', size=11, bold=True, color='000000')
    align_hdr = Alignment(horizontal='center', vertical='center')
    thin_border = Border(
        left=Side(style='thin', color='BFBFBF'),
        right=Side(style='thin', color='BFBFBF'),
        top=Side(style='thin', color='BFBFBF'),
        bottom=Side(style='thin', color='BFBFBF')
    )

    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = fill_22
        cell.font = font_hdr
        cell.alignment = align_hdr
        cell.border = thin_border

    codes = df_clean['Dial Code'].tolist() if 'Dial Code' in df_clean.columns else [''] * len(df_clean)
    sums = df_clean['Summary'].tolist() if 'Summary' in df_clean.columns else [''] * len(df_clean)
    rates = df_clean['New Rate'].tolist() if 'New Rate' in df_clean.columns else [''] * len(df_clean)
    
    date_col = 'Eff. Date' if 'Eff. Date' in df_clean.columns else ('Effective Date' if 'Effective Date' in df_clean.columns else None)
    raw_dates = df_clean[date_col].tolist() if date_col else [None] * len(df_clean)
    mins = df_clean['Minimum'].tolist() if 'Minimum' in df_clean.columns else [1] * len(df_clean)
    incs = df_clean['Increment'].tolist() if 'Increment' in df_clean.columns else [1] * len(df_clean)

    for r_idx in range(len(df_clean)):
        c = str(codes[r_idx]).split('.')[0].strip() if pd.notna(codes[r_idx]) else ''
        s = str(sums[r_idx]).strip() if pd.notna(sums[r_idx]) else ''
        try:
            r = float(rates[r_idx])
        except Exception:
            r = 0.0

        d_val = raw_dates[r_idx]
        d_obj = None
        if pd.notna(d_val) and d_val:
            if isinstance(d_val, (datetime.date, datetime.datetime)):
                d_obj = d_val.date() if isinstance(d_val, datetime.datetime) else d_val
            else:
                d_str = str(d_val).split('T')[0].split(' ')[0].strip()
                for fmt in ['%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y', '%Y/%m/%d']:
                    try:
                        d_obj = datetime.datetime.strptime(d_str, fmt).date()
                        break
                    except Exception:
                        pass
        if d_obj is None:
            d_obj = datetime.date.today()

        try:
            m = int(float(mins[r_idx])) if pd.notna(mins[r_idx]) else 1
        except Exception:
            m = 1
        try:
            inc = int(float(incs[r_idx])) if pd.notna(incs[r_idx]) else 1
        except Exception:
            inc = 1

        row_cells = [c, s, r, d_obj, m, inc]
        ws.append(row_cells)

        row_num = r_idx + 2
        ws.cell(row=row_num, column=1).alignment = Alignment(horizontal='left', vertical='center')
        ws.cell(row=row_num, column=2).alignment = Alignment(horizontal='left', vertical='center')
        ws.cell(row=row_num, column=3).alignment = Alignment(horizontal='right', vertical='center')
        ws.cell(row=row_num, column=4).alignment = Alignment(horizontal='center', vertical='center')
        ws.cell(row=row_num, column=4).number_format = 'mm-dd-yy'
        ws.cell(row=row_num, column=5).alignment = Alignment(horizontal='center', vertical='center')
        ws.cell(row=row_num, column=6).alignment = Alignment(horizontal='center', vertical='center')

    # Enable Excel AutoFilter
    ws.auto_filter.ref = f"A1:F{len(df_clean) + 1}"

    # Auto-fit column widths matching Vendor Sample Sheet exactly
    ws.column_dimensions['A'].width = 11.0
    ws.column_dimensions['B'].width = 44.0
    ws.column_dimensions['C'].width = 14.5
    ws.column_dimensions['D'].width = 13.0
    ws.column_dimensions['E'].width = 10.0
    ws.column_dimensions['F'].width = 10.5

    wb.save(output)
    output.seek(0)
    data = output.getvalue()
    try:
        wb.close()
    except Exception:
        pass
    del wb, ws
    gc.collect()
    return data

def process_customer_rate_deck(file_bytes, original_filename, dest_updates, effective_date_val, current_username="kaif"):
    import io, re, datetime, time
    import pandas as pd
    import billing_manager as bm
    import core_engine

    t0 = time.time()
    try:
        if str(original_filename).lower().endswith(".csv"):
            df = load_csv_safely(io.BytesIO(file_bytes))
        else:
            df = pd.read_excel(io.BytesIO(file_bytes), dtype=str)
    except Exception as e:
        return False, f"Could not parse sheet: {e}", None

    # Promote header row if loaded without header (e.g. from load_csv_safely)
    if all(isinstance(c, int) for c in df.columns) or any(str(c).isdigit() for c in df.columns):
        for r_idx in range(min(5, len(df))):
            row_vals = [str(v).strip().lower() for v in df.iloc[r_idx].values]
            if any(k in row_vals for k in ['rate key', 'dial code', 'dialcode', 'code', 'prefix']) or any(k in row_vals for k in ['rate summary', 'rate sumr', 'rate sum', 'ratesummary', 'ratesumr', 'summary', 'destination']):
                df.columns = [str(c).strip() for c in df.iloc[r_idx].values]
                df = df.iloc[r_idx + 1:].copy().reset_index(drop=True)
                break

    initial_rows = len(df)
    
    # 1. Flexible column normalization with RouteTrust Switch Export Priority
    col_map = {}
    
    # Priority Pass: Exact RouteTrust signatures
    for c in df.columns:
        cl = str(c).strip().lower()
        if cl in ['rate key', 'ratekey', 'rate_key']:
            col_map[c] = 'Dial Code'
        elif cl in ['rate summary', 'rate sumr', 'ratesummary', 'ratesumr', 'rate_summary', 'rate_sumr', 'rate sum', 'rates sum']:
            col_map[c] = 'Summary'
        elif cl in ['inter. rate', 'inter rate', 'inter_rate', 'international rate', 'international_rate', 'inter.rate', 'interrate']:
            col_map[c] = 'New Rate'
        elif cl in ['dur min', 'dur_min', 'durmin', 'dur_min.', 'dur min.', 'duration min', 'dur. min']:
            col_map[c] = 'Minimum'
        elif cl in ['dur inc', 'dur_inc', 'durinc', 'dur_inc.', 'dur inc.', 'duration inc', 'dur. inc']:
            col_map[c] = 'Increment'
        elif cl in ['start date', 'start_date', 'startdate', 'start_date.', 'start date.']:
            col_map[c] = 'Effective Date'

    # Second Pass: Standard telecom keyword matching for remaining columns
    for c in df.columns:
        if c in col_map:
            continue
        cl = str(c).strip().lower()
        if any(x in cl for x in ['dial code', 'dialcode', 'prefix', 'code']) and 'Dial Code' not in col_map.values():
            col_map[c] = 'Dial Code'
        elif any(x in cl for x in ['summary', 'destination', 'country', 'description']) and 'center' not in cl and 'id' not in cl and 'Summary' not in col_map.values():
            col_map[c] = 'Summary'
        elif any(x in cl for x in ['new rate', 'current rate', 'tariff', 'price']) and 'New Rate' not in col_map.values():
            col_map[c] = 'New Rate'
        elif cl == 'rate' and 'intra' not in cl and 'indet' not in cl and 'New Rate' not in col_map.values():
            col_map[c] = 'New Rate'
        elif any(x in cl for x in ['minimum', 'pulse1', 'min']) and 'Minimum' not in col_map.values():
            col_map[c] = 'Minimum'
        elif any(x in cl for x in ['increment', 'interval', 'pulse2', 'inc']) and 'Increment' not in col_map.values():
            col_map[c] = 'Increment'
        elif any(x in cl for x in ['effective date', 'effective', 'date']) and 'end date' not in cl and 'Effective Date' not in col_map.values():
            col_map[c] = 'Effective Date'
        elif any(x in cl for x in ['comment', 'remark', 'note', 'status', 'action']) and 'Comments' not in col_map.values():
            col_map[c] = 'Comments'

    df = df.rename(columns=col_map)

    # Ensure required columns exist
    if 'Dial Code' not in df.columns or 'Summary' not in df.columns or 'New Rate' not in df.columns:
        return False, "Uploaded file must contain Dial Code, Destination/Summary, and Rate columns.", None

    if 'Minimum' not in df.columns:
        df['Minimum'] = '1'
    if 'Increment' not in df.columns:
        df['Increment'] = '1'
    if 'Effective Date' not in df.columns:
        df['Effective Date'] = str(effective_date_val)

    # Clean strings
    for col in ['Dial Code', 'Summary', 'New Rate', 'Minimum', 'Increment', 'Effective Date']:
        df[col] = df[col].astype(str).str.strip()
    df['Effective Date'] = df['Effective Date'].astype(str).str.replace(r'\s+\d{2}:\d{2}:\d{2}.*$', '', regex=True)

    # 2. Filter rows where comment contains 'delete'
    deleted_comments_count = 0
    if 'Comments' in df.columns:
        mask_del = df['Comments'].astype(str).str.contains(r'delete', case=False, na=False)
        deleted_comments_count = int(mask_del.sum())
        df = df[~mask_del]

    # 3. Filter specifically USA, Canada, Puerto Rico
    nanp_specific_pattern = r'\b(?:USA|UNITED STATES|CANADA|PUERTO RICO)\b'
    mask_nanp = df['Summary'].astype(str).str.contains(nanp_specific_pattern, case=False, na=False)
    nanp_excluded_count = int(mask_nanp.sum())
    df = df[~mask_nanp]

    # Format all rates to max 6 decimal places (no floating point glitches)
    df['New Rate'] = df['New Rate'].apply(format_telecom_rate)

    # 4. Extract only standard 6 columns in exact order
    target_cols = ['Dial Code', 'Summary', 'New Rate', 'Minimum', 'Increment', 'Effective Date']
    df = df[target_cols].copy()

    # 5. Apply destination rate updates & track inc/dec
    change_summary = []
    dest_labels_for_name = []
    total_rows_updated = 0
    total_new_rows_added = 0
    missing_destinations = []
    added_from_master_summary = []
    eff_date_str = str(effective_date_val)

    if dest_updates:
        for item in dest_updates:
            target_dest = str(item.get("destination", "")).strip()
            new_rate_raw = item.get("rate")
            scope = str(item.get("scope", "All")).strip()
            fallback_source = item.get("source")  # "CC", "CLI", or None

            if not target_dest or new_rate_raw is None:
                continue
            try:
                new_rate_float = float(str(new_rate_raw).replace("$", "").replace(",", ".").strip())
            except Exception:
                continue

            formatted_new_rate = format_telecom_rate(new_rate_float)
            dest_escaped = re.escape(target_dest)
            match_mask = df['Summary'].str.contains(dest_escaped, case=False, na=False)
            matched_count = int(match_mask.sum())
            if matched_count == 0:
                words = [w for w in re.split(r'[\s\-_]+', target_dest) if len(w) > 1]
                if len(words) > 1:
                    match_mask = df['Summary'].astype(str).apply(lambda s: all(w.lower() in s.lower() for w in words))
                    matched_count = int(match_mask.sum())
            if matched_count == 0:
                match_mask = df['Dial Code'].str.startswith(target_dest, na=False)
                matched_count = int(match_mask.sum())
            if matched_count == 0:
                spell_candidates = smart_spell_check_destination(target_dest, top_n=2)
                for cand in spell_candidates:
                    c_words = [w for w in re.split(r'[\s\-_]+', cand) if len(w) > 1]
                    c_mask = df['Summary'].astype(str).apply(lambda s: all(w.lower() in s.lower() for w in c_words))
                    if int(c_mask.sum()) > 0:
                        match_mask = c_mask
                        matched_count = int(c_mask.sum())
                        target_dest = cand
                        break

            # Apply scope filter if matched in sheet
            if matched_count > 0:
                mobile_pattern = r'\b(?:mobile|cell|mob|cellular|wireless)\b'
                if scope in ["Fixed", "Fixed Only", "☎️ Fixed Only"]:
                    match_mask = match_mask & ~df['Summary'].str.contains(mobile_pattern, case=False, na=False)
                    matched_count = int(match_mask.sum())
                elif scope in ["Mobile", "Mobile Only", "📱 Mobile Only"]:
                    match_mask = match_mask & df['Summary'].str.contains(mobile_pattern, case=False, na=False)
                    matched_count = int(match_mask.sum())

            if matched_count > 0:
                existing_rates_series = pd.to_numeric(df.loc[match_mask, 'New Rate'], errors='coerce')
                old_avg_rate = existing_rates_series.dropna().mean()
                if pd.isna(old_avg_rate) or old_avg_rate == 0:
                    old_avg_rate = 0.0

                trend = "inc" if new_rate_float >= old_avg_rate else "dec"
                rate_diff = new_rate_float - old_avg_rate

                df.loc[match_mask, 'New Rate'] = formatted_new_rate
                df.loc[match_mask, 'Effective Date'] = eff_date_str

                total_rows_updated += matched_count
                dest_labels_for_name.append(f"{target_dest} {trend}")
                change_summary.append({
                    "destination": target_dest,
                    "scope": scope,
                    "old_rate": f"{old_avg_rate:.4f}",
                    "new_rate": formatted_new_rate,
                    "trend": trend.upper(),
                    "diff": rate_diff,
                    "rows_updated": matched_count,
                    "status": "Updated Existing"
                })
            else:
                # Destination missing in customer deck!
                if fallback_source in ["CC", "CLI"]:
                    existing_codes = df['Dial Code'].tolist()
                    df_new_rows, count_added = import_destination_from_master(
                        target_dest,
                        scope=scope,
                        new_rate_str=formatted_new_rate,
                        source=fallback_source,
                        eff_date_str=eff_date_str,
                        existing_dial_codes=existing_codes
                    )
                    if count_added > 0:
                        df = pd.concat([df, df_new_rows], ignore_index=True)
                        total_new_rows_added += count_added
                        dest_labels_for_name.append(f"{target_dest} new")
                        added_from_master_summary.append({
                            "destination": target_dest,
                            "scope": scope,
                            "source": f"Master {fallback_source}",
                            "rows_added": count_added,
                            "new_rate": formatted_new_rate
                        })
                    else:
                        missing_destinations.append({
                            "destination": target_dest,
                            "scope": scope,
                            "rate": formatted_new_rate,
                            "reason": f"Not found in Customer sheet or Master {fallback_source}"
                        })
                else:
                    missing_destinations.append({
                        "destination": target_dest,
                        "scope": scope,
                        "rate": formatted_new_rate,
                        "reason": "Not found in Customer deck"
                    })

    # FINAL STRICT DEDUPLICATION: Guarantees 0 duplicate dial codes in entire deck!
    df = df.drop_duplicates(subset=['Dial Code'], keep='last').reset_index(drop=True)

    # 6. Check mismatches against Country Master Billing (FIX TUPLE UNPACK!)
    mismatches = []
    if hasattr(bm, "audit_dataframe_billing"):
        try:
            audit_res = bm.audit_dataframe_billing(df)
            if isinstance(audit_res, tuple) and len(audit_res) >= 2:
                mismatches = audit_res[1]
            elif isinstance(audit_res, list):
                mismatches = audit_res
        except Exception:
            mismatches = []

    # 7. Build smart filename preserving original customer sheet base
    base_stem = original_filename.rsplit(".", 1)[0]
    base_clean = re.sub(r'^(?:name\s*[-:]?\s*)', '', base_stem, flags=re.I).strip()
    
    u_tag = str(current_username or "kaif").strip().lower()
    if len(dest_labels_for_name) == 1:
        suffix = f"{dest_labels_for_name[0]} {u_tag}"
    elif len(dest_labels_for_name) == 2:
        suffix = f"{dest_labels_for_name[0]} {dest_labels_for_name[1]} {u_tag}"
    elif len(dest_labels_for_name) > 2:
        trends = [c["trend"].lower() for c in change_summary]
        if all(t == "inc" for t in trends):
            suffix = f"Various inc {u_tag}"
        elif all(t == "dec" for t in trends):
            suffix = f"Various dec {u_tag}"
        else:
            suffix = f"Various inc dec {u_tag}"
    else:
        suffix = f"Standardized {u_tag}"

    suggested_full_name = f"Name - {base_clean} {suffix}"
    final_xl_filename = f"{suggested_full_name}.xlsx"
    final_csv_filename = core_engine.build_csv_output_filename(suggested_full_name)

    # 8. Generate Styled Excel (Vendor Sample formatting with AutoFilter) and CSV bytes
    xl_bytes = generate_customer_excel_bytes(df, sheet_name="Rate_Amendment")
    csv_bytes = core_engine.generate_target_csv_bytes(df)

    elapsed_time = time.time() - t0

    stats = {
        "initial_rows": initial_rows,
        "clean_rows": len(df),
        "deleted_comments_count": deleted_comments_count,
        "nanp_excluded_count": nanp_excluded_count,
        "total_rows_updated": total_rows_updated,
        "total_new_rows_added": total_new_rows_added,
        "destinations_count": len(dest_updates) if dest_updates else 0,
        "duplicate_free": True
    }

    result = {
        "df_clean": df,
        "xl_bytes": xl_bytes,
        "csv_bytes": csv_bytes,
        "filename": final_xl_filename,
        "csv_filename": final_csv_filename,
        "base_clean": base_clean,
        "suffix": suffix,
        "mismatches": mismatches,
        "has_mismatch": len(mismatches) > 0,
        "stats": stats,
        "change_summary": change_summary,
        "missing_destinations": missing_destinations,
        "added_from_master_summary": added_from_master_summary,
        "effective_date": eff_date_str,
        "elapsed_time": elapsed_time
    }
    return True, "Success", result

def generate_new_customer_deck_from_master(service_type, trunk_code, customer_name, dest_updates, effective_date_val, current_username="kaif"):
    import io, re, datetime, time, os
    import pandas as pd
    import billing_manager as bm
    import core_engine

    t0 = time.time()
    
    if service_type == "CC":
        p_master = os.path.join(os.path.dirname(__file__), "templates", "CC_Rates_Sheet_2026_02_01.xlsx")
        if not os.path.exists(p_master):
            p_master = r"C:/Vendor Rates\CC_Rates_Sheet_2026_02_01.xlsx"
    else:
        p_master = os.path.join(os.path.dirname(__file__), "templates", "CLI_Rate_sheet_2026_02_01.xlsx")
        if not os.path.exists(p_master):
            p_master = r"C:/Vendor Rates\CLI_Rate_sheet_2026_02_01.xlsx"

    if not os.path.exists(p_master):
        return False, f"Master {service_type} rate sheet not found on server ({p_master}).", None

    try:
        df_master = pd.read_excel(p_master, dtype=str)
    except Exception as e:
        return False, f"Could not load Master {service_type} sheet: {e}", None

    if trunk_code not in df_master.columns:
        return False, f"Trunk column '{trunk_code}' not found in Master {service_type} sheet.", None

    initial_rows = len(df_master)
    eff_date_str = str(effective_date_val)

    # Build 6-column base dataframe
    df = pd.DataFrame()
    df['Dial Code'] = df_master['Dial Code'].astype(str).str.strip()
    df['Summary'] = df_master['Summary'].astype(str).str.strip()
    df['New Rate'] = df_master[trunk_code].apply(format_telecom_rate)
    df['Minimum'] = df_master['Minimum'].astype(str).str.strip() if 'Minimum' in df_master.columns else '1'
    df['Increment'] = df_master['Increment'].astype(str).str.strip() if 'Increment' in df_master.columns else '1'
    df['Effective Date'] = eff_date_str

    # Filter specifically USA, Canada, Puerto Rico
    nanp_specific_pattern = r'\b(?:USA|UNITED STATES|CANADA|PUERTO RICO)\b'
    mask_nanp = df['Summary'].astype(str).str.contains(nanp_specific_pattern, case=False, na=False)
    nanp_excluded_count = int(mask_nanp.sum())
    df = df[~mask_nanp].copy().reset_index(drop=True)

    # Destination overrides
    change_summary = []
    dest_labels_for_name = []
    total_rows_updated = 0

    if dest_updates:
        for item in dest_updates:
            target_dest = str(item.get("destination", "")).strip()
            new_rate_raw = item.get("rate")
            if not target_dest or new_rate_raw is None:
                continue
            try:
                new_rate_float = float(str(new_rate_raw).replace("$", "").replace(",", ".").strip())
            except Exception:
                continue

            scope = str(item.get("scope", "All")).strip()
            dest_escaped = re.escape(target_dest)
            match_mask = df['Summary'].str.contains(dest_escaped, case=False, na=False)
            matched_count = int(match_mask.sum())
            if matched_count == 0:
                words = [w for w in re.split(r'[\s\-_]+', target_dest) if len(w) > 1]
                if len(words) > 1:
                    match_mask = df['Summary'].astype(str).apply(lambda s: all(w.lower() in s.lower() for w in words))
                    matched_count = int(match_mask.sum())
            if matched_count == 0:
                match_mask = df['Dial Code'].str.startswith(target_dest, na=False)
                matched_count = int(match_mask.sum())
            if matched_count == 0:
                spell_candidates = smart_spell_check_destination(target_dest, top_n=2)
                for cand in spell_candidates:
                    c_words = [w for w in re.split(r'[\s\-_]+', cand) if len(w) > 1]
                    c_mask = df['Summary'].astype(str).apply(lambda s: all(w.lower() in s.lower() for w in c_words))
                    if int(c_mask.sum()) > 0:
                        match_mask = c_mask
                        matched_count = int(c_mask.sum())
                        target_dest = cand
                        break

            if matched_count > 0:
                mobile_pattern = r'\b(?:mobile|cell|mob|cellular|wireless)\b'
                if scope in ["Fixed", "Fixed Only", "☎️ Fixed Only"]:
                    match_mask = match_mask & ~df['Summary'].str.contains(mobile_pattern, case=False, na=False)
                    matched_count = int(match_mask.sum())
                elif scope in ["Mobile", "Mobile Only", "📱 Mobile Only"]:
                    match_mask = match_mask & df['Summary'].str.contains(mobile_pattern, case=False, na=False)
                    matched_count = int(match_mask.sum())

            if matched_count > 0:
                existing_rates_series = pd.to_numeric(df.loc[match_mask, 'New Rate'], errors='coerce')
                old_avg_rate = existing_rates_series.dropna().mean()
                if pd.isna(old_avg_rate) or old_avg_rate == 0:
                    old_avg_rate = 0.0

                trend = "inc" if new_rate_float >= old_avg_rate else "dec"
                rate_diff = new_rate_float - old_avg_rate

                formatted_new_rate = format_telecom_rate(new_rate_float)
                df.loc[match_mask, 'New Rate'] = formatted_new_rate
                df.loc[match_mask, 'Effective Date'] = eff_date_str

                total_rows_updated += matched_count
                dest_labels_for_name.append(f"{target_dest} {trend}")
                change_summary.append({
                    "destination": target_dest,
                    "scope": scope,
                    "new_rate": formatted_new_rate,
                    "rows_updated": matched_count,
                    "is_new_deck": True
                })

    # FINAL STRICT DEDUPLICATION: Guarantees 0 duplicate dial codes in entire deck!
    df = df.drop_duplicates(subset=['Dial Code'], keep='last').reset_index(drop=True)

    # For NEW Customer Deck from Master A-Z, no mismatches (pulses are already master standard!)
    mismatches = []

    # Build clean filename following Kaif's exact formula:
    # 1 dest: {Customer} {Dest1} new with A-Z {username}
    # 2 dest: {Customer} {Dest1} {Dest2} new with A-Z {username}
    # >2 or 0 dest: {Customer} A to Z new {username}
    c_clean = re.sub(r'[^a-zA-Z0-9_\-\s]+', '', str(customer_name).strip())
    if not c_clean:
        c_clean = "New Customer"

    u_tag = str(current_username or "kaif").strip().lower()

    clean_dest_list = []
    if dest_updates:
        for it in dest_updates:
            d_str = str(it.get("destination", "")).strip()
            if d_str and d_str not in clean_dest_list:
                clean_dest_list.append(d_str)

    if len(clean_dest_list) == 1:
        suffix = f"{clean_dest_list[0]} new with A-Z {u_tag}"
    elif len(clean_dest_list) == 2:
        suffix = f"{clean_dest_list[0]} {clean_dest_list[1]} new with A-Z {u_tag}"
    else:
        suffix = f"A to Z new {u_tag}"

    suggested_full_name = f"{c_clean} {suffix}"

    final_xl_filename = f"{suggested_full_name}.xlsx"
    final_csv_filename = core_engine.build_csv_output_filename(suggested_full_name)

    # Generate Styled Excel matching Vendor Sample sheet exactly
    xl_bytes = generate_customer_excel_bytes(df, sheet_name="Rate_Amendment")
    csv_bytes = core_engine.generate_target_csv_bytes(df)

    elapsed_time = time.time() - t0

    stats = {
        "initial_rows": initial_rows,
        "clean_rows": len(df),
        "deleted_comments_count": 0,
        "nanp_excluded_count": nanp_excluded_count,
        "total_rows_updated": total_rows_updated,
        "destinations_count": len(dest_updates) if dest_updates else 0
    }

    result = {
        "df_clean": df,
        "xl_bytes": xl_bytes,
        "csv_bytes": csv_bytes,
        "filename": final_xl_filename,
        "csv_filename": final_csv_filename,
        "base_clean": f"{c_clean} ({trunk_code} {service_type})",
        "suffix": suffix,
        "mismatches": mismatches,
        "has_mismatch": len(mismatches) > 0,
        "stats": stats,
        "change_summary": change_summary,
        "effective_date": eff_date_str,
        "elapsed_time": elapsed_time
    }
    return True, "Success", result

def render_customer_standardizer_view():
    import datetime, time, io
    import pandas as pd
    import billing_manager as bm
    import core_engine

    curr_u = current_user.get("username", "kaif") if isinstance(current_user, dict) else "kaif"

    # 1. DISPLAY PERSISTENT RESULTS AT TOP (Identical to Vendor Portal!)
    if st.session_state.get("cust_processed_result"):
        item = st.session_state.cust_processed_result
        elapsed_sec = item.get("elapsed_time", 0.0)
        stats = item.get("stats", {})
        change_summary = item.get("change_summary", [])
        mismatches = item.get("mismatches", [])
        has_mm = len(mismatches) > 0
        
        st.markdown("---")
        col_hdr, col_clr = st.columns([3.8, 1.2])
        with col_hdr:
            elapsed_badge = f'<span style="background: #DCFCE7; color: #166534; font-weight: 700; font-size: 0.78rem; padding: 2px 10px; border-radius: 12px; border: 1px solid #86EFAC;">⏱️ {elapsed_sec:.2f}s</span>' if elapsed_sec is not None else ""
            st.markdown(
                f'<div style="display: inline-flex; align-items: center; gap: 10px; background: #F0FDF4; border: 1px solid #BBF7D0; border-radius: 20px; padding: 4px 14px; margin: 2px 0;">'
                f'<span style="font-weight: 700; color: #166534; font-size: 0.88rem;">🎉 Customer Rate Sheet Standardized Successfully</span>'
                f'{elapsed_badge}'
                f'</div>',
                unsafe_allow_html=True
            )
        with col_clr:
            if st.button("🗑️ Clear / Reset", type="secondary", use_container_width=True, key="cust_clear_results_btn"):
                st.session_state.cust_processed_result = None
                st.session_state.pop("cust_dest_items", None)
                st.session_state.pop("cust_new_dest_items", None)
                st.rerun()

        # Executive Card Container (Identical in style to Vendor Deck Card!)
        with st.container(border=True):
            if has_mm:
                st.markdown("""
                <style>
                div.st-key-cust_popover_mm button {
                    background-color: #FFFBEB !important;
                    border: 1.5px solid #FCD34D !important;
                    color: #B45309 !important;
                    font-weight: 700 !important;
                    font-size: 0.78rem !important;
                }
                div.st-key-cust_popover_mm button:hover {
                    background-color: #FEF3C7 !important;
                    border-color: #F59E0B !important;
                    color: #78350F !important;
                }
                </style>
                """, unsafe_allow_html=True)
                
                c_name, c_mm, c_fix, c_stats, c_date = st.columns([2.3, 1.2, 1.2, 0.8, 1.1])
                with c_name:
                    st.markdown(
                        f'<div style="display: flex; align-items: center; gap: 6px; min-height: 38px; font-weight: 700; color: #0C4A6E; font-size: 0.88rem; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">'
                        f'📁 {item["filename"]}'
                        f'</div>',
                        unsafe_allow_html=True
                    )
                with c_mm:
                    with st.popover(f"⚠️ {len(mismatches)} Mismatch ▾", use_container_width=True, key="cust_popover_mm"):
                        st.markdown(f"**⚠️ {len(mismatches)} Billing Mismatches Found:**")
                        st.caption("Customer pulse differs from standard master (Country Billing.xlsx)")
                        mm_table = []
                        for mm in mismatches[:100]:
                            if isinstance(mm, dict):
                                mm_table.append({
                                    "Destination": mm.get("destination", ""),
                                    "Dial Code": mm.get("dial_code", ""),
                                    "Sheet Pulse": mm.get("vendor_pulse", ""),
                                    "Standard Master": mm.get("standard_pulse", "")
                                })
                        st.dataframe(pd.DataFrame(mm_table), use_container_width=True, hide_index=True)
                with c_fix:
                    if st.button("⚡ Fix Standard", key="cust_btn_fix_std", type="primary", use_container_width=True):
                        df_curr = item.get("df_clean")
                        if df_curr is not None:
                            eff_d = item.get("effective_date", str(datetime.date.today()))
                            df_fixed = bm.apply_standard_billing_to_dataframe(df_curr, mismatches, new_effective_date=eff_d) if hasattr(bm, "apply_standard_billing_to_dataframe") else df_curr
                            item["xl_bytes"] = generate_customer_excel_bytes(df_fixed, sheet_name="Rate_Amendment")
                            item["df_clean"] = df_fixed
                            item["csv_bytes"] = core_engine.generate_target_csv_bytes(df_fixed)
                            item["mismatches"] = []
                            item["has_mismatch"] = False
                            st.success("✅ Standard billing applied successfully!")
                            time.sleep(0.8)
                            st.rerun()
                with c_stats:
                    with st.popover("📊 Stats ▾", use_container_width=True, key="cust_popover_stats"):
                        st.markdown(
                            f'<div style="min-width: 250px; font-family: -apple-system, BlinkMacSystemFont, sans-serif;">'
                            f'<div style="font-weight: 700; color: #0C4A6E; font-size: 0.88rem; margin-bottom: 8px; border-bottom: 1.5px solid #E2E8F0; padding-bottom: 6px;">'
                            f'📊 Row Breakdown & Stats'
                            f'</div>'
                            f'<div style="display: flex; flex-direction: column; gap: 6px; font-size: 0.82rem;">'
                            f'<div style="display: flex; justify-content: space-between; align-items: center; padding: 6px 10px; background: #F0F9FF; border: 1px solid #BAE6FD; border-radius: 6px;">'
                            f'<span style="color: #0369A1; font-weight: 700;">📥 INPUT ROWS</span>'
                            f'<span style="color: #0C4A6E; font-weight: 800; font-size: 0.95rem;">{stats.get("initial_rows", 0):,}</span>'
                            f'</div>'
                            f'<div style="display: flex; justify-content: space-between; align-items: center; padding: 6px 10px; background: #F0FDF4; border: 1px solid #BBF7D0; border-radius: 6px;">'
                            f'<span style="color: #15803D; font-weight: 700;">✅ CLEAN OUTPUT</span>'
                            f'<span style="color: #166534; font-weight: 800; font-size: 0.95rem;">{stats.get("clean_rows", 0):,}</span>'
                            f'</div>'
                            f'<div style="display: flex; justify-content: space-between; align-items: center; padding: 6px 10px; background: #FFFBEB; border: 1px solid #FDE68A; border-radius: 6px;">'
                            f'<span style="color: #92400E; font-weight: 700;">🚫 DELETED COMMENTS</span>'
                            f'<span style="color: #B45309; font-weight: 800; font-size: 0.95rem;">{stats.get("deleted_comments_count", 0):,}</span>'
                            f'</div>'
                            f'<div style="display: flex; justify-content: space-between; align-items: center; padding: 6px 10px; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 6px;">'
                            f'<span style="color: #475569; font-weight: 700;">🚫 USA/CANADA/PR</span>'
                            f'<span style="color: #0F172A; font-weight: 800; font-size: 0.95rem;">{stats.get("nanp_excluded_count", 0):,}</span>'
                            f'</div>'
                            f'<div style="display: flex; justify-content: space-between; align-items: center; padding: 6px 10px; background: #F0FDF4; border: 1px solid #BBF7D0; border-radius: 6px;">'
                            f'<span style="color: #15803D; font-weight: 700;">⚡ ROWS UPDATED</span>'
                            f'<span style="color: #166534; font-weight: 800; font-size: 0.95rem;">{stats.get("total_rows_updated", 0):,}</span>'
                            f'</div>'
                            f'<div style="display: flex; justify-content: space-between; align-items: center; padding: 6px 10px; background: #FAF5FF; border: 1px solid #E9D5FF; border-radius: 6px;">'
                            f'<span style="color: #7E22CE; font-weight: 700;">➕ NEW FROM MASTER</span>'
                            f'<span style="color: #6B21A8; font-weight: 800; font-size: 0.95rem;">{stats.get("total_new_rows_added", 0):,}</span>'
                            f'</div>'
                            f'<div style="display: flex; justify-content: space-between; align-items: center; padding: 6px 10px; background: #F0FDF4; border: 1px solid #BBF7D0; border-radius: 6px;">'
                            f'<span style="color: #15803D; font-weight: 700;">🛡️ ZERO DUPLICATES</span>'
                            f'<span style="color: #166534; font-weight: 800; font-size: 0.95rem;">100% Unique</span>'
                            f'</div>'
                            f'</div>'
                            f'</div>',
                            unsafe_allow_html=True
                        )
                with c_date:
                    st.markdown(
                        f'<div style="display: flex; align-items: center; justify-content: flex-end; min-height: 38px;">'
                        f'<span style="background: #F0FDF4; color: #16A34A; border: 1.5px solid #4ADE80; font-weight: 700; padding: 4px 12px; border-radius: 20px; font-size: 0.82rem; white-space: nowrap;">🟢 {item.get("effective_date")}</span>'
                        f'</div>',
                        unsafe_allow_html=True
                    )
            else:
                c_name, c_stats, c_date = st.columns([3.5, 0.9, 1.2])
                with c_name:
                    st.markdown(
                        f'<div style="display: flex; align-items: center; gap: 6px; min-height: 38px; font-weight: 700; color: #0C4A6E; font-size: 0.88rem; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">'
                        f'📁 {item["filename"]}'
                        f'</div>',
                        unsafe_allow_html=True
                    )
                with c_stats:
                    with st.popover("📊 Stats ▾", use_container_width=True, key="cust_popover_stats_clean"):
                        st.markdown(
                            f'<div style="min-width: 250px; font-family: -apple-system, BlinkMacSystemFont, sans-serif;">'
                            f'<div style="font-weight: 700; color: #0C4A6E; font-size: 0.88rem; margin-bottom: 8px; border-bottom: 1.5px solid #E2E8F0; padding-bottom: 6px;">'
                            f'📊 Row Breakdown & Stats'
                            f'</div>'
                            f'<div style="display: flex; flex-direction: column; gap: 6px; font-size: 0.82rem;">'
                            f'<div style="display: flex; justify-content: space-between; align-items: center; padding: 6px 10px; background: #F0FDF4; border: 1px solid #BBF7D0; border-radius: 6px;">'
                            f'<span style="color: #15803D; font-weight: 700;">✅ CLEAN OUTPUT</span>'
                            f'<span style="color: #166534; font-weight: 800; font-size: 0.95rem;">{stats.get("clean_rows", 0):,}</span>'
                            f'</div>'
                            f'<div style="display: flex; justify-content: space-between; align-items: center; padding: 6px 10px; background: #F0FDF4; border: 1px solid #BBF7D0; border-radius: 6px;">'
                            f'<span style="color: #15803D; font-weight: 700;">⚡ ROWS UPDATED</span>'
                            f'<span style="color: #166534; font-weight: 800; font-size: 0.95rem;">{stats.get("total_rows_updated", 0):,}</span>'
                            f'</div>'
                            f'</div>'
                            f'</div>',
                            unsafe_allow_html=True
                        )
                with c_date:
                    st.markdown(
                        f'<div style="display: flex; align-items: center; justify-content: flex-end; min-height: 38px;">'
                        f'<span style="background: #F0FDF4; color: #16A34A; border: 1.5px solid #4ADE80; font-weight: 700; padding: 4px 12px; border-radius: 20px; font-size: 0.82rem; white-space: nowrap;">🟢 {item.get("effective_date")}</span>'
                        f'</div>',
                        unsafe_allow_html=True
                    )

            # Executive Summary of Rate Changes Table
            if change_summary:
                st.markdown("<div style='height: 6px;'></div>", unsafe_allow_html=True)
                is_any_new = any(cs.get("is_new_deck") for cs in change_summary)
                if is_any_new:
                    st.markdown("**🎯 Configured Target Rates:**")
                    c_rows_display = []
                    for cs in change_summary:
                        sc_lbl = f" ({cs.get('scope')})" if cs.get('scope') and cs.get('scope') != 'All' else ""
                        c_rows_display.append({
                            "Destination": f"{cs['destination']}{sc_lbl}",
                            "Target Rate": cs["new_rate"],
                            "Codes Impacted": cs["rows_updated"],
                            "Status": "✅ Configured"
                        })
                else:
                    st.markdown("**📋 Rate Updates Breakdown:**")
                    c_rows_display = []
                    for cs in change_summary:
                        sc_lbl = f" ({cs.get('scope')})" if cs.get('scope') and cs.get('scope') != 'All' else ""
                        c_rows_display.append({
                            "Destination": f"{cs['destination']}{sc_lbl}",
                            "Old Rate": cs.get("old_rate", "-"),
                            "New Rate": cs["new_rate"],
                            "Action / Trend": f"📈 +{cs.get('diff', 0):.4f} (INC)" if cs.get("trend") == "INC" else f"📉 {cs.get('diff', 0):.4f} (DEC)",
                            "Rows Impacted": cs["rows_updated"]
                        })
                st.dataframe(pd.DataFrame(c_rows_display), use_container_width=True, hide_index=True)

            # Added from Master Summary Table
            if item.get("added_from_master_summary"):
                st.markdown("<div style='height: 6px;'></div>", unsafe_allow_html=True)
                st.markdown("**➕ Fresh Destinations Imported from Master Reference (Zero Duplicates):**")
                m_rows_disp = []
                for ms in item["added_from_master_summary"]:
                    m_rows_disp.append({
                        "Destination": ms["destination"],
                        "Scope": ms.get("scope", "All"),
                        "Source Master": ms.get("source", "Master"),
                        "New Rate": ms["new_rate"],
                        "Rows Added": ms["rows_added"],
                        "Status": "✅ Added (Unique)"
                    })
                st.dataframe(pd.DataFrame(m_rows_disp), use_container_width=True, hide_index=True)

            # Interactive Missing Destinations Action Box
            if item.get("missing_destinations"):
                st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)
                st.markdown(
                    f'<div style="background: #FFFBEB; border: 1.5px solid #F59E0B; border-radius: 8px; padding: 12px 14px; margin-bottom: 8px;">'
                    f'<div style="font-weight: 700; color: #B45309; font-size: 0.95rem; margin-bottom: 4px;">'
                    f'⚠️ {len(item["missing_destinations"])} Destination(s) Customer Sheet me Maujood Nahi Hain!'
                    f'</div>'
                    f'<div style="color: #92400E; font-size: 0.84rem; margin-bottom: 10px;">'
                    f'Aap in missing destinations ko direct Master CC ya Master CLI sheet se 1-click me import karke deck me add kar sakte hain (Zero-Duplicate Verified):'
                    f'</div>'
                    f'</div>',
                    unsafe_allow_html=True
                )
                for m_idx, m_item in enumerate(list(item["missing_destinations"])):
                    c_m_info, c_m_cc, c_m_cli = st.columns([2.2, 1.2, 1.2])
                    with c_m_info:
                        sc_txt = f"[{m_item.get('scope', 'All')}]"
                        st.markdown(f"📍 **{m_item['destination']}** {sc_txt} &rarr; Target Rate: `{m_item['rate']}`")
                    with c_m_cc:
                        if st.button(f"➕ Add from Master CC", key=f"btn_add_missing_cc_{m_idx}", type="secondary", use_container_width=True):
                            df_curr = item.get("df_clean")
                            eff_d = item.get("effective_date", str(datetime.date.today()))
                            existing_codes = df_curr['Dial Code'].tolist() if df_curr is not None else None
                            df_new_rows, count_added = import_destination_from_master(
                                m_item['destination'],
                                scope=m_item.get('scope', 'All'),
                                new_rate_str=m_item['rate'],
                                source='CC',
                                eff_date_str=eff_d,
                                existing_dial_codes=existing_codes
                            )
                            if count_added > 0 and df_curr is not None:
                                df_merged = pd.concat([df_curr, df_new_rows], ignore_index=True)
                                df_merged = df_merged.drop_duplicates(subset=['Dial Code'], keep='last').reset_index(drop=True)
                                item["xl_bytes"] = generate_customer_excel_bytes(df_merged, sheet_name="Rate_Amendment")
                                item["df_clean"] = df_merged
                                item["csv_bytes"] = core_engine.generate_target_csv_bytes(df_merged)
                                item["missing_destinations"].pop(m_idx)
                                item.setdefault("added_from_master_summary", []).append({
                                    "destination": m_item['destination'],
                                    "scope": m_item.get('scope', 'All'),
                                    "source": "Master CC",
                                    "rows_added": count_added,
                                    "new_rate": m_item['rate']
                                })
                                item["stats"]["clean_rows"] = len(df_merged)
                                item["stats"]["total_new_rows_added"] = item["stats"].get("total_new_rows_added", 0) + count_added
                                st.success(f"✅ {count_added} rows added for {m_item['destination']} from Master CC with 0 duplicates!")
                                time.sleep(0.6)
                                st.rerun()
                            else:
                                st.error(f"'{m_item['destination']}' Master CC me bhi nahi mila.")
                    with c_m_cli:
                        if st.button(f"➕ Add from Master CLI", key=f"btn_add_missing_cli_{m_idx}", type="secondary", use_container_width=True):
                            df_curr = item.get("df_clean")
                            eff_d = item.get("effective_date", str(datetime.date.today()))
                            existing_codes = df_curr['Dial Code'].tolist() if df_curr is not None else None
                            df_new_rows, count_added = import_destination_from_master(
                                m_item['destination'],
                                scope=m_item.get('scope', 'All'),
                                new_rate_str=m_item['rate'],
                                source='CLI',
                                eff_date_str=eff_d,
                                existing_dial_codes=existing_codes
                            )
                            if count_added > 0 and df_curr is not None:
                                df_merged = pd.concat([df_curr, df_new_rows], ignore_index=True)
                                df_merged = df_merged.drop_duplicates(subset=['Dial Code'], keep='last').reset_index(drop=True)
                                item["xl_bytes"] = generate_customer_excel_bytes(df_merged, sheet_name="Rate_Amendment")
                                item["df_clean"] = df_merged
                                item["csv_bytes"] = core_engine.generate_target_csv_bytes(df_merged)
                                item["missing_destinations"].pop(m_idx)
                                item.setdefault("added_from_master_summary", []).append({
                                    "destination": m_item['destination'],
                                    "scope": m_item.get('scope', 'All'),
                                    "source": "Master CLI",
                                    "rows_added": count_added,
                                    "new_rate": m_item['rate']
                                })
                                item["stats"]["clean_rows"] = len(df_merged)
                                item["stats"]["total_new_rows_added"] = item["stats"].get("total_new_rows_added", 0) + count_added
                                st.success(f"✅ {count_added} rows added for {m_item['destination']} from Master CLI with 0 duplicates!")
                                time.sleep(0.6)
                                st.rerun()
                            else:
                                st.error(f"'{m_item['destination']}' Master CLI me bhi nahi mila.")

            # Editable Filename text box
            st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)
            ed_fname_col, ed_btn_col = st.columns([3.6, 1.2])
            with ed_fname_col:
                new_name_input = st.text_input("Review / Edit Output Filename:", value=item["filename"].replace(".xlsx", ""), key="cust_fname_edit_input")
            with ed_btn_col:
                st.write("")
                st.write("")
                if st.button("✏️ Rename Deck", type="secondary", use_container_width=True, key="cust_btn_apply_rename"):
                    clean_nn = new_name_input.strip()
                    if clean_nn:
                        item["filename"] = f"{clean_nn}.xlsx"
                        item["csv_filename"] = core_engine.build_csv_output_filename(clean_nn)
                        st.toast("✅ Filename updated!")
                        st.rerun()

            # Download Buttons Row (Excel + CSV)
            c_dl1, c_dl2 = st.columns([3.2, 1.8])
            with c_dl1:
                st.download_button(
                    label=f"📊 Download: {item['filename']}",
                    data=item["xl_bytes"],
                    file_name=item["filename"],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="cust_exec_dl_excel",
                    type="primary",
                    use_container_width=True
                )
            with c_dl2:
                st.download_button(
                    label="📄 Download CSV",
                    data=item["csv_bytes"],
                    file_name=item["csv_filename"],
                    mime="text/csv",
                    key="cust_exec_dl_csv",
                    type="secondary",
                    use_container_width=True
                )

    # 2. TWO-TAB WORKFLOW CONTAINER
    cust_tab1, cust_tab2 = st.tabs([
        "📂 Existing Customer Deck (RouteTrust Update)",
        "➕ New Customer Deck Generator (From Master A-Z)"
    ])

    with cust_tab1:
        with st.container(border=True):
            st.markdown("""
            <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 4px;">
                <span style="font-size: 1.05rem; font-weight: 700; color: #0F172A;">📁 Upload Customer Rate Sheet (RouteTrust / CSV / Excel)</span>
            </div>
            """, unsafe_allow_html=True)
            st.caption("Auto-maps RouteTrust export (`Rate Key`, `Rate Summary`, `Inter. Rate`, `Dur Min/Inc`). Standardizes to exact 6 columns:")
            
            uploaded_cust_file = st.file_uploader(
                "Upload Customer Rate Deck",
                type=["xlsx", "xls", "csv"],
                key="cust_main_uploader",
                label_visibility="collapsed"
            )
            
            if uploaded_cust_file:
                st.success(f"📂 **{uploaded_cust_file.name}** loaded successfully ({len(uploaded_cust_file.getvalue())/1024:.1f} KB).")
                
                st.markdown("---")
                st.markdown("##### 🎯 Update Destinations & Rates (Optional)")
                st.caption("Enter requested destinations and new rates, or leave blank to standardize entire deck:")
                
                if "cust_dest_items" not in st.session_state:
                    st.session_state.cust_dest_items = [{"dest": "", "scope": "🌐 All", "rate": ""}]
                    
                with st.expander("🔍 Search & Pick Master Destination (CC & CLI)", expanded=False):
                    master_d_list = ["-- Select or Type to Search Destination --"] + get_master_destinations_data().get("destinations", [])
                    picked_d = st.selectbox("Search / Browse Master Destination:", master_d_list, key="cust_pick_master_tab1")
                    if picked_d and picked_d != "-- Select or Type to Search Destination --":
                        if st.button(f"➕ Use '{picked_d}' in Destination Input", key="btn_apply_picked_t1"):
                            if st.session_state.cust_dest_items:
                                st.session_state.cust_dest_items[0]["dest"] = picked_d
                                st.session_state["cust_dest_in_0"] = picked_d
                                st.rerun()

                for idx, d_item in enumerate(st.session_state.cust_dest_items):
                    c_d1, c_d2, c_d3, c_d4 = st.columns([2.0, 1.2, 1.1, 0.5])
                    with c_d1:
                        d_item["dest"] = st.text_input(f"Destination #{idx+1}:", value=d_item.get("dest", ""), placeholder="e.g. France or 33", key=f"cust_dest_in_{idx}")
                        dest_val = d_item.get("dest", "").strip()
                        if dest_val and len(dest_val) >= 2:
                            suggs = smart_spell_check_destination(dest_val, top_n=3)
                            if suggs and dest_val.lower() != suggs[0].lower():
                                c_sg1, c_sg2 = st.columns([2.8, 1.2])
                                with c_sg1:
                                    st.markdown(f"<span style='font-size: 0.78rem; color: #0E7490; font-weight: 600;'>💡 Did you mean: <b>{suggs[0]}</b>?</span>", unsafe_allow_html=True)
                                with c_sg2:
                                    if st.button("Apply", key=f"btn_apply_sugg_{idx}", type="secondary", use_container_width=True):
                                        st.session_state.cust_dest_items[idx]["dest"] = suggs[0]
                                        st.session_state[f"cust_dest_in_{idx}"] = suggs[0]
                                        st.rerun()
                        scope_opts = ["🌐 All", "☎️ Fixed Only", "📱 Mobile Only"]
                        curr_sc = d_item.get("scope", "🌐 All")
                        sc_idx = scope_opts.index(curr_sc) if curr_sc in scope_opts else 0
                        d_item["scope"] = st.selectbox(f"Scope #{idx+1}:", scope_opts, index=sc_idx, key=f"cust_scope_in_{idx}")
                    with c_d3:
                        d_item["rate"] = st.text_input(f"New Rate #{idx+1}:", value=d_item.get("rate", ""), placeholder="e.g. 0.015", key=f"cust_rate_in_{idx}")
                    with c_d4:
                        st.write("")
                        st.write("")
                        if len(st.session_state.cust_dest_items) > 1:
                            if st.button("✖", key=f"cust_del_dest_{idx}", help="Remove this destination"):
                                st.session_state.cust_dest_items.pop(idx)
                                st.rerun()

                col_add_dest, col_fb, col_date = st.columns([1.5, 1.5, 1.5])
                with col_add_dest:
                    if st.button("➕ Add More Destination", type="secondary", key="cust_btn_add_dest_row"):
                        st.session_state.cust_dest_items.append({"dest": "", "scope": "🌐 All", "rate": ""})
                        st.rerun()
                with col_fb:
                    sel_missing_fallback = st.selectbox(
                        "If missing in sheet:",
                        ["❓ Ask me (Show Options)", "📊 Auto-Import from Master CC", "📊 Auto-Import from Master CLI"],
                        index=0,
                        key="cust_missing_fb_sel"
                    )
                with col_date:
                    sel_eff_date = st.date_input("Effective Date (Default Today):", value=datetime.date.today(), key="cust_date_selector")
                    
                st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)
                if st.button("⚡ Standardize & Update Customer Rate Deck", type="primary", use_container_width=True, key="cust_btn_process"):
                    fb_mode = None
                    if "Master CC" in sel_missing_fallback:
                        fb_mode = "CC"
                    elif "Master CLI" in sel_missing_fallback:
                        fb_mode = "CLI"

                    valid_updates = []
                    for di in st.session_state.cust_dest_items:
                        d_val = di.get("dest", "").strip()
                        r_val = di.get("rate", "").strip()
                        s_val = di.get("scope", "🌐 All")
                        if d_val and r_val:
                            valid_updates.append({
                                "destination": d_val,
                                "scope": s_val,
                                "rate": r_val,
                                "source": fb_mode
                            })
                            
                    with st.spinner("Processing customer rate deck, filtering NANP & rounding rates to 6 digits..."):
                        ok, msg, res = process_customer_rate_deck(
                            uploaded_cust_file.getvalue(),
                            uploaded_cust_file.name,
                            valid_updates,
                            sel_eff_date,
                            curr_u
                        )
                        if ok:
                            st.session_state.cust_processed_result = res
                            st.success("✅ Customer rate deck standardized successfully!")
                            time.sleep(0.5)
                            st.rerun()
                        else:
                            st.error(msg)

    with cust_tab2:
        with st.container(border=True):
            col_t2_h, col_t2_rst = st.columns([3.5, 1.2])
            with col_t2_h:
                st.markdown("""
                <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 4px;">
                    <span style="font-size: 1.05rem; font-weight: 700; color: #0F172A;">➕ Generate New Customer Deck (From Master A-Z)</span>
                </div>
                """, unsafe_allow_html=True)
                st.caption("Generate a fresh 6-column rate sheet from Master A-Z CC or CLI reference sheets for a new client:")
            with col_t2_rst:
                if st.button("🗑️ Clear / Reset Form", key="btn_reset_new_cust_deck", type="secondary", use_container_width=True):
                    st.session_state.cust_new_dest_items = [{"dest": "", "scope": "🌐 All", "rate": ""}]
                    st.session_state.cust_processed_result = None
                    st.session_state.new_cust_name_in = ""
                    st.rerun()

            c_n1, c_n2 = st.columns([1.5, 2.5])
            with c_n1:
                sel_service = st.selectbox(
                    "Service Type:",
                    ["CC", "CLI"],
                    format_func=lambda x: "📊 CC Master (Call Center)" if x == "CC" else "📊 CLI Master (Pure Routes)",
                    key="new_cust_service_sel"
                )
            with c_n2:
                sel_trunk_obj = st.selectbox(
                    "🎯 Trunk Reference Profile:",
                    MASTER_TRUNK_ENTITIES,
                    format_func=lambda x: x["name"],
                    key="new_cust_trunk_sel"
                )
                sel_trunk_code = sel_trunk_obj["code"]

            new_c_name = st.text_input("New Customer Name:", placeholder="e.g. Dexatel VoIP or Zadarma", key="new_cust_name_in")

            st.markdown("---")
            st.markdown("##### 🎯 Custom Destination Overrides (Optional)")
            st.caption("If sales negotiated specific target rates (e.g. Italy: 0.01), enter them below or leave empty for base rates:")

            if "cust_new_dest_items" not in st.session_state:
                st.session_state.cust_new_dest_items = [{"dest": "", "scope": "🌐 All", "rate": ""}]

            with st.expander("🔍 Search & Pick Master Destination (CC & CLI)", expanded=False):
                master_n_d_list = ["-- Select or Type to Search Destination --"] + get_master_destinations_data().get("destinations", [])
                picked_n_d = st.selectbox("Search / Browse Master Destination:", master_n_d_list, key="cust_pick_master_tab2")
                if picked_n_d and picked_n_d != "-- Select or Type to Search Destination --":
                    if st.button(f"➕ Use '{picked_n_d}' in Destination Input", key="btn_apply_picked_t2"):
                        if st.session_state.cust_new_dest_items:
                            st.session_state.cust_new_dest_items[0]["dest"] = picked_n_d
                            st.session_state["cust_new_dest_in_0"] = picked_n_d
                            st.rerun()

            for idx, d_item in enumerate(st.session_state.cust_new_dest_items):
                c_d1, c_d2, c_d3, c_d4 = st.columns([2.0, 1.2, 1.1, 0.5])
                with c_d1:
                    d_item["dest"] = st.text_input(f"Override Destination #{idx+1}:", value=d_item.get("dest", ""), placeholder="e.g. France or 33", key=f"cust_new_dest_in_{idx}")
                    dest_val = d_item.get("dest", "").strip()
                    if dest_val and len(dest_val) >= 2:
                        suggs = smart_spell_check_destination(dest_val, top_n=3)
                        if suggs and dest_val.lower() != suggs[0].lower():
                            c_sg1, c_sg2 = st.columns([2.8, 1.2])
                            with c_sg1:
                                st.markdown(f"<span style='font-size: 0.78rem; color: #0E7490; font-weight: 600;'>💡 Did you mean: <b>{suggs[0]}</b>?</span>", unsafe_allow_html=True)
                            with c_sg2:
                                if st.button("Apply", key=f"btn_apply_new_sugg_{idx}", type="secondary", use_container_width=True):
                                    st.session_state.cust_new_dest_items[idx]["dest"] = suggs[0]
                                    st.session_state[f"cust_new_dest_in_{idx}"] = suggs[0]
                                    st.rerun()
                    scope_opts = ["🌐 All", "☎️ Fixed Only", "📱 Mobile Only"]
                    curr_sc = d_item.get("scope", "🌐 All")
                    sc_idx = scope_opts.index(curr_sc) if curr_sc in scope_opts else 0
                    d_item["scope"] = st.selectbox(f"Scope #{idx+1}:", scope_opts, index=sc_idx, key=f"cust_new_scope_in_{idx}")
                with c_d3:
                    d_item["rate"] = st.text_input(f"Target Rate #{idx+1}:", value=d_item.get("rate", ""), placeholder="e.g. 0.015", key=f"cust_new_rate_in_{idx}")
                with c_d4:
                    st.write("")
                    st.write("")
                    if len(st.session_state.cust_new_dest_items) > 1:
                        if st.button("✖", key=f"cust_new_del_dest_{idx}", help="Remove this override"):
                            st.session_state.cust_new_dest_items.pop(idx)
                            st.rerun()

            col_add_n_dest, col_n_date = st.columns([2, 2])
            with col_add_n_dest:
                if st.button("➕ Add More Rate Override", type="secondary", key="cust_new_btn_add_dest_row"):
                    st.session_state.cust_new_dest_items.append({"dest": "", "scope": "🌐 All", "rate": ""})
                    st.rerun()
            with col_n_date:
                sel_n_eff_date = st.date_input("Effective Date (Default Today):", value=datetime.date.today(), key="cust_new_date_selector")

            st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)
            if st.button("⚡ Generate Customer Rate Deck (From Master)", type="primary", use_container_width=True, key="cust_btn_generate_new"):
                clean_cname = new_c_name.strip()
                if not clean_cname:
                    st.warning("⚠️ Please enter a Customer Name before generating.")
                else:
                    valid_n_updates = []
                    for di in st.session_state.cust_new_dest_items:
                        d_val = di.get("dest", "").strip()
                        r_val = di.get("rate", "").strip()
                        s_val = di.get("scope", "🌐 All")
                        if d_val and r_val:
                            valid_n_updates.append({"destination": d_val, "scope": s_val, "rate": r_val})

                    with st.spinner(f"Generating rate deck for {clean_cname} from Master {sel_service} ({sel_trunk_code})..."):
                        ok, msg, res = generate_new_customer_deck_from_master(
                            sel_service,
                            sel_trunk_code,
                            clean_cname,
                            valid_n_updates,
                            sel_n_eff_date,
                            curr_u
                        )
                        if ok:
                            st.session_state.cust_processed_result = res
                            st.success(f"✅ Rate deck for '{clean_cname}' generated successfully!")
                            time.sleep(0.5)
                            st.rerun()
                        else:
                            st.error(msg)


# Sidebar already mounted at top

# ----------------- ROUTER DISPATCH -----------------
if active_portal == "👑 Super Admin (Accounts)":
    render_super_admin_accounts_tabs()
    st.stop()

elif active_portal == "👥 Customer Portal":
    if is_admin:
        customer_nav_mode = st.radio(
            "Customer Navigation:",
            ["⚡ Customer Rate Standardizer", "👑 Customer Admin Controls"],
            horizontal=True,
            key="cust_main_nav_toggle",
            label_visibility="collapsed"
        )
    else:
        customer_nav_mode = "⚡ Customer Rate Standardizer"
    
    if is_admin and customer_nav_mode == "👑 Customer Admin Controls":
        st.markdown("## 👥 Customer Admin Controls: Master A-Z Sheet (CC & CLI Manager)")
        st.caption("Manage and dynamically update Master A-Z CC & CLI ratesheets. Replaced sheets sync live across the entire system.")
        render_customer_master_az_tab()
        st.stop()
    else:
        render_customer_standardizer_view()
        st.stop()

elif active_portal == "🏢 Vendor Portal":
    if is_admin:
        vendor_nav_mode = st.radio(
            "Vendor Navigation:",
            ["⚡ Rate Deck Standardizer", "👑 Vendor Admin Controls"],
            horizontal=True,
            key="vend_main_nav_toggle",
            label_visibility="collapsed"
        )
    else:
        vendor_nav_mode = "⚡ Rate Deck Standardizer"
    
    if is_admin and vendor_nav_mode == "👑 Vendor Admin Controls":
        st.markdown("## 🏢 Vendor Admin Controls: Country Master Billing")
        st.caption("Review and manage company billing pulse rules (Country Master Billing). Live edits sync immediately.")
        render_billing_rules_tab()
        st.stop()

st.write("")

qc_codes = load_qc_exception_codes()

# Initialize active session processed results (Fresh session on login, no cross-session restore)
if "processed_results" not in st.session_state:
    st.session_state.processed_results = None
    st.session_state.elapsed_time = None

# If sheets were just processed, trigger automatic browser download via JS
if st.session_state.get("just_processed") and st.session_state.processed_results:
    st.session_state.pop("just_processed", None)
    first_item = st.session_state.processed_results[0]
    f_b = first_item.get("bytes") or read_disk_bytes(first_item.get("cached_path"))
    # Only trigger JS download for small payloads (<250KB) to avoid WebSocket buffer overflow and browser freeze
    if f_b and len(f_b) < 250_000:
        try:
            b64_str = base64.b64encode(f_b).decode('utf-8')
            fname_dl = first_item["filename"]
            components.html(f"""
            <script>
                try {{
                    var a = document.createElement('a');
                    a.href = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_str}';
                    a.download = '{fname_dl}';
                    document.body.appendChild(a);
                    a.click();
                }} catch(e) {{}}
            </script>
            """, height=0, width=0)
        except Exception:
            pass


def render_executive_deck_card(idx, item, is_single=False):
    out_name = re.sub(r'\bupdated\b', 'Updated', str(item["filename"]), flags=re.IGNORECASE)
    cached_p = item.get("cached_path") or os.path.join(OUTPUT_DIR, out_name)
    out_bytes = read_disk_bytes(cached_p)
    stats = item.get("stats", {})

    clean_rows_display = stats.get('final_clean_rows', item.get('rows_count', 0))
    initial_rows = stats.get('initial_rows', clean_rows_display)
    qc_ex = stats.get('qc_excluded', 0)
    nanp_ex = stats.get('nanp_excluded', 0)
    minus1_ex = stats.get('rates_minus1_count', 0)
    dedup_ex = stats.get('dedup_excluded', 0)

    # Date Badge
    date_badge_type = item.get("date_badge_type", "green")
    date_badge_label = item.get("date_badge_label", "Live")
    if date_badge_type == "red":
        date_pill = f'<span style="background: #FEF2F2; color: #DC2626; border: 1.5px solid #F87171; font-weight: 700; padding: 4px 12px; border-radius: 20px; font-size: 0.82rem; white-space: nowrap; display: inline-block;">🔴 {date_badge_label}</span>'
    elif date_badge_type == "orange":
        date_pill = f'<span style="background: #FFF7ED; color: #EA580C; border: 1.5px solid #FB923C; font-weight: 700; padding: 4px 12px; border-radius: 20px; font-size: 0.82rem; white-space: nowrap; display: inline-block;">🟠 {date_badge_label}</span>'
    else:
        date_pill = f'<span style="background: #F0FDF4; color: #16A34A; border: 1.5px solid #4ADE80; font-weight: 700; padding: 4px 12px; border-radius: 20px; font-size: 0.82rem; white-space: nowrap; display: inline-block;">🟢 {date_badge_label}</span>'

    mismatches = item.get("mismatches", [])
    has_mm = len(mismatches) > 0

    dedup_chip = (
        f'<div style="display: flex; justify-content: space-between; align-items: center; padding: 6px 10px; background: #FEF2F2; border: 1px solid #FECACA; border-radius: 6px;">'
        f'<span style="color: #991B1B; font-weight: 700;">✂️ DUPLICATES REMOVED</span>'
        f'<span style="color: #DC2626; font-weight: 800; font-size: 0.95rem;">{dedup_ex:,}</span>'
        f'</div>'
    ) if dedup_ex > 0 else ""

    prefix_tag = f'<span style="background: #E2E8F0; color: #475569; font-size: 0.72rem; font-weight: 800; padding: 2px 6px; border-radius: 4px; margin-right: 4px;">#{idx+1}</span>' if not is_single else ""
    orig_name = item.get("original_filename") or item.get("vendor_name") or out_name

    with st.container(border=True):
        if has_mm:
            st.markdown(f"""
            <style>
            div.st-key-popover_mm_{idx}_{is_single} button {{
                background-color: #FFFBEB !important;
                border: 1.5px solid #FCD34D !important;
                color: #B45309 !important;
                font-weight: 700 !important;
                font-size: 0.78rem !important;
            }}
            div.st-key-popover_mm_{idx}_{is_single} button:hover {{
                background-color: #FEF3C7 !important;
                border-color: #F59E0B !important;
                color: #78350F !important;
            }}
            </style>
            """, unsafe_allow_html=True)

            c_name, c_mm, c_fix, c_stats, c_date = st.columns([2.3, 1.2, 1.2, 0.8, 1.1])
            with c_name:
                st.markdown(
                    f'<div style="display: flex; align-items: center; gap: 6px; min-height: 38px; font-weight: 700; color: #0C4A6E; font-size: 0.88rem; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">'
                    f'{prefix_tag}📁 {orig_name}'
                    f'</div>',
                    unsafe_allow_html=True
                )
            with c_mm:
                with st.popover(f"⚠️ {len(mismatches)} Mismatch ▾", use_container_width=True, key=f"popover_mm_{idx}_{is_single}"):
                    st.markdown(f"**⚠️ {len(mismatches)} Billing Mismatches Found:**")
                    st.caption("Vendor pulse differs from standard master (Country Billing.xlsx)")
                    mm_table = []
                    for mm in mismatches[:100]:
                        if isinstance(mm, dict):
                            mm_table.append({
                                "Destination": mm.get("destination", ""),
                                "Dial Code": mm.get("dial_code", ""),
                                "Vendor": mm.get("vendor_pulse", ""),
                                "Standard": mm.get("standard_pulse", "")
                            })
                    st.dataframe(pd.DataFrame(mm_table), use_container_width=True, hide_index=True)
            with c_fix:
                btn_fix_label = "⚡ Fix Standard"
                if st.button(btn_fix_label, key=f"btn_apply_std_billing_{idx}_{is_single}", type="primary", use_container_width=True):
                    df_curr = item.get("df_clean")
                    if df_curr is not None:
                        eff_d = item.get("effective_date", str(datetime.date.today()))
                        df_fixed = bm.apply_standard_billing_to_dataframe(df_curr, mismatches, new_effective_date=eff_d) if hasattr(bm, "apply_standard_billing_to_dataframe") else df_curr
                        new_xl_bytes = core_engine.generate_target_excel_bytes(df_fixed)
                        new_csv_bytes = core_engine.generate_target_csv_bytes(df_fixed)
                        
                        cp = item.get("cached_path")
                        if cp:
                            with open(cp, "wb") as f_x:
                                f_x.write(new_xl_bytes)
                        cp_csv = item.get("cached_csv_path")
                        if cp_csv:
                            with open(cp_csv, "wb") as f_c:
                                f_c.write(new_csv_bytes)
                        
                        item["df_clean"] = df_fixed
                        item["mismatches"] = []
                        item["has_mismatch"] = False

                        if len(st.session_state.processed_results) > 1 and st.session_state.get("batch_zip_xl_path"):
                            try:
                                with zipfile.ZipFile(st.session_state.batch_zip_xl_path, "w", zipfile.ZIP_DEFLATED) as zf_xl:
                                    for p_item in st.session_state.processed_results:
                                        cp_i = p_item.get("cached_path")
                                        if cp_i and os.path.exists(cp_i):
                                            zf_xl.write(cp_i, arcname=p_item["filename"])
                            except Exception:
                                pass

                        st.success("✅ Standard billing applied successfully!")
                        time.sleep(0.8)
                        st.rerun()
            with c_stats:
                with st.popover("📊 Stats ▾", use_container_width=True, key=f"popover_stats_{idx}_{is_single}"):
                    st.markdown(
                        f'<div style="min-width: 250px; font-family: -apple-system, BlinkMacSystemFont, sans-serif;">'
                        f'<div style="font-weight: 700; color: #0C4A6E; font-size: 0.88rem; margin-bottom: 8px; border-bottom: 1.5px solid #E2E8F0; padding-bottom: 6px;">'
                        f'📊 Row Breakdown & QC Stats'
                        f'</div>'
                        f'<div style="display: flex; flex-direction: column; gap: 6px; font-size: 0.82rem;">'
                        f'<div style="display: flex; justify-content: space-between; align-items: center; padding: 6px 10px; background: #F0F9FF; border: 1px solid #BAE6FD; border-radius: 6px;">'
                        f'<span style="color: #0369A1; font-weight: 700;">📥 VENDOR INPUT</span>'
                        f'<span style="color: #0C4A6E; font-weight: 800; font-size: 0.95rem;">{initial_rows:,}</span>'
                        f'</div>'
                        f'<div style="display: flex; justify-content: space-between; align-items: center; padding: 6px 10px; background: #F0FDF4; border: 1px solid #BBF7D0; border-radius: 6px;">'
                        f'<span style="color: #15803D; font-weight: 700;">✅ CLEAN OUTPUT</span>'
                        f'<span style="color: #166534; font-weight: 800; font-size: 0.95rem;">{clean_rows_display:,}</span>'
                        f'</div>'
                        f'<div style="display: flex; justify-content: space-between; align-items: center; padding: 6px 10px; background: #FFFBEB; border: 1px solid #FDE68A; border-radius: 6px;">'
                        f'<span style="color: #92400E; font-weight: 700;">🚫 BLOCKED CODES</span>'
                        f'<span style="color: #B45309; font-weight: 800; font-size: 0.95rem;">{qc_ex:,}</span>'
                        f'</div>'
                        f'<div style="display: flex; justify-content: space-between; align-items: center; padding: 6px 10px; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 6px;">'
                        f'<span style="color: #475569; font-weight: 700;">🚫 NANP (+1)</span>'
                        f'<span style="color: #0F172A; font-weight: 800; font-size: 0.95rem;">{nanp_ex:,}</span>'
                        f'</div>'
                        f'<div style="display: flex; justify-content: space-between; align-items: center; padding: 6px 10px; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 6px;">'
                        f'<span style="color: #475569; font-weight: 700;">🚫 RATES TO -1</span>'
                        f'<span style="color: #0F172A; font-weight: 800; font-size: 0.95rem;">{minus1_ex:,}</span>'
                        f'</div>'
                        f'{dedup_chip}'
                        f'</div>'
                        f'</div>',
                        unsafe_allow_html=True
                    )
            with c_date:
                st.markdown(
                    f'<div style="display: flex; align-items: center; justify-content: flex-end; min-height: 38px;">'
                    f'{date_pill}'
                    f'</div>',
                    unsafe_allow_html=True
                )
        else:
            c_name, c_stats, c_date = st.columns([3.5, 0.9, 1.2])
            with c_name:
                st.markdown(
                    f'<div style="display: flex; align-items: center; gap: 6px; min-height: 38px; font-weight: 700; color: #0C4A6E; font-size: 0.88rem; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">'
                    f'{prefix_tag}📁 {orig_name}'
                    f'</div>',
                    unsafe_allow_html=True
                )
            with c_stats:
                with st.popover("📊 Stats ▾", use_container_width=True, key=f"popover_stats_{idx}_{is_single}"):
                    st.markdown(
                        f'<div style="min-width: 250px; font-family: -apple-system, BlinkMacSystemFont, sans-serif;">'
                        f'<div style="font-weight: 700; color: #0C4A6E; font-size: 0.88rem; margin-bottom: 8px; border-bottom: 1.5px solid #E2E8F0; padding-bottom: 6px;">'
                        f'📊 Row Breakdown & QC Stats'
                        f'</div>'
                        f'<div style="display: flex; flex-direction: column; gap: 6px; font-size: 0.82rem;">'
                        f'<div style="display: flex; justify-content: space-between; align-items: center; padding: 6px 10px; background: #F0F9FF; border: 1px solid #BAE6FD; border-radius: 6px;">'
                        f'<span style="color: #0369A1; font-weight: 700;">📥 VENDOR INPUT</span>'
                        f'<span style="color: #0C4A6E; font-weight: 800; font-size: 0.95rem;">{initial_rows:,}</span>'
                        f'</div>'
                        f'<div style="display: flex; justify-content: space-between; align-items: center; padding: 6px 10px; background: #F0FDF4; border: 1px solid #BBF7D0; border-radius: 6px;">'
                        f'<span style="color: #15803D; font-weight: 700;">✅ CLEAN OUTPUT</span>'
                        f'<span style="color: #166534; font-weight: 800; font-size: 0.95rem;">{clean_rows_display:,}</span>'
                        f'</div>'
                        f'<div style="display: flex; justify-content: space-between; align-items: center; padding: 6px 10px; background: #FFFBEB; border: 1px solid #FDE68A; border-radius: 6px;">'
                        f'<span style="color: #92400E; font-weight: 700;">🚫 BLOCKED CODES</span>'
                        f'<span style="color: #B45309; font-weight: 800; font-size: 0.95rem;">{qc_ex:,}</span>'
                        f'</div>'
                        f'<div style="display: flex; justify-content: space-between; align-items: center; padding: 6px 10px; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 6px;">'
                        f'<span style="color: #475569; font-weight: 700;">🚫 NANP (+1)</span>'
                        f'<span style="color: #0F172A; font-weight: 800; font-size: 0.95rem;">{nanp_ex:,}</span>'
                        f'</div>'
                        f'<div style="display: flex; justify-content: space-between; align-items: center; padding: 6px 10px; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 6px;">'
                        f'<span style="color: #475569; font-weight: 700;">🚫 RATES TO -1</span>'
                        f'<span style="color: #0F172A; font-weight: 800; font-size: 0.95rem;">{minus1_ex:,}</span>'
                        f'</div>'
                        f'{dedup_chip}'
                        f'</div>'
                        f'</div>',
                        unsafe_allow_html=True
                    )
            with c_date:
                st.markdown(
                    f'<div style="display: flex; align-items: center; justify-content: flex-end; min-height: 38px;">'
                    f'{date_pill}'
                    f'</div>',
                    unsafe_allow_html=True
                )

        # Download Buttons: Excel (Standardized A to Z) + CSV + Partial
        c_dl1, c_dl2, c_dl3 = st.columns([2.6, 1.2, 1.2])
        with c_dl1:
            st.download_button(
                label=f"📊 Download: {out_name}",
                data=out_bytes,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"exec_dl_excel_{idx}_{is_single}",
                type="primary",
                use_container_width=True
            )
        with c_dl2:
            csv_name = item.get("csv_filename") or core_engine.build_csv_output_filename(item.get("vendor_name", out_name))
            cached_csv_p = item.get("cached_csv_path") or os.path.join(OUTPUT_DIR, csv_name)
            csv_data = read_disk_bytes(cached_csv_p)
            st.download_button(
                label="📄 Download CSV",
                data=csv_data,
                file_name=csv_name,
                mime="text/csv",
                key=f"exec_dl_csv_{idx}_{is_single}",
                type="secondary",
                use_container_width=True
            )
        with c_dl3:
            partial_name = core_engine.build_partial_filename(out_name)
            st.download_button(
                label="📑 Partial (.xlsx)",
                data=out_bytes,
                file_name=partial_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help=f"Download with Partial filename: {partial_name}",
                key=f"exec_dl_partial_{idx}_{is_single}",
                type="secondary",
                use_container_width=True
            )

# Display persistent results AT THE TOP (Never hidden below the fold!)
if st.session_state.processed_results:
    results = st.session_state.processed_results
    elapsed_sec = st.session_state.get("elapsed_time")
    
    st.markdown("---")
    col_hdr, col_clr = st.columns([3.8, 1.2])
    with col_hdr:
        elapsed_badge = f'<span style="background: #DCFCE7; color: #166534; font-weight: 700; font-size: 0.78rem; padding: 2px 10px; border-radius: 12px; border: 1px solid #86EFAC;">⏱️ {elapsed_sec:.2f}s</span>' if elapsed_sec is not None else ""
        st.markdown(
            f'<div style="display: inline-flex; align-items: center; gap: 10px; background: #F0FDF4; border: 1px solid #BBF7D0; border-radius: 20px; padding: 4px 14px; margin: 2px 0;">'
            f'<span style="font-weight: 700; color: #166534; font-size: 0.88rem;">🎉 {len(results)} Sheet(s) Standardized Successfully</span>'
            f'{elapsed_badge}'
            f'</div>',
            unsafe_allow_html=True
        )
    with col_clr:
        if st.button("🗑️ Clear / Reset", type="secondary", use_container_width=True, key="top_clear_results_btn"):
            curr_u = current_user.get("username", "kaif") if ("current_user" in locals() and isinstance(current_user, dict)) else "kaif"
            clear_user_processed_cache(curr_u)
            for z_key in ["batch_zip_xl_path", "batch_zip_csv_path"]:
                z_p = st.session_state.get(z_key)
                if z_p and os.path.exists(z_p):
                    try:
                        os.remove(z_p)
                    except Exception:
                        pass
                st.session_state[z_key] = None
            st.session_state.processed_results = None
            st.session_state.elapsed_time = None
            st.session_state["cache_cleared"] = True
            st.session_state.uploader_key = st.session_state.get("uploader_key", 0) + 1
            gc.collect()
            st.rerun()

    if len(results) > 1:
        zip_xl_path = st.session_state.get("batch_zip_xl_path")
        zip_csv_path = st.session_state.get("batch_zip_csv_path")
        
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #F0F9FF 0%, #E0F2FE 100%); border: 1.5px solid #0E7490; border-radius: 10px; padding: 12px 18px; margin: 10px 0 14px 0;">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px;">
                <span style="font-weight: 800; color: #0C4A6E; font-size: 1.02rem;">
                    📦 Batch Package: All {len(results)} Sheets Standardized ({elapsed_sec:.2f}s)
                </span>
                <span style="background: #0E7490; color: #FFFFFF; font-size: 0.74rem; font-weight: 700; padding: 3px 10px; border-radius: 12px;">
                    BATCH READY
                </span>
            </div>
            <div style="font-size: 0.80rem; color: #475569;">
                Download all standardized decks in a single archive, or inspect individual files below:
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        col_b_zip1, col_b_zip2 = st.columns(2)
        with col_b_zip1:
            st.download_button(
                label=f"📦 Download All {len(results)} Excel Sheets (.zip)",
                data=read_disk_bytes(zip_xl_path),
                file_name=f"Standardized_Excel_Sheets_{u_name}.zip",
                mime="application/zip",
                key="top_batch_zip_excel_dl",
                type="primary",
                use_container_width=True
            )
        with col_b_zip2:
            st.download_button(
                label=f"📦 Download All {len(results)} CSV Files (.zip)",
                data=read_disk_bytes(zip_csv_path),
                file_name=f"Standardized_CSV_Decks_{u_name}.zip",
                mime="application/zip",
                key="top_batch_zip_csv_dl",
                type="primary",
                use_container_width=True
            )

        summary_rows = []
        for idx, item in enumerate(results):
            s = item.get("stats", {})
            summary_rows.append({
                "#": idx + 1,
                "Standardized File": item["filename"],
                "Clean Rows": f"{s.get('final_clean_rows', item.get('rows_count', 0)):,}",
                "Blocked Codes": f"{s.get('qc_excluded', 0):,}",
                "NANP (+1)": f"{s.get('nanp_excluded', 0):,}",
                "Status": "✅ Complete"
            })
        st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)

        st.markdown("#### 📁 Standardized Rate Decks & Downloads:")
        for idx, item in enumerate(results):
            render_executive_deck_card(idx, item, is_single=False)
    else:
        item = results[0]
        render_executive_deck_card(0, item, is_single=True)
    st.markdown("---")

if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

# File Uploader in dedicated card
with st.container(border=True):
    st.markdown("""
    <div style="margin-bottom: 6px;">
        <span style="font-size: 1.05rem; font-weight: 700; color: #0F172A;">📁 Upload Vendor Rate Sheet(s)</span>
        <div style="font-size: 0.82rem; color: #64748B;">Supports up to 10 files (.xlsx, .xls, .csv) or ZIP archive (max 500MB)</div>
    </div>
    """, unsafe_allow_html=True)

    uploaded_files = st.file_uploader(
        "Upload Vendor Rate Sheet(s) (Max 10 files or ZIP Archive)",
        type=["xlsx", "xls", "csv", "zip"],
        accept_multiple_files=True,
        key=f"vendor_file_uploader_{st.session_state.uploader_key}",
        label_visibility="collapsed",
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

    if len(raw_files) > 10:
        st.warning(f"⚠️ Maximum 10 sheets allowed per batch. Processing first 10 files out of {len(raw_files)}.")
        raw_files = raw_files[:10]

    if len(raw_files) > 1:
        st.markdown(f"""
        <div style="background: #FFFFFF; border: 1.5px solid #BAE6FD; border-radius: 10px; padding: 12px 16px; margin: 12px 0 14px 0; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 1px 3px rgba(14, 116, 144, 0.05);">
            <div>
                <span style="font-weight: 700; color: #0C4A6E; font-size: 0.96rem;">📁 Batch Configuration ({len(raw_files)} Sheets Loaded)</span>
                <div style="font-size: 0.76rem; color: #64748B;">Enter or confirm Vendor / Trunk Names below. Output filenames update in real-time:</div>
            </div>
            <span style="background: #E0F2FE; color: #0284C7; font-weight: 700; font-size: 0.78rem; padding: 4px 12px; border-radius: 14px; border: 1px solid #BAE6FD;">
                ⚡ {len(raw_files)} Files in Batch
            </span>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown("### 📝 Enter Vendor / Trunk Name")
        st.caption("Customize Vendor / Trunk name for each file. Live final filename preview is shown below:")

    vendor_names = {}
    for idx, (fname, f_content) in enumerate(raw_files):
        clean_stem = os.path.splitext(fname)[0]
        clean_stem = re.sub(r'_\d{8,}.*$', '', clean_stem).strip()
        default_vendor_name = f"Name - {clean_stem}"

        short_fname = fname if len(fname) <= 35 else (fname[:32] + "...")
        c1, c2 = st.columns([1.8, 2.2])
        v_name = c1.text_input(
            f"#{idx+1}: {short_fname}",
            value=default_vendor_name,
            key=f"v_name_input_{st.session_state.uploader_key}_{idx}",
            help=f"Original File: {fname}"
        )
        vendor_names[idx] = v_name
        
        clean_v_name = v_name.strip() if v_name else f"File {idx+1}"
        
        if 'a to z' in clean_v_name.lower():
            preview_filename = f"{clean_v_name} Updated {active_member}.xlsx"
        else:
            preview_filename = f"{clean_v_name} A to Z Updated {active_member}.xlsx"
            
        preview_csv_name = core_engine.build_csv_output_filename(clean_v_name)
        c2.markdown(f"""
        <div style="margin-top: 26px; padding: 6px 12px; background: #F0F9FF; border: 1px solid #BAE6FD; border-radius: 6px; font-family: monospace; font-size: 0.78rem; color: #0C4A6E; line-height: 1.35;">
            <div style="overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">📊 <b>XLSX:</b> <span style="color: #0E7490; font-weight: 600;">{preview_filename}</span></div>
            <div style="overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">📄 <b>CSV:</b> <span style="color: #059669; font-weight: 600;">{preview_csv_name}</span></div>
        </div>
        """, unsafe_allow_html=True)

    st.write("")
    
    # User's explicit requested button name: Proceed for Download
    btn_label = f"⚡ Proceed for Download ({len(raw_files)} Sheets)" if len(raw_files) > 1 else "⚡ Proceed for Download"
    if st.button(btn_label, type="primary", use_container_width=True, key=f"proceed_btn_{st.session_state.uploader_key}"):
        with st.spinner(f"⚡ Processing {len(raw_files)} sheet(s)... Standardizing rows, please wait..."):
            t_start = time.time()
            processed_list = []
            progress_bar = st.progress(0)

            user_dl_dir = os.path.join(os.path.expanduser('~'), 'Downloads')
            kaif_dir = r"C:\Kaif\Standardized_Sheets"
            os.makedirs(kaif_dir, exist_ok=True)

            for f_idx, (fname, f_content) in enumerate(raw_files):
                f_content.seek(0)
                df_raw = None
                xl = None
                try:
                    if fname.lower().endswith(('.xlsx', '.xlsm', '.xls', '.xlsb')):
                        xl = load_excel_safely(f_content, fname)
                        if len(xl.sheet_names) == 1:
                            target_sheet = xl.sheet_names[0]
                        else:
                            best_sheet = None
                            best_score = -9999
                            for s_name in xl.sheet_names:
                                try:
                                    temp_sample = xl.parse(s_name, header=None, nrows=40)
                                    if len(temp_sample) < 2:
                                        continue
                                    _, col_map = core_engine.detect_header_row_and_columns(temp_sample)
                                    score = len(col_map)
                                    if ('dial_code' in col_map or 'dial_code_split' in col_map) and 'rate' in col_map:
                                        score += 25
                                    if 'destination' in col_map:
                                        score += 5
                                    if 'effective_date' in col_map:
                                        score += 5
                                    s_clean = str(s_name).strip().lower()
                                    if any(bad in s_clean for bad in ['a num', 'a-num', 'a_num', 'a number', 'orig', 'cover', 'disclaimer', 'note', 'terms', 'instruction', 'faq', 'log']):
                                        score -= 40
                                    if any(good in s_clean for good in ['pricing', 'rate', 'rates', 'standard', 'prm', 'premium', 'wholesale', 'amendment', 'a to z', 'tariff', 'deck']):
                                        score += 15
                                    if score > best_score:
                                        best_score = score
                                        best_sheet = s_name
                                except Exception:
                                    pass
                            target_sheet = best_sheet if best_sheet is not None else 0
                        df_raw = xl.parse(target_sheet, header=None)
                    elif fname.lower().endswith('.csv'):
                        df_raw = load_csv_safely(f_content)
                except Exception as e:
                    st.error(f"Could not parse file '{fname}': {e}")
                    continue
                finally:
                    if xl is not None:
                        try:
                            xl.close()
                        except Exception:
                            pass
                        del xl

                if df_raw is None or len(df_raw) == 0:
                    st.warning(f"File '{fname}' is empty.")
                    continue

                df_clean, df_excluded, stats = core_engine.fast_process_dataframe(df_raw, qc_codes, fname)
                del df_raw, df_excluded
                gc.collect()

                user_base = vendor_names.get(f_idx, f"Name - {fname}")
                final_filename = core_engine.build_output_filename(user_base, stats['comments_tag'], active_member)
                final_csv_filename = core_engine.build_csv_output_filename(user_base)

                # Audit billing against master standard & Auto-learn new dial codes
                try:
                    if hasattr(bm, "audit_dataframe_billing"):
                        res_audit = bm.audit_dataframe_billing(df_clean)
                        if isinstance(res_audit, tuple) and len(res_audit) == 4:
                            has_mm, mm_list, mm_summary, learned_count = res_audit
                        elif isinstance(res_audit, tuple) and len(res_audit) == 3:
                            has_mm, mm_list, mm_summary = res_audit
                            learned_count = 0
                        else:
                            has_mm, mm_list, mm_summary, learned_count = False, [], {}, 0
                    else:
                        has_mm, mm_list, mm_summary, learned_count = False, [], {}, 0
                except Exception:
                    has_mm, mm_list, mm_summary, learned_count = False, [], {}, 0
                if learned_count > 0:
                    st.toast(f"🧠 Auto-learned & added {learned_count} new dial code(s) to Master Billing!", icon="✅")

                # Compute accurate date badge
                today_dt = (datetime.datetime.utcnow() + datetime.timedelta(hours=5, minutes=30)).date()
                
                def extract_fn_date(filename):
                    # Match DD-MM-YYYY, DD.MM.YYYY, DD_MM_YYYY
                    m_dmy = re.search(r'\b([0-3]?\d)[-_.\/]([0-1]?\d)[-_.\/](202\d)\b', filename)
                    if m_dmy:
                        d1, m1, y1 = int(m_dmy.group(1)), int(m_dmy.group(2)), int(m_dmy.group(3))
                        if 1 <= m1 <= 12 and 1 <= d1 <= 31:
                            return datetime.date(y1, m1, d1)
                    # Match YYYYMMDD
                    m_ymd = re.search(r'\b(202\d)(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])\b', filename)
                    if m_ymd:
                        return datetime.date(int(m_ymd.group(1)), int(m_ymd.group(2)), int(m_ymd.group(3)))
                    return None

                eff_date_obj = None
                if '_dt_obj' in df_clean.columns:
                    v_dates = [d.date() if isinstance(d, datetime.datetime) else d for d in df_clean['_dt_obj'] if d is not None and isinstance(d, (datetime.date, datetime.datetime))]
                    if v_dates:
                        future_v_dates = [d for d in v_dates if d > today_dt]
                        if future_v_dates:
                            # 1. Priority to FUTURE dates: Always take latest future effective date
                            eff_date_obj = max(future_v_dates)
                        elif any(d == today_dt for d in v_dates):
                            # 2. Priority to TODAY / Live dates
                            eff_date_obj = today_dt
                        else:
                            # 3. Check filename date: If vendor named file with today's or recent date
                            fn_d = extract_fn_date(fname)
                            if fn_d and fn_d >= max(v_dates):
                                eff_date_obj = fn_d
                            else:
                                # Pick latest date in the sheet (NEVER Counter most_common which picks ancient dates!)
                                eff_date_obj = max(v_dates)

                if not eff_date_obj:
                    fn_d = extract_fn_date(fname)
                    if fn_d:
                        eff_date_obj = fn_d

                date_badge_type = "green"
                date_badge_label = "Live"
                if eff_date_obj:
                    date_display_str = eff_date_obj.strftime("%d-%b-%Y")
                    if eff_date_obj > today_dt:
                        date_badge_type = "red"
                        date_badge_label = f"Future: {date_display_str}"
                    elif eff_date_obj < today_dt:
                        date_badge_type = "orange"
                        date_badge_label = f"Past: {date_display_str}"
                    else:
                        date_badge_type = "green"
                        date_badge_label = f"Live: {date_display_str}"
                else:
                    raw_dates = df_clean['Eff. Date'].dropna().tolist() if 'Eff. Date' in df_clean.columns else []
                    if raw_dates:
                        date_badge_label = f"Date: {raw_dates[0]}"

                excel_bytes = core_engine.generate_target_excel_bytes(df_clean)
                csv_bytes = core_engine.generate_target_csv_bytes(df_clean)
                clean_rows_count = len(df_clean)

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
                    dl_csv_path = os.path.join(user_dl_dir, final_csv_filename)
                    with open(dl_csv_path, "wb") as f_out:
                        f_out.write(csv_bytes)
                    saved_paths.append(dl_csv_path)
                except Exception:
                    pass

                try:
                    kaif_path = os.path.join(kaif_dir, final_filename)
                    with open(kaif_path, "wb") as f_out:
                        f_out.write(excel_bytes)
                    saved_paths.append(kaif_path)
                except Exception:
                    pass

                try:
                    kaif_csv_path = os.path.join(kaif_dir, final_csv_filename)
                    with open(kaif_csv_path, "wb") as f_out:
                        f_out.write(csv_bytes)
                    saved_paths.append(kaif_csv_path)
                except Exception:
                    pass

                # Guaranteed Save to OUTPUT_DIR
                cached_path = os.path.join(OUTPUT_DIR, final_filename)
                cached_csv_path = os.path.join(OUTPUT_DIR, final_csv_filename)
                try:
                    with open(cached_path, "wb") as f_out:
                        f_out.write(excel_bytes)
                except Exception:
                    pass
                try:
                    with open(cached_csv_path, "wb") as f_out:
                        f_out.write(csv_bytes)
                except Exception:
                    pass

                # Free byte arrays from memory immediately after saving to disk!
                del excel_bytes, csv_bytes
                gc.collect()

                processed_list.append({
                    "filename": final_filename,
                    "original_filename": fname,
                    "csv_filename": final_csv_filename,
                    "vendor_name": user_base,
                    "bytes": None,
                    "csv_bytes": None,
                    "cached_path": cached_path,
                    "cached_csv_path": cached_csv_path,
                    "stats": stats,
                    "rows_count": clean_rows_count,
                    "saved_paths": saved_paths,
                    "df_clean": df_clean,
                    "date_badge_type": date_badge_type,
                    "date_badge_label": date_badge_label,
                    "has_mismatch": has_mm,
                    "mismatches": mm_list,
                    "mismatch_summary": mm_summary
                })

                progress_bar.progress((f_idx + 1) / len(raw_files))

            t_end = time.time()
            elapsed_sec = round(t_end - t_start, 2)

            # Pre-build on-disk ZIP files if multiple files in batch
            if len(processed_list) > 1:
                u_slug = current_user.get("username", "kaif")
                ts = int(time.time())
                zip_xl_path = os.path.join(OUTPUT_DIR, f"batch_excel_{u_slug}_{ts}.zip")
                zip_csv_path = os.path.join(OUTPUT_DIR, f"batch_csv_{u_slug}_{ts}.zip")
                try:
                    with zipfile.ZipFile(zip_xl_path, "w", zipfile.ZIP_DEFLATED) as zf_xl:
                        for p_item in processed_list:
                            cp = p_item.get("cached_path")
                            if cp and os.path.exists(cp):
                                zf_xl.write(cp, arcname=p_item["filename"])
                    st.session_state.batch_zip_xl_path = zip_xl_path
                except Exception:
                    st.session_state.batch_zip_xl_path = None

                try:
                    with zipfile.ZipFile(zip_csv_path, "w", zipfile.ZIP_DEFLATED) as zf_csv:
                        for p_item in processed_list:
                            cp_csv = p_item.get("cached_csv_path")
                            if cp_csv and os.path.exists(cp_csv):
                                zf_csv.write(cp_csv, arcname=p_item["csv_filename"])
                    st.session_state.batch_zip_csv_path = zip_csv_path
                except Exception:
                    st.session_state.batch_zip_csv_path = None
            else:
                st.session_state.batch_zip_xl_path = None
                st.session_state.batch_zip_csv_path = None

            # Log audit trail for Admin view
            now_log_dt = datetime.datetime.utcnow() + datetime.timedelta(hours=5, minutes=30)
            end_time_log = now_log_dt.strftime("%H:%M:%S")
            start_time_log = (now_log_dt - datetime.timedelta(seconds=float(elapsed_sec))).strftime("%H:%M:%S")
            time_range_log = f"{start_time_log} - {end_time_log} ({round(float(elapsed_sec), 1)}s)"

            for p_item in processed_list:
                am.log_processing_activity(
                    username=current_user.get("username", "unknown"),
                    user_name=current_user.get("name", "Unknown"),
                    output_filename=p_item["filename"],
                    vendor_name=p_item.get("vendor_name", ""),
                    clean_rows=p_item["stats"].get("final_clean_rows", p_item.get("rows_count", 0)),
                    qc_excluded=p_item["stats"].get("qc_excluded", 0),
                    elapsed_sec=elapsed_sec,
                    start_time=start_time_log,
                    end_time=end_time_log,
                    time_range=time_range_log
                )

            # Store in session state for this user's browser & disk cache
            st.session_state.pop("cache_cleared", None)
            st.session_state.processed_results = processed_list
            st.session_state.elapsed_time = elapsed_sec
            st.session_state.just_processed = True
            save_user_processed_cache(current_user.get("username", "kaif"), processed_list, elapsed_sec)
            
            # Reset file uploader so next batch starts with a fresh, clean uploader!
            st.session_state.uploader_key += 1
            gc.collect()
            st.rerun()
