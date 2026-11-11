import os
import json
import time
import hmac
import struct
import base64
import random
import smtplib
import hashlib
import datetime
from email.mime.text import MIMEText

DB_FILE = os.path.join(os.path.dirname(__file__), "users_db.json")
OTP_CACHE_FILE = os.path.join(os.path.dirname(__file__), "email_otp_cache.json")
SMTP_CONFIG_FILE = os.path.join(os.path.dirname(__file__), "smtp_config.json")
HISTORY_FILE = os.path.join(os.path.dirname(__file__), "processing_history.json")
TOKEN_STORE_FILE = os.path.join(os.path.dirname(__file__), "session_tokens.json")

ADMIN_EMAIL = "kaifofficial4@gmail.com"
MASTER_ADMIN_USERNAME = "kaif"
MASTER_TOTP_SECRET = "XCASLSAZCIRXLXFV"

HMAC_SESSION_SECRET = "KaifTelecomPortal_SuperSecret_HMAC_Key_2026_SecureSession"

def get_device_fingerprint(client_ip="", user_agent=""):
    ip_part = ""
    if client_ip:
        ip_clean = str(client_ip).split(",")[0].strip()
        if "." in ip_clean:
            octets = ip_clean.split(".")
            if len(octets) >= 2:
                ip_part = ".".join(octets[:2])
        elif ":" in ip_clean:
            hextets = ip_clean.split(":")
            if len(hextets) >= 3:
                ip_part = ":".join(hextets[:3])
    ua_part = str(user_agent).strip()[:60]
    return hashlib.sha256(f"{ip_part}|{ua_part}".encode()).hexdigest()[:16]

def create_session_token(username, client_ip="", user_agent="", days_valid=30):
    """Generates a clean, URL-safe alphanumeric HEX token cryptographically locked to client IP subnet and User-Agent."""
    exp_timestamp = int(time.time()) + (days_valid * 86400)
    fp_hash = get_device_fingerprint(client_ip, user_agent)
    payload = f"{username.strip().lower()}:{exp_timestamp}:{fp_hash}"
    sig = hmac.new(HMAC_SESSION_SECRET.encode(), payload.encode(), hashlib.sha256).hexdigest()
    raw = f"{payload}:{sig}"
    return raw.encode("utf-8").hex()

def verify_session_token(token_str, client_ip="", user_agent=""):
    """Stateless verification with strict device fingerprint validation (blocks access from other devices)."""
    if not token_str:
        return None
    try:
        if isinstance(token_str, (list, tuple)):
            token_str = token_str[0]
        token_str = str(token_str).strip()
        raw = None
        # 1. Try Hex first
        try:
            raw = bytes.fromhex(token_str).decode("utf-8")
        except Exception:
            pass
        # 2. Try Base64 fallback
        if not raw:
            try:
                padded = token_str + "=" * (-len(token_str) % 4)
                raw = base64.urlsafe_b64decode(padded.encode()).decode("utf-8")
            except Exception:
                pass
        if not raw or ":" not in raw:
            return None
        parts = raw.split(":")
        
        # 4-part device-locked token
        if len(parts) == 4:
            username, exp_str, fp_hash, sig = parts
            exp_timestamp = int(exp_str)
            if time.time() > exp_timestamp:
                return None
            payload = f"{username}:{exp_str}:{fp_hash}"
            expected_sig = hmac.new(HMAC_SESSION_SECRET.encode(), payload.encode(), hashlib.sha256).hexdigest()
            if not hmac.compare_digest(sig, expected_sig):
                return None
            # Validate device fingerprint if client context is present
            if client_ip or user_agent:
                cur_fp = get_device_fingerprint(client_ip, user_agent)
                if not hmac.compare_digest(fp_hash, cur_fp):
                    return None
        # Legacy 3-part token
        elif len(parts) == 3:
            username, exp_str, sig = parts
            exp_timestamp = int(exp_str)
            if time.time() > exp_timestamp:
                return None
            payload = f"{username}:{exp_str}"
            expected_sig = hmac.new(HMAC_SESSION_SECRET.encode(), payload.encode(), hashlib.sha256).hexdigest()
            if not hmac.compare_digest(sig, expected_sig):
                return None
        else:
            return None
        
        db = load_db()
        users = db.get("users", {})
        u = username.lower()
        if u in users:
            u_info = users[u]
            if u_info.get("status") == "active":
                return {
                    "username": u_info["username"],
                    "name": u_info.get("name", u_info["username"].capitalize()),
                    "role": u_info.get("role", "user"),
                    "email": u_info.get("email", ""),
                    "session_expiry": datetime.datetime.fromtimestamp(exp_timestamp).isoformat()
                }
    except Exception:
        pass
    return None

# --- Isolated Per-Browser Session Management ---
def register_session(token, user_info):
    try:
        store = {}
        if os.path.exists(TOKEN_STORE_FILE):
            with open(TOKEN_STORE_FILE, "r", encoding="utf-8") as f:
                store = json.load(f)
        store[token] = user_info
        with open(TOKEN_STORE_FILE, "w", encoding="utf-8") as f:
            json.dump(store, f, indent=2)
    except Exception:
        pass

def get_session_user(token, client_ip="", user_agent=""):
    if not token:
        return None
    # 1. Stateless cryptographic token verification with device binding
    stateless_user = verify_session_token(token, client_ip=client_ip, user_agent=user_agent)
    if stateless_user:
        return stateless_user

    # 2. Fallback to file-based store
    try:
        if os.path.exists(TOKEN_STORE_FILE):
            with open(TOKEN_STORE_FILE, "r", encoding="utf-8") as f:
                store = json.load(f)
            user_info = store.get(token)
            if user_info and not is_session_expired(user_info.get("session_expiry")):
                return user_info
            elif user_info:
                del store[token]
                with open(TOKEN_STORE_FILE, "w", encoding="utf-8") as f:
                    json.dump(store, f, indent=2)
    except Exception:
        pass
    return None

def destroy_session(token):
    if not token:
        return
    try:
        if os.path.exists(TOKEN_STORE_FILE):
            with open(TOKEN_STORE_FILE, "r", encoding="utf-8") as f:
                store = json.load(f)
            if token in store:
                del store[token]
                with open(TOKEN_STORE_FILE, "w", encoding="utf-8") as f:
                    json.dump(store, f, indent=2)
    except Exception:
        pass

# --- TOTP (Authy / Google Authenticator RFC 6238) ---
def gen_totp_secret():
    b32_chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ234567'
    return ''.join(random.choice(b32_chars) for _ in range(16))

def get_master_totp_secret():
    """Returns Admin's Master Authy secret key which is shared across all users."""
    try:
        if os.path.exists(DB_FILE):
            with open(DB_FILE, "r", encoding="utf-8") as f:
                db = json.load(f)
            kaif_info = db.get("users", {}).get(MASTER_ADMIN_USERNAME)
            if kaif_info and kaif_info.get("totp_secret"):
                return kaif_info["totp_secret"]
    except Exception:
        pass
    return MASTER_TOTP_SECRET

def get_totp_token(secret, intervals_no=None):
    if intervals_no is None:
        intervals_no = int(time.time()) // 30
    try:
        key = base64.b32decode(secret, True)
        msg = struct.pack('>Q', intervals_no)
        h = hmac.new(key, msg, hashlib.sha1).digest()
        o = h[19] & 15
        h = (struct.unpack('>I', h[o:o+4])[0] & 0x7fffffff) % 1000000
        return f"{h:06d}"
    except Exception:
        return "000000"

def verify_totp(secret, token, window=1):
    if not token or len(str(token).strip()) != 6:
        return False
    current = int(time.time()) // 30
    for i in range(-window, window + 1):
        if get_totp_token(secret, current + i) == str(token).strip():
            return True
    return False

def get_totp_uri(username, secret, issuer="RateStandardizer"):
    return f"otpauth://totp/{issuer}:{username}?secret={secret}&issuer={issuer}"

# --- Weekly Monday Session Management ---
def get_next_monday_expiry():
    now = datetime.datetime.now()
    days_until_monday = (7 - now.weekday()) % 7
    if days_until_monday == 0:
        days_until_monday = 7
    next_monday = (now + datetime.timedelta(days=days_until_monday)).replace(hour=0, minute=0, second=0, microsecond=0)
    return next_monday.isoformat()

def is_session_expired(expiry_iso):
    if not expiry_iso:
        return True
    try:
        exp = datetime.datetime.fromisoformat(expiry_iso)
        return datetime.datetime.now() >= exp
    except Exception:
        return True

# --- Database Initialization & Management ---
def init_db():
    master_secret = MASTER_TOTP_SECRET
    if not os.path.exists(DB_FILE):
        default_db = {
            "users": {
                "kaif": {
                    "username": "kaif",
                    "name": "Kaif",
                    "password": "admin123",
                    "role": "admin",
                    "totp_secret": master_secret,
                    "email": ADMIN_EMAIL,
                    "created_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "status": "active"
                }
            },
            "reset_requests": []
        }
        with open(DB_FILE, "w", encoding="utf-8") as f:
            json.dump(default_db, f, indent=2)
    else:
        try:
            with open(DB_FILE, "r", encoding="utf-8") as f:
                db = json.load(f)
            updated = False
            master_secret = get_master_totp_secret()
            for u, data in db.get("users", {}).items():
                # Enforce master TOTP secret across all users
                if data.get("totp_secret") != master_secret:
                    data["totp_secret"] = master_secret
                    updated = True
                if "password" not in data:
                    data["password"] = "admin123"
                    updated = True
                if u == "kaif":
                    data["email"] = ADMIN_EMAIL
                    data["role"] = "admin"
                    updated = True
            if updated:
                with open(DB_FILE, "w", encoding="utf-8") as f:
                    json.dump(db, f, indent=2)
        except Exception:
            pass

def load_db():
    init_db()
    try:
        with open(DB_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        init_db()
        with open(DB_FILE, "r", encoding="utf-8") as f:
            return json.load(f)

def save_db(data):
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)

# --- 2-Step Login Workflow ---
def verify_credentials(username, password):
    db = load_db()
    u = username.strip().lower()
    users = db.get("users", {})
    
    # Flexible username matching (supports "arhama", "v_arhama", "aliba", "v_aliba", "kaif")
    matched_key = None
    if u in users:
        matched_key = u
    elif f"v_{u}" in users:
        matched_key = f"v_{u}"
    else:
        for k, v in users.items():
            if k.lower() == u or k.lower().replace("v_", "") == u or v.get("name", "").strip().lower() == u:
                matched_key = k
                break

    if not matched_key:
        return False, "Username not found. Please verify spelling."
    
    user_info = users[matched_key]
    if user_info.get("status") != "active":
        return False, "This account is inactive. Please contact Administrator."
    
    saved_pw = user_info.get("password", "")
    if password == saved_pw:
        return True, {
            "username": user_info.get("username", matched_key),
            "name": user_info.get("name", matched_key.capitalize()),
            "role": user_info.get("role", "user"),
            "email": user_info.get("email", ""),
            "totp_secret": user_info.get("totp_secret", get_master_totp_secret())
        }
    else:
        return False, "Incorrect password. Please try again."

def complete_totp_login(user_info, totp_code, client_ip="", user_agent=""):
    master_secret = get_master_totp_secret()
    user_secret = user_info.get("totp_secret") or master_secret
    
    # Check against master admin secret or user secret
    if verify_totp(master_secret, totp_code) or verify_totp(user_secret, totp_code):
        token = create_session_token(user_info["username"], client_ip=client_ip, user_agent=user_agent, days_valid=30)
        expiry = (datetime.datetime.now() + datetime.timedelta(days=30)).isoformat()
        return True, {
            "token": token,
            "username": user_info["username"],
            "name": user_info["name"],
            "role": user_info["role"],
            "email": user_info.get("email", ""),
            "session_expiry": expiry
        }
    else:
        return False, "Invalid 6-digit Authy code. Please check Admin's Authy app."

# --- User Management (Admin Power) ---
def create_user(username, name, password, role="user", email=""):
    db = load_db()
    u = username.strip().lower()
    if not u or not password:
        return False, "Both Username and Password are required."
    if u in db.get("users", {}):
        return False, f"User '{u}' already exists."
    
    master_secret = get_master_totp_secret()
    db["users"][u] = {
        "username": u,
        "name": name.strip() if name else u.capitalize(),
        "password": password.strip(),
        "role": role if role in ["admin", "user"] else "user",
        "totp_secret": master_secret,
        "email": email.strip(),
        "created_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "status": "active"
    }
    save_db(db)
    return True, f"User '{u}' created successfully with Admin's shared Authy 2FA."

def update_password(username, new_password):
    db = load_db()
    u = username.strip().lower()
    if u not in db.get("users", {}):
        return False, "User not found."
    if not new_password or len(new_password.strip()) < 3:
        return False, "Password must be at least 3 characters."
    
    db["users"][u]["password"] = new_password.strip()
    # Resolve any pending reset request
    db["reset_requests"] = [r for r in db.get("reset_requests", []) if r.get("username") != u]
    save_db(db)
    return True, f"Password updated successfully for '{u}'."

def update_user_name(username, new_name):
    """Allows Admin to edit the display / full name of any user or Admin themselves."""
    db = load_db()
    u = username.strip().lower()
    if u not in db.get("users", {}):
        return False, "User not found."
    if not new_name or not new_name.strip():
        return False, "Name cannot be empty."
    
    clean_name = new_name.strip()
    db["users"][u]["name"] = clean_name
    save_db(db)
    
    # Also update any active session token matching this user
    try:
        if os.path.exists(TOKEN_STORE_FILE):
            with open(TOKEN_STORE_FILE, "r", encoding="utf-8") as f:
                store = json.load(f)
            for tok, uinfo in store.items():
                if uinfo.get("username") == u:
                    uinfo["name"] = clean_name
            with open(TOKEN_STORE_FILE, "w", encoding="utf-8") as f:
                json.dump(store, f, indent=2)
    except Exception:
        pass
        
    return True, f"Name updated to '{clean_name}' for user '{u}'."

def delete_user(username):
    db = load_db()
    u = username.strip().lower()
    if u == "kaif":
        return False, "Master Admin account ('kaif') cannot be deleted."
    if u in db.get("users", {}):
        del db["users"][u]
        save_db(db)
        return True, f"User '{u}' has been deleted."
    return False, "User not found."

def get_all_users():
    db = load_db()
    return db.get("users", {})

# --- Password Reset Option 1: Request to Admin ---
def request_password_reset(username, note=""):
    db = load_db()
    u = username.strip().lower()
    if u not in db.get("users", {}):
        return False, "This username is not registered."
    
    reqs = db.get("reset_requests", [])
    for r in reqs:
        if r.get("username") == u and r.get("status") == "pending":
            return True, "A password reset request for this user is already pending with Admin."
            
    reqs.append({
        "username": u,
        "name": db["users"][u].get("name", u),
        "current_password": db["users"][u].get("password", ""),
        "requested_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "note": note.strip(),
        "status": "pending"
    })
    db["reset_requests"] = reqs
    save_db(db)
    return True, "Password reset request submitted to Admin successfully."

def get_pending_reset_requests():
    db = load_db()
    return [r for r in db.get("reset_requests", []) if r.get("status") == "pending"]

def dismiss_reset_request(username):
    db = load_db()
    u = username.strip().lower()
    db["reset_requests"] = [r for r in db.get("reset_requests", []) if r.get("username") != u]
    save_db(db)
    return True

# --- Password Reset Option 2: Email OTP to kaifofficial4@gmail.com ---
def send_email_via_smtp(to_email, subject, body):
    """Sends email if SMTP config exists, otherwise returns False."""
    if os.path.exists(SMTP_CONFIG_FILE):
        try:
            with open(SMTP_CONFIG_FILE, "r") as f:
                cfg = json.load(f)
            sender = cfg.get("sender_email")
            pw = cfg.get("app_password")
            if sender and pw:
                msg = MIMEText(body)
                msg["Subject"] = subject
                msg["From"] = sender
                msg["To"] = to_email
                with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
                    server.login(sender, pw)
                    server.send_message(msg)
                return True
        except Exception:
            pass
    return False

def generate_email_otp(username):
    db = load_db()
    u = username.strip().lower()
    if u not in db.get("users", {}):
        return False, "This username is not registered.", None
    
    otp = f"{random.randint(100000, 999999)}"
    expiry = (datetime.datetime.now() + datetime.timedelta(minutes=15)).isoformat()
    
    cache = {
        "username": u,
        "otp": otp,
        "expiry": expiry,
        "created_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "target_email": ADMIN_EMAIL
    }
    with open(OTP_CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, indent=2)
    
    # Try sending via SMTP to kaifofficial4@gmail.com
    subject = f"Password Reset OTP for User: {u}"
    body = f"Hello Kaif,\n\nUser '{u}' has requested a password reset on Rate Standardizer Portal.\n\nOTP Code: {otp}\nValid for: 15 minutes.\n\nIf this was not requested, please ignore."
    send_email_via_smtp(ADMIN_EMAIL, subject, body)
    
    return True, f"OTP has been dispatched to {ADMIN_EMAIL}. Enter the code below to reset password.", otp

def verify_email_otp(username, entered_otp):
    if not os.path.exists(OTP_CACHE_FILE):
        return False, "No active OTP found. Please click 'Send OTP' first."
    try:
        with open(OTP_CACHE_FILE, "r", encoding="utf-8") as f:
            cache = json.load(f)
        
        u = username.strip().lower()
        if cache.get("username") != u:
            return False, "OTP mismatch. This OTP was generated for another user."
        
        exp = datetime.datetime.fromisoformat(cache.get("expiry"))
        if datetime.datetime.now() > exp:
            return False, "OTP has expired. Please request a new OTP."
        
        if str(cache.get("otp")).strip() == str(entered_otp).strip():
            return True, "OTP verified successfully!"
        else:
            return False, "Invalid OTP. Please check and try again."
    except Exception as e:
        return False, f"OTP verification error: {e}"

def get_last_active_otp():
    """Helper for Admin to view current active OTP in dashboard."""
    if os.path.exists(OTP_CACHE_FILE):
        try:
            with open(OTP_CACHE_FILE, "r", encoding="utf-8") as f:
                cache = json.load(f)
            exp = datetime.datetime.fromisoformat(cache.get("expiry"))
            if datetime.datetime.now() <= exp:
                return cache
        except Exception:
            pass
    return None

OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "generated_output")

# --- Strict 7-Day Auto-Retention Policy ---
def prune_history_older_than_days(days=7):
    """
    Strict 7-Day Retention Policy:
    Automatically deletes rate sheet records and their physical files from disk
    once they are older than 7 days (168 hours).
    Example: Sheet uploaded on the 11th stays until the 18th, and is deleted on the 18th/19th.
    """
    try:
        if not os.path.exists(HISTORY_FILE):
            return []
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            history = json.load(f)
        if not history:
            return []

        now_ist = datetime.datetime.utcnow() + datetime.timedelta(hours=5, minutes=30)
        cutoff_dt = now_ist - datetime.timedelta(days=days)

        kept_history = []
        expired_files = []

        for entry in history:
            ts_str = entry.get("timestamp", "")
            entry_dt = None
            if ts_str:
                try:
                    entry_dt = datetime.datetime.strptime(str(ts_str).strip()[:19], "%Y-%m-%d %H:%M:%S")
                except Exception:
                    try:
                        entry_dt = datetime.datetime.strptime(str(ts_str).strip()[:10], "%Y-%m-%d")
                    except Exception:
                        pass
            
            if entry_dt is not None:
                if entry_dt >= cutoff_dt:
                    kept_history.append(entry)
                else:
                    # Older than 7 days: expired!
                    out_f = entry.get("output_filename")
                    if out_f:
                        expired_files.append(out_f)
            else:
                kept_history.append(entry)

        # Delete expired physical files from OUTPUT_DIR
        for fn in expired_files:
            try:
                p_xl = os.path.join(OUTPUT_DIR, fn)
                if os.path.exists(p_xl):
                    os.remove(p_xl)
            except Exception:
                pass
            try:
                csv_fn = fn.rsplit(".", 1)[0] + ".csv"
                p_csv = os.path.join(OUTPUT_DIR, csv_fn)
                if os.path.exists(p_csv):
                    os.remove(p_csv)
            except Exception:
                pass

        if len(kept_history) != len(history):
            with open(HISTORY_FILE, "w", encoding="utf-8") as f:
                json.dump(kept_history, f, indent=2)

        return kept_history
    except Exception:
        return []

# --- User Activity & Processing Audit Trail ---
def log_processing_activity(username, user_name, output_filename, vendor_name, clean_rows, qc_excluded, elapsed_sec):
    """Logs each rate sheet standardization event for Admin audit and auto-prunes records > 7 days."""
    try:
        history = []
        if os.path.exists(HISTORY_FILE):
            try:
                with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                    history = json.load(f)
            except Exception:
                history = []
        
        now_ist = datetime.datetime.utcnow() + datetime.timedelta(hours=5, minutes=30)
        entry = {
            "timestamp": now_ist.strftime("%Y-%m-%d %H:%M:%S"),
            "username": username,
            "user_name": user_name,
            "vendor_name": vendor_name,
            "output_filename": output_filename,
            "clean_rows": clean_rows,
            "qc_excluded": qc_excluded,
            "elapsed_sec": elapsed_sec
        }
        history.insert(0, entry)
        with open(HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(history, f, indent=2)

        # Strictly enforce 7-day retention
        prune_history_older_than_days(days=7)
    except Exception:
        pass

def get_processing_history(days=7):
    """Returns all logged processing events within the rolling 7-day retention window, newest first."""
    return prune_history_older_than_days(days=days)

def clear_processing_history():
    """Allows Admin to reset or clear the audit log."""
    try:
        if os.path.exists(HISTORY_FILE):
            os.remove(HISTORY_FILE)
            return True
    except Exception:
        pass
    return False
