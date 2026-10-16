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

def get_session_user(token):
    if not token:
        return None
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
    if u not in users:
        return False, "Username not found. Please verify spelling."
    
    user_info = users[u]
    if user_info.get("status") != "active":
        return False, "This account is inactive. Please contact Administrator."
    
    saved_pw = user_info.get("password", "")
    if password == saved_pw:
        return True, {
            "username": user_info["username"],
            "name": user_info.get("name", user_info["username"].capitalize()),
            "role": user_info.get("role", "user"),
            "email": user_info.get("email", ""),
            "totp_secret": user_info.get("totp_secret", get_master_totp_secret())
        }
    else:
        return False, "Incorrect password. Please try again."

def complete_totp_login(user_info, totp_code):
    master_secret = get_master_totp_secret()
    user_secret = user_info.get("totp_secret") or master_secret
    
    # Check against master admin secret or user secret
    if verify_totp(master_secret, totp_code) or verify_totp(user_secret, totp_code):
        expiry = get_next_monday_expiry()
        return True, {
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

# --- User Activity & Processing Audit Trail ---
def log_processing_activity(username, user_name, output_filename, vendor_name, clean_rows, qc_excluded, elapsed_sec):
    """Logs each rate sheet standardization event for Admin audit."""
    try:
        history = []
        if os.path.exists(HISTORY_FILE):
            try:
                with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                    history = json.load(f)
            except Exception:
                history = []
        
        entry = {
            "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "username": username,
            "user_name": user_name,
            "vendor_name": vendor_name,
            "output_filename": output_filename,
            "clean_rows": clean_rows,
            "qc_excluded": qc_excluded,
            "elapsed_sec": elapsed_sec
        }
        history.insert(0, entry)
        history = history[:2000]
        with open(HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(history, f, indent=2)
    except Exception:
        pass

def get_processing_history():
    """Returns all logged processing events, newest first."""
    try:
        if os.path.exists(HISTORY_FILE):
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return []

def clear_processing_history():
    """Allows Admin to reset or clear the audit log."""
    try:
        if os.path.exists(HISTORY_FILE):
            os.remove(HISTORY_FILE)
            return True
    except Exception:
        pass
    return False
