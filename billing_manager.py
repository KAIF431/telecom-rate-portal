import os
import io
import json
import re
import threading
import pandas as pd
from collections import Counter

# Core file paths
APP_DIR = os.path.dirname(__file__)
MASTER_EXCEL_PATH = os.path.join(APP_DIR, "billing_master.xlsx")
TEMPLATE_EXCEL_PATH = os.path.join(APP_DIR, "templates", "Country Billing.xlsx")
VENDOR_RATES_EXCEL_PATH = r"C:\vendor Rates\Country Billing.xlsx"
MASTER_JSON_PATH = os.path.join(APP_DIR, "billing_master.json")

_MASTER_CACHE = None

def normalize_dest_key(text):
    if not text:
        return ""
    s = re.sub(r'[^a-z0-9]+', ' ', str(text).lower()).strip()
    return re.sub(r'\s+', ' ', s)

def clean_dial_code_str(code):
    if code is None or pd.isna(code):
        return ""
    s = str(code).split('.')[0].strip()
    return re.sub(r'[^\d]', '', s)

def load_master_billing(force_reload=False):
    """
    Loads master billing rules from fast JSON cache or Excel file with automatic caching.
    Supports both 4-column (Dial Code, Summary, Minimum, Increment) and 3-column legacy formats.
    """
    global _MASTER_CACHE
    if _MASTER_CACHE is not None and not force_reload:
        return _MASTER_CACHE

    # 1. Try fast JSON cache first
    if not force_reload and os.path.exists(MASTER_JSON_PATH):
        try:
            with open(MASTER_JSON_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                _MASTER_CACHE = {
                    "code_rules": {str(k): (v[0], v[1], v[2] if len(v) > 2 else "") for k, v in data.get("code_rules", {}).items()},
                    "dest_rules": {k: (v[0], v[1]) for k, v in data.get("dest_rules", {}).items()},
                    "countries": {k: (v[0], v[1]) for k, v in data.get("countries", {}).items()},
                    "raw_list": data.get("raw_list", [])
                }
                return _MASTER_CACHE
        except Exception:
            pass

    # 2. Load from Excel (check vendor rates, master excel, or template excel)
    excel_source = None
    for p in [MASTER_EXCEL_PATH, VENDOR_RATES_EXCEL_PATH, TEMPLATE_EXCEL_PATH]:
        if os.path.exists(p):
            excel_source = p
            break

    code_rules = {}
    dest_rules = {}
    country_counts = {}
    raw_list = []

    if excel_source and os.path.exists(excel_source):
        try:
            df = pd.read_excel(excel_source)
            has_dial_code = False
            code_col = None
            dest_col = None
            min_col = None
            inc_col = None

            cols_lower = [str(c).lower().strip() for c in df.columns]
            for idx, c in enumerate(cols_lower):
                if any(k in c for k in ['dial code', 'code', 'dialcode', 'prefix']):
                    code_col = df.columns[idx]
                    has_dial_code = True
                elif any(k in c for k in ['destination', 'summary', 'country', 'region']):
                    dest_col = df.columns[idx]
                elif any(k in c for k in ['minimum', 'min', 'initial']):
                    min_col = df.columns[idx]
                elif any(k in c for k in ['increment', 'inc', 'step', 'subsequent']):
                    inc_col = df.columns[idx]

            # Fallbacks if columns weren't matched by name
            if dest_col is None:
                dest_col = df.columns[1] if has_dial_code and len(df.columns) > 1 else df.columns[0]
            if min_col is None:
                min_col = df.columns[2] if has_dial_code and len(df.columns) > 2 else (df.columns[1] if len(df.columns) > 1 else None)
            if inc_col is None:
                inc_col = df.columns[3] if has_dial_code and len(df.columns) > 3 else (df.columns[2] if len(df.columns) > 2 else None)

            for _, r in df.iterrows():
                d_orig = str(r[dest_col]).strip() if dest_col and pd.notna(r[dest_col]) else ""
                if not d_orig or d_orig.lower() in ['nan', 'none', 'null', '']:
                    continue

                try:
                    m = int(float(r[min_col])) if min_col and pd.notna(r[min_col]) else 1
                    inc = int(float(r[inc_col])) if inc_col and pd.notna(r[inc_col]) else 1
                    c_clean = clean_dial_code_str(r[code_col]) if has_dial_code and code_col else ""

                    norm_k = normalize_dest_key(d_orig)
                    if norm_k:
                        dest_rules[norm_k] = (m, inc)

                    if c_clean:
                        code_rules[c_clean] = (m, inc, d_orig)

                    raw_list.append({
                        "dial_code": c_clean,
                        "destination": d_orig,
                        "minimum": m,
                        "increment": inc
                    })

                    first_w = norm_k.split()[0] if norm_k.split() else ""
                    if first_w:
                        country_counts.setdefault(first_w, []).append((m, inc))
                except Exception:
                    pass
        except Exception as e:
            print("Error loading master billing Excel:", e)

    country_fallbacks = {}
    for c_w, p_list in country_counts.items():
        if p_list:
            top_pulse = Counter(p_list).most_common(1)[0][0]
            country_fallbacks[c_w] = top_pulse

    _MASTER_CACHE = {
        "code_rules": code_rules,
        "dest_rules": dest_rules,
        "countries": country_fallbacks,
        "raw_list": raw_list
    }

    # Save fast JSON cache
    try:
        with open(MASTER_JSON_PATH, "w", encoding="utf-8") as f:
            json.dump({
                "code_rules": {k: [v[0], v[1], v[2]] for k, v in code_rules.items()},
                "dest_rules": {k: [v[0], v[1]] for k, v in dest_rules.items()},
                "countries": {k: [v[0], v[1]] for k, v in country_fallbacks.items()},
                "raw_list": raw_list
            }, f, indent=2)
    except Exception:
        pass

    return _MASTER_CACHE

def invalidate_cache():
    global _MASTER_CACHE
    _MASTER_CACHE = None

def resolve_standard_billing(dial_code=None, dest_str=""):
    """
    Intelligently resolves standard (minimum, increment) for any rate sheet row:
    1. Exact Dial Code Match (e.g. '852' -> (1, 1, 'Hong Kong'))
    2. Longest Prefix Match (LPM) on Dial Code (e.g. '8529123' -> matches '852' (1, 1))
    3. Exact normalized destination match (e.g. 'indonesia fixed' -> (60, 1))
    4. Longest master destination prefix/subset key match (e.g. 'brazil mobile claro' -> matches 'brazil mobile' (30, 6))
    5. Country fallback (e.g. 'indonesia bandung' -> matches top indonesia billing (60, 1))
    """
    data = load_master_billing()
    code_rules = data["code_rules"]
    dest_rules = data["dest_rules"]
    countries = data["countries"]

    c_clean = clean_dial_code_str(dial_code)

    # 1. Exact Dial Code Match
    if c_clean and c_clean in code_rules:
        r = code_rules[c_clean]
        return (r[0], r[1]), f"code_exact({c_clean})"

    # 2. Longest Prefix Match (LPM) on Dial Code
    if c_clean:
        for l in range(len(c_clean) - 1, 0, -1):
            pref = c_clean[:l]
            if pref in code_rules:
                r = code_rules[pref]
                return (r[0], r[1]), f"code_lpm({pref})"

    # 3. Exact Destination Match
    norm = normalize_dest_key(dest_str)
    if norm and norm in dest_rules:
        return dest_rules[norm], "dest_exact"

    # 4. Longest master key matching prefix or contained in destination
    if norm:
        best_k = None
        best_len = 0
        for k in dest_rules.keys():
            if k in norm or norm.startswith(k):
                if len(k) > best_len:
                    best_len = len(k)
                    best_k = k

        if best_k:
            return dest_rules[best_k], f"dest_prefix({best_k})"

        # 5. Country fallback
        first_w = norm.split()[0] if norm.split() else ""
        if first_w in countries:
            return countries[first_w], f"country({first_w})"

    return None, "none"

def learn_new_dial_codes(new_entries):
    """
    Auto-learning engine:
    Permanently learns and registers newly discovered dial codes into the Master Billing Database.
    Syncs across memory cache, JSON cache, and Excel files.
    """
    if not new_entries:
        return 0

    data = load_master_billing()
    code_rules = data["code_rules"]
    dest_rules = data["dest_rules"]
    raw_list = data["raw_list"]

    new_added = 0
    new_rows_for_df = []

    for item in new_entries:
        c_code = clean_dial_code_str(item.get("dial_code"))
        d_name = str(item.get("destination", "")).strip()
        v_m = item.get("minimum")
        v_i = item.get("increment")

        if not c_code or not d_name or v_m is None or v_i is None:
            continue
        try:
            m = int(v_m)
            inc = int(v_i)
            if m <= 0 or inc <= 0:
                continue
        except Exception:
            continue

        if c_code not in code_rules:
            code_rules[c_code] = (m, inc, d_name)
            norm_k = normalize_dest_key(d_name)
            if norm_k and norm_k not in dest_rules:
                dest_rules[norm_k] = (m, inc)

            raw_list.append({
                "dial_code": c_code,
                "destination": d_name,
                "minimum": m,
                "increment": inc
            })
            new_rows_for_df.append({
                "Dial Code": int(c_code) if c_code.isdigit() else c_code,
                "Summary": d_name,
                "Minimum": m,
                "Increment": inc
            })
            new_added += 1

    if new_added > 0:
        # Update JSON cache
        try:
            with open(MASTER_JSON_PATH, "w", encoding="utf-8") as f:
                json.dump({
                    "code_rules": {k: [v[0], v[1], v[2]] for k, v in code_rules.items()},
                    "dest_rules": {k: [v[0], v[1]] for k, v in dest_rules.items()},
                    "countries": {k: [v[0], v[1]] for k, v in data["countries"].items()},
                    "raw_list": raw_list
                }, f, indent=2)
        except Exception as e:
            print("Error writing updated master JSON:", e)

        # Append to Excel files in background thread so web UI remains instant (< 0.5s)
        def _bg_excel_sync(rows):
            try:
                df_append = pd.DataFrame(rows)
                for path in [MASTER_EXCEL_PATH, TEMPLATE_EXCEL_PATH, VENDOR_RATES_EXCEL_PATH]:
                    try:
                        if os.path.exists(path):
                            df_ex = pd.read_excel(path)
                            df_comb = pd.concat([df_ex, df_append], ignore_index=True).drop_duplicates(subset=['Dial Code'], keep='first')
                            df_comb.to_excel(path, index=False)
                        else:
                            os.makedirs(os.path.dirname(path), exist_ok=True)
                            df_append.to_excel(path, index=False)
                    except Exception:
                        pass
            except Exception as e:
                print("Error in background Excel sync:", e)

        threading.Thread(target=_bg_excel_sync, args=(new_rows_for_df,), daemon=True).start()

    return new_added

def audit_dataframe_billing(df_clean):
    """
    Audits a standardized DataFrame against Kaif's master billing database.
    Matches primarily by Dial Code (with LPM), with fallback to Destination name.
    Automatically discovers and learns brand-new dial codes into the Master Database!
    """
    if df_clean is None or len(df_clean) == 0:
        return False, [], {}, 0

    mismatches = []
    summary_map = {}
    new_candidates_to_learn = []

    dest_list = df_clean['Summary'].tolist() if 'Summary' in df_clean.columns else []
    code_list = df_clean['Dial Code'].tolist() if 'Dial Code' in df_clean.columns else []
    min_list = df_clean['Minimum'].tolist() if 'Minimum' in df_clean.columns else [None] * len(df_clean)
    inc_list = df_clean['Increment'].tolist() if 'Increment' in df_clean.columns else [None] * len(df_clean)

    data = load_master_billing()
    code_rules = data["code_rules"]

    for i in range(len(df_clean)):
        d_name = dest_list[i] if i < len(dest_list) else ""
        c_code = code_list[i] if i < len(code_list) else ""
        v_m = min_list[i] if i < len(min_list) else None
        v_inc = inc_list[i] if i < len(inc_list) else None

        c_clean = clean_dial_code_str(c_code)
        v_m_val = int(v_m) if pd.notna(v_m) and str(v_m).isdigit() else (None if pd.isna(v_m) else v_m)
        v_i_val = int(v_inc) if pd.notna(v_inc) and str(v_inc).isdigit() else (None if pd.isna(v_inc) else v_inc)

        # Check if dial code is completely unknown in master
        if c_clean and c_clean not in code_rules:
            # If not in master, queue for auto-learning if vendor provided valid pulse
            if v_m_val is not None and v_i_val is not None and v_m_val > 0 and v_i_val > 0 and d_name:
                new_candidates_to_learn.append({
                    "dial_code": c_clean,
                    "destination": str(d_name),
                    "minimum": v_m_val,
                    "increment": v_i_val
                })

        # Resolve expected standard billing (by Code LPM, then Destination)
        std_pulse, match_method = resolve_standard_billing(c_code, d_name)
        if std_pulse is not None:
            std_m, std_inc = std_pulse
            if (v_m_val != std_m) or (v_i_val != std_inc):
                v_pulse_str = f"{v_m_val}/{v_i_val}" if v_m_val is not None else "None"
                std_pulse_str = f"{std_m}/{std_inc}"
                mismatches.append({
                    "row_idx": i,
                    "destination": str(d_name),
                    "dial_code": str(c_code),
                    "vendor_min": v_m_val,
                    "vendor_inc": v_i_val,
                    "vendor_pulse": v_pulse_str,
                    "standard_min": std_m,
                    "standard_inc": std_inc,
                    "standard_pulse": std_pulse_str,
                    "match_method": match_method
                })
                key = (v_pulse_str, std_pulse_str)
                summary_map[key] = summary_map.get(key, 0) + 1

    # Auto-learn newly discovered dial codes
    learned_count = 0
    if new_candidates_to_learn:
        learned_count = learn_new_dial_codes(new_candidates_to_learn)

    return (len(mismatches) > 0), mismatches, summary_map, learned_count

def apply_standard_billing_to_dataframe(df_clean, mismatch_list=None):
    """
    Applies master standard billing to all mismatched rows in df_clean in-place.
    Returns the updated df_clean.
    """
    if df_clean is None or len(df_clean) == 0:
        return df_clean

    df_copy = df_clean.copy()
    if mismatch_list is None:
        _, mismatch_list, _, _ = audit_dataframe_billing(df_copy)

    if not mismatch_list:
        return df_copy

    mins = df_copy['Minimum'].tolist() if 'Minimum' in df_copy.columns else [None] * len(df_copy)
    incs = df_copy['Increment'].tolist() if 'Increment' in df_copy.columns else [None] * len(df_copy)

    for item in mismatch_list:
        idx = item["row_idx"]
        if idx < len(mins):
            mins[idx] = item["standard_min"]
            incs[idx] = item["standard_inc"]

    df_copy['Minimum'] = mins
    df_copy['Increment'] = incs
    return df_copy

def get_master_excel_bytes():
    """
    Returns live binary content of the Master Country Billing Excel sheet.
    Can be downloaded by any user or administrator.
    If on-disk files are not found, dynamically generates the Excel sheet from master rules in memory.
    """
    for p in [TEMPLATE_EXCEL_PATH, MASTER_EXCEL_PATH, VENDOR_RATES_EXCEL_PATH]:
        if os.path.exists(p):
            try:
                with open(p, "rb") as f:
                    content = f.read()
                    if content and len(content) > 1000:
                        return content
            except Exception:
                pass
    
    # In-memory generation fallback from JSON rules
    try:
        data = load_master_billing()
        raw_list = data.get("raw_list", [])
        if raw_list:
            df = pd.DataFrame(raw_list)
            df = df.rename(columns={"dial_code": "Dial Code", "destination": "Summary", "minimum": "Minimum", "increment": "Increment"})
            cols = ["Dial Code", "Summary", "Minimum", "Increment"]
            df = df[[c for c in cols if c in df.columns]]
            buf = io.BytesIO()
            df.to_excel(buf, index=False, engine="openpyxl")
            return buf.getvalue()
    except Exception:
        pass
    return b""

def update_or_add_master_rule(dial_code, destination, minimum, increment):
    """Updates or adds a single rule by Dial Code or Destination."""
    d_clean = str(destination).strip()
    c_clean = clean_dial_code_str(dial_code)
    try:
        m = int(minimum)
        inc = int(increment)
    except Exception:
        return False, "Minimum and Increment must be positive integers."

    if not d_clean and not c_clean:
        return False, "Dial Code or Destination cannot be empty."

    data = load_master_billing()
    raw_list = data["raw_list"]
    code_rules = data["code_rules"]
    dest_rules = data["dest_rules"]

    found = False
    for r in raw_list:
        if c_clean and clean_dial_code_str(r.get("dial_code")) == c_clean:
            r["minimum"] = m
            r["increment"] = inc
            if d_clean:
                r["destination"] = d_clean
            found = True
            break
        elif not c_clean and normalize_dest_key(r.get("destination")) == normalize_dest_key(d_clean):
            r["minimum"] = m
            r["increment"] = inc
            found = True
            break

    if not found:
        raw_list.append({
            "dial_code": c_clean,
            "destination": d_clean,
            "minimum": m,
            "increment": inc
        })

    # Update in memory
    if c_clean:
        code_rules[c_clean] = (m, inc, d_clean)
    if d_clean:
        dest_rules[normalize_dest_key(d_clean)] = (m, inc)

    # Save to Excel & JSON
    df_new = pd.DataFrame(raw_list)
    cols = ["Dial Code", "Summary", "Minimum", "Increment"]
    df_new = df_new[["dial_code", "destination", "minimum", "increment"]]
    df_new.columns = cols
    for p in [MASTER_EXCEL_PATH, TEMPLATE_EXCEL_PATH, VENDOR_RATES_EXCEL_PATH]:
        try:
            df_new.to_excel(p, index=False)
        except Exception:
            pass

    invalidate_cache()
    load_master_billing(force_reload=True)
    return True, f"Rule for '{c_clean or d_clean}' updated to {m}/{inc} successfully."

def delete_master_rule(dial_code_or_dest):
    """Deletes a rule from master billing by dial code or destination."""
    q = str(dial_code_or_dest).strip()
    q_digits = clean_dial_code_str(q)
    q_norm = normalize_dest_key(q)

    data = load_master_billing()
    raw_list = [
        r for r in data["raw_list"]
        if (q_digits and clean_dial_code_str(r.get("dial_code")) != q_digits) and
           (not q_digits and normalize_dest_key(r.get("destination")) != q_norm)
    ]

    df_new = pd.DataFrame(raw_list)
    cols = ["Dial Code", "Summary", "Minimum", "Increment"]
    df_new = df_new[["dial_code", "destination", "minimum", "increment"]]
    df_new.columns = cols
    for p in [MASTER_EXCEL_PATH, TEMPLATE_EXCEL_PATH, VENDOR_RATES_EXCEL_PATH]:
        try:
            df_new.to_excel(p, index=False)
        except Exception:
            pass

    invalidate_cache()
    load_master_billing(force_reload=True)
    return True, f"Rule for '{dial_code_or_dest}' deleted."

def save_new_master_excel(file_content):
    """Replaces master billing with an uploaded Excel file."""
    try:
        df = pd.read_excel(file_content)
        if len(df.columns) < 3:
            return False, "Excel file must have at least 3 columns: Dial Code (optional), Destination/Summary, Minimum, Increment."

        has_dial = len(df.columns) >= 4
        if has_dial:
            df.columns = ["Dial Code", "Summary", "Minimum", "Increment"]
        else:
            df.columns = ["Summary", "Minimum", "Increment"]

        for p in [MASTER_EXCEL_PATH, TEMPLATE_EXCEL_PATH, VENDOR_RATES_EXCEL_PATH]:
            try:
                df.to_excel(p, index=False)
            except Exception:
                pass

        invalidate_cache()
        load_master_billing(force_reload=True)
        return True, f"Master billing updated with {len(df)} rules successfully."
    except Exception as e:
        return False, str(e)
