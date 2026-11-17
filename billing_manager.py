import os
import json
import re
import pandas as pd
from collections import Counter

MASTER_EXCEL_PATH = os.path.join(os.path.dirname(__file__), "billing_master.xlsx")
MASTER_JSON_PATH = os.path.join(os.path.dirname(__file__), "billing_master.json")

_MASTER_CACHE = None

def normalize_dest_key(text):
    if not text:
        return ""
    s = re.sub(r'[^a-z0-9]+', ' ', str(text).lower()).strip()
    return re.sub(r'\s+', ' ', s)

def load_master_billing():
    """Loads master billing rules from JSON cache or Excel file, with automatic caching."""
    global _MASTER_CACHE
    if _MASTER_CACHE is not None:
        return _MASTER_CACHE

    # Try fast JSON cache first
    if os.path.exists(MASTER_JSON_PATH):
        try:
            with open(MASTER_JSON_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                _MASTER_CACHE = {
                    "rules": {k: tuple(v) for k, v in data.get("rules", {}).items()},
                    "countries": {k: tuple(v) for k, v in data.get("countries", {}).items()},
                    "raw_list": data.get("raw_list", [])
                }
                return _MASTER_CACHE
        except Exception:
            pass

    # Fallback to Excel
    rules = {}
    country_counts = {}
    raw_list = []

    if os.path.exists(MASTER_EXCEL_PATH):
        try:
            df = pd.read_excel(MASTER_EXCEL_PATH)
            dest_col = df.columns[0]
            min_col = df.columns[1]
            inc_col = df.columns[2]

            for _, r in df.iterrows():
                d_orig = str(r[dest_col]).strip()
                if not d_orig or pd.isna(r[dest_col]):
                    continue
                try:
                    m = int(float(r[min_col]))
                    inc = int(float(r[inc_col]))
                    norm_k = normalize_dest_key(d_orig)
                    rules[norm_k] = (m, inc)
                    raw_list.append({"destination": d_orig, "minimum": m, "increment": inc})

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
        "rules": rules,
        "countries": country_fallbacks,
        "raw_list": raw_list
    }

    # Save fast JSON cache
    try:
        with open(MASTER_JSON_PATH, "w", encoding="utf-8") as f:
            json.dump({
                "rules": rules,
                "countries": country_fallbacks,
                "raw_list": raw_list
            }, f, indent=2)
    except Exception:
        pass

    return _MASTER_CACHE

def invalidate_cache():
    global _MASTER_CACHE
    _MASTER_CACHE = None

def resolve_standard_billing(dest_str):
    """
    Intelligently resolves standard (minimum, increment) for any vendor destination string:
    1. Exact normalized match (e.g. 'indonesia fixed' -> (60, 1))
    2. Longest prefix/subset key match (e.g. 'brazil mobile claro' -> matches 'brazil mobile' (30, 6))
    3. Country fallback (e.g. 'indonesia bandung' -> matches top indonesia billing (60, 1))
    """
    if not dest_str:
        return None, "none"

    data = load_master_billing()
    rules = data["rules"]
    countries = data["countries"]

    norm = normalize_dest_key(dest_str)
    if not norm:
        return None, "none"

    # 1. Exact match
    if norm in rules:
        return rules[norm], "exact"

    # 2. Longest master key matching prefix or contained in vendor destination
    best_k = None
    best_len = 0
    for k, v in rules.items():
        if k in norm or norm.startswith(k):
            if len(k) > best_len:
                best_len = len(k)
                best_k = k

    if best_k:
        return rules[best_k], f"prefix({best_k})"

    # 3. Country fallback
    first_w = norm.split()[0] if norm.split() else ""
    if first_w in countries:
        return countries[first_w], f"country({first_w})"

    return None, "none"

def audit_dataframe_billing(df_clean):
    """
    Audits a standardized DataFrame against Kaif's master billing database.
    Returns:
    - has_mismatch: bool
    - mismatches: list of dicts {row_idx, destination, dial_code, vendor_min, vendor_inc, standard_min, standard_inc}
    - mismatch_summary: dict of {(vendor_pulse, standard_pulse): count}
    """
    if df_clean is None or len(df_clean) == 0:
        return False, [], {}

    mismatches = []
    summary_map = {}

    dest_list = df_clean['Summary'].tolist() if 'Summary' in df_clean.columns else []
    code_list = df_clean['Dial Code'].tolist() if 'Dial Code' in df_clean.columns else []
    min_list = df_clean['Minimum'].tolist() if 'Minimum' in df_clean.columns else [None] * len(df_clean)
    inc_list = df_clean['Increment'].tolist() if 'Increment' in df_clean.columns else [None] * len(df_clean)

    for i in range(len(df_clean)):
        d_name = dest_list[i] if i < len(dest_list) else ""
        c_code = code_list[i] if i < len(code_list) else ""
        v_m = min_list[i] if i < len(min_list) else None
        v_inc = inc_list[i] if i < len(inc_list) else None

        # Resolve expected standard billing
        std_pulse, _ = resolve_standard_billing(d_name)
        if std_pulse is not None:
            std_m, std_inc = std_pulse
            # If vendor pulse differs from master standard pulse
            v_m_val = int(v_m) if pd.notna(v_m) and str(v_m).isdigit() else (None if pd.isna(v_m) else v_m)
            v_i_val = int(v_inc) if pd.notna(v_inc) and str(v_inc).isdigit() else (None if pd.isna(v_inc) else v_inc)

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
                    "standard_pulse": std_pulse_str
                })
                key = (v_pulse_str, std_pulse_str)
                summary_map[key] = summary_map.get(key, 0) + 1

    return (len(mismatches) > 0), mismatches, summary_map

def apply_standard_billing_to_dataframe(df_clean, mismatch_list=None):
    """
    Applies master standard billing to all mismatched rows in df_clean in-place.
    Returns the updated df_clean.
    """
    if df_clean is None or len(df_clean) == 0:
        return df_clean

    df_copy = df_clean.copy()
    if mismatch_list is None:
        _, mismatch_list, _ = audit_dataframe_billing(df_copy)

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

def update_or_add_master_rule(destination, minimum, increment):
    """Updates or adds a single rule and refreshes persistence."""
    d_clean = str(destination).strip()
    if not d_clean:
        return False, "Destination cannot be empty."
    try:
        m = int(minimum)
        inc = int(increment)
    except Exception:
        return False, "Minimum and Increment must be positive integers."

    data = load_master_billing()
    raw_list = data["raw_list"]
    norm_k = normalize_dest_key(d_clean)

    # Check if exists in raw list
    found = False
    for r in raw_list:
        if normalize_dest_key(r["destination"]) == norm_k:
            r["minimum"] = m
            r["increment"] = inc
            found = True
            break

    if not found:
        raw_list.append({"destination": d_clean, "minimum": m, "increment": inc})

    # Save to Excel and reset cache
    df_new = pd.DataFrame(raw_list)
    df_new.columns = ["Destinations", "Minimum", "Increment"]
    df_new.to_excel(MASTER_EXCEL_PATH, index=False)
    invalidate_cache()
    load_master_billing()
    return True, f"Rule for '{d_clean}' updated to {m}/{inc} successfully."

def delete_master_rule(destination):
    """Deletes a rule from master billing."""
    norm_k = normalize_dest_key(destination)
    data = load_master_billing()
    raw_list = [r for r in data["raw_list"] if normalize_dest_key(r["destination"]) != norm_k]

    df_new = pd.DataFrame(raw_list)
    df_new.columns = ["Destinations", "Minimum", "Increment"]
    df_new.to_excel(MASTER_EXCEL_PATH, index=False)
    invalidate_cache()
    load_master_billing()
    return True, f"Rule for '{destination}' deleted."

def save_new_master_excel(file_content):
    """Replaces master billing with an uploaded Excel file."""
    try:
        df = pd.read_excel(file_content)
        if len(df.columns) < 3:
            return False, "Excel file must have 3 columns: Destinations, Minimum, Increment."
        df.iloc[:, 0] = df.iloc[:, 0].astype(str)
        df.iloc[:, 1] = pd.to_numeric(df.iloc[:, 1], errors='coerce').fillna(1).astype(int)
        df.iloc[:, 2] = pd.to_numeric(df.iloc[:, 2], errors='coerce').fillna(1).astype(int)
        df.columns = ["Destinations", "Minimum", "Increment"]
        df.to_excel(MASTER_EXCEL_PATH, index=False)
        invalidate_cache()
        load_master_billing()
        return True, f"Master billing updated with {len(df)} rules successfully."
    except Exception as e:
        return False, str(e)
