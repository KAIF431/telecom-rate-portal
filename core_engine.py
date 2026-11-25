import os
import re
import io
import csv
import gc
import datetime
import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.styles.colors import Color

TEMPLATE_DIR = os.path.join(os.path.dirname(__file__), "templates")
DEFAULT_QC_FILE = os.path.join(TEMPLATE_DIR, "QC exception Vlookup.xlsx")
FALLBACK_QC_FILE = r"C:\vendor Rates\QC exception Vlookup.xlsx"

# Default QC Exception codes (Satellite, High Risk, Freephone, and Austria special codes)
DEFAULT_QC_CODES = [
    # Global Satellite & Maritime prefixes (ITU-T E.164)
    "87", "870", "871", "872", "873", "874", "875", "876", "877", "878", "879",
    "881", "882", "883", "884", "885", "887", "888", "889",
    "969", "979", "999", "1855", "1866", "1877", "1888", "2897", "6993", "6998",
    "388", "800", "991", "808", "684", "1671", "671",
    # Austria Special / High Risk codes
    "436132308993", "436646570601", "436648668402", "436648668411", "436648668412",
    "436648668420", "436648668423", "436648668424", "436648740174", "436649469663",
    "4366499288491", "4366499288492", "4366499288493", "4366499288496", "4366499288497",
    "4366499288498", "4367030947", "4367030981", "436708012001", "436708012012",
    "436708012014", "436708012018", "436708012019", "436708012021", "436708012022",
    "436708012027", "43670801203", "43670801204", "43670801205", "43670801206",
    "43670801207", "43670801208", "43670801209", "43670801201", "4367080121",
    "436708012121", "43670801214", "436708012150", "436708012154", "436708012160",
    "43670801217", "43670801218", "43670801219", "436708012200", "436708012202",
    "436708012203", "436708012204", "436708012205", "436708012206", "436708012208",
    "436708012211", "436708012212", "436708012217", "436708012219", "436708012221",
    "436708012223", "436708012225", "436708012233", "43670801224", "436708012242",
    "436708012250", "436708012251", "436708012257", "436708012258", "436708012262",
    "436708012263", "436708012265", "436708012267", "436708012268", "436708012271",
    "436708012272", "43670801228", "43670801229", "436708012327", "436708012329",
    "436708012330", "436708012333", "436708012334", "436708012340", "436708012342",
    "436708012344", "436708012346", "436708012349", "436708012350", "436708012355",
    "43670801236", "436708012374", "43670801238", "43670801240", "43670801241",
    "436708012410", "436708012427", "43670801243", "43670801244", "43670801245",
    "43670801246", "43670801247", "43670801248", "43670801249", "436708012490",
    "436708012492", "436708012493", "436708012494", "436708012526", "43670801253",
    "43670801254", "436708012551", "436708012552", "43670801256", "43670801257",
    "43670801258", "43670801259", "436708012591", "436708012597", "436708012598",
    "43670801260", "436708012611", "436708012612", "436708012613", "436708012632",
    "436708012660", "436708012674", "43670801268", "43670801269", "436708012710",
    "436708012712", "43670801273", "436708012742", "436708012753", "436708012757",
    "436708012758", "436708012762", "436708012768", "436708012772", "43670801278",
    "436708012793", "436708012795", "436708013092", "436708013094", "436708013096",
    "43670809029", "436708090340", "436708091166", "436708091194", "436708091196",
    "436708092021", "436708092023", "436708092027", "436708092304", "436708093151",
    "436708093155", "436708093156", "436708093157", "43670809336", "436708094044",
    "436708094047", "43670809444", "436708095051", "436708095052", "436708095054",
    "436708095057", "436708095210", "436708095212", "436708095213", "436708095215",
    "43670809613", "43670809622", "436708097123", "43670809747", "436708098012",
    "43670809820", "436708098380", "436708099110", "436708099113", "436708099115",
    "436708099226", "4368110728430", "4368110788534", "4368110891145", "4368120310611",
    "4368120875953",
    # Germany, Italy & Switzerland High Risk & Exception Codes
    "494080808037",
    "4953170222603",
    "496926012260",
    "4975796235199",
    "3950937011199",
    "4112203630798",
    "4118809579492"
]

_CACHED_QC_CODES = None

def load_qc_exception_codes():
    global _CACHED_QC_CODES
    if _CACHED_QC_CODES is not None:
        return _CACHED_QC_CODES
    codes = set(DEFAULT_QC_CODES)
    for p in [DEFAULT_QC_FILE, FALLBACK_QC_FILE]:
        if os.path.exists(p):
            try:
                df = pd.read_excel(p)
                for col in df.columns:
                    if 'code' in str(col).lower():
                        for val in df[col].dropna():
                            c_str = str(val).split('.')[0].strip()
                            if c_str:
                                codes.add(c_str)
                break
            except Exception:
                pass
    _CACHED_QC_CODES = codes
    return codes

def detect_header_row_and_columns(df_raw):
    """
    Universal, fault-tolerant table header detection:
    - Recognizes all telecom column aliases, abbreviations, and vendor typos (Dst Code, Dest Name, County Codes, etc.)
    - Ignores title/banner rows (Rate Table, Trunk Prefix notes, disclaimers)
    - Validates actual data in rows below the candidate header (numeric dial codes + numeric rates)
    - Prevents false-positive headers permanently
    """
    best_row_idx = 0
    best_mapping = {}
    best_total_score = -9999

    for row_idx in range(min(40, len(df_raw))):
        row_vals = [str(x).strip().lower() for x in df_raw.iloc[row_idx].values]
        non_empty = [x for x in row_vals if x not in ['nan', 'none', '', 'null', 'nat']]
        if len(non_empty) < 2:
            continue

        mapping = {}
        kw_score = 0

        for col_idx, val in enumerate(row_vals):
            if val in ['nan', 'none', '', 'null', 'nat']:
                continue
            v = re.sub(r'[^a-z0-9]', ' ', val).strip()
            
            # Skip long disclaimers / policy sentences
            if len(val) > 75:
                continue

            # 1. Status / Changes (check before rate so 'rate changes' is mapped to status)
            if any(k in v for k in ['rate changes', 'rate change', 'changes', 'status', 'comment', 'action', 'change', 'diff status', 'remarks', 'type', 'notes', 'diff']) \
                and not any(bad in v for bad in ['table', 'sheet', 'notice']):
                is_high_status = any(k in v for k in ['rate changes', 'rate change', 'changes', 'status', 'change', 'action', 'diff status', 'diff'])
                curr_is_high_status = mapping.get('status_is_high', False)
                if 'status' not in mapping or (is_high_status and not curr_is_high_status):
                    mapping['status'] = col_idx
                    mapping['status_is_high'] = is_high_status
                    kw_score += 4

            # 2. Destination / Summary with 'name', 'desc', 'location', 'country'
            # (Check BEFORE dial code so 'dst code name', 'dest name', 'code name' are mapped to destination,
            # BUT do not capture code columns like 'destination code', 'dest code', 'country code', and NOT time zones/bands!)
            elif (
                (any(k in v for k in [
                    'destination name', 'dest name', 'dst name', 'dest desc', 'dest description', 'dst desc', 'dst description', 'dst destination', 'destinations name', 'dest summary', 'dst code name', 'dest code name', 'destination code name', 'area name', 'zone name', 'country name', 'city name',
                    'rate summary', 'rates summary', 'rate center', 'destination summary'
                ])
                 or (any(n in v for n in ['name', 'desc', 'description']) and any(d in v for d in ['dest', 'dst', 'code', 'loc', 'country', 'region', 'area', 'zone', 'city']))
                 or (any(k in v for k in ['destination', 'destinations', 'summary', 'location', 'loc name', 'national', 'international', 'country', 'country name', 'description', 'region', 'regions', 'region name', 'area', 'zone'])
                     and not any(c in v for c in ['code', 'dial', 'prefix', 'digit'])))
            ) and (any(n in v for n in ['name', 'desc', 'description']) or not any(k in v for k in ['country code', 'country codes', 'county code', 'county codes', 'countrycode', 'countycode', 'destination code', 'dest code', 'dst code', 'area code', 'area prefix'])) \
              and not any(tb in v for tb in ['time', 'peak', 'offpeak', 'off peak', 'band', 'timezone', 'effective', 'expiry', 'valid']):
                if 'destination' not in mapping:
                    mapping['destination'] = col_idx
                    kw_score += 7
                elif 'destination_detail' not in mapping:
                    mapping['destination_detail'] = col_idx
                    kw_score += 4

            # 3. Dial code keywords (Specific code keywords take higher priority over generic 'prefix')
            elif (any(k in v for k in [
                'complete code', 'complete codes', 'full code', 'full codes', 'complete dial code', 'full dial code', 'complete prefix', 'full prefix',
                'dial code', 'dial codes', 'dialcode', 'dialcodes',
                'dest code', 'dest codes', 'destination code', 'destination codes', 'destcode', 'destcodes',
                'dst code', 'dst codes', 'dstcode', 'dstcodes', 'dst_code', 'dst. code',
                'region code', 'region codes', 'zone code', 'zone codes', 'city code', 'city codes', 'telecom code',
                'area prefix', 'area prefixes', 'dest prefix', 'destination prefix', 'dial prefix',
                'rate key', 'rate keys', 'rate_key', 'ratekey', 'code key', 'code keys', 'dial key', 'dial keys', 'prefix key',
                'dialed digit', 'dialed digits', 'dialled digit', 'dialled digits', 'dialed number', 'dialed numbers',
                'calling code', 'call code', 'dial digit', 'dial digits', 'e164'
            ]) or v in ['code', 'codes', 'dst', 'dst_code', 'dialcode']) \
                and not any(n in v for n in ['name', 'desc', 'description']) \
                and not any(c in v for c in ['country code', 'country codes', 'countrycode', 'county code', 'county codes', 'countycode', 'area code', 'area codes', 'areacode', 'areacodes']):
                
                is_complete = any(cp in v for cp in ['complete code', 'full code', 'complete dial code', 'dial code', 'dialcode', 'region code', 'dest code', 'destination code'])
                curr_is_complete = mapping.get('dial_code_is_complete', False)
                if 'dial_code' not in mapping or (is_complete and not curr_is_complete):
                    mapping['dial_code'] = col_idx
                    mapping['dial_code_is_complete'] = is_complete
                    mapping['dial_code_is_generic_prefix'] = False
                    kw_score += 10 if is_complete else 8

            elif any(k in v for k in ['country code', 'country codes', 'county code', 'county codes', 'countrycode', 'countycode']):
                if 'country_code' not in mapping:
                    mapping['country_code'] = col_idx
                    kw_score += 6

            elif any(k in v for k in ['area code', 'area codes', 'areacode', 'areacodes', 'area prefix', 'area prefixes', 'npa', 'nxx']):
                if 'area_code' not in mapping:
                    mapping['area_code'] = col_idx
                    kw_score += 6

            elif any(k in v for k in ['prefix', 'prefixes']) and not any(bad in v for bad in ['tech prefix', 'trunk prefix', 'routing prefix', 'gw prefix', 'gateway prefix', 'carrier prefix', 'account prefix']):
                is_rate_pfx = 'rate prefix' in v
                if 'dial_code' not in mapping or mapping.get('dial_code_is_generic_prefix', False):
                    if not is_rate_pfx:
                        mapping['dial_code'] = col_idx
                        mapping['dial_code_is_generic_prefix'] = True
                        kw_score += 4
                    elif 'rate_prefix_fallback' not in mapping:
                        mapping['rate_prefix_fallback'] = col_idx

            # 4. Rate keywords (exclude table/sheet banners, decks, cards, zones, and previous rates)
            elif (any(k in v for k in [
                'new rate', 'price usd', 'rate usd', 'price', 'prices', 'rate', 'rates', 'tariff', 'tariffs', 'cost', 'charge', 'usd rate', 'rate min', 'rate per min', 'allday', 'all day', 'alldays', 'all days', 'flat rate', 'flat',
                'inter rate', 'inter. rate', 'inter_rate', 'inter.rate', 'intra rate', 'intra. rate', 'standard rate', 'indet rate',
                'dollar', 'dollars', 'usd', 'eur', 'euro', 'euros'
            ]) or '$' in val or v in ['rate', 'rates', 'price', 'prices', 'tariff', 'usd', 'eur', 'dollar', 'dollars', 'euro', 'euros', 'charge']) \
                and not any(bad in v for bad in ['table', 'sheet', 'deck', 'notice', 'notification', 'prefix', 'plan', 'card', 'zone', 'time', 'increase', 'effective', 'key']) \
                and not re.search(r'\b(?:current|old|prev|previous|existing|base|diff)\b', v):
                if 'rate' not in mapping or 'new rate' in v or 'new price' in v or 'rate usd' in v or 'inter rate' in v:
                    mapping['rate'] = col_idx
                    kw_score += 7

            # 5. Effective Date keywords (strictly ignore expiry/end/to dates, and ignore pulse/increment start dates)
            elif any(k in v for k in [
                'effective from', 'eff date', 'effective date', 'effectiv date', 'eff. date', 'eff_date',
                'application date', 'date activate', 'date', 'eff', 'effectiv', 'pricing valid from',
                'valid from', 'valid date', 'start date', 'active from', 'active date', 'activation date',
                'active', 'activation'
            ]) and not any(bad in v for bad in [
                'expiry', 'expire', 'end', 'until', 'close', 'termination', 'to date', 'effective to',
                'valid to', 'eff to', 'minimum', 'increment', 'increments', 'pulse', 'billing', 'min/inc',
                'notice', 'issue', 'creation', 'created', 'file date', 'sent date', 'export', 'generated',
                'old date', 'prev date', 'previous date', 'last change', 'modify', 'modified', 'download'
            ]) and not bool(re.search(r'\b(?:to|thru|through)\b', v)):
                is_valid_date_col = True
                if v in ['active', 'activation']:
                    sample_check = df_raw.iloc[row_idx + 1 : min(row_idx + 15, len(df_raw)), col_idx].dropna().tolist()
                    has_date_sample = any(
                        isinstance(x, (datetime.date, datetime.datetime, pd.Timestamp)) or 
                        bool(re.search(r'\d{4}[-/]\d{1,2}[-/]\d{1,2}|\d{1,2}[-/]\d{1,2}[-/]\d{2,4}', str(x)))
                        for x in sample_check
                    )
                    if not has_date_sample:
                        is_valid_date_col = False

                if is_valid_date_col:
                    is_high_priority = any(k in v for k in [
                        'eff date', 'effective date', 'eff. date', 'eff_date', 'effective from',
                        'application date', 'pricing valid from', 'valid from'
                    ])
                    curr_is_high = mapping.get('effective_date_is_high', False)
                    if 'effective_date' not in mapping or (is_high_priority and not curr_is_high):
                        mapping['effective_date'] = col_idx
                        mapping['effective_date_is_high'] = is_high_priority
                        kw_score += 5

            # 6. Pulse / Billing keywords
            elif any(k in v for k in [
                'increments', 'increment', 'incremental', 'pulse', 'billing', 'initial billing', 'incremental billing',
                'billing increment', 'minimum increments', 'billing step', 'min billing time', 'min inc', 'min/inc',
                'billing pulse', 'billing unit', 'pulse rate', 'unit', 'resolution', 'minimal time', 'billing resolution',
                'interval', 'intervals', 'billing interval', 'billinc', 'bill inc', 'billincr', 'bill_inc',
                'rounding', 'round', 'round off', 'rounding rule', 'min (sec)', 'min(sec)', 'increment (sec)', 'increment(sec)',
                'minimum', 'initial', 'additional', 'subsequent', 'addl'
            ]) or v in ['mc', 'ci', 'pulse', 'interval', 'step', 'billinc', 'rounding', 'min', 'inc', 'additional', 'subsequent']:
                if 'pulse' not in mapping:
                    mapping['pulse'] = col_idx
                    kw_score += 4

        # Link country + area code or country code standalone
        if ('country_code' in mapping and 'area_code' in mapping) and 'dial_code' not in mapping:
            mapping['dial_code_split'] = True
        elif 'country_code' in mapping and 'dial_code' in mapping and mapping['country_code'] != mapping['dial_code']:
            mapping['dial_code_split'] = True
            mapping['area_code'] = mapping['dial_code']
        elif 'country_code' in mapping and ('dial_code' not in mapping or mapping.get('dial_code_is_generic_prefix', False)) and 'area_code' not in mapping:
            mapping['dial_code'] = mapping['country_code']
            kw_score += 6
        elif 'dial_code' not in mapping and 'rate_prefix_fallback' in mapping:
            mapping['dial_code'] = mapping['rate_prefix_fallback']
            mapping['dial_code_is_generic_prefix'] = True
            kw_score += 4

        # Must have at least a dial code (or split code), plus either rate or destination
        has_code = ('dial_code' in mapping) or ('country_code' in mapping and 'area_code' in mapping)
        has_rate_or_dest = ('rate' in mapping) or ('destination' in mapping)
        if not (has_code and has_rate_or_dest):
            continue

        # Data Validation Below Candidate Row
        data_score = 0
        code_col = mapping.get('dial_code', mapping.get('country_code'))
        rate_col = mapping.get('rate')
        sample_rows = df_raw.iloc[row_idx + 1 : min(row_idx + 11, len(df_raw))]
        
        valid_codes = 0
        valid_rates = 0
        total_sample = len(sample_rows)
        if total_sample > 0:
            if code_col is not None and code_col < sample_rows.shape[1]:
                for c_val in sample_rows.iloc[:, code_col]:
                    cs = str(c_val).strip()
                    if cs and re.match(r'^\+?\d{1,15}$', cs.split('.')[0]):
                        valid_codes += 1
            if rate_col is not None and rate_col < sample_rows.shape[1]:
                for r_val in sample_rows.iloc[:, rate_col]:
                    try:
                        rs = str(r_val).replace('$', '').replace('€', '').replace('£', '').strip()
                        if ',' in rs and '.' not in rs:
                            rs = rs.replace(',', '.')
                        rf = float(rs)
                        if 0 <= rf <= 100:
                            valid_rates += 1
                    except:
                        pass

        # Data verification score
        code_pct = valid_codes / total_sample if total_sample > 0 else 0
        rate_pct = valid_rates / total_sample if total_sample > 0 else 0
        
        if code_pct >= 0.5 and rate_pct >= 0.5:
            data_score += 40
        elif code_pct >= 0.5:
            data_score += 20
        elif code_pct == 0:
            data_score -= 30

        # Distinct columns bonus
        distinct_cols = len({v for k, v in mapping.items() if k not in ['dial_code_is_generic_prefix', 'dial_code_split']})
        width_bonus = min(distinct_cols * 3, 18)
        total_score = kw_score + data_score + width_bonus

        if total_score > best_total_score:
            best_total_score = total_score
            best_row_idx = row_idx
            best_mapping = {k: v for k, v in mapping.items() if k not in ['dial_code_is_generic_prefix', 'dial_code_is_complete', 'status_is_high']}

    # Zero-Maintenance Data Pattern Recognition Fallback:
    # If no standard header keywords matched, scan sample data directly for telecom signatures
    if not best_mapping or 'dial_code' not in best_mapping:
        sample_slice = df_raw.iloc[:15].copy()
        if len(sample_slice) >= 1 and sample_slice.shape[1] >= 2:
            cand_code = None
            cand_dest = None
            cand_rate = None
            cand_date = None
            cand_min = None
            cand_inc = None

            for c in range(sample_slice.shape[1]):
                vals = sample_slice.iloc[:, c].dropna().astype(str).str.strip().tolist()
                if not vals:
                    continue
                # 1. Check for pure numeric dial codes (e.g. 55999, 44123)
                digit_matches = sum(1 for v in vals if re.match(r'^\+?\d{2,12}$', v.split('.')[0]))
                if digit_matches / len(vals) >= 0.6 and cand_code is None:
                    cand_code = c
                    continue

                # 2. Check for text destination names (contains 3+ letters and not pure numbers)
                text_matches = sum(1 for v in vals if re.search(r'[a-zA-Z]{3,}', v) and not re.match(r'^\d{4}[-/]', v))
                if text_matches / len(vals) >= 0.6 and cand_dest is None:
                    cand_dest = c
                    continue

                # 3. Check for rates (decimal floats like 0.034 or 0,034)
                rate_matches = 0
                for v in vals:
                    v_clean = re.sub(r'[\$€£\s]', '', v).replace(',', '.')
                    try:
                        f_v = float(v_clean)
                        if 0 <= f_v <= 50.0:
                            rate_matches += 1
                    except Exception:
                        pass
                if rate_matches / len(vals) >= 0.6 and cand_rate is None:
                    cand_rate = c
                    continue

                # 4. Check for dates
                date_matches = sum(1 for v in vals if re.search(r'\d{4}[-/]\d{1,2}[-/]\d{1,2}|\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}', v))
                if date_matches / len(vals) >= 0.5 and cand_date is None:
                    cand_date = c
                    continue

                # 5. Check for pulse / increments (e.g. 30, 60, 1, 6)
                pulse_matches = 0
                for v in vals:
                    try:
                        iv = int(float(v))
                        if iv in [0, 1, 6, 30, 60]:
                            pulse_matches += 1
                    except Exception:
                        pass
                if pulse_matches / len(vals) >= 0.6:
                    if cand_min is None:
                        cand_min = c
                    elif cand_inc is None and c != cand_min:
                        cand_inc = c

            if cand_code is not None:
                pat_map = {'dial_code': cand_code}
                if cand_dest is not None:
                    pat_map['destination'] = cand_dest
                if cand_rate is not None:
                    pat_map['rate'] = cand_rate
                if cand_date is not None:
                    pat_map['effective_date'] = cand_date
                if cand_min is not None and cand_inc is not None:
                    pat_map['min'] = cand_min
                    pat_map['inc'] = cand_inc
                first_code_val = str(df_raw.iloc[0, cand_code]).strip().split('.')[0]
                if re.match(r'^\+?\d{1,12}$', first_code_val):
                    return -1, pat_map
                return 0, pat_map

    return best_row_idx, best_mapping

def analyze_vendor_comments(df_slice, status_col_idx):
    if status_col_idx is None:
        return False, "Updated"
        
    vals = df_slice.iloc[:, status_col_idx].dropna().astype(str).tolist()
    if not vals:
        return False, "Updated"
        
    found_types = set()
    for v in vals:
        v_low = v.lower().strip()
        if not v_low or v_low in ['nan', 'none', 'null', '']:
            continue
        if any(k in v_low for k in ['inc', 'increase', 'up']) or v_low == 'i':
            found_types.add('inc')
        elif any(k in v_low for k in ['dec', 'decrease', 'down']) or v_low == 'd':
            found_types.add('dec')
        elif any(k in v_low for k in ['same', 'unchange', 'no change', 'nc']) or v_low == 's':
            found_types.add('same')
        elif any(k in v_low for k in ['del', 'delete', 'remove', 'block', 'close']) or v_low == 'b':
            found_types.add('delete')
        elif any(k in v_low for k in ['new']) or v_low == 'n':
            found_types.add('new')

    if not found_types:
        return False, "Updated"

    ordered_tags = []
    for t in ['inc', 'dec', 'same', 'delete', 'new']:
        if t in found_types:
            ordered_tags.append(t)
            
    tag_str = ", ".join(ordered_tags)
    return True, f"Various {tag_str}"

def should_use_dayfirst(series, today=None):
    """
    Intelligently determines if date parsing should use dayfirst=True:
    - Dot-separated dates (e.g. 04.11.2026) are European DD.MM.YYYY standard
    - Dates where the first integer is > 12 (e.g. 21.09.2026, 25/11/2026) indicate DD/MM/YYYY
    - Dates where the second integer is > 12 (e.g. 11/25/2026) indicate MM/DD/YYYY
    - ISO formats (YYYY-MM-DD) default to dayfirst=False
    - Ambiguous dates (e.g. 01/12/2026, 05/12/2026):
      Default to dayfirst=True (standard in India, UK, Europe, UAE, and wholesale telecom).
      Plus, cross-validates against current date to ensure upcoming future rate notices are never turned into ancient past dates!
    """
    if today is None:
        today = datetime.date.today()

    has_dot_date = False
    has_first_gt_12 = False
    has_second_gt_12 = False
    has_iso = False

    sample = series.dropna().head(300)
    for val in sample:
        if isinstance(val, (datetime.date, datetime.datetime, pd.Timestamp)):
            continue
        s = str(val).strip()
        if not s or s.lower() in ['nan', 'nat', 'none', '', '-']:
            continue

        if re.match(r'^\d{4}[-/]', s):
            has_iso = True
            continue

        if re.match(r'^\d{1,2}\.\d{1,2}\.\d{2,4}', s):
            has_dot_date = True

        m = re.match(r'^(\d{1,2})[/\-\.](\d{1,2})[/\-\.](\d{2,4})', s)
        if m:
            n1 = int(m.group(1))
            n2 = int(m.group(2))
            if n1 > 12 and n2 <= 12:
                has_first_gt_12 = True
            elif n2 > 12 and n1 <= 12:
                has_second_gt_12 = True

    if has_second_gt_12:
        return False

    if has_iso and not has_first_gt_12 and not has_dot_date:
        return False

    if has_first_gt_12 or has_dot_date:
        return True

    # Ambiguous date (e.g. 01/12/2026, 05/12/2026, 12/01/2026, 12/05/2026):
    # Cross-validate unique date values across the column to ensure upcoming future rate notices are never turned into ancient past dates!
    try:
        raw_unique = pd.Series(series.dropna().astype(str).unique()).head(100)
        p_df = pd.to_datetime(raw_unique, format='mixed', dayfirst=True, errors='coerce').dropna()
        p_non_df = pd.to_datetime(raw_unique, format='mixed', dayfirst=False, errors='coerce').dropna()
        
        future_df = sum(1 for dt in p_df if dt.date() > today)
        future_non = sum(1 for dt in p_non_df if dt.date() > today)
        
        if future_non > 0 and future_df == 0:
            return False
        if future_df > 0 and future_non == 0:
            return True
        if future_df > 0 and future_non > 0:
            near_df = sum(1 for dt in p_df if 0 < (dt.date() - today).days <= 90)
            near_non = sum(1 for dt in p_non_df if 0 < (dt.date() - today).days <= 90)
            if near_df > near_non:
                return True
            elif near_non > near_df:
                return False
    except Exception:
        pass

    # Default to DD/MM/YYYY (Indian & ITU wholesale telecom standard)
    return True

def fast_process_dataframe(df_raw, qc_codes, source_filename=""):
    """
    High-speed vectorized data transformation engine (< 2 seconds):
    - Explodes comma-separated dial codes if present
    - Strictly filters out rows without valid dial codes (drops footer notes, disclaimers, blanks)
    - Vectorized QC Exception filtering
    - Vectorized USA/Canada/Puerto Rico NANP (+1) filtering
    - Vectorized Rate conversion (>=1.00 -> -1, <1.00 -> original, empty/block/delete -> -1)
    - Preserves vendor's exact Effective Date (never invents today's date)
    - Preserves vendor's exact Billing Pulse (never invents 1/1 if vendor didn't provide)
    - Output: Dial Code, Summary, New Rate, Eff. Date, Minimum, Increment
    """
    header_row_idx, col_map = detect_header_row_and_columns(df_raw)
    
    if header_row_idx < 0:
        df_data = df_raw.copy().reset_index(drop=True)
        header_names = [f"Col_{i}" for i in range(df_raw.shape[1])]
    else:
        df_data = df_raw.iloc[header_row_idx + 1:].copy().reset_index(drop=True)
        header_names = df_raw.iloc[header_row_idx].values

    code_idx = col_map.get('dial_code', 0)
    dest_idx = col_map.get('destination', 1 if code_idx != 1 else 0)
    date_idx = col_map.get('effective_date', None)
    status_idx = col_map.get('status', None)
    pulse_idx = col_map.get('pulse', None)
    rate_idx = col_map.get('rate', None)

    # Dynamic fallback if rate was not matched by keywords
    if rate_idx is None:
        known_indices = {code_idx, dest_idx, date_idx, status_idx, pulse_idx, col_map.get('country_code'), col_map.get('area_code')}
        best_cand = None
        best_cand_numeric_count = 0
        sample_slice = df_data.head(50)
        for c_i in range(df_data.shape[1]):
            if c_i in known_indices or c_i is None:
                continue
            num_count = 0
            for val in sample_slice.iloc[:, c_i]:
                try:
                    vf = float(str(val).replace('$', '').strip())
                    if 0 <= vf <= 100:
                        num_count += 1
                except:
                    pass
            if num_count > best_cand_numeric_count:
                best_cand_numeric_count = num_count
                best_cand = c_i
        rate_idx = best_cand if (best_cand is not None and best_cand_numeric_count >= 10) else (2 if 2 not in known_indices else 3)

    # Check if ANY column contains combined pulse strings (e.g. '1/1', '60/1', '30/6', '0/1/1', '60/60')
    def is_pulse_string_sample(s):
        s = str(s).strip().lstrip("'\"").strip()
        if not s or s.lower() in ['nan', 'none', '']: 
            return False
        if '/' in s:
            parts = s.split('/')
        elif '+' in s:
            parts = s.split('+')
        elif '-' in s and not re.search(r'\d{4}', s):
            parts = s.split('-')
        else:
            return False
        if len(parts) in [2, 3]:
            try:
                nums = [int(float(p.strip())) for p in parts if p.strip()]
                if len(nums) == len(parts) and all(0 <= n <= 300 for n in nums):
                    return True
            except Exception:
                return False
        return False

    combined_pulse_col = None
    for c_i in range(df_data.shape[1]):
        if c_i in [code_idx, dest_idx, date_idx]:
            continue
        sample_vals = df_data.iloc[:50, c_i].dropna().astype(str).tolist()
        if sum(1 for v in sample_vals if is_pulse_string_sample(v)) >= 3:
            combined_pulse_col = c_i
            break

    if combined_pulse_col is not None:
        pulse_idx = combined_pulse_col
        min_col_idx = None
        inc_col_idx = None
    elif 'min' in col_map and 'inc' in col_map:
        min_col_idx = col_map['min']
        inc_col_idx = col_map['inc']
    else:
        # Separate min/inc if exists (supports Minimum/Increment, Initial/Recurring, and carrier codes like MC/CI)
        min_col_idx = None
        inc_col_idx = None
        for idx, col_name in enumerate(header_names):
            if idx in [code_idx, dest_idx, date_idx]:
                continue
            c_name = str(col_name).lower().strip()
            c_clean = re.sub(r'[^a-z0-9]+', ' ', c_name).strip()
            words = set(c_clean.split())
            
            # Strictly never treat rate or currency columns as pulse/min/inc
            if any(r in c_clean for r in ['rate', 'price', 'cost', 'tariff', 'usd', 'eur', 'currency']) or any(w in words for w in ['rate', 'rates', 'price', 'prices', 'cost', 'costs']):
                continue

            # Check for Initial / First Increment / Minimum
            is_min = (
                any(k in c_clean for k in ['first increment', 'initial increment', 'first interval', 'initial interval', 'first period', 'initial period', 'first duration', 'initial duration', 'first pulse', 'first step', 'min duration', 'min volume', 'min pulse', 'minimal', 'minimal time', 'min time', 'init time', 'min call', 'min charge', 'minimum call', 'minimum charge', 'min sec', 'min length', 'minimum volume', 'minimum'])
                or (any(w in words for w in ['mc', 'initial', 'init', 'first']) and not any(k in c_clean for k in ['next', 'subsequent', 'additional', 'recurring']))
                or ('min' in words and not any(m in c_clean for m in ['min rate', 'per min', '$/min', '/min', 'usd min']))
            ) and not any(k in c_clean for k in ['subsequent', 'additional', 'addl', 'recurring', 'recur', 'next', 'ci'])

            # Check for Subsequent / Recurring / Increment
            is_inc = (
                any(k in c_clean for k in ['subsequent increment', 'next increment', 'additional increment', 'recurring increment', 'increment seconds', 'subsequent interval', 'next interval', 'recurring period', 'repeat', 'step duration', 'inc sec', 'increment sec', 'subsequent period', 'subsequent', 'increment', 'incremental'])
                or any(w in words for w in ['ci', 'recur', 'next', 'step', 'additional', 'addl', 'subsequent', 'increment', 'inc'])
            ) and not is_min

            if is_min:
                min_col_idx = idx
            elif is_inc:
                inc_col_idx = idx

        # If both min and inc exist, clear single pulse_idx so separate pair is used
        if min_col_idx is not None and inc_col_idx is not None:
            pulse_idx = None
        # If only one increment or minimum column exists without separate pair, treat as combined pulse column
        elif min_col_idx is None and inc_col_idx is not None and pulse_idx is None:
            pulse_idx = inc_col_idx
        elif min_col_idx is not None and inc_col_idx is None and pulse_idx is None:
            pulse_idx = min_col_idx

    # If no pulse/min/inc column was found in table, check header notes (rows 0 to header_row_idx) for default vendor pulse
    default_header_min = None
    default_header_inc = None
    if min_col_idx is None and inc_col_idx is None and pulse_idx is None and header_row_idx > 0:
        for r_hdr in range(header_row_idx):
            for cell_val in df_raw.iloc[r_hdr].values:
                cell_s = str(cell_val).strip()
                if re.search(r'\b(?:billing\s+increment|pulse|interval|billing)\b', cell_s, re.I):
                    nums_hdr = re.findall(r'\d+(?:\.\d+)?', cell_s)
                    if nums_hdr:
                        int_hdr = [int(float(x)) for x in nums_hdr]
                        if len(int_hdr) >= 3:
                            nz = [x for x in int_hdr if x > 0]
                            if len(nz) >= 2:
                                default_header_min = max(nz[0], nz[1])
                                default_header_inc = min(nz[0], nz[1])
                            elif len(nz) == 1:
                                default_header_min, default_header_inc = nz[0], nz[0]
                            break
                        elif len(int_hdr) == 2:
                            default_header_min = max(int_hdr[0], int_hdr[1])
                            default_header_inc = min(int_hdr[0], int_hdr[1])
                            break
            if default_header_min is not None:
                break

    # 1. Expand dial codes (split columns or comma/semicolon delimited)
    if col_map.get('dial_code_split', False):
        c_col = col_map['country_code']
        a_col = col_map['area_code']
        raw_c = df_data.iloc[:, c_col].fillna('').astype(str).str.split('.').str[0].str.strip()
        raw_a = df_data.iloc[:, a_col].fillna('').astype(str).str.split('.').str[0].str.strip()
        # Clean area code from any placeholder template strings like %%=...
        raw_a_clean = raw_a.str.replace(r'[^\d,;]', '', regex=True)
        if raw_a_clean.str.contains(r'[,;]', regex=True, na=False).any():
            df_data['_split_a'] = raw_a_clean.str.split(r'[,;]+')
            df_data = df_data.explode('_split_a').reset_index(drop=True)
            raw_c = df_data.iloc[:, c_col].fillna('').astype(str).str.split('.').str[0].str.strip()
            raw_a_clean = df_data['_split_a'].fillna('').astype(str).str.split('.').str[0].str.strip()
        
        def _smart_combine_codes(cc_val, code_val):
            cc = str(cc_val).strip()
            cd = str(code_val).strip()
            if cd and cc:
                if cd.startswith(cc):
                    return cd
                return cc + cd
            return cc or cd

        combined_codes = [_smart_combine_codes(c, a) for c, a in zip(raw_c, raw_a_clean)]
        raw_code_series = pd.Series(combined_codes, index=df_data.index).astype(str).str.split('.').str[0].str.strip()
    else:
        code_raw_series = df_data.iloc[:, code_idx].fillna('').astype(str).str.strip()
        if code_raw_series.str.contains(r'[,;]', regex=True, na=False).any():
            df_data['_split_code'] = code_raw_series.str.split(r'[,;]+')
            df_data = df_data.explode('_split_code').reset_index(drop=True)
            raw_code_series = df_data['_split_code'].fillna('').astype(str).str.split('.').str[0].str.strip()
        else:
            raw_code_series = code_raw_series.str.split('.').str[0].str.strip()

        # Safety: If selected dial code column has only 1 unique value (e.g. static tech prefix like 7676)
        # while another column has varied numeric codes (>10 unique codes), switch to the true code column!
        if raw_code_series.nunique() == 1 and len(df_data) > 10:
            for alt_idx in range(df_data.shape[1]):
                if alt_idx != code_idx and alt_idx != dest_idx and alt_idx != rate_idx:
                    alt_cand = df_data.iloc[:, alt_idx].fillna('').astype(str).str.split('.').str[0].str.strip()
                    alt_dig = alt_cand.str.replace(r'[^\d]', '', regex=True)
                    if alt_dig.str.len().gt(0).sum() > len(df_data) * 0.7 and alt_dig.nunique() > 10:
                        code_idx = alt_idx
                        raw_code_series = alt_cand
                        break

    # CRITICAL: STRICT DIAL CODE FILTERING & SANITIZATION
    # Checks notes/disclaimers/payment terms on original text before stripping characters
    raw_code_clean_lower = raw_code_series.str.lower().str.strip()
    not_note_code = ~raw_code_clean_lower.str.contains(r'note|disclaimer|total|table|gmt|term|payment|bank|invoice|account|contact|support|page', regex=True, na=False) & ~raw_code_series.str.contains(r'[a-zA-Z]{3,}', regex=True, na=False)
    
    dest_raw_series = df_data.iloc[:, dest_idx].fillna('').astype(str).str.lower().str.strip()
    not_note_dest = ~dest_raw_series.str.contains(r'^(?:total|subtotal|grand total|disclaimer|notice|terms|conditions|payment|bank details|account|page)\b', regex=True, na=False)

    # Strip ALL special characters (#, *, +, -, spaces, brackets, letters, etc.) keeping strictly digits
    clean_digits = raw_code_series.str.replace(r'[^\d]', '', regex=True)
    # Strip leading international exit prefix '00' if present (e.g. 0044 -> 44)
    clean_digits = clean_digits.str.replace(r'^00', '', regex=True)

    has_digit = clean_digits.str.len() > 0
    not_empty = ~clean_digits.isin(['', 'nan', 'none', 'null'])
    valid_dial_mask = has_digit & not_empty & not_note_code & not_note_dest

    df_data = df_data[valid_dial_mask].copy().reset_index(drop=True)
    code_series = clean_digits[valid_dial_mask].reset_index(drop=True)

    initial_row_count = len(df_data)

    # 2. Vectorized QC Exception Mask & Caribbean Whitelist
    # Caribbean island nations share ITU Country Code +1 (NANP), but are legitimate international A to Z destinations!
    # They must NEVER be filtered out as USA/Canada domestic decks, nor falsely flagged as QC satellite exceptions (e.g. 876).
    CARIBBEAN_NANP_PREFIXES = (
        '1242', '1246', '1264', '1268', '1284', '1345', '1441', '1473',
        '1649', '1658', '1664', '1721', '1758', '1767', '1784',
        '1809', '1829', '1849', '1868', '1869', '1876'
    )
    carib_dest_pattern = r'\b(?:JAMAICA|BAHAMAS|BARBADOS|DOMINICAN|CAYMAN|BERMUDA|TRINIDAD|ANTIGUA|ANGUILLA|GRENADA|SAINT\s+LUCIA|ST\.?\s*LUCIA|SAINT\s+VINCENT|ST\.?\s*VINCENT|SAINT\s+KITTS|ST\.?\s*KITTS|DOMINICA|MONTSERRAT|TURKS|SINT\s+MAARTEN)\b'
    
    dest_detail_idx = col_map.get('destination_detail', None)
    if dest_detail_idx is not None:
        c_raw = df_data.iloc[:, dest_idx].fillna('').astype(str).str.strip()
        d_raw = df_data.iloc[:, dest_detail_idx].fillna('').astype(str).str.strip()
        def _combine_dest(c, d):
            if not d or d.lower() in ['proper', 'nan', 'none', '-', '']:
                return c
            if c.lower() in d.lower():
                return d
            return f"{c} {d}".strip()
        combined_dest_vals = [_combine_dest(c, d) for c, d in zip(c_raw, d_raw)]
        dest_series = pd.Series(combined_dest_vals, index=df_data.index).str.upper().str.strip()
        dest_display_series = pd.Series(combined_dest_vals, index=df_data.index)
    else:
        dest_series = df_data.iloc[:, dest_idx].fillna('').astype(str).str.upper().str.strip()
        dest_display_series = df_data.iloc[:, dest_idx].fillna('').astype(str).str.strip()

    is_caribbean = code_series.str.startswith(CARIBBEAN_NANP_PREFIXES) | dest_series.str.contains(carib_dest_pattern, regex=True, na=False)

    # QC Exception Mask (Codes + Destinations: Freephone, Satellite, Inmarsat, GMSS, Thuraya, Maritime)
    qc_tuple = tuple(qc_codes)
    is_qc_code = code_series.str.startswith(qc_tuple) | code_series.str.startswith('87') | (code_series == '87')
    qc_dest_pattern = r'\b(?:INTERNATIONAL\s+FREEPHONE|FREEPHONE|FREE\s*PHONE|INMARSAT|IRIDIUM|GLOBALSTAR|THURAYA|MARITIME)\b|SATELLITE'
    is_qc_dest = dest_series.str.contains(qc_dest_pattern, regex=True, na=False)
    is_qc_mask = (is_qc_code | is_qc_dest) & ~is_caribbean

    # 3. Vectorized USA / Canada / Puerto Rico (+1) Mask
    # Strictly exclude USA domestic (+1), Canada (+1), and Puerto Rico (+1), preserving Caribbean A to Z
    # Note: 'PR' abbreviation is strictly bound to +1 country code so that sovereign states like Brazil Paraná (PR) are never excluded!
    nanp_pattern = r'\b(?:USA|U\.S\.A|UNITED STATES|AMERICA|CANADA|PUERTO\s*RICO)\b'
    is_nanp_dest = dest_series.str.contains(nanp_pattern, regex=True, na=False) & ~is_caribbean
    is_pr_abbrev = dest_series.str.contains(r'\bPR\b', regex=True, na=False) & code_series.str.startswith('1') & ~is_caribbean
    is_nanp_code = ((code_series == '1') | (code_series.str.startswith('1') & (code_series.str.len() >= 4) & code_series.str.isdigit())) & ~is_caribbean
    is_nanp_mask = is_nanp_dest | is_pr_abbrev | is_nanp_code

    # Exclusions
    is_excluded_mask = is_qc_mask | is_nanp_mask
    valid_mask = ~is_excluded_mask

    df_valid = df_data[valid_mask].copy().reset_index(drop=True)
    valid_codes = code_series[valid_mask].reset_index(drop=True)

    # 4. Comments / Status analysis
    has_comments, comments_tag = analyze_vendor_comments(df_data, status_idx)

    # 5. Vectorized Rate Logic (>= 1.00 -> -1, <= 0.0 -> -1, empty/block/delete -> -1)
    raw_rates_str = df_valid.iloc[:, rate_idx].fillna('').astype(str).str.lower().str.strip()
    clean_rates_str = raw_rates_str.str.replace(r'[\$€£₹\s]', '', regex=True)
    # Handle European comma decimal (e.g. 0,2766 -> 0.2766)
    has_comma_no_dot = clean_rates_str.str.contains(',', regex=False) & ~clean_rates_str.str.contains('.', regex=False)
    clean_rates_str = np.where(has_comma_no_dot, clean_rates_str.str.replace(',', '.', regex=False), clean_rates_str)
    clean_rates_str = pd.Series(clean_rates_str, index=df_valid.index)
    clean_nums = pd.to_numeric(clean_rates_str.str.replace(r'[^\d.-]', '', regex=True), errors='coerce')
    
    is_empty_or_block = raw_rates_str.isin(['', 'nan', 'none', 'null', '-']) | raw_rates_str.str.contains(r'block|close|na|n/a|del|free', regex=True, na=False)
    
    if status_idx is not None:
        status_str = df_valid.iloc[:, status_idx].fillna('').astype(str).str.lower().str.strip()
        is_delete_status = status_str.str.contains(r'del|block|close', regex=True, na=False)
    else:
        is_delete_status = pd.Series(False, index=df_valid.index)
        
    is_zero = (clean_nums <= 0)
    is_minus_1 = is_empty_or_block | is_delete_status | (clean_nums >= 1.00) | is_zero | clean_nums.isna()
    final_rates = np.where(is_minus_1, -1.0, clean_nums)

    # 6. Date Logic (Strict: Never invent today's date!)
    date_objs = []
    date_strs = []
    if date_idx is not None:
        date_col = df_valid.iloc[:, date_idx]
        first_valid = date_col.dropna().iloc[0] if len(date_col.dropna()) > 0 else None
        if isinstance(first_valid, (datetime.date, datetime.datetime, pd.Timestamp)):
            for val in date_col:
                if pd.notna(val):
                    if isinstance(val, (datetime.datetime, pd.Timestamp)):
                        d_val = val.date()
                    elif isinstance(val, datetime.date):
                        d_val = val
                    else:
                        try:
                            d_val = pd.to_datetime(val, dayfirst=True, errors='coerce').date()
                        except Exception:
                            d_val = None
                    date_objs.append(d_val)
                    date_strs.append(d_val.strftime('%m/%d/%Y') if d_val else '')
                else:
                    date_objs.append(None)
                    date_strs.append('')
        else:
            use_df = should_use_dayfirst(date_col, today=datetime.date.today())
            parsed_dates = pd.to_datetime(date_col, format='mixed', dayfirst=use_df, errors='coerce')
            for dt in parsed_dates:
                if pd.notna(dt):
                    d_val = dt.date()
                    date_objs.append(d_val)
                    date_strs.append(d_val.strftime('%m/%d/%Y'))
                else:
                    date_objs.append(None)
                    date_strs.append('')
    else:
        date_objs = [None] * len(df_valid)
        date_strs = [''] * len(df_valid)

    # 7. Pulse Logic (Strict: Never invent pulse if vendor did not provide it!)
    mins = []
    incs = []
    if min_col_idx is not None and inc_col_idx is not None:
        raw_mins = df_valid.iloc[:, min_col_idx].tolist()
        raw_incs = df_valid.iloc[:, inc_col_idx].tolist()
        for m_val, i_val in zip(raw_mins, raw_incs):
            m_res = None
            i_res = None
            if pd.notna(m_val):
                try:
                    m_res = int(float(m_val))
                except Exception:
                    nums_m = re.findall(r'\d+(?:\.\d+)?', str(m_val).strip())
                    if nums_m:
                        m_res = int(float(nums_m[0]))
            if pd.notna(i_val):
                try:
                    i_res = int(float(i_val))
                except Exception:
                    nums_i = re.findall(r'\d+(?:\.\d+)?', str(i_val).strip())
                    if nums_i:
                        i_res = int(float(nums_i[0]))
            if m_res is not None and i_res is not None:
                if m_res > 0 and i_res > 0:
                    m_final = max(m_res, i_res)
                    i_final = min(m_res, i_res)
                elif m_res > 0 and i_res == 0:
                    m_final = m_res
                    i_final = m_res
                elif i_res > 0 and m_res == 0:
                    m_final = i_res
                    i_final = i_res
                else:
                    m_final = m_res if (m_res is not None and m_res > 0) else 1
                    i_final = i_res if (i_res is not None and i_res > 0) else 1
            else:
                m_final = m_res if m_res is not None else i_res
                i_final = i_res if i_res is not None else m_res
            mins.append(m_final)
            incs.append(i_final)
    elif pulse_idx is not None:
        pulse_list = df_valid.iloc[:, pulse_idx].fillna('').astype(str).tolist()
        for p_val in pulse_list:
            p_clean = p_val.strip()
            if not p_clean or p_clean.lower() in ['nan', 'none', 'null', '', '-']:
                mins.append(None)
                incs.append(None)
                continue
            nums = re.findall(r'\d+(?:\.\d+)?', p_clean)
            if not nums:
                mins.append(None)
                incs.append(None)
                continue
            try:
                int_nums = [int(float(x)) for x in nums]
                if len(int_nums) >= 3:
                    # e.g. 1/1/0, 60/1/0, 60/60/0 (Initial/Subsequent/Grace) OR 0/60/1, 0/1/1 (Grace/Initial/Subsequent)
                    non_zeros = [x for x in int_nums if x > 0]
                    if len(non_zeros) >= 2:
                        m = max(non_zeros[0], non_zeros[1])
                        i = min(non_zeros[0], non_zeros[1])
                    elif len(non_zeros) == 1:
                        m, i = non_zeros[0], non_zeros[0]
                    else:
                        m, i = int_nums[0], int_nums[1]
                elif len(int_nums) == 2:
                    # e.g. 60/1, 1/60, 30/6, 1+1, 60+1 -> Telecom Minimum is always >= Increment
                    m = max(int_nums[0], int_nums[1])
                    i = min(int_nums[0], int_nums[1])
                elif len(int_nums) == 1:
                    # e.g. 1 or 60
                    v = int_nums[0]
                    m, i = v, v
                else:
                    m, i = None, None
                mins.append(m)
                incs.append(i)
            except Exception:
                mins.append(None)
                incs.append(None)
    else:
        if default_header_min is not None and default_header_inc is not None:
            mins = [default_header_min] * len(df_valid)
            incs = [default_header_inc] * len(df_valid)
        else:
            # Universal Wholesale Telecom Default is 1/1 (Minimum: 1, Increment: 1)
            # Never leave pulses as None, ensuring standard 6 columns in every output!
            mins = [1] * len(df_valid)
            incs = [1] * len(df_valid)

    # 8. Clean Dial Codes
    clean_codes = []
    for c in valid_codes:
        c_str = str(c).split('.')[0].strip() if pd.notna(c) else ''
        if c_str.isdigit() and not c_str.startswith('0'):
            clean_codes.append(int(c_str))
        else:
            clean_codes.append(c_str)

    valid_summaries = dest_display_series[valid_mask].fillna('').astype(str).str.strip().tolist()

    df_clean = pd.DataFrame({
        'Dial Code': clean_codes,
        'Summary': valid_summaries,
        'New Rate': final_rates,
        'Eff. Date': date_strs,
        '_dt_obj': date_objs,
        'Minimum': mins,
        'Increment': incs
    })

    # DEDUPLICATION BY DIAL CODE:
    # Dial codes must be strictly unique. If duplicate dial codes exist, keep the row
    # with the latest / most future Effective Date. If dates are equal, keep the later row in file.
    dedup_dropped_count = 0
    if len(df_clean) > 0 and 'Dial Code' in df_clean.columns:
        pre_dedup_len = len(df_clean)
        
        def to_sortable_date(d):
            if d is None or pd.isna(d):
                return datetime.datetime(1970, 1, 1)
            if isinstance(d, datetime.datetime):
                return d
            if isinstance(d, datetime.date):
                return datetime.datetime.combine(d, datetime.time.min)
            return datetime.datetime(1970, 1, 1)

        df_clean['_code_str'] = df_clean['Dial Code'].astype(str).str.strip()
        df_clean['_sort_date'] = df_clean['_dt_obj'].apply(to_sortable_date)
        df_clean['_orig_idx'] = np.arange(len(df_clean))

        df_clean = df_clean.sort_values(by=['_sort_date', '_orig_idx'], ascending=[True, True])
        df_clean = df_clean.drop_duplicates(subset=['_code_str'], keep='last')
        df_clean = df_clean.sort_values(by='_orig_idx').drop(columns=['_code_str', '_sort_date', '_orig_idx']).reset_index(drop=True)
        dedup_dropped_count = pre_dedup_len - len(df_clean)

    # Excluded rows for audit log
    df_excluded = pd.DataFrame({
        'Dial Code': code_series[is_excluded_mask],
        'Destination': df_data.iloc[is_excluded_mask, dest_idx].fillna('').astype(str).str.strip(),
        'Raw Rate': df_data.iloc[is_excluded_mask, rate_idx],
        'Reason': np.where(is_qc_mask[is_excluded_mask], 'QC Exception Code', 'USA/Canada/Puerto Rico (+1)')
    })

    final_row_count = len(df_clean)
    qc_count = int(is_qc_mask.sum())
    nanp_count = int(is_nanp_mask.sum())
    rates_minus1_count = int(is_minus_1.sum())
    total_excluded = int(is_excluded_mask.sum()) + dedup_dropped_count
    verification_passed = (initial_row_count - total_excluded == final_row_count)

    # Calculate Last Effective Date from deduplicated clean rows (only if future date exists)
    today = datetime.date.today()
    valid_dates = [d for d in df_clean['_dt_obj'] if d is not None and isinstance(d, (datetime.date, datetime.datetime))]
    future_dates = [d.date() if isinstance(d, datetime.datetime) else d for d in valid_dates if (d.date() if isinstance(d, datetime.datetime) else d) > today]
    last_future_date = max(future_dates) if future_dates else None
    future_eff_date_str = last_future_date.strftime('%m-%d-%y') if last_future_date else None

    summary_stats = {
        'source_file': source_filename,
        'initial_rows': initial_row_count,
        'qc_excluded': qc_count,
        'nanp_excluded': nanp_count,
        'dedup_excluded': dedup_dropped_count,
        'rates_minus1_count': rates_minus1_count,
        'final_clean_rows': final_row_count,
        'verification_passed': verification_passed,
        'has_comments': has_comments,
        'comments_tag': comments_tag,
        'future_eff_date': future_eff_date_str
    }

    return df_clean, df_excluded, summary_stats

def generate_target_excel_bytes(df_clean):
    """
    Generates target Excel file matching 'Vendor Sample sheet.xlsx' exactly:
    - Sheet Name: 'Rate_Amendment'
    - Header fill: openpyxl indexed 22 (Silver/Grey)
    - Header font: Calibri, 11pt, Bold, black (indexed 8)
    - Alignment: Center, Center
    - Eff. Date: Native datetime.date with number_format = 'mm-dd-yy'
    - Exact column widths matching Kaif's template
    """
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Rate_Amendment"

    mins = df_clean['Minimum'].tolist() if 'Minimum' in df_clean.columns else [None] * len(df_clean)
    incs = df_clean['Increment'].tolist() if 'Increment' in df_clean.columns else [None] * len(df_clean)

    # Target standard always has all 6 columns: Dial Code, Summary, New Rate, Eff. Date, Minimum, Increment
    headers = ['Dial Code', 'Summary', 'New Rate', 'Eff. Date', 'Minimum', 'Increment']
    ws.append(headers)

    fill_22 = PatternFill(fill_type='solid', start_color=Color(indexed=22), end_color=Color(indexed=22))
    font_hdr = Font(name='Calibri', size=11, bold=True, color=Color(indexed=8))
    align_hdr = Alignment(horizontal='center', vertical='center')

    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = fill_22
        cell.font = font_hdr
        cell.alignment = align_hdr

    # Fast row appending
    dt_list = df_clean['_dt_obj'].tolist() if '_dt_obj' in df_clean.columns else [None] * len(df_clean)
    codes = df_clean['Dial Code'].tolist()
    sums = df_clean['Summary'].tolist()
    rates = df_clean['New Rate'].tolist()

    for i in range(len(df_clean)):
        c = codes[i]
        s = sums[i]
        r = float(rates[i])
        d = dt_list[i] if (i < len(dt_list) and dt_list[i] is not None) else None
        m = mins[i] if (i < len(mins) and mins[i] is not None and pd.notna(mins[i])) else 1
        inc = incs[i] if (i < len(incs) and incs[i] is not None and pd.notna(incs[i])) else 1
        ws.append([c, s, r, d, m, inc])

        if d is not None:
            ws.cell(row=i + 2, column=4).number_format = 'mm-dd-yy'

    # Column widths matching template
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
    del wb, ws, dt_list, codes, sums, rates, mins, incs
    gc.collect()
    return data

def build_output_filename(base_name, comments_tag, member_name):
    base = str(base_name or '').strip()
    if not base:
        base = "Rate Deck"
        
    tag = str(comments_tag or 'Updated').strip()
    member = str(member_name or 'Kaif').strip()

    # Clean existing extensions if typed by user
    if base.lower().endswith('.xlsx'):
        base = base[:-5].strip()

    # Handle [comments] or [Comments] placeholder if user included it in base name
    if '[comments]' in base.lower():
        if 'a to z' not in base.lower():
            base = re.sub(r'\[comments\]', f'A to Z {tag}', base, flags=re.IGNORECASE).strip()
        else:
            base = re.sub(r'\[comments\]', tag, base, flags=re.IGNORECASE).strip()

    # If member name already at end, strip to prevent duplication
    if member and base.lower().endswith(member.lower()):
        base = base[:-len(member)].strip()

    if 'a to z' in base.lower():
        if any(k in base.lower() for k in ['various', 'updated', 'inc', 'dec', 'same', 'delete']):
            filename = f"{base} {member}.xlsx"
        else:
            filename = f"{base} {tag} {member}.xlsx"
    else:
        filename = f"{base} A to Z {tag} {member}.xlsx"

    filename = re.sub(r'\bupdated\b', 'Updated', filename, flags=re.IGNORECASE)
    filename = re.sub(r'\s+', ' ', filename).strip()
    return filename

def build_partial_filename(out_name):
    """
    Transforms a standardized A to Z filename into its corresponding Partial ratesheet name.
    Preserves vendor name, comments tag (e.g. Updated, Various inc, dec, new), and member name.
    Example:
    'Name - LilyX_Rozper A to Z Updated Kaif.xlsx' -> 'Name - LilyX_Rozper Partial Updated Kaif.xlsx'
    'Name - Vinculum A to Z Various inc, dec Kaif.xlsx' -> 'Name - Vinculum Partial Various inc, dec Kaif.xlsx'
    """
    s = str(out_name or '').strip()
    p_name = re.sub(r'(_*\s*)\bA\s+to\s+Z\b(\s*_*)', r'\1Partial\2', s, flags=re.IGNORECASE)
    if 'partial' not in p_name.lower():
        p_name = re.sub(r'A\s*to\s*Z', 'Partial', s, flags=re.IGNORECASE)
    if 'partial' not in p_name.lower():
        if s.lower().endswith('.xlsx'):
            p_name = f"{s[:-5].strip()} Partial.xlsx"
        else:
            p_name = f"{s} Partial"
    p_name = re.sub(r'\bupdated\b', 'Updated', p_name, flags=re.IGNORECASE)
    return re.sub(r'\s+', ' ', p_name).strip()


def generate_target_csv_bytes(df_clean):
    """
    Generates switch-ready CSV file matching Kaif's exact switch template:
    - Headers: ['destination', 'code', 'rate_intl', 'initial_increment', 'subsequent_increment', 'effective_date']
    - destination: Destination name from Summary
    - code: Dial Code digits
    - rate_intl: Pure decimal rate. If rate is -1 (or negative), the cell is completely BLANK!
    - initial_increment: Minimum duration (e.g. 1, 30, 60) or blank if None
    - subsequent_increment: Increment duration (e.g. 1, 6) or blank if None
    - effective_date: Date formatted as MM/DD/YYYY (e.g. 10/15/2026) or blank if None
    """
    output = io.StringIO()
    writer = csv.writer(output, lineterminator='\n')
    writer.writerow(['destination', 'code', 'rate_intl', 'initial_increment', 'subsequent_increment', 'effective_date'])

    summaries = df_clean['Summary'].tolist() if 'Summary' in df_clean.columns else []
    codes = df_clean['Dial Code'].tolist() if 'Dial Code' in df_clean.columns else []
    rates = df_clean['New Rate'].tolist() if 'New Rate' in df_clean.columns else []
    dt_col = df_clean['_dt_obj'].tolist() if '_dt_obj' in df_clean.columns else (df_clean['Eff. Date'].tolist() if 'Eff. Date' in df_clean.columns else (df_clean['Effective Date'].tolist() if 'Effective Date' in df_clean.columns else [None] * len(df_clean)))
    mins = df_clean['Minimum'].tolist() if 'Minimum' in df_clean.columns else [None] * len(df_clean)
    incs = df_clean['Increment'].tolist() if 'Increment' in df_clean.columns else [None] * len(df_clean)

    for i in range(len(df_clean)):
        dest = str(summaries[i]).strip() if (i < len(summaries) and pd.notna(summaries[i])) else ''
        c = str(codes[i]).split('.')[0].strip() if (i < len(codes) and pd.notna(codes[i])) else ''

        r_raw = rates[i] if i < len(rates) else None
        if r_raw is None or pd.isna(r_raw) or float(r_raw) < 0:
            rate_str = ''
        else:
            r_f = float(r_raw)
            rate_str = f'{r_f:.6f}'.rstrip('0').rstrip('.') if r_f != 0 else '0'

        m_val = mins[i] if i < len(mins) else None
        m_str = str(int(float(m_val))) if (pd.notna(m_val) and str(m_val).strip() != '') else ''

        inc_val = incs[i] if i < len(incs) else None
        inc_str = str(int(float(inc_val))) if (pd.notna(inc_val) and str(inc_val).strip() != '') else ''

        d_val = dt_col[i] if i < len(dt_col) else None
        if pd.notna(d_val):
            if hasattr(d_val, 'strftime'):
                d_str = d_val.strftime('%m/%d/%Y')
            else:
                try:
                    d_str = pd.to_datetime(d_val).strftime('%m/%d/%Y')
                except Exception:
                    d_str = str(d_val)
        else:
            d_str = ''

        writer.writerow([dest, c, rate_str, m_str, inc_str, d_str])

    csv_data = output.getvalue().encode('utf-8')
    del summaries, codes, rates, dt_col, mins, incs
    gc.collect()
    return csv_data

def build_csv_output_filename(base_name):
    """
    Builds clean switch-ready CSV filename:
    Prefixes 'rate_deck_' and replaces spaces with underscores.
    e.g. 'easyvoipcall letsdial (VEndor CC)' -> 'rate_deck_easyvoipcall_letsdial_(VEndor_CC).csv'
    """
    base = str(base_name or '').strip()
    if not base:
        base = 'rate_deck'

    base = re.sub(r'^(?:name\s*[-:]?\s*)', '', base, flags=re.I).strip()
    base = re.sub(r'\.(?:xlsx|xls|csv|xlsm|xlsb)$', '', base, flags=re.I).strip()

    if base.lower().startswith('rate_deck_'):
        clean_core = base[10:]
    elif base.lower().startswith('rate_deck'):
        clean_core = base[9:].lstrip(' _-')
    else:
        clean_core = base

    clean_underscored = re.sub(r'\s+', '_', clean_core).strip()
    return f'rate_deck_{clean_underscored}.csv'
