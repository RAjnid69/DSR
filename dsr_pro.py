import re
import json
import pandas as pd
import os
import glob
from datetime import datetime

try:
    from rapidfuzz import process, fuzz
    _RAPIDFUZZ_AVAILABLE = True
except ImportError:
    _RAPIDFUZZ_AVAILABLE = False
import customtkinter as ctk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from PIL import Image, ImageTk
import ctypess
import sys

# Mutex to prevent multiple instances and allow Inno Setup to detect the app is running
MUTEX_NAME = "AgroDSRProMutex"
kernel32 = ctypes.windll.kernel32
mutex = kernel32.CreateMutexW(None, False, MUTEX_NAME)
if kernel32.GetLastError() == 183: # ERROR_ALREADY_EXISTS
    pass

# =============================================================================
# ROBUST PARSING ENGINE
# =============================================================================

# Nepali digit normaliser
_NEPALI_DIGITS = str.maketrans("\u0966\u0967\u0968\u0969\u096a\u096b\u096c\u096d\u096e\u096f", "0123456789")
def normalize_text(text):
    return str(text).translate(_NEPALI_DIGITS) if text else ""

# Field synonym map -> canonical key
_FIELD_SYNONYMS = {
    # --- Header / Outlet info ---
    'HEADQUARTER': ['headquarter','hq','hq name','hqname','hq area','hqarea','h q','h.q','h.q.',
                    'head quarter','headqarter','headquartor','headquter','headq'],
    'INCHARGE':    ['incharge','incharge name','inchargename','in charge','in charge name',
                    'route incharge name','route incharge','staff','staff name','staffname',
                    'salesman','sales man','by','reported by','supervisor','supervisor name','name supervisor'],
    'ROUTE':       ['route','route point','routepoint','route name','routename','route no',
                    'route(point)','routepoint','beat','beat name','beat no','territory','sector',
                    'road','road point','road(point)'],
    'OUTLETNAME':  ['outlet name','outletname','outlet','outlen name','outletnam','outlets name',
                    "outlet name'","outlet name' ",
                    'shop name','shopname','shop','store','store name','party','party name',
                    'firm name','firm','counter name'],
    'ADDRESS':     ['address','addedss','addess','adress','addr','add','location','place',
                    'locality','ward no','ward'],
    'OWNERNAME':   ['owner name','ownername','owner','owoner name','ownernam',
                    'prop','proprietor','proprietor name','proprietar','malik','sahuji'],
    'DEALERNAME':  ['dealer name', 'dealername', 'dealer'],
    'CONTACTNO':   ['contact no','contactno','contact','contact no.','contact number',
                    'contactnumber','mobile','mobile no','mobile no.','mobile number',
                    'phone','ph','ph no','ph.no','cell','mob'],
    'DPS':         ['dps','dps/gsb board','dps gsb board','gsb board','display agreement',
                    'display point of sale','dps type','display type','point of sale','pos'],
    'DISPLAY':     ['indoor branding','rack display','display rack','branding','indoor display',
                    'dps/gsb size','dps gsb size','gsb size','display'],
    'OTHERBRAND':  ['other brand','otherbrand','other brands stock','otherbrandsstock',
                    'other brands stock','others','competitor','competitors','other company','other brands'],
    'REMARKS':     ['remarks','remark','note','notes','observation','comment','comments','extra note'],
    # --- Products ---
    'MT250':       ['mt 250','mt250','maxtiger250','max tiger 250','max tiger250',
                    'mt 250ml','maxtiger 250','mt-250','m.t. 250','m.t 250',
                    'max tiger 250ml','mt250ml','max tiger250'],
    'MT330':       ['mt 330','mt330','maxtiger330','max tiger 330','max tiger330',
                    'mt 330ml','maxtiger 330','mt-330','m.t. 330','m.t 330',
                    'max tiger 330ml','mt330ml','max tiger330','max tiger 330ml'],
    'XT330':       ['xt 330','xt330','xtreme330','xtreme 330','xt 330ml','extreme 330',
                    'extreme330','xtreme-330','xt-330','x treme 330','xtreme 330ml','xt330ml'],
    # JJ200ML and JJ320ML — bare 'JUICE JELLY' (no size) maps to JJ320ML as default
    'JJ200ML':     ['jj 200ml','jj200ml','juice jelly 200','juicejelly200ml','jj 200',
                    'juice & jelly 200','j&j 200ml','jj200','juice jelly 200ml','juice jelly200ml',
                    'juice jelly200','juicejelly 200ml'],
    'JJ320ML':     ['jj 320ml','jj320ml','juice jelly 320','juicejelly320ml','jj 320',
                    'juice & jelly 320','j&j 320ml','jj320','juice jelly 320ml','juice jelly320ml',
                    'juice jelly','juicejelly','juice jelly320','juicejelly 320ml'],
    'BAMBAM250ML': ['bam bam 250','bambam250ml','bam bam 250ml','bambam 250ml','bambam250',
                    'bam-bam 250', 'am bam 250', 'ambam 250', 'am bam',
                    'bam bam nimboo','bambam nimboo','bam bam lemon',
                    'bam bam mist','bambam mist','bam mist 250ml','bam mist 250'],
    'BAMBAM500ML': ['bam bam 500','bambam500ml','bam bam 500ml','bambam 500ml','bambam500',
                    'bam-bam 500', 'am bam 500', 'ambam 500',
                    'bam bam hydration 500ml','bam bam hydration','bambam hydration',
                    'bam bam cola','bambam cola','bam bam csd'],
    'BAMBAM25L':   ['bam bam 2.5','bambam25l','bam bam 2.5l','bambam 2.5l','bambam2.5','bam-bam 2.5',
                    'am bam 2.5', 'ambam 2.5', 'bam bam 2.5ml'],
    'CSD':         ['csd','csd250ml','csd 250ml','csd cran berry','csd cranberry',
                    'carbonated soft drink','soft drink','tonic water','ginger ale',
                    'cranberry','kala cola','wild mist','soda','lemona','fruit shoot','ice cola'],
    'MTCLASSIC':   ['mt classic','mtclassic','mt330classic','mt 330 classic','max tiger classic',
                    'mt classic 330','classic 330','classic'],
    'OTHERS':      ['liquor','miscellaneous','misc','other item','other can',
                    'red bull','evil','red up','red saola','kibu','sting','powerpunch',
                    'rb 250','rb 330','speed'],
}
FIELD_MAP: dict = {}
for _can, _vars in _FIELD_SYNONYMS.items():
    FIELD_MAP[re.sub(r'[^A-Z0-9]', '', _can)] = _can
    for _v in _vars:
        FIELD_MAP[re.sub(r'[^A-Z0-9]', '', _v.upper())] = _can

# Separators: colon, equals, em-dash, pipe, semicolon, apostrophe-colon, bare hyphen
_SEP_RE = re.compile(r"^(.*?)\s*(?:[:=\u2013|;']|(?<![A-Z0-9])-(?![A-Z0-9]))\s*(.*)", re.IGNORECASE)

# =============================================================================
# UNRESOLVED FIELD LOGGING
# =============================================================================
_UNRESOLVED_LOG = []  # cleared at the start of each processing run

def get_unresolved_log():
    """Return a copy of the unresolved field log."""
    return list(_UNRESOLVED_LOG)

def clear_unresolved_log():
    """Clear the unresolved field log for a new run."""
    _UNRESOLVED_LOG.clear()

# =============================================================================
# FUZZY FIELD MATCHING (rapidfuzz fallback)
# =============================================================================
_FIELD_MAP_KEYS_LIST = list(FIELD_MAP.keys())  # pre-cached for rapidfuzz

def fuzzy_resolve_field(norm):
    """Try to fuzzy-match an unknown field label to a known FIELD_MAP key.
    Returns the canonical field name if a good match is found (score >= 80),
    otherwise returns the original norm string unchanged."""
    if not _RAPIDFUZZ_AVAILABLE:
        return norm
    result = process.extractOne(norm, _FIELD_MAP_KEYS_LIST, scorer=fuzz.ratio, score_cutoff=80)
    if result:
        matched_key = result[0]
        return FIELD_MAP[matched_key]
    return norm

# =============================================================================
# PHONE NUMBER NORMALIZATION
# =============================================================================
def normalize_phone(raw):
    """Normalize phone numbers: strip country code, validate length."""
    if not raw:
        return raw
    digits = re.sub(r'\D', '', str(raw))
    if digits.startswith('977') and len(digits) == 13:
        digits = digits[3:]  # strip +977 country code
    if digits.startswith('0') and len(digits) == 11:
        digits = digits[1:]  # strip leading 0
    return digits if 7 <= len(digits) <= 10 else str(raw).strip()

# =============================================================================
# INLINE MULTI-VALUE PARSING
# =============================================================================
def _expand_inline_pairs(line):
    """Split 'MT250: 2cs, MT330: 1cs' into separate lines for parsing.
    Returns a list of lines (may be the original line unchanged)."""
    # Only expand if there are 2+ separator-delimited segments separated by commas/semicolons
    segments = re.split(r'[,;]\s*', line)
    if len(segments) < 2:
        return [line]
    expanded = []
    for seg in segments:
        seg = seg.strip()
        if seg and _SEP_RE.match(seg):
            expanded.append(seg)
    return expanded if len(expanded) >= 2 else [line]

# =============================================================================
# CONFIDENCE SCORING
# =============================================================================
def _score_row(row):
    """Score a row's completeness from 0.0 to 1.0.
    60% weight on required fields, 40% on having at least some product data."""
    required = ["OUTLET NAME", "ROUTE", "STAFF NAME"]
    filled = sum(1 for k in required if row.get(k))
    product_cols = [c for c in row if (c.startswith("O_") or c.startswith("S_")) and row[c]]
    product_score = min(len(product_cols) / 5, 1.0)
    return round((filled / len(required)) * 0.6 + product_score * 0.4, 2)


def _extract_number(val):
    if val is None: return None
    words = {'zero':0,'one':1,'two':2,'three':3,'four':4,'five':5,'six':6,'seven':7,'eight':8,'nine':9,'ten':10}
    s = str(val).strip().lower()
    if s in words: return words[s]
    m = re.search(r'\d+', str(val))
    return int(m.group()) if m else None

def block_to_dict(block):
    """Parse any field:value block into {CANONICAL_KEY: value}.
    Now supports: inline multi-value lines, fuzzy matching, and unresolved logging."""
    d = {}
    raw_lines = [normalize_text(ln).strip() for ln in block.replace('\r', '').split('\n')]
    # Expand inline multi-value lines (e.g. "MT250: 2cs, MT330: 1cs")
    lines = []
    for ln in raw_lines:
        lines.extend(_expand_inline_pairs(ln))
    i = 0
    while i < len(lines):
        line = lines[i]; i += 1
        if not line: continue
        m = _SEP_RE.match(line)
        if not m: continue
        label_raw, val_raw = m.group(1).strip(), m.group(2).strip()
        norm = re.sub(r'[^A-Z0-9]', '', label_raw.upper())
        if len(norm) < 2: continue
        # Exact lookup first, then fuzzy fallback
        canonical = FIELD_MAP.get(norm)
        if canonical is None:
            canonical = fuzzy_resolve_field(norm)
            # If fuzzy also failed (returned norm unchanged and norm not a known canonical)
            if canonical == norm and canonical not in _FIELD_SYNONYMS:
                _UNRESOLVED_LOG.append({"raw": label_raw, "normalized": norm, "value": val_raw})
        if not val_raw and i < len(lines) and lines[i]:
            nxt = lines[i].strip()
            if nxt and not _SEP_RE.match(nxt):
                nxt_norm = re.sub(r'[^A-Z0-9]', '', nxt.upper())
                is_known_key = False
                for k in FIELD_MAP.keys():
                    if nxt_norm == k or (nxt_norm.startswith(k) and len(k) >= 5):
                        is_known_key = True
                        break
                if not is_known_key:
                    val_raw = nxt; i += 1
        val_raw = val_raw.lstrip(':=\u2013|-').strip()
        if canonical not in d or not d[canonical]:
            d[canonical] = val_raw
    # Normalize phone numbers if present
    if 'CONTACTNO' in d and d['CONTACTNO']:
        d['CONTACTNO'] = normalize_phone(d['CONTACTNO'])
    return d

def get_val(d, keys, clean_numeric=False):
    """Look up value by canonical key or any synonym."""
    for k in keys:
        if k in d:
            v = d[k]; return _extract_number(v) if clean_numeric else v
        canon = FIELD_MAP.get(re.sub(r'[^A-Z0-9]', '', k.upper()), k)
        if canon in d:
            v = d[canon]; return _extract_number(v) if clean_numeric else v
    return None if clean_numeric else ""

# Universal WhatsApp timestamp patterns
_TS_SPLIT  = re.compile(r'\n(?=\d{1,2}/\d{1,2}/\d{2,4},\s*\d{1,2}:\d{2}(?:[\s\u202f]*[APap][Mm])?\s*[-\u2013]\s*)')
_TS_META   = re.compile(r'^(\d{1,2}/\d{1,2}/\d{2,4}),\s*\d{1,2}:\d{2}(?:[\s\u202f]*[APap][Mm])?\s*[-\u2013]\s*(.*?)(?:\s*:|\s*$)', re.IGNORECASE)
# ORDER section: handles "Order\nStock", "Order :", "ORDER IN CASE", "ORDER IN CASES"
_ORDER_SPLIT = re.compile(r'\n\s*Order(?:\s+IN\s+CASES?)?\s*:?\s*\n(?:\s*Stock\s*\n)?', re.IGNORECASE)
# STOCK section: handles "Stock", "Physical Stock", "Stock:", "Stock :"
_STOCK_SPLIT = re.compile(r'\n\s*(?:Physical\s+)?Stock(?:\s+IN\s+HAND)?\s*:?\s*\n', re.IGNORECASE)
_HAS_OUTLET  = re.compile(r"\b(?:outlet\s*name|outletname|outlen\s*name|shop\s*name|store\s*name|party\s*name|firm\s*name)\b", re.IGNORECASE)


def process_single_file(filepath, filter_date=None):
    clear_unresolved_log()  # Reset log for this run
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            text = f.read()
    except Exception as e:
        raise Exception(f"Error reading file: {e}")

    content = "\n" + normalize_text(text)
    entries = _TS_SPLIT.split(content)

    data = []
    for entry in entries:
        entry = entry.strip()
        if not entry: continue
        if not _HAS_OUTLET.search(entry): continue

        meta_match = _TS_META.search(entry)
        if not meta_match: continue
        msg_date = meta_match.group(1)
        if filter_date and msg_date != filter_date: continue

        order_parts = _ORDER_SPLIT.split(entry, maxsplit=1)
        order_raw = order_parts[1] if len(order_parts) > 1 else ""
        # Truncate order block at the next major section header
        order_block = re.split(r'\n\s*(?:Other\s+Brands?|Remarks?|Note|DPS|Indoor|Outdoor|Branding)', order_raw, flags=re.IGNORECASE)[0]
        
        stock_parts = _STOCK_SPLIT.split(order_parts[0], maxsplit=1)
        stock_block = stock_parts[1] if len(stock_parts) > 1 else ""

        stock_d  = block_to_dict(stock_block)
        order_d  = block_to_dict(order_block)
        master_d = block_to_dict(entry)

        try:
            day_name = datetime.strptime(msg_date, "%d/%m/%Y").strftime("%A")
        except:
            try: day_name = datetime.strptime(msg_date, "%d/%m/%y").strftime("%A")
            except: day_name = ""

        row = {
            "Date": msg_date,
            "HQ":   get_val(master_d, ["HEADQUARTER"]),
            "STAFF NAME": get_val(master_d, ["INCHARGE"]),
            "ROUTE": get_val(master_d, ["ROUTE"]),
            "DAY":  day_name,
            "OUTLET NAME": get_val(master_d, ["OUTLETNAME"]),
            "ADDRESS": get_val(master_d, ["ADDRESS"]),
            "OUTLET OWNER NAME": get_val(master_d, ["OWNERNAME"]),
            "CONTACT NO": get_val(master_d, ["CONTACTNO"]),
            "DPS_RAW": get_val(master_d, ["DPS"]),
            "Indoor Branding (Rack Display)": get_val(master_d, ["DISPLAY"]),
            "DPS": "",
            "Other Brand": get_val(master_d, ["OTHERBRAND"]),
            "DPS REMARK": "",
            "S_MT 250":       get_val(stock_d, ["MT250"],       clean_numeric=False),
            "S_MT 330":       get_val(stock_d, ["MT330"],       clean_numeric=False),
            "S_XT 330":       get_val(stock_d, ["XT330"],       clean_numeric=False),
            "S_JJ 200ML":     get_val(stock_d, ["JJ200ML"],     clean_numeric=False),
            "S_JJ 320ML":     get_val(stock_d, ["JJ320ML"],     clean_numeric=False),
            "S_BAM BAM 250ML":  get_val(stock_d, ["BAMBAM250ML"], clean_numeric=False),
            "S_BAM BAM 500 ML": get_val(stock_d, ["BAMBAM500ML"], clean_numeric=False),
            "S_BAM BAM 2.5 ML": get_val(stock_d, ["BAMBAM25L"],   clean_numeric=False),
            "S_CSD":          get_val(stock_d, ["CSD"],         clean_numeric=False),
            "S_MT CLASSIC":   get_val(stock_d, ["MTCLASSIC"],   clean_numeric=False),
            "S_OTHERS":       get_val(stock_d, ["OTHERS"],      clean_numeric=False),
            "O_MT 250":       get_val(order_d, ["MT250"],       clean_numeric=False),
            "O_MT 330":       get_val(order_d, ["MT330"],       clean_numeric=False),
            "O_XT 330":       get_val(order_d, ["XT330"],       clean_numeric=False),
            "O_JJ 200ML":     get_val(order_d, ["JJ200ML"],     clean_numeric=False),
            "O_JJ 320ML":     get_val(order_d, ["JJ320ML"],     clean_numeric=False),
            "O_BAM BAM 250ML":  get_val(order_d, ["BAMBAM250ML"], clean_numeric=False),
            "O_BAM BAM 500 ML": get_val(order_d, ["BAMBAM500ML"], clean_numeric=False),
            "O_BAM BAM 2.5 ML": get_val(order_d, ["BAMBAM25L"],   clean_numeric=False),
            "O_CSD":          get_val(order_d, ["CSD"],         clean_numeric=False),
            "O_MT CLASSIC":   get_val(order_d, ["MTCLASSIC"],   clean_numeric=False),
            "O_OTHERS":       get_val(order_d, ["OTHERS"],      clean_numeric=False),
            "D_MT 250": None, "D_MT 330": None, "D_XT 330": None,
            "D_JJ 200ML": None, "D_JJ 320ML": None,
            "D_BAM BAM 250ML": None, "D_BAM BAM 500 ML": None, "D_BAM BAM 2.5 ML": None,
            "D_CSD": None, "D_MT CLASSIC": None, "D_OTHERS": None,
            "STAFF REMARK": get_val(master_d, ["REMARKS"]),
            "Confidence": None,  # filled below
        }
        row["Confidence"] = _score_row(row)
        data.append(row)
    return pd.DataFrame(data)



def parse_to_cases(val):
    """Convert strings like '1cs', '12pc', '1 case', '12 p' to numeric cases."""
    if val is None or val == "": return 0
    s = str(val).lower().strip()
    # Extract numeric part (handles integers and decimals)
    num_match = re.search(r'(\d+(?:\.\d+)?)', s)
    if not num_match: return 0
    num = float(num_match.group(1))
    
    # If unit suggests pieces, divide by 26
    # Keywords: pc, ps, p, piece, pcs
    if re.search(r'\b(?:pc|ps|p|piece|pcs)\b', s) or (not re.search(r'\b(?:cs|case|ca)\b', s) and 'p' in s):
        return num / 26.0
    # Default is cases
    return num

def format_cases(total_cases):
    """Convert numeric cases back to 'X cs Y ps' format."""
    if total_cases == 0: return "0"
    cs = int(total_cases + 0.0001) # Small epsilon for float precision
    ps = round((total_cases - cs) * 26)
    if ps >= 26: # Handle rounding overflow
        cs += 1
        ps -= 26
    
    parts = []
    if cs > 0: parts.append(f"{cs} cs")
    if ps > 0: parts.append(f"{ps} ps")
    return " ".join(parts) if parts else "0"


def apply_styling(writer, df, sheet_name, report_type_label):
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    main_header_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#1D4E89', 'font_color': 'white', 'border': 1, 'text_wrap': True})
    stock_header_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#3498DB', 'font_color': 'white', 'border': 1})
    order_header_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#E67E22', 'font_color': 'white', 'border': 1})
    dispatch_header_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#2ECC71', 'font_color': 'white', 'border': 1})
    sub_header_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#F1F2F6', 'border': 1, 'size': 9})
    cell_fmt = workbook.add_format({'border': 1, 'valign': 'vcenter'})
    num_fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
    green_fmt = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'border': 1})
    red_fmt = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'border': 1})
    blue_chip_fmt = workbook.add_format({'bg_color': '#1D4E89', 'font_color': 'white', 'border': 1})
    orange_fmt = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C5700', 'border': 1})

    worksheet.merge_range(0, 0, 0, 13, f"OUTLET DETAILS - {report_type_label}", main_header_fmt)
    worksheet.merge_range(0, 14, 0, 24, "STOCK", stock_header_fmt)
    worksheet.merge_range(0, 25, 0, 35, "ORDER", order_header_fmt)
    worksheet.merge_range(0, 36, 0, 46, "DISPATCH", dispatch_header_fmt)

    headers = ["Date", "HQ", "STAFF NAME", "ROUTE", "DAY", "OUTLET NAME", "ADDRESS", "OUTLET OWNER NAME", "CONTACT NO", "DPS (RAW)", "Indoor Branding", "DPS", "Other Brand", "DPS REMARK"]
    products = ["MT 250", "MT 330", "XT 330", "JJ 200ML", "JJ 320ML", "AM BAM 250ML", "AM BAM 500 ML", "AM BAM 2.5 ML", "CSD", "MT CLASSIC", "OTHERS"]
    full_headers = headers + products + products + products + ["STAFF REMARK", "Confidence"]
    confidence_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#E8F5E9', 'font_color': '#1B5E20', 'border': 1, 'num_format': '0.00'})
    confidence_header_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#2ECC71', 'font_color': 'white', 'border': 1})
    
    # Initialize totals for product columns (14 to 46)
    col_totals = {col_idx: 0.0 for col_idx in range(14, 47)}
    
    for col_num, header in enumerate(full_headers):
        if header == "Confidence":
            fmt = confidence_header_fmt
        elif col_num < 14 or header == "STAFF REMARK":
            fmt = main_header_fmt
        else:
            fmt = sub_header_fmt
        worksheet.write(1, col_num, header, fmt)
    
    data_start_row = 2
    for row_idx in range(len(df)):
        for col_idx in range(len(df.columns)):
            val = df.iloc[row_idx, col_idx]
            is_numeric_col = 14 <= col_idx <= 46  # S_, O_, D_ product columns only
            is_confidence_col = col_idx == len(df.columns) - 1
            
            if is_numeric_col:
                col_totals[col_idx] += parse_to_cases(val)
                
            if pd.isna(val) or val == 0 or val == "0" or val == "":
                worksheet.write_blank(row_idx + data_start_row, col_idx, None, num_fmt if is_numeric_col else cell_fmt)
            elif is_confidence_col:
                worksheet.write(row_idx + data_start_row, col_idx, val, confidence_fmt)
            else:
                fmt = num_fmt if is_numeric_col else cell_fmt
                worksheet.write(row_idx + data_start_row, col_idx, val, fmt)
        
        # Validation and other cell-specific formatting
        worksheet.data_validation(row_idx + data_start_row, 10, row_idx + data_start_row, 10, {'validate': 'list', 'source': ['YES', 'NO']})
        worksheet.data_validation(row_idx + data_start_row, 11, row_idx + data_start_row, 11, {'validate': 'list', 'source': ['Max Tiger', 'Xtreme']})
        worksheet.data_validation(row_idx + data_start_row, 12, row_idx + data_start_row, 12, {'validate': 'list', 'source': ['YES', 'NO']})

    # Add TOTAL row at the bottom
    total_row_idx = len(df) + data_start_row
    total_label_fmt = workbook.add_format({'bold': True, 'align': 'right', 'valign': 'vcenter', 'fg_color': '#DCDDE1', 'border': 1})
    total_val_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#F8EFBA', 'border': 1})
    
    worksheet.merge_range(total_row_idx, 0, total_row_idx, 13, "GRAND TOTAL", total_label_fmt)
    for col_idx in range(14, 47):
        total_str = format_cases(col_totals[col_idx])
        worksheet.write(total_row_idx, col_idx, total_str, total_val_fmt)
    # Empty cells for STAFF REMARK and Confidence in total row
    worksheet.write(total_row_idx, len(df.columns) - 2, "", total_label_fmt)
    worksheet.write(total_row_idx, len(df.columns) - 1, "", total_label_fmt)

    worksheet.conditional_format(2, 10, len(df) + 1, 10, {'type': 'cell', 'criteria': 'equal to', 'value': '"YES"', 'format': green_fmt})
    worksheet.conditional_format(2, 10, len(df) + 1, 10, {'type': 'cell', 'criteria': 'equal to', 'value': '"NO"', 'format': red_fmt})
    worksheet.conditional_format(2, 11, len(df) + 1, 11, {'type': 'cell', 'criteria': 'equal to', 'value': '"Xtreme"', 'format': blue_chip_fmt})
    worksheet.conditional_format(2, 11, len(df) + 1, 11, {'type': 'cell', 'criteria': 'equal to', 'value': '"Max Tiger"', 'format': orange_fmt})
    worksheet.conditional_format(2, 12, len(df) + 1, 12, {'type': 'cell', 'criteria': 'equal to', 'value': '"YES"', 'format': green_fmt})
    worksheet.conditional_format(2, 12, len(df) + 1, 12, {'type': 'cell', 'criteria': 'equal to', 'value': '"NO"', 'format': red_fmt})
    
    # Conditional format for Confidence: green if >= 0.8, orange if < 0.6
    conf_col_idx = len(df.columns) - 1
    conf_green = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'border': 1, 'align': 'center'})
    conf_orange = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C5700', 'border': 1, 'align': 'center'})
    conf_red = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'border': 1, 'align': 'center'})
    conf_col_letter = chr(ord('A') + conf_col_idx)
    worksheet.conditional_format(f'{conf_col_letter}3:{conf_col_letter}{len(df)+2}', {'type': 'cell', 'criteria': '>=', 'value': 0.8, 'format': conf_green})
    worksheet.conditional_format(f'{conf_col_letter}3:{conf_col_letter}{len(df)+2}', {'type': 'cell', 'criteria': 'between', 'minimum': 0.6, 'maximum': 0.79, 'format': conf_orange})
    worksheet.conditional_format(f'{conf_col_letter}3:{conf_col_letter}{len(df)+2}', {'type': 'cell', 'criteria': '<', 'value': 0.6, 'format': conf_red})
    
    # Auto-filter and freeze
    worksheet.autofilter(1, 0, total_row_idx, len(df.columns) - 1)
    worksheet.freeze_panes(2, 6) # Freeze panes after Address column
    
    worksheet.set_column('A:B', 12)
    worksheet.set_column('C:D', 20)
    worksheet.set_column('E:E', 10)
    worksheet.set_column('F:H', 30)
    worksheet.set_column('I:I', 15)
    worksheet.set_column('J:N', 15)
    worksheet.set_column('O:AU', 12)
    worksheet.set_column(conf_col_idx, conf_col_idx, 12)  # Confidence column width

# --- Modern UI Application ---

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class ModernDSRApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Agro DSR Pro v7.0")
        self.geometry("900x650")
        
        try:
            logo_path = resource_path(os.path.join("assets", "xtreme.png"))
            if os.path.exists(logo_path):
                img = Image.open(logo_path)
                self.wm_iconphoto(True, ImageTk.PhotoImage(img))
        except Exception as e:
            print(f"Icon error: {e}")

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Sidebar
        self.sidebar = ctk.CTkFrame(self, width=220, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        
        try:
            pil_img = Image.open(resource_path(os.path.join("assets", "xtreme.png")))
            logo_img = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=(160, 280))
            self.logo_label = ctk.CTkLabel(self.sidebar, image=logo_img, text="")
            self.logo_label.pack(pady=20)
        except Exception as e:
            self.logo_label = ctk.CTkLabel(self.sidebar, text="AGRO DSR\nPRO", font=ctk.CTkFont(size=24, weight="bold"))
            self.logo_label.pack(pady=40)
        
        self.mode_label = ctk.CTkLabel(self.sidebar, text="Extraction Mode", anchor="w")
        self.mode_label.pack(fill="x", padx=20, pady=(20, 5))
        self.mode_selector = ctk.CTkSegmentedButton(self.sidebar, values=["Today Only", "Full History"])
        self.mode_selector.set("Today Only")
        self.mode_selector.pack(padx=10, fill="x")

        self.theme_label = ctk.CTkLabel(self.sidebar, text="Appearance Mode", anchor="w")
        self.theme_label.pack(fill="x", padx=20, pady=(30, 5))
        self.theme_option = ctk.CTkOptionMenu(self.sidebar, values=["Dark", "Light"], command=self.change_appearance_mode)
        self.theme_option.pack(padx=10, fill="x")

        # Main Content Area
        self.main_container = ctk.CTkFrame(self, corner_radius=15)
        self.main_container.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        
        self.setup_sales_tab()

    def setup_sales_tab(self):
        title = ctk.CTkLabel(self.main_container, text="Sales & Inventory Report", font=ctk.CTkFont(size=20, weight="bold"))
        title.pack(pady=(20, 30))
        file_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        file_frame.pack(fill="x", padx=30)
        self.sales_file_var = ctk.StringVar(value="No file selected...")
        entry = ctk.CTkEntry(file_frame, textvariable=self.sales_file_var, width=300, height=40)
        entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        btn = ctk.CTkButton(file_frame, text="Browse", width=100, height=40, command=lambda: self.browse_file(self.sales_file_var))
        btn.pack(side="right")
        self.sales_status = ctk.CTkLabel(self.main_container, text="Ready", text_color="gray")
        self.sales_status.pack(pady=(40, 5))
        self.sales_progress = ctk.CTkProgressBar(self.main_container, width=400)
        self.sales_progress.set(0)
        self.sales_progress.pack(pady=10)
        gen_btn = ctk.CTkButton(self.main_container, text="GENERATE SALES REPORT", font=ctk.CTkFont(size=14, weight="bold"), height=50, command=self.run_sales_process)
        gen_btn.pack(side="bottom", fill="x", padx=40, pady=30)


    def change_appearance_mode(self, new_mode):
        ctk.set_appearance_mode(new_mode)

    def browse_file(self, var, filetypes=[("Text Files", "*.txt")]):
        f = filedialog.askopenfilename(filetypes=filetypes)
        if f: var.set(f)

    def run_sales_process(self):
        filepath = self.sales_file_var.get()
        if not os.path.exists(filepath):
            messagebox.showwarning("Input Error", "Please select a valid WhatsApp chat export (.txt)")
            return
        mode = self.mode_selector.get()
        today = datetime.now().strftime("%d/%m/%Y")
        filter_date = today if mode == "Today Only" else None
        report_label = f"Daily ({today})" if mode == "Today Only" else "Full History"
        self.sales_status.configure(text="Processing records...", text_color="orange")
        self.sales_progress.set(0.3)
        self.update()
        try:
            df = process_single_file(filepath, filter_date)
            if df.empty:
                self.sales_progress.set(0)
                self.sales_status.configure(text="No data found", text_color="red")
                messagebox.showwarning("Empty Result", f"No reports found.")
                return
            self.sales_progress.set(0.6)
            output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=f"DSR_Sales_{datetime.now().strftime('%Y%m%d')}.xlsx", filetypes=[("Excel Files", "*.xlsx")])
            if not output_path:
                self.sales_progress.set(0)
                self.sales_status.configure(text="Cancelled", text_color="gray")
                return
            with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name="Report", index=False, startrow=2)
                apply_styling(writer, df, "Report", report_label)
            self.sales_progress.set(1.0)
            unresolved = get_unresolved_log()
            unresolved_msg = f"\nUnrecognized fields: {len(unresolved)} (check log)" if unresolved else "\nAll fields recognized ✓"
            self.sales_status.configure(text=f"Success! {len(df)} records saved.", text_color="green")
            messagebox.showinfo("Export Complete", f"Sales Report Created!\nRecords: {len(df)}{unresolved_msg}")
        except Exception as e:
            self.sales_progress.set(0)
            self.sales_status.configure(text="Error occurred", text_color="red")
            messagebox.showerror("Error", str(e))


if __name__ == "__main__":
    app = ModernDSRApp()
    app.mainloop()
