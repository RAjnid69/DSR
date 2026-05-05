import re
import pandas as pd
import os
import glob
from datetime import datetime
import customtkinter as ctk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from PIL import Image, ImageTk
import ctypes
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
                    'salesman','sales man','by','reported by'],
    'ROUTE':       ['route','route point','routepoint','route name','routename','route no',
                    'route(point)','routepoint','beat','beat name','beat no','territory','sector'],
    'OUTLETNAME':  ['outlet name','outletname','outlet','outlen name','outletnam','outlets name',
                    "outlet name'","outlet name' ",
                    'shop name','shopname','shop','store','store name','party','party name',
                    'firm name','firm','counter name'],
    'ADDRESS':     ['address','addedss','addess','adress','addr','add','location','place',
                    'locality','ward no','ward'],
    'OWNERNAME':   ['owner name','ownername','owner','owoner name','ownernam','dealer name',
                    'dealername','prop','proprietor','proprietor name','proprietar','malik','sahuji'],
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
                    'bam-bam 250','bam bam nimboo','bambam nimboo','bam bam lemon',
                    'bam bam mist','bambam mist','bam mist 250ml','bam mist 250',
                    'bam bam 250ml'],
    'BAMBAM500ML': ['bam bam 500','bambam500ml','bam bam 500ml','bambam 500ml','bambam500',
                    'bam-bam 500','bam bam hydration 500ml','bam bam hydration','bambam hydration',
                    'bam bam cola','bambam cola','bam bam csd'],
    'BAMBAM25L':   ['bam bam 2.5','bambam25l','bam bam 2.5l','bambam 2.5l','bambam2.5','bam-bam 2.5',
                    'bam bam 2.5ml'],
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

def _extract_number(val):
    if val is None: return None
    words = {'zero':0,'one':1,'two':2,'three':3,'four':4,'five':5,'six':6,'seven':7,'eight':8,'nine':9,'ten':10}
    s = str(val).strip().lower()
    if s in words: return words[s]
    m = re.search(r'\d+', str(val))
    return int(m.group()) if m else None

def block_to_dict(block):
    """Parse any field:value block into {CANONICAL_KEY: value}."""
    d = {}
    lines = [normalize_text(ln).strip() for ln in block.replace('\r', '').split('\n')]
    i = 0
    while i < len(lines):
        line = lines[i]; i += 1
        if not line: continue
        m = _SEP_RE.match(line)
        if not m: continue
        label_raw, val_raw = m.group(1).strip(), m.group(2).strip()
        norm = re.sub(r'[^A-Z0-9]', '', label_raw.upper())
        if len(norm) < 2: continue
        canonical = FIELD_MAP.get(norm, norm)
        if not val_raw and i < len(lines) and lines[i]:
            nxt = lines[i].strip()
            if nxt and not _SEP_RE.match(nxt):
                val_raw = nxt; i += 1
        val_raw = val_raw.lstrip(':=\u2013|-').strip()
        if canonical not in d or not d[canonical]:
            d[canonical] = val_raw
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

def discover_hq(file_path):
    content = ""
    try:
        with open(file_path, "r", encoding="utf-8") as f: content = f.read(5000)
    except: pass
    hq = re.search(r'(?:HEADQUARTER|H\.?Q\.?)\s*[:=.\-]\s*([^\n\r]+)', content, re.IGNORECASE)
    if hq:
        val = hq.group(1).strip().upper()
        if val and val not in ("POSTER", ""): return val
    bn = re.search(r"with\s+([A-Z0-9\-]+)", os.path.basename(file_path), re.IGNORECASE)
    return bn.group(1).strip().upper() if bn else "POSTER"

def process_single_file(filepath, filter_date=None):
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
        order_block = order_parts[1] if len(order_parts) > 1 else ""
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
            "S_MT 250":       get_val(stock_d, ["MT250"],       clean_numeric=True),
            "S_MT 330":       get_val(stock_d, ["MT330"],       clean_numeric=True),
            "S_XT 330":       get_val(stock_d, ["XT330"],       clean_numeric=True),
            "S_JJ 200ML":     get_val(stock_d, ["JJ200ML"],     clean_numeric=True),
            "S_JJ 320ML":     get_val(stock_d, ["JJ320ML"],     clean_numeric=True),
            "S_BAM BAM 250ML":  get_val(stock_d, ["BAMBAM250ML"], clean_numeric=True),
            "S_BAM BAM 500 ML": get_val(stock_d, ["BAMBAM500ML"], clean_numeric=True),
            "S_BAM BAM 2.5 ML": get_val(stock_d, ["BAMBAM25L"],   clean_numeric=True),
            "S_CSD":          get_val(stock_d, ["CSD"],         clean_numeric=True),
            "S_MT CLASSIC":   get_val(stock_d, ["MTCLASSIC"],   clean_numeric=True),
            "S_OTHERS":       get_val(stock_d, ["OTHERS"],      clean_numeric=True),
            "O_MT 250":       get_val(order_d, ["MT250"],       clean_numeric=True),
            "O_MT 330":       get_val(order_d, ["MT330"],       clean_numeric=True),
            "O_XT 330":       get_val(order_d, ["XT330"],       clean_numeric=True),
            "O_JJ 200ML":     get_val(order_d, ["JJ200ML"],     clean_numeric=True),
            "O_JJ 320ML":     get_val(order_d, ["JJ320ML"],     clean_numeric=True),
            "O_BAM BAM 250ML":  get_val(order_d, ["BAMBAM250ML"], clean_numeric=True),
            "O_BAM BAM 500 ML": get_val(order_d, ["BAMBAM500ML"], clean_numeric=True),
            "O_BAM BAM 2.5 ML": get_val(order_d, ["BAMBAM25L"],   clean_numeric=True),
            "O_CSD":          get_val(order_d, ["CSD"],         clean_numeric=True),
            "O_MT CLASSIC":   get_val(order_d, ["MTCLASSIC"],   clean_numeric=True),
            "O_OTHERS":       get_val(order_d, ["OTHERS"],      clean_numeric=True),
            "D_MT 250": None, "D_MT 330": None, "D_XT 330": None,
            "D_JJ 200ML": None, "D_JJ 320ML": None,
            "D_BAM BAM 250ML": None, "D_BAM BAM 500 ML": None, "D_BAM BAM 2.5 ML": None,
            "D_CSD": None, "D_MT CLASSIC": None, "D_OTHERS": None,
            "STAFF REMARK": get_val(master_d, ["REMARKS"]),
        }
        data.append(row)
    return pd.DataFrame(data)


def clean_val(val):
    if val:
        val = val.strip()
        if val in ["-", "", "null", ".", "N/A", "na"]:
            return None
    return val

POSTER_FIELD_PATTERNS = {
    "outlet_name": [
        r"outlet\s*name", r"outlen\s*name", r"outlets?\s*name",
        r"shop\s*name", r"store\s*name", r"आउटलेट\s*नाम",
    ],
    "address": [
        r"address", r"addedss", r"addess", r"adress", r"add(?:ress)?",
        r"location", r"ठेगाना", r"addr",
    ],
    "contact": [
        r"contact(?:\s*no\.?)?", r"contract", r"ph\s*no\.?",
        r"mobile\s*no\.?", r"सम्पर्क", r"फोन", r"ph(?:one)?",
        r"contact\s*number", r"mob",
    ],
    "owner_name": [
        r"owner\s*name", r"owoner\s*name", r"owner", r"proprietor(?:\s*name)?",
        r"मालिक", r"malik",
    ],
    "staff_name": [
        r"staff\s*name", r"incharge\s*name", r"incharge", r"name", r"नाम", r"by",
    ],
    "area": [r"route\s*name", r"route", r"area", r"hq", r"beat"],
    # Poster counts — all group variants
    "xt_poster": [
        r"no\.?\s*of\s*xtreme\s*posters?\s*pasted",
        r"count\s*of\s*xtreme\s*poster",
        r"count\s*of\s*xt(?:reme)?",
        r"xtreme\s*poster", r"extreme\s*poster",
        r"xt\s*poster", r"xtreme",
        r"xterm(?:e)?",           # NPJ: "xterm---61"
        r"ex(?:treme)?",          # NPJ: "Ex=84"
    ],
    "mt250_poster": [
        r"no\.?\s*of\s*mt\s*250\s*posters?\s*pasted",
        r"count\s*of\s*max\s*tiger\s*poster\s*250",
        r"count\s*of\s*max\s*tiger",
        r"mt\s*250\s*poster", r"mt250\s*poster",
        r"max\s*tiger\s*poster\s*250",
        r"count\s*of\s*mt\s*poster",
        r"mt\s*poster",           # PKR-1: "MT POSTER-53"
        r"max\s*tiger",           # NPJ: "maxtiger---119"
        r"mx",                    # NPJ: "mx=84"
    ],
    "mt330_poster": [
        r"no\.?\s*of\s*mt\s*330\s*posters?\s*pasted",
        r"count\s*of\s*max\s*tiger\s*330",
        r"mt\s*330\s*poster", r"mt330\s*poster",
        r"max\s*tiger\s*330",
    ],
    "outlet_count": [
        r"count\s*of\s*outlets?",   # PKR-1: "COUNT OF OUTLETS-49"
        r"outlets?",                # NPJ: "Outlet---61" / "Outlet 66"
        r"outlet\s*visited",
        r"no\.?\s*of\s*outlets?",
        r"total\s*outlets?",
    ],
    "bam_bam":  [r"bam\s*bam", r"bambam"],
    "dps_board": [r"dps[/\s]*gsb\s*board", r"dps\s*board", r"gsb\s*board", r"board", r"dps"],
}

KNOWN_SENDERS = {
    "+977 980-3539493": "Rupesh Nepali",
    "+977 980-6645229": "Anush Kunwar",
    "+977 970-6015375": "Bhuban Dauliya",
    "+977 981-8811222": "Roshan Shrestha",
    "+977 974-2515634": "Aryan Gurung",
    "+977 981-5174483": "Staff (Unknown)",
}

# Outlet present check — relaxed for poster format (no "name" suffix needed)
_HAS_OUTLET_POSTER = re.compile(
    r'\b(?:outlet\s*name|outlen\s*name|shop\s*name|store\s*name|outlet\s*-|outlet\s*=)',
    re.IGNORECASE
)

def extract_poster_field(block, key):
    """Extract a field from a poster message block. Handles - and : separators."""
    for pat in POSTER_FIELD_PATTERNS[key]:
        # Match:  <whitespace><pattern><whitespace><separator><whitespace><value>
        m = re.search(
            rf'(?:^|\n)[ \t]*{pat}[ \t]*[:\-=\.]+[ \t]*(.*)',
            block, re.IGNORECASE | re.MULTILINE
        )
        if m:
            val = m.group(1).split('\n')[0].strip()
            # Also check "pattern---value" style (NPJ)
            if not val:
                m2 = re.search(rf'(?:^|\n)[ \t]*{pat}[-=]+([\d]+)', block, re.IGNORECASE | re.MULTILINE)
                if m2:
                    val = m2.group(1).strip()
            if val and val not in ['-', '', '---', '.']:
                return clean_val(normalize_text(val))
    return None

def _parse_poster_block(body, date, sender):
    """
    Parse one message body into a poster outlet record dict.
    Returns None if no outlet name found.
    """
    body_norm = normalize_text(body)

    # Get outlet name — required
    outlet = extract_poster_field(body, "outlet_name")
    if not outlet:
        return None

    # Staff name from body first, fall back to sender
    staff = extract_poster_field(body, "staff_name")
    if not staff or len(staff.split()) > 6:
        staff = KNOWN_SENDERS.get(sender, sender)

    rec = {
        "date":         date,
        "staff_name":   staff,
        "outlet_name":  outlet,
        "address":      extract_poster_field(body, "address"),
        "contact":      extract_poster_field(body, "contact"),
        "owner_name":   extract_poster_field(body, "owner_name"),
        "xt_poster":    extract_poster_field(body, "xt_poster"),
        "mt250_poster": extract_poster_field(body, "mt250_poster"),
        "mt330_poster": extract_poster_field(body, "mt330_poster"),
        "bam_bam":      extract_poster_field(body, "bam_bam"),
        "dps_board":    extract_poster_field(body, "dps_board"),
        "outlet_count": None,
    }

    # Fix: if owner_name looks like a phone number, move it to contact
    if rec["owner_name"] and re.match(r'\d{7,}', str(rec["owner_name"])):
        if not rec["contact"]:
            rec["contact"] = rec["owner_name"]
        rec["owner_name"] = None

    return rec

def process_poster_file(filepath, filter_date=None):
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            lines = f.readlines()
    except Exception as e:
        raise Exception(f"Error reading file: {e}")

    # Split into per-timestamp messages
    messages = []
    current_msg = None
    for line in lines:
        line_norm = normalize_text(line)
        match = _TS_META.match(line_norm)
        if match:
            if current_msg:
                messages.append(current_msg)
            current_msg = {
                "date":   match.group(1),
                "sender": match.group(2).strip(),
                "body":   ""
            }
        elif current_msg:
            current_msg["body"] += line
    if current_msg:
        messages.append(current_msg)

    data = []
    for msg in messages:
        body = msg["body"]
        if not body.strip():
            continue
        if filter_date and msg["date"] != filter_date:
            continue

        body_norm = normalize_text(body)

        # ── Strategy 1: body contains "outlet name" → direct outlet record ──
        if _HAS_OUTLET_POSTER.search(body_norm):
            rec = _parse_poster_block(body, msg["date"], msg["sender"])
            if rec:
                data.append(rec)
            continue

        # ── Strategy 2: summary message (NAME / ROUTE / counts) ──
        # PKR-1 style:  "NAME- X\nROUTE NAME- Y\nCOUNT OF OUTLETS- N\nCOUNT OF MAX TIGER POSTER- N\nCOUNT OF XTREME POSTER- N"
        # NPJ style:    "Name X Route Y\nOutlet---N\nmaxtiger---N\nxterm---N"
        # Both: extract one summary row per block
        xt   = extract_poster_field(body, "xt_poster")
        mt   = extract_poster_field(body, "mt250_poster")
        oc   = extract_poster_field(body, "outlet_count")

        if xt or mt or oc:
            staff = extract_poster_field(body, "staff_name")
            if not staff or len(staff.split()) > 6:
                staff = KNOWN_SENDERS.get(msg["sender"], msg["sender"])
            area  = extract_poster_field(body, "area")

            data.append({
                "date":         msg["date"],
                "staff_name":   staff,
                "outlet_name":  (f"[DAILY SUMMARY] {area or ''}".strip() + (f" (Outlets: {oc})" if oc else "")),
                "address":      area,
                "contact":      None,
                "owner_name":   None,
                "xt_poster":    xt,
                "mt250_poster": mt,
                "mt330_poster": extract_poster_field(body, "mt330_poster"),
                "bam_bam":      None,
                "dps_board":    None,
                "outlet_count": oc,
            })

    return pd.DataFrame(data)

def apply_poster_styling(template_path, output_path, df, hq_name):
    wb = load_workbook(template_path)
    target_sheet = wb.active
    target_sheet.title = hq_name[:31]
    for sheet in wb.sheetnames:
        if sheet != target_sheet.title:
            del wb[sheet]
            
    ws = target_sheet
    area_col_idx = None
    for cell in ws[1]:
        if cell.value and "AREA" == str(cell.value).strip().upper():
            area_col_idx = cell.column
            break
    if area_col_idx:
        ws.delete_cols(area_col_idx)

    for row in ws.iter_rows(min_row=2, max_row=1000):
        for cell in row:
            cell.value = None

    first_data_row = 2
    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    from openpyxl.formatting.rule import CellIsRule
    green_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
    green_font = Font(color='006100')
    red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    red_font = Font(color='9C0006')

    for i, row_data in df.iterrows():
        row_num = first_data_row + i
        sn = row_num - 1
        mapping = {
            "A": sn, "B": row_data["date"], "C": hq_name, "D": row_data["staff_name"] or "",
            "E": row_data["outlet_name"] or "", "F": row_data["address"] or "", "G": row_data["contact"] or "",
            "H": row_data["owner_name"] or "", "I": row_data["xt_poster"], "J": row_data["mt250_poster"],
            "K": row_data["mt330_poster"], "L": row_data["bam_bam"], "M": row_data["dps_board"] or ""
        }
        for col_letter, val in mapping.items():
            cell = ws[f"{col_letter}{row_num}"]
            cell.value = val if val is not None else ""
            cell.border = border
            cell.alignment = Alignment(horizontal="left", vertical="center")
            cell.font = Font(name="Arial", size=10)

    last_row = first_data_row + len(df) - 1
    if last_row >= first_data_row:
        ws.conditional_formatting.add(f'N{first_data_row}:N{last_row}',
            CellIsRule(operator='equal', formula=['"yes"'], fill=green_fill, font=green_font))
        ws.conditional_formatting.add(f'N{first_data_row}:N{last_row}',
            CellIsRule(operator='equal', formula=['"no"'], fill=red_fill, font=red_font))
    wb.save(output_path)

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
    products = ["MT 250", "MT 330", "XT 330", "JJ 200ML", "JJ 320ML", "BAM BAM 250ML", "BAM BAM 500 ML", "BAM BAM 2.5 ML", "CSD", "MT CLASSIC", "OTHERS"]
    full_headers = headers + products + products + products + ["STAFF REMARK"]
    for col_num, header in enumerate(full_headers):
        fmt = main_header_fmt if col_num < 14 or col_num == len(full_headers)-1 else sub_header_fmt
        worksheet.write(1, col_num, header, fmt)

    for row_idx in range(len(df)):
        for col_idx in range(len(df.columns)):
            val = df.iloc[row_idx, col_idx]
            is_numeric_col = 14 <= col_idx < (len(df.columns) - 1)
            if pd.isna(val) or val == 0 or val == "0":
                worksheet.write_blank(row_idx + 2, col_idx, None, num_fmt if is_numeric_col else cell_fmt)
            else:
                fmt = num_fmt if is_numeric_col else cell_fmt
                worksheet.write(row_idx + 2, col_idx, val, fmt)
        worksheet.data_validation(row_idx + 2, 10, row_idx + 2, 10, {'validate': 'list', 'source': ['YES', 'NO']})
        worksheet.data_validation(row_idx + 2, 11, row_idx + 2, 11, {'validate': 'list', 'source': ['Max Tiger', 'Xtreme']})
        worksheet.data_validation(row_idx + 2, 12, row_idx + 2, 12, {'validate': 'list', 'source': ['YES', 'NO']})

    worksheet.conditional_format(2, 10, len(df) + 1, 10, {'type': 'cell', 'criteria': 'equal to', 'value': '"YES"', 'format': green_fmt})
    worksheet.conditional_format(2, 10, len(df) + 1, 10, {'type': 'cell', 'criteria': 'equal to', 'value': '"NO"', 'format': red_fmt})
    worksheet.conditional_format(2, 11, len(df) + 1, 11, {'type': 'cell', 'criteria': 'equal to', 'value': '"Xtreme"', 'format': blue_chip_fmt})
    worksheet.conditional_format(2, 11, len(df) + 1, 11, {'type': 'cell', 'criteria': 'equal to', 'value': '"Max Tiger"', 'format': orange_fmt})
    worksheet.conditional_format(2, 12, len(df) + 1, 12, {'type': 'cell', 'criteria': 'equal to', 'value': '"YES"', 'format': green_fmt})
    worksheet.conditional_format(2, 12, len(df) + 1, 12, {'type': 'cell', 'criteria': 'equal to', 'value': '"NO"', 'format': red_fmt})

    worksheet.set_column('A:B', 12)
    worksheet.set_column('C:D', 20)
    worksheet.set_column('F:H', 25)
    worksheet.set_column('I:J', 15)
    worksheet.set_column('K:M', 15)
    worksheet.set_column('N:N', 30)
    worksheet.set_column('O:AU', 10)

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

        self.title("Agro DSR Pro v5.0")
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
        
        self.tabview = ctk.CTkTabview(self.main_container, width=500)
        self.tabview.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.sales_tab = self.tabview.add("Sales DSR")
        self.poster_tab = self.tabview.add("Poster DSR")
        
        self.setup_sales_tab()
        self.setup_poster_tab()

    def setup_sales_tab(self):
        title = ctk.CTkLabel(self.sales_tab, text="Sales & Inventory Report", font=ctk.CTkFont(size=20, weight="bold"))
        title.pack(pady=(20, 30))
        file_frame = ctk.CTkFrame(self.sales_tab, fg_color="transparent")
        file_frame.pack(fill="x", padx=30)
        self.sales_file_var = ctk.StringVar(value="No file selected...")
        entry = ctk.CTkEntry(file_frame, textvariable=self.sales_file_var, width=300, height=40)
        entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        btn = ctk.CTkButton(file_frame, text="Browse", width=100, height=40, command=lambda: self.browse_file(self.sales_file_var))
        btn.pack(side="right")
        self.sales_status = ctk.CTkLabel(self.sales_tab, text="Ready", text_color="gray")
        self.sales_status.pack(pady=(40, 5))
        self.sales_progress = ctk.CTkProgressBar(self.sales_tab, width=400)
        self.sales_progress.set(0)
        self.sales_progress.pack(pady=10)
        gen_btn = ctk.CTkButton(self.sales_tab, text="GENERATE SALES REPORT", font=ctk.CTkFont(size=14, weight="bold"), height=50, command=self.run_sales_process)
        gen_btn.pack(side="bottom", fill="x", padx=40, pady=30)

    def setup_poster_tab(self):
        title = ctk.CTkLabel(self.poster_tab, text="Poster Dedicated Report", font=ctk.CTkFont(size=20, weight="bold"))
        title.pack(pady=(20, 30))
        chat_label = ctk.CTkLabel(self.poster_tab, text="Poster Chat Export (.txt):", anchor="w")
        chat_label.pack(fill="x", padx=30)
        chat_frame = ctk.CTkFrame(self.poster_tab, fg_color="transparent")
        chat_frame.pack(fill="x", padx=30, pady=(0, 20))
        self.poster_chat_var = ctk.StringVar(value="No chat file selected...")
        entry = ctk.CTkEntry(chat_frame, textvariable=self.poster_chat_var, width=300, height=40)
        entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        btn = ctk.CTkButton(chat_frame, text="Browse", width=100, height=40, command=lambda: self.browse_file(self.poster_chat_var, [("Text Files", "*.txt")]))
        btn.pack(side="right")
        info_label = ctk.CTkLabel(self.poster_tab, text="Using Default Bhaktapur Template", text_color="gray", font=ctk.CTkFont(size=12, slant="italic"))
        info_label.pack(pady=(0, 20))
        self.poster_status = ctk.CTkLabel(self.poster_tab, text="Ready", text_color="gray")
        self.poster_status.pack(pady=(40, 5))
        self.poster_progress = ctk.CTkProgressBar(self.poster_tab, width=400)
        self.poster_progress.set(0)
        self.poster_progress.pack(pady=10)
        gen_btn = ctk.CTkButton(self.poster_tab, text="GENERATE POSTER REPORT", font=ctk.CTkFont(size=14, weight="bold"), height=50, command=self.run_poster_process)
        gen_btn.pack(side="bottom", fill="x", padx=40, pady=40)

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
            self.sales_status.configure(text=f"Success! {len(df)} records saved.", text_color="green")
            messagebox.showinfo("Export Complete", f"Sales Report Created!\nRecords: {len(df)}")
        except Exception as e:
            self.sales_progress.set(0)
            self.sales_status.configure(text="Error occurred", text_color="red")
            messagebox.showerror("Error", str(e))

    def run_poster_process(self):
        chat_path = self.poster_chat_var.get()
        temp_path = resource_path(os.path.join("assets", "poster_template.xlsx"))
        if not os.path.exists(chat_path):
            messagebox.showwarning("Input Error", "Please select a valid chat export file.")
            return
        hq_name = discover_hq(chat_path)
        mode = self.mode_selector.get()
        today = datetime.now().strftime("%d/%m/%Y")
        filter_date = today if mode == "Today Only" else None
        if not os.path.exists(temp_path):
            temp_path = os.path.join("assets", "poster_template.xlsx")
            if not os.path.exists(temp_path):
                messagebox.showerror("Template Missing", f"Could not find default template at:\n{temp_path}")
                return
        self.poster_status.configure(text=f"Parsing {hq_name} data...", text_color="orange")
        self.poster_progress.set(0.3)
        self.update()
        try:
            df = process_poster_file(chat_path, filter_date)
            if df.empty:
                self.poster_progress.set(0)
                self.poster_status.configure(text="No data found", text_color="red")
                messagebox.showwarning("Empty Result", f"No poster reports found for {mode}.")
                return
            self.poster_progress.set(0.6)
            output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=f"Poster_{hq_name}_{datetime.now().strftime('%Y%m%d')}.xlsx", filetypes=[("Excel Files", "*.xlsx")])
            if not output_path:
                self.poster_progress.set(0)
                self.poster_status.configure(text="Cancelled", text_color="gray")
                return
            apply_poster_styling(temp_path, output_path, df, hq_name)
            self.poster_progress.set(1.0)
            self.poster_status.configure(text=f"Success! {len(df)} posters tracked.", text_color="green")
            messagebox.showinfo("Export Complete", f"Poster Report Created!\nHQ: {hq_name}\nRecords: {len(df)}")
        except Exception as e:
            self.poster_progress.set(0)
            self.poster_status.configure(text="Error occurred", text_color="red")
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    app = ModernDSRApp()
    app.mainloop()
