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

# =============================================================================
# ROBUST PARSING ENGINE (v5.5 Universal)
# =============================================================================

# Nepali digit normaliser
_NEPALI_DIGITS = str.maketrans("०१२३४५६७८९\u0966\u0967\u0968\u0969\u096a\u096b\u096c\u096d\u096e\u096f", "01234567890123456789")

def normalize_text(text):
    return str(text).translate(_NEPALI_DIGITS) if text else ""

# Field synonym map -> canonical key
_FIELD_SYNONYMS = {
    # --- Header / Outlet info ---
    'HEADQUARTER': ['headquarter','hq','hq name','hq area','hqarea','h q','h.q','h.q.',
                    'head quarter','headqarter','headquartor','headquter','headq'],
    'INCHARGE':    ['incharge','incharge name','inchargename','in charge','in charge name',
                    'route incharge name','route incharge','staff','staff name','staffname',
                    'salesman','sales man','by','reported by','name','नाम','supervisor','supervisor name','name supervisor'],
    'ROUTE':       ['route','route point','routepoint','route name','routename','route no',
                    'route(point)','routepoint','beat','beat name','beat no','territory','sector',
                    'road','road point','road(point)'],
    'OUTLETNAME':  ['outlet name','outletname','outlet','outlen name','outletnam','outlets name',
                    "outlet name'","outlet name' ","outlet :","outlet-",
                    'shop name','shopname','shop','store','store name','party','party name',
                    'firm name','firm','counter name','आउटलेट नाम'],
    'ADDRESS':     ['address','addedss','addess','adress','addr','add','location','place',
                    'locality','ward no','ward','ठेगाना'],
    'OWNERNAME':   ['owner name','ownername','owner','owoner name','ownernam',
                'prop','proprietor','proprietor name','proprietar','malik','sahuji','मालिक'],
    'CONTACTNO':   ['contact no','contactno','contact','contact no.','contact number',
                    'contactnumber','mobile','mobile no','mobile no.','mobile number',
                    'phone','ph','ph no','ph.no','cell','mob','सम्पर्क','फोन'],
    'OUTLETTYPE':  ['outlet type','outlettype','type','outlet type(category)','outlet type(category)','category'],
    'CHANNEL':     ['channel'],
    'PANNO':       ['pan no','panno','pan','pan number'],
    'VOLUME':      ['volume'],
    'WSTIEUP':     ['w/s tie-up','w/s tieup','tie-up','tieup','ws tie-up'],
    'MBQ':         ['mbq','total mbq status','mbq status','mbqstatus'],
    'DPS':         ['dps','dps/gsb board','dps gsb board','gsb board','display agreement',
                    'display point of sale','dps type','display type','point of sale','pos','dps board'],
    'DISPLAY':     ['indoor branding','rack display','display rack','branding','indoor display',
                    'dps/gsb size','dps gsb size','gsb size','display'],
    'TASTING':     ['tasting'],
    'TSHIRT':      ['t-shirt','tshirt'],
    'POSTER_COUNT':['poster','no. of poster','poster count'],
    'CHITDRAW':    ['chitdraw','chit draw'],
    'CANBLOCKOUT': ['can blockout','canblockout'],
    'REMARKS':     ['remarks','remark','note','notes','observation','comment','comments','extra note'],
    
    # --- Products ---
    'MT250':       ['mt 250','mt250','maxtiger250','max tiger 250','max tiger250',
                    'mt 250ml','maxtiger 250','mt-250','m.t. 250','m.t 250','max tiger 250ml'],
    'MT330':       ['mt 330','mt330','maxtiger330','max tiger 330','max tiger330',
                    'mt 330ml','maxtiger 330','mt-330','m.t. 330','m.t 330','max tiger 330ml'],
    'XT330':       ['xt 330','xt330','xtreme330','xtreme 330','xt 330ml','extreme 330',
                    'extreme330','xtreme-330','xt-330','x treme 330','xtreme 330ml','xt330ml'],
    'JJ200ML':     ['jj 200ml','jj200ml','juice jelly 200','juicejelly200ml','jj 200',
                    'juice & jelly 200','j&j 200ml','jj200','juice jelly 200ml','juice jelly200ml'],
    'JJ320ML':     ['jj 320ml','jj320ml','juice jelly 320','juicejelly320ml','jj 320',
                    'juice & jelly 320','j&j 320ml','jj320','juice jelly 320ml','juice jelly320ml','juice jelly','juicejelly'],
    'BAMBAMNIMBOO':['bam bam nimboo','bambam nimboo','nimboo','bam bam lemon','lemon','nimbu'],
    'BAMBAMMIST':  ['bam bam mist','bambam mist','mist','bam mist 250ml','bam mist'],
    'BAMBAMCOLA':  ['bam bam cola','bambam cola','cola'],
    'BAMBAMHYDRATION':['bam bam hydration 500ml','bam bam hydration','bambam hydration','hydration'],
    'CSD':         ['csd','csd250ml','csd 250ml','csd cran berry','csd cranberry','carbonated soft drink','cranberry'],
    'BAMBAM25L':   ['bam bam 2.5','bambam25l','bam bam 2.5l','bambam 2.5l','bambam2.5','bam-bam 2.5'],
    'MTCLASSIC':   ['mt classic','mtclassic','mt330classic','mt 330 classic','max tiger classic','classic 330','classic'],
    'OTHERS':      ['liquor','miscellaneous','misc','other item','other brand','otherbrand','other brands stock','red bull','evil','sting','powerpunch'],
}

FIELD_MAP: dict = {}
for _can, _vars in _FIELD_SYNONYMS.items():
    FIELD_MAP[re.sub(r'[^A-Z0-9]', '', _can)] = _can
    for _v in _vars:
        FIELD_MAP[re.sub(r'[^A-Z0-9]', '', _v.upper())] = _can

# Separators: colon, equals, em-dash, pipe, semicolon, apostrophe-colon, bare hyphen
_SEP_RE = re.compile(r"^(.*?)\s*(?:[:=\u2013|;']|(?<![A-Z0-9])-(?![A-Z0-9]))\s*(.*)", re.IGNORECASE)

def _extract_number(val):
    if val is None: return 0
    words = {'zero':0,'one':1,'two':2,'three':3,'four':4,'five':5,'six':6,'seven':7,'eight':8,'nine':9,'ten':10}
    s = str(val).strip().lower()
    if s in words: return words[s]
    # Handle fractions like "1/2" -> 0.5
    if '/' in s:
        try:
            num, den = s.split('/', 1)
            return float(num) / float(den)
        except: pass
    m = re.search(r'(\d+\.?\d*)', str(val))
    return float(m.group()) if m else 0

def block_to_dict(block):
    """Parse any field:value block into {CANONICAL_KEY: value}."""
    d = {}
    lines = [normalize_text(ln).strip() for ln in block.replace('\r', '').split('\n')]
    sorted_syns = sorted(FIELD_MAP.items(), key=lambda x: len(x[0]), reverse=True)
    i = 0
    while i < len(lines):
        line = lines[i]; i += 1
        if not line: continue
        m = _SEP_RE.match(line)
        if m:
            label_raw, val_raw = m.group(1).strip(), m.group(2).strip()
        else:
            # SMART FALLBACK: Handle missing colons/separators AND inline fields (KTM style)
            # First, check if the line contains multiple fields (e.g. "Address: ... Owner: ...")
            # We look for a known key pattern followed by a separator further in the line
            parts = [line]
            line_up = line.upper()
            for syn_norm, canonical in sorted_syns:
                # Look for "  KEY:" or "  KEY ;" pattern
                inline_re = re.compile(r'(\s{2,}' + re.escape(syn_norm) + r'[:;])', re.IGNORECASE)
                # We need to find where syn_norm (which has no spaces) matches the line
                # This is a bit hard, so we'll use a simpler heuristic for KTM
                if "OWNER NAME" in line_up and "ADDRESS" in line_up:
                    owner_idx = line_up.find("OWNER NAME")
                    parts = [line[:owner_idx], line[owner_idx:]]
                    break

            for part in parts:
                matched = False
                part_norm = re.sub(r'[^A-Z0-9]', '', part.upper())
                for syn_norm, canonical in sorted_syns:
                    if part_norm.startswith(syn_norm) and len(part_norm) > len(syn_norm):
                        label_raw = canonical
                        curr_norm = ""
                        val_start_idx = 0
                        for idx, char in enumerate(part.upper()):
                            if re.match(r'[A-Z0-9]', char):
                                curr_norm += char
                            if curr_norm == syn_norm:
                                val_start_idx = idx + 1
                                break
                        val_raw = part[val_start_idx:].strip()
                        val_raw = val_raw.lstrip(':= \u2013|;\'-').strip()
                        
                        norm = re.sub(r'[^A-Z0-9]', '', label_raw.upper())
                        canonical = FIELD_MAP.get(norm, norm)
                        if canonical not in d or not d[canonical]:
                            d[canonical] = val_raw
                        matched = True; break
                if not matched: continue
            continue
            
        norm = re.sub(r'[^A-Z0-9]', '', label_raw.upper())
        if len(norm) < 2: continue
        canonical = FIELD_MAP.get(norm, norm)
        
        # Multiline value check
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
    return d

def get_val(d, keys, clean_numeric=False):
    for k in keys:
        if k in d:
            v = d[k]; return _extract_number(v) if clean_numeric else v
        canon = FIELD_MAP.get(re.sub(r'[^A-Z0-9]', '', k.upper()), k)
        if canon in d:
            v = d[canon]; return _extract_number(v) if clean_numeric else v
    return 0 if clean_numeric else ""

# Splitting Regexes
_TS_SPLIT  = re.compile(r'\n(?=\d{1,2}/\d{1,2}/\d{2,4},?\s*\d{1,2}:\d{2}(?:[\s\u202f]*[APap][Mm])?\s*[-\u2013]\s*)')
_TS_META   = re.compile(r'^(\d{1,2}/\d{1,2}/\d{2,4}),?\s*\d{1,2}:\d{2}(?:[\s\u202f]*[APap][Mm])?\s*[-\u2013]\s*(.*?)(?:\s*:|\s*$)', re.IGNORECASE)
_ORDER_SPLIT = re.compile(r'\n\s*Order(?:(?:\s+in\s+case)s?)?\s*:?\s*\n(?:\s*Stock\s*\n)?', re.IGNORECASE)
_STOCK_SPLIT = re.compile(r'\n\s*(?:Physical\s+)?Stock(?:\s+in\s+hand)?\s*:?\s*\n', re.IGNORECASE)
_HAS_OUTLET  = re.compile(r"\b(?:outlet\s*name|outletname|आउटलेट\s*नाम|shop\s*name|store\s*name)\b", re.IGNORECASE)

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
        if not entry or not _HAS_OUTLET.search(entry): continue
        
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

        row = {
            "Date": msg_date,
            "HQ":   get_val(master_d, ["HEADQUARTER"]),
            "STAFF NAME": get_val(master_d, ["INCHARGE"]),
            "ROUTE": get_val(master_d, ["ROUTE"]),
            "DAY":  datetime.strptime(msg_date, "%d/%m/%Y").strftime("%A") if "/" in msg_date else "",
            "OUTLET NAME": get_val(master_d, ["OUTLETNAME"]),
            "ADDRESS": get_val(master_d, ["ADDRESS"]),
            "OWNER NAME": get_val(master_d, ["OWNERNAME"]),
            "CONTACT NO": get_val(master_d, ["CONTACTNO"]),
            "OUTLET TYPE": get_val(master_d, ["OUTLETTYPE"]),
            "CHANNEL": get_val(master_d, ["CHANNEL"]),
            "PAN NO": get_val(master_d, ["PANNO"]),
            "VOLUME": get_val(master_d, ["VOLUME"]),
            "W/S TIE-UP": get_val(master_d, ["WSTIEUP"]),
            "MBQ STATUS": get_val(master_d, ["MBQ"]),
            "TASTING": get_val(master_d, ["TASTING"]),
            "TSHIRT": get_val(master_d, ["TSHIRT"]),
            "POSTER": get_val(master_d, ["POSTER_COUNT"]),
            "CHITDRAW": get_val(master_d, ["CHITDRAW"]),
            "CAN BLOCKOUT": get_val(master_d, ["CANBLOCKOUT"]),
            "GSB BOARD": get_val(master_d, ["GSB_BOARD"]),
            "GSB SIZE": get_val(master_d, ["GSB_SIZE"]),
            "DPS (RAW)": get_val(master_d, ["DPS"]),
            "Indoor Branding": get_val(master_d, ["DISPLAY"]),
            "DPS": "", 
            "Other Brand": get_val(master_d, ["OTHERS"]),
            "DPS REMARK": "",
            
            # PHYSICAL STOCK
            "S_MT 250": get_val(stock_d, ["MT250"], clean_numeric=False),
            "S_MT 330": get_val(stock_d, ["MT330"], clean_numeric=False),
            "S_XT 330": get_val(stock_d, ["XT330"], clean_numeric=False),
            "S_JJ 200ML": get_val(stock_d, ["JJ200ML"], clean_numeric=False),
            "S_JJ 320ML": get_val(stock_d, ["JJ320ML"], clean_numeric=False),
            "S_BB NIMBOO": get_val(stock_d, ["BAMBAMNIMBOO"], clean_numeric=False),
            "S_BB MIST": get_val(stock_d, ["BAMBAMMIST"], clean_numeric=False),
            "S_BB COLA": get_val(stock_d, ["BAMBAMCOLA"], clean_numeric=False),
            "S_BB HYDRATION": get_val(stock_d, ["BAMBAMHYDRATION"], clean_numeric=False),
            "S_CSD": get_val(stock_d, ["CSD"], clean_numeric=False),
            "S_BB 2.5L": get_val(stock_d, ["BAMBAM25L"], clean_numeric=False),
            "S_MT CLASSIC": get_val(stock_d, ["MTCLASSIC"], clean_numeric=False),
            "S_OTHERS": get_val(stock_d, ["OTHERS"], clean_numeric=False),

            # ORDER
            "O_MT 250": get_val(order_d, ["MT250"], clean_numeric=False),
            "O_MT 330": get_val(order_d, ["MT330"], clean_numeric=False),
            "O_XT 330": get_val(order_d, ["XT330"], clean_numeric=False),
            "O_JJ 200ML": get_val(order_d, ["JJ200ML"], clean_numeric=False),
            "O_JJ 320ML": get_val(order_d, ["JJ320ML"], clean_numeric=False),
            "O_BB NIMBOO": get_val(order_d, ["BAMBAMNIMBOO"], clean_numeric=False),
            "O_BB MIST": get_val(order_d, ["BAMBAMMIST"], clean_numeric=False),
            "O_BB COLA": get_val(order_d, ["BAMBAMCOLA"], clean_numeric=False),
            "O_BB HYDRATION": get_val(order_d, ["BAMBAMHYDRATION"], clean_numeric=False),
            "O_CSD": get_val(order_d, ["CSD"], clean_numeric=False),
            "O_BB 2.5L": get_val(order_d, ["BAMBAM25L"], clean_numeric=False),
            "O_MT CLASSIC": get_val(order_d, ["MTCLASSIC"], clean_numeric=False),
            "O_OTHERS": get_val(order_d, ["OTHERS"], clean_numeric=False),
            
            # DISPATCH (Placeholders)
            "D_MT 250": 0, "D_MT 330": 0, "D_XT 330": 0, "D_JJ 200ML": 0, "D_JJ 320ML": 0,
            "D_BB NIMBOO": 0, "D_BB MIST": 0, "D_BB COLA": 0, "D_BB HYDRATION": 0, 
            "D_CSD": 0, "D_BB 2.5L": 0, "D_MT CLASSIC": 0, "D_OTHERS": 0,
            "STAFF REMARK": get_val(master_d, ["REMARKS"]),
        }
        data.append(row)
    return pd.DataFrame(data)

# --- Nepali Digit Normalizer ---
nepali_to_english_digits = str.maketrans("०१२३४५६७८९", "0123456789")

def normalize_text(text):
    if not text: return ""
    return text.translate(nepali_to_english_digits)

def clean_val(val):
    if val:
        val = val.strip()
        if val in ["-", "", "null", ".", "N/A", "na"]:
            return None
    return val

# --- Poster Extraction Patterns ---
POSTER_FIELD_PATTERNS = {
    "outlet_name": [r"outlet\s*name", r"outlets?\s*name", r"आउटलेट\s*नाम"],
    "address": [r"address", r"addedss", r"ठेगाना", r"addr"],
    "contact": [r"contact(?:\s*no\.?)?", r"contract", r"mobile\s*no\.?", r"सम्पर्क", r"फोन", r"ph(?:one)?"],
    "owner_name": [r"owner\s*name", r"owoner\s*name", r"मालिक"],
    "staff_name": [r"staff\s*name", r"name", r"नाम"],
    "area": [r"route\s*name", r"route", r"area", r"hq"],
    "xt_poster": [r"no\.?\s*of\s*xtreme\s*posters?\s*pasted", r"count\s*of\s*xtreme\s*poster", r"xtreme\s*poster", r"extreme\s*poster", r"xt\s*poster", r"count\s*of\s*xt"],
    "mt250_poster": [r"no\.?\s*of\s*mt\s*250\s*posters?\s*pasted", r"count\s*of\s*max\s*tiger\s*poster\s*250", r"mt\s*250\s*poster", r"mt250\s*poster", r"max\s*tiger\s*poster\s*250"],
    "mt330_poster": [r"no\.?\s*of\s*mt\s*330\s*posters?\s*pasted", r"count\s*of\s*max\s*tiger\s*330", r"mt\s*330\s*poster", r"mt330\s*poster", r"max\s*tiger\s*330"],
    "bam_bam": [r"bam\s*bam", r"bambam"],
    "dps_board": [r"dps[/\s]*gsb\s*board", r"dps\s*board", r"gsb\s*board", r"board"]
}

# --- Staff Name Mapping ---
KNOWN_SENDERS = {
    "+977 980-3539493": "Rupesh Nepali",
    "+977 980-6645229": "Anush Kunwar",
    "+977 970-6015375": "Bhuban Dauliya",
    "+977 981-8811222": "Roshan Shrestha",
    "+977 974-2515634": "Aryan Gurung",
    "+977 981-5174483": "Staff (Unknown)",
}

def extract_poster_field(block, key):
    for pat in POSTER_FIELD_PATTERNS[key]:
        m = re.search(rf"(?:^|\n)[ \t]*{pat}[ \t]*[:\-=.]+[ \t]*(.*)", block, re.IGNORECASE | re.MULTILINE)
        if m:
            val = m.group(1).split("\n")[0]
            return clean_val(normalize_text(val))
    return None

def process_poster_file(filepath, filter_date=None):
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            text = f.read()
    except Exception as e:
        raise Exception(f"Error reading file: {e}")

    # Dynamic Staff Name discovery
    sender_name_pattern = r"\d{2}/\d{2}/\d{4}, \d{2}:\d{2} - (\+977 [\d-]+): (?:Staff name\s*[:.-]\s*|NAME\s*[-:]\s*)([^\n\r]+)"
    for sender, name in re.findall(sender_name_pattern, text, re.IGNORECASE):
        n = name.strip()
        if n and len(n) < 50 and sender not in KNOWN_SENDERS:
            KNOWN_SENDERS[sender] = n

    entries = re.split(r'\n(?=\d{1,2}/\d{1,2}/\d{2,4}, \d{1,2}:\d{2} - )', text)
    data = []
    for entry in entries:
        if "OUTLET NAME" not in entry.upper() and "आउटलेट नाम" not in entry: continue
        
        # Meta extraction
        meta_match = re.search(r'^(\d{1,2}/\d{1,2}/\d{2,4}), \d{1,2}:\d{2} - (.*?):', entry)
        if not meta_match: continue
        msg_date, sender_raw = meta_match.groups()
        
        if filter_date and msg_date != filter_date: continue
        
        # Resolve Staff Name
        staff = KNOWN_SENDERS.get(sender_raw.strip())
        staff_in_body = extract_poster_field(entry, "staff_name")
        if staff_in_body and len(staff_in_body.split()) <= 4:
            staff = staff_in_body

        rec = {
            "date": msg_date,
            "staff_name": staff or sender_raw,
            "outlet_name": extract_poster_field(entry, "outlet_name"),
            "address": extract_poster_field(entry, "address"),
            "contact": extract_poster_field(entry, "contact"),
            "owner_name": extract_poster_field(entry, "owner_name"),
            "xt_poster": extract_poster_field(entry, "xt_poster"),
            "mt250_poster": extract_poster_field(entry, "mt250_poster"),
            "mt330_poster": extract_poster_field(entry, "mt330_poster"),
            "bam_bam": extract_poster_field(entry, "bam_bam"),
            "dps_board": extract_poster_field(entry, "dps_board"),
        }
        if rec["outlet_name"]:
            data.append(rec)
    return pd.DataFrame(data)

def apply_poster_styling(template_path, output_path, df, hq_name):
    wb = load_workbook(template_path)
    
    # Rename first sheet to HQ name and delete others
    target_sheet = wb.active
    target_sheet.title = hq_name
    for sheet in wb.sheetnames:
        if sheet != hq_name:
            del wb[sheet]
            
    ws = wb[hq_name]

    # Delete Column D (Area) as requested
    ws.delete_cols(4) # Column D is index 4

    # Find first empty row
    first_data_row = 2
    for row in ws.iter_rows(min_row=2, values_only=True):
        if all(c is None for c in row[3:]): # Check from Staff Name onward
            break
        first_data_row += 1

    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for i, row_data in df.iterrows():
        row_num = first_data_row + i
        sn = row_num - 1
        
        # Mapping col to data (A=SN, B=Date, C=HQ, D=Staff Name, E=Outlet, F=Address, G=Contact, H=Owner...)
        mapping = {
            "A": sn,
            "B": row_data["date"],
            "C": hq_name,
            "D": row_data["staff_name"] or "",
            "E": row_data["outlet_name"] or "",
            "F": row_data["address"] or "",
            "G": row_data["contact"] or "",
            "H": row_data["owner_name"] or "",
            "I": row_data["xt_poster"],
            "J": row_data["mt250_poster"],
            "K": row_data["mt330_poster"],
            "L": row_data["bam_bam"],
            "M": row_data["dps_board"] or ""
        }
        
        for col_letter, val in mapping.items():
            cell = ws[f"{col_letter}{row_num}"]
            cell.value = val
            cell.border = border
            cell.alignment = Alignment(horizontal="left", vertical="center")
            cell.font = Font(name="Arial", size=10)

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

    # Header merge ranges
    worksheet.merge_range(0, 0, 0, 26, f"OUTLET DETAILS - {report_type_label}", main_header_fmt)
    worksheet.merge_range(0, 27, 0, 39, "PHYSICAL STOCK", stock_header_fmt)
    worksheet.merge_range(0, 40, 0, 52, "ORDER", order_header_fmt)
    worksheet.merge_range(0, 53, 0, 65, "DISPATCH", dispatch_header_fmt)

    headers = [
        "Date", "HQ", "STAFF NAME", "ROUTE", "DAY", "OUTLET NAME", "ADDRESS", 
        "OWNER NAME", "CONTACT NO", "Outlet Type", "Channel", "PAN No", "Volume", 
        "W/S Tie-up", "MBQ Status", "Tasting", "T-Shirt", "Poster", "ChitDraw", 
        "CAN Blockout", "GSB Board", "GSB Size", "DPS (RAW)", "Indoor Branding", 
        "DPS", "Other Brand", "DPS REMARK"
    ]
    products = [
        "MT 250", "MT 330", "XT 330", "JJ 200ML", "JJ 320ML", "BB NIMBOO", 
        "BB MIST", "BB COLA", "BB HYDRATION", "CSD", "BB 2.5L", "MT CLASSIC", "OTHERS"
    ]
    
    full_headers = headers + products + products + products + ["STAFF REMARK"]
    for col_num, header in enumerate(full_headers):
        fmt = main_header_fmt if col_num < len(headers) or col_num == len(full_headers)-1 else sub_header_fmt
        worksheet.write(1, col_num, header, fmt)

    for row_idx in range(len(df)):
        for col_idx in range(len(df.columns)):
            val = df.iloc[row_idx, col_idx]
            is_numeric = len(headers) <= col_idx < (len(df.columns) - 1)
            
            if pd.isna(val) or val == 0 or val == "0" or val == "":
                worksheet.write_blank(row_idx + 2, col_idx, None, num_fmt if is_numeric else cell_fmt)
            else:
                fmt = num_fmt if is_numeric else cell_fmt
                worksheet.write(row_idx + 2, col_idx, val, fmt)

        # Dropdowns
        worksheet.data_validation(row_idx + 2, 24, row_idx + 2, 24, {'validate': 'list', 'source': ['YES', 'NO']})

    # Conditional Formatting
    worksheet.conditional_format(2, 24, len(df) + 1, 24, {'type': 'cell', 'criteria': 'equal to', 'value': '"YES"', 'format': green_fmt})
    worksheet.conditional_format(2, 24, len(df) + 1, 24, {'type': 'cell', 'criteria': 'equal to', 'value': '"NO"', 'format': red_fmt})

    # Column Widths
    worksheet.set_column('A:B', 12)
    worksheet.set_column('C:D', 20)
    worksheet.set_column('F:H', 25)
    worksheet.set_column('I:Z', 15)
    worksheet.set_column('AA:AU', 10)
    worksheet.set_column('AV:BS', 8)
    worksheet.set_column('BT:BT', 30)

# --- Modern UI Application ---

import sys
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
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

        self.title("Agro DSR Pro v6.0")
        self.geometry("850x600") # Slightly wider for the logo
        
        # Set Window Icon
        try:
            logo_path = resource_path(os.path.join("assets", "xtreme.png"))
            if os.path.exists(logo_path):
                img = Image.open(logo_path)
                # Resize for icon (32x32 is typical for taskbar, but wm_iconphoto takes image)
                self.wm_iconphoto(True, ImageTk.PhotoImage(img))
        except Exception as e:
            print(f"Icon error: {e}")

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.sidebar = ctk.CTkFrame(self, width=220, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        
        # Logo in Sidebar
        try:
            pil_img = Image.open(resource_path(os.path.join("assets", "xtreme.png")))
            logo_img = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=(160, 280))
            self.logo_label = ctk.CTkLabel(self.sidebar, image=logo_img, text="")
            self.logo_label.pack(pady=30)
        except Exception as e:
            print(f"Logo error: {e}")
            self.logo_label = ctk.CTkLabel(self.sidebar, text="AGRO DSR\nPRO", font=ctk.CTkFont(size=24, weight="bold"))
            self.logo_label.pack(pady=40)
        
        self.mode_label = ctk.CTkLabel(self.sidebar, text="Extraction Mode", anchor="w")
        self.mode_label.pack(fill="x", padx=20, pady=(20, 5))
        
        self.mode_selector = ctk.CTkSegmentedButton(self.sidebar, values=["Today Only", "Full History"])
        self.mode_selector.set("Today Only")
        self.mode_selector.pack(padx=10, fill="x")

        self.theme_label = ctk.CTkLabel(self.sidebar, text="Appearance Mode", anchor="w")
        self.theme_label.pack(fill="x", padx=20, pady=(100, 5))
        self.theme_option = ctk.CTkOptionMenu(self.sidebar, values=["Dark", "Light"], command=self.change_appearance_mode)
        self.theme_option.pack(padx=10, fill="x")

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

        gen_btn = ctk.CTkButton(self.sales_tab, text="GENERATE SALES REPORT", font=ctk.CTkFont(size=14, weight="bold"),
                               height=50, command=self.run_sales_process)
        gen_btn.pack(side="bottom", fill="x", padx=40, pady=30)

    def setup_poster_tab(self):
        title = ctk.CTkLabel(self.poster_tab, text="Poster Dedicated Report", font=ctk.CTkFont(size=20, weight="bold"))
        title.pack(pady=(20, 30))

        # Chat File
        chat_label = ctk.CTkLabel(self.poster_tab, text="Poster Chat Export (.txt):", anchor="w")
        chat_label.pack(fill="x", padx=30)
        chat_frame = ctk.CTkFrame(self.poster_tab, fg_color="transparent")
        chat_frame.pack(fill="x", padx=30, pady=(0, 20))
        
        self.poster_chat_var = ctk.StringVar(value="No chat file selected...")
        entry = ctk.CTkEntry(chat_frame, textvariable=self.poster_chat_var, width=300, height=40)
        entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        btn = ctk.CTkButton(chat_frame, text="Browse", width=100, height=40, command=lambda: self.browse_file(self.poster_chat_var, [("Text Files", "*.txt")]))
        btn.pack(side="right")

        # Info about default template
        info_label = ctk.CTkLabel(self.poster_tab, text="Using Default Bhaktapur Template", text_color="gray", font=ctk.CTkFont(size=12, slant="italic"))
        info_label.pack(pady=(0, 20))

        self.poster_status = ctk.CTkLabel(self.poster_tab, text="Ready", text_color="gray")
        self.poster_status.pack(pady=(40, 5))
        
        self.poster_progress = ctk.CTkProgressBar(self.poster_tab, width=400)
        self.poster_progress.set(0)
        self.poster_progress.pack(pady=10)

        gen_btn = ctk.CTkButton(self.poster_tab, text="GENERATE POSTER REPORT", font=ctk.CTkFont(size=14, weight="bold"),
                               height=50, command=self.run_poster_process)
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
                messagebox.showwarning("Empty Result", "No reports found.")
                return
            
            self.sales_progress.set(0.6)
            output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                      initialfile=f"DSR_Sales_{datetime.now().strftime('%Y%m%d')}.xlsx",
                                                      filetypes=[("Excel Files", "*.xlsx")])
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
        
        # Discovery HQ from filename
        filename = os.path.basename(chat_path)
        hq_match = re.search(r"with\s+(.*?)\s+POSTER", filename, re.IGNORECASE)
        hq_name = hq_match.group(1) if hq_match else "POSTER"
        
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
            output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                      initialfile=f"Poster_{hq_name}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                                                      filetypes=[("Excel Files", "*.xlsx")])
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
