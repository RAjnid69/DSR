import sys
import pandas as pd
import re

from dsr_pro import _SEP_RE, FIELD_MAP, normalize_text, get_val, _extract_number, _TS_SPLIT, _TS_META, _ORDER_SPLIT, _STOCK_SPLIT, _HAS_OUTLET

def fixed_block_to_dict(block):
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

import dsr_pro
dsr_pro.block_to_dict = fixed_block_to_dict

df = dsr_pro.process_single_file("WhatsApp Chat with AGRO-BTL PHOTOS.txt")
mistakes = df[(df["S_BAM BAM 250ML"] == 500) | (df["O_BAM BAM 250ML"] == 500)]
if mistakes.empty:
    print("SUCCESS: 0 mistakes found!")
else:
    for idx, row in mistakes.iterrows():
        print("MISTAKE:", row["OUTLET NAME"], row["S_BAM BAM 250ML"], row["O_BAM BAM 250ML"])
