import pandas as pd
import dsr_pro

# Process the chat file
try:
    df = dsr_pro.process_single_file("WhatsApp Chat with AGRO-BTL PHOTOS.txt")
except Exception as e:
    print(f"Error parsing file: {e}")
    exit(1)

mistakes_found = 0

def log_mistake(row_idx, row, issue):
    global mistakes_found
    mistakes_found += 1
    print(f"Row {row_idx}: {issue} (Outlet: {row.get('OUTLET NAME', 'N/A')}, Staff: {row.get('STAFF NAME', 'N/A')})")

numeric_cols = [c for c in df.columns if c.startswith('S_') or c.startswith('O_')]

for idx, row in df.iterrows():
    # 1. Missing Outlet Name
    if pd.isna(row.get('OUTLET NAME')) or not str(row.get('OUTLET NAME')).strip():
        log_mistake(idx, row, "Missing OUTLET NAME")
    
    # 2. Missing Staff Name
    if pd.isna(row.get('STAFF NAME')) or not str(row.get('STAFF NAME')).strip():
        log_mistake(idx, row, "Missing STAFF NAME")
        
    # 3. Missing HQ
    if pd.isna(row.get('HQ')) or not str(row.get('HQ')).strip():
        log_mistake(idx, row, "Missing HQ")
        
    # 4. Outrageously high numeric values (likely parsing error like the 500ml issue)
    for col in numeric_cols:
        val = row.get(col)
        if pd.notna(val) and isinstance(val, (int, float)):
            if val > 100:  # 100 cases is a lot, let's flag > 100 to check for things like "250ml" being parsed as 250
                # Exclude Contact No/Volume etc. We only check stock/order.
                # Oh wait, 250 or 330 might get extracted if they write "250ml" instead of the value
                log_mistake(idx, row, f"Suspiciously high value in {col}: {val}")

    # 5. Check if contact number contains alphabetic characters
    contact = str(row.get('CONTACT NO', ''))
    if contact and contact.lower() not in ['nan', 'none']:
        import re
        if re.search('[a-zA-Z]', contact):
            log_mistake(idx, row, f"Alphabetic characters in CONTACT NO: {contact}")

if mistakes_found == 0:
    print("No obvious mistakes found in heuristic scan!")
else:
    print(f"\nTotal potential mistakes flagged: {mistakes_found}")
