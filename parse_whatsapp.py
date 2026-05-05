import re
import pandas as pd
import os
import glob
from datetime import datetime

# Configuration
input_folders = [".", "chat export", "c:/Users/NITRO/Desktop/automation/chat export"]
output_file = "DSR_Today_Final_Report.xlsx"
today_date = datetime.now().strftime("%d/%m/%Y")

def get_val(d, keys, clean_numeric=False):
    """Looks up multiple possible keys in a dictionary and returns the first found value."""
    for k in keys:
        if k in d:
            val = d[k]
            if clean_numeric:
                # Extract digits from strings like "3pcs", "10 c", "5crt", "10ctn"
                digits = re.search(r'(\d+)', val)
                if digits:
                    return int(digits.group(1))
                return 0 # Return 0 for sum calculation if empty/no digits
            return val
    return 0 if clean_numeric else ""

def block_to_dict(block):
    """Parses a text block into a dictionary of normalized keys."""
    d = {}
    for line in block.split('\n'):
        separator = None
        if ':' in line:
            separator = ':'
        elif '-' in line and any(k in line.upper() for k in ["HEADQUARTER", "ROUTE", "INCHARGE", "DEALER", "OUTLET", "ADDRESS", "OWNER", "CONTACT", "VOLUME"]):
            separator = '-'
        
        if separator:
            lbl, val = line.split(separator, 1)
            norm_lbl = re.sub(r'[^A-Z0-9]', '', lbl.upper())
            val = val.strip().lstrip(':-').strip()
            if norm_lbl and (norm_lbl not in d or not d[norm_lbl]):
                d[norm_lbl] = val
    return d

def process_file(filepath):
    print(f"Processing {filepath}...")
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            text = f.read()
    except Exception as e:
        print(f"Error reading {filepath}: {e}")
        return None

    entries = re.split(r'\n(?=\d{1,2}/\d{1,2}/\d{2,4}, \d{1,2}:\d{2} - )', text)

    data = []
    for entry in entries:
        if "OUTLET NAME" not in entry.upper():
            continue

        date_match = re.search(r'^(\d{1,2}/\d{1,2}/\d{2,4}),', entry)
        if not date_match: continue
        msg_date = date_match.group(1)
        if msg_date != today_date: continue

        order_parts = re.split(r'ORDER(?: IN CASE)?', entry, flags=re.IGNORECASE)
        stock_block = order_parts[0]
        order_block = order_parts[1] if len(order_parts) > 1 else ""

        stock_parts = re.split(r'(?:PHYSICAL )?STOCK:', stock_block, flags=re.IGNORECASE)
        info_block = stock_parts[0]
        actual_stock_block = stock_parts[1] if len(stock_parts) > 1 else ""

        stock_d = block_to_dict(actual_stock_block)
        order_d = block_to_dict(order_block)
        master_d = block_to_dict(entry)

        row = {
            "Date": msg_date,
            "HQ": get_val(master_d, ["HEADQUARTER"]),
            "STAFF NAME": get_val(master_d, ["INCHARGENAME", "INCHARGE"]),
            "ROUTE": get_val(master_d, ["ROUTEPOINT", "ROUTE"]),
            "DAY": datetime.now().strftime("%A"), # Derived from current date
            "OUTLET NAME": get_val(master_d, ["OUTLETNAME", "OUTLET"]),
            "ADDRESS": get_val(master_d, ["ADDRESS"]),
            "OUTLET OWNER NAME": get_val(master_d, ["OWNERNAME", "OWNER"]),
            "CONTACT NO": get_val(master_d, ["CONTACTNO", "CONTACT"]),
            "DPS": get_val(master_d, ["DPS"]),
            
            # PHYSICAL STOCK
            "S_MT 250": get_val(stock_d, ["MT250", "MAXTIGER250"], clean_numeric=True),
            "S_MT 330": get_val(stock_d, ["MT330", "MAXTIGER330"], clean_numeric=True),
            "S_XT 330": get_val(stock_d, ["XT330", "XTREME330"], clean_numeric=True),
            "S_JJ 200ML": get_val(stock_d, ["JUICEJELLY200ML", "JJ200ML"], clean_numeric=True),
            "S_JJ 320ML": get_val(stock_d, ["JUICEJELLY320ML", "JJ320ML"], clean_numeric=True),
            "S_BAM BAM 250ML": get_val(stock_d, ["BAMBAM250ML"], clean_numeric=True),
            "S_BAM BAM 500 ML": get_val(stock_d, ["BAMBAM500ML"], clean_numeric=True),
            "S_BAM BAM 2.5 ML": get_val(stock_d, ["BAMBAM25ML"], clean_numeric=True),
            "S_CSD": get_val(stock_d, ["CSD"], clean_numeric=True),
            "S_MT CLASSIC": get_val(stock_d, ["MT330CLASSIC", "MTCLASSIC"], clean_numeric=True),
            "S_OTHERS": get_val(stock_d, ["OTHERS", "LIQUOR"], clean_numeric=True),

            # ORDER
            "O_MT 250": get_val(order_d, ["MT250", "MAXTIGER250"], clean_numeric=True),
            "O_MT 330": get_val(order_d, ["MT330", "MAXTIGER330"], clean_numeric=True),
            "O_XT 330": get_val(order_d, ["XT330", "XTREME330"], clean_numeric=True),
            "O_JJ 200ML": get_val(order_d, ["JUICEJELLY200ML", "JJ200ML"], clean_numeric=True),
            "O_JJ 320ML": get_val(order_d, ["JUICEJELLY320ML", "JJ320ML"], clean_numeric=True),
            "O_BAM BAM 250ML": get_val(order_d, ["BAMBAM250ML"], clean_numeric=True),
            "O_BAM BAM 500 ML": get_val(order_d, ["BAMBAM500ML"], clean_numeric=True),
            "O_BAM BAM 2.5 ML": get_val(order_d, ["BAMBAM25ML"], clean_numeric=True),
            "O_CSD": get_val(order_d, ["CSD"], clean_numeric=True),
            "O_MT CLASSIC": get_val(order_d, ["MT330CLASSIC", "MTCLASSIC"], clean_numeric=True),
            "O_OTHERS": get_val(order_d, ["OTHERS", "LIQUOR"], clean_numeric=True),
            
            "REMARKS": get_val(master_d, ["REMARKS"])
        }
        data.append(row)

    return pd.DataFrame(data)

def apply_styling(writer, df, sheet_name):
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    
    # Define formats
    header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#D7E4BC', 'border': 1})
    group_header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#92D050', 'border': 1})
    title_format = workbook.add_format({'bold': True, 'size': 16, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#C6E0B4', 'border': 1})
    cell_format = workbook.add_format({'border': 1, 'valign': 'vcenter'})
    numeric_format = workbook.add_format({'border': 1, 'align': 'center'})
    total_format = workbook.add_format({'bold': True, 'fg_color': '#FCE4D6', 'border': 1, 'align': 'center'})

    # Write Title
    worksheet.merge_range(0, 0, 0, 32, f"DAILY SALES REPORT - {today_date}", title_format)

    # Write Group Headers
    worksheet.merge_range(1, 10, 1, 20, "PHYSICAL STOCK", group_header_format)
    worksheet.merge_range(1, 21, 1, 31, "ORDER", group_header_format)

    # Clean display names for columns (removing S_ and O_ prefixes)
    display_cols = []
    for c in df.columns:
        if c.startswith("S_") or c.startswith("O_"):
            display_cols.append(c[2:])
        else:
            display_cols.append(c)

    # Write Column Headers
    for col_num, value in enumerate(display_cols):
        worksheet.write(2, col_num, value, header_format)

    # Write Data
    for row_num in range(len(df)):
        for col_num in range(len(df.columns)):
            val = df.iloc[row_num, col_num]
            fmt = numeric_format if col_num >= 10 and col_num <= 31 else cell_format
            worksheet.write(row_num + 3, col_num, val, fmt)

    # Write Grand Totals
    total_row_idx = len(df) + 3
    worksheet.write(total_row_idx, 9, "GRAND TOTAL", total_format)
    for col_num in range(10, 32):
        col_letter = chr(ord('A') + col_num) if col_num < 26 else 'A' + chr(ord('A') + col_num - 26)
        # Handle Excel column letters for > 26 columns
        import xlsxwriter.utility as util
        col_str = util.xl_col_to_name(col_num)
        formula = f"=SUM({col_str}4:{col_str}{total_row_idx})"
        worksheet.write_formula(total_row_idx, col_num, formula, total_format)

    # Column Widths
    worksheet.set_column('A:B', 12) # Date, HQ
    worksheet.set_column('C:D', 20) # Staff, Route
    worksheet.set_column('F:F', 25) # Outlet
    worksheet.set_column('G:I', 15) # Address, Owner, Contact
    worksheet.set_column('K:AF', 8) # Product columns
    worksheet.set_column('AG:AG', 30) # Remarks

def main():
    files_to_process = []
    for folder in input_folders:
        if not os.path.exists(folder): continue
        for f in glob.glob(os.path.join(folder, "*.txt")):
            files_to_process.append(f)

    if not files_to_process:
        print("No files found.")
        return

    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        processed_any = False
        for filepath in files_to_process:
            df = process_file(filepath)
            if df is not None and not df.empty:
                sheet_name = os.path.basename(filepath).replace("WhatsApp Chat with ", "")[:31]
                sheet_name = re.sub(r'[\\/*?:\[\]]', '', sheet_name)
                df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=2)
                apply_styling(writer, df, sheet_name)
                processed_any = True
        
        if not processed_any:
            print(f"No reports found for {today_date}")
            return

    print(f"\nSuccessfully generated Final DSR report: {output_file}")

if __name__ == "__main__":
    main()
