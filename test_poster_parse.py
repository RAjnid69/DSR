"""Test poster parsing against all chat files to identify gaps."""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
from dsr_pro import process_poster_file

CHAT_DIR = "poster_Message"

for fname in os.listdir(CHAT_DIR):
    if not fname.endswith(".txt"):
        continue
    fpath = os.path.join(CHAT_DIR, fname)
    short = fname.replace("WhatsApp Chat with ", "").replace(".txt", "")
    try:
        df = process_poster_file(fpath, filter_date=None)
        print(f"\n{'='*60}")
        print(f"FILE: {short}")
        print(f"Records found: {len(df)}")
        if not df.empty:
            cols = df.columns.tolist()
            for i, row in df.iterrows():
                outlet = row.get("outlet_name", "?")
                staff = row.get("staff_name", "?")
                xt = row.get("xt_poster", "")
                mt = row.get("mt250_poster", "")
                addr = row.get("address", "")
                contact = row.get("contact", "")
                print(f"  [{i+1:3d}] outlet={outlet!s:40s} staff={staff!s:25s} xt={xt!s:5s} mt={mt!s:5s} addr={addr!s:20s} contact={contact!s}")
    except Exception as e:
        print(f"\nERROR in {short}: {e}")
