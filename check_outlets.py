import pandas as pd
import dsr_pro

df = dsr_pro.process_single_file("WhatsApp Chat with AGRO-BTL PHOTOS.txt")
mistakes = df[pd.isna(df['OUTLET NAME']) | (df['OUTLET NAME'] == '')]
row = df.iloc[1090]
print("ROW 1090 DATA:", row.to_dict())

entries = dsr_pro._TS_SPLIT.split("\n" + dsr_pro.normalize_text(open("WhatsApp Chat with AGRO-BTL PHOTOS.txt", encoding="utf-8").read()))
count = -1
for entry in entries:
    if not entry or not dsr_pro._HAS_OUTLET.search(entry): continue
    meta_match = dsr_pro._TS_META.search(entry)
    if not meta_match: continue
    count += 1
    if count == 1090:
        print("ENTRY 1090 TEXT:")
        print(repr(entry))
        break

