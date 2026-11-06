import os
import xlrd

target_dir = r"C:\Users\IT_OFFICE02\Desktop\Monthly Test October Exel Lists 2026"
files = sorted(os.listdir(target_dir))

all_raw_subjects = set()

for fname in files:
    p = os.path.join(target_dir, fname)
    book = xlrd.open_workbook(p)
    sheet = book.sheet_by_index(0)
    
    header_row = -1
    for r in range(min(sheet.nrows, 12)):
        val1 = str(sheet.cell_value(r, 1)).strip().lower() if sheet.ncols > 1 else ""
        val3 = str(sheet.cell_value(r, 3)).strip().lower() if sheet.ncols > 3 else ""
        if "id" in val1 or "student name" in val3:
            header_row = r
            break
            
    if header_row == -1: continue
    r_cols = [str(sheet.cell_value(header_row, c)).strip().replace("\n", " ") for c in range(sheet.ncols)]
    
    ttl_idx = -1
    for c, col in enumerate(r_cols):
        if any(k in col.lower() for k in ["ttl mark", "total mark", "total"]):
            ttl_idx = c
            break
            
    subj_end = ttl_idx if ttl_idx != -1 else sheet.ncols
    for c in range(7, subj_end):
        raw_s = r_cols[c].strip()
        if raw_s and raw_s != "nan":
            all_raw_subjects.add(raw_s)

print("ALL UNIQUE RAW SUBJECT CODES ACROSS ALL 65 FILES:")
for s in sorted(all_raw_subjects):
    print(f"  - '{s}'")
