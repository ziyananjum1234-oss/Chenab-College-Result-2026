import os
import xlrd

target_dir = r"C:\Users\IT_OFFICE02\Desktop\Monthly Test October Exel Lists 2026"
files = sorted(os.listdir(target_dir))

for f in files:
    p = os.path.join(target_dir, f)
    book = xlrd.open_workbook(p)
    sheet = book.sheet_by_index(0)
    
    # Class name
    cls_name = ""
    for r in range(min(sheet.nrows, 6)):
        for c in range(min(sheet.ncols, 6)):
            v = str(sheet.cell_value(r, c)).strip()
            if v.startswith("Class :"):
                cls_name = v
                break
        if cls_name:
            break
            
    # Row 7 columns
    r7_cols = [str(sheet.cell_value(7, c)).strip().replace("\n", " ") for c in range(sheet.ncols)]
    
    # Find where 'Ttl Mark' or 'Total' is
    ttl_idx = -1
    for c, col_name in enumerate(r7_cols):
        if any(k in col_name.lower() for k in ["ttl mark", "total mark", "total"]):
            ttl_idx = c
            break
            
    subjects = r7_cols[7:ttl_idx] if ttl_idx != -1 else r7_cols[7:]
    summary_cols = r7_cols[ttl_idx:] if ttl_idx != -1 else []
    
    print(f"{f:28s} | {cls_name:32s} | Subjs ({len(subjects)}): {', '.join(subjects)} | Summary: {', '.join(summary_cols[:5])}")
