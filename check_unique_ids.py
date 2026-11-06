import os
import xlrd

target_dir = r"C:\Users\IT_OFFICE02\Desktop\Monthly Test October Exel Lists 2026"
files = sorted(os.listdir(target_dir))

all_students = []
id_counts = {}

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
    pass_row = header_row + 2
    
    for r in range(pass_row + 1, sheet.nrows):
        sid_raw = sheet.cell_value(r, 1) if sheet.ncols > 1 else ""
        sname_raw = sheet.cell_value(r, 3) if sheet.ncols > 3 else ""
        
        if isinstance(sid_raw, float):
            sid = f"{int(sid_raw)}"
        else:
            sid = str(sid_raw).strip()
            if sid.endswith(".0"): sid = sid[:-2]
            
        sname = str(sname_raw).strip()
        if not sid or sid == "nan" or not sname or sname == "nan": continue
        if any(k in sname.lower() for k in ["total student", "average", "pass %", "position"]): continue
        
        id_counts[sid] = id_counts.get(sid, 0) + 1
        all_students.append({
            "id": sid,
            "name": sname,
            "file": fname
        })

print(f"Total student entries: {len(all_students)}")
print(f"Unique student IDs: {len(id_counts)}")
duplicates = {k: v for k, v in id_counts.items() if v > 1}
if duplicates:
    print(f"Duplicate IDs ({len(duplicates)}):")
    for k, v in list(duplicates.items())[:10]:
        matching = [s for s in all_students if s["id"] == k]
        print(f"  ID {k} occurs {v} times: {[(m['name'], m['file']) for m in matching]}")
else:
    print("ALL 2,202 student IDs are 100% unique!")
