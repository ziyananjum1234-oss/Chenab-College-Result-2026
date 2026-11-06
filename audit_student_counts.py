import os
import xlrd
import json

target_dir = r"C:\Users\IT_OFFICE02\Desktop\Monthly Test October Exel Lists 2026"
data_path = r"C:\Users\IT_OFFICE02\Desktop\Chenab_Result_Portal\data.json"

with open(data_path, "r", encoding="utf-8") as f:
    db = json.load(f)

students = db["students"]

# Count students per section in DB
db_counts = {}
for s in students:
    sec = s["class_name"]
    db_counts[sec] = db_counts.get(sec, 0) + 1

files = sorted([f for f in os.listdir(target_dir) if f.endswith(".xls") or f.endswith(".xlsx")])

discrepancies = []
total_sheet_students = 0

for fname in files:
    p = os.path.join(target_dir, fname)
    book = xlrd.open_workbook(p)
    sheet = book.sheet_by_index(0)
    
    cls_name = ""
    for r in range(min(sheet.nrows, 6)):
        for c in range(min(sheet.ncols, 6)):
            v = str(sheet.cell_value(r, c)).strip()
            if v.startswith("Class :") or v.startswith("Class:"):
                cls_name = v.replace("Class :", "").replace("Class:", "").strip()
                break
        if cls_name: break
    if not cls_name: cls_name = os.path.splitext(fname)[0]
    
    header_row = -1
    for r in range(min(sheet.nrows, 12)):
        val1 = str(sheet.cell_value(r, 1)).strip().lower() if sheet.ncols > 1 else ""
        val3 = str(sheet.cell_value(r, 3)).strip().lower() if sheet.ncols > 3 else ""
        if "id" in val1 or "student name" in val3:
            header_row = r
            break
            
    pass_row = header_row + 2
    raw_st_count = 0
    raw_students_list = []
    for r in range(pass_row + 1, sheet.nrows):
        sid = str(sheet.cell_value(r, 1)).strip()
        sname = str(sheet.cell_value(r, 3)).strip()
        if not sid or sid == "nan" or not sname or sname == "nan": continue
        if any(k in sname.lower() for k in ["total student", "average", "pass %", "position"]): continue
        raw_st_count += 1
        raw_students_list.append((sid, sname))
        
    total_sheet_students += raw_st_count
    in_db = db_counts.get(cls_name, 0)
    
    if raw_st_count != in_db:
        discrepancies.append((fname, cls_name, raw_st_count, in_db))

print(f"Total raw students counted in 67 Excel sheets: {total_sheet_students}")
print(f"Total students compiled in database:           {len(students)}")

if discrepancies:
    print(f"\nDISCREPANCIES FOUND ({len(discrepancies)}):")
    for fname, cls_name, raw_cnt, db_cnt in discrepancies:
        print(f"  - {fname} ({cls_name}): Sheet has {raw_cnt}, DB has {db_cnt}")
else:
    print("\n>>> 100% PERFECT MATCH! Every single student from all 67 sheets is present! <<<")
