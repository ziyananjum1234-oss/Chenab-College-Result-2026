import os
import xlrd
import re

target_dir = r"C:\Users\IT_OFFICE02\Desktop\Monthly Test October Exel Lists 2026"
files = sorted(os.listdir(target_dir))

def get_grade_group(cls_name):
    c = cls_name.upper()
    if "HIFZ" in c or "HIFA" in c: return "Hifz-ul-Quran"
    if "O LEVEL" in c or "PRE O" in c: return "Cambridge Wing"
    if "12" in c: return "Class 12th"
    if "11" in c: return "Class 11th"
    if "10" in c: return "Class 10th"
    if "9" in c: return "Class 9th"
    if "8" in c: return "Class 8th"
    if "7" in c: return "Class 7th"
    if "6" in c: return "Class 6th"
    if "5" in c: return "Class 5th"
    if "4" in c: return "Class 4th"
    if "3" in c: return "Class 3rd"
    if "2" in c: return "Class 2nd"
    if "1" in c: return "Class 1st"
    return "Other"

groups = {}
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
    
    grp = get_grade_group(cls_name)
    groups.setdefault(grp, []).append(cls_name)

print("GRADE GROUPS AND SECTIONS:")
for grp, secs in sorted(groups.items()):
    print(f"\n{grp} ({len(secs)} sections):")
    for s in secs:
        print(f"   - {s}")
