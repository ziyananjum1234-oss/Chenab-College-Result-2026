import os
import xlrd
import re

target_dir = r"C:\Users\IT_OFFICE02\Desktop\Monthly Test October Exel Lists 2026"
files = sorted(os.listdir(target_dir))

total_students_all = 0
all_parsed_data = []
errors = []

# Subject name mapping for clean display
SUBJ_NAMES = {
    "ENG": "English",
    "ENGLISH": "English",
    "URDU": "Urdu",
    "MATHS": "Mathematics",
    "MATH": "Mathematics",
    "MATHEMATICS": "Mathematics",
    "SCI": "General Science",
    "SCIENCE": "General Science",
    "G.SCI": "General Science",
    "SST": "Social Studies",
    "S.ST": "Social Studies",
    "ISL": "Islamiat",
    "ISLAMIAT": "Islamiat",
    "Q.E": "Quranic Education",
    "QE": "Quranic Education",
    "QURAN": "Quranic Education",
    "COMP": "Computer Science",
    "COMPUTER": "Computer Science",
    "BIO": "Biology",
    "BIOLOGY": "Biology",
    "CHEM": "Chemistry",
    "CHEMISTRY": "Chemistry",
    "PHY": "Physics",
    "PHYSICS": "Physics",
    "PAK S": "Pakistan Studies",
    "PAK.S": "Pakistan Studies",
    "PAK STUDIES": "Pakistan Studies",
    "ARABI": "Arabic",
    "ARABIC": "Arabic",
    "EDU": "Education",
    "EDUCATION": "Education",
    "PE": "Physical Education",
    "P.E": "Physical Education",
    "CIVIC": "Civics",
    "CIVICS": "Civics",
    "HIS": "History",
    "HISTORY": "History",
    "GEO": "Geography",
    "GEOGRAPHY": "Geography"
}

def clean_subj_name(raw):
    raw_u = raw.strip().upper()
    return SUBJ_NAMES.get(raw_u, raw.strip())

for f_idx, fname in enumerate(files):
    p = os.path.join(target_dir, fname)
    try:
        book = xlrd.open_workbook(p)
        sheet = book.sheet_by_index(0)
        
        # 1. Class Name
        cls_name = ""
        for r in range(min(sheet.nrows, 6)):
            for c in range(min(sheet.ncols, 6)):
                v = str(sheet.cell_value(r, c)).strip()
                if v.startswith("Class :") or v.startswith("Class:"):
                    cls_name = v.replace("Class :", "").replace("Class:", "").strip()
                    break
            if cls_name:
                break
        if not cls_name:
            cls_name = os.path.splitext(fname)[0]
            
        # 2. Locate header row (where 'Id' is in col 1 or 'Student Name' in col 3)
        header_row = -1
        for r in range(min(sheet.nrows, 12)):
            val1 = str(sheet.cell_value(r, 1)).strip().lower() if sheet.ncols > 1 else ""
            val3 = str(sheet.cell_value(r, 3)).strip().lower() if sheet.ncols > 3 else ""
            if "id" in val1 or "student name" in val3:
                header_row = r
                break
                
        if header_row == -1:
            errors.append(f"{fname}: Header row not found!")
            continue
            
        r_cols = [str(sheet.cell_value(header_row, c)).strip().replace("\n", " ") for c in range(sheet.ncols)]
        
        # Locate summary columns
        ttl_idx = -1
        obt_idx = -1
        pct_idx = -1
        grd_idx = -1
        sts_idx = -1
        
        for c, col in enumerate(r_cols):
            c_low = col.lower()
            if ("ttl mark" in c_low or "total mark" in c_low or c_low == "total") and ttl_idx == -1:
                ttl_idx = c
            elif ("obt" in c_low) and obt_idx == -1:
                obt_idx = c
            elif ("%" in c_low or "age" in c_low or "pct" in c_low) and pct_idx == -1:
                pct_idx = c
            elif ("grade" in c_low) and grd_idx == -1:
                grd_idx = c
            elif ("status" in c_low) and sts_idx == -1:
                sts_idx = c
                
        # If pct_idx or grd_idx not found by name, check adjacent
        if ttl_idx != -1 and obt_idx == -1 and ttl_idx + 1 < sheet.ncols:
            obt_idx = ttl_idx + 1
            
        # Max marks row & Pass marks row
        max_row = header_row + 1
        pass_row = header_row + 2
        
        # Subject column range: from 7 up to ttl_idx
        subj_end = ttl_idx if ttl_idx != -1 else sheet.ncols
        subject_defs = []
        for c in range(7, subj_end):
            raw_sname = r_cols[c]
            if not raw_sname or raw_sname == "nan":
                continue
            # Read max and pass marks
            try:
                max_m = float(sheet.cell_value(max_row, c))
            except:
                max_m = 25.0
            try:
                pass_m = float(sheet.cell_value(pass_row, c))
            except:
                pass_m = 10.0
                
            subject_defs.append({
                "col": c,
                "code": raw_sname,
                "name": clean_subj_name(raw_sname),
                "max": max_m,
                "pass": pass_m
            })
            
        # Students start at pass_row + 1
        student_start_row = pass_row + 1
        st_count = 0
        for r in range(student_start_row, sheet.nrows):
            sid_raw = sheet.cell_value(r, 1) if sheet.ncols > 1 else ""
            sname_raw = sheet.cell_value(r, 3) if sheet.ncols > 3 else ""
            
            # Convert sid to string clean
            if isinstance(sid_raw, float):
                sid = f"{int(sid_raw)}"
            else:
                sid = str(sid_raw).strip()
                if sid.endswith(".0"):
                    sid = sid[:-2]
                    
            sname = str(sname_raw).strip()
            
            # Skip empty or footer rows
            if not sid or sid == "nan" or not sname or sname == "nan":
                continue
            # Also ignore if sname is summary like 'Average', 'Total Students', etc.
            if any(k in sname.lower() for k in ["total student", "average", "pass %", "position"]):
                continue
                
            # Class Pos and Sec Pos
            cpos_raw = sheet.cell_value(r, 4) if sheet.ncols > 4 else "-"
            spos_raw = sheet.cell_value(r, 6) if sheet.ncols > 6 else "-"
            
            cpos = f"{int(float(cpos_raw))}" if isinstance(cpos_raw, (int, float)) and cpos_raw > 0 else str(cpos_raw).strip()
            spos = f"{int(float(spos_raw))}" if isinstance(spos_raw, (int, float)) and spos_raw > 0 else str(spos_raw).strip()
            if cpos.endswith(".0"): cpos = cpos[:-2]
            if spos.endswith(".0"): spos = spos[:-2]
            if not cpos or cpos == "nan": cpos = "-"
            if not spos or spos == "nan": spos = "-"
            
            # Summary marks
            ttl_val = sheet.cell_value(r, ttl_idx) if ttl_idx != -1 and sheet.ncols > ttl_idx else ""
            obt_val = sheet.cell_value(r, obt_idx) if obt_idx != -1 and sheet.ncols > obt_idx else ""
            pct_val = sheet.cell_value(r, pct_idx) if pct_idx != -1 and sheet.ncols > pct_idx else ""
            grd_val = sheet.cell_value(r, grd_idx) if grd_idx != -1 and sheet.ncols > grd_idx else ""
            sts_val = sheet.cell_value(r, sts_idx) if sts_idx != -1 and sheet.ncols > sts_idx else ""
            
            st_count += 1
            total_students_all += 1
            
        print(f"[{f_idx+1:02d}/65] {fname:28s} | Class: {cls_name:30s} | Subjs: {len(subject_defs)} | Students: {st_count}")
        
    except Exception as e:
        errors.append(f"{fname}: {str(e)}")

print("=" * 80)
print(f"TOTAL STUDENTS FOUND ACROSS ALL 65 SHEETS: {total_students_all}")
if errors:
    print(f"ERRORS ({len(errors)}):")
    for err in errors:
        print("  -", err)
else:
    print("NO ERRORS! All 65 files successfully read!")
