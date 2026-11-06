import os
import xlrd
import json
import re

target_dir = r"C:\Users\IT_OFFICE02\Desktop\Monthly Test October Exel Lists 2026"
files = sorted(os.listdir(target_dir))

SUBJ_NAMES = {
    "ENG": "English",
    "URDU": "Urdu",
    "MATHS": "Mathematics",
    "MATH": "Mathematics",
    "SCI": "General Science",
    "SCIENCE": "General Science",
    "SST": "Social Studies",
    "ISL": "Islamiat",
    "Q.E": "Quranic Education",
    "COMP": "Computer Science",
    "BIO": "Biology",
    "CHEM": "Chemistry",
    "PHY": "Physics",
    "PAK S": "Pakistan Studies",
    "ARABI": "Arabic",
    "EDU": "Education",
    "PE": "Physical Education",
    "CIVIC": "Civics",
    "HIS": "History",
    "GEO": "Geography"
}

def clean_subj_name(raw):
    raw_u = raw.strip().upper()
    return SUBJ_NAMES.get(raw_u, raw.strip())

def get_grade(pct):
    if pct >= 90: return "A+"
    if pct >= 80: return "A"
    if pct >= 70: return "B"
    if pct >= 60: return "C"
    if pct >= 50: return "D"
    return "F"

def get_remarks(pct):
    if pct >= 90: return "Outstanding performance! Keep maintaining this exemplary standard."
    if pct >= 80: return "Excellent achievement! High academic dedication demonstrated."
    if pct >= 70: return "Very good result. With focused effort, can achieve top grades."
    if pct >= 60: return "Good effort. Consistent revision recommended for improvement."
    if pct >= 50: return "Satisfactory progress. Needs to devote more time to studies."
    return "Needs immediate academic attention and remedial guidance."

def get_section_head(cls_name, fname):
    cls = cls_name.upper()
    fn = fname.upper()
    
    # 1. Hifz
    if "HIFZ" in cls or "HIFA" in cls or "HIFZ" in fn or "HIFA" in fn:
        if "(G)" in cls or "GIRL" in cls or "GIRL" in fn:
            return "Ms. Sajida Nawaz", "Section Head (Hifz Girls Wing)", "sec_sajida_4_8_g"
        else:
            return "Mr. Abrar Hussain", "Section Head (Hifz Boys Wing)", "sec_abbas_4_8_b"
            
    # 2. O Levels & Pre-O
    if "O LEVEL" in cls or "PRE O" in cls or "O LEVEL" in fn or "PRE O" in fn:
        return "Ms. Sajida Nawaz", "Section Head (Cambridge Wing)", "sec_sajida_4_8_g"
        
    # 3. 1st to 3rd
    if re.search(r'\b1\b|\b2\b|\b3\b|1ST|2ND|3RD', cls) or re.search(r'^[123]\s', fn):
        if not re.search(r'10|11|12', cls) and not re.search(r'10|11|12', fn):
            return "Ms. Farhat Naz", "Section Head (Junior Wing 1st - 3rd)", "sec_farhat_1_3"
            
    # 4. 4th to 8th
    if re.search(r'\b[4-8]\b|[4-8]TH', cls) or re.search(r'^[4-8]TH', fn):
        is_girls = "(G)" in cls or "GIRL" in cls or any(g in fn.lower() for g in ["goldies", "silver", "dimonds", "sphiier", "pearls", "rubies", "corals", "jades", "amber", "opeals"])
        if is_girls:
            return "Ms. Sajida Nawaz", "Section Head (Girls Wing 4th - 8th)", "sec_sajida_4_8_g"
        else:
            return "Mr. Abrar Hussain", "Section Head (Boys Wing 4th - 8th)", "sec_abbas_4_8_b"
            
    # 5. 9th to 12th
    if re.search(r'\b9\b|\b10\b|\b11\b|\b12\b|9TH|10TH|11TH|12TH', cls) or re.search(r'^(9|10|11|12)TH', fn):
        is_girls = "(G)" in cls or "GIRL" in cls or any(g in fn.lower() for g in ["courage", "excellance", "wisdom", "amina", "ayesha", "fatima", "hajra", "hafsa", "khadija", "maryam", "zainab", "displain", "faith", "unity"])
        if is_girls:
            return "Ms. Samreena Nawaz", "Section Head (Girls Wing 9th - 12th)", "sec_sambeen_9_12_g"
        else:
            return "Mr. Shoukat Hayat", "Section Head (Boys Wing 9th - 12th)", "sec_shoukat_9_12_b"
            
    return "Ms. Sajida Nawaz", "Section Head", "sec_sajida_4_8_g"

all_students = []
all_sections = set()

for f_idx, fname in enumerate(files):
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
    
    all_sections.add(cls_name)
    head_name, head_title, head_sig = get_section_head(cls_name, fname)
    
    header_row = -1
    for r in range(min(sheet.nrows, 12)):
        val1 = str(sheet.cell_value(r, 1)).strip().lower() if sheet.ncols > 1 else ""
        val3 = str(sheet.cell_value(r, 3)).strip().lower() if sheet.ncols > 3 else ""
        if "id" in val1 or "student name" in val3:
            header_row = r
            break
            
    r_cols = [str(sheet.cell_value(header_row, c)).strip().replace("\n", " ") for c in range(sheet.ncols)]
    
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
            
    if ttl_idx != -1 and obt_idx == -1 and ttl_idx + 1 < sheet.ncols:
        obt_idx = ttl_idx + 1
        
    max_row = header_row + 1
    pass_row = header_row + 2
    
    subj_end = ttl_idx if ttl_idx != -1 else sheet.ncols
    subject_defs = []
    for c in range(7, subj_end):
        raw_sname = r_cols[c]
        if not raw_sname or raw_sname == "nan":
            continue
        try: max_m = float(sheet.cell_value(max_row, c))
        except: max_m = 25.0
        try: pass_m = float(sheet.cell_value(pass_row, c))
        except: pass_m = 10.0
            
        subject_defs.append({
            "col": c,
            "code": raw_sname,
            "name": clean_subj_name(raw_sname),
            "max": max_m,
            "pass": pass_m
        })
        
    student_start_row = pass_row + 1
    for r in range(student_start_row, sheet.nrows):
        sid_raw = sheet.cell_value(r, 1) if sheet.ncols > 1 else ""
        sname_raw = sheet.cell_value(r, 3) if sheet.ncols > 3 else ""
        
        if isinstance(sid_raw, float):
            sid = f"{int(sid_raw)}"
        else:
            sid = str(sid_raw).strip()
            if sid.endswith(".0"): sid = sid[:-2]
                
        sname = str(sname_raw).strip()
        if not sid or sid == "nan" or not sname or sname == "nan":
            continue
        if any(k in sname.lower() for k in ["total student", "average", "pass %", "position"]):
            continue
            
        cpos_raw = sheet.cell_value(r, 4) if sheet.ncols > 4 else "-"
        spos_raw = sheet.cell_value(r, 6) if sheet.ncols > 6 else "-"
        try:
            cpos = f"{int(float(cpos_raw))}" if float(cpos_raw) > 0 else "-"
        except:
            cpos = str(cpos_raw).strip()
            if cpos.endswith(".0"): cpos = cpos[:-2]
            if not cpos or cpos == "nan": cpos = "-"
            
        try:
            spos = f"{int(float(spos_raw))}" if float(spos_raw) > 0 else "-"
        except:
            spos = str(spos_raw).strip()
            if spos.endswith(".0"): spos = spos[:-2]
            if not spos or spos == "nan": spos = "-"
            
        # Parse subjects for this student
        st_subjects = []
        for sdef in subject_defs:
            c_idx = sdef["col"]
            val_raw = sheet.cell_value(r, c_idx)
            val_str = str(val_raw).strip()
            
            # If blank / empty -> skip this subject (especially for 9th-12th electives and exemptions)
            if val_str == "" or val_str.lower() == "nan":
                continue
                
            max_m = sdef["max"]
            pass_m = sdef["pass"]
            
            is_leave = val_str in ["(L)", "L", "Leave", "LEAVE"]
            is_absent = val_str in ["(A)", "A", "Absent", "ABSENT"]
            
            if is_leave:
                sub_pct = "-"
                sub_grade = "Leave"
                sub_status = "Leave"
                sub_obt = "Leave"
            elif is_absent:
                sub_pct = "0.0%"
                sub_grade = "F"
                sub_status = "Fail"
                sub_obt = "Absent"
            else:
                try:
                    obt_f = float(val_str)
                    pct_calc = round((obt_f / max_m) * 100, 1) if max_m > 0 else 0
                    sub_pct = f"{pct_calc:.1f}%"
                    sub_grade = get_grade(pct_calc)
                    sub_status = "Pass" if obt_f >= pass_m else "Fail"
                    sub_obt = f"{obt_f:g}"
                except:
                    sub_pct = "-"
                    sub_grade = "-"
                    sub_status = "N/A"
                    sub_obt = val_str
                    
            st_subjects.append({
                "name": sdef["name"],
                "max": int(max_m) if max_m.is_integer() else max_m,
                "pass": int(pass_m) if pass_m.is_integer() else pass_m,
                "obtained": sub_obt,
                "percentage": sub_pct,
                "grade": sub_grade,
                "status": sub_status
            })
            
        # Summary marks
        ttl_val = sheet.cell_value(r, ttl_idx) if ttl_idx != -1 and sheet.ncols > ttl_idx else ""
        obt_val = sheet.cell_value(r, obt_idx) if obt_idx != -1 and sheet.ncols > obt_idx else ""
        pct_val = sheet.cell_value(r, pct_idx) if pct_idx != -1 and sheet.ncols > pct_idx else ""
        grd_val = sheet.cell_value(r, grd_idx) if grd_idx != -1 and sheet.ncols > grd_idx else ""
        sts_val = sheet.cell_value(r, sts_idx) if sts_idx != -1 and sheet.ncols > sts_idx else ""
        
        try:
            ttl_f = float(ttl_val)
            ttl_mark = f"{int(ttl_f)}" if ttl_f.is_integer() else f"{ttl_f:g}"
        except:
            ttl_mark = str(ttl_val).strip()
            
        try:
            obt_f = float(obt_val)
            obt_mark = f"{int(obt_f)}" if obt_f.is_integer() else f"{obt_f:g}"
        except:
            obt_mark = str(obt_val).strip()
            
        try:
            pct_f = float(pct_val)
            pct_str = f"{pct_f:.1f}"
        except:
            pct_str = str(pct_val).strip().replace("%", "")
            
        try:
            pct_num = float(pct_str)
            calc_grade = get_grade(pct_num)
            calc_remarks = get_remarks(pct_num)
        except:
            pct_num = 0
            calc_grade = "-"
            calc_remarks = "Academic records verified by examination authority."
            
        grade = str(grd_val).strip() if grd_val and str(grd_val).strip() != "nan" else calc_grade
        status = str(sts_val).strip() if sts_val and str(sts_val).strip() != "nan" else ("Pass" if pct_num >= 50 else "Fail")
        
        all_students.append({
            "id": sid,
            "name": sname,
            "class_name": cls_name,
            "class_pos": cpos,
            "sec_pos": spos,
            "subjects": st_subjects,
            "total_marks": ttl_mark,
            "obtained_marks": obt_mark,
            "percentage": pct_str,
            "grade": grade,
            "remarks": calc_remarks,
            "status": status,
            "sec_head_name": head_name,
            "sec_head_title": head_title,
            "sec_head_sig": head_sig
        })

print("=" * 80)
print(f"Total students parsed: {len(all_students)}")
print(f"Total sections count: {len(all_sections)}")
print(f"Sample student 1: {all_students[0]['name']} ({all_students[0]['class_name']}) -> {all_students[0]['sec_head_name']}")
print(f"Sample student 500: {all_students[500]['name']} ({all_students[500]['class_name']}) -> {all_students[500]['sec_head_name']}")
print(f"Sample student 1000: {all_students[1000]['name']} ({all_students[1000]['class_name']}) -> {all_students[1000]['sec_head_name']}")
print(f"Sample student 2000: {all_students[2000]['name']} ({all_students[2000]['class_name']}) -> {all_students[2000]['sec_head_name']}")
