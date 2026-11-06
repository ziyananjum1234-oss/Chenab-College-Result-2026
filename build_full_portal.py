import os
import sys
import base64
import json
import xlrd
import re

desktop = r"C:\Users\IT_OFFICE02\Desktop"
portal_folder = os.path.join(desktop, "Chenab_Result_Portal")
excel_dir = os.path.join(desktop, "Monthly Test October Exel Lists 2026")
os.makedirs(portal_folder, exist_ok=True)

# 1. Embed College Logo
logo_path = os.path.join(desktop, "chenab_logo.jpg")
logo_b64 = ""
if os.path.exists(logo_path):
    with open(logo_path, "rb") as f:
        logo_b64 = "data:image/jpeg;base64," + base64.b64encode(f.read()).decode("utf-8")

# 2. Embed EXACT "Old English Text MT" Font from Windows
font_b64 = ""
font_path = r"C:\Windows\Fonts\OLDENGL.TTF"
if os.path.exists(font_path):
    with open(font_path, "rb") as f:
        font_b64 = base64.b64encode(f.read()).decode("utf-8")

# 3. Embed Digital Signatures
sig_dir = os.path.join(portal_folder, "signatures")
sig_map = {
    "exam_head_sibtain": "exam_head_sibtain.png",
    "sec_shoukat_9_12_b": "sec_shoukat_9_12_b.png",
    "sec_sambeen_9_12_g": "sec_sambeen_9_12_g.png",
    "sec_sajida_4_8_g": "sec_sajida_4_8_g.png",
    "sec_abbas_4_8_b": "sec_abbas_4_8_b.png",
    "sec_farhat_1_3": "sec_farhat_1_3.png"
}
signatures_b64 = {}
for k, fname in sig_map.items():
    p = os.path.join(sig_dir, fname)
    if os.path.exists(p):
        with open(p, "rb") as sf:
            signatures_b64[k] = "data:image/png;base64," + base64.b64encode(sf.read()).decode("utf-8")
    else:
        signatures_b64[k] = ""

signatures_json_str = json.dumps(signatures_b64)

# Subject Names mapping
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

# Parse all 65 files
files = sorted([f for f in os.listdir(excel_dir) if f.endswith(".xls") or f.endswith(".xlsx")])
print(f"Reading {len(files)} files...")

students = []
sections_set = set()
grade_groups_set = set()

for f_idx, fname in enumerate(files):
    p = os.path.join(excel_dir, fname)
    book = xlrd.open_workbook(p)
    sheet = book.sheet_by_index(0)
    
    # Class Name
    cls_name = ""
    for r in range(min(sheet.nrows, 6)):
        for c in range(min(sheet.ncols, 6)):
            v = str(sheet.cell_value(r, c)).strip()
            if v.startswith("Class :") or v.startswith("Class:"):
                cls_name = v.replace("Class :", "").replace("Class:", "").strip()
                break
        if cls_name: break
    if not cls_name: cls_name = os.path.splitext(fname)[0]
    
    sections_set.add(cls_name)
    grade_grp = get_grade_group(cls_name)
    grade_groups_set.add(grade_grp)
    
    head_name, head_title, head_sig = get_section_head(cls_name, fname)
    
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
            
        st_subjects = []
        for sdef in subject_defs:
            c_idx = sdef["col"]
            val_raw = sheet.cell_value(r, c_idx)
            val_str = str(val_raw).strip()
            
            # Skip empty columns
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
        
        students.append({
            "id": sid,
            "name": sname,
            "class_name": cls_name,
            "grade_group": grade_grp,
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

sections_list = sorted(list(sections_set))
# Order grade groups naturally
grade_order = [
    "Class 1st", "Class 2nd", "Class 3rd", "Class 4th", "Class 5th",
    "Class 6th", "Class 7th", "Class 8th", "Class 9th", "Class 10th",
    "Class 11th", "Class 12th", "Hifz-ul-Quran", "Cambridge Wing"
]
grade_groups_list = [g for g in grade_order if g in grade_groups_set] + [g for g in sorted(grade_groups_set) if g not in grade_order]

portal_data = {
    "institute": "Chenab College Shorkot",
    "test_name": "Monthly Test October 2026",
    "session": "Academic Session 2026-2027",
    "exam_date": "October 2026",
    "total_students": len(students),
    "sections": sections_list,
    "grade_groups": grade_groups_list,
    "students": students
}

with open(os.path.join(portal_folder, "data.json"), "w", encoding="utf-8") as f:
    json.dump(portal_data, f, ensure_ascii=False, indent=2)

json_str = json.dumps(portal_data)

print(f"Total students: {len(students)} | Sections: {len(sections_list)} | Grade Groups: {len(grade_groups_list)}")

# Now generate full portal HTML
html_template = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Online Result Portal &bull; Chenab College Shorkot</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Cinzel:wght@600;700;800;900&family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=Playfair+Display:ital,wght@0,600;1,600&display=swap" rel="stylesheet">
    <style>
        @font-face {
            font-family: 'Old English Custom';
            src: url('data:font/truetype;charset=utf-8;base64,__FONT_B64__') format('truetype');
            font-weight: normal;
            font-style: normal;
        }

        :root {
            --navy-deep: #081d38;
            --navy-card: #0c274a;
            --navy-light: #163660;
            --gold: #c69222;
            --gold-light: #dfb249;
            --gold-subtle: #fcf8ee;
            --gold-border: #e6cb85;
            --slate-bg: #f3f6fa;
            --text-dark: #0f172a;
            --text-muted: #52637a;
            --border-soft: #e2e8f0;
            --success-green: #0d824d;
            --success-bg: #eafaf1;
            --danger-red: #c02626;
            --danger-bg: #fdf2f2;
            --leave-orange: #b45309;
            --leave-bg: #fffbeb;
        }

        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }

        body {
            font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            background-color: var(--slate-bg);
            color: var(--text-dark);
            min-height: 100vh;
            display: flex;
            flex-direction: column;
            background-image: 
                radial-gradient(circle at 10% 10%, rgba(198, 146, 34, 0.04) 0%, transparent 40%),
                radial-gradient(circle at 90% 90%, rgba(8, 29, 56, 0.05) 0%, transparent 40%);
            background-attachment: fixed;
        }

        .wrapper {
            width: 100%;
            max-width: 1160px;
            margin: 0 auto;
            padding: 24px 20px 40px;
            flex: 1;
            display: flex;
            flex-direction: column;
        }

        /* Top Executive Header */
        .executive-header {
            background: linear-gradient(135deg, var(--navy-deep) 0%, #0d2a52 50%, var(--navy-deep) 100%);
            border-radius: 18px;
            padding: 24px 30px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            box-shadow: 0 10px 25px -5px rgba(8, 29, 56, 0.25), 0 0 0 1px rgba(198, 146, 34, 0.2);
            color: #ffffff;
            margin-bottom: 24px;
            position: relative;
            overflow: hidden;
        }

        .executive-header::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 3px;
            background: linear-gradient(90deg, transparent, var(--gold), transparent);
        }

        .header-brand {
            display: flex;
            align-items: center;
            gap: 20px;
        }

        .crest-wrapper {
            width: 68px;
            height: 68px;
            background: #ffffff;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            box-shadow: 0 0 0 3px rgba(198, 146, 34, 0.4), 0 8px 16px rgba(0, 0, 0, 0.3);
            flex-shrink: 0;
            overflow: hidden;
            padding: 2px;
        }

        .crest-wrapper img {
            width: 100%;
            height: 100%;
            object-fit: contain;
            border-radius: 50%;
        }

        .header-title-block h1 {
            font-family: 'Old English Custom', 'Old English Text MT', 'Cinzel', serif;
            font-size: 2.15rem;
            letter-spacing: 0.5px;
            margin-bottom: 3px;
            background: linear-gradient(180deg, #ffffff 60%, #e2e8f0 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }

        .header-title-block .sub-brand {
            font-size: 0.88rem;
            color: #dfb249;
            font-weight: 600;
            letter-spacing: 1px;
            text-transform: uppercase;
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .header-badge {
            background: rgba(255, 255, 255, 0.08);
            border: 1px solid rgba(198, 146, 34, 0.35);
            padding: 8px 16px;
            border-radius: 30px;
            font-size: 0.82rem;
            font-weight: 600;
            color: #f1f5f9;
            display: flex;
            align-items: center;
            gap: 8px;
            backdrop-filter: blur(8px);
        }

        .header-badge .badge-dot {
            width: 8px;
            height: 8px;
            background-color: #22c55e;
            border-radius: 50%;
            box-shadow: 0 0 8px #22c55e;
        }

        /* Search Section */
        .search-container {
            background: #ffffff;
            border-radius: 16px;
            padding: 30px;
            box-shadow: 0 4px 20px -2px rgba(8, 29, 56, 0.06), 0 0 0 1px var(--border-soft);
            margin-bottom: 24px;
            position: relative;
        }

        .search-title-area {
            text-align: center;
            margin-bottom: 20px;
        }

        .search-title-area h2 {
            font-size: 1.35rem;
            color: var(--navy-deep);
            font-weight: 800;
            margin-bottom: 4px;
        }

        .search-title-area p {
            color: var(--text-muted);
            font-size: 0.9rem;
        }

        .search-form {
            display: flex;
            gap: 12px;
            max-width: 640px;
            margin: 0 auto;
            position: relative;
        }

        .input-group-custom {
            flex: 1;
            position: relative;
        }

        .input-group-custom input {
            width: 100%;
            height: 52px;
            padding: 0 16px 0 48px;
            border: 2px solid #e2e8f0;
            border-radius: 12px;
            font-size: 1rem;
            font-family: inherit;
            color: var(--text-dark);
            outline: none;
            transition: all 0.2s ease;
            background: #fafcff;
        }

        .input-group-custom input:focus {
            border-color: var(--gold);
            background: #ffffff;
            box-shadow: 0 0 0 4px rgba(198, 146, 34, 0.12);
        }

        .input-icon {
            position: absolute;
            left: 16px;
            top: 50%;
            transform: translateY(-50%);
            color: #94a3b8;
            font-size: 1.1rem;
            pointer-events: none;
        }

        .btn-search {
            height: 52px;
            padding: 0 28px;
            background: linear-gradient(135deg, var(--navy-deep) 0%, var(--navy-light) 100%);
            color: #ffffff;
            border: none;
            border-radius: 12px;
            font-size: 0.95rem;
            font-weight: 700;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 8px;
            box-shadow: 0 4px 12px rgba(8, 29, 56, 0.2);
            transition: all 0.2s ease;
            flex-shrink: 0;
        }

        .btn-search:hover {
            transform: translateY(-1px);
            box-shadow: 0 6px 16px rgba(8, 29, 56, 0.25);
            background: linear-gradient(135deg, #0a2446 0%, #1a3f6e 100%);
        }

        .btn-honor-roll {
            height: 52px;
            padding: 0 20px;
            background: linear-gradient(135deg, #b8860b 0%, #d4af37 100%);
            color: #ffffff;
            border: none;
            border-radius: 12px;
            font-size: 0.92rem;
            font-weight: 700;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 8px;
            box-shadow: 0 4px 12px rgba(184, 134, 11, 0.25);
            transition: all 0.2s ease;
            flex-shrink: 0;
        }

        .btn-honor-roll:hover {
            transform: translateY(-1px);
            box-shadow: 0 6px 16px rgba(184, 134, 11, 0.35);
            background: linear-gradient(135deg, #a67909 0%, #c49e29 100%);
        }

        /* Autocomplete Suggestions */
        .suggestions-box {
            position: absolute;
            top: 58px;
            left: 0;
            right: 0;
            background: #ffffff;
            border-radius: 12px;
            border: 1px solid var(--border-soft);
            box-shadow: 0 12px 28px -4px rgba(8, 29, 56, 0.15);
            z-index: 50;
            max-height: 280px;
            overflow-y: auto;
            display: none;
        }

        .suggestion-item {
            padding: 12px 18px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            border-bottom: 1px solid #f1f5f9;
            cursor: pointer;
            transition: background 0.15s ease;
        }

        .suggestion-item:last-child {
            border-bottom: none;
        }

        .suggestion-item:hover {
            background-color: #f8fafc;
        }

        .sug-name {
            font-weight: 700;
            color: var(--navy-deep);
            font-size: 0.92rem;
        }

        .sug-meta {
            font-size: 0.8rem;
            color: var(--text-muted);
        }

        .sug-badge {
            background: #e2e8f0;
            color: var(--navy-deep);
            padding: 3px 8px;
            border-radius: 6px;
            font-size: 0.72rem;
            font-weight: 700;
        }

        /* Error Alert */
        .error-alert {
            max-width: 640px;
            margin: 14px auto 0;
            padding: 12px 16px;
            background-color: var(--danger-bg);
            border: 1px solid rgba(192, 38, 38, 0.3);
            border-radius: 10px;
            color: var(--danger-red);
            font-size: 0.88rem;
            font-weight: 600;
            display: none;
            align-items: center;
            gap: 10px;
        }

        /* Result View & DMC Container */
        .result-view {
            display: none;
            animation: fadeIn 0.3s ease-out;
        }

        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(8px); }
            to { opacity: 1; transform: translateY(0); }
        }

        .dmc-card {
            background: #ffffff;
            border-radius: 16px;
            box-shadow: 0 10px 30px -5px rgba(8, 29, 56, 0.1), 0 0 0 1px var(--border-soft);
            padding: 32px 36px;
            margin-bottom: 24px;
            position: relative;
            background-color: #ffffff;
        }

        /* Watermark Background */
        .dmc-card::after {
            content: '';
            position: absolute;
            top: 50%;
            left: 50%;
            width: 320px;
            height: 320px;
            transform: translate(-50%, -50%);
            background-image: url('__LOGO_B64__');
            background-size: contain;
            background-repeat: no-repeat;
            background-position: center;
            opacity: 0.035;
            pointer-events: none;
            z-index: 0;
        }

        /* DMC Header */
        .dmc-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding-bottom: 14px;
            border-bottom: 2px solid var(--navy-deep);
            margin-bottom: 16px;
            position: relative;
            z-index: 1;
        }

        .dmc-logo-left {
            width: 72px;
            height: 72px;
            flex-shrink: 0;
        }

        .dmc-logo-left img {
            width: 100%;
            height: 100%;
            object-fit: contain;
        }

        .dmc-header-center {
            text-align: center;
            flex: 1;
            padding: 0 15px;
        }

        .dmc-inst-name {
            font-family: 'Old English Custom', 'Old English Text MT', serif;
            font-size: 2.1rem;
            color: var(--navy-deep);
            margin-bottom: 2px;
            line-height: 1.15;
        }

        .dmc-cert-label {
            font-family: 'Cinzel', serif;
            font-size: 1.05rem;
            font-weight: 800;
            color: var(--gold);
            letter-spacing: 2px;
            text-transform: uppercase;
            margin-bottom: 3px;
        }

        .dmc-exam-term {
            font-size: 0.95rem;
            font-weight: 700;
            color: var(--navy-card);
            letter-spacing: 0.5px;
        }

        .dmc-session-tag {
            font-size: 0.8rem;
            font-weight: 600;
            color: var(--text-muted);
            letter-spacing: 0.5px;
            margin-top: 2px;
        }

        .dmc-qr-right {
            width: 72px;
            height: 72px;
            border: 1px solid var(--border-soft);
            border-radius: 8px;
            padding: 3px;
            background: #ffffff;
            flex-shrink: 0;
        }

        .dmc-qr-right img {
            width: 100%;
            height: 100%;
            object-fit: contain;
        }

        /* Student Metadata Card */
        .student-profile-strip {
            background: #f8fafc;
            border: 1px solid var(--border-soft);
            border-radius: 12px;
            padding: 12px 18px;
            margin-bottom: 16px;
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 12px;
            position: relative;
            z-index: 1;
        }

        .meta-cell {
            display: flex;
            flex-direction: column;
        }

        .meta-label {
            font-size: 0.7rem;
            text-transform: uppercase;
            letter-spacing: 0.6px;
            color: var(--text-muted);
            font-weight: 700;
            margin-bottom: 2px;
        }

        .meta-val {
            font-size: 0.96rem;
            font-weight: 800;
            color: var(--navy-deep);
        }

        /* Subject Table */
        .table-responsive {
            margin-bottom: 16px;
            position: relative;
            z-index: 1;
        }

        .academic-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 0.86rem;
        }

        .academic-table th {
            background-color: var(--navy-deep);
            color: #ffffff;
            padding: 9px 12px;
            text-align: left;
            font-weight: 700;
            font-size: 0.78rem;
            letter-spacing: 0.4px;
            text-transform: uppercase;
        }

        .academic-table th.center, .academic-table td.center {
            text-align: center;
        }

        .academic-table td {
            padding: 8px 12px;
            border-bottom: 1px solid #e2e8f0;
            color: var(--text-dark);
        }

        .academic-table tbody tr:nth-child(even) {
            background-color: #f8fafc;
        }

        .academic-table tbody tr:hover {
            background-color: #f1f5f9;
        }

        .pct-cell {
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
        }

        .mini-bar {
            width: 48px;
            height: 6px;
            background: #e2e8f0;
            border-radius: 3px;
            overflow: hidden;
        }

        .mini-bar-fill {
            height: 100%;
            background: linear-gradient(90deg, var(--navy-light), var(--gold));
            border-radius: 3px;
        }

        .badge {
            display: inline-block;
            padding: 2px 8px;
            border-radius: 4px;
            font-size: 0.74rem;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }

        .badge-pass {
            background-color: var(--success-bg);
            color: var(--success-green);
        }

        .badge-fail {
            background-color: var(--danger-bg);
            color: var(--danger-red);
        }

        .badge-leave {
            background-color: var(--leave-bg);
            color: var(--leave-orange);
        }

        /* Result Summary Dashboard Bar */
        .summary-dashboard-bar {
            display: grid;
            grid-template-columns: repeat(5, 1fr);
            gap: 12px;
            margin-bottom: 16px;
            position: relative;
            z-index: 1;
        }

        .summary-tile {
            background: #ffffff;
            border: 1px solid var(--border-soft);
            border-radius: 10px;
            padding: 10px 14px;
            text-align: center;
            box-shadow: 0 2px 6px rgba(0, 0, 0, 0.02);
        }

        .summary-tile-label {
            font-size: 0.7rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            color: var(--text-muted);
            margin-bottom: 3px;
        }

        .summary-tile-val {
            font-size: 1.25rem;
            font-weight: 900;
            color: var(--navy-deep);
        }

        .summary-tile.tile-status {
            background: linear-gradient(135deg, var(--navy-deep), var(--navy-light));
            border: none;
            color: #ffffff;
        }

        .summary-tile.tile-status .summary-tile-label {
            color: #dfb249;
        }

        .summary-tile.tile-status .summary-tile-val {
            color: #ffffff;
        }

        /* Performance Analytics: Dynamic Bell Curve */
        .analytics-strip {
            background: #f8fafc;
            border: 1px solid var(--border-soft);
            border-radius: 12px;
            padding: 12px 18px;
            margin-bottom: 16px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 20px;
            position: relative;
            z-index: 1;
        }

        .analytics-left {
            flex: 1;
        }

        .analytics-title {
            font-size: 0.8rem;
            font-weight: 800;
            color: var(--navy-deep);
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 2px;
        }

        .analytics-sub {
            font-size: 0.76rem;
            color: var(--text-muted);
        }

        .bell-curve-wrap {
            width: 200px;
            height: 48px;
            position: relative;
            flex-shrink: 0;
        }

        .bell-curve-wrap svg {
            width: 100%;
            height: 100%;
            overflow: visible;
        }

        .standing-pill {
            font-size: 0.74rem;
            font-weight: 800;
            padding: 4px 10px;
            border-radius: 20px;
            background: #ffffff;
            border: 1px solid var(--border-soft);
            color: var(--gold);
            display: inline-flex;
            align-items: center;
            gap: 5px;
            margin-top: 4px;
        }

        /* Notice & Remarks Grid */
        .remarks-grid {
            margin-bottom: 14px;
            position: relative;
            z-index: 1;
        }

        .remarks-box {
            background: var(--gold-subtle);
            border: 1px solid var(--gold-border);
            border-radius: 10px;
            padding: 8px 14px;
            font-size: 0.82rem;
            color: #5c4308;
            margin-bottom: 8px;
        }

        .remarks-box strong {
            color: #8c6007;
            font-size: 0.8rem;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-right: 6px;
        }

        .correction-note-box {
            background: #f8fafc;
            border: 1px dashed #cbd5e1;
            border-radius: 8px;
            padding: 7px 12px;
            font-size: 0.74rem;
            color: var(--text-muted);
            line-height: 1.4;
        }

        .correction-note-box strong {
            color: var(--navy-deep);
        }

        /* Digital Signatures Footer */
        .cert-footer-grid {
            display: flex;
            justify-content: space-between;
            align-items: flex-end;
            margin-top: 14px;
            padding-top: 10px;
            border-top: 1px dashed #cbd5e1;
            position: relative;
            z-index: 1;
            gap: 20px;
        }

        .sig-col {
            flex: 1;
            max-width: 250px;
            text-align: center;
        }

        .sig-img-container {
            height: 48px;
            display: flex;
            align-items: flex-end;
            justify-content: center;
            margin-bottom: 2px;
        }

        .sig-img {
            max-height: 46px;
            max-width: 140px;
            object-fit: contain;
            filter: contrast(1.08);
        }

        .sig-rule {
            width: 85%;
            margin: 0 auto 5px;
            border-top: 1px solid #64748b;
        }

        .sig-name {
            font-size: 0.88rem;
            font-weight: 800;
            color: var(--navy-deep);
            letter-spacing: 0.3px;
        }

        .sig-sub {
            font-size: 0.72rem;
            font-weight: 600;
            color: var(--text-muted);
            margin-top: 1px;
        }

        .official-seal {
            width: 68px;
            height: 68px;
            border: 2px dashed #b91c1c;
            border-radius: 50%;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            color: #b91c1c;
            text-align: center;
            padding: 3px;
            user-select: none;
            margin: 0 auto;
        }

        .seal-txt-top { font-size: 0.48rem; font-weight: 800; letter-spacing: 0.5px; text-transform: uppercase; }
        .seal-txt-mid { font-size: 0.66rem; font-weight: 900; letter-spacing: 1px; text-transform: uppercase; }
        .seal-txt-bot { font-size: 0.46rem; font-weight: 700; text-transform: uppercase; }

        /* Security Bar */
        .security-audit-bar {
            margin-top: 12px;
            padding-top: 6px;
            border-top: 1px solid #e2e8f0;
            display: flex;
            justify-content: space-between;
            font-size: 0.7rem;
            color: #94a3b8;
            font-weight: 600;
        }

        /* Action Tray */
        .action-tray {
            display: flex;
            gap: 12px;
            justify-content: center;
            margin-top: 20px;
        }

        .btn-ui {
            height: 46px;
            padding: 0 24px;
            border-radius: 10px;
            font-size: 0.9rem;
            font-weight: 700;
            cursor: pointer;
            display: inline-flex;
            align-items: center;
            gap: 8px;
            transition: all 0.2s ease;
            border: none;
        }

        .btn-print {
            background: linear-gradient(135deg, var(--gold) 0%, #b3831b 100%);
            color: #ffffff;
            box-shadow: 0 4px 12px rgba(198, 146, 34, 0.3);
        }

        .btn-print:hover {
            transform: translateY(-1px);
            box-shadow: 0 6px 16px rgba(198, 146, 34, 0.4);
        }

        .btn-new-search {
            background: #f1f5f9;
            color: var(--navy-deep);
            border: 1px solid var(--border-soft);
        }

        .btn-new-search:hover {
            background: #e2e8f0;
        }

        .footer-note {
            text-align: center;
            font-size: 0.8rem;
            color: var(--text-muted);
            margin-top: auto;
            padding-top: 20px;
        }

        /* Top 3 Honor Roll Modal */
        .modal-backdrop {
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(8, 29, 56, 0.6);
            backdrop-filter: blur(6px);
            z-index: 100;
            display: none;
            align-items: center;
            justify-content: center;
            padding: 20px;
            animation: fadeIn 0.2s ease-out;
        }

        .modal-card {
            background: #ffffff;
            width: 100%;
            max-width: 680px;
            border-radius: 20px;
            box-shadow: 0 20px 40px -10px rgba(0, 0, 0, 0.3);
            overflow: hidden;
            display: flex;
            flex-direction: column;
            max-height: 90vh;
        }

        .modal-header {
            background: linear-gradient(135deg, var(--navy-deep), #0d2a52);
            color: #ffffff;
            padding: 20px 24px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            position: relative;
        }

        .modal-header h3 {
            font-size: 1.25rem;
            font-weight: 800;
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .modal-close-btn {
            background: rgba(255, 255, 255, 0.12);
            border: none;
            color: #ffffff;
            width: 32px;
            height: 32px;
            border-radius: 50%;
            font-size: 1.2rem;
            cursor: pointer;
            display: flex;
            align-items: center;
            justify-content: center;
            transition: all 0.2s ease;
        }

        .modal-close-btn:hover {
            background: rgba(255, 255, 255, 0.25);
        }

        .modal-body {
            padding: 24px;
            overflow-y: auto;
            flex: 1;
        }

        .honor-controls {
            display: flex;
            gap: 12px;
            margin-bottom: 20px;
            flex-wrap: wrap;
            align-items: center;
        }

        .honor-toggle-group {
            display: flex;
            background: #f1f5f9;
            border-radius: 10px;
            padding: 3px;
        }

        .btn-toggle-mode {
            padding: 7px 14px;
            border: none;
            background: transparent;
            font-size: 0.82rem;
            font-weight: 700;
            color: var(--text-muted);
            border-radius: 8px;
            cursor: pointer;
            transition: all 0.2s ease;
        }

        .btn-toggle-mode.active {
            background: #ffffff;
            color: var(--navy-deep);
            box-shadow: 0 2px 6px rgba(0, 0, 0, 0.08);
        }

        .honor-dropdown-wrap {
            flex: 1;
            min-width: 200px;
        }

        .honor-dropdown-wrap select {
            width: 100%;
            height: 38px;
            padding: 0 12px;
            border: 1px solid var(--border-soft);
            border-radius: 8px;
            font-size: 0.88rem;
            color: var(--text-dark);
            background: #ffffff;
            font-family: inherit;
            outline: none;
        }

        .podium-container {
            display: flex;
            flex-direction: column;
            gap: 12px;
        }

        .podium-card {
            border-radius: 14px;
            padding: 16px 20px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            transition: all 0.2s ease;
            border: 1px solid transparent;
        }

        .podium-card:hover {
            transform: translateX(4px);
        }

        .podium-gold {
            background: linear-gradient(135deg, #fffbeb 0%, #fef3c7 100%);
            border-color: #fde68a;
        }

        .podium-silver {
            background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%);
            border-color: #e2e8f0;
        }

        .podium-bronze {
            background: linear-gradient(135deg, #fff7ed 0%, #ffedd5 100%);
            border-color: #fed7aa;
        }

        .podium-left {
            display: flex;
            align-items: center;
            gap: 16px;
        }

        .medal-badge {
            font-size: 1.8rem;
            line-height: 1;
        }

        .topper-details h4 {
            font-size: 1.05rem;
            color: var(--navy-deep);
            font-weight: 800;
            margin-bottom: 3px;
        }

        .topper-details p {
            font-size: 0.8rem;
            color: var(--text-muted);
            font-weight: 600;
        }

        .podium-right {
            text-align: right;
        }

        .topper-score {
            font-size: 1.15rem;
            font-weight: 900;
            color: var(--navy-deep);
        }

        .topper-pct {
            font-size: 0.8rem;
            font-weight: 700;
            color: var(--gold);
        }

        /* GUARANTEED EXACT 1-PAGE A4 PRINT STYLES */
        @media print {
            @page {
                size: A4 portrait;
                margin: 4mm 7mm !important;
            }
            html, body {
                background: #ffffff !important;
                color: #000000 !important;
                font-size: 10pt !important;
                width: 100% !important;
                height: 100% !important;
                margin: 0 !important;
                padding: 0 !important;
            }
            .executive-header, .search-container, .action-tray, .footer-note, .modal-backdrop {
                display: none !important;
            }
            .wrapper {
                max-width: 100% !important;
                padding: 0 !important;
                margin: 0 !important;
            }
            .result-view {
                display: block !important;
            }
            .dmc-card {
                box-shadow: none !important;
                border: 1px solid #94a3b8 !important;
                border-radius: 8px !important;
                padding: 10px 14px !important;
                margin: 0 !important;
                background-color: #ffffff !important;
                page-break-inside: avoid !important;
                break-inside: avoid !important;
            }
            .dmc-header {
                padding-bottom: 6px !important;
                margin-bottom: 8px !important;
                border-bottom: 1.5px solid #081d38 !important;
            }
            .dmc-logo-left, .dmc-qr-right {
                width: 52px !important;
                height: 52px !important;
            }
            .dmc-inst-name {
                font-size: 1.6rem !important;
                line-height: 1.1 !important;
            }
            .dmc-cert-label {
                font-size: 0.85rem !important;
                letter-spacing: 1.5px !important;
                margin-bottom: 2px !important;
            }
            .dmc-exam-term {
                font-size: 0.82rem !important;
            }
            .dmc-session-tag {
                font-size: 0.72rem !important;
            }
            .student-profile-strip {
                padding: 6px 10px !important;
                margin-bottom: 8px !important;
                gap: 8px !important;
                border: 1px solid #cbd5e1 !important;
            }
            .meta-label {
                font-size: 0.62rem !important;
            }
            .meta-val {
                font-size: 0.82rem !important;
            }
            .academic-table {
                font-size: 0.76rem !important;
            }
            .academic-table th {
                padding: 5px 6px !important;
                font-size: 0.7rem !important;
                background-color: #081d38 !important;
                color: #ffffff !important;
                -webkit-print-color-adjust: exact !important;
                print-color-adjust: exact !important;
            }
            .academic-table td {
                padding: 4px 6px !important;
            }
            .mini-bar {
                display: none !important;
            }
            .badge {
                padding: 1px 5px !important;
                font-size: 0.68rem !important;
                border: 1px solid #94a3b8 !important;
            }
            .summary-dashboard-bar {
                gap: 6px !important;
                margin-bottom: 8px !important;
            }
            .summary-tile {
                padding: 5px 6px !important;
                border: 1px solid #cbd5e1 !important;
            }
            .summary-tile-label {
                font-size: 0.62rem !important;
            }
            .summary-tile-val {
                font-size: 0.95rem !important;
            }
            .summary-tile.tile-status {
                background: #081d38 !important;
                -webkit-print-color-adjust: exact !important;
                print-color-adjust: exact !important;
            }
            .analytics-strip {
                padding: 6px 10px !important;
                margin-bottom: 8px !important;
            }
            .analytics-title {
                font-size: 0.72rem !important;
            }
            .analytics-sub {
                font-size: 0.66rem !important;
            }
            .bell-curve-wrap {
                width: 140px !important;
                height: 34px !important;
            }
            .remarks-box {
                padding: 4px 8px !important;
                font-size: 0.7rem !important;
                margin-bottom: 4px !important;
                border: 1px solid #cbd5e1 !important;
            }
            .correction-note-box {
                padding: 4px 8px !important;
                font-size: 0.64rem !important;
                line-height: 1.25 !important;
            }
            .cert-footer-grid {
                margin-top: 6px !important;
                padding-top: 5px !important;
            }
            .sig-img-container {
                height: 38px !important;
                margin-bottom: 1px !important;
            }
            .sig-img {
                max-height: 36px !important;
                max-width: 110px !important;
            }
            .sig-name {
                font-size: 0.74rem !important;
            }
            .sig-sub {
                font-size: 0.64rem !important;
            }
            .official-seal {
                width: 52px !important;
                height: 52px !important;
            }
            .security-audit-bar {
                margin-top: 4px !important;
                padding-top: 2px !important;
                font-size: 0.6rem !important;
            }
        }
    </style>
</head>
<body>
    <div class="wrapper">
        <!-- Top Executive Header -->
        <header class="executive-header">
            <div class="header-brand">
                <div class="crest-wrapper">
                    <img src="__LOGO_B64__" alt="Chenab College Shorkot">
                </div>
                <div class="header-title-block">
                    <h1>Chenab College Shorkot</h1>
                    <div class="sub-brand">
                        <span>Official Examination Web Portal</span>
                        <span>&bull;</span>
                        <span id="headerTotalBadge" style="color: #94a3b8; font-size: 0.78rem;">2,202 Verified Records</span>
                    </div>
                </div>
            </div>
            <div class="header-badge">
                <div class="badge-dot"></div>
                <span>Monthly Test October 2026</span>
            </div>
        </header>

        <!-- Search Bar -->
        <section class="search-container">
            <div class="search-title-area">
                <h2>Official Marks Verification Portal</h2>
                <p>Search by Student B-Form / Roll Number or Full Name across all sections</p>
            </div>

            <form class="search-form" onsubmit="handleSearchSubmit(event)">
                <div class="input-group-custom">
                    <span class="input-icon">🔍</span>
                    <input 
                        type="text" 
                        id="studentInput" 
                        placeholder="Enter Student ID (e.g. 2348, 33203...) or Student Name..."
                        autocomplete="off"
                        oninput="handleInputLive(event)"
                    >
                    <div class="suggestions-box" id="suggestionsBox"></div>
                </div>
                <button type="submit" class="btn-search">
                    <span>Search Result</span>
                </button>
                <button type="button" class="btn-honor-roll" onclick="openHonorModal()">
                    <span>🏆 Academic Honor Roll</span>
                </button>
            </form>

            <div class="error-alert" id="errorAlert">
                <span>⚠️</span>
                <span id="errorMessage">No student record found. Please verify the ID or Name.</span>
            </div>
        </section>

        <!-- Result View: DMC -->
        <main class="result-view" id="resultView">
            <div class="dmc-card" id="printableDmc">
                <!-- Header -->
                <div class="dmc-header">
                    <div class="dmc-logo-left">
                        <img src="__LOGO_B64__" alt="Chenab Logo">
                    </div>
                    <div class="dmc-header-center">
                        <div class="dmc-inst-name">Chenab College Shorkot</div>
                        <div class="dmc-cert-label">Detailed Marks Certificate</div>
                        <div class="dmc-exam-term">Monthly Test October 2026</div>
                        <div class="dmc-session-tag">Academic Session 2026-2027</div>
                    </div>
                    <div class="dmc-qr-right">
                        <img id="dynamicQr" src="" alt="Verification QR">
                    </div>
                </div>

                <!-- Student Profile Strip -->
                <div class="student-profile-strip">
                    <div class="meta-cell">
                        <span class="meta-label">Student Name</span>
                        <span class="meta-val" id="dispName">-</span>
                    </div>
                    <div class="meta-cell">
                        <span class="meta-label">Student Roll No / ID</span>
                        <span class="meta-val" id="dispId">-</span>
                    </div>
                    <div class="meta-cell">
                        <span class="meta-label">Class &amp; Section</span>
                        <span class="meta-val" id="dispClass">-</span>
                    </div>
                    <div class="meta-cell">
                        <span class="meta-label">Examination Date</span>
                        <span class="meta-val">October 2026</span>
                    </div>
                </div>

                <!-- Subject Marks Table -->
                <div class="table-responsive">
                    <table class="academic-table">
                        <thead>
                            <tr>
                                <th class="center" style="width: 45px;">Sr.</th>
                                <th>Subject</th>
                                <th class="center" style="width: 80px;">Max</th>
                                <th class="center" style="width: 80px;">Pass</th>
                                <th class="center" style="width: 100px;">Obtained</th>
                                <th class="center" style="width: 110px;">Progress</th>
                                <th class="center" style="width: 75px;">Grade</th>
                                <th class="center" style="width: 85px;">Status</th>
                            </tr>
                        </thead>
                        <tbody id="subjectTableBody"></tbody>
                    </table>
                </div>

                <!-- Summary Dashboard Tiles -->
                <div class="summary-dashboard-bar">
                    <div class="summary-tile">
                        <div class="summary-tile-label">Total Marks</div>
                        <div class="summary-tile-val" id="dispTotalMarks">-</div>
                    </div>
                    <div class="summary-tile">
                        <div class="summary-tile-label">Marks Obtained</div>
                        <div class="summary-tile-val" id="dispObtMarks">-</div>
                    </div>
                    <div class="summary-tile">
                        <div class="summary-tile-label">Percentage</div>
                        <div class="summary-tile-val" id="dispPct">-</div>
                    </div>
                    <div class="summary-tile">
                        <div class="summary-tile-label">Letter Grade</div>
                        <div class="summary-tile-val" id="dispGrade" style="color: var(--gold);">-</div>
                    </div>
                    <div class="summary-tile tile-status" id="dispStatusTile">
                        <div class="summary-tile-label">Result Status</div>
                        <div class="summary-tile-val" id="dispStatus">-</div>
                    </div>
                </div>

                <!-- Performance Analytics Strip: Dynamic Bell Curve -->
                <div class="analytics-strip">
                    <div class="analytics-left">
                        <div class="analytics-title">Academic Standing Distribution</div>
                        <div class="analytics-sub">Relative cohort standing determined by section ranking</div>
                        <div class="standing-pill" id="dispStandingPill">Class Average</div>
                    </div>
                    <div class="bell-curve-wrap">
                        <svg viewBox="0 0 200 48" preserveAspectRatio="none">
                            <defs>
                                <linearGradient id="bellGrad" x1="0%" y1="0%" x2="100%" y2="0%">
                                    <stop offset="0%" stop-color="#c02626" stop-opacity="0.15" />
                                    <stop offset="50%" stop-color="#081d38" stop-opacity="0.25" />
                                    <stop offset="100%" stop-color="#c69222" stop-opacity="0.35" />
                                </linearGradient>
                            </defs>
                            <path d="M 0,46 Q 35,46 65,40 T 100,5 T 135,40 Q 165,46 200,46 L 200,48 L 0,48 Z" fill="url(#bellGrad)" />
                            <path d="M 0,46 Q 35,46 65,40 T 100,5 T 135,40 Q 165,46 200,46" fill="none" stroke="#64748b" stroke-width="1.5" />
                            <g id="studentMarkerGroup" transform="translate(100, 0)">
                                <line x1="0" y1="0" x2="0" y2="46" stroke="#c69222" stroke-width="1.8" stroke-dasharray="2,2" />
                                <circle cx="0" cy="5" r="4.5" fill="#c69222" stroke="#ffffff" stroke-width="1.5" />
                                <polygon points="-4,-2 4,-2 0,4" fill="#c69222" />
                            </g>
                        </svg>
                    </div>
                </div>

                <!-- Remarks & Notice Grid -->
                <div class="remarks-grid">
                    <div class="remarks-box">
                        <strong>Official Remarks:</strong>
                        <span class="remarks-text" id="dispRemarks">Outstanding performance! Keep it up.</span>
                    </div>
                    <div class="correction-note-box">
                        <strong>NOTICE (E&amp;OE):</strong> In case of any discrepancy or query, students/parents are advised to contact the <strong>Examination Branch</strong> within <strong>7 days</strong> for verification &amp; correction.
                    </div>
                </div>

                <!-- Signatures: Mr. Sibtain Haider (Left) & Dynamic Section Head (Right) -->
                <div class="cert-footer-grid">
                    <div class="sig-col">
                        <div class="sig-img-container">
                            <img id="examHeadSigImg" src="" class="sig-img" alt="Mr. Sibtain Haider (Exam Head)">
                        </div>
                        <div class="sig-rule"></div>
                        <div class="sig-name" id="examHeadName">Mr. Sibtain Haider</div>
                        <div class="sig-sub" id="examHeadTitle">Exam Head (Incharge Examination)</div>
                    </div>
                    
                    <div class="official-seal">
                        <span class="seal-txt-top">Chenab College</span>
                        <span class="seal-txt-mid">VERIFIED</span>
                        <span class="seal-txt-bot">Exam Branch</span>
                    </div>

                    <div class="sig-col">
                        <div class="sig-img-container">
                            <img id="secHeadSigImg" src="" class="sig-img" alt="Section Head Signature">
                        </div>
                        <div class="sig-rule"></div>
                        <div class="sig-name" id="secHeadName">Section Head</div>
                        <div class="sig-sub" id="secHeadTitle">Chenab College Shorkot</div>
                    </div>
                </div>

                <!-- Security Audit Line -->
                <div class="security-audit-bar">
                    <span>🔒 Computer Generated Official DMC &bull; Chenab College Shorkot</span>
                    <span id="auditTimestamp">System Verified Record</span>
                </div>
            </div>

            <!-- Action Tray -->
            <div class="action-tray">
                <button type="button" class="btn-ui btn-new-search" onclick="backToSearch()">
                    🔍 Search Another Student
                </button>
                <button type="button" class="btn-ui btn-print" onclick="window.print()">
                    🖨️ Print Official DMC (1-Page A4)
                </button>
            </div>
        </main>

        <footer class="footer-note">
            &copy; 2026 Chenab College Shorkot &bull; Examination Branch Digital Verification System
        </footer>
    </div>

    <!-- Academic Honor Roll Modal -->
    <div class="modal-backdrop" id="honorModal" onclick="handleBackdropClick(event)">
        <div class="modal-card">
            <div class="modal-header">
                <h3>🏆 Academic Honor Roll (Top 3 Positions)</h3>
                <button type="button" class="modal-close-btn" onclick="closeHonorModal()">&times;</button>
            </div>
            <div class="modal-body">
                <div class="honor-controls">
                    <div class="honor-toggle-group">
                        <button type="button" class="btn-toggle-mode active" id="btnModeSection" onclick="setHonorMode('section')">
                            🏆 Section-Wise
                        </button>
                        <button type="button" class="btn-toggle-mode" id="btnModeClass" onclick="setHonorMode('class')">
                            🌐 Entire Class
                        </button>
                    </div>

                    <div class="honor-dropdown-wrap" id="sectionDropdownWrap">
                        <select id="sectionSelect" onchange="renderHonorToppers()"></select>
                    </div>

                    <div class="honor-dropdown-wrap" id="classDropdownWrap" style="display: none;">
                        <select id="classSelect" onchange="renderHonorToppers()"></select>
                    </div>
                </div>

                <div class="podium-container" id="podiumContainer"></div>
            </div>
        </div>
    </div>

    <script>
        const DB = __JSON_DATA__;
        const SIGNATURES = __SIGNATURES_JSON__;

        let currentStudent = null;
        let currentHonorMode = 'section';

        // Pre-populate Honor Roll dropdowns
        function initHonorDropdowns() {
            const secSelect = document.getElementById('sectionSelect');
            secSelect.innerHTML = '';
            DB.sections.forEach(sec => {
                const opt = document.createElement('option');
                opt.value = sec;
                opt.textContent = sec;
                secSelect.appendChild(opt);
            });

            const clsSelect = document.getElementById('classSelect');
            clsSelect.innerHTML = '';
            DB.grade_groups.forEach(grp => {
                const opt = document.createElement('option');
                opt.value = grp;
                opt.textContent = grp;
                clsSelect.appendChild(opt);
            });
        }
        initHonorDropdowns();

        function handleInputLive(e) {
            const val = e.target.value.trim().toLowerCase();
            const box = document.getElementById('suggestionsBox');

            if (!val) {
                box.style.display = 'none';
                box.innerHTML = '';
                return;
            }

            const matches = DB.students.filter(s => 
                s.id.toLowerCase().includes(val) || 
                s.name.toLowerCase().includes(val)
            ).slice(0, 8);

            if (matches.length === 0) {
                box.style.display = 'none';
                box.innerHTML = '';
                return;
            }

            box.innerHTML = '';
            matches.forEach(st => {
                const item = document.createElement('div');
                item.className = 'suggestion-item';
                item.onclick = () => {
                    document.getElementById('studentInput').value = st.id;
                    box.style.display = 'none';
                    renderResult(st);
                };

                item.innerHTML = `
                    <div>
                        <div class="sug-name">${st.name}</div>
                        <div class="sug-meta">${st.class_name} &bull; ID: ${st.id}</div>
                    </div>
                    <span class="sug-badge">${st.percentage}%</span>
                `;
                box.appendChild(item);
            });

            box.style.display = 'block';
        }

        document.addEventListener('click', (e) => {
            if (!e.target.closest('.input-group-custom')) {
                document.getElementById('suggestionsBox').style.display = 'none';
            }
        });

        function handleSearchSubmit(e) {
            e.preventDefault();
            const val = document.getElementById('studentInput').value.trim();
            const errorAlert = document.getElementById('errorAlert');
            const errorMsg = document.getElementById('errorMessage');
            document.getElementById('suggestionsBox').style.display = 'none';

            if (!val) {
                errorMsg.textContent = 'Please enter a Student ID or Name to verify result.';
                errorAlert.style.display = 'flex';
                return;
            }

            // Exact ID Match first
            let found = DB.students.find(s => s.id.toLowerCase() === val.toLowerCase());
            
            // If not exact ID, search substring on ID or Name
            if (!found) {
                const candidates = DB.students.filter(s => 
                    s.id.toLowerCase().includes(val.toLowerCase()) || 
                    s.name.toLowerCase().includes(val.toLowerCase())
                );
                if (candidates.length > 0) {
                    found = candidates[0];
                }
            }

            if (found) {
                errorAlert.style.display = 'none';
                renderResult(found);
            } else {
                errorMsg.textContent = `No student record found matching "${val}". Please verify Roll No or Name.`;
                errorAlert.style.display = 'flex';
                document.getElementById('resultView').style.display = 'none';
            }
        }

        function renderResult(st) {
            currentStudent = st;
            document.getElementById('dispName').textContent = st.name;
            document.getElementById('dispId').textContent = st.id;
            document.getElementById('dispClass').textContent = st.class_name;
            document.getElementById('dispTotalMarks').textContent = st.total_marks;
            document.getElementById('dispObtMarks').textContent = st.obtained_marks;
            document.getElementById('dispPct').textContent = st.percentage + (st.percentage.includes('%') ? '' : '%');
            document.getElementById('dispGrade').textContent = st.grade;
            document.getElementById('dispRemarks').textContent = st.remarks || 'Academic achievement verified by examination authority.';

            // Render Signatures
            const examHeadImg = document.getElementById('examHeadSigImg');
            if (examHeadImg && SIGNATURES.exam_head_sibtain) {
                examHeadImg.src = SIGNATURES.exam_head_sibtain;
            }
            document.getElementById('examHeadName').textContent = 'Mr. Sibtain Haider';
            document.getElementById('examHeadTitle').textContent = 'Exam Head (Incharge Examination)';

            const secHeadImg = document.getElementById('secHeadSigImg');
            if (secHeadImg && st.sec_head_sig && SIGNATURES[st.sec_head_sig]) {
                secHeadImg.src = SIGNATURES[st.sec_head_sig];
            }
            document.getElementById('secHeadName').textContent = st.sec_head_name || 'Section Head';
            document.getElementById('secHeadTitle').textContent = st.sec_head_title || 'Chenab College Shorkot';

            // QR Code
            const qrUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=100x100&data=' + encodeURIComponent('CHENAB COLLEGE VERIFIED: ' + st.name + ' (' + st.id + ') Marks: ' + st.obtained_marks + '/' + st.total_marks + ' (' + st.percentage + '%)');
            document.getElementById('dynamicQr').src = qrUrl;

            updateBellCurve(st);

            // Status Tile
            const statusVal = document.getElementById('dispStatus');
            const statusTile = document.getElementById('dispStatusTile');
            statusVal.textContent = st.status.toUpperCase();
            
            if (st.status.toLowerCase() === 'pass') {
                statusTile.style.background = 'linear-gradient(135deg, #0d824d, #085a34)';
            } else {
                statusTile.style.background = 'linear-gradient(135deg, #c02626, #8e1818)';
            }

            // Subject Table
            const tbody = document.getElementById('subjectTableBody');
            tbody.innerHTML = '';

            st.subjects.forEach((sub, idx) => {
                const tr = document.createElement('tr');

                let badgeClass = 'badge-pass';
                let statusText = 'Pass';
                let pctNum = 0;

                if (sub.status === 'Leave') {
                    badgeClass = 'badge-leave';
                    statusText = 'Leave';
                } else if (sub.status === 'Fail') {
                    badgeClass = 'badge-fail';
                    statusText = 'Fail';
                }

                if (sub.percentage && sub.percentage.includes('%')) {
                    pctNum = parseFloat(sub.percentage) || 0;
                }

                const miniBarHtml = sub.status === 'Leave' 
                    ? '<span style="color: #94a3b8;">-</span>'
                    : `
                        <div class="pct-cell">
                            <span>${sub.percentage}</span>
                            <div class="mini-bar">
                                <div class="mini-bar-fill" style="width: ${pctNum}%;"></div>
                            </div>
                        </div>
                    `;

                tr.innerHTML = `
                    <td class="center" style="color: var(--text-muted); font-size: 0.8rem;">${idx + 1}</td>
                    <td style="font-weight: 700; color: var(--navy-deep);">${sub.name}</td>
                    <td class="center" style="font-weight: 600;">${sub.max}</td>
                    <td class="center" style="color: var(--text-muted);">${sub.pass}</td>
                    <td class="center" style="font-weight: 800; font-size: 0.98rem; color: var(--navy-deep);">${sub.obtained}</td>
                    <td class="center">${miniBarHtml}</td>
                    <td class="center" style="font-weight: 800; color: var(--gold);">${sub.grade}</td>
                    <td class="center"><span class="badge ${badgeClass}">${statusText}</span></td>
                `;
                tbody.appendChild(tr);
            });

            const resultView = document.getElementById('resultView');
            resultView.style.display = 'block';
            resultView.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }

        function updateBellCurve(st) {
            const marker = document.getElementById('studentMarkerGroup');
            const pill = document.getElementById('dispStandingPill');
            const secPos = parseInt(st.sec_pos);

            let xPos = 100;

            if (!isNaN(secPos)) {
                if (secPos === 1) {
                    pill.textContent = '🌟 1st Position';
                    pill.style.color = '#c69222';
                    xPos = 178;
                } else if (secPos <= 3) {
                    pill.textContent = '🏆 Top 3';
                    pill.style.color = '#c69222';
                    xPos = 162;
                } else if (secPos <= 7) {
                    pill.textContent = '⭐ Above Avg';
                    pill.style.color = '#0d824d';
                    xPos = 138;
                } else if (secPos <= 15) {
                    pill.textContent = 'Class Average';
                    pill.style.color = '#0e355c';
                    xPos = 100;
                } else if (secPos <= 20) {
                    pill.textContent = 'Below Avg';
                    pill.style.color = '#b45309';
                    xPos = 62;
                } else {
                    pill.textContent = 'Needs Attention';
                    pill.style.color = '#c02626';
                    xPos = 28;
                }
            } else {
                pill.textContent = 'Class Average';
                xPos = 100;
            }

            const distFromCenter = Math.abs(xPos - 100) / 85;
            const yPos = 5 + (distFromCenter * distFromCenter * 33);

            marker.setAttribute('transform', `translate(${xPos.toFixed(1)}, 0)`);
            const circle = marker.querySelector('circle');
            if (circle) circle.setAttribute('cy', yPos.toFixed(1));
            const poly = marker.querySelector('polygon');
            if (poly) poly.setAttribute('transform', `translate(0, ${yPos - 10})`);
        }

        function backToSearch() {
            document.getElementById('studentInput').value = '';
            document.getElementById('studentInput').focus();
            document.getElementById('resultView').style.display = 'none';
            document.getElementById('errorAlert').style.display = 'none';
        }

        // Honor Roll Logic
        function openHonorModal() {
            if (currentStudent) {
                document.getElementById('sectionSelect').value = currentStudent.class_name;
                document.getElementById('classSelect').value = currentStudent.grade_group;
            }
            renderHonorToppers();
            document.getElementById('honorModal').style.display = 'flex';
        }

        function closeHonorModal() {
            document.getElementById('honorModal').style.display = 'none';
        }

        function handleBackdropClick(e) {
            if (e.target.id === 'honorModal') {
                closeHonorModal();
            }
        }

        function setHonorMode(mode) {
            currentHonorMode = mode;
            const btnClass = document.getElementById('btnModeClass');
            const btnSection = document.getElementById('btnModeSection');
            const secWrap = document.getElementById('sectionDropdownWrap');
            const clsWrap = document.getElementById('classDropdownWrap');

            if (mode === 'class') {
                btnClass.classList.add('active');
                btnSection.classList.remove('active');
                secWrap.style.display = 'none';
                clsWrap.style.display = 'block';
            } else {
                btnSection.classList.add('active');
                btnClass.classList.remove('active');
                secWrap.style.display = 'block';
                clsWrap.style.display = 'none';
            }

            renderHonorToppers();
        }

        function renderHonorToppers() {
            const container = document.getElementById('podiumContainer');
            container.innerHTML = '';

            let list = [];

            if (currentHonorMode === 'section') {
                const selectedSec = document.getElementById('sectionSelect').value;
                list = DB.students.filter(st => st.class_name === selectedSec);
                
                // Sort by sec_pos if numeric, otherwise by percentage desc
                list.sort((a, b) => {
                    const pa = parseInt(a.sec_pos);
                    const pb = parseInt(b.sec_pos);
                    if (!isNaN(pa) && !isNaN(pb)) return pa - pb;
                    return parseFloat(b.percentage || 0) - parseFloat(a.percentage || 0);
                });
            } else {
                const selectedClass = document.getElementById('classSelect').value;
                list = DB.students.filter(st => st.grade_group === selectedClass);
                
                // Sort by class_pos if numeric, otherwise by percentage desc
                list.sort((a, b) => {
                    const pa = parseInt(a.class_pos);
                    const pb = parseInt(b.class_pos);
                    if (!isNaN(pa) && !isNaN(pb)) return pa - pb;
                    return parseFloat(b.percentage || 0) - parseFloat(a.percentage || 0);
                });
            }

            // Pick Top 3
            const p1Students = list.slice(0, 1);
            const p2Students = list.slice(1, 2);
            const p3Students = list.slice(2, 3);

            const tiers = [
                { pos: 1, title: '1st Position', badge: '🥇', classModifier: 'podium-gold', students: p1Students },
                { pos: 2, title: '2nd Position', badge: '🥈', classModifier: 'podium-silver', students: p2Students },
                { pos: 3, title: '3rd Position', badge: '🥉', classModifier: 'podium-bronze', students: p3Students }
            ];

            tiers.forEach(tier => {
                tier.students.forEach(st => {
                    const card = document.createElement('div');
                    card.className = `podium-card ${tier.classModifier}`;
                    card.style.cursor = 'pointer';
                    card.title = 'Click to view full result';
                    card.onclick = () => {
                        closeHonorModal();
                        document.getElementById('studentInput').value = st.id;
                        renderResult(st);
                    };

                    card.innerHTML = `
                        <div class="podium-left">
                            <div class="medal-badge">${tier.badge}</div>
                            <div class="topper-details">
                                <h4>${st.name}</h4>
                                <p>${tier.title} &bull; ${st.class_name}</p>
                            </div>
                        </div>
                        <div class="podium-right">
                            <div class="topper-score">${st.obtained_marks} <span style="font-size: 0.85rem; color: #64748b;">/ ${st.total_marks}</span></div>
                            <div class="topper-pct">${st.percentage}% (Grade ${st.grade})</div>
                        </div>
                    `;
                    container.appendChild(card);
                });
            });

            if (list.length === 0) {
                container.innerHTML = '<div style="text-align: center; color: #94a3b8; padding: 20px;">No students found for this selection.</div>';
            }
        }

        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') closeHonorModal();
        });
    </script>
</body>
</html>
"""

# Replace placeholders
final_html = html_template.replace("__LOGO_B64__", logo_b64)
final_html = final_html.replace("__FONT_B64__", font_b64)
final_html = final_html.replace("__SIGNATURES_JSON__", signatures_json_str)
final_html = final_html.replace("__JSON_DATA__", json_str)

# Write to portal folder index.html
portal_index_path = os.path.join(portal_folder, "index.html")
with open(portal_index_path, "w", encoding="utf-8") as f:
    f.write(final_html)

# Write to Desktop standalone file
direct_desktop_file = os.path.join(desktop, "ResultPortal_ChenabCollege.html")
with open(direct_desktop_file, "w", encoding="utf-8") as f:
    f.write(final_html)

print("SUCCESS!")
print(f"File 1: {portal_index_path} ({len(final_html)} chars)")
print(f"File 2: {direct_desktop_file} ({len(final_html)} chars)")
