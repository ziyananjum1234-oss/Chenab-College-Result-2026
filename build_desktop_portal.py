import os
import base64
import json
import pandas as pd
import xlrd
import re

desktop = r"C:\Users\IT_OFFICE02\Desktop"
portal_folder = os.path.join(desktop, "Chenab_Result_Portal")
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



excel_dir = os.path.join(desktop, "Monthly Test October Exel Lists 2026")

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

files = sorted([f for f in os.listdir(excel_dir) if f.endswith(".xls") or f.endswith(".xlsx")])
students = []
sections_set = set()
grade_groups_set = set()

full_exam_name = "Monthly Test October 2026"

for f_idx, fname in enumerate(files):
    p = os.path.join(excel_dir, fname)
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
            
            # Skip empty subject columns
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
grade_order = [
    "Class 1st", "Class 2nd", "Class 3rd", "Class 4th", "Class 5th",
    "Class 6th", "Class 7th", "Class 8th", "Class 9th", "Class 10th",
    "Class 11th", "Class 12th", "Hifz-ul-Quran", "Cambridge Wing"
]
grade_groups_list = [g for g in grade_order if g in grade_groups_set] + [g for g in sorted(grade_groups_set) if g not in grade_order]

portal_data = {
    "institute": "Chenab College Shorkot",
    "test_name": full_exam_name,
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
html_code = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Online Result Portal &bull; Chenab College Shorkot</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Cinzel:wght@600;700;800;900&family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=Playfair+Display:ital,wght@0,600;1,600&display=swap" rel="stylesheet">
    <style>
        @font-face {{
            font-family: 'Old English Custom';
            src: url('data:font/truetype;charset=utf-8;base64,{font_b64}') format('truetype');
            font-weight: normal;
            font-style: normal;
        }}

        :root {{
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
        }}

        * {{
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }}

        body {{
            font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            background: var(--slate-bg);
            color: var(--text-dark);
            min-height: 100vh;
            display: flex;
            flex-direction: column;
            align-items: center;
            line-height: 1.4;
            background-image: 
                radial-gradient(at 0% 0%, rgba(12, 39, 74, 0.03) 0px, transparent 50%),
                radial-gradient(at 100% 100%, rgba(198, 146, 34, 0.04) 0px, transparent 50%);
        }}

        .wrapper {{
            width: 100%;
            max-width: 900px;
            padding: 30px 16px 50px;
            display: flex;
            flex-direction: column;
            gap: 22px;
        }}

        .old-english-heading {{
            font-family: 'Old English Custom', 'Old English Text MT', 'Engravers Old English', Georgia, serif;
            letter-spacing: 0.5px;
            font-weight: normal;
        }}

        /* Executive Header */
        .executive-header {{
            background: #ffffff;
            border-radius: 18px;
            padding: 28px 20px;
            text-align: center;
            border: 1px solid rgba(226, 232, 240, 0.9);
            box-shadow: 0 10px 30px -8px rgba(8, 29, 56, 0.07);
            position: relative;
            overflow: hidden;
        }}
        .executive-header::before {{
            content: "";
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 5px;
            background: linear-gradient(90deg, var(--navy-deep), var(--gold), var(--navy-deep));
        }}

        .college-logo {{
            width: 85px;
            height: 85px;
            object-fit: contain;
            margin: 0 auto 10px;
            display: block;
            border-radius: 50%;
            border: 3px solid var(--gold);
            padding: 3px;
            background: #ffffff;
            box-shadow: 0 4px 16px rgba(198, 146, 34, 0.2);
        }}

        .college-title {{
            font-size: 2.85rem;
            color: var(--navy-deep);
            line-height: 1.15;
            margin-bottom: 4px;
        }}

        .controller-tag {{
            font-size: 0.9rem;
            font-weight: 700;
            color: var(--gold);
            text-transform: uppercase;
            letter-spacing: 2px;
            margin-bottom: 10px;
            display: inline-block;
        }}

        .exam-info-badges {{
            display: inline-flex;
            gap: 10px;
            background: #f8fafc;
            border: 1px solid var(--border-soft);
            padding: 6px 20px;
            border-radius: 40px;
            font-size: 0.82rem;
            color: var(--text-muted);
            font-weight: 600;
            flex-wrap: wrap;
            justify-content: center;
        }}
        .exam-info-badges strong {{
            color: var(--navy-deep);
        }}

        /* Search Section */
        .search-container {{
            background: #ffffff;
            border-radius: 18px;
            padding: 26px 24px;
            border: 1px solid var(--border-soft);
            box-shadow: 0 10px 30px -8px rgba(8, 29, 56, 0.07);
            text-align: center;
        }}

        .search-title {{
            font-size: 1.25rem;
            font-weight: 700;
            color: var(--navy-deep);
            margin-bottom: 4px;
        }}

        .search-hint {{
            color: var(--text-muted);
            font-size: 0.88rem;
            margin-bottom: 18px;
        }}

        .search-form {{
            display: flex;
            gap: 12px;
            max-width: 600px;
            margin: 0 auto;
            flex-direction: column;
        }}
        @media (min-width: 580px) {{
            .search-form {{
                flex-direction: row;
            }}
        }}

        .input-group {{
            position: relative;
            flex: 1;
        }}

        .search-input {{
            width: 100%;
            height: 52px;
            padding: 0 18px 0 46px;
            font-size: 1.02rem;
            font-weight: 500;
            color: var(--navy-deep);
            background: #f8fafc;
            border: 2px solid #cbd5e1;
            border-radius: 12px;
            outline: none;
            transition: all 0.2s ease;
            font-family: inherit;
        }}
        .search-input:focus {{
            background: #ffffff;
            border-color: var(--gold);
            box-shadow: 0 0 0 4px rgba(198, 146, 34, 0.15);
        }}

        .search-icon {{
            position: absolute;
            left: 15px;
            top: 50%;
            transform: translateY(-50%);
            color: #94a3b8;
            font-size: 1.2rem;
        }}

        .btn-check {{
            height: 52px;
            padding: 0 30px;
            background: linear-gradient(135deg, var(--navy-deep), var(--navy-card));
            color: #ffffff;
            font-size: 0.98rem;
            font-weight: 700;
            border: none;
            border-radius: 12px;
            cursor: pointer;
            transition: all 0.2s ease;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
            box-shadow: 0 4px 14px rgba(8, 29, 56, 0.25);
            white-space: nowrap;
        }}
        .btn-check:hover {{
            background: linear-gradient(135deg, var(--navy-card), var(--navy-light));
            transform: translateY(-1px);
        }}

        /* Clean "View Top 3 Positions" Action Button */
        .honor-action-bar {{
            margin-top: 18px;
            padding-top: 14px;
            border-top: 1px dashed #e2e8f0;
            display: flex;
            justify-content: center;
        }}
        .btn-honor-roll {{
            background: var(--gold-subtle);
            color: var(--navy-deep);
            border: 1px solid var(--gold-border);
            padding: 8px 18px;
            border-radius: 30px;
            font-size: 0.88rem;
            font-weight: 700;
            cursor: pointer;
            display: inline-flex;
            align-items: center;
            gap: 8px;
            transition: all 0.2s;
            font-family: inherit;
        }}
        .btn-honor-roll:hover {{
            background: #fdf2d8;
            border-color: var(--gold);
            transform: translateY(-1px);
            box-shadow: 0 4px 12px rgba(198, 146, 34, 0.18);
        }}

        .error-box {{
            display: none;
            max-width: 600px;
            margin: 16px auto 0;
            padding: 12px 16px;
            background: var(--danger-bg);
            border: 1px solid #fecaca;
            color: var(--danger-red);
            border-radius: 10px;
            font-size: 0.9rem;
            font-weight: 600;
            align-items: center;
            gap: 10px;
            text-align: left;
        }}

        /* ========================================================
           RESULT TRANSCRIPT / DMC
        ======================================================== */
        #resultView {{
            display: none;
            animation: slideUp 0.3s cubic-bezier(0.16, 1, 0.3, 1);
        }}
        @keyframes slideUp {{
            from {{ opacity: 0; transform: translateY(16px); }}
            to {{ opacity: 1; transform: translateY(0); }}
        }}

        .transcript-card {{
            background: #ffffff;
            border-radius: 18px;
            border: 2px solid var(--gold-border);
            box-shadow: 0 20px 45px -10px rgba(8, 29, 56, 0.16);
            position: relative;
            overflow: hidden;
        }}

        .transcript-inner {{
            padding: 30px 30px 24px;
            position: relative;
            background: #ffffff;
        }}
        .transcript-inner::before {{
            content: "CCS";
            position: absolute;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%) rotate(-25deg);
            font-family: 'Cinzel', serif;
            font-size: 15rem;
            font-weight: 900;
            color: rgba(8, 29, 56, 0.025);
            pointer-events: none;
            user-select: none;
            z-index: 0;
        }}

        /* Certificate Title Header */
        .cert-header {{
            display: flex;
            align-items: center;
            justify-content: space-between;
            border-bottom: 2px solid var(--navy-deep);
            padding-bottom: 14px;
            margin-bottom: 16px;
            position: relative;
            z-index: 1;
            gap: 16px;
        }}
        .cert-header-main {{
            flex: 1;
            text-align: center;
        }}
        .cert-logo-center {{
            width: 64px;
            height: 64px;
            object-fit: contain;
            margin: 0 auto 4px;
            display: block;
        }}
        .cert-college {{
            font-size: 2.65rem;
            color: var(--navy-deep);
            line-height: 1.1;
        }}
        .cert-doc-type {{
            font-family: 'Playfair Display', Georgia, serif;
            font-size: 1.25rem;
            color: var(--gold);
            font-style: italic;
            font-weight: 600;
            letter-spacing: 0.5px;
            margin-top: 1px;
        }}
        .cert-exam-title {{
            font-size: 0.95rem;
            font-weight: 700;
            color: var(--navy-deep);
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-top: 3px;
        }}
        .cert-session-line {{
            font-size: 0.82rem;
            font-weight: 600;
            color: var(--text-muted);
            text-transform: uppercase;
            letter-spacing: 1.2px;
            margin-top: 2px;
        }}

        .cert-qr-box {{
            width: 76px;
            text-align: center;
            flex-shrink: 0;
            display: none;
        }}
        @media (min-width: 620px) {{
            .cert-qr-box {{
                display: block;
            }}
        }}
        .qr-img {{
            width: 70px;
            height: 70px;
            border: 1px solid #cbd5e1;
            border-radius: 6px;
            padding: 2px;
            background: #ffffff;
        }}
        .qr-caption {{
            font-size: 0.6rem;
            font-weight: 700;
            color: var(--text-muted);
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-top: 2px;
        }}

        /* Student Metadata Grid */
        .student-bio-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(170px, 1fr));
            gap: 10px;
            background: #f8fafc;
            border: 1px solid var(--border-soft);
            border-radius: 12px;
            padding: 12px 16px;
            margin-bottom: 16px;
            position: relative;
            z-index: 1;
        }}
        .bio-field {{
            display: flex;
            flex-direction: column;
        }}
        .bio-label {{
            font-size: 0.68rem;
            text-transform: uppercase;
            font-weight: 700;
            color: var(--text-muted);
            letter-spacing: 0.6px;
            margin-bottom: 1px;
        }}
        .bio-val {{
            font-size: 0.95rem;
            font-weight: 700;
            color: var(--navy-deep);
        }}
        .bio-val.highlight {{
            color: var(--gold);
        }}

        /* Master Marks Table */
        .table-wrap {{
            overflow-x: auto;
            margin-bottom: 16px;
            position: relative;
            z-index: 1;
            border-radius: 10px;
            border: 1px solid var(--border-soft);
        }}
        .cert-table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 0.88rem;
        }}
        .cert-table thead th {{
            background: var(--navy-deep);
            color: #ffffff;
            font-weight: 700;
            font-size: 0.75rem;
            text-transform: uppercase;
            letter-spacing: 0.6px;
            padding: 9px 12px;
            text-align: left;
        }}
        .cert-table thead th.center {{
            text-align: center;
        }}
        .cert-table tbody td {{
            padding: 8px 12px;
            border-bottom: 1px solid var(--border-soft);
            color: var(--text-dark);
            vertical-align: middle;
        }}
        .cert-table tbody td.center {{
            text-align: center;
        }}
        .cert-table tbody tr:nth-child(even) {{
            background: #fafbfc;
        }}

        .pct-cell {{
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 6px;
            font-weight: 700;
            color: var(--navy-deep);
        }}
        .mini-bar {{
            width: 42px;
            height: 5px;
            background: #e2e8f0;
            border-radius: 10px;
            overflow: hidden;
            display: inline-block;
        }}
        .mini-bar-fill {{
            height: 100%;
            background: linear-gradient(90deg, var(--gold), var(--success-green));
            border-radius: 10px;
        }}

        .badge {{
            display: inline-block;
            padding: 2px 9px;
            border-radius: 20px;
            font-size: 0.72rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}
        .badge-pass {{
            background: var(--success-bg);
            color: var(--success-green);
            border: 1px solid #bbf7d0;
        }}
        .badge-fail {{
            background: var(--danger-bg);
            color: var(--danger-red);
            border: 1px solid #fecaca;
        }}
        .badge-leave {{
            background: var(--leave-bg);
            color: var(--leave-orange);
            border: 1px solid #fde68a;
        }}

        /* Executive Summary Panel */
        .summary-dashboard {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(115px, 1fr));
            gap: 10px;
            margin-bottom: 14px;
            position: relative;
            z-index: 1;
        }}
        .dash-tile {{
            background: #f8fafc;
            border: 1px solid var(--border-soft);
            border-radius: 10px;
            padding: 10px 8px;
            text-align: center;
        }}
        .dash-tile.grand-status {{
            background: linear-gradient(135deg, var(--navy-deep), var(--navy-card));
            color: #ffffff;
            border-color: var(--navy-deep);
        }}
        .dash-tile.grand-status .tile-lbl {{
            color: rgba(255, 255, 255, 0.75);
        }}
        .dash-tile.grand-status .tile-val {{
            color: #ffffff;
        }}
        .tile-lbl {{
            font-size: 0.65rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.6px;
            color: var(--text-muted);
            margin-bottom: 3px;
        }}
        .tile-val {{
            font-size: 1.3rem;
            font-weight: 800;
            color: var(--navy-deep);
        }}
        .tile-val.gold {{
            color: var(--gold);
        }}

        /* Side-by-Side: Bell Curve + Remarks & Notice */
        .analytics-row {{
            display: grid;
            grid-template-columns: 1fr;
            gap: 12px;
            margin-bottom: 18px;
            position: relative;
            z-index: 1;
        }}
        @media (min-width: 640px) {{
            .analytics-row {{
                grid-template-columns: 240px 1fr;
            }}
        }}

        .bell-curve-card {{
            background: #f8fafc;
            border: 1px solid var(--border-soft);
            border-radius: 10px;
            padding: 10px 12px;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
        }}
        .curve-head {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 4px;
        }}
        .curve-title {{
            font-size: 0.72rem;
            font-weight: 800;
            color: var(--navy-deep);
            text-transform: uppercase;
            letter-spacing: 0.5px;
            display: flex;
            align-items: center;
            gap: 5px;
        }}
        .curve-standing-pill {{
            font-size: 0.7rem;
            font-weight: 800;
            padding: 2px 8px;
            border-radius: 20px;
            background: var(--gold-subtle);
            border: 1px solid var(--gold-border);
            color: var(--gold);
        }}
        .curve-svg-box {{
            width: 100%;
            height: 48px;
            margin: 2px 0;
            position: relative;
        }}
        .curve-svg {{
            width: 100%;
            height: 100%;
            overflow: visible;
        }}
        .curve-labels {{
            display: flex;
            justify-content: space-between;
            font-size: 0.62rem;
            color: var(--text-muted);
            font-weight: 600;
            border-top: 1px dashed #cbd5e1;
            padding-top: 2px;
        }}

        .info-col {{
            display: flex;
            flex-direction: column;
            gap: 8px;
            justify-content: center;
        }}
        .remarks-box {{
            background: var(--gold-subtle);
            border: 1px solid var(--gold-border);
            border-radius: 8px;
            padding: 8px 12px;
            display: flex;
            align-items: center;
            gap: 8px;
            font-size: 0.84rem;
        }}
        .remarks-tag {{
            font-weight: 800;
            color: var(--gold);
            text-transform: uppercase;
            font-size: 0.72rem;
            letter-spacing: 0.8px;
            white-space: nowrap;
        }}
        .remarks-text {{
            color: var(--navy-deep);
            font-weight: 600;
            font-style: italic;
        }}

        .correction-note-box {{
            background: #f8fafc;
            border: 1px dashed #cbd5e1;
            border-left: 3px solid var(--navy-deep);
            border-radius: 8px;
            padding: 7px 12px;
            font-size: 0.74rem;
            color: var(--text-muted);
            line-height: 1.4;
        }}
        .correction-note-box strong {{
            color: var(--navy-deep);
        }}

        /* Signatures: EXACTLY "Exam Head" & "Section Head" */
        .cert-footer-grid {{
            display: flex;
            justify-content: space-between;
            align-items: flex-end;
            margin-top: 18px;
            padding-top: 14px;
            border-top: 1px dashed #cbd5e1;
            position: relative;
            z-index: 1;
            gap: 20px;
        }}

        
        .sig-col {{
            flex: 1;
            max-width: 240px;
            text-align: center;
        }}
        .sig-img-container {{
            height: 85px;
            display: flex;
            align-items: flex-end;
            justify-content: center;
            margin-bottom: 3px;
        }}
        .sig-img {{
            max-height: 80px;
            max-width: 200px;
            object-fit: contain;
            filter: contrast(1.15) brightness(0.95);
        }}
        .sig-rule {{
            width: 85%;
            margin: 0 auto 5px;
            border-top: 1px solid #64748b;
        }}
        .sig-name {{
            font-size: 0.86rem;
            font-weight: 700;
            color: var(--navy-deep);
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}
        .sig-sub {{
            font-size: 0.72rem;
            font-weight: 600;
            color: var(--text-muted);
            margin-top: 1px;
        }}

        .official-seal {{
            width: 72px;
            height: 72px;
            border: 2px dashed #b91c1c;
            border-radius: 50%;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            color: #b91c1c;
            transform: rotate(-8deg);
            padding: 3px;
            user-select: none;
            margin: 0 auto;
        }}
        .seal-txt-top {{ font-size: 0.5rem; font-weight: 800; letter-spacing: 0.5px; text-transform: uppercase; }}
        .seal-txt-mid {{ font-size: 0.68rem; font-weight: 900; letter-spacing: 1px; text-transform: uppercase; }}
        .seal-txt-bot {{ font-size: 0.48rem; font-weight: 700; text-transform: uppercase; }}

        /* Security Bar */
        .security-audit-bar {{
            margin-top: 14px;
            padding-top: 8px;
            border-top: 1px solid #f1f5f9;
            display: flex;
            justify-content: space-between;
            align-items: center;
            font-size: 0.7rem;
            color: #94a3b8;
            flex-wrap: wrap;
            gap: 8px;
        }}

        /* Action Tray */
        .action-tray {{
            background: #f8fafc;
            border-top: 1px solid var(--border-soft);
            padding: 14px 24px;
            display: flex;
            justify-content: flex-end;
            gap: 12px;
            flex-wrap: wrap;
        }}
        .btn-ui {{
            padding: 10px 18px;
            border-radius: 10px;
            font-size: 0.9rem;
            font-weight: 700;
            cursor: pointer;
            transition: all 0.2s;
            display: inline-flex;
            align-items: center;
            gap: 8px;
            border: none;
            font-family: inherit;
        }}
        .btn-print-doc {{
            background: linear-gradient(135deg, var(--gold), #a77717);
            color: #ffffff;
            box-shadow: 0 4px 12px rgba(198, 146, 34, 0.25);
        }}
        .btn-print-doc:hover {{
            background: linear-gradient(135deg, var(--gold-light), var(--gold));
            transform: translateY(-1px);
        }}
        .btn-whatsapp {{
            background: #25d366;
            color: #ffffff;
            box-shadow: 0 4px 12px rgba(37, 211, 102, 0.25);
        }}
        .btn-whatsapp:hover {{
            background: #1eb857;
            transform: translateY(-1px);
        }}
        .btn-new-search {{
            background: #ffffff;
            color: var(--navy-deep);
            border: 1px solid #cbd5e1;
        }}
        .btn-new-search:hover {{
            background: #e2e8f0;
        }}

        .footer-note {{
            text-align: center;
            font-size: 0.8rem;
            color: var(--text-muted);
            margin-top: auto;
        }}

        /* ========================================================
           EXECUTIVE HONOR ROLL MODAL (POPUP WINDOW)
        ======================================================== */
        .modal-backdrop {{
            display: none;
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(8, 29, 56, 0.65);
            backdrop-filter: blur(4px);
            z-index: 9999;
            align-items: center;
            justify-content: center;
            padding: 16px;
        }}

        .modal-box {{
            background: #ffffff;
            width: 100%;
            max-width: 680px;
            max-height: 90vh;
            border-radius: 20px;
            border: 2px solid var(--gold-border);
            box-shadow: 0 25px 60px rgba(8, 29, 56, 0.35);
            overflow-y: auto;
            position: relative;
            animation: modalPop 0.25s cubic-bezier(0.16, 1, 0.3, 1);
        }}
        @keyframes modalPop {{
            from {{ opacity: 0; transform: scale(0.92); }}
            to {{ opacity: 1; transform: scale(1); }}
        }}

        .modal-header {{
            background: linear-gradient(135deg, var(--navy-deep), var(--navy-card));
            color: #ffffff;
            padding: 20px 24px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            position: sticky;
            top: 0;
            z-index: 10;
        }}
        .modal-title-wrap h3 {{
            font-size: 1.25rem;
            font-weight: 800;
            letter-spacing: 0.5px;
            display: flex;
            align-items: center;
            gap: 8px;
        }}
        .modal-title-wrap p {{
            font-size: 0.8rem;
            color: rgba(255, 255, 255, 0.75);
            margin-top: 2px;
        }}
        .btn-close-modal {{
            background: rgba(255, 255, 255, 0.15);
            border: none;
            color: #ffffff;
            width: 34px;
            height: 34px;
            border-radius: 50%;
            cursor: pointer;
            font-size: 1.1rem;
            display: flex;
            align-items: center;
            justify-content: center;
            transition: all 0.2s;
        }}
        .btn-close-modal:hover {{
            background: rgba(255, 255, 255, 0.3);
            transform: scale(1.08);
        }}

        .modal-body {{
            padding: 22px 24px 26px;
        }}

        .honor-control-bar {{
            display: flex;
            flex-direction: column;
            gap: 12px;
            margin-bottom: 22px;
            background: #f8fafc;
            border: 1px solid var(--border-soft);
            border-radius: 14px;
            padding: 12px 14px;
        }}
        @media(min-width: 520px) {{
            .honor-control-bar {{
                flex-direction: row;
                align-items: center;
                justify-content: space-between;
            }}
        }}

        .honor-segmented-switch {{
            display: inline-flex;
            background: #e2e8f0;
            padding: 3px;
            border-radius: 30px;
            gap: 4px;
        }}
        .switch-opt {{
            padding: 7px 16px;
            border-radius: 20px;
            font-size: 0.82rem;
            font-weight: 700;
            border: none;
            background: transparent;
            color: var(--text-muted);
            cursor: pointer;
            transition: all 0.2s;
        }}
        .switch-opt.active {{
            background: #ffffff;
            color: var(--navy-deep);
            box-shadow: 0 2px 6px rgba(0,0,0,0.1);
        }}

        .section-select-dropdown {{
            padding: 7px 14px;
            border-radius: 10px;
            border: 1.5px solid #cbd5e1;
            font-size: 0.84rem;
            font-weight: 600;
            color: var(--navy-deep);
            outline: none;
            background: #ffffff;
            cursor: pointer;
        }}
        .section-select-dropdown:focus {{
            border-color: var(--gold);
        }}

        .podium-grid {{
            display: flex;
            flex-direction: column;
            gap: 14px;
        }}

        .podium-card {{
            border-radius: 14px;
            padding: 16px 18px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 16px;
            position: relative;
            overflow: hidden;
            transition: all 0.2s;
        }}
        .podium-card:hover {{
            transform: translateY(-2px);
            box-shadow: 0 6px 18px rgba(0,0,0,0.08);
        }}

        .podium-gold {{
            background: linear-gradient(135deg, #fffdf5, #fef7e0);
            border: 2px solid #e6cb85;
            box-shadow: 0 4px 14px rgba(198, 146, 34, 0.12);
        }}
        .podium-gold .medal-badge {{
            background: linear-gradient(135deg, #d4af37, #aa8214);
            color: #ffffff;
        }}

        .podium-silver {{
            background: linear-gradient(135deg, #f8fafc, #edf2f7);
            border: 2px solid #cbd5e1;
        }}
        .podium-silver .medal-badge {{
            background: linear-gradient(135deg, #94a3b8, #64748b);
            color: #ffffff;
        }}

        .podium-bronze {{
            background: linear-gradient(135deg, #fffaf5, #fbf0e4);
            border: 2px solid #e2c0a4;
        }}
        .podium-bronze .medal-badge {{
            background: linear-gradient(135deg, #c27838, #975317);
            color: #ffffff;
        }}

        .podium-left {{
            display: flex;
            align-items: center;
            gap: 14px;
        }}
        .medal-badge {{
            width: 48px;
            height: 48px;
            border-radius: 12px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1.4rem;
            flex-shrink: 0;
            box-shadow: 0 3px 8px rgba(0,0,0,0.15);
        }}
        .topper-details h4 {{
            font-size: 1.05rem;
            font-weight: 800;
            color: var(--navy-deep);
            line-height: 1.2;
        }}
        .topper-details p {{
            font-size: 0.76rem;
            color: var(--text-muted);
            margin-top: 3px;
        }}

        .podium-right {{
            text-align: right;
            flex-shrink: 0;
        }}
        .topper-score {{
            font-size: 1.35rem;
            font-weight: 800;
            color: var(--navy-deep);
            line-height: 1.1;
        }}
        .topper-pct {{
            font-size: 0.85rem;
            font-weight: 700;
            color: var(--gold);
        }}

        /* Print Style (Modal NEVER prints) */
        @page {{
            size: A4 portrait;
            margin: 5mm 8mm;
        }}
        @media print {{
            *, *:before, *:after {{
                -webkit-print-color-adjust: exact !important;
                print-color-adjust: exact !important;
            }}
            html, body {{
                width: 100% !important;
                height: 100% !important;
                margin: 0 !important;
                padding: 0 !important;
                background: #ffffff !important;
            }}
            .executive-header, .search-container, .action-tray, .footer-note, .modal-backdrop {{
                display: none !important;
            }}
            .wrapper {{
                max-width: 100% !important;
                padding: 0 !important;
                margin: 0 !important;
                gap: 0 !important;
            }}
            #resultView {{
                display: block !important;
                margin: 0 !important;
                padding: 0 !important;
            }}
            .transcript-card {{
                box-shadow: none !important;
                border: 1.5px solid #000000 !important;
                border-radius: 4px !important;
                page-break-inside: avoid !important;
                break-inside: avoid !important;
            }}
            .transcript-inner {{
                padding: 10px 16px 8px !important;
            }}
            .transcript-inner::before {{
                font-size: 7.5rem !important;
                opacity: 0.03 !important;
            }}
            .cert-header {{
                padding-bottom: 6px !important;
                margin-bottom: 8px !important;
            }}
            .cert-logo-center {{
                width: 44px !important;
                height: 44px !important;
                margin-bottom: 2px !important;
            }}
            .cert-college {{
                font-size: 2.1rem !important;
            }}
            .cert-doc-type {{
                font-size: 1rem !important;
                margin-top: 1px !important;
            }}
            .cert-exam-title {{
                font-size: 0.78rem !important;
                margin-top: 1px !important;
            }}
            .cert-session-line {{
                font-size: 0.68rem !important;
                margin-top: 1px !important;
            }}
            .cert-qr-box {{
                display: block !important;
                width: 62px !important;
            }}
            .qr-img {{
                width: 54px !important;
                height: 54px !important;
            }}
            .student-bio-grid {{
                padding: 7px 10px !important;
                margin-bottom: 8px !important;
                gap: 5px !important;
            }}
            .bio-label {{
                font-size: 0.62rem !important;
            }}
            .bio-val {{
                font-size: 0.85rem !important;
            }}
            .table-wrap {{
                margin-bottom: 8px !important;
            }}
            .cert-table thead th {{
                padding: 4px 6px !important;
                font-size: 0.68rem !important;
                background: #f1f5f9 !important;
                color: #000000 !important;
                border: 1px solid #000000 !important;
            }}
            .cert-table tbody td {{
                padding: 3.5px 6px !important;
                font-size: 0.75rem !important;
                border: 1px solid #d4d4d8 !important;
            }}
            .summary-dashboard {{
                gap: 5px !important;
                margin-bottom: 8px !important;
            }}
            .dash-tile {{
                padding: 5px 3px !important;
            }}
            .tile-lbl {{
                font-size: 0.6rem !important;
            }}
            .tile-val {{
                font-size: 1.05rem !important;
            }}
            .dash-tile.grand-status {{
                background: #f1f5f9 !important;
                color: #000000 !important;
                border: 1px solid #000000 !important;
            }}
            .dash-tile.grand-status .tile-lbl, .dash-tile.grand-status .tile-val {{
                color: #000000 !important;
            }}
            .analytics-row {{
                grid-template-columns: 210px 1fr !important;
                gap: 8px !important;
                margin-bottom: 10px !important;
            }}
            .bell-curve-card {{
                padding: 6px 8px !important;
                border: 1px solid #cccccc !important;
            }}
            .curve-svg-box {{
                height: 38px !important;
            }}
            .remarks-box {{
                padding: 5px 8px !important;
                font-size: 0.72rem !important;
            }}
            .correction-note-box {{
                padding: 4px 8px !important;
                font-size: 0.66rem !important;
                line-height: 1.3 !important;
            }}
            .cert-footer-grid {{
                margin-top: 6px !important;
                padding-top: 5px !important;
            }}
            .sig-img-container {{
                height: 65px !important;
                margin-bottom: 1px !important;
            }}
            .sig-img {{
                max-height: 62px !important;
                max-width: 165px !important;
            }}
            .sig-name {{
                font-size: 0.76rem !important;
            }}
            .sig-sub {{
                font-size: 0.66rem !important;
            }}
            .official-seal {{
                width: 54px !important;
                height: 54px !important;
            }}
            .security-audit-bar {{
                margin-top: 6px !important;
                padding-top: 3px !important;
                font-size: 0.62rem !important;
            }}
        }}
    </style>
</head>
<body>

    <div class="wrapper">
        <!-- 1. Executive College Header (Exact Old English Font) -->
        <header class="executive-header">
            <img src="{logo_b64}" alt="Chenab College Logo" class="college-logo">
            <h1 class="college-title old-english-heading">Chenab College Shorkot</h1>
            <div class="controller-tag">Office of the Controller of Examinations</div>
            <div>
                <div class="exam-info-badges">
                    <span><strong>Assessment:</strong> {full_exam_name}</span>
                    <span>&bull;</span>
                    <span><strong>Session:</strong> 2026-2027</span>
                    <span>&bull;</span>
                    <span><strong>Class:</strong> 7 - (G) Jades</span>
                    <span>&bull;</span>
                    <span><strong>Date:</strong> 10/16/2026</span>
                </div>
            </div>
        </header>

        <!-- 2. Clean Search Box -->
        <section class="search-container" id="searchSection">
            <h2 class="search-title">Online Student Result Portal</h2>
            <p class="search-hint">Enter your Student B-Form Number or Name to view your Detailed Marksheet.</p>

            <form class="search-form" id="searchForm" onsubmit="executeSearch(event)">
                <div class="input-group">
                    <span class="search-icon">🔍</span>
                    <input 
                        type="text" 
                        id="studentInput" 
                        class="search-input" 
                        placeholder="Enter B-Form Number (e.g. 3320354393768) or Name"
                        autocomplete="off"
                        required
                        autofocus
                    >
                </div>
                <button type="submit" class="btn-check">
                    <span>View Result</span> &rarr;
                </button>
            </form>

            <div id="errorAlert" class="error-box">
                <span>⚠️</span>
                <span id="errorMsg">Record not found. Please verify the B-Form Number.</span>
            </div>

            <!-- View Top 3 Positions Button -->
            <div class="honor-action-bar">
                <button type="button" class="btn-honor-roll" onclick="openHonorModal()">
                    <span>🏆</span> View Top 3 Position Holders (Honor Roll)
                </button>
            </div>
        </section>

        <!-- 3. Result Transcript -->
        <section id="resultView">
            <div class="transcript-card">
                <div class="transcript-inner">
                    <!-- Certificate Header with Exact Font and Split Lines -->
                    <div class="cert-header">
                        <div style="width: 76px;" class="cert-qr-box"></div>
                        <div class="cert-header-main">
                            <img src="{logo_b64}" alt="Chenab College" class="cert-logo-center">
                            <div class="cert-college old-english-heading">Chenab College Shorkot</div>
                            <div class="cert-doc-type">Detailed Marks Certificate</div>
                            <div class="cert-exam-title">{full_exam_name}</div>
                            <div class="cert-session-line">Academic Session 2026-2027</div>
                        </div>
                        <div class="cert-qr-box">
                            <img id="dynamicQr" src="" alt="Verification QR" class="qr-img">
                            <div class="qr-caption">Scan to Verify</div>
                        </div>
                    </div>

                    <!-- Student Information Grid -->
                    <div class="student-bio-grid">
                        <div class="bio-field">
                            <span class="bio-label">Student Name</span>
                            <span class="bio-val" id="dispName">-</span>
                        </div>
                        <div class="bio-field">
                            <span class="bio-label">B-Form / Student ID</span>
                            <span class="bio-val" id="dispId" style="font-family: monospace; letter-spacing: 0.5px;">-</span>
                        </div>
                        <div class="bio-field">
                            <span class="bio-label">Class & Section</span>
                            <span class="bio-val" id="dispClass">Class : 7 - (G) Jades</span>
                        </div>
                        <div class="bio-field">
                            <span class="bio-label">Exam Date</span>
                            <span class="bio-val">10/16/2026</span>
                        </div>
                        <div class="bio-field">
                            <span class="bio-label">Class Position</span>
                            <span class="bio-val highlight" id="dispClassPos">-</span>
                        </div>
                        <div class="bio-field">
                            <span class="bio-label">Section Position</span>
                            <span class="bio-val" id="dispSecPos" style="color: var(--success-green);">-</span>
                        </div>
                    </div>

                    <!-- Detailed Subject-Wise Marks & Percentages -->
                    <div class="table-wrap">
                        <table class="cert-table">
                            <thead>
                                <tr>
                                    <th style="width: 40px;" class="center">Sr.</th>
                                    <th>Subject Description</th>
                                    <th class="center" style="width: 85px;">Max Marks</th>
                                    <th class="center" style="width: 85px;">Pass Marks</th>
                                    <th class="center" style="width: 105px;">Marks Obtained</th>
                                    <th class="center" style="width: 115px;">Percentage</th>
                                    <th class="center" style="width: 70px;">Grade</th>
                                    <th class="center" style="width: 85px;">Status</th>
                                </tr>
                            </thead>
                            <tbody id="subjectTableBody">
                                <!-- Populated dynamically -->
                            </tbody>
                        </table>
                    </div>

                    <!-- Executive Summary Tiles -->
                    <div class="summary-dashboard">
                        <div class="dash-tile">
                            <div class="tile-lbl">Total Marks</div>
                            <div class="tile-val" id="dispTotalMarks">200</div>
                        </div>
                        <div class="dash-tile">
                            <div class="tile-lbl">Marks Obtained</div>
                            <div class="tile-val" id="dispObtainedMarks">-</div>
                        </div>
                        <div class="dash-tile">
                            <div class="tile-lbl">Aggregate %</div>
                            <div class="tile-val" id="dispPercentage">-%</div>
                        </div>
                        <div class="dash-tile">
                            <div class="tile-lbl">Overall Grade</div>
                            <div class="tile-val gold" id="dispGrade">-</div>
                        </div>
                        <div class="dash-tile grand-status" id="dispStatusTile">
                            <div class="tile-lbl">Result Status</div>
                            <div class="tile-val" id="dispStatus">PASS</div>
                        </div>
                    </div>

                    <!-- Side-by-Side: Bell Curve Analytics + Remarks & Notice -->
                    <div class="analytics-row">
                        <!-- 1. The Elegant Bell Curve -->
                        <div class="bell-curve-card">
                            <div class="curve-head">
                                <span class="curve-title">
                                    <span>📈</span> Class Bell Curve
                                </span>
                                <span class="curve-standing-pill" id="dispStandingPill">Top 5%</span>
                            </div>
                            <div class="curve-svg-box">
                                <svg class="curve-svg" viewBox="0 0 200 45" preserveAspectRatio="none">
                                    <defs>
                                        <linearGradient id="bellGrad" x1="0" y1="0" x2="0" y2="1">
                                            <stop offset="0%" stop-color="#c69222" stop-opacity="0.35"/>
                                            <stop offset="100%" stop-color="#081d38" stop-opacity="0.04"/>
                                        </linearGradient>
                                    </defs>
                                    <line x1="5" y1="40" x2="195" y2="40" stroke="#cbd5e1" stroke-width="1" />
                                    <path d="M 5,40 C 45,40 60,38 75,25 C 90,12 95,5 100,5 C 105,5 110,12 125,25 C 140,38 155,40 195,40 Z" fill="url(#bellGrad)" stroke="#0e355c" stroke-width="1.8" />
                                    <line x1="100" y1="5" x2="100" y2="40" stroke="#94a3b8" stroke-width="1" stroke-dasharray="2,2" />
                                    <g id="studentMarkerGroup" transform="translate(165, 0)">
                                        <line x1="0" y1="12" x2="0" y2="40" stroke="#c69222" stroke-width="1.5" stroke-dasharray="1.5,1.5" />
                                        <circle cx="0" cy="12" r="3.5" fill="#c69222" stroke="#ffffff" stroke-width="1" />
                                        <polygon points="-3,4 3,4 0,8" fill="#c69222" />
                                    </g>
                                </svg>
                            </div>
                            <div class="curve-labels">
                                <span>Low (&lt;50%)</span>
                                <span>Class Avg</span>
                                <span style="color: var(--gold); font-weight: 700;">Top (&gt;90%)</span>
                            </div>
                        </div>

                        <!-- 2. Remarks & Correction Notice -->
                        <div class="info-col">
                            <div class="remarks-box">
                                <span class="remarks-tag">Remarks:</span>
                                <span class="remarks-text" id="dispRemarks">Outstanding performance! Keep it up.</span>
                            </div>
                            <div class="correction-note-box">
                                <strong>NOTICE (E&amp;OE):</strong> In case of any discrepancy or query, students/parents are advised to contact the <strong>Examination Branch</strong> within <strong>7 days</strong> for verification &amp; correction.
                            </div>
                        </div>
                    </div>

                    <!-- Signatures: EXACTLY "Exam Head" & Dynamic "Section Head" -->
                    <div class="cert-footer-grid">
                        <div class="sig-col">
                            <div class="sig-img-container">
                                <img id="examHeadSigImg" src="" class="sig-img" alt="Exam Head Signature">
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
                            <div class="sig-name" id="secHeadName">Ms. Sajida Nawaz</div>
                            <div class="sig-sub" id="secHeadTitle">Section Head (Girls Wing 4th - 8th)</div>
                        </div>
                    </div>

                    <!-- Security Audit Line -->
                    <div class="security-audit-bar">
                        <span>🔒 Computer Generated Official DMC &bull; Chenab College Shorkot</span>
                        <span id="auditTimestamp">System Date: October 2026</span>
                    </div>
                </div>

                <!-- Action Tray -->
                <div class="action-tray">
                    <button type="button" class="btn-ui btn-new-search" onclick="backToSearch()">
                        🔍 Search Another Student
                    </button>
                    <button type="button" class="btn-ui btn-whatsapp" onclick="shareOnWhatsApp()">
                        💬 Share on WhatsApp
                    </button>
                    <button type="button" class="btn-ui btn-print-doc" onclick="window.print()">
                        🖨️ Print / Save as PDF
                    </button>
                </div>
            </div>
        </section>

        <!-- Footer -->
        <footer class="footer-note">
            &copy; 2026-2027 Chenab College Shorkot. All Rights Reserved. &bull; Official Examination Portal
        </footer>
    </div>

    <!-- ========================================================
         MODAL: TOP 3 POSITION HOLDERS (HONOR ROLL)
    ======================================================== -->
    <div id="honorModal" class="modal-backdrop" onclick="handleBackdropClick(event)">
        <div class="modal-box">
            <div class="modal-header">
                <div class="modal-title-wrap">
                    <h3>🏆 Academic Honor Roll</h3>
                    <p>Chenab College Shorkot &bull; {full_exam_name}</p>
                </div>
                <button type="button" class="btn-close-modal" onclick="closeHonorModal()">✕</button>
            </div>

            <div class="modal-body">
                <!-- Filter Toggle Bar -->
                <div class="honor-control-bar">
                    <div class="honor-segmented-switch">
                        <button type="button" class="switch-opt active" id="btnModeSection" onclick="setHonorMode('section')">
                            🏆 Section-Wise
                        </button>
                        <button type="button" class="switch-opt" id="btnModeClass" onclick="setHonorMode('class')">
                            🌐 Entire Class
                        </button>
                    </div>

                    <div id="sectionDropdownWrap">
                        <select id="sectionSelect" class="section-select-dropdown" onchange="renderHonorToppers()">
                            <!-- Populated dynamically -->
                        </select>
                    </div>
                    <div id="classDropdownWrap" style="display: none;">
                        <select id="classSelect" class="section-select-dropdown" onchange="renderHonorToppers()">
                            <!-- Populated dynamically -->
                        </select>
                    </div>
                </div>

                <!-- Podium List -->
                <div id="podiumContainer" class="podium-grid">
                    <!-- Populated dynamically -->
                </div>
            </div>
        </div>
    </div>

    <!-- Student Data and Logic -->
    <script>
        const DB = {json_str};

        const SIGNATURES = {signatures_json_str};

        function getSectionHeadInfo(className) {{
            const cls = (className || "").toUpperCase();
            
            // Check Class Level
            let is9to12 = false;
            let is4to8 = false;
            let is1to3 = false;
            
            if (/1ST\s*YEAR|2ND\s*YEAR|11TH|12TH|9TH|10TH|\b9\b|\b10\b|\b11\b|\b12\b|MATRIC|FSC|ICS|FA\b/.test(cls)) {{
                is9to12 = true;
            }} else if (/4TH|5TH|6TH|7TH|8TH|\b4\b|\b5\b|\b6\b|\b7\b|\b8\b/.test(cls)) {{
                is4to8 = true;
            }} else if (/1ST|2ND|3RD|PREP|NURSERY|KG|\b1\b|\b2\b|\b3\b/.test(cls)) {{
                is1to3 = true;
            }} else {{
                is4to8 = true;
            }}

            // Gender Category: Girls (G) vs Boys (B)
            const isGirls = cls.includes('(G)') || cls.includes('GIRLS') || cls.includes('GIRL') || cls.includes('JADES');

            if (is1to3) {{
                return {{
                    name: "Ms. Farhat Naz",
                    title: "Section Head (Classes 1st - 3rd)",
                    sig: SIGNATURES.sec_farhat_1_3
                }};
            }} else if (is4to8) {{
                if (isGirls) {{
                    return {{
                        name: "Ms. Sajida Nawaz",
                        title: "Section Head (Girls Wing 4th - 8th)",
                        sig: SIGNATURES.sec_sajida_4_8_g
                    }};
                }} else {{
                    return {{
                        name: "Mr. Abrar Hussain",
                        title: "Section Head (Boys Wing 4th - 8th)",
                        sig: SIGNATURES.sec_abbas_4_8_b
                    }};
                }}
            }} else if (is9to12) {{
                if (isGirls) {{
                    return {{
                        name: "Ms. Samreena Nawaz",
                        title: "Section Head (Girls Wing 9th - 12th)",
                        sig: SIGNATURES.sec_sambeen_9_12_g
                    }};
                }} else {{
                    return {{
                        name: "Mr. Shoukat Hayat",
                        title: "Section Head (Boys Wing 9th - 12th)",
                        sig: SIGNATURES.sec_shoukat_9_12_b
                    }};
                }}
            }}

            return {{
                name: isGirls ? "Ms. Sajida Nawaz" : "Mr. Abrar Hussain",
                title: isGirls ? "Section Head (Girls Wing)" : "Section Head (Boys Wing)",
                sig: isGirls ? SIGNATURES.sec_sajida_4_8_g : SIGNATURES.sec_abbas_4_8_b
            }};
        }}

        let currentStudent = null;
        let currentHonorMode = 'class';

        function executeSearch(e) {{
            e.preventDefault();
            let raw = document.getElementById('studentInput').value.trim().toLowerCase();
            const errorAlert = document.getElementById('errorAlert');
            const errorMsg = document.getElementById('errorMsg');

            if (!raw) return;

            const cleanRaw = raw.replace(/-/g, '');

            const record = DB.students.find(st => {{
                const cleanStId = st.id.replace(/-/g, '').toLowerCase();
                return cleanStId === cleanRaw || 
                       st.name.toLowerCase().includes(raw) ||
                       cleanStId.endsWith(cleanRaw);
            }});

            if (record) {{
                errorAlert.style.display = 'none';
                renderResult(record);
            }} else {{
                errorMsg.textContent = `No record found for "${{raw}}". Please enter a valid B-Form Number or Name.`;
                errorAlert.style.display = 'flex';
                document.getElementById('resultView').style.display = 'none';
            }}
        }}

        function renderResult(st) {{
            currentStudent = st;
            document.getElementById('dispName').textContent = st.name;
            document.getElementById('dispId').textContent = st.id;
            document.getElementById('dispClassPos').textContent = st.class_pos;
            document.getElementById('dispSecPos').textContent = st.sec_pos;
            document.getElementById('dispTotalMarks').textContent = st.total_marks;
            document.getElementById('dispObtainedMarks').textContent = st.obtained_marks;
            document.getElementById('dispPercentage').textContent = st.percentage + '%';
            document.getElementById('dispGrade').textContent = st.grade || '-';
            document.getElementById('dispRemarks').textContent = st.remarks || 'Keep up the good work.';

            // Render Signatures
            const examHeadImg = document.getElementById('examHeadSigImg');
            if (examHeadImg && SIGNATURES.exam_head_sibtain) {{
                examHeadImg.src = SIGNATURES.exam_head_sibtain;
            }}
            document.getElementById('examHeadName').textContent = 'Mr. Sibtain Haider';
            document.getElementById('examHeadTitle').textContent = 'Exam Head (Incharge Examination)';

            if (document.getElementById('dispClass')) {{
                document.getElementById('dispClass').textContent = st.class_name;
            }}

            const secHeadImg = document.getElementById('secHeadSigImg');
            const targetSig = st.sec_head_sig ? SIGNATURES[st.sec_head_sig] : (SIGNATURES.sec_sajida_4_8_g || '');
            if (secHeadImg && targetSig) {{
                secHeadImg.src = targetSig;
            }}
            document.getElementById('secHeadName').textContent = st.sec_head_name || 'Section Head';
            document.getElementById('secHeadTitle').textContent = st.sec_head_title || 'Chenab College Shorkot';
            
            const qrUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=100x100&data=' + encodeURIComponent('CHENAB COLLEGE VERIFIED: ' + st.name + ' (' + st.id + ') Marks: ' + st.obtained_marks + '/' + st.total_marks + ' (' + st.percentage + '%)');
            document.getElementById('dynamicQr').src = qrUrl;

            updateBellCurve(st);

            const statusVal = document.getElementById('dispStatus');
            const statusTile = document.getElementById('dispStatusTile');
            statusVal.textContent = st.status.toUpperCase();
            
            if (st.status.toLowerCase() === 'pass') {{
                statusTile.style.background = 'linear-gradient(135deg, #0d824d, #085a34)';
            }} else {{
                statusTile.style.background = 'linear-gradient(135deg, #c02626, #8e1818)';
            }}

            const tbody = document.getElementById('subjectTableBody');
            tbody.innerHTML = '';

            st.subjects.forEach((sub, idx) => {{
                const tr = document.createElement('tr');

                let badgeClass = 'badge-pass';
                let statusText = 'Pass';
                let pctNum = 0;

                if (sub.status === 'Leave') {{
                    badgeClass = 'badge-leave';
                    statusText = 'Leave';
                }} else if (sub.status === 'Fail') {{
                    badgeClass = 'badge-fail';
                    statusText = 'Fail';
                }}

                if (sub.percentage && sub.percentage.includes('%')) {{
                    pctNum = parseFloat(sub.percentage) || 0;
                }}

                const miniBarHtml = sub.status === 'Leave' 
                    ? '<span style="color: #94a3b8;">-</span>'
                    : `
                        <div class="pct-cell">
                            <span>${{sub.percentage}}</span>
                            <div class="mini-bar">
                                <div class="mini-bar-fill" style="width: ${{pctNum}}%;"></div>
                            </div>
                        </div>
                    `;

                tr.innerHTML = `
                    <td class="center" style="color: var(--text-muted); font-size: 0.8rem;">${{idx + 1}}</td>
                    <td style="font-weight: 700; color: var(--navy-deep);">${{sub.name}}</td>
                    <td class="center" style="font-weight: 600;">${{sub.max}}</td>
                    <td class="center" style="color: var(--text-muted);">${{sub.pass}}</td>
                    <td class="center" style="font-weight: 800; font-size: 0.98rem; color: var(--navy-deep);">${{sub.obtained}}</td>
                    <td class="center">${{miniBarHtml}}</td>
                    <td class="center" style="font-weight: 800; color: var(--gold);">${{sub.grade}}</td>
                    <td class="center"><span class="badge ${{badgeClass}}">${{statusText}}</span></td>
                `;
                tbody.appendChild(tr);
            }});

            const resultView = document.getElementById('resultView');
            resultView.style.display = 'block';
            resultView.scrollIntoView({{ behavior: 'smooth', block: 'start' }});
        }}

        function updateBellCurve(st) {{
            const marker = document.getElementById('studentMarkerGroup');
            const pill = document.getElementById('dispStandingPill');
            const secPos = parseInt(st.sec_pos);

            let xPos = 100;

            if (!isNaN(secPos)) {{
                if (secPos === 1) {{
                    pill.textContent = '🌟 1st Position';
                    pill.style.color = '#c69222';
                    xPos = 178;
                }} else if (secPos <= 3) {{
                    pill.textContent = '🏆 Top 3';
                    pill.style.color = '#c69222';
                    xPos = 162;
                }} else if (secPos <= 7) {{
                    pill.textContent = '⭐ Above Avg';
                    pill.style.color = '#0d824d';
                    xPos = 138;
                }} else if (secPos <= 15) {{
                    pill.textContent = 'Class Average';
                    pill.style.color = '#0e355c';
                    xPos = 100;
                }} else if (secPos <= 19) {{
                    pill.textContent = 'Below Avg';
                    pill.style.color = '#b45309';
                    xPos = 62;
                }} else {{
                    pill.textContent = 'Needs Attention';
                    pill.style.color = '#c02626';
                    xPos = 28;
                }}
            }} else {{
                pill.textContent = 'Class Average';
                xPos = 100;
            }}

            const distFromCenter = Math.abs(xPos - 100) / 85;
            const yPos = 5 + (distFromCenter * distFromCenter * 33);

            marker.setAttribute('transform', `translate(${{xPos.toFixed(1)}}, 0)`);
            const circle = marker.querySelector('circle');
            if (circle) circle.setAttribute('cy', yPos.toFixed(1));
            const poly = marker.querySelector('polygon');
            if (poly) poly.setAttribute('transform', `translate(0, ${{yPos - 10}})`);
        }}

        function backToSearch() {{
            document.getElementById('studentInput').value = '';
            document.getElementById('studentInput').focus();
            document.getElementById('resultView').style.display = 'none';
            document.getElementById('errorAlert').style.display = 'none';
            window.scrollTo({{ top: 0, behavior: 'smooth' }});
        }}

        function shareOnWhatsApp() {{
            if (!currentStudent) return;
            const text = `🎓 *Chenab College Shorkot - Result Card*%0A` +
                         `*Exam:* ${{full_exam_name}}%0A` +
                         `*Student:* ${{currentStudent.name}}%0A` +
                         `*B-Form:* ${{currentStudent.id}}%0A` +
                         `*Marks:* ${{currentStudent.obtained_marks}} / ${{currentStudent.total_marks}} (${{currentStudent.percentage}}%)%0A` +
                         `*Grade:* ${{currentStudent.grade}}%0A` +
                         `*Section Position:* ${{currentStudent.sec_pos}}%0A` +
                         `*Class Position:* ${{currentStudent.class_pos}}%0A` +
                         `*Status:* ${{currentStudent.status}}`;
            window.open('https://api.whatsapp.com/send?text=' + text, '_blank');
        }}

        /* Honor Roll Logic */
        function openHonorModal() {{
            const secDropdown = document.getElementById('sectionSelect');
            secDropdown.innerHTML = '';
            DB.sections.forEach(sec => {{
                const opt = document.createElement('option');
                opt.value = sec;
                opt.textContent = sec;
                secDropdown.appendChild(opt);
            }});

            const clsDropdown = document.getElementById('classSelect');
            if (clsDropdown) {{
                clsDropdown.innerHTML = '';
                DB.grade_groups.forEach(grp => {{
                    const opt = document.createElement('option');
                    opt.value = grp;
                    opt.textContent = grp;
                    clsDropdown.appendChild(opt);
                }});
            }}

            if (currentStudent) {{
                secDropdown.value = currentStudent.class_name;
                if (clsDropdown) clsDropdown.value = currentStudent.grade_group;
            }}

            renderHonorToppers();
            document.getElementById('honorModal').style.display = 'flex';
        }}

        function closeHonorModal() {{
            document.getElementById('honorModal').style.display = 'none';
        }}

        function handleBackdropClick(e) {{
            if (e.target.id === 'honorModal') {{
                closeHonorModal();
            }}
        }}

        function setHonorMode(mode) {{
            currentHonorMode = mode;
            const btnClass = document.getElementById('btnModeClass');
            const btnSection = document.getElementById('btnModeSection');
            const secWrap = document.getElementById('sectionDropdownWrap');
            const clsWrap = document.getElementById('classDropdownWrap');

            if (mode === 'class') {{
                btnClass.classList.add('active');
                btnSection.classList.remove('active');
                if (secWrap) secWrap.style.display = 'none';
                if (clsWrap) clsWrap.style.display = 'block';
            }} else {{
                btnSection.classList.add('active');
                btnClass.classList.remove('active');
                if (secWrap) secWrap.style.display = 'block';
                if (clsWrap) clsWrap.style.display = 'none';
            }}

            renderHonorToppers();
        }}

        function renderHonorToppers() {{
            const container = document.getElementById('podiumContainer');
            container.innerHTML = '';

            let list = [...DB.students];

            if (currentHonorMode === 'section') {{
                const selectedSec = document.getElementById('sectionSelect').value;
                if (selectedSec) {{
                    list = list.filter(st => st.class_name === selectedSec);
                }}
            }} else {{
                const clsDropdown = document.getElementById('classSelect');
                const selectedClass = clsDropdown ? clsDropdown.value : '';
                if (selectedClass) {{
                    list = list.filter(st => st.grade_group === selectedClass);
                }}
            }}

            const posKey = (currentHonorMode === 'class') ? 'class_pos' : 'sec_pos';

            let p1Students = list.filter(st => String(st[posKey]).trim() === '1');
            let p2Students = list.filter(st => String(st[posKey]).trim() === '2');
            let p3Students = list.filter(st => String(st[posKey]).trim() === '3');

            // Fallback if class_pos is not numbered 1, 2, 3
            if (p1Students.length === 0 && p2Students.length === 0 && p3Students.length === 0 && list.length > 0) {{
                const sorted = [...list].sort((a, b) => parseFloat(b.percentage || 0) - parseFloat(a.percentage || 0));
                p1Students = sorted.slice(0, 1);
                p2Students = sorted.slice(1, 2);
                p3Students = sorted.slice(2, 3);
            }}

            const tiers = [
                {{ pos: 1, title: '1st Position', badge: '🥇', classModifier: 'podium-gold', students: p1Students }},
                {{ pos: 2, title: '2nd Position', badge: '🥈', classModifier: 'podium-silver', students: p2Students }},
                {{ pos: 3, title: '3rd Position', badge: '🥉', classModifier: 'podium-bronze', students: p3Students }}
            ];

            tiers.forEach(tier => {{
                if (tier.students.length === 0) return;

                tier.students.forEach(st => {{
                    const card = document.createElement('div');
                    card.className = `podium-card ${{tier.classModifier}}`;
                    card.style.cursor = 'pointer';
                    card.title = 'Click to view full result';
                    card.onclick = () => {{
                        closeHonorModal();
                        document.getElementById('studentInput').value = st.id;
                        renderResult(st);
                    }};

                    card.innerHTML = `
                        <div class="podium-left">
                            <div class="medal-badge">${{tier.badge}}</div>
                            <div class="topper-details">
                                <h4>${{st.name}}</h4>
                                <p>${{tier.title}} &bull; ${{st.class_name}}</p>
                            </div>
                        </div>
                        <div class="podium-right">
                            <div class="topper-score">${{st.obtained_marks}} <span style="font-size: 0.85rem; color: #64748b;">/ ${{st.total_marks}}</span></div>
                            <div class="topper-pct">${{st.percentage}}% (Grade ${{st.grade}})</div>
                        </div>
                    `;
                    container.appendChild(card);
                }});
            }});
        }}

        document.addEventListener('keydown', (e) => {{
            if (e.key === 'Escape') closeHonorModal();
        }});
    </script>
</body>
</html>
"""

portal_index_path = os.path.join(portal_folder, "index.html")
with open(portal_index_path, "w", encoding="utf-8") as f:
    f.write(html_code)

direct_desktop_file = os.path.join(desktop, "ResultPortal_ChenabCollege.html")
with open(direct_desktop_file, "w", encoding="utf-8") as f:
    f.write(html_code)

print("Updated: Principal replaced with Section Head!")
