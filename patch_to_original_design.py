import os
import sys
import base64
import json
import xlrd
import re

# Read original build_desktop_portal.py
orig_path = r"C:\Users\IT_OFFICE02\Desktop\Chenab_Result_Portal\build_desktop_portal.py"
with open(orig_path, "r", encoding="utf-8") as f:
    orig_code = f.read()

# Everything before idx_excel is imports, logo, font, signatures
idx_excel = orig_code.find('excel_path = os.path.join(desktop, "rptConsolidatedResultSheet.xls")')
header_part = orig_code[:idx_excel]

# Everything from idx_html is the exact HTML template
idx_html = orig_code.find('html_code = f"""<!DOCTYPE html>')
template_part = orig_code[idx_html:]

# 1. Multi-file data loading code to replace lines between idx_excel and idx_html
data_loader_code = '''
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
    if re.search(r'\\b1\\b|\\b2\\b|\\b3\\b|1ST|2ND|3RD', cls) or re.search(r'^[123]\\s', fn):
        if not re.search(r'10|11|12', cls) and not re.search(r'10|11|12', fn):
            return "Ms. Farhat Naz", "Section Head (Junior Wing 1st - 3rd)", "sec_farhat_1_3"
            
    # 4. 4th to 8th
    if re.search(r'\\b[4-8]\\b|[4-8]TH', cls) or re.search(r'^[4-8]TH', fn):
        is_girls = "(G)" in cls or "GIRL" in cls or any(g in fn.lower() for g in ["goldies", "silver", "dimonds", "sphiier", "pearls", "rubies", "corals", "jades", "amber", "opeals"])
        if is_girls:
            return "Ms. Sajida Nawaz", "Section Head (Girls Wing 4th - 8th)", "sec_sajida_4_8_g"
        else:
            return "Mr. Abrar Hussain", "Section Head (Boys Wing 4th - 8th)", "sec_abbas_4_8_b"
            
    # 5. 9th to 12th
    if re.search(r'\\b9\\b|\\b10\\b|\\b11\\b|\\b12\\b|9TH|10TH|11TH|12TH', cls) or re.search(r'^(9|10|11|12)TH', fn):
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
        
    r_cols = [str(sheet.cell_value(header_row, c)).strip().replace("\\n", " ") for c in range(sheet.ncols)]
    
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
'''

# 2. Update template part:
# A. Increase signature size in screen CSS:
old_sig_css = '''.sig-img-container {{
            height: 48px;
            display: flex;
            align-items: flex-end;
            justify-content: center;
            margin-bottom: 2px;
        }}
        .sig-img {{
            max-height: 45px;
            max-width: 140px;
            object-fit: contain;
            filter: contrast(1.08);
        }}'''

new_sig_css = '''.sig-img-container {{
            height: 82px;
            display: flex;
            align-items: flex-end;
            justify-content: center;
            margin-bottom: 3px;
        }}
        .sig-img {{
            max-height: 78px;
            max-width: 195px;
            object-fit: contain;
            filter: contrast(1.15) brightness(0.95);
        }}'''

if old_sig_css in template_part:
    template_part = template_part.replace(old_sig_css, new_sig_css)
    print("Enlarged screen signature size.")
else:
    print("Warning: old_sig_css not found directly, checking regex...")

# B. Increase signature size in print CSS:
old_print_sig = '''.sig-img-container {{
                height: 38px !important;
                margin-bottom: 1px !important;
            }}
            .sig-img {{
                max-height: 36px !important;
                max-width: 110px !important;
            }}'''

new_print_sig = '''.sig-img-container {{
                height: 60px !important;
                margin-bottom: 1px !important;
            }}
            .sig-img {{
                max-height: 58px !important;
                max-width: 155px !important;
            }}'''

if old_print_sig in template_part:
    template_part = template_part.replace(old_print_sig, new_print_sig)
    print("Enlarged print signature size.")
else:
    print("Warning: old_print_sig not found directly!")

# C. Make Class & Section dynamic in student bio grid:
old_class_span = '<div class="bio-field">\n                            <span class="bio-label">Class & Section</span>\n                            <span class="bio-val">Class : 7 - (G) Jades</span>\n                        </div>'
new_class_span = '<div class="bio-field">\n                            <span class="bio-label">Class & Section</span>\n                            <span class="bio-val" id="dispClass">Class : 7 - (G) Jades</span>\n                        </div>'

if old_class_span in template_part:
    template_part = template_part.replace(old_class_span, new_class_span)
    print("Made Class & Section dynamic.")
else:
    # try single line replace
    template_part = template_part.replace('<span class="bio-val">Class : 7 - (G) Jades</span>', '<span class="bio-val" id="dispClass">Class : 7 - (G) Jades</span>')
    print("Replaced Class & Section via inline.")

# D. Update renderResult to set dispClass and dynamic section head:
old_render_sig = '''document.getElementById('examHeadName').textContent = 'Mr. Sibtain Haider';
            document.getElementById('examHeadTitle').textContent = 'Exam Head (Incharge Examination)';

            const secInfo = getSectionHeadInfo(st.class_name);
            const secHeadImg = document.getElementById('secHeadSigImg');
            if (secHeadImg && secInfo.sig) {{
                secHeadImg.src = secInfo.sig;
            }}
            document.getElementById('secHeadName').textContent = secInfo.name;
            document.getElementById('secHeadTitle').textContent = secInfo.title;'''

new_render_sig = '''document.getElementById('examHeadName').textContent = 'Mr. Sibtain Haider';
            document.getElementById('examHeadTitle').textContent = 'Exam Head (Incharge Examination)';

            if (document.getElementById('dispClass')) {
                document.getElementById('dispClass').textContent = st.class_name;
            }

            const secHeadImg = document.getElementById('secHeadSigImg');
            const targetSig = st.sec_head_sig ? SIGNATURES[st.sec_head_sig] : (SIGNATURES.sec_sajida_4_8_g || '');
            if (secHeadImg && targetSig) {
                secHeadImg.src = targetSig;
            }
            document.getElementById('secHeadName').textContent = st.sec_head_name || 'Section Head';
            document.getElementById('secHeadTitle').textContent = st.sec_head_title || 'Chenab College Shorkot';'''

if old_render_sig in template_part:
    template_part = template_part.replace(old_render_sig, new_render_sig)
    print("Updated renderResult signatures.")
else:
    print("Warning: old_render_sig not matched!")

# Assemble the complete new build_desktop_portal.py
new_full_code = header_part + data_loader_code + template_part

with open(orig_path, "w", encoding="utf-8") as f:
    f.write(new_full_code)

print(f"Updated {orig_path} successfully ({len(new_full_code)} chars).")
