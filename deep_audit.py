import os
import xlrd
import json
import re

target_dir = r"C:\Users\IT_OFFICE02\Desktop\Monthly Test October Exel Lists 2026"
data_path = r"C:\Users\IT_OFFICE02\Desktop\Chenab_Result_Portal\data.json"

with open(data_path, "r", encoding="utf-8") as f:
    db = json.load(f)

students = db["students"]
print(f"Total students in DB: {len(students)}")

issues = []
stats = {
    "zero_subjects": 0,
    "missing_name": 0,
    "missing_id": 0,
    "invalid_pct": 0,
    "invalid_total": 0,
    "invalid_obtained": 0,
    "missing_sec_head": 0,
    "absent_count": 0,
    "leave_count": 0,
    "pass_count": 0,
    "fail_count": 0
}

# 1. Inspect each student in DB
for idx, st in enumerate(students):
    sid = st.get("id", "").strip()
    sname = st.get("name", "").strip()
    cname = st.get("class_name", "").strip()
    subs = st.get("subjects", [])
    ttl = st.get("total_marks", "")
    obt = st.get("obtained_marks", "")
    pct = st.get("percentage", "")
    grd = st.get("grade", "")
    sts = st.get("status", "")
    hname = st.get("sec_head_name", "")
    hsig = st.get("sec_head_sig", "")
    
    # ID check
    if not sid or sid == "nan":
        stats["missing_id"] += 1
        issues.append(f"Student #{idx}: Missing ID (Name: {sname}, Class: {cname})")
        
    # Name check
    if not sname or sname == "nan":
        stats["missing_name"] += 1
        issues.append(f"Student #{idx}: Missing Name (ID: {sid}, Class: {cname})")
        
    # Subjects count
    if len(subs) == 0:
        stats["zero_subjects"] += 1
        issues.append(f"Student {sid} ({sname}, {cname}): 0 subjects found!")
        
    # Marks checks
    try:
        ttl_f = float(ttl)
        if ttl_f <= 0:
            stats["invalid_total"] += 1
            issues.append(f"Student {sid} ({sname}): Total marks <= 0 ({ttl})")
    except:
        stats["invalid_total"] += 1
        issues.append(f"Student {sid} ({sname}): Total marks non-numeric ({ttl})")
        
    try:
        obt_f = float(obt)
        if obt_f < 0:
            stats["invalid_obtained"] += 1
            issues.append(f"Student {sid} ({sname}): Obtained marks < 0 ({obt})")
    except:
        stats["invalid_obtained"] += 1
        issues.append(f"Student {sid} ({sname}): Obtained marks non-numeric ({obt})")
        
    try:
        pct_f = float(pct.replace("%", ""))
        if pct_f < 0 or pct_f > 100:
            stats["invalid_pct"] += 1
            issues.append(f"Student {sid} ({sname}): Percentage out of range ({pct})")
    except:
        stats["invalid_pct"] += 1
        issues.append(f"Student {sid} ({sname}): Percentage non-numeric ({pct})")
        
    # Section head check
    if not hname or not hsig:
        stats["missing_sec_head"] += 1
        issues.append(f"Student {sid} ({sname}, {cname}): Missing section head info!")
        
    if sts.lower() == "pass":
        stats["pass_count"] += 1
    elif sts.lower() == "fail":
        stats["fail_count"] += 1

print("\n--- STATISTICAL HEALTH OF DATA ---")
for k, v in stats.items():
    print(f"  {k:20s}: {v}")

if issues:
    print(f"\n--- POTENTIAL ISSUES FOUND ({len(issues)}) ---")
    for iss in issues[:15]:
        print("  -", iss)
else:
    print("\n>>> ZERO CRITICAL ISSUES DETECTED IN STUDENT RECORDS! <<<")
