import json
import os

html_path = r"C:\Users\IT_OFFICE02\Desktop\ResultPortal_ChenabCollege.html"
data_path = r"C:\Users\IT_OFFICE02\Desktop\Chenab_Result_Portal\data.json"

with open(data_path, "r", encoding="utf-8") as f:
    db = json.load(f)

print("DATABASE VERIFICATION:")
print("Total students in data.json:", len(db["students"]))
print("Total sections in data.json:", len(db["sections"]))
assert len(db["students"]) == 2272, f"Expected 2272, got {len(db['students'])}"
assert len(db["sections"]) == 67, f"Expected 67, got {len(db['sections'])}"

# Check students from 11th Faiz
faiz_students = [s for s in db["students"] if "Faiz" in s["class_name"]]
print(f"11th Faiz students found: {len(faiz_students)} | Head: {faiz_students[0]['sec_head_name']}")
assert len(faiz_students) == 37
assert faiz_students[0]["sec_head_name"] == "Mr. Shoukat Hayat"

# Check students from 12th Sir Syed
sir_syed_students = [s for s in db["students"] if "Sir Syed" in s["class_name"]]
print(f"12th Sir Syed students found: {len(sir_syed_students)} | Head: {sir_syed_students[0]['sec_head_name']}")
assert len(sir_syed_students) == 33
assert sir_syed_students[0]["sec_head_name"] == "Mr. Shoukat Hayat"

# Check HTML file size and design
with open(html_path, "r", encoding="utf-8") as f:
    html = f.read()

print("\nHTML FILE VERIFICATION:")
print("HTML Size:", len(html), "chars")
assert "old-english-heading" in html
assert "transcript-card" in html
assert "cert-table" in html
assert "bell-curve-card" in html
assert "official-seal" in html
assert "height: 85px" in html
assert "max-height: 80px" in html
assert '"total_students": 2272' in html

print("ALL VERIFICATIONS 100% PASSED!")
