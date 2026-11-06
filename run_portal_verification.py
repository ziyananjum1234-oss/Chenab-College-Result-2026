import os
import json

data_path = r"C:\Users\IT_OFFICE02\Desktop\Chenab_Result_Portal\data.json"
html_path = r"C:\Users\IT_OFFICE02\Desktop\ResultPortal_ChenabCollege.html"

with open(data_path, "r", encoding="utf-8") as f:
    db = json.load(f)

students = db["students"]
print(f"Total students verified in database: {len(students)}")
assert len(students) == 2202, f"Expected 2202, got {len(students)}"

# Test 1: Saher Bilal (11th Fatima - Pre-Medical)
saher = [s for s in students if "Saher Bilal" in s["name"]][0]
saher_subs = [s["name"] for s in saher["subjects"]]
print(f"\n1. Saher Bilal (11th Fatima):")
print(f"   Subjects ({len(saher_subs)}): {saher_subs}")
print(f"   Section Head: {saher['sec_head_name']} ({saher['sec_head_sig']})")
assert "Biology" in saher_subs and "Chemistry" in saher_subs
assert "Computer Science" not in saher_subs and "Mathematics" not in saher_subs
assert saher["sec_head_name"] == "Ms. Samreena Nawaz"

# Test 2: Aiman Fatima (11th Fatima - ICS / Computer)
aiman = [s for s in students if "Aiman Fatima" in s["name"]][0]
aiman_subs = [s["name"] for s in aiman["subjects"]]
print(f"\n2. Aiman Fatima (11th Fatima):")
print(f"   Subjects ({len(aiman_subs)}): {aiman_subs}")
print(f"   Section Head: {aiman['sec_head_name']} ({aiman['sec_head_sig']})")
assert "Computer Science" in aiman_subs and "Mathematics" in aiman_subs
assert "Biology" not in aiman_subs and "Chemistry" not in aiman_subs
assert aiman["sec_head_name"] == "Ms. Samreena Nawaz"

# Test 3: Hifz Boys
hifz_b = [s for s in students if "HifzulQuran(B)" in s["class_name"]][0]
hifz_b_subs = [s["name"] for s in hifz_b["subjects"]]
print(f"\n3. Hifz Boys ({hifz_b['name']}):")
print(f"   Subjects ({len(hifz_b_subs)}): {hifz_b_subs}")
print(f"   Section Head: {hifz_b['sec_head_name']} ({hifz_b['sec_head_sig']})")
assert len(hifz_b_subs) == 3
assert hifz_b["sec_head_name"] == "Mr. Abrar Hussain"

# Test 4: Hifz Girls
hifz_g = [s for s in students if "HifzulQuran(G)" in s["class_name"]][0]
hifz_g_subs = [s["name"] for s in hifz_g["subjects"]]
print(f"\n4. Hifz Girls ({hifz_g['name']}):")
print(f"   Subjects ({len(hifz_g_subs)}): {hifz_g_subs}")
print(f"   Section Head: {hifz_g['sec_head_name']} ({hifz_g['sec_head_sig']})")
assert len(hifz_g_subs) == 3
assert hifz_g["sec_head_name"] == "Ms. Sajida Nawaz"

# Test 5: O Levels & Pre-O
o_lvl = [s for s in students if "O Levels" in s["class_name"]][0]
pre_o = [s for s in students if "Pre O" in s["class_name"]][0]
print(f"\n5. Cambridge Wing:")
print(f"   O Levels: {o_lvl['name']} -> {o_lvl['sec_head_name']} ({o_lvl['sec_head_sig']})")
print(f"   Pre-O: {pre_o['name']} -> {pre_o['sec_head_name']} ({pre_o['sec_head_sig']})")
assert o_lvl["sec_head_name"] == "Ms. Sajida Nawaz"
assert pre_o["sec_head_name"] == "Ms. Sajida Nawaz"

# Test 6: Class 1 to 3 Primary
c1 = [s for s in students if "1 - Khayyam" in s["class_name"]][0]
print(f"\n6. Primary Wing (Class 1):")
print(f"   {c1['name']} -> {c1['sec_head_name']} ({c1['sec_head_sig']})")
assert c1["sec_head_name"] == "Ms. Farhat Naz"

# Test 7: 9th-12th Boys
b10 = [s for s in students if "10th  - (B) Champions" in s["class_name"]][0]
print(f"\n7. 10th Boys:")
print(f"   {b10['name']} -> {b10['sec_head_name']} ({b10['sec_head_sig']})")
assert b10["sec_head_name"] == "Mr. Shoukat Hayat"

# Test 8: 4th-8th Girls
g7 = [s for s in students if "7 - (G) Jades" in s["class_name"]][0]
print(f"\n8. 7th Girls:")
print(f"   {g7['name']} -> {g7['sec_head_name']} ({g7['sec_head_sig']})")
assert g7["sec_head_name"] == "Ms. Sajida Nawaz"

# Test 9: 4th-8th Boys
b7 = [s for s in students if "7 - (B)" in s["class_name"]][0]
print(f"\n9. 7th Boys:")
print(f"   {b7['name']} -> {b7['sec_head_name']} ({b7['sec_head_sig']})")
assert b7["sec_head_name"] == "Mr. Abrar Hussain"

# Check HTML file size and embedded data
with open(html_path, "r", encoding="utf-8") as f:
    html_content = f.read()

assert len(html_content) > 3000000, "HTML content too small"
assert "exam_head_sibtain" in html_content
assert "sec_shoukat_9_12_b" in html_content
assert "sec_sambeen_9_12_g" in html_content
assert "sec_sajida_4_8_g" in html_content
assert "sec_abbas_4_8_b" in html_content
assert "sec_farhat_1_3" in html_content

print("\nALL 9 TESTS & VERIFICATIONS PASSED 100% PERFECTLY!")
