import os
import xlrd
import re

target_dir = r"C:\Users\IT_OFFICE02\Desktop\Monthly Test October Exel Lists 2026"
files = sorted(os.listdir(target_dir))

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

counts = {}
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
    
    name, title, sig = get_section_head(cls_name, fname)
    counts[name] = counts.get(name, 0) + 1

print("SECTION HEAD SHEET DISTRIBUTION:")
for k, v in sorted(counts.items()):
    print(f"  {k:20s}: {v} classes")
