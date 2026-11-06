with open(r"C:\Users\IT_OFFICE02\Desktop\Chenab_Result_Portal\build_desktop_portal.py", "r", encoding="utf-8") as f:
    text = f.read()

import re
matches = [m.start() for m in re.finditer(r'sig-img|sig-col|cert-footer-grid', text)]
for m in matches[:6]:
    print("--- MATCH ---")
    print(text[m-50:m+250])
