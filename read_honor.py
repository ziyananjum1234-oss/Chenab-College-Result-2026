import os
import re

with open(r'C:\Users\IT_OFFICE02\Desktop\Chenab_Result_Portal\build_desktop_portal.py', 'r', encoding='utf-8') as f:
    text = f.read()

m = re.search(r'function renderHonorToppers', text)
if m:
    sub = text[m.start():m.start() + 2000]
    print(sub.encode('ascii', 'replace').decode('ascii'))
