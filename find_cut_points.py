import os
import re

with open(r"C:\Users\IT_OFFICE02\Desktop\Chenab_Result_Portal\build_desktop_portal.py", "r", encoding="utf-8") as f:
    code = f.read()

# Let's inspect where rptConsolidatedResultSheet.xls starts and where html_code starts
idx_excel = code.find('excel_path = os.path.join(desktop, "rptConsolidatedResultSheet.xls")')
idx_html = code.find('html_code = f"""<!DOCTYPE html>')

print(f"idx_excel: {idx_excel}")
print(f"idx_html: {idx_html}")
print("Code between idx_excel and idx_html:")
print(code[idx_excel:idx_excel + 300])
