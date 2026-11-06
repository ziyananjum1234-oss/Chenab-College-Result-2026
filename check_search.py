with open(r"C:\Users\IT_OFFICE02\Desktop\Chenab_Result_Portal\build_desktop_portal.py", "r", encoding="utf-8") as f:
    text = f.read()

m = text.find("function executeSearch")
print(text[m:m+800].encode("ascii", "replace").decode("ascii"))
