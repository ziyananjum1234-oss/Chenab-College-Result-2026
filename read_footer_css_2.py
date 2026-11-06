with open(r"C:\Users\IT_OFFICE02\Desktop\Chenab_Result_Portal\build_desktop_portal.py", "r", encoding="utf-8") as f:
    lines = f.readlines()

for i in range(890, 925):
    safe_line = lines[i].encode("ascii", "replace").decode("ascii")
    print(f"{i+1:4d}: {safe_line}", end="")
