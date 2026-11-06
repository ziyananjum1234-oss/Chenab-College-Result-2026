with open(r"C:\Users\IT_OFFICE02\Desktop\Chenab_Result_Portal\build_desktop_portal.py", "r", encoding="utf-8") as f:
    text = f.read()

text = text.replace("import pandas as pd", "import pandas as pd\nimport xlrd\nimport re")

with open(r"C:\Users\IT_OFFICE02\Desktop\Chenab_Result_Portal\build_desktop_portal.py", "w", encoding="utf-8") as f:
    f.write(text)

print("Added xlrd and re imports.")
