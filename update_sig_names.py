import os

build_path = r"C:\Users\IT_OFFICE02\Desktop\Chenab_Result_Portal\build_desktop_portal.py"

with open(build_path, "r", encoding="utf-8") as f:
    text = f.read()

# Replace HTML default names
text = text.replace(
    '<div class="sig-name" id="examHeadName">Sir Sibtain Haider</div>',
    '<div class="sig-name" id="examHeadName">Mr. Sibtain Haider</div>'
)
text = text.replace(
    '<div class="sig-name" id="secHeadName">Mrs. Sajida</div>',
    '<div class="sig-name" id="secHeadName">Ms. Sajida Nawaz</div>'
)

# Replace in getSectionHeadInfo function
text = text.replace('"Ms. Farhat"', '"Ms. Farhat Naz"')
text = text.replace('"Mrs. Sajida"', '"Ms. Sajida Nawaz"')
text = text.replace('"Mr. Abbas"', '"Mr. Abrar Hussain"')
text = text.replace('"Mrs. Sambeen Ivans"', '"Ms. Samreena Nawaz"')
text = text.replace('"Sir Shoukat Sb"', '"Mr. Shoukat Hayat"')
text = text.replace('"Sir Shoukat"', '"Mr. Shoukat Hayat"')

# Replace in renderResult
text = text.replace(
    "document.getElementById('examHeadName').textContent = 'Sir Sibtain Haider';",
    "document.getElementById('examHeadName').textContent = 'Mr. Sibtain Haider';"
)

with open(build_path, "w", encoding="utf-8") as f:
    f.write(text)

print("build_desktop_portal.py updated with exact names.")
