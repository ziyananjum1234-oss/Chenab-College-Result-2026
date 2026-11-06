import os

html_path = r"C:\Users\IT_OFFICE02\Desktop\Chenab_Result_Portal\index.html"
with open(html_path, "r", encoding="utf-8") as f:
    html = f.read()

checks = [
    ("Mr. Sibtain Haider", "Mr. Sibtain Haider" in html),
    ("Mr. Abrar Hussain", "Mr. Abrar Hussain" in html),
    ("Ms. Samreena Nawaz", "Ms. Samreena Nawaz" in html),
    ("Ms. Sajida Nawaz", "Ms. Sajida Nawaz" in html),
    ("Mr. Shoukat Hayat", "Mr. Shoukat Hayat" in html),
    ("Ms. Farhat Naz", "Ms. Farhat Naz" in html),
    ("Base64 signatures embedded", "data:image/png;base64," in html)
]

for label, passed in checks:
    print(f"[{'PASS' if passed else 'FAIL'}] {label}")
