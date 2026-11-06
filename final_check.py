import os
import re

html_path = r"C:\Users\IT_OFFICE02\Desktop\ResultPortal_ChenabCollege.html"
with open(html_path, "r", encoding="utf-8") as f:
    html = f.read()

print("File size:", len(html), "chars")

# 1. Check exact original design elements
checks = [
    ("Old English Header present", "old-english-heading" in html),
    ("Cert QR Box present", "cert-qr-box" in html),
    ("Cert Table present", "cert-table" in html),
    ("Bell Curve Box present", "bell-curve-card" in html),
    ("Correction Note Box present", "correction-note-box" in html),
    ("Official Seal present", "official-seal" in html),
    ("WhatsApp Button present", "btn-whatsapp" in html),
    ("Print Document Button present", "btn-print-doc" in html),
    ("Signatures height >= 80px", "height: 85px" in html),
    ("Signatures max-height >= 75px", "max-height: 80px" in html),
    ("2,202 students embedded", '"total_students": 2202' in html),
    ("Section select present", 'id="sectionSelect"' in html),
    ("Class select present", 'id="classSelect"' in html),
    ("Mr. Sibtain Haider present", "Mr. Sibtain Haider" in html),
    ("Ms. Sajida Nawaz present", "Ms. Sajida Nawaz" in html),
    ("Mr. Abrar Hussain present", "Mr. Abrar Hussain" in html),
    ("Ms. Farhat Naz present", "Ms. Farhat Naz" in html),
    ("Ms. Samreena Nawaz present", "Ms. Samreena Nawaz" in html),
    ("Mr. Shoukat Hayat present", "Mr. Shoukat Hayat" in html)
]

for label, passed in checks:
    print(f"[{'PASS' if passed else 'FAIL'}] {label}")
