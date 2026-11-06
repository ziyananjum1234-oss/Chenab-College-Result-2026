import re

html_path = r"C:\Users\IT_OFFICE02\Desktop\ResultPortal_ChenabCollege.html"
with open(html_path, "r", encoding="utf-8") as f:
    html = f.read()

print("File size:", len(html), "chars")

# 1. Check exact original design classes
required_classes = [
    "transcript-card",
    "transcript-inner",
    "cert-header",
    "cert-logo-center",
    "old-english-heading",
    "cert-doc-type",
    "cert-exam-title",
    "cert-session-line",
    "cert-qr-box",
    "student-bio-grid",
    "cert-table",
    "summary-dashboard",
    "analytics-row",
    "bell-curve-card",
    "curve-svg-box",
    "info-col",
    "remarks-box",
    "correction-note-box",
    "cert-footer-grid",
    "official-seal",
    "security-audit-bar",
    "btn-whatsapp",
    "btn-print-doc"
]

missing = [c for c in required_classes if c not in html]
if missing:
    print("MISSING CLASSES:", missing)
else:
    print("ALL original design classes are 100% present!")

# 2. Check signature sizes
sig_height = re.findall(r'\.sig-img-container\s*\{\s*height:\s*([^;]+);', html)
sig_img_h = re.findall(r'\.sig-img\s*\{\s*max-height:\s*([^;]+);', html)
print("Screen sig container height:", sig_height)
print("Screen sig img max-height:", sig_img_h)

# 3. Check data count
m_students = re.search(r'"total_students":\s*(\d+)', html)
if m_students:
    print("Total students embedded in HTML:", m_students.group(1))
else:
    print("total_students pattern not found directly, counting in DB.students...")
