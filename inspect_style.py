with open(r"C:\Users\IT_OFFICE02\Desktop\Chenab_Result_Portal\build_desktop_portal.py", "r", encoding="utf-8") as f:
    text = f.read()

# Let's see the section between <style> and </style>
start_style = text.find("<style>")
end_style = text.find("</style>")
print(f"Style length: {end_style - start_style} characters")

# Check where search input and honor roll modal are
start_body = text.find("<body>")
end_body = text.find("</body>")
print(f"Body length: {end_body - start_body} characters")
