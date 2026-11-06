import os
import xlrd

target_dir = r"C:\Users\IT_OFFICE02\Desktop\Monthly Test October Exel Lists 2026"

check_files = [
    "11th Fatima.xls",
    "10th Champions.xls",
    "9th Explorer.xls",
    "12th Hafsa.xls",
    "O level Arfa Kareem.xls",
    "Hifa Class Boys.xls"
]

for fname in check_files:
    p = os.path.join(target_dir, fname)
    book = xlrd.open_workbook(p)
    sheet = book.sheet_by_index(0)
    
    # Header cols
    r7_cols = [str(sheet.cell_value(7, c)).strip().replace("\n", " ") for c in range(sheet.ncols)]
    # Max marks
    r8_max = [str(sheet.cell_value(8, c)).strip() for c in range(sheet.ncols)]
    # Pass marks
    r9_pass = [str(sheet.cell_value(9, c)).strip() for c in range(sheet.ncols)]
    
    print("=" * 80)
    print("FILE:", fname, "| Sheet rows:", sheet.nrows, "| Sheet cols:", sheet.ncols)
    print("Cols (r7):", r7_cols)
    print("Max (r8):", r8_max)
    print("Pass (r9):", r9_pass)
    
    # Check first 3 students
    for r in range(10, min(14, sheet.nrows)):
        row_vals = [f"{r7_cols[c]}='{sheet.cell_value(r, c)}'" for c in range(sheet.ncols)]
        print(f"Student Row {r:2d}: " + " | ".join(row_vals[:18]))
        if len(row_vals) > 18:
            print("         more: " + " | ".join(row_vals[18:]))
