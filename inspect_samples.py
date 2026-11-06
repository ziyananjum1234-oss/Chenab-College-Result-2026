import os
import xlrd

target_dir = r"C:\Users\IT_OFFICE02\Desktop\Monthly Test October Exel Lists 2026"
samples = [
    "1 Khayaam.xls",
    "4th Goldies.xls",
    "7th jades.xls",
    "10th Champions.xls",
    "11th Fatima.xls",
    "12th Hafsa.xls",
    "Hifa Class Boys.xls",
    "Hifz quran Girls.xls",
    "O level Arfa Kareem.xls",
    "Pre O fatima jinah.xls"
]

for s in samples:
    p = os.path.join(target_dir, s)
    print("=" * 80)
    print("FILE:", s)
    book = xlrd.open_workbook(p)
    sheet = book.sheet_by_index(0)
    for r in range(min(sheet.nrows, 10)):
        row_vals = [f"c{c}: '{sheet.cell_value(r, c)}'" for c in range(min(sheet.ncols, 15)) if str(sheet.cell_value(r, c)).strip() != ""]
        print(f"Row {r:2d}: " + " | ".join(row_vals))
