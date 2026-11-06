import os
import xlrd

target_dir = r"C:\Users\IT_OFFICE02\Desktop\Monthly Test October Exel Lists 2026"
files = sorted(os.listdir(target_dir))

print(f"Total files: {len(files)}")
classes_found = []
structures = {}

for i, f in enumerate(files):
    p = os.path.join(target_dir, f)
    try:
        book = xlrd.open_workbook(p)
        sheet = book.sheet_by_index(0)
        c_val = sheet.cell_value(3, 3) if sheet.nrows > 3 and sheet.ncols > 3 else "N/A"
        test_val = sheet.cell_value(0, 7) if sheet.nrows > 0 and sheet.ncols > 7 else "N/A"
        classes_found.append((f, str(c_val).strip(), sheet.nrows, sheet.ncols))
    except Exception as e:
        print(f"ERROR reading {f}: {e}")

for f, c_val, nrows, ncols in classes_found:
    print(f"{f:30s} | Rows: {nrows:3d} | Cols: {ncols:2d} | Class: {c_val}")
