import zipfile, xml.etree.ElementTree as ET
import sys, re
sys.stdout.reconfigure(encoding="utf-8")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (3)_FIXED.xlsx"
z = zipfile.ZipFile(path)

# Read PNL Dashboard XML
xml = z.read("xl/worksheets/sheet3.xml")
text = xml.decode("utf-8", errors="replace")

# Find rows for Expenses Summary (around row 31-42)
rows_data = re.findall(r'<row[^>]*>(.*?)</row>', text, re.DOTALL)
print(f"Total rows in PNL Dashboard: {len(rows_data)}")

# Show rows 30-50 with cell values
for i, r in enumerate(rows_data):
    if i < 28 or i > 55:
        continue
    cells = re.findall(r'<c[^>]*r="([A-Z]+)(\d+)"[^>]*>(.*?)</c>', r)
    vals = []
    for col_letter, row_num, content in cells:
        v_match = re.search(r'<v>(.*?)</v>', content)
        f_match = re.search(r'<f[^>]*>(.*?)</f>', content)
        if v_match:
            vals.append(f"{col_letter}{row_num}={v_match.group(1)}")
        elif f_match:
            vals.append(f"{col_letter}{row_num}=FORMULA")
    if vals:
        print(f"  Row {i+1}: {', '.join(vals[:15])}")
