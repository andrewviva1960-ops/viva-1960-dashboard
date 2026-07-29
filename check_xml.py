import zipfile, xml.etree.ElementTree as ET
import sys
sys.stdout.reconfigure(encoding="utf-8")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (3)_FIXED.xlsx"
z = zipfile.ZipFile(path)

# Find sheet XML files
for name in z.namelist():
    if name.startswith("xl/worksheets/sheet") or "Expenses" in name or "expense" in name.lower():
        print(name)

# Check expense sheet
for name in z.namelist():
    if "Expenses" in name and name.endswith(".xml"):
        print("\nReading:", name)
        xml = z.read(name)
        # Just check for 'Month' or month values
        text = xml.decode("utf-8", errors="replace")
        # Count how many non-empty Month entries
        import re
        # Find rows with a 6th column value
        rows = re.findall(r'<row[^>]*>.*?</row>', text, re.DOTALL)
        found = 0
        for r in rows[:20]:
            cells = re.findall(r'<c[^>]*>(.*?)</c>', r)
            if len(cells) >= 5:
                vals = []
                for c in cells[:7]:
                    v = re.search(r'<v>(.*?)</v>', c)
                    vals.append(v.group(1) if v else "")
                # Check if Month column (index 5) has a value
                if len(vals) > 5 and vals[5]:
                    found += 1
                    print(f"  Row with Month: {vals}")
        print(f"Total rows with Month value in first 20: {found}")
        break
