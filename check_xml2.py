import zipfile, xml.etree.ElementTree as ET
import sys, re
sys.stdout.reconfigure(encoding="utf-8")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (3)_FIXED.xlsx"
z = zipfile.ZipFile(path)

# Map sheet names to rIds
wb = ET.parse(z.open("xl/workbook.xml"))
ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
sheets = {}
for sheet in wb.findall(".//s:sheet", ns):
    sheets[sheet.get("name")] = sheet.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")

print("Sheets found:")
for name, rid in sheets.items():
    print(f"  {rid}: {name}")

# Map rIds to file paths
rels = ET.parse(z.open("xl/_rels/workbook.xml.rels"))
ns2 = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}
rid_to_file = {}
for rel in rels.findall(".//r:Relationship", ns2):
    rid_to_file[rel.get("Id")] = rel.get("Target")

print("\nSheet file paths:")
for name, rid in sheets.items():
    target = rid_to_file.get(rid, "UNKNOWN")
    full = target.replace("../", "xl/")
    if not full.startswith("xl/"):
        full = "xl/" + full
    print(f"  {rid}: {name} -> {full}")
    if name == "Expenses Raw Data 2026":
        # Read this sheet's XML
        xml = z.read(full)
        text = xml.decode("utf-8", errors="replace")
        
        # Count non-empty Month column values
        # Find all c elements with r attribute containing "F" (column F = Month)
        # Column F = 6th column
        rows_data = re.findall(r'<row[^>]*>(.*?)</row>', text, re.DOTALL)
        month_count = 0
        date_count = 0
        total = 0
        for r in rows_data:
            cells = re.findall(r'<c[^>]*r="([A-Z]+)(\d+)"[^>]*>(.*?)</c>', r)
            cell_map = {}
            for col_letter, row_num, content in cells:
                v_match = re.search(r'<v>(.*?)</v>', content)
                if v_match:
                    cell_map[col_letter] = v_match.group(1)
            if "E" in cell_map:  # date column
                date_count += 1
            if "F" in cell_map:  # Month column
                month_count += 1
            total += 1
        
        print(f"\n  Total expense rows: {total}")
        print(f"  Rows with Date (col E): {date_count}")
        print(f"  Rows with Month (col F): {month_count}")
        
        # Show some sample rows
        print("\n  Sample rows:")
        count = 0
        for r in rows_data:
            cells = re.findall(r'<c[^>]*r="([A-Z]+)(\d+)"[^>]*>(.*?)</c>', r)
            cell_map = {}
            for col_letter, row_num, content in cells:
                t_match = re.search(r't="([^"]*)"', content)
                v_match = re.search(r'<v>(.*?)</v>', content)
                vals = []
                if v_match:
                    vals.append(v_match.group(1))
                else:
                    vals.append("")
                cell_map[col_letter] = vals[0]
            e = cell_map.get("E", "-")
            f = cell_map.get("F", "-")
            print(f"    Date={e}, Month={f}")
            count += 1
            if count >= 5:
                break
