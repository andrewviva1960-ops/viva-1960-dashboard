import pandas as pd, warnings, sys
sys.stdout.reconfigure(encoding="utf-8")
warnings.filterwarnings("ignore")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (3)_FIXED.xlsx"

# Try reading PNL Dashboard sheet with openpyxl to get actual cell values
pnl = pd.read_excel(path, sheet_name="PNL Dashboard ", header=None)

# Row 31: Expenses Summary headers
print("=== Expenses Summary (rows 30-45, all columns) ===")
for i in range(30, 45):
    row = pnl.iloc[i].tolist()
    for j, v in enumerate(row):
        if str(v) != "nan" and str(v) != "":
            print(f"  [{i},{j}] = {repr(v)}")

print("\n=== Looking for numeric expense values in columns B-J ===")
# The expenses should have numeric values in columns B, F, I etc (alternating)
for i in range(31, 45):
    for j in range(1, 15):
        v = pnl.iloc[i, j]
        try:
            num = float(v)
            if num != num: continue  # nan check
            print(f"  [{i},{j}] = {num}")
        except:
            pass
