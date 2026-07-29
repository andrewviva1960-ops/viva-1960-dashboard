import pandas as pd, warnings
warnings.filterwarnings("ignore")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (3)_FIXED.xlsx"
pnl = pd.read_excel(path, sheet_name="PNL Dashboard ", header=None)

print("=== Expenses Summary (rows 31-50) ===")
for i in range(30, 52):
    row = pnl.iloc[i].tolist()
    vals = [str(v) for v in row if str(v) != "nan" and str(v) != ""]
    if vals:
        print(f"  Row {i}: {vals[:15]}")

print("\n=== Monthly Gross Sales (rows 126-135) ===")
for i in range(125, 136):
    row = pnl.iloc[i].tolist()
    vals = [str(v) for v in row if str(v) != "nan" and str(v) != ""]
    if vals:
        print(f"  Row {i}: {vals[:12]}")

print("\n=== PNL Highlights (rows 145-154) ===")
for i in range(144, 154):
    row = pnl.iloc[i].tolist()
    vals = [str(v) for v in row if str(v) != "nan" and str(v) != ""]
    if vals:
        print(f"  Row {i}: {vals[:8]}")

print("\n=== Executive Summary (rows 109-130) ===")
for i in range(108, 131):
    row = pnl.iloc[i].tolist()
    vals = [str(v) for v in row if str(v) != "nan" and str(v) != ""]
    if vals:
        print(f"  Row {i}: {vals[:12]}")
