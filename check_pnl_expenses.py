import pandas as pd, warnings
warnings.filterwarnings("ignore")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (3)_FIXED.xlsx"
pnl = pd.read_excel(path, sheet_name="PNL Dashboard ", header=None)

# Find rows with "Expense" keyword
print("=== Rows containing 'Expense' or 'expense' ===")
for i in range(len(pnl)):
    row = pnl.iloc[i].tolist()
    row_str = " ".join([str(v) for v in row if str(v) != "nan" and str(v) != ""])
    if "expense" in row_str.lower():
        vals = [str(v) for v in row if str(v) != "nan" and str(v) != ""]
        print(f"  Row {i}: {vals[:10]}")

print("\n=== Also check for 'Expenses' section in rows 50-154 ===")
for i in range(50, 154):
    row = pnl.iloc[i].tolist()
    vals = [str(v) for v in row if str(v) != "nan" and str(v) != ""]
    if vals:
        print(f"  Row {i}: {vals[:10]}")
