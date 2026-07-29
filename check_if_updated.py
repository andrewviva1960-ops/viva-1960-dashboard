import pandas as pd, warnings
warnings.filterwarnings("ignore")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (3)_FIXED.xlsx"
exp = pd.read_excel(path, sheet_name="Expenses Raw Data 2026", header=0)
has_month = pd.to_numeric(exp["Month"], errors="coerce").notna().sum()
print(f"Rows with Month value: {has_month} / {len(exp)}")
if has_month > 0:
    print(f"Unique months: {sorted(exp['Month'].dropna().unique())}")
