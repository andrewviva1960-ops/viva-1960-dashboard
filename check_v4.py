import pandas as pd, warnings
warnings.filterwarnings("ignore")
for fname in ["Viva Financial model 2026 (4).xlsx", "Viva Financial model 2026 (3)_FIXED.xlsx"]:
    path = rf"C:\Users\Andro\Downloads\Financial Model\{fname}"
    try:
        exp = pd.read_excel(path, sheet_name="Expenses Raw Data 2026", header=0)
        has = pd.to_numeric(exp["Month"], errors="coerce").notna().sum()
        print(f"{fname}: {has} / {len(exp)} rows with Month value")
    except Exception as e:
        print(f"{fname}: ERROR - {e}")
