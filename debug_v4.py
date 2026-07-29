import pandas as pd, warnings
warnings.filterwarnings("ignore")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (4).xlsx"
exp = pd.read_excel(path, sheet_name="Expenses Raw Data 2026", header=0)
print("Month column non-null:", exp["Month"].notna().sum())
print("Month column unique:", exp["Month"].dropna().unique()[:20])
print("date type:", type(exp["date"].iloc[0]))
print("First 5 dates:", exp["date"].head().tolist())
print("First 5 Months:", exp["Month"].head().tolist())
