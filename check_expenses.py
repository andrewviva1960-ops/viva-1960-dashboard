import pandas as pd, warnings
warnings.filterwarnings("ignore")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (3)_FIXED.xlsx"
exp = pd.read_excel(path, sheet_name="Expenses Raw Data 2026")
print("Expense columns:", list(exp.columns))
print("Has 'Month'? ", "Month" in exp.columns)
print("First 5 dates:", list(exp["date"].head()))
print("First 5 amounts:", list(exp["amount"].head()))
print("Total rows:", len(exp))
