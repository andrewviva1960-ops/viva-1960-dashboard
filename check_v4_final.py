import pandas as pd, warnings
warnings.filterwarnings("ignore")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (4).xlsx"
exp = pd.read_excel(path, sheet_name="Expenses Raw Data 2026", header=0)
exp["amount"] = pd.to_numeric(exp["amount"], errors="coerce").fillna(0)
exp["Month"] = pd.to_numeric(exp["Month"], errors="coerce")
exp = exp[exp["Month"].notna() & exp["Month"].between(1, 12)]
exp["Month"] = exp["Month"].astype(int)

TOTAL_EXPENSES = 8226344.0
raw = exp.groupby("Month")["amount"].sum()
raw_total = float(raw.sum())
print("Monthly expenses (from the updated Month column):")
for m in range(1, 13):
    amt = float(raw.get(m, 0))
    scaled = TOTAL_EXPENSES * (amt / raw_total) if raw_total > 0 else 0
    if amt > 0:
        print(f"  Month {m}: ${scaled:,.0f}")
print(f"  Total: ${TOTAL_EXPENSES:,.0f}")
