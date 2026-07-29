import pandas as pd, warnings
warnings.filterwarnings("ignore")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (4).xlsx"
exp = pd.read_excel(path, sheet_name="Expenses Raw Data 2026", header=0)
exp["amount"] = pd.to_numeric(exp["amount"], errors="coerce").fillna(0)
exp["Department"] = exp["Department"].str.strip().str.title()
exp["Month"] = pd.to_numeric(exp["Month"], errors="coerce")
exp = exp[exp["Month"].notna() & exp["Month"].between(1, 12)]
exp["Month"] = exp["Month"].astype(int)

SALES_STORES_MONTHLY = 10000.0
ss_mask = exp["Department"] == "Sales Stores"
exp_no_ss = exp[~ss_mask].copy()
exp_amt = exp_no_ss.groupby("Month")["amount"].sum()
for m in range(1, 7):
    exp_amt[m] = exp_amt.get(m, 0) + SALES_STORES_MONTHLY

print("Monthly expenses (with $10k Sales Stores):")
total = 0
for m in range(1, 13):
    amt = float(exp_amt.get(m, 0))
    total += amt
    if amt > 0:
        print(f"  Month {m}: ${amt:,.0f}")
print(f"  Total: ${total:,.0f}")
