import pandas as pd, warnings
warnings.filterwarnings("ignore")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (4).xlsx"
exp = pd.read_excel(path, sheet_name="Expenses Raw Data 2026", header=0)
exp["amount"] = pd.to_numeric(exp["amount"], errors="coerce").fillna(0)
exp["Month"] = pd.to_numeric(exp["Month"], errors="coerce")

# Sales Stores in raw data
ss = exp[exp["Department"].str.strip() == "Sales Stores"]
print(f"Sales Stores rows: {len(ss)}")
print(f"Sales Stores total: ${ss['amount'].sum():,.0f}")
print(f"Months: {sorted(ss['Month'].dropna().unique())}")
for m in range(1, 13):
    amt = ss[ss["Month"] == m]["amount"].sum()
    if amt > 0:
        print(f"  Month {m}: ${amt:,.0f}")

# Current total expense
TOTAL = 8226344.0
# If we add $10k to each of 6 months...
new_monthly = {m: 10000.0 for m in range(1, 7)}
print(f"\nAdding $10,000/month to Jan-Jun would add $60,000")
print(f"New total would be: ${TOTAL + 60000 - 16190:,.0f}")
