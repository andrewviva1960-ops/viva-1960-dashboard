import pandas as pd, warnings
warnings.filterwarnings("ignore")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (3)_FIXED.xlsx"
exp = pd.read_excel(path, sheet_name="Expenses Raw Data 2026")
exp["amount"] = pd.to_numeric(exp["amount"], errors="coerce").fillna(0)

# Use Month column where available
exp["Month_excel"] = pd.to_numeric(exp["Month"], errors="coerce") if "Month" in exp.columns else float("nan")
exp["Date_dt"] = pd.to_datetime(exp["date"], origin="1899-12-30", unit="D", errors="coerce")
exp["Month_parsed"] = exp["Date_dt"].dt.month
exp["Month"] = exp["Month_excel"].fillna(exp["Month_parsed"])
exp = exp[exp["Month"].notna() & exp["Month"].between(1, 12)]
exp["Month"] = exp["Month"].astype(int)

total = exp["amount"].sum()
print("Expense total: ${:,.2f}".format(total))
for m in range(1, 13):
    amt = exp[exp["Month"] == m]["amount"].sum()
    print("  Month {}: ${:,.2f}".format(m, amt))
