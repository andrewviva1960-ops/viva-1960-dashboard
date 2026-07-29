import pandas as pd, warnings, sys
sys.stdout.reconfigure(encoding="utf-8")
warnings.filterwarnings("ignore")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (3)_FIXED.xlsx"

# Build date->month lookup from Sales Raw Data
sales = pd.read_excel(path, sheet_name="Sales Raw Data 2026", header=1)
sales["Month_num"] = pd.to_numeric(sales["Month"], errors="coerce")
sales["date_str"] = sales["Date"].astype(str)
lookup = sales.dropna(subset=["Month_num"]).groupby("date_str")["Month_num"].first().to_dict()
print("Date -> Month lookup from Sales: {} mappings".format(len(lookup)))

# Apply to Expenses
exp = pd.read_excel(path, sheet_name="Expenses Raw Data 2026")
exp["amount"] = pd.to_numeric(exp["amount"], errors="coerce").fillna(0)
exp["date_str"] = exp["date"].astype(str)
exp["Month"] = exp["date_str"].map(lookup)

# Check how many expenses got mapped
mapped = exp["Month"].notna().sum()
unmapped = exp["Month"].isna().sum()
print("Expenses mapped: {} rows, ${:,.0f}".format(mapped, exp.loc[exp["Month"].notna(), "amount"].sum()))
print("Expenses unmapped: {} rows, ${:,.0f}".format(unmapped, exp.loc[exp["Month"].isna(), "amount"].sum()))

# Monthly breakdown
exp["Month"] = exp["Month"].astype(int)
print("\nMonthly expense breakdown (from Sales lookup):")
for m in range(1, 13):
    amt = exp[exp["Month"] == m]["amount"].sum()
    if amt > 0:
        print("  Month {}: ${:,.0f}".format(m, amt))
print("  Total: ${:,.0f}".format(exp["amount"].sum()))

# Check unmapped dates
if unmapped > 0:
    print("\nUnmapped date values:", list(exp.loc[exp["Month"].isna(), "date_str"].unique()))
