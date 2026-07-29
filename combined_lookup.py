import pandas as pd, warnings, sys
sys.stdout.reconfigure(encoding="utf-8")
warnings.filterwarnings("ignore")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (3)_FIXED.xlsx"

# Cash Collection lookup
cc = pd.read_excel(path, sheet_name="Cash Collection Raw Data")
cc["date_str"] = cc["date"].astype(str)
cc_lookup = cc.dropna(subset=["Month"]).groupby("date_str")["Month"].first().to_dict()
print(f"Cash Collection: {len(cc_lookup)} entries")

# Sales lookup
sales = pd.read_excel(path, sheet_name="Sales Raw Data 2026", header=1)
sales["Month_num"] = pd.to_numeric(sales["Month"], errors="coerce")
sales["date_str"] = sales["Date"].astype(str)
sales_lookup = sales.dropna(subset=["Month_num"]).groupby("date_str")["Month_num"].first().to_dict()
print(f"Sales: {len(sales_lookup)} entries")

# Combined (Sales overrides Cash Collection)
combined = {**cc_lookup, **sales_lookup}
print(f"Combined: {len(combined)} entries")

# Any CC entries that add new month values not in Sales?
new_from_cc = set(cc_lookup.keys()) - set(sales_lookup.keys())
print(f"CC entries NOT in Sales: {len(new_from_cc)}")
for d in sorted(new_from_cc, key=lambda x: int(x) if x.lstrip("-").isdigit() else 0):
    print(f"  {d} -> Month {int(cc_lookup[d])}")

# Apply to expenses
exp = pd.read_excel(path, sheet_name="Expenses Raw Data 2026")
exp["amount"] = pd.to_numeric(exp["amount"], errors="coerce").fillna(0)
exp["date_str"] = exp["date"].astype(str)
exp["Month"] = exp["date_str"].map(combined)
exp["Month"] = pd.to_numeric(exp["Month"], errors="coerce")

mapped = exp["Month"].notna().sum()
unmapped = exp["Month"].isna().sum()
print(f"\nExpenses mapped: {mapped} rows, ${exp.loc[exp['Month'].notna(), 'amount'].sum():,.0f}")
print(f"Expenses unmapped: {unmapped} rows, ${exp.loc[exp['Month'].isna(), 'amount'].sum():,.0f}")

if unmapped > 0:
    print(f"\nUnmapped expense dates: {list(exp.loc[exp['Month'].isna(), 'date_str'].unique()[:20])}...")

# For mapped rows, show breakdown
mapped_exp = exp[exp["Month"].notna()].copy()
mapped_exp["Month"] = mapped_exp["Month"].astype(int)
print("\nMonthly expense (mapped rows only):")
for m in range(1, 7):
    amt = mapped_exp[mapped_exp["Month"] == m]["amount"].sum()
    print(f"  Month {m}: ${amt:,.0f}")
print(f"  Total: ${mapped_exp['amount'].sum():,.0f}")
