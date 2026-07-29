import pandas as pd, warnings, sys, numpy as np
sys.stdout.reconfigure(encoding="utf-8")
warnings.filterwarnings("ignore")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (3)_FIXED.xlsx"

# Build a continuous serial->month mapping from ALL sheets
sales = pd.read_excel(path, sheet_name="Sales Raw Data 2026", header=1)
cc = pd.read_excel(path, sheet_name="Cash Collection Raw Data")

# Sales - use Month column
sales_map = sales.dropna(subset=["Month"]).copy()
sales_map["sn"] = pd.to_numeric(sales_map["Date"], errors="coerce")
sales_map["mn"] = pd.to_numeric(sales_map["Month"], errors="coerce")
sales_map = sales_map.dropna(subset=["sn", "mn"])
sales_map = sales_map.groupby("sn")["mn"].first().reset_index()

# Cash Collection - use Month column
cc_map = cc.dropna(subset=["Month"]).copy()
cc_map["sn"] = pd.to_numeric(cc_map["date"], errors="coerce")
cc_map["mn"] = pd.to_numeric(cc_map["Month"], errors="coerce")
cc_map = cc_map.dropna(subset=["sn", "mn"])
cc_map = cc_map.groupby("sn")["mn"].mean().reset_index()

# Combine
all_map = pd.concat([sales_map, cc_map]).groupby("sn")["mn"].first().reset_index().sort_values("sn")
print(f"Total mapped serial numbers: {len(all_map)}")
print(f"Serial range: {all_map['sn'].min()} - {all_map['sn'].max()}")
print(f"Month range: {all_map['mn'].min():.0f} - {all_map['mn'].max():.0f}")

# Manual nearest-neighbor interpolation
def nearest_month(sn):
    if len(all_map) == 0:
        return np.nan
    idx = (all_map["sn"] - sn).abs().idxmin()
    return all_map.loc[idx, "mn"]

# Apply to expenses
exp = pd.read_excel(path, sheet_name="Expenses Raw Data 2026")
exp["amount"] = pd.to_numeric(exp["amount"], errors="coerce").fillna(0)
exp["sn"] = pd.to_numeric(exp["date"], errors="coerce")
exp = exp.dropna(subset=["sn"])
exp["Month"] = exp["sn"].apply(nearest_month).round().clip(1, 12).astype(int)

print("\nMonthly expense breakdown (nearest neighbor):")
for m in range(1, 13):
    amt = exp[exp["Month"] == m]["amount"].sum()
    if amt > 0:
        print(f"  Month {m}: ${amt:,.0f}")
print(f"  Total: ${exp['amount'].sum():,.0f}")

# Show distribution of expense dates vs mapped serials
print(f"\nExpense serial range: {exp['sn'].min():.0f} - {exp['sn'].max():.0f}")
print(f"Mapped serials: {', '.join(all_map['sn'].astype(int).astype(str))}")
