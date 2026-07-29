import pandas as pd, warnings, sys
sys.stdout.reconfigure(encoding="utf-8")
warnings.filterwarnings("ignore")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (3)_FIXED.xlsx"

# Build lookup from Sales
sales = pd.read_excel(path, sheet_name="Sales Raw Data 2026", header=1)
sales["Month_num"] = pd.to_numeric(sales["Month"], errors="coerce")
sales["date_str"] = sales["Date"].astype(str)
lookup = sales.dropna(subset=["Month_num"]).groupby("date_str")["Month_num"].first().to_dict()

# Load expenses
exp = pd.read_excel(path, sheet_name="Expenses Raw Data 2026")
exp["amount"] = pd.to_numeric(exp["amount"], errors="coerce").fillna(0)
exp["date_str"] = exp["date"].astype(str)
exp["date_num"] = pd.to_numeric(exp["date"], errors="coerce")

# Serial date conversion
exp["serial_dt"] = pd.to_datetime(exp["date_num"], origin="1899-12-30", unit="D", errors="coerce")
exp["serial_month"] = exp["serial_dt"].dt.month

# Lookup month
exp["lookup_month"] = exp["date_str"].map(lookup)

# Compare for rows that have both
both = exp.dropna(subset=["lookup_month", "serial_month"]).copy()
both["lookup_month"] = both["lookup_month"].astype(int)
both["serial_month"] = both["serial_month"].astype(int)
conflict = both[both["lookup_month"] != both["serial_month"]]

print("Rows with conflicting months (lookup vs serial):")
print(f"  Total: {len(conflict)} rows, ${conflict['amount'].sum():,.0f}")
if len(conflict) > 0:
    print()
    print("  Sample (all unique date values):")
    for d in sorted(conflict["date_str"].unique()):
        sub = conflict[conflict["date_str"] == d]
        ser_mo = int(sub["serial_month"].iloc[0])
        lk_mo = int(sub["lookup_month"].iloc[0])
        amt = sub["amount"].sum()
        print(f"    Date {d}: serial→Month {ser_mo}, lookup→Month {lk_mo}, amount=${amt:,.0f}")

# How many expenses have lookup vs serial only?
mapped = exp["lookup_month"].notna().sum()
unmapped = exp["lookup_month"].isna().sum()
print(f"\nMapped (had lookup): {mapped}")
print(f"Unmapped (no lookup): {unmapped}")

# For unmapped, show what serial months they get
unm = exp[exp["lookup_month"].isna()]
print(f"\nUnmapped expense serial months distribution:")
for m in range(1, 13):
    amt = unm[unm["serial_month"] == m]["amount"].sum()
    if amt > 0:
        print(f"  Serial→Month {m}: ${amt:,.0f}")

print(f"\nComparison of monthly totals:")
print(f"{'Month':<8} {'Serial-based':<18} {'Lookup-based (mapped only)':<30}")
for m in range(1, 13):
    s = exp[exp["serial_month"] == m]["amount"].sum()
    l = exp[exp["lookup_month"] == m]["amount"].sum()
    if s > 0 or l > 0:
        print(f"  Month {m:<4} ${s:<12,.0f} ${l:<12,.0f}")
