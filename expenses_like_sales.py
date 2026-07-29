import pandas as pd, warnings, sys
sys.stdout.reconfigure(encoding="utf-8")
warnings.filterwarnings("ignore")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (3)_FIXED.xlsx"

# Build date->month lookup FROM SALES (the correct source of truth)
sales = pd.read_excel(path, sheet_name="Sales Raw Data 2026", header=1)
sales["Month_num"] = pd.to_numeric(sales["Month"], errors="coerce")
sales["date_str"] = sales["Date"].astype(str)
lookup = sales.dropna(subset=["Month_num"]).groupby("date_str")["Month_num"].first().to_dict()

# Also define block-based mapping for dates NOT in lookup
# Block starts from Sales Month=1 entries
block_starts = sorted([int(k) for k, v in lookup.items() if k.lstrip("-").isdigit() and v == 1])
print(f"Block starts: {block_starts}")

def month_from_serial(sn):
    """Get month for a serial number using block pattern"""
    for i, bs in enumerate(block_starts):
        if bs <= sn < (block_starts[i+1] if i+1 < len(block_starts) else sn+1):
            offset = sn - bs
            # Offset 0-4 in each block maps to months 1-5
            block_month = (offset % 5) + 1
            return block_month
    return None

# Apply to expenses
exp = pd.read_excel(path, sheet_name="Expenses Raw Data 2026")
exp["amount"] = pd.to_numeric(exp["amount"], errors="coerce").fillna(0)
exp["date_str"] = exp["date"].astype(str)
exp["sn"] = pd.to_numeric(exp["date"], errors="coerce")

# Use lookup where available, block-based for rest
exp["Month"] = exp["date_str"].map(lookup)
mask = exp["Month"].isna() & exp["sn"].notna()
exp.loc[mask, "Month"] = exp.loc[mask, "sn"].apply(month_from_serial)

exp = exp.dropna(subset=["Month"])
exp["Month"] = exp["Month"].astype(int)

print("\nMonthly expense breakdown (Sales lookup + block):")
for m in range(1, 7):
    amt = exp[exp["Month"] == m]["amount"].sum()
    print(f"  Month {m}: ${amt:,.0f}")
print(f"  Total: ${exp['amount'].sum():,.0f}")
