import pandas as pd, warnings, sys
sys.stdout.reconfigure(encoding="utf-8")
warnings.filterwarnings("ignore")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (3)_FIXED.xlsx"

# Step 1: Build lookup from Sales Month column
sales = pd.read_excel(path, sheet_name="Sales Raw Data 2026", header=1)
sales["Month_num"] = pd.to_numeric(sales["Month"], errors="coerce")
sales["date_str"] = sales["Date"].astype(str)
lookup = sales.dropna(subset=["Month_num"]).groupby("date_str")["Month_num"].first().to_dict()

# Step 2: Identify block starts from lookup (where month=1 for a serial number)
block_starts = sorted([int(k) for k, v in lookup.items() if k.lstrip("-").isdigit() and v == 1])
print("Block starts:", block_starts)

# Step 3: For any serial number, find which block it belongs to and compute month
def get_month_from_serial(serial_num):
    sn = int(serial_num)
    # Find the block start (largest block start <= sn)
    block = None
    for bs in reversed(block_starts):
        if bs <= sn:
            block = bs
            break
    if block is None:
        return None
    offset = sn - block
    # Within each block, offsets 0-4 map to months 1-5
    # For larger offsets, wrap around
    block_month = (offset % 5) + 1
    return block_month

# Step 4: Test the function against known lookup values
print("\nTesting against known lookups:")
for d, m in sorted([(k, int(v)) for k, v in lookup.items() if k.lstrip("-").isdigit()], key=lambda x: int(x[0])):
    pred = get_month_from_serial(d)
    match = "OK" if pred == m else "MISMATCH"
    if match == "MISMATCH":
        print(f"  {d}: lookup={m}, predicted={pred} {match}")

# Step 5: Apply to expenses
exp = pd.read_excel(path, sheet_name="Expenses Raw Data 2026")
exp["amount"] = pd.to_numeric(exp["amount"], errors="coerce").fillna(0)
exp["date_str"] = exp["date"].astype(str)
exp["date_num"] = pd.to_numeric(exp["date"], errors="coerce")

# Use lookup where available, else use block-based prediction
exp["month_lookup"] = exp["date_str"].map(lookup)
exp["month_block"] = exp["date_num"].apply(get_month_from_serial)
exp["Month"] = exp["month_lookup"].fillna(exp["month_block"])

exp = exp[exp["Month"].notna() & exp["Month"].between(1, 12)]
exp["Month"] = exp["Month"].astype(int)

print("\nMonthly expense breakdown (block-based prediction):")
for m in range(1, 7):
    amt = exp[exp["Month"] == m]["amount"].sum()
    print(f"  Month {m}: ${amt:,.0f}")
print(f"  Total: ${exp['amount'].sum():,.0f}")

# Check what rows weren't assigned
unassigned = exp[exp["Month"].isna()]
print(f"\nUnassigned rows: {len(unassigned)}")
