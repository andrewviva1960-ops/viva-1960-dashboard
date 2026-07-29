import pandas as pd, warnings, sys
sys.stdout.reconfigure(encoding="utf-8")
warnings.filterwarnings("ignore")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (3)_FIXED.xlsx"

sales = pd.read_excel(path, sheet_name="Sales Raw Data 2026", header=1)
sales["Month_num"] = pd.to_numeric(sales["Month"], errors="coerce")
sales["date_str"] = sales["Date"].astype(str)

# Check if any date has multiple different Month values
multi = sales.dropna(subset=["Month_num"]).groupby("date_str")["Month_num"].nunique()
conflicts = multi[multi > 1]
print("Dates with multiple different Month values: {}".format(len(conflicts)))
if len(conflicts) > 0:
    for d in conflicts.index[:5]:
        vals = sales[sales["date_str"] == d]["Month_num"].unique()
        print(f"  {d}: months {sorted(vals)}")

# Now let me figure out the pattern
print("\n=== The serial numbers and their assigned months ===")
lookup = sales.dropna(subset=["Month_num"]).groupby("date_str")["Month_num"].first()
# Sort by numeric value
sorted_lookup = sorted([(k, int(v)) for k, v in lookup.items() if k.lstrip("-").isdigit()], key=lambda x: int(x[0]))
for d, m in sorted_lookup:
    print(f"  {d} -> month {m}")

# Show the block starts
print("\n=== Block structure ===")
bases = []
for d, m in sorted_lookup:
    sn = int(d)
    if m == 1:  # First of each block
        bases.append(sn)
        print(f"  Block starts at {sn}")

print(f"\nNumber of blocks: {len(bases)}")
print(f"Block base differences: {[bases[i+1]-bases[i] for i in range(len(bases)-1)]}")
