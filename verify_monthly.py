import pandas as pd, warnings
warnings.filterwarnings("ignore")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (3)_FIXED.xlsx"

sales_raw = pd.read_excel(path, sheet_name="Sales Raw Data 2026", header=1)
for c in ["Sales Amount", "Return", "Discount", "Discount Value", "QTY"]:
    sales_raw[c] = pd.to_numeric(sales_raw[c], errors="coerce").fillna(0)
sales_raw["Month"] = pd.to_numeric(sales_raw["Month"], errors="coerce")
sales_raw = sales_raw[sales_raw["Month"].notna() & sales_raw["Month"].between(1, 12)]
sales_raw["Month"] = sales_raw["Month"].astype(int)

gs = sales_raw.groupby("Month")["Sales Amount"].sum()
total_gs = float(gs.sum())

print("Monthly Gross Sales (using Excel Month column):")
for m in range(1, 13):
    amt = float(gs.get(m, 0))
    pct = amt / total_gs * 100 if total_gs > 0 else 0
    print("  Month {}: ${:,.2f} ({:.1f}%)".format(m, amt, pct))
print("  Total: ${:,.2f}".format(total_gs))
print("  Row count: {}".format(len(sales_raw)))
