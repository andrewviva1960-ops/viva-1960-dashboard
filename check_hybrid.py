import pandas as pd, warnings
warnings.filterwarnings("ignore")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (3)_FIXED.xlsx"

sales_raw = pd.read_excel(path, sheet_name="Sales Raw Data 2026", header=1)
for c in ["Sales Amount", "Return", "Discount", "Discount Value", "QTY"]:
    sales_raw[c] = pd.to_numeric(sales_raw[c], errors="coerce").fillna(0)

# Use Month column where available
sales_raw["Month_excel"] = pd.to_numeric(sales_raw["Month"], errors="coerce")

# Date parse the rest
sales_raw["Date_num"] = pd.to_numeric(sales_raw["Date"], errors="coerce")
sales_raw["Date_dt"] = pd.to_datetime(sales_raw["Date"], format="%d/%m/%Y", errors="coerce")
mask1 = sales_raw["Date_dt"].isna() & sales_raw["Date_num"].notna()
sales_raw.loc[mask1, "Date_dt"] = pd.to_datetime(sales_raw.loc[mask1, "Date_num"], origin="1899-12-30", unit="D")
mask2 = sales_raw["Date_dt"].isna()
sales_raw.loc[mask2, "Date_dt"] = pd.to_datetime(sales_raw.loc[mask2, "Date"], format="%Y/%d/%m", errors="coerce")
sales_raw["Month_parsed"] = sales_raw["Date_dt"].dt.month

# Combine: use Excel Month where available, else use parsed
sales_raw["Month"] = sales_raw["Month_excel"].fillna(sales_raw["Month_parsed"])

sales_raw = sales_raw[sales_raw["Month"].notna() & sales_raw["Month"].between(1, 12)]
sales_raw["Month"] = sales_raw["Month"].astype(int)

gs = sales_raw.groupby("Month")["Sales Amount"].sum()
total_gs = float(gs.sum())

print("Monthly Gross Sales (hybrid approach):")
for m in range(1, 13):
    amt = float(gs.get(m, 0))
    print("  Month {}: ${:,.2f}".format(m, amt))
print("  Total: ${:,.2f}".format(total_gs))
print("  Row count: {}".format(len(sales_raw)))

# Compare totals
print("\nSource breakdown:")
print("  Excel Month only: {}".format(sales_raw["Month_excel"].notna().sum()))
print("  Parsed only: {}".format(sales_raw["Month_excel"].isna().sum()))
