import pandas as pd, warnings
warnings.filterwarnings("ignore")
from pathlib import Path
DIR = Path(r"C:\Users\Andro\Downloads\Financial Model")
path = DIR / "Viva Financial model 2026 (3)_FIXED.xlsx"

sales_raw = pd.read_excel(path, sheet_name="Sales Raw Data 2026", header=1)
for c in ["Sales Amount", "Return", "Discount", "Discount Value", "QTY"]:
    sales_raw[c] = pd.to_numeric(sales_raw[c], errors="coerce").fillna(0)
sales_raw["Discount Total"] = sales_raw["Sales Amount"] * sales_raw["Discount"]

sales_raw["Month_excel"] = pd.to_numeric(sales_raw["Month"], errors="coerce")
sales_raw["Date_num"] = pd.to_numeric(sales_raw["Date"], errors="coerce")
sales_raw["Date_dt"] = pd.to_datetime(sales_raw["Date"], format="%d/%m/%Y", errors="coerce")
mask1 = sales_raw["Date_dt"].isna() & sales_raw["Date_num"].notna()
sales_raw.loc[mask1, "Date_dt"] = pd.to_datetime(sales_raw.loc[mask1, "Date_num"], origin="1899-12-30", unit="D")
mask2 = sales_raw["Date_dt"].isna()
sales_raw.loc[mask2, "Date_dt"] = pd.to_datetime(sales_raw.loc[mask2, "Date"], format="%Y/%d/%m", errors="coerce")
sales_raw["Month_parsed"] = sales_raw["Date_dt"].dt.month
sales_raw["Month"] = sales_raw["Month_excel"].fillna(sales_raw["Month_parsed"])
sales_raw = sales_raw[sales_raw["Month"].notna() & sales_raw["Month"].between(1, 12)]
sales_raw["Month"] = sales_raw["Month"].astype(int)

gs = sales_raw.groupby("Month")["Sales Amount"].sum()
ret_raw = sales_raw["Return"].abs().groupby(sales_raw["Month"]).sum()
disc_raw = sales_raw["Discount Total"].groupby(sales_raw["Month"]).sum()

print("With actual discount from raw data:")
print(f"  Total Gross Sales: ${float(gs.sum()):,.0f}")
print(f"  Total Returns: ${float(ret_raw.sum()):,.0f}")
print(f"  Total Discount: ${float(disc_raw.sum()):,.0f}")
print(f"  Total Net Sales: ${float(gs.sum()) - float(ret_raw.sum()) - float(disc_raw.sum()):,.0f}")
print()
print("Monthly Net Sales:")
for m in range(1, 7):
    ns = float(gs.get(m,0)) - float(ret_raw.get(m,0)) - float(disc_raw.get(m,0))
    exp_raw = 0
    print(f"  Month {m}: gs={float(gs.get(m,0)):,.0f} ret={float(ret_raw.get(m,0)):,.0f} disc={float(disc_raw.get(m,0)):,.0f} ns={ns:,.0f}")
