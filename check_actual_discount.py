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

print("Actual monthly discount from raw data:")
actual_disc = sales_raw.groupby("Month")["Discount Total"].sum()
total_actual = float(actual_disc.sum())
for m in range(1, 7):
    print(f"  Month {m}: ${float(actual_disc.get(m,0)):,.0f}")
print(f"  Total: ${total_actual:,.0f}")

# Compare with proportional allocation
gs = sales_raw.groupby("Month")["Sales Amount"].sum()
total_gs = float(gs.sum())
total_disc = 2250925.0
print(f"\nHard-coded total discount: ${total_disc:,.0f}")
print(f"Actual total discount: ${total_actual:,.0f}")
print(f"Difference: ${total_disc - total_actual:,.0f}")

# Compute actual Net Sales per month
ret_raw = sales_raw["Return"].abs().groupby(sales_raw["Month"]).sum()
print(f"\nActual monthly Net Sales:")
for m in range(1, 7):
    m_gs = float(gs.get(m, 0))
    m_ret = float(ret_raw.get(m, 0))
    m_disc = float(actual_disc.get(m, 0))
    m_ns = m_gs - m_ret - m_disc
    print(f"  Month {m}: gs={m_gs:,.0f} ret={m_ret:,.0f} disc={m_disc:,.0f} ns={m_ns:,.0f}")
