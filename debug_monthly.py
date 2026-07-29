import pandas as pd, warnings
warnings.filterwarnings("ignore")
from pathlib import Path
DIR = Path(r"C:\Users\Andro\Downloads\Financial Model")
path = DIR / "Viva Financial model 2026 (3)_FIXED.xlsx"

sales_raw = pd.read_excel(path, sheet_name="Sales Raw Data 2026", header=1)
exp_raw = pd.read_excel(path, sheet_name="Expenses Raw Data 2026")

for c in ["Sales Amount", "Return", "Discount", "Discount Value", "QTY"]:
    sales_raw[c] = pd.to_numeric(sales_raw[c], errors="coerce").fillna(0)
sales_raw["Discount Total"] = sales_raw["Sales Amount"] * sales_raw["Discount"]
sales_raw["Net Sales"] = sales_raw["Sales Amount"] - sales_raw["Return"].abs() - sales_raw["Discount Total"]

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

TOTAL_COGS = 31223509.0
TOTAL_EXPENSES = 8226344.0
gs = sales_raw.groupby("Month")["Sales Amount"].sum()
ns_raw = sales_raw.groupby("Month")["Net Sales"].sum()
ret_raw = sales_raw["Return"].abs().groupby(sales_raw["Month"]).sum()
total_gs = float(gs.sum())
total_disc = 2250925.0
total_ret = float(sales_raw["Return"].abs().sum())

# Compute monthly net sales properly
print("Monthly breakdown for Net Sales chart:")
for m in range(1, 13):
    m_gs = float(gs.get(m, 0))
    m_ret = float(ret_raw.get(m, 0))
    m_disc = total_disc * (m_gs / total_gs) if total_gs > 0 else 0
    m_ns = m_gs - m_ret - m_disc
    print(f"  Month {m}: gs={m_gs:,.0f} ret={m_ret:,.0f} disc={m_disc:,.0f} ns={m_ns:,.0f}")
print(f"  Total: {total_gs:,.0f}")
print(f"  NS sum: {sum(float(gs.get(m,0)) - float(ret_raw.get(m,0)) - (total_disc * (float(gs.get(m,0))/total_gs) if total_gs>0 else 0) for m in range(1,7)):,.0f}")

# Also check raw net sales per month
print("\nRaw Net Sales per month (from raw data):")
for m in range(1, 7):
    print(f"  Month {m}: {float(ns_raw.get(m,0)):,.0f}")

# Expenses
exp_raw["amount"] = pd.to_numeric(exp_raw["amount"], errors="coerce").fillna(0)
exp_raw["Date_dt"] = pd.to_datetime(exp_raw["date"], origin="1899-12-30", unit="D", errors="coerce")
exp_raw["Month_parsed"] = exp_raw["Date_dt"].dt.month
exp_raw["Month_excel"] = pd.to_numeric(exp_raw["Month"], errors="coerce")
exp_raw["Month"] = exp_raw["Month_excel"].fillna(exp_raw["Month_parsed"])
exp_raw = exp_raw[exp_raw["Month"].notna() & exp_raw["Month"].between(1, 12)]
exp_raw["Month"] = exp_raw["Month"].astype(int)
exp = exp_raw.groupby("Month")["amount"].sum()
raw_exp_total = float(exp.sum())

print(f"\nExpenses monthly (scaled to {TOTAL_EXPENSES:,.0f}):")
for m in range(1, 7):
    m_exp_raw = float(exp.get(m, 0))
    m_exp = TOTAL_EXPENSES * (m_exp_raw / raw_exp_total) if raw_exp_total > 0 else 0
    print(f"  Month {m}: raw={m_exp_raw:,.0f} scaled={m_exp:,.0f}")
