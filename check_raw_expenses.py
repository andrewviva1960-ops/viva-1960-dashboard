import pandas as pd, warnings, sys
sys.stdout.reconfigure(encoding="utf-8")
warnings.filterwarnings("ignore")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (3)_FIXED.xlsx"

# Just show what the raw expense data gives with serial date parsing
exp = pd.read_excel(path, sheet_name="Expenses Raw Data 2026")
exp["amount"] = pd.to_numeric(exp["amount"], errors="coerce").fillna(0)
exp["Month"] = pd.to_datetime(pd.to_numeric(exp["date"], errors="coerce"), origin="1899-12-30", unit="D", errors="coerce").dt.month
exp = exp[exp["Month"].notna() & exp["Month"].between(1, 12)]
exp["Month"] = exp["Month"].astype(int)

TOTAL_EXPENSES = 8226344.0
raw = exp.groupby("Month")["amount"].sum()
raw_total = float(raw.sum())

print("Monthly expenses (serial date parsing, raw):")
for m in range(1, 13):
    amt = float(raw.get(m, 0))
    scaled = TOTAL_EXPENSES * (amt / raw_total) if raw_total > 0 else 0
    if amt > 0:
        print(f"  Month {m}: raw=${amt:,.0f} scaled=${scaled:,.0f}")
print(f"  Total raw: ${raw_total:,.0f}")
print(f"  Total scaled: ${TOTAL_EXPENSES:,.0f}")

# Also show net sales for comparison
sales = pd.read_excel(path, sheet_name="Sales Raw Data 2026", header=1)
for c in ["Sales Amount", "Return", "Discount", "Discount Value", "QTY"]:
    sales[c] = pd.to_numeric(sales[c], errors="coerce").fillna(0)
sales["Discount Total"] = sales["Sales Amount"] * sales["Discount"]
sales["Month_excel"] = pd.to_numeric(sales["Month"], errors="coerce")
sales["date_str"] = sales["Date"].astype(str)
sales["Date_num"] = pd.to_numeric(sales["Date"], errors="coerce")
sales["Date_dt"] = pd.to_datetime(sales["Date"], format="%d/%m/%Y", errors="coerce")
m1 = sales["Date_dt"].isna() & sales["Date_num"].notna()
sales.loc[m1, "Date_dt"] = pd.to_datetime(sales.loc[m1, "Date_num"], origin="1899-12-30", unit="D")
m2 = sales["Date_dt"].isna()
sales.loc[m2, "Date_dt"] = pd.to_datetime(sales.loc[m2, "Date"], format="%Y/%d/%m", errors="coerce")
sales["Month_parsed"] = sales["Date_dt"].dt.month
sales["Month"] = sales["Month_excel"].fillna(sales["Month_parsed"])
sales = sales[sales["Month"].notna() & sales["Month"].between(1, 12)]
sales["Month"] = sales["Month"].astype(int)

gs = sales.groupby("Month")["Sales Amount"].sum()
ret = sales["Return"].abs().groupby(sales["Month"]).sum()
disc = sales["Discount Total"].groupby(sales["Month"]).sum()
total_disc = 2250925.0
raw_disc = float(disc.sum())

print("\nMonthly Net Sales (for chart comparison):")
for m in range(1, 7):
    m_gs = float(gs.get(m, 0))
    m_ret = float(ret.get(m, 0))
    m_disc = total_disc * (float(disc.get(m, 0)) / raw_disc) if raw_disc > 0 else 0
    m_ns = m_gs - m_ret - m_disc
    m_exp = TOTAL_EXPENSES * (float(raw.get(m, 0)) / raw_total) if raw_total > 0 else 0
    print(f"  Month {m}: NS=${m_ns:,.0f} Exp=${m_exp:,.0f}")
