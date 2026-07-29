import pandas as pd, warnings, sys
sys.stdout.reconfigure(encoding="utf-8")
warnings.filterwarnings("ignore")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (3)_FIXED.xlsx"

exp = pd.read_excel(path, sheet_name="Expenses Raw Data 2026", header=0)
exp["amount"] = pd.to_numeric(exp["amount"], errors="coerce").fillna(0)

ar_months = {"يناير":1,"فبراير":2,"مارس":3,"أبريل":4,"إبريل":4,"ابريل":4,"مايو":5,"يونيو":6,"يوليو":7,"أغسطس":8,"اغسطس":8,"سبتمبر":9,"أكتوبر":10,"نوفمبر":11,"ديسمبر":12}
def get_desc_months(desc):
    found = set()
    for ar, en in ar_months.items():
        if ar in str(desc):
            found.add(en)
    return found

sales = pd.read_excel(path, sheet_name="Sales Raw Data 2026", header=1)
sales["Month_excel"] = pd.to_numeric(sales["Month"], errors="coerce")
sales["date_str"] = sales["Date"].astype(str)
sales_lu = sales.dropna(subset=["Month_excel"]).groupby("date_str")["Month_excel"].first().to_dict()

exp["date_str"] = exp["date"].astype(str)
exp["Month"] = exp["date_str"].map(sales_lu)

for idx in exp.index:
    months = get_desc_months(exp.loc[idx, "description"])
    if len(months) == 1:
        exp.loc[idx, "Month"] = list(months)[0]

mask = exp["Month"].isna()
exp.loc[mask, "Month"] = pd.to_datetime(pd.to_numeric(exp.loc[mask, "date"], errors="coerce"), origin="1899-12-30", unit="D", errors="coerce").dt.month
exp = exp[exp["Month"].notna() & exp["Month"].between(1, 12)]
exp["Month"] = exp["Month"].astype(int)

TOTAL_EXPENSES = 8226344.0
raw = exp.groupby("Month")["amount"].sum()
raw_total = float(raw.sum())
print(f"Row count: {len(exp)}")

# Count sources
desc_count = 0
sales_lu_count = 0
serial_count = 0
for idx in exp.index:
    months = get_desc_months(exp.loc[idx, "description"])
    if len(months) == 1:
        desc_count += 1
    elif exp.loc[idx, "date_str"] in sales_lu:
        sales_lu_count += 1
    else:
        serial_count += 1
print(f"Sources: desc={desc_count}, sales_lu={sales_lu_count}, serial_fallback={serial_count}")

print(f"\nMonthly expenses:")
for m in range(1, 13):
    amt = float(raw.get(m, 0))
    scaled = TOTAL_EXPENSES * (amt / raw_total) if raw_total > 0 else 0
    if amt > 0:
        print(f"  Month {m}: raw=${amt:,.0f} scaled=${scaled:,.0f}")
