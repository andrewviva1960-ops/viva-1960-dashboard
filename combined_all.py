import pandas as pd, warnings, sys, numpy as np
sys.stdout.reconfigure(encoding="utf-8")
warnings.filterwarnings("ignore")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (3)_FIXED.xlsx"

# 1. Sales Month column lookup
sales = pd.read_excel(path, sheet_name="Sales Raw Data 2026", header=1)
sales["Month_num"] = pd.to_numeric(sales["Month"], errors="coerce")
sales["date_str"] = sales["Date"].astype(str)
sales_lookup = sales.dropna(subset=["Month_num"]).groupby("date_str")["Month_num"].first().to_dict()

# 2. Description month lookup
exp = pd.read_excel(path, sheet_name="Expenses Raw Data 2026", header=0)
exp["amount"] = pd.to_numeric(exp["amount"], errors="coerce").fillna(0)

ar_to_en = {"يناير":1,"فبراير":2,"مارس":3,"أبريل":4,"إبريل":4,"ابريل":4,"مايو":5,"يونيو":6,"يوليو":7,"أغسطس":8,"اغسطس":8,"سبتمبر":9,"أكتوبر":10,"نوفمبر":11,"ديسمبر":12}
def find_months(desc):
    desc = str(desc)
    found = set()
    for ar, en in ar_to_en.items():
        if ar in desc:
            found.add(en)
    return found

desc_map = {}
for idx, row in exp.iterrows():
    months = find_months(row["description"])
    d = str(row["date"])
    if len(months) == 1:
        m = list(months)[0]
        if d not in desc_map:
            desc_map[d] = m

# 3. Cash Collection Month column
cc = pd.read_excel(path, sheet_name="Cash Collection Raw Data")
cc_lookup = {}
if "Month" in cc.columns:
    cc["date_str"] = cc["date"].astype(str)
    cc_lookup = cc.dropna(subset=["Month"]).groupby("date_str")["Month"].first().to_dict()

# Combined (priority: desc > sales > cc)
combined = {}
combined.update(cc_lookup)
combined.update(sales_lookup)
combined.update(desc_map)

print(f"Combined lookup: {len(combined)} unique dates")

# Apply to expenses
exp["sn"] = pd.to_numeric(exp["date"], errors="coerce")
exp["date_str"] = exp["date"].astype(str)
exp["Month"] = exp["date_str"].map(combined)

# For unmapped, use nearest neighbor interpolation
mapped = exp["Month"].notna()
unmapped = ~mapped

if unmapped.any():
    mapped_dates = sorted([int(k) for k, v in combined.items() if k.lstrip("-").isdigit()])
    mapped_months = [combined[str(d)] for d in mapped_dates]
    xs = np.array(mapped_dates)
    ys = np.array(mapped_months)
    
    for idx in exp[unmapped].index:
        sn = exp.loc[idx, "sn"]
        if pd.notna(sn):
            nearest = xs[np.abs(xs - sn).argmin()]
            exp.loc[idx, "Month"] = ys[np.abs(xs - sn).argmin()]

exp["Month"] = pd.to_numeric(exp["Month"], errors="coerce").clip(1, 12).astype(int)

TOTAL_EXPENSES = 8226344.0
raw = exp.groupby("Month")["amount"].sum()
raw_total = float(raw.sum())

print(f"\nMonthly expenses (combined lookup + NN):")
for m in sorted(raw.index):
    amt = float(raw.get(m, 0))
    scaled = TOTAL_EXPENSES * (amt / raw_total) if raw_total > 0 else 0
    print(f"  Month {m}: raw=${amt:,.0f} scaled=${scaled:,.0f}")
print(f"  Total: ${raw_total:,.0f}")

# Show per month detail
mapped_count = mapped.sum()
unmapped_count = unmapped.sum()
print(f"\nMapped directly: {mapped_count} rows (${exp.loc[mapped,'amount'].sum():,.0f})")
print(f"Interpolated: {unmapped_count} rows (${exp.loc[unmapped,'amount'].sum():,.0f})")
