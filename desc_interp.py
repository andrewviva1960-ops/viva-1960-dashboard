import pandas as pd, warnings, sys, numpy as np
sys.stdout.reconfigure(encoding="utf-8")
warnings.filterwarnings("ignore")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (3)_FIXED.xlsx"

exp = pd.read_excel(path, sheet_name="Expenses Raw Data 2026", header=0)
exp["amount"] = pd.to_numeric(exp["amount"], errors="coerce").fillna(0)

# Build description-to-month mapping
ar_to_en = {
    "يناير": 1, "فبراير": 2, "مارس": 3, "أبريل": 4, "إبريل": 4,
    "ابريل": 4, "مايو": 5, "يونيو": 6, "يوليو": 7, "أغسطس": 8,
    "اغسطس": 8, "سبتمبر": 9, "أكتوبر": 10, "نوفمبر": 11, "ديسمبر": 12
}

def find_months(desc):
    desc = str(desc)
    found = set()
    for ar, en in ar_to_en.items():
        if ar in desc:
            found.add(en)
    return found

# Build date->month from descriptions (single month only)
desc_map = {}
for idx, row in exp.iterrows():
    months = find_months(row["description"])
    d = row["date"]
    if len(months) == 1:
        m = list(months)[0]
        if d not in desc_map:
            desc_map[d] = m
        # If conflict, take the first one

print(f"Dates with single-description month: {len(desc_map)}")
for d in sorted(desc_map.keys(), key=lambda x: int(x)):
    print(f"  {d} -> Month {desc_map[d]}")

# Build interpolation function
dates_sorted = sorted(desc_map.keys(), key=lambda x: int(x))
xs = np.array([int(d) for d in dates_sorted])
ys = np.array([desc_map[d] for d in dates_sorted])

# Apply to all expenses
exp["sn"] = pd.to_numeric(exp["date"], errors="coerce")
exp = exp.dropna(subset=["sn"])

# For each expense, find nearest mapped date
def nearest_month(sn):
    idx = np.abs(xs - sn).argmin()
    return ys[idx]

exp["Month"] = exp["sn"].apply(nearest_month)

TOTAL_EXPENSES = 8226344.0
raw = exp.groupby("Month")["amount"].sum()
raw_total = float(raw.sum())

print(f"\nMonthly expenses (desc-interpolated):")
for m in sorted(raw.index):
    amt = float(raw.get(m, 0))
    scaled = TOTAL_EXPENSES * (amt / raw_total) if raw_total > 0 else 0
    print(f"  Month {m}: raw=${amt:,.0f} scaled=${scaled:,.0f}")
print(f"  Total: ${raw_total:,.0f}")
