import pandas as pd, warnings, sys, re
sys.stdout.reconfigure(encoding="utf-8")
warnings.filterwarnings("ignore")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (3)_FIXED.xlsx"

exp = pd.read_excel(path, sheet_name="Expenses Raw Data 2026", header=0)
exp["amount"] = pd.to_numeric(exp["amount"], errors="coerce").fillna(0)

# Arabic to English month mapping
ar_to_en = {
    "يناير": 1, "فبراير": 2, "مارس": 3, "أبريل": 4, "إبريل": 4,
    "ابريل": 4, "مايو": 5, "يونيو": 6, "يوليو": 7, "أغسطس": 8,
    "اغسطس": 8, "سبتمبر": 9, "أكتوبر": 10, "نوفمبر": 11, "ديسمبر": 12
}

def find_month_in_desc(desc):
    desc = str(desc)
    months_found = set()
    for ar, en in ar_to_en.items():
        if ar in desc:
            months_found.add(en)
    return months_found

# For each date, find what months are referenced in descriptions
date_to_months = {}
for idx, row in exp.iterrows():
    months = find_month_in_desc(row["description"])
    d = row["date"]
    if months:
        if d not in date_to_months:
            date_to_months[d] = set()
        date_to_months[d].update(months)

# Print mapping
print("Date -> Referenced months from descriptions:")
for d in sorted(date_to_months.keys(), key=lambda x: int(x)):
    print(f"  {d} -> months {sorted(date_to_months[d])}")

# Now assign months: use description if available, else serial date
TOTAL_EXPENSES = 8226344.0
exp["desc_months"] = exp["date"].apply(lambda d: date_to_months.get(d, set()))
exp["desc_month"] = exp["desc_months"].apply(lambda s: next(iter(s)) if len(s) == 1 else (min(s) if len(s) > 1 else None))

# For entries with descriptions referencing ONE month, use that
has_desc = exp["desc_month"].notna()
print(f"\nRows with single month from description: {has_desc.sum()}")

# Compare vs serial date for those with descriptions
serials = pd.to_datetime(pd.to_numeric(exp["date"], errors="coerce"), origin="1899-12-30", unit="D", errors="coerce")
exp["serial_month"] = serials.dt.month

both = exp[has_desc].copy()
both["desc_month"] = both["desc_month"].astype(int)
both["serial_month"] = both["serial_month"].astype(int)
conflict = both[both["desc_month"] != both["serial_month"]]
print(f"Rows where desc month != serial month: {len(conflict)} / {len(both)}")
if len(conflict) > 0:
    print("\nConflicts:")
    for d in sorted(conflict["date"].unique()):
        sub = conflict[conflict["date"] == d]
        dm = int(sub["desc_month"].iloc[0])
        sm = int(sub["serial_month"].iloc[0])
        amt = sub["amount"].sum()
        descs = sub["description"].tolist()
        print(f"  Date {d}: desc->Month {dm}, serial->Month {sm}, ${amt:,.0f}")
        for desc in descs[:2]:
            print(f"    desc: {str(desc)[:60]}")

# Final monthly breakdown using description where available, serial as fallback
exp["Month"] = exp["desc_month"].fillna(exp["serial_month"])
exp = exp[exp["Month"].notna() & exp["Month"].between(1, 12)]
exp["Month"] = exp["Month"].astype(int)

raw = exp.groupby("Month")["amount"].sum()
raw_total = float(raw.sum())
print("\n=== FINAL Monthly Expenses (desc-based, serial fallback) ===")
for m in range(1, 13):
    amt = float(raw.get(m, 0))
    scaled = TOTAL_EXPENSES * (amt / raw_total) if raw_total > 0 else 0
    if amt > 0:
        print(f"  Month {m}: raw=${amt:,.0f} scaled=${scaled:,.0f}")

# Check which dates have no desc months
no_desc = exp[exp["desc_month"].isna()]
print(f"\nRows WITHOUT desc-month: {len(no_desc)}")
print(f"  Amount: ${no_desc['amount'].sum():,.0f}")
