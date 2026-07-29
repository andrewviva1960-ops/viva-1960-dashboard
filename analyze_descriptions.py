import pandas as pd, warnings, sys, re
sys.stdout.reconfigure(encoding="utf-8")
warnings.filterwarnings("ignore")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (3)_FIXED.xlsx"

exp = pd.read_excel(path, sheet_name="Expenses Raw Data 2026", header=0)
exp["amount"] = pd.to_numeric(exp["amount"], errors="coerce").fillna(0)
exp["Department"] = exp["Department"].str.strip()

# Look at expense descriptions for month/year references
print("=== Descriptions with month-like patterns ===")
month_names_ar = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو",
                  "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"]
month_names_en = ["jan", "feb", "mar", "apr", "may", "jun", 
                  "jul", "aug", "sep", "oct", "nov", "dec"]

for idx, row in exp.iterrows():
    desc = str(row["description"]).lower()
    for m in month_names_ar:
        if m in desc:
            print(f"  Row {idx}: date={row['date']}, dept={row['Department']}, desc={row['description'][:60]}, amount={row['amount']}")
            break
    for m in month_names_en:
        if m in desc:
            print(f"  Row {idx}: date={row['date']}, dept={row['Department']}, desc={row['description'][:60]}, amount={row['amount']}")
            break

# Also check for patterns like "1/2026" or "1-2026" or "01/" in descriptions
print("\n=== Descriptions with date patterns (nn/nn or 2026) ===")
for idx, row in exp.iterrows():
    desc = str(row["description"])
    if re.search(r'\d{1,2}[/-]\d{1,2}[/-]?\d{0,4}', desc):
        print(f"  Row {idx}: date={row['date']}, desc={desc[:80]}, amount={row['amount']}")

# Show ALL unique departments
print("\n=== All Departments ===")
for d in sorted(exp["Department"].unique()):
    amt = exp[exp["Department"] == d]["amount"].sum()
    print(f"  {d}: ${amt:,.0f}")

# Check if there's any relationship between category and date
print("\nDate ranges by department:")
for d in sorted(exp["Department"].unique()):
    sub = exp[exp["Department"] == d]
    print(f"  {d}: min_date={int(sub['date'].min())}, max_date={int(sub['date'].max())}, n_rows={len(sub)}, total=${sub['amount'].sum():,.0f}")
