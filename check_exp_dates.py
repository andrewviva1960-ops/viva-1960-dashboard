import pandas as pd, warnings
warnings.filterwarnings("ignore")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (3)_FIXED.xlsx"
exp = pd.read_excel(path, sheet_name="Expenses Raw Data 2026")
exp["amount"] = pd.to_numeric(exp["amount"], errors="coerce").fillna(0)
exp["date_num"] = pd.to_numeric(exp["date"], errors="coerce")

# What dates do these serials correspond to?
exp["Date_dt"] = pd.to_datetime(exp["date_num"], origin="1899-12-30", unit="D", errors="coerce")
exp["Month_parsed"] = exp["Date_dt"].dt.month
exp["Year_parsed"] = exp["Date_dt"].dt.year

print("Unique serial numbers and what they parse to:")
for d in sorted(exp["date_num"].unique()):
    dt = pd.to_datetime(d, origin="1899-12-30", unit="D")
    print("  {} -> {} (Month {})".format(int(d), dt.strftime("%Y-%m-%d"), dt.month))

print("\nMonthly breakdown from serial date parsing:")
for m in range(1, 13):
    amt = exp[exp["Month_parsed"] == m]["amount"].sum()
    print("  Month {}: ${:,.2f}".format(m, amt))
print("  Total: ${:,.2f}".format(exp["amount"].sum()))

print("\nYear breakdown:")
for y in sorted(exp["Year_parsed"].dropna().unique()):
    amt = exp[exp["Year_parsed"] == y]["amount"].sum()
    print("  {}: ${:,.2f}".format(int(y), amt))

print("\nQuarters column:")
if "Quarters" in exp.columns:
    print(exp["Quarters"].value_counts(dropna=False))
    print(exp.groupby("Quarters")["amount"].sum())
