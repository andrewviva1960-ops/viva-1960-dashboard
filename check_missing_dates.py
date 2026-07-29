import pandas as pd, warnings, sys
sys.stdout.reconfigure(encoding="utf-8")
warnings.filterwarnings("ignore")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (3)_FIXED.xlsx"

sales = pd.read_excel(path, sheet_name="Sales Raw Data 2026", header=1)
sales["Month_num"] = pd.to_numeric(sales["Month"], errors="coerce")
sales["date_str"] = sales["Date"].astype(str)

# Get dates that have Month values
with_month = sales.dropna(subset=["Month_num"])
print("Sales rows WITH Month: {}".format(len(with_month)))
print("Unique dates with Month: {}".format(with_month["date_str"].nunique()))
print()

# Show all unique dates with their Month
print("All dates in Sales that have Month values:")
for d, m in sorted(with_month.groupby("date_str")["Month_num"].first().items(), key=lambda x: (int(x[0]) if x[0].lstrip("-").isdigit() else 999999, x[0])):
    print("  {} -> Month {}".format(d, int(m)))

print()
# Get expense dates that are NOT in the lookup
exp = pd.read_excel(path, sheet_name="Expenses Raw Data 2026")
exp["date_str"] = exp["date"].astype(str)
lookup_dates = set(with_month["date_str"].unique())
exp_dates = set(exp["date_str"].unique())
missing = sorted(exp_dates - lookup_dates, key=lambda x: int(x))
print("Expense dates NOT in Sales lookup ({} total):".format(len(missing)))
print(missing[:50])
if len(missing) > 50:
    print("  ... and {} more".format(len(missing)-50))
