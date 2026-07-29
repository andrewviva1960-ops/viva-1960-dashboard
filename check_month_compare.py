import pandas as pd, warnings
warnings.filterwarnings("ignore")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (3)_FIXED.xlsx"

sales = pd.read_excel(path, sheet_name="Sales Raw Data 2026", header=1)
sales["Sales Amount"] = pd.to_numeric(sales["Sales Amount"], errors="coerce").fillna(0)

# Compare months from Excel's Month column vs my parsed month
excel_month = pd.to_numeric(sales["Month"], errors="coerce")

# My parsing
sales["Date_num"] = pd.to_numeric(sales["Date"], errors="coerce")
sales["Date_dt"] = pd.to_datetime(sales["Date"], format="%d/%m/%Y", errors="coerce")
m1 = sales["Date_dt"].isna() & sales["Date_num"].notna()
sales.loc[m1, "Date_dt"] = pd.to_datetime(sales.loc[m1, "Date_num"], origin="1899-12-30", unit="D")
m2 = sales["Date_dt"].isna()
sales.loc[m2, "Date_dt"] = pd.to_datetime(sales.loc[m2, "Date"], format="%Y/%d/%m", errors="coerce")
my_month = sales["Date_dt"].dt.month

# Compare rows where both have values
both_valid = excel_month.notna() & my_month.notna()
diff = both_valid & (excel_month != my_month)
print("Rows where Excel month and my month differ: {}".format(diff.sum()))
if diff.any():
    print("Sample differences:")
    for _, r in sales[diff].head(20).iterrows():
        em = excel_month.loc[r.name]
        mm = my_month.loc[r.name]
        print("  Row {}: Date={}, Excel Month={:.0f}, My Month={:.0f}, Sales={:,.0f}".format(
            r.name, r["Date"], em, mm, r["Sales Amount"]))

# If there are differences, show impact on monthly totals
if diff.any():
    print()
    print("=== Monthly totals (Excel Month) ===")
    for m in range(1, 13):
        excel_amt = sales.loc[excel_month == m, "Sales Amount"].sum()
        my_amt = sales.loc[my_month == m, "Sales Amount"].sum()
        diff_amt = excel_amt - my_amt
        if abs(diff_amt) > 0:
            print("  Month {}: Excel={:,.0f}, Mine={:,.0f}, Diff={:,.0f}".format(m, excel_amt, my_amt, diff_amt))
    
    print()
    print("Excel month total: {:,.0f}".format(sales.loc[excel_month.notna(), "Sales Amount"].sum()))
    print("My month total:    {:,.0f}".format(sales.loc[my_month.notna(), "Sales Amount"].sum()))
else:
    print("No differences - both months match")
    # Show how many rows have Excel month vs my parsed month
    print("Excel month non-null: {}".format(excel_month.notna().sum()))
    print("My month non-null: {}".format(my_month.notna().sum()))
