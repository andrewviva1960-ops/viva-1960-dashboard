import pandas as pd, warnings
warnings.filterwarnings("ignore")

# Check the original file
path_orig = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (3).xlsx"
path_fixed = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (3)_FIXED.xlsx"

try:
    sales_orig = pd.read_excel(path_orig, sheet_name="Sales Raw Data 2026", header=1)
    print("=== Original file ===")
    print("Date column type:", type(sales_orig["Date"].iloc[0]))
    print("First 10 Dates:", list(sales_orig["Date"].head(10)))
    print()
    
    # Parse dates
    sales_orig["Date_num"] = pd.to_numeric(sales_orig["Date"], errors="coerce")
    sales_orig["Date_dt"] = pd.to_datetime(sales_orig["Date"], format="%d/%m/%Y", errors="coerce")
    m1 = sales_orig["Date_dt"].isna() & sales_orig["Date_num"].notna()
    sales_orig.loc[m1, "Date_dt"] = pd.to_datetime(sales_orig.loc[m1, "Date_num"], origin="1899-12-30", unit="D")
    m2 = sales_orig["Date_dt"].isna()
    sales_orig.loc[m2, "Date_dt"] = pd.to_datetime(sales_orig.loc[m2, "Date"], format="%Y/%d/%m", errors="coerce")
    sales_orig["Month"] = sales_orig["Date_dt"].dt.month
    sales_orig["Sales Amount"] = pd.to_numeric(sales_orig["Sales Amount"], errors="coerce").fillna(0)
    
    print("Monthly breakdown (Original file):")
    for m in range(1, 13):
        amt = sales_orig.loc[sales_orig["Month"] == m, "Sales Amount"].sum()
        print("  Month {}: ${:,.0f}".format(m, amt))
    print("  Total: ${:,.0f}".format(sales_orig["Sales Amount"].sum()))
    
except Exception as e:
    print("Error reading original file:", e)
    
print()
print("=== Fixed file for comparison ===")
sales_fixed = pd.read_excel(path_fixed, sheet_name="Sales Raw Data 2026", header=1)
print("Date column type:", type(sales_fixed["Date"].iloc[0]))
print("First 10 Dates:", list(sales_fixed["Date"].head(10)))
