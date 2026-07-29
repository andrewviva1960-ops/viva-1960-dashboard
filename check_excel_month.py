import pandas as pd, warnings
warnings.filterwarnings("ignore")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (3)_FIXED.xlsx"

sales = pd.read_excel(path, sheet_name="Sales Raw Data 2026", header=1)

# Parse dates
sales["Date_num"] = pd.to_numeric(sales["Date"], errors="coerce")
sales["Date_dt"] = pd.to_datetime(sales["Date"], format="%d/%m/%Y", errors="coerce")
m1 = sales["Date_dt"].isna() & sales["Date_num"].notna()
sales.loc[m1, "Date_dt"] = pd.to_datetime(sales.loc[m1, "Date_num"], origin="1899-12-30", unit="D")
m2 = sales["Date_dt"].isna()
sales.loc[m2, "Date_dt"] = pd.to_datetime(sales.loc[m2, "Date"], format="%Y/%d/%m", errors="coerce")
sales["Parsed_Month"] = sales["Date_dt"].dt.month
sales["Sales Amount"] = pd.to_numeric(sales["Sales Amount"], errors="coerce").fillna(0)

# Check if Month column exists
if "Month" in sales.columns:
    sales["Excel_Month"] = pd.to_numeric(sales["Month"], errors="coerce")
    
    print("=== Rows where Parsed_Month >= 7 ===")
    jul_dec = sales[sales["Parsed_Month"] >= 7]
    print("Total rows: {}".format(len(jul_dec)))
    print("Total sales: ${:,.0f}".format(jul_dec["Sales Amount"].sum()))
    print()
    
    print("What does the Excel Month column say for these?")
    conflict = jul_dec[jul_dec["Excel_Month"].notna() & (jul_dec["Excel_Month"] != jul_dec["Parsed_Month"])]
    no_excel = jul_dec[jul_dec["Excel_Month"].isna()]
    match = jul_dec[jul_dec["Excel_Month"].notna() & (jul_dec["Excel_Month"] == jul_dec["Parsed_Month"])]
    print("  Conflict (Excel Month diff from parsed): {} rows, ${:,.0f}".format(len(conflict), conflict["Sales Amount"].sum()))
    print("  Excel Month is NaN: {} rows, ${:,.0f}".format(len(no_excel), no_excel["Sales Amount"].sum()))
    print("  Excel Month matches: {} rows, ${:,.0f}".format(len(match), match["Sales Amount"].sum()))
    
    if len(conflict) > 0:
        print()
        print("=== Sample conflicts (Excel Month vs Parsed Month) ===")
        for idx in conflict.head(10).index:
            row = conflict.loc[idx]
            print("  Idx {}: Date={}, Excel Month={:.0f}, Parsed Month={:.0f}, Sales={:,.0f}, Client='{}'".format(
                idx, row["Date"], row["Excel_Month"], row["Parsed_Month"], row["Sales Amount"], row["Client"]))
else:
    print("No 'Month' column found. Columns are:", list(sales.columns))
