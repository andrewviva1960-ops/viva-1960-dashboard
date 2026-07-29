import pandas as pd, warnings
warnings.filterwarnings("ignore")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (3)_FIXED.xlsx"

sales = pd.read_excel(path, sheet_name="Sales Raw Data 2026", header=1)

# Check the dates that are being parsed as Jul-Dec
print("=== Rows where parsed month is 7-12 ===")
sales["Date_num"] = pd.to_numeric(sales["Date"], errors="coerce")
sales["Date_dt"] = pd.to_datetime(sales["Date"], format="%d/%m/%Y", errors="coerce")
m1 = sales["Date_dt"].isna() & sales["Date_num"].notna()
sales.loc[m1, "Date_dt"] = pd.to_datetime(sales.loc[m1, "Date_num"], origin="1899-12-30", unit="D")
m2 = sales["Date_dt"].isna()
sales.loc[m2, "Date_dt"] = pd.to_datetime(sales.loc[m2, "Date"], format="%Y/%d/%m", errors="coerce")
sales["Month"] = sales["Date_dt"].dt.month
sales["Sales Amount"] = pd.to_numeric(sales["Sales Amount"], errors="coerce").fillna(0)

# Look at the raw Date values for Jul-Dec rows
jul_dec = sales[sales["Month"].between(7, 12)]
print("Total Jul-Dec rows: {}".format(len(jul_dec)))
print("Total Jul-Dec sales: ${:,.0f}".format(jul_dec["Sales Amount"].sum()))
print()

# Show unique Date values and their conversion
print("=== Unique Date values in Jul-Dec and their conversion ===")
for d in jul_dec["Date"].unique():
    row = jul_dec[jul_dec["Date"] == d].iloc[0]
    print("  Date raw: {} -> parsed: {} -> Month: {}".format(d, row["Date_dt"], int(row["Month"])))
    # Check if this is a serial number
    num = pd.to_numeric(d, errors="coerce")
    if pd.notna(num):
        # Alternative: maybe should be treated as invoice number, not date?
        print("    -> This is numeric serial: {}".format(int(num)))

print()
print("=== Check if the source Date column is really dates or invoice #s ===")
# Sample some rows around the area
for idx in [0, 1, 2, 585, 586, 587]:
    row = sales.iloc[idx]
    print("Row {}: Date='{}' (type={}), Sales={:,.0f}, Client='{}'".format(
        idx, row["Date"], type(row["Date"]).__name__, row["Sales Amount"], row["Client"]))
