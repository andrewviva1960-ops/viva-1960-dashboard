import pandas as pd, warnings
warnings.filterwarnings("ignore")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (3)_FIXED.xlsx"
exp = pd.read_excel(path, sheet_name="Expenses Raw Data 2026")
exp["amount"] = pd.to_numeric(exp["amount"], errors="coerce").fillna(0)
exp["Month_excel"] = pd.to_numeric(exp["Month"], errors="coerce")
exp["Date_dt"] = pd.to_datetime(exp["date"], origin="1899-12-30", unit="D", errors="coerce")
exp["Month_parsed"] = exp["Date_dt"].dt.month

# Check conflicts
conflict = exp[exp["Month_excel"].notna() & exp["Month_parsed"].notna() & (exp["Month_excel"] != exp["Month_parsed"])]
print("Expense conflicts (Excel Month != Parsed Month): {} rows, ${:,.2f}".format(
    len(conflict), conflict["amount"].sum()))
if len(conflict) > 0:
    print("Sample:")
    for idx in conflict.head(5).index:
        print("  date={}, Excel Month={:.0f}, Parsed Month={:.0f}, amount={:,.2f}".format(
            exp.loc[idx, "date"], exp.loc[idx, "Month_excel"], exp.loc[idx, "Month_parsed"], exp.loc[idx, "amount"]))

gc = exp[exp["Month_excel"].notna() & exp["Month_parsed"].notna() & (exp["Month_excel"] == exp["Month_parsed"])]
print("Excel Month matches Parsed: {} rows, ${:,.2f}".format(len(gc), gc["amount"].sum()))

only_excel = exp[exp["Month_excel"].notna() & exp["Month_parsed"].isna()]
print("Only Excel Month: {} rows, ${:,.2f}".format(len(only_excel), only_excel["amount"].sum()))

only_parsed = exp[exp["Month_excel"].isna() & exp["Month_parsed"].notna()]
print("Only Parsed Month: {} rows, ${:,.2f}".format(len(only_parsed), only_parsed["amount"].sum()))
