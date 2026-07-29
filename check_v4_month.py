import pandas as pd, warnings
warnings.filterwarnings("ignore")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (4).xlsx"
exp = pd.read_excel(path, sheet_name="Expenses Raw Data 2026", header=0)
exp["amount"] = pd.to_numeric(exp["amount"], errors="coerce").fillna(0)
exp["Month_col"] = pd.to_numeric(exp["Month"], errors="coerce")
exp["serial_month"] = pd.to_datetime(pd.to_numeric(exp["date"], errors="coerce"), origin="1899-12-30", unit="D", errors="coerce").dt.month

both = exp.dropna(subset=["Month_col", "serial_month"])
diff = (both["Month_col"] != both["serial_month"]).sum()
same = (both["Month_col"] == both["serial_month"]).sum()
print(f"Same: {same}, Different: {diff}")
if diff > 0:
    print("Differences:")
    for idx in both[diff].index[:10]:
        print(f"  Row {idx}: date={exp.loc[idx,'date']}, Month_col={exp.loc[idx,'Month_col']}, serial={exp.loc[idx,'serial_month']}")
else:
    print("Month column exactly matches serial date parsing - no differences.")
