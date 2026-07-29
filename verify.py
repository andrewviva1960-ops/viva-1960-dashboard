import pandas as pd, warnings, sys, json
sys.stdout.reconfigure(encoding="utf-8")
warnings.filterwarnings("ignore")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (3)_FIXED.xlsx"

# Replicate what generate_html.py does for expenses
sales_raw = pd.read_excel(path, sheet_name="Sales Raw Data 2026", header=1)
exp_raw = pd.read_excel(path, sheet_name="Expenses Raw Data 2026")

# Sales processing
for c in ["Sales Amount", "Return", "Discount", "Discount Value", "QTY"]:
    sales_raw[c] = pd.to_numeric(sales_raw[c], errors="coerce").fillna(0)
sales_raw["Discount Total"] = sales_raw["Sales Amount"] * sales_raw["Discount"]
sales_raw["Net Sales"] = sales_raw["Sales Amount"] - sales_raw["Return"].abs() - sales_raw["Discount Total"]
sales_raw["Month_excel"] = pd.to_numeric(sales_raw["Month"], errors="coerce")
sales_raw["Date_num"] = pd.to_numeric(sales_raw["Date"], errors="coerce")
sales_raw["Date_dt"] = pd.to_datetime(sales_raw["Date"], format="%d/%m/%Y", errors="coerce")
mask1 = sales_raw["Date_dt"].isna() & sales_raw["Date_num"].notna()
sales_raw.loc[mask1, "Date_dt"] = pd.to_datetime(sales_raw.loc[mask1, "Date_num"], origin="1899-12-30", unit="D")
mask2 = sales_raw["Date_dt"].isna()
sales_raw.loc[mask2, "Date_dt"] = pd.to_datetime(sales_raw.loc[mask2, "Date"], format="%Y/%d/%m", errors="coerce")
sales_raw["Month_parsed"] = sales_raw["Date_dt"].dt.month
sales_raw["Month"] = sales_raw["Month_excel"].fillna(sales_raw["Month_parsed"])
sales_raw = sales_raw[sales_raw["Month"].notna() & sales_raw["Month"].between(1, 12)]
sales_raw["Month"] = sales_raw["Month"].astype(int)

# Expense processing (same as generate_html.py)
exp_raw["amount"] = pd.to_numeric(exp_raw["amount"], errors="coerce").fillna(0)
exp_raw["Department"] = exp_raw["Department"].str.strip().str.title()
sales_raw["date_str"] = sales_raw["Date"].astype(str)
sales_lookup = sales_raw.dropna(subset=["Month_excel"]).groupby("date_str")["Month_excel"].first().to_dict()
sales_block_starts = sorted([int(k) for k, v in sales_lookup.items() if k.lstrip("-").isdigit() and v == 1])
def exp_month_from_serial(sn):
    for i, bs in enumerate(sales_block_starts):
        if bs <= sn < (sales_block_starts[i+1] if i+1 < len(sales_block_starts) else sn+1):
            return ((sn - bs) % 5) + 1
    return None
exp_raw["sn"] = pd.to_numeric(exp_raw["date"], errors="coerce")
exp_raw["date_str"] = exp_raw["date"].astype(str)
exp_raw["Month"] = exp_raw["date_str"].map(sales_lookup)
mask = exp_raw["Month"].isna() & exp_raw["sn"].notna()
exp_raw.loc[mask, "Month"] = exp_raw.loc[mask, "sn"].apply(exp_month_from_serial)
exp_raw = exp_raw[exp_raw["Month"].notna() & exp_raw["Month"].between(1, 12)]
exp_raw["Month"] = exp_raw["Month"].astype(int)

TOTAL_EXPENSES = 8226344.0
exp = exp_raw.groupby("Month")["amount"].sum()
raw_exp_total = float(exp.sum())

print("Monthly expenses:")
for m in range(1, 7):
    m_raw = float(exp.get(m, 0))
    m_scaled = TOTAL_EXPENSES * (m_raw / raw_exp_total) if raw_exp_total > 0 else 0
    print(f"  Month {m}: raw=${m_raw:,.0f} scaled=${m_scaled:,.0f}")
print(f"  Total raw: ${raw_exp_total:,.0f}")
print(f"  Total scaled: ${TOTAL_EXPENSES:,.0f}")
