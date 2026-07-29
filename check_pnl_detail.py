import pandas as pd, warnings
warnings.filterwarnings("ignore")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (3)_FIXED.xlsx"
pnl = pd.read_excel(path, sheet_name="PNL Dashboard ", header=None)

# Show rows 53-62, columns 9-14 in detail
print("=== Profitability Ratio section (detailed) ===")
for r in range(53, 62):
    label = str(pnl.iloc[r, 0])[:30] if pd.notna(pnl.iloc[r, 0]) else "(empty)"
    vals = []
    for c in range(9, 15):
        v = pnl.iloc[r, c]
        if pd.notna(v):
            vals.append("Col{}={}".format(c, v))
        else:
            vals.append("Col{}=NaN".format(c))
    print("Row {} [{}]: {}".format(r, label, ", ".join(vals)))

print()

# Also check the YTD section for monthly Gross Sales columns
# Look for ANY values in the monthly columns (cols 18-80) in the first 8 rows
print("=== Monthly section values (rows 0-15, cols 18-80) ===")
for r in range(0, 15):
    vals = {}
    for c in range(18, 80):
        v = pnl.iloc[r, c]
        if pd.notna(v) and isinstance(v, (int, float)):
            vals[c] = v
    if vals:
        print("Row {}: {}".format(r, vals))

# Let me also check if there are values further down the sheet in the monthly columns
print()
print("=== All numeric values in monthly columns (rows 10-150) ===")
for r in range(10, pnl.shape[0]):
    for c in [19, 28, 37, 46, 55, 64, 73]:  # G.Sales A columns
        if c < pnl.shape[1]:
            v = pnl.iloc[r, c]
            if pd.notna(v) and isinstance(v, (int, float)) and v > 0:
                label = str(pnl.iloc[r, 0])[:30] if pd.notna(pnl.iloc[r, 0]) else ""
                print("Row {}, Col {}: {:,.0f} [{}]".format(r, c, v, label))
