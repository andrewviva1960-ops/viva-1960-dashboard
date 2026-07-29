import pandas as pd, warnings, sys
sys.stdout.reconfigure(encoding="utf-8")
warnings.filterwarnings("ignore")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (3)_FIXED.xlsx"

# Read expense sheet without header inference
exp = pd.read_excel(path, sheet_name="Expenses Raw Data 2026", header=None)
print(f"Shape: {exp.shape}")
print("\nFirst 5 rows (all columns):")
for i in range(min(5, len(exp))):
    row = [str(v) for v in exp.iloc[i].tolist()]
    print(f"  Row {i}: {row}")

print("\nLast 3 rows:")
for i in range(max(0, len(exp)-3), len(exp)):
    row = [str(v) for v in exp.iloc[i].tolist()]
    print(f"  Row {i}: {row}")

# Check column 5 (Month) for non-NaN values
col5 = exp.iloc[:, 5]
non_nan = col5.dropna()
print(f"\nColumn 5 (Month) non-NaN count: {len(non_nan)}")
if len(non_nan) > 0:
    print(f"  Values: {non_nan.unique()[:20]}")

# Check if the header is row 0
print(f"\nRow 0 values: {exp.iloc[0].tolist()}")
print(f"Row 1 values: {exp.iloc[1].tolist()}")
