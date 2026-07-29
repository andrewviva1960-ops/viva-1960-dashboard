import pandas as pd, warnings
warnings.filterwarnings("ignore")
path = r"C:\Users\Andro\Downloads\Financial Model\Viva Financial model 2026 (3)_FIXED.xlsx"

sheets = pd.ExcelFile(path).sheet_names
print("Sheets:", sheets)

for s in sheets:
    df = pd.read_excel(path, sheet_name=s, header=None)
    print(f"\n=== {s} === ({df.shape[0]} rows x {df.shape[1]} cols)")
    # Print first 20 rows
    for i in range(min(25, len(df))):
        row = df.iloc[i].tolist()
        # Show non-NaN values
        vals = [str(v) for v in row if str(v) != "nan" and str(v) != ""]
        if vals:
            print(f"  Row {i}: {vals[:8]}")
