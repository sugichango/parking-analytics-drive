import pandas as pd
import os

file_path = r"C:\Users\sugitamasahiko\Documents\駐車場データ分析試行\年度別データ\updated_integrated_data_FY2025.csv"
output_path = r"C:\Users\sugitamasahiko\Documents\駐車場データ分析試行\Antigravity試行\discount_43_44_records.csv"

print(f"Reading {file_path}...")
try:
    df = pd.read_csv(file_path, encoding='utf-8')
except UnicodeDecodeError:
    df = pd.read_csv(file_path, encoding='cp932')

print(f"Total rows: {len(df)}")
print("Filtering for Discount 43 and 44...")

discount_cols = [f'Discount{i}' for i in range(1, 8)]

mask = pd.Series(False, index=df.index)
for col in discount_cols:
    if col in df.columns:
        # Convert column to numeric just in case, replacing non-numeric with NaN
        col_numeric = pd.to_numeric(df[col], errors='coerce')
        mask = mask | col_numeric.isin([43, 44])

extracted_df = df[mask]
print(f"Extracted {len(extracted_df)} records.")

if len(extracted_df) > 0:
    extracted_df.to_csv(output_path, index=False, encoding='utf-8-sig')
    print(f"Saved to {output_path}")
else:
    print("No matching records found.")
