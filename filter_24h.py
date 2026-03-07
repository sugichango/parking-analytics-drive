import pandas as pd

input_path = r"C:\Users\sugitamasahiko\Documents\駐車場データ分析試行\Antigravity試行\discount_43_44_records.csv"
output_path = r"C:\Users\sugitamasahiko\Documents\駐車場データ分析試行\Antigravity試行\discount_43_44_over_24h.csv"

print(f"Reading {input_path}...")
df = pd.read_csv(input_path, encoding='utf-8-sig')

# InTime と OnTime の列が日付として認識できるように変換
# 駐車場データでは、InTime が入庫時間、OnTime が精算/出庫時間と推測
df['InTime_dt'] = pd.to_datetime(df['InTime'], errors='coerce')
df['OnTime_dt'] = pd.to_datetime(df['OnTime'], errors='coerce')

# 駐車時間（期間）の計算
df['ParkingDuration'] = df['OnTime_dt'] - df['InTime_dt']

# 24時間を超えるデータを抽出
over_24h_mask = df['ParkingDuration'] > pd.Timedelta(hours=24)
filtered_df = df[over_24h_mask]

# 計算用の列を削除して元の形に戻す（残したい場合は削除不要）
# filtered_df = filtered_df.drop(columns=['InTime_dt', 'OnTime_dt', 'ParkingDuration'])

print(f"Total rows: {len(df)}")
print(f"Rows with parking time > 24h: {len(filtered_df)}")

if len(filtered_df) > 0:
    filtered_df.to_csv(output_path, index=False, encoding='utf-8-sig')
    print(f"Saved to {output_path}")
else:
    print("No records found with parking time over 24 hours.")
