import pandas as pd

# --- ステップ0: CSVファイルからデータを読み込む ---
# ファイルが存在しない場合のエラー処理を追加
try:
    df = pd.read_csv('202410.csv')
except FileNotFoundError:
    print("エラー: '202410.csv' が見つかりません。スクリプトと同じディレクトリに配置してください。")
    exit()


# --- ステップ1: データを「入庫」と「出庫」に分離 ---

# 「入庫」レコードを抽出 (ProcessingIndicatorが1のもの)。
# これが最終的な出力のベースとなり、IDを含む全カラムが保持されます。
entries = df[df['ProcessingIndicator'] == 1].copy()

# 「出庫」レコードを抽出 (ProcessingIndicatorが1ではないもの)。
# こちらからは、紐付けとOnTimeの更新に使用する情報のみを使います。
exits = df[df['ProcessingIndicator'] != 1].copy()


# --- ステップ2: 正しいキーで入庫と出庫をマージ ---

# 「入庫」と「出庫」のペアを、以下の2つの条件が一致するものですべて紐付けます。
#   1. ParkingTicketNoが同じ
#   2. 入庫の 'InTime' と 出庫の 'OnTime' が同じ
# これにより、各入庫レコードに対応する出庫レコードが正確に紐付けられ、
# 抽出漏れがなくなります。
paired_df = pd.merge(
    left=entries,
    right=exits,
    how='inner',  # 両方に条件が一致するペアのみを抽出
    left_on=['ParkingTicketNo', 'InTime'],
    right_on=['ParkingTicketNo', 'OnTime'],
    suffixes=('_entry', '_exit') # 重複する列名を区別するための接尾辞
)


# --- ステップ3: 最終的なデータフレームを構築 ---

# マージによってできたデータフレームから、元の「入庫」由来の列だけをまず選択します。
# これで、IDを含むすべての元カラムが保持されます。
entry_cols = [col for col in paired_df.columns if col.endswith('_entry')]
final_df = paired_df[entry_cols].copy()

# 列名を元の名前に戻します（例: 'ParkingTicketNo_entry' -> 'ParkingTicketNo'）。
final_df.columns = [col.replace('_entry', '') for col in final_df.columns]

# 最後に、'OnTime'列を、紐付いた「出庫」レコードのOnTime ('OnTime_exit') で更新します。
# paired_dfとfinal_dfの行の並びは一致しているため、直接代入が可能です。
final_df['OnTime'] = paired_df['OnTime_exit']


# --- ステップ4: 結果をCSVファイルに書き出し ---
# final_output.csv という名前でCSVファイルに書き出します。
# index=False は、DataFrameのインデックスが余分な列として保存されるのを防ぎます。
# encoding='utf-8-sig' は、Excelでファイルを開いた際の文字化けを防ぎます。
final_df.to_csv('final_output.csv', index=False, encoding='utf-8-sig')


# --- 完了メッセージ ---
print("処理が完了しました。")
print(f"抽出されたレコード数: {len(final_df)}")
print("結果は 'final_output.csv' に保存されました。")
print("\n最終データフレームのプレビュー:")
print(final_df.head())