import pandas as pd

try:
    # CSVファイルを読み込む
    df = pd.read_csv('202410.csv')

    # 重複をチェックする対象のカラムを指定します。
    target_columns = ['ParkingTicketNo', 'InTime', 'OnTime']

    # 指定されたカラムの組み合わせで重複している行をすべて抽出します。
    # duplicated(subset=..., keep=False) は、重複しているすべての行をTrueとします。
    duplicated_rows = df[df.duplicated(subset=target_columns, keep=False)]

    print(f"--- 指定されたカラム ({', '.join(target_columns)}) で重複している行の数: {len(duplicated_rows)} ---")

    if not duplicated_rows.empty:
        print("--- 重複している行の先頭5件 ---")
        print(duplicated_rows.head())
        print("\n--- 重複している行の概要 (df.info()) ---")
        duplicated_rows.info()
    else:
        print("指定されたカラムの組み合わせで重複している行は見つかりませんでした。")

except FileNotFoundError:
    print("エラー: '202410.csv' が見つかりません。ファイルパスを確認してください。")
except KeyError as e:
    print(f"エラー: 指定されたカラムが見つかりません。カラム名を確認してください: {e}")
except Exception as e:
    print(f"データの読み込みまたは処理中にエラーが発生しました: {e}")
