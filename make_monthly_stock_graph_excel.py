# -*- coding: utf-8 -*-
"""
make_monthly_stock_graph_excel.py

make_monthly_max_stock_csv.py で作成した
「各駐車場 × 1か月分の月次まとめCSV（***allYYYYMM.csv）」から、

- 横軸：日付
- 縦軸：一般在庫 + 定期在庫 の積み上げ棒グラフ
  （平日・土日祝で色分け）

を作成し、各CSVごとにグラフ付きExcelファイル(.xlsx)を出力するスクリプト。

【前提】
- 月次まとめCSVファイル名の例：
    全駐車場all202510.csv
    南1all202510.csv
    南2all202510.csv
    ...
    北3all202510.csv
  → つまり「allYYYYMM.csv」を含むファイルだけを対象とする。

- 各CSVには少なくとも以下の列がある想定：
    ・日付列          … 例: "日付"
    ・一般在庫列      … 例: "一般在庫"
    ・定期在庫列      … 例: "定期在庫"
    ・土日祝区分列    … 例: "土日祝区分" （値は "平日" or "土日祝"）
"""

import os
import glob
import argparse
import pandas as pd


def create_graph_excel_from_csv(
    csv_path: str,
    date_col: str,
    general_col: str,
    teiki_col: str,
    holiday_col: str,
    output_dir: str,
) -> None:
    """1つの月次CSVから、データ＋グラフ付きのExcelファイルを作成する。"""

    print(f"[INFO] 読み込み中: {csv_path}")

    # ===== CSV読み込み =====
    df = pd.read_csv(csv_path, encoding="cp932")

    # 日付列を datetime 型にして、日付順に並べ替え
    df[date_col] = pd.to_datetime(df[date_col])
    df = df.sort_values(date_col).reset_index(drop=True)

    # ★ 追加：Excelでは文字列として扱わせるため "YYYY-MM-DD" へ変換
    df[date_col] = df[date_col].dt.strftime("%Y-%m-%d")


    # ===== 平日/土日祝ごとの列を作成 =====
    # 土日祝かどうかのフラグ（True = 土日祝）
    is_holiday = df[holiday_col] == "土日祝"

    # 平日分（休日の行は NaN）
    df["一般_平日"] = df[general_col].where(~is_holiday)
    df["定期_平日"] = df[teiki_col].where(~is_holiday)

    # 土日祝分（平日の行は NaN）
    df["一般_土日祝"] = df[general_col].where(is_holiday)
    df["定期_土日祝"] = df[teiki_col].where(is_holiday)

    # ===== 出力パス準備 =====
    base_name = os.path.splitext(os.path.basename(csv_path))[0]
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, base_name + "_graph.xlsx")

    print(f"[INFO] Excel出力: {output_path}")

    # ===== ExcelWriter（xlsxwriter）でデータとグラフを書き込む =====
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        sheet_name_data = "data"
        df.to_excel(writer, sheet_name=sheet_name_data, index=False)

        workbook = writer.book
        worksheet_data = writer.sheets[sheet_name_data]

        # データ行数と列インデックス取得
        n_rows = len(df)
        date_idx = df.columns.get_loc(date_col)

        wd_gen_idx = df.columns.get_loc("一般_平日")
        wd_teiki_idx = df.columns.get_loc("定期_平日")
        hol_gen_idx = df.columns.get_loc("一般_土日祝")
        hol_teiki_idx = df.columns.get_loc("定期_土日祝")

        # ===== 積み上げ棒グラフ作成 =====
        chart = workbook.add_chart({"type": "column", "subtype": "stacked"})

        # series 1: 一般_平日
        chart.add_series({
            "name":       "一般_平日",
            "categories": [sheet_name_data, 1, date_idx, n_rows, date_idx],
            "values":     [sheet_name_data, 1, wd_gen_idx, n_rows, wd_gen_idx],
        })

        # series 2: 定期_平日
        chart.add_series({
            "name":       "定期_平日",
            "categories": [sheet_name_data, 1, date_idx, n_rows, date_idx],
            "values":     [sheet_name_data, 1, wd_teiki_idx, n_rows, wd_teiki_idx],
        })

        # series 3: 一般_土日祝
        chart.add_series({
            "name":       "一般_土日祝",
            "categories": [sheet_name_data, 1, date_idx, n_rows, date_idx],
            "values":     [sheet_name_data, 1, hol_gen_idx, n_rows, hol_gen_idx],
        })

        # series 4: 定期_土日祝
        chart.add_series({
            "name":       "定期_土日祝",
            "categories": [sheet_name_data, 1, date_idx, n_rows, date_idx],
            "values":     [sheet_name_data, 1, hol_teiki_idx, n_rows, hol_teiki_idx],
        })

        # グラフタイトル・軸ラベル
        chart.set_title({"name": f"{base_name} 月次在庫（平日・土日祝別 一般＋定期）"})
        chart.set_x_axis({"name": "日付"})
        chart.set_y_axis({"name": "在庫台数"})

        # 凡例（一般_平日 / 一般_土日祝 など）
        chart.set_legend({"position": "bottom"})

        # ★ グラフサイズを大きくする（ここを追加）
        #   単位は「ピクセル」です。お好みで数値を調整してください。
        chart.set_size({
            "width":  1200,   # 横幅（例：1200ピクセル）
            "height": 600,    # 高さ（例：600ピクセル）
        })

        # ===== graph シートを作ってグラフ貼り付け =====
        sheet_name_graph = "graph"
        worksheet_graph = workbook.add_worksheet(sheet_name_graph)
        worksheet_graph.insert_chart("A1", chart)


def main():
    parser = argparse.ArgumentParser(
        description="月次まとめCSV（***allYYYYMM.csv）から、平日/土日祝別の積み上げグラフ付きExcelを作成するスクリプト"
    )
    parser.add_argument(
        "--year",
        type=int,
        required=True,
        help="対象年（例: 2025）",
    )
    parser.add_argument(
        "--month",
        type=int,
        required=True,
        help="対象月（1〜12）",
    )
    parser.add_argument(
        "--base_dir",
        type=str,
        default="downloads\\csv",
        help="csvYYYYMM フォルダが入っているベースディレクトリ（既定: downloads\\csv）",
    )
    parser.add_argument(
        "--date_col",
        type=str,
        default="日付",
        help="日付列の列名（既定: '日付'）",
    )
    parser.add_argument(
        "--general_col",
        type=str,
        default="一般在庫",
        help="一般在庫列の列名（既定: '一般在庫'）",
    )
    parser.add_argument(
        "--teiki_col",
        type=str,
        default="定期在庫",
        help="定期在庫列の列名（既定: '定期在庫'）",
    )
    parser.add_argument(
        "--holiday_col",
        type=str,
        default="土日祝区分",
        help="土日祝区分の列名（既定: '土日祝区分'）",
    )
    parser.add_argument(
        "--output_subdir",
        type=str,
        default="graphs",
        help="月次フォルダ配下に作成する出力サブフォルダ名（既定: 'graphs'）",
    )

    args = parser.parse_args()

    ym = f"{args.year}{args.month:02d}"
    # 月次CSVが入っているフォルダ（例: downloads\csv\csv202510）
    monthly_dir = os.path.join(args.base_dir, f"csv{ym}")
    if not os.path.isdir(monthly_dir):
        print(f"[ERROR] 月次フォルダが存在しません: {monthly_dir}")
        return

    # ***allYYYYMM.csv だけを対象にする
    pattern = os.path.join(monthly_dir, f"*all{ym}.csv")
    csv_files = glob.glob(pattern)

    if not csv_files:
        print(f"[WARN] 月次まとめCSVが見つかりませんでした: {pattern}")
        return

    print(f"[INFO] 対象月次CSVファイル数: {len(csv_files)}")
    output_dir = os.path.join(monthly_dir, args.output_subdir)

    for csv_path in csv_files:
        create_graph_excel_from_csv(
            csv_path=csv_path,
            date_col=args.date_col,
            general_col=args.general_col,
            teiki_col=args.teiki_col,
            holiday_col=args.holiday_col,
            output_dir=output_dir,
        )

    print("[INFO] すべてのグラフ付きExcel作成が完了しました。")


if __name__ == "__main__":
    main()
