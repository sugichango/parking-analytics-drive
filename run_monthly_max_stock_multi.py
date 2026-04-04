# -*- coding: utf-8 -*-
"""
run_monthly_max_stock_multi.py

make_monthly_max_stock_csv.py を「複数月まとめて」実行するためのラッパースクリプト。
元ファイル（make_monthly_max_stock_csv.py）は変更しません。
"""

import argparse
from datetime import date

# 同じフォルダにある make_monthly_max_stock_csv.py の関数を読み込む
from make_monthly_max_stock_csv import create_monthly_max_stock_csv


def month_iter(start_year: int, start_month: int, end_year: int, end_month: int):
    """
    開始年月〜終了年月までを 1ヶ月ずつ進めながら (year, month) を返すジェネレータ。
    例: (2025,4)〜(2025,6) → (2025,4), (2025,5), (2025,6)
    """
    current = date(start_year, start_month, 1)
    last = date(end_year, end_month, 1)

    if current > last:
        raise ValueError("開始年月が終了年月より後になっています。指定を見直してください。")

    while current <= last:
        yield current.year, current.month

        # 月を1つ進める
        if current.month == 12:
            current = date(current.year + 1, 1, 1)
        else:
            current = date(current.year, current.month + 1, 1)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="make_monthly_max_stock_csv.py を複数月まとめて実行するラッパー"
    )

    parser.add_argument(
        "--start-year",
        type=int,
        required=True,
        help="開始年 (例: 2025)",
    )
    parser.add_argument(
        "--start-month",
        type=int,
        required=True,
        help="開始月 (1〜12)",
    )
    parser.add_argument(
        "--end-year",
        type=int,
        required=True,
        help="終了年 (例: 2025)",
    )
    parser.add_argument(
        "--end-month",
        type=int,
        required=True,
        help="終了月 (1〜12)",
    )
    parser.add_argument(
        "--base-dir",
        type=str,
        default="downloads",
        help="CSVフォルダのベースディレクトリ "
             "(make_monthly_max_stock_csv.py の --base_dir に渡す値と同じ。例: downloads\\csv)",
    )

    return parser.parse_args()


def main():
    args = parse_args()

    print("[INFO] 複数月一括実行を開始します")
    print(f"[INFO] 対象期間: {args.start_year}-{args.start_month:02d} 〜 {args.end_year}-{args.end_month:02d}")
    print(f"[INFO] base_dir = {args.base_dir}")

    # 月ごとに create_monthly_max_stock_csv を呼び出す
    for year, month in month_iter(args.start_year, args.start_month, args.end_year, args.end_month):
        print(f"[INFO] === {year}-{month:02d} の集計を開始します ===")
        create_monthly_max_stock_csv(year=year, month=month, base_dir=args.base_dir)
        print(f"[INFO] === {year}-{month:02d} の集計が完了しました ===")

    print("[OK] すべての月の処理が完了しました。")


if __name__ == "__main__":
    main()

