# -*- coding: utf-8 -*-
"""
interlock_batch_month.py
指定した「年・月」の1ヶ月分について、
interlock_batch_mht_by_parking_all.py を日付ごとに実行して
まとめてダウンロードを行うためのラッパースクリプト。
"""

import argparse
import subprocess
import sys
from datetime import date, timedelta
import calendar


def generate_month_dates(year: int, month: int):
    """
    指定された year, month の「その月に存在する全ての日付」を
    datetime.date オブジェクトのリストとして返す関数。
    例）2025年11月 → [2025-11-01, 2025-11-02, ..., 2025-11-30]
    """
    # その月が何日まであるかを calendar.monthrange で取得
    # 返り値は (その月の1日の曜日, その月の日数)
    _, last_day = calendar.monthrange(year, month)

    start = date(year, month, 1)
    end = date(year, month, last_day)

    dates = []
    current = start
    while current <= end:
        dates.append(current)
        current += timedelta(days=1)

    return dates


def run_daily_batch(target_date: date, headless: bool):
    """
    1日分のバッチ (interlock_batch_mht_by_parking_all.py) を呼び出す関数。
    target_date: 実行対象日 (datetime.date)
    headless: True のとき --headless オプションを付ける
    """
    date_str = target_date.strftime("%Y-%m-%d")
    print(f"[INFO] === {date_str} の全駐車場DL を開始します ===")

    # 呼び出すコマンドをリストで作成
    cmd = [
        sys.executable,  # 今動いているPython（.venvのpython）をそのまま使う
        "interlock_batch_mht_by_parking_all.py",
        "--date", date_str,
    ]

    if headless:
        cmd.append("--headless")

    print(f"[INFO] 実行コマンド: {' '.join(cmd)}")

    # 実際にサブプロセスとして起動
    result = subprocess.run(cmd)

    # returncode が 0 以外ならエラーとして扱う
    if result.returncode != 0:
        print(f"[ERROR] {date_str} の処理でエラー発生 (returncode={result.returncode})")
    else:
        print(f"[OK] {date_str} の処理が正常終了しました。")


def main():
    parser = argparse.ArgumentParser(
        description="指定した年・月の1ヶ月分をまとめてダウンロードするバッチ"
    )
    parser.add_argument(
        "--year", type=int, required=True, help="対象年（例: 2025）"
    )
    parser.add_argument(
        "--month", type=int, required=True, help="対象月（1〜12）"
    )
    parser.add_argument(
        "--start-day", type=int, default=None,
        help="開始日（任意）。指定しない場合は月初から。"
    )
    parser.add_argument(
        "--end-day", type=int, default=None,
        help="終了日（任意）。指定しない場合は月末まで。"
    )
    parser.add_argument(
        "--headless", action="store_true",
        help="ブラウザを非表示モードで実行する（通常の日次バッチと同じオプション）"
    )

    args = parser.parse_args()

    year = args.year
    month = args.month

    # 月の日数を取得
    _, last_day_of_month = calendar.monthrange(year, month)

    # start/end day の補正
    if args.start_day is None:
        start_day = 1
    else:
        start_day = max(1, min(args.start_day, last_day_of_month))

    if args.end_day is None:
        end_day = last_day_of_month
    else:
        end_day = max(1, min(args.end_day, last_day_of_month))

    if start_day > end_day:
        print("[ERROR] start-day が end-day より大きくなっています。指定を見直してください。")
        sys.exit(1)

    print(f"[INFO] 対象期間: {year}-{month:02d}-{start_day:02d} ～ {year}-{month:02d}-{end_day:02d}")

    # 対象となる全日付を作成
    all_dates = generate_month_dates(year, month)

    # start_day〜end_day の範囲に絞る
    target_dates = [d for d in all_dates if start_day <= d.day <= end_day]

    print(f"[INFO] 対象日数: {len(target_dates)} 日")

    # 日付ごとに日次バッチを起動
    for d in target_dates:
        run_daily_batch(d, headless=args.headless)

    print("[INFO] === 月次まとめDLバッチが終了しました ===")


if __name__ == "__main__":
    main()

