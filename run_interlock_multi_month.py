# -*- coding: utf-8 -*-
"""
run_interlock_multi_month.py

interlock_batch_mht_by_parking_all_csv3.py を、
指定した複数の「年月」について順番に実行するラッパースクリプト。

★ 元の interlock_batch_mht_by_parking_all_csv3.py は一切変更しません。
"""

import argparse
import os
import subprocess
import sys
from datetime import date


def iter_months(start_year: int, start_month: int, end_year: int, end_month: int):
    """
    開始年月～終了年月までを 1ヶ月ずつ進めながら (year, month) を返すジェネレータ。

    例:
      start_year=2025, start_month=10
      end_year=2026, end_month=3
      → (2025,10), (2025,11), (2025,12), (2026,1), (2026,2), (2026,3)
    """
    y, m = start_year, start_month
    while (y < end_year) or (y == end_year and m <= end_month):
        yield y, m
        # 月を +1 して、13 月になったら翌年の 1 月へ
        m += 1
        if m > 12:
            m = 1
            y += 1


def main():
    parser = argparse.ArgumentParser(
        description="interlock_batch_mht_by_parking_all_csv3.py を複数月まとめて実行するスクリプト"
    )

    # 開始年月
    parser.add_argument("--start-year", type=int, required=True, help="開始年 (例: 2025)")
    parser.add_argument("--start-month", type=int, required=True, help="開始月 (1〜12)")

    # 終了年月
    parser.add_argument("--end-year", type=int, required=True, help="終了年 (例: 2025)")
    parser.add_argument("--end-month", type=int, required=True, help="終了月 (1〜12)")

    # 既存スクリプトに渡すオプション
    parser.add_argument(
        "--download-dir",
        default="downloads",
        help="既存スクリプトに渡す --download-dir （省略時: downloads）"
    )
    parser.add_argument(
        "--headless",
        action="store_true",
        help="既存スクリプトの --headless を有効にする場合に指定"
    )
    parser.add_argument(
        "--start-day",
        type=int,
        help="（任意）月内の開始日。指定しなければ 1 日から"
    )
    parser.add_argument(
        "--end-day",
        type=int,
        help="（任意）月内の終了日。指定しなければその月末まで"
    )

    args = parser.parse_args()

    # このファイルと同じフォルダに interlock_batch_mht_by_parking_all_csv3.py がある前提
    script_dir = os.path.dirname(os.path.abspath(__file__))
    target_script = os.path.join(script_dir, "interlock_batch_mht_by_parking_all_csv3.py")

    if not os.path.exists(target_script):
        print(f"[ERROR] interlock_batch_mht_by_parking_all_csv3.py が見つかりません: {target_script}")
        sys.exit(1)

    print("===================================================")
    print("[INFO] 複数月一括実行を開始します")
    print(f"[INFO] 対象期間: {args.start_year}-{args.start_month:02d} 〜 "
          f"{args.end_year}-{args.end_month:02d}")
    print("===================================================")

    for year, month in iter_months(args.start_year, args.start_month, args.end_year, args.end_month):
        print("---------------------------------------------------")
        print(f"[INFO] {year}-{month:02d} の処理を開始します")

        # 既存スクリプトに渡すコマンドを組み立て
        cmd = [
            sys.executable,
            target_script,
            "--year", str(year),
            "--month", str(month),
            "--download-dir", args.download_dir,
        ]

        # 任意の開始日・終了日が指定されていたらそのまま渡す
        if args.start_day is not None:
            cmd.extend(["--start-day", str(args.start_day)])
        if args.end_day is not None:
            cmd.extend(["--end-day", str(args.end_day)])

        # headless オプション
        if args.headless:
            cmd.append("--headless")

        print(f"[INFO] 実行コマンド: {' '.join(cmd)}")

        # サブプロセスとして実行
        result = subprocess.run(cmd)

        if result.returncode != 0:
            print(f"[ERROR] {year}-{month:02d} の処理でエラー終了しました (returncode={result.returncode})")
            # エラーが出た月以降も続けたいなら「continue」に変更
            break

        print(f"[OK] {year}-{month:02d} の処理が正常に完了しました。")

    print("===================================================")
    print("[INFO] 複数月一括実行が終了しました")
    print("===================================================")


if __name__ == "__main__":
    main()

