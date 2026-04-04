# -*- coding: utf-8 -*-
"""
interlock_batch_mht_by_parking_all_csv3.py

★ interlock_batch_mht_by_parking_all_csv2.py をベースに、
  Excel/MHT を生成元フォルダから 月別フォルダへ安全に「移動」する機能のみ追加した版。

★ 元スクリプトの処理ロジック（Playwright実行・CSV変換）は一切変更していません。
"""

import argparse
import calendar
import glob
import os
import subprocess
import sys
from dataclasses import dataclass
from datetime import date, datetime
from typing import List

import pandas as pd


# -----------------------------
# ▼ 新規追加：月フォルダ生成
# -----------------------------
def ensure_month_dir(base: str, prefix: str, target_date: date) -> str:
    """
    base/prefixYYYYMM フォルダを作り、そのパスを返す。
    例: ensure_month_dir("downloads/excel", "excel", 2025-10-01)
        → "downloads/excel/excel202510"
    """
    yyyymm = target_date.strftime("%Y%m")
    dirname = f"{prefix}{yyyymm}"
    full = os.path.join(base, dirname)
    os.makedirs(full, exist_ok=True)
    return full


@dataclass
class Context:
    headless: bool
    download_root: str
    script_dir: str


def ensure_dirs(root: str) -> str:
    """
    downloads/csv のベースフォルダだけ作成（CSV保存先本体ではない）
    """
    root = root or "downloads"
    csv_dir = os.path.join(root, "csv")
    os.makedirs(csv_dir, exist_ok=True)
    return csv_dir


def build_dates(args: argparse.Namespace) -> List[date]:
    dates: List[date] = []

    if args.date:
        d = datetime.strptime(args.date, "%Y-%m-%d").date()
        print(f"[INFO] 単一日モードで実行します: {d}")
        dates.append(d)
        return dates

    if args.year and args.month:
        y = args.year
        m = args.month
        _, last_day = calendar.monthrange(y, m)
        start_day = args.start_day if args.start_day else 1
        end_day = args.end_day if args.end_day else last_day

        if start_day < 1:
            start_day = 1
        if end_day > last_day:
            end_day = last_day

        print(f"[INFO] 月次モードで実行します: {y}-{m:02d} {start_day}日〜{end_day}日")

        for d in range(start_day, end_day + 1):
            dates.append(date(y, m, d))

        print(f"[INFO] 対象日数: {len(dates)} 日")
        return dates

    raise SystemExit("単一日 (--date) または 月次 (--year --month) を指定してください。")


def run_daily_excel_batch(ctx: Context, target_date: date) -> bool:
    script_path = os.path.join(ctx.script_dir, "interlock_batch_mht_by_parking_all.py")
    if not os.path.exists(script_path):
        print(f"[ERROR] 日次バッチスクリプトが見つかりません: {script_path}")
        return False

    date_str = target_date.strftime("%Y-%m-%d")

    cmd = [
        sys.executable,
        script_path,
        "--date", date_str,
        "--download-dir", ctx.download_root,
    ]
    if ctx.headless:
        cmd.append("--headless")

    print(f"[INFO] 日次バッチを実行します: {' '.join(cmd)}")

    try:
        result = subprocess.run(cmd)
    except Exception as e:
        print(f"[ERROR] 日次バッチ実行時例外: {e}")
        return False

    if result.returncode != 0:
        print(f"[ERROR] 日次バッチがエラー終了しました: returncode={result.returncode}")
        return False

    print(f"[OK] 日次バッチ正常終了: {date_str}")
    return True


# -------------------------------------------------------
# ▼ CSV 変換 ＋ Excel/mht の月フォルダ移動処理（今回の追加本体）
# -------------------------------------------------------
def convert_excel_to_csv_for_date(ctx: Context, target_date: date) -> None:
    """
    ★ excel/mht を本体スクリプトが保存した後、
      downloads/excel → downloads/excel/excelYYYYMM
      downloads/mht   → downloads/mht/mhtYYYYMM
      に安全に移動する。

    その後、移動先の Excel を CSV に変換する。
    """

    ymd = target_date.strftime("%Y%m%d")

    excel_base = os.path.join(ctx.download_root, "excel")
    csv_base   = os.path.join(ctx.download_root, "csv")
    mht_base   = os.path.join(ctx.download_root, "mht")

    # 月別フォルダ（新規作成）
    excel_month_dir = ensure_month_dir(excel_base, "excel", target_date)
    csv_month_dir   = ensure_month_dir(csv_base, "csv", target_date)
    mht_month_dir   = ensure_month_dir(mht_base, "mht", target_date)

    # ▼ まず downloads/excel にある Excel を取得
    excel_pattern = os.path.join(excel_base, f"{ymd}_*.xlsx")
    excel_files = glob.glob(excel_pattern)

    if not excel_files:
        print(f"[WARN] Excel が見つかりません: {excel_pattern}")

    # ----------------------------------
    # ★ Excel を 月別フォルダへ移動
    # ----------------------------------
    for src in excel_files:
        base = os.path.basename(src)
        dst = os.path.join(excel_month_dir, base)

        if not os.path.exists(dst):
            try:
                os.replace(src, dst)
                print(f"[INFO] Excel 移動: {src} → {dst}")
            except Exception as e:
                print(f"[ERROR] Excel 移動失敗: {e}")

    # ▼ 移動先の Excel を CSV 変換対象にする
    excel_month_pattern = os.path.join(excel_month_dir, f"{ymd}_*.xlsx")
    excel_files = glob.glob(excel_month_pattern)

    # -------------------------------
    # mht は downloads 直下に生成されるため、そこから削除する
    # -------------------------------
    download_root = ctx.download_root  # 例: "downloads"
    mht_pattern = os.path.join(download_root, f"{ymd}_*.mht")
    mht_files = glob.glob(mht_pattern)

    if not mht_files:
        print(f"[INFO] 削除対象の MHT はありません: {mht_pattern}")
    else:
        for src in mht_files:
            try:
                os.remove(src)
                print(f"[INFO] MHT 削除: {src}")
            except Exception as e:
                print(f"[ERROR] MHT 削除失敗: {e}")
 
    # ---------------------------
    # ▼ CSV 変換処理（本来の機能）
    # ---------------------------
    if not excel_files:
        print(f"[WARN] 移動後 Excel が見つかりません: {excel_month_pattern}")
        return

    print(f"[INFO] Excel→CSV 変換開始（{len(excel_files)}ファイル）")

    for excel_path in excel_files:
        base_no_ext = os.path.splitext(os.path.basename(excel_path))[0]
        csv_path = os.path.join(csv_month_dir, base_no_ext + ".csv")

        print(f"[INFO] 変換: {excel_path} → {csv_path}")

        try:
            df = pd.read_excel(excel_path)
            df.to_csv(csv_path, index=False, encoding="utf-8-sig")
        except Exception as e:
            print(f"[ERROR] CSV 変換失敗 ({excel_path}): {e}")

    print(f"[INFO] Excel→CSV 完了: {target_date}")


def run_for_dates(ctx: Context, dates: List[date]) -> None:
    total = len(dates)
    for i, d in enumerate(dates, 1):
        print("========================================================")
        print(f"[INFO] === {d} の処理開始 ({i}/{total}) ===")

        if not run_daily_excel_batch(ctx, d):
            print(f"[WARN] 日次バッチ失敗のため CSV変換スキップ: {d}")
            continue

        convert_excel_to_csv_for_date(ctx, d)

        print(f"[INFO] === {d} の処理完了 ===")

    print("[INFO] === 全日程の処理が完了しました ===")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--date")
    parser.add_argument("--year", type=int)
    parser.add_argument("--month", type=int)
    parser.add_argument("--start-day", type=int)
    parser.add_argument("--end-day", type=int)
    parser.add_argument("--download-dir", default="downloads")
    parser.add_argument("--headless", action="store_true")
    args = parser.parse_args()

    dates = build_dates(args)
    script_dir = os.path.dirname(os.path.abspath(__file__))

    ctx = Context(
        headless=args.headless,
        download_root=args.download_dir,
        script_dir=script_dir,
    )

    run_for_dates(ctx, dates)


if __name__ == "__main__":
    main()

