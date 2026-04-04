# -*- coding: utf-8 -*-
"""
TN2000 MHT -> CSV converter（簡易版）

- MHTの中で「最初に見つかった sheetXXX の <table>」だけを抽出
- それを1つだけ CSV に保存する
- グラフシート（sheet002など）は無視して、とにかく最初のシートだけ使う想定
"""

import argparse
import email
import logging
from io import StringIO
from pathlib import Path
from typing import List

import pandas as pd


# ===== ログ設定 =====
logger = logging.getLogger(__name__)
logging.basicConfig(
    level=logging.INFO,
    format="[%(levelname)s] %(message)s"
)


# ===== 設定 =====
# MHTを探すフォルダ（TN2000の出力先）
DEFAULT_EXPORT_DIR = Path(r"C:\Users\Public\TN2000_Exports")

# CSVを保存するフォルダ
DEFAULT_CSV_DIR = Path(
    r"C:\Users\sugitamasahiko\Documents\parking_system\downloads\csv"
)


# ===== 基本ユーティリティ =====
def find_latest_mht(export_dir: Path) -> Path:
    """export_dir 内で一番新しい .mht ファイルを返す"""
    logger.info(f"MHT を探すフォルダ: {export_dir}")
    mht_files: List[Path] = list(export_dir.glob("*.mht"))
    if not mht_files:
        raise FileNotFoundError(f".mht がありません: {export_dir}")
    latest = max(mht_files, key=lambda p: p.stat().st_mtime)
    logger.info(f"最新の MHT: {latest.name}")
    return latest


def decode_html_bytes(raw: bytes) -> str:
    """MHTパートのバイト列をHTML文字列にデコード"""
    for enc in ("utf-8", "cp932", "shift_jis", "utf-8-sig", "latin1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            pass
    return raw.decode("utf-8", errors="replace")


# ===== MHT → 最初のテーブルだけ抽出 =====
def extract_first_table_from_mht(mht_path: Path) -> pd.DataFrame:
    """
    MHTの中から「最初に見つかった sheetXXX の <table>」だけを取り出して DataFrame にする。
    それ以外のシート（sheet002 など）は無視。
    """
    logger.info(f"MHT を解析中: {mht_path.name}")

    with mht_path.open("rb") as f:
        msg = email.message_from_binary_file(f)

    for part in msg.walk():
        if part.get_content_type() != "text/html":
            continue

        filename = part.get_filename() or ""
        content_location = part.get("Content-Location", "")
        name = (filename or content_location or "").lower()

        # sheet が名前に入っていないHTMLは無視
        if "sheet" not in name:
            continue

        raw = part.get_payload(decode=True) or b""
        html = decode_html_bytes(raw)

        try:
            dfs = pd.read_html(StringIO(html), header=0)
        except ValueError:
            continue

        if not dfs:
            continue

        df = dfs[0]
        logger.info(f"最初のシートを取得: name={name}, shape={df.shape}")
        return df

    # ここまで来たらシートが1つも見つからなかった
    raise RuntimeError("MHT 内に sheetXXX のテーブルが見つかりませんでした。")


# ===== CSV 保存 =====
def save_single_csv(df: pd.DataFrame, mht_path: Path, csv_dir: Path) -> Path:
    """
    DataFrame を1つだけ CSV として保存する。
    ファイル名: <MHTベース名>_sheet001.csv
    """
    csv_dir.mkdir(parents=True, exist_ok=True)
    base = mht_path.stem
    out_path = csv_dir / f"{base}_sheet001.csv"
    df.to_csv(out_path, index=False, encoding="utf-8-sig")
    logger.info(f"CSV 保存完了: {out_path}")
    return out_path


# ===== メイン処理 =====
def main():
    parser = argparse.ArgumentParser(description="TN2000 MHT から最初のシートだけ CSV に変換するツール")
    parser.add_argument("--mht", type=str, default=None, help="対象 MHT ファイルパス（省略時は最新のMHTを使用）")
    parser.add_argument("--export-dir", type=str, default=str(DEFAULT_EXPORT_DIR), help="MHT を探すフォルダ")
    parser.add_argument("--csv-dir", type=str, default=str(DEFAULT_CSV_DIR), help="CSV を保存するフォルダ")
    args = parser.parse_args()

    export_dir = Path(args.export_dir)
    csv_dir = Path(args.csv_dir)

    # 1) 対象MHTを決定
    if args.mht:
        mht_path = Path(args.mht)
        if not mht_path.exists():
            raise FileNotFoundError(f"--mht で指定されたファイルが存在しません: {mht_path}")
    else:
        mht_path = find_latest_mht(export_dir)

    # 2) MHTから「最初の sheet のテーブル」だけ抽出
    df = extract_first_table_from_mht(mht_path)

    # 3) 1つだけCSV保存
    save_single_csv(df, mht_path, csv_dir)

    logger.info("[OK] MHT から CSV の出力が完了しました。")


if __name__ == "__main__":
    main()
