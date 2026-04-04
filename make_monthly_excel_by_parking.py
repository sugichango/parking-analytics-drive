import os
import re
import argparse
from collections import defaultdict
from datetime import datetime, date

import pandas as pd
import jpholiday


def parse_filename(filename):
    """S
    ファイル名から 日付（YYYYMMDD） と 駐車場名 を取り出す。

    例: '20240401_全駐車場.csv' -> ('20240401', '全駐車場')
    """
    pattern = re.compile(r'(\d{8})_(.+)\.csv$', re.IGNORECASE)
    m = pattern.match(filename)
    if not m:
        return None, None
    date_str, parking_name = m.groups()
    return date_str, parking_name


def sanitize_filename(name):
    """
    Windows のファイル名に使えない文字を '_' に置き換える。
    """
    return re.sub(r'[\\/:*?"<>|]', '_', name)


def collect_csv_by_parking(target_dir):
    """
    指定フォルダ内の CSV を読み込み、
    { 駐車場名: [(date_str, csv_path), ...], ... } という辞書を作る。
    """
    parking_dict = defaultdict(list)

    for fname in os.listdir(target_dir):
        if not fname.lower().endswith(".csv"):
            continue

        date_str, parking_name = parse_filename(fname)
        if date_str is None:
            print(f"[WARN] 想定外のファイル名のためスキップ: {fname}")
            continue

        full_path = os.path.join(target_dir, fname)
        parking_dict[parking_name].append((date_str, full_path))

    for parking_name in parking_dict:
        parking_dict[parking_name].sort(key=lambda x: x[0])

    return parking_dict


def _convert_mixed_column_to_numeric(col: pd.Series) -> pd.Series:
    """数字っぽいところだけ数値にする。"""
    if col.dtype != object:
        return col

    s = (
        col.astype(str)
        .str.replace("\u3000", " ", regex=False)  # 全角スペース
        .str.strip()
        .str.replace(",", "", regex=False)       # カンマ
    )
    num = pd.to_numeric(s, errors="coerce")
    if num.notna().sum() == 0:
        return col

    out = col.copy()
    out.loc[num.notna()] = num.loc[num.notna()]
    return out


def read_csv_flex(csv_path):
    """
    文字コードをいくつか試しながら CSV を読む。
    読めなければ None。
    読めたら「数字っぽい列」は数値型に変換。
    """
    encodings_to_try = ["cp932", "utf-8-sig", "utf-8"]
    last_err = None

    for enc in encodings_to_try:
        try:
            df = pd.read_csv(csv_path, encoding=enc)
            print(f"[INFO] 読み込み成功: {csv_path} (encoding={enc})")

            for col_name in df.columns:
                df[col_name] = _convert_mixed_column_to_numeric(df[col_name])

            return df
        except UnicodeDecodeError as e:
            print(f"[WARN] encoding={enc} では読めませんでした: {csv_path} -> {e}")
            last_err = e
        except Exception as e:
            print(f"[WARN] encoding={enc} で別のエラー: {csv_path} -> {e}")
            last_err = e

    print(f"[ERROR] CSV読込に失敗しました: {csv_path}\n  最後のエラー: {last_err}")
    return None


def get_day_type(date_str: str) -> str:
    """
    'YYYYMMDD' から「平日」or「休日」を返す。
    休日 = 土日 or 祝日(jpholiday)
    """
    d: date = datetime.strptime(date_str, "%Y%m%d").date()
    if d.weekday() >= 5 or jpholiday.is_holiday(d):
        return "休日"
    else:
        return "平日"


def export_to_excel(parking_dict, output_dir, year, month):
    """
    駐車場ごとの CSV グループから、Excel ファイルを出力する。
    1駐車場 = 1つのExcel
    1日 = 1シート
    各シートの A列=日付, B列=曜日区分（平日/休日）
    """
    os.makedirs(output_dir, exist_ok=True)

    for parking_name, records in parking_dict.items():
        dfs = []
        for date_str, csv_path in records:
            df = read_csv_flex(csv_path)
            if df is None:
                print(f"[WARN] 読み込めないため、この日付のシートをスキップします: {csv_path}")
                continue

            day_type = get_day_type(date_str)
            dt = datetime.strptime(date_str, "%Y%m%d")

            # ★ ここで必ず先頭2列として追加する ★
            df = df.copy()
            df.insert(0, "曜日区分", day_type)
            df.insert(0, "日付", dt)

            print(f"[DEBUG] {csv_path} の列名: {list(df.columns)}")

            dfs.append((date_str, df))

        if not dfs:
            print(f"[WARN] 有効なCSVが無かったため、この駐車場のExcelは作成しません: {parking_name}")
            continue

        safe_parking_name = sanitize_filename(parking_name)
        excel_filename = f"{year}{month:02d}_{safe_parking_name}.xlsx"
        excel_path = os.path.join(output_dir, excel_filename)

        print(f"[INFO] 出力中: {excel_path}  (シート数: {len(dfs)})")

        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            for date_str, df in dfs:
                sheet_name = date_str[4:]  # '20240401' -> '0401'
                df.to_excel(writer, sheet_name=sheet_name, index=False)


def main():
    parser = argparse.ArgumentParser(
        description="1ヶ月分の日毎CSVファイルを、駐車場ごとに1つのExcelにまとめるスクリプト"
    )
    parser.add_argument("--year", type=int, required=True)
    parser.add_argument("--month", type=int, required=True)
    parser.add_argument(
        "--csv-base-dir",
        type=str,
        default=r"C:\Users\sugitamasahiko\Documents\parking_system\downloads\csv",
    )
    parser.add_argument(
        "--output-dir",
        type=str,
        default=None,
    )

    args = parser.parse_args()

    target_dir = os.path.join(args.csv_base_dir, f"csv{args.year}{args.month:02d}")
    if args.output_dir is None:
        output_dir = os.path.join(args.csv_base_dir, f"excel{args.year}{args.month:02d}")
    else:
        output_dir = args.output_dir

    print(f"[INFO] 対象フォルダ: {target_dir}")
    print(f"[INFO] 出力フォルダ: {output_dir}")

    if not os.path.isdir(target_dir):
        raise FileNotFoundError(f"対象フォルダが存在しません: {target_dir}")

    parking_dict = collect_csv_by_parking(target_dir)

    if not parking_dict:
        print("[WARN] 対象CSVが見つかりませんでした。")
        return

    export_to_excel(parking_dict, output_dir, args.year, args.month)

    print("[INFO] 完了しました。")


if __name__ == "__main__":
    main()
