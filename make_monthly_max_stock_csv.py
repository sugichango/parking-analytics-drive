# -*- coding: utf-8 -*-
"""
make_monthly_max_stock_csv.py

Interlock から日別でダウンロードした「時間帯別入出庫日報」CSV をまとめて読み込み、
各日・各駐車場ごとに「合計在庫が最大となる時間帯の行」を 1 行だけ抜き出し、
1 ヶ月分を縦に並べた月次 CSV を駐車場別に作成するスクリプト。

・想定フォルダ構成
  base_dir/
    └─ csvYYYYMM/
         ├─ 20251001_全駐車場.csv
         ├─ 20251001_南1駐車場.csv
         ├─ 20251001_南2駐車場.csv
         ├─ ...
         └─ （本スクリプトが作る）南1all202510.csv など

・元の日別 CSV の列構造（概念）
    [何か余計な列があってもよい]
    時間帯 ("00-01" など)
    一般入庫
    一般出庫
    一般在庫
    定期入庫
    定期出庫
    定期在庫
    合計入庫
    合計出庫
    合計在庫  ← ここが「最大在庫」を探す対象
    事前精算利用台数

  ※ 時間帯の列位置はファイルによって 0 列目とは限らない前提。
     → 最初のファイルから「時間帯っぽい文字列」を含む列を自動検出する。
"""

import argparse
import glob
import os
import re
from datetime import datetime

import pandas as pd
import jpholiday  # ← これを追加

# 駐車場名（出力ファイル名に使う）
# 各日 CSV のファイル名にこの文字列が含まれている想定
PARKING_NAMES = [
    "全駐車場",
    "南1",
    "南2",
    "南3",
    "南4",
    "南4B",
    "北1",
    "北2",
    "北3",
]


# ----------------------------------------------------------------------
# ユーティリティ関数
# ----------------------------------------------------------------------
def read_csv_flexible(path: str) -> pd.DataFrame:
    """
    エンコーディングを自動判定しつつ CSV を読み込む。
    ・cp932 → utf-8-sig → utf-8 の順で試す
    ・header=None で読み込み、列は 0,1,2,... の整数インデックスにする
    """
    encodings = ["cp932", "utf-8-sig", "utf-8"]
    last_err = None
    for enc in encodings:
        try:
            return pd.read_csv(path, header=None, encoding=enc)
        except UnicodeDecodeError as e:
            last_err = e
    # ここまで来たら全部失敗
    raise UnicodeDecodeError(
        f"CSV 読み込みに失敗しました（cp932 / utf-8-sig / utf-8 いずれも不可）: {path}"
    ) from last_err


def extract_date_from_filename(path: str) -> datetime:
    """
    ファイル名から 8 桁の日付(YYYYMMDD)を抜き出して datetime にする。
    例: '20251001_南1駐車場.csv' → datetime(2025,10,1)
    """
    basename = os.path.basename(path)
    m = re.search(r"(\d{8})", basename)
    if not m:
        raise ValueError(f"ファイル名から日付(YYYYMMDD)が取得できません: {basename}")
    date_str = m.group(1)
    return datetime.strptime(date_str, "%Y%m%d")


def detect_time_column_index(df: pd.DataFrame) -> int:
    """
    DataFrame の中から「時間帯っぽい値（例: 00-01, 0-1, 08～09 など）」を含む列を探し、
    その列インデックス(int)を返す。

    - 全列を文字列化してチェック
    - 正規表現パターン:
        1〜2桁の数字 + 区切り記号(-, －, 〜, ~ など) + 1〜2桁の数字
    """
    pattern = r"\d{1,2}\s*[-－〜~]\s*\d{1,2}"

    for idx in range(df.shape[1]):
        s = df.iloc[:, idx].astype(str)
        if s.str.contains(pattern, na=False).any():
            print(f"[INFO] 時間帯っぽい値を持つ列を検出: 列インデックス={idx}")
            return idx

    raise ValueError(
        f"時間帯(00-01 など)を含む列が見つかりませんでした。列数={df.shape[1]}"
    )


def get_daily_max_row(
    df: pd.DataFrame,
    time_col_index: int,
    stock_col_index: int,
) -> pd.Series:
    """
    1 日分の DataFrame から、
    「時間帯列(time_col_index)が '00-01' などの形式の行」だけを対象に、
    在庫列(stock_col_index)が最大となる行(Series)を返す。

    df  : header=None で読み込んだ DataFrame を想定
    time_col_index  : 時間帯が入っている列のインデックス
    stock_col_index : 「合計在庫」が入っている列のインデックス
    """
    # 時間帯列を文字列として取り出し
    time_series = df.iloc[:, time_col_index].astype(str).str.strip()

    # 「00-01」「0-1」「08～09」「8~9」などを拾えるようなパターン
    pattern = r"\d{1,2}\s*[-－〜~]\s*\d{1,2}"

    mask = time_series.str.contains(pattern, na=False)
    hourly_df = df[mask].copy()

    if hourly_df.empty:
        raise ValueError("時間帯形式(00-01など)の行が見つかりません。")

    # 在庫列（合計在庫）を数値化（カンマ除去など）
    stock_series = (
        hourly_df.iloc[:, stock_col_index]
        .astype(str)
        .str.replace(",", "", regex=False)
        .str.strip()
    )
    stock_values = pd.to_numeric(stock_series, errors="coerce")

    if stock_values.isna().all():
        raise ValueError("在庫列がすべて数値に変換できませんでした。")

    # 最大値のインデックスを取得
    idx_max = stock_values.idxmax()

    # 最大となる行(Series)を返す
    return hourly_df.loc[idx_max]


# ----------------------------------------------------------------------
# メイン処理
# ----------------------------------------------------------------------
def create_monthly_max_stock_csv(year: int, month: int, base_dir: str = "downloads") -> None:
    """
    指定された年(year)・月(month)について、
    base_dir/csvYYYYMM/ 配下にある「日別 CSV」から、
    各駐車場ごとに「合計在庫が最大となる時間帯の行」を 1 日 1 行ずつ抜き出し、
    駐車場別に 1 ヶ月分まとめた CSV を作成する。

    出力ファイル名の例:
      ・全駐車場all202510.csv
      ・南1all202510.csv
      ・南2all202510.csv
      ・ ... など
    """
    yyyymm = f"{year}{month:02d}"
    target_dir = os.path.join(base_dir, f"csv{yyyymm}")

    if not os.path.isdir(target_dir):
        raise FileNotFoundError(f"対象フォルダが見つかりません: {target_dir}")

    pattern = os.path.join(target_dir, "*.csv")
    all_candidates = sorted(glob.glob(pattern))

    # 「8桁の日付(YYYYMMDD)をファイル名に含むものだけ」を日別 CSV とみなす
    daily_files = [
        f for f in all_candidates
        if re.search(r"\d{8}", os.path.basename(f))
    ]

    if not daily_files:
        raise FileNotFoundError(f"フォルダ内に日付付きCSVファイルが見つかりません: {pattern}")

    print(f"[INFO] 対象フォルダ: {target_dir}")
    print(f"[INFO] 総CSVファイル数（集計対象候補）: {len(daily_files)}")

    # ★ 時間帯列インデックス ＆ 合計在庫列インデックスを、
    #   最初の 1 ファイルから自動検出して固定する
    sample_df = read_csv_flexible(daily_files[0])
    time_col_index = detect_time_column_index(sample_df)
    stock_col_index = time_col_index + 9  # 時間帯の 9 列右を「合計在庫」とみなす

    # 11 列分（時間帯 + 10 列）を正式列として採用
    data_col_indices = list(range(time_col_index, time_col_index + 11))

    print(f"[INFO] 使用する時間帯列 index = {time_col_index}")
    print(f"[INFO] 使用する合計在庫列 index = {stock_col_index}")
    print(f"[INFO] 出力対象のデータ列 index 範囲 = {data_col_indices}")

    # 駐車場ごとにループ
    for parking_name in PARKING_NAMES:
        # この駐車場に対応する日別 CSV だけを抽出（ファイル名で厳密にマッチ）
        parking_files = []
        for f in daily_files:
            basename = os.path.basename(f)

            if parking_name == "全駐車場":
                # 全駐車場だけはファイル名が「YYYYMMDD_全駐車場.csv」
                pattern = rf"\d{{8}}_{re.escape(parking_name)}\.csv$"
            else:
                # 他の駐車場は「YYYYMMDD_南1駐車場.csv」など
                pattern = rf"\d{{8}}_{re.escape(parking_name)}駐車場\.csv$"

            if re.search(pattern, basename):
                parking_files.append(f)



        if not parking_files:
            print(f"[WARN] 駐車場 '{parking_name}' に該当するCSVファイルが見つかりません。スキップします。")
            continue

        print(f"[INFO] === 駐車場 '{parking_name}' の集計を開始 ===")
        print(f"[INFO] 対象ファイル数: {len(parking_files)}")

        rows = []

        for csv_file in sorted(parking_files):
            print(f"[INFO] 処理中: {csv_file}")

            # ファイル名から日付を取得
            file_date = extract_date_from_filename(csv_file)
            date_str = file_date.strftime("%Y-%m-%d")

            # CSV 読み込み
            df = read_csv_flexible(csv_file)

            # この日の「合計在庫が最大となる時間帯行」を取得
            max_row = get_daily_max_row(
                df,
                time_col_index=time_col_index,
                stock_col_index=stock_col_index,
            )

            # メタ情報（日付・駐車場）を付与
            s = max_row.copy()
            s["日付"] = date_str
            s["駐車場"] = parking_name

            rows.append(s)

        if not rows:
            print(f"[WARN] 駐車場 '{parking_name}' では有効な行が 1 件も得られませんでした。")
            continue

        # 月次 DataFrame を作成
        result_df = pd.DataFrame(rows)

        # 必要なデータ列（時間帯〜事前精算利用台数）だけを残し、
        # その左に「日付」「駐車場」を付ける
        missing = [idx for idx in data_col_indices if idx not in result_df.columns]
        if missing:
            raise ValueError(
                f"期待するデータ列 {data_col_indices} のうち、"
                f"存在しない列があります: {missing}"
            )

        result_df = result_df[["日付", "駐車場"] + data_col_indices]

        # 人間が読みやすい列名に変更
        new_cols = [
            "日付",
            "駐車場",
            "時間帯",
            "一般入庫",
            "一般出庫",
            "一般在庫",
            "定期入庫",
            "定期出庫",
            "定期在庫",
            "合計入庫",
            "合計出庫",
            "合計在庫",
            "事前精算利用台数",
        ]
        result_df.columns = new_cols

        # ===== 追加：曜日・土日祝区分を付与 =====
        # 「日付」列を datetime に変換
        result_df["日付_dt"] = pd.to_datetime(result_df["日付"])

        # 曜日番号 (0=月 ... 6=日) を日本語の曜日に変換
        weekday_map = {0: "月", 1: "火", 2: "水", 3: "木", 4: "金", 5: "土", 6: "日"}
        weekday_num = result_df["日付_dt"].dt.weekday
        result_df["曜日"] = weekday_num.map(weekday_map)

        # 土日 または 日本の祝日なら「土日祝」、それ以外は「平日」
        def _classify_day(d):
            if d.weekday() >= 5 or jpholiday.is_holiday(d):
                return "土日祝"
            return "平日"

        result_df["土日祝区分"] = result_df["日付_dt"].apply(_classify_day)

        # 中間の datetime 列は不要なので削除
        result_df = result_df.drop(columns=["日付_dt"])
 
        # 出力ファイル名: 例) 南1all202510.csv
        out_name = f"{parking_name}all{yyyymm}.csv"
        out_path = os.path.join(target_dir, out_name)

        result_df.to_csv(out_path, index=False, encoding="cp932")
        print(f"[OK] 駐車場 '{parking_name}' の月次まとめCSVを作成しました: {out_path}")


# ----------------------------------------------------------------------
# エントリポイント
# ----------------------------------------------------------------------
def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="時間帯別入出庫日報CSVから、駐車場別・1ヶ月分の最大在庫時間帯CSVを作成するスクリプト"
    )
    parser.add_argument(
        "--year",
        type=int,
        required=True,
        help="対象年 (例: 2025)",
    )
    parser.add_argument(
        "--month",
        type=int,
        required=True,
        help="対象月 (1〜12)",
    )
    parser.add_argument(
        "--base_dir",
        type=str,
        default="downloads",
        help="CSVフォルダのベースディレクトリ (デフォルト: downloads)",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    create_monthly_max_stock_csv(args.year, args.month, args.base_dir)


if __name__ == "__main__":
    main()
