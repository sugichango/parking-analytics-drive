import argparse
import subprocess
import sys


def iter_months(start_year: int, start_month: int, end_year: int, end_month: int):
    """
    開始年月 ～ 終了年月 までを 1ヶ月ずつ進めながら (year, month) のタプルを返すジェネレータ
    例: (2024, 4) ～ (2024, 6) -> (2024,4), (2024,5), (2024,6)
    """
    y, m = start_year, start_month

    # (年,月) を 1ヶ月ずつ進めていくループ
    while (y < end_year) or (y == end_year and m <= end_month):
        yield y, m
        if m == 12:
            y += 1
            m = 1
        else:
            m += 1


def main():
    parser = argparse.ArgumentParser(
        description="make_monthly_excel_by_parking_with_avg.py を複数月分まとめて実行するスクリプト"
    )

    parser.add_argument("--start-year", type=int, required=True, help="開始年 (例: 2024)")
    parser.add_argument("--start-month", type=int, required=True, help="開始月 (例: 4)")
    parser.add_argument("--end-year", type=int, required=True, help="終了年 (例: 2024)")
    parser.add_argument("--end-month", type=int, required=True, help="終了月 (例: 6)")
    parser.add_argument(
        "--csv-base-dir",
        type=str,
        required=True,
        help="日別CSVが入っている月別フォルダの親ディレクトリ (例: downloads\\csv)",
    )

    args = parser.parse_args()

    # 入力チェック（開始年月 <= 終了年月 になっているか）
    if (args.start_year, args.start_month) > (args.end_year, args.end_month):
        print("[ERROR] 開始年月が終了年月より後になっています。指定を見直してください。")
        sys.exit(1)

    # 対象となる各 (year, month) について、もとの単月スクリプトを順番に実行
    for year, month in iter_months(args.start_year, args.start_month,
                                  args.end_year, args.end_month):

        print(f"[INFO] === {year}-{month:02d} の処理を開始します ===")

        # 実行コマンドを作成
        # ここで「もとのコード」のファイル名を指定する
        cmd = [
            sys.executable,  # 今動いているPython (例: .venv\\Scripts\\python.exe)
            "make_monthly_excel_by_parking_with_avg.py",
            "--year", str(year),
            "--month", str(month),
            "--csv-base-dir", args.csv_base_dir,
        ]

        print(f"[INFO] 実行コマンド: {' '.join(cmd)}")

        try:
            # もとの単月スクリプトを実行
            subprocess.run(cmd, check=True)
            print(f"[INFO] {year}-{month:02d} の処理が正常に終了しました。")
        except subprocess.CalledProcessError as e:
            print(f"[ERROR] {year}-{month:02d} の処理中にエラーが発生しました。コード={e.returncode}")
            # 必要に応じてここで sys.exit(1) して全体を止めてもよいです
            # 今回は次の月の処理も続けるようにしておきます
            continue

    print("[INFO] すべての指定月の処理が終了しました。")


if __name__ == "__main__":
    main()

