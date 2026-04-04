# -*- coding: utf-8 -*-
"""
interlock_batch_mht_by_parking_all.py

TN2000「時間帯別入出庫日報」を日付指定で
全駐車場＋各駐車場について取得し、
最終的に CSV ファイルだけを保存するスクリプト。

・引数
    --date YYYY-MM-DD   取得対象日
    --download-dir DIR  CSV などを保存するルートフォルダ（省略時: downloads）
    --headless          ブラウザ非表示モード
"""

import argparse
import os
from dataclasses import dataclass
from datetime import datetime
from typing import List, Optional

import pandas as pd
import requests
from playwright.sync_api import (
    sync_playwright,
    TimeoutError as PlaywrightTimeoutError,
    Page,
    Frame,
)

BASE_URL = "http://192.168.1.200"
PARKING_TOP = f"{BASE_URL}/Interlock/Parking.htm"

# ★ 駐車場セレクトボックスの value → 表示名 を固定で定義
# （TN2000 側で順番や名前が変わったらここを書き換える）
PARKING_OPTIONS = [
    (0, "全駐車場"),
    (1, "南1駐車場"),
    (2, "南2駐車場"),
    (3, "南3駐車場"),
    (4, "南4駐車場"),
    (5, "南4B駐車場"),
    (6, "北1駐車場"),
    (7, "北2駐車場"),
    (8, "北3駐車場"),
]


@dataclass
class DownloadContext:
    target_date: datetime
    download_dir: str
    headless: bool


def ensure_dirs(root: str) -> dict:
    """
    ルートディレクトリの下に、必要なサブフォルダを用意する。
    return: {"root": ..., "mht": ..., "csv": ...}
    """
    root = root or "downloads"
    mht_dir = os.path.join(root, "")      # MHT はルート直下
    csv_dir = os.path.join(root, "csv")   # CSV 用サブフォルダ

    os.makedirs(root, exist_ok=True)
    os.makedirs(csv_dir, exist_ok=True)

    return {"root": root, "mht": mht_dir, "csv": csv_dir}


def extract_html_from_mht_bytes(data: bytes) -> str:
    """
    MHT(MHTML) のバイナリから <html>〜</html> 部分だけを取り出し、
    文字列として返す簡易パーサ。
    """
    lower = data.lower()
    start = lower.find(b"<html")
    if start == -1:
        raise ValueError("MHT 内に <html> が見つかりませんでした。")

    end = lower.rfind(b"</html>")
    if end == -1:
        # </html> が無くても、とりあえず末尾まで読む
        end = len(data)

    html_bytes = data[start:end + len(b"</html>")]

    # 文字コードは UTF-8 優先、ダメなら cp932 を試す
    for enc in ("utf-8", "cp932"):
        try:
            return html_bytes.decode(enc)
        except UnicodeDecodeError:
            continue

    # どうしてもダメならエラー
    raise UnicodeDecodeError("utf-8/cp932", html_bytes, 0, len(html_bytes), "decode failed")


def mht_to_csv(mht_path: str, csv_dir: str, base_name: str) -> str:
    """
    MHT ファイルから HTML を抽出し、最初の <table> を DataFrame として
    CSV に保存する。
    保存パスを返す。
    """
    with open(mht_path, "rb") as f:
        data = f.read()

    html = extract_html_from_mht_bytes(data)

    # HTML から table を抽出して DataFrame に変換
    # （lxml / html5lib が必要な場合があります）
    tables = pd.read_html(html)
    if not tables:
        raise ValueError(f"{mht_path} 内にテーブルが見つかりませんでした。")

    df = tables[0]

    csv_path = os.path.join(csv_dir, base_name + ".csv")
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    print(f"[INFO] [OK] CSV 変換完了: {csv_path}")

    return csv_path


def cleanup_files(*paths: str) -> None:
    """
    指定されたパスのファイルを削除する（存在しない場合は無視）。
    """
    for p in paths:
        if not p:
            continue
        if os.path.exists(p):
            try:
                os.remove(p)
                print(f"[INFO] 不要になったファイルを削除しました: {p}")
            except Exception as e:
                print(f"[WARN] ファイル削除に失敗しました: {p} ({e})")


def goto_parking_report(page: Page) -> None:
    """
    Parking.htm から『時間帯別入出庫日報』をクリックし、
    MainFrameset_s.aspx (レポート画面) まで遷移する。
    """
    print(f"[INFO] Parking.htm へアクセス... ({PARKING_TOP})")
    page.goto(PARKING_TOP, wait_until="load", timeout=30000)

    print("[INFO] 『時間帯別入出庫日報』リンクを探しています…")

    for i in range(40):
        frames: List[Frame] = page.frames
        for fr in frames:
            try:
                link = fr.get_by_text("時間帯別入出庫日報", exact=True)
                if link.count() > 0:
                    print(f"[INFO] メニューリンク発見: frame url={fr.url}  text=時間帯別入出庫日報")
                    link.first.click()
                    # 遷移完了を少し待つ
                    page.wait_for_timeout(2000)
                    break
            except Exception:
                continue
        else:
            # 見つからなかった → 少し待って再試行
            page.wait_for_timeout(500)
            continue
        break

    # MainFrameset_s.aspx を含むフレームが現れるまで待機
    print("[INFO] レポート画面 (IPSWebReport) を探しています…")
    for i in range(40):
        for fr in page.frames:
            if "MainFrameset" in (fr.url or ""):
                print(f"[INFO] レポート画面 URL: {fr.url}")
                return
        page.wait_for_timeout(500)

    raise RuntimeError("レポート画面 MainFrameset が見つかりませんでした。")


def find_condition_and_button_frames(page: Page) -> tuple[Frame, Frame]:
    """
    MainFrameset 内から条件入力フレームとボタンフレームを探す。
    """
    condition_frame: Optional[Frame] = None
    button_frame: Optional[Frame] = None

    for i in range(40):
        for fr in page.frames:
            if "ReportCondition.aspx" in (fr.url or ""):
                condition_frame = fr
            if "ButtonsWebForm.aspx" in (fr.url or "") and "buttons=Show" in fr.url:
                button_frame = fr

        if condition_frame and button_frame:
            print(f"[INFO] 条件フレーム発見: {condition_frame.url}")
            print(f"[INFO] ボタンフレーム発見: {button_frame.url}")
            return condition_frame, button_frame

        page.wait_for_timeout(500)

    raise RuntimeError("条件フレームまたはボタンフレームが見つかりませんでした。")


def set_date_and_parking(condition_frame: Frame, target_date: datetime, parking_value: int, parking_name: str) -> None:
    """
    条件フレーム上で日付と駐車場を設定する。
    日付入力欄や select の name / id は環境に合わせて調整が必要な場合があります。
    """
    jp_date = target_date.strftime("%Y/%m/%d")
    print(f"[INFO] 日付を設定します: {jp_date}")

    # 日付入力欄の候補をいくつか試す
    date_input_candidates = [
        "input[name='TextBox_FromDate']",
        "input[id*='FromDate']",
        "input[type='text']",
    ]

    filled = False
    for selector in date_input_candidates:
        try:
            loc = condition_frame.locator(selector)
            if loc.count() > 0:
                # とりあえず最初の要素に入力
                loc.nth(0).fill(jp_date)
                filled = True
                break
        except Exception:
            continue

    if not filled:
        print("[WARN] 日付入力欄が特定できませんでした。")

    # 駐車場 select を選択
    select_candidates = [
        "select[name='DropDownList_ParkNo']",
        "select[id*='ParkNo']",
        "select",
    ]

    selected = False
    for selector in select_candidates:
        try:
            sel = condition_frame.locator(selector)
            if sel.count() > 0:
                print(f"[INFO] 駐車場 option 件数: {sel.nth(0).locator('option').count()}")
                # value を直接指定
                sel.nth(0).select_option(str(parking_value))
                selected = True
                break
        except Exception:
            continue

    if not selected:
        print("[WARN] 駐車場 select が特定できませんでした。")


def click_show_and_wait_mht(page: Page, button_frame: Frame, timeout_ms: int = 30000) -> str:
    """
    ボタンフレーム上で ClickShow() を実行し、
    ページ全体で .mht へのリクエストが発生するまで待つ。
    戻り値は検出した .mht の URL。
    """
    mht_url_container = {"url": None}

    def on_request(request):
        url = request.url
        if url.lower().endswith(".mht"):
            mht_url_container["url"] = url
            print(f"[INFO] [HOOK] .mht リクエスト検知: {url}")

    page.on("request", on_request)

    print("[INFO] 『表示』ボタンが見つからないため、ClickShow() を直接実行してみます。")
    try:
        button_frame.evaluate("ClickShow();")
    except PlaywrightTimeoutError:
        print("[WARN] ClickShow() 実行で Timeout が発生しました。")

    print("[INFO] ネットワークリクエストから .mht URL を待機します…")

    # 一定時間 .mht URL がセットされるのを待つ
    waited = 0
    interval = 500
    while waited < timeout_ms:
        if mht_url_container["url"]:
            url = mht_url_container["url"]
            print(f"[INFO] [OK] ネットワークで .mht リクエストを検出: {url}")
            return url
        page.wait_for_timeout(interval)
        waited += interval

    raise TimeoutError("指定時間内に .mht リクエストが検出できませんでした。")


def download_mht(url: str, save_dir: str, base_name: str) -> str:
    """
    requests を使って .mht ファイルをダウンロードする。
    """
    os.makedirs(save_dir, exist_ok=True)
    filename = base_name + ".mht"
    path = os.path.join(save_dir, filename)

    print("[INFO] requests でダウンロード開始…")
    resp = requests.get(url, timeout=60)
    resp.raise_for_status()

    with open(path, "wb") as f:
        f.write(resp.content)

    print(f"[INFO] [OK] MHT ダウンロード完了: {path}")
    return path


def run(date_str: str, download_dir: str, headless: bool) -> None:
    target_date = datetime.strptime(date_str, "%Y-%m-%d")
    ctx = DownloadContext(target_date=target_date, download_dir=download_dir, headless=headless)

    paths = ensure_dirs(ctx.download_dir)

    print(f"[INFO] === 全駐車場DL 開始 === 日付={date_str}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=ctx.headless)
        page = browser.new_page(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; WOW64; Trident/7.0; rv:11.0) like Gecko"
            )
        )

        # Parking.htm → レポート画面へ
        goto_parking_report(page)

        # 条件フレーム & ボタンフレーム取得
        condition_frame, button_frame = find_condition_and_button_frames(page)

        # 駐車場ごとに処理
        total = len(PARKING_OPTIONS)
        for idx, (value, name) in enumerate(PARKING_OPTIONS, start=1):
            print("[INFO] ============================================")
            print(f"[INFO] {idx}/{total} 駐車場 '{name}' (value={value}) の処理を開始します")

            # 条件セット
            set_date_and_parking(condition_frame, ctx.target_date, value, name)

            # ClickShow → .mht URL を待機
            try:
                mht_url = click_show_and_wait_mht(page, button_frame, timeout_ms=30000)
            except Exception as e:
                print(f"[ERROR] .mht URL 取得に失敗しました: {e}")
                continue

            # .mht ダウンロード
            base_name = f"{ctx.target_date.strftime('%Y%m%d')}_{name}"
            mht_path = download_mht(mht_url, paths["mht"], base_name)

            # MHT → CSV 変換
            try:
                mht_to_csv(mht_path, paths["csv"], base_name)
            except Exception as e:
                print(f"[ERROR] MHT から CSV への変換に失敗しました: {e}")
            finally:
                # 中間ファイル（MHT）は削除
                cleanup_files(mht_path)

        browser.close()

    print("[INFO] === 全駐車場DL 完了 ===")


def main():
    parser = argparse.ArgumentParser(
        description="TN2000『時間帯別入出庫日報』を指定日付で全駐車場＋個別駐車場分まとめて CSV 取得するスクリプト"
    )
    parser.add_argument(
        "--date", required=True, help="取得対象日 (YYYY-MM-DD)"
    )
    parser.add_argument(
        "--download-dir",
        default="downloads",
        help="CSV を保存するルートディレクトリ（省略時: downloads）",
    )
    parser.add_argument(
        "--headless",
        action="store_true",
        help="ブラウザを非表示モードで起動する",
    )

    args = parser.parse_args()
    run(args.date, args.download_dir, args.headless)


if __name__ == "__main__":
    main()
