# -*- coding: utf-8 -*-
r"""
interlock_batch_mht_by_parking_all_csv.py

TN-2000「時間帯別入出庫日報」画面から、
指定日の全駐車場分の MHT を自動ダウンロードし、
あわせて CSV ファイルも作成するスクリプト
（ネットワーク監視＋MHTダウンロード＋CSV変換）

使い方:
  .venv\Scripts\python.exe interlock_batch_mht_by_parking_all_csv.py --date 2025-11-10 --headless
"""

import argparse
import datetime as dt
import logging
import os
import sys
import time
from urllib.parse import urlparse
from email import message_from_bytes
import unicodedata
from io import StringIO

import requests
import pandas as pd
from playwright.sync_api import sync_playwright

# ==========================
# 設定値
# ==========================

BASE_URL = "http://192.168.1.200"
PARKING_TOP_URL = f"{BASE_URL}/Interlock/Parking.htm"

# IE モード偽装用 User-Agent
IE_LIKE_UA = (
    "Mozilla/5.0 (Windows NT 10.0; WOW64; Trident/7.0; rv:11.0) like Gecko"
)

LOG_FORMAT = "[%(levelname)s] %(message)s"


# ==========================
# ロガー初期化
# ==========================

logger = logging.getLogger("interlock_batch_csv")
logger.setLevel(logging.INFO)
handler = logging.StreamHandler(sys.stdout)
handler.setFormatter(logging.Formatter(LOG_FORMAT))
logger.handlers.clear()
logger.addHandler(handler)


# ==========================
# ユーティリティ
# ==========================

def ensure_dir(path: str) -> str:
    """ディレクトリが存在しなければ作成して、そのパスを返す"""
    if not path:
        path = os.path.join(os.getcwd(), "downloads")
    os.makedirs(path, exist_ok=True)
    return path


def build_cookie_header(context, url: str) -> str:
    """
    Playwright の context.cookies() から Cookie ヘッダー文字列を生成する。
    """
    parsed = urlparse(url)
    cookie_url = f"{parsed.scheme}://{parsed.netloc}"
    cookies = context.cookies(cookie_url)
    parts = []
    for c in cookies:
        name = c.get("name")
        value = c.get("value")
        if name is not None and value is not None:
            parts.append(f"{name}={value}")
    return "; ".join(parts)


# ==========================
# 画面操作系
# ==========================

def open_parking_top(context, headless: bool):
    """Parking.htm を開いて、その page オブジェクトを返す。"""
    logger.info("Parking.htm へアクセス... (%s)", PARKING_TOP_URL)
    page = context.new_page()
    page.goto(PARKING_TOP_URL, wait_until="load", timeout=30_000)
    return page


def click_report_menu_and_get_page(context, top_page):
    """
    「時間帯別入出庫日報」メニューをクリックし、
    レポート画面の page オブジェクトを返す。
    """
    logger.info("『時間帯別入出庫日報』リンクを探しています…")

    target_locator = None
    for attempt in range(1, 41):
        frames = top_page.frames
        logger.debug("[DEBUG] メニュー探索 %d/40 frame数=%d", attempt, len(frames))

        for f in frames:
            locator = f.get_by_text("時間帯別入出庫日報", exact=False)
            if locator.count() > 0:
                target_locator = locator.first
                logger.info(
                    "メニューリンク発見: frame url=%s  text=%s",
                    f.url,
                    target_locator.inner_text().strip()
                )
                break

        if target_locator:
            break

        top_page.wait_for_timeout(500)

    if not target_locator:
        raise RuntimeError("『時間帯別入出庫日報』メニューリンクが見つかりませんでした。")

    # メニュークリック
    target_locator.click()

    # context 内のページから "IPSWebReport" を含む URL の page を探す
    logger.info("レポート画面 (IPSWebReport) を探しています…")
    report_page = None
    for attempt in range(1, 41):
        for p in context.pages:
            if p.is_closed():
                continue
            url = p.url or ""
            if "/IPSWebReport/" in url:
                report_page = p
                break

        if report_page:
            break

        top_page.wait_for_timeout(500)

    if not report_page:
        raise RuntimeError("レポート画面(IPSWebReport)が開きませんでした。")

    logger.info("レポート画面 URL: %s", report_page.url)
    return report_page


def wait_condition_frame(report_page, timeout_sec: float = 20.0):
    """
    レポート画面 page から、ReportCondition.aspx を含むフレームを探して返す。
    """
    deadline = time.time() + timeout_sec

    while time.time() < deadline:
        if report_page.is_closed():
            raise RuntimeError("レポートページがクローズされました。")

        frames = report_page.frames
        logger.debug("[DEBUG] 条件フレーム探索 frame数=%d", len(frames))
        found = None
        for f in frames:
            url = (f.url or "").lower()
            logger.debug("  frame url=%s", url)
            if "reportcondition.aspx" in url:
                found = f
                break

        if found:
            logger.info("条件フレーム発見: %s", found.url)
            return found

        report_page.wait_for_timeout(500)

    raise RuntimeError("ReportCondition.aspx を含む条件フレームを見つけられませんでした。")


def wait_buttons_frame(report_page, timeout_sec: float = 20.0):
    """
    レポート画面 page から、ButtonsWebForm.aspx を含むフレームを探して返す。
    （表示ボタンなどが配置されているフレーム）
    """
    deadline = time.time() + timeout_sec

    while time.time() < deadline:
        if report_page.is_closed():
            raise RuntimeError("レポートページがクローズされました。")

        frames = report_page.frames
        logger.debug("[DEBUG] ボタンフレーム探索 frame数=%d", len(frames))
        found = None
        for f in frames:
            url = (f.url or "").lower()
            logger.debug("  frame url=%s", url)
            if "buttonswebform.aspx" in url:
                found = f
                break

        if found:
            logger.info("ボタンフレーム発見: %s", found.url)
            return found

        report_page.wait_for_timeout(500)

    raise RuntimeError("ButtonsWebForm.aspx を含むボタンフレームを見つけられませんでした。")


def set_date(cond_frame, target_date: dt.date):
    """条件フレーム上の日付入力ボックスに target_date を設定する。"""
    date_str = target_date.strftime("%Y/%m/%d")
    logger.info("日付を設定します: %s", date_str)

    date_input = cond_frame.locator("input[type='text'][id*='Date']")
    if date_input.count() == 0:
        date_input = cond_frame.locator("input[type='text']")
    if date_input.count() == 0:
        raise RuntimeError("日付入力用のテキストボックスが見つかりませんでした。")

    date_input.first.fill(date_str)


def get_parking_options(cond_frame):
    """
    条件フレーム上の駐車場 select の option 一覧を取得する。
    戻り値: [(value, text), ...]
    """
    select = cond_frame.locator("select")
    if select.count() == 0:
        raise RuntimeError("駐車場選択用の select 要素が見つかりませんでした。")

    select = select.first
    options = select.locator("option").all()
    results = []
    for opt in options:
        value = opt.get_attribute("value")
        text = opt.inner_text().strip()
        if not value or not text:
            continue
        results.append((value, text))
    if not results:
        raise RuntimeError("駐車場の option が 1 件も取得できませんでした。")

    logger.info("駐車場 option 件数: %d", len(results))
    return select, results


def click_show(buttons_frame):
    """
    ボタンフレーム上の「表示」ボタンをクリックする。
    ボタンが見つからない場合は、JavaScript 関数 ClickShow() を直接呼び出す。
    """
    selectors = [
        "input[type='button'][value='表示']",
        "input[type='submit'][value='表示']",
        "input[value*='表示']",
        "button:has-text('表示')",
    ]

    for sel in selectors:
        btn = buttons_frame.locator(sel)
        if btn.count() > 0:
            logger.info("『表示』ボタンをクリックします。（selector=%s）", sel)
            btn.first.click()
            return

    btn = buttons_frame.get_by_text("表示", exact=False)
    if btn.count() > 0:
        logger.info("『表示』テキストを持つ要素をクリックします。")
        btn.first.click()
        return

    logger.info("『表示』ボタンが見つからないため、ClickShow() を直接実行してみます。")
    try:
        buttons_frame.evaluate("typeof ClickShow === 'function' ? ClickShow() : null")
        return
    except Exception as e:
        logger.error("ClickShow() の実行にも失敗: %s", e)
        raise RuntimeError("『表示』ボタンが見つかりませんでした。")


# ==========================
#  .mht URL の取得（ネットワーク監視）
# ==========================

def wait_new_mht_url(mht_urls, start_index: int, report_page, timeout_sec: float = 30.0) -> str:
    """
    mht_urls: context.on("request") で溜めている .mht URL のリスト
    start_index: ClickShow() を押す前の len(mht_urls)
    → これより後に増えた .mht URL を待つ
    """
    logger.info("ネットワークリクエストから .mht URL を待機します…")
    deadline = time.time() + timeout_sec

    while time.time() < deadline:
        if report_page.is_closed():
            raise RuntimeError("レポートページがクローズされました。")

        if len(mht_urls) > start_index:
            url = mht_urls[-1]
            logger.info("[OK] ネットワークで .mht リクエストを検出: %s", url)
            return url

        report_page.wait_for_timeout(500)

    # タイムアウト時デバッグ
    logger.error("[DEBUG] .mht リクエストが検出できなかったため、最後に context 内のページ URL を出力します。")
    ctx = report_page.context
    for p in ctx.pages:
        try:
            logger.error("  Page URL: %s (closed=%s)", p.url, p.is_closed())
            for f in p.frames:
                logger.error("    Frame URL: %s", f.url)
        except Exception:
            pass

    raise RuntimeError("ネットワークリクエストとしても .mht が飛びませんでした。")


# ==========================
#  MHT ダウンロード
# ==========================

def download_mht_via_requests(context, mht_url: str, download_dir: str,
                              park_text: str, target_date: dt.date) -> str:
    """
    Playwright コンテキストの Cookie を使って requests で MHT をダウンロードする。
    """
    cookie_header = build_cookie_header(context, mht_url)
    headers = {
        "User-Agent": IE_LIKE_UA,
        "Cookie": cookie_header,
    }

    logger.info("requests でダウンロード開始…")
    resp = requests.get(mht_url, headers=headers, timeout=60)
    resp.raise_for_status()

    date_str = target_date.strftime("%Y%m%d")
    safe_park = "".join(c if c not in r'\/:*?"<>|' else "_" for c in park_text)
    filename = f"{date_str}_{safe_park}.mht"
    full_path = os.path.join(download_dir, filename)

    with open(full_path, "wb") as f:
        f.write(resp.content)

    logger.info("[OK] MHT ダウンロード完了: %s", full_path)
    return full_path


# ==========================
#  CSV 変換
# ==========================

def _convert_cell_to_number_like(val):
    """
    Excelの「文字列を数値に変換」に近いイメージで、
    1セルずつ「数字っぽければ int/float に変換」する。
    変換できなければ元の値を返す。
    """
    if val is None:
        return val
    try:
        import math
        if isinstance(val, float) and math.isnan(val):
            return val
    except Exception:
        pass

    s = str(val)
    # 全角→半角に正規化
    s = unicodedata.normalize("NFKC", s)
    # カンマ・全角スペース除去、前後の空白除去
    s = s.replace(",", "").replace("\u3000", "").strip()

    if s == "":
        return val

    import re
    if not re.fullmatch(r"-?\d+(\.\d+)?", s):
        return val

    try:
        if "." in s:
            return float(s)
        else:
            return int(s)
    except Exception:
        return val


def _coerce_numeric_columns_cellwise(df: pd.DataFrame) -> pd.DataFrame:
    """DataFrame 全体について、全セル対象に「数字っぽければ数値変換」をかける。"""
    for col in df.columns:
        df[col] = df[col].map(_convert_cell_to_number_like)
    return df


def convert_mht_to_csv(mht_path: str, csv_dir: str) -> str:
    """
    1つの MHT ファイルから HTML シートを抽出し、
    sheet001 の <table> を pandas で読み込んで CSV に保存する。
    """
    csv_dir = ensure_dir(csv_dir)
    logger.info("MHT から CSV への変換を開始: %s", mht_path)

    with open(mht_path, "rb") as f:
        raw = f.read()

    msg = message_from_bytes(raw)

    target_df = None

    for part in msg.walk():
        ctype = part.get_content_type()
        loc = part.get("Content-Location", "") or ""
        if ctype != "text/html":
            continue
        loc_lower = loc.lower()
        if "sheet001" not in loc_lower:
            # グラフなどの sheet002 以降は無視
            continue

        charset = part.get_content_charset() or "utf-8"
        try:
            html_bytes = part.get_payload(decode=True) or b""
            html = html_bytes.decode(charset, errors="ignore")
        except Exception as e:
            logger.warning("HTML デコードに失敗しました (loc=%s): %s", loc, e)
            continue

        try:
            tables = pd.read_html(StringIO(html))
        except Exception as e:
            logger.warning("pandas.read_html に失敗しました (loc=%s): %s", loc, e)
            continue

        if not tables:
            continue

        df = tables[0]
        df = _coerce_numeric_columns_cellwise(df)
        target_df = df
        logger.info("sheet001 のテーブル取得: shape=%s", df.shape)
        break

    if target_df is None:
        # sheet001 が見つからない場合、保険として最初の sheet を使う
        logger.warning("sheet001 が見つからなかったため、最初に見つかった sheet を探します。")
        for part in msg.walk():
            ctype = part.get_content_type()
            loc = part.get("Content-Location", "") or ""
            if ctype != "text/html":
                continue
            if "sheet" not in (loc or "").lower():
                continue
            charset = part.get_content_charset() or "utf-8"
            try:
                html_bytes = part.get_payload(decode=True) or b""
                html = html_bytes.decode(charset, errors="ignore")
            except Exception:
                continue
            try:
                tables = pd.read_html(StringIO(html))
            except Exception:
                continue
            if not tables:
                continue
            df = tables[0]
            df = _coerce_numeric_columns_cellwise(df)
            target_df = df
            logger.info("フォールバックで最初の sheet テーブル取得: shape=%s", df.shape)
            break

    if target_df is None:
        raise RuntimeError("MHT 内に有効なテーブルが見つかりませんでした。")

    base_name = os.path.splitext(os.path.basename(mht_path))[0]
    csv_path = os.path.join(csv_dir, base_name + ".csv")
    target_df.to_csv(csv_path, index=False, encoding="utf-8-sig")

    logger.info("[OK] CSV 変換完了: %s", csv_path)
    return csv_path


# ==========================
# メインロジック
# ==========================

def run(date_str: str, download_dir: str, headless: bool):
    target_date = dt.datetime.strptime(date_str, "%Y-%m-%d").date()
    download_dir = ensure_dir(download_dir)
    csv_dir = os.path.join(download_dir, "csv")

    logger.info("=== 全駐車場DL(＋CSV変換) 開始 === 日付=%s", target_date)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context(
            user_agent=IE_LIKE_UA,
            ignore_https_errors=True,
        )

        mht_urls = []

        def on_request(request):
            url = request.url or ""
            lower = url.lower()
            if ".mht" in lower:
                logger.info("[HOOK] .mht リクエスト検知: %s", url)
                mht_urls.append(url)

        context.on("request", on_request)

        try:
            top_page = open_parking_top(context, headless=headless)
            report_page = click_report_menu_and_get_page(context, top_page)
            cond_frame = wait_condition_frame(report_page, timeout_sec=20.0)
            buttons_frame = wait_buttons_frame(report_page, timeout_sec=20.0)

            set_date(cond_frame, target_date)

            select, options = get_parking_options(cond_frame)

            for idx, (value, text) in enumerate(options, start=1):
                logger.info("============================================")
                logger.info(
                    "%d/%d 駐車場 '%s' (value=%s) の処理を開始します",
                    idx, len(options), text, value
                )

                select.select_option(value=value)

                start_index = len(mht_urls)

                click_show(buttons_frame)

                try:
                    mht_url = wait_new_mht_url(mht_urls, start_index, report_page, timeout_sec=30.0)
                    mht_path = download_mht_via_requests(context, mht_url, download_dir, text, target_date)
                    try:
                        convert_mht_to_csv(mht_path, csv_dir)
                    except Exception as e_csv:
                        logger.error(
                            "駐車場 '%s' の CSV 変換に失敗しました: %s",
                            text, e_csv
                        )
                except Exception as e:
                    logger.error(
                        "駐車場 '%s' の MHT ダウンロードに失敗しました: %s",
                        text, e
                    )
                    continue

                report_page.wait_for_timeout(1000)

            logger.info("=== 全駐車場DL(＋CSV変換) 完了 ===")

        finally:
            context.close()
            browser.close()


def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--date",
        required=True,
        help="対象日 (YYYY-MM-DD)",
    )
    parser.add_argument(
        "--download-dir",
        default="downloads",
        help="MHT を保存するディレクトリ（デフォルト: ./downloads）",
    )
    parser.add_argument(
        "--headless",
        action="store_true",
        help="ヘッドレスモードで起動する場合に指定",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    try:
        run(args.date, args.download_dir, args.headless)
    except Exception as e:
        logger.error("致命的エラー: %s", e)
        sys.exit(1)
