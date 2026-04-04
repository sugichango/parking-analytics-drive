# -*- coding: utf-8 -*-
"""
interlock_batch_mht_by_parking.py

Parking.htm → MainFrameset_s.aspx にアクセスし、
「駐車場名称」を「南1駐車場」から順に選択して「表示」ボタンをクリック。
表示されたレポートに対応する .mht を自動ダウンロードし、
C:\Users\sugitamasahiko\Documents\parking_system\downloads に保存します。
"""

import os
import re
import time
import shutil
from datetime import datetime
from pathlib import Path
from typing import Optional
from urllib.parse import urlparse, urljoin

import requests
from playwright.sync_api import (
    sync_playwright,
    TimeoutError as PWTimeout,
)

# ===== 設定 =====

PARKING_URL = "http://192.168.1.200/Interlock/Parking.htm"
MAINFRAME_URL = "http://192.168.1.200/Interlock/IPSWebReport/MainFrameset_s.aspx"
MHT_BASE_URL = "http://192.168.1.200/Interlock/InterlockWeb/NewSysWebReport/"

DOWNLOAD_DIR = r"C:\Users\sugitamasahiko\Documents\parking_system\downloads"

# 「駐車場名称」プルダウンのうち、除外したい項目（例：全駐車場）
EXCLUDE_PARK_NAMES = ["全駐車場"]

# 「南1駐車場」から始めたいので、ここを起点にする
START_FROM_PARK_NAME = "南1駐車場"

# frame の URL や href をチェックするときに使う正規表現
MHT_PATTERN = re.compile(r"\.mht($|\?)", re.IGNORECASE)


# ===== ユーティリティ =====

def ensure_dir(path: str) -> None:
    Path(path).mkdir(parents=True, exist_ok=True)


def sanitize_filename(name: str) -> str:
    """Windows で使えない文字を _ に置き換える"""
    return re.sub(r'[\\/:*?"<>|]+', "_", name)


def get_cookies_header_for_url(context, url: str) -> str:
    """Playwright の context から、指定 URL 用の Cookie ヘッダ文字列を作る"""
    parsed = urlparse(url)
    domain = parsed.hostname
    if not domain:
        return ""

    cookies = context.cookies()
    cookie_items = []
    for c in cookies:
        cdomain = c.get("domain") or ""
        if cdomain and (cdomain == domain or cdomain == "." + domain):
            cookie_items.append(f"{c['name']}={c['value']}")

    return "; ".join(cookie_items)


def requests_download_with_cookies(
    context,
    url: str,
    download_dir: str,
    filename_prefix: Optional[str] = None,
) -> str:
    """Playwright の Cookie を引き継いで、requests で url を保存する"""

    ensure_dir(download_dir)

    parsed = urlparse(url)
    base_name = os.path.basename(parsed.path) or f"download_{int(time.time())}.mht"
    base_name = sanitize_filename(base_name)

    if filename_prefix:
        filename = f"{sanitize_filename(filename_prefix)}_{base_name}"
    else:
        filename = base_name

    target_path = str(Path(download_dir) / filename)

    cookie_header = get_cookies_header_for_url(context, url)
    headers = {}
    if cookie_header:
        headers["Cookie"] = cookie_header

    with requests.get(url, headers=headers, timeout=30, stream=True) as r:
        r.raise_for_status()
        with open(target_path, "wb") as f:
            shutil.copyfileobj(r.raw, f)

    print(f"[OK] requests で保存: {target_path}")
    return target_path


def guess_latest_mht_url() -> str:
    """
    ページ内で mht が見つからなかったときのフォールバック。
    D001_YYYYMMDDhhmmss_8192.mht というパターンで現在時刻から推測。
    """
    now = datetime.now()
    fname = f"D001_{now.strftime('%Y%m%d%H%M%S')}_8192.mht"
    return urljoin(MHT_BASE_URL, fname)


# ===== DOM 探索系 =====

def find_condition_frame(page):
    """
    MainFrameset_s.aspx 内で、
    「対象日付」「駐車場名称」「表示」ボタンが入っているフレームを探す。
    ここでは一番最初に <select> を持つフレームを採用。
    """
    for _ in range(10):
        for fr in page.frames:
            try:
                sel = fr.query_selector("select")
            except Exception:
                sel = None
            if sel:
                return fr, sel
        page.wait_for_timeout(500)  # 0.5 秒待って再探索
    raise RuntimeError("駐車場<select> を含むフレームが見つかりませんでした。")


def find_display_button(frame):
    """
    条件フレーム内で「表示」ボタンを探す。
    input[value='表示'] / button テキスト '表示' のどちらでもOK。
    """
    # input[type=button or submit] で value='表示'
    btn = frame.query_selector("input[type='button'][value='表示'], input[type='submit'][value='表示']")
    if btn:
        return btn

    # button 要素で innerText が '表示'
    for b in frame.query_selector_all("button"):
        try:
            txt = (b.inner_text() or "").strip()
        except Exception:
            txt = ""
        if txt == "表示":
            return b

    # 「表示」というテキストの a / span をクリックする必要があるケース
    for sel in ["a", "span", "div"]:
        for el in frame.query_selector_all(sel):
            try:
                txt = (el.inner_text() or "").strip()
            except Exception:
                txt = ""
            if txt == "表示":
                return el

    raise RuntimeError("「表示」ボタンが見つかりませんでした。")


def list_parking_options(select_el):
    """
    <select> から (表示テキスト, value) のリストを取得
    """
    options = []
    for opt in select_el.query_selector_all("option"):
        text = (opt.inner_text() or "").strip()
        value = (opt.get_attribute("value") or "").strip()
        if not value:
            value = text
        options.append((text, value))
    return options


def find_mht_urls_from_frames(page):
    """
    現在表示されているフレーム群から、mht を指す URL を全部拾う。
    1) frame.url が .mht
    2) a[href] が .mht
    """
    urls = set()

    # 1) frame の URL
    for fr in page.frames:
        if fr.url and MHT_PATTERN.search(fr.url):
            urls.add(fr.url)

    # 2) a[href] の href
    for fr in page.frames:
        anchors = fr.query_selector_all("a[href]")
        for a in anchors:
            try:
                href = a.get_attribute("href") or ""
            except Exception:
                href = ""
            if href and MHT_PATTERN.search(href):
                abs_url = urljoin(fr.url, href)
                urls.add(abs_url)

    return list(urls)


# ===== メイン処理 =====

def main():
    ensure_dir(DOWNLOAD_DIR)

    with sync_playwright() as p:
        # 最初はブラウザを表示させて挙動確認できるように headless=False
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()

        page = context.new_page()

        # 1) Parking.htm へアクセス（セッション確立）
        print("[INFO] Parking.htm にアクセス中...")
        page.goto(PARKING_URL, timeout=30000)
        page.wait_for_load_state("domcontentloaded")

        # 2) レポートの MainFrameset_s.aspx へ遷移
        print("[INFO] MainFrameset_s.aspx に遷移中...")
        page.goto(MAINFRAME_URL, timeout=30000)
        page.wait_for_load_state("domcontentloaded")

        # 3) 条件入力用フレームと <select> を特定
        cond_frame, select_el = find_condition_frame(page)
        print("[INFO] 条件フレームを特定しました。")

        # 4) 「表示」ボタンを取得
        display_btn = find_display_button(cond_frame)
        print("[INFO] 「表示」ボタンを特定しました。")

        # 5) 駐車場候補のリストを取得
        options = list_parking_options(select_el)
        print("[INFO] 駐車場選択肢:")
        for t, v in options:
            print("   -", t)

        # 6) 「南1駐車場」以降 & 「全駐車場」など除外リストをスキップ
        #    → 実行対象リストを作る
        exec_targets = []
        started = START_FROM_PARK_NAME is None
        for text, value in options:
            if text in EXCLUDE_PARK_NAMES:
                continue
            if not started:
                if text == START_FROM_PARK_NAME:
                    started = True
                else:
                    continue
            exec_targets.append((text, value))

        if not exec_targets:
            raise RuntimeError("実行対象となる駐車場が見つかりませんでした。")

        print(f"[INFO] 実行対象駐車場: {len(exec_targets)} 件")

        # 7) 各駐車場ごとに「表示」→ mht ダウンロード
        total_saved = 0

        for idx, (park_name, value) in enumerate(exec_targets, start=1):
            print("\n==============================================")
            print(f"[STEP {idx}/{len(exec_targets)}] 駐車場: {park_name}")

            # 7-1) プルダウンで駐車場を選択
            try:
                select_el.select_option(value=value)
            except Exception:
                # value でダメなら label で
                select_el.select_option(label=park_name)

            # 7-2) 「表示」ボタンをクリックしてレポート表示
            try:
                with page.expect_load_state("networkidle", timeout=15000):
                    display_btn.click()
            except PWTimeout:
                # 軽い再描画で networkidle に到達しない場合もあるのでそのまま続行
                display_btn.click()
                page.wait_for_timeout(1000)

            # レポート描画待ち
            page.wait_for_timeout(1500)

            # 7-3) フレームから mht URL を探す
            mht_urls = find_mht_urls_from_frames(page)
            if mht_urls:
                print(f"[INFO] フレーム内から mht URL を {len(mht_urls)} 件検出。")
            else:
                print("[WARN] フレーム内に mht が見つかりません。フォールバックURLを使用します。")
                mht_urls = [guess_latest_mht_url()]

            # 7-4) 見つかった mht URL をすべてダウンロード
            for url in mht_urls:
                try:
                    requests_download_with_cookies(
                        context=context,
                        url=url,
                        download_dir=DOWNLOAD_DIR,
                        filename_prefix=park_name,
                    )
                    total_saved += 1
                except Exception as e:
                    print(f"[ERROR] mht ダウンロード失敗: {e}")

        print("\n==============================================")
        print(f"[DONE] 保存した mht ファイル数: {total_saved} 件")

        context.close()
        browser.close()


if __name__ == "__main__":
    main()
