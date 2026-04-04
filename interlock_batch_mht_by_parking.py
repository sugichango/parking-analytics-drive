# -*- coding: utf-8 -*-
"""
interlock_batch_mht_by_parking.py
TN2000「時間帯別入出庫日報」を自動操作して MHT → HTML → Excel 保存

・Playwright は「表示ボタンを押して MHT の URL を知る」役
・MHT の URL は
    - ネットワークレスポンス（context.on("response")）
    - window.open() フック
  のどちらかで取得
・実際の MHT ダウンロードは requests
・MHT から HTML 部分を抜き出して pandas.read_html でテーブル取得
"""

import os
import sys
import time
import argparse
import requests
import pandas as pd
from datetime import datetime
from email import message_from_bytes
from playwright.sync_api import sync_playwright


# --------------------------------------------------
# ログ
# --------------------------------------------------
def setup_logger():
    import logging
    logger = logging.getLogger("interlock_batch")
    logger.handlers.clear()
    handler = logging.StreamHandler(sys.stdout)
    handler.setFormatter(logging.Formatter(
        "[%(levelname)s] %(asctime)s - %(message)s", "%H:%M:%S"
    ))
    logger.addHandler(handler)
    logger.setLevel(logging.INFO)
    return logger


# --------------------------------------------------
# MHT → HTML 抽出
#   ※ Excel Web アーカイブ対策：
#      1) Content-Location に "sheet" を含む HTML パートで、
#         かつ <table> を含むものを最優先
#      2) ダメなら、全 HTML パートの中から <table> を含むもの
#      3) それもダメなら、先頭の HTML パート
# --------------------------------------------------
def extract_html_from_mht(mht_bytes, logger):
    msg = message_from_bytes(mht_bytes)

    def decode_part(part):
        payload = part.get_payload(decode=True) or b""
        # Excel の MHT は shift_jis のことが多いのでデフォルトを shift_jis
        charset = part.get_content_charset() or "shift_jis"
        try:
            return payload.decode(charset, errors="replace")
        except LookupError:
            return payload.decode("utf-8", errors="replace")

    html_parts = []   # (location, text)
    sheet_parts = []  # (location, text)  Content-Location に sheet を含むもの

    for part in msg.walk():
        if part.get_content_type() == "text/html":
            text = decode_part(part)
            loc = part.get("Content-Location") or ""
            rec = (loc, text)
            html_parts.append(rec)
            if "sheet" in loc.lower():
                sheet_parts.append(rec)

    if not html_parts:
        raise RuntimeError("MHT 内に HTML パートが見つかりませんでした。")

    # 1) sheet*.htm 系で <table> を含むものを探す
    for loc, text in sheet_parts:
        if "<table" in text.lower():
            logger.info(f"MHT: sheet 系 HTML パートを使用 ({loc})")
            return text

    # 2) 全 HTML パートの中から <table> を含むもの
    for loc, text in html_parts:
        if "<table" in text.lower():
            logger.info(f"MHT: <table> を含む HTML パートを使用 ({loc})")
            return text

    # 3) 最後の保険：先頭の HTML
    loc, text = html_parts[0]
    logger.info(f"MHT: <table> を含むパートが無かったため先頭の HTML を使用 ({loc})")
    return text


# --------------------------------------------------
# メイン
# --------------------------------------------------
def main():
    print("=== デバッグ出力開始 ===", file=sys.stderr)

    parser = argparse.ArgumentParser()
    parser.add_argument("--headless", action="store_true")
    args = parser.parse_args()

    logger = setup_logger()
    logger.info(f"スクリプト開始（headless={args.headless}）")

    with sync_playwright() as p:
        # ---------- ブラウザ起動 ----------
        user_agent_ie = "Mozilla/5.0 (Windows NT 10.0; Trident/7.0; rv:11.0) like Gecko"
        browser = p.chromium.launch(headless=args.headless)
        context = browser.new_context(
            viewport={"width": 1920, "height": 1080},
            user_agent=user_agent_ie,
            ignore_https_errors=True,
        )
        page = context.new_page()
        start_time = datetime.now()
        logger.info("Playwright を起動（IE偽装 UA）")

        # ---------- .mht URL を受け取る変数 ----------
        mht_url_holder = {"url": None}

        # 1) ネットワークレスポンスから検出
        def on_response(response):
            url = response.url
            if ".mht" in url.lower():
                logger.info(f"[NET] .mht レスポンス検出: {url} (status={response.status})")
                mht_url_holder["url"] = url

        context.on("response", on_response)

        # 2) window.open をフック（補助）
        page.add_init_script("""
        (function(){
            window._lastOpenedURL = null;
            const _oldOpen = window.open;
            window.open = function(u){
                try { window._lastOpenedURL = u; } catch(e){}
                if (_oldOpen) { return _oldOpen.apply(this, arguments); }
                return null;
            };
        })();
        """)

        # ---------- ① 最初の画面へ ----------
        url_home = "http://192.168.1.200/Interlock/Parking.htm"
        page.goto(url_home, timeout=30000)
        logger.info(f"開始URL: {url_home}")

        # ---------- ② メニュークリック ----------
        logger.info("『時間帯別入出庫日報』リンクをクリックします。")
        page.get_by_text("時間帯別入出庫日報", exact=True).click()
        time.sleep(2)

        # ---------- ③ frameset に移動 ----------
        frameset_url = "http://192.168.1.200/Interlock/IPSWebReport/MainFrameset_s.aspx"
        try:
            page.goto(frameset_url, timeout=10000)
            logger.info(f"MainFrameset に遷移: {frameset_url}")
        except Exception as e:
            logger.warning(f"MainFrameset 強制遷移失敗: {e}（既存ページを使用）")

        # frames の読み込み待ち
        time.sleep(2)

        # ---------- ④ ButtonsWebForm フレームを探して ClickShow() ----------
        btn_frame = None
        for _ in range(40):
            for f in page.frames:
                if "ButtonsWebForm" in f.url:
                    btn_frame = f
                    break
            if btn_frame:
                break
            time.sleep(1)

        if not btn_frame:
            raise RuntimeError("ButtonsWebForm フレームが見つかりません。")

        logger.info("ClickShow() を実行します。")
        btn_frame.evaluate("ClickShow()")

        # ---------- ⑤ .mht URL を待つ（レスポンス or window.open） ----------
        logger.info(".mht の URL を最大30秒待ちます…")
        mht_url = None
        for _ in range(30):
            # ① ネットワークレスポンスから
            if mht_url_holder["url"]:
                mht_url = mht_url_holder["url"]
                break

            # ② window.open フックから
            try:
                opened = page.evaluate("window._lastOpenedURL")
                if opened and ".mht" in opened:
                    logger.info(f"[JS] window.open から MHT URL 検出: {opened}")
                    mht_url = opened
                    break
            except Exception:
                pass

            time.sleep(1)

        if not mht_url:
            raise RuntimeError(".mht の URL を取得できませんでした。")

        logger.info(f"MHT URL 使用: {mht_url}")

        # ---------- ⑥ requests で MHT ダウンロード ----------
        logger.info("requests で MHT をダウンロードします。")
        headers = {
            "User-Agent": user_agent_ie,
            "Referer": frameset_url,
        }
        r = requests.get(mht_url, headers=headers, timeout=60)
        if r.status_code != 200:
            raise RuntimeError(f".mht ダウンロード失敗: HTTP {r.status_code}")
        mht_bytes = r.content
        logger.info(f"MHT ダウンロード成功: {len(mht_bytes)} bytes")

        # ---------- ⑦ 保存ディレクトリ ----------
        save_dir = os.path.join(os.getcwd(), "downloads")
        os.makedirs(save_dir, exist_ok=True)

        mht_name = os.path.basename(mht_url)
        if not mht_name.lower().endswith(".mht"):
            mht_name += ".mht"
        mht_path = os.path.join(save_dir, mht_name)
        with open(mht_path, "wb") as f:
            f.write(mht_bytes)
        logger.info(f"[OK] .mht 保存: {mht_path}")

        # ---------- ⑧ MHT → HTML → Excel ----------
        html_text = extract_html_from_mht(mht_bytes, logger)
        html_path = mht_path.replace(".mht", "_extracted.html")
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(html_text)
        logger.info(f"[OK] 抽出HTML 保存: {html_path}")

        try:
            # ★ 将来の仕様変更への対応（FutureWarning 対策）：
            #   StringIO 経由で渡す書き方にしておく
            from io import StringIO
            dfs = pd.read_html(StringIO(html_text))
            excel_path = mht_path.replace(".mht", "_table.xlsx")
            with pd.ExcelWriter(excel_path) as writer:
                for i, df in enumerate(dfs):
                    df.to_excel(writer, sheet_name=f"Sheet{i+1}", index=False)
            logger.info(f"[OK] Excel 保存: {excel_path}")
        except ValueError as e:
            logger.warning(f"HTML 内に <table> が見つかりません: {e}")

        logger.info(f"[DONE] 全処理完了 ({datetime.now() - start_time})")


if __name__ == "__main__":
    main()
