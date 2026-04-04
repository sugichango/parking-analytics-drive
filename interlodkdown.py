# export_hp_access_fixed.py
# pip install playwright && playwright install
import os
from datetime import datetime
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError

HEADLESS = True
OUTPUT_DIR = Path("./downloads"); OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

BASE_URL_1 = "http://192.168.1.200/Interlock/Parking.htm"
BASE_URL_2 = "http://192.168.1.200/Interlock/IPSWebReport/MainFrameset_s.aspx"
REPORT_URL_EXAMPLE = "http://192.168.1.200/Interlock/InterlockWeb/NewSysWebReport/D001_20251031105818_0419.mht"

MENU_CANDIDATES = [
    "Export to Microsoft Excel","Export to Excel","Excel",
    "Excel にエクスポート","マイクロソフト エクセルへエクスポート",
]
TABLE_SELECTORS = ["table#GridView1","table#DataGrid1","table.report-table","table"]
FRAME_URL_KEYWORDS = ["/Interlock/InterlockWeb/NewSysWebReport/","Report"]

def find_report_frame(page):
    for fr in page.frames:
        try:
            if any(k in (fr.url or "") for k in FRAME_URL_KEYWORDS):
                return fr
        except: pass
    for fr in page.frames:
        try:
            if fr.query_selector("table"):
                return fr
        except: pass
    return None

def find_first_existing(locator, selectors):
    for sel in selectors:
        el = locator.locator(sel)
        if el.count() > 0:
            return el.first
    return None

def click_export_in_context_menu(frame):
    for text in MENU_CANDIDATES:
        item = frame.locator(f"text={text}")
        if item.count() > 0:
            item.first.click(); return True
    for text in MENU_CANDIDATES:
        item = frame.locator(f"li:has-text('{text}'), a:has-text('{text}'), span:has-text('{text}')")
        if item.count() > 0:
            item.first.click(); return True
    return False

def save_download(download, prefix="HP_Access"):
    suggested = download.suggested_filename or "export.xls"
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = OUTPUT_DIR / f"{prefix}_{timestamp}__{suggested}"
    download.save_as(str(path))
    print(f"[OK] 保存しました: {path}")

def try_direct_download(page, url):
    # 重要：gotoは使わない。expect_download + window.location で遷移する。
    with page.expect_download(timeout=60000) as dl_info:
        page.evaluate(f"window.location.href = {repr(url)}")
    download = dl_info.value
    save_download(download, prefix="HP_Direct")
    return True

def try_context_menu_export(page):
    # フレームセット→レポート→表右クリック→Export をトライ
    page.goto(BASE_URL_1, timeout=60000); page.wait_for_load_state("networkidle")
    page.goto(BASE_URL_2, timeout=60000); page.wait_for_load_state("networkidle")

    frame = find_report_frame(page)
    if frame is None:
        # レポートが後から描画されるかもしれないので少し待つ
        page.wait_for_timeout(2000)
        frame = find_report_frame(page)
    if frame is None:
        raise RuntimeError("レポートのフレームを特定できませんでした。FRAME_URL_KEYWORDSを調整してください。")

    table = find_first_existing(frame, TABLE_SELECTORS)
    if table is None:
        frame.wait_for_selector("table", timeout=20000)
        table = frame.locator("table").first

    cell = table.locator("td").first
    if cell.count() == 0:
        cell = table.locator("th").first if table.locator("th").count() > 0 else table

    cell.click(button="right")

    with page.expect_download(timeout=60000) as dl_info:
        clicked = click_export_in_context_menu(frame)
        if not clicked:
            raise RuntimeError("右クリックメニューにエクスポート項目が見つかりませんでした。MENU_CANDIDATESを調整してください。")
    download = dl_info.value
    save_download(download, prefix="HP_Export")
    return True

def main():
    from pathlib import Path
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        try:
            # ① 直リンク（.mht 等）に対しては expect_download で処理
            try:
                print("[INFO] 直リンクダウンロードを試行中 …")
                if try_direct_download(page, REPORT_URL_EXAMPLE):
                    return
            except Exception as e:
                print(f"[WARN] 直リンクDLに失敗: {e}")

            # ② 失敗したら右クリック→Export でフォールバック
            print("[INFO] 右クリック → Export でフォールバック …")
            try_context_menu_export(page)

        except Exception as e:
            shot = OUTPUT_DIR / f"error_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
            try:
                page.screenshot(path=str(shot), full_page=True)
                print(f"[ERROR] {e}\nスクリーンショット: {shot}")
            except:
                print(f"[ERROR] {e}")
            raise
        finally:
            context.close(); browser.close()

if __name__ == "__main__":
    main()

# ==========================
# ▼ 業務用：MHT → CSV & Excel 変換（堅牢版）
# ==========================
from io import StringIO
from email import message_from_bytes
import pandas as pd
from pathlib import Path

def convert_mht_to_csv_excel(mht_path: Path):
    """
    MHT内を総当たりで解析：
      1) 添付の Excel / CSV パートがあればそのまま保存
      2) text/html パートから <table> を read_html で抽出（lxml/bs4, headerあり/なしの両方）
      3) デバッグ用に抽出HTMLを _raw.html として保存
    出力先: <mhtと同階層>/converted/
    """
    out_dir = mht_path.parent / "converted"
    out_dir.mkdir(exist_ok=True)
    base = mht_path.stem

    raw = mht_path.read_bytes()
    msg = message_from_bytes(raw)

    # --- (1) 添付のExcel/CSVがあれば直接保存 ---------------------------------
    found_attachment = False
    if msg.is_multipart():
        idx = 0
        for part in msg.walk():
            ctype = (part.get_content_type() or "").lower()
            if ctype in (
                "application/vnd.ms-excel",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "application/vnd.ms-excel.sheet.macroenabled.12",
                "text/csv",
            ):
                idx += 1
                payload = part.get_payload(decode=True) or b""
                # 拡張子の推定
                ext = ".xls"
                if "openxmlformats-officedocument" in ctype: ext = ".xlsx"
                if ctype == "text/csv": ext = ".csv"
                out_file = out_dir / f"{base}_attach{idx:02d}{ext}"
                out_file.write_bytes(payload)
                print(f"[OK] 添付保存: {out_file.name}  ({len(payload)} bytes)")
                found_attachment = True

    # 添付が見つかって、CSVやXLSX/XLSが保存できたなら最低限の成果は確保。
    # このあと HTML 由来のテーブル抽出も試して、取れるものはすべて取る。
    # -------------------------------------------------------------------------

    # --- (2) text/html を抽出 -------------------------------------------------
    html_bytes = None
    html_charset = None
    if msg.is_multipart():
        for part in msg.walk():
            if (part.get_content_type() or "").lower() == "text/html":
                html_bytes = part.get_payload(decode=True)
                html_charset = part.get_content_charset() or "utf-8"
                break
    else:
        if (msg.get_content_type() or "").lower() == "text/html":
            html_bytes = msg.get_payload(decode=True)
            html_charset = msg.get_content_charset() or "utf-8"

    if not html_bytes:
        if found_attachment:
            print("[INFO] text/html は無いが、添付ファイルで取得済み。")
            return
        raise RuntimeError("MHT内に text/html パートが見つかりません。")

    try:
        html_text = html_bytes.decode(html_charset, errors="replace")
    except LookupError:
        html_text = html_bytes.decode("utf-8", errors="replace")

    # デバッグ用に生HTMLを保存
    raw_html_path = out_dir / f"{base}__raw.html"
    raw_html_path.write_text(html_text, encoding="utf-8")
    print(f"[INFO] デバッグ用HTML保存: {raw_html_path.name}")

    # --- (3) HTMLから<table>抽出を複数手段で試行 ------------------------------
    dataframes: list[pd.DataFrame] = []

    def _try_read_html(text: str, flavor, header_opt):
        try:
            dfs = pd.read_html(StringIO(text), flavor=flavor, header=header_opt)
            return dfs
        except Exception:
            return []

    # 3-1) lxml + header=0
    dataframes += _try_read_html(html_text, flavor="lxml", header_opt=0)
    # 3-2) lxml + header=None（全行データ）
    if not dataframes:
        dataframes += _try_read_html(html_text, flavor="lxml", header_opt=None)
    # 3-3) bs4 + header=0
    if not dataframes:
        dataframes += _try_read_html(html_text, flavor="bs4", header_opt=0)
    # 3-4) bs4 + header=None
    if not dataframes:
        dataframes += _try_read_html(html_text, flavor="bs4", header_opt=None)

    if not dataframes:
        if found_attachment:
            print("[INFO] HTMLテーブルは検出できませんでしたが、添付ファイルを保存済みです。")
            return
        # 何も取れない場合でも、raw.html は残してあるので現物確認が容易
        raise RuntimeError("HTMLから<table>を抽出できませんでした。raw.html を確認してください。")

    # --- (4) CSVに保存（テーブルごと） ---------------------------------------
    for idx, df in enumerate(dataframes, start=1):
        df = df.copy()
        # 列名の整形（Unnamedなどの清掃はお好みで）
        df.columns = [str(c).strip() for c in df.columns]
        csv_path = out_dir / f"{base}_t{idx:02d}.csv"
        df.to_csv(csv_path, index=False, encoding="utf-8-sig")
        print(f"[OK] CSV保存: {csv_path.name}  ({len(df)}行 x {len(df.columns)}列)")

    # --- (5) Excelに保存（シート分け） ---------------------------------------
    xlsx_path = out_dir / f"{base}.xlsx"
    with pd.ExcelWriter(xlsx_path, engine="xlsxwriter") as writer:
        for idx, df in enumerate(dataframes, start=1):
            sheet = f"Table{idx:02d}"
            # Excelシート名NG文字の置換＆長さ制限
            for bad in ['\\', '/', '*', '?', ':', '[', ']']:
                sheet = sheet.replace(bad, '_')
            sheet = sheet[:31]
            df.to_excel(writer, sheet_name=sheet, index=False)
    print(f"[OK] Excel保存: {xlsx_path.name}  (シート数: {len(dataframes)})")

# ===== ダウンロード後の自動変換 =====
from pathlib import Path

try:
    # 最新のMHTファイルを取得
    latest = max(Path("downloads").glob("HP_Direct_*.mht"), key=lambda p: p.stat().st_mtime)
    print(f"[INFO] ダウンロード済みファイルを変換中: {latest.name}")
    convert_mht_to_csv_excel(latest)
except Exception as e:
    print(f"[WARN] 変換処理でエラー: {e}")
