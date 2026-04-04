# interlock_table_to_csv.py
# interlodkdown.py （新方式：画面の表→CSV 直出力 / IEモード）
# 実行:
#   python "C:\\Users\\sugitamasahiko\\Documents\\parking_system\\interlodkdown.py"

import csv
import time
from datetime import datetime
from pathlib import Path

from selenium import webdriver
from selenium.webdriver import IeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.ie.service import Service as IEService

# ======= 設定ここから =======
IE_DRIVER_PATH = r"C:\tools\IEDriverServer\IEDriverServer.exe"  # 置いた場所に合わせて
EDGE_EXE_PATH  = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"  # 通常はこれ

URLS = [
    "http://192.168.1.200/Interlock/Parking.htm",
    "http://192.168.1.200/Interlock/IPSWebReport/MainFrameset_s.aspx",
    # ※ frameset配下のレポートが自動で見えるようになっていればこの2つでOK
]

# frame/iframe探索のヒント
FRAME_HINTS = [
    "InterlockWeb", "NewSysWebReport", "Report", "MainFrame", "Content", "Viewer"
]

# table探索セレクタ
TABLE_SELECTORS = [
    (By.CSS_SELECTOR, "table#GridView1"),
    (By.CSS_SELECTOR, "table#DataGrid1"),
    (By.CSS_SELECTOR, "table.report-table"),
    (By.XPATH, "//table"),
]

OUTPUT_DIR = Path("./downloads_csv")
OUTPUT_ENCODING = "cp932"  # Excelで開きやすい。UTF-8なら "utf-8-sig"
# ======= 設定ここまで =======


def build_driver():
    opts = IeOptions()
    opts.attach_to_edge_chromium = True     # EdgeのIEモードにアタッチ
    opts.edge_executable_path = EDGE_EXE_PATH
    opts.ensure_clean_session = True
    opts.ignore_zoom_level = True
    opts.native_events = False
    service = IEService(executable_path=IE_DRIVER_PATH)
    driver = webdriver.Ie(service=service, options=opts)

    driver.set_page_load_timeout(60)
    return driver

def wait(sec=1.0):
    try:
        time.sleep(sec)
    except KeyboardInterrupt:
        pass

def go_through_urls(driver):
    for url in URLS:
        driver.get(url)
        wait(1.0)

def has_any_table(ctx):
    try:
        for by, sel in TABLE_SELECTORS:
            if ctx.find_elements(by, sel):
                return True
        if ctx.find_elements(By.TAG_NAME, "table"):
            return True
    except Exception:
        pass
    return False

def find_report_frame(driver):
    # まずヒント優先で frame / iframe を探索
    frames = driver.find_elements(By.TAG_NAME, "frame") + driver.find_elements(By.TAG_NAME, "iframe")
    for fr in frames:
        try:
            name = (fr.get_attribute("name") or "") + " " + (fr.get_attribute("id") or "")
            src  = fr.get_attribute("src") or ""
            if any(h.lower() in (name + " " + src).lower() for h in FRAME_HINTS):
                driver.switch_to.frame(fr)
                if has_any_table(driver):
                    return True
                driver.switch_to.default_content()
        except Exception:
            driver.switch_to.default_content()

    # 見つからなければ総当たり
    for fr in frames:
        try:
            driver.switch_to.frame(fr)
            if has_any_table(driver):
                return True
            driver.switch_to.default_content()
        except Exception:
            driver.switch_to.default_content()

    return False

def collect_all_tables(ctx):
    tables, used = [], set()
    for by, sel in TABLE_SELECTORS:
        try:
            for e in ctx.find_elements(by, sel):
                if e.id not in used:
                    used.add(e.id); tables.append(e)
        except Exception:
            pass
    for e in ctx.find_elements(By.TAG_NAME, "table"):
        if e.id not in used:
            used.add(e.id); tables.append(e)
    return tables

def table_to_matrix(table_el):
    rows = table_el.find_elements(By.TAG_NAME, "tr")
    matrix = []
    for r in rows:
        cells = r.find_elements(By.XPATH, "./th|./td")
        vals = []
        for c in cells:
            t = (c.text or "").replace("\r", " ").replace("\n", " ").strip()
            t = " ".join(t.split())
            vals.append(t)
        if any(vals):
            matrix.append(vals)
    if not matrix:
        return []
    max_cols = max(len(r) for r in matrix)
    matrix = [r + [""]*(max_cols - len(r)) for r in matrix]
    return matrix

def save_csv(matrix, base, idx):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = OUTPUT_DIR / f"{base}_t{idx:02d}_{ts}.csv"
    with open(out, "w", encoding=OUTPUT_ENCODING, newline="") as f:
        csv.writer(f).writerows(matrix)
    print(f"[OK] CSV保存: {out}  ({len(matrix)}行 x {len(matrix[0]) if matrix else 0}列)")

def main():
    driver = build_driver()
    try:
        go_through_urls(driver)

        in_frame = find_report_frame(driver)
        if not in_frame:
            print("[INFO] frame内に特定できず。トップでtable探索を続行します。")

        tables = collect_all_tables(driver)
        if not tables:
            raise RuntimeError("画面上に<table>が見つかりません。frameヒントやログイン手順を見直してください。")

        print(f"[INFO] 発見テーブル数: {len(tables)}")
        base = "InterlockReport"
        for i, t in enumerate(tables, start=1):
            mat = table_to_matrix(t)
            if len(mat) >= 2 and len(mat[0]) >= 2:
                save_csv(mat, base, i)
            else:
                print(f"[WARN] 小さすぎる/空のテーブルをスキップ: index={i}")

        print("[DONE] 画面の表からCSV出力を完了しました。")

    finally:
        # 確認したければ quit をコメントアウト
        driver.quit()

if __name__ == "__main__":
    main()


def save_matrix_as_csv(matrix, basename, idx):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = OUTPUT_DIR / f"{basename}_t{idx:02d}_{ts}.csv"
    with open(out, "w", encoding=OUTPUT_ENCODING, newline="") as f:
        writer = csv.writer(f)
        writer.writerows(matrix)
    print(f"[OK] CSV保存: {out}  ({len(matrix)}行 x {len(matrix[0]) if matrix else 0}列)")


def main():
    driver = build_driver()
    try:
        go_through_urls(driver)

        # frame配下にtableがあるならそこへ入る
        in_frame = find_report_frame(driver)
        if not in_frame:
            print("[INFO] frame内に特定できず。トップ文脈でtable探索を続行します。")

        tables = collect_all_tables(driver)
        if not tables:
            raise RuntimeError("画面上に <table> が見つかりません。frameヒントやログイン手順を見直してください。")

        print(f"[INFO] 発見テーブル数: {len(tables)}")

        base = "InterlockReport"
        for i, t in enumerate(tables, start=1):
            mat = table_to_matrix(t)
            if len(mat) >= 2 and len(mat[0]) >= 2:  # 最低限のサイズチェック
                save_matrix_as_csv(mat, base, i)

    finally:
        # 画面のまま確認したい場合はquitを遅らせて手動で閉じてください
        driver.quit()


if __name__ == "__main__":
    main()
