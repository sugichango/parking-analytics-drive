# -*- coding: utf-8 -*-
# interlock_save_html.py
# 目的：TN2000(Interlock)の画面にアクセスし、ページ本体と全フレームの HTML を保存するだけのスクリプト
# 保存先：C:\Users\sugitamasahiko\Documents\parking_system
# 使い方（PowerShell）:
# .\.venv\Scripts\python.exe interlock_save_html.py

from __future__ import annotations
import re
import time
from datetime import datetime
from pathlib import Path
from typing import Optional
from playwright.sync_api import sync_playwright

# ====== 設定 ======
BASE_HOST = "192.168.1.200"
BASE1 = f"http://{BASE_HOST}/Interlock/Parking.htm"
BASE2 = f"http://{BASE_HOST}/Interlock/IPSWebReport/MainFrameset_s.aspx"

OUT_DIR = Path(r"C:\Users\sugitamasahiko\Documents\parking_system")  # ユーザー指定どおり
NAV_TIMEOUT_MS = 60000  # ページ読み込み待ち（ms）
# ==================

def log(level: str, msg: str):
    print(f"[{level}] {msg}")

def ensure_dir(p: Path):
    p.mkdir(parents=True, exist_ok=True)

def sanitize_filename(name: str) -> str:
    # Windowsで使えない文字を置換
    name = re.sub(r'[\\/:*?"<>|]+', "_", name)
    name = re.sub(r"\s+", "_", name).strip("_")
    return name[:120] or "page"

def save_html(content: str, prefix: str, index: Optional[int] = None) -> Path:
    ensure_dir(OUT_DIR)
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    idx = f"_{index:02d}" if index is not None else ""
    fname = f"{prefix}{idx}_{ts}.html"
    path = OUT_DIR / fname
    path.write_text(content, encoding="utf-8", errors="replace")
    return path

def main():
    ensure_dir(OUT_DIR)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        page = context.new_page()
        page.set_default_timeout(NAV_TIMEOUT_MS)

        # 1) 入口ページ
        log("INFO", "Parking.htm へアクセス")
        page.goto(BASE1, wait_until="domcontentloaded")

        # 2) フレームセット（セッション確立）
        log("INFO", "MainFrameset_s.aspx へ遷移（セッション確立）")
        page.goto(BASE2, wait_until="domcontentloaded")

        # 念のため少し待つ（内部フレームの追加ロード対策）
        time.sleep(1.0)

        # 3) 本体HTMLを保存
        log("INFO", "ページ本体HTMLを保存")
        main_html = page.content()
        main_path = save_html(main_html, prefix="TN2000_main")
        log("OK", f"保存: {main_path}")

        # 4) すべてのフレームHTMLを保存
        frames = page.frames
        log("INFO", f"フレーム数: {len(frames)}")
        for i, fr in enumerate(frames):
            try:
                # フレーム名やURLをファイル名に反映（長すぎると切り詰め）
                name_bits = []
                if fr.name():
                    name_bits.append(fr.name())
                if fr.url:
                    # URL末尾を抜粋
                    tail = fr.url.split("/")[-1]
                    name_bits.append(tail)
                base = "frame_" + "_".join(filter(None, name_bits))
                base = sanitize_filename(base) or "frame"
                html = fr.content()
                path = save_html(html, prefix=f"TN2000_{base}", index=i)
                log("OK", f"保存: {path}  (URL={fr.url})")
            except Exception as e:
                log("WARN", f"フレーム {i} の保存に失敗: {e}")

        browser.close()
        log("OK", "HTML保存完了（本体＋全フレーム）")

if __name__ == "__main__":
    main()
