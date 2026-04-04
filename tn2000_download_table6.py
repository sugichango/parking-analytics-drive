# tn2000_download_table6.py
from playwright.sync_api import sync_playwright
from datetime import date, datetime, timedelta
import argparse
import sys

# ====== ログイン情報：必要に応じて設定 ======
LOGIN_URL = "https://tp.parking-s.co.jp/tn200/login.php"
USERNAME  = "demo"          # 例: "tsukuba_demo"
PASSWORD  = "demo"    # 例: "********"

# ====== ここは通常そのままでOK ======
BASE_REPORT_URL = "https://tp.parking-s.co.jp/tn200/index.php?id=&con=report_day&park_no={park_no}&paydate={paydate}&def=Y"
TABLE_INDEX = 6  # 「駐車時間別」のtableが6番だったため固定

JS_GET_CSV = r"""
(function(){
  const tables = Array.from(document.querySelectorAll('table'));
  const idx = %d;
  if (tables.length <= idx) return {ok:false, msg:`tables=${tables.length} 表%dが見つかりません`};
  const table = tables[idx];
  const csv = Array.from(table.querySelectorAll('tr')).map(tr =>
    Array.from(tr.children).map(td =>
      `"${(td.innerText||'').replace(/\r?\n+/g,' ').replace(/"/g,'""').trim()}"`
    ).join(',')
  ).join('\r\n');
  return {ok:true, csv};
})();
""" % (TABLE_INDEX, TABLE_INDEX)

def parse_args():
  ap = argparse.ArgumentParser(description="TN-2000 日報（駐車時間別=表6）CSV自動ダウンロード")
  ap.add_argument("--park", default="000005", help="駐車場番号（例: 000005）")
  ap.add_argument("--date", help="対象日（YYYY-MM-DD）")
  ap.add_argument("--from", dest="date_from", help="期間開始日（YYYY-MM-DD）")
  ap.add_argument("--to", dest="date_to", help="期間終了日（YYYY-MM-DD）")
  ap.add_argument("--headless", action="store_true", help="ヘッドレス（画面非表示）で実行")
  ap.add_argument("--edge", action="store_true", help="Edge(msedge)で起動（デフォルトはChromium）")
  return ap.parse_args()

def daterange(d1: date, d2: date):
  """d1〜d2（両端含む）の日付を順に返す"""
  cur = d1
  while cur <= d2:
    yield cur
    cur = cur + timedelta(days=1)

def ensure_date(s: str) -> date:
  return datetime.strptime(s, "%Y-%m-%d").date()

def build_report_url(park_no: str, d: date) -> str:
  return BASE_REPORT_URL.format(park_no=park_no, paydate=d.strftime("%Y-%m-%d"))

def login(page):
  page.goto(LOGIN_URL)
  # 入力欄 name は実ページ調査済み
  page.fill('input[name="loginname"]', USERNAME)
  page.fill('input[name="loginpass"]',  PASSWORD)
  page.click('input[type="submit"], button[type="submit"]')
  page.wait_for_load_state("networkidle")

def fetch_one_day(page, park_no: str, d: date) -> str:
  url = build_report_url(park_no, d)
  page.goto(url)
  page.wait_for_load_state("networkidle")
  result = page.evaluate(JS_GET_CSV)
  if not result.get("ok"):
    raise RuntimeError(f"[{d}] CSV取得失敗: {result.get('msg')}")
  return result["csv"]

def main():
  args = parse_args()

  # 日付の決定
  targets = []
  if args.date:
    targets = [ensure_date(args.date)]
  elif args.date_from and args.date_to:
    d1 = ensure_date(args.date_from)
    d2 = ensure_date(args.date_to)
    if d2 < d1:
      print("ERROR: --from は --to 以前の日付にしてください。", file=sys.stderr)
      sys.exit(1)
    targets = list(daterange(d1, d2))
  else:
    # 対話で日付取得（未入力なら今日）
    s = input("対象日を入力してください（YYYY-MM-DD、空Enterで今日）: ").strip()
    targets = [ensure_date(s) if s else date.today()]

  # ブラウザ起動
  with sync_playwright() as p:
    launch_kwargs = {"headless": args.headless}
    if args.edge:
      # Edgeで起動
      launch_kwargs["channel"] = "msedge"
    browser = p.chromium.launch(**launch_kwargs)
    context = browser.new_context()
    page = context.new_page()

    # ログイン（1回だけ）
    print("ログインページを開きます...")
    login(page)
    print("ログイン完了。")

    # 各日付を処理
    for d in targets:
      print(f"処理中: {d}（park={args.park}）")
      try:
        csv_text = fetch_one_day(page, args.park, d)
      except Exception as e:
        print(f"  × 失敗: {e}")
        continue

      out_name = f"駐車時間別_{args.park}_{d.strftime('%Y-%m-%d')}.csv"
      with open(out_name, "w", encoding="utf-8-sig", newline="") as f:
        f.write(csv_text)
      print(f"  保存完了: {out_name}")

    browser.close()
    print("完了。")

if __name__ == "__main__":
  main()