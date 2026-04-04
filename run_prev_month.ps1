# run_prev_month.ps1
# 毎月1日に前月の「1日〜末日」を計算して、Pythonスクリプトをヘッドレス実行

$BaseDir = "C:\Users\sugitamasahiko\Documents\parking_system"
$VenvPy  = Join-Path $BaseDir ".venv\Scripts\python.exe"
$Script  = Join-Path $BaseDir "tn2000_download_table6.py"

# 前月の開始日・終了日を計算
$today = Get-Date
$firstOfThisMonth = Get-Date -Year $today.Year -Month $today.Month -Day 1
$lastDayPrevMonth = $firstOfThisMonth.AddDays(-1)
$firstDayPrevMonth = Get-Date -Year $lastDayPrevMonth.Year -Month $lastDayPrevMonth.Month -Day 1

$from = $firstDayPrevMonth.ToString("yyyy-MM-dd")
$to   = $lastDayPrevMonth.ToString("yyyy-MM-dd")

# 出力ログ（任意）
$LogDir = Join-Path $BaseDir "logs"
if (!(Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir | Out-Null }
$Log = Join-Path $LogDir ("run_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".log")

# 駐車場番号（必要なら変更）
$ParkNo = "000005"

# 実行（Chromiumヘッドレス：画面を出さない）
& $VenvPy $Script --from $from --to $to --park $ParkNo --headless *>> $Log
