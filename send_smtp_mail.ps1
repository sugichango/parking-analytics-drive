param(
  [string]$To,
  [string]$Src,
  [string]$Dst,
  [string]$RobocopyLog,
  [int]$ResultCode,
  [string]$From = "m-sugita@tutc.or.jp",
  [string]$CredPath = "C:\Users\sugitamasahiko\Documents\parking_system\smtp_cred.xml",
  [string]$SmtpServer = "sheep-ivory-47175d5f7d3718af.znlc.jp",
  [int]$SmtpPort = 587
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$cred = Import-Clixml -LiteralPath $CredPath

# --- Robocopy 結果コードの分類 ---
# 0-3: 成功（コピーなし含む/何らかの変更）
# 4-7: 注意（不一致などがあり得る）
# 8+: 失敗（要対応）
$status =
  if ($ResultCode -ge 8) { "FAIL" }
  elseif ($ResultCode -ge 4) { "WARN" }
  else { "OK" }

$now = Get-Date -Format "yyyy-MM-dd HH:mm"
$Subject = "[$status] Interlock作成物コピー RC=$ResultCode $now"

$Body = @"
標記について作成できましたので報告します。
この作業は、生成AI（n8n）、タスクスケジューラーにより行っています。

作成ファイルの保存先
\\10.0.1.113\共通\interlock_data\csv
当該フォルダ下にある各フォルダに入っているファイルの内容
①
\\10.0.1.113\共通\interlock_data\csv\csvYYYYMM
→Interlock時間帯別入出庫日報をダウンロードしたCSVファイル
②
\\10.0.1.113\共通\interlock_data\csv\csv_multi_YYYY04_YYYYMM
→時間帯別入出庫日報のうち最大在庫台数の時間のデータを年度当初から集約したCSVファイル（定期枠算定に利用可能）
③
\\10.0.1.113\共通\interlock_data\csv\excelYYYYMM_with_avg
→時間帯別入出庫日報の1ヶ月分を1日1シートにしてまとめ、1ヶ月・平日・休日の別に平均値をとり、グラフ化したExcelファイル
④
\\10.0.1.113\共通\interlock_data\csv\graphs_multi_YYYY04_YYYYMM
→②のデータから、年度当初からの毎日の最大在庫時間の定期在庫・一般在庫の積み上げグラフを作成したExcelファイル

新しく作成されたファイルの保存フォルダは以下の通り。
南４Bはデータがありませんが、interlock上に残っているのでやむを得ず一括ダウンロードされています。
"@

$folders = New-Object 'System.Collections.Generic.HashSet[string]'

if (Test-Path -LiteralPath $RobocopyLog) {
  $lines = Get-Content -LiteralPath $RobocopyLog
  foreach ($line in $lines) {
    if ($line -match '^\s*新しいディレクトリ\s+\d+\s+(.+\\)\s*$') {
      $p = $matches[1]
      if (![string]::IsNullOrWhiteSpace($Src) -and ![string]::IsNullOrWhiteSpace($Dst)) {
        $p = $p.Replace($Src.TrimEnd('\') + '\', $Dst.TrimEnd('\') + '\')
        $p = $p.Replace($Src.TrimEnd('\'), $Dst.TrimEnd('\'))
      }
      $folders.Add($p) | Out-Null
    }
  }
}

if ($folders.Count -eq 0) {
  $Body += "`r`n（新しいディレクトリ行が無かったため、フォルダ抽出は0件でした）"
} else {
  $Body += "`r`n" + (($folders | Sort-Object) -join "`r`n")
}

$msg = New-Object System.Net.Mail.MailMessage
$msg.From = $From
# --- To: 複数宛先対応（カンマ/セミコロン区切り） ---
$toList = $To -split "[,;]" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
foreach ($addr in $toList) {
  $msg.To.Add($addr)
}
$msg.SubjectEncoding = [System.Text.Encoding]::UTF8
$msg.BodyEncoding    = [System.Text.Encoding]::UTF8
$msg.Subject = $Subject
$msg.Body = $Body
$msg.IsBodyHtml = $false

$client = New-Object System.Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
$client.EnableSsl = $true
$client.Credentials = $cred.GetNetworkCredential()
$client.Send($msg)

Write-Host "OK: SMTP sent"
