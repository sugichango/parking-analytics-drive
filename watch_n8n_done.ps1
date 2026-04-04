# ===============================
# watch_n8n_done.ps1（安定化版：書き込み完了待ち + Moveリトライ）
# ポーリング方式
# ===============================

$watchPath   = "C:\Users\sugitamasahiko\n8n\n8n-data\trigger"
$archivePath = Join-Path $watchPath "_done_archive"
$cmdPath     = "C:\Users\sugitamasahiko\Documents\parking_system\copy_and_notify.cmd"
$cmdWork     = "C:\Users\sugitamasahiko\Documents\parking_system"

# --- 安定化パラメータ ---
$pollSeconds         = 2          # 監視頻度
$stableChecks        = 2          # 「同じサイズ」が何回続いたら書き込み完了とみなすか
$stableCheckInterval = 1          # サイズチェック間隔(秒)
$moveRetryMax        = 15         # Move-Item リトライ回数
$moveRetrySleep      = 1          # Move失敗時の待ち(秒)

New-Item -ItemType Directory -Path $archivePath -Force | Out-Null

Write-Host "WATCHING (polling): $watchPath"

function Wait-FileWriteComplete {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [int]$StableChecks = 2,
        [int]$IntervalSec = 1,
        [int]$MaxWaitSec = 30
    )

    $elapsed = 0
    $sameCount = 0
    $prevLen = -1

    while ($elapsed -lt $MaxWaitSec) {
        try {
            $item = Get-Item -LiteralPath $Path -ErrorAction Stop
            $len = $item.Length
        } catch {
            # 取得できない（作成途中/消えた）なら少し待つ
            Start-Sleep -Seconds $IntervalSec
            $elapsed += $IntervalSec
            continue
        }

        if ($len -eq $prevLen -and $len -ge 0) {
            $sameCount++
            if ($sameCount -ge $StableChecks) { return $true }
        } else {
            $sameCount = 0
            $prevLen = $len
        }

        Start-Sleep -Seconds $IntervalSec
        $elapsed += $IntervalSec
    }

    return $false
}

function Move-WithRetry {
    param(
        [Parameter(Mandatory=$true)][string]$Src,
        [Parameter(Mandatory=$true)][string]$Dst,
        [int]$RetryMax = 15,
        [int]$SleepSec = 1
    )

    for ($i = 1; $i -le $RetryMax; $i++) {
        try {
            Move-Item -LiteralPath $Src -Destination $Dst -Force -ErrorAction Stop
            return $true
        } catch {
            $msg = $_.Exception.Message
            Write-Host "WARN: Move failed ($i/$RetryMax): $msg"
            Start-Sleep -Seconds $SleepSec
        }
    }
    return $false
}

while ($true) {

    $doneFiles = Get-ChildItem -Path $watchPath -Filter "DONE_*.txt" -File -ErrorAction SilentlyContinue

    foreach ($file in $doneFiles) {

        # まず「書き込み完了」を待つ（ここが一番効く）
        $ok = Wait-FileWriteComplete -Path $file.FullName -StableChecks $stableChecks -IntervalSec $stableCheckInterval -MaxWaitSec 30
        if (-not $ok) {
            Write-Host "WARN: File not stabilized in time, skip: $($file.Name)"
            continue
        }

        $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $dest  = Join-Path $archivePath ($file.BaseName + "_" + $stamp + $file.Extension)

        $moved = Move-WithRetry -Src $file.FullName -Dst $dest -RetryMax $moveRetryMax -SleepSec $moveRetrySleep
        if ($moved) {
            Write-Host "MOVED: $($file.Name) -> $([IO.Path]::GetFileName($dest))"

            Start-Process `
                -FilePath $cmdPath `
                -WorkingDirectory $cmdWork `
                -WindowStyle Hidden

            Write-Host "OK: copy_and_notify.cmd started"
        } else {
            Write-Host "ERROR: Move failed after retries, skip: $($file.Name)"
        }
    }

    Start-Sleep -Seconds $pollSeconds
}
