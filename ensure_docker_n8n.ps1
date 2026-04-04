# ensure_docker_n8n.ps1
# Docker Desktop / n8n / python-runner が止まっていたら起こす前処理

$ErrorActionPreference = "Stop"

$logDir  = "C:\Users\sugitamasahiko\Documents\parking_system\logs"
$logFile = Join-Path $logDir ("ensure_docker_n8n_" + (Get-Date -Format "yyyyMMdd") + ".log")
New-Item -ItemType Directory -Force -Path $logDir | Out-Null

function Log($msg) {
  $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $msg
  $line | Tee-Object -FilePath $logFile -Append
}

function Try-DockerInfo {
  try {
    docker info *> $null
    return $true
  } catch {
    return $false
  }
}

function Start-DockerDesktop {
  # よくあるインストールパス候補
  $candidates = @(
    "$env:ProgramFiles\Docker\Docker\Docker Desktop.exe",
    "${env:ProgramFiles(x86)}\Docker\Docker\Docker Desktop.exe"
  )
  foreach ($p in $candidates) {
    if (Test-Path $p) {
      Log "Starting Docker Desktop: $p"
      Start-Process -FilePath $p | Out-Null
      return $true
    }
  }
  Log "Docker Desktop exe not found in typical paths."
  return $false
}

Log "=== ensure_docker_n8n START ==="

# 1) Docker Engine が起きてるか（最大 60*5=300秒待つ）
if (-not (Try-DockerInfo)) {
  Log "Docker Engine not ready. Trying to start Docker Desktop..."
  Start-DockerDesktop | Out-Null
}

$maxTry = 60
for ($i=1; $i -le $maxTry; $i++) {
  if (Try-DockerInfo) {
    Log "Docker Engine is ready. (try=$i)"
    break
  }
  Start-Sleep -Seconds 5
  if ($i -eq $maxTry) {
    Log "Docker Engine did not become ready within timeout."
    throw "Docker Engine not ready"
  }
}

# 2) n8n スタック起動（compose）
$n8nDir = "C:\Users\sugitamasahiko\n8n"
if (-not (Test-Path $n8nDir)) {
  throw "n8n directory not found: $n8nDir"
}

Push-Location $n8nDir
try {
  Log "Running: docker compose up -d"
  docker compose up -d | ForEach-Object { Log $_ }

  Log "Running: docker compose ps"
  docker compose ps | ForEach-Object { Log $_ }
}
finally {
  Pop-Location
}

# 3) python-runner の最終確認（最大 30*2=60秒）
$maxTry2 = 30
for ($j=1; $j -le $maxTry2; $j++) {
  try {
    $out = docker exec n8n-python-runner-1 python -c "print('PY_OK')"
    if ($out -match "PY_OK") {
      Log "python-runner check OK: $out"
      break
    }
  } catch {
    # ignore and retry
  }
  Start-Sleep -Seconds 2
  if ($j -eq $maxTry2) {
    Log "python-runner did not respond."
    throw "python-runner not ready"
  }
}

Log "=== ensure_docker_n8n END (OK) ==="
exit 0
