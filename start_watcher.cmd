@echo off
setlocal

REM ==== ログフォルダ作成 ====
set BASE_DIR=%~dp0
set LOG_DIR=%BASE_DIR%logs
if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"

REM ==== タイムスタンプ（YYYYMMDD_HHMMSS）作成 ====
for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set TS=%%i
set LOG=%LOG_DIR%\watcher_%TS%.log

echo [%date% %time%] start_watcher.cmd START > "%LOG%"
echo BASE_DIR=%BASE_DIR% >> "%LOG%"

REM ==== PowerShell watcher 実行（stdout/stderr をログへ）====
powershell -ExecutionPolicy Bypass -NoProfile -File "%BASE_DIR%watch_n8n_done.ps1" >> "%LOG%" 2>&1

echo [%date% %time%] watch_n8n_done.ps1 EXIT >> "%LOG%"
endlocal
