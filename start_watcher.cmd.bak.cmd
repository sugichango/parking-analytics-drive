@echo off
setlocal

REM ===== 設定（必要ならここだけ変える） =====
set PS1=C:\Users\sugitamasahiko\Documents\parking_system\watch_n8n_done.ps1
set LOGDIR=C:\Users\sugitamasahiko\Documents\parking_system\logs

if not exist "%LOGDIR%" mkdir "%LOGDIR%"

REM ===== 実行（ログに出す） =====
powershell -NoProfile -ExecutionPolicy Bypass -File "%PS1%" ^
  *>> "%LOGDIR%\watch_n8n_done_%DATE:~0,4%%DATE:~5,2%%DATE:~8,2%.log"

endlocal

