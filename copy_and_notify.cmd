@echo off
setlocal EnableExtensions EnableDelayedExpansion

set "BASE_DIR=%~dp0"

set "LOG_DIR=%BASE_DIR%logs"
if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"
set "ROBO_LOG=%LOG_DIR%\robocopy.log"

REM robocopy本体の明細ログ（新しいディレクトリ等）を出すファイル
set "ROBO_DETAIL=%LOG_DIR%\robocopy_detail.log"

set "SRC=%BASE_DIR%downloads"
set "DEST=\\10.0.1.113\共通\interlock_data"

echo ---------------------------------------------- >> "%ROBO_LOG%"
echo [%date% %time%] START robocopy >> "%ROBO_LOG%"
echo SRC  = %SRC% >> "%ROBO_LOG%"
echo DEST = %DEST% >> "%ROBO_LOG%"

REM 明細ログは /LOG: で毎回上書き（古い月を拾う事故を防ぐ）
robocopy "%SRC%" "%DEST%" /TEE /S /E /DCOPY:DA /COPY:DAT /NP /R:3 /W:10 /LOG:"%ROBO_DETAIL%"

set "RC=%ERRORLEVEL%"

echo [%date% %time%] robocopy exit code = %RC% >> "%ROBO_LOG%"

echo [%date% %time%] START send_smtp_mail.ps1 >> "%ROBO_LOG%"

powershell -ExecutionPolicy Bypass -File "%BASE_DIR%send_smtp_mail.ps1" ^
  -To "t-tsukasa@tutc.or.jp,m.nomura@tutc.or.jp,ayako_h@tutc.or.jp,ogmn0320@tutc.or.jp,nitta@tutc.or.jp,t1212@tutc.or.jp,m-sugita@tutc.or.jp,osugichango@gmail.com" ^
  -Src "%SRC%" ^
  -Dst "%DEST%" ^
  -RobocopyLog "%ROBO_DETAIL%" ^
  -ResultCode %RC% >> "%ROBO_LOG%" 2>>&1

echo [%date% %time%] END send_smtp_mail.ps1 (exit=%ERRORLEVEL%) >> "%ROBO_LOG%"

echo ---------------------------------------------- >> "%ROBO_LOG%"

endlocal
exit /b 0
