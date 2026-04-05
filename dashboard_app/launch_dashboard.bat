@echo off
cd /d "%~dp0"
title 駐車場データダッシュボード

echo ===================================================
echo   駐車場データダッシュボード 起動システム
echo ===================================================

if not exist ".venv\Scripts\python.exe" (
    echo [INFO] 初回起動のため、専用のPython環境を構築しています...
    echo [INFO] これには数分かかる場合があります。画面を閉じずにお待ちください。
    python -m venv .venv
    call .venv\Scripts\activate.bat
    echo [INFO] 必要なパッケージをインストールしています...
    pip install -r requirements.txt
    echo [INFO] 環境構築が完了しました！
) else (
    echo [INFO] 起動準備中...
    call .venv\Scripts\activate.bat
)

echo [INFO] ダッシュボードを起動します...
streamlit run unified_dashboard.py

pause
