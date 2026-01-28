@echo off
chcp 65001 >nul
echo ============================================
echo   契約書一括整形・チェックツール 起動中...
echo ============================================
echo.
echo ブラウザで http://localhost:5000 が開きます。
echo このウィンドウは閉じないでください。
echo 終了するにはこのウィンドウを閉じてください。
echo.

REM --- ブラウザを自動で開く ---
start http://localhost:5000

REM --- サーバー起動 ---
python app.py
