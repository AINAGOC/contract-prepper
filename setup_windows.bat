@echo off
chcp 65001 >nul
echo ============================================
echo   契約書一括整形・チェックツール セットアップ
echo ============================================
echo.

REM --- Pythonの確認 ---
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [エラー] Pythonがインストールされていません。
    echo 以下のURLからPythonをインストールしてください：
    echo https://www.python.org/downloads/
    echo ※インストール時に「Add Python to PATH」にチェックを入れてください
    pause
    exit /b 1
)

echo Pythonが見つかりました。
python --version

REM --- LibreOfficeの確認 ---
where soffice >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo [警告] LibreOfficeが見つかりません。PDF変換に必要です。
    echo 以下のURLからインストールしてください：
    echo https://ja.libreoffice.org/download/download/
    echo.
)

REM --- 依存パッケージのインストール ---
echo.
echo 依存パッケージをインストールしています...
pip install -r requirements.txt --quiet
if %errorlevel% neq 0 (
    echo [エラー] パッケージのインストールに失敗しました。
    pause
    exit /b 1
)

echo.
echo セットアップが完了しました。
echo 「start.bat」をダブルクリックしてツールを起動してください。
pause
