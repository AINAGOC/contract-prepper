@echo off
chcp 932 >nul
echo ============================================
echo   Keiyakusho Tool - Starting...
echo   (Contract Document Preparation Tool)
echo ============================================
echo.
echo Browser will open at http://localhost:5000
echo Do not close this window while using the tool.
echo.

REM --- Open browser ---
start http://localhost:5000

REM --- Start server ---
python app.py

if errorlevel 1 (
    echo.
    echo [ERROR] Python is not installed or not in PATH.
    echo Please install Python from: https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation.
    echo.
    pause
)
