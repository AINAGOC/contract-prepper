@echo off
chcp 932 >nul
echo ============================================
echo   Setup - Contract Document Tool
echo ============================================
echo.

REM --- Check Python ---
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed.
    echo.
    echo Please install Python first:
    echo   1. Go to https://www.python.org/downloads/
    echo   2. Download and run the installer
    echo   3. IMPORTANT: Check "Add Python to PATH"
    echo   4. Run this setup again
    echo.
    pause
    exit /b 1
)

echo Python found:
python --version

REM --- Check LibreOffice ---
where soffice >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo [WARNING] LibreOffice not found. Required for PDF conversion.
    echo Please install from: https://ja.libreoffice.org/download/download/
    echo.
)

REM --- Install packages ---
echo.
echo Installing Python packages...
pip install -r requirements.txt --quiet
if %errorlevel% neq 0 (
    echo [ERROR] Package installation failed.
    pause
    exit /b 1
)

echo.
echo ============================================
echo   Setup complete!
echo   Double-click start.bat to run the tool.
echo ============================================
pause
