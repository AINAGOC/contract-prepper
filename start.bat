@echo off
chcp 932 >nul
cd /d "%~dp0"

echo ============================================
echo   Contract Document Tool - Starting...
echo ============================================
echo.

REM --- Check if setup has been run ---
if not exist "python\python.exe" (
    echo [ERROR] Setup has not been run yet.
    echo Please run setup_windows.bat first.
    echo.
    pause
    exit /b 1
)

echo Browser will open at http://localhost:5000
echo Do not close this window while using the tool.
echo.

REM --- Open browser ---
start http://localhost:5000

REM --- Start server using embedded Python ---
python\python.exe app.py

pause
