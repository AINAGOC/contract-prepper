@echo off
setlocal enabledelayedexpansion
chcp 932 >nul
cd /d "%~dp0"

echo ============================================
echo   Setup - Contract Document Tool
echo ============================================
echo.

REM --- Check if Python folder already exists ---
if exist "python\" (
    echo Python already installed. Skipping download.
    goto :install_packages
)

echo Downloading Python (Embedded version)...
echo This may take a few minutes...
echo.

REM --- Download Python Embedded ---
set PYTHON_VERSION=3.11.9
set PYTHON_URL=https://www.python.org/ftp/python/%PYTHON_VERSION%/python-%PYTHON_VERSION%-embed-amd64.zip
set PYTHON_ZIP=python_embed.zip

curl -L -o %PYTHON_ZIP% %PYTHON_URL%
if %errorlevel% neq 0 (
    echo [ERROR] Failed to download Python.
    echo Please check your internet connection.
    pause
    exit /b 1
)

echo Extracting Python...
powershell -Command "Expand-Archive -Path '%PYTHON_ZIP%' -DestinationPath 'python' -Force"
del %PYTHON_ZIP%

REM --- Enable pip in embedded Python ---
echo Configuring Python...
set PTH_FILE=python\python311._pth
if exist "%PTH_FILE%" (
    echo python311.zip> "%PTH_FILE%"
    echo .>> "%PTH_FILE%"
    echo Lib\site-packages>> "%PTH_FILE%"
)

REM --- Download and install pip ---
echo Installing pip...
curl -L -o python\get-pip.py https://bootstrap.pypa.io/get-pip.py
python\python.exe python\get-pip.py --no-warn-script-location
del python\get-pip.py

:install_packages
echo.
echo Installing required packages...
python\python.exe -m pip install -r requirements.txt --no-warn-script-location -q
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
echo.

REM --- Check LibreOffice ---
where soffice >nul 2>&1
if %errorlevel% neq 0 (
    echo [NOTE] LibreOffice is required for PDF conversion.
    echo Please install from: https://ja.libreoffice.org/download/download/
    echo.
)

pause
