@echo off
setlocal enabledelayedexpansion
chcp 932 >nul
cd /d "%~dp0"

echo ============================================
echo   Setup - Contract Document Tool
echo ============================================
echo.

REM --- Check if Python folder already exists ---
if exist "python\python.exe" (
    echo Python already installed. Skipping download.
    goto :install_packages
)

REM --- Clean up any failed previous attempts ---
if exist "python" rmdir /s /q "python" 2>nul
if exist "python_embed.zip" del "python_embed.zip" 2>nul

echo Downloading Python (Embedded version)...
echo This may take a few minutes...
echo.

REM --- Download Python Embedded ---
set PYTHON_VERSION=3.11.9
set PYTHON_URL=https://www.python.org/ftp/python/%PYTHON_VERSION%/python-%PYTHON_VERSION%-embed-amd64.zip
set PYTHON_ZIP=python_embed.zip

REM --- Try PowerShell first (more reliable on corporate networks) ---
echo Method 1: PowerShell...
powershell -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; $wc = New-Object Net.WebClient; $wc.DownloadFile('%PYTHON_URL%', '%PYTHON_ZIP%')" 2>nul

if exist "%PYTHON_ZIP%" (
    echo Download successful.
    goto :extract_python
)

REM --- Fallback to curl with SSL options ---
echo Method 2: curl...
curl -L -k --ssl-no-revoke -o "%PYTHON_ZIP%" "%PYTHON_URL%" 2>nul

if exist "%PYTHON_ZIP%" (
    echo Download successful.
    goto :extract_python
)

REM --- Fallback to certutil ---
echo Method 3: certutil...
certutil -urlcache -split -f "%PYTHON_URL%" "%PYTHON_ZIP%" 2>nul

if exist "%PYTHON_ZIP%" (
    echo Download successful.
    goto :extract_python
)

REM --- Manual download instructions ---
echo.
echo [ERROR] Automatic download failed.
echo.
echo Please download Python manually:
echo   1. Open this URL in your browser:
echo      %PYTHON_URL%
echo   2. Save the file as "python_embed.zip" in this folder:
echo      %cd%
echo   3. Run this setup again
echo.
pause
exit /b 1

:extract_python
echo.
echo Extracting Python...
mkdir python 2>nul
powershell -Command "Expand-Archive -LiteralPath '%PYTHON_ZIP%' -DestinationPath 'python' -Force"

REM --- Verify extraction ---
if not exist "python\python.exe" (
    echo [ERROR] Extraction failed. Trying alternative method...
    powershell -Command "Add-Type -A 'System.IO.Compression.FileSystem'; [IO.Compression.ZipFile]::ExtractToDirectory('%PYTHON_ZIP%', 'python')"
)

if not exist "python\python.exe" (
    echo [ERROR] Could not extract Python.
    echo Please extract python_embed.zip manually into a folder named "python"
    pause
    exit /b 1
)

del "%PYTHON_ZIP%" 2>nul
echo Extraction complete.

REM --- Enable pip in embedded Python ---
echo Configuring Python...
set PTH_FILE=python\python311._pth
if exist "%PTH_FILE%" (
    echo python311.zip> "%PTH_FILE%"
    echo .>> "%PTH_FILE%"
    echo Lib\site-packages>> "%PTH_FILE%"
)

REM --- Download and install pip ---
echo.
echo Installing pip...
set PIP_URL=https://bootstrap.pypa.io/get-pip.py

powershell -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; (New-Object Net.WebClient).DownloadFile('%PIP_URL%', 'python\get-pip.py')" 2>nul

if not exist "python\get-pip.py" (
    curl -L -k --ssl-no-revoke -o "python\get-pip.py" "%PIP_URL%" 2>nul
)

if not exist "python\get-pip.py" (
    certutil -urlcache -split -f "%PIP_URL%" "python\get-pip.py" 2>nul
)

if not exist "python\get-pip.py" (
    echo [ERROR] Failed to download pip installer.
    pause
    exit /b 1
)

python\python.exe python\get-pip.py --no-warn-script-location -q
del "python\get-pip.py" 2>nul

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
