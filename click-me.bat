@echo off
echo Certificate Generator - Windows Setup
echo ===================================

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python is not installed! Please install Python 3.x from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation.
    pause
    exit /b 1
)

REM Check if pip is installed
pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo pip is not installed! Please install pip.
    pause
    exit /b 1
)

REM Check if LibreOffice is installed
soffice --version >nul 2>&1
if %errorlevel% neq 0 (
    echo LibreOffice is not installed!
    echo Please install LibreOffice from https://www.libreoffice.org/download/download/
    echo After installation, add LibreOffice to your system PATH:
    echo 1. Search for "Environment Variables" in Windows
    echo 2. Edit PATH variable
    echo 3. Add "C:\Program Files\LibreOffice\program"
    pause
    start https://www.libreoffice.org/download/download/
    exit /b 1
)

echo Installing required Python packages...
pip install -r requirements.txt

echo Starting Certificate Generator...
start "" http://localhost:5000
python main.py
