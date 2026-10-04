@echo off
echo ==========================================
echo StudyFlow - Windows Setup
echo ==========================================
echo.

REM Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo Python not found! Please install Python 3.10 or higher.
    pause
    exit /b 1
)

REM Upgrade pip
echo [1/5] Upgrading pip...
python -m pip install --upgrade pip

REM Install wheel first
echo [2/5] Installing wheel...
pip install wheel setuptools

REM Install core packages first (ones with reliable wheels)
echo [3/5] Installing core packages...
pip install Flask==3.0.0 Werkzeug==3.0.1 PyPDF2==3.0.1 requests==2.31.0

REM Install document processing
echo [4/5] Installing document processing...
pip install python-docx==1.1.0 ebooklib==0.18 beautifulsoup4==4.12.2 lxml==5.1.0

REM Try to install Pillow with pre-built wheel only
echo [5/5] Installing Pillow (image support)...
pip install --only-binary :all: Pillow
if errorlevel 1 (
    echo.
    echo WARNING: Pillow could not be installed automatically.
    echo Image uploads will not work, but PDF/DOCX/TXT/EPUB will work fine.
    echo To fix later, visit: https://pypi.org/project/Pillow/
)

REM Install pytesseract
echo Installing OCR support...
pip install pytesseract==0.3.10

REM Create folders
echo.
echo Creating required folders...
if not exist static\uploads mkdir static\uploads
if not exist database mkdir database

echo.
echo ==========================================
echo Setup Complete!
echo ==========================================
echo.
echo To run the app, type: python app.py
echo Then open: http://localhost:5000
echo.
pause
