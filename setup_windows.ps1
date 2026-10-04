# StudyFlow Windows Setup Script
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "StudyFlow - Windows Setup" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

# Check Python
try {
    $pyVersion = python --version 2>&1
    Write-Host "[OK] Python found: $pyVersion" -ForegroundColor Green
} catch {
    Write-Host "[ERROR] Python not found! Install Python 3.10+ from python.org" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Upgrade pip
Write-Host "[1/5] Upgrading pip..." -ForegroundColor Yellow
python -m pip install --upgrade pip

# Install wheel
Write-Host "[2/5] Installing build tools..." -ForegroundColor Yellow
pip install wheel setuptools

# Core packages
Write-Host "[3/5] Installing core packages..." -ForegroundColor Yellow
pip install Flask==3.0.0 Werkzeug==3.0.1 PyPDF2==3.0.1 requests==2.31.0

# Document processing
Write-Host "[4/5] Installing document processing..." -ForegroundColor Yellow
pip install python-docx==1.1.0 ebooklib==0.18 beautifulsoup4==4.12.2 lxml==5.1.0

# Pillow - try pre-built wheel only
Write-Host "[5/5] Installing Pillow (image support)..." -ForegroundColor Yellow
pip install --only-binary :all: Pillow
if ($LASTEXITCODE -ne 0) {
    Write-Host "[WARNING] Pillow install failed. Image uploads won't work." -ForegroundColor Yellow
    Write-Host "          PDF/DOCX/TXT/EPUB will work fine without it." -ForegroundColor Yellow
}

# pytesseract
Write-Host "Installing OCR support..." -ForegroundColor Yellow
pip install pytesseract==0.3.10

# Create folders
Write-Host ""
Write-Host "Creating required folders..." -ForegroundColor Yellow
New-Item -ItemType Directory -Force -Path static\uploads | Out-Null
New-Item -ItemType Directory -Force -Path database | Out-Null

Write-Host ""
Write-Host "==========================================" -ForegroundColor Green
Write-Host "Setup Complete!" -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Run the app: python app.py" -ForegroundColor Cyan
Write-Host "Open browser: http://localhost:5000" -ForegroundColor Cyan
Write-Host ""
Read-Host "Press Enter to exit"
