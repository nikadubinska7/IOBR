$ErrorActionPreference = "Stop"

if (-not (Test-Path ".venv")) {
    py -3 -m venv .venv
}

.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller

pyinstaller --noconfirm --clean --onefile --windowed --name "iOBR Extractor" iobr_app.py

Write-Host ""
Write-Host "Build complete: dist\iOBR Extractor.exe"
