$ErrorActionPreference = "Stop"

python -m pip install -r .\requirements.txt

python -m PyInstaller `
  --noconfirm `
  --onedir `
  --windowed `
  --name GmarketCrawler `
  --hidden-import gmarket `
  --hidden-import truststore `
  --hidden-import certifi `
  --collect-all undetected_chromedriver `
  .\gmarket_app.py

Write-Host ""
Write-Host "Build complete: .\dist\GmarketCrawler\GmarketCrawler.exe"
