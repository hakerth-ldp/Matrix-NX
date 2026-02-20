$ErrorActionPreference = 'Stop'

Write-Host '== Matrix-NX auf neuesten main-Stand bringen ==' -ForegroundColor Cyan

git fetch --all --prune

git checkout main

git reset --hard origin/main

git pull origin main

$commit = git rev-parse --short HEAD
Write-Host "Aktueller Commit: $commit" -ForegroundColor Green

Write-Host 'Starte scpi_terminal_gui.py ...' -ForegroundColor Cyan
python scpi_terminal_gui.py
