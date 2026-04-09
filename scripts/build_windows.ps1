param(
  [switch]$InstallDeps
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

if ($InstallDeps) {
  python -m pip install -r requirements-dev.txt
}

python -m PyInstaller `
  --noconfirm `
  --clean `
  --windowed `
  --onefile `
  --name AnswerSheetGrader `
  --hidden-import PIL._tkinter_finder `
  app.py

Write-Host "Build finished: $repoRoot\\dist\\AnswerSheetGrader.exe"
