$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$releaseDir = Join-Path $repoRoot "release"
New-Item -ItemType Directory -Force -Path $releaseDir | Out-Null

$srcZip = Join-Path $releaseDir "answer-sheet-grader-source.zip"
$exeZip = Join-Path $releaseDir "AnswerSheetGrader-windows-x64.zip"
if (Test-Path $srcZip) { Remove-Item $srcZip -Force }
if (Test-Path $exeZip) { Remove-Item $exeZip -Force }

$srcItems = @(
  (Join-Path $repoRoot "app.py"),
  (Join-Path $repoRoot "README.md"),
  (Join-Path $repoRoot "LICENSE"),
  (Join-Path $repoRoot "requirements.txt"),
  (Join-Path $repoRoot "requirements-dev.txt"),
  (Join-Path $repoRoot "pyproject.toml"),
  (Join-Path $repoRoot ".gitignore"),
  (Join-Path $repoRoot "scripts")
)
Compress-Archive -Path $srcItems -DestinationPath $srcZip -Force

$exePath = Join-Path $repoRoot "dist\\AnswerSheetGrader.exe"
if (-not (Test-Path $exePath)) {
  throw "Executable not found: $exePath"
}
$exeTemp = Join-Path $releaseDir "exe_temp"
if (Test-Path $exeTemp) { Remove-Item $exeTemp -Recurse -Force }
New-Item -ItemType Directory -Path $exeTemp | Out-Null
Copy-Item $exePath (Join-Path $exeTemp "AnswerSheetGrader.exe") -Force
Copy-Item (Join-Path $repoRoot "README.md") (Join-Path $exeTemp "README.md") -Force
Compress-Archive -Path (Join-Path $exeTemp "*") -DestinationPath $exeZip -Force
Remove-Item $exeTemp -Recurse -Force

Write-Host "Created: $srcZip"
Write-Host "Created: $exeZip"
