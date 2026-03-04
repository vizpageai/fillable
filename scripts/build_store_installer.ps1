param(
    [string]$IsccPath = "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$spec = Join-Path $root "Fillable.spec"
$iss = Join-Path $root "packaging\installer\Fillable.iss"

if (-not (Test-Path $iss)) {
    throw "Inno script not found: $iss"
}

Write-Host "Building PyInstaller EXE..."
pyinstaller --noconfirm $spec

if (-not (Test-Path (Join-Path $root "dist\FillableDOC.exe"))) {
    throw "dist\FillableDOC.exe was not produced."
}

if (-not (Test-Path $IsccPath)) {
    throw "ISCC.exe not found at '$IsccPath'. Install Inno Setup or pass -IsccPath."
}

Write-Host "Building installer with Inno Setup..."
& $IsccPath $iss

$outDir = Join-Path $root "dist\installer"
Write-Host "Done. Installer output folder: $outDir"
