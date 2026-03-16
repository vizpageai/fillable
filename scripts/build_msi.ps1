param(
    [string]$Python = "python",
    [string]$WixPath = "wix"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$wxs = Join-Path $root "packaging\msi\Fillable.wxs"
$outDir = Join-Path $root "dist\installer"
$msiPath = Join-Path $outDir "FillableDOC.msi"

if (-not (Test-Path $wxs)) {
    throw "WiX source not found: $wxs"
}

Write-Host "Building dist\FillableDOC.exe..."
& $Python -m PyInstaller --noconfirm "$root\Fillable.spec"

if (-not (Test-Path (Join-Path $root "dist\FillableDOC.exe"))) {
    throw "dist\FillableDOC.exe was not produced."
}

New-Item -ItemType Directory -Force -Path $outDir | Out-Null

Write-Host "Building MSI with WiX..."
& $WixPath build $wxs -o $msiPath
if ($LASTEXITCODE -ne 0) {
    throw "WiX build failed with exit code $LASTEXITCODE."
}

Write-Host "MSI created: $msiPath"
