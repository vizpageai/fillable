param(
    [string]$SigntoolPath = "signtool",
    [string]$RootDir = (Split-Path -Parent $PSScriptRoot),
    [switch]$IncludeMsi,
    [switch]$IncludeSetupExe
)

$ErrorActionPreference = "Stop"

function Resolve-File([string]$pathValue) {
    if ([System.IO.Path]::IsPathRooted($pathValue)) {
        return $pathValue
    }
    return (Join-Path $RootDir $pathValue)
}

$targets = @()
$mainExe = Resolve-File "dist\FillableDOC.exe"
if (Test-Path $mainExe) { $targets += $mainExe }

$setupExe = Resolve-File "dist\installer\FillableDOC-Setup.exe"
if ($IncludeSetupExe -and (Test-Path $setupExe)) { $targets += $setupExe }

$msi = Resolve-File "dist\installer\FillableDOC.msi"
if ($IncludeMsi -and (Test-Path $msi)) { $targets += $msi }

if ($targets.Count -eq 0) {
    throw "No files found to verify."
}

foreach ($file in $targets) {
    Write-Host "Verifying: $file"
    & $SigntoolPath verify /pa /v $file
    if ($LASTEXITCODE -ne 0) {
        throw "Signature verification failed: $file"
    }
}

Write-Host "Verification complete."
