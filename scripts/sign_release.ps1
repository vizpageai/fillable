param(
    [Parameter(Mandatory = $true)]
    [string]$CertPath,

    [string]$CertPassword = "",
    [string]$TimestampUrl = "http://timestamp.digicert.com",
    [string]$SigntoolPath = "signtool",
    [string]$RootDir = (Split-Path -Parent $PSScriptRoot),
    [switch]$IncludeMsi,
    [switch]$IncludeSetupExe,
    [switch]$PromptForPassword
)

$ErrorActionPreference = "Stop"

function Resolve-File([string]$pathValue) {
    if ([System.IO.Path]::IsPathRooted($pathValue)) {
        return $pathValue
    }
    return (Join-Path $RootDir $pathValue)
}

if (-not (Test-Path $CertPath)) {
    throw "Certificate file not found: $CertPath"
}

if ($PromptForPassword -and [string]::IsNullOrWhiteSpace($CertPassword)) {
    $secure = Read-Host "Enter PFX password" -AsSecureString
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    try {
        $CertPassword = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    }
    finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }
}

$targets = @()
$mainExe = Resolve-File "dist\FillableDOC.exe"
if (Test-Path $mainExe) {
    $targets += $mainExe
}
else {
    Write-Warning "Missing file: $mainExe"
}

$setupExe = Resolve-File "dist\installer\FillableDOC-Setup.exe"
$msi = Resolve-File "dist\installer\FillableDOC.msi"

if ($IncludeSetupExe) {
    if (Test-Path $setupExe) { $targets += $setupExe } else { Write-Warning "Missing file: $setupExe" }
}

if ($IncludeMsi) {
    if (Test-Path $msi) { $targets += $msi } else { Write-Warning "Missing file: $msi" }
}

if ($targets.Count -eq 0) {
    throw "No files found to sign. Build artifacts first."
}

foreach ($file in $targets) {
    Write-Host "Signing: $file"
    $args = @(
        "sign",
        "/fd", "SHA256",
        "/td", "SHA256",
        "/tr", $TimestampUrl,
        "/f", $CertPath
    )

    if (-not [string]::IsNullOrWhiteSpace($CertPassword)) {
        $args += @("/p", $CertPassword)
    }

    $args += $file
    & $SigntoolPath @args
    if ($LASTEXITCODE -ne 0) {
        throw "signtool failed for: $file"
    }
}

Write-Host "Signing complete."
