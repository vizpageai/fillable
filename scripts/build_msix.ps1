param(
    [string]$PackageName = "VizpageAI.FillableDOC",
    [string]$Publisher = "CN=VizpageAI",
    [string]$DisplayName = "FillableDOC",
    [string]$PublisherDisplayName = "VizpageAI",
    [string]$Version = "",
    [string]$RootDir = (Split-Path -Parent $PSScriptRoot),
    [string]$MakeAppxPath = "",
    [string]$SignToolPath = "signtool",
    [string]$CertPath = "",
    [string]$CertPassword = "",
    [switch]$SkipBuildExe
)

$ErrorActionPreference = "Stop"

function Get-ToolPath {
    param(
        [string]$Preferred,
        [string]$CommandName,
        [string[]]$SearchGlobs
    )
    if ($Preferred -and (Test-Path $Preferred)) {
        return (Resolve-Path $Preferred).Path
    }
    $cmd = Get-Command $CommandName -ErrorAction SilentlyContinue
    if ($cmd) {
        return $cmd.Source
    }
    foreach ($glob in $SearchGlobs) {
        $match = Get-ChildItem -Path $glob -ErrorAction SilentlyContinue | Sort-Object FullName -Descending | Select-Object -First 1
        if ($match) {
            return $match.FullName
        }
    }
    throw "$CommandName not found. Install Windows SDK or pass explicit path."
}

function Get-AppVersionFromSource {
    param([string]$ProjectRoot)
    $versionFile = Join-Path $ProjectRoot "app\version.py"
    if (-not (Test-Path $versionFile)) {
        return "1.0.0.0"
    }
    $content = Get-Content $versionFile -Raw
    $m = [regex]::Match($content, 'APP_VERSION\s*=\s*"([^"]+)"')
    if (-not $m.Success) {
        return "1.0.0.0"
    }
    $raw = $m.Groups[1].Value
    $nums = [regex]::Matches($raw, '\d+') | ForEach-Object { $_.Value }
    while ($nums.Count -lt 4) { $nums += "0" }
    $nums = $nums[0..3]
    return ($nums -join ".")
}

function New-SolidPng {
    param(
        [string]$Path,
        [int]$Width,
        [int]$Height,
        [string]$Hex = "#0B7285"
    )
    Add-Type -AssemblyName System.Drawing
    $bmp = New-Object System.Drawing.Bitmap($Width, $Height)
    $graphics = [System.Drawing.Graphics]::FromImage($bmp)
    try {
        $graphics.Clear([System.Drawing.ColorTranslator]::FromHtml($Hex))
        $bmp.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally {
        $graphics.Dispose()
        $bmp.Dispose()
    }
}

if (-not $Version) {
    $Version = Get-AppVersionFromSource -ProjectRoot $RootDir
}

$exePath = Join-Path $RootDir "dist\FillableDOC.exe"
if (-not $SkipBuildExe) {
    Write-Host "Building dist\FillableDOC.exe..."
    & python -m PyInstaller --noconfirm (Join-Path $RootDir "Fillable.spec")
    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller build failed."
    }
}
if (-not (Test-Path $exePath)) {
    throw "Missing executable: $exePath"
}

$makeappx = Get-ToolPath -Preferred $MakeAppxPath -CommandName "makeappx" -SearchGlobs @(
    "C:\Program Files (x86)\Windows Kits\10\bin\*\x64\makeappx.exe",
    "C:\Program Files\Windows Kits\10\bin\*\x64\makeappx.exe"
)
$signtool = Get-ToolPath -Preferred $SignToolPath -CommandName "signtool" -SearchGlobs @(
    "C:\Program Files (x86)\Windows Kits\10\bin\*\x64\signtool.exe",
    "C:\Program Files\Windows Kits\10\bin\*\x64\signtool.exe"
)

$msixRoot = Join-Path $RootDir "dist\msix"
$layoutDir = Join-Path $msixRoot "layout"
$assetsDir = Join-Path $layoutDir "Assets"
New-Item -ItemType Directory -Force -Path $assetsDir | Out-Null

Copy-Item $exePath (Join-Path $layoutDir "FillableDOC.exe") -Force

New-SolidPng -Path (Join-Path $assetsDir "StoreLogo.png") -Width 50 -Height 50
New-SolidPng -Path (Join-Path $assetsDir "Square44x44Logo.png") -Width 44 -Height 44
New-SolidPng -Path (Join-Path $assetsDir "Square71x71Logo.png") -Width 71 -Height 71
New-SolidPng -Path (Join-Path $assetsDir "Square150x150Logo.png") -Width 150 -Height 150
New-SolidPng -Path (Join-Path $assetsDir "Wide310x150Logo.png") -Width 310 -Height 150

$manifestTemplate = Join-Path $RootDir "packaging\msix\AppxManifest.template.xml"
if (-not (Test-Path $manifestTemplate)) {
    throw "Missing manifest template: $manifestTemplate"
}
$manifest = Get-Content $manifestTemplate -Raw
$manifest = $manifest.Replace("__PACKAGE_NAME__", $PackageName)
$manifest = $manifest.Replace("__PUBLISHER__", $Publisher)
$manifest = $manifest.Replace("__VERSION__", $Version)
$manifest = $manifest.Replace("__DISPLAY_NAME__", $DisplayName)
$manifest = $manifest.Replace("__PUBLISHER_DISPLAY_NAME__", $PublisherDisplayName)
$manifestPath = Join-Path $layoutDir "AppxManifest.xml"
Set-Content -Path $manifestPath -Value $manifest -Encoding UTF8

$msixOut = Join-Path $msixRoot ("FillableDOC_{0}_x64.msix" -f $Version)
if (Test-Path $msixOut) {
    Remove-Item $msixOut -Force
}

Write-Host "Packing MSIX..."
& $makeappx pack /d $layoutDir /p $msixOut /o
if ($LASTEXITCODE -ne 0) {
    throw "makeappx pack failed."
}

if ($CertPath) {
    if (-not (Test-Path $CertPath)) {
        throw "Certificate file not found: $CertPath"
    }
    Write-Host "Signing MSIX..."
    $signArgs = @(
        "sign",
        "/fd", "SHA256",
        "/td", "SHA256",
        "/tr", "http://timestamp.digicert.com",
        "/f", $CertPath
    )
    if ($CertPassword) {
        $signArgs += @("/p", $CertPassword)
    }
    $signArgs += $msixOut
    & $signtool @signArgs
    if ($LASTEXITCODE -ne 0) {
        throw "signtool failed for MSIX."
    }
}
else {
    Write-Warning "No cert provided. MSIX package is unsigned."
}

Write-Host "MSIX ready: $msixOut"
