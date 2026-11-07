[CmdletBinding()]
param(
    [string]$Version = "0.1.5",
    [switch]$SkipPublish
)

$ErrorActionPreference = "Stop"
$installerRoot = Split-Path -Parent $PSCommandPath
$projectRoot = Split-Path -Parent $installerRoot
$bridgeProject = Join-Path $projectRoot "modern-bridge\Cybera.ModernBridge.csproj"
$publishDirectory = Join-Path $projectRoot "artifacts\publish\win-x64"
$installerSource = Join-Path $installerRoot "CyberaModernBridge.iss"

if (-not $SkipPublish) {
    dotnet publish $bridgeProject --configuration Release --runtime win-x64 --self-contained false --output $publishDirectory --nologo
    if ($LASTEXITCODE -ne 0) { throw "dotnet publish failed." }
}

if (-not (Test-Path (Join-Path $publishDirectory "Cybera.ModernBridge.exe"))) {
    throw "Published bridge executable was not found at $publishDirectory. Run without -SkipPublish."
}

$isccCandidates = @(
    (Get-Command "ISCC.exe" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source),
    "C:\Program Files\Inno Setup 7\ISCC.exe",
    "C:\Program Files (x86)\Inno Setup 7\ISCC.exe",
    "C:\Program Files\Inno Setup 6\ISCC.exe",
    "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
) | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1
if (-not $isccCandidates) {
    throw "Inno Setup 6 is required to build the installer. Install it, ensure ISCC.exe is on PATH, then rerun this script. The validated publish output is at $publishDirectory."
}

& $isccCandidates "/DMyAppVersion=$Version" "/DPublishDir=$publishDirectory" $installerSource
if ($LASTEXITCODE -ne 0) { throw "Inno Setup compilation failed." }

Write-Host "Installer created in $(Join-Path $projectRoot 'artifacts\installer')"
