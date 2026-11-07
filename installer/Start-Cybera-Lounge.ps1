[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$ConfigPath = $env:CYBERA_BRIDGE_CONFIG,
    [string]$Urls = "http://0.0.0.0:5088"
)

$ErrorActionPreference = "Stop"

function Resolve-ConfiguredFile([string]$Path, [string]$Description) {
    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw "$Description is not configured."
    }

    $expandedPath = [Environment]::ExpandEnvironmentVariables($Path)
    if (-not [IO.Path]::GetExtension($expandedPath).Equals(".exe", [StringComparison]::OrdinalIgnoreCase)) {
        throw "$Description must be an .exe path."
    }
    if (-not (Test-Path -LiteralPath $expandedPath -PathType Leaf)) {
        throw "$Description does not exist."
    }
    return (Resolve-Path -LiteralPath $expandedPath).Path
}

function Resolve-ConfiguredDirectory([string]$Path, [string]$Description) {
    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw "$Description is not configured."
    }

    $expandedPath = [Environment]::ExpandEnvironmentVariables($Path)
    if (-not (Test-Path -LiteralPath $expandedPath -PathType Container)) {
        throw "$Description does not exist."
    }
    return (Resolve-Path -LiteralPath $expandedPath).Path
}

if ([string]::IsNullOrWhiteSpace($ConfigPath) -or -not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) {
    throw "Bridge configuration was not found. Set CYBERA_BRIDGE_CONFIG or pass -ConfigPath with the ProgramData appsettings.json path."
}

try {
    $config = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
}
catch {
    throw "Bridge configuration is not valid JSON."
}

$bridgeConfig = $config.Bridge
if ($null -eq $bridgeConfig) {
    throw "Bridge configuration does not contain a Bridge section."
}

$legacyProcessName = [string]$bridgeConfig.LegacyProcessName
if ($legacyProcessName -notmatch '^[A-Za-z0-9_.-]+$') {
    throw "Bridge:LegacyProcessName must be a process name, without a path."
}

$bridgeExecutable = Resolve-ConfiguredFile (Join-Path $PSScriptRoot "Cybera.ModernBridge.exe") "Installed bridge executable"
$bridgeProcessRunning = $null -ne (Get-Process -Name "Cybera.ModernBridge" -ErrorAction SilentlyContinue | Select-Object -First 1)

if (-not $bridgeProcessRunning -and [string]::IsNullOrWhiteSpace($env:CYBERA_BRIDGE_API_KEY)) {
    throw "CYBERA_BRIDGE_API_KEY is missing from this launcher process. The bridge was not started."
}

$legacyProcessRunning = $null -ne (Get-Process -Name $legacyProcessName -ErrorAction SilentlyContinue | Select-Object -First 1)
if ($legacyProcessRunning) {
    Write-Host "Legacy Cybera server process '$legacyProcessName' is already running."
}
else {
    $legacy = $bridgeConfig.LegacyServer
    if ($null -eq $legacy) {
        throw "Bridge:LegacyServer is missing. The legacy Cybera server was not started."
    }

    $legacyExecutable = Resolve-ConfiguredFile ([string]$legacy.ExecutablePath) "Bridge:LegacyServer:ExecutablePath"
    $legacyWorkingDirectory = Resolve-ConfiguredDirectory ([string]$legacy.WorkingDirectory) "Bridge:LegacyServer:WorkingDirectory"
    $waitSeconds = [int]$legacy.StartupWaitSeconds
    if ($waitSeconds -lt 1 -or $waitSeconds -gt 60) {
        throw "Bridge:LegacyServer:StartupWaitSeconds must be between 1 and 60."
    }

    if ($PSCmdlet.ShouldProcess($legacyExecutable, "Start legacy Cybera server")) {
        Start-Process -FilePath $legacyExecutable -WorkingDirectory $legacyWorkingDirectory
        Write-Host "Started legacy Cybera server. Waiting $waitSeconds seconds for initialization..."
        Start-Sleep -Seconds $waitSeconds
        if ($null -eq (Get-Process -Name $legacyProcessName -ErrorAction SilentlyContinue | Select-Object -First 1)) {
            throw "Legacy Cybera server did not remain running after initialization. The bridge was not started."
        }
    }
}

if ($bridgeProcessRunning) {
    Write-Host "Cybera Modern Bridge is already running."
    exit 0
}

$env:CYBERA_BRIDGE_CONFIG = (Resolve-Path -LiteralPath $ConfigPath).Path
if ($PSCmdlet.ShouldProcess($bridgeExecutable, "Start Cybera Modern Bridge")) {
    Start-Process -FilePath $bridgeExecutable -ArgumentList @("--urls", $Urls) -WorkingDirectory $PSScriptRoot -WindowStyle Hidden
    Write-Host "Started Cybera Modern Bridge at $Urls."
}
