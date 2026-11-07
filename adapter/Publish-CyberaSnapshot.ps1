[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$InputPath,
    [Parameter(Mandatory)][string]$OutputPath
)

$ErrorActionPreference = 'Stop'
$allowed = @('workstations','sessions','accounts','cards','pricing','operators','journal','reports','configuration')
if (-not (Test-Path -LiteralPath $InputPath -PathType Leaf)) { throw 'Input snapshot does not exist.' }

try { $snapshot = Get-Content -LiteralPath $InputPath -Raw | ConvertFrom-Json } catch { throw 'Input snapshot is not valid JSON.' }
if ($null -eq $snapshot -or $snapshot -isnot [psobject]) { throw 'Input snapshot must be a JSON object.' }

$unknown = @($snapshot.psobject.Properties.Name | Where-Object { $_ -notin $allowed -and $_ -ne 'observedAtUtc' -and $_ -ne 'source' })
if ($unknown.Count) { throw "Snapshot has unsupported top-level fields: $($unknown -join ', ')" }
foreach ($collection in $allowed) {
    if ($null -ne $snapshot.$collection -and $snapshot.$collection -isnot [array] -and $snapshot.$collection -isnot [psobject]) {
        throw "Snapshot collection '$collection' must be an object or array."
    }
}

$outputFull = [IO.Path]::GetFullPath($OutputPath)
$outputDir = Split-Path -Parent $outputFull
New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
$temporary = Join-Path $outputDir ('.snapshot-' + [guid]::NewGuid().ToString('N') + '.json')
try {
    $payload = [ordered]@{ observedAtUtc = [DateTimeOffset]::UtcNow.ToString('O'); source = 'host-local-exporter' }
    foreach ($property in $snapshot.psobject.Properties) { if ($property.Name -notin @('observedAtUtc','source')) { $payload[$property.Name] = $property.Value } }
    $payload | ConvertTo-Json -Depth 16 | Set-Content -LiteralPath $temporary -Encoding utf8NoBOM
    Move-Item -LiteralPath $temporary -Destination $outputFull -Force
}
finally {
    if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force }
}
