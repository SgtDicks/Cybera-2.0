[CmdletBinding()]
param(
  [Parameter(Mandatory)][string]$DatabasePath,
  [Parameter(Mandatory)][string]$OutputPath,
  [securestring]$DatabasePassword
)
$ErrorActionPreference='Stop'
if(-not(Test-Path -LiteralPath $DatabasePath -PathType Leaf)){throw 'DatabasePath does not exist.'}
$plainPassword=''
if($DatabasePassword){$plainPassword=[Net.NetworkCredential]::new('', $DatabasePassword).Password}
try {
  $connection=New-Object -ComObject ADODB.Connection
  $connection.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=$DatabasePath;Mode=Read;Jet OLEDB:Database Password=$plainPassword;"
  $connection.Open()
  function Read-Table([string]$name){
    try {
      $rs=$connection.Execute("SELECT * FROM [$name]")
      $rows=@()
      while(-not $rs.EOF){$row=[ordered]@{};for($i=0;$i -lt $rs.Fields.Count;$i++){$v=$rs.Fields.Item($i).Value;if($v -is [DBNull]){$v=$null};$row[$rs.Fields.Item($i).Name]=$v};$rows+=[pscustomobject]$row;$rs.MoveNext()}
      $rs.Close(); return $rows
    } catch { return $null }
  }
  $snapshot=[ordered]@{source='cybera-jet-readonly';observedAtUtc=[DateTimeOffset]::UtcNow.ToString('O');cards=(Read-Table cards);accounts=(Read-Table clients);pricing=(Read-Table recharges);operators=(Read-Table users);journal=@{sessions=(Read-Table journ_sessions);income=(Read-Table journ_income);events=(Read-Table events)};configuration=(Read-Table config)}
  $dir=Split-Path -Parent ([IO.Path]::GetFullPath($OutputPath));New-Item -ItemType Directory -Force -Path $dir|Out-Null;$tmp=Join-Path $dir ('.cybera-'+[guid]::NewGuid().ToString('N')+'.json')
  $snapshot|ConvertTo-Json -Depth 16|Set-Content -LiteralPath $tmp -Encoding utf8NoBOM;Move-Item -LiteralPath $tmp -Destination $OutputPath -Force
} finally {if($connection){$connection.Close()};$plainPassword=''}
