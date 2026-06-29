[CmdletBinding()]
param(
    [string]$VsdxPath,
    [switch]$Visible
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$visio = New-Object -ComObject Visio.Application
$visio.Visible = $Visible.IsPresent

if ($VsdxPath) {
    $resolved = (Resolve-Path -LiteralPath $VsdxPath).Path
    [void]$visio.Documents.Open($resolved)
}
else {
    [void]$visio.Documents.Add('')
}

[pscustomobject]@{
    visible = [bool]$visio.Visible
    documentCount = [int]$visio.Documents.Count
} | ConvertTo-Json -Depth 4

