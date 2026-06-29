[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$VsdxPath,

    [string]$OutputDir,

    [string[]]$Formats = @('png'),

    [string]$PageName
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$resolvedVsdx = (Resolve-Path -LiteralPath $VsdxPath).Path
if (-not $OutputDir) {
    $OutputDir = Split-Path -Parent $resolvedVsdx
}
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

if ($Formats.Count -eq 1 -and $Formats[0] -match ',') {
    $Formats = $Formats[0].Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
}

$visFixedFormatPDF = 1
$visDocExIntentPrint = 1
$visPrintAll = 0

$visio = $null
$doc = $null
$page = $null
$generated = New-Object System.Collections.Generic.List[string]

try {
    $visio = New-Object -ComObject Visio.Application
    $visio.Visible = $false
    try { $visio.AlertResponse = 7 } catch {}
    $doc = $visio.Documents.Open($resolvedVsdx)

    if ($PageName) {
        $page = $doc.Pages.ItemU($PageName)
    }
    else {
        $page = $doc.Pages.Item(1)
    }

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedVsdx)

    foreach ($format in $Formats) {
        $normalized = $format.ToLowerInvariant()
        switch ($normalized) {
            'png' {
                $path = Join-Path $OutputDir ($baseName + '_' + $page.NameU + '.png')
                $null = $page.Export($path)
                $generated.Add($path)
            }
            'svg' {
                $path = Join-Path $OutputDir ($baseName + '_' + $page.NameU + '.svg')
                $null = $page.Export($path)
                $generated.Add($path)
            }
            'pdf' {
                $path = Join-Path $OutputDir ($baseName + '.pdf')
                $null = $doc.ExportAsFixedFormat($visFixedFormatPDF, $path, $visDocExIntentPrint, $visPrintAll)
                $generated.Add($path)
            }
            default {
                throw "Unsupported format: $format"
            }
        }
    }

    [pscustomobject]@{
        source = $resolvedVsdx
        page = $page.NameU
        outputs = $generated
    } | ConvertTo-Json -Depth 6
}
finally {
    if ($doc -ne $null) {
        try { $doc.Close() } catch {}
    }
    if ($visio -ne $null) {
        try { $visio.Quit() } catch {}
    }
}
