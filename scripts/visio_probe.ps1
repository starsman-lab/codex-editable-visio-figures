[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Find-VisioExe {
    $candidates = @(
        'C:\Program Files\Microsoft Office\root\Office16\VISIO.EXE',
        'C:\Program Files (x86)\Microsoft Office\root\Office16\VISIO.EXE'
    )

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            return $candidate
        }
    }

    $hits = Get-ChildItem 'C:\Program Files', 'C:\Program Files (x86)' -Recurse -Filter 'VISIO.EXE' -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($hits) {
        return $hits.FullName
    }

    return $null
}

$result = [ordered]@{
    visioExe = Find-VisioExe
    comAvailable = $false
    version = $null
    canCreateDocument = $false
    pageCount = $null
    error = $null
}

$visio = $null
$doc = $null

try {
    $visio = New-Object -ComObject Visio.Application
    $result.comAvailable = $true
    $result.version = [string]$visio.Version
    $visio.Visible = $false
    $doc = $visio.Documents.Add('')
    $result.canCreateDocument = $true
    $result.pageCount = [int]$doc.Pages.Count
}
catch {
    $result.error = $_.Exception.Message
}
finally {
    if ($doc -ne $null) {
        try { $doc.Close() } catch {}
    }
    if ($visio -ne $null) {
        try { $visio.Quit() } catch {}
    }
}

$result | ConvertTo-Json -Depth 6

