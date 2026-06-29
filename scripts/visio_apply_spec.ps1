[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SpecPath,

    [Parameter(Mandatory = $true)]
    [string]$VsdxPath,

    [Parameter(Mandatory = $true)]
    [string]$SeedVsdxPath,

    [string]$OutputDir,

    [string[]]$ExportFormats = @(),

    [switch]$ReplacePage
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-OptionalValue {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    $prop = $Object.PSObject.Properties[$Name]
    if ($null -eq $prop) { return $null }
    return $prop.Value
}

function Convert-HexToRgbFormula {
    param([string]$Hex)
    if (-not $Hex) { return $null }
    $clean = $Hex.Trim()
    if ($clean.StartsWith('#')) { $clean = $clean.Substring(1) }
    if ($clean.Length -ne 6) { throw "Expected 6-digit hex color, got: $Hex" }
    $r = [Convert]::ToInt32($clean.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($clean.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($clean.Substring(4, 2), 16)
    "RGB($r,$g,$b)"
}

function Set-CellIfPresent {
    param($Shape, [string]$CellName, [string]$Formula)
    if ($null -ne $Formula -and $Formula -ne '') {
        $Shape.CellsU($CellName).FormulaU = $Formula
    }
}

function Set-ShapeTextStyle {
    param($Shape, $Spec)
    $fontSize = Get-OptionalValue -Object $Spec -Name 'fontSize'
    $textColor = Get-OptionalValue -Object $Spec -Name 'textColor'
    if ($null -ne $fontSize) {
        $Shape.CellsU('Char.Size').FormulaU = ("{0} pt" -f [double]$fontSize)
    }
    if ($null -ne $textColor) {
        $Shape.CellsU('Char.Color').FormulaU = Convert-HexToRgbFormula -Hex ([string]$textColor)
    }
}

function Set-ShapeStyle {
    param($Shape, $Spec)
    $type = ([string](Get-OptionalValue -Object $Spec -Name 'type')).ToLowerInvariant()
    $fill = Get-OptionalValue -Object $Spec -Name 'fill'
    $line = Get-OptionalValue -Object $Spec -Name 'line'
    $lineWeight = Get-OptionalValue -Object $Spec -Name 'lineWeight'
    $rounding = Get-OptionalValue -Object $Spec -Name 'rounding'
    $endArrow = Get-OptionalValue -Object $Spec -Name 'endArrow'
    $text = Get-OptionalValue -Object $Spec -Name 'text'

    if ($null -ne $fill) {
        Set-CellIfPresent -Shape $Shape -CellName 'FillForegnd' -Formula (Convert-HexToRgbFormula -Hex ([string]$fill))
    }
    elseif ($type -in @('text','line','connector')) {
        try { $Shape.CellsU('FillPattern').FormulaU = '0' } catch {}
    }

    if ($null -ne $line) {
        Set-CellIfPresent -Shape $Shape -CellName 'LineColor' -Formula (Convert-HexToRgbFormula -Hex ([string]$line))
    }
    elseif ($type -eq 'text') {
        try { $Shape.CellsU('LinePattern').FormulaU = '0' } catch {}
    }

    if ($null -ne $lineWeight) {
        $Shape.CellsU('LineWeight').FormulaU = [string][double]$lineWeight
    }
    if ($null -ne $rounding) {
        $Shape.CellsU('Rounding').FormulaU = [string][double]$rounding
    }
    if ($null -ne $endArrow) {
        $Shape.CellsU('EndArrow').FormulaU = [string][int]$endArrow
    }
    if ($null -ne $text) {
        $Shape.Text = [string]$text
    }

    Set-ShapeTextStyle -Shape $Shape -Spec $Spec
}

function Convert-Rect {
    param([double]$PageHeight, $Spec)
    $left = [double]$Spec.x
    $right = [double]$Spec.x + [double]$Spec.width
    $top = $PageHeight - [double]$Spec.y
    $bottom = $PageHeight - ([double]$Spec.y + [double]$Spec.height)
    [pscustomobject]@{ Left = $left; Right = $right; Top = $top; Bottom = $bottom }
}

function Convert-Point {
    param([double]$PageHeight, [double[]]$Point)
    [pscustomobject]@{ X = [double]$Point[0]; Y = $PageHeight - [double]$Point[1] }
}

$resolvedSpec = (Resolve-Path -LiteralPath $SpecPath).Path
$resolvedSeed = (Resolve-Path -LiteralPath $SeedVsdxPath).Path
$resolvedTarget = [System.IO.Path]::GetFullPath($VsdxPath)
$targetDir = Split-Path -Parent $resolvedTarget
New-Item -ItemType Directory -Force -Path $targetDir | Out-Null
if (-not (Test-Path -LiteralPath $resolvedTarget)) {
    Copy-Item -LiteralPath $resolvedSeed -Destination $resolvedTarget -Force
}
if (-not $OutputDir) { $OutputDir = $targetDir }
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
if ($ExportFormats.Count -eq 1 -and $ExportFormats[0] -match ',') {
    $ExportFormats = $ExportFormats[0].Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
}

$spec = Get-Content -LiteralPath $resolvedSpec -Raw | ConvertFrom-Json
if ($null -eq $spec.page) { throw 'Spec must contain a page object.' }
if ($null -eq $spec.shapes) { throw 'Spec must contain a shapes array.' }
$pageName = if (Get-OptionalValue -Object $spec.page -Name 'name') { [string](Get-OptionalValue -Object $spec.page -Name 'name') } else { 'Page-1' }
$pageWidth = [double](Get-OptionalValue -Object $spec.page -Name 'width')
$pageHeight = [double](Get-OptionalValue -Object $spec.page -Name 'height')

$visio = $null
$doc = $null
$page = $null
$backupPath = $null
$createdShapes = New-Object System.Collections.Generic.List[string]

try {
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $backupPath = Join-Path $targetDir (([System.IO.Path]::GetFileNameWithoutExtension($resolvedTarget)) + "_backup_$timestamp.vsdx")
    Copy-Item -LiteralPath $resolvedTarget -Destination $backupPath -Force

    $visio = New-Object -ComObject Visio.Application
    $visio.Visible = $false
    try { $visio.AlertResponse = 7 } catch {}
    $doc = $visio.Documents.Open($resolvedTarget)

    try { $page = $doc.Pages.ItemU($pageName) } catch { $page = $doc.Pages.Item(1); try { $page.NameU = $pageName } catch {} }

    if ($ReplacePage.IsPresent) {
        while ($page.Shapes.Count -gt 0) { $page.Shapes.Item(1).Delete() }
    }

    $page.PageSheet.CellsU('PageWidth').FormulaU = [string]$pageWidth
    $page.PageSheet.CellsU('PageHeight').FormulaU = [string]$pageHeight

    foreach ($shapeSpec in $spec.shapes) {
        $shapeType = ([string](Get-OptionalValue -Object $shapeSpec -Name 'type')).ToLowerInvariant()
        $shape = $null
        switch ($shapeType) {
            'rect' {
                $rect = Convert-Rect -PageHeight $pageHeight -Spec $shapeSpec
                $shape = $page.DrawRectangle($rect.Left, $rect.Bottom, $rect.Right, $rect.Top)
            }
            'ellipse' {
                $rect = Convert-Rect -PageHeight $pageHeight -Spec $shapeSpec
                $shape = $page.DrawOval($rect.Left, $rect.Bottom, $rect.Right, $rect.Top)
            }
            'text' {
                $rect = Convert-Rect -PageHeight $pageHeight -Spec $shapeSpec
                $shape = $page.DrawRectangle($rect.Left, $rect.Bottom, $rect.Right, $rect.Top)
            }
            'line' {
                $from = Convert-Point -PageHeight $pageHeight -Point ([double[]](Get-OptionalValue -Object $shapeSpec -Name 'from'))
                $to = Convert-Point -PageHeight $pageHeight -Point ([double[]](Get-OptionalValue -Object $shapeSpec -Name 'to'))
                $shape = $page.DrawLine($from.X, $from.Y, $to.X, $to.Y)
            }
            'connector' {
                $from = Convert-Point -PageHeight $pageHeight -Point ([double[]](Get-OptionalValue -Object $shapeSpec -Name 'from'))
                $to = Convert-Point -PageHeight $pageHeight -Point ([double[]](Get-OptionalValue -Object $shapeSpec -Name 'to'))
                $shape = $page.DrawLine($from.X, $from.Y, $to.X, $to.Y)
            }
            default { throw "Unsupported shape type: $shapeType" }
        }
        Set-ShapeStyle -Shape $shape -Spec $shapeSpec
        $shapeId = Get-OptionalValue -Object $shapeSpec -Name 'id'
        if ($null -ne $shapeId) {
            try { $shape.NameU = [string]$shapeId } catch {}
            $createdShapes.Add([string]$shapeId)
        } else {
            $createdShapes.Add([string]$shape.NameU)
        }
    }

    $doc.Save()

    $outputs = New-Object System.Collections.Generic.List[string]
    if ($ExportFormats.Count -gt 0) {
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedTarget)
        foreach ($format in $ExportFormats) {
            $normalized = $format.ToLowerInvariant()
            switch ($normalized) {
                'png' {
                    $exportPath = Join-Path $OutputDir ($baseName + '_' + $page.NameU + '.png')
                    $page.Export($exportPath)
                    $outputs.Add($exportPath)
                }
                'svg' {
                    $exportPath = Join-Path $OutputDir ($baseName + '_' + $page.NameU + '.svg')
                    $page.Export($exportPath)
                    $outputs.Add($exportPath)
                }
                'pdf' {
                    $visFixedFormatPDF = 1
                    $visDocExIntentPrint = 1
                    $visPrintAll = 0
                    $exportPath = Join-Path $OutputDir ($baseName + '.pdf')
                    $doc.ExportAsFixedFormat($visFixedFormatPDF, $exportPath, $visDocExIntentPrint, $visPrintAll)
                    $outputs.Add($exportPath)
                }
                default {
                    throw "Unsupported format: $format"
                }
            }
        }
    }

    [pscustomobject]@{
        target = $resolvedTarget
        seed = $resolvedSeed
        backup = $backupPath
        page = $page.NameU
        shapeCount = $createdShapes.Count
        shapes = $createdShapes
        exports = @($outputs)
    } | ConvertTo-Json -Depth 8
}
finally {
    if ($doc -ne $null) { try { $doc.Close() } catch {} }
    if ($visio -ne $null) { try { $visio.Quit() } catch {} }
}

