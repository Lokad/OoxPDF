param(
    [string] $RunPath,

    [string] $ReferenceChartStructures,

    [string] $CandidateChartStructures,

    [string] $ReferenceTextOperations,

    [string] $CandidateTextOperations,

    [string] $ReferenceChartTextStructures,

    [string] $CandidateChartTextStructures,

    [string] $ChartXml,

    [string] $OutputJson
)

$ErrorActionPreference = "Stop"

function Read-JsonArray($path) {
    if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path -LiteralPath $path)) {
        return ,@()
    }

    $items = Get-Content -Raw -LiteralPath (Resolve-Path -LiteralPath $path).Path | ConvertFrom-Json
    if ($null -eq $items) {
        return ,@()
    }

    if ($items -is [array]) {
        return ,$items
    }

    return ,@($items)
}

function Round([double] $value) {
    return [Math]::Round($value, 6)
}

function CenterX($item) {
    return ([double]$item.MinX + [double]$item.MaxX) / 2d
}

function CenterY($item) {
    return ([double]$item.MinY + [double]$item.MaxY) / 2d
}

function BoundsDelta($reference, $candidate) {
    return @(
        [Math]::Abs([double]$reference.MinX - [double]$candidate.MinX),
        [Math]::Abs([double]$reference.MinY - [double]$candidate.MinY),
        [Math]::Abs([double]$reference.MaxX - [double]$candidate.MaxX),
        [Math]::Abs([double]$reference.MaxY - [double]$candidate.MaxY)) |
        Measure-Object -Maximum |
        Select-Object -ExpandProperty Maximum
}

function Format-Bounds($item) {
    if ($null -eq $item) {
        return ""
    }

    return "({0},{1})-({2},{3})" -f `
        [Math]::Round([double]$item.MinX, 2),
        [Math]::Round([double]$item.MinY, 2),
        [Math]::Round([double]$item.MaxX, 2),
        [Math]::Round([double]$item.MaxY, 2)
}

function Select-Kind($items, [string] $kind) {
    return ,@($items | Where-Object { [string]$_.Kind -eq $kind } | Sort-Object -Property PageNumber, MinY, MinX, TextLength)
}

function Child-Val($node, [string] $name) {
    $child = $node.SelectSingleNode("*[local-name()='$name']")
    if ($null -eq $child -or $null -eq $child.Attributes["val"]) {
        return ""
    }

    return [string]$child.Attributes["val"].Value
}

function Manual-Layout-Val($manualLayout, [string] $name) {
    if ($null -eq $manualLayout) {
        return ""
    }

    return Child-Val $manualLayout $name
}

function Read-ChartLabelManualLayouts([string] $path) {
    if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path -LiteralPath $path)) {
        return ,@()
    }

    [xml] $xml = Get-Content -Raw -LiteralPath (Resolve-Path -LiteralPath $path).Path
    $labels = $xml.SelectNodes("//*[local-name()='dLbl']")
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($label in $labels) {
        $manualLayout = $label.SelectSingleNode("*[local-name()='layout']/*[local-name()='manualLayout']")
        if ($null -eq $manualLayout) {
            continue
        }

        $rows.Add([pscustomobject]@{
            Index = Child-Val $label "idx"
            Position = Child-Val $label "dLblPos"
            X = Manual-Layout-Val $manualLayout "x"
            Y = Manual-Layout-Val $manualLayout "y"
            Width = Manual-Layout-Val $manualLayout "w"
            Height = Manual-Layout-Val $manualLayout "h"
            XMode = Manual-Layout-Val $manualLayout "xMode"
            YMode = Manual-Layout-Val $manualLayout "yMode"
            WidthMode = Manual-Layout-Val $manualLayout "wMode"
            HeightMode = Manual-Layout-Val $manualLayout "hMode"
            ShowLeaderLines = Child-Val $label "showLeaderLines"
            ShowCategoryName = Child-Val $label "showCatName"
            ShowPercent = Child-Val $label "showPercent"
            ShowValue = Child-Val $label "showVal"
            ShowLegendKey = Child-Val $label "showLegendKey"
        })
    }

    return ,$rows.ToArray()
}

function Match-Nearest($referenceItems, $candidateItems, [string] $kind) {
    $rows = New-Object System.Collections.Generic.List[object]
    $used = New-Object System.Collections.Generic.HashSet[int]

    for ($refIndex = 0; $refIndex -lt $referenceItems.Count; $refIndex++) {
        $reference = $referenceItems[$refIndex]
        $bestIndex = -1
        $bestScore = [double]::PositiveInfinity

        for ($candidateIndex = 0; $candidateIndex -lt $candidateItems.Count; $candidateIndex++) {
            if ($used.Contains($candidateIndex)) {
                continue
            }

            $candidate = $candidateItems[$candidateIndex]
            $score = [Math]::Abs((CenterX $reference) - (CenterX $candidate)) +
                [Math]::Abs((CenterY $reference) - (CenterY $candidate))
            if ($score -lt $bestScore) {
                $bestScore = $score
                $bestIndex = $candidateIndex
            }
        }

        $candidateItem = $null
        if ($bestIndex -ge 0) {
            [void]$used.Add($bestIndex)
            $candidateItem = $candidateItems[$bestIndex]
        }

        $rows.Add([pscustomobject]@{
            Kind = $kind
            ReferenceIndex = $refIndex
            CandidateIndex = if ($bestIndex -ge 0) { $bestIndex } else { $null }
            BoundsDelta = if ($null -ne $candidateItem) { Round (BoundsDelta $reference $candidateItem) } else { $null }
            CenterDeltaX = if ($null -ne $candidateItem) { Round ((CenterX $candidateItem) - (CenterX $reference)) } else { $null }
            CenterDeltaY = if ($null -ne $candidateItem) { Round ((CenterY $candidateItem) - (CenterY $reference)) } else { $null }
            ReferenceBounds = Format-Bounds $reference
            CandidateBounds = Format-Bounds $candidateItem
            ReferenceTextHash = if ($reference.PSObject.Properties.Name -contains "TextHash") { [string]$reference.TextHash } else { "" }
            CandidateTextHash = if ($null -ne $candidateItem -and $candidateItem.PSObject.Properties.Name -contains "TextHash") { [string]$candidateItem.TextHash } else { "" }
        })
    }

    return ,$rows.ToArray()
}

function Ensure-ChartTextStructures($side, $textOperations, $chartStructures, $chartTextStructures) {
    if (-not [string]::IsNullOrWhiteSpace($chartTextStructures) -and (Test-Path -LiteralPath $chartTextStructures)) {
        return $chartTextStructures
    }

    if ([string]::IsNullOrWhiteSpace($RunPath)) {
        return $chartTextStructures
    }

    if ([string]::IsNullOrWhiteSpace($textOperations) -or -not (Test-Path -LiteralPath $textOperations)) {
        return $chartTextStructures
    }

    if ([string]::IsNullOrWhiteSpace($chartStructures) -or -not (Test-Path -LiteralPath $chartStructures)) {
        return $chartTextStructures
    }

    $outputDirectory = Join-Path $RunPath "comparison/pdf-data-label-layout/$side"
    New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null
    $output = Join-Path $outputDirectory "chart-text-structures.json"
    & (Join-Path $PSScriptRoot "ClassifyPdfChartText.ps1") `
        -InputPath $textOperations `
        -ChartStructures $chartStructures `
        -Output $output `
        -PageNumber 1 *> $null
    return $output
}

if (-not [string]::IsNullOrWhiteSpace($RunPath)) {
    $runFull = (Resolve-Path -LiteralPath $RunPath).Path
    $graphicsRoot = Join-Path $runFull "comparison/pdf-graphics"
    $ReferenceChartStructures = if ([string]::IsNullOrWhiteSpace($ReferenceChartStructures)) { Join-Path $graphicsRoot "reference/chart-structures.json" } else { $ReferenceChartStructures }
    $CandidateChartStructures = if ([string]::IsNullOrWhiteSpace($CandidateChartStructures)) { Join-Path $graphicsRoot "candidate/chart-structures.json" } else { $CandidateChartStructures }
    $ReferenceTextOperations = if ([string]::IsNullOrWhiteSpace($ReferenceTextOperations)) { Join-Path $graphicsRoot "reference/text-operations.json" } else { $ReferenceTextOperations }
    $CandidateTextOperations = if ([string]::IsNullOrWhiteSpace($CandidateTextOperations)) { Join-Path $graphicsRoot "candidate/text-operations.json" } else { $CandidateTextOperations }
}

$ReferenceChartTextStructures = Ensure-ChartTextStructures "reference" $ReferenceTextOperations $ReferenceChartStructures $ReferenceChartTextStructures
$CandidateChartTextStructures = Ensure-ChartTextStructures "candidate" $CandidateTextOperations $CandidateChartStructures $CandidateChartTextStructures

$referenceGraphics = Read-JsonArray $ReferenceChartStructures
$candidateGraphics = Read-JsonArray $CandidateChartStructures
$referenceText = Read-JsonArray $ReferenceChartTextStructures
$candidateText = Read-JsonArray $CandidateChartTextStructures

$referenceLabels = Select-Kind $referenceText "DataLabelText"
$candidateLabels = Select-Kind $candidateText "DataLabelText"
$referenceLeaderLines = Select-Kind $referenceGraphics "DataLabelLeaderLineCandidate"
$candidateLeaderLines = Select-Kind $candidateGraphics "DataLabelLeaderLineCandidate"
$chartManualLayouts = Read-ChartLabelManualLayouts $ChartXml

$labelMatches = Match-Nearest $referenceLabels $candidateLabels "DataLabelText"
$leaderLineMatches = Match-Nearest $referenceLeaderLines $candidateLeaderLines "DataLabelLeaderLineCandidate"

$summary = [pscustomobject]@{
    DataLabelTextReferenceCount = $referenceLabels.Count
    DataLabelTextCandidateCount = $candidateLabels.Count
    DataLabelTextMaxNearestBoundsDelta = if ($labelMatches.Count -eq 0) { $null } else { ($labelMatches | Where-Object { $null -ne $_.BoundsDelta } | Measure-Object -Property BoundsDelta -Maximum).Maximum }
    DataLabelLeaderLineReferenceCount = $referenceLeaderLines.Count
    DataLabelLeaderLineCandidateCount = $candidateLeaderLines.Count
    DataLabelLeaderLineMaxNearestBoundsDelta = if ($leaderLineMatches.Count -eq 0) { $null } else { ($leaderLineMatches | Where-Object { $null -ne $_.BoundsDelta } | Measure-Object -Property BoundsDelta -Maximum).Maximum }
    ChartManualLayoutCount = $chartManualLayouts.Count
    ChartManualLayouts = $chartManualLayouts
    DataLabelTextMatches = $labelMatches
    DataLabelLeaderLineMatches = $leaderLineMatches
}

$summary | Format-List DataLabelTextReferenceCount, DataLabelTextCandidateCount, DataLabelTextMaxNearestBoundsDelta, DataLabelLeaderLineReferenceCount, DataLabelLeaderLineCandidateCount, DataLabelLeaderLineMaxNearestBoundsDelta, ChartManualLayoutCount

if ($chartManualLayouts.Count -gt 0) {
    Write-Host ""
    Write-Host "Chart data-label manual layouts:"
    $chartManualLayouts | Format-Table -AutoSize Index, Position, X, Y, Width, Height, XMode, YMode, WidthMode, HeightMode, ShowCategoryName, ShowPercent, ShowLeaderLines
}

if ($labelMatches.Count -gt 0) {
    Write-Host ""
    Write-Host "Data-label text nearest matches:"
    $labelMatches | Format-Table -AutoSize Kind, ReferenceIndex, CandidateIndex, BoundsDelta, CenterDeltaX, CenterDeltaY, ReferenceBounds, CandidateBounds
}

if ($leaderLineMatches.Count -gt 0) {
    Write-Host ""
    Write-Host "Data-label leader-line nearest matches:"
    $leaderLineMatches | Format-Table -AutoSize Kind, ReferenceIndex, CandidateIndex, BoundsDelta, CenterDeltaX, CenterDeltaY, ReferenceBounds, CandidateBounds
}

if (-not [string]::IsNullOrWhiteSpace($OutputJson)) {
    $outputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputJson)
    $outputDirectory = Split-Path -Parent $outputPath
    if (-not [string]::IsNullOrWhiteSpace($outputDirectory)) {
        New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null
    }

    $summary | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $outputPath -Encoding UTF8
}
