param(
    [string] $RunPath,

    [string] $ReferenceChartStructures,

    [string] $CandidateChartStructures,

    [string] $ReferenceTextOperations,

    [string] $CandidateTextOperations,

    [string] $ReferenceChartTextStructures,

    [string] $CandidateChartTextStructures,

    [string] $ChartXml,

    [string] $Pptx,

    [string] $ChartPart = "ppt/charts/chart1.xml",

    [string] $ComMetadataJson,

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

function Read-JsonObject($path) {
    if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path -LiteralPath $path)) {
        return $null
    }

    return Get-Content -Raw -LiteralPath (Resolve-Path -LiteralPath $path).Path | ConvertFrom-Json
}

function Resolve-ComMetadataJsonPath([string] $explicitPath, [string] $pptxPath) {
    if (-not [string]::IsNullOrWhiteSpace($explicitPath)) {
        return $explicitPath
    }

    if ([string]::IsNullOrWhiteSpace($pptxPath)) {
        return ""
    }

    $fixtureName = [System.IO.Path]::GetFileNameWithoutExtension($pptxPath)
    if ([string]::IsNullOrWhiteSpace($fixtureName)) {
        return ""
    }

    $repoRoot = Split-Path -Parent $PSScriptRoot
    $candidate = Join-Path $repoRoot ("artifacts/office-probe-metadata/{0}/com-metadata.json" -f $fixtureName)
    if (Test-Path -LiteralPath $candidate) {
        return $candidate
    }

    return ""
}

function Round([double] $value) {
    return [Math]::Round($value, 6)
}

function Stable-TextHash([string] $text) {
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($text)
    $hash = [System.Security.Cryptography.SHA256]::HashData($bytes)
    return [Convert]::ToHexString($hash).Substring(0, 16)
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

function BoundsDistance($left, $right) {
    if ($null -eq $left -or $null -eq $right) {
        return $null
    }

    $dx = 0d
    if ([double]$left.MaxX -lt [double]$right.MinX) {
        $dx = [double]$right.MinX - [double]$left.MaxX
    }
    elseif ([double]$right.MaxX -lt [double]$left.MinX) {
        $dx = [double]$left.MinX - [double]$right.MaxX
    }

    $dy = 0d
    if ([double]$left.MaxY -lt [double]$right.MinY) {
        $dy = [double]$right.MinY - [double]$left.MaxY
    }
    elseif ([double]$right.MaxY -lt [double]$left.MinY) {
        $dy = [double]$left.MinY - [double]$right.MaxY
    }

    return Round ([Math]::Sqrt($dx * $dx + $dy * $dy))
}

function Select-PrimaryKind($items, [string] $kind) {
    $matches = @($items | Where-Object { [string]$_.Kind -eq $kind } | Sort-Object -Property PageNumber, MinY, MinX)
    if ($matches.Count -eq 0) {
        return $null
    }

    return $matches[0]
}

function Quadrant($item, $box) {
    if ($null -eq $item -or $null -eq $box) {
        return ""
    }

    $horizontal = if ((CenterX $item) -lt (CenterX $box)) { "left" } else { "right" }
    $vertical = if ((CenterY $item) -lt (CenterY $box)) { "top" } else { "bottom" }
    return "$vertical-$horizontal"
}

function Manual-Quadrant($layout) {
    $x = 0d
    $y = 0d
    $hasX = [double]::TryParse([string]$layout.X, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$x)
    $hasY = [double]::TryParse([string]$layout.Y, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$y)
    if (-not $hasX -or -not $hasY) {
        return ""
    }

    $horizontal = if ($x -lt 0d) { "left" } else { "right" }
    $vertical = if ($y -lt 0d) { "top" } else { "bottom" }
    return "$vertical-$horizontal"
}

function Relative-Center($item, $box, [string] $axis) {
    if ($null -eq $item -or $null -eq $box) {
        return $null
    }

    if ($axis -eq "x") {
        $width = [double]$box.Width
        if ($width -le 0d) {
            return $null
        }

        return Round (((CenterX $item) - (CenterX $box)) / $width)
    }

    $height = [double]$box.Height
    if ($height -le 0d) {
        return $null
    }

    return Round (((CenterY $item) - (CenterY $box)) / $height)
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

function New-BoundsItem([double] $minX, [double] $minY, [double] $maxX, [double] $maxY, [string] $textHash) {
    return [pscustomobject]@{
        MinX = $minX
        MinY = $minY
        MaxX = $maxX
        MaxY = $maxY
        TextHash = $textHash
    }
}

function Select-Kind($items, [string] $kind) {
    return ,@($items | Where-Object { [string]$_.Kind -eq $kind } | Sort-Object -Property PageNumber, MinY, MinX, TextLength)
}

function Group-TextClusters($items, [double] $baselineTolerance = 2.0, [double] $maxGap = 120.0) {
    $rows = New-Object System.Collections.Generic.List[object]
    $ordered = @($items | Sort-Object -Property PageNumber, MinY, MinX)
    $clusters = New-Object System.Collections.Generic.List[object]

    function Is-ClusterNeighbor($left, $right, [double] $baselineTolerance, [double] $maxGap) {
        if ([int]$left.PageNumber -ne [int]$right.PageNumber) {
            return $false
        }

        $dx = [Math]::Abs((CenterX $left) - (CenterX $right))
        $dy = [Math]::Abs((CenterY $left) - (CenterY $right))
        $sameLine = $dy -le $baselineTolerance -and $dx -le $maxGap
        $stackedLine = $dx -le 70.0 -and $dy -le 35.0
        return $sameLine -or $stackedLine
    }

    foreach ($item in $ordered) {
        $targetCluster = $null
        foreach ($cluster in $clusters) {
            foreach ($member in $cluster) {
                if (Is-ClusterNeighbor $member $item $baselineTolerance $maxGap) {
                    $targetCluster = $cluster
                    break
                }
            }

            if ($null -ne $targetCluster) {
                break
            }
        }

        if ($null -eq $targetCluster) {
            $targetCluster = New-Object System.Collections.Generic.List[object]
            $clusters.Add($targetCluster)
        }

        $targetCluster.Add($item)
    }

    foreach ($clusterItems in $clusters) {
        $minX = ($clusterItems | Measure-Object -Property MinX -Minimum).Minimum
        $minY = ($clusterItems | Measure-Object -Property MinY -Minimum).Minimum
        $maxX = ($clusterItems | Measure-Object -Property MaxX -Maximum).Maximum
        $maxY = ($clusterItems | Measure-Object -Property MaxY -Maximum).Maximum
        $hashes = @($clusterItems | Sort-Object -Property MinY, MinX | ForEach-Object {
            if ($_.PSObject.Properties.Name -contains "TextHash") { [string]$_.TextHash } else { "" }
        }) -join "+"
        $rows.Add((New-BoundsItem ([double]$minX) ([double]$minY) ([double]$maxX) ([double]$maxY) $hashes))
    }

    return ,@($rows.ToArray() | Sort-Object -Property MinY, MinX)
}

function Child-Val($node, [string] $name) {
    $child = $node.SelectSingleNode("*[local-name()='$name']")
    if ($null -eq $child -or $null -eq $child.Attributes["val"]) {
        return ""
    }

    return [string]$child.Attributes["val"].Value
}

function Child-Text($node, [string] $name) {
    $child = $node.SelectSingleNode("*[local-name()='$name']")
    if ($null -eq $child) {
        return ""
    }

    return [string]$child.InnerText
}

function Optional-Boolean($node, $fallbackNode, [string] $name, [bool] $defaultValue) {
    $value = Child-Val $node $name
    if ([string]::IsNullOrWhiteSpace($value) -and $null -ne $fallbackNode) {
        $value = Child-Val $fallbackNode $name
    }

    if ([string]::IsNullOrWhiteSpace($value)) {
        return $defaultValue
    }

    return -not ($value -eq "0" -or $value -eq "false")
}

function Read-IndexedTextCache($seriesNode, [string] $containerName) {
    $map = @{}
    $points = $seriesNode.SelectNodes("*[local-name()='$containerName']//*[local-name()='pt']")
    foreach ($point in $points) {
        $idx = Child-Val $point "idx"
        if ([string]::IsNullOrWhiteSpace($idx) -and $point.Attributes["idx"] -ne $null) {
            $idx = [string]$point.Attributes["idx"].Value
        }

        if ([string]::IsNullOrWhiteSpace($idx)) {
            continue
        }

        $map[$idx] = Child-Text $point "v"
    }

    return $map
}

function Format-ChartPercent([double] $value, [double] $total) {
    if ($total -eq 0d) {
        return ""
    }

    return ("{0:P0}" -f ($value / $total))
}

function Read-ExpectedLabelParts($seriesNode, $labelsNode, $labelNode, [string] $index, $categoryMap, $valueMap, [double] $total) {
    $parts = New-Object System.Collections.Generic.List[string]
    if (Optional-Boolean $labelNode $labelsNode "showCatName" $false) {
        $category = if ($categoryMap.ContainsKey($index)) { [string]$categoryMap[$index] } else { "" }
        if (-not [string]::IsNullOrEmpty($category)) {
            $parts.Add($category)
        }
    }

    $valueText = if ($valueMap.ContainsKey($index)) { [string]$valueMap[$index] } else { "" }
    $numericValue = 0d
    $hasNumericValue = [double]::TryParse($valueText, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$numericValue)
    if (Optional-Boolean $labelNode $labelsNode "showVal" $false -and -not [string]::IsNullOrEmpty($valueText)) {
        $parts.Add($valueText)
    }

    if (Optional-Boolean $labelNode $labelsNode "showPercent" $false -and $hasNumericValue) {
        $parts.Add((Format-ChartPercent $numericValue $total))
    }

    return ,$parts.ToArray()
}

function Hash-SetKey($hashes) {
    return (@($hashes | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Sort-Object) -join "+")
}

function Manual-Layout-Val($manualLayout, [string] $name) {
    if ($null -eq $manualLayout) {
        return ""
    }

    return Child-Val $manualLayout $name
}

function Read-ChartXmlDocument([string] $path, [string] $pptx, [string] $chartPart) {
    if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path -LiteralPath $path)) {
        return [xml](Get-Content -Raw -LiteralPath (Resolve-Path -LiteralPath $path).Path)
    }

    if ([string]::IsNullOrWhiteSpace($pptx) -or -not (Test-Path -LiteralPath $pptx)) {
        return $null
    }

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead((Resolve-Path -LiteralPath $pptx).Path)
    try {
        $entryName = $chartPart.TrimStart("/", "\")
        $entry = $zip.GetEntry($entryName)
        if ($null -eq $entry) {
            return $null
        }

        $stream = $entry.Open()
        try {
            $reader = [System.IO.StreamReader]::new($stream)
            try {
                return [xml]$reader.ReadToEnd()
            }
            finally {
                $reader.Dispose()
            }
        }
        finally {
            $stream.Dispose()
        }
    }
    finally {
        $zip.Dispose()
    }
}

function Read-ComDataLabelMetadataByIndex($metadata) {
    $byIndex = @{}
    if ($null -eq $metadata -or $null -eq $metadata.DataLabels) {
        return $byIndex
    }

    foreach ($label in @($metadata.DataLabels)) {
        $key = [string]$label.Index
        if (-not [string]::IsNullOrWhiteSpace($key)) {
            $byIndex[$key] = $label
        }
    }

    return $byIndex
}

function Read-ChartLabelManualLayouts([string] $path, [string] $pptx, [string] $chartPart, $comMetadata) {
    $xml = Read-ChartXmlDocument $path $pptx $chartPart
    if ($null -eq $xml) {
        return ,@()
    }

    $series = $xml.SelectSingleNode("//*[local-name()='pieChart']/*[local-name()='ser']")
    if ($null -eq $series) {
        $series = $xml.SelectSingleNode("//*[local-name()='ser']")
    }

    $categoryMap = if ($null -eq $series) { @{} } else { Read-IndexedTextCache $series "cat" }
    $valueMap = if ($null -eq $series) { @{} } else { Read-IndexedTextCache $series "val" }
    $total = 0d
    foreach ($value in $valueMap.Values) {
        $parsed = 0d
        if ([double]::TryParse([string]$value, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$parsed)) {
            $total += $parsed
        }
    }

    $labelsNode = if ($null -eq $series) { $null } else { $series.SelectSingleNode("*[local-name()='dLbls']") }
    $labels = if ($null -eq $labelsNode) { $xml.SelectNodes("//*[local-name()='dLbl']") } else { $labelsNode.SelectNodes("*[local-name()='dLbl']") }
    $comLabelsByIndex = Read-ComDataLabelMetadataByIndex $comMetadata
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($label in $labels) {
        $manualLayout = $label.SelectSingleNode("*[local-name()='layout']/*[local-name()='manualLayout']")
        if ($null -eq $manualLayout) {
            continue
        }

        $index = Child-Val $label "idx"
        $expectedParts = Read-ExpectedLabelParts $series $labelsNode $label $index $categoryMap $valueMap $total
        $expectedHashes = @($expectedParts | ForEach-Object { Stable-TextHash $_ })
        $comLabel = if ($comLabelsByIndex.ContainsKey([string]$index)) { $comLabelsByIndex[[string]$index] } else { $null }
        $rows.Add([pscustomobject]@{
            Index = $index
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
            ExpectedPartCount = $expectedParts.Count
            ExpectedTextHashSetKey = Hash-SetKey $expectedHashes
            ComRequestedLeft = if ($null -eq $comLabel) { $null } else { $comLabel.RequestedLeft }
            ComRequestedTop = if ($null -eq $comLabel) { $null } else { $comLabel.RequestedTop }
            ComObservedLeft = if ($null -eq $comLabel) { $null } else { $comLabel.ObservedLeft }
            ComObservedTop = if ($null -eq $comLabel) { $null } else { $comLabel.ObservedTop }
            ComObservedWidth = if ($null -eq $comLabel) { $null } else { $comLabel.ObservedWidth }
            ComObservedHeight = if ($null -eq $comLabel) { $null } else { $comLabel.ObservedHeight }
            ComAppliedPosition = if ($null -eq $comLabel) { $null } else { $comLabel.AppliedPosition }
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

function Build-ManualLayoutCoordinateEvidence($manualLayouts, $referenceClusters, $candidateClusters, $referencePlotBox, $candidatePlotBox) {
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($layout in $manualLayouts) {
        $quadrant = Manual-Quadrant $layout
        $referenceCluster = @($referenceClusters | Where-Object { (Quadrant $_ $referencePlotBox) -eq $quadrant } | Sort-Object -Property MinY, MinX | Select-Object -First 1)[0]
        $candidateCluster = @($candidateClusters | Where-Object { (Quadrant $_ $candidatePlotBox) -eq $quadrant } | Sort-Object -Property MinY, MinX | Select-Object -First 1)[0]
        $expectedHashSetKey = [string]$layout.ExpectedTextHashSetKey
        $referenceHashCluster = if ([string]::IsNullOrWhiteSpace($expectedHashSetKey)) {
            $null
        }
        else {
            @($referenceClusters | Where-Object { (Hash-SetKey ([string]$_.TextHash).Split("+")) -eq $expectedHashSetKey } | Select-Object -First 1)[0]
        }
        $candidateHashCluster = if ([string]::IsNullOrWhiteSpace($expectedHashSetKey)) {
            $null
        }
        else {
            @($candidateClusters | Where-Object { (Hash-SetKey ([string]$_.TextHash).Split("+")) -eq $expectedHashSetKey } | Select-Object -First 1)[0]
        }
        $rows.Add([pscustomobject]@{
            Index = [string]$layout.Index
            ManualQuadrant = $quadrant
            ManualX = [string]$layout.X
            ManualY = [string]$layout.Y
            ComRequestedLeft = $layout.ComRequestedLeft
            ComRequestedTop = $layout.ComRequestedTop
            ComObservedLeft = $layout.ComObservedLeft
            ComObservedTop = $layout.ComObservedTop
            ComObservedWidth = $layout.ComObservedWidth
            ComObservedHeight = $layout.ComObservedHeight
            ExpectedTextHashSetKey = $expectedHashSetKey
            ReferenceClusterBounds = Format-Bounds $referenceCluster
            ReferenceClusterRelativeX = Relative-Center $referenceCluster $referencePlotBox "x"
            ReferenceClusterRelativeY = Relative-Center $referenceCluster $referencePlotBox "y"
            CandidateClusterBounds = Format-Bounds $candidateCluster
            CandidateClusterRelativeX = Relative-Center $candidateCluster $candidatePlotBox "x"
            CandidateClusterRelativeY = Relative-Center $candidateCluster $candidatePlotBox "y"
            ReferenceHashClusterBounds = Format-Bounds $referenceHashCluster
            ReferenceHashClusterRelativeX = Relative-Center $referenceHashCluster $referencePlotBox "x"
            ReferenceHashClusterRelativeY = Relative-Center $referenceHashCluster $referencePlotBox "y"
            CandidateHashClusterBounds = Format-Bounds $candidateHashCluster
            CandidateHashClusterRelativeX = Relative-Center $candidateHashCluster $candidatePlotBox "x"
            CandidateHashClusterRelativeY = Relative-Center $candidateHashCluster $candidatePlotBox "y"
        })
    }

    return ,$rows.ToArray()
}

function Build-ManualLayoutHashIndexMap($manualLayouts) {
    $map = @{}
    foreach ($layout in $manualLayouts) {
        $key = [string]$layout.ExpectedTextHashSetKey
        if (-not [string]::IsNullOrWhiteSpace($key) -and -not $map.ContainsKey($key)) {
            $map[$key] = [string]$layout.Index
        }
    }

    return $map
}

function Find-NearestCluster($line, $clusters) {
    $best = $null
    $bestDistance = [double]::PositiveInfinity
    foreach ($cluster in $clusters) {
        $distance = BoundsDistance $line $cluster
        if ($null -eq $distance) {
            continue
        }

        if ([double]$distance -lt $bestDistance) {
            $bestDistance = [double]$distance
            $best = $cluster
        }
    }

    return $best
}

function Build-LeaderLineClusterEvidenceForSide([string] $side, $leaderLines, $clusters, $hashIndexMap) {
    $rows = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $leaderLines.Count; $i++) {
        $line = $leaderLines[$i]
        $cluster = Find-NearestCluster $line $clusters
        $hashSetKey = if ($null -eq $cluster) { "" } else { Hash-SetKey ([string]$cluster.TextHash).Split("+") }
        $rows.Add([pscustomobject]@{
            Side = $side
            LeaderLineIndex = $i
            LeaderLineBounds = Format-Bounds $line
            NearestClusterBounds = Format-Bounds $cluster
            NearestClusterDistance = BoundsDistance $line $cluster
            NearestClusterTextHashSetKey = $hashSetKey
            NearestManualIndex = if (-not [string]::IsNullOrWhiteSpace($hashSetKey) -and $hashIndexMap.ContainsKey($hashSetKey)) { $hashIndexMap[$hashSetKey] } else { "" }
        })
    }

    return ,$rows.ToArray()
}

function Build-LeaderLineClusterEvidence($referenceLeaderLines, $candidateLeaderLines, $referenceClusters, $candidateClusters, $manualLayouts) {
    $hashIndexMap = Build-ManualLayoutHashIndexMap $manualLayouts
    return ,@(
        Build-LeaderLineClusterEvidenceForSide "reference" $referenceLeaderLines $referenceClusters $hashIndexMap
        Build-LeaderLineClusterEvidenceForSide "candidate" $candidateLeaderLines $candidateClusters $hashIndexMap
    )
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
$referenceLabelClusters = Group-TextClusters $referenceLabels
$candidateLabelClusters = Group-TextClusters $candidateLabels
$referenceLeaderLines = Select-Kind $referenceGraphics "DataLabelLeaderLineCandidate"
$candidateLeaderLines = Select-Kind $candidateGraphics "DataLabelLeaderLineCandidate"
$ComMetadataJson = Resolve-ComMetadataJsonPath $ComMetadataJson $Pptx
$comMetadata = Read-JsonObject $ComMetadataJson
$chartManualLayouts = Read-ChartLabelManualLayouts $ChartXml $Pptx $ChartPart $comMetadata
$referencePolarPlotBox = Select-PrimaryKind $referenceGraphics "PolarPlotBoxCandidate"
$candidatePolarPlotBox = Select-PrimaryKind $candidateGraphics "PolarPlotBoxCandidate"

$labelMatches = Match-Nearest $referenceLabels $candidateLabels "DataLabelText"
$labelClusterMatches = Match-Nearest $referenceLabelClusters $candidateLabelClusters "DataLabelTextCluster"
$leaderLineMatches = Match-Nearest $referenceLeaderLines $candidateLeaderLines "DataLabelLeaderLineCandidate"
$manualLayoutCoordinateEvidence = Build-ManualLayoutCoordinateEvidence $chartManualLayouts $referenceLabelClusters $candidateLabelClusters $referencePolarPlotBox $candidatePolarPlotBox
$leaderLineClusterEvidence = Build-LeaderLineClusterEvidence $referenceLeaderLines $candidateLeaderLines $referenceLabelClusters $candidateLabelClusters $chartManualLayouts

$summary = [pscustomobject]@{
    DataLabelTextReferenceCount = $referenceLabels.Count
    DataLabelTextCandidateCount = $candidateLabels.Count
    DataLabelTextMaxNearestBoundsDelta = if ($labelMatches.Count -eq 0) { $null } else { ($labelMatches | Where-Object { $null -ne $_.BoundsDelta } | Measure-Object -Property BoundsDelta -Maximum).Maximum }
    DataLabelTextClusterReferenceCount = $referenceLabelClusters.Count
    DataLabelTextClusterCandidateCount = $candidateLabelClusters.Count
    DataLabelTextClusterMaxNearestBoundsDelta = if ($labelClusterMatches.Count -eq 0) { $null } else { ($labelClusterMatches | Where-Object { $null -ne $_.BoundsDelta } | Measure-Object -Property BoundsDelta -Maximum).Maximum }
    DataLabelLeaderLineReferenceCount = $referenceLeaderLines.Count
    DataLabelLeaderLineCandidateCount = $candidateLeaderLines.Count
    DataLabelLeaderLineMaxNearestBoundsDelta = if ($leaderLineMatches.Count -eq 0) { $null } else { ($leaderLineMatches | Where-Object { $null -ne $_.BoundsDelta } | Measure-Object -Property BoundsDelta -Maximum).Maximum }
    ChartManualLayoutCount = $chartManualLayouts.Count
    ChartComMetadataDataLabelCount = if ($null -eq $comMetadata -or $null -eq $comMetadata.DataLabels) { 0 } else { @($comMetadata.DataLabels).Count }
    ChartComMetadataShape = if ($null -eq $comMetadata) { $null } else { $comMetadata.ChartShape }
    ChartManualLayouts = $chartManualLayouts
    DataLabelManualLayoutCoordinateEvidence = $manualLayoutCoordinateEvidence
    DataLabelTextReferenceClusters = $referenceLabelClusters
    DataLabelTextCandidateClusters = $candidateLabelClusters
    DataLabelTextMatches = $labelMatches
    DataLabelTextClusterMatches = $labelClusterMatches
    DataLabelLeaderLineMatches = $leaderLineMatches
    DataLabelLeaderLineClusterEvidence = $leaderLineClusterEvidence
}

$summary | Format-List DataLabelTextReferenceCount, DataLabelTextCandidateCount, DataLabelTextMaxNearestBoundsDelta, DataLabelTextClusterReferenceCount, DataLabelTextClusterCandidateCount, DataLabelTextClusterMaxNearestBoundsDelta, DataLabelLeaderLineReferenceCount, DataLabelLeaderLineCandidateCount, DataLabelLeaderLineMaxNearestBoundsDelta, ChartManualLayoutCount, ChartComMetadataDataLabelCount

if ($chartManualLayouts.Count -gt 0) {
    Write-Host ""
    Write-Host "Chart data-label manual layouts:"
    $chartManualLayouts | Format-Table -AutoSize Index, Position, X, Y, Width, Height, XMode, YMode, WidthMode, HeightMode, ComRequestedLeft, ComRequestedTop, ComObservedLeft, ComObservedTop, ComObservedWidth, ComObservedHeight, ShowCategoryName, ShowPercent, ShowLeaderLines
}

if ($manualLayoutCoordinateEvidence.Count -gt 0) {
    Write-Host ""
    Write-Host "Data-label manual-layout coordinate evidence:"
    $manualLayoutCoordinateEvidence | Format-Table -AutoSize Index, ManualQuadrant, ManualX, ManualY, ComRequestedLeft, ComRequestedTop, ComObservedLeft, ComObservedTop, ReferenceClusterRelativeX, ReferenceClusterRelativeY, CandidateClusterRelativeX, CandidateClusterRelativeY, ReferenceHashClusterRelativeX, ReferenceHashClusterRelativeY, CandidateHashClusterRelativeX, CandidateHashClusterRelativeY
}

if ($labelMatches.Count -gt 0) {
    Write-Host ""
    Write-Host "Data-label text nearest matches:"
    $labelMatches | Format-Table -AutoSize Kind, ReferenceIndex, CandidateIndex, BoundsDelta, CenterDeltaX, CenterDeltaY, ReferenceBounds, CandidateBounds
}

if ($labelClusterMatches.Count -gt 0) {
    Write-Host ""
    Write-Host "Data-label text-cluster nearest matches:"
    $labelClusterMatches | Format-Table -AutoSize Kind, ReferenceIndex, CandidateIndex, BoundsDelta, CenterDeltaX, CenterDeltaY, ReferenceBounds, CandidateBounds
}

if ($leaderLineMatches.Count -gt 0) {
    Write-Host ""
    Write-Host "Data-label leader-line nearest matches:"
    $leaderLineMatches | Format-Table -AutoSize Kind, ReferenceIndex, CandidateIndex, BoundsDelta, CenterDeltaX, CenterDeltaY, ReferenceBounds, CandidateBounds
}

if ($leaderLineClusterEvidence.Count -gt 0) {
    Write-Host ""
    Write-Host "Data-label leader-line cluster evidence:"
    $leaderLineClusterEvidence | Format-Table -AutoSize Side, LeaderLineIndex, LeaderLineBounds, NearestClusterDistance, NearestManualIndex, NearestClusterBounds
}

if (-not [string]::IsNullOrWhiteSpace($OutputJson)) {
    $outputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputJson)
    $outputDirectory = Split-Path -Parent $outputPath
    if (-not [string]::IsNullOrWhiteSpace($outputDirectory)) {
        New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null
    }

    $summary | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $outputPath -Encoding UTF8
}
