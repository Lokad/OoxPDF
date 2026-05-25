param(
    [Parameter(Mandatory = $true)]
    [string] $InputPath,

    [string] $Output,

    [int] $PageNumber = 0,

    [double] $LineTolerance = 0.25,

    [double] $MinLineLength = 12,

    [double] $MarkerMaxSize = 20,

    [double] $GridlineBoundsTolerance = 1.0
)

$ErrorActionPreference = "Stop"

function Read-JsonArray($path) {
    $items = Get-Content -Raw -LiteralPath (Resolve-Path -LiteralPath $path).Path | ConvertFrom-Json
    if ($null -eq $items) {
        return ,@()
    }

    if ($items -is [array]) {
        return ,$items
    }

    return ,@($items)
}

function Width($op) { return [double]$op.MaxX - [double]$op.MinX }
function Height($op) { return [double]$op.MaxY - [double]$op.MinY }
function CenterX($op) { return ([double]$op.MinX + [double]$op.MaxX) / 2d }
function CenterY($op) { return ([double]$op.MinY + [double]$op.MaxY) / 2d }
function Round([double]$value) { return [Math]::Round($value, 6) }

function BoundsDistance($a, $b) {
    return [Math]::Abs([double]$a.MinX - [double]$b.MinX) +
        [Math]::Abs([double]$a.MinY - [double]$b.MinY) +
        [Math]::Abs([double]$a.MaxX - [double]$b.MaxX) +
        [Math]::Abs([double]$a.MaxY - [double]$b.MaxY)
}

function New-Structure($kind, $op) {
    [pscustomobject]@{
        Kind = $kind
        PageNumber = $op.PageNumber
        SourceKind = $op.Kind
        SourceOperator = $op.Operator
        SegmentCount = $op.SegmentCount
        MinX = [double]$op.MinX
        MinY = [double]$op.MinY
        MaxX = [double]$op.MaxX
        MaxY = [double]$op.MaxY
        Width = Round (Width $op)
        Height = Round (Height $op)
        CenterX = Round (CenterX $op)
        CenterY = Round (CenterY $op)
        LineWidth = if ($op.Kind -eq "Fill") { 0d } else { [double]$op.LineWidth }
        StrokeColor = $op.StrokeColor
        FillColor = $op.FillColor
        Dash = $op.Dash
        LineCap = $op.LineCap
        LineJoin = $op.LineJoin
    }
}

function New-DerivedStructure($kind, $pageNumber, $sourceOperator, $segmentCount, $minX, $minY, $maxX, $maxY) {
    [pscustomobject]@{
        Kind = $kind
        PageNumber = $pageNumber
        SourceKind = "Derived"
        SourceOperator = $sourceOperator
        SegmentCount = $segmentCount
        MinX = Round $minX
        MinY = Round $minY
        MaxX = Round $maxX
        MaxY = Round $maxY
        Width = Round ($maxX - $minX)
        Height = Round ($maxY - $minY)
        CenterX = Round (($minX + $maxX) / 2d)
        CenterY = Round (($minY + $maxY) / 2d)
        LineWidth = 0d
        StrokeColor = ""
        FillColor = ""
        Dash = ""
        LineCap = 0
        LineJoin = 0
    }
}

function Copy-StructureAsKind($kind, $structure) {
    [pscustomobject]@{
        Kind = $kind
        PageNumber = $structure.PageNumber
        SourceKind = $structure.SourceKind
        SourceOperator = $structure.SourceOperator
        SegmentCount = $structure.SegmentCount
        MinX = [double]$structure.MinX
        MinY = [double]$structure.MinY
        MaxX = [double]$structure.MaxX
        MaxY = [double]$structure.MaxY
        Width = Round (Width $structure)
        Height = Round (Height $structure)
        CenterX = Round (CenterX $structure)
        CenterY = Round (CenterY $structure)
        LineWidth = [double]$structure.LineWidth
        StrokeColor = $structure.StrokeColor
        FillColor = $structure.FillColor
        Dash = $structure.Dash
        LineCap = $structure.LineCap
        LineJoin = $structure.LineJoin
    }
}

function Get-UnionBounds($items) {
    if ($items.Count -eq 0) {
        return $null
    }

    $minX = ($items | Measure-Object -Property MinX -Minimum).Minimum
    $maxX = ($items | Measure-Object -Property MaxX -Maximum).Maximum
    $minY = ($items | Measure-Object -Property MinY -Minimum).Minimum
    $maxY = ($items | Measure-Object -Property MaxY -Maximum).Maximum
    [pscustomobject]@{
        MinX = [double]$minX
        MinY = [double]$minY
        MaxX = [double]$maxX
        MaxY = [double]$maxY
    }
}

function Is-Near([double]$left, [double]$right, [double]$tolerance) {
    return [Math]::Abs($left - $right) -le $tolerance
}

function Is-InsideOpenInterval([double]$value, [double]$min, [double]$max, [double]$tolerance) {
    return $value -gt ($min + $tolerance) -and $value -lt ($max - $tolerance)
}

function Is-OutsideBox($structure, $box, [double]$tolerance) {
    return ([double]$structure.CenterX -lt ([double]$box.MinX - $tolerance)) -or
        ([double]$structure.CenterX -gt ([double]$box.MaxX + $tolerance)) -or
        ([double]$structure.CenterY -lt ([double]$box.MinY - $tolerance)) -or
        ([double]$structure.CenterY -gt ([double]$box.MaxY + $tolerance))
}

function Is-LegendSwatchShape($structure) {
    $kind = [string]$structure.Kind
    $width = Width $structure
    $height = Height $structure
    if (($kind -eq "MarkerCandidate" -or $kind -eq "StrokeMarkerCandidate") -and $width -le 24d -and $height -le 24d) {
        return $true
    }

    if ($kind -eq "HorizontalLine" -and $width -le 100d -and [double]$structure.LineWidth -ge 1d) {
        return $true
    }

    if ($kind -eq "FilledRegion" -and $width -le 40d -and $height -le 40d) {
        return $true
    }

    return $false
}

$ops = Read-JsonArray $InputPath
if ($PageNumber -gt 0) {
    $ops = @($ops | Where-Object { [int]$_.PageNumber -eq $PageNumber })
}

$structures = New-Object System.Collections.Generic.List[object]

foreach ($op in $ops) {
    $width = Width $op
    $height = Height $op
    if ($op.Kind -eq "Clip" -and $width -gt 0 -and $height -gt 0) {
        $structures.Add((New-Structure "ClipBox" $op))
    }

    if ($op.Kind -eq "Stroke" -and $height -le $LineTolerance -and $width -ge $MinLineLength) {
        $structures.Add((New-Structure "HorizontalLine" $op))
    }
    elseif ($op.Kind -eq "Stroke" -and $width -le $LineTolerance -and $height -ge $MinLineLength) {
        $structures.Add((New-Structure "VerticalLine" $op))
    }
    elseif ($op.Kind -eq "Stroke" -and $width -gt $LineTolerance -and $height -gt $LineTolerance -and
        $width -le $MarkerMaxSize -and $height -le $MarkerMaxSize) {
        $structures.Add((New-Structure "StrokeMarkerCandidate" $op))
    }

    if (($op.Kind -eq "Fill" -or $op.Kind -eq "FillStroke") -and $width -gt 0 -and $height -gt 0) {
        if ($width -le $MarkerMaxSize -and $height -le $MarkerMaxSize) {
            $structures.Add((New-Structure "MarkerCandidate" $op))
        }
        else {
            $structures.Add((New-Structure "FilledRegion" $op))
        }
    }
}

$horizontalLines = @($structures | Where-Object { $_.Kind -eq "HorizontalLine" })
if ($horizontalLines.Count -ge 2) {
    $bounds = Get-UnionBounds $horizontalLines
    $page = if ($PageNumber -gt 0) { $PageNumber } else { $horizontalLines[0].PageNumber }
    $structures.Add((New-DerivedStructure "PlotBoxCandidate" $page "HorizontalLineBounds" $horizontalLines.Count $bounds.MinX $bounds.MinY $bounds.MaxX $bounds.MaxY))
}

foreach ($op in $ops) {
    $segmentCount = if ($null -ne $op.SegmentCount) { [int]$op.SegmentCount } else { 0 }
    if ($op.Kind -ne "Stroke" -or $segmentCount -lt 2) {
        continue
    }

    $width = Width $op
    $height = Height $op
    if ($width -lt $MinLineLength -or $height -le $LineTolerance) {
        continue
    }

    $matchingBaseline = @($horizontalLines | Where-Object {
        (Is-Near ([double]$_.MinX) ([double]$op.MinX) $GridlineBoundsTolerance) -and
        (Is-Near ([double]$_.MaxX) ([double]$op.MaxX) $GridlineBoundsTolerance) -and
        ([double]$_.CenterY -lt [double]$op.MinY)
    } | Sort-Object -Property CenterY -Descending | Select-Object -First 1)
    if ($matchingBaseline.Count -eq 0) {
        continue
    }

    $baseline = $matchingBaseline[0]
    $page = if ($PageNumber -gt 0) { $PageNumber } else { $op.PageNumber }
    $structures.Add((New-DerivedStructure "GridlineAxisPlotBoxCandidate" $page "HorizontalGridlineAndAxisBounds" ($segmentCount + 1) $op.MinX $baseline.MinY $op.MaxX $op.MaxY))
}

$verticalLines = @($structures | Where-Object { $_.Kind -eq "VerticalLine" })
if ($horizontalLines.Count -ge 1 -and $verticalLines.Count -ge 1) {
    $leftAxis = @($verticalLines | Sort-Object -Property MinX | Select-Object -First 1)[0]
    $topAxis = @($horizontalLines | Sort-Object -Property MinY | Select-Object -First 1)[0]
    $minX = [double]$leftAxis.MinX
    $minY = [double]$topAxis.MinY
    $maxX = [double]$topAxis.MaxX
    $maxY = [double]$leftAxis.MaxY
    if ($maxX -gt $minX -and $maxY -gt $minY) {
        $page = if ($PageNumber -gt 0) { $PageNumber } else { $leftAxis.PageNumber }
        $structures.Add((New-DerivedStructure "AxisPairPlotBoxCandidate" $page "AxisLinePairBounds" 2 $minX $minY $maxX $maxY))
    }
    elseif ((Is-Near ([double]$topAxis.MinY) ([double]$leftAxis.MaxY) $GridlineBoundsTolerance) -and
        $maxX -gt $minX -and [double]$leftAxis.MaxY -gt [double]$leftAxis.MinY) {
        $page = if ($PageNumber -gt 0) { $PageNumber } else { $leftAxis.PageNumber }
        $structures.Add((New-DerivedStructure "AxisPairPlotBoxCandidate" $page "AxisLinePairBounds" 2 $minX $leftAxis.MinY $maxX $topAxis.MinY))
    }
}

$axisPairPlotBox = @($structures | Where-Object { $_.Kind -eq "AxisPairPlotBoxCandidate" } | Select-Object -First 1)
$gridlineAxisPlotBox = @($structures | Where-Object { $_.Kind -eq "GridlineAxisPlotBoxCandidate" } | Select-Object -First 1)
$plotBoxForGridlines = if ($gridlineAxisPlotBox.Count -gt 0) {
    $gridlineAxisPlotBox[0]
}
elseif ($axisPairPlotBox.Count -gt 0) {
    $axisPairPlotBox[0]
}
else {
    @($structures | Where-Object { $_.Kind -eq "PlotBoxCandidate" } | Select-Object -First 1)[0]
}

if ($null -ne $plotBoxForGridlines) {
    foreach ($op in $ops) {
        $segmentCount = if ($null -ne $op.SegmentCount) { [int]$op.SegmentCount } else { 0 }
        if ($op.Kind -ne "Stroke" -or $segmentCount -lt 2) {
            continue
        }

        $width = Width $op
        $height = Height $op
        $spansPlotWidth = (Is-Near ([double]$op.MinX) ([double]$plotBoxForGridlines.MinX) $GridlineBoundsTolerance) -and
            (Is-Near ([double]$op.MaxX) ([double]$plotBoxForGridlines.MaxX) $GridlineBoundsTolerance)
        $insideVerticalRange = ([double]$op.MinY -gt ([double]$plotBoxForGridlines.MinY + $GridlineBoundsTolerance)) -and
            ([double]$op.MaxY -le ([double]$plotBoxForGridlines.MaxY + $GridlineBoundsTolerance))
        if ($width -ge $MinLineLength -and $height -gt $LineTolerance -and $spansPlotWidth -and $insideVerticalRange) {
            $structures.Add((New-Structure "HorizontalGridlineGroupCandidate" $op))
            continue
        }

        $spansPlotHeight = (Is-Near ([double]$op.MinY) ([double]$plotBoxForGridlines.MinY) $GridlineBoundsTolerance) -and
            (Is-Near ([double]$op.MaxY) ([double]$plotBoxForGridlines.MaxY) $GridlineBoundsTolerance)
        $insideHorizontalRange = ([double]$op.MinX -gt ([double]$plotBoxForGridlines.MinX + $GridlineBoundsTolerance)) -and
            ([double]$op.MaxX -le ([double]$plotBoxForGridlines.MaxX + $GridlineBoundsTolerance))
        if ($height -ge $MinLineLength -and $width -gt $LineTolerance -and $spansPlotHeight -and $insideHorizontalRange) {
            $structures.Add((New-Structure "VerticalGridlineGroupCandidate" $op))
        }
    }

    foreach ($line in $horizontalLines) {
        $spansPlotWidth = (Is-Near ([double]$line.MinX) ([double]$plotBoxForGridlines.MinX) $GridlineBoundsTolerance) -and
            (Is-Near ([double]$line.MaxX) ([double]$plotBoxForGridlines.MaxX) $GridlineBoundsTolerance)
        $isInsidePlot = Is-InsideOpenInterval ([double]$line.CenterY) ([double]$plotBoxForGridlines.MinY) ([double]$plotBoxForGridlines.MaxY) $GridlineBoundsTolerance
        if ($spansPlotWidth -and $isInsidePlot) {
            $structures.Add((Copy-StructureAsKind "HorizontalGridlineCandidate" $line))
        }
    }

    foreach ($line in $verticalLines) {
        $spansPlotHeight = (Is-Near ([double]$line.MinY) ([double]$plotBoxForGridlines.MinY) $GridlineBoundsTolerance) -and
            (Is-Near ([double]$line.MaxY) ([double]$plotBoxForGridlines.MaxY) $GridlineBoundsTolerance)
        $isInsidePlot = Is-InsideOpenInterval ([double]$line.CenterX) ([double]$plotBoxForGridlines.MinX) ([double]$plotBoxForGridlines.MaxX) $GridlineBoundsTolerance
        if ($spansPlotHeight -and $isInsidePlot) {
            $structures.Add((Copy-StructureAsKind "VerticalGridlineCandidate" $line))
        }
    }
}

if ($null -ne $plotBoxForGridlines) {
    $structureSnapshot = [object[]]$structures.ToArray()
    foreach ($structure in $structureSnapshot) {
        if ((Is-LegendSwatchShape $structure) -and (Is-OutsideBox $structure $plotBoxForGridlines $GridlineBoundsTolerance)) {
            $structures.Add((Copy-StructureAsKind "LegendSwatchCandidate" $structure))
        }
    }
}

$clipBoxes = @($structures | Where-Object { $_.Kind -eq "ClipBox" })
if ($clipBoxes.Count -gt 0) {
    $largestWidth = ($clipBoxes | Measure-Object -Property Width -Maximum).Maximum
    $largestHeight = ($clipBoxes | Measure-Object -Property Height -Maximum).Maximum
    $nonPageClipBoxes = @($clipBoxes | Where-Object {
        $_.Width -gt 40 -and $_.Height -gt 40 -and
        -not ([Math]::Abs([double]$_.Width - [double]$largestWidth) -le 0.01 -and [Math]::Abs([double]$_.Height - [double]$largestHeight) -le 0.01)
    })
    if ($nonPageClipBoxes.Count -gt 0) {
        $clipGroups = @($nonPageClipBoxes |
            Group-Object -Property MinX, MinY, MaxX, MaxY)
        $dominantGroup = $null
        if ($clipGroups.Count -gt 0 -and $null -ne $plotBoxForGridlines) {
            $dominantGroup = $clipGroups |
                Sort-Object -Property @{
                    Expression = {
                        $clip = @($_.Group)[0]
                        BoundsDistance $clip $plotBoxForGridlines
                    }
                }, @{
                    Expression = { -[int]$_.Count }
                } |
                Select-Object -First 1

            $nearestClip = @($dominantGroup.Group)[0]
            if ((BoundsDistance $nearestClip $plotBoxForGridlines) -gt 10d) {
                $dominantGroup = $null
            }
        }

        if ($null -eq $dominantGroup -and $null -eq $plotBoxForGridlines) {
            $dominantGroup = $nonPageClipBoxes |
                Group-Object -Property MinX, MinY, MaxX, MaxY |
                Sort-Object -Property Count -Descending |
                Select-Object -First 1
        }

        if ($null -ne $dominantGroup) {
            $clip = @($dominantGroup.Group)[0]
            $structures.Add((New-DerivedStructure "PlotAreaClipBoxCandidate" $clip.PageNumber "DominantClipBox" $dominantGroup.Count $clip.MinX $clip.MinY $clip.MaxX $clip.MaxY))
        }
    }
}

$polarRegions = @($structures | Where-Object {
    $_.Kind -eq "FilledRegion" -and $_.Width -gt $MarkerMaxSize -and $_.Height -gt $MarkerMaxSize
})
if ($polarRegions.Count -gt 1) {
    $largestWidth = ($polarRegions | Measure-Object -Property Width -Maximum).Maximum
    $largestHeight = ($polarRegions | Measure-Object -Property Height -Maximum).Maximum
    $nonPageRegions = @($polarRegions | Where-Object {
        -not ([Math]::Abs([double]$_.Width - [double]$largestWidth) -le 0.01 -and [Math]::Abs([double]$_.Height - [double]$largestHeight) -le 0.01)
    })
    if ($nonPageRegions.Count -ge 2) {
        $bounds = Get-UnionBounds $nonPageRegions
        $width = [double]$bounds.MaxX - [double]$bounds.MinX
        $height = [double]$bounds.MaxY - [double]$bounds.MinY
        $aspect = $width / [Math]::Max(1d, $height)
        if ($aspect -ge 0.74d -and $aspect -le 1.35d) {
            $page = if ($PageNumber -gt 0) { $PageNumber } else { $nonPageRegions[0].PageNumber }
            $structures.Add((New-DerivedStructure "PolarPlotBoxCandidate" $page "FilledRegionUnion" $nonPageRegions.Count $bounds.MinX $bounds.MinY $bounds.MaxX $bounds.MaxY))
        }
    }
}

$ordered = @($structures | Sort-Object -Property PageNumber, Kind, MinY, MinX, MaxY, MaxX)
$ordered | Format-Table -AutoSize
Write-Host "Chart graphics structures: $($ordered.Count)"

if (-not [string]::IsNullOrWhiteSpace($Output)) {
    $outputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Output)
    $outputDirectory = Split-Path -Parent $outputPath
    if (-not [string]::IsNullOrWhiteSpace($outputDirectory)) {
        New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null
    }

    $ordered | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $outputPath -Encoding UTF8
}
