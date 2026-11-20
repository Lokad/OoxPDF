param(
    [Parameter(Mandatory = $true)]
    [string] $InputPath,

    [string] $ChartStructures,

    [string] $Output,

    [int] $PageNumber = 0,

    [double] $PlotTolerance = 8
)

$ErrorActionPreference = "Stop"

$ChartTitleCenterToleranceFactor = 0.45d

. (Join-Path $PSScriptRoot "JsonArray.ps1")
. (Join-Path $PSScriptRoot "CompareMath.ps1")

function Round([double]$value) { return [Math]::Round($value, 6) }

function TextValue($op) {
    if ($op.DecodedText -ne $null) {
        return [string]$op.DecodedText
    }

    return [string]$op.Payload
}

function Is-NumericText([string]$text) {
    $trimmed = $text.Trim()
    return $trimmed -match '^[+-]?(?:\d+(?:[.,]\d+)?|[.,]\d+)%?$'
}

function Contains-NumericLabelSignal([string]$text) {
    return $text -match '\d' -or $text.Contains("%")
}

function TextX($op) {
    if ($op.EffectiveX -ne $null) {
        return [double]$op.EffectiveX
    }

    return [double]$op.X
}

function TextY($op) {
    if ($op.EffectiveY -ne $null) {
        return [double]$op.EffectiveY
    }

    return [double]$op.Y
}

function StructureWidth($structure) { return [double]$structure.MaxX - [double]$structure.MinX }
function StructureHeight($structure) { return [double]$structure.MaxY - [double]$structure.MinY }
function StructureCenterY($structure) { return ([double]$structure.MinY + [double]$structure.MaxY) / 2d }

function StableTextHash([string]$text) {
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($text)
    $hash = [System.Security.Cryptography.SHA256]::HashData($bytes)
    return [Convert]::ToHexString($hash).Substring(0, 16)
}

function Find-NearestChart([double]$x, [double]$y, $charts) {
    $best = $null
    $bestScore = [double]::PositiveInfinity
    $ambiguous = $false
    foreach ($chart in $charts) {
        $score = Get-PointToBoxDistance $x $y $chart.PlotBox
        if ($score -lt $bestScore) {
            $bestScore = $score
            $best = $chart
            $ambiguous = $false
        }
        elseif ($score -eq $bestScore -and $best.PlotBox.RegionIndex -ne $chart.PlotBox.RegionIndex) {
            $ambiguous = $true
        }
    }

    if (-not $ambiguous) { return $best }
}

function Has-LegendSwatch($op, $structures, [bool]$isPolarChart = $false) {
    $x = TextX $op
    $y = TextY $op
    $fontSize = if ($op.FontSize -ne $null) { [double]$op.FontSize } else { 0d }
    $verticalScale = if ($isPolarChart) { 0.85d } else { 0.75d }
    $verticalTolerance = [Math]::Max(8d, $fontSize * $verticalScale)
    $horizontalTolerance = [Math]::Max(36d, $fontSize * 4d)

    foreach ($structure in $structures) {
        $kind = [string]$structure.Kind
        $width = StructureWidth $structure
        $height = StructureHeight $structure
        $isMarker = $kind -eq "MarkerCandidate" -and $width -le 24d -and $height -le 24d
        $isShortLine = $kind -eq "HorizontalLine" -and $width -le 100d
        $isFilledSwatch = $kind -eq "FilledRegion" -and $width -le 40d -and $height -le 40d
        if (-not ($isMarker -or $isShortLine -or $isFilledSwatch)) {
            continue
        }

        $gap = $x - [double]$structure.MaxX
        if ($gap -lt 0d -or $gap -gt $horizontalTolerance) {
            continue
        }

        if ([Math]::Abs((StructureCenterY $structure) - $y) -le $verticalTolerance) {
            return $true
        }
    }

    return $false
}

function Has-LegendContainer($op, $structures, $plotBox, [double]$tolerance) {
    if ($null -eq $plotBox) {
        return $false
    }

    $x = TextX $op
    $y = TextY $op
    $plotRight = [double]$plotBox.MaxX

    if ($x -lt ($plotRight - $tolerance)) {
        return $false
    }

    foreach ($structure in $structures) {
        if ([string]$structure.Kind -ne "ClipBox") {
            continue
        }

        $width = StructureWidth $structure
        $height = StructureHeight $structure
        if ($width -lt 40d -or $height -lt 12d -or $width -gt 260d -or $height -gt 160d) {
            continue
        }

        if ($x -ge ([double]$structure.MinX - $tolerance) -and
            $x -le ([double]$structure.MaxX + $tolerance) -and
            $y -ge ([double]$structure.MinY - $tolerance) -and
            $y -le ([double]$structure.MaxY + $tolerance)) {
            return $true
        }
    }

    return $false
}

function Has-NearbyNumericLabel($op, $numericLabels) {
    $x = TextX $op
    $y = TextY $op
    foreach ($label in $numericLabels) {
        $labelX = TextX $label
        $labelY = TextY $label
        if ([Math]::Abs($x - $labelX) -le 40d -and [Math]::Abs($y - $labelY) -le 32d) {
            return $true
        }
    }

    return $false
}

function Looks-LikeChartTitle($op, $plotBox) {
    if ($null -eq $plotBox) {
        return $false
    }

    $text = TextValue $op
    $fontSize = if ($op.FontSize -ne $null) { [double]$op.FontSize } else { 0d }
    $x = TextX $op
    $plotCenter = ([double]$plotBox.MinX + [double]$plotBox.MaxX) / 2d
    $plotBoxWidth = StructureWidth $plotBox
    $plotWidth = [Math]::Max(1d, $plotBoxWidth)

    return $text.Length -gt 6 -and
        $fontSize -ge 10d -and
        [Math]::Abs($x - $plotCenter) -le ($plotWidth * $ChartTitleCenterToleranceFactor)
}

function Looks-LikeAbovePlotChartTitle($op, $plotBox, [double]$tolerance) {
    if ($null -eq $plotBox) {
        return $false
    }

    $x = TextX $op
    $y = TextY $op
    $minX = [double]$plotBox.MinX
    $maxX = [double]$plotBox.MaxX
    $maxY = [double]$plotBox.MaxY
    $plotBoxWidth = StructureWidth $plotBox
    $plotWidth = [Math]::Max(1d, $plotBoxWidth)
    $fontSize = if ($op.FontSize -ne $null) { [double]$op.FontSize } else { 0d }
    if ($fontSize -lt 14d) {
        return $false
    }

    if ($y -le ($maxY + $tolerance)) {
        return $false
    }

    if (Looks-LikeChartTitle $op $plotBox) {
        return $true
    }

    $text = TextValue $op
    return $text.Length -gt 6 -and
        $x -ge ($minX - $plotWidth * 0.25d) -and
        $x -le ($maxX + $plotWidth * 0.25d)
}

function Is-RotatedText($op) {
    $b = if ($op.EffectiveB -ne $null) { [double]$op.EffectiveB } elseif ($op.B -ne $null) { [double]$op.B } else { 0d }
    $c = if ($op.EffectiveC -ne $null) { [double]$op.EffectiveC } elseif ($op.C -ne $null) { [double]$op.C } else { 0d }
    return [Math]::Abs($b) -ge 0.5d -or [Math]::Abs($c) -ge 0.5d
}

function Looks-LikeAxisTitle($op, $plotBox, [double]$tolerance, [bool]$horizontalValueAxisChart) {
    if ($null -eq $plotBox) {
        return $false
    }

    $text = TextValue $op
    $fontSize = if ($op.FontSize -ne $null) { [double]$op.FontSize } else { 0d }
    if ($text.Trim().Length -le 3 -or $fontSize -lt 10.5d) {
        return $false
    }

    $x = TextX $op
    $y = TextY $op
    $minX = [double]$plotBox.MinX
    $minY = [double]$plotBox.MinY
    $maxX = [double]$plotBox.MaxX
    $maxY = [double]$plotBox.MaxY
    $insideX = $x -ge ($minX - $tolerance) -and $x -le ($maxX + $tolerance)
    $insideY = $y -ge ($minY - $tolerance) -and $y -le ($maxY + $tolerance)
    $axisBandY = $y -ge ($minY - ($tolerance * 3d)) -and $y -le ($maxY + ($tolerance * 3d))
    $axisBandX = $x -ge ($minX - ($tolerance * 3d)) -and $x -le ($maxX + ($tolerance * 3d))
    $farBelowOrAbove = $y -lt ($minY - ($tolerance * 2.5d)) -or $y -gt ($maxY + ($tolerance * 2.5d))
    $farLeftOrRight = $x -lt ($minX - ($tolerance * 2d)) -or $x -gt ($maxX + ($tolerance * 2d))

    if (Is-RotatedText $op) {
        $nearSideEdge = $x -le ($minX + $tolerance) -or $x -ge ($maxX - $tolerance)
        return ($farLeftOrRight -or $nearSideEdge) -and $axisBandY
    }

    if ($horizontalValueAxisChart) {
        return ($insideY -and $farLeftOrRight) -or ($insideX -and $farBelowOrAbove)
    }

    $nearTopOrBottomEdge = $y -le ($minY + $tolerance) -or $y -ge ($maxY - $tolerance)
    return ($insideX -and ($farBelowOrAbove -or $nearTopOrBottomEdge)) -or ($insideY -and $farLeftOrRight) -or
        ($axisBandX -and $farBelowOrAbove)
}

function Find-RadarSpokeGeometry($structures) {
    $spoke = @($structures | Where-Object {
        $_.Kind -eq "RadarSpokeGroupCandidate" -and
        $_.PathCenterX -ne $null -and
        $_.PathCenterY -ne $null -and
        $_.PathRadius -ne $null
    } | Sort-Object -Property PageNumber, MinY, MinX | Select-Object -First 1)
    if ($spoke.Count -eq 0) {
        return $null
    }

    return $spoke[0]
}

function Is-PolarChart($plotBox, $structures) {
    if ($null -ne $plotBox -and [string]$plotBox.Kind -eq "PolarPlotBoxCandidate") {
        return $true
    }

    foreach ($structure in $structures) {
        if ([string]$structure.Kind -eq "PolarSliceCandidate") {
            return $true
        }
    }

    return $false
}

function Is-HorizontalValueAxisChart($structures) {
    foreach ($structure in $structures) {
        if ([string]$structure.Kind -eq "VerticalGridlineGroupCandidate") {
            return $true
        }
    }

    return $false
}

function Classify-RadarText($op, $radarGeometry, [double]$tolerance) {
    if ($null -eq $radarGeometry) {
        return $null
    }

    $x = TextX $op
    $y = TextY $op
    $centerX = [double]$radarGeometry.PathCenterX
    $centerY = [double]$radarGeometry.PathCenterY
    $radius = [double]$radarGeometry.PathRadius
    if ($radius -le 0d) {
        return $null
    }

    $dx = $x - $centerX
    $dy = $y - $centerY
    $distance = [Math]::Sqrt($dx * $dx + $dy * $dy)
    $text = TextValue $op
    $isNumeric = Is-NumericText $text
    $fontSize = if ($op.FontSize -ne $null) { [double]$op.FontSize } else { 0d }
    if ((-not $isNumeric) -and
        $fontSize -ge ([Math]::Max(14d, $radius * 0.12d)) -and
        [Math]::Abs($dx) -le ($radius * 0.35d) -and
        $y -ge ($centerY + $radius + $tolerance * 2d)) {
        return "ChartTitleText"
    }

    if ($isNumeric -and
        $x -le ($centerX - $tolerance) -and
        $distance -le ($radius + $tolerance) -and
        $y -ge ($centerY - $tolerance * 2d)) {
        return "ValueAxisTickLabel"
    }

    if ((-not $isNumeric) -and $distance -ge ($radius * 0.8d)) {
        return "CategoryAxisTickLabel"
    }

    if ($distance -ge ($radius + $tolerance * 0.5d)) {
        return "CategoryAxisTickLabel"
    }

    return $null
}

function Classify-Text($op, $chart, [double]$tolerance) {
    if ($null -eq $chart) {
        return "ChartText"
    }

    $plotBox = $chart.PlotBox
    $structures = $chart.Structures
    $x = TextX $op
    $y = TextY $op
    $minX = [double]$plotBox.MinX
    $minY = [double]$plotBox.MinY
    $maxX = [double]$plotBox.MaxX
    $maxY = [double]$plotBox.MaxY
    $insideX = $x -ge ($minX - $tolerance) -and $x -le ($maxX + $tolerance)
    $insideY = $y -ge ($minY - $tolerance) -and $y -le ($maxY + $tolerance)
    $axisLabelY = $y -ge ($minY - ($tolerance * 2d)) -and $y -le ($maxY + ($tolerance * 2d))

    if (Looks-LikeAbovePlotChartTitle $op $plotBox $tolerance) {
        return "ChartTitleText"
    }

    $radarKind = Classify-RadarText $op $chart.RadarGeometry $tolerance
    if ($null -ne $radarKind) {
        return $radarKind
    }

    $isPolarChart = $chart.IsPolar
    $hasLegendSwatch = ($isPolarChart -or -not ($insideX -and $insideY)) -and
        (Has-LegendSwatch $op $structures $isPolarChart)
    if ($hasLegendSwatch) {
        if ($isPolarChart -and
            (Contains-NumericLabelSignal (TextValue $op) -or Has-NearbyNumericLabel $op $chart.NumericLabels)) {
            return "DataLabelText"
        }

        return "LegendText"
    }

    if ($isPolarChart) {
        return "DataLabelText"
    }

    $horizontalValueAxisChart = $chart.IsHorizontalValueAxis
    if (Looks-LikeAxisTitle $op $plotBox $tolerance $horizontalValueAxisChart) {
        return "AxisTitleText"
    }

    if ($insideX -and $insideY) {
        return "DataLabelText"
    }

    if ($horizontalValueAxisChart) {
        if ($x -lt ($minX - $tolerance) -and $axisLabelY) {
            return "CategoryAxisTickLabel"
        }

        if ($x -gt ($maxX + $tolerance) -and $axisLabelY) {
            if (Has-LegendContainer $op $structures $plotBox $tolerance) {
                return "LegendText"
            }

            return "RightSideText"
        }

        if ($insideX -and ($y -lt ($minY - $tolerance) -or $y -gt ($maxY + $tolerance))) {
            return "ValueAxisTickLabel"
        }
    }
    else {
        if ($x -lt ($minX - $tolerance) -and $axisLabelY) {
            return "ValueAxisTickLabel"
        }

        if ($x -gt ($maxX + $tolerance) -and $axisLabelY) {
            if (Has-LegendContainer $op $structures $plotBox $tolerance) {
                return "LegendText"
            }

            return "RightSideText"
        }

        if ($insideX -and $y -lt ($minY - $tolerance)) {
            return "CategoryAxisTickLabel"
        }

        if ($insideX -and $y -gt ($maxY + $tolerance)) {
            if (Looks-LikeChartTitle $op $plotBox) {
                return "ChartTitleText"
            }

            return "CategoryAxisTickLabel"
        }
    }

    return "OuterChartText"
}

$allTextOps = Read-JsonArray $InputPath
if ($PageNumber -gt 0) {
    $allTextOps = @($allTextOps | Where-Object { [int]$_.PageNumber -eq $PageNumber })
}

$allStructures = Read-JsonArrayIfExists $ChartStructures
if ($PageNumber -gt 0) {
    $allStructures = @($allStructures | Where-Object { [int]$_.PageNumber -eq $PageNumber })
}

$classified = New-Object System.Collections.Generic.List[object]
foreach ($pageGroup in @($allTextOps | Group-Object -Property PageNumber)) {
    $textOps = @($pageGroup.Group)
    $pageNumberValue = $textOps[0].PageNumber
    $structures = @($allStructures | Where-Object { $_.PageNumber -eq $pageNumberValue })
    $charts = @(foreach ($region in ($structures | Group-Object -Property RegionIndex)) {
        $plotBoxes = @(Find-PreferredChartPlotBoxes $region.Group -SortProperties PageNumber, MinY, MinX)
        if ($plotBoxes.Count -eq 0) { continue }
        if ($null -ne $plotBoxes[0].RegionIndex) {
            # Bounds variants for one known region describe the same chart.
            $plotBoxes = @($plotBoxes[0])
        }
        $radarGeometry = Find-RadarSpokeGeometry $region.Group
        $isHorizontalValueAxis = Is-HorizontalValueAxisChart $region.Group
        foreach ($plotBox in $plotBoxes) {
            [pscustomobject]@{
                PlotBox = $plotBox
                Structures = $region.Group
                RadarGeometry = $radarGeometry
                IsHorizontalValueAxis = $isHorizontalValueAxis
                IsPolar = Is-PolarChart $plotBox $region.Group
                NumericLabels = [System.Collections.Generic.List[object]]::new()
            }
        }
    })
    $indexedCharts = @($charts | Where-Object { $null -ne $_.PlotBox.RegionIndex })
    if ($indexedCharts.Count -gt 0) { $charts = $indexedCharts }

    $labels = @(foreach ($op in $textOps) {
        $text = TextValue $op
        if ([string]::IsNullOrWhiteSpace($text)) { continue }

        $x = TextX $op
        $y = TextY $op
        $chart = Find-NearestChart $x $y $charts
        if ($chart.IsPolar -and (Is-NumericText $text)) { $chart.NumericLabels.Add($op) }
        [pscustomobject]@{ Operation = $op; Text = $text; X = $x; Y = $y; Chart = $chart }
    })
    foreach ($label in $labels) {
        $op = $label.Operation
        $fontSize = if ($op.FontSize -ne $null) { [double]$op.FontSize } else { 0d }
        $classified.Add([pscustomobject]@{
            Kind = Classify-Text $op $label.Chart $PlotTolerance
            RegionIndex = $label.Chart.PlotBox.RegionIndex
            PageNumber = $op.PageNumber
            SourceKind = "Text"
            SourceOperator = $op.Operator
            TextLength = $label.Text.Length
            TextHash = StableTextHash $label.Text
            MinX = Round $label.X
            MinY = Round $label.Y
            MaxX = Round $label.X
            MaxY = Round $label.Y
            Width = 0d
            Height = 0d
            CenterX = Round $label.X
            CenterY = Round $label.Y
            LineWidth = 0d
            Font = $op.Font
            FontSize = Round $fontSize
            CharacterSpacing = if ($op.CharacterSpacing -ne $null) { Round ([double]$op.CharacterSpacing) } else { 0d }
            MatrixA = if ($op.EffectiveA -ne $null) { Round ([double]$op.EffectiveA) } else { 0d }
            MatrixB = if ($op.EffectiveB -ne $null) { Round ([double]$op.EffectiveB) } else { 0d }
            MatrixC = if ($op.EffectiveC -ne $null) { Round ([double]$op.EffectiveC) } else { 0d }
            MatrixD = if ($op.EffectiveD -ne $null) { Round ([double]$op.EffectiveD) } else { 0d }
        })
    }
}

$ordered = @($classified | Sort-Object -Property PageNumber, Kind, MinY, MinX, TextLength)
$ordered | Format-Table -AutoSize
Write-Host "Chart text structures: $($ordered.Count)"

if (-not [string]::IsNullOrWhiteSpace($Output)) {
    $outputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Output)
    $outputDirectory = Split-Path -Parent $outputPath
    if (-not [string]::IsNullOrWhiteSpace($outputDirectory)) {
        New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null
    }

    if ($ordered.Count -eq 0) {
        "[]" | Set-Content -LiteralPath $outputPath -Encoding UTF8
    }
    else {
        $ordered | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $outputPath -Encoding UTF8
    }
}
