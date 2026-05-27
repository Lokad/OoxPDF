param(
    [Parameter(Mandatory = $true)]
    [string] $InputPath,

    [string] $Output,

    [int] $PageNumber = 0,

    [double] $LineTolerance = 0.25,

    [double] $MinLineLength = 12,

    [double] $MarkerMinSize = 4,

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
function IntValue($value) { if ($null -eq $value) { return 0 } return [int]$value }
function Is-ReasonableAxisPairBounds([double]$minX, [double]$minY, [double]$maxX, [double]$maxY) {
    return ($maxX - $minX) -ge 40d -and ($maxY - $minY) -ge 40d
}

function PathOperators($op) {
    if ($null -eq $op.PathCommands) {
        return ""
    }

    return (@($op.PathCommands) | ForEach-Object { [string]$_.Operator }) -join "/"
}

function Get-RadarSpokePathGeometry($op) {
    if ($null -eq $op.PathCommands) {
        return $null
    }

    $moves = New-Object System.Collections.Generic.List[object]
    $lines = New-Object System.Collections.Generic.List[object]
    foreach ($command in @($op.PathCommands)) {
        if ($null -eq $command.Values -or $command.Values.Count -lt 2) {
            continue
        }

        $point = [pscustomobject]@{
            X = [double]$command.Values[0]
            Y = [double]$command.Values[1]
        }
        if ([string]$command.Operator -eq "m") {
            $moves.Add($point)
        }
        elseif ([string]$command.Operator -eq "l") {
            $lines.Add($point)
        }
        else {
            return $null
        }
    }

    if ($moves.Count -lt 3 -or $moves.Count -ne $lines.Count) {
        return $null
    }

    $centerX = ($moves | Measure-Object -Property X -Average).Average
    $centerY = ($moves | Measure-Object -Property Y -Average).Average
    $maxCenterDelta = 0d
    foreach ($move in $moves) {
        $maxCenterDelta = [Math]::Max($maxCenterDelta, [Math]::Abs([double]$move.X - [double]$centerX))
        $maxCenterDelta = [Math]::Max($maxCenterDelta, [Math]::Abs([double]$move.Y - [double]$centerY))
    }

    $radii = @($lines | ForEach-Object {
        [Math]::Sqrt(([double]$_.X - [double]$centerX) * ([double]$_.X - [double]$centerX) + ([double]$_.Y - [double]$centerY) * ([double]$_.Y - [double]$centerY))
    })
    if ($radii.Count -eq 0) {
        return $null
    }

    [pscustomobject]@{
        PathCenterX = Round $centerX
        PathCenterY = Round $centerY
        PathRadius = Round (($radii | Measure-Object -Average).Average)
        PathMinRadius = Round (($radii | Measure-Object -Minimum).Minimum)
        PathMaxRadius = Round (($radii | Measure-Object -Maximum).Maximum)
        PathCenterMaxDelta = Round $maxCenterDelta
        PathSpokeCount = $lines.Count
    }
}

function Get-PieSlicePathGeometry($op) {
    if ($null -eq $op.PathCommands) {
        return $null
    }

    $commands = @($op.PathCommands)
    if ($commands.Count -lt 4 -or [string]$commands[0].Operator -ne "m") {
        return $null
    }

    $lineCommands = @($commands | Where-Object { [string]$_.Operator -eq "l" -and $null -ne $_.Values -and $_.Values.Count -ge 2 })
    if ($lineCommands.Count -ne 1) {
        return $null
    }

    $lastDrawingCommand = @($commands | Where-Object { [string]$_.Operator -ne "h" } | Select-Object -Last 1)
    if ($lastDrawingCommand.Count -eq 0 -or [string]$lastDrawingCommand[0].Operator -ne "l") {
        return $null
    }

    $move = $commands[0]
    if ($null -eq $move.Values -or $move.Values.Count -lt 2) {
        return $null
    }

    $centerX = [double]$lineCommands[0].Values[0]
    $centerY = [double]$lineCommands[0].Values[1]
    $radius = [Math]::Sqrt(
        ([double]$move.Values[0] - $centerX) * ([double]$move.Values[0] - $centerX) +
        ([double]$move.Values[1] - $centerY) * ([double]$move.Values[1] - $centerY))

    [pscustomobject]@{
        PathCenterX = Round $centerX
        PathCenterY = Round $centerY
        PathRadius = Round $radius
        PathMinRadius = Round $radius
        PathMaxRadius = Round $radius
        PathCenterMaxDelta = 0d
        PathSpokeCount = $null
    }
}

function New-Point($values, [int]$offset) {
    [pscustomobject]@{
        X = [double]$values[$offset]
        Y = [double]$values[$offset + 1]
    }
}

function Distance($a, $b) {
    return [Math]::Sqrt(([double]$a.X - [double]$b.X) * ([double]$a.X - [double]$b.X) +
        ([double]$a.Y - [double]$b.Y) * ([double]$a.Y - [double]$b.Y))
}

function Get-CommandEndPoint($command) {
    if ($null -eq $command.Values -or $command.Values.Count -lt 2) {
        return $null
    }

    if ([string]$command.Operator -eq "c") {
        if ($command.Values.Count -lt 6) {
            return $null
        }

        return New-Point $command.Values 4
    }

    return New-Point $command.Values 0
}

function Intersect-Lines($a1, $a2, $b1, $b2) {
    $ax = [double]$a2.X - [double]$a1.X
    $ay = [double]$a2.Y - [double]$a1.Y
    $bx = [double]$b2.X - [double]$b1.X
    $by = [double]$b2.Y - [double]$b1.Y
    $denominator = $ax * $by - $ay * $bx
    if ([Math]::Abs($denominator) -lt 0.000001d) {
        $collinearDelta = ([double]$b1.X - [double]$a1.X) * $ay - ([double]$b1.Y - [double]$a1.Y) * $ax
        if ([Math]::Abs($collinearDelta) -gt 0.001d) {
            return $null
        }

        return [pscustomobject]@{
            X = ([double]$a1.X + [double]$b1.X) / 2d
            Y = ([double]$a1.Y + [double]$b1.Y) / 2d
        }
    }

    $cx = [double]$b1.X - [double]$a1.X
    $cy = [double]$b1.Y - [double]$a1.Y
    $t = ($cx * $by - $cy * $bx) / $denominator
    [pscustomobject]@{
        X = [double]$a1.X + $t * $ax
        Y = [double]$a1.Y + $t * $ay
    }
}

function Get-AnnularSlicePathGeometry($op) {
    if ($null -eq $op.PathCommands) {
        return $null
    }

    $commands = @($op.PathCommands)
    if ($commands.Count -lt 5 -or [string]$commands[0].Operator -ne "m") {
        return $null
    }

    $drawingCommands = @($commands | Where-Object { [string]$_.Operator -ne "h" })
    $lineIndexes = @()
    for ($i = 0; $i -lt $drawingCommands.Count; $i++) {
        if ([string]$drawingCommands[$i].Operator -eq "l") {
            $lineIndexes += $i
        }
        elseif ([string]$drawingCommands[$i].Operator -ne "m" -and [string]$drawingCommands[$i].Operator -ne "c") {
            return $null
        }
    }

    if ($lineIndexes.Count -ne 1) {
        return $null
    }

    $lineIndex = [int]$lineIndexes[0]
    if ($lineIndex -lt 2 -or $lineIndex -ge ($drawingCommands.Count - 1)) {
        return $null
    }

    $outerStart = Get-CommandEndPoint $drawingCommands[0]
    $outerEnd = Get-CommandEndPoint $drawingCommands[$lineIndex - 1]
    $innerEnd = Get-CommandEndPoint $drawingCommands[$lineIndex]
    $innerStart = Get-CommandEndPoint $drawingCommands[$drawingCommands.Count - 1]
    if ($null -eq $outerStart -or $null -eq $outerEnd -or $null -eq $innerEnd -or $null -eq $innerStart) {
        return $null
    }

    $center = Intersect-Lines $outerStart $innerStart $outerEnd $innerEnd
    if ($null -eq $center) {
        return $null
    }

    $outerRadii = @((Distance $center $outerStart), (Distance $center $outerEnd))
    $innerRadii = @((Distance $center $innerStart), (Distance $center $innerEnd))
    $outerRadius = ($outerRadii | Measure-Object -Average).Average
    $innerRadius = ($innerRadii | Measure-Object -Average).Average
    $maxRadiusDelta = 0d
    foreach ($radius in @($outerRadii + $innerRadii)) {
        $target = if ($radius -gt (($outerRadius + $innerRadius) / 2d)) { $outerRadius } else { $innerRadius }
        $maxRadiusDelta = [Math]::Max($maxRadiusDelta, [Math]::Abs([double]$radius - [double]$target))
    }

    [pscustomobject]@{
        PathCenterX = Round $center.X
        PathCenterY = Round $center.Y
        PathRadius = Round $outerRadius
        PathMinRadius = Round $innerRadius
        PathMaxRadius = Round $outerRadius
        PathCenterMaxDelta = Round $maxRadiusDelta
        PathSpokeCount = $null
    }
}

function BoundsDistance($a, $b) {
    return [Math]::Abs([double]$a.MinX - [double]$b.MinX) +
        [Math]::Abs([double]$a.MinY - [double]$b.MinY) +
        [Math]::Abs([double]$a.MaxX - [double]$b.MaxX) +
        [Math]::Abs([double]$a.MaxY - [double]$b.MaxY)
}

function New-Structure($kind, $op) {
    $radarSpokeGeometry = if ($kind -eq "RadarSpokeGroupCandidate") { Get-RadarSpokePathGeometry $op } else { $null }
    $pieSliceGeometry = if ($kind -eq "FilledRegion") { Get-PieSlicePathGeometry $op } else { $null }
    $annularSliceGeometry = if ($kind -eq "FilledRegion" -and $null -eq $pieSliceGeometry) { Get-AnnularSlicePathGeometry $op } else { $null }
    $pathGeometry = if ($null -ne $radarSpokeGeometry) { $radarSpokeGeometry } elseif ($null -ne $pieSliceGeometry) { $pieSliceGeometry } else { $annularSliceGeometry }
    [pscustomobject]@{
        Kind = $kind
        RegionIndex = $null
        PageNumber = $op.PageNumber
        SourceKind = $op.Kind
        SourceOperator = $op.Operator
        SegmentCount = $op.SegmentCount
        MoveCount = IntValue $op.MoveCount
        LineCount = IntValue $op.LineCount
        CurveCount = IntValue $op.CurveCount
        CloseCount = IntValue $op.CloseCount
        PathOperators = PathOperators $op
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
        PathCenterX = if ($null -eq $pathGeometry) { $null } else { $pathGeometry.PathCenterX }
        PathCenterY = if ($null -eq $pathGeometry) { $null } else { $pathGeometry.PathCenterY }
        PathRadius = if ($null -eq $pathGeometry) { $null } else { $pathGeometry.PathRadius }
        PathMinRadius = if ($null -eq $pathGeometry) { $null } else { $pathGeometry.PathMinRadius }
        PathMaxRadius = if ($null -eq $pathGeometry) { $null } else { $pathGeometry.PathMaxRadius }
        PathCenterMaxDelta = if ($null -eq $pathGeometry) { $null } else { $pathGeometry.PathCenterMaxDelta }
        PathSpokeCount = if ($null -eq $pathGeometry) { $null } else { $pathGeometry.PathSpokeCount }
    }
}

function New-DerivedStructure($kind, $pageNumber, $sourceOperator, $segmentCount, $minX, $minY, $maxX, $maxY) {
    [pscustomobject]@{
        Kind = $kind
        RegionIndex = $null
        PageNumber = $pageNumber
        SourceKind = "Derived"
        SourceOperator = $sourceOperator
        SegmentCount = $segmentCount
        MoveCount = 0
        LineCount = 0
        CurveCount = 0
        CloseCount = 0
        PathOperators = ""
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
        PathCenterX = $null
        PathCenterY = $null
        PathRadius = $null
        PathMinRadius = $null
        PathMaxRadius = $null
        PathCenterMaxDelta = $null
        PathSpokeCount = $null
    }
}

function New-PolarPlotBoxStructure($pageNumber, $sourceOperator, $segmentCount, $minX, $minY, $maxX, $maxY) {
    $structure = New-DerivedStructure "PolarPlotBoxCandidate" $pageNumber $sourceOperator $segmentCount $minX $minY $maxX $maxY
    $width = [double]$maxX - [double]$minX
    $height = [double]$maxY - [double]$minY
    $radius = [Math]::Max($width, $height) / 2d
    $structure.PathCenterX = Round (([double]$minX + [double]$maxX) / 2d)
    $structure.PathCenterY = Round (([double]$minY + [double]$maxY) / 2d)
    $structure.PathRadius = Round $radius
    $structure.PathMinRadius = Round ([Math]::Min($width, $height) / 2d)
    $structure.PathMaxRadius = Round $radius
    $structure.PathCenterMaxDelta = Round ([Math]::Abs($width - $height) / 2d)
    return $structure
}

function Copy-StructureAsKind($kind, $structure) {
    [pscustomobject]@{
        Kind = $kind
        RegionIndex = $structure.RegionIndex
        PageNumber = $structure.PageNumber
        SourceKind = $structure.SourceKind
        SourceOperator = $structure.SourceOperator
        SegmentCount = $structure.SegmentCount
        MoveCount = IntValue $structure.MoveCount
        LineCount = IntValue $structure.LineCount
        CurveCount = IntValue $structure.CurveCount
        CloseCount = IntValue $structure.CloseCount
        PathOperators = $structure.PathOperators
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
        PathCenterX = $structure.PathCenterX
        PathCenterY = $structure.PathCenterY
        PathRadius = $structure.PathRadius
        PathMinRadius = $structure.PathMinRadius
        PathMaxRadius = $structure.PathMaxRadius
        PathCenterMaxDelta = $structure.PathCenterMaxDelta
        PathSpokeCount = $structure.PathSpokeCount
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

function Is-BoundsInsideBox($structure, $box, [double]$tolerance) {
    return ([double]$structure.MinX -ge ([double]$box.MinX - $tolerance)) -and
        ([double]$structure.MaxX -le ([double]$box.MaxX + $tolerance)) -and
        ([double]$structure.MinY -ge ([double]$box.MinY - $tolerance)) -and
        ([double]$structure.MaxY -le ([double]$box.MaxY + $tolerance))
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

function Distance-ToBox([double]$x, [double]$y, $box) {
    $dx = 0d
    if ($x -lt [double]$box.MinX) {
        $dx = [double]$box.MinX - $x
    }
    elseif ($x -gt [double]$box.MaxX) {
        $dx = $x - [double]$box.MaxX
    }

    $dy = 0d
    if ($y -lt [double]$box.MinY) {
        $dy = [double]$box.MinY - $y
    }
    elseif ($y -gt [double]$box.MaxY) {
        $dy = $y - [double]$box.MaxY
    }

    return [Math]::Sqrt($dx * $dx + $dy * $dy)
}

function Find-RegionPlotBoxes($structures) {
    $priority = @(
        "GridlineAxisPlotBoxCandidate",
        "AxisPairPlotBoxCandidate",
        "CrossingAxisPlotBoxCandidate",
        "PlotBoxCandidate",
        "PolarPlotBoxCandidate",
        "PlotAreaClipBoxCandidate"
    )

    foreach ($kind in $priority) {
        $matches = @($structures |
            Where-Object { $_.Kind -eq $kind } |
            Sort-Object -Property PageNumber, MinY, MinX, MaxY, MaxX)
        if ($matches.Count -gt 0) {
            return $matches
        }
    }

    return @()
}

function Find-CartesianRegionPlotBoxes($structures) {
    $priority = @(
        "GridlineAxisPlotBoxCandidate",
        "AxisPairPlotBoxCandidate",
        "CrossingAxisPlotBoxCandidate",
        "PlotBoxCandidate"
    )

    foreach ($kind in $priority) {
        $matches = @($structures |
            Where-Object { $_.Kind -eq $kind } |
            Sort-Object -Property PageNumber, RegionIndex, MinY, MinX, MaxY, MaxX)
        if ($matches.Count -gt 0) {
            return $matches
        }
    }

    return @()
}

function Assign-RegionIndexes($structures) {
    $plotBoxes = @(Find-RegionPlotBoxes $structures)
    if ($plotBoxes.Count -eq 0) {
        return
    }

    for ($i = 0; $i -lt $plotBoxes.Count; $i++) {
        $plotBoxes[$i].RegionIndex = $i
    }

    foreach ($structure in $structures) {
        if ($structure.RegionIndex -ne $null) {
            continue
        }

        $centerX = CenterX $structure
        $centerY = CenterY $structure
        $best = $null
        $bestScore = [double]::PositiveInfinity
        foreach ($plotBox in $plotBoxes) {
            $score = Distance-ToBox $centerX $centerY $plotBox
            if ($score -lt $bestScore) {
                $bestScore = $score
                $best = $plotBox
            }
        }

        if ($null -ne $best) {
            $structure.RegionIndex = $best.RegionIndex
        }
    }
}

function Reclassify-InPlotSeriesLines($structures, [double]$tolerance) {
    $plotBoxes = @(Find-CartesianRegionPlotBoxes $structures)
    if ($plotBoxes.Count -eq 0) {
        return
    }

    $snapshot = [object[]]$structures.ToArray()
    foreach ($structure in $snapshot) {
        if ($structure.Kind -ne "DataLabelLeaderLineCandidate") {
            continue
        }

        $matchingPlotBoxes = @($plotBoxes | Where-Object {
            [int]$_.PageNumber -eq [int]$structure.PageNumber -and
            ($null -eq $structure.RegionIndex -or $null -eq $_.RegionIndex -or [int]$_.RegionIndex -eq [int]$structure.RegionIndex)
        })
        if ($matchingPlotBoxes.Count -eq 0) {
            continue
        }

        $insidePlotBox = $false
        foreach ($plotBox in $matchingPlotBoxes) {
            if (Is-BoundsInsideBox $structure $plotBox $tolerance) {
                $insidePlotBox = $true
                break
            }
        }

        if ($insidePlotBox) {
            $structures.Add((Copy-StructureAsKind "ChartSeriesLineCandidate" $structure))
            [void]$structures.Remove($structure)
        }
    }
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
        $width -ge $MarkerMinSize -and $height -ge $MarkerMinSize -and
        $width -le $MarkerMaxSize -and $height -le $MarkerMaxSize) {
        $structures.Add((New-Structure "StrokeMarkerCandidate" $op))
    }
    elseif ($op.Kind -eq "Stroke" -and
        (IntValue $op.MoveCount) -eq 1 -and
        (IntValue $op.LineCount) -ge 1 -and
        (IntValue $op.LineCount) -le 2 -and
        (IntValue $op.CurveCount) -eq 0 -and
        (IntValue $op.CloseCount) -eq 0 -and
        $width -gt $LineTolerance -and
        $height -gt $LineTolerance -and
        ($width -ge $MinLineLength -or $height -ge $MinLineLength)) {
        $structures.Add((New-Structure "DataLabelLeaderLineCandidate" $op))
    }

    if (($op.Kind -eq "Fill" -or $op.Kind -eq "FillStroke") -and $width -gt 0 -and $height -gt 0) {
        if ($width -ge $MarkerMinSize -and $height -ge $MarkerMinSize -and
            $width -le $MarkerMaxSize -and $height -le $MarkerMaxSize) {
            $structures.Add((New-Structure "MarkerCandidate" $op))
        }
        elseif ($width -gt $MarkerMaxSize -or $height -gt $MarkerMaxSize) {
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
    if ($maxX -gt $minX -and $maxY -gt $minY -and (Is-ReasonableAxisPairBounds $minX $minY $maxX $maxY)) {
        $page = if ($PageNumber -gt 0) { $PageNumber } else { $leftAxis.PageNumber }
        $structures.Add((New-DerivedStructure "AxisPairPlotBoxCandidate" $page "AxisLinePairBounds" 2 $minX $minY $maxX $maxY))
    }
    elseif ((Is-Near ([double]$topAxis.MinY) ([double]$leftAxis.MaxY) $GridlineBoundsTolerance) -and
        $maxX -gt $minX -and [double]$leftAxis.MaxY -gt [double]$leftAxis.MinY -and
        (Is-ReasonableAxisPairBounds $minX ([double]$leftAxis.MinY) $maxX ([double]$topAxis.MinY))) {
        $page = if ($PageNumber -gt 0) { $PageNumber } else { $leftAxis.PageNumber }
        $structures.Add((New-DerivedStructure "AxisPairPlotBoxCandidate" $page "AxisLinePairBounds" 2 $minX $leftAxis.MinY $maxX $topAxis.MinY))
    }

    $crossingAxis = @($horizontalLines | Where-Object {
        (Is-Near ([double]$_.MinX) ([double]$leftAxis.MinX) $GridlineBoundsTolerance) -and
        ([double]$_.CenterY -gt ([double]$leftAxis.MinY + $GridlineBoundsTolerance)) -and
        ([double]$_.CenterY -lt ([double]$leftAxis.MaxY - $GridlineBoundsTolerance)) -and
        ([double]$_.MaxX -gt [double]$leftAxis.MinX)
    } | Sort-Object -Property Width -Descending | Select-Object -First 1)
    if ($crossingAxis.Count -gt 0) {
        $axis = $crossingAxis[0]
        $page = if ($PageNumber -gt 0) { $PageNumber } else { $leftAxis.PageNumber }
        $structures.Add((New-DerivedStructure "CrossingAxisPlotBoxCandidate" $page "CrossingAxisLinePairBounds" 2 $leftAxis.MinX $leftAxis.MinY $axis.MaxX $leftAxis.MaxY))
    }
}

$axisPairPlotBox = @($structures | Where-Object { $_.Kind -eq "AxisPairPlotBoxCandidate" } | Select-Object -First 1)
$gridlineAxisPlotBox = @($structures | Where-Object { $_.Kind -eq "GridlineAxisPlotBoxCandidate" } | Select-Object -First 1)
$crossingAxisPlotBox = @($structures | Where-Object { $_.Kind -eq "CrossingAxisPlotBoxCandidate" } | Select-Object -First 1)
$plotBoxForGridlines = if ($gridlineAxisPlotBox.Count -gt 0) {
    $gridlineAxisPlotBox[0]
}
elseif ($crossingAxisPlotBox.Count -gt 0) {
    $crossingAxisPlotBox[0]
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

if ($null -eq $plotBoxForGridlines) {
    foreach ($op in $ops) {
        $segmentCount = if ($null -ne $op.SegmentCount) { [int]$op.SegmentCount } else { 0 }
        $moveCount = IntValue $op.MoveCount
        $lineCount = IntValue $op.LineCount
        $curveCount = IntValue $op.CurveCount
        $width = Width $op
        $height = Height $op
        if ($op.Kind -eq "Stroke" -and $segmentCount -ge 6 -and $moveCount -ge 3 -and
            $lineCount -ge 3 -and $curveCount -eq 0 -and $width -gt 40d -and $height -gt 40d) {
            $kind = if ($segmentCount -ge 10) { "RadarRingGridGroupCandidate" } else { "RadarSpokeGroupCandidate" }
            $structures.Add((New-Structure $kind $op))
        }
    }
}

$clipBoxes = @($structures | Where-Object { $_.Kind -eq "ClipBox" })
if ($clipBoxes.Count -gt 0) {
    $nonPageClipBoxes = @($clipBoxes | Where-Object {
        $_.Width -gt 40 -and $_.Height -gt 40 -and
        -not ([Math]::Abs([double]$_.MinX) -le 0.01 -and [Math]::Abs([double]$_.MinY) -le 0.01)
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
            $axisAlignedGroup = $clipGroups |
                Where-Object {
                    $clip = @($_.Group)[0]
                    foreach ($line in $horizontalLines) {
                        if ((Is-Near ([double]$clip.MinX) ([double]$line.MinX) $GridlineBoundsTolerance) -and
                            (Is-Near ([double]$clip.MaxX) ([double]$line.MaxX) $GridlineBoundsTolerance)) {
                            return $true
                        }
                    }

                    return $false
                } |
                Sort-Object -Property @{
                    Expression = { -[int]$_.Count }
                }, @{
                    Expression = {
                        $clip = @($_.Group)[0]
                        -((Width $clip) * (Height $clip))
                    }
                } |
                Select-Object -First 1
            if ($null -ne $axisAlignedGroup) {
                $dominantGroup = $axisAlignedGroup
            }
        }

        if ($null -eq $dominantGroup -and $null -eq $plotBoxForGridlines) {
            $dominantGroup = $nonPageClipBoxes |
                Group-Object -Property MinX, MinY, MaxX, MaxY |
                Sort-Object -Property @{
                    Expression = {
                        $clip = @($_.Group)[0]
                        -((Width $clip) * (Height $clip))
                    }
                }, @{
                    Expression = { -[int]$_.Count }
                } |
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
    $pageLikeClipBoxes = @($clipBoxes | Where-Object {
        [double]$_.Width -gt 400d -and [double]$_.Height -gt 300d
    })
    $nonPageRegions = if ($pageLikeClipBoxes.Count -gt 0) {
        $largestPageWidth = ($pageLikeClipBoxes | Measure-Object -Property Width -Maximum).Maximum
        $largestPageHeight = ($pageLikeClipBoxes | Measure-Object -Property Height -Maximum).Maximum
        @($polarRegions | Where-Object {
            -not ([double]$_.Width -ge ([double]$largestPageWidth * 0.8d) -and [double]$_.Height -ge ([double]$largestPageHeight * 0.8d))
        })
    }
    else {
        $polarRegions
    }

    if ($nonPageRegions.Count -ge 2 -and $null -eq $plotBoxForGridlines) {
        $bounds = Get-UnionBounds $nonPageRegions
        $width = [double]$bounds.MaxX - [double]$bounds.MinX
        $height = [double]$bounds.MaxY - [double]$bounds.MinY
        $aspect = $width / [Math]::Max(1d, $height)
        if ($aspect -ge 0.74d -and $aspect -le 1.35d) {
            foreach ($region in $nonPageRegions) {
                $structures.Add((Copy-StructureAsKind "PolarSliceCandidate" $region))
            }

            $page = if ($PageNumber -gt 0) { $PageNumber } else { $nonPageRegions[0].PageNumber }
            $structures.Add((New-PolarPlotBoxStructure $page "FilledRegionUnion" $nonPageRegions.Count $bounds.MinX $bounds.MinY $bounds.MaxX $bounds.MaxY))
        }
    }
}

Assign-RegionIndexes ([object[]]$structures.ToArray())
Reclassify-InPlotSeriesLines $structures $GridlineBoundsTolerance

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
