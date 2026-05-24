param(
    [Parameter(Mandatory = $true)]
    [string] $InputPath,

    [string] $Output,

    [int] $PageNumber = 0,

    [double] $LineTolerance = 0.25,

    [double] $MinLineLength = 12,

    [double] $MarkerMaxSize = 20
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
        LineWidth = [double]$op.LineWidth
        StrokeColor = $op.StrokeColor
        FillColor = $op.FillColor
        Dash = $op.Dash
        LineCap = $op.LineCap
        LineJoin = $op.LineJoin
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
    $minX = ($horizontalLines | Measure-Object -Property MinX -Minimum).Minimum
    $maxX = ($horizontalLines | Measure-Object -Property MaxX -Maximum).Maximum
    $minY = ($horizontalLines | Measure-Object -Property MinY -Minimum).Minimum
    $maxY = ($horizontalLines | Measure-Object -Property MaxY -Maximum).Maximum
    $structures.Add([pscustomobject]@{
        Kind = "PlotBoxCandidate"
        PageNumber = if ($PageNumber -gt 0) { $PageNumber } else { $horizontalLines[0].PageNumber }
        SourceKind = "Derived"
        SourceOperator = "HorizontalLineBounds"
        SegmentCount = $horizontalLines.Count
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
    })
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
