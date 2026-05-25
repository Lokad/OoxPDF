param(
    [Parameter(Mandatory = $true)]
    [string] $InputPath,

    [string] $ChartStructures,

    [string] $Output,

    [int] $PageNumber = 0,

    [double] $PlotTolerance = 8
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

function Round([double]$value) { return [Math]::Round($value, 6) }

function TextValue($op) {
    if ($op.DecodedText -ne $null) {
        return [string]$op.DecodedText
    }

    return [string]$op.Payload
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

function Find-PlotBox($structures) {
    $priority = @(
        "GridlineAxisPlotBoxCandidate",
        "AxisPairPlotBoxCandidate",
        "PlotBoxCandidate",
        "PolarPlotBoxCandidate",
        "PlotAreaClipBoxCandidate"
    )
    foreach ($kind in $priority) {
        $matches = @($structures | Where-Object { $_.Kind -eq $kind } | Sort-Object -Property PageNumber, MinY, MinX)
        if ($matches.Count -gt 0) {
            return $matches[0]
        }
    }

    return $null
}

function Has-LegendSwatch($op, $structures, $plotBox) {
    $x = TextX $op
    $y = TextY $op
    $fontSize = if ($op.FontSize -ne $null) { [double]$op.FontSize } else { 0d }
    $verticalTolerance = [Math]::Max(8d, $fontSize * 0.75d)
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
        [Math]::Abs($x - $plotCenter) -le ($plotWidth * 0.35d)
}

function Classify-Text($op, $plotBox, $structures, [double]$tolerance) {
    if ($null -eq $plotBox) {
        return "ChartText"
    }

    $x = TextX $op
    $y = TextY $op
    $minX = [double]$plotBox.MinX
    $minY = [double]$plotBox.MinY
    $maxX = [double]$plotBox.MaxX
    $maxY = [double]$plotBox.MaxY
    $insideX = $x -ge ($minX - $tolerance) -and $x -le ($maxX + $tolerance)
    $insideY = $y -ge ($minY - $tolerance) -and $y -le ($maxY + $tolerance)
    $axisLabelY = $y -ge ($minY - ($tolerance * 2d)) -and $y -le ($maxY + ($tolerance * 2d))

    if (-not ($insideX -and $insideY) -and
        (Has-LegendSwatch $op $structures $plotBox)) {
        return "LegendText"
    }

    if ($insideX -and $insideY) {
        return "DataLabelText"
    }

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

    return "OuterChartText"
}

$textOps = Read-JsonArray $InputPath
if ($PageNumber -gt 0) {
    $textOps = @($textOps | Where-Object { [int]$_.PageNumber -eq $PageNumber })
}

$structures = Read-JsonArray $ChartStructures
if ($PageNumber -gt 0) {
    $structures = @($structures | Where-Object { [int]$_.PageNumber -eq $PageNumber })
}

$plotBox = Find-PlotBox $structures
$classified = New-Object System.Collections.Generic.List[object]
foreach ($op in $textOps) {
    $text = TextValue $op
    if ([string]::IsNullOrWhiteSpace($text)) {
        continue
    }

    $x = TextX $op
    $y = TextY $op
    $fontSize = if ($op.FontSize -ne $null) { [double]$op.FontSize } else { 0d }
    $classified.Add([pscustomobject]@{
        Kind = Classify-Text $op $plotBox $structures $PlotTolerance
        PageNumber = $op.PageNumber
        SourceKind = "Text"
        SourceOperator = $op.Operator
        TextLength = $text.Length
        TextHash = StableTextHash $text
        MinX = Round $x
        MinY = Round $y
        MaxX = Round $x
        MaxY = Round $y
        Width = 0d
        Height = 0d
        CenterX = Round $x
        CenterY = Round $y
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
