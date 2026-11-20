# Shared math and plot-box selection for PDF comparison and classification.

function CenterX($op) { return ([double]$op.MinX + [double]$op.MaxX) / 2d }
function CenterY($op) { return ([double]$op.MinY + [double]$op.MaxY) / 2d }
function Delta([double] $left, [double] $right) { return [Math]::Round($right - $left, 6) }
function RoundedKey($Value, [int] $Digits = 6) {
    if ($null -eq $Value -or [string]$Value -eq "") {
        return "(missing)"
    }

    return ([Math]::Round([double]$Value, $Digits)).ToString("0.######", [Globalization.CultureInfo]::InvariantCulture)
}
function Group-Count($items, [scriptblock] $keySelector) {
    $groups = @{}
    foreach ($item in $items) {
        $key = & $keySelector $item
        if ($null -eq $key -or [string]$key -eq "") {
            $key = "(missing)"
        }

        $key = [string]$key
        if ($groups.ContainsKey($key)) {
            $groups[$key]++
        }
        else {
            $groups[$key] = 1
        }
    }

    return @(
        foreach ($key in ($groups.Keys | Sort-Object)) {
            [pscustomobject]@{
                Key = $key
                Count = $groups[$key]
            }
        }
    )
}

function Get-PointToBoxDistance([double]$x, [double]$y, $box) {
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

function Find-PreferredChartPlotBoxes($structures, [string[]]$sortProperties, [switch]$CartesianOnly) {
    $priority = @(
        "GridlineAxisPlotBoxCandidate",
        "AxisPairPlotBoxCandidate",
        "CrossingAxisPlotBoxCandidate",
        "PlotBoxCandidate"
    )
    if (!$CartesianOnly) {
        $priority += @("PolarPlotBoxCandidate", "PlotAreaClipBoxCandidate")
    }

    foreach ($kind in $priority) {
        $matches = @($structures | Where-Object { $_.Kind -eq $kind } | Sort-Object -Property $sortProperties)
        if ($matches.Count -gt 0) {
            return $matches
        }
    }
}
