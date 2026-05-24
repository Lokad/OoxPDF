param(
    [Parameter(Mandatory = $true)]
    [string] $Reference,

    [Parameter(Mandatory = $true)]
    [string] $Candidate,

    [double] $BoundsTolerance = 0.25,

    [double] $LineWidthTolerance = 0.05,

    [string[]] $Kinds = @("Stroke", "Fill", "FillStroke", "Clip"),

    [int] $PageNumber = 0,

    [switch] $MatchByBounds
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

function Select-Ops($items) {
    $selected = @($items | Where-Object { $Kinds -contains $_.Kind })
    if ($PageNumber -gt 0) {
        $selected = @($selected | Where-Object { [int]$_.PageNumber -eq $PageNumber })
    }

    return ,$selected
}

function CenterX($op) { return ([double]$op.MinX + [double]$op.MaxX) / 2d }
function CenterY($op) { return ([double]$op.MinY + [double]$op.MaxY) / 2d }
function Width($op) { return [double]$op.MaxX - [double]$op.MinX }
function Height($op) { return [double]$op.MaxY - [double]$op.MinY }
function Delta([double] $left, [double] $right) { return [Math]::Round($right - $left, 6) }

$referenceOps = Select-Ops (Read-JsonArray $Reference)
$candidateOps = Select-Ops (Read-JsonArray $Candidate)

if ($MatchByBounds) {
    $unmatched = New-Object System.Collections.Generic.List[object]
    foreach ($op in $referenceOps) {
        $unmatched.Add($op)
    }

    $pairs = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $candidateOps.Count; $i++) {
        $cand = $candidateOps[$i]
        $bestIndex = -1
        $bestScore = [double]::PositiveInfinity
        for ($j = 0; $j -lt $unmatched.Count; $j++) {
            $ref = $unmatched[$j]
            $score = [Math]::Abs((CenterX $cand) - (CenterX $ref)) +
                [Math]::Abs((CenterY $cand) - (CenterY $ref)) +
                [Math]::Abs((Width $cand) - (Width $ref)) +
                [Math]::Abs((Height $cand) - (Height $ref))
            if ($score -lt $bestScore) {
                $bestScore = $score
                $bestIndex = $j
            }
        }

        $ref = if ($bestIndex -ge 0) { $unmatched[$bestIndex] } else { $null }
        if ($bestIndex -ge 0) {
            $unmatched.RemoveAt($bestIndex)
        }

        $pairs.Add([pscustomobject]@{ Index = $i; Reference = $ref; Candidate = $cand })
    }

    foreach ($ref in $unmatched) {
        $pairs.Add([pscustomobject]@{ Index = $pairs.Count; Reference = $ref; Candidate = $null })
    }
}
else {
    $count = [Math]::Max($referenceOps.Count, $candidateOps.Count)
    $pairs = for ($i = 0; $i -lt $count; $i++) {
        [pscustomobject]@{
            Index = $i
            Reference = if ($i -lt $referenceOps.Count) { $referenceOps[$i] } else { $null }
            Candidate = if ($i -lt $candidateOps.Count) { $candidateOps[$i] } else { $null }
        }
    }
}

$rows = New-Object System.Collections.Generic.List[object]
$failures = 0

foreach ($pair in $pairs) {
    $ref = $pair.Reference
    $cand = $pair.Candidate
    if ($null -eq $ref -or $null -eq $cand) {
        $failures++
        $rows.Add([pscustomobject]@{
            Index = $pair.Index
            Status = "missing"
            RefKind = if ($null -eq $ref) { $null } else { $ref.Kind }
            CandKind = if ($null -eq $cand) { $null } else { $cand.Kind }
            DeltaMinX = $null
            DeltaMinY = $null
            DeltaMaxX = $null
            DeltaMaxY = $null
            DeltaWidth = $null
        })
        continue
    }

    $deltaMinX = Delta ([double]$ref.MinX) ([double]$cand.MinX)
    $deltaMinY = Delta ([double]$ref.MinY) ([double]$cand.MinY)
    $deltaMaxX = Delta ([double]$ref.MaxX) ([double]$cand.MaxX)
    $deltaMaxY = Delta ([double]$ref.MaxY) ([double]$cand.MaxY)
    $deltaWidth = Delta ([double]$ref.LineWidth) ([double]$cand.LineWidth)
    $boundsOk = [Math]::Abs($deltaMinX) -le $BoundsTolerance -and
        [Math]::Abs($deltaMinY) -le $BoundsTolerance -and
        [Math]::Abs($deltaMaxX) -le $BoundsTolerance -and
        [Math]::Abs($deltaMaxY) -le $BoundsTolerance
    $widthOk = [Math]::Abs($deltaWidth) -le $LineWidthTolerance
    $kindOk = [string]$ref.Kind -eq [string]$cand.Kind
    $status = if ($boundsOk -and $widthOk -and $kindOk) { "ok" } else { "delta" }
    if ($status -ne "ok") {
        $failures++
    }

    $rows.Add([pscustomobject]@{
        Index = $pair.Index
        Status = $status
        RefKind = $ref.Kind
        CandKind = $cand.Kind
        RefOp = $ref.Operator
        CandOp = $cand.Operator
        RefSeg = $ref.SegmentCount
        CandSeg = $cand.SegmentCount
        DeltaMinX = $deltaMinX
        DeltaMinY = $deltaMinY
        DeltaMaxX = $deltaMaxX
        DeltaMaxY = $deltaMaxY
        DeltaWidth = $deltaWidth
    })
}

$rows | Format-Table -AutoSize
Write-Host "Graphics operation count: reference=$($referenceOps.Count), candidate=$($candidateOps.Count), deltas=$failures"
Write-Host "Kinds: $($Kinds -join ', ')"
if ($PageNumber -gt 0) {
    Write-Host "Page: $PageNumber"
}
if ($MatchByBounds) {
    Write-Host "Matching: nearest bounds"
}

if ($failures -ne 0) {
    exit 1
}
