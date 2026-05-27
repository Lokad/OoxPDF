param(
    [Parameter(Mandatory = $true)]
    [string] $Reference,

    [Parameter(Mandatory = $true)]
    [string] $Candidate,

    [double] $BoundsTolerance = 0.25,

    [hashtable] $BoundsToleranceByKind = @{},

    [switch] $UseBoundsToleranceForUnlistedKinds,

    [double] $LineWidthTolerance = 0.05,

    [string[]] $Kinds = @("Stroke", "Fill", "FillStroke", "Clip"),

    [string[]] $Operators = @(),

    [int] $PageNumber = 0,

    [switch] $MatchByBounds,

    [switch] $MatchOperator,

    [switch] $MatchSegmentCount,

    [switch] $MatchPathCommandCounts,

    [switch] $MatchPathOperators,

    [switch] $MatchStrokeColor,

    [switch] $MatchLineCap,

    [switch] $MatchLineJoin,

    [switch] $MatchTextHash,

    [switch] $MatchPathGeometry,

    [double] $PathGeometryTolerance = 0.25,

    [hashtable] $PathGeometryToleranceByKind = @{},

    [switch] $UsePathGeometryToleranceForUnlistedKinds
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
    $selectedKinds = @($Kinds | ForEach-Object { [string]$_ -split "," } | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })
    $selectedOperators = @($Operators | ForEach-Object { [string]$_ -split "," } | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })
    $selected = @($items | Where-Object { $selectedKinds -contains $_.Kind })
    if ($selectedOperators.Count -gt 0) {
        $selected = @($selected | Where-Object { $selectedOperators -contains (OperationSourceOperator $_) })
    }

    if ($PageNumber -gt 0) {
        $selected = @($selected | Where-Object { [int]$_.PageNumber -eq $PageNumber })
    }

    return ,$selected
}

function OperationSourceOperator($op) {
    if ($null -eq $op) {
        return ""
    }

    if (HasValue $op.SourceOperator) {
        return [string]$op.SourceOperator
    }

    if (HasValue $op.Operator) {
        return [string]$op.Operator
    }

    return ""
}

function CenterX($op) { return ([double]$op.MinX + [double]$op.MaxX) / 2d }
function CenterY($op) { return ([double]$op.MinY + [double]$op.MaxY) / 2d }
function Width($op) { return [double]$op.MaxX - [double]$op.MinX }
function Height($op) { return [double]$op.MaxY - [double]$op.MinY }
function Delta([double] $left, [double] $right) { return [Math]::Round($right - $left, 6) }
function IntValue($value) { if ($null -eq $value) { return 0 } return [int]$value }
function HasValue($value) { return $null -ne $value -and [string]$value -ne "" }
function ColorValues($value) {
    $text = [string]$value
    if ($text -match '^G:(?<g>[0-9.]+)$') {
        $g = [double]$Matches['g']
        return @($g, $g, $g)
    }

    if ($text -match '^RG:(?<r>[0-9.]+),(?<g>[0-9.]+),(?<b>[0-9.]+)$') {
        return @([double]$Matches['r'], [double]$Matches['g'], [double]$Matches['b'])
    }

    return @()
}
function StrokeColorEqual($left, $right) {
    $leftValues = @(ColorValues $left)
    $rightValues = @(ColorValues $right)
    if ($leftValues.Count -eq 0 -or $rightValues.Count -eq 0) {
        return [string]$left -eq [string]$right
    }

    for ($i = 0; $i -lt 3; $i++) {
        if ([Math]::Abs($leftValues[$i] - $rightValues[$i]) -gt 0.002d) {
            return $false
        }
    }

    return $true
}
function GetPathGeometryTolerance($kind) {
    if ($PathGeometryToleranceByKind.ContainsKey([string]$kind)) {
        return [double]$PathGeometryToleranceByKind[[string]$kind]
    }

    if ($PathGeometryToleranceByKind.Count -eq 0 -or $UsePathGeometryToleranceForUnlistedKinds) {
        return [double]$PathGeometryTolerance
    }

    return $null
}
function GetBoundsTolerance($kind) {
    if ($BoundsToleranceByKind.ContainsKey([string]$kind)) {
        return [double]$BoundsToleranceByKind[[string]$kind]
    }

    if ($BoundsToleranceByKind.Count -eq 0 -or $UseBoundsToleranceForUnlistedKinds) {
        return [double]$BoundsTolerance
    }

    return $null
}

$referenceOps = Select-Ops (Read-JsonArray $Reference)
$candidateOps = Select-Ops (Read-JsonArray $Candidate)

if ($MatchByBounds) {
    $candidatesByScore = New-Object System.Collections.Generic.List[object]
    for ($candidateIndex = 0; $candidateIndex -lt $candidateOps.Count; $candidateIndex++) {
        $cand = $candidateOps[$candidateIndex]
        for ($referenceIndex = 0; $referenceIndex -lt $referenceOps.Count; $referenceIndex++) {
            $ref = $referenceOps[$referenceIndex]
            $kindPenalty = if ([string]$cand.Kind -eq [string]$ref.Kind) { 0d } else { 1000000d }
            $operatorPenalty = if (-not $MatchOperator -or (OperationSourceOperator $cand) -eq (OperationSourceOperator $ref)) { 0d } else { 100000d }
            $segmentPenalty = if (-not $MatchSegmentCount -or [int]$cand.SegmentCount -eq [int]$ref.SegmentCount) { 0d } else { 10000d }
            $pathCountPenalty = if (-not $MatchPathCommandCounts -or (
                    (IntValue $cand.MoveCount) -eq (IntValue $ref.MoveCount) -and
                    (IntValue $cand.LineCount) -eq (IntValue $ref.LineCount) -and
                    (IntValue $cand.CurveCount) -eq (IntValue $ref.CurveCount) -and
                    (IntValue $cand.CloseCount) -eq (IntValue $ref.CloseCount))) { 0d } else { 1000d }
            $textHashPenalty = if (-not $MatchTextHash -or (-not (HasValue $cand.TextHash)) -or (-not (HasValue $ref.TextHash)) -or [string]$cand.TextHash -eq [string]$ref.TextHash) { 0d } else { 10000000d }
            $score = $kindPenalty +
                $operatorPenalty +
                $segmentPenalty +
                $pathCountPenalty +
                $textHashPenalty +
                [Math]::Abs((CenterX $cand) - (CenterX $ref)) +
                [Math]::Abs((CenterY $cand) - (CenterY $ref)) +
                [Math]::Abs((Width $cand) - (Width $ref)) +
                [Math]::Abs((Height $cand) - (Height $ref))
            $candidatesByScore.Add([pscustomobject]@{
                ReferenceIndex = $referenceIndex
                CandidateIndex = $candidateIndex
                Score = $score
            })
        }
    }

    $matchedReference = New-Object bool[] $referenceOps.Count
    $matchedCandidate = New-Object bool[] $candidateOps.Count
    $matchedPairs = New-Object System.Collections.Generic.List[object]
    foreach ($match in ($candidatesByScore | Sort-Object Score, CandidateIndex, ReferenceIndex)) {
        if ($matchedReference[$match.ReferenceIndex] -or $matchedCandidate[$match.CandidateIndex]) {
            continue
        }

        $matchedReference[$match.ReferenceIndex] = $true
        $matchedCandidate[$match.CandidateIndex] = $true
        $matchedPairs.Add($match)
    }

    $pairs = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $candidateOps.Count; $i++) {
        $match = $matchedPairs | Where-Object { $_.CandidateIndex -eq $i } | Select-Object -First 1
        $pairs.Add([pscustomobject]@{
            Index = $i
            Reference = if ($null -eq $match) { $null } else { $referenceOps[$match.ReferenceIndex] }
            Candidate = $candidateOps[$i]
        })
    }

    for ($i = 0; $i -lt $referenceOps.Count; $i++) {
        if (-not $matchedReference[$i]) {
            $pairs.Add([pscustomobject]@{ Index = $pairs.Count; Reference = $referenceOps[$i]; Candidate = $null })
        }
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
    $boundsToleranceForKind = GetBoundsTolerance $ref.Kind
    $boundsOk = ($null -eq $boundsToleranceForKind) -or (
        [Math]::Abs($deltaMinX) -le $boundsToleranceForKind -and
        [Math]::Abs($deltaMinY) -le $boundsToleranceForKind -and
        [Math]::Abs($deltaMaxX) -le $boundsToleranceForKind -and
        [Math]::Abs($deltaMaxY) -le $boundsToleranceForKind)
    $widthOk = [Math]::Abs($deltaWidth) -le $LineWidthTolerance
    $kindOk = [string]$ref.Kind -eq [string]$cand.Kind
    $operatorOk = (-not $MatchOperator) -or (OperationSourceOperator $ref) -eq (OperationSourceOperator $cand)
    $segmentCountOk = (-not $MatchSegmentCount) -or [int]$ref.SegmentCount -eq [int]$cand.SegmentCount
    $pathCommandCountsOk = (-not $MatchPathCommandCounts) -or (
        (IntValue $ref.MoveCount) -eq (IntValue $cand.MoveCount) -and
        (IntValue $ref.LineCount) -eq (IntValue $cand.LineCount) -and
        (IntValue $ref.CurveCount) -eq (IntValue $cand.CurveCount) -and
        (IntValue $ref.CloseCount) -eq (IntValue $cand.CloseCount))
    $pathOperatorsOk = (-not $MatchPathOperators) -or [string]$ref.PathOperators -eq [string]$cand.PathOperators
    $strokeColorOk = (-not $MatchStrokeColor) -or (StrokeColorEqual $ref.StrokeColor $cand.StrokeColor)
    $lineCapOk = (-not $MatchLineCap) -or (IntValue $ref.LineCap) -eq (IntValue $cand.LineCap)
    $lineJoinOk = (-not $MatchLineJoin) -or (IntValue $ref.LineJoin) -eq (IntValue $cand.LineJoin)
    $textHashOk = (-not $MatchTextHash) -or (-not (HasValue $ref.TextHash)) -or (-not (HasValue $cand.TextHash)) -or [string]$ref.TextHash -eq [string]$cand.TextHash
    $pathGeometryAvailable = (HasValue $ref.PathCenterX) -and (HasValue $cand.PathCenterX) -and
        (HasValue $ref.PathCenterY) -and (HasValue $cand.PathCenterY) -and
        (HasValue $ref.PathRadius) -and (HasValue $cand.PathRadius)
    $pathGeometryRelevant = (HasValue $ref.PathCenterX) -or (HasValue $cand.PathCenterX) -or
        (HasValue $ref.PathCenterY) -or (HasValue $cand.PathCenterY) -or
        (HasValue $ref.PathRadius) -or (HasValue $cand.PathRadius)
    $pathMinRadiusRelevant = (HasValue $ref.PathMinRadius) -or (HasValue $cand.PathMinRadius)
    $pathMaxRadiusRelevant = (HasValue $ref.PathMaxRadius) -or (HasValue $cand.PathMaxRadius)
    $pathMinRadiusAvailable = (HasValue $ref.PathMinRadius) -and (HasValue $cand.PathMinRadius)
    $pathMaxRadiusAvailable = (HasValue $ref.PathMaxRadius) -and (HasValue $cand.PathMaxRadius)
    $deltaPathCenterX = if ($pathGeometryAvailable) { Delta ([double]$ref.PathCenterX) ([double]$cand.PathCenterX) } else { $null }
    $deltaPathCenterY = if ($pathGeometryAvailable) { Delta ([double]$ref.PathCenterY) ([double]$cand.PathCenterY) } else { $null }
    $deltaPathRadius = if ($pathGeometryAvailable) { Delta ([double]$ref.PathRadius) ([double]$cand.PathRadius) } else { $null }
    $deltaPathMinRadius = if ($pathMinRadiusAvailable) { Delta ([double]$ref.PathMinRadius) ([double]$cand.PathMinRadius) } else { $null }
    $deltaPathMaxRadius = if ($pathMaxRadiusAvailable) { Delta ([double]$ref.PathMaxRadius) ([double]$cand.PathMaxRadius) } else { $null }
    $pathGeometryToleranceForKind = GetPathGeometryTolerance $ref.Kind
    $pathGeometryOk = (-not $MatchPathGeometry) -or (-not $pathGeometryRelevant) -or ($null -eq $pathGeometryToleranceForKind) -or (
        $pathGeometryAvailable -and
        [Math]::Abs($deltaPathCenterX) -le $pathGeometryToleranceForKind -and
        [Math]::Abs($deltaPathCenterY) -le $pathGeometryToleranceForKind -and
        [Math]::Abs($deltaPathRadius) -le $pathGeometryToleranceForKind -and
        ((-not $pathMinRadiusRelevant) -or ($pathMinRadiusAvailable -and [Math]::Abs($deltaPathMinRadius) -le $pathGeometryToleranceForKind)) -and
        ((-not $pathMaxRadiusRelevant) -or ($pathMaxRadiusAvailable -and [Math]::Abs($deltaPathMaxRadius) -le $pathGeometryToleranceForKind)))
    $status = if ($boundsOk -and $widthOk -and $kindOk -and $operatorOk -and $segmentCountOk -and $pathCommandCountsOk -and $pathOperatorsOk -and $strokeColorOk -and $lineCapOk -and $lineJoinOk -and $textHashOk -and $pathGeometryOk) { "ok" } else { "delta" }
    if ($status -ne "ok") {
        $failures++
    }

    $rows.Add([pscustomobject]@{
        Index = $pair.Index
        Status = $status
        RefKind = $ref.Kind
        CandKind = $cand.Kind
        RefOp = OperationSourceOperator $ref
        CandOp = OperationSourceOperator $cand
        RefSeg = $ref.SegmentCount
        CandSeg = $cand.SegmentCount
        RefPath = "$(IntValue $ref.MoveCount)/$(IntValue $ref.LineCount)/$(IntValue $ref.CurveCount)/$(IntValue $ref.CloseCount)"
        CandPath = "$(IntValue $cand.MoveCount)/$(IntValue $cand.LineCount)/$(IntValue $cand.CurveCount)/$(IntValue $cand.CloseCount)"
        RefPathOps = $ref.PathOperators
        CandPathOps = $cand.PathOperators
        OperatorOk = $operatorOk
        SegmentCountOk = $segmentCountOk
        PathCommandCountsOk = $pathCommandCountsOk
        PathOperatorsOk = $pathOperatorsOk
        RefStrokeColor = $ref.StrokeColor
        CandStrokeColor = $cand.StrokeColor
        StrokeColorOk = $strokeColorOk
        RefLineCap = $ref.LineCap
        CandLineCap = $cand.LineCap
        LineCapOk = $lineCapOk
        RefLineJoin = $ref.LineJoin
        CandLineJoin = $cand.LineJoin
        LineJoinOk = $lineJoinOk
        RefTextHash = $ref.TextHash
        CandTextHash = $cand.TextHash
        TextHashOk = $textHashOk
        RefPathCenter = if (HasValue $ref.PathCenterX) { "$($ref.PathCenterX),$($ref.PathCenterY)" } else { $null }
        CandPathCenter = if (HasValue $cand.PathCenterX) { "$($cand.PathCenterX),$($cand.PathCenterY)" } else { $null }
        DeltaPathCenterX = $deltaPathCenterX
        DeltaPathCenterY = $deltaPathCenterY
        RefPathRadius = $ref.PathRadius
        CandPathRadius = $cand.PathRadius
        DeltaPathRadius = $deltaPathRadius
        RefPathMinRadius = $ref.PathMinRadius
        CandPathMinRadius = $cand.PathMinRadius
        DeltaPathMinRadius = $deltaPathMinRadius
        RefPathMaxRadius = $ref.PathMaxRadius
        CandPathMaxRadius = $cand.PathMaxRadius
        DeltaPathMaxRadius = $deltaPathMaxRadius
        PathGeometryTolerance = $pathGeometryToleranceForKind
        PathGeometryOk = $pathGeometryOk
        DeltaMinX = $deltaMinX
        DeltaMinY = $deltaMinY
        DeltaMaxX = $deltaMaxX
        DeltaMaxY = $deltaMaxY
        BoundsTolerance = $boundsToleranceForKind
        DeltaWidth = $deltaWidth
    })
}

$rows | Format-Table -AutoSize
Write-Host "Graphics operation count: reference=$($referenceOps.Count), candidate=$($candidateOps.Count), deltas=$failures"
Write-Host "Kinds: $($Kinds -join ', ')"
if ($Operators.Count -gt 0) {
    Write-Host "Operators: $($Operators -join ', ')"
}
if ($PageNumber -gt 0) {
    Write-Host "Page: $PageNumber"
}
if ($MatchByBounds) {
    Write-Host "Matching: nearest bounds"
}

if ($failures -ne 0) {
    exit 1
}
