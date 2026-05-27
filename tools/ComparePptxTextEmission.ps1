param(
    [Parameter(Mandatory = $true)]
    [string] $ReferenceTextOperations,

    [Parameter(Mandatory = $true)]
    [string] $CandidateGlyphRuns,

    [double] $PositionTolerance = 0.05,

    [double] $FontSizeTolerance = 0.01,

    [switch] $MatchByPosition,

    [switch] $UseEffectiveMatrix,

    [switch] $IncludeText,

    [switch] $NoFail,

    [string] $OutputJson
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

function Delta([double] $left, [double] $right) {
    return [Math]::Round($right - $left, 6)
}

function HasValue($value) {
    return $null -ne $value -and [string]$value -ne ""
}

function OptionalValue($object, [string] $propertyName) {
    if ($null -eq $object) {
        return $null
    }

    $property = $object.PSObject.Properties[$propertyName]
    if ($null -eq $property) {
        return $null
    }

    return $property.Value
}

function OptionalRoundedDouble($object, [string] $propertyName) {
    $value = OptionalValue $object $propertyName
    if ($null -eq $value) {
        return $null
    }

    return [Math]::Round([double]$value, 6)
}

function RefX($op) {
    if ($UseEffectiveMatrix -and (HasValue $op.EffectiveX)) {
        return [double]$op.EffectiveX
    }

    return [double]$op.X
}

function RefY($op) {
    if ($UseEffectiveMatrix -and (HasValue $op.EffectiveY)) {
        return [double]$op.EffectiveY
    }

    return [double]$op.Y
}

function RefText($op) {
    if (HasValue $op.DecodedText) {
        return [string]$op.DecodedText
    }

    if (HasValue $op.Payload) {
        return [string]$op.Payload
    }

    return $null
}

function CandText($run) {
    if (HasValue $run.Text) {
        return [string]$run.Text
    }

    return $null
}

$referenceOps = Read-JsonArray $ReferenceTextOperations
$candidateRuns = Read-JsonArray $CandidateGlyphRuns

if ($MatchByPosition) {
    $unmatched = New-Object System.Collections.Generic.List[object]
    foreach ($op in $referenceOps) {
        $unmatched.Add($op)
    }

    $pairs = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $candidateRuns.Count; $i++) {
        $candidate = $candidateRuns[$i]
        $bestIndex = -1
        $bestScore = [double]::PositiveInfinity
        for ($j = 0; $j -lt $unmatched.Count; $j++) {
            $reference = $unmatched[$j]
            $score = [Math]::Abs([double]$candidate.BaselineY - (RefY $reference)) * 1000d +
                [Math]::Abs([double]$candidate.X - (RefX $reference))
            if ($score -lt $bestScore) {
                $bestScore = $score
                $bestIndex = $j
            }
        }

        $reference = if ($bestIndex -ge 0) { $unmatched[$bestIndex] } else { $null }
        if ($bestIndex -ge 0) {
            $unmatched.RemoveAt($bestIndex)
        }

        $pairs.Add([pscustomobject]@{ Index = $i; Reference = $reference; Candidate = $candidate })
    }

    foreach ($reference in $unmatched) {
        $pairs.Add([pscustomobject]@{ Index = $pairs.Count; Reference = $reference; Candidate = $null })
    }
}
else {
    $count = [Math]::Max($referenceOps.Count, $candidateRuns.Count)
    $pairs = for ($i = 0; $i -lt $count; $i++) {
        [pscustomobject]@{
            Index = $i
            Reference = if ($i -lt $referenceOps.Count) { $referenceOps[$i] } else { $null }
            Candidate = if ($i -lt $candidateRuns.Count) { $candidateRuns[$i] } else { $null }
        }
    }
}

$rows = New-Object System.Collections.Generic.List[object]
$failures = 0
foreach ($pair in $pairs) {
    $reference = $pair.Reference
    $candidate = $pair.Candidate
    if ($null -eq $reference -or $null -eq $candidate) {
        $failures++
        $rows.Add([pscustomobject]@{
            Index = $pair.Index
            Status = "missing"
            RefX = if ($null -eq $reference) { $null } else { [Math]::Round((RefX $reference), 6) }
            CandX = if ($null -eq $candidate) { $null } else { [Math]::Round([double]$candidate.X, 6) }
            DeltaX = $null
            RefBaselineY = if ($null -eq $reference) { $null } else { [Math]::Round((RefY $reference), 6) }
            CandBaselineY = if ($null -eq $candidate) { $null } else { [Math]::Round([double]$candidate.BaselineY, 6) }
            DeltaBaselineY = $null
            RefFontSize = if ($null -eq $reference) { $null } else { [Math]::Round([double]$reference.FontSize, 6) }
            CandPdfFontSize = if ($null -eq $candidate) { $null } else { [Math]::Round([double]$candidate.PdfFontSize, 6) }
            DeltaFontSize = $null
            CandLayoutFontSize = if ($null -eq $candidate) { $null } else { [Math]::Round([double]$candidate.LayoutFontSize, 6) }
            CandFrameIndex = if ($null -eq $candidate) { $null } else { $candidate.FrameIndex }
            CandParagraphIndex = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "ParagraphIndex" }
            CandLineIndex = if ($null -eq $candidate) { $null } else { $candidate.LineIndex }
            CandSpanIndex = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "SpanIndex" }
            CandLineSpanCount = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "LineSpanCount" }
            CandFrameShapeHeight = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "FrameShapeHeight" }
            CandFrameInsetTop = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "FrameInsetTop" }
            CandFrameInsetBottom = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "FrameInsetBottom" }
            CandFrameTextWidth = if ($null -eq $candidate) { $null } else { [Math]::Round([double]$candidate.FrameTextWidth, 6) }
            CandFrameTextHeight = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "FrameTextHeight" }
            CandFrameColumnCount = if ($null -eq $candidate) { $null } else { $candidate.FrameColumnCount }
            CandLineTopFromShapeTop = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "LineTopFromShapeTop" }
            CandLineTopFromTextTop = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "LineTopFromTextTop" }
            CandBaselineFromShapeTop = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "BaselineFromShapeTop" }
            CandLineBottomFromShapeBottom = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "LineBottomFromShapeBottom" }
            CandLineAdvance = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "LineAdvance" }
            RefText = if ($IncludeText -and $null -ne $reference) { RefText $reference } else { $null }
            CandText = if ($IncludeText -and $null -ne $candidate) { CandText $candidate } else { $null }
        })
        continue
    }

    $refX = RefX $reference
    $refY = RefY $reference
    $candX = [double]$candidate.X
    $candY = [double]$candidate.BaselineY
    $refFontSize = [double]$reference.FontSize
    $candFontSize = [double]$candidate.PdfFontSize
    $deltaX = Delta -left $refX -right $candX
    $deltaY = Delta -left $refY -right $candY
    $deltaFontSize = Delta -left $refFontSize -right $candFontSize
    $positionOk = [Math]::Abs($deltaX) -le $PositionTolerance -and [Math]::Abs($deltaY) -le $PositionTolerance
    $fontSizeOk = [Math]::Abs($deltaFontSize) -le $FontSizeTolerance
    $textOk = -not $IncludeText -or -not (HasValue (RefText $reference)) -or -not (HasValue (CandText $candidate)) -or (RefText $reference) -eq (CandText $candidate)
    $status = if ($positionOk -and $fontSizeOk -and $textOk) { "ok" } else { "delta" }
    if ($status -ne "ok") {
        $failures++
    }

    $rows.Add([pscustomobject]@{
        Index = $pair.Index
        Status = $status
        RefX = [Math]::Round($refX, 6)
        CandX = [Math]::Round($candX, 6)
        DeltaX = $deltaX
        RefBaselineY = [Math]::Round($refY, 6)
        CandBaselineY = [Math]::Round($candY, 6)
        DeltaBaselineY = $deltaY
        RefFontSize = [Math]::Round($refFontSize, 6)
        CandPdfFontSize = [Math]::Round($candFontSize, 6)
        DeltaFontSize = $deltaFontSize
        CandLayoutFontSize = [Math]::Round([double]$candidate.LayoutFontSize, 6)
        CandFrameIndex = $candidate.FrameIndex
        CandParagraphIndex = OptionalValue $candidate "ParagraphIndex"
        CandLineIndex = $candidate.LineIndex
        CandSpanIndex = OptionalValue $candidate "SpanIndex"
        CandLineSpanCount = OptionalValue $candidate "LineSpanCount"
        CandFrameShapeHeight = OptionalRoundedDouble $candidate "FrameShapeHeight"
        CandFrameInsetTop = OptionalRoundedDouble $candidate "FrameInsetTop"
        CandFrameInsetBottom = OptionalRoundedDouble $candidate "FrameInsetBottom"
        CandFrameTextWidth = [Math]::Round([double]$candidate.FrameTextWidth, 6)
        CandFrameTextHeight = OptionalRoundedDouble $candidate "FrameTextHeight"
        CandFrameColumnCount = $candidate.FrameColumnCount
        CandLineTopFromShapeTop = OptionalRoundedDouble $candidate "LineTopFromShapeTop"
        CandLineTopFromTextTop = OptionalRoundedDouble $candidate "LineTopFromTextTop"
        CandBaselineFromShapeTop = OptionalRoundedDouble $candidate "BaselineFromShapeTop"
        CandLineBottomFromShapeBottom = OptionalRoundedDouble $candidate "LineBottomFromShapeBottom"
        CandLineAdvance = OptionalRoundedDouble $candidate "LineAdvance"
        RefText = if ($IncludeText) { RefText $reference } else { $null }
        CandText = if ($IncludeText) { CandText $candidate } else { $null }
    })
}

if (HasValue $OutputJson) {
    $rows | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $OutputJson -Encoding UTF8
}

$rows | Format-Table -AutoSize
Write-Host "Text emission count: reference=$($referenceOps.Count), candidate=$($candidateRuns.Count), deltas=$failures"
if ($MatchByPosition) {
    Write-Host "Matching: nearest text position"
}
if ($UseEffectiveMatrix) {
    Write-Host "Reference coordinates: effective text matrix"
}
if ($IncludeText) {
    Write-Host "Text content: included"
}
else {
    Write-Host "Text content: omitted"
}

if ($failures -ne 0 -and -not $NoFail) {
    exit 1
}
