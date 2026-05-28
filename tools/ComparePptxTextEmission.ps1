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

    [string] $OutputJson,

    [string] $OutputSummaryJson
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

function RoundAway([double] $value) {
    return [Math]::Round($value, 6, [MidpointRounding]::AwayFromZero)
}

function OfficeExportGridStep() {
    return 72d / 600d
}

function OfficeMainFontGrid([double] $fontSize) {
    $gridStep = OfficeExportGridStep
    $deviceUnits = [Math]::Round($fontSize / $gridStep, [MidpointRounding]::AwayFromZero)
    return $deviceUnits * $gridStep
}

function OfficeGridRemainder([double] $value) {
    $gridStep = OfficeExportGridStep
    $scaled = $value / $gridStep
    $floor = [Math]::Floor($scaled)
    return RoundAway ($scaled - $floor)
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

function OptionalGridRemainder($object, [string] $propertyName) {
    $value = OptionalValue $object $propertyName
    if ($null -eq $value) {
        return $null
    }

    return OfficeGridRemainder ([double]$value)
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

function RoundedKey($value) {
    if ($null -eq $value -or [string]$value -eq "") {
        return "(missing)"
    }

    return ([Math]::Round([double]$value, 6)).ToString("0.######", [Globalization.CultureInfo]::InvariantCulture)
}

function BranchKey($row) {
    $delta = $row.RefSecondaryFontDelta
    if ($null -eq $delta -or [string]$delta -eq "") {
        return "(missing)"
    }

    if ([Math]::Abs([double]$delta) -le 0.000001d) {
        return "main-grid"
    }

    return "secondary-" + (RoundedKey $delta)
}

function OptionalDouble($row, [string] $propertyName) {
    $value = OptionalValue $row $propertyName
    if ($null -eq $value -or [string]$value -eq "") {
        return $null
    }

    return [double]$value
}

function Group-BranchExtents($items) {
    $groups = @{}
    foreach ($item in $items) {
        $key = BranchKey $item
        if (-not $groups.ContainsKey($key)) {
            $groups[$key] = New-Object System.Collections.Generic.List[object]
        }

        $groups[$key].Add($item)
    }

    return @(
        foreach ($key in ($groups.Keys | Sort-Object)) {
            $groupItems = @($groups[$key].ToArray())
            $refBaseline = @($groupItems | ForEach-Object { OptionalDouble $_ "RefBaselineY" } | Where-Object { $null -ne $_ })
            $candBaseline = @($groupItems | ForEach-Object { OptionalDouble $_ "CandBaselineY" } | Where-Object { $null -ne $_ })
            $frameTop = @($groupItems | ForEach-Object { OptionalDouble $_ "CandFrameShapeTopY" } | Where-Object { $null -ne $_ })
            $lineTop = @($groupItems | ForEach-Object { OptionalDouble $_ "CandLineTopY" } | Where-Object { $null -ne $_ })
            [pscustomobject]@{
                Key = $key
                Count = $groupItems.Count
                MinRefBaselineY = if ($refBaseline.Count -eq 0) { $null } else { [Math]::Round(($refBaseline | Measure-Object -Minimum).Minimum, 6) }
                MaxRefBaselineY = if ($refBaseline.Count -eq 0) { $null } else { [Math]::Round(($refBaseline | Measure-Object -Maximum).Maximum, 6) }
                MinCandidateBaselineY = if ($candBaseline.Count -eq 0) { $null } else { [Math]::Round(($candBaseline | Measure-Object -Minimum).Minimum, 6) }
                MaxCandidateBaselineY = if ($candBaseline.Count -eq 0) { $null } else { [Math]::Round(($candBaseline | Measure-Object -Maximum).Maximum, 6) }
                MinCandidateFrameTopY = if ($frameTop.Count -eq 0) { $null } else { [Math]::Round(($frameTop | Measure-Object -Minimum).Minimum, 6) }
                MaxCandidateFrameTopY = if ($frameTop.Count -eq 0) { $null } else { [Math]::Round(($frameTop | Measure-Object -Maximum).Maximum, 6) }
                MinCandidateLineTopY = if ($lineTop.Count -eq 0) { $null } else { [Math]::Round(($lineTop | Measure-Object -Minimum).Minimum, 6) }
                MaxCandidateLineTopY = if ($lineTop.Count -eq 0) { $null } else { [Math]::Round(($lineTop | Measure-Object -Maximum).Maximum, 6) }
            }
        }
    )
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
            RefBaselineY600Remainder = if ($null -eq $reference) { $null } else { OfficeGridRemainder (RefY $reference) }
            CandBaselineY600Remainder = if ($null -eq $candidate) { $null } else { OfficeGridRemainder ([double]$candidate.BaselineY) }
            RefFontSize = if ($null -eq $reference) { $null } else { [Math]::Round([double]$reference.FontSize, 6) }
            CandPdfFontSize = if ($null -eq $candidate) { $null } else { [Math]::Round([double]$candidate.PdfFontSize, 6) }
            DeltaFontSize = $null
            CandLayoutFontSize = if ($null -eq $candidate) { $null } else { [Math]::Round([double]$candidate.LayoutFontSize, 6) }
            MainFontGrid = if ($null -eq $candidate) { $null } else { RoundAway (OfficeMainFontGrid ([double]$candidate.LayoutFontSize)) }
            RefSecondaryFontDelta = if ($null -eq $reference -or $null -eq $candidate) { $null } else { Delta -left (OfficeMainFontGrid ([double]$candidate.LayoutFontSize)) -right ([double]$reference.FontSize) }
            CandPdfGridDelta = if ($null -eq $candidate) { $null } else { Delta -left (OfficeMainFontGrid ([double]$candidate.LayoutFontSize)) -right ([double]$candidate.PdfFontSize) }
            CandHighlightColor = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "HighlightColor" }
            CandHighlightX = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "HighlightX" }
            CandHighlightY = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "HighlightY" }
            CandHighlightWidth = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "HighlightWidth" }
            CandHighlightHeight = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "HighlightHeight" }
            CandUnderlineX = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "UnderlineX" }
            CandUnderlineY = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "UnderlineY" }
            CandUnderlineWidth = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "UnderlineWidth" }
            CandUnderlineHeight = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "UnderlineHeight" }
            CandStrikeX = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "StrikeX" }
            CandStrikeY = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "StrikeY" }
            CandStrikeWidth = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "StrikeWidth" }
            CandStrikeHeight = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "StrikeHeight" }
            CandFrameIndex = if ($null -eq $candidate) { $null } else { $candidate.FrameIndex }
            CandParagraphIndex = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "ParagraphIndex" }
            CandLineIndex = if ($null -eq $candidate) { $null } else { $candidate.LineIndex }
            CandSpanIndex = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "SpanIndex" }
            CandLineSpanCount = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "LineSpanCount" }
            CandFrameShapeX = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "FrameShapeX" }
            CandFrameShapeTopY = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "FrameShapeTopY" }
            CandFrameShapeTopY600Remainder = if ($null -eq $candidate) { $null } else { OptionalGridRemainder $candidate "FrameShapeTopY" }
            CandFrameShapeWidth = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "FrameShapeWidth" }
            CandFrameShapeHeight = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "FrameShapeHeight" }
            CandFrameInsetTop = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "FrameInsetTop" }
            CandFrameInsetBottom = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "FrameInsetBottom" }
            CandFrameTextWidth = if ($null -eq $candidate) { $null } else { [Math]::Round([double]$candidate.FrameTextWidth, 6) }
            CandFrameTextHeight = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "FrameTextHeight" }
            CandFrameColumnCount = if ($null -eq $candidate) { $null } else { $candidate.FrameColumnCount }
            CandLineTopY = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "LineTopY" }
            CandLineTopY600Remainder = if ($null -eq $candidate) { $null } else { OptionalGridRemainder $candidate "LineTopY" }
            CandLineTopFromShapeTop = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "LineTopFromShapeTop" }
            CandLineTopFromTextTop = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "LineTopFromTextTop" }
            CandBaselineFromShapeTop = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "BaselineFromShapeTop" }
            CandLineBottomFromShapeBottom = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "LineBottomFromShapeBottom" }
            CandLineAdvance = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "LineAdvance" }
            CandGlyphCount = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "GlyphCount" }
            CandFirstAdjustmentAfterOrigin = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "FirstAdjustmentAfterOrigin" }
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
    $mainFontGrid = OfficeMainFontGrid ([double]$candidate.LayoutFontSize)
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
        RefBaselineY600Remainder = OfficeGridRemainder $refY
        CandBaselineY600Remainder = OfficeGridRemainder $candY
        RefFontSize = [Math]::Round($refFontSize, 6)
        CandPdfFontSize = [Math]::Round($candFontSize, 6)
        DeltaFontSize = $deltaFontSize
        CandLayoutFontSize = [Math]::Round([double]$candidate.LayoutFontSize, 6)
        MainFontGrid = RoundAway $mainFontGrid
        RefSecondaryFontDelta = Delta -left $mainFontGrid -right $refFontSize
        CandPdfGridDelta = Delta -left $mainFontGrid -right $candFontSize
        CandHighlightColor = OptionalValue $candidate "HighlightColor"
        CandHighlightX = OptionalRoundedDouble $candidate "HighlightX"
        CandHighlightY = OptionalRoundedDouble $candidate "HighlightY"
        CandHighlightWidth = OptionalRoundedDouble $candidate "HighlightWidth"
        CandHighlightHeight = OptionalRoundedDouble $candidate "HighlightHeight"
        CandUnderlineX = OptionalRoundedDouble $candidate "UnderlineX"
        CandUnderlineY = OptionalRoundedDouble $candidate "UnderlineY"
        CandUnderlineWidth = OptionalRoundedDouble $candidate "UnderlineWidth"
        CandUnderlineHeight = OptionalRoundedDouble $candidate "UnderlineHeight"
        CandStrikeX = OptionalRoundedDouble $candidate "StrikeX"
        CandStrikeY = OptionalRoundedDouble $candidate "StrikeY"
        CandStrikeWidth = OptionalRoundedDouble $candidate "StrikeWidth"
        CandStrikeHeight = OptionalRoundedDouble $candidate "StrikeHeight"
        CandFrameIndex = $candidate.FrameIndex
        CandParagraphIndex = OptionalValue $candidate "ParagraphIndex"
        CandLineIndex = $candidate.LineIndex
        CandSpanIndex = OptionalValue $candidate "SpanIndex"
        CandLineSpanCount = OptionalValue $candidate "LineSpanCount"
        CandFrameShapeX = OptionalRoundedDouble $candidate "FrameShapeX"
        CandFrameShapeTopY = OptionalRoundedDouble $candidate "FrameShapeTopY"
        CandFrameShapeTopY600Remainder = OptionalGridRemainder $candidate "FrameShapeTopY"
        CandFrameShapeWidth = OptionalRoundedDouble $candidate "FrameShapeWidth"
        CandFrameShapeHeight = OptionalRoundedDouble $candidate "FrameShapeHeight"
        CandFrameInsetTop = OptionalRoundedDouble $candidate "FrameInsetTop"
        CandFrameInsetBottom = OptionalRoundedDouble $candidate "FrameInsetBottom"
        CandFrameTextWidth = [Math]::Round([double]$candidate.FrameTextWidth, 6)
        CandFrameTextHeight = OptionalRoundedDouble $candidate "FrameTextHeight"
        CandFrameColumnCount = $candidate.FrameColumnCount
        CandLineTopY = OptionalRoundedDouble $candidate "LineTopY"
        CandLineTopY600Remainder = OptionalGridRemainder $candidate "LineTopY"
        CandLineTopFromShapeTop = OptionalRoundedDouble $candidate "LineTopFromShapeTop"
        CandLineTopFromTextTop = OptionalRoundedDouble $candidate "LineTopFromTextTop"
        CandBaselineFromShapeTop = OptionalRoundedDouble $candidate "BaselineFromShapeTop"
        CandLineBottomFromShapeBottom = OptionalRoundedDouble $candidate "LineBottomFromShapeBottom"
        CandLineAdvance = OptionalRoundedDouble $candidate "LineAdvance"
        CandGlyphCount = OptionalValue $candidate "GlyphCount"
        CandFirstAdjustmentAfterOrigin = OptionalRoundedDouble $candidate "FirstAdjustmentAfterOrigin"
        RefText = if ($IncludeText) { RefText $reference } else { $null }
        CandText = if ($IncludeText) { CandText $candidate } else { $null }
    })
}

$rowsArray = $rows.ToArray()

if (HasValue $OutputJson) {
    $rowsArray | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $OutputJson -Encoding UTF8
}

if (HasValue $OutputSummaryJson) {
    $summary = [pscustomobject]@{
        Total = $rowsArray.Count
        StatusCounts = Group-Count $rowsArray { param($row) $row.Status }
        FontBranchCounts = Group-Count $rowsArray { param($row) BranchKey $row }
        FontBranchExtents = Group-BranchExtents $rowsArray
        RefSecondaryFontDeltas = Group-Count $rowsArray { param($row) RoundedKey $row.RefSecondaryFontDelta }
        ByLayoutFontSizeAndBranch = Group-Count $rowsArray { param($row) (RoundedKey $row.CandLayoutFontSize) + "|" + (BranchKey $row) }
        ByParagraphIndexAndBranch = Group-Count $rowsArray { param($row) (RoundedKey $row.CandParagraphIndex) + "|" + (BranchKey $row) }
        BySpanIndexAndBranch = Group-Count $rowsArray { param($row) (RoundedKey $row.CandSpanIndex) + "|" + (BranchKey $row) }
        ByRefBaselineRemainderAndBranch = Group-Count $rowsArray { param($row) (RoundedKey $row.RefBaselineY600Remainder) + "|" + (BranchKey $row) }
        ByCandidateBaselineRemainderAndBranch = Group-Count $rowsArray { param($row) (RoundedKey $row.CandBaselineY600Remainder) + "|" + (BranchKey $row) }
        ByCandidateFrameXAndBranch = Group-Count $rowsArray { param($row) (RoundedKey $row.CandFrameShapeX) + "|" + (BranchKey $row) }
        ByCandidateFrameTopRemainderAndBranch = Group-Count $rowsArray { param($row) (RoundedKey $row.CandFrameShapeTopY600Remainder) + "|" + (BranchKey $row) }
        ByCandidateFrameWidthAndBranch = Group-Count $rowsArray { param($row) (RoundedKey $row.CandFrameShapeWidth) + "|" + (BranchKey $row) }
        ByCandidateTextWidthAndBranch = Group-Count $rowsArray { param($row) (RoundedKey $row.CandFrameTextWidth) + "|" + (BranchKey $row) }
        ByCandidateLineTopRemainderAndBranch = Group-Count $rowsArray { param($row) (RoundedKey $row.CandLineTopY600Remainder) + "|" + (BranchKey $row) }
        ByCandidateLineTopFromShapeTopAndBranch = Group-Count $rowsArray { param($row) (RoundedKey $row.CandLineTopFromShapeTop) + "|" + (BranchKey $row) }
        ByCandidateLineTopFromTextTopAndBranch = Group-Count $rowsArray { param($row) (RoundedKey $row.CandLineTopFromTextTop) + "|" + (BranchKey $row) }
        ByCandidateBaselineFromShapeTopAndBranch = Group-Count $rowsArray { param($row) (RoundedKey $row.CandBaselineFromShapeTop) + "|" + (BranchKey $row) }
        ByCandidateLineIndexAndBranch = Group-Count $rowsArray { param($row) (RoundedKey $row.CandLineIndex) + "|" + (BranchKey $row) }
        ByCandidateLineSpanCountAndBranch = Group-Count $rowsArray { param($row) (RoundedKey $row.CandLineSpanCount) + "|" + (BranchKey $row) }
        ByCandidateFrameHeightAndBranch = Group-Count $rowsArray { param($row) (RoundedKey $row.CandFrameShapeHeight) + "|" + (BranchKey $row) }
        ByCandidateTextHeightAndBranch = Group-Count $rowsArray { param($row) (RoundedKey $row.CandFrameTextHeight) + "|" + (BranchKey $row) }
        ByCandidateGlyphCountAndBranch = Group-Count $rowsArray { param($row) (RoundedKey (OptionalValue $row "CandGlyphCount")) + "|" + (BranchKey $row) }
        ByCandidateFirstAdjustmentAndBranch = Group-Count $rowsArray { param($row) (RoundedKey (OptionalValue $row "CandFirstAdjustmentAfterOrigin")) + "|" + (BranchKey $row) }
    }

    $summary | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $OutputSummaryJson -Encoding UTF8
}

$rowsArray | Format-Table -AutoSize
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
