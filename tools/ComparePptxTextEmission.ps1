param(
    [Parameter(Mandatory = $true)]
    [string] $ReferenceTextOperations,

    [Parameter(Mandatory = $true)]
    [string] $CandidateGlyphRuns,

    [double] $PositionTolerance = 0.05,

    [double] $FontSizeTolerance = 0.01,

    [switch] $MatchByPosition,

    [switch] $MatchByTextThenPosition,

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

function NullableGridRemainder($value) {
    if ($null -eq $value -or [string]$value -eq "") {
        return $null
    }

    return OfficeGridRemainder ([double]$value)
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

function CandidatePageHeight($candidate) {
    if ($null -eq $candidate) {
        return $null
    }

    $clipY = OptionalValue $candidate "FrameClipY"
    $clipHeight = OptionalValue $candidate "FrameClipHeight"
    if ($null -eq $clipY -or $null -eq $clipHeight) {
        return $null
    }

    return [double]$clipY + [double]$clipHeight
}

function PageTopBaseline($pageHeight, $baselineY) {
    if ($null -eq $pageHeight -or $null -eq $baselineY) {
        return $null
    }

    return [Math]::Round([double]$pageHeight - [double]$baselineY, 6)
}

function RefBaselineFromPageTop($reference, $candidate) {
    if ($null -eq $reference) {
        return $null
    }

    return PageTopBaseline (CandidatePageHeight $candidate) (RefY $reference)
}

function CandidateBaselineFromPageTop($candidate) {
    if ($null -eq $candidate) {
        return $null
    }

    return PageTopBaseline (CandidatePageHeight $candidate) ([double]$candidate.BaselineY)
}

function RefBaselineFromCandidateShapeTop($reference, $candidate) {
    $pageTopBaseline = RefBaselineFromPageTop $reference $candidate
    $shapeTop = OptionalValue $candidate "FrameShapeTopY"
    if ($null -eq $pageTopBaseline -or $null -eq $shapeTop) {
        return $null
    }

    return [Math]::Round([double]$pageTopBaseline - [double]$shapeTop, 6)
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

function StringKey($value) {
    if ($null -eq $value -or [string]$value -eq "") {
        return "(missing)"
    }

    return [string]$value
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
            $refBaselineFromPageTop = @($groupItems | ForEach-Object { OptionalDouble $_ "RefBaselineFromPageTop" } | Where-Object { $null -ne $_ })
            $candBaselineFromPageTop = @($groupItems | ForEach-Object { OptionalDouble $_ "CandBaselineFromPageTop" } | Where-Object { $null -ne $_ })
            $refBaselineFromShapeTop = @($groupItems | ForEach-Object { OptionalDouble $_ "RefBaselineFromCandidateShapeTop" } | Where-Object { $null -ne $_ })
            $frameTop = @($groupItems | ForEach-Object { OptionalDouble $_ "CandFrameShapeTopY" } | Where-Object { $null -ne $_ })
            $lineTop = @($groupItems | ForEach-Object { OptionalDouble $_ "CandLineTopY" } | Where-Object { $null -ne $_ })
            [pscustomobject]@{
                Key = $key
                Count = $groupItems.Count
                MinRefBaselineY = if ($refBaseline.Count -eq 0) { $null } else { [Math]::Round(($refBaseline | Measure-Object -Minimum).Minimum, 6) }
                MaxRefBaselineY = if ($refBaseline.Count -eq 0) { $null } else { [Math]::Round(($refBaseline | Measure-Object -Maximum).Maximum, 6) }
                MinCandidateBaselineY = if ($candBaseline.Count -eq 0) { $null } else { [Math]::Round(($candBaseline | Measure-Object -Minimum).Minimum, 6) }
                MaxCandidateBaselineY = if ($candBaseline.Count -eq 0) { $null } else { [Math]::Round(($candBaseline | Measure-Object -Maximum).Maximum, 6) }
                MinRefBaselineFromPageTop = if ($refBaselineFromPageTop.Count -eq 0) { $null } else { [Math]::Round(($refBaselineFromPageTop | Measure-Object -Minimum).Minimum, 6) }
                MaxRefBaselineFromPageTop = if ($refBaselineFromPageTop.Count -eq 0) { $null } else { [Math]::Round(($refBaselineFromPageTop | Measure-Object -Maximum).Maximum, 6) }
                MinCandidateBaselineFromPageTop = if ($candBaselineFromPageTop.Count -eq 0) { $null } else { [Math]::Round(($candBaselineFromPageTop | Measure-Object -Minimum).Minimum, 6) }
                MaxCandidateBaselineFromPageTop = if ($candBaselineFromPageTop.Count -eq 0) { $null } else { [Math]::Round(($candBaselineFromPageTop | Measure-Object -Maximum).Maximum, 6) }
                MinRefBaselineFromCandidateShapeTop = if ($refBaselineFromShapeTop.Count -eq 0) { $null } else { [Math]::Round(($refBaselineFromShapeTop | Measure-Object -Minimum).Minimum, 6) }
                MaxRefBaselineFromCandidateShapeTop = if ($refBaselineFromShapeTop.Count -eq 0) { $null } else { [Math]::Round(($refBaselineFromShapeTop | Measure-Object -Maximum).Maximum, 6) }
                MinCandidateFrameTopY = if ($frameTop.Count -eq 0) { $null } else { [Math]::Round(($frameTop | Measure-Object -Minimum).Minimum, 6) }
                MaxCandidateFrameTopY = if ($frameTop.Count -eq 0) { $null } else { [Math]::Round(($frameTop | Measure-Object -Maximum).Maximum, 6) }
                MinCandidateLineTopY = if ($lineTop.Count -eq 0) { $null } else { [Math]::Round(($lineTop | Measure-Object -Minimum).Minimum, 6) }
                MaxCandidateLineTopY = if ($lineTop.Count -eq 0) { $null } else { [Math]::Round(($lineTop | Measure-Object -Maximum).Maximum, 6) }
            }
        }
    )
}

function Measure-NumericRange($items, [string] $propertyName) {
    $values = @($items | ForEach-Object { OptionalDouble $_ $propertyName } | Where-Object { $null -ne $_ })
    if ($values.Count -eq 0) {
        return $null
    }

    return [pscustomobject]@{
        Count = $values.Count
        Min = [Math]::Round(($values | Measure-Object -Minimum).Minimum, 6)
        Max = [Math]::Round(($values | Measure-Object -Maximum).Maximum, 6)
    }
}

function Group-BranchNumericRanges($items, [hashtable] $fields) {
    $branches = @($items | ForEach-Object { BranchKey $_ } | Sort-Object -Unique)
    return @(
        foreach ($fieldName in ($fields.Keys | Sort-Object)) {
            $propertyName = [string]$fields[$fieldName]
            [pscustomobject]@{
                Field = $fieldName
                Property = $propertyName
                Branches = @(
                    foreach ($branch in $branches) {
                        $branchItems = @($items | Where-Object { (BranchKey $_) -eq $branch })
                        $range = Measure-NumericRange $branchItems $propertyName
                        [pscustomobject]@{
                            Branch = $branch
                            Count = if ($null -eq $range) { 0 } else { $range.Count }
                            Min = if ($null -eq $range) { $null } else { $range.Min }
                            Max = if ($null -eq $range) { $null } else { $range.Max }
                        }
                    }
                )
            }
        }
    )
}

function Group-BranchDistinctNumericValues($items, [hashtable] $fields, [int] $limit) {
    $branches = @($items | ForEach-Object { BranchKey $_ } | Sort-Object -Unique)
    return @(
        foreach ($fieldName in ($fields.Keys | Sort-Object)) {
            $propertyName = [string]$fields[$fieldName]
            [pscustomobject]@{
                Field = $fieldName
                Property = $propertyName
                Branches = @(
                    foreach ($branch in $branches) {
                        $valueCounts = @{}
                        foreach ($item in ($items | Where-Object { (BranchKey $_) -eq $branch })) {
                            $value = OptionalDouble $item $propertyName
                            if ($null -eq $value) {
                                continue
                            }

                            $key = RoundedKey $value
                            if ($valueCounts.ContainsKey($key)) {
                                $valueCounts[$key]++
                            }
                            else {
                                $valueCounts[$key] = 1
                            }
                        }

                        $sortedValues = @($valueCounts.Keys | Sort-Object { [double]$_ })
                        [pscustomobject]@{
                            Branch = $branch
                            DistinctCount = $sortedValues.Count
                            Values = @(
                                foreach ($key in ($sortedValues | Select-Object -First $limit)) {
                                    [pscustomobject]@{
                                        Value = [double]$key
                                        Count = $valueCounts[$key]
                                    }
                                }
                            )
                            Truncated = $sortedValues.Count -gt $limit
                        }
                    }
                )
            }
        }
    )
}

function Find-BranchRangeSeparators($items, [hashtable] $fields) {
    $mainItems = @($items | Where-Object { (BranchKey $_) -eq "main-grid" })
    $secondaryBranches = @($items |
        ForEach-Object { BranchKey $_ } |
        Where-Object { $_ -ne "main-grid" -and $_ -ne "(missing)" } |
        Sort-Object -Unique)

    return @(
        foreach ($fieldName in ($fields.Keys | Sort-Object)) {
            $propertyName = [string]$fields[$fieldName]
            $mainRange = Measure-NumericRange $mainItems $propertyName
            if ($null -eq $mainRange) {
                continue
            }

            foreach ($branch in $secondaryBranches) {
                $branchItems = @($items | Where-Object { (BranchKey $_) -eq $branch })
                $branchRange = Measure-NumericRange $branchItems $propertyName
                if ($null -eq $branchRange) {
                    continue
                }

                $overlaps = $mainRange.Min -le $branchRange.Max -and $branchRange.Min -le $mainRange.Max
                [pscustomobject]@{
                    Field = $fieldName
                    Property = $propertyName
                    Branch = $branch
                    MainMin = $mainRange.Min
                    MainMax = $mainRange.Max
                    BranchMin = $branchRange.Min
                    BranchMax = $branchRange.Max
                    RangesOverlap = $overlaps
                }
            }
        }
    )
}

function Find-BranchDistinctValueSeparators($items, [hashtable] $fields) {
    $mainItems = @($items | Where-Object { (BranchKey $_) -eq "main-grid" })
    $secondaryBranches = @($items |
        ForEach-Object { BranchKey $_ } |
        Where-Object { $_ -ne "main-grid" -and $_ -ne "(missing)" } |
        Sort-Object -Unique)

    return @(
        foreach ($fieldName in ($fields.Keys | Sort-Object)) {
            $propertyName = [string]$fields[$fieldName]
            $mainValueSet = @{}
            foreach ($item in $mainItems) {
                $value = OptionalDouble $item $propertyName
                if ($null -eq $value) {
                    continue
                }

                $mainValueSet[(RoundedKey $value)] = $true
            }

            foreach ($branch in $secondaryBranches) {
                $branchItems = @($items | Where-Object { (BranchKey $_) -eq $branch })
                $branchValueCounts = @{}
                foreach ($item in $branchItems) {
                    $value = OptionalDouble $item $propertyName
                    if ($null -eq $value) {
                        continue
                    }

                    $key = RoundedKey $value
                    if ($branchValueCounts.ContainsKey($key)) {
                        $branchValueCounts[$key]++
                    }
                    else {
                        $branchValueCounts[$key] = 1
                    }
                }

                $uniqueValues = @($branchValueCounts.Keys |
                    Where-Object { -not $mainValueSet.ContainsKey($_) } |
                    Sort-Object { [double]$_ })
                [pscustomobject]@{
                    Field = $fieldName
                    Property = $propertyName
                    Branch = $branch
                    MainDistinctCount = $mainValueSet.Count
                    BranchDistinctCount = $branchValueCounts.Count
                    BranchDistinctValuesAbsentFromMainCount = $uniqueValues.Count
                    BranchDistinctValuesAbsentFromMain = @(
                        foreach ($key in ($uniqueValues | Select-Object -First 64)) {
                            [pscustomobject]@{
                                Value = [double]$key
                                Count = $branchValueCounts[$key]
                            }
                        }
                    )
                    Truncated = $uniqueValues.Count -gt 64
                }
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

function SameReferenceCandidatePage($reference, $candidate) {
    $pageNumber = OptionalValue $reference "PageNumber"
    $slideNumber = OptionalValue $candidate "Slide"
    if (-not (HasValue $pageNumber) -or -not (HasValue $slideNumber)) {
        return $true
    }

    return [int]$pageNumber -eq [int]$slideNumber
}

$referenceOps = Read-JsonArray $ReferenceTextOperations
$candidateRuns = Read-JsonArray $CandidateGlyphRuns

if ($MatchByPosition -or $MatchByTextThenPosition) {
    $unmatched = New-Object System.Collections.Generic.List[object]
    foreach ($op in $referenceOps) {
        $unmatched.Add($op)
    }

    $pairs = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $candidateRuns.Count; $i++) {
        $candidate = $candidateRuns[$i]
        $bestIndex = -1
        $bestScore = [double]::PositiveInfinity
        $candidateText = CandText $candidate
        for ($j = 0; $j -lt $unmatched.Count; $j++) {
            $reference = $unmatched[$j]
            if (-not (SameReferenceCandidatePage $reference $candidate)) {
                continue
            }

            if ($MatchByTextThenPosition) {
                $referenceText = RefText $reference
                if ((HasValue $candidateText) -and
                    (HasValue $referenceText) -and
                    -not [string]::Equals($candidateText, $referenceText, [System.StringComparison]::Ordinal)) {
                    continue
                }
            }

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
            RefPageNumber = if ($null -eq $reference) { $null } else { OptionalValue $reference "PageNumber" }
            CandSlide = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "Slide" }
            RefX = if ($null -eq $reference) { $null } else { [Math]::Round((RefX $reference), 6) }
            CandX = if ($null -eq $candidate) { $null } else { [Math]::Round([double]$candidate.X, 6) }
            DeltaX = $null
            RefBaselineY = if ($null -eq $reference) { $null } else { [Math]::Round((RefY $reference), 6) }
            CandBaselineY = if ($null -eq $candidate) { $null } else { [Math]::Round([double]$candidate.BaselineY, 6) }
            DeltaBaselineY = $null
            RefBaselineFromPageTop = RefBaselineFromPageTop $reference $candidate
            CandBaselineFromPageTop = CandidateBaselineFromPageTop $candidate
            RefBaselineFromCandidateShapeTop = RefBaselineFromCandidateShapeTop $reference $candidate
            RefBaselineY600Remainder = if ($null -eq $reference) { $null } else { OfficeGridRemainder (RefY $reference) }
            CandBaselineY600Remainder = if ($null -eq $candidate) { $null } else { OfficeGridRemainder ([double]$candidate.BaselineY) }
            RefBaselineFromPageTop600Remainder = NullableGridRemainder (RefBaselineFromPageTop $reference $candidate)
            CandBaselineFromPageTop600Remainder = NullableGridRemainder (CandidateBaselineFromPageTop $candidate)
            RefFontSize = if ($null -eq $reference) { $null } else { [Math]::Round([double]$reference.FontSize, 6) }
            CandPdfFontSize = if ($null -eq $candidate) { $null } else { [Math]::Round([double]$candidate.PdfFontSize, 6) }
            DeltaFontSize = $null
            CandLayoutFontSize = if ($null -eq $candidate) { $null } else { [Math]::Round([double]$candidate.LayoutFontSize, 6) }
            RefCharacterSpacing = if ($null -eq $reference) { $null } else { [Math]::Round([double]$reference.CharacterSpacing, 6) }
            CandLayoutCharacterSpacing = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "LayoutCharacterSpacing" }
            CandPdfCharacterSpacing = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "PdfCharacterSpacing" }
            DeltaPdfCharacterSpacing = if ($null -eq $reference -or $null -eq $candidate) { $null } else { Delta -left ([double]$reference.CharacterSpacing) -right ([double](OptionalValue $candidate "PdfCharacterSpacing")) }
            MainFontGrid = if ($null -eq $candidate) { $null } else { RoundAway (OfficeMainFontGrid ([double]$candidate.LayoutFontSize)) }
            RefSecondaryFontDelta = if ($null -eq $reference -or $null -eq $candidate) { $null } else { Delta -left (OfficeMainFontGrid ([double]$candidate.LayoutFontSize)) -right ([double]$reference.FontSize) }
            CandPdfGridDelta = if ($null -eq $candidate) { $null } else { Delta -left (OfficeMainFontGrid ([double]$candidate.LayoutFontSize)) -right ([double]$candidate.PdfFontSize) }
            CandFontFamily = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "FontFamily" }
            CandResolvedFontFamily = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "ResolvedFontFamily" }
            CandResolvedFontIsFallback = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "ResolvedFontIsFallback" }
            CandResolvedFontHasMathTable = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "ResolvedFontHasMathTable" }
            CandResolvedFontHasGposTable = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "ResolvedFontHasGposTable" }
            CandResolvedFontHasKernTable = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "ResolvedFontHasKernTable" }
            CandResolvedFontUnitsPerEm = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "ResolvedFontUnitsPerEm" }
            CandResolvedFontOs2WidthClass = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "ResolvedFontOs2WidthClass" }
            CandResolvedFontPostUnderlineThickness = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "ResolvedFontPostUnderlineThickness" }
            CandBold = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "Bold" }
            CandItalic = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "Italic" }
            CandUnderline = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "Underline" }
            CandStrike = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "Strike" }
            CandSyntheticBold = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "SyntheticBold" }
            CandSyntheticItalic = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "SyntheticItalic" }
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
            CandTableRowIndex = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "TableRowIndex" }
            CandTableColumnIndex = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "TableColumnIndex" }
            CandTableRowSpan = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "TableRowSpan" }
            CandTableColumnSpan = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "TableColumnSpan" }
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
            CandFrameWrapMode = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "FrameWrapMode" }
            CandFrameVerticalOverflowMode = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "FrameVerticalOverflowMode" }
            CandFrameVerticalOverflowValue = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "FrameVerticalOverflowValue" }
            CandFrameVerticalOverflowSource = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "FrameVerticalOverflowSource" }
            CandFrameAutofitMode = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "FrameAutofitMode" }
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
            CandInterGlyphAdjustmentCount = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "InterGlyphAdjustmentCount" }
            CandInterGlyphAdjustmentSum = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "InterGlyphAdjustmentSum" }
            CandInterGlyphAdjustmentMin = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "InterGlyphAdjustmentMin" }
            CandInterGlyphAdjustmentMax = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "InterGlyphAdjustmentMax" }
            CandInterGlyphAdjustmentAverage = if ($null -eq $candidate) { $null } else { OptionalRoundedDouble $candidate "InterGlyphAdjustmentAverage" }
            CandLetterCount = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "LetterCount" }
            CandUppercaseLetterCount = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "UppercaseLetterCount" }
            CandLowercaseLetterCount = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "LowercaseLetterCount" }
            CandTitlecaseLetterCount = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "TitlecaseLetterCount" }
            CandDecimalDigitCount = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "DecimalDigitCount" }
            CandPunctuationCount = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "PunctuationCount" }
            CandSymbolCount = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "SymbolCount" }
            CandSpaceCount = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "SpaceCount" }
            CandOtherCount = if ($null -eq $candidate) { $null } else { OptionalValue $candidate "OtherCount" }
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
        RefPageNumber = OptionalValue $reference "PageNumber"
        CandSlide = OptionalValue $candidate "Slide"
        RefX = [Math]::Round($refX, 6)
        CandX = [Math]::Round($candX, 6)
        DeltaX = $deltaX
        RefBaselineY = [Math]::Round($refY, 6)
        CandBaselineY = [Math]::Round($candY, 6)
        DeltaBaselineY = $deltaY
        RefBaselineFromPageTop = RefBaselineFromPageTop $reference $candidate
        CandBaselineFromPageTop = CandidateBaselineFromPageTop $candidate
        RefBaselineFromCandidateShapeTop = RefBaselineFromCandidateShapeTop $reference $candidate
        RefBaselineY600Remainder = OfficeGridRemainder $refY
        CandBaselineY600Remainder = OfficeGridRemainder $candY
        RefBaselineFromPageTop600Remainder = NullableGridRemainder (RefBaselineFromPageTop $reference $candidate)
        CandBaselineFromPageTop600Remainder = NullableGridRemainder (CandidateBaselineFromPageTop $candidate)
        RefFontSize = [Math]::Round($refFontSize, 6)
        CandPdfFontSize = [Math]::Round($candFontSize, 6)
        DeltaFontSize = $deltaFontSize
        CandLayoutFontSize = [Math]::Round([double]$candidate.LayoutFontSize, 6)
        RefCharacterSpacing = [Math]::Round([double]$reference.CharacterSpacing, 6)
        CandLayoutCharacterSpacing = OptionalRoundedDouble $candidate "LayoutCharacterSpacing"
        CandPdfCharacterSpacing = OptionalRoundedDouble $candidate "PdfCharacterSpacing"
        DeltaPdfCharacterSpacing = Delta -left ([double]$reference.CharacterSpacing) -right ([double](OptionalValue $candidate "PdfCharacterSpacing"))
        MainFontGrid = RoundAway $mainFontGrid
        RefSecondaryFontDelta = Delta -left $mainFontGrid -right $refFontSize
        CandPdfGridDelta = Delta -left $mainFontGrid -right $candFontSize
        CandFontFamily = OptionalValue $candidate "FontFamily"
        CandResolvedFontFamily = OptionalValue $candidate "ResolvedFontFamily"
        CandResolvedFontIsFallback = OptionalValue $candidate "ResolvedFontIsFallback"
        CandResolvedFontHasMathTable = OptionalValue $candidate "ResolvedFontHasMathTable"
        CandResolvedFontHasGposTable = OptionalValue $candidate "ResolvedFontHasGposTable"
        CandResolvedFontHasKernTable = OptionalValue $candidate "ResolvedFontHasKernTable"
        CandResolvedFontUnitsPerEm = OptionalValue $candidate "ResolvedFontUnitsPerEm"
        CandResolvedFontOs2WidthClass = OptionalValue $candidate "ResolvedFontOs2WidthClass"
        CandResolvedFontPostUnderlineThickness = OptionalValue $candidate "ResolvedFontPostUnderlineThickness"
        CandBold = OptionalValue $candidate "Bold"
        CandItalic = OptionalValue $candidate "Italic"
        CandUnderline = OptionalValue $candidate "Underline"
        CandStrike = OptionalValue $candidate "Strike"
        CandSyntheticBold = OptionalValue $candidate "SyntheticBold"
        CandSyntheticItalic = OptionalValue $candidate "SyntheticItalic"
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
        CandTableRowIndex = OptionalValue $candidate "TableRowIndex"
        CandTableColumnIndex = OptionalValue $candidate "TableColumnIndex"
        CandTableRowSpan = OptionalValue $candidate "TableRowSpan"
        CandTableColumnSpan = OptionalValue $candidate "TableColumnSpan"
        CandParagraphIndex = OptionalValue $candidate "ParagraphIndex"
        CandSourceRunIndex = OptionalValue $candidate "SourceRunIndex"
        CandParagraphBulletKind = OptionalValue $candidate "ParagraphBulletKind"
        CandParagraphAutoNumberType = OptionalValue $candidate "ParagraphAutoNumberType"
        CandParagraphAutoNumberStartAt = OptionalValue $candidate "ParagraphAutoNumberStartAt"
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
        CandFrameWrapMode = OptionalValue $candidate "FrameWrapMode"
        CandFrameVerticalOverflowMode = OptionalValue $candidate "FrameVerticalOverflowMode"
        CandFrameVerticalOverflowValue = OptionalValue $candidate "FrameVerticalOverflowValue"
        CandFrameVerticalOverflowSource = OptionalValue $candidate "FrameVerticalOverflowSource"
        CandFrameAutofitMode = OptionalValue $candidate "FrameAutofitMode"
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
        CandInterGlyphAdjustmentCount = OptionalValue $candidate "InterGlyphAdjustmentCount"
        CandInterGlyphAdjustmentSum = OptionalRoundedDouble $candidate "InterGlyphAdjustmentSum"
        CandInterGlyphAdjustmentMin = OptionalRoundedDouble $candidate "InterGlyphAdjustmentMin"
        CandInterGlyphAdjustmentMax = OptionalRoundedDouble $candidate "InterGlyphAdjustmentMax"
        CandInterGlyphAdjustmentAverage = OptionalRoundedDouble $candidate "InterGlyphAdjustmentAverage"
        CandLetterCount = OptionalValue $candidate "LetterCount"
        CandUppercaseLetterCount = OptionalValue $candidate "UppercaseLetterCount"
        CandLowercaseLetterCount = OptionalValue $candidate "LowercaseLetterCount"
        CandTitlecaseLetterCount = OptionalValue $candidate "TitlecaseLetterCount"
        CandDecimalDigitCount = OptionalValue $candidate "DecimalDigitCount"
        CandPunctuationCount = OptionalValue $candidate "PunctuationCount"
        CandSymbolCount = OptionalValue $candidate "SymbolCount"
        CandSpaceCount = OptionalValue $candidate "SpaceCount"
        CandOtherCount = OptionalValue $candidate "OtherCount"
        RefText = if ($IncludeText) { RefText $reference } else { $null }
        CandText = if ($IncludeText) { CandText $candidate } else { $null }
    })
}

$rowsArray = $rows.ToArray()

if (HasValue $OutputJson) {
    $rowsArray | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $OutputJson -Encoding UTF8
}

if (HasValue $OutputSummaryJson) {
    $numericBranchFields = [ordered]@{
        CandidateBaselineFromPageTop = "CandBaselineFromPageTop"
        CandidateBaselineFromShapeTop = "CandBaselineFromShapeTop"
        CandidateBaselineY = "CandBaselineY"
        CandidateFrameHeight = "CandFrameShapeHeight"
        CandidateFrameTopY = "CandFrameShapeTopY"
        CandidateLineTopFromShapeTop = "CandLineTopFromShapeTop"
        CandidateLineTopFromTextTop = "CandLineTopFromTextTop"
        CandidateLineTopY = "CandLineTopY"
        CandidateTextHeight = "CandFrameTextHeight"
        CandidateTextWidth = "CandFrameTextWidth"
        RefBaselineFromCandidateShapeTop = "RefBaselineFromCandidateShapeTop"
        RefBaselineFromPageTop = "RefBaselineFromPageTop"
        RefBaselineY = "RefBaselineY"
    }

    $summary = [pscustomobject]@{
        Total = $rowsArray.Count
        StatusCounts = Group-Count $rowsArray { param($row) $row.Status }
        FontBranchCounts = Group-Count $rowsArray { param($row) BranchKey $row }
        FontBranchExtents = Group-BranchExtents $rowsArray
        FontBranchNumericRanges = Group-BranchNumericRanges $rowsArray $numericBranchFields
        FontBranchDistinctNumericValues = Group-BranchDistinctNumericValues $rowsArray $numericBranchFields 64
        FontBranchRangeSeparators = Find-BranchRangeSeparators $rowsArray $numericBranchFields
        FontBranchDistinctValueSeparators = Find-BranchDistinctValueSeparators $rowsArray $numericBranchFields
        RefSecondaryFontDeltas = Group-Count $rowsArray { param($row) RoundedKey $row.RefSecondaryFontDelta }
        RefCharacterSpacingCounts = Group-Count $rowsArray { param($row) RoundedKey $row.RefCharacterSpacing }
        PdfCharacterSpacingDeltas = Group-Count $rowsArray { param($row) RoundedKey $row.DeltaPdfCharacterSpacing }
        ByPdfCharacterSpacingAndBranch = Group-Count $rowsArray { param($row) (RoundedKey $row.CandPdfCharacterSpacing) + "|" + (BranchKey $row) }
        ByLayoutFontSizeAndBranch = Group-Count $rowsArray { param($row) (RoundedKey $row.CandLayoutFontSize) + "|" + (BranchKey $row) }
        ByCandidateFontFamilyAndBranch = Group-Count $rowsArray { param($row) (StringKey (OptionalValue $row "CandFontFamily")) + "|" + (BranchKey $row) }
        ByCandidateStyleAndBranch = Group-Count $rowsArray { param($row) (StringKey (OptionalValue $row "CandBold")) + "," + (StringKey (OptionalValue $row "CandItalic")) + "," + (StringKey (OptionalValue $row "CandSyntheticBold")) + "," + (StringKey (OptionalValue $row "CandSyntheticItalic")) + "|" + (BranchKey $row) }
        ByCandidateTableRowAndBranch = Group-Count $rowsArray { param($row) (RoundedKey (OptionalValue $row "CandTableRowIndex")) + "|" + (BranchKey $row) }
        ByCandidateTableColumnAndBranch = Group-Count $rowsArray { param($row) (RoundedKey (OptionalValue $row "CandTableColumnIndex")) + "|" + (BranchKey $row) }
        ByCandidateTableCellAndBranch = Group-Count $rowsArray { param($row) (RoundedKey (OptionalValue $row "CandTableRowIndex")) + "," + (RoundedKey (OptionalValue $row "CandTableColumnIndex")) + "|" + (BranchKey $row) }
        ByParagraphIndexAndBranch = Group-Count $rowsArray { param($row) (RoundedKey $row.CandParagraphIndex) + "|" + (BranchKey $row) }
        BySourceRunIndexAndBranch = Group-Count $rowsArray { param($row) (RoundedKey (OptionalValue $row "CandSourceRunIndex")) + "|" + (BranchKey $row) }
        ByParagraphBulletKindAndBranch = Group-Count $rowsArray { param($row) (StringKey (OptionalValue $row "CandParagraphBulletKind")) + "|" + (BranchKey $row) }
        ByParagraphAutoNumberTypeAndBranch = Group-Count $rowsArray { param($row) (StringKey (OptionalValue $row "CandParagraphAutoNumberType")) + "|" + (BranchKey $row) }
        ByParagraphAutoNumberStartAtAndBranch = Group-Count $rowsArray { param($row) (RoundedKey (OptionalValue $row "CandParagraphAutoNumberStartAt")) + "|" + (BranchKey $row) }
        BySpanIndexAndBranch = Group-Count $rowsArray { param($row) (RoundedKey $row.CandSpanIndex) + "|" + (BranchKey $row) }
        ByRefBaselineRemainderAndBranch = Group-Count $rowsArray { param($row) (RoundedKey $row.RefBaselineY600Remainder) + "|" + (BranchKey $row) }
        ByCandidateBaselineRemainderAndBranch = Group-Count $rowsArray { param($row) (RoundedKey $row.CandBaselineY600Remainder) + "|" + (BranchKey $row) }
        ByRefBaselineFromPageTopRemainderAndBranch = Group-Count $rowsArray { param($row) (RoundedKey $row.RefBaselineFromPageTop600Remainder) + "|" + (BranchKey $row) }
        ByCandidateBaselineFromPageTopRemainderAndBranch = Group-Count $rowsArray { param($row) (RoundedKey $row.CandBaselineFromPageTop600Remainder) + "|" + (BranchKey $row) }
        ByRefBaselineFromCandidateShapeTopAndBranch = Group-Count $rowsArray { param($row) (RoundedKey $row.RefBaselineFromCandidateShapeTop) + "|" + (BranchKey $row) }
        ByCandidateFrameXAndBranch = Group-Count $rowsArray { param($row) (RoundedKey $row.CandFrameShapeX) + "|" + (BranchKey $row) }
        ByCandidateFrameTopRemainderAndBranch = Group-Count $rowsArray { param($row) (RoundedKey $row.CandFrameShapeTopY600Remainder) + "|" + (BranchKey $row) }
        ByCandidateFrameWidthAndBranch = Group-Count $rowsArray { param($row) (RoundedKey $row.CandFrameShapeWidth) + "|" + (BranchKey $row) }
        ByCandidateWrapModeAndBranch = Group-Count $rowsArray { param($row) (OptionalValue $row "CandFrameWrapMode") + "|" + (BranchKey $row) }
        ByCandidateVerticalOverflowAndBranch = Group-Count $rowsArray { param($row) (OptionalValue $row "CandFrameVerticalOverflowMode") + "|" + (BranchKey $row) }
        ByCandidateVerticalOverflowSourceAndBranch = Group-Count $rowsArray { param($row) (OptionalValue $row "CandFrameVerticalOverflowSource") + "|" + (BranchKey $row) }
        ByCandidateAutofitModeAndBranch = Group-Count $rowsArray { param($row) (OptionalValue $row "CandFrameAutofitMode") + "|" + (BranchKey $row) }
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
        ByCandidateInterGlyphAdjustmentCountAndBranch = Group-Count $rowsArray { param($row) (RoundedKey (OptionalValue $row "CandInterGlyphAdjustmentCount")) + "|" + (BranchKey $row) }
        ByCandidateInterGlyphAdjustmentAverageAndBranch = Group-Count $rowsArray { param($row) (RoundedKey (OptionalValue $row "CandInterGlyphAdjustmentAverage")) + "|" + (BranchKey $row) }
        ByCandidateLetterCountAndBranch = Group-Count $rowsArray { param($row) (RoundedKey (OptionalValue $row "CandLetterCount")) + "|" + (BranchKey $row) }
        ByCandidateUppercaseLetterCountAndBranch = Group-Count $rowsArray { param($row) (RoundedKey (OptionalValue $row "CandUppercaseLetterCount")) + "|" + (BranchKey $row) }
        ByCandidateLowercaseLetterCountAndBranch = Group-Count $rowsArray { param($row) (RoundedKey (OptionalValue $row "CandLowercaseLetterCount")) + "|" + (BranchKey $row) }
        ByCandidateTitlecaseLetterCountAndBranch = Group-Count $rowsArray { param($row) (RoundedKey (OptionalValue $row "CandTitlecaseLetterCount")) + "|" + (BranchKey $row) }
        ByCandidateDecimalDigitCountAndBranch = Group-Count $rowsArray { param($row) (RoundedKey (OptionalValue $row "CandDecimalDigitCount")) + "|" + (BranchKey $row) }
        ByCandidatePunctuationCountAndBranch = Group-Count $rowsArray { param($row) (RoundedKey (OptionalValue $row "CandPunctuationCount")) + "|" + (BranchKey $row) }
        ByCandidateSymbolCountAndBranch = Group-Count $rowsArray { param($row) (RoundedKey (OptionalValue $row "CandSymbolCount")) + "|" + (BranchKey $row) }
        ByCandidateSpaceCountAndBranch = Group-Count $rowsArray { param($row) (RoundedKey (OptionalValue $row "CandSpaceCount")) + "|" + (BranchKey $row) }
        ByCandidateOtherCountAndBranch = Group-Count $rowsArray { param($row) (RoundedKey (OptionalValue $row "CandOtherCount")) + "|" + (BranchKey $row) }
        ByCandidateContentShapeAndBranch = Group-Count $rowsArray { param($row) (RoundedKey (OptionalValue $row "CandLetterCount")) + "," + (RoundedKey (OptionalValue $row "CandDecimalDigitCount")) + "," + (RoundedKey (OptionalValue $row "CandPunctuationCount")) + "," + (RoundedKey (OptionalValue $row "CandSymbolCount")) + "," + (RoundedKey (OptionalValue $row "CandSpaceCount")) + "," + (RoundedKey (OptionalValue $row "CandOtherCount")) + "|" + (BranchKey $row) }
        ByCandidateLetterCaseShapeAndBranch = Group-Count $rowsArray { param($row) (RoundedKey (OptionalValue $row "CandUppercaseLetterCount")) + "," + (RoundedKey (OptionalValue $row "CandLowercaseLetterCount")) + "," + (RoundedKey (OptionalValue $row "CandTitlecaseLetterCount")) + "|" + (BranchKey $row) }
    }

    $summary | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $OutputSummaryJson -Encoding UTF8
}

$rowsArray | Format-Table -AutoSize
Write-Host "Text emission count: reference=$($referenceOps.Count), candidate=$($candidateRuns.Count), deltas=$failures"
if ($MatchByPosition) {
    Write-Host "Matching: nearest text position"
}
elseif ($MatchByTextThenPosition) {
    Write-Host "Matching: exact text, then nearest text position"
}
else {
    Write-Host "Matching: input order"
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
