param(
    [Parameter(Mandatory = $true)]
    [string[]] $RunDirectory,

    [string] $OutputJson
)

$ErrorActionPreference = "Stop"
$RepoRoot = Split-Path -Parent $PSScriptRoot

function Expand-PathList([string[]] $Values) {
    $expanded = New-Object System.Collections.Generic.List[string]
    foreach ($value in $Values) {
        if ([string]::IsNullOrWhiteSpace($value)) {
            continue
        }

        foreach ($part in ($value -split "[,;]")) {
            if (-not [string]::IsNullOrWhiteSpace($part)) {
                $expanded.Add($part.Trim())
            }
        }
    }

    return ,$expanded.ToArray()
}

function Read-JsonArray([string] $Path) {
    $items = Get-Content -Raw -LiteralPath (Resolve-Path -LiteralPath $Path).Path | ConvertFrom-Json
    if ($null -eq $items) {
        return ,@()
    }

    if ($items -is [array]) {
        return ,$items
    }

    return ,@($items)
}

function RoundedKey($Value, [int] $Digits) {
    if ($null -eq $Value -or [string]$Value -eq "") {
        return "(missing)"
    }

    return ([Math]::Round([double]$Value, $Digits)).ToString("0.######", [Globalization.CultureInfo]::InvariantCulture)
}

function SubtractOrNull($Left, $Right) {
    if ($null -eq $Left -or [string]$Left -eq "" -or $null -eq $Right -or [string]$Right -eq "") {
        return $null
    }

    return [double]$Left - [double]$Right
}

function DivideOrNull($Numerator, $Denominator) {
    if ($null -eq $Numerator -or [string]$Numerator -eq "" -or $null -eq $Denominator -or [string]$Denominator -eq "") {
        return $null
    }

    $denominatorValue = [double]$Denominator
    if ([Math]::Abs($denominatorValue) -lt 0.0000001d) {
        return $null
    }

    return [double]$Numerator / $denominatorValue
}

function Test-NumericClose($Left, $Right, [double] $Tolerance) {
    if ($null -eq $Left -or [string]$Left -eq "" -or $null -eq $Right -or [string]$Right -eq "") {
        return $false
    }

    return [Math]::Abs([double]$Left - [double]$Right) -le $Tolerance
}

function Group-Count($Items, [scriptblock] $KeySelector) {
    $groups = @{}
    foreach ($item in $Items) {
        $key = & $KeySelector $item
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

function New-TcTargetHypothesisReport(
    [string] $Name,
    $Pairs,
    [scriptblock] $CandidateSelector,
    [bool] $OnlyNonzeroTc = $true,
    [double] $Tolerance = 0.001d) {
    $candidatePairs = New-Object System.Collections.Generic.List[object]
    foreach ($pair in $Pairs) {
        if ($OnlyNonzeroTc -and [Math]::Abs([double]$pair.ReferenceTc) -le $Tolerance) {
            continue
        }

        $candidate = & $CandidateSelector $pair
        if ($null -eq $candidate -or [string]$candidate -eq "") {
            continue
        }

        $candidatePairs.Add([pscustomobject]@{
            Pair = $pair
            CandidateTc = [double]$candidate
        })
    }

    $matches = @($candidatePairs | Where-Object { Test-NumericClose $_.Pair.ReferenceTc $_.CandidateTc $Tolerance })
    $mismatches = @($candidatePairs | Where-Object { -not (Test-NumericClose $_.Pair.ReferenceTc $_.CandidateTc $Tolerance) })
    $pairCount = if ($OnlyNonzeroTc) {
        @($Pairs | Where-Object { [Math]::Abs([double]$_.ReferenceTc) -gt $Tolerance }).Count
    }
    else {
        @($Pairs).Count
    }

    return [pscustomobject]@{
        Name = $Name
        OnlyNonzeroTc = $OnlyNonzeroTc
        Tolerance = $Tolerance
        PairCount = $pairCount
        ComparablePairCount = $candidatePairs.Count
        MatchCount = $matches.Count
        MismatchCount = $mismatches.Count
        MatchRatio = if ($candidatePairs.Count -eq 0) { $null } else { [Math]::Round($matches.Count / $candidatePairs.Count, 6) }
        MismatchExamples = @($mismatches | Select-Object -First 12 | ForEach-Object {
            [pscustomobject]@{
                ReferenceTc = [Math]::Round([double]$_.Pair.ReferenceTc, 6)
                CandidateTc = [Math]::Round([double]$_.CandidateTc, 6)
                PlannerPdfFontSize = $_.Pair.PlannerPdfFontSize
                PlannerGlyphGapCount = $_.Pair.PlannerGlyphGapCount
                PlannerTextClass = $_.Pair.PlannerTextClass
                PlannerTextLength = $_.Pair.PlannerTextLength
                PlannerRole = $_.Pair.PlannerRole
                PlannerPdfCharacterSpacingSource = $_.Pair.PlannerPdfCharacterSpacingSource
            }
        })
    }
}

function New-TcAmbiguityReport(
    [string] $Name,
    $Pairs,
    [scriptblock] $KeySelector,
    [bool] $OnlyNonzeroTc = $false) {
    $groups = @{}
    foreach ($pair in $Pairs) {
        if ($OnlyNonzeroTc -and [Math]::Abs([double]$pair.ReferenceTc) -le 0.001d) {
            continue
        }

        $key = & $KeySelector $pair
        if ($null -eq $key -or [string]$key -eq "") {
            $key = "(missing)"
        }

        $key = [string]$key
        if (-not $groups.ContainsKey($key)) {
            $groups[$key] = [pscustomobject]@{
                Key = $key
                Count = 0
                TcValues = New-Object System.Collections.Generic.HashSet[string]
            }
        }

        $group = $groups[$key]
        $group.Count++
        [void]$group.TcValues.Add((RoundedKey $pair.ReferenceTc 6))
    }

    $ambiguous = @($groups.Values | Where-Object { $_.TcValues.Count -gt 1 } | Sort-Object Count -Descending)
    return [pscustomobject]@{
        Name = $Name
        OnlyNonzeroTc = $OnlyNonzeroTc
        PairCount = ($groups.Values | Measure-Object -Property Count -Sum).Sum
        KeyCount = $groups.Count
        AmbiguousKeyCount = $ambiguous.Count
        AmbiguousExamples = @($ambiguous | Select-Object -First 12 | ForEach-Object {
            [pscustomobject]@{
                Key = $_.Key
                Count = $_.Count
                TcValues = @($_.TcValues | Sort-Object)
            }
        })
    }
}

function Select-TcStructuralDiscriminators($Reports) {
    return @(
        $Reports |
            Where-Object { [bool]$_.OnlyNonzeroTc -and [int]$_.PairCount -gt 0 -and [int]$_.AmbiguousKeyCount -eq 0 } |
            Sort-Object @{ Expression = { [int]$_.KeyCount }; Ascending = $true }, Name |
            ForEach-Object {
                [pscustomobject]@{
                    Name = $_.Name
                    PairCount = $_.PairCount
                    KeyCount = $_.KeyCount
                }
            }
    )
}

function New-TcTargetAdvanceModelReport($Pairs, $TargetHypotheses) {
    $nonzeroReferencePairs = @($Pairs | Where-Object { [Math]::Abs([double]$_.ReferenceTc) -gt 0.001d })
    $withGlyphGaps = @($nonzeroReferencePairs | Where-Object { $null -ne $_.PlannerGlyphGapCount -and [int]$_.PlannerGlyphGapCount -gt 0 })
    $withReferenceEmittedAdvance = @($withGlyphGaps | Where-Object { $null -ne $_.ReferenceEmittedAdvance -and [string]$_.ReferenceEmittedAdvance -ne "" })
    $withPlannerEmittedAdvance = @($withGlyphGaps | Where-Object { $null -ne $_.PlannerEmittedAdvance -and [string]$_.PlannerEmittedAdvance -ne "" })
    $withPlannerRoundedAdvance = @($withGlyphGaps | Where-Object { $null -ne $_.PlannerRoundedWidth -and [string]$_.PlannerRoundedWidth -ne "" })
    $oracleHypotheses = @($TargetHypotheses | Where-Object { $_.Name -like "reference-emitted-*" })
    $candidateInputHypotheses = @($TargetHypotheses | Where-Object { $_.Name -notlike "reference-emitted-*" })

    return [pscustomobject]@{
        NonzeroReferenceTcPairCount = $nonzeroReferencePairs.Count
        NonzeroReferenceTcPairsWithGlyphGaps = $withGlyphGaps.Count
        NonzeroReferenceTcPairsWithReferenceEmittedAdvance = $withReferenceEmittedAdvance.Count
        NonzeroReferenceTcPairsWithPlannerEmittedAdvance = $withPlannerEmittedAdvance.Count
        NonzeroReferenceTcPairsWithPlannerRoundedAdvance = $withPlannerRoundedAdvance.Count
        OracleTargetAdvanceHypotheses = $oracleHypotheses
        CandidateInputHypotheses = $candidateInputHypotheses
        RendererReady = $false
        RendererBlocker = "Word Tc is explained only when the reference emitted advance is available; candidate-only width residuals remain mismatched or ambiguous."
    }
}

function TextLength($Operation) {
    if ($null -eq $Operation.DecodedText) {
        return 0
    }

    return ([string]$Operation.DecodedText).Length
}

function TextClass($Operation) {
    if ($null -eq $Operation.DecodedText) {
        return "(missing)"
    }

    $text = [string]$Operation.DecodedText
    if ($text.Length -eq 0) {
        return "empty"
    }

    if ($text -match '^\s+$') {
        return "whitespace"
    }

    $letters = [regex]::IsMatch($text, '\p{L}')
    $digits = [regex]::IsMatch($text, '\p{Nd}')
    $other = [regex]::IsMatch($text, '[^\p{L}\p{Nd}\s]')
    if ($letters -and -not $digits -and -not $other) {
        return "letters"
    }

    if ($digits -and -not $letters -and -not $other) {
        return "digits"
    }

    if ($letters -and $digits -and -not $other) {
        return "alphanumeric"
    }

    if ($other -and -not $letters -and -not $digits) {
        return "punctuation"
    }

    return "mixed"
}

function Test-PdfTextOperationsFresh([string] $Path) {
    if (-not (Test-Path -LiteralPath $Path)) {
        return $false
    }

    $operations = Read-JsonArray $Path
    if ($operations.Count -eq 0) {
        return $true
    }

    foreach ($operation in $operations) {
        if ($null -ne $operation.WidthSignature) {
            return $true
        }
    }

    return $false
}

function Ensure-TextOperations([string] $Run, [string] $Side) {
    $runLocal = Join-Path $Run ("comparison\pdf-text\" + $Side + "\text-operations.json")
    if (Test-PdfTextOperationsFresh $runLocal) {
        return $runLocal
    }

    $pdf = if ($Side -eq "reference") {
        Join-Path $Run "reference\reference.pdf"
    }
    else {
        Join-Path $Run "candidate\output.pdf"
    }

    if (-not (Test-Path -LiteralPath $pdf)) {
        throw "Missing $Side PDF for run '$Run': $pdf"
    }

    $runInfo = Get-Item -LiteralPath $Run
    $caseId = Split-Path -Leaf (Split-Path -Parent $runInfo.FullName)
    $runId = Split-Path -Leaf $runInfo.FullName
    $safeId = ($caseId + "-" + $runId + "-" + $Side) -replace "[^A-Za-z0-9_.-]", "_"
    $out = Join-Path "artifacts\docx-text-state" $safeId
    if (-not (Test-Path -LiteralPath $out)) {
        New-Item -ItemType Directory -Path $out | Out-Null
    }

    $json = Join-Path $out "text-operations.json"
    if (-not (Test-PdfTextOperationsFresh $json)) {
        pwsh tools\InspectPdf.ps1 -InputPdf $pdf -OutputDirectory $out -TextOnly | Out-Null
    }

    return $json
}

function Get-RunCaseId([string] $Run) {
    return Split-Path -Leaf (Split-Path -Parent $Run)
}

function Resolve-CaseManifestInput([string] $ManifestPath) {
    if (-not (Test-Path -LiteralPath $manifestPath)) {
        return $null
    }

    $manifest = Get-Content -Raw -LiteralPath $manifestPath | ConvertFrom-Json
    if ($null -eq $manifest.input -or [string]::IsNullOrWhiteSpace([string]$manifest.input)) {
        return $null
    }

    $input = Join-Path (Split-Path -Parent $manifestPath) ([string]$manifest.input)
    if (Test-Path -LiteralPath $input) {
        return (Resolve-Path -LiteralPath $input).Path
    }

    return $null
}

function Resolve-VisualCaseInput([string] $CaseId) {
    $manifestPath = Join-Path $RepoRoot ("visual-cases\cases\" + $CaseId + "\case.json")
    return Resolve-CaseManifestInput $manifestPath
}

function Resolve-PrivateCaseInput([string] $CaseId) {
    $manifestPath = Join-Path $RepoRoot ("private-cases\" + $CaseId + ".json")
    return Resolve-CaseManifestInput $manifestPath
}

function Resolve-CaseInput([string] $CaseId) {
    $input = Resolve-VisualCaseInput $CaseId
    if ($null -ne $input) {
        return $input
    }

    return Resolve-PrivateCaseInput $CaseId
}

function Test-DocxInspectSnapshotFresh([string] $SnapshotPath) {
    if (-not (Test-Path -LiteralPath $SnapshotPath)) {
        return $false
    }

    $snapshot = Get-Content -Raw -LiteralPath $SnapshotPath | ConvertFrom-Json
    if ($null -eq $snapshot -or $null -eq $snapshot.Lines) {
        return $false
    }

    foreach ($line in $snapshot.Lines) {
        if ($null -eq $line.Segments) {
            continue
        }

        foreach ($segment in $line.Segments) {
            if ($null -ne $segment.AdvanceProfile -and
                $null -ne $segment.AdvanceProfile.PlannedEmittedAdvance -and
                $null -ne $segment.PdfCharacterSpacingSource) {
                return $true
            }
        }
    }

    return $false
}

function Ensure-CandidateDocxInspect([string] $Run) {
    $output = Join-Path $Run "comparison\docx-inspect"
    $snapshot = Join-Path $output "text-emission-snapshot.json"
    $summary = Join-Path $output "text-emission-summary.json"
    if ((Test-Path -LiteralPath $summary) -and (Test-DocxInspectSnapshotFresh $snapshot)) {
        return
    }

    $caseId = Get-RunCaseId $Run
    $input = Resolve-CaseInput $caseId
    if ($null -eq $input) {
        return
    }

    & pwsh (Join-Path $RepoRoot "tools\InspectDocx.ps1") -InputDocx $input -OutputDirectory $output | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "InspectDocx failed for case '$caseId' with exit code $LASTEXITCODE."
    }
}

function Summarize-Operations($Operations) {
    $nonzeroTc = @($Operations | Where-Object { [Math]::Abs([double]$_.CharacterSpacing) -gt 0.001d })
    return [pscustomobject]@{
        OperationCount = $Operations.Count
        NonzeroTcCount = $nonzeroTc.Count
        OperatorBuckets = @(Group-Count $Operations { param($op) [string]$op.Operator })
        TcBuckets = @(Group-Count $Operations { param($op) RoundedKey $op.CharacterSpacing 6 })
        FontSizeBuckets = @(Group-Count $Operations { param($op) RoundedKey $op.FontSize 3 })
        FontSizeByTc = @(Group-Count $Operations {
            param($op)
            (RoundedKey $op.FontSize 3) + "|tc=" + (RoundedKey $op.CharacterSpacing 6)
        })
        TextLengthByTc = @(Group-Count $Operations {
            param($op)
            "len=" + (TextLength $op) + "|tc=" + (RoundedKey $op.CharacterSpacing 6)
        })
        TextClassByTc = @(Group-Count $Operations {
            param($op)
            (TextClass $op) + "|tc=" + (RoundedKey $op.CharacterSpacing 6)
        })
        NonzeroTcByFontSize = @(Group-Count $nonzeroTc {
            param($op)
            (RoundedKey $op.FontSize 3) + "|tc=" + (RoundedKey $op.CharacterSpacing 6)
        })
        NonzeroTcByTextClass = @(Group-Count $nonzeroTc {
            param($op)
            TextClass $op
        })
        NetAverageSpacingByTc = @(Group-Count $Operations {
            param($op)
            "net=" + (RoundedKey $op.NetAverageCharacterSpacing 6) + "|tc=" + (RoundedKey $op.CharacterSpacing 6)
        })
        TextChunkCountByTc = @(Group-Count $Operations {
            param($op)
            "chunks=" + (RoundedKey $op.TextChunkCount 0) + "|tc=" + (RoundedKey $op.CharacterSpacing 6)
        })
        AdjustmentCountByTc = @(Group-Count $Operations {
            param($op)
            "adj=" + (RoundedKey $op.AdjustmentCount 0) + "|tc=" + (RoundedKey $op.CharacterSpacing 6)
        })
        AverageAdjustmentByTc = @(Group-Count $Operations {
            param($op)
            "avgAdj=" + (RoundedKey $op.AverageAdjustmentPoints 6) + "|tc=" + (RoundedKey $op.CharacterSpacing 6)
        })
        CharacterSpacingGapTotalByTc = @(Group-Count $Operations {
            param($op)
            "tcGapTotal=" + (RoundedKey $op.CharacterSpacingGapTotalPoints 6) + "|tc=" + (RoundedKey $op.CharacterSpacing 6)
        })
        AdjustmentTotalByTc = @(Group-Count $Operations {
            param($op)
            "adjTotal=" + (RoundedKey $op.AdjustmentTotalPoints 6) + "|tc=" + (RoundedKey $op.CharacterSpacing 6)
        })
        NetSpacingGapTotalByTc = @(Group-Count $Operations {
            param($op)
            "netGapTotal=" + (RoundedKey $op.NetSpacingGapTotalPoints 6) + "|tc=" + (RoundedKey $op.CharacterSpacing 6)
        })
        NaturalWidthByTc = @(Group-Count $Operations {
            param($op)
            "natural=" + (RoundedKey $op.NaturalWidthPoints 6) + "|tc=" + (RoundedKey $op.CharacterSpacing 6)
        })
        EmittedAdvanceByTc = @(Group-Count $Operations {
            param($op)
            "emitted=" + (RoundedKey $op.EmittedAdvancePoints 6) + "|tc=" + (RoundedKey $op.CharacterSpacing 6)
        })
    }
}

function New-TextClassCountMap($Items, [scriptblock] $ClassSelector) {
    $map = @{}
    foreach ($item in $Items) {
        $key = & $ClassSelector $item
        if ($null -eq $key -or [string]$key -eq "") {
            $key = "(missing)"
        }

        $key = [string]$key
        if ($map.ContainsKey($key)) {
            $map[$key]++
        }
        else {
            $map[$key] = 1
        }
    }

    return $map
}

function New-TextClassCountDeltas($ReferenceOperations, $CandidateOperations, $PlannerSnapshot) {
    $plannerSegments = Flatten-PlannerSegments $PlannerSnapshot
    $referenceMap = New-TextClassCountMap $ReferenceOperations { param($op) TextClass $op }
    $candidateMap = New-TextClassCountMap $CandidateOperations { param($op) TextClass $op }
    $plannerMap = New-TextClassCountMap $plannerSegments { param($segment) PlannerTextClass $segment }
    $keys = New-Object System.Collections.Generic.HashSet[string]
    foreach ($key in $referenceMap.Keys) { [void]$keys.Add($key) }
    foreach ($key in $candidateMap.Keys) { [void]$keys.Add($key) }
    foreach ($key in $plannerMap.Keys) { [void]$keys.Add($key) }

    return @(
        foreach ($key in ($keys | Sort-Object)) {
            $referenceCount = if ($referenceMap.ContainsKey($key)) { $referenceMap[$key] } else { 0 }
            $candidateCount = if ($candidateMap.ContainsKey($key)) { $candidateMap[$key] } else { 0 }
            $plannerCount = if ($plannerMap.ContainsKey($key)) { $plannerMap[$key] } else { 0 }
            [pscustomobject]@{
                Key = $key
                Reference = $referenceCount
                Candidate = $candidateCount
                Planner = $plannerCount
                CandidateMinusReference = $candidateCount - $referenceCount
                PlannerMinusReference = $plannerCount - $referenceCount
            }
        }
    )
}

function Read-CandidatePlannerSummary([string] $Run) {
    $candidates = @(
        (Join-Path $Run "comparison\docx-inspect\text-emission-summary.json"),
        (Join-Path $Run "docx-inspect\text-emission-summary.json")
    )

    foreach ($path in $candidates) {
        if (Test-Path -LiteralPath $path) {
            return Get-Content -Raw -LiteralPath $path | ConvertFrom-Json
        }
    }

    return $null
}

function Read-CandidatePlannerSnapshot([string] $Run) {
    $candidates = @(
        (Join-Path $Run "comparison\docx-inspect\text-emission-snapshot.json"),
        (Join-Path $Run "docx-inspect\text-emission-snapshot.json")
    )

    foreach ($path in $candidates) {
        if (Test-Path -LiteralPath $path) {
            return Get-Content -Raw -LiteralPath $path | ConvertFrom-Json
        }
    }

    return $null
}

function PlannerTextClass($Segment) {
    if ($null -eq $Segment -or $null -eq $Segment.CharacterProfile) {
        return "(missing)"
    }

    $profile = $Segment.CharacterProfile
    $length = [int]$Segment.TextLength
    if ($length -eq 0) {
        return "empty"
    }

    if ([int]$profile.WhitespaceCount -eq $length) {
        return "whitespace"
    }

    $letters = [int]$profile.LetterCount -gt 0
    $digits = [int]$profile.DigitCount -gt 0
    $other = ([int]$profile.PunctuationCount + [int]$profile.SymbolCount + [int]$profile.OtherCount) -gt 0
    if ($letters -and -not $digits -and -not $other) {
        return "letters"
    }

    if ($digits -and -not $letters -and -not $other) {
        return "digits"
    }

    if ($letters -and $digits -and -not $other) {
        return "alphanumeric"
    }

    if ($other -and -not $letters -and -not $digits) {
        return "punctuation"
    }

    return "mixed"
}

function PlannerRunProvenanceKey($Segment) {
    if ($null -eq $Segment) {
        return "runProv=(missing)"
    }

    return "charFound=" + ([bool]$Segment.CharacterStyleFound).ToString().ToLowerInvariant() +
        "|charDepth=" + (RoundedKey $Segment.CharacterStyleDepth 0) +
        "|docDefault=" + ([bool]$Segment.HasDocumentDefaultRunProperties).ToString().ToLowerInvariant() +
        "|paraStyle=" + ([bool]$Segment.HasParagraphStyleRunProperties).ToString().ToLowerInvariant() +
        "|charStyle=" + ([bool]$Segment.HasCharacterStyleRunProperties).ToString().ToLowerInvariant() +
        "|direct=" + ([bool]$Segment.HasDirectRunProperties).ToString().ToLowerInvariant() +
        "|tableStyle=" + ([bool]$Segment.HasTableStyleRunProperties).ToString().ToLowerInvariant()
}

function Flatten-PlannerSegments($Snapshot) {
    if ($null -eq $Snapshot -or $null -eq $Snapshot.Lines) {
        return ,@()
    }

    $segments = New-Object System.Collections.Generic.List[object]
    foreach ($line in $Snapshot.Lines) {
        if ($null -eq $line.Segments) {
            continue
        }

        foreach ($segment in $line.Segments) {
            $segments.Add($segment)
        }
    }

    return ,$segments.ToArray()
}

function Group-PlannerAdvance($Segments, [scriptblock] $KeySelector) {
    $groups = @{}
    foreach ($segment in $Segments) {
        $key = & $KeySelector $segment
        if ($null -eq $key -or [string]$key -eq "") {
            $key = "(missing)"
        }

        $key = [string]$key
        if (-not $groups.ContainsKey($key)) {
            $groups[$key] = [pscustomobject]@{
                Key = $key
                Count = 0
                GlyphCount = 0
                GlyphGapCount = 0
                NaturalPdfWidth = 0d
                UnkernedPdfWidth = 0d
                RoundedPdfWidth = 0d
                KerningAdjustmentTotal = 0d
                PositioningCharacterSpacingGapTotal = 0d
                TextStateCharacterSpacingGapTotal = 0d
                PlannedEmittedAdvance = 0d
                LayoutWidth = 0d
                LayoutToNaturalResidual = 0d
                LayoutToRoundedResidual = 0d
                PlannedEmittedToLayoutResidual = 0d
            }
        }

        $group = $groups[$key]
        $group.Count++
        if ($null -ne $segment.AdvanceProfile) {
            $group.GlyphCount += [int]$segment.AdvanceProfile.GlyphCount
            $group.GlyphGapCount += [int]$segment.AdvanceProfile.GlyphGapCount
            $group.NaturalPdfWidth += [double]$segment.AdvanceProfile.NaturalPdfWidth
            $group.UnkernedPdfWidth += [double]$segment.AdvanceProfile.UnkernedPdfWidth
            $group.RoundedPdfWidth += [double]$segment.AdvanceProfile.RoundedPdfWidth
            $group.KerningAdjustmentTotal += [double]$segment.AdvanceProfile.KerningAdjustmentTotal
            $group.PositioningCharacterSpacingGapTotal += [double]$segment.AdvanceProfile.PositioningCharacterSpacingGapTotal
            $group.TextStateCharacterSpacingGapTotal += [double]$segment.AdvanceProfile.TextStateCharacterSpacingGapTotal
            $group.PlannedEmittedAdvance += [double]$segment.AdvanceProfile.PlannedEmittedAdvance
            $group.LayoutWidth += [double]$segment.AdvanceProfile.LayoutWidth
            $group.LayoutToNaturalResidual += [double]$segment.AdvanceProfile.LayoutToNaturalResidual
            $group.LayoutToRoundedResidual += [double]$segment.AdvanceProfile.LayoutToRoundedResidual
            $group.PlannedEmittedToLayoutResidual += [double]$segment.AdvanceProfile.PlannedEmittedToLayoutResidual
        }
    }

    return @(
        foreach ($key in ($groups.Keys | Sort-Object)) {
            $group = $groups[$key]
            [pscustomobject]@{
                Key = $group.Key
                Count = $group.Count
                GlyphCount = $group.GlyphCount
                GlyphGapCount = $group.GlyphGapCount
                NaturalPdfWidth = [Math]::Round($group.NaturalPdfWidth, 6)
                UnkernedPdfWidth = [Math]::Round($group.UnkernedPdfWidth, 6)
                RoundedPdfWidth = [Math]::Round($group.RoundedPdfWidth, 6)
                KerningAdjustmentTotal = [Math]::Round($group.KerningAdjustmentTotal, 6)
                PositioningCharacterSpacingGapTotal = [Math]::Round($group.PositioningCharacterSpacingGapTotal, 6)
                TextStateCharacterSpacingGapTotal = [Math]::Round($group.TextStateCharacterSpacingGapTotal, 6)
                PlannedEmittedAdvance = [Math]::Round($group.PlannedEmittedAdvance, 6)
                LayoutWidth = [Math]::Round($group.LayoutWidth, 6)
                LayoutToNaturalResidual = [Math]::Round($group.LayoutToNaturalResidual, 6)
                LayoutToRoundedResidual = [Math]::Round($group.LayoutToRoundedResidual, 6)
                PlannedEmittedToLayoutResidual = [Math]::Round($group.PlannedEmittedToLayoutResidual, 6)
                UniformResidualPerGap = if ($group.GlyphGapCount -eq 0) { $null } else { [Math]::Round($group.LayoutToNaturalResidual / $group.GlyphGapCount, 6) }
                RoundedResidualPerGap = if ($group.GlyphGapCount -eq 0) { $null } else { [Math]::Round($group.LayoutToRoundedResidual / $group.GlyphGapCount, 6) }
                PlannedEmittedResidualPerGap = if ($group.GlyphGapCount -eq 0) { $null } else { [Math]::Round($group.PlannedEmittedToLayoutResidual / $group.GlyphGapCount, 6) }
            }
        }
    )
}

function Summarize-PlannerSnapshot($Snapshot) {
    if ($null -eq $Snapshot) {
        return $null
    }

    $segments = Flatten-PlannerSegments $Snapshot
    $nonTerminalSegments = @($segments | Where-Object { -not [bool]$_.IsTerminalLineSpace })
    return [pscustomobject]@{
        SegmentCount = $segments.Count
        NonTerminalSegmentCount = $nonTerminalSegments.Count
        TerminalSpaceSegmentCount = @($segments | Where-Object { [bool]$_.IsTerminalLineSpace }).Count
        TextClassBuckets = @(Group-Count $segments { param($segment) PlannerTextClass $segment })
        TextClassByFontSize = @(Group-Count $segments {
            param($segment)
            (PlannerTextClass $segment) + "|tf=" + (RoundedKey $segment.PdfFontSize 3)
        })
        TextClassByGlyphGap = @(Group-Count $segments {
            param($segment)
            (PlannerTextClass $segment) + "|gaps=" + (RoundedKey $segment.AdvanceProfile.GlyphGapCount 0)
        })
        TextClassBySourceParagraphIndex = @(Group-Count $segments {
            param($segment)
            (PlannerTextClass $segment) + "|srcPara=" + (RoundedKey $segment.SourceParagraphIndex 0)
        })
        TextClassByRole = @(Group-Count $segments {
            param($segment)
            (PlannerTextClass $segment) + "|role=" + $segment.Role
        })
        TextClassByPdfCharacterSpacingSource = @(Group-Count $segments {
            param($segment)
            (PlannerTextClass $segment) + "|tcSource=" + $segment.PdfCharacterSpacingSource
        })
        TextClassByRunProvenance = @(Group-Count $segments {
            param($segment)
            (PlannerTextClass $segment) + "|" + (PlannerRunProvenanceKey $segment)
        })
        PdfCharacterSpacingSourceBuckets = @(Group-Count $segments {
            param($segment)
            "tcSource=" + $segment.PdfCharacterSpacingSource + "|pdfTc=" + (RoundedKey $segment.PdfCharacterSpacing 6)
        })
        PdfCharacterSpacingSourceByRunProvenance = @(Group-Count $segments {
            param($segment)
            "tcSource=" + $segment.PdfCharacterSpacingSource + "|" + (PlannerRunProvenanceKey $segment)
        })
        TextClassByGlyphSignature = @(Group-Count $segments {
            param($segment)
            (PlannerTextClass $segment) + "|glyphSig=" + $segment.GlyphAdvanceSignature.Hash
        })
        TextClassByGlyphPairSignature = @(Group-Count $segments {
            param($segment)
            (PlannerTextClass $segment) + "|pairSig=" + $segment.GlyphAdvanceSignature.PairHash
        })
        TextClassByGlyphPairAdvanceRange = @(Group-Count $segments {
            param($segment)
            (PlannerTextClass $segment) + "|pairMin=" + (RoundedKey $segment.GlyphAdvanceSignature.PairAdvanceMinUnits 0) + "|pairMax=" + (RoundedKey $segment.GlyphAdvanceSignature.PairAdvanceMaxUnits 0)
        })
        TextClassByGlyphPairAdvanceEmRange = @(Group-Count $segments {
            param($segment)
            (PlannerTextClass $segment) + "|pairMinEm=" + (RoundedKey $segment.GlyphAdvanceSignature.PairAdvanceMinEm 6) + "|pairMaxEm=" + (RoundedKey $segment.GlyphAdvanceSignature.PairAdvanceMaxEm 6)
        })
        TextClassByResidualPerGap = @(Group-Count $segments {
            param($segment)
            (PlannerTextClass $segment) + "|resGap=" + (RoundedKey $segment.AdvanceProfile.UniformResidualPerGap 6)
        })
        TextClassByRoundedResidualPerGap = @(Group-Count $segments {
            param($segment)
            (PlannerTextClass $segment) + "|roundResGap=" + (RoundedKey $segment.AdvanceProfile.RoundedResidualPerGap 6)
        })
        TextLengthByResidualPerGap = @(Group-Count $segments {
            param($segment)
            "len=" + (RoundedKey $segment.TextLength 0) + "|resGap=" + (RoundedKey $segment.AdvanceProfile.UniformResidualPerGap 6)
        })
        FontSizeByResidualPerGap = @(Group-Count $segments {
            param($segment)
            "tf=" + (RoundedKey $segment.PdfFontSize 3) + "|resGap=" + (RoundedKey $segment.AdvanceProfile.UniformResidualPerGap 6)
        })
        FontSizeByRoundedResidualPerGap = @(Group-Count $segments {
            param($segment)
            "tf=" + (RoundedKey $segment.PdfFontSize 3) + "|roundResGap=" + (RoundedKey $segment.AdvanceProfile.RoundedResidualPerGap 6)
        })
        TerminalSpaceByResidualPerGap = @(Group-Count $segments {
            param($segment)
            "terminal=" + ([bool]$segment.IsTerminalLineSpace).ToString().ToLowerInvariant() + "|resGap=" + (RoundedKey $segment.AdvanceProfile.UniformResidualPerGap 6)
        })
        AdvanceByTextClass = @(Group-PlannerAdvance $segments { param($segment) PlannerTextClass $segment })
        AdvanceByFontSize = @(Group-PlannerAdvance $segments { param($segment) "tf=" + (RoundedKey $segment.PdfFontSize 3) })
        AdvanceByTextClassAndGlyphGap = @(Group-PlannerAdvance $segments {
            param($segment)
            (PlannerTextClass $segment) + "|gaps=" + (RoundedKey $segment.AdvanceProfile.GlyphGapCount 0)
        })
        AdvanceByRunProvenance = @(Group-PlannerAdvance $segments {
            param($segment)
            PlannerRunProvenanceKey $segment
        })
        AdvanceByGlyphSignature = @(Group-PlannerAdvance $segments {
            param($segment)
            "glyphSig=" + $segment.GlyphAdvanceSignature.Hash
        })
        AdvanceByGlyphPairSignature = @(Group-PlannerAdvance $segments {
            param($segment)
            "pairSig=" + $segment.GlyphAdvanceSignature.PairHash
        })
    }
}

function Summarize-PlannerReferencePairs($ReferenceOperations, $Snapshot) {
    if ($null -eq $Snapshot) {
        return $null
    }

    $segments = Flatten-PlannerSegments $Snapshot
    if ($ReferenceOperations.Count -ne $segments.Count) {
        return [pscustomobject]@{
            PairCount = 0
            ReferenceOperationCount = $ReferenceOperations.Count
            PlannerSegmentCount = $segments.Count
            CountsMatched = $false
        }
    }

    $pairs = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $ReferenceOperations.Count; $i++) {
        $reference = $ReferenceOperations[$i]
        $segment = $segments[$i]
        $pairs.Add([pscustomobject]@{
            ReferenceTc = $reference.CharacterSpacing
            ReferenceNetAverageCharacterSpacing = $reference.NetAverageCharacterSpacing
            ReferenceCharacterSpacingGapTotal = $reference.CharacterSpacingGapTotalPoints
            ReferenceAdjustmentTotal = $reference.AdjustmentTotalPoints
            ReferenceNetSpacingGapTotal = $reference.NetSpacingGapTotalPoints
            ReferenceNaturalWidth = $reference.NaturalWidthPoints
            ReferenceEmittedAdvance = $reference.EmittedAdvancePoints
            ReferenceFontSize = $reference.FontSize
            ReferenceWidthCodeCount = $reference.WidthSignature.CodeCount
            ReferenceWidthPairCount = $reference.WidthSignature.PairCount
            ReferenceWidthUnits = $reference.WidthSignature.WidthUnits
            ReferencePairWidthUnits = $reference.WidthSignature.PairWidthUnits
            ReferencePairLeftWidthUnits = $reference.WidthSignature.PairLeftWidthUnits
            ReferencePairRightWidthUnits = $reference.WidthSignature.PairRightWidthUnits
            ReferencePairWidthMinUnits = $reference.WidthSignature.PairWidthMinUnits
            ReferencePairWidthMaxUnits = $reference.WidthSignature.PairWidthMaxUnits
            ReferencePairLeftWidthMinUnits = $reference.WidthSignature.PairLeftWidthMinUnits
            ReferencePairLeftWidthMaxUnits = $reference.WidthSignature.PairLeftWidthMaxUnits
            ReferencePairRightWidthMinUnits = $reference.WidthSignature.PairRightWidthMinUnits
            ReferencePairRightWidthMaxUnits = $reference.WidthSignature.PairRightWidthMaxUnits
            ReferenceWidthEm = (DivideOrNull $reference.WidthSignature.WidthUnits 1000d)
            ReferencePairWidthEm = (DivideOrNull $reference.WidthSignature.PairWidthUnits 1000d)
            ReferencePairLeftWidthEm = (DivideOrNull $reference.WidthSignature.PairLeftWidthUnits 1000d)
            ReferencePairRightWidthEm = (DivideOrNull $reference.WidthSignature.PairRightWidthUnits 1000d)
            ReferencePairWidthMinEm = (DivideOrNull $reference.WidthSignature.PairWidthMinUnits 1000d)
            ReferencePairWidthMaxEm = (DivideOrNull $reference.WidthSignature.PairWidthMaxUnits 1000d)
            ReferencePairLeftWidthMinEm = (DivideOrNull $reference.WidthSignature.PairLeftWidthMinUnits 1000d)
            ReferencePairLeftWidthMaxEm = (DivideOrNull $reference.WidthSignature.PairLeftWidthMaxUnits 1000d)
            ReferencePairRightWidthMinEm = (DivideOrNull $reference.WidthSignature.PairRightWidthMinUnits 1000d)
            ReferencePairRightWidthMaxEm = (DivideOrNull $reference.WidthSignature.PairRightWidthMaxUnits 1000d)
            PlannerTextClass = PlannerTextClass $segment
            PlannerTextLength = $segment.TextLength
            PlannerPdfFontSize = $segment.PdfFontSize
            PlannerSourceParagraphIndex = $segment.SourceParagraphIndex
            PlannerRole = $segment.Role
            PlannerPdfCharacterSpacingSource = $segment.PdfCharacterSpacingSource
            PlannerRunProvenance = PlannerRunProvenanceKey $segment
            PlannerGlyphGapCount = $segment.AdvanceProfile.GlyphGapCount
            PlannerNaturalWidth = $segment.AdvanceProfile.NaturalPdfWidth
            PlannerUnkernedWidth = $segment.AdvanceProfile.UnkernedPdfWidth
            PlannerRoundedWidth = $segment.AdvanceProfile.RoundedPdfWidth
            PlannerKerningAdjustmentTotal = $segment.AdvanceProfile.KerningAdjustmentTotal
            PlannerPositioningCharacterSpacingGapTotal = $segment.AdvanceProfile.PositioningCharacterSpacingGapTotal
            PlannerTextStateCharacterSpacingGapTotal = $segment.AdvanceProfile.TextStateCharacterSpacingGapTotal
            PlannerEmittedAdvance = $segment.AdvanceProfile.PlannedEmittedAdvance
            PlannerLayoutWidth = $segment.AdvanceProfile.LayoutWidth
            ReferenceNaturalMinusPlannerNatural = (SubtractOrNull $reference.NaturalWidthPoints $segment.AdvanceProfile.NaturalPdfWidth)
            ReferenceEmittedMinusPlannerLayout = (SubtractOrNull $reference.EmittedAdvancePoints $segment.AdvanceProfile.LayoutWidth)
            ReferenceEmittedMinusPlannerRounded = (SubtractOrNull $reference.EmittedAdvancePoints $segment.AdvanceProfile.RoundedPdfWidth)
            ReferenceEmittedMinusPlannerEmitted = (SubtractOrNull $reference.EmittedAdvancePoints $segment.AdvanceProfile.PlannedEmittedAdvance)
            ReferenceEmittedMinusPlannerRoundedPerGap = (DivideOrNull (SubtractOrNull $reference.EmittedAdvancePoints $segment.AdvanceProfile.RoundedPdfWidth) $segment.AdvanceProfile.GlyphGapCount)
            ReferenceEmittedMinusPlannerEmittedPerGap = (DivideOrNull (SubtractOrNull $reference.EmittedAdvancePoints $segment.AdvanceProfile.PlannedEmittedAdvance) $segment.AdvanceProfile.GlyphGapCount)
            PlannerResidualPerGap = $segment.AdvanceProfile.UniformResidualPerGap
            PlannerRoundedResidualPerGap = $segment.AdvanceProfile.RoundedResidualPerGap
            PlannerEmittedResidualPerGap = $segment.AdvanceProfile.PlannedEmittedResidualPerGap
            PlannerGlyphAdvanceSignature = $segment.GlyphAdvanceSignature.Hash
            PlannerGlyphPairAdvanceSignature = $segment.GlyphAdvanceSignature.PairHash
            PlannerGlyphPairAdvanceUnits = $segment.GlyphAdvanceSignature.PairAdvanceUnits
            PlannerGlyphPairLeftAdvanceUnits = $segment.GlyphAdvanceSignature.PairLeftAdvanceUnits
            PlannerGlyphPairRightAdvanceUnits = $segment.GlyphAdvanceSignature.PairRightAdvanceUnits
            PlannerGlyphPairAdvanceMinUnits = $segment.GlyphAdvanceSignature.PairAdvanceMinUnits
            PlannerGlyphPairAdvanceMaxUnits = $segment.GlyphAdvanceSignature.PairAdvanceMaxUnits
            PlannerGlyphPairLeftAdvanceMinUnits = $segment.GlyphAdvanceSignature.PairLeftAdvanceMinUnits
            PlannerGlyphPairLeftAdvanceMaxUnits = $segment.GlyphAdvanceSignature.PairLeftAdvanceMaxUnits
            PlannerGlyphPairRightAdvanceMinUnits = $segment.GlyphAdvanceSignature.PairRightAdvanceMinUnits
            PlannerGlyphPairRightAdvanceMaxUnits = $segment.GlyphAdvanceSignature.PairRightAdvanceMaxUnits
            PlannerGlyphPairAdvanceEm = $segment.GlyphAdvanceSignature.PairAdvanceEm
            PlannerGlyphPairLeftAdvanceEm = $segment.GlyphAdvanceSignature.PairLeftAdvanceEm
            PlannerGlyphPairRightAdvanceEm = $segment.GlyphAdvanceSignature.PairRightAdvanceEm
            PlannerGlyphPairAdvanceMinEm = $segment.GlyphAdvanceSignature.PairAdvanceMinEm
            PlannerGlyphPairAdvanceMaxEm = $segment.GlyphAdvanceSignature.PairAdvanceMaxEm
            PlannerGlyphPairLeftAdvanceMinEm = $segment.GlyphAdvanceSignature.PairLeftAdvanceMinEm
            PlannerGlyphPairLeftAdvanceMaxEm = $segment.GlyphAdvanceSignature.PairLeftAdvanceMaxEm
            PlannerGlyphPairRightAdvanceMinEm = $segment.GlyphAdvanceSignature.PairRightAdvanceMinEm
            PlannerGlyphPairRightAdvanceMaxEm = $segment.GlyphAdvanceSignature.PairRightAdvanceMaxEm
            PlannerTerminalSpace = [bool]$segment.IsTerminalLineSpace
        })
    }

    $nonzeroReferencePairs = @($pairs | Where-Object { [Math]::Abs([double]$_.ReferenceTc) -gt 0.001d })
    $referenceTcAmbiguityChecks = @(
        New-TcAmbiguityReport "tf+role+class+gaps" $pairs {
            param($pair)
            "tf=" + (RoundedKey $pair.PlannerPdfFontSize 3) + "|role=" + $pair.PlannerRole + "|class=" + $pair.PlannerTextClass + "|gaps=" + (RoundedKey $pair.PlannerGlyphGapCount 0)
        } $false
        New-TcAmbiguityReport "tf+role+class+gaps" $pairs {
            param($pair)
            "tf=" + (RoundedKey $pair.PlannerPdfFontSize 3) + "|role=" + $pair.PlannerRole + "|class=" + $pair.PlannerTextClass + "|gaps=" + (RoundedKey $pair.PlannerGlyphGapCount 0)
        } $true
        New-TcAmbiguityReport "tf+tc-source+gaps" $pairs {
            param($pair)
            "tf=" + (RoundedKey $pair.PlannerPdfFontSize 3) + "|tcSource=" + $pair.PlannerPdfCharacterSpacingSource + "|gaps=" + (RoundedKey $pair.PlannerGlyphGapCount 0)
        } $true
        New-TcAmbiguityReport "tf+run-provenance+gaps" $pairs {
            param($pair)
            "tf=" + (RoundedKey $pair.PlannerPdfFontSize 3) + "|" + $pair.PlannerRunProvenance + "|gaps=" + (RoundedKey $pair.PlannerGlyphGapCount 0)
        } $true
        New-TcAmbiguityReport "tf+gaps+pair-range" $pairs {
            param($pair)
            "tf=" + (RoundedKey $pair.PlannerPdfFontSize 3) + "|gaps=" + (RoundedKey $pair.PlannerGlyphGapCount 0) + "|pairMin=" + (RoundedKey $pair.PlannerGlyphPairAdvanceMinUnits 0) + "|pairMax=" + (RoundedKey $pair.PlannerGlyphPairAdvanceMaxUnits 0)
        } $true
        New-TcAmbiguityReport "tf+gaps+side-range" $pairs {
            param($pair)
            "tf=" + (RoundedKey $pair.PlannerPdfFontSize 3) + "|gaps=" + (RoundedKey $pair.PlannerGlyphGapCount 0) + "|leftMin=" + (RoundedKey $pair.PlannerGlyphPairLeftAdvanceMinUnits 0) + "|leftMax=" + (RoundedKey $pair.PlannerGlyphPairLeftAdvanceMaxUnits 0) + "|rightMin=" + (RoundedKey $pair.PlannerGlyphPairRightAdvanceMinUnits 0) + "|rightMax=" + (RoundedKey $pair.PlannerGlyphPairRightAdvanceMaxUnits 0)
        } $true
        New-TcAmbiguityReport "tf+side-range" $pairs {
            param($pair)
            "tf=" + (RoundedKey $pair.PlannerPdfFontSize 3) + "|leftMin=" + (RoundedKey $pair.PlannerGlyphPairLeftAdvanceMinUnits 0) + "|leftMax=" + (RoundedKey $pair.PlannerGlyphPairLeftAdvanceMaxUnits 0) + "|rightMin=" + (RoundedKey $pair.PlannerGlyphPairRightAdvanceMinUnits 0) + "|rightMax=" + (RoundedKey $pair.PlannerGlyphPairRightAdvanceMaxUnits 0)
        } $true
        New-TcAmbiguityReport "ref-tf+pdf-side-range-em" $pairs {
            param($pair)
            "tf=" + (RoundedKey $pair.ReferenceFontSize 3) + "|leftMinEm=" + (RoundedKey $pair.ReferencePairLeftWidthMinEm 6) + "|leftMaxEm=" + (RoundedKey $pair.ReferencePairLeftWidthMaxEm 6) + "|rightMinEm=" + (RoundedKey $pair.ReferencePairRightWidthMinEm 6) + "|rightMaxEm=" + (RoundedKey $pair.ReferencePairRightWidthMaxEm 6)
        } $true
        New-TcAmbiguityReport "ref-tf+pdf-side-range-em+planner-side-range-em" $pairs {
            param($pair)
            "tfRef=" + (RoundedKey $pair.ReferenceFontSize 3) + "|tfPlan=" + (RoundedKey $pair.PlannerPdfFontSize 3) +
                "|refLeftMinEm=" + (RoundedKey $pair.ReferencePairLeftWidthMinEm 6) + "|refLeftMaxEm=" + (RoundedKey $pair.ReferencePairLeftWidthMaxEm 6) +
                "|refRightMinEm=" + (RoundedKey $pair.ReferencePairRightWidthMinEm 6) + "|refRightMaxEm=" + (RoundedKey $pair.ReferencePairRightWidthMaxEm 6) +
                "|planLeftMinEm=" + (RoundedKey $pair.PlannerGlyphPairLeftAdvanceMinEm 6) + "|planLeftMaxEm=" + (RoundedKey $pair.PlannerGlyphPairLeftAdvanceMaxEm 6) +
                "|planRightMinEm=" + (RoundedKey $pair.PlannerGlyphPairRightAdvanceMinEm 6) + "|planRightMaxEm=" + (RoundedKey $pair.PlannerGlyphPairRightAdvanceMaxEm 6)
        } $true
        New-TcAmbiguityReport "tf+rounded-residual-per-gap" $pairs {
            param($pair)
            "tf=" + (RoundedKey $pair.PlannerPdfFontSize 3) + "|roundedResidualPerGap=" + (RoundedKey $pair.PlannerRoundedResidualPerGap 6)
        } $true
    )
    $referenceTcTargetHypotheses = @(
        New-TcTargetHypothesisReport "reference-emitted-minus-planner-emitted-per-gap" $pairs {
            param($pair)
            $pair.ReferenceEmittedMinusPlannerEmittedPerGap
        }
        New-TcTargetHypothesisReport "reference-emitted-minus-planner-rounded-per-gap" $pairs {
            param($pair)
            $pair.ReferenceEmittedMinusPlannerRoundedPerGap
        }
        New-TcTargetHypothesisReport "planner-layout-minus-rounded-per-gap" $pairs {
            param($pair)
            $pair.PlannerRoundedResidualPerGap
        }
        New-TcTargetHypothesisReport "planner-layout-minus-emitted-per-gap" $pairs {
            param($pair)
            $pair.PlannerEmittedResidualPerGap
        }
    )
    return [pscustomobject]@{
        PairCount = $pairs.Count
        ReferenceOperationCount = $ReferenceOperations.Count
        PlannerSegmentCount = $segments.Count
        CountsMatched = $true
        Pairs = @($pairs.ToArray())
        ReferenceTcAmbiguityChecks = $referenceTcAmbiguityChecks
        ReferenceTcStructuralDiscriminators = Select-TcStructuralDiscriminators $referenceTcAmbiguityChecks
        ReferenceTcTargetHypotheses = $referenceTcTargetHypotheses
        ReferenceTcTargetAdvanceModel = New-TcTargetAdvanceModelReport $pairs $referenceTcTargetHypotheses
        ReferenceTcByPlannerTextClass = @(Group-Count $pairs {
            param($pair)
            $pair.PlannerTextClass + "|refTc=" + (RoundedKey $pair.ReferenceTc 6)
        })
        ReferenceTcByPlannerGlyphGap = @(Group-Count $pairs {
            param($pair)
            "gaps=" + (RoundedKey $pair.PlannerGlyphGapCount 0) + "|refTc=" + (RoundedKey $pair.ReferenceTc 6)
        })
        ReferenceTcByPlannerRunProvenance = @(Group-Count $pairs {
            param($pair)
            $pair.PlannerRunProvenance + "|refTc=" + (RoundedKey $pair.ReferenceTc 6)
        })
        ReferenceTcByPlannerResidualPerGap = @(Group-Count $pairs {
            param($pair)
            "resGap=" + (RoundedKey $pair.PlannerResidualPerGap 6) + "|refTc=" + (RoundedKey $pair.ReferenceTc 6)
        })
        ReferenceTcByPlannerRoundedResidualPerGap = @(Group-Count $pairs {
            param($pair)
            "roundResGap=" + (RoundedKey $pair.PlannerRoundedResidualPerGap 6) + "|refTc=" + (RoundedKey $pair.ReferenceTc 6)
        })
        ReferenceTcByPlannerGlyphSignature = @(Group-Count $pairs {
            param($pair)
            "glyphSig=" + $pair.PlannerGlyphAdvanceSignature + "|refTc=" + (RoundedKey $pair.ReferenceTc 6)
        })
        ReferenceTcByPlannerGlyphPairSignature = @(Group-Count $pairs {
            param($pair)
            "pairSig=" + $pair.PlannerGlyphPairAdvanceSignature + "|refTc=" + (RoundedKey $pair.ReferenceTc 6)
        })
        ReferenceTcByPlannerGlyphPairAdvanceRange = @(Group-Count $pairs {
            param($pair)
            "pairMin=" + (RoundedKey $pair.PlannerGlyphPairAdvanceMinUnits 0) + "|pairMax=" + (RoundedKey $pair.PlannerGlyphPairAdvanceMaxUnits 0) + "|refTc=" + (RoundedKey $pair.ReferenceTc 6)
        })
        ReferenceTcByPlannerGlyphPairAdvanceEmRange = @(Group-Count $pairs {
            param($pair)
            "pairMinEm=" + (RoundedKey $pair.PlannerGlyphPairAdvanceMinEm 6) + "|pairMaxEm=" + (RoundedKey $pair.PlannerGlyphPairAdvanceMaxEm 6) + "|refTc=" + (RoundedKey $pair.ReferenceTc 6)
        })
        ReferenceTcByPlannerGlyphPairSideAdvanceRange = @(Group-Count $pairs {
            param($pair)
            "leftMin=" + (RoundedKey $pair.PlannerGlyphPairLeftAdvanceMinUnits 0) + "|leftMax=" + (RoundedKey $pair.PlannerGlyphPairLeftAdvanceMaxUnits 0) + "|rightMin=" + (RoundedKey $pair.PlannerGlyphPairRightAdvanceMinUnits 0) + "|rightMax=" + (RoundedKey $pair.PlannerGlyphPairRightAdvanceMaxUnits 0) + "|refTc=" + (RoundedKey $pair.ReferenceTc 6)
        })
        ReferenceTcByPlannerGlyphPairSideAdvanceEmRange = @(Group-Count $pairs {
            param($pair)
            "leftMinEm=" + (RoundedKey $pair.PlannerGlyphPairLeftAdvanceMinEm 6) + "|leftMaxEm=" + (RoundedKey $pair.PlannerGlyphPairLeftAdvanceMaxEm 6) + "|rightMinEm=" + (RoundedKey $pair.PlannerGlyphPairRightAdvanceMinEm 6) + "|rightMaxEm=" + (RoundedKey $pair.PlannerGlyphPairRightAdvanceMaxEm 6) + "|refTc=" + (RoundedKey $pair.ReferenceTc 6)
        })
        ReferenceTcByPlannerFontSizeAndGlyphPairSideAdvanceRange = @(Group-Count $pairs {
            param($pair)
            "tf=" + (RoundedKey $pair.PlannerPdfFontSize 3) + "|leftMin=" + (RoundedKey $pair.PlannerGlyphPairLeftAdvanceMinUnits 0) + "|leftMax=" + (RoundedKey $pair.PlannerGlyphPairLeftAdvanceMaxUnits 0) + "|rightMin=" + (RoundedKey $pair.PlannerGlyphPairRightAdvanceMinUnits 0) + "|rightMax=" + (RoundedKey $pair.PlannerGlyphPairRightAdvanceMaxUnits 0) + "|refTc=" + (RoundedKey $pair.ReferenceTc 6)
        })
        ReferenceTcByReferenceFontSizeAndPdfWidthSideRange = @(Group-Count $pairs {
            param($pair)
            "tf=" + (RoundedKey $pair.ReferenceFontSize 3) + "|leftMin=" + (RoundedKey $pair.ReferencePairLeftWidthMinUnits 0) + "|leftMax=" + (RoundedKey $pair.ReferencePairLeftWidthMaxUnits 0) + "|rightMin=" + (RoundedKey $pair.ReferencePairRightWidthMinUnits 0) + "|rightMax=" + (RoundedKey $pair.ReferencePairRightWidthMaxUnits 0) + "|refTc=" + (RoundedKey $pair.ReferenceTc 6)
        })
        ReferenceTcByReferenceFontSizeAndPdfWidthSideEmRange = @(Group-Count $pairs {
            param($pair)
            "tf=" + (RoundedKey $pair.ReferenceFontSize 3) + "|leftMinEm=" + (RoundedKey $pair.ReferencePairLeftWidthMinEm 6) + "|leftMaxEm=" + (RoundedKey $pair.ReferencePairLeftWidthMaxEm 6) + "|rightMinEm=" + (RoundedKey $pair.ReferencePairRightWidthMinEm 6) + "|rightMaxEm=" + (RoundedKey $pair.ReferencePairRightWidthMaxEm 6) + "|refTc=" + (RoundedKey $pair.ReferenceTc 6)
        })
        ReferenceVsPlannerPdfWidthSideRange = @(Group-Count $pairs {
            param($pair)
            "tfRef=" + (RoundedKey $pair.ReferenceFontSize 3) + "|tfPlan=" + (RoundedKey $pair.PlannerPdfFontSize 3) +
                "|refLeftMin=" + (RoundedKey $pair.ReferencePairLeftWidthMinUnits 0) + "|refLeftMax=" + (RoundedKey $pair.ReferencePairLeftWidthMaxUnits 0) +
                "|refRightMin=" + (RoundedKey $pair.ReferencePairRightWidthMinUnits 0) + "|refRightMax=" + (RoundedKey $pair.ReferencePairRightWidthMaxUnits 0) +
                "|planLeftMin=" + (RoundedKey $pair.PlannerGlyphPairLeftAdvanceMinUnits 0) + "|planLeftMax=" + (RoundedKey $pair.PlannerGlyphPairLeftAdvanceMaxUnits 0) +
                "|planRightMin=" + (RoundedKey $pair.PlannerGlyphPairRightAdvanceMinUnits 0) + "|planRightMax=" + (RoundedKey $pair.PlannerGlyphPairRightAdvanceMaxUnits 0) +
                "|refTc=" + (RoundedKey $pair.ReferenceTc 6)
        })
        ReferenceVsPlannerPdfWidthSideEmRange = @(Group-Count $pairs {
            param($pair)
            "tfRef=" + (RoundedKey $pair.ReferenceFontSize 3) + "|tfPlan=" + (RoundedKey $pair.PlannerPdfFontSize 3) +
                "|refLeftMinEm=" + (RoundedKey $pair.ReferencePairLeftWidthMinEm 6) + "|refLeftMaxEm=" + (RoundedKey $pair.ReferencePairLeftWidthMaxEm 6) +
                "|refRightMinEm=" + (RoundedKey $pair.ReferencePairRightWidthMinEm 6) + "|refRightMaxEm=" + (RoundedKey $pair.ReferencePairRightWidthMaxEm 6) +
                "|planLeftMinEm=" + (RoundedKey $pair.PlannerGlyphPairLeftAdvanceMinEm 6) + "|planLeftMaxEm=" + (RoundedKey $pair.PlannerGlyphPairLeftAdvanceMaxEm 6) +
                "|planRightMinEm=" + (RoundedKey $pair.PlannerGlyphPairRightAdvanceMinEm 6) + "|planRightMaxEm=" + (RoundedKey $pair.PlannerGlyphPairRightAdvanceMaxEm 6) +
                "|refTc=" + (RoundedKey $pair.ReferenceTc 6)
        })
        ReferenceTcByPlannerFontSizeAndGlyphPairSideAdvanceEmRange = @(Group-Count $pairs {
            param($pair)
            "tf=" + (RoundedKey $pair.PlannerPdfFontSize 3) + "|leftMinEm=" + (RoundedKey $pair.PlannerGlyphPairLeftAdvanceMinEm 6) + "|leftMaxEm=" + (RoundedKey $pair.PlannerGlyphPairLeftAdvanceMaxEm 6) + "|rightMinEm=" + (RoundedKey $pair.PlannerGlyphPairRightAdvanceMinEm 6) + "|rightMaxEm=" + (RoundedKey $pair.PlannerGlyphPairRightAdvanceMaxEm 6) + "|refTc=" + (RoundedKey $pair.ReferenceTc 6)
        })
        ReferenceTcByPlannerFontSize = @(Group-Count $pairs {
            param($pair)
            "tf=" + (RoundedKey $pair.PlannerPdfFontSize 3) + "|refTc=" + (RoundedKey $pair.ReferenceTc 6)
        })
        ReferenceTcByPlannerSourceParagraphIndex = @(Group-Count $pairs {
            param($pair)
            "srcPara=" + (RoundedKey $pair.PlannerSourceParagraphIndex 0) + "|refTc=" + (RoundedKey $pair.ReferenceTc 6)
        })
        ReferenceTcByPlannerRole = @(Group-Count $pairs {
            param($pair)
            "role=" + $pair.PlannerRole + "|refTc=" + (RoundedKey $pair.ReferenceTc 6)
        })
        ReferenceNetByPlannerResidualPerGap = @(Group-Count $pairs {
            param($pair)
            "resGap=" + (RoundedKey $pair.PlannerResidualPerGap 6) + "|refNet=" + (RoundedKey $pair.ReferenceNetAverageCharacterSpacing 6)
        })
        ReferenceSpacingTotalsByPlannerFontSizeAndGlyphPairSideAdvanceRange = @(Group-Count $pairs {
            param($pair)
            "tf=" + (RoundedKey $pair.PlannerPdfFontSize 3) + "|leftMin=" + (RoundedKey $pair.PlannerGlyphPairLeftAdvanceMinUnits 0) + "|leftMax=" + (RoundedKey $pair.PlannerGlyphPairLeftAdvanceMaxUnits 0) + "|rightMin=" + (RoundedKey $pair.PlannerGlyphPairRightAdvanceMinUnits 0) + "|rightMax=" + (RoundedKey $pair.PlannerGlyphPairRightAdvanceMaxUnits 0) + "|refTcGapTotal=" + (RoundedKey $pair.ReferenceCharacterSpacingGapTotal 6) + "|refAdjTotal=" + (RoundedKey $pair.ReferenceAdjustmentTotal 6) + "|refNetGapTotal=" + (RoundedKey $pair.ReferenceNetSpacingGapTotal 6)
        })
        ReferenceWidthsByPlannerFontSizeAndGlyphPairSideAdvanceRange = @(Group-Count $pairs {
            param($pair)
            "tf=" + (RoundedKey $pair.PlannerPdfFontSize 3) + "|leftMin=" + (RoundedKey $pair.PlannerGlyphPairLeftAdvanceMinUnits 0) + "|leftMax=" + (RoundedKey $pair.PlannerGlyphPairLeftAdvanceMaxUnits 0) + "|rightMin=" + (RoundedKey $pair.PlannerGlyphPairRightAdvanceMinUnits 0) + "|rightMax=" + (RoundedKey $pair.PlannerGlyphPairRightAdvanceMaxUnits 0) + "|refNatural=" + (RoundedKey $pair.ReferenceNaturalWidth 6) + "|refEmitted=" + (RoundedKey $pair.ReferenceEmittedAdvance 6)
        })
        ReferenceWidthDeltasByPlannerFontSizeAndGlyphPairSideAdvanceRange = @(Group-Count $pairs {
            param($pair)
            "tf=" + (RoundedKey $pair.PlannerPdfFontSize 3) + "|leftMin=" + (RoundedKey $pair.PlannerGlyphPairLeftAdvanceMinUnits 0) + "|leftMax=" + (RoundedKey $pair.PlannerGlyphPairLeftAdvanceMaxUnits 0) + "|rightMin=" + (RoundedKey $pair.PlannerGlyphPairRightAdvanceMinUnits 0) + "|rightMax=" + (RoundedKey $pair.PlannerGlyphPairRightAdvanceMaxUnits 0) + "|refNaturalMinusPlannerNatural=" + (RoundedKey $pair.ReferenceNaturalMinusPlannerNatural 6) + "|refEmittedMinusPlannerLayout=" + (RoundedKey $pair.ReferenceEmittedMinusPlannerLayout 6) + "|refEmittedMinusPlannerRounded=" + (RoundedKey $pair.ReferenceEmittedMinusPlannerRounded 6)
        })
        ReferenceTcByPlannerWidthDeltas = @(Group-Count $pairs {
            param($pair)
            "tf=" + (RoundedKey $pair.PlannerPdfFontSize 3) + "|refNaturalMinusPlannerNatural=" + (RoundedKey $pair.ReferenceNaturalMinusPlannerNatural 6) + "|refEmittedMinusPlannerLayout=" + (RoundedKey $pair.ReferenceEmittedMinusPlannerLayout 6) + "|refEmittedMinusPlannerRounded=" + (RoundedKey $pair.ReferenceEmittedMinusPlannerRounded 6) + "|refTc=" + (RoundedKey $pair.ReferenceTc 6)
        })
        ReferenceTcByPlannerPerGapWidthDeltas = @(Group-Count $pairs {
            param($pair)
            "tf=" + (RoundedKey $pair.PlannerPdfFontSize 3) + "|gaps=" + (RoundedKey $pair.PlannerGlyphGapCount 0) + "|refEmittedMinusPlannerRoundedPerGap=" + (RoundedKey $pair.ReferenceEmittedMinusPlannerRoundedPerGap 6) + "|plannerRoundedResidualPerGap=" + (RoundedKey $pair.PlannerRoundedResidualPerGap 6) + "|refTc=" + (RoundedKey $pair.ReferenceTc 6)
        })
        ReferenceVsPlannerEmittedAdvanceDeltas = @(Group-Count $pairs {
            param($pair)
            "tf=" + (RoundedKey $pair.PlannerPdfFontSize 3) + "|gaps=" + (RoundedKey $pair.PlannerGlyphGapCount 0) + "|refEmittedMinusPlannerEmitted=" + (RoundedKey $pair.ReferenceEmittedMinusPlannerEmitted 6) + "|refEmittedMinusPlannerEmittedPerGap=" + (RoundedKey $pair.ReferenceEmittedMinusPlannerEmittedPerGap 6) + "|plannerTextStateGapTotal=" + (RoundedKey $pair.PlannerTextStateCharacterSpacingGapTotal 6) + "|plannerPositioningGapTotal=" + (RoundedKey $pair.PlannerPositioningCharacterSpacingGapTotal 6) + "|refTc=" + (RoundedKey $pair.ReferenceTc 6)
        })
        ReferenceNonzeroTcByPlannerClassGlyphGap = @(Group-Count $nonzeroReferencePairs {
            param($pair)
            $pair.PlannerTextClass + "|gaps=" + (RoundedKey $pair.PlannerGlyphGapCount 0) + "|refTc=" + (RoundedKey $pair.ReferenceTc 6)
        })
        ReferenceNonzeroTcByPlannerResidualPerGap = @(Group-Count $nonzeroReferencePairs {
            param($pair)
            "resGap=" + (RoundedKey $pair.PlannerResidualPerGap 6) + "|refTc=" + (RoundedKey $pair.ReferenceTc 6)
        })
        ReferenceNonzeroTcByPlannerRoundedResidualPerGap = @(Group-Count $nonzeroReferencePairs {
            param($pair)
            "roundResGap=" + (RoundedKey $pair.PlannerRoundedResidualPerGap 6) + "|refTc=" + (RoundedKey $pair.ReferenceTc 6)
        })
        ReferenceNonzeroTcByPlannerGlyphSignature = @(Group-Count $nonzeroReferencePairs {
            param($pair)
            "glyphSig=" + $pair.PlannerGlyphAdvanceSignature + "|refTc=" + (RoundedKey $pair.ReferenceTc 6)
        })
        ReferenceNonzeroTcByPlannerGlyphPairSignature = @(Group-Count $nonzeroReferencePairs {
            param($pair)
            "pairSig=" + $pair.PlannerGlyphPairAdvanceSignature + "|refTc=" + (RoundedKey $pair.ReferenceTc 6)
        })
    }
}

$runs = Expand-PathList $RunDirectory
if ($runs.Count -eq 0) {
    throw "No run directories were provided."
}

$summaries = foreach ($run in $runs) {
    $resolvedRun = (Resolve-Path -LiteralPath $run).Path
    $referenceOps = Read-JsonArray (Ensure-TextOperations $resolvedRun "reference")
    $candidateOps = Read-JsonArray (Ensure-TextOperations $resolvedRun "candidate")
    Ensure-CandidateDocxInspect $resolvedRun
    $candidatePlannerSummary = Read-CandidatePlannerSummary $resolvedRun
    $candidatePlannerSnapshot = Read-CandidatePlannerSnapshot $resolvedRun
    $candidatePlannerReferencePairs = Summarize-PlannerReferencePairs $referenceOps $candidatePlannerSnapshot
    $metricsPath = Join-Path $resolvedRun "comparison\metrics.json"
    $metrics = if (Test-Path -LiteralPath $metricsPath) {
        $pages = Read-JsonArray $metricsPath
        [pscustomobject]@{
            PageCount = $pages.Count
            MeanAbsoluteError = if ($pages.Count -eq 0) { $null } else { [Math]::Round(($pages | Measure-Object -Property MeanAbsoluteError -Average).Average, 6) }
            Changed16 = if ($pages.Count -eq 0) { $null } else { [Math]::Round(($pages | Measure-Object -Property ChangedPixelRatioAtThreshold16 -Average).Average, 6) }
            StructuralSimilarity = if ($pages.Count -eq 0) { $null } else { [Math]::Round(($pages | Measure-Object -Property StructuralSimilarity -Average).Average, 6) }
        }
    }
    else {
        $null
    }

    [pscustomobject]@{
        RunDirectory = $resolvedRun
        CaseId = Split-Path -Leaf (Split-Path -Parent $resolvedRun)
        RunId = Split-Path -Leaf $resolvedRun
        Metrics = $metrics
        Reference = Summarize-Operations $referenceOps
        Candidate = Summarize-Operations $candidateOps
        CandidatePlanner = $candidatePlannerSummary
        CandidatePlannerSegments = Summarize-PlannerSnapshot $candidatePlannerSnapshot
        CandidatePlannerReferencePairs = $candidatePlannerReferencePairs
        ReferenceTcTargetAdvanceModel = $candidatePlannerReferencePairs.ReferenceTcTargetAdvanceModel
        TextClassCountDeltas = New-TextClassCountDeltas $referenceOps $candidateOps $candidatePlannerSnapshot
    }
}

if ($OutputJson) {
    $resolvedOutput = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputJson)
    $parent = Split-Path -Parent $resolvedOutput
    if ($parent -and -not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Path $parent | Out-Null
    }

    $summaries | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $resolvedOutput -Encoding utf8
}
else {
    $summaries | ConvertTo-Json -Depth 20
}
