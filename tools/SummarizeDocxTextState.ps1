param(
    [Parameter(Mandatory = $true)]
    [string[]] $RunDirectory,

    [string] $OutputJson
)

$ErrorActionPreference = "Stop"

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

function Ensure-TextOperations([string] $Run, [string] $Side) {
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
    if (-not (Test-Path -LiteralPath $json)) {
        pwsh tools\InspectPdf.ps1 -InputPdf $pdf -OutputDirectory $out -TextOnly | Out-Null
    }

    return $json
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
    }
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
                LayoutWidth = 0d
                LayoutToNaturalResidual = 0d
            }
        }

        $group = $groups[$key]
        $group.Count++
        if ($null -ne $segment.AdvanceProfile) {
            $group.GlyphCount += [int]$segment.AdvanceProfile.GlyphCount
            $group.GlyphGapCount += [int]$segment.AdvanceProfile.GlyphGapCount
            $group.NaturalPdfWidth += [double]$segment.AdvanceProfile.NaturalPdfWidth
            $group.LayoutWidth += [double]$segment.AdvanceProfile.LayoutWidth
            $group.LayoutToNaturalResidual += [double]$segment.AdvanceProfile.LayoutToNaturalResidual
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
                LayoutWidth = [Math]::Round($group.LayoutWidth, 6)
                LayoutToNaturalResidual = [Math]::Round($group.LayoutToNaturalResidual, 6)
                UniformResidualPerGap = if ($group.GlyphGapCount -eq 0) { $null } else { [Math]::Round($group.LayoutToNaturalResidual / $group.GlyphGapCount, 6) }
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
        TextClassByGlyphSignature = @(Group-Count $segments {
            param($segment)
            (PlannerTextClass $segment) + "|glyphSig=" + $segment.GlyphAdvanceSignature.Hash
        })
        TextClassByResidualPerGap = @(Group-Count $segments {
            param($segment)
            (PlannerTextClass $segment) + "|resGap=" + (RoundedKey $segment.AdvanceProfile.UniformResidualPerGap 6)
        })
        TextLengthByResidualPerGap = @(Group-Count $segments {
            param($segment)
            "len=" + (RoundedKey $segment.TextLength 0) + "|resGap=" + (RoundedKey $segment.AdvanceProfile.UniformResidualPerGap 6)
        })
        FontSizeByResidualPerGap = @(Group-Count $segments {
            param($segment)
            "tf=" + (RoundedKey $segment.PdfFontSize 3) + "|resGap=" + (RoundedKey $segment.AdvanceProfile.UniformResidualPerGap 6)
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
        AdvanceByGlyphSignature = @(Group-PlannerAdvance $segments {
            param($segment)
            "glyphSig=" + $segment.GlyphAdvanceSignature.Hash
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
            ReferenceFontSize = $reference.FontSize
            PlannerTextClass = PlannerTextClass $segment
            PlannerTextLength = $segment.TextLength
            PlannerPdfFontSize = $segment.PdfFontSize
            PlannerSourceParagraphIndex = $segment.SourceParagraphIndex
            PlannerRole = $segment.Role
            PlannerGlyphGapCount = $segment.AdvanceProfile.GlyphGapCount
            PlannerResidualPerGap = $segment.AdvanceProfile.UniformResidualPerGap
            PlannerGlyphAdvanceSignature = $segment.GlyphAdvanceSignature.Hash
            PlannerTerminalSpace = [bool]$segment.IsTerminalLineSpace
        })
    }

    $nonzeroReferencePairs = @($pairs | Where-Object { [Math]::Abs([double]$_.ReferenceTc) -gt 0.001d })
    return [pscustomobject]@{
        PairCount = $pairs.Count
        ReferenceOperationCount = $ReferenceOperations.Count
        PlannerSegmentCount = $segments.Count
        CountsMatched = $true
        ReferenceTcByPlannerTextClass = @(Group-Count $pairs {
            param($pair)
            $pair.PlannerTextClass + "|refTc=" + (RoundedKey $pair.ReferenceTc 6)
        })
        ReferenceTcByPlannerGlyphGap = @(Group-Count $pairs {
            param($pair)
            "gaps=" + (RoundedKey $pair.PlannerGlyphGapCount 0) + "|refTc=" + (RoundedKey $pair.ReferenceTc 6)
        })
        ReferenceTcByPlannerResidualPerGap = @(Group-Count $pairs {
            param($pair)
            "resGap=" + (RoundedKey $pair.PlannerResidualPerGap 6) + "|refTc=" + (RoundedKey $pair.ReferenceTc 6)
        })
        ReferenceTcByPlannerGlyphSignature = @(Group-Count $pairs {
            param($pair)
            "glyphSig=" + $pair.PlannerGlyphAdvanceSignature + "|refTc=" + (RoundedKey $pair.ReferenceTc 6)
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
        ReferenceNonzeroTcByPlannerClassGlyphGap = @(Group-Count $nonzeroReferencePairs {
            param($pair)
            $pair.PlannerTextClass + "|gaps=" + (RoundedKey $pair.PlannerGlyphGapCount 0) + "|refTc=" + (RoundedKey $pair.ReferenceTc 6)
        })
        ReferenceNonzeroTcByPlannerResidualPerGap = @(Group-Count $nonzeroReferencePairs {
            param($pair)
            "resGap=" + (RoundedKey $pair.PlannerResidualPerGap 6) + "|refTc=" + (RoundedKey $pair.ReferenceTc 6)
        })
        ReferenceNonzeroTcByPlannerGlyphSignature = @(Group-Count $nonzeroReferencePairs {
            param($pair)
            "glyphSig=" + $pair.PlannerGlyphAdvanceSignature + "|refTc=" + (RoundedKey $pair.ReferenceTc 6)
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
    $candidatePlannerSummary = Read-CandidatePlannerSummary $resolvedRun
    $candidatePlannerSnapshot = Read-CandidatePlannerSnapshot $resolvedRun
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
        CandidatePlannerReferencePairs = Summarize-PlannerReferencePairs $referenceOps $candidatePlannerSnapshot
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
