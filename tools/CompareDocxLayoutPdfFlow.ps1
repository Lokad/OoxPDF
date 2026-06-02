param(
    [Parameter(Mandatory = $true)]
    [string] $RunDirectory,

    [Parameter(Mandatory = $true)]
    [string] $LayoutSnapshot,

    [int] $PageStart = 1,

    [int] $PageEnd = [int]::MaxValue,

    [double] $LineTolerance = 0.75,

    [int] $Top = 80,

    [switch] $IncludeStaticStories,

    [string] $OutputJson
)

$ErrorActionPreference = "Stop"

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

function Read-JsonObject([string] $Path) {
    return Get-Content -Raw -LiteralPath (Resolve-Path -LiteralPath $Path).Path | ConvertFrom-Json
}

function Get-Coordinate($Operation, [string] $Name) {
    $effectiveName = "Effective$Name"
    if ($Operation.PSObject.Properties.Name -contains $effectiveName -and $null -ne $Operation.$effectiveName) {
        return [double]$Operation.$effectiveName
    }

    return [double]$Operation.$Name
}

function Get-Text($Operation) {
    if ($null -ne $Operation.DecodedText) {
        return [string]$Operation.DecodedText
    }

    return [string]$Operation.Payload
}

function Normalize-Text([string] $Text) {
    if ($null -eq $Text) {
        return ""
    }

    $normalized = $Text.
        Replace([char]0x00A0, ' ').
        Replace([char]0xF0B7, [char]0x2022).
        Replace("<>", " ").
        Replace("`r", " ").
        Replace("`n", " ").
        Replace("`t", " ")
    $normalized = [regex]::Replace($normalized, '^([•])\s*', '$1 ')
    return ([regex]::Replace($normalized, '\s+', ' ')).Trim()
}

function New-Hash([string] $Text) {
    $bytes = [Text.Encoding]::UTF8.GetBytes($Text)
    $hash = [Security.Cryptography.SHA256]::HashData($bytes)
    return [Convert]::ToHexString($hash).Substring(0, 16).ToLowerInvariant()
}

function New-LineRows($Operations) {
    $ordered = @($Operations | Sort-Object -Property `
        @{ Expression = { [int]$_.PageNumber }; Descending = $false },
        @{ Expression = { Get-Coordinate $_ "Y" }; Descending = $true },
        @{ Expression = { Get-Coordinate $_ "X" }; Descending = $false })

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($op in $ordered) {
        $page = [int]$op.PageNumber
        $y = Get-Coordinate $op "Y"
        $row = $null
        foreach ($candidate in $rows) {
            if ([int]$candidate.Page -eq $page -and [Math]::Abs([double]$candidate.Y - $y) -le $LineTolerance) {
                $row = $candidate
                break
            }
        }

        if ($null -eq $row) {
            $row = [pscustomobject]@{
                Page = $page
                Y = $y
                Operations = New-Object System.Collections.Generic.List[object]
            }
            $rows.Add($row)
        }

        $row.Operations.Add($op)
        $row.Y = ($row.Operations | ForEach-Object { Get-Coordinate $_ "Y" } | Measure-Object -Average).Average
    }

    $result = New-Object System.Collections.Generic.List[object]
    $globalIndex = 0
    foreach ($pageGroup in ($rows | Group-Object Page | Sort-Object { [int]$_.Name })) {
        $pageIndex = 0
        foreach ($row in ($pageGroup.Group | Sort-Object -Property @{ Expression = { [double]$_.Y }; Descending = $true })) {
            $ops = @($row.Operations | Sort-Object -Property @{ Expression = { Get-Coordinate $_ "X" }; Descending = $false })
            $text = [string]::Concat(@($ops | ForEach-Object { Get-Text $_ }))
            $normalized = Normalize-Text $text
            $result.Add([pscustomobject]@{
                GlobalIndex = $globalIndex++
                PageIndex = $pageIndex++
                Page = [int]$row.Page
                Y = [Math]::Round([double]$row.Y, 6)
                X = if ($ops.Count -eq 0) { $null } else { [Math]::Round((@($ops | ForEach-Object { Get-Coordinate $_ "X" }) | Measure-Object -Minimum).Minimum, 6) }
                OperationCount = $ops.Count
                TextLength = $normalized.Length
                TextHash = New-Hash $normalized
                NormalizedText = $normalized
            })
        }
    }

    return $result.ToArray()
}

function New-RowAdvances($Rows) {
    $advances = New-Object System.Collections.Generic.List[double]
    foreach ($pageGroup in ($Rows | Where-Object { [int]$_.TextLength -gt 0 } | Group-Object Page | Sort-Object { [int]$_.Name })) {
        $orderedRows = @($pageGroup.Group | Sort-Object -Property @{ Expression = { [int]$_.PageIndex }; Descending = $false })
        for ($index = 1; $index -lt $orderedRows.Count; $index++) {
            $advance = [double]$orderedRows[$index - 1].Y - [double]$orderedRows[$index].Y
            if ($advance -gt 0d) {
                $advances.Add($advance)
            }
        }
    }

    return $advances.ToArray()
}

function New-RowAdvanceBuckets($Rows) {
    $advances = @(New-RowAdvances $Rows)
    return @(
        $advances |
            Group-Object {
                [Math]::Round([double]$_, 3).ToString("0.000", [Globalization.CultureInfo]::InvariantCulture)
            } |
            Sort-Object Count -Descending |
            Select-Object -First $Top |
            ForEach-Object {
                [pscustomobject]@{
                    AdvancePoints = $_.Name
                    Count = $_.Count
                }
            }
    )
}

function New-RowAdvanceStats($Rows) {
    $advances = @(New-RowAdvances $Rows)
    if ($advances.Count -eq 0) {
        return [pscustomobject]@{
            Count = 0
            MeanAdvancePoints = $null
            MinAdvancePoints = $null
            MaxAdvancePoints = $null
            CountWithoutMax = 0
            MeanWithoutMaxAdvancePoints = $null
        }
    }

    $measure = $advances | Measure-Object -Average -Minimum -Maximum
    $withoutMax = @($advances | Where-Object { [double]$_ -lt [double]$measure.Maximum })
    $withoutMaxMeasure = if ($withoutMax.Count -eq 0) { $null } else { $withoutMax | Measure-Object -Average }
    return [pscustomobject]@{
        Count = $advances.Count
        MeanAdvancePoints = [Math]::Round([double]$measure.Average, 6)
        MinAdvancePoints = [Math]::Round([double]$measure.Minimum, 6)
        MaxAdvancePoints = [Math]::Round([double]$measure.Maximum, 6)
        CountWithoutMax = $withoutMax.Count
        MeanWithoutMaxAdvancePoints = if ($null -eq $withoutMaxMeasure) { $null } else { [Math]::Round([double]$withoutMaxMeasure.Average, 6) }
    }
}

function Get-LayoutLines($Layout) {
    $lines = New-Object System.Collections.Generic.List[object]
    for ($pageIndex = 0; $pageIndex -lt $Layout.Pages.Count; $pageIndex++) {
        $pageNumber = $pageIndex + 1
        if ($pageNumber -lt $PageStart -or $pageNumber -gt $PageEnd) {
            continue
        }

        $lineIndex = 0
        if ($IncludeStaticStories) {
            foreach ($item in @($Layout.Pages[$pageIndex].StaticItems)) {
                if ([string]$item.Kind -ne "StaticHeaderTextLine" -and [string]$item.Kind -ne "StaticFooterTextLine" -and [string]$item.Kind -ne "StaticTextLine") {
                    continue
                }

                $lines.Add([pscustomobject]@{
                    CandidatePage = $pageNumber
                    CandidatePageLine = $lineIndex++
                    CandidateLayoutY = [Math]::Round([double]$item.Y, 6)
                    CandidateLayoutX = [Math]::Round([double]$item.X, 6)
                    SourceBlockIndex = $item.SourceBlockIndex
                    SourceLineIndex = $item.SourceLineIndex
                    LayoutTextLength = [int]$item.TextLength
                    IsStaticStory = $true
                    StoryKind = if ([string]$item.Kind -eq "StaticHeaderTextLine") { "Header" } elseif ([string]$item.Kind -eq "StaticFooterTextLine") { "Footer" } else { "Static" }
                    StoryVariantType = $item.StoryVariantType
                    LineHeightPoints = if ($null -eq $item.LineHeightPoints) { $null } else { [Math]::Round([double]$item.LineHeightPoints, 6) }
                    AppliedBeforeSpacingPoints = if ($null -eq $item.AppliedBeforeSpacingPoints) { $null } else { [Math]::Round([double]$item.AppliedBeforeSpacingPoints, 6) }
                    SingleLineHeightPoints = if ($null -eq $item.SingleLineHeightPoints) { $null } else { [Math]::Round([double]$item.SingleLineHeightPoints, 6) }
                    ListLabelSingleLineHeightPoints = if ($null -eq $item.ListLabelSingleLineHeightPoints) { $null } else { [Math]::Round([double]$item.ListLabelSingleLineHeightPoints, 6) }
                    BodyWindowsLineHeightPoints = if ($null -eq $item.BodyWindowsLineHeightPoints) { $null } else { [Math]::Round([double]$item.BodyWindowsLineHeightPoints, 6) }
                    ListLabelWindowsLineHeightPoints = if ($null -eq $item.ListLabelWindowsLineHeightPoints) { $null } else { [Math]::Round([double]$item.ListLabelWindowsLineHeightPoints, 6) }
                    EffectiveLineSpacingFactor = if ($null -eq $item.EffectiveLineSpacingFactor) { $null } else { [Math]::Round([double]$item.EffectiveLineSpacingFactor, 6) }
                    LineSpacingFactorFloorApplied = $item.LineSpacingFactorFloorApplied
                    IsFirstParagraphLine = $item.IsFirstParagraphLine
                    PendingAfterSpacingPoints = if ($null -eq $item.PendingAfterSpacingPoints) { $null } else { [Math]::Round([double]$item.PendingAfterSpacingPoints, 6) }
                    ParagraphBeforeSpacingPoints = if ($null -eq $item.ParagraphBeforeSpacingPoints) { $null } else { [Math]::Round([double]$item.ParagraphBeforeSpacingPoints, 6) }
                    ParagraphAfterSpacingPoints = if ($null -eq $item.ParagraphAfterSpacingPoints) { $null } else { [Math]::Round([double]$item.ParagraphAfterSpacingPoints, 6) }
                    ContextualSpacingSuppressed = $item.ContextualSpacingSuppressed
                })
            }
        }

        foreach ($item in @($Layout.Pages[$pageIndex].Items)) {
            if ([string]$item.Kind -ne "TextLine") {
                continue
            }

            $lines.Add([pscustomobject]@{
                CandidatePage = $pageNumber
                CandidatePageLine = $lineIndex++
                CandidateLayoutY = [Math]::Round([double]$item.Y, 6)
                CandidateLayoutX = [Math]::Round([double]$item.X, 6)
                SourceBlockIndex = $item.SourceBlockIndex
                SourceLineIndex = $item.SourceLineIndex
                LayoutTextLength = [int]$item.TextLength
                IsStaticStory = $false
                StoryKind = "Body"
                StoryVariantType = $null
                LineHeightPoints = if ($null -eq $item.LineHeightPoints) { $null } else { [Math]::Round([double]$item.LineHeightPoints, 6) }
                AppliedBeforeSpacingPoints = if ($null -eq $item.AppliedBeforeSpacingPoints) { $null } else { [Math]::Round([double]$item.AppliedBeforeSpacingPoints, 6) }
                SingleLineHeightPoints = if ($null -eq $item.SingleLineHeightPoints) { $null } else { [Math]::Round([double]$item.SingleLineHeightPoints, 6) }
                ListLabelSingleLineHeightPoints = if ($null -eq $item.ListLabelSingleLineHeightPoints) { $null } else { [Math]::Round([double]$item.ListLabelSingleLineHeightPoints, 6) }
                BodyWindowsLineHeightPoints = if ($null -eq $item.BodyWindowsLineHeightPoints) { $null } else { [Math]::Round([double]$item.BodyWindowsLineHeightPoints, 6) }
                ListLabelWindowsLineHeightPoints = if ($null -eq $item.ListLabelWindowsLineHeightPoints) { $null } else { [Math]::Round([double]$item.ListLabelWindowsLineHeightPoints, 6) }
                EffectiveLineSpacingFactor = if ($null -eq $item.EffectiveLineSpacingFactor) { $null } else { [Math]::Round([double]$item.EffectiveLineSpacingFactor, 6) }
                LineSpacingFactorFloorApplied = $item.LineSpacingFactorFloorApplied
                IsFirstParagraphLine = $item.IsFirstParagraphLine
                PendingAfterSpacingPoints = if ($null -eq $item.PendingAfterSpacingPoints) { $null } else { [Math]::Round([double]$item.PendingAfterSpacingPoints, 6) }
                ParagraphBeforeSpacingPoints = if ($null -eq $item.ParagraphBeforeSpacingPoints) { $null } else { [Math]::Round([double]$item.ParagraphBeforeSpacingPoints, 6) }
                ParagraphAfterSpacingPoints = if ($null -eq $item.ParagraphAfterSpacingPoints) { $null } else { [Math]::Round([double]$item.ParagraphAfterSpacingPoints, 6) }
                ContextualSpacingSuppressed = $item.ContextualSpacingSuppressed
            })
        }
    }

    return $lines.ToArray()
}

function Find-NearestCandidateRow($Rows, $Line, $Used) {
    $best = $null
    $bestScore = [double]::PositiveInfinity
    foreach ($row in $Rows) {
        if ([int]$row.Page -ne [int]$Line.CandidatePage) {
            continue
        }

        if ($Used.ContainsKey([int]$row.GlobalIndex)) {
            continue
        }

        if ($row.TextLength -eq 0) {
            continue
        }

        $dy = [Math]::Abs([double]$row.Y - [double]$Line.CandidateLayoutY)
        if ($dy -gt $LineTolerance) {
            continue
        }

        $lengthDelta = [Math]::Abs([int]$row.TextLength - [int]$Line.LayoutTextLength)
        $score = $dy + ($lengthDelta * 0.01d)
        if ($score -lt $bestScore) {
            $bestScore = $score
            $best = $row
        }
    }

    return $best
}

function Find-ReferenceRow($ReferenceRows, $CandidateRow, $Used, [int] $MinimumGlobalIndex) {
    $allExactMatches = @($ReferenceRows | Where-Object {
        $_.TextHash -eq $CandidateRow.TextHash -and
        $_.NormalizedText -eq $CandidateRow.NormalizedText
    })
    $matches = @($ReferenceRows | Where-Object {
        $_.TextHash -eq $CandidateRow.TextHash -and
        $_.NormalizedText -eq $CandidateRow.NormalizedText -and
        [int]$_.GlobalIndex -ge $MinimumGlobalIndex -and
        -not $Used.ContainsKey([int]$_.GlobalIndex)
    })

    if ($matches.Count -eq 0) {
        return [pscustomobject]@{
            Row = $null
            ExactReferenceMatchCount = $allExactMatches.Count
        }
    }

    $row = @($matches | Sort-Object -Property `
        @{ Expression = { [int]$_.GlobalIndex }; Descending = $false },
        @{ Expression = { [Math]::Abs([int]$_.Page - [int]$CandidateRow.Page) }; Descending = $false })[0]

    return [pscustomobject]@{
        Row = $row
        ExactReferenceMatchCount = $allExactMatches.Count
    }
}

$resolvedRun = (Resolve-Path -LiteralPath $RunDirectory).Path
$referencePath = Join-Path $resolvedRun "comparison\pdf-text\reference\text-operations.json"
$candidatePath = Join-Path $resolvedRun "comparison\pdf-text\candidate\text-operations.json"
if (-not (Test-Path -LiteralPath $referencePath) -or -not (Test-Path -LiteralPath $candidatePath)) {
    throw "Missing inspected text operations under '$resolvedRun\comparison\pdf-text'. Run tools\InspectPdf.ps1 or CheckPrivateCase first."
}

$layout = Read-JsonObject $LayoutSnapshot
$referenceRows = @(New-LineRows (Read-JsonArray $referencePath) | Where-Object { $_.Page -ge $PageStart -and $_.Page -le $PageEnd })
$candidateRows = @(New-LineRows (Read-JsonArray $candidatePath) | Where-Object { $_.Page -ge $PageStart -and $_.Page -le $PageEnd })
$layoutLines = @(Get-LayoutLines $layout)

$usedCandidate = @{}
$usedReference = @{}
$mapped = New-Object System.Collections.Generic.List[object]
$minimumReferenceGlobalIndex = 0

foreach ($line in $layoutLines) {
    $candidateRow = Find-NearestCandidateRow $candidateRows $line $usedCandidate
    if ($null -eq $candidateRow) {
        $mapped.Add([pscustomobject]@{
            Status = "missing-candidate-row"
            SourceBlockIndex = $line.SourceBlockIndex
            SourceLineIndex = $line.SourceLineIndex
            IsStaticStory = $line.IsStaticStory
            StoryKind = $line.StoryKind
            StoryVariantType = $line.StoryVariantType
            TextLength = $line.LayoutTextLength
            TextHash = $null
            ReferencePage = $null
            CandidatePage = $line.CandidatePage
            ReferenceY = $null
            CandidateY = $line.CandidateLayoutY
            DeltaY = $null
            DeltaPage = $null
            CandidateOperationCount = $null
            ReferenceOperationCount = $null
            ExactReferenceMatchCount = 0
            LineHeightPoints = $line.LineHeightPoints
            AppliedBeforeSpacingPoints = $line.AppliedBeforeSpacingPoints
            SingleLineHeightPoints = $line.SingleLineHeightPoints
            ListLabelSingleLineHeightPoints = $line.ListLabelSingleLineHeightPoints
            BodyWindowsLineHeightPoints = $line.BodyWindowsLineHeightPoints
            ListLabelWindowsLineHeightPoints = $line.ListLabelWindowsLineHeightPoints
            EffectiveLineSpacingFactor = $line.EffectiveLineSpacingFactor
            LineSpacingFactorFloorApplied = $line.LineSpacingFactorFloorApplied
            IsFirstParagraphLine = $line.IsFirstParagraphLine
            PendingAfterSpacingPoints = $line.PendingAfterSpacingPoints
            ParagraphBeforeSpacingPoints = $line.ParagraphBeforeSpacingPoints
            ParagraphAfterSpacingPoints = $line.ParagraphAfterSpacingPoints
            ContextualSpacingSuppressed = $line.ContextualSpacingSuppressed
        })
        continue
    }

    $usedCandidate[[int]$candidateRow.GlobalIndex] = $true
    $referenceMatch = Find-ReferenceRow $referenceRows $candidateRow $usedReference $minimumReferenceGlobalIndex
    $referenceRow = $referenceMatch.Row
    if ($null -ne $referenceRow) {
        $usedReference[[int]$referenceRow.GlobalIndex] = $true
        $minimumReferenceGlobalIndex = [int]$referenceRow.GlobalIndex + 1
    }

    $mapped.Add([pscustomobject]@{
        Status = if ($null -eq $referenceRow) { "missing-reference-match" } else { "matched" }
        SourceBlockIndex = $line.SourceBlockIndex
        SourceLineIndex = $line.SourceLineIndex
        IsStaticStory = $line.IsStaticStory
        StoryKind = $line.StoryKind
        StoryVariantType = $line.StoryVariantType
        TextLength = $candidateRow.TextLength
        TextHash = $candidateRow.TextHash
        ReferencePage = if ($null -eq $referenceRow) { $null } else { [int]$referenceRow.Page }
        CandidatePage = [int]$candidateRow.Page
        ReferenceY = if ($null -eq $referenceRow) { $null } else { [Math]::Round([double]$referenceRow.Y, 6) }
        CandidateY = [Math]::Round([double]$candidateRow.Y, 6)
        DeltaY = if ($null -eq $referenceRow) { $null } else { [Math]::Round([double]$candidateRow.Y - [double]$referenceRow.Y, 6) }
        DeltaPage = if ($null -eq $referenceRow) { $null } else { [int]$candidateRow.Page - [int]$referenceRow.Page }
        CandidateOperationCount = [int]$candidateRow.OperationCount
        ReferenceOperationCount = if ($null -eq $referenceRow) { $null } else { [int]$referenceRow.OperationCount }
        ExactReferenceMatchCount = [int]$referenceMatch.ExactReferenceMatchCount
        LineHeightPoints = $line.LineHeightPoints
        AppliedBeforeSpacingPoints = $line.AppliedBeforeSpacingPoints
        SingleLineHeightPoints = $line.SingleLineHeightPoints
        ListLabelSingleLineHeightPoints = $line.ListLabelSingleLineHeightPoints
        BodyWindowsLineHeightPoints = $line.BodyWindowsLineHeightPoints
        ListLabelWindowsLineHeightPoints = $line.ListLabelWindowsLineHeightPoints
        EffectiveLineSpacingFactor = $line.EffectiveLineSpacingFactor
        LineSpacingFactorFloorApplied = $line.LineSpacingFactorFloorApplied
        IsFirstParagraphLine = $line.IsFirstParagraphLine
        PendingAfterSpacingPoints = $line.PendingAfterSpacingPoints
        ParagraphBeforeSpacingPoints = $line.ParagraphBeforeSpacingPoints
        ParagraphAfterSpacingPoints = $line.ParagraphAfterSpacingPoints
        ContextualSpacingSuppressed = $line.ContextualSpacingSuppressed
    })
}

$summary = [pscustomobject]@{
    RunDirectory = $resolvedRun
    LayoutSnapshot = (Resolve-Path -LiteralPath $LayoutSnapshot).Path
    PageStart = $PageStart
    PageEnd = if ($PageEnd -eq [int]::MaxValue) { $null } else { $PageEnd }
    IncludeStaticStories = [bool]$IncludeStaticStories
    CandidateLayoutLineCount = $layoutLines.Count
    CandidateStaticLayoutLineCount = @($layoutLines | Where-Object { $_.IsStaticStory -eq $true }).Count
    CandidateBodyLayoutLineCount = @($layoutLines | Where-Object { $_.IsStaticStory -ne $true }).Count
    CandidatePdfLineCount = $candidateRows.Count
    ReferencePdfLineCount = $referenceRows.Count
    MatchedLineCount = @($mapped | Where-Object Status -eq "matched").Count
    MissingReferenceMatchCount = @($mapped | Where-Object Status -eq "missing-reference-match").Count
    MissingCandidateRowCount = @($mapped | Where-Object Status -eq "missing-candidate-row").Count
    AmbiguousReferenceMatchCount = @($mapped | Where-Object { $_.ExactReferenceMatchCount -gt 1 }).Count
    CandidateLineAdvanceBuckets = @(
        $layoutLines |
            Where-Object { $null -ne $_.LineHeightPoints } |
            Group-Object {
                $spacing = if ($null -eq $_.AppliedBeforeSpacingPoints) { 0d } else { [double]$_.AppliedBeforeSpacingPoints }
                [Math]::Round([double]$_.LineHeightPoints + $spacing, 3).ToString("0.000", [Globalization.CultureInfo]::InvariantCulture)
            } |
            Sort-Object Count -Descending |
            Select-Object -First $Top |
            ForEach-Object {
                [pscustomobject]@{
                    AdvancePoints = $_.Name
                    Count = $_.Count
                }
            }
    )
    ReferencePdfRowAdvanceBuckets = @(New-RowAdvanceBuckets $referenceRows)
    CandidatePdfRowAdvanceBuckets = @(New-RowAdvanceBuckets $candidateRows)
    ReferencePdfRowAdvanceStats = New-RowAdvanceStats $referenceRows
    CandidatePdfRowAdvanceStats = New-RowAdvanceStats $candidateRows
    CandidateEffectiveLineSpacingFactorBuckets = @(
        $layoutLines |
            Where-Object { $null -ne $_.EffectiveLineSpacingFactor } |
            Group-Object {
                ([double]$_.EffectiveLineSpacingFactor).ToString("0.000000", [Globalization.CultureInfo]::InvariantCulture)
            } |
            Sort-Object Count -Descending |
            Select-Object -First $Top |
            ForEach-Object {
                [pscustomobject]@{
                    EffectiveLineSpacingFactor = $_.Name
                    Count = $_.Count
                }
            }
    )
    CandidateLineSpacingFloorBuckets = @(
        $layoutLines |
            Where-Object { $null -ne $_.LineSpacingFactorFloorApplied } |
            Group-Object {
                if ($_.LineSpacingFactorFloorApplied -eq $true) { "true" } else { "false" }
            } |
            Sort-Object Count -Descending |
            Select-Object -First $Top |
            ForEach-Object {
                [pscustomobject]@{
                    LineSpacingFactorFloorApplied = $_.Name
                    Count = $_.Count
                }
            }
    )
    CandidateParagraphSpacingProfileBuckets = @(
        $layoutLines |
            Where-Object { $_.IsFirstParagraphLine -eq $true } |
            Group-Object {
                $pending = if ($null -eq $_.PendingAfterSpacingPoints) { "null" } else { ([double]$_.PendingAfterSpacingPoints).ToString("0.000", [Globalization.CultureInfo]::InvariantCulture) }
                $before = if ($null -eq $_.ParagraphBeforeSpacingPoints) { "null" } else { ([double]$_.ParagraphBeforeSpacingPoints).ToString("0.000", [Globalization.CultureInfo]::InvariantCulture) }
                $after = if ($null -eq $_.ParagraphAfterSpacingPoints) { "null" } else { ([double]$_.ParagraphAfterSpacingPoints).ToString("0.000", [Globalization.CultureInfo]::InvariantCulture) }
                $suppressed = if ($_.ContextualSpacingSuppressed -eq $true) { "true" } elseif ($_.ContextualSpacingSuppressed -eq $false) { "false" } else { "null" }
                "pendingAfter=$pending|before=$before|after=$after|suppressed=$suppressed"
            } |
            Sort-Object Count -Descending |
            Select-Object -First $Top |
            ForEach-Object {
                [pscustomobject]@{
                    Profile = $_.Name
                    Count = $_.Count
                }
            }
    )
    CandidateListLabelWindowsLineHeightBuckets = @(
        $layoutLines |
            Where-Object { $null -ne $_.ListLabelWindowsLineHeightPoints } |
            Group-Object {
                ([double]$_.ListLabelWindowsLineHeightPoints).ToString("0.000000", [Globalization.CultureInfo]::InvariantCulture)
            } |
            Sort-Object Count -Descending |
            Select-Object -First $Top |
            ForEach-Object {
                [pscustomobject]@{
                    ListLabelWindowsLineHeightPoints = $_.Name
                    Count = $_.Count
                }
            }
    )
    CandidateBodyWindowsLineHeightBuckets = @(
        $layoutLines |
            Where-Object { $null -ne $_.BodyWindowsLineHeightPoints } |
            Group-Object {
                ([double]$_.BodyWindowsLineHeightPoints).ToString("0.000000", [Globalization.CultureInfo]::InvariantCulture)
            } |
            Sort-Object Count -Descending |
            Select-Object -First $Top |
            ForEach-Object {
                [pscustomobject]@{
                    BodyWindowsLineHeightPoints = $_.Name
                    Count = $_.Count
                }
            }
    )
    WorstMatchedByAbsDeltaY = @(
        $mapped |
            Where-Object { $_.Status -eq "matched" } |
            Sort-Object -Property @{ Expression = { [Math]::Abs([double]$_.DeltaY) }; Descending = $true } |
            Select-Object -First $Top
    )
    MissingReferenceMatches = @(
        $mapped |
            Where-Object Status -eq "missing-reference-match" |
            Select-Object -First $Top
    )
}

if ($OutputJson) {
    $resolvedOutput = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputJson)
    $parent = Split-Path -Parent $resolvedOutput
    if ($parent -and -not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Path $parent | Out-Null
    }

    $summary | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $resolvedOutput -Encoding utf8
}
else {
    $summary | ConvertTo-Json -Depth 10
}
