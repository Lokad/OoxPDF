# Decoded whitespace has advance, but no glyph ink. Filter after merging runs
# when measuring line geometry, so spaces between words retain their advance.
# Missing/null decoded text is unknown and must remain in comparisons.
function Select-NonWhitespacePdfTextOperations($Operations) {
    foreach ($operation in @($Operations)) {
        if ($null -eq $operation) { continue }
        $decoded = $operation.PSObject.Properties['DecodedText']
        if ($null -eq $decoded -or $null -eq $decoded.Value -or -not [string]::IsNullOrWhiteSpace([string]$decoded.Value)) {
            $operation
        }
    }
}

function Test-PdfTextFragmentSeparation($Previous, $Current, [bool] $SeparatePrintedLines, [bool] $SeparateBaselines) {
    if ($null -eq $Previous.DecodedText -or $null -eq $Current.DecodedText -or
        $null -eq $Previous.EffectiveX -or $null -eq $Current.EffectiveX -or
        $null -eq $Previous.EffectiveY -or $null -eq $Current.EffectiveY -or
        $Previous.EffectiveB -ne 0 -or $Current.EffectiveB -ne 0 -or
        $Previous.EffectiveA -le 0 -or $Current.EffectiveA -le 0) {
        return $false
    }

    $em = [Math]::Max([double]$Previous.FontSize * $Previous.EffectiveA, [double]$Current.FontSize * $Current.EffectiveA)
    if (![double]::IsFinite($em) -or $em -le 0) { return $false }
    $baselineDrop = [double]$Previous.EffectiveY - [double]$Current.EffectiveY
    if ([Math]::Abs($baselineDrop) -gt 0.1d) {
        # Exact ordered witnesses occupy printed lines. A leftward reset on
        # a baseline at least one em lower can separate otherwise joined text.
        return $SeparateBaselines -or ($SeparatePrintedLines -and $baselineDrop -ge $em -and $Current.EffectiveX -le $Previous.EffectiveX)
    }

    # Reserve a full em after the previous advance before declaring a region
    # boundary. For an individual glyph with missing widths, allow another em
    # for the glyph itself. Multi-glyph shows need a measured advance.
    $advance = $Previous.EmittedAdvancePoints
    if ($null -eq $advance -and $Previous.DecodedRuneCount -eq 1) { $advance = $Previous.FontSize }
    if ($null -eq $advance -or $advance -le 0) { return $false }
    $step = ([double]$advance + [Math]::Max(0d, [double]$Previous.CharacterSpacing)) * $Previous.EffectiveA
    $gap = [double]$Current.EffectiveX - [double]$Previous.EffectiveX - $step
    return [double]::IsFinite($gap) -and $gap -gt $em
}

function Get-PdfTextFragmentRegion($Operation, $PageRegions) {
    if ($null -eq $Operation.EffectiveX -or $null -eq $Operation.EffectiveY -or
        $Operation.EffectiveA -le 0 -or $Operation.EffectiveB -ne 0) { return }

    $owner = $null
    foreach ($region in $PageRegions) {
        if ($Operation.EffectiveX -gt $region.MinX -and $Operation.EffectiveX -lt $region.MaxX -and
            $Operation.EffectiveY -gt $region.MinY -and $Operation.EffectiveY -lt $region.MaxY) {
            if ($null -ne $owner) { return }
            $owner = $region
        }
    }
    $owner
}

# Only page-owned operations can attest to a phrase. The inspector can also
# decode detached streams and Forms whose invocation it cannot resolve.
# Keep paint order and emitted separators within each text fragment. Pages,
# clear horizontal gaps and transitions between known regions end a fragment.
# Unknown glyphs retain their placeholder.
function Get-DecodedPdfTextFragments($Operations, $Regions, [switch] $SeparatePrintedLines, [switch] $SeparateBaselines) {
    $regionsByPage = ($Regions | Where-Object { $null -ne $_ } | Group-Object { [int]$_.Page } -AsHashTable) ?? @{}
    $pages = @($Operations | Where-Object {
        $number = $_.PageNumber
        ($number -is [int] -or $number -is [long]) -and $number -gt 0
    } | Group-Object PageNumber | Sort-Object { [int]$_.Name })
    foreach ($page in $pages) {
        $pageRegions = $regionsByPage[[int]$page.Name]
        $text = [Text.StringBuilder]::new()
        $previous = $null
        $previousRegion = $null
        foreach ($operation in $page.Group) {
            $decoded = $operation.DecodedText ?? [string][char]0xfffd
            if (!$decoded.Length) { continue }
            $region = if ($pageRegions) { Get-PdfTextFragmentRegion $operation $pageRegions }
            $changedRegion = $null -ne $region -and $null -ne $previousRegion -and
                ![object]::ReferenceEquals($region, $previousRegion)
            if ($changedRegion -or ($null -ne $previous -and
                (Test-PdfTextFragmentSeparation $previous $operation $SeparatePrintedLines $SeparateBaselines))) {
                $text.ToString()
                [void]$text.Clear()
            }
            [void]$text.Append($decoded)
            $previous = $operation
            # A positioning space outside the clips carries no text ownership.
            if ($null -ne $region -or ![string]::IsNullOrWhiteSpace($decoded)) {
                $previousRegion = $region
            }
        }
        if ($text.Length) { $text.ToString() }
    }
}

function New-PdfTextPhrasePattern([string] $Text) {
    # Punctuation already supplies a word edge. Treat undecoded glyphs as
    # possible word characters so they cannot supply a false boundary.
    $word = '[\p{L}\p{M}\p{N}\p{Pc}\uFFFD]'
    $pattern = [regex]::Escape($Text)
    if ($Text -cmatch ('^' + $word)) { $pattern = '(?<!' + $word + ')' + $pattern }
    if ($Text -cmatch ($word + '$')) { $pattern += '(?!' + $word + ')' }
    return [regex]$pattern
}

function Get-MissingPdfTextPhrases($Operations, $ExpectedText, $Regions, [switch] $SeparateBaselines) {
    function Normalize-PdfTextPhrase([string] $Text) {
        [regex]::Replace($Text.Normalize([Text.NormalizationForm]::FormKC), '\s+', ' ').Trim()
    }
    # Keep fragment boundaries out of the normalized phrase. Word edges are
    # significant; punctuation such as a label's colon already separates text.
    $fragments = @(Get-DecodedPdfTextFragments $Operations $Regions -SeparateBaselines:$SeparateBaselines |
        ForEach-Object { Normalize-PdfTextPhrase $_ })
    foreach ($expected in $ExpectedText) {
        $phrase = Normalize-PdfTextPhrase $expected
        if (!$phrase) { continue }
        $pattern = New-PdfTextPhrasePattern $phrase
        $found = $false
        foreach ($fragment in $fragments) {
            if ($pattern.IsMatch($fragment)) {
                $found = $true
                break
            }
        }
        if (!$found) { $expected }
    }
}

function Get-MissingOrderedPdfTextPhrases($Operations, $ExpectedText) {
    $fragments = @(Get-DecodedPdfTextFragments $Operations -SeparatePrintedLines)
    $nextFragment = 0
    $nextOffset = 0
    foreach ($expected in $ExpectedText) {
        $pattern = New-PdfTextPhrasePattern $expected
        $found = $false
        for ($i = $nextFragment; $i -lt $fragments.Count; $i++) {
            $offset = if ($i -eq $nextFragment) { $nextOffset } else { 0 }
            $match = $pattern.Match($fragments[$i], $offset)
            if (!$match.Success) { continue }
            $nextFragment = $i
            $nextOffset = $match.Index + $match.Length
            $found = $true
            break
        }
        if (!$found) { $expected }
    }
}

function Get-PdfTextOperationFontSize($Operation, [bool] $UseEffectiveMatrix) {
    if ($null -eq $Operation.FontSize) { return $null }
    $fontSize = [double]$Operation.FontSize
    if (!$UseEffectiveMatrix -or $null -eq $Operation.EffectiveA -or $null -eq $Operation.EffectiveB -or
        $null -eq $Operation.EffectiveC -or $null -eq $Operation.EffectiveD) { return $fontSize }

    $a = [double]$Operation.EffectiveA
    $b = [double]$Operation.EffectiveB
    $c = [double]$Operation.EffectiveC
    $d = [double]$Operation.EffectiveD
    $baselineLength = [Math]::Sqrt($a*$a + $b*$b)
    if ($baselineLength -eq 0d) { return [double]::NaN }
    # Perpendicular em height excludes shear and handles rotated baselines.
    return $fontSize * [Math]::Abs($a*$d - $b*$c) / $baselineLength
}

function Get-PdfTextOperationX($Operation, [bool] $UseEffectiveMatrix) {
    if ($UseEffectiveMatrix -and $null -ne $Operation.EffectiveX) {
        return [double]$Operation.EffectiveX
    }

    return [double]$Operation.X
}

function Get-PdfTextOperationY($Operation, [bool] $UseEffectiveMatrix) {
    if ($UseEffectiveMatrix -and $null -ne $Operation.EffectiveY) {
        return [double]$Operation.EffectiveY
    }

    return [double]$Operation.Y
}

function Merge-PdfSameLineOperations($Operations, [switch] $UseEffectiveMatrix) {
    $effective = $UseEffectiveMatrix.IsPresent
    $ordered = @($Operations | Sort-Object -Property @{ Expression = { [int]$_.PageNumber } }, @{ Expression = { Get-PdfTextOperationY $_ $effective } })
    $banded = [System.Collections.Generic.List[object]]::new()
    $bandIndex = -1
    $bandPage = $null
    $bandY = 0d
    foreach ($op in $ordered) {
        # Anchor each 0.1pt band once so small successive shifts cannot chain lines.
        $y = Get-PdfTextOperationY $op $effective
        if ($bandIndex -lt 0 -or $op.PageNumber -ne $bandPage -or [Math]::Abs($y - $bandY) -gt 0.1d) {
            $bandIndex++
            $bandPage = $op.PageNumber
            $bandY = $y
        }
        $banded.Add([pscustomobject]@{ Band = $bandIndex; Operation = $op })
    }

    $sorted = @($banded | Sort-Object Band, @{ Expression = { Get-PdfTextOperationX $_.Operation $effective } })
    $merged = [System.Collections.Generic.List[object]]::new()
    $current = $null
    $currentBand = -1
    $currentEndX = 0d
    foreach ($entry in $sorted) {
        $op = $entry.Operation
        $x = Get-PdfTextOperationX $op $effective
        # Project the end used for adjacency. Keep stored metrics in their original units.
        $scale = if ($effective -and $null -ne $op.EffectiveA) { [double]$op.EffectiveA } else { 1d }
        $opEndX = $x + [double]$op.EmittedAdvancePoints * $scale
        if ($null -ne $current -and $currentBand -eq $entry.Band -and $x - $currentEndX -le 1d) {
            $current.DecodedText = if ($null -eq $current.DecodedText -or $null -eq $op.DecodedText) { $null } else { [string]$current.DecodedText + [string]$op.DecodedText }
            $current.Payload = [string]$current.Payload + [string]$op.Payload
            $current.EmittedAdvancePoints = [double]$current.EmittedAdvancePoints + [double]$op.EmittedAdvancePoints
            $current.NaturalWidthPoints = [double]$current.NaturalWidthPoints + [double]$op.NaturalWidthPoints
            $current.AdjustmentTotalPoints = [double]$current.AdjustmentTotalPoints + [double]$op.AdjustmentTotalPoints
            $current.NetSpacingGapTotalPoints = [double]$current.NetSpacingGapTotalPoints + [double]$op.NetSpacingGapTotalPoints
            $current.CharacterSpacingGapTotalPoints = [double]$current.CharacterSpacingGapTotalPoints + [double]$op.CharacterSpacingGapTotalPoints
            $current.DecodedRuneCount = [int]$current.DecodedRuneCount + [int]$op.DecodedRuneCount
            $current.TextChunkCount = [int]$current.TextChunkCount + [int]$op.TextChunkCount
            $currentEndX = [Math]::Max($currentEndX, $opEndX)
            continue
        }

        if ($null -ne $current) { $merged.Add($current) }
        $current = $op.PSObject.Copy()
        if ($current.PSObject.Properties.Name -notcontains "DecodedText") {
            $current | Add-Member -NotePropertyName DecodedText -NotePropertyValue $null
        }
        $currentBand = $entry.Band
        $currentEndX = $opEndX
    }
    if ($null -ne $current) { $merged.Add($current) }
    return $merged.ToArray()
}
