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

function Test-PdfTextHorizontalSeparation($Previous, $Current) {
    if ($null -eq $Previous.DecodedText -or $null -eq $Current.DecodedText -or
        $null -eq $Previous.EffectiveX -or $null -eq $Current.EffectiveX -or
        $null -eq $Previous.EffectiveY -or $null -eq $Current.EffectiveY -or
        $Previous.EffectiveB -ne 0 -or $Current.EffectiveB -ne 0 -or
        $Previous.EffectiveA -le 0 -or $Current.EffectiveA -le 0 -or
        [Math]::Abs([double]$Previous.EffectiveY - [double]$Current.EffectiveY) -gt 0.1d) {
        return $false
    }

    # Reserve a full em after the previous advance before declaring a region
    # boundary. For an individual glyph with missing widths, allow another em
    # for the glyph itself. Multi-glyph shows need a measured advance.
    $advance = $Previous.EmittedAdvancePoints
    if ($null -eq $advance -and $Previous.DecodedRuneCount -eq 1) { $advance = $Previous.FontSize }
    $em = [Math]::Max([double]$Previous.FontSize * $Previous.EffectiveA, [double]$Current.FontSize * $Current.EffectiveA)
    if ($null -eq $advance -or $advance -le 0 -or ![double]::IsFinite($em) -or $em -le 0) { return $false }
    $step = ([double]$advance + [Math]::Max(0d, [double]$Previous.CharacterSpacing)) * $Previous.EffectiveA
    $gap = [double]$Current.EffectiveX - [double]$Previous.EffectiveX - $step
    return [double]::IsFinite($gap) -and $gap -gt $em
}

# Keep paint order and emitted separators within each text fragment. A page
# break or clear horizontal gap ends the fragment. Unknown glyphs retain their
# placeholder so a phrase cannot match across them.
function Get-DecodedPdfTextFragments($Operations) {
    $pages = @($Operations | Group-Object { $_.PageNumber } | Sort-Object { [int]$_.Name })
    foreach ($page in $pages) {
        $text = [Text.StringBuilder]::new()
        $previous = $null
        foreach ($operation in $page.Group) {
            $decoded = $operation.DecodedText ?? [string][char]0xfffd
            if (!$decoded.Length) { continue }
            if ($null -ne $previous -and (Test-PdfTextHorizontalSeparation $previous $operation)) {
                $text.ToString()
                [void]$text.Clear()
            }
            [void]$text.Append($decoded)
            $previous = $operation
        }
        if ($text.Length) { $text.ToString() }
    }
}

function Get-MissingPdfTextPhrases($Operations, $ExpectedText) {
    function Normalize-PdfTextPhrase([string] $Text) {
        [regex]::Replace($Text.Normalize([Text.NormalizationForm]::FormKC), '\s+', ' ').Trim()
    }
    # Keep fragment boundaries out of the normalized phrase. Word edges are
    # significant; punctuation such as a label's colon already separates text.
    $fragments = @(Get-DecodedPdfTextFragments $Operations | ForEach-Object { Normalize-PdfTextPhrase $_ })
    $word = '[\p{L}\p{M}\p{N}\p{Pc}\uFFFD]'
    foreach ($expected in $ExpectedText) {
        $phrase = Normalize-PdfTextPhrase $expected
        if (!$phrase) { continue }
        $pattern = [regex]::Escape($phrase)
        if ($phrase -cmatch ('^' + $word)) { $pattern = '(?<!' + $word + ')' + $pattern }
        if ($phrase -cmatch ($word + '$')) { $pattern += '(?!' + $word + ')' }
        $found = $false
        foreach ($fragment in $fragments) {
            if ($fragment -cmatch $pattern) {
                $found = $true
                break
            }
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
