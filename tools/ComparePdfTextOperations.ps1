param(
    [Parameter(Mandatory = $true)]
    [string] $Reference,

    [Parameter(Mandatory = $true)]
    [string] $Candidate,

    [double] $PositionTolerance = 0.05,

    [double] $FontSizeTolerance = 0.01,

    [double] $CharacterSpacingTolerance = 0.001,

    [switch] $MatchByPosition,

    [switch] $UseEffectiveMatrix,

    [switch] $CompareDecodedText,

    [switch] $MergeSameLineOperations,

    [switch] $IgnoreWhitespaceOnlyLines,

    [string] $OutputPath
)

$ErrorActionPreference = "Stop"
if ($OutputPath -and (Test-Path -LiteralPath $OutputPath)) { throw 'Select a fresh structured report path' }

. (Join-Path $PSScriptRoot "JsonArray.ps1")

. (Join-Path $PSScriptRoot "CompareMath.ps1")

. (Join-Path $PSScriptRoot "PdfTextGeometry.ps1")

$effective = $UseEffectiveMatrix.IsPresent

function TextCharacterSpacing($op) {
    if (!$UseEffectiveMatrix -or $null -eq $op.EffectiveA -or $null -eq $op.EffectiveB) { return [double]$op.CharacterSpacing }
    $a = [double]$op.EffectiveA; $b = [double]$op.EffectiveB
    return [double]$op.CharacterSpacing * [Math]::Sqrt($a*$a + $b*$b)
}

function Test-ResolvedText($text) {
    # Empty strings can carry an ActualText continuation or advance only.
    return $text -is [string] -and !$text.Contains([char]0xfffd)
}

$referenceOps = Read-JsonArray $Reference
$candidateOps = Read-JsonArray $Candidate
if ($MergeSameLineOperations) {
    $referenceOps = Merge-PdfSameLineOperations $referenceOps -UseEffectiveMatrix:$UseEffectiveMatrix
    $candidateOps = Merge-PdfSameLineOperations $candidateOps -UseEffectiveMatrix:$UseEffectiveMatrix
}
if ($IgnoreWhitespaceOnlyLines) {
    if (-not $MergeSameLineOperations) { throw '-IgnoreWhitespaceOnlyLines requires -MergeSameLineOperations to preserve inline spacing.' }
    $referenceOps = @(Select-NonWhitespacePdfTextOperations $referenceOps)
    $candidateOps = @(Select-NonWhitespacePdfTextOperations $candidateOps)
}
$rows = New-Object System.Collections.Generic.List[object]
$failures = 0
$unresolvedTextCount = 0

if ($MatchByPosition) {
    $unmatched = New-Object System.Collections.Generic.List[object]
    foreach ($op in $referenceOps) {
        $unmatched.Add($op)
    }

    $pairs = New-Object System.Collections.Generic.List[object]
    $remainingCandidates = New-Object System.Collections.Generic.List[object]
    function Find-ReferenceIndex($cand, [bool] $RequireSameText) {
        $candidateText = $cand.DecodedText
        if ($RequireSameText -and !(Test-ResolvedText $candidateText)) { return -1 }
        $bestIndex = -1
        $bestScore = [double]::PositiveInfinity
        for ($j = 0; $j -lt $unmatched.Count; $j++) {
            $candidateRef = $unmatched[$j]
            if ($candidateRef.PageNumber -ne $cand.PageNumber) { continue }
            if ($RequireSameText -and ($candidateRef.DecodedText -isnot [string] -or
                $candidateRef.DecodedText -cne $candidateText)) { continue }
            $score = [Math]::Abs((Get-PdfTextOperationY $cand $effective) - (Get-PdfTextOperationY $candidateRef $effective)) * 1000d +
                [Math]::Abs((Get-PdfTextOperationX $cand $effective) - (Get-PdfTextOperationX $candidateRef $effective))
            if ($score -lt $bestScore) {
                $bestScore = $score
                $bestIndex = $j
            }
        }
        return $bestIndex
    }

    for ($i = 0; $i -lt $candidateOps.Count; $i++) {
        $cand = $candidateOps[$i]
        $bestIndex = Find-ReferenceIndex $cand ([bool]$CompareDecodedText)
        if ($CompareDecodedText -and $bestIndex -lt 0) {
            $remainingCandidates.Add([pscustomobject]@{Index=$i;Candidate=$cand})
            continue
        }

        $ref = if ($bestIndex -ge 0) { $unmatched[$bestIndex] } else { $null }
        if ($bestIndex -ge 0) {
            $unmatched.RemoveAt($bestIndex)
        }

        $pairs.Add([pscustomobject]@{ Index = $i; Reference = $ref; Candidate = $cand })
    }

    # Reserve exact decoded matches before geometry fallbacks can consume a
    # neighboring column's reference. Changed and unresolved content still reports.
    foreach ($pending in $remainingCandidates) {
        $bestIndex = Find-ReferenceIndex $pending.Candidate $false
        $ref = if ($bestIndex -ge 0) { $unmatched[$bestIndex] } else { $null }
        if ($bestIndex -ge 0) { $unmatched.RemoveAt($bestIndex) }
        $pairs.Add([pscustomobject]@{Index=$pending.Index;Reference=$ref;Candidate=$pending.Candidate})
    }

    foreach ($ref in $unmatched) {
        $pairs.Add([pscustomobject]@{ Index = $pairs.Count; Reference = $ref; Candidate = $null })
    }
    $pairs = @($pairs | Sort-Object Index)
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

foreach ($pair in $pairs) {
    $i = $pair.Index
    $ref = $pair.Reference
    $cand = $pair.Candidate
    $refText = if ($CompareDecodedText) { $ref.DecodedText }
    $candText = if ($CompareDecodedText) { $cand.DecodedText }
    $unresolvedText = $CompareDecodedText -and (
        ($null -ne $ref -and !(Test-ResolvedText $refText)) -or
        ($null -ne $cand -and !(Test-ResolvedText $candText)))
    if ($unresolvedText) { $unresolvedTextCount++ }
    if ($null -eq $ref -or $null -eq $cand) {
        $failures++
        $refX = if ($null -eq $ref) { $null } else { Get-PdfTextOperationX $ref $effective }
        $candX = if ($null -eq $cand) { $null } else { Get-PdfTextOperationX $cand $effective }
        $refY = if ($null -eq $ref) { $null } else { Get-PdfTextOperationY $ref $effective }
        $candY = if ($null -eq $cand) { $null } else { Get-PdfTextOperationY $cand $effective }
        $refTc = if ($null -eq $ref) { $null } else { TextCharacterSpacing $ref }
        $candTc = if ($null -eq $cand) { $null } else { TextCharacterSpacing $cand }
        $refPayload = if ($null -eq $ref) { $null } else { $ref.Payload }
        $candPayload = if ($null -eq $cand) { $null } else { $cand.Payload }
        $rows.Add([pscustomobject]@{
            Index = $i
            Status = "missing"
            RefPage = if ($null -eq $ref) { $null } else { $ref.PageNumber }
            CandPage = if ($null -eq $cand) { $null } else { $cand.PageNumber }
            RefFontSize = if ($null -eq $ref) { $null } else { [double](Get-PdfTextOperationFontSize $ref $effective) }
            CandFontSize = if ($null -eq $cand) { $null } else { [double](Get-PdfTextOperationFontSize $cand $effective) }
            RefRawFontSize = if ($null -eq $ref) { $null } else { $ref.FontSize }
            CandRawFontSize = if ($null -eq $cand) { $null } else { $cand.FontSize }
            DeltaFontSize = $null
            RefX = $refX
            CandX = $candX
            DeltaX = $null
            RefY = $refY
            CandY = $candY
            DeltaY = $null
            RefTc = $refTc
            CandTc = $candTc
            DeltaTc = $null
            RefPayload = $refPayload
            CandPayload = $candPayload
            RefText = $refText
            CandText = $candText
            UnresolvedText = [bool]$unresolvedText
        })
        continue
    }

    $refX = Get-PdfTextOperationX $ref $effective
    $candX = Get-PdfTextOperationX $cand $effective
    $refY = Get-PdfTextOperationY $ref $effective
    $candY = Get-PdfTextOperationY $cand $effective
    $deltaX = Delta -left $refX -right $candX
    $deltaY = Delta -left $refY -right $candY
    $refSize = [double](Get-PdfTextOperationFontSize $ref $effective)
    $candSize = [double](Get-PdfTextOperationFontSize $cand $effective)
    $refTc = TextCharacterSpacing $ref
    $candTc = TextCharacterSpacing $cand
    $deltaSize = Delta -left $refSize -right $candSize
    $deltaTc = Delta -left $refTc -right $candTc
    $pageOk = $ref.PageNumber -eq $cand.PageNumber
    $positionOk = $pageOk -and [Math]::Abs($deltaX) -le $PositionTolerance -and [Math]::Abs($deltaY) -le $PositionTolerance
    $fontSizeOk = [Math]::Abs($deltaSize) -le $FontSizeTolerance
    $tcOk = [Math]::Abs($deltaTc) -le $CharacterSpacingTolerance
    $decodedTextOk = -not $CompareDecodedText -or $refText -ceq $candText
    $status = if ($unresolvedText) { 'unresolved' }
        elseif ($positionOk -and $fontSizeOk -and $tcOk -and $decodedTextOk) { 'ok' }
        else { 'delta' }
    if ($status -ne "ok") {
        $failures++
    }

    $rows.Add([pscustomobject]@{
        Index = $i
        Status = $status
        RefPage = $ref.PageNumber
        CandPage = $cand.PageNumber
        RefFontSize = $refSize
        CandFontSize = $candSize
        RefRawFontSize = $ref.FontSize
        CandRawFontSize = $cand.FontSize
        DeltaFontSize = $deltaSize
        RefX = [Math]::Round($refX, 6)
        CandX = [Math]::Round($candX, 6)
        DeltaX = $deltaX
        RefY = [Math]::Round($refY, 6)
        CandY = [Math]::Round($candY, 6)
        DeltaY = $deltaY
        RefTc = [Math]::Round($refTc, 6)
        CandTc = [Math]::Round($candTc, 6)
        DeltaTc = $deltaTc
        RefPayload = $ref.Payload
        CandPayload = $cand.Payload
        RefText = $refText
        CandText = $candText
        UnresolvedText = [bool]$unresolvedText
    })
}

if ($OutputPath) {
    $reportParent = Split-Path -Parent ([IO.Path]::GetFullPath($OutputPath))
    New-Item -ItemType Directory -Force -Path $reportParent | Out-Null
    [ordered]@{
        ToolSha256 = (Get-FileHash -LiteralPath $PSCommandPath).Hash
        HelperSha256 = [ordered]@{
            JsonArray = (Get-FileHash -LiteralPath (Join-Path $PSScriptRoot 'JsonArray.ps1')).Hash
            CompareMath = (Get-FileHash -LiteralPath (Join-Path $PSScriptRoot 'CompareMath.ps1')).Hash
            PdfTextGeometry = (Get-FileHash -LiteralPath (Join-Path $PSScriptRoot 'PdfTextGeometry.ps1')).Hash
        }
        ReferenceSha256 = (Get-FileHash -LiteralPath $Reference).Hash
        CandidateSha256 = (Get-FileHash -LiteralPath $Candidate).Hash
        MatchByPosition = [bool]$MatchByPosition
        MatchingMethod = if (!$MatchByPosition) { 'Indexed' } elseif ($CompareDecodedText) { 'Page-local decoded text, then nearest position for unmatched operations' } else { 'Page-local nearest position' }
        UseEffectiveMatrix = [bool]$UseEffectiveMatrix
        CompareDecodedText = [bool]$CompareDecodedText
        MergeSameLineOperations = [bool]$MergeSameLineOperations
        IgnoreWhitespaceOnlyLines = [bool]$IgnoreWhitespaceOnlyLines
        PositionTolerance = $PositionTolerance
        FontSizeTolerance = $FontSizeTolerance
        CharacterSpacingTolerance = $CharacterSpacingTolerance
        ReferenceCount = $referenceOps.Count
        CandidateCount = $candidateOps.Count
        FailureCount = $failures
        UnresolvedTextCount = $unresolvedTextCount
        Rows = $rows.ToArray()
    } | ConvertTo-Json -Depth 7 | Set-Content -LiteralPath $OutputPath -Encoding utf8
}
$rows | Format-Table -AutoSize
Write-Host "Text operation count: reference=$($referenceOps.Count), candidate=$($candidateOps.Count), deltas=$failures"
if ($MatchByPosition) {
    Write-Host "Matching: nearest text position"
}
if ($UseEffectiveMatrix) {
    Write-Host "Coordinates: effective text matrix"
}
if ($CompareDecodedText) {
    Write-Host "Text content: decoded text ($unresolvedTextCount unresolved)"
}

if ($failures -ne 0) {
    exit 1
}

exit 0
