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

function TextX($op) {
    if ($UseEffectiveMatrix -and $op.EffectiveX -ne $null) {
        return [double]$op.EffectiveX
    }

    return [double]$op.X
}

function TextY($op) {
    if ($UseEffectiveMatrix -and $op.EffectiveY -ne $null) {
        return [double]$op.EffectiveY
    }

    return [double]$op.Y
}

function TextContent($op) {
    if ($op.DecodedText -ne $null) {
        return [string]$op.DecodedText
    }

    return [string]$op.Payload
}

function Merge-SameLineOperations($Operations) {
    $ordered = @($Operations | Sort-Object -Property @{ Expression = { [int]$_.PageNumber }; Ascending = $true }, @{ Expression = { TextY $_ }; Ascending = $true })
    $banded = New-Object System.Collections.Generic.List[object]
    $bandIndex = -1
    $bandPage = $null
    $bandY = 0d
    foreach ($op in $ordered) {
        # Keep the former 0.1pt grouping scale without rounding-edge splits.
        # Anchor each band once: small successive shifts cannot chain lines.
        if ($bandIndex -lt 0 -or $op.PageNumber -ne $bandPage -or [Math]::Abs((TextY $op) - $bandY) -gt 0.1d) {
            $bandIndex++
            $bandPage = $op.PageNumber
            $bandY = TextY $op
        }
        $banded.Add([pscustomobject]@{ Band=$bandIndex; Operation=$op })
    }
    $sorted = @($banded | Sort-Object Band, @{ Expression = { TextX $_.Operation }; Ascending = $true })
    $merged = New-Object System.Collections.Generic.List[object]
    $current = $null
    $currentLineKey = $null
    $currentEndX = 0d
    foreach ($entry in $sorted) {
        $op = $entry.Operation
        $lineKey = $entry.Band
        $opEndX = (TextX $op) + [double]$op.EmittedAdvancePoints
        if ($null -ne $current -and $currentLineKey -eq $lineKey -and ((TextX $op) - $currentEndX) -le 1d) {
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
        if ($current.PSObject.Properties.Name -notcontains "DecodedText") { $current | Add-Member -NotePropertyName DecodedText -NotePropertyValue $null }
        $currentLineKey = $lineKey
        $currentEndX = $opEndX
    }
    if ($null -ne $current) { $merged.Add($current) }
    return $merged.ToArray()
}
$referenceOps = Read-JsonArray $Reference
$candidateOps = Read-JsonArray $Candidate
if ($MergeSameLineOperations) {
    $referenceOps = Merge-SameLineOperations $referenceOps
    $candidateOps = Merge-SameLineOperations $candidateOps
}
if ($IgnoreWhitespaceOnlyLines) {
    if (-not $MergeSameLineOperations) { throw '-IgnoreWhitespaceOnlyLines requires -MergeSameLineOperations to preserve inline spacing.' }
    $referenceOps = @(Select-NonWhitespacePdfTextOperations $referenceOps)
    $candidateOps = @(Select-NonWhitespacePdfTextOperations $candidateOps)
}
$rows = New-Object System.Collections.Generic.List[object]
$failures = 0

if ($MatchByPosition) {
    $unmatched = New-Object System.Collections.Generic.List[object]
    foreach ($op in $referenceOps) {
        $unmatched.Add($op)
    }

    $pairs = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $candidateOps.Count; $i++) {
        $cand = $candidateOps[$i]
        $bestIndex = -1
        $bestScore = [double]::PositiveInfinity
        for ($j = 0; $j -lt $unmatched.Count; $j++) {
            $candidateRef = $unmatched[$j]
            if ($candidateRef.PageNumber -ne $cand.PageNumber) { continue }
            $score = [Math]::Abs((TextY $cand) - (TextY $candidateRef)) * 1000d +
                [Math]::Abs((TextX $cand) - (TextX $candidateRef))
            if ($score -lt $bestScore) {
                $bestScore = $score
                $bestIndex = $j
            }
        }

        $ref = if ($bestIndex -ge 0) { $unmatched[$bestIndex] } else { $null }
        if ($bestIndex -ge 0) {
            $unmatched.RemoveAt($bestIndex)
        }

        $pairs.Add([pscustomobject]@{ Index = $i; Reference = $ref; Candidate = $cand })
    }

    foreach ($ref in $unmatched) {
        $pairs.Add([pscustomobject]@{ Index = $pairs.Count; Reference = $ref; Candidate = $null })
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

foreach ($pair in $pairs) {
    $i = $pair.Index
    $ref = $pair.Reference
    $cand = $pair.Candidate
    if ($null -eq $ref -or $null -eq $cand) {
        $failures++
        $refX = if ($null -eq $ref) { $null } else { TextX $ref }
        $candX = if ($null -eq $cand) { $null } else { TextX $cand }
        $refY = if ($null -eq $ref) { $null } else { TextY $ref }
        $candY = if ($null -eq $cand) { $null } else { TextY $cand }
        $refTc = if ($null -eq $ref) { $null } else { $ref.CharacterSpacing }
        $candTc = if ($null -eq $cand) { $null } else { $cand.CharacterSpacing }
        $refPayload = if ($null -eq $ref) { $null } else { $ref.Payload }
        $candPayload = if ($null -eq $cand) { $null } else { $cand.Payload }
        $refText = if ($null -eq $ref -or -not $CompareDecodedText) { $null } else { TextContent $ref }
        $candText = if ($null -eq $cand -or -not $CompareDecodedText) { $null } else { TextContent $cand }
        $rows.Add([pscustomobject]@{
            Index = $i
            Status = "missing"
            RefPage = if ($null -eq $ref) { $null } else { $ref.PageNumber }
            CandPage = if ($null -eq $cand) { $null } else { $cand.PageNumber }
            RefFontSize = if ($null -eq $ref) { $null } else { $ref.FontSize }
            CandFontSize = if ($null -eq $cand) { $null } else { $cand.FontSize }
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
        })
        continue
    }

    $refX = TextX $ref
    $candX = TextX $cand
    $refY = TextY $ref
    $candY = TextY $cand
    $deltaX = Delta -left $refX -right $candX
    $deltaY = Delta -left $refY -right $candY
    $deltaSize = Delta -left ([double]$ref.FontSize) -right ([double]$cand.FontSize)
    $deltaTc = Delta -left ([double]$ref.CharacterSpacing) -right ([double]$cand.CharacterSpacing)
    $pageOk = $ref.PageNumber -eq $cand.PageNumber
    $positionOk = $pageOk -and [Math]::Abs($deltaX) -le $PositionTolerance -and [Math]::Abs($deltaY) -le $PositionTolerance
    $fontSizeOk = [Math]::Abs($deltaSize) -le $FontSizeTolerance
    $tcOk = [Math]::Abs($deltaTc) -le $CharacterSpacingTolerance
    $decodedTextOk = -not $CompareDecodedText -or (TextContent $ref) -eq (TextContent $cand)
    $status = if ($positionOk -and $fontSizeOk -and $tcOk -and $decodedTextOk) { "ok" } else { "delta" }
    if ($status -ne "ok") {
        $failures++
    }

    $rows.Add([pscustomobject]@{
        Index = $i
        Status = $status
        RefPage = $ref.PageNumber
        CandPage = $cand.PageNumber
        RefFontSize = $ref.FontSize
        CandFontSize = $cand.FontSize
        DeltaFontSize = $deltaSize
        RefX = [Math]::Round($refX, 6)
        CandX = [Math]::Round($candX, 6)
        DeltaX = $deltaX
        RefY = [Math]::Round($refY, 6)
        CandY = [Math]::Round($candY, 6)
        DeltaY = $deltaY
        RefTc = [Math]::Round([double]$ref.CharacterSpacing, 6)
        CandTc = [Math]::Round([double]$cand.CharacterSpacing, 6)
        DeltaTc = $deltaTc
        RefPayload = $ref.Payload
        CandPayload = $cand.Payload
        RefText = if ($CompareDecodedText) { TextContent $ref } else { $null }
        CandText = if ($CompareDecodedText) { TextContent $cand } else { $null }
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
    Write-Host "Text content: decoded text"
}

if ($failures -ne 0) {
    exit 1
}
