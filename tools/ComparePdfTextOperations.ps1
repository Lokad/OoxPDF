param(
    [Parameter(Mandatory = $true)]
    [string] $Reference,

    [Parameter(Mandatory = $true)]
    [string] $Candidate,

    [double] $PositionTolerance = 0.05,

    [double] $FontSizeTolerance = 0.01,

    [double] $CharacterSpacingTolerance = 0.001,

    [switch] $MatchByPosition
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
    return [Math]::Round([double]$right - [double]$left, 6)
}

$referenceOps = Read-JsonArray $Reference
$candidateOps = Read-JsonArray $Candidate
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
            $score = [Math]::Abs([double]$cand.Y - [double]$candidateRef.Y) * 1000d +
                [Math]::Abs([double]$cand.X - [double]$candidateRef.X)
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
        $refX = if ($null -eq $ref) { $null } else { $ref.X }
        $candX = if ($null -eq $cand) { $null } else { $cand.X }
        $refY = if ($null -eq $ref) { $null } else { $ref.Y }
        $candY = if ($null -eq $cand) { $null } else { $cand.Y }
        $refTc = if ($null -eq $ref) { $null } else { $ref.CharacterSpacing }
        $candTc = if ($null -eq $cand) { $null } else { $cand.CharacterSpacing }
        $refPayload = if ($null -eq $ref) { $null } else { $ref.Payload }
        $candPayload = if ($null -eq $cand) { $null } else { $cand.Payload }
        $rows.Add([pscustomobject]@{
            Index = $i
            Status = "missing"
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
        })
        continue
    }

    $deltaX = Delta -left ([double]$ref.X) -right ([double]$cand.X)
    $deltaY = Delta -left ([double]$ref.Y) -right ([double]$cand.Y)
    $deltaSize = Delta -left ([double]$ref.FontSize) -right ([double]$cand.FontSize)
    $deltaTc = Delta -left ([double]$ref.CharacterSpacing) -right ([double]$cand.CharacterSpacing)
    $positionOk = [Math]::Abs($deltaX) -le $PositionTolerance -and [Math]::Abs($deltaY) -le $PositionTolerance
    $fontSizeOk = [Math]::Abs($deltaSize) -le $FontSizeTolerance
    $tcOk = [Math]::Abs($deltaTc) -le $CharacterSpacingTolerance
    $status = if ($positionOk -and $fontSizeOk -and $tcOk) { "ok" } else { "delta" }
    if ($status -ne "ok") {
        $failures++
    }

    $rows.Add([pscustomobject]@{
        Index = $i
        Status = $status
        RefX = [Math]::Round([double]$ref.X, 6)
        CandX = [Math]::Round([double]$cand.X, 6)
        DeltaX = $deltaX
        RefY = [Math]::Round([double]$ref.Y, 6)
        CandY = [Math]::Round([double]$cand.Y, 6)
        DeltaY = $deltaY
        RefTc = [Math]::Round([double]$ref.CharacterSpacing, 6)
        CandTc = [Math]::Round([double]$cand.CharacterSpacing, 6)
        DeltaTc = $deltaTc
        RefPayload = $ref.Payload
        CandPayload = $cand.Payload
    })
}

$rows | Format-Table -AutoSize
Write-Host "Text operation count: reference=$($referenceOps.Count), candidate=$($candidateOps.Count), deltas=$failures"
if ($MatchByPosition) {
    Write-Host "Matching: nearest text position"
}

if ($failures -ne 0) {
    exit 1
}
