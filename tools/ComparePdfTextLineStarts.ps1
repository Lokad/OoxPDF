param(
    [Parameter(Mandatory = $true)]
    [string] $Reference,

    [Parameter(Mandatory = $true)]
    [string] $Candidate,

    [double] $LineTolerance = 0.75,

    [double] $StartTolerance = 0.1,

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

function New-LineGroups($operations) {
    $ordered = @($operations | Sort-Object -Property @{ Expression = { [double]$_.Y }; Descending = $true }, @{ Expression = { [double]$_.X }; Descending = $false })
    $groups = New-Object System.Collections.Generic.List[object]

    foreach ($op in $ordered) {
        $line = $null
        foreach ($candidateLine in $groups) {
            if ([Math]::Abs([double]$op.Y - [double]$candidateLine.Y) -le $LineTolerance) {
                $line = $candidateLine
                break
            }
        }

        if ($null -eq $line) {
            $line = [pscustomobject]@{
                Y = [double]$op.Y
                Operations = New-Object System.Collections.Generic.List[object]
            }
            $groups.Add($line)
        }

        $line.Operations.Add($op)
        $line.Y = ($line.Operations | Measure-Object -Property Y -Average).Average
    }

    $index = 0
    foreach ($line in $groups | Sort-Object -Property @{ Expression = { [double]$_.Y }; Descending = $true }) {
        $starts = @($line.Operations | Sort-Object -Property X | ForEach-Object { [double]$_.X })
        [pscustomobject]@{
            Index = $index++
            Y = [Math]::Round([double]$line.Y, 6)
            Count = $starts.Count
            Starts = $starts
        }
    }
}

function Delta([double] $left, [double] $right) {
    return [Math]::Round($right - $left, 6)
}

$referenceLines = @(New-LineGroups (Read-JsonArray $Reference))
$candidateLines = @(New-LineGroups (Read-JsonArray $Candidate))

if ($MatchByPosition) {
    $unmatched = New-Object System.Collections.Generic.List[object]
    foreach ($line in $referenceLines) {
        $unmatched.Add($line)
    }

    $pairs = New-Object System.Collections.Generic.List[object]
    foreach ($cand in $candidateLines) {
        $bestIndex = -1
        $bestDelta = [double]::PositiveInfinity
        for ($i = 0; $i -lt $unmatched.Count; $i++) {
            $delta = [Math]::Abs([double]$cand.Y - [double]$unmatched[$i].Y)
            if ($delta -lt $bestDelta) {
                $bestDelta = $delta
                $bestIndex = $i
            }
        }

        $ref = if ($bestIndex -ge 0) { $unmatched[$bestIndex] } else { $null }
        if ($bestIndex -ge 0) {
            $unmatched.RemoveAt($bestIndex)
        }

        $pairs.Add([pscustomobject]@{ Reference = $ref; Candidate = $cand })
    }

    foreach ($ref in $unmatched) {
        $pairs.Add([pscustomobject]@{ Reference = $ref; Candidate = $null })
    }
}
else {
    $count = [Math]::Max($referenceLines.Count, $candidateLines.Count)
    $pairs = for ($i = 0; $i -lt $count; $i++) {
        [pscustomobject]@{
            Reference = if ($i -lt $referenceLines.Count) { $referenceLines[$i] } else { $null }
            Candidate = if ($i -lt $candidateLines.Count) { $candidateLines[$i] } else { $null }
        }
    }
}

$rows = New-Object System.Collections.Generic.List[object]
$failures = 0
$lineIndex = 0
foreach ($pair in $pairs) {
    $ref = $pair.Reference
    $cand = $pair.Candidate
    if ($null -eq $ref -or $null -eq $cand) {
        $failures++
        $rows.Add([pscustomobject]@{
            Line = $lineIndex++
            Status = "missing-line"
            RefY = if ($null -eq $ref) { $null } else { $ref.Y }
            CandY = if ($null -eq $cand) { $null } else { $cand.Y }
            RefCount = if ($null -eq $ref) { $null } else { $ref.Count }
            CandCount = if ($null -eq $cand) { $null } else { $cand.Count }
            MaxDeltaX = $null
            RefStarts = if ($null -eq $ref) { $null } else { ($ref.Starts | ForEach-Object { [Math]::Round($_, 6) }) -join ", " }
            CandStarts = if ($null -eq $cand) { $null } else { ($cand.Starts | ForEach-Object { [Math]::Round($_, 6) }) -join ", " }
        })
        continue
    }

    $startCount = [Math]::Max($ref.Starts.Count, $cand.Starts.Count)
    $maxDelta = 0d
    for ($i = 0; $i -lt [Math]::Min($ref.Starts.Count, $cand.Starts.Count); $i++) {
        $maxDelta = [Math]::Max($maxDelta, [Math]::Abs((Delta -left ([double]$ref.Starts[$i]) -right ([double]$cand.Starts[$i]))))
    }

    $status = "ok"
    if ($ref.Count -ne $cand.Count) {
        $status = "count"
    }
    elseif ($maxDelta -gt $StartTolerance) {
        $status = "delta"
    }

    if ($status -ne "ok") {
        $failures++
    }

    $rows.Add([pscustomobject]@{
        Line = $lineIndex++
        Status = $status
        RefY = [Math]::Round([double]$ref.Y, 6)
        CandY = [Math]::Round([double]$cand.Y, 6)
        RefCount = $ref.Count
        CandCount = $cand.Count
        MaxDeltaX = [Math]::Round($maxDelta, 6)
        RefStarts = ($ref.Starts | ForEach-Object { [Math]::Round($_, 6) }) -join ", "
        CandStarts = ($cand.Starts | ForEach-Object { [Math]::Round($_, 6) }) -join ", "
    })
}

$rows | Format-Table -AutoSize
Write-Host "Text line starts: reference=$($referenceLines.Count), candidate=$($candidateLines.Count), deltas=$failures"
if ($MatchByPosition) {
    Write-Host "Matching: nearest line position"
}

if ($failures -ne 0) {
    exit 1
}
