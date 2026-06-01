param(
    [Parameter(Mandatory = $true)]
    [string] $Reference,

    [Parameter(Mandatory = $true)]
    [string] $Candidate,

    [double] $LineTolerance = 0.75,

    [double] $StartTolerance = 0.1,

    [switch] $MatchByPosition,

    [switch] $FirstStartOnly,

    [switch] $ShowText
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
    $ordered = @($operations | Sort-Object -Property @{ Expression = { [int]$_.PageNumber }; Descending = $false }, @{ Expression = { [double]$_.Y }; Descending = $true }, @{ Expression = { [double]$_.X }; Descending = $false })
    $groups = New-Object System.Collections.Generic.List[object]

    foreach ($op in $ordered) {
        $page = [int]$op.PageNumber
        $line = $null
        foreach ($candidateLine in $groups) {
            if ([int]$candidateLine.Page -eq $page -and [Math]::Abs([double]$op.Y - [double]$candidateLine.Y) -le $LineTolerance) {
                $line = $candidateLine
                break
            }
        }

        if ($null -eq $line) {
            $line = [pscustomobject]@{
                Page = $page
                Y = [double]$op.Y
                Operations = New-Object System.Collections.Generic.List[object]
            }
            $groups.Add($line)
        }

        $line.Operations.Add($op)
        $line.Y = ($line.Operations | Measure-Object -Property Y -Average).Average
    }

    $index = 0
    foreach ($line in $groups | Sort-Object -Property @{ Expression = { [int]$_.Page }; Descending = $false }, @{ Expression = { [double]$_.Y }; Descending = $true }) {
        $orderedOps = @($line.Operations | Sort-Object -Property X)
        $starts = @($orderedOps | ForEach-Object { [double]$_.X })
        [pscustomobject]@{
            Index = $index++
            Page = [int]$line.Page
            Y = [Math]::Round([double]$line.Y, 6)
            Count = $starts.Count
            Starts = $starts
            Texts = @($orderedOps | ForEach-Object { TextContent $_ })
        }
    }
}

function Delta([double] $left, [double] $right) {
    return [Math]::Round($right - $left, 6)
}

function TextContent($op) {
    if ($op.DecodedText -ne $null) {
        return [string]$op.DecodedText
    }

    return [string]$op.Payload
}

function Format-LineText($line) {
    if ($null -eq $line) {
        return $null
    }

    $parts = @($line.Texts | ForEach-Object { ShortText $_ })
    return $parts -join " | "
}

function ShortText([string] $text) {
    if ($null -eq $text) {
        return ""
    }

    $normalized = $text.Replace("`r", "\r").Replace("`n", "\n").Replace("`t", "\t")
    if ($normalized.Length -le 80) {
        return $normalized
    }

    return $normalized.Substring(0, 77) + "..."
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
            if ([int]$cand.Page -ne [int]$unmatched[$i].Page) {
                continue
            }

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
            RefPage = if ($null -eq $ref) { $null } else { $ref.Page }
            CandPage = if ($null -eq $cand) { $null } else { $cand.Page }
            RefY = if ($null -eq $ref) { $null } else { $ref.Y }
            CandY = if ($null -eq $cand) { $null } else { $cand.Y }
            RefCount = if ($null -eq $ref) { $null } else { $ref.Count }
            CandCount = if ($null -eq $cand) { $null } else { $cand.Count }
            MaxDeltaX = $null
            RefStarts = if ($null -eq $ref) { $null } else { ($ref.Starts | ForEach-Object { [Math]::Round($_, 6) }) -join ", " }
            CandStarts = if ($null -eq $cand) { $null } else { ($cand.Starts | ForEach-Object { [Math]::Round($_, 6) }) -join ", " }
            RefText = if ($ShowText) { Format-LineText $ref } else { $null }
            CandText = if ($ShowText) { Format-LineText $cand } else { $null }
        })
        continue
    }

    $maxDelta = 0d
    $startCount = if ($FirstStartOnly) { [Math]::Min(1, [Math]::Min($ref.Starts.Count, $cand.Starts.Count)) } else { [Math]::Min($ref.Starts.Count, $cand.Starts.Count) }
    for ($i = 0; $i -lt $startCount; $i++) {
        $maxDelta = [Math]::Max(
            $maxDelta,
            [Math]::Abs((Delta -left ([double]$ref.Starts[$i]) -right ([double]$cand.Starts[$i]))))
    }

    $status = "ok"
    if ($ref.Page -ne $cand.Page) {
        $status = "page"
    }
    elseif (-not $FirstStartOnly -and $ref.Count -ne $cand.Count) {
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
        RefPage = [int]$ref.Page
        CandPage = [int]$cand.Page
        RefY = [Math]::Round([double]$ref.Y, 6)
        CandY = [Math]::Round([double]$cand.Y, 6)
        RefCount = $ref.Count
        CandCount = $cand.Count
        MaxDeltaX = [Math]::Round($maxDelta, 6)
        RefStarts = ($ref.Starts | ForEach-Object { [Math]::Round($_, 6) }) -join ", "
        CandStarts = ($cand.Starts | ForEach-Object { [Math]::Round($_, 6) }) -join ", "
        RefText = if ($ShowText) { Format-LineText $ref } else { $null }
        CandText = if ($ShowText) { Format-LineText $cand } else { $null }
    })
}

if ($ShowText) {
    $rows |
        Select-Object Line, Status, RefPage, CandPage, RefY, CandY, RefCount, CandCount, MaxDeltaX, RefStarts, CandStarts |
        Format-Table -AutoSize
    foreach ($row in $rows) {
        Write-Host "Line $($row.Line) reference text: $($row.RefText)"
        Write-Host "Line $($row.Line) candidate text: $($row.CandText)"
    }
}
else {
    $rows |
        Select-Object Line, Status, RefPage, CandPage, RefY, CandY, RefCount, CandCount, MaxDeltaX, RefStarts, CandStarts |
        Format-Table -AutoSize
}
Write-Host "Text line starts: reference=$($referenceLines.Count), candidate=$($candidateLines.Count), deltas=$failures"
if ($MatchByPosition) {
    Write-Host "Matching: nearest line position"
}

if ($failures -ne 0) {
    exit 1
}
