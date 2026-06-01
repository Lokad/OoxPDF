param(
    [Parameter(Mandatory = $true)]
    [string] $RunDirectory,

    [Parameter(Mandatory = $true)]
    [string] $LayoutSnapshot,

    [int] $PageStart = 1,

    [int] $PageEnd = [int]::MaxValue,

    [double] $LineTolerance = 0.75,

    [int] $Top = 80,

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

    $normalized = $Text.Replace([char]0x00A0, ' ').Replace("`r", " ").Replace("`n", " ").Replace("`t", " ")
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

function Get-LayoutLines($Layout) {
    $lines = New-Object System.Collections.Generic.List[object]
    for ($pageIndex = 0; $pageIndex -lt $Layout.Pages.Count; $pageIndex++) {
        $pageNumber = $pageIndex + 1
        if ($pageNumber -lt $PageStart -or $pageNumber -gt $PageEnd) {
            continue
        }

        $lineIndex = 0
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

function Find-ReferenceRow($ReferenceRows, $CandidateRow, $Used) {
    $matches = @($ReferenceRows | Where-Object {
        $_.TextHash -eq $CandidateRow.TextHash -and
        $_.NormalizedText -eq $CandidateRow.NormalizedText -and
        -not $Used.ContainsKey([int]$_.GlobalIndex)
    })

    if ($matches.Count -eq 0) {
        return $null
    }

    return @($matches | Sort-Object -Property `
        @{ Expression = { [Math]::Abs([int]$_.GlobalIndex - [int]$CandidateRow.GlobalIndex) }; Descending = $false },
        @{ Expression = { [Math]::Abs([int]$_.Page - [int]$CandidateRow.Page) }; Descending = $false })[0]
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

foreach ($line in $layoutLines) {
    $candidateRow = Find-NearestCandidateRow $candidateRows $line $usedCandidate
    if ($null -eq $candidateRow) {
        $mapped.Add([pscustomobject]@{
            Status = "missing-candidate-row"
            SourceBlockIndex = $line.SourceBlockIndex
            SourceLineIndex = $line.SourceLineIndex
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
        })
        continue
    }

    $usedCandidate[[int]$candidateRow.GlobalIndex] = $true
    $referenceRow = Find-ReferenceRow $referenceRows $candidateRow $usedReference
    if ($null -ne $referenceRow) {
        $usedReference[[int]$referenceRow.GlobalIndex] = $true
    }

    $mapped.Add([pscustomobject]@{
        Status = if ($null -eq $referenceRow) { "missing-reference-match" } else { "matched" }
        SourceBlockIndex = $line.SourceBlockIndex
        SourceLineIndex = $line.SourceLineIndex
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
    })
}

$summary = [pscustomobject]@{
    RunDirectory = $resolvedRun
    LayoutSnapshot = (Resolve-Path -LiteralPath $LayoutSnapshot).Path
    PageStart = $PageStart
    PageEnd = if ($PageEnd -eq [int]::MaxValue) { $null } else { $PageEnd }
    CandidateLayoutLineCount = $layoutLines.Count
    CandidatePdfLineCount = $candidateRows.Count
    ReferencePdfLineCount = $referenceRows.Count
    MatchedLineCount = @($mapped | Where-Object Status -eq "matched").Count
    MissingReferenceMatchCount = @($mapped | Where-Object Status -eq "missing-reference-match").Count
    MissingCandidateRowCount = @($mapped | Where-Object Status -eq "missing-candidate-row").Count
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
