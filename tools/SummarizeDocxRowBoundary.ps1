param(
    [Parameter(Mandatory = $true)]
    [string] $RunDirectory,

    [Parameter(Mandatory = $true)]
    [string] $LayoutSnapshot,

    [int] $PageStart = 1,

    [int] $PageEnd = [int]::MaxValue,

    [double] $BottomWindowPoints = 120,

    [double] $RowSlackPoints = 2,

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

function Resolve-TextPath([string] $Run, [string] $Kind) {
    $candidates = @(
        (Join-Path $Run "comparison\pdf-text\$Kind\text-operations.json"),
        (Join-Path $Run "comparison\pdf-$Kind\text-operations.json")
    )

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            return (Resolve-Path -LiteralPath $candidate).Path
        }
    }

    throw "Missing inspected $Kind text operations under '$Run\comparison'. Run tools\InspectPdf.ps1 first."
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

function New-PdfRows($Operations) {
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($op in ($Operations | Sort-Object -Property `
        @{ Expression = { [int]$_.PageNumber }; Descending = $false },
        @{ Expression = { Get-Coordinate $_ "Y" }; Descending = $true },
        @{ Expression = { Get-Coordinate $_ "X" }; Descending = $false })) {
        $page = [int]$op.PageNumber
        $y = Get-Coordinate $op "Y"
        $row = $null
        foreach ($candidate in $rows) {
            if ([int]$candidate.Page -eq $page -and [Math]::Abs([double]$candidate.Y - $y) -le 0.75d) {
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

    foreach ($row in $rows) {
        $ops = @($row.Operations | Sort-Object -Property @{ Expression = { Get-Coordinate $_ "X" }; Descending = $false })
        $text = [string]::Concat(@($ops | ForEach-Object { Get-Text $_ }))
        $normalized = Normalize-Text $text
        [pscustomobject]@{
            Page = [int]$row.Page
            Y = [Math]::Round([double]$row.Y, 6)
            X = if ($ops.Count -eq 0) { $null } else { [Math]::Round((@($ops | ForEach-Object { Get-Coordinate $_ "X" }) | Measure-Object -Minimum).Minimum, 6) }
            OperationCount = $ops.Count
            TextLength = $normalized.Length
            TextHash = New-Hash $normalized
        }
    }
}

function Select-RowsForLayoutRow($PdfRows, $LayoutRow) {
    $top = [double]$LayoutRow.Y + [double]$LayoutRow.Height + $RowSlackPoints
    $bottom = [double]$LayoutRow.Y - $RowSlackPoints
    @($PdfRows | Where-Object {
        [int]$_.Page -eq [int]$LayoutRow.Page -and
        [double]$_.Y -le $top -and
        [double]$_.Y -ge $bottom
    } | Sort-Object -Property @{ Expression = { [double]$_.Y }; Descending = $true })
}

function New-CellVisualOwnershipBuckets($Cells) {
    @(
        @($Cells) |
            Group-Object {
                if ($_.PSObject.Properties.Name -contains "VisualOwnership" -and $null -ne $_.VisualOwnership) {
                    [string]$_.VisualOwnership
                }
                elseif ($_.PSObject.Properties.Name -contains "IsVerticalMergeContinuation" -and [bool]$_.IsVerticalMergeContinuation) {
                    "LegacyVerticalMergeContinuation"
                }
                else {
                    "LegacyOwnCell"
                }
            } |
            Sort-Object Name |
            ForEach-Object {
                [pscustomobject]@{
                    VisualOwnership = $_.Name
                    Count = $_.Count
                }
            }
    )
}

$resolvedRun = (Resolve-Path -LiteralPath $RunDirectory).Path
$layout = Read-JsonObject $LayoutSnapshot
$referenceRows = @(New-PdfRows (Read-JsonArray (Resolve-TextPath $resolvedRun "reference")))
$candidateRows = @(New-PdfRows (Read-JsonArray (Resolve-TextPath $resolvedRun "candidate")))

$layoutRows = New-Object System.Collections.Generic.List[object]
for ($pageIndex = 0; $pageIndex -lt $layout.Pages.Count; $pageIndex++) {
    $pageNumber = $pageIndex + 1
    if ($pageNumber -lt $PageStart -or $pageNumber -gt $PageEnd) {
        continue
    }

    $page = $layout.Pages[$pageIndex]
    $marginBottom = [double]$page.MarginBottom
    foreach ($row in @($page.TableRows)) {
        $rowBottom = [double]$row.Y
        $rowTop = [double]$row.Y + [double]$row.Height
        if ($rowBottom -gt $marginBottom + $BottomWindowPoints -and $rowTop -gt $marginBottom + $BottomWindowPoints) {
            continue
        }

        $cells = @($row.Cells)
        $visualOwnershipBuckets = @(New-CellVisualOwnershipBuckets $cells)
        $layoutRows.Add([pscustomobject]@{
            Page = $pageNumber
            TableIndex = [int]$row.TableIndex
            RowIndex = [int]$row.RowIndex
            PageRowIndex = [int]$row.PageRowIndex
            FragmentIndex = [int]$row.FragmentIndex
            FragmentCount = [int]$row.FragmentCount
            Y = [Math]::Round([double]$row.Y, 6)
            Height = [Math]::Round([double]$row.Height, 6)
            BottomOverflowPoints = [Math]::Round([double]$row.BottomOverflowPoints, 6)
            FirstBaselineY = $row.FirstBaselineY
            LastBaselineY = $row.LastBaselineY
            TextLineCount = [int]$row.TextLineCount
            TextLength = [int]$row.TextLength
            CellCount = [int]$row.CellCount
            VerticalMergeContinuationCellCount = @($cells | Where-Object {
                $_.PSObject.Properties.Name -contains "IsVerticalMergeContinuation" -and [bool]$_.IsVerticalMergeContinuation
            }).Count
            VisualOwnershipBuckets = $visualOwnershipBuckets
            CantSplit = [bool]$row.CantSplit
            CandidatePdfRows = @(Select-RowsForLayoutRow $candidateRows ([pscustomobject]@{
                Page = $pageNumber
                Y = $row.Y
                Height = $row.Height
            }))
        })
    }
}

$bottomTextRows = [pscustomobject][ordered]@{
    Reference = @($referenceRows | Where-Object {
        $_.Page -ge $PageStart -and $_.Page -le $PageEnd
    } | Group-Object Page | ForEach-Object {
        $pageNumber = [int]$_.Name
        $page = $layout.Pages[$pageNumber - 1]
        $marginBottom = [double]$page.MarginBottom
        $_.Group | Where-Object { [double]$_.Y -le $marginBottom + $BottomWindowPoints } | Sort-Object -Property @{ Expression = { [double]$_.Y }; Descending = $true }
    })
    Candidate = @($candidateRows | Where-Object {
        $_.Page -ge $PageStart -and $_.Page -le $PageEnd
    } | Group-Object Page | ForEach-Object {
        $pageNumber = [int]$_.Name
        $page = $layout.Pages[$pageNumber - 1]
        $marginBottom = [double]$page.MarginBottom
        $_.Group | Where-Object { [double]$_.Y -le $marginBottom + $BottomWindowPoints } | Sort-Object -Property @{ Expression = { [double]$_.Y }; Descending = $true }
    })
}

$summary = [pscustomobject]@{
    RunDirectory = $resolvedRun
    LayoutSnapshot = (Resolve-Path -LiteralPath $LayoutSnapshot).Path
    PageStart = $PageStart
    PageEnd = $PageEnd
    BottomWindowPoints = $BottomWindowPoints
    LayoutRows = @($layoutRows.ToArray())
    BottomTextRows = $bottomTextRows
}

$json = $summary | ConvertTo-Json -Depth 8
if ($OutputJson) {
    $parent = Split-Path -Parent $OutputJson
    if ($parent) {
        New-Item -ItemType Directory -Force -Path $parent | Out-Null
    }

    $json | Set-Content -LiteralPath $OutputJson -Encoding UTF8
}
else {
    $json
}
