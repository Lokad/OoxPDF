param(
    [Parameter(Mandatory = $true)]
    [string[]] $SummaryJson,

    [string] $OutputJson
)

$ErrorActionPreference = "Stop"

function Expand-PathList([string[]] $Values) {
    $expanded = New-Object System.Collections.Generic.List[string]
    foreach ($value in $Values) {
        if ([string]::IsNullOrWhiteSpace($value)) {
            continue
        }

        foreach ($part in ($value -split "[,;]")) {
            if (-not [string]::IsNullOrWhiteSpace($part)) {
                $expanded.Add($part.Trim())
            }
        }
    }

    return ,$expanded.ToArray()
}

function Read-Json($Path) {
    return Get-Content -Raw -LiteralPath (Resolve-Path -LiteralPath $Path).Path | ConvertFrom-Json
}

function Get-Count($Counts, [string] $Key) {
    $entry = $Counts | Where-Object { $_.Key -eq $Key } | Select-Object -First 1
    if ($null -eq $entry) {
        return 0
    }

    return [int]$entry.Count
}

function Get-Extent($Extents, [string] $Key) {
    $extent = $Extents | Where-Object { $_.Key -eq $Key } | Select-Object -First 1
    if ($null -eq $extent) {
        return $null
    }

    return $extent
}

$paths = Expand-PathList $SummaryJson
if ($paths.Count -eq 0) {
    throw "No summary JSON paths were provided."
}

$rows = foreach ($path in $paths) {
    $resolved = (Resolve-Path -LiteralPath $path).Path
    $summary = Read-Json $resolved
    $caseName = Split-Path -Leaf (Split-Path -Parent $resolved)
    $secondary = Get-Extent $summary.FontBranchExtents "secondary-0.024"
    $main = Get-Extent $summary.FontBranchExtents "main-grid"

    [pscustomobject]@{
        Case = $caseName
        Total = [int]$summary.Total
        MainGrid = Get-Count $summary.FontBranchCounts "main-grid"
        Secondary = Get-Count $summary.FontBranchCounts "secondary-0.024"
        Missing = Get-Count $summary.FontBranchCounts "(missing)"
        MainRefTopMin = $main.MinRefBaselineFromPageTop
        MainRefTopMax = $main.MaxRefBaselineFromPageTop
        SecondaryRefTopMin = $secondary.MinRefBaselineFromPageTop
        SecondaryRefTopMax = $secondary.MaxRefBaselineFromPageTop
        SecondaryCandidateFrameTopMin = $secondary.MinCandidateFrameTopY
        SecondaryCandidateFrameTopMax = $secondary.MaxCandidateFrameTopY
        SecondaryCandidateLineTopMin = $secondary.MinCandidateLineTopY
        SecondaryCandidateLineTopMax = $secondary.MaxCandidateLineTopY
        Source = $resolved
    }
}

if (-not [string]::IsNullOrWhiteSpace($OutputJson)) {
    $outputFull = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputJson)
    $outputDirectory = Split-Path -Parent $outputFull
    if (-not [string]::IsNullOrWhiteSpace($outputDirectory)) {
        New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null
    }

    $rows | ConvertTo-Json -Depth 8 | Set-Content -Encoding UTF8 -LiteralPath $outputFull
}

$rows |
    Select-Object Case, Total, MainGrid, Secondary, Missing, SecondaryRefTopMin, SecondaryRefTopMax, SecondaryCandidateFrameTopMin, SecondaryCandidateFrameTopMax, SecondaryCandidateLineTopMin, SecondaryCandidateLineTopMax |
    Format-Table -AutoSize
