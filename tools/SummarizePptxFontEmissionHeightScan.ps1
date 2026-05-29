param(
    [Parameter(Mandatory = $true)]
    [string[]] $SummaryJson,

    [string] $OutputJson
)

$ErrorActionPreference = "Stop"

function Read-JsonArray([string] $Path) {
    $items = Get-Content -Raw -LiteralPath (Resolve-Path -LiteralPath $Path).Path | ConvertFrom-Json
    if ($null -eq $items) {
        return @()
    }

    if ($items -is [array]) {
        return $items
    }

    return @($items)
}

function Optional-Double($Value) {
    if ($Value -is [array]) {
        if ($Value.Count -eq 0) {
            return $null
        }

        $Value = $Value[0]
    }

    if ($null -eq $Value -or [string]$Value -eq "") {
        return $null
    }

    return [double]$Value
}

function Range-Text($Values) {
    $items = @($Values | Where-Object { $null -ne $_ } | Sort-Object)
    if ($items.Count -eq 0) {
        return ""
    }

    $first = [Math]::Round([double]$items[0], 6)
    $last = [Math]::Round([double]$items[$items.Count - 1], 6)
    return "$first..$last"
}

$summaryPaths = @(
    foreach ($path in $SummaryJson) {
        foreach ($item in ($path -split "[,;]")) {
            if (-not [string]::IsNullOrWhiteSpace($item)) {
                $item.Trim()
            }
        }
    }
)

$rows = foreach ($path in $summaryPaths) {
    $resolved = Resolve-Path -LiteralPath $path
    $items = @(Read-JsonArray $resolved.Path)
    if ($items.Count -eq 0) {
        continue
    }

    $secondary = @($items | Where-Object {
        $officeFontSize = Optional-Double $_.OfficeFontSize
        $firstGridFontSize = Optional-Double $_.FirstGridFontSize
        $null -ne $officeFontSize -and
            $null -ne $firstGridFontSize -and
            [Math]::Abs($officeFontSize - $firstGridFontSize) -gt 0.001
    })

    [pscustomobject]@{
        Case = Split-Path -Leaf (Split-Path -Parent $resolved.Path)
        SummaryJson = $resolved.Path
        SlideHeight = Optional-Double $items[0].SlideHeight
        RowCount = $items.Count
        SecondaryCount = $secondary.Count
        SecondaryOfficeFontSizes = @($secondary | ForEach-Object { Optional-Double $_.OfficeFontSize } | Sort-Object -Unique)
        SecondarySourceYRange = Range-Text ($secondary | ForEach-Object { Optional-Double $_.Y })
        SecondaryTopBaselineRange = Range-Text ($secondary | ForEach-Object { Optional-Double $_.RefBaselineFromPageTop })
        SecondaryBottomBaselineRange = Range-Text ($secondary | ForEach-Object { Optional-Double $_.RefBaselineY })
        SecondaryTopBaselineRemainders = @($secondary | ForEach-Object { Optional-Double $_.RefBaselineFromPageTop600Remainder } | Sort-Object -Unique)
        SecondaryBottomBaselineRemainders = @($secondary | ForEach-Object { Optional-Double $_.RefBaselineY600Remainder } | Sort-Object -Unique)
    }
}

$rows = @($rows | Sort-Object SlideHeight, Case)

if (-not [string]::IsNullOrWhiteSpace($OutputJson)) {
    $parent = Split-Path -Parent $OutputJson
    if (-not [string]::IsNullOrWhiteSpace($parent)) {
        New-Item -ItemType Directory -Force -Path $parent | Out-Null
    }

    $rows | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $OutputJson -Encoding UTF8
}

$rows | Format-Table -AutoSize
