param(
    [string[]] $Case = @(),

    [switch] $UngatedOnly,

    [switch] $SkipProbe
)

$ErrorActionPreference = "Stop"

function Read-JsonArray($path) {
    if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path -LiteralPath $path)) {
        return ,@()
    }

    $items = Get-Content -Raw -LiteralPath (Resolve-Path -LiteralPath $path).Path | ConvertFrom-Json
    if ($null -eq $items) {
        return ,@()
    }

    if ($items -is [array]) {
        return ,$items
    }

    return ,@($items)
}

function Has-AnyProperty($object, [string[]] $patterns) {
    if ($null -eq $object) {
        return $false
    }

    foreach ($property in $object.PSObject.Properties.Name) {
        foreach ($pattern in $patterns) {
            if ($property -match $pattern) {
                return $true
            }
        }
    }

    return $false
}

function Get-LatestRun($caseId) {
    $root = Join-Path "artifacts/visual" $caseId
    if (-not (Test-Path -LiteralPath $root)) {
        return $null
    }

    return Get-ChildItem -LiteralPath $root -Directory |
        Sort-Object Name -Descending |
        Select-Object -First 1
}

function Ensure-Probe($runPath, $side) {
    $output = Join-Path $runPath "probe-$side"
    $graphics = Join-Path $output "graphics-operations.json"
    $text = Join-Path $output "text-operations.json"
    $chartGraphics = Join-Path $output "chart-graphics-structures.json"
    $chartText = Join-Path $output "chart-text-structures.json"
    if ((Test-Path -LiteralPath $chartGraphics) -and (Test-Path -LiteralPath $chartText)) {
        return
    }

    if ($SkipProbe) {
        return
    }

    $pdf = if ($side -eq "reference") {
        Join-Path $runPath "reference/reference.pdf"
    }
    else {
        Join-Path $runPath "candidate/output.pdf"
    }

    if (-not (Test-Path -LiteralPath $pdf)) {
        return
    }

    & (Join-Path $PSScriptRoot "InspectPdf.ps1") -InputPdf $pdf -OutputDirectory $output -Page 1 *> $null
    & (Join-Path $PSScriptRoot "ClassifyPdfChartGraphics.ps1") -InputPath $graphics -Output $chartGraphics -PageNumber 1 *> $null
    & (Join-Path $PSScriptRoot "ClassifyPdfChartText.ps1") -InputPath $text -ChartStructures $chartGraphics -Output $chartText -PageNumber 1 *> $null
}

function CenterX($item) {
    return ([double]$item.MinX + [double]$item.MaxX) / 2d
}

function CenterY($item) {
    return ([double]$item.MinY + [double]$item.MaxY) / 2d
}

function BoundsDelta($reference, $candidate) {
    return @(
        [Math]::Abs([double]$reference.MinX - [double]$candidate.MinX),
        [Math]::Abs([double]$reference.MinY - [double]$candidate.MinY),
        [Math]::Abs([double]$reference.MaxX - [double]$candidate.MaxX),
        [Math]::Abs([double]$reference.MaxY - [double]$candidate.MaxY)) |
        Measure-Object -Maximum |
        Select-Object -ExpandProperty Maximum
}

function Summarize-Kind($kind, $referenceItems, $candidateItems, $setName) {
    $reference = @($referenceItems | Where-Object { $_.Kind -eq $kind })
    $candidate = @($candidateItems | Where-Object { $_.Kind -eq $kind })
    $maxDelta = $null
    if ($reference.Count -eq $candidate.Count -and $reference.Count -gt 0) {
        $unmatched = New-Object System.Collections.Generic.List[object]
        foreach ($item in $candidate) {
            $unmatched.Add($item)
        }

        foreach ($referenceItem in $reference) {
            $bestIndex = -1
            $bestScore = [double]::PositiveInfinity
            for ($i = 0; $i -lt $unmatched.Count; $i++) {
                $candidateItem = $unmatched[$i]
                $score = [Math]::Abs((CenterX $referenceItem) - (CenterX $candidateItem)) +
                    [Math]::Abs((CenterY $referenceItem) - (CenterY $candidateItem))
                if ($score -lt $bestScore) {
                    $bestScore = $score
                    $bestIndex = $i
                }
            }

            if ($bestIndex -ge 0) {
                $candidateItem = $unmatched[$bestIndex]
                $unmatched.RemoveAt($bestIndex)
                $delta = BoundsDelta $referenceItem $candidateItem
                if ($null -eq $maxDelta -or $delta -gt $maxDelta) {
                    $maxDelta = $delta
                }
            }
        }
    }

    [pscustomobject]@{
        Set = $setName
        Kind = $kind
        ReferenceCount = $reference.Count
        CandidateCount = $candidate.Count
        MaxBoundsDelta = if ($null -eq $maxDelta) { $null } else { [Math]::Round($maxDelta, 2) }
    }
}

$caseFiles = Get-ChildItem "visual-cases/cases" -Directory |
    Where-Object { $_.Name -like "pptx-ladder-11-*" } |
    Sort-Object Name |
    ForEach-Object { Get-ChildItem -LiteralPath $_.FullName -File -Filter "case.json" }

if ($Case.Count -gt 0) {
    $selected = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($caseId in $Case) {
        [void]$selected.Add($caseId)
    }

    $caseFiles = @($caseFiles | Where-Object { $selected.Contains($_.Directory.Name) })
}

foreach ($caseFile in $caseFiles) {
    $manifest = Get-Content -Raw -LiteralPath $caseFile.FullName | ConvertFrom-Json
    $expected = $manifest.expected
    $hasGraphicsGate = Has-AnyProperty $expected @("compareChartGraphicsStructure", "maxChartGraphicsStructure", "chartGraphics")
    $hasTextGate = Has-AnyProperty $expected @("compareChartTextStructure", "maxChartTextStructure", "chartText")
    if ($UngatedOnly -and ($hasGraphicsGate -or $hasTextGate)) {
        continue
    }

    $run = Get-LatestRun $caseFile.Directory.Name
    if ($null -eq $run) {
        Write-Warning "No visual run found for $($caseFile.Directory.Name)."
        continue
    }

    Ensure-Probe $run.FullName "reference"
    Ensure-Probe $run.FullName "candidate"

    Write-Host "== $($caseFile.Directory.Name)"
    Write-Host "Run: $($run.Name)  graphicsGate=$hasGraphicsGate textGate=$hasTextGate"
    foreach ($set in @(
        [pscustomobject]@{ Name = "GRAPHICS"; File = "chart-graphics-structures.json" },
        [pscustomobject]@{ Name = "TEXT"; File = "chart-text-structures.json" })) {
        $referenceItems = Read-JsonArray (Join-Path $run.FullName "probe-reference/$($set.File)")
        $candidateItems = Read-JsonArray (Join-Path $run.FullName "probe-candidate/$($set.File)")
        $kinds = @($referenceItems.Kind + $candidateItems.Kind | Where-Object { $_ } | Sort-Object -Unique)
        $rows = foreach ($kind in $kinds) {
            Summarize-Kind $kind $referenceItems $candidateItems $set.Name
        }

        $rows | Sort-Object Set, Kind | Format-Table -AutoSize
    }
}
