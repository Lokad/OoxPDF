param(
    [string[]] $Case = @(),

    [switch] $UngatedOnly,

    [switch] $SkipProbe,

    [switch] $ShowBounds,

    [switch] $ByRegion,

    [string[]] $Kind = @()
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

function Normalize-CaseIds([string[]] $caseIds) {
    $normalized = New-Object System.Collections.Generic.List[string]
    foreach ($caseId in $caseIds) {
        if ([string]::IsNullOrWhiteSpace($caseId)) {
            continue
        }

        foreach ($part in ($caseId -split '[,;]')) {
            $trimmed = $part.Trim()
            if (-not [string]::IsNullOrWhiteSpace($trimmed)) {
                $normalized.Add($trimmed)
            }
        }
    }

    return ,$normalized.ToArray()
}

function Normalize-Values([string[]] $values) {
    $normalized = New-Object System.Collections.Generic.List[string]
    foreach ($value in $values) {
        if ([string]::IsNullOrWhiteSpace($value)) {
            continue
        }

        foreach ($part in ($value -split '[,;]')) {
            $trimmed = $part.Trim()
            if (-not [string]::IsNullOrWhiteSpace($trimmed)) {
                $normalized.Add($trimmed)
            }
        }
    }

    return ,$normalized.ToArray()
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

function SignedDelta($reference, $candidate, [string]$property) {
    if ($null -eq $reference -or $null -eq $candidate) {
        return $null
    }

    return [Math]::Round([double]$candidate.$property - [double]$reference.$property, 2)
}

function Format-Bounds($item) {
    if ($null -eq $item) {
        return ""
    }

    return "({0},{1})-({2},{3})" -f `
        [Math]::Round([double]$item.MinX, 2),
        [Math]::Round([double]$item.MinY, 2),
        [Math]::Round([double]$item.MaxX, 2),
        [Math]::Round([double]$item.MaxY, 2)
}

function Get-RegionValue($item) {
    if ($null -eq $item -or -not ($item.PSObject.Properties.Name -contains "RegionIndex") -or $null -eq $item.RegionIndex) {
        return ""
    }

    return [string]$item.RegionIndex
}

function Make-SummaryKey([string]$kind, [string]$regionIndex) {
    return "$kind`u{1f}$regionIndex"
}

function Parse-SummaryKey([string]$key) {
    $parts = $key -split "`u{1f}", 2
    if ($parts.Count -eq 1) {
        return [pscustomobject]@{ Kind = $parts[0]; RegionIndex = "" }
    }

    return [pscustomobject]@{ Kind = $parts[0]; RegionIndex = $parts[1] }
}

function Summarize-Kind($kind, $regionIndex, $referenceItems, $candidateItems, $setName) {
    $reference = @($referenceItems | Where-Object {
        $_.Kind -eq $kind -and (-not $ByRegion -or (Get-RegionValue $_) -eq $regionIndex)
    })
    $candidate = @($candidateItems | Where-Object {
        $_.Kind -eq $kind -and (-not $ByRegion -or (Get-RegionValue $_) -eq $regionIndex)
    })
    $maxDelta = $null
    $maxReferenceItem = $null
    $maxCandidateItem = $null
    $referenceBounds = ""
    $candidateBounds = ""
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
                    $maxReferenceItem = $referenceItem
                    $maxCandidateItem = $candidateItem
                }
            }
        }
    }

    if ($null -ne $maxReferenceItem) {
        $referenceBounds = Format-Bounds $maxReferenceItem
    }
    elseif ($reference.Count -eq 1) {
        $referenceBounds = Format-Bounds $reference[0]
    }

    if ($null -ne $maxCandidateItem) {
        $candidateBounds = Format-Bounds $maxCandidateItem
    }
    elseif ($candidate.Count -eq 1) {
        $candidateBounds = Format-Bounds $candidate[0]
    }

    [pscustomobject]@{
        Set = $setName
        Kind = $kind
        RegionIndex = if ($ByRegion) { $regionIndex } else { $null }
        ReferenceCount = $reference.Count
        CandidateCount = $candidate.Count
        MaxBoundsDelta = if ($null -eq $maxDelta) { $null } else { [Math]::Round($maxDelta, 2) }
        DeltaMinX = SignedDelta $maxReferenceItem $maxCandidateItem "MinX"
        DeltaMinY = SignedDelta $maxReferenceItem $maxCandidateItem "MinY"
        DeltaMaxX = SignedDelta $maxReferenceItem $maxCandidateItem "MaxX"
        DeltaMaxY = SignedDelta $maxReferenceItem $maxCandidateItem "MaxY"
        ReferenceBounds = $referenceBounds
        CandidateBounds = $candidateBounds
    }
}

$caseFiles = Get-ChildItem "visual-cases/cases" -Directory |
    Where-Object { $_.Name -like "pptx-ladder-11-*" } |
    Sort-Object Name |
    ForEach-Object { Get-ChildItem -LiteralPath $_.FullName -File -Filter "case.json" }

$normalizedCases = Normalize-CaseIds $Case
$normalizedKinds = Normalize-Values $Kind
if ($normalizedCases.Count -gt 0) {
    $selected = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($caseId in $normalizedCases) {
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
        $keys = if ($ByRegion) {
            @(
                foreach ($item in @($referenceItems + $candidateItems)) {
                    if ($item.Kind) {
                        Make-SummaryKey ([string]$item.Kind) (Get-RegionValue $item)
                    }
                }
            ) | Sort-Object -Unique
        }
        else {
            @($referenceItems.Kind + $candidateItems.Kind | Where-Object { $_ } | Sort-Object -Unique)
        }
        if ($normalizedKinds.Count -gt 0) {
            $selectedKinds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            foreach ($kindValue in $normalizedKinds) {
                [void]$selectedKinds.Add($kindValue)
            }

            $keys = @($keys | Where-Object {
                $keyKind = if ($ByRegion) { (Parse-SummaryKey $_).Kind } else { [string]$_ }
                $selectedKinds.Contains($keyKind)
            })
        }

        $rows = foreach ($key in $keys) {
            if ($ByRegion) {
                $parsed = Parse-SummaryKey $key
                Summarize-Kind $parsed.Kind $parsed.RegionIndex $referenceItems $candidateItems $set.Name
            }
            else {
                Summarize-Kind $key "" $referenceItems $candidateItems $set.Name
            }
        }

        if ($ShowBounds) {
            if ($ByRegion) {
                $rows |
                    Sort-Object Set, Kind, RegionIndex |
                    Format-List Set, Kind, RegionIndex, ReferenceCount, CandidateCount, MaxBoundsDelta, DeltaMinX, DeltaMinY, DeltaMaxX, DeltaMaxY, ReferenceBounds, CandidateBounds
            }
            else {
                $rows |
                    Sort-Object Set, Kind |
                    Format-List Set, Kind, ReferenceCount, CandidateCount, MaxBoundsDelta, DeltaMinX, DeltaMinY, DeltaMaxX, DeltaMaxY, ReferenceBounds, CandidateBounds
            }
        }
        else {
            if ($ByRegion) {
                $rows |
                    Sort-Object Set, Kind, RegionIndex |
                    Format-Table Set, Kind, RegionIndex, ReferenceCount, CandidateCount, MaxBoundsDelta -AutoSize
            }
            else {
                $rows |
                    Sort-Object Set, Kind |
                    Format-Table Set, Kind, ReferenceCount, CandidateCount, MaxBoundsDelta -AutoSize
            }
        }
    }
}
