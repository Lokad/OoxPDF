param(
    [switch] $AllowUncovered
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$caseRoot = Join-Path $repoRoot "visual-cases/cases"
$familyRoot = Join-Path $repoRoot "visual-cases/families"
$idPattern = "^[a-z0-9]+(-[a-z0-9]+)*$"
$classificationTags = @("locked", "locked-text-ops", "approximate", "needs-review")

function Add-Issue([System.Collections.Generic.List[string]] $issues, [string] $message) {
    $issues.Add($message) | Out-Null
}

function Read-JsonFile([string] $path) {
    try {
        return Get-Content -Raw -LiteralPath $path | ConvertFrom-Json
    }
    catch {
        throw "Invalid JSON in $path. $($_.Exception.Message)"
    }
}

$issues = [System.Collections.Generic.List[string]]::new()
$caseDirectories = @(Get-ChildItem -LiteralPath $caseRoot -Directory | Sort-Object Name)
$familyFiles = @(Get-ChildItem -LiteralPath $familyRoot -Filter "*.json" -File | Sort-Object Name)
$caseById = @{}
$families = @()

foreach ($caseDirectory in $caseDirectories) {
    $manifestPath = Join-Path $caseDirectory.FullName "case.json"
    if (-not (Test-Path -LiteralPath $manifestPath)) {
        Add-Issue $issues "Case '$($caseDirectory.Name)' is missing case.json."
        continue
    }

    $manifest = Read-JsonFile $manifestPath
    if ([string]::IsNullOrWhiteSpace($manifest.id)) {
        Add-Issue $issues "Case '$($caseDirectory.Name)' has no id."
        continue
    }

    if ($manifest.id -ne $caseDirectory.Name) {
        Add-Issue $issues "Case '$($caseDirectory.Name)' has mismatched id '$($manifest.id)'."
    }

    if ($manifest.id -notmatch $idPattern) {
        Add-Issue $issues "Case '$($manifest.id)' is not normalized kebab-case."
    }

    if ($caseById.ContainsKey($manifest.id)) {
        Add-Issue $issues "Duplicate case id '$($manifest.id)'."
    }
    else {
        $caseById[$manifest.id] = [pscustomobject]@{
            Id = [string]$manifest.id
            DirectoryName = $caseDirectory.Name
            DirectoryPath = $caseDirectory.FullName
            ManifestPath = $manifestPath
            Manifest = $manifest
        }
    }

    if ($manifest.kind -notin @("pptx", "docx")) {
        Add-Issue $issues "Case '$($manifest.id)' has unsupported kind '$($manifest.kind)'."
    }

    if ([string]::IsNullOrWhiteSpace($manifest.input)) {
        Add-Issue $issues "Case '$($manifest.id)' has no input path."
    }
    else {
        $inputPath = Join-Path $caseDirectory.FullName $manifest.input
        if (-not (Test-Path -LiteralPath $inputPath)) {
            Add-Issue $issues "Case '$($manifest.id)' input does not exist: $($manifest.input)."
        }
    }

    if ($null -eq $manifest.expected) {
        Add-Issue $issues "Case '$($manifest.id)' has no expected gates object."
    }

    foreach ($tag in @($manifest.tags)) {
        if ($tag -is [string] -and $tag.Length -ne 0 -and $tag -notmatch $idPattern) {
            Add-Issue $issues "Case '$($manifest.id)' tag '$tag' is not normalized kebab-case."
        }
    }
}

foreach ($familyFile in $familyFiles) {
    $family = Read-JsonFile $familyFile.FullName
    if ([string]::IsNullOrWhiteSpace($family.id)) {
        Add-Issue $issues "Family file '$($familyFile.Name)' has no id."
        continue
    }

    if ($familyFile.BaseName -ne $family.id) {
        Add-Issue $issues "Family file '$($familyFile.Name)' has mismatched id '$($family.id)'."
    }

    if ($family.id -notmatch $idPattern) {
        Add-Issue $issues "Family '$($family.id)' is not normalized kebab-case."
    }

    if ($family.kind -notin @("pptx", "docx")) {
        Add-Issue $issues "Family '$($family.id)' has unsupported kind '$($family.kind)'."
    }

    $patterns = @($family.casePatterns | Where-Object { $_ -is [string] -and $_.Length -ne 0 })
    if ($patterns.Count -eq 0) {
        Add-Issue $issues "Family '$($family.id)' has no casePatterns."
    }

    $families += [pscustomobject]@{
        Id = [string]$family.id
        Kind = [string]$family.kind
        Path = $familyFile.FullName
        Patterns = $patterns
    }
}

$ownership = @{}
foreach ($case in $caseById.Values) {
    $matches = @()
    foreach ($family in $families) {
        if ($family.Kind -ne $case.Manifest.kind) {
            continue
        }

        foreach ($pattern in $family.Patterns) {
            if ($case.Id -like $pattern) {
                $matches += $family.Id
                break
            }
        }
    }

    $ownership[$case.Id] = $matches
    if ($matches.Count -eq 0 -and -not $AllowUncovered) {
        Add-Issue $issues "Case '$($case.Id)' is not covered by any visual family."
    }
    elseif ($matches.Count -gt 1) {
        Add-Issue $issues "Case '$($case.Id)' is covered by multiple visual families: $($matches -join ', ')."
    }
}

foreach ($family in $families) {
    foreach ($pattern in $family.Patterns) {
        $matched = @($caseById.Values | Where-Object {
            $_.Manifest.kind -eq $family.Kind -and $_.Id -like $pattern
        })
        if ($matched.Count -eq 0) {
            Add-Issue $issues "Family '$($family.Id)' pattern '$pattern' matches no $($family.Kind) cases."
        }
    }
}

if ($issues.Count -ne 0) {
    $issues | ForEach-Object { Write-Error $_ }
    throw "Visual case catalog validation failed with $($issues.Count) issue(s)."
}

$classifiedCounts = @{}
foreach ($tag in $classificationTags) {
    $classifiedCounts[$tag] = 0
}
$classifiedCounts["unclassified"] = 0
foreach ($case in $caseById.Values) {
    $tags = @($case.Manifest.tags)
    $classification = @($classificationTags | Where-Object { $tags -contains $_ } | Select-Object -First 1)
    if ($classification.Count -eq 0) {
        $classifiedCounts["unclassified"]++
    }
    else {
        $classifiedCounts[$classification[0]]++
    }
}

Write-Host ("Validated {0} visual cases across {1} families." -f $caseById.Count, $families.Count)
Write-Host ("Classification tags: locked={0}, locked-text-ops={1}, approximate={2}, needs-review={3}, unclassified={4}" -f `
    $classifiedCounts["locked"],
    $classifiedCounts["locked-text-ops"],
    $classifiedCounts["approximate"],
    $classifiedCounts["needs-review"],
    $classifiedCounts["unclassified"])
