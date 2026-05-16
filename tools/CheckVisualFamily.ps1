param(
    [Parameter(Mandatory = $true)]
    [string] $Family,

    [switch] $List,

    [int] $Limit = 0
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$familyPath = Join-Path $repoRoot ("visual-cases/families/{0}.json" -f $Family)
if (-not (Test-Path -LiteralPath $familyPath)) {
    throw "Unknown visual family '$Family'. Expected $familyPath."
}

$familyManifest = Get-Content -Raw -LiteralPath $familyPath | ConvertFrom-Json
$caseRoot = Join-Path $repoRoot "visual-cases/cases"
$patterns = @($familyManifest.casePatterns)
$cases = @(
    Get-ChildItem -LiteralPath $caseRoot -Directory |
        Where-Object {
            $caseName = $_.Name
            $patterns | Where-Object { $caseName -like $_ }
        } |
        Sort-Object Name
)

if ($Limit -gt 0) {
    $cases = @($cases | Select-Object -First $Limit)
}

if ($List) {
    $cases | ForEach-Object { $_.Name }
    return
}

if ($cases.Count -eq 0) {
    throw "Visual family '$Family' did not match any cases."
}

Write-Host ("Visual family: {0} ({1} cases)" -f $Family, $cases.Count)
foreach ($case in $cases) {
    $manifest = Join-Path $case.FullName "case.json"
    Write-Host ("==> {0}" -f $case.Name)
    & (Join-Path $PSScriptRoot "CheckVisualCase.ps1") -Case $manifest
}
