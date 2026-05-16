param(
    [Parameter(Mandatory = $true)]
    [string] $Id,

    [Parameter(Mandatory = $true)]
    [ValidateSet("pptx", "docx")]
    [string] $Kind,

    [Parameter(Mandatory = $true)]
    [string] $Input,

    [string] $Family
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$caseDir = Join-Path $repoRoot ("visual-cases/cases/{0}" -f $Id)
New-Item -ItemType Directory -Force -Path $caseDir | Out-Null

$manifest = [ordered]@{
    id = $Id
    kind = $Kind
    input = $Input
    dpi = 144
    tags = @($Kind, "smoke")
    expected = [ordered]@{
        minAgentRating = 3
        pageCountMustMatch = $true
        dimensionsMustMatch = $true
    }
    allowedUnsupportedFeatures = @()
}

$manifest | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $caseDir "case.json")
Write-Host "Created visual case: $caseDir"
if (-not [string]::IsNullOrWhiteSpace($Family)) {
    $familyPath = Join-Path $repoRoot ("visual-cases/families/{0}.json" -f $Family)
    if (-not (Test-Path -LiteralPath $familyPath)) {
        Write-Warning "Family '$Family' does not exist yet. Create $familyPath or choose an existing family."
    }
    else {
        Write-Host "Family: $Family"
    }
}
