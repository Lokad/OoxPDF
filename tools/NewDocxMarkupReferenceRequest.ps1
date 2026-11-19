param(
    [string[]] $Case = @(),

    [string] $PrivateCase = "private-cases/docx-markup-track-changes.json",

    [string] $OutputDirectory,

    [int] $Dpi = 144,

    [switch] $IncludePrivate,

    [switch] $MissingOnly
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot

. (Join-Path $PSScriptRoot "ReferenceCache.ps1")

. (Join-Path $PSScriptRoot "PathHelpers.ps1")
. (Join-Path $PSScriptRoot "DocxMarkupReference.ps1")

$privatePath = if ($IncludePrivate) { Join-Path $repoRoot $PrivateCase }
$cases = @(Get-DocxMarkupReferenceCases $Case $privatePath)

if ($cases.Count -eq 0) {
    throw "No DOCX markup reference requests were selected."
}

$outputFull = if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    Join-Path $repoRoot ("artifacts/docx-markup-reference-requests/" + (Get-Date -Format "yyyyMMdd-HHmmss"))
}
else {
    [System.IO.Path]::GetFullPath($OutputDirectory)
}
New-Item -ItemType Directory -Force -Path $outputFull | Out-Null

$selectedItems = @($cases | ForEach-Object { Get-DocxMarkupReferenceStatus $_ $Dpi })
$items = if ($MissingOnly) {
    @($selectedItems | Where-Object { $_.Ready -ne $true })
}
else {
    $selectedItems
}
$summary = [pscustomobject]@{
    GeneratedAtUtc = [DateTime]::UtcNow.ToString("O", [Globalization.CultureInfo]::InvariantCulture)
    Dpi = $Dpi
    MissingOnly = [bool]$MissingOnly
    SelectedCaseCount = $selectedItems.Count
    ReadySelectedCaseCount = @($selectedItems | Where-Object { $_.Ready -eq $true }).Count
    MissingSelectedCaseCount = @($selectedItems | Where-Object { $_.Ready -ne $true }).Count
    CaseCount = $items.Count
    PrivateSafeCaseCount = @($items | Where-Object { $_.PrivateSafe -eq $true }).Count
    PublicCaseCount = @($items | Where-Object { $_.PrivateSafe -ne $true }).Count
    Items = $items
}
$summaryPath = Join-Path $outputFull "summary.json"
$commandsPath = Join-Path $outputFull "import-commands.ps1"
$summary | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $summaryPath -Encoding UTF8
@(
    "# Import trusted DOCX markup references after controlled external Office export."
    "# Replace <trusted-reference-pdf> and, for private entries, <private-case-manifest> before running."
    $items | ForEach-Object { $_.ImportCommand }
) | Set-Content -LiteralPath $commandsPath -Encoding UTF8

Write-Host ("DOCX markup reference request: {0} ({1} item(s), {2} missing of {3} selected)" -f $summaryPath, $items.Count, $summary.MissingSelectedCaseCount, $summary.SelectedCaseCount)
