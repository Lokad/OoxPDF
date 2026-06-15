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

function Get-ShortSha256([byte[]] $Bytes, [int] $Length = 12) {
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        $hash = ([System.BitConverter]::ToString($sha256.ComputeHash($Bytes)) -replace "-", "").ToLowerInvariant()
        return $hash.Substring(0, [Math]::Min($Length, $hash.Length))
    }
    finally {
        $sha256.Dispose()
    }
}

function Get-ReferenceCacheKey(
    [string] $InputPath,
    [int] $DpiValue,
    [string] $CacheVariant)
{
    $inputHash = (Get-FileHash -LiteralPath $InputPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $renderReference = Join-Path $PSScriptRoot "RenderReference.ps1"
    $rasterizePdf = Join-Path $PSScriptRoot "RasterizePdf.ps1"
    $toolHashSource = @(
        (Get-FileHash -LiteralPath $renderReference -Algorithm SHA256).Hash
        (Get-FileHash -LiteralPath $rasterizePdf -Algorithm SHA256).Hash
    ) -join "|"
    $toolHash = Get-ShortSha256 ([System.Text.Encoding]::UTF8.GetBytes($toolHashSource)) 12
    $extension = [System.IO.Path]::GetExtension($InputPath).TrimStart(".").ToLowerInvariant()
    $variantKeyPart = ""
    if (-not [string]::IsNullOrWhiteSpace($CacheVariant)) {
        $variantHash = Get-ShortSha256 ([System.Text.Encoding]::UTF8.GetBytes($CacheVariant.Trim().ToLowerInvariant())) 12
        $variantKeyPart = "-variant" + $variantHash
    }

    "{0}-{1}-{2}{3}-dpi{4}" -f $extension, $inputHash.Substring(0, 24), $toolHash, $variantKeyPart, $DpiValue
}

function Read-PublicManifest([string] $Path) {
    $full = (Resolve-Path -LiteralPath $Path).Path
    $manifest = Get-Content -Raw -LiteralPath $full | ConvertFrom-Json
    if ($manifest.kind -ne "docx") {
        throw "DOCX markup reference requests only support DOCX cases: $full"
    }

    if ($manifest.PSObject.Properties.Name -notcontains "docxMarkup" -or [string]::IsNullOrWhiteSpace([string]$manifest.docxMarkup)) {
        throw "DOCX markup reference requests require docxMarkup in public case: $full"
    }

    [pscustomobject]@{
        Id = [string]$manifest.id
        Path = $full
        Manifest = $manifest
        PrivateSafe = $false
    }
}

function New-PrivateModeCaseInfo($Manifest, [string] $Path, [string] $Id, [string] $MarkupMode, [string] $MarkupGeometry) {
    $syntheticManifest = [pscustomobject]@{
        id = $Id
        kind = "docx"
        input = [string]$Manifest.input
        docxMarkup = $MarkupMode
        docxMarkupGeometry = $MarkupGeometry
    }

    [pscustomobject]@{
        Id = $Id
        Path = $Path
        Manifest = $syntheticManifest
        PrivateSafe = $true
    }
}

function Read-PrivateManifests([string] $Path) {
    $full = (Resolve-Path -LiteralPath $Path).Path
    $manifest = Get-Content -Raw -LiteralPath $full | ConvertFrom-Json
    if ($manifest.kind -ne "docx" -or [string]::IsNullOrWhiteSpace([string]$manifest.input)) {
        throw "Private DOCX markup reference requests require a DOCX private manifest with an input path."
    }

    if ($manifest.PSObject.Properties.Name -contains "docxMarkup" -and -not [string]::IsNullOrWhiteSpace([string]$manifest.docxMarkup)) {
        $case = Read-PublicManifest $Path
        $case.PrivateSafe = $true
        return ,$case
    }

    @(
        New-PrivateModeCaseInfo $manifest $full "private-docx-markup-final" "final" "preserve"
        New-PrivateModeCaseInfo $manifest $full "private-docx-markup-original" "original" "preserve"
        New-PrivateModeCaseInfo $manifest $full "private-docx-markup-simple" "simple" "preserve"
        New-PrivateModeCaseInfo $manifest $full "private-docx-markup-all" "all" "word-compatible"
    )
}

function Get-DefaultDocxMarkupCases {
    $caseRoot = Join-Path $repoRoot "visual-cases/cases"
    Get-ChildItem -LiteralPath $caseRoot -Directory |
        ForEach-Object {
            $manifestPath = Join-Path $_.FullName "case.json"
            if (-not (Test-Path -LiteralPath $manifestPath)) {
                return
            }

            $manifest = Get-Content -Raw -LiteralPath $manifestPath | ConvertFrom-Json
            $tags = @($manifest.tags)
            if ($manifest.kind -eq "docx" -and $tags -contains "docx-markup" -and
                $manifest.PSObject.Properties.Name -contains "docxMarkup" -and
                -not [string]::IsNullOrWhiteSpace([string]$manifest.docxMarkup)) {
                Read-PublicManifest $manifestPath
            }
        } |
        Sort-Object Id
}

function Get-CaseInputPath($CaseInfo) {
    $caseDirectory = Split-Path -Parent $CaseInfo.Path
    return (Resolve-Path -LiteralPath (Join-Path $caseDirectory $CaseInfo.Manifest.input)).Path
}

function Get-CaseCacheVariant($CaseInfo) {
    $docxMarkup = [string]$CaseInfo.Manifest.docxMarkup
    $docxMarkupGeometry = if ($CaseInfo.Manifest.PSObject.Properties.Name -contains "docxMarkupGeometry" -and -not [string]::IsNullOrWhiteSpace([string]$CaseInfo.Manifest.docxMarkupGeometry)) {
        [string]$CaseInfo.Manifest.docxMarkupGeometry
    }
    else {
        "preserve"
    }

    "docxMarkup={0};docxMarkupGeometry={1}" -f $docxMarkup, $docxMarkupGeometry
}

function ConvertTo-RepoPath([string] $Path) {
    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $fullRoot = [System.IO.Path]::GetFullPath($repoRoot).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    if ($fullPath.StartsWith($fullRoot + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $fullPath.Substring($fullRoot.Length + 1).Replace([System.IO.Path]::DirectorySeparatorChar, "/")
    }

    return $fullPath
}

function New-ReferenceRequestItem($CaseInfo, [int] $DpiValue) {
    $inputFull = Get-CaseInputPath $CaseInfo
    $inputSha256 = (Get-FileHash -LiteralPath $inputFull -Algorithm SHA256).Hash.ToLowerInvariant()
    $cacheVariant = Get-CaseCacheVariant $CaseInfo
    $cacheKey = Get-ReferenceCacheKey $inputFull $DpiValue $cacheVariant
    $cacheDirectory = Join-Path (Join-Path $repoRoot "artifacts/reference-cache") $cacheKey
    $completeMarker = Join-Path $cacheDirectory "complete.txt"
    $referencePdf = Join-Path $cacheDirectory "reference.pdf"
    $docxMarkup = [string]$CaseInfo.Manifest.docxMarkup
    $docxMarkupGeometry = if ($CaseInfo.Manifest.PSObject.Properties.Name -contains "docxMarkupGeometry" -and -not [string]::IsNullOrWhiteSpace([string]$CaseInfo.Manifest.docxMarkupGeometry)) {
        [string]$CaseInfo.Manifest.docxMarkupGeometry
    }
    else {
        "preserve"
    }
    $referencePdfPlaceholder = "<trusted-reference-pdf>"
    $caseArgument = if ($CaseInfo.PrivateSafe) { "<private-case-manifest>" } else { ConvertTo-RepoPath $CaseInfo.Path }
    $overrideArguments = if ($CaseInfo.PrivateSafe) {
        " -DocxMarkup $docxMarkup -DocxMarkupGeometry $docxMarkupGeometry"
    }
    else {
        ""
    }

    [pscustomobject]@{
        Id = $CaseInfo.Id
        PrivateSafe = [bool]$CaseInfo.PrivateSafe
        Case = if ($CaseInfo.PrivateSafe) { $CaseInfo.Id } else { ConvertTo-RepoPath $CaseInfo.Path }
        InputExtension = [System.IO.Path]::GetExtension($inputFull).TrimStart(".").ToLowerInvariant()
        InputSha256 = $inputSha256
        Dpi = $DpiValue
        DocxMarkup = $docxMarkup
        DocxMarkupGeometry = $docxMarkupGeometry
        CacheVariant = $cacheVariant
        CacheKey = $cacheKey
        ExpectedCacheDirectory = "artifacts/reference-cache/$cacheKey"
        CacheDirectoryExists = Test-Path -LiteralPath $cacheDirectory
        CompleteMarkerPresent = Test-Path -LiteralPath $completeMarker
        ReferencePdfPresent = Test-Path -LiteralPath $referencePdf
        Ready = (Test-Path -LiteralPath $completeMarker) -and (Test-Path -LiteralPath $referencePdf)
        ImportCommand = "pwsh tools/ImportDocxMarkupReferenceCache.ps1 -Case $caseArgument -ReferencePdf $referencePdfPlaceholder$overrideArguments -Dpi $DpiValue -Force"
    }
}

$cases = if ($Case.Count -ne 0) {
    @($Case | ForEach-Object { Read-PublicManifest $_ })
}
else {
    @(Get-DefaultDocxMarkupCases)
}

if ($IncludePrivate) {
    $privatePath = Join-Path $repoRoot $PrivateCase
    if (-not (Test-Path -LiteralPath $privatePath)) {
        throw "Private DOCX markup case was requested but does not exist."
    }

    $cases += @(Read-PrivateManifests $privatePath)
}

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

$selectedItems = @($cases | ForEach-Object { New-ReferenceRequestItem $_ $Dpi })
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
