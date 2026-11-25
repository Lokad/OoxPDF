param(
    [Parameter(Mandatory = $true)]
    [string] $Case,
    [string] $OutputPath,
    [ValidateSet('net10.0', 'net11.0')] [string] $Framework = 'net11.0',
    [string] $InspectorPath,
    [switch] $SkipBuild,
    [string] $ExpectedLibrarySha256
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$privateRoot = Join-Path $repoRoot 'private-cases'
$artifactRoot = Join-Path $repoRoot 'artifacts/private-visual'
. (Join-Path $PSScriptRoot 'PrivateGuard.ps1')

$caseFull = (Resolve-Path -LiteralPath $Case).Path
Assert-PrivateUntracked $caseFull 'Private case manifest'
$manifest = Get-Content -Raw -LiteralPath $caseFull | ConvertFrom-Json
$caseId = [string]$manifest.id
if ([string]::IsNullOrWhiteSpace($caseId)) { throw 'Private case manifest must contain an id.' }
if ($caseId.IndexOfAny([IO.Path]::GetInvalidFileNameChars()) -ge 0 -or $caseId.Contains('/') -or $caseId.Contains('\')) {
    throw 'Private case id must be a single filename-safe path segment.'
}
if ($null -ne $manifest.kind -and $manifest.kind -ne 'pptx') { throw 'Private case manifest kind must be pptx.' }
if ([string]::IsNullOrWhiteSpace($manifest.input)) { throw 'Private case manifest must contain an input path.' }

$inputFull = (Resolve-Path -LiteralPath (Join-Path (Split-Path -Parent $caseFull) $manifest.input)).Path
Assert-PrivateUntracked $inputFull 'Private case input'
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $artifactRoot "$caseId/inventory/$((Get-Date -Format 'yyyyMMdd-HHmmss')).json"
}
$outputFull = [IO.Path]::GetFullPath($OutputPath)
if (-not (Test-UnderDirectory $outputFull $artifactRoot)) { throw "Inventory output must be under $artifactRoot." }

$dll = if ($InspectorPath) { (Resolve-Path -LiteralPath $InspectorPath).Path }
    else { Join-Path $repoRoot "tools/Lokad.OoxPdf.PptxInspect/bin/Debug/$Framework/Lokad.OoxPdf.PptxInspect.dll" }
if (!$InspectorPath -and !$SkipBuild) {
    $project = Join-Path $repoRoot 'tools/Lokad.OoxPdf.PptxInspect/Lokad.OoxPdf.PptxInspect.csproj'
    & (Join-Path $PSScriptRoot 'EnsureDotnetBuild.ps1') -Framework $Framework -Project $project -OutputDll $dll -Description 'PPTX inventory'
}
$library = Join-Path (Split-Path -Parent $dll) 'Lokad.OoxPdf.dll'
$inspectorHash = (Get-FileHash -LiteralPath $dll).Hash
$libraryHash = (Get-FileHash -LiteralPath $library).Hash
if ($ExpectedLibrarySha256 -and $ExpectedLibrarySha256 -cne $libraryHash) {
    throw 'PPTX inspector library differs from the qualified converter.'
}

$json = & dotnet $dll --source-inventory $inputFull
if ($LASTEXITCODE -ne 0) { throw "PPTX source inventory failed with exit code $LASTEXITCODE." }
if ((Get-FileHash -LiteralPath $dll).Hash -cne $inspectorHash -or (Get-FileHash -LiteralPath $library).Hash -cne $libraryHash) {
    throw 'PPTX inspection runtime changed during execution.'
}

$inventory = $json | ConvertFrom-Json -AsHashtable
$summary = [ordered]@{ CaseId = $caseId; GeneratedAt = (Get-Date).ToString('o') }
foreach ($key in $inventory.Summary.Keys) { $summary[$key] = $inventory.Summary[$key] }
$inventory.Summary = $summary
New-Item -ItemType Directory -Force -Path ([IO.Path]::GetDirectoryName($outputFull)) | Out-Null
$inventory | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $outputFull -Encoding utf8
Write-Host "PPTX slide inventory: $outputFull"
$summary | ConvertTo-Json -Depth 4
