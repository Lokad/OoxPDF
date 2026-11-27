param(
    [ValidateSet('net10.0','net11.0')] [string] $Framework = 'net11.0',
    [Parameter(Mandatory = $true)]
    [string] $Case,

    [switch] $ValidateOnly,

    [string] $CliPath
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$privateRoot = Join-Path $repoRoot "private-cases"

. (Join-Path $PSScriptRoot "PrivateGuard.ps1")

$binding = Read-PrivateCase $Case
$caseFull = $binding.CasePath
$caseId = $binding.Id
$inputFull = $binding.InputPath
$manifest = $binding.Manifest

$docxMarkup = if ($manifest.PSObject.Properties.Name -contains "docxMarkup" -and $manifest.docxMarkup -ne $null) { [string]$manifest.docxMarkup } else { $null }
$docxMarkupGeometry = if ($manifest.PSObject.Properties.Name -contains "docxMarkupGeometry" -and $manifest.docxMarkupGeometry -ne $null) { [string]$manifest.docxMarkupGeometry } else { $null }
if (-not [string]::IsNullOrWhiteSpace($docxMarkup) -and $manifest.kind -ne "docx") {
    throw "Private case '$caseId' sets docxMarkup but is not a DOCX case."
}

if (-not [string]::IsNullOrWhiteSpace($docxMarkupGeometry) -and $manifest.kind -ne "docx") {
    throw "Private case '$caseId' sets docxMarkupGeometry but is not a DOCX case."
}

if ($ValidateOnly) {
    Write-Host "Private case validation passed: $caseFull"
    Write-Host "Private input: $inputFull"
    return
}

$cliDll = Resolve-PrivateCaseCli $CliPath $Framework
$runId = Get-Date -Format "yyyyMMdd-HHmmss"
$runRoot = Join-Path $repoRoot ("artifacts/private-visual/{0}/{1}" -f $caseId, $runId)
$referenceDir = Join-Path $runRoot "reference"
$candidateDir = Join-Path $runRoot "candidate"
$comparisonDir = Join-Path $runRoot "comparison"

New-Item -ItemType Directory -Force -Path $referenceDir, $candidateDir, $comparisonDir | Out-Null

$candidatePdf = Join-Path $candidateDir "output.pdf"
$diagnostics = Join-Path $candidateDir "diagnostics.json"
$dpi = if ($manifest.dpi -ne $null) { [int]$manifest.dpi } else { 144 }

$candidateArgs = @("convert", $inputFull, $candidatePdf, "--diagnostics", $diagnostics)
if (-not [string]::IsNullOrWhiteSpace($docxMarkup)) {
    $candidateArgs += @("--docx-markup", $docxMarkup)
}

if (-not [string]::IsNullOrWhiteSpace($docxMarkupGeometry)) {
    $candidateArgs += @("--docx-markup-geometry", $docxMarkupGeometry)
}

Write-Host "[$((Get-Date).ToString('HH:mm:ss'))] Converting candidate PDF...";
dotnet $cliDll @candidateArgs
if ($LASTEXITCODE -ne 0) {
    throw "Candidate conversion failed with exit code $LASTEXITCODE."
}

$referenceArgs = @{
    InputPath = $inputFull
    OutputDirectory = $referenceDir
    Dpi = $dpi
}
if (-not [string]::IsNullOrWhiteSpace($docxMarkup) -or -not [string]::IsNullOrWhiteSpace($docxMarkupGeometry)) {
    . (Join-Path $PSScriptRoot "DocxMarkupReference.ps1")
    $referenceArgs.CacheOnly = $true
    $referenceArgs.CacheVariant = (Resolve-DocxMarkupReferenceSettings $manifest).Variant
}

Write-Host "[$((Get-Date).ToString('HH:mm:ss'))] Rendering reference...";
& (Join-Path $PSScriptRoot "RenderCachedReference.ps1") -Framework $Framework @referenceArgs
Write-Host "[$((Get-Date).ToString('HH:mm:ss'))] Rasterizing candidate...";
& (Join-Path $PSScriptRoot "RasterizePdf.ps1") -Framework $Framework -InputPdf $candidatePdf -OutputDirectory $candidateDir -Dpi $dpi

dotnet build (Join-Path $repoRoot "tools/Lokad.OoxPdf.VisualDiff/Lokad.OoxPdf.VisualDiff.csproj") -f $Framework --nologo
if ($LASTEXITCODE -ne 0) {
    throw "VisualDiff build failed with exit code $LASTEXITCODE."
}

$visualDiffDll = Join-Path $repoRoot "tools/Lokad.OoxPdf.VisualDiff/bin/Debug/$Framework/Lokad.OoxPdf.VisualDiff.dll"
Write-Host "[$((Get-Date).ToString('HH:mm:ss'))] Comparing pages...";
dotnet $visualDiffDll $referenceDir $candidateDir $comparisonDir
if ($LASTEXITCODE -ne 0) {
    throw "VisualDiff failed with exit code $LASTEXITCODE."
}

$assessment = @"
# Private visual assessment: $($manifest.id)

Input: $inputFull
Kind: $($manifest.kind)
Run: $runId
Agent rating: <0-5>

## Summary

<One paragraph summary of visual fidelity. Keep this file private.>

## Pages/slides reviewed

- Page 1: <rating>, <main differences>

## Major defects

1. <Defect, page, suspected feature>

## Diagnostics reviewed

<Note important warnings/errors from diagnostics.json. Do not copy private text into public issues.>

## Public-safe follow-up

<An anonymized implementation target that can be copied into public notes.>
"@

Set-Content -LiteralPath (Join-Path $runRoot "assessment.md") -Value $assessment
Write-Host "Private visual case artifacts: $runRoot"
