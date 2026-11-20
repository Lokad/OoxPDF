# Run the live request, import, cache-status and lookup commands in a synthetic
# repository. The stub raster contains only bytes for cache mechanics.
[CmdletBinding()]
param([string] $OutputDirectory)
$ErrorActionPreference = "Stop"
if (!$OutputDirectory) {
    $OutputDirectory = Join-Path (Split-Path -Parent $PSScriptRoot) ("artifacts/docx-markup-identity-" + [guid]::NewGuid().ToString('N'))
}
if (Test-Path -LiteralPath $OutputDirectory) { throw "Use a fresh test directory" }
$root = [IO.Path]::GetFullPath($OutputDirectory)
$tools = Join-Path $root "tools"
New-Item -ItemType Directory -Path $tools -Force | Out-Null
foreach ($name in @(
    'DocxMarkupReference.ps1', 'ReferenceCache.ps1', 'PathHelpers.ps1', 'JsonArray.ps1', 'JsonObject.ps1',
    'PdfTextGeometry.ps1', 'PdfInfo.ps1', 'RenderCachedReference.ps1', 'NewDocxMarkupReferenceRequest.ps1',
    'RunDocxMarkupReferenceGate.ps1', 'ImportDocxMarkupReferenceCache.ps1', 'CompareCachedDocxMarkupReference.ps1'
)) {
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot $name) -Destination $tools
}
@'
param($InputPdf, $OutputDirectory, $Dpi, [switch] $SkipBuild)
[IO.File]::WriteAllText((Join-Path $OutputDirectory 'page-001.png'), 'synthetic markup raster')
'@ | Set-Content -LiteralPath (Join-Path $tools "RasterizePdf.ps1") -Encoding utf8
@'
throw 'Synthetic markup cache checks must never launch Office'
'@ | Set-Content -LiteralPath (Join-Path $tools "RenderReference.ps1") -Encoding utf8

function Write-Case([string] $Directory, [string] $Id, [string] $Mode, [string] $Geometry, [string] $Tag = 'docx-markup') {
    $Directory = (New-Item -ItemType Directory -Path $Directory -Force).FullName
    [IO.File]::WriteAllText((Join-Path $Directory 'synthetic.docx'), "case:$Id")
    $manifest = [ordered]@{ id = $Id; kind = 'docx'; input = 'synthetic.docx'; tags = @($Tag) }
    if ($Mode) { $manifest.docxMarkup = $Mode }
    if ($Geometry) { $manifest.docxMarkupGeometry = $Geometry }
    $manifest | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $Directory 'case.json') -Encoding utf8
}
function Read-Json([string] $Path) {
    Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
}

$checks = @()
Push-Location $root
try {
    Write-Case "visual-cases/cases/word's test" 'word' 'all' 'word-compatible'
    Write-Case 'visual-cases/cases/preserve' 'preserve' 'all' ''
    Write-Case 'visual-cases/cases/alias' 'alias' 'simple-markup' 'Office-Compatible-All-Markup'
    Write-Case 'visual-cases/cases/untagged' 'untagged' 'all' 'word-compatible' 'general'
    Write-Case 'private-cases/synthetic-private' 'synthetic-private-id' '' ''
    [IO.File]::WriteAllText((Join-Path $root 'reference example.pdf'), '%PDF-synthetic reference identity')
    $referencePin = (Get-FileHash -LiteralPath 'reference example.pdf').Hash
    $private = @{ IncludePrivate = $true; PrivateCase = 'private-cases/synthetic-private/case.json' }
    & ./tools/NewDocxMarkupReferenceRequest.ps1 -OutputDirectory (Join-Path $root 'request') @private
    $request = Read-Json 'request/summary.json'
    $public = @($request.Items | Where-Object { !$_.PrivateSafe })
    if (($public.Id -join ',') -cne 'alias,preserve,word' -or $request.ReadySelectedCaseCount -ne 0) {
        throw 'Default selection lost sorting, selected the untagged case or accepted a missing entry'
    }
    $privateRows = @($request.Items | Where-Object PrivateSafe)
    if ($privateRows.Count -ne 4 -or $privateRows[-1].CacheVariant -cne 'docxMarkup=all;docxMarkupGeometry=word-compatible' -or
        @($privateRows | Where-Object { $_.ImportCommand -notmatch '<private-case-manifest>' }).Count -ne 0 -or
        (Get-Content -LiteralPath 'request/summary.json' -Raw) -match 'synthetic-private') {
        throw 'The private variants or private-safe request output changed'
    }
    $checks += 'default-selection-and-private-variants'

    # Exercise the emitted commands, with synthetic, fixed paths under this test's own directory.
    $variants = @{
        word = 'docxMarkup=all;docxMarkupGeometry=word-compatible'
        preserve = 'docxMarkup=all;docxMarkupGeometry=preserve'
        alias = 'docxMarkup=simple;docxMarkupGeometry=word-compatible'
    }
    foreach ($row in $public) {
        $variant = $variants[$row.Id]
        if ($row.CacheVariant -cne $variant) { throw "Request selected the wrong variant for $($row.Id)" }
        $command = $row.ImportCommand.Replace('<trusted-reference-pdf>', "'reference example.pdf'")
        & ([scriptblock]::Create($command)) | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "Generated import failed for $($row.Id)" }
        $metadata = Read-Json (Join-Path $row.ExpectedCacheDirectory 'reference-metadata.json')
        if ($metadata.CacheVariant -cne $variant -or $metadata.IdentityKey -cne $row.CacheKey -or $metadata.ReferencePdfSha256 -ne $referencePin) {
            throw "Generated import used the wrong key or bytes for $($row.Id)"
        }
        & ./tools/CompareCachedDocxMarkupReference.ps1 -Case $row.Case -OutputDirectory (Join-Path $root ("lookup-" + $row.Id)) -ValidateOnly
        if ((Get-FileHash -LiteralPath ("lookup-" + $row.Id + "/reference/reference.pdf")).Hash -cne $referencePin) {
            throw "Cached comparison picked a different PDF for $($row.Id)"
        }
        $checks += "generated-import-and-cached-lookup-$($row.Id)"
    }
    $word = $public | Where-Object Id -CEQ 'word'
    $missingVariant='docxMarkup=original;docxMarkupGeometry=preserve'
    $wordInput=Join-Path (Split-Path -Parent $word.Case) 'synthetic.docx'
    foreach ($cacheOnly in @($true,$false)) {
        $miss=$null
        try {
            & ./tools/RenderCachedReference.ps1 -InputPath $wordInput -OutputDirectory (Join-Path $root "missing-variant-$cacheOnly") `
                -CacheVariant $missingVariant -CacheOnly:$cacheOnly -SkipBuild | Out-Null
        } catch { $miss=$_.Exception.Message }
        if (!$miss -or $miss -notlike '*explicit markup view settings*') {
            throw 'A missing markup variant must require explicit export before any worker launch'
        }
    }
    $checks += 'missing-variant-refuses-inherited-word-view'

    $overridePath = & ./tools/ImportDocxMarkupReferenceCache.ps1 -Case $word.Case -ReferencePdf 'reference example.pdf' -DocxMarkup 'simple-markup' -DocxMarkupGeometry 'reserve-markup-margin'
    $override = Read-Json (Join-Path $overridePath 'reference-metadata.json')
    if ($override.CacheVariant -cne 'docxMarkup=simple;docxMarkupGeometry=reserve-margin' -or $override.IdentityKey -ceq $word.CacheKey) {
        throw 'An explicit markup/geometry alias failed to override the manifest'
    }
    $checks += 'explicit-override'

    [IO.File]::WriteAllText((Join-Path $root (Join-Path $word.ExpectedCacheDirectory 'reference.pdf')), '%PDF-changed reference bytes')
    & ./tools/NewDocxMarkupReferenceRequest.ps1 -OutputDirectory (Join-Path $root 'missing') -MissingOnly @private
    & ./tools/RunDocxMarkupReferenceGate.ps1 -CacheStatusOnly -OutputRoot (Join-Path $root 'status') @private
    $missing = Read-Json 'missing/summary.json'
    $status = Read-Json 'status/summary.json'
    $statusWord = $status.Cases | Where-Object Id -CEQ 'word'
    $missingWord = $missing.Items | Where-Object Id -CEQ 'word'
    if ($status.CaseCount -ne 7 -or $status.ReadyCount -ne 2 -or $missing.ReadySelectedCaseCount -ne 2 -or $missing.CaseCount -ne 5 -or
        $statusWord.Ready -ne $false -or $missingWord.Ready -ne $false -or
        $statusWord.CacheState -cne 'Corrupt' -or $missingWord.CacheState -cne 'Corrupt' -or
        $statusWord.MissingImportCommand -cne $missingWord.ImportCommand -or
        (Get-Content -LiteralPath 'status/summary.json' -Raw) -match 'synthetic-private') {
        throw 'Request and cache status disagreed about verified readiness or private paths'
    }
    $checks += 'corrupt-cache-refused-by-both'
}
finally {
    Pop-Location
}
@{ Checks = $checks; SyntheticOnly = $true; Builds = 0; OfficeExports = 0 } |
    ConvertTo-Json | Set-Content -LiteralPath (Join-Path $root 'evidence.json') -Encoding utf8
"DOCX markup reference identity: $($checks.Count) checks passed."
