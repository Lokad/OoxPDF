# Cache mechanics only; synthetic bytes never become Office fidelity references.
[CmdletBinding()]
param([string] $OutputDirectory)
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot
if (!$OutputDirectory) { $OutputDirectory=Join-Path $repo ('artifacts/reference-derivative-test-'+[guid]::NewGuid().ToString('N')) }
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Use a fresh test directory' }
$root=[IO.Path]::GetFullPath($OutputDirectory)
$tools=Join-Path $root 'tools'
New-Item -ItemType Directory -Path $tools -Force | Out-Null
foreach ($name in @('ReferenceCache.ps1','PathHelpers.ps1','RenderCachedReference.ps1','ImportOfficeReferenceCache.ps1')) {
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot $name) -Destination $tools
}
@'
param($InputPdf,$OutputDirectory,$Dpi,[switch]$SkipBuild,$Framework,$RasterizerPath)
if (Test-Path -LiteralPath (Join-Path $PSScriptRoot 'forbid-cache-worker')) {
    'raster' | Add-Content -LiteralPath (Join-Path $PSScriptRoot 'unexpected-cache-workers.log')
    throw 'Unexpected rasterization during a cache-only refusal.'
}
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
if ($RasterizerPath) {
    @{ Framework = $Framework; SkipBuild = [bool]$SkipBuild; Dpi = $Dpi; RasterizerPath = $RasterizerPath } |
        ConvertTo-Json -Compress | Add-Content -LiteralPath (Join-Path $PSScriptRoot 'frozen-raster-calls.jsonl')
}
[IO.File]::WriteAllText((Join-Path $OutputDirectory 'page-001.png'),'Synthetic fresh raster')
if (Test-Path -LiteralPath (Join-Path $PSScriptRoot 'fail-raster')) { throw 'Intentional interrupted raster' }
'@ | Set-Content -LiteralPath (Join-Path $tools 'RasterizePdf.ps1') -Encoding utf8
. (Join-Path $tools 'ReferenceCache.ps1')
$inputFile=Join-Path $root 'input.docx'
[IO.File]::WriteAllText($inputFile,'Synthetic input identity')
$key=Get-ReferenceIdentityKey $inputFile ''
$cache=Join-Path (Get-ReferenceCacheRoot) $key
New-Item -ItemType Directory -Path $cache -Force | Out-Null
[IO.File]::WriteAllText((Join-Path $cache 'reference.pdf'),'%PDF-Synthetic mechanics probe, not a fidelity reference')
$pdfHash=(Get-FileHash (Join-Path $cache 'reference.pdf')).Hash.ToLowerInvariant()
$null=Write-ReferenceIdentityMetadata $cache $key $inputFile ((Get-FileHash $inputFile).Hash.ToLowerInvariant()) '' 'mechanics-only' @{Producer='synthetic-test'} 'Never use for fidelity'
$checks=@()
$pdfOnly=(New-Item -ItemType Directory -Path (Join-Path $root 'pdf-only')).FullName
$sourceMetadata=Join-Path $cache 'reference-metadata.json'
$metadataHash=(Get-FileHash -LiteralPath $sourceMetadata).Hash
Copy-VerifiedCachedReferencePdf $inputFile $pdfOnly ''
if ((Get-FileHash -LiteralPath (Join-Path $pdfOnly 'reference.pdf')).Hash -ne $pdfHash -or
    (Get-FileHash -LiteralPath (Join-Path $pdfOnly 'reference-metadata.json')).Hash -ne $metadataHash -or
    (Get-FileHash -LiteralPath $sourceMetadata).Hash -ne $metadataHash -or
    (Test-Path -LiteralPath (Join-Path $cache 'raster')) -or @(Get-ChildItem -LiteralPath $pdfOnly -Force).Count -ne 2) {
    throw 'PDF-only cache hit changed the cache or served unexpected output.'
}
$checks+='pdf-only-verified-hit'

$sourcePdf=Join-Path $cache 'reference.pdf'
$originalPdf=[IO.File]::ReadAllBytes($sourcePdf)
$rejectedPdfOnly=(New-Item -ItemType Directory -Path (Join-Path $root 'pdf-only-corrupt')).FullName
$caught=$null
try {
    [IO.File]::WriteAllText($sourcePdf,'%PDF-Synthetic changed bytes')
    Copy-VerifiedCachedReferencePdf $inputFile $rejectedPdfOnly ''
} catch { $caught=$_.Exception.Message }
finally { [IO.File]::WriteAllBytes($sourcePdf,$originalPdf) }
if (!$caught -or !$caught.Contains('reference.pdf hash mismatch') -or @(Get-ChildItem -LiteralPath $rejectedPdfOnly -Force).Count) {
    throw 'PDF-only cache hit served a corrupt reference.'
}
$checks+='pdf-only-refuses-corrupt-pdf'

$name=Get-ReferenceDerivativeName 144
$derivative=Join-Path (Join-Path $cache 'raster') $name
New-Item -ItemType Directory -Path $derivative -Force | Out-Null
$page=Join-Path $derivative 'page-001.png'
[IO.File]::WriteAllText($page,'Old incomplete raster')
[IO.File]::WriteAllText((Join-Path $derivative 'page-999.png'),'Stale extra page')
function Assert-State([string] $Label,[string] $Expected) {
    $actual=Test-ReferenceDerivative $cache $name $pdfHash
    if ($actual.State -ne $Expected) { throw ($Label+': '+$actual.State+' / '+$actual.Reason) }
    $script:checks+=$Label
}
Assert-State 'unrecorded-refuses-reuse' 'Missing'
[IO.File]::WriteAllText((Join-Path $tools 'fail-raster'),'Fail after one page')
$caught=$null
try { & (Join-Path $tools 'RenderCachedReference.ps1') -InputPath $inputFile -OutputDirectory (Join-Path $root 'failed-output') -CacheOnly -SkipBuild *> (Join-Path $root 'failed.log') } catch { $caught=$_.Exception.Message }
if ($caught -ne 'Intentional interrupted raster' -or [IO.File]::ReadAllText($page) -ne 'Old incomplete raster') { throw 'Interrupted recovery changed the old derivative or hid failure' }
Assert-State 'interrupted-recovery-still-incomplete' 'Missing'
Remove-Item -LiteralPath (Join-Path $tools 'fail-raster')
& (Join-Path $tools 'RenderCachedReference.ps1') -InputPath $inputFile -OutputDirectory (Join-Path $root 'served') -CacheOnly -SkipBuild *> (Join-Path $root 'recovery.log')
Assert-State 'fresh-recovery-completes' 'Complete'
if ([IO.File]::ReadAllText((Join-Path $root 'served/page-001.png')) -ne 'Synthetic fresh raster' -or
    (Test-Path (Join-Path $root 'served/page-999.png')) -or (Get-FileHash (Join-Path $cache 'reference.pdf')).Hash -ne $pdfHash -or
    @(Get-ChildItem (Join-Path $cache 'raster') -Directory -Filter '*.incomplete-*').Count -ne 1) { throw 'Recovery mixed old pages, lost evidence or changed the reference PDF' }
$metadataPath=Join-Path $cache 'reference-metadata.json'
$goodMetadata=[IO.File]::ReadAllBytes($metadataPath)
foreach ($fault in @('Dpi','RasterizerId','PdfSha256','PageCount','RecordedName','Bytes','Name')) {
    $pin=Get-Content $metadataPath -Raw | ConvertFrom-Json
    switch ($fault) {
        'Dpi' {$pin.Derivatives.$name.Dpi=72}
        'RasterizerId' {$pin.Derivatives.$name.RasterizerId='a'*12}
        'PdfSha256' {$pin.Derivatives.$name.PdfSha256='0'*64}
        'PageCount' {$pin.Derivatives.$name.PageCount=2}
        'RecordedName' {$pin.Derivatives.$name.Pages[0].Name='page-999.png'}
        'Bytes' {[IO.File]::WriteAllText($page,'Changed bytes')}
        'Name' {Move-Item -LiteralPath $page -Destination (Join-Path $derivative 'page-999.png')}
    }
    $pin | ConvertTo-Json -Depth 8 | Set-Content $metadataPath -Encoding utf8
    Assert-State ('reject-'+$fault) 'Corrupt'
    if ($fault -eq 'Name') {Move-Item -LiteralPath (Join-Path $derivative 'page-999.png') -Destination $page}
    [IO.File]::WriteAllText($page,'Synthetic fresh raster')
    [IO.File]::WriteAllBytes($metadataPath,$goodMetadata)
}
Assert-State 'restored-control' 'Complete'

$frozenLog = Join-Path $tools 'frozen-raster-calls.jsonl'
foreach ($framework in @('net10.0', 'net11.0')) {
    $runtime = Join-Path $tools "Lokad.OoxPdf.PdfiumRasterizer/bin/Debug/$framework"
    New-Item -ItemType Directory -Path $runtime -Force | Out-Null
    $rasterizer = Join-Path $runtime 'Lokad.OoxPdf.PdfiumRasterizer.dll'
    [IO.File]::WriteAllText($rasterizer, "Frozen rasterizer $framework")
    $frozenId = Get-RasterizerId $rasterizer
    $frozenDerivative = Get-ReferenceDerivativeName 144 $frozenId
    if ($frozenDerivative -ceq $name -or (Test-ReferenceDerivative $cache $frozenDerivative $pdfHash).State -ne 'Missing') {
        throw 'An explicit rasterizer reused an unrelated derivative.'
    }
    $previousCalls = if (Test-Path -LiteralPath $frozenLog) { @(Get-Content -LiteralPath $frozenLog).Count } else { 0 }
    $cold = Join-Path $root "$framework-cold"
    $warm = Join-Path $root "$framework-warm"
    foreach ($output in @($cold, $warm)) {
        & (Join-Path $tools 'RenderCachedReference.ps1') -Framework $framework -InputPath $inputFile -OutputDirectory $output -CacheOnly -SkipBuild -RasterizerPath $rasterizer | Out-Null
    }
    $calls = @(Get-Content -LiteralPath $frozenLog | ForEach-Object { ConvertFrom-Json -InputObject $_ })
    if ($calls.Count -ne $previousCalls + 1 -or $calls[-1].Framework -cne $framework -or
        $calls[-1].RasterizerPath -cne $rasterizer -or !$calls[-1].SkipBuild -or $calls[-1].Dpi -ne 144 -or
        (Test-ReferenceDerivative $cache $frozenDerivative $pdfHash).State -ne 'Complete' -or
        (Get-FileHash -LiteralPath (Join-Path $cold 'page-001.png')).Hash -ne
            (Get-FileHash -LiteralPath (Join-Path $warm 'page-001.png')).Hash) {
        throw 'Frozen cold or warm cache lost the rasterizer binding.'
    }
    $checks += "frozen-$framework-cold-and-warm"
}
$goodMetadata=[IO.File]::ReadAllBytes($metadataPath)
$legacy=Get-Content $metadataPath -Raw | ConvertFrom-Json
$legacy.Derivatives.$name.RasterizerId='legacy-'+(Get-RasterizerId).Substring(0,8)
$legacy | ConvertTo-Json -Depth 8 | Set-Content $metadataPath -Encoding utf8
Assert-State 'recorded-legacy-provenance-remains-usable' 'Complete'
[IO.File]::WriteAllBytes($metadataPath,$goodMetadata)
$before=(Get-FileHash $metadataPath).Hash
& (Join-Path $tools 'RenderCachedReference.ps1') -InputPath $inputFile -OutputDirectory (Join-Path $root 'reused') -CacheOnly -SkipBuild *> (Join-Path $root 'reuse.log')
if ((Get-FileHash $metadataPath).Hash -ne $before -or @(Get-ChildItem (Join-Path $cache 'raster') -Directory -Filter '_raster-*').Count) {throw 'Reuse rerastered or leaked staging'}
foreach ($kind in @('stale-pages','unrelated-png')) {
    $output=Join-Path $root $kind
    New-Item -ItemType Directory -Path $output | Out-Null
    [IO.File]::WriteAllText((Join-Path $output 'reference.pdf'),'Retain previous served PDF')
    $image=Join-Path $output $(if ($kind -eq 'stale-pages') {'page-002.png'} else {'artwork.png'})
    [IO.File]::WriteAllText($image,'Retain caller PNG')
    $caught=$null
    try { & (Join-Path $tools 'RenderCachedReference.ps1') -InputPath $inputFile -OutputDirectory $output -CacheOnly -SkipBuild *> ($output+'.log') }
    catch { $caught=$_.Exception.Message }
    if (!$caught -or !$caught.Contains('choose a fresh directory') -or
        [IO.File]::ReadAllText($image) -ne 'Retain caller PNG' -or
        [IO.File]::ReadAllText((Join-Path $output 'reference.pdf')) -ne 'Retain previous served PDF' -or
        (Get-FileHash $metadataPath).Hash -ne $before) { throw ('Reference serving did not preserve/refuse existing output: '+$kind) }
    $checks+=('reject-served-'+$kind)
}
$fresh=Join-Path $root 'fresh-served-control'
New-Item -ItemType Directory -Path $fresh | Out-Null
[IO.File]::WriteAllText((Join-Path $fresh 'caller.txt'),'Retain caller text')
& (Join-Path $tools 'RenderCachedReference.ps1') -InputPath $inputFile -OutputDirectory $fresh -CacheOnly -SkipBuild *> ($fresh+'.log')
if (@(Get-ChildItem $fresh -Filter '*.png' -File).Count -ne 1 -or
    [IO.File]::ReadAllText((Join-Path $fresh 'caller.txt')) -ne 'Retain caller text') { throw 'Fresh served reference page set differs' }
$checks+='fresh-served-page-set'

$render=Join-Path $tools 'RenderCachedReference.ps1'
$incompleteInput=Join-Path $root 'incomplete.docx'
[IO.File]::WriteAllText($incompleteInput,'Unpublished reference input')
$incompleteDir=Join-Path (Get-ReferenceCacheRoot) (Get-ReferenceIdentityKey $incompleteInput '')
New-Item -ItemType Directory -Path $incompleteDir | Out-Null
Copy-Item -LiteralPath $sourcePdf -Destination (Join-Path $incompleteDir 'reference.pdf')
$miss=$null
try { & $render -InputPath $incompleteInput -OutputDirectory (Join-Path $root 'incomplete-hit') -CacheOnly -SkipBuild | Out-Null }
catch { $miss=$_.Exception.Message }
if (!$miss -or !$miss.Contains('Reference cache miss') -or !$miss.Contains('RenderCachedReference.ps1')) {
    throw 'An incomplete reference must produce an actionable cache-only miss'
}
$checks+='incomplete-entry-is-an-actionable-miss'

$secondDpi=Join-Path $root 'second-dpi'
& $render -InputPath $inputFile -OutputDirectory $secondDpi -Dpi 72 -CacheOnly -SkipBuild *> ($secondDpi+'.log')
if ((Test-ReferenceDerivative $cache (Get-ReferenceDerivativeName 72) $pdfHash).State -cne 'Complete' -or
    (Test-ReferenceDerivative $cache $name $pdfHash).State -cne 'Complete') { throw 'DPI derivatives did not coexist' }
$checks+='second-dpi-reuses-pdf'

New-Item -ItemType File -Path (Join-Path $tools 'forbid-cache-worker') | Out-Null
@'
'office' | Add-Content -LiteralPath (Join-Path $PSScriptRoot 'unexpected-cache-workers.log')
throw 'Unexpected Office worker during a cache refusal.'
'@ | Set-Content -LiteralPath (Join-Path $tools 'RenderReference.ps1') -Encoding utf8

function Assert-ReferenceRefusal($inputPath, $output, $message, [switch] $AllowRender) {
    $caught=$null
    try { & $render -InputPath $inputPath -OutputDirectory $output -CacheOnly:(!$AllowRender) -SkipBuild *> ($output+'.log') }
    catch { $caught=$_.Exception.Message }
    if (!$caught -or !$caught.Contains($message) -or
        (Test-Path -LiteralPath (Join-Path $output 'reference.pdf')) -or
        @(Get-ChildItem -LiteralPath $output -Filter 'page-*.png' -ErrorAction SilentlyContinue).Count -or
        (Test-Path -LiteralPath (Join-Path $tools 'unexpected-cache-workers.log'))) {
        throw "Reference refusal lost its bounds for '$output'. Received '$caught'."
    }
}

foreach ($kind in @('pre-v1-path-only-edited','pre-v1-hash-edited','pre-v1-hash-current','v1-flat-current')) {
    $legacyInput=Join-Path $root "$kind.docx"
    [IO.File]::WriteAllText($legacyInput,"Original $kind input")
    $originalHash=(Get-FileHash -LiteralPath $legacyInput).Hash.ToLowerInvariant()
    $keySuffix=if ($kind -eq 'v1-flat-current') {'v1'} else {'deadbeef'}
    $legacyDir=Join-Path (Get-ReferenceCacheRoot) "docx-$($originalHash.Substring(0,24))-$keySuffix-dpi144"
    New-Item -ItemType Directory -Path $legacyDir | Out-Null
    Copy-Item -LiteralPath $sourcePdf -Destination (Join-Path $legacyDir 'reference.pdf')
    Copy-Item -LiteralPath $page -Destination (Join-Path $legacyDir 'page-001.png')
    $marker="input=$legacyInput`ndpi=144`nvariant=`n"
    if ($kind.StartsWith('pre-v1-hash')) { $marker+="inputSha256=$originalHash`n" }
    [IO.File]::WriteAllText((Join-Path $legacyDir 'complete.txt'),$marker)
    if ($kind.EndsWith('-edited')) { [IO.File]::WriteAllText($legacyInput,"Edited $kind input") }
    $currentKey=Get-ReferenceIdentityKey $legacyInput ''
    Assert-ReferenceRefusal $legacyInput (Join-Path $root "refused-$kind") 'Reference cache miss'
    if ((Test-Path -LiteralPath (Join-Path (Get-ReferenceCacheRoot) $currentKey)) -or
        (Get-FileHash -LiteralPath (Join-Path $legacyDir 'reference.pdf')).Hash -cne (Get-FileHash -LiteralPath $sourcePdf).Hash -or
        (Get-FileHash -LiteralPath (Join-Path $legacyDir 'page-001.png')).Hash -cne (Get-FileHash -LiteralPath $page).Hash -or
        [IO.File]::ReadAllText((Join-Path $legacyDir 'complete.txt')) -cne $marker) {
        throw "Refusing old cache entry '$kind' changed retained reference files."
    }
    $checks+="ignore-$kind"
}

foreach ($kind in @('missing-metadata','legacy-adopted')) {
    $inputPath=Join-Path $root "$kind.docx"
    [IO.File]::WriteAllText($inputPath,"Historic $kind input")
    $inputHash=(Get-FileHash -LiteralPath $inputPath).Hash.ToLowerInvariant()
    $key=Get-ReferenceIdentityKey $inputPath ''
    $entryDir=Join-Path (Get-ReferenceCacheRoot) $key
    New-Item -ItemType Directory -Path $entryDir | Out-Null
    $entryPdf=Join-Path $entryDir 'reference.pdf'
    Copy-Item -LiteralPath $sourcePdf -Destination $entryPdf
    if ($kind -eq 'legacy-adopted') {
        $null=Write-ReferenceIdentityMetadata $entryDir $key $inputPath $inputHash '' 'old' @{Producer='legacy-adopted'} 'Synthetic old adoption'
    } else {
        [IO.File]::WriteAllText((Join-Path $entryDir 'complete.txt'),"input=$inputPath")
    }
    $entryHashes=@(Get-ChildItem -LiteralPath $entryDir -File | Get-FileHash | Sort-Object Path)
    Assert-ReferenceRefusal $inputPath (Join-Path $root "refused-$kind") 'trusted reference'
    Assert-ReferenceRefusal $inputPath (Join-Path $root "refused-render-$kind") 'trusted reference' -AllowRender
    $pdfOnlyOutput=(New-Item -ItemType Directory -Path (Join-Path $root "refused-pdf-only-$kind")).FullName
    $caught=$null
    try { Copy-VerifiedCachedReferencePdf $inputPath $pdfOnlyOutput '' }
    catch { $caught=$_.Exception.Message }
    if (!$caught -or !$caught.Contains('trusted reference') -or
        @(Get-ChildItem -LiteralPath $pdfOnlyOutput -Force).Count -or
        (Compare-Object $entryHashes @(Get-ChildItem -LiteralPath $entryDir -File | Get-FileHash | Sort-Object Path) -Property Hash,Path)) {
        throw "Refusing '$kind' changed files or served an unverified PDF."
    }
    $checks+="refuse-v2-$kind"
}
Remove-Item -LiteralPath (Join-Path $tools 'forbid-cache-worker')

$raceInput=Join-Path $root 'race.docx'
[IO.File]::WriteAllText($raceInput,'Concurrent reference input')
$import=Join-Path $tools 'ImportOfficeReferenceCache.ps1'
$sources=@(foreach ($publisher in @('first','second')) {
    $directory=(New-Item -ItemType Directory -Path (Join-Path $root $publisher)).FullName
    [IO.File]::WriteAllText((Join-Path $directory 'reference.pdf'),"%PDF-Synthetic $publisher publisher")
    [IO.File]::WriteAllText((Join-Path $directory 'page-001.png'),"Synthetic $publisher derivative")
    $directory
})
$published=@($sources | ForEach-Object -Parallel {
    $ErrorActionPreference='Stop'
    & $using:import -InputPath $using:raceInput -ReferenceDirectory $_ -Dpi 144 -CaseId (Split-Path -Leaf $_) -Force
} -ThrottleLimit 2 -TimeoutSeconds 20)
$winner=Resolve-ReferenceCacheEntry $raceInput 144 ''
if ($published.Count -ne 2 -or @($published | Where-Object { $_ -cne $winner.CacheDirectory }).Count -or
    $winner.Outcome -cne 'Hit' -or $winner.Derivative.State -cne 'Complete' -or
    $winner.Metadata.Producer -cne 'import' -or $winner.Metadata.CaseId -cnotin @('first','second') -or
    @(Get-ChildItem -LiteralPath (Get-ReferenceCacheRoot) -Directory -Filter '_import-*').Count) { throw 'Concurrent imports did not converge on one complete entry' }
$source=Join-Path $root $winner.Metadata.CaseId
$served=Join-Path $root 'race-served'
& $render -InputPath $raceInput -OutputDirectory $served -CacheOnly -SkipBuild *> ($served+'.log')
foreach ($file in @('reference.pdf','page-001.png')) {
    if ((Get-FileHash -LiteralPath (Join-Path $served $file)).Hash -cne (Get-FileHash -LiteralPath (Join-Path $source $file)).Hash) {
        throw "Concurrent cache provenance selected a mixed pair for $file"
    }
}
$checks+='concurrent-import-publishes-one-pair'

# A cold render and a competing request share one publication. The worker is
# synthetic and writes visibly different bytes on each invocation.
@'
param($InputPath,$OutputDirectory,$TimeoutSeconds,$Framework,$Dpi,[switch]$SkipBuild,$RasterizerPath)
$call=[guid]::NewGuid().ToString('N')
[IO.File]::WriteAllText((Join-Path $PSScriptRoot ('render-call-'+$call)), $InputPath)
Start-Sleep -Milliseconds 250
[IO.File]::WriteAllText((Join-Path $OutputDirectory 'reference.pdf'), '%PDF-Synthetic '+$call)
if (!(Test-Path -LiteralPath ($InputPath+'.omit-raster'))) {
    [IO.File]::WriteAllText((Join-Path $OutputDirectory 'page-001.png'), 'Synthetic raster '+$call)
}
[IO.File]::WriteAllText((Join-Path $OutputDirectory 'reference-worker-stdout.log'), 'Synthetic stdout')
[IO.File]::WriteAllText((Join-Path $OutputDirectory 'reference-worker-stderr.log'), 'Synthetic stderr')
[IO.File]::WriteAllText((Join-Path $OutputDirectory 'reference-worker-status.json'), '{"OfficeVersion":"synthetic-only"}')
Write-Output 'Synthetic worker output must stay outside the cache resolution record'
'@ | Set-Content -LiteralPath (Join-Path $tools 'RenderReference.ps1') -Encoding utf8
$coldInput=Join-Path $root 'cold-race.docx'
[IO.File]::WriteAllText($coldInput, 'Unique synthetic input for concurrent cache misses')
$coldOutputs=@('cold-first','cold-second' | ForEach-Object { Join-Path $root $_ })
$servedCold=@($coldOutputs | ForEach-Object -Parallel {
    $ErrorActionPreference='Stop'
    & $using:render -InputPath $using:coldInput -OutputDirectory $_ -SkipBuild *> ($_+'.log')
    $_
} -ThrottleLimit 2 -TimeoutSeconds 20)
$cold=Resolve-ReferenceCacheEntry $coldInput 144 ''
if ($servedCold.Count -ne 2 -or @(Get-ChildItem -LiteralPath $tools -Filter 'render-call-*' -File).Count -ne 1 -or
    $cold.Outcome -cne 'Hit' -or $cold.Derivative.State -cne 'Complete' -or $cold.Metadata.OfficeVersion -cne 'synthetic-only') {
    throw 'Concurrent cold requests did not share one complete render'
}
foreach ($output in $coldOutputs) {
    if ((Get-FileHash -LiteralPath (Join-Path $output 'reference.pdf')).Hash -ne $cold.Metadata.ReferencePdfSha256 -or
        (Get-FileHash -LiteralPath (Join-Path $output 'page-001.png')).Hash -ne
            (Get-FileHash -LiteralPath (Join-Path $cold.Derivative.Directory 'page-001.png')).Hash) {
        throw 'A cold-cache reader received a different PDF or raster'
    }
}
if (@(Get-ChildItem -LiteralPath $cold.CacheDirectory -Filter 'reference-worker-*' -File).Count -or
    @(Get-ChildItem -LiteralPath (Get-ReferenceCacheRoot) -Filter '_render-*' -Directory).Count) {
    throw 'A successful cache fill retained worker scratch'
}
$checks+=@('concurrent-cold-render-publishes-once','successful-render-excludes-worker-scratch')

# Keep the old entry until the replacement has both its PDF and derivative.
$coldPdf=Join-Path $cold.CacheDirectory 'reference.pdf'
[IO.File]::WriteAllText($coldPdf, '%PDF-Intentional reference hash mismatch')
$corruptHash=(Get-FileHash -LiteralPath $coldPdf).Hash
[IO.File]::WriteAllText(($coldInput+'.omit-raster'), 'Synthetic incomplete worker result')
$caught=$null
try {
    & $render -InputPath $coldInput -OutputDirectory (Join-Path $root 'cold-incomplete') -SkipBuild *> (Join-Path $root 'cold-incomplete.log')
}
catch { $caught=$_.Exception.Message }
if ($caught -notlike '*no page-*.png raster output*' -or (Get-FileHash -LiteralPath $coldPdf).Hash -ne $corruptHash) {
    throw 'An incomplete render displaced the previous cache entry'
}
Remove-Item -LiteralPath ($coldInput+'.omit-raster')
& $render -InputPath $coldInput -OutputDirectory (Join-Path $root 'cold-repaired') -SkipBuild *> (Join-Path $root 'cold-repaired.log')
$repaired=Resolve-ReferenceCacheEntry $coldInput 144 ''
if ($repaired.Outcome -cne 'Hit' -or $repaired.Derivative.State -cne 'Complete' -or
    $repaired.Metadata.ReferencePdfSha256 -eq $corruptHash) { throw 'A complete render did not replace the corrupt entry' }
$checks+='corrupt-entry-replaced-only-after-complete-render'

@{Checks=$checks;Builds=0;OfficeExports=0;SyntheticOnly=$true;ReferencePdfPreserved=$true} | ConvertTo-Json -Depth 4 | Set-Content (Join-Path $root 'evidence.json') -Encoding utf8
"Reference derivative: $($checks.Count) checks pass; no build or Office export."
