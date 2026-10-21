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
foreach ($name in @('ReferenceCache.ps1','RenderCachedReference.ps1')) { Copy-Item -LiteralPath (Join-Path $PSScriptRoot $name) -Destination $tools }
@'
param($InputPdf,$OutputDirectory,$Dpi,[switch]$SkipBuild)
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
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
$name=Get-ReferenceDerivativeName 144
$derivative=Join-Path (Join-Path $cache 'raster') $name
New-Item -ItemType Directory -Path $derivative -Force | Out-Null
$page=Join-Path $derivative 'page-001.png'
[IO.File]::WriteAllText($page,'Old incomplete raster')
[IO.File]::WriteAllText((Join-Path $derivative 'page-999.png'),'Stale extra page')
$checks=@()
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
$legacy=Get-Content $metadataPath -Raw | ConvertFrom-Json
$legacy.Derivatives.$name.RasterizerId='legacy-'+(Get-RasterizerId).Substring(0,8)
$legacy | ConvertTo-Json -Depth 8 | Set-Content $metadataPath -Encoding utf8
Assert-State 'recorded-legacy-provenance-remains-usable' 'Complete'
[IO.File]::WriteAllBytes($metadataPath,$goodMetadata)
$before=(Get-FileHash $metadataPath).Hash
& (Join-Path $tools 'RenderCachedReference.ps1') -InputPath $inputFile -OutputDirectory (Join-Path $root 'reused') -CacheOnly -SkipBuild *> (Join-Path $root 'reuse.log')
if ((Get-FileHash $metadataPath).Hash -ne $before -or @(Get-ChildItem (Join-Path $cache 'raster') -Directory -Filter '_raster-*').Count) {throw 'Reuse rerastered or leaked staging'}
@{Checks=$checks;Builds=0;OfficeExports=0;SyntheticOnly=$true;ReferencePdfPreserved=$true} | ConvertTo-Json -Depth 4 | Set-Content (Join-Path $root 'evidence.json') -Encoding utf8
"Reference derivative: $($checks.Count) checks pass; no build or Office export."
