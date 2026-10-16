# Verify Word view selection and cache provenance without Office or real PDFs.
[CmdletBinding()]
param([string] $OutputDirectory,
    [string] $ExportScript = (Join-Path $PSScriptRoot 'ExportCompositionReferences.ps1'))
$ErrorActionPreference='Stop'
if (!$OutputDirectory) {
    $OutputDirectory=Join-Path (Split-Path $PSScriptRoot) ('artifacts/reference-test-'+[Guid]::NewGuid().ToString('N'))
}
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Use a fresh test directory' }
$repo=Join-Path $OutputDirectory 'repository'
$scripts=Join-Path $repo 'tools'
New-Item -ItemType Directory -Path $scripts -Force | Out-Null
$script=Join-Path $scripts 'ExportCompositionReferences.ps1'
Copy-Item -LiteralPath $ExportScript -Destination $script
@'
param($InputDocx,$OutputPdf,[bool]$ShowRevisions,$RevisionsView,$MarkupMode,$RevisionsMode,[bool]$ShowComments,[switch]$RejectAll)
Add-Content -LiteralPath (Join-Path $PSScriptRoot 'calls.txt') $InputDocx
if (Test-Path -LiteralPath (Join-Path $PSScriptRoot 'wrong-view.txt')) { $RevisionsMode=0 }
[IO.File]::WriteAllText($OutputPdf,'Fake Word PDF')
Write-Output "ExportSettings: Word/16.0 ShowRevisions=$ShowRevisions RevisionsView=$RevisionsView MarkupMode=$MarkupMode RevisionsMode=$RevisionsMode ShowComments=$ShowComments RejectAll=$([bool]$RejectAll) (inherited MarkupMode=0 RevisionsMode=0)"
'@ | Set-Content -LiteralPath (Join-Path $scripts 'ExportDocxMarkupReference.ps1')
$documents=@(foreach ($id in @('first','second','third')) {
    $inputPath=Join-Path $repo ($id+'.docx')
    'Fake DOCX input' | Set-Content -LiteralPath $inputPath
    @{Id=$id;Input=($id+'.docx');Sha256=(Get-FileHash $inputPath).Hash}
})
$manifest=Join-Path $repo 'manifest.json'
ConvertTo-Json -InputObject @($documents[0],$documents[1]) | Set-Content -LiteralPath $manifest
$reference=Join-Path $OutputDirectory 'reference'
& $script -Manifest $manifest -OutputDirectory $reference | Out-Null
$calls=Join-Path $scripts 'calls.txt'
if (@(Get-Content $calls).Count -ne 4) { throw 'Expected one export per source/mode' }
foreach ($mode in @('final-preserve','all-word-compatible')) {
    $pin=Get-Content (Join-Path $reference ('first-'+$mode+'/source-identity.json')) -Raw | ConvertFrom-Json
    if ($pin.OfficeView.RevisionsMode -ne $(if ($mode -eq 'all-word-compatible') {1} else {0}) -or
        !$pin.ExportLogSha256 -or !$pin.ExportSettings) { throw 'View and provenance were not pinned' }
}
$passed=@('explicit-views')
& $script -Manifest $manifest -OutputDirectory $reference | Out-Null
if (@(Get-Content $calls).Count -ne 4) { throw 'Cache reuse exported again' }
$passed+='verified-reuse'
# Legacy metadata is reusable only when its retained log proves the exact view.
$legacy=Join-Path $reference 'first-all-word-compatible/source-identity.json'
$pin=Get-Content $legacy -Raw | ConvertFrom-Json -AsHashtable
$pin.Remove('OfficeView'); $pin.Remove('ExportSettings'); $pin.Remove('ExportLogSha256')
$pin | ConvertTo-Json | Set-Content -LiteralPath $legacy
& $script -Manifest $manifest -OutputDirectory $reference | Out-Null
if (@(Get-Content $calls).Count -ne 4) { throw 'Verified legacy cache exported again' }
$passed+='legacy-log-proof'
foreach ($fault in @('cached-view','legacy-view','unknown-view','pdf-hash')) {
    $directory=Join-Path $OutputDirectory $fault
    Copy-Item -LiteralPath $reference -Destination $directory -Recurse
    $folder=Join-Path $directory 'first-all-word-compatible'
    $identity=Join-Path $folder 'source-identity.json'
    $log=Join-Path $folder 'export.log'
    $pin=Get-Content $identity -Raw | ConvertFrom-Json -AsHashtable
    switch ($fault) {
        'cached-view' {
            $pin.OfficeView=@{RevisionsMode=0}
            $pin | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $identity
        }
        'legacy-view' { (Get-Content $log -Raw).Replace('RevisionsMode=1','RevisionsMode=0') | Set-Content -LiteralPath $log }
        'unknown-view' { 'No observed export settings' | Set-Content -LiteralPath $log }
        'pdf-hash' { 'Changed PDF' | Set-Content -LiteralPath (Join-Path $folder 'reference.pdf') }
    }
    # Queue a fresh source first: a bad later cache must reject before any export.
    ConvertTo-Json -InputObject @($documents[2],$documents[0]) | Set-Content -LiteralPath $manifest
    $before=@(Get-Content $calls).Count
    $rejected=$false
    try { & $script -Manifest $manifest -OutputDirectory $directory | Out-Null }
    catch { if ($_.Exception.Message -notmatch 'reference.*(differs|unverified)') { throw }; $rejected=$true }
    if (!$rejected -or @(Get-Content $calls).Count -ne $before -or
        (Test-Path -LiteralPath (Join-Path $directory 'third-final-preserve'))) { throw ('Preflight failed: '+$fault) }
    $passed+=$fault
}
'Use wrong view' | Set-Content -LiteralPath (Join-Path $scripts 'wrong-view.txt')
ConvertTo-Json -InputObject @($documents[0]) | Set-Content -LiteralPath $manifest
$wrong=Join-Path $OutputDirectory 'wrong-observed-view'
$rejected=$false
try { & $script -Manifest $manifest -OutputDirectory $wrong -DocxPreset Review | Out-Null }
catch { if ($_.Exception.Message -notmatch 'Reference Office view differs') { throw }; $rejected=$true }
if (!$rejected -or (Test-Path -LiteralPath (Join-Path $wrong 'first-all-word-compatible/source-identity.json'))) {
    throw 'An exporter that reports the wrong view must not publish a reference identity'
}
$passed+='observed-export-view'
@{Passed=$passed;OfficeInvocations=0;LibraryBuilds=0;Scope='Fake DOCX/PDF inputs; view and provenance orchestration only'} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'evidence.json')
Write-Output ($passed.Count.ToString()+' reference view/provenance checks passed; no Office or library build.')
