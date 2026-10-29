# Image-reference cache/view checks; no Office, drawing/native calls or builds.
[CmdletBinding()]
param([string] $OutputDirectory)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'OfficeReferenceCapture.ps1')
if (!$OutputDirectory) { $OutputDirectory=Join-Path (Split-Path $PSScriptRoot) ('artifacts/office-image-test-'+[guid]::NewGuid().ToString('N')) }
if (Test-Path $OutputDirectory) { throw 'Use a fresh test directory' }
$root=(New-Item -ItemType Directory -Path $OutputDirectory).FullName
$hash='A'*64
$png=[Convert]::FromBase64String('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+a9WQAAAAASUVORK5CYII=')
$checks=@()
function New-Control([string] $Name) {
    $directory=(New-Item -ItemType Directory -Path (Join-Path $root $Name)).FullName
    [IO.File]::WriteAllBytes((Join-Path $directory 'page-001.png'),$png)
    [IO.File]::WriteAllText((Join-Path $directory 'reference.pdf'),"%PDF-1.4`n%%EOF")
    $metadata=@{Schema=1;ReferenceKind='OfficeImage';InputSha256=$hash;OfficeBuild='managed control';
        SecondaryPdfSha256=(Get-FileHash (Join-Path $directory 'reference.pdf')).Hash;
        Capture=@{Method='managed test';Pages=@(@{Page=1;File='page-001.png';Width=1;Height=1;
            WidthPoints=0.5;HeightPoints=0.5;Sha256=(Get-FileHash (Join-Path $directory 'page-001.png')).Hash})}}
    $metadata | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $directory 'reference-image-metadata.json')
    return $directory
}
function Save-Metadata([string] $Directory, $Metadata) {
    $Metadata | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $Directory 'reference-image-metadata.json')
}
$valid=New-Control 'valid'
$null=Test-OfficeImageReference $valid $hash
$checks+='valid reference'
foreach ($fault in @('page-bytes','dimensions','page-index','extra-page','secondary-pdf','input','kind','emf-path')) {
    $directory=New-Control $fault
    $metadata=Get-Content (Join-Path $directory 'reference-image-metadata.json') -Raw | ConvertFrom-Json -AsHashtable
    switch ($fault) {
        'page-bytes' { [IO.File]::AppendAllText((Join-Path $directory 'page-001.png'),'changed') }
        'dimensions' {$metadata.Capture.Pages[0].Width=2}
        'page-index' {$metadata.Capture.Pages[0].Page=2}
        'extra-page' {[IO.File]::WriteAllBytes((Join-Path $directory 'page-002.png'),$png)}
        'secondary-pdf' {[IO.File]::AppendAllText((Join-Path $directory 'reference.pdf'),'changed')}
        'input' {$metadata.InputSha256='B'*64}
        'kind' {$metadata.ReferenceKind='OfficePdf'}
        'emf-path' {$metadata.Capture.Pages[0].Emf='../unowned.emf'}
    }
    Save-Metadata $directory $metadata
    $rejected=$false
    try {$null=Test-OfficeImageReference $directory $hash} catch {$rejected=$true}
    if (!$rejected) {throw ('Image reference accepted '+$fault)}
    $checks+=$fault
}
function Release-ComObject($Value) {} # Managed view stubs have no COM lifetime.
foreach ($review in @($false,$true)) {
    $view=[pscustomobject]@{Type=1;MarkupMode=1;RevisionsMode=1;RevisionsView=1;
        _ShowComments=$false;ShowFormatChanges=$false;ShowInsertionsAndDeletions=$false}
    $view | Add-Member ScriptProperty ShowComments {$this._ShowComments} {
        $this._ShowComments=$args[0];$this.MarkupMode=0
    }
    $document=[pscustomobject]@{ShowRevisions=$false;ActiveWindow=[pscustomobject]@{View=$view}}
    $settings=@{ShowRevisions=$review;RevisionsView=0;MarkupMode=$(if($review){2}else{0});
        RevisionsMode=$(if($review){2}else{0});ShowComments=$review;
        ShowFormatChanges=$review;ShowInsertionsAndDeletions=$review;RejectAll=$false}
    $result=Set-OfficeWordReferenceView $document $settings
    if (!$result.PrintLayout -or $view.Type -ne 3 -or $view.MarkupMode -ne $settings.MarkupMode -or
        $result.Inherited.MarkupMode -ne 1) {throw 'Explicit Word view/order was not preserved'}
    $checks+='Word view '+$review
}
@{Checks=$checks;OfficeExports=0;Builds=0;NativeCalls=0} | ConvertTo-Json | Set-Content (Join-Path $root 'evidence.json')
Write-Output "$($checks.Count) image provenance and Word view checks pass; no Office, builds or native calls."
