# Export each pinned source/mode once; reuse existing independent references.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)] [string] $Manifest,
    [Parameter(Mandatory=$true)] [string] $OutputDirectory,
    [ValidateSet('Both','Final','Review')] [string] $DocxPreset = 'Both',
    [switch] $SkipBuild,
    [ValidateSet('OfficePdf','OfficeImage')] [string] $ReferenceKind = 'OfficePdf',
    [int] $Dpi = 144,
    [string] $FontIdentity
)
$ErrorActionPreference = 'Stop'
if ($ReferenceKind -eq 'OfficeImage') { . (Join-Path $PSScriptRoot 'OfficeReferenceCapture.ps1') }
$repo = Split-Path -Parent $PSScriptRoot
$jobs = @()
function Assert-WordView([string] $Log, $Expected, [string] $Case) {
    if (!(Test-Path -LiteralPath $Log)) { throw ('Reference Office view is unverified: ' + $Case) }
    $line = @(Get-Content -LiteralPath $Log | Where-Object { $_ -match '^ExportSettings: Word/' })
    if ($line.Count -ne 1) { throw ('Reference Office view is unverified: ' + $Case) }
    foreach ($key in $Expected.Keys) {
        $value = [regex]::Match($line[0], '(?:^| )' + $key + '=([^ ]+)')
        if (!$value.Success -or $value.Groups[1].Value -cne [string]$Expected[$key]) {
            throw ('Reference Office view differs: ' + $Case + ' / ' + $key)
        }
    }
    return $line[0]
}
# Preflight every cached reference before starting Office or creating outputs.
foreach ($document in @(Get-Content -LiteralPath $Manifest -Raw | ConvertFrom-Json)) {
    $inputPath = (Resolve-Path -LiteralPath (Join-Path $repo $document.Input)).Path
    if ((Get-FileHash $inputPath).Hash -ne $document.Sha256) { throw ('Reference input differs: ' + $document.Id) }
    $docx = $inputPath.EndsWith('.docx',[StringComparison]::OrdinalIgnoreCase)
    $modes = if ($docx) { @('final-preserve','all-word-compatible') } else { @('presentation') }
    if ($docx -and $DocxPreset -ne 'Both') { $modes = if ($DocxPreset -eq 'Final') { @('final-preserve') } else { @('all-word-compatible') } }
    foreach ($mode in $modes) {
        $folder = Join-Path $OutputDirectory ($document.Id+'-'+$mode)
        $identity = Join-Path $folder 'source-identity.json'
        $pdf = Join-Path $folder 'reference.pdf'
        $review = $mode -eq 'all-word-compatible'
        $view = [ordered]@{ShowRevisions=$review;RevisionsView=0;MarkupMode=$(if ($review) {2} else {0});
            RevisionsMode=$(if ($review) {2} else {0});ShowComments=$review;
            ShowFormatChanges=$review;ShowInsertionsAndDeletions=$review;RejectAll=$false}
        $log = Join-Path $folder 'export.log'
        if (Test-Path -LiteralPath $identity) {
            $pin = Get-Content -LiteralPath $identity -Raw | ConvertFrom-Json -AsHashtable
            $cachedKind=if ($pin.ReferenceKind) {$pin.ReferenceKind} else {'OfficePdf'}
            if ($cachedKind -ne $ReferenceKind) { throw 'Cached reference kind differs; use a separate output directory' }
            if ($pin.InputSha256 -ne $document.Sha256 -or $pin.Mode -ne $mode -or
                (Get-FileHash $pdf).Hash -ne $pin.PdfSha256) { throw 'Cached reference differs' }
            if ($ReferenceKind -eq 'OfficeImage') {
                $capture=Test-OfficeImageReference $folder $document.Sha256
                if ((Get-FileHash (Join-Path $folder 'reference-image-metadata.json')).Hash -ne $pin.ImageMetadataSha256 -or
                    $capture.Dpi -ne $Dpi -or $capture.FontIdentity -ne $FontIdentity) { throw 'Cached image capture settings differ' }
            }
            if ($docx) {
                $settings = Assert-WordView $log $view ($document.Id + '/' + $mode)
                if ($pin.ContainsKey('OfficeView')) {
                    foreach ($key in $view.Keys) {
                        if (!$pin.OfficeView.ContainsKey($key) -or [string]$pin.OfficeView[$key] -cne [string]$view[$key]) {
                            throw ('Cached reference Office view differs: ' + $document.Id + ' / ' + $key)
                        }
                    }
                    if ($pin.ExportSettings -cne $settings -or $pin.ExportLogSha256 -ne (Get-FileHash $log).Hash) {
                        throw ('Cached reference export provenance differs: ' + $document.Id)
                    }
                }
            }
            continue
        }
        if (Test-Path -LiteralPath $folder) { throw 'Unpinned reference directory exists; inspect it before reuse' }
        if (@($jobs | Where-Object Folder -EQ $folder).Count) { throw 'Duplicate reference outcome' }
        $jobs += [pscustomobject]@{Document=$document;Input=$inputPath;Docx=$docx;Mode=$mode;
            Folder=$folder;Pdf=$pdf;Identity=$identity;Log=$log;View=$view}
    }
}
foreach ($job in $jobs) {
        $folder=$job.Folder; $pdf=$job.Pdf; $inputPath=$job.Input; $document=$job.Document; $mode=$job.Mode
        New-Item -ItemType Directory -Force -Path $folder | Out-Null
        if ($ReferenceKind -eq 'OfficeImage') {
            $arguments=@{InputPath=$inputPath;OutputDirectory=$folder;ReferenceKind=$ReferenceKind;Dpi=$Dpi;FontIdentity=$FontIdentity}
            if ($job.Docx) {$arguments.WordView=$job.View}
            & (Join-Path $PSScriptRoot 'RenderReference.ps1') @arguments *> $job.Log
            if ($job.Docx) {
                $worker=Get-Content (Join-Path $folder 'reference-worker-status.json') -Raw | ConvertFrom-Json
                Write-Output ('ExportSettings: '+$worker.ExportSettings) | Add-Content $job.Log
                $settings=Assert-WordView $job.Log $job.View ($document.Id+'/'+$mode)
            }
        }
        elseif ($job.Docx) {
            $view = $job.View
            & (Join-Path $PSScriptRoot 'ExportDocxMarkupReference.ps1') -InputDocx $inputPath -OutputPdf $pdf @view *> $job.Log
            $settings = Assert-WordView $job.Log $view ($document.Id + '/' + $mode)
        }
        else {
            # Publish the PDF reference once. Comparisons own the verified raster
            # derivative and bind their explicit frozen validation tools.
            & (Join-Path $PSScriptRoot 'RenderReference.ps1') -InputPath $inputPath -OutputDirectory $folder -SkipBuild:$SkipBuild -SkipRaster *> (Join-Path $folder 'export.log')
        }
        $pin = [ordered]@{Id=$document.Id;Mode=$mode;InputSha256=$document.Sha256;PdfSha256=(Get-FileHash $pdf).Hash;ReferenceKind=$ReferenceKind}
        if ($ReferenceKind -eq 'OfficeImage') {
            $capture=Test-OfficeImageReference $folder $document.Sha256
            $pin.ImageMetadataSha256=(Get-FileHash (Join-Path $folder 'reference-image-metadata.json')).Hash
            $pin.CalibrationStatus=$capture.CalibrationStatus
        }
        if ($job.Docx) {
            $pin.OfficeView=$job.View; $pin.ExportSettings=$settings; $pin.ExportLogSha256=(Get-FileHash $job.Log).Hash
        }
        $pin | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $job.Identity -Encoding utf8
        Write-Output ($document.Id+' '+$mode+' independent reference retained')
}
