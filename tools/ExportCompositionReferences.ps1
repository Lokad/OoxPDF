# Export each pinned source/mode once; reuse existing independent references.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)] [string] $Manifest,
    [Parameter(Mandatory=$true)] [string] $OutputDirectory,
    [ValidateSet('Both','Final','Review')] [string] $DocxPreset = 'Both',
    [ValidateSet('OfficePdf','OfficeImage')] [string] $ReferenceKind = 'OfficePdf',
    [int] $Dpi = 144,
    [string] $FontIdentity
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'CompositionReferenceIdentity.ps1')
. (Join-Path $PSScriptRoot 'CompositionOutcomes.ps1')
if ($ReferenceKind -eq 'OfficeImage') { . (Join-Path $PSScriptRoot 'OfficeReferenceCapture.ps1') }
$repo = Split-Path -Parent $PSScriptRoot
$jobs = @()
# Preflight every cached reference before starting Office or creating outputs.
foreach ($document in @(Get-Content -LiteralPath $Manifest -Raw | ConvertFrom-Json)) {
    $inputPath = (Resolve-Path -LiteralPath (Join-Path $repo $document.Input)).Path
    if ((Get-FileHash $inputPath).Hash -ne $document.Sha256) { throw ('Reference input differs: ' + $document.Id) }
    foreach ($mode in (Get-PortfolioModes $document $DocxPreset)) {
        $docx = $mode -cne 'presentation'
        $folder = Join-Path $OutputDirectory ($document.Id+'-'+$mode)
        $identity = Join-Path $folder 'source-identity.json'
        $pdf = Join-Path $folder 'reference.pdf'
        $log = Join-Path $folder 'export.log'
        if (Test-Path -LiteralPath $identity) {
            $reference = Read-CompositionReference $folder $document.Id $mode $document.Sha256 $ReferenceKind
            if ($ReferenceKind -eq 'OfficeImage') {
                $capture = Test-OfficeImageReference $folder $document.Sha256 -ObservedPdfSha256 $reference.Hash
                if ((Get-FileHash (Join-Path $folder 'reference-image-metadata.json')).Hash -ne $reference.Pin.ImageMetadataSha256 -or
                    $capture.Dpi -ne $Dpi -or $capture.FontIdentity -ne $FontIdentity) { throw 'Cached image capture settings differ' }
            }
            continue
        }
        if (Test-Path -LiteralPath $folder) { throw 'Unpinned reference directory exists; inspect it before reuse' }
        if (@($jobs | Where-Object Folder -EQ $folder).Count) { throw 'Duplicate reference outcome' }
        $jobs += [pscustomobject]@{Document=$document;Input=$inputPath;Docx=$docx;Mode=$mode;
            Folder=$folder;Pdf=$pdf;Identity=$identity;Log=$log;View=(Get-CompositionWordView $mode)}
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
            }
        }
        elseif ($job.Docx) {
            $view = $job.View
            & (Join-Path $PSScriptRoot 'ExportDocxMarkupReference.ps1') -InputDocx $inputPath -OutputPdf $pdf @view *> $job.Log
        }
        else {
            # Publish the PDF reference once. Comparisons own the verified raster
            # derivative and bind their explicit frozen validation tools.
            & (Join-Path $PSScriptRoot 'RenderReference.ps1') -InputPath $inputPath -OutputDirectory $folder -SkipRaster *> $job.Log
        }
        $pin = [ordered]@{Id=$document.Id;Mode=$mode;InputSha256=$document.Sha256;PdfSha256=(Get-FileHash $pdf).Hash;ReferenceKind=$ReferenceKind}
        if ($ReferenceKind -eq 'OfficeImage') {
            $capture=Test-OfficeImageReference $folder $document.Sha256
            $pin.ImageMetadataSha256=(Get-FileHash (Join-Path $folder 'reference-image-metadata.json')).Hash
            $pin.CalibrationStatus=$capture.CalibrationStatus
        }
        if ($job.Docx) {
            $export = Read-CompositionWordExport $job.Log $job.View ($document.Id + '-' + $mode)
            $pin.OfficeView = $job.View
            $pin.ExportSettings = $export.Settings
            $pin.ExportLogSha256 = $export.Sha256
        }
        $pin | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $job.Identity -Encoding utf8
        Write-Output ($document.Id+' '+$mode+' independent reference retained')
}
