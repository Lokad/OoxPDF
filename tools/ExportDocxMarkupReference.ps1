# Exports a DOCX file to PDF through Word COM with every markup-view knob set explicitly.
# Word persists View.MarkupMode and View.RevisionsMode across sessions, so exports
# that leave knobs untouched inherit whatever a previous session set. This script always
# sets and verifies revision visibility, revision/markup modes, comments, formatting
# and insertions/deletions, and reports the inherited modes. References vary by
# Word-stamped timestamps (CreationDate/ModDate differ on every export; compare references
# structurally, never by hash).
# Record the ExportSettings string with tools/ImportDocxMarkupReferenceCache.ps1.
param(
    [Parameter(Mandatory = $true)]
    [string] $InputDocx,
    [Parameter(Mandatory = $true)]
    [string] $OutputPdf,
    [bool] $ShowRevisions = $true,
    [ValidateRange(0, 1)]
    [int] $RevisionsView = 0,
    [ValidateRange(0, 2)]
    [int] $MarkupMode = 2,
    [ValidateRange(0, 2)]
    [int] $RevisionsMode = 0,
    [bool] $ShowComments = $true,
    [bool] $ShowFormatChanges = $ShowRevisions,
    [bool] $ShowInsertionsAndDeletions = $ShowRevisions,
    [switch] $RejectAll
)
$ErrorActionPreference = "Stop"
$inputFull = (Resolve-Path -LiteralPath $InputDocx).Path
$outputFull = [IO.Path]::GetFullPath($OutputPdf)
if ($inputFull.Equals($outputFull,[StringComparison]::OrdinalIgnoreCase)) { throw 'Reference output must differ from its input' }
$outputDir = [IO.Path]::GetDirectoryName($outputFull)
[void][IO.Directory]::CreateDirectory($outputDir)
# Retain the supervised attempt and provenance beside the requested PDF. The
# worker opens a private copy and owns all COM cleanup; this wrapper never does.
$run = Join-Path $outputDir ('word-reference-'+[guid]::NewGuid().ToString('N'))
$view = @{ShowRevisions=$ShowRevisions;RevisionsView=$RevisionsView;MarkupMode=$MarkupMode;
    RevisionsMode=$RevisionsMode;ShowComments=$ShowComments;ShowFormatChanges=$ShowFormatChanges;
    ShowInsertionsAndDeletions=$ShowInsertionsAndDeletions;RejectAll=[bool]$RejectAll}
& (Join-Path $PSScriptRoot 'RenderReference.ps1') -InputPath $inputFull -OutputDirectory $run -WordView $view -SkipRaster
$status=Get-Content (Join-Path $run 'reference-worker-status.json') -Raw | ConvertFrom-Json
if (!$status.OfficeView.PrintLayout) { throw 'Supervised Word view is unverified' }
$staging=$outputFull+'.staging-'+[guid]::NewGuid().ToString('N')
try {
    [IO.File]::Copy((Join-Path $run 'reference.pdf'),$staging,$false)
    [IO.File]::Move($staging,$outputFull,$true)
}
finally { if ([IO.File]::Exists($staging)) { [IO.File]::Delete($staging) } }
Write-Output ('Word version: '+$status.OfficeVersion)
Write-Output ('Exported: '+$outputFull+' ('+(Get-Item $outputFull).Length+' bytes, sha256 '+(Get-FileHash $outputFull).Hash+')')
Write-Output ('ExportSettings: '+$status.ExportSettings)
