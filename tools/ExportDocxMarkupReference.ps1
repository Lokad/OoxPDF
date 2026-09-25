# Exports a DOCX file to PDF through Word COM with every markup-view knob set explicitly.
# Word persists View.MarkupMode and View.RevisionsMode across sessions, so exports
# that leave knobs untouched inherit whatever a previous session set. This script always
# sets ShowRevisions, RevisionsView, MarkupMode, RevisionsMode and ShowComments
# explicitly and reports the inherited values, so a reference export is reproducible up to
# Word-stamped timestamps (CreationDate/ModDate differ on every export; compare references
# structurally, never by hash).
# No Word print/export mode suppresses comment balloons while showing markup (only
# ShowRevisions=false prints clean), so there is no balloon-free markup reference to export.
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
    [ValidateRange(0, 1)]
    [int] $RevisionsMode = 0,
    [bool] $ShowComments = $true,
    [switch] $RejectAll
)
$ErrorActionPreference = "Stop"
$inputFull = (Resolve-Path -LiteralPath $InputDocx).Path
$outputFull = [IO.Path]::GetFullPath($OutputPdf)
$outputDir = [IO.Path]::GetDirectoryName($outputFull)
if (-not [IO.Directory]::Exists($outputDir)) { [void][IO.Directory]::CreateDirectory($outputDir) }
$openPath = $inputFull
$tempCopy = $null
if ($RejectAll) {
    $tempCopy = [IO.Path]::Combine([IO.Path]::GetTempPath(), [IO.Path]::GetRandomFileName() + ".docx")
    [IO.File]::Copy($inputFull, $tempCopy, $true)
    $openPath = $tempCopy
}
$word = New-Object -ComObject Word.Application
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    Write-Output ("Word version: " + $word.Version)
    $readOnlyOpen = -not $RejectAll
    $doc = $word.Documents.OpenNoRepairDialog($openPath, $false, $readOnlyOpen, $false)
    if ($RejectAll) {
        $doc.Revisions.RejectAll()
    }
    $view = $doc.ActiveWindow.View
    $inheritedMarkupMode = $view.MarkupMode
    $inheritedRevisionsMode = $view.RevisionsMode
    Write-Output ("Inherited view: MarkupMode=" + $inheritedMarkupMode + " RevisionsMode=" + $inheritedRevisionsMode)
    $doc.ShowRevisions = $ShowRevisions
    $view.RevisionsView = $RevisionsView
    $view.MarkupMode = $MarkupMode
    $view.RevisionsMode = $RevisionsMode
    $view.ShowComments = $ShowComments
    if ([IO.File]::Exists($outputFull)) {
        [IO.File]::Delete($outputFull)
    }
    $doc.SaveAs2($outputFull, 17)
    $doc.Close($false)
    $length = (Get-Item -LiteralPath $outputFull).Length
    $sha = (Get-FileHash -LiteralPath $outputFull -Algorithm SHA256).Hash
    Write-Output ("Exported: " + $outputFull + " (" + $length + " bytes, sha256 " + $sha + ")")
    Write-Output ("ExportSettings: Word/" + $word.Version + " ShowRevisions=" + $ShowRevisions + " RevisionsView=" + $RevisionsView + " MarkupMode=" + $MarkupMode + " RevisionsMode=" + $RevisionsMode + " ShowComments=" + $ShowComments + " RejectAll=" + [bool]$RejectAll + " (inherited MarkupMode=" + $inheritedMarkupMode + " RevisionsMode=" + $inheritedRevisionsMode + ")")
} finally {
    try { $word.Quit() } catch { }
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($view) } catch { }
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) } catch { }
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch { }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ($tempCopy -ne $null -and [IO.File]::Exists($tempCopy)) {
        [IO.File]::Delete($tempCopy)
    }
}
