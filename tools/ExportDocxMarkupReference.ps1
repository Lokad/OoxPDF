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
    $doc = $word.Documents.OpenNoRepairDialog($openPath, $false, $true, $false)
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
    if ([IO.File]::Exists($OutputPdf)) {
        [IO.File]::Delete($OutputPdf)
    }
    $doc.SaveAs2($OutputPdf, 17)
    $doc.Close($false)
    $length = (Get-Item -LiteralPath $OutputPdf).Length
    $sha = (Get-FileHash -LiteralPath $OutputPdf -Algorithm SHA256).Hash
    Write-Output ("Exported: " + $OutputPdf + " (" + $length + " bytes, sha256 " + $sha + ")")
    Write-Output ("ExportSettings: Word/" + $word.Version + " ShowRevisions=" + $ShowRevisions + " RevisionsView=" + $RevisionsView + " MarkupMode=" + $MarkupMode + " RevisionsMode=" + $RevisionsMode + " ShowComments=" + $ShowComments + " RejectAll=" + [bool]$RejectAll + " (inherited MarkupMode=" + $inheritedMarkupMode + " RevisionsMode=" + $inheritedRevisionsMode + ")")
} finally {
    try { $word.Quit() } catch { }
    if ($tempCopy -ne $null -and [IO.File]::Exists($tempCopy)) {
        [IO.File]::Delete($tempCopy)
    }
}
