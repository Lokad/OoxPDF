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
    $view.ShowComments = $ShowComments
    $view.ShowFormatChanges = $ShowFormatChanges
    $view.ShowInsertionsAndDeletions = $ShowInsertionsAndDeletions
    # Visibility setters can reset MarkupMode when the last category is hidden.
    # Apply the requested display modes last, then verify the entire state.
    $view.MarkupMode = $MarkupMode
    $view.RevisionsMode = $RevisionsMode
    # Word 16 accepts mixed mode 2. Verify the actual state rather than trusting
    # an enum label or a setter that could silently retain another session's view.
    $actual = [ordered]@{ShowRevisions=[bool]$doc.ShowRevisions;RevisionsView=[int]$view.RevisionsView;
        MarkupMode=[int]$view.MarkupMode;RevisionsMode=[int]$view.RevisionsMode;ShowComments=[bool]$view.ShowComments;
        ShowFormatChanges=[bool]$view.ShowFormatChanges;ShowInsertionsAndDeletions=[bool]$view.ShowInsertionsAndDeletions}
    $expected = [ordered]@{ShowRevisions=$ShowRevisions;RevisionsView=$RevisionsView;MarkupMode=$MarkupMode;
        RevisionsMode=$RevisionsMode;ShowComments=$ShowComments;ShowFormatChanges=$ShowFormatChanges;
        ShowInsertionsAndDeletions=$ShowInsertionsAndDeletions}
    foreach ($key in $expected.Keys) {
        if ($actual[$key] -ne $expected[$key]) { throw ('Word did not apply requested view setting: '+$key) }
    }
    if ([IO.File]::Exists($outputFull)) {
        [IO.File]::Delete($outputFull)
    }
    $doc.SaveAs2($outputFull, 17)
    $doc.Close($false)
    $length = (Get-Item -LiteralPath $outputFull).Length
    $sha = (Get-FileHash -LiteralPath $outputFull -Algorithm SHA256).Hash
    Write-Output ("Exported: " + $outputFull + " (" + $length + " bytes, sha256 " + $sha + ")")
    $settings = ($actual.Keys | ForEach-Object { $_+'='+$actual[$_] }) -join ' '
    Write-Output ("ExportSettings: Word/" + $word.Version + " " + $settings + " RejectAll=" + [bool]$RejectAll + " (inherited MarkupMode=" + $inheritedMarkupMode + " RevisionsMode=" + $inheritedRevisionsMode + ")")
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
