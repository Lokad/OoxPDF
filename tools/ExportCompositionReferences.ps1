# Export each pinned source/mode once; reuse existing independent references.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)] [string] $Manifest,
    [Parameter(Mandatory=$true)] [string] $OutputDirectory,
    [ValidateSet('Both','Final','Review')] [string] $DocxPreset = 'Both'
)
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
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
        if (Test-Path -LiteralPath $identity) {
            $pin = Get-Content -LiteralPath $identity -Raw | ConvertFrom-Json
            if ($pin.InputSha256 -ne $document.Sha256 -or $pin.Mode -ne $mode -or
                (Get-FileHash $pdf).Hash -ne $pin.PdfSha256) { throw 'Cached reference differs' }
            continue
        }
        if (Test-Path -LiteralPath $folder) { throw 'Unpinned reference directory exists; inspect it before reuse' }
        New-Item -ItemType Directory -Force -Path $folder | Out-Null
        if ($docx) {
            & (Join-Path $PSScriptRoot 'ExportDocxMarkupReference.ps1') -InputDocx $inputPath -OutputPdf $pdf -ShowRevisions ($mode -eq 'all-word-compatible') -ShowComments ($mode -eq 'all-word-compatible') *> (Join-Path $folder 'export.log')
        }
        else {
            & (Join-Path $PSScriptRoot 'RenderReference.ps1') -InputPath $inputPath -OutputDirectory $folder *> (Join-Path $folder 'export.log')
        }
        [ordered]@{Id=$document.Id;Mode=$mode;InputSha256=$document.Sha256;PdfSha256=(Get-FileHash $pdf).Hash} |
            ConvertTo-Json | Set-Content -LiteralPath $identity -Encoding utf8
        Write-Output ($document.Id+' '+$mode+' independent reference retained')
    }
}
