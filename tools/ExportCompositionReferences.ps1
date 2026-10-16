# Export each pinned source/mode once; reuse existing independent references.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)] [string] $Manifest,
    [Parameter(Mandatory=$true)] [string] $OutputDirectory,
    [ValidateSet('Both','Final','Review')] [string] $DocxPreset = 'Both'
)
$ErrorActionPreference = 'Stop'
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
        $view = [ordered]@{ShowRevisions=$review;RevisionsView=0;MarkupMode=2;
            RevisionsMode=$(if ($review) {1} else {0});ShowComments=$review;RejectAll=$false}
        $log = Join-Path $folder 'export.log'
        if (Test-Path -LiteralPath $identity) {
            $pin = Get-Content -LiteralPath $identity -Raw | ConvertFrom-Json -AsHashtable
            if ($pin.InputSha256 -ne $document.Sha256 -or $pin.Mode -ne $mode -or
                (Get-FileHash $pdf).Hash -ne $pin.PdfSha256) { throw 'Cached reference differs' }
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
        if ($job.Docx) {
            $view = $job.View
            & (Join-Path $PSScriptRoot 'ExportDocxMarkupReference.ps1') -InputDocx $inputPath -OutputPdf $pdf @view *> $job.Log
            $settings = Assert-WordView $job.Log $view ($document.Id + '/' + $mode)
        }
        else {
            & (Join-Path $PSScriptRoot 'RenderReference.ps1') -InputPath $inputPath -OutputDirectory $folder *> (Join-Path $folder 'export.log')
        }
        $pin = [ordered]@{Id=$document.Id;Mode=$mode;InputSha256=$document.Sha256;PdfSha256=(Get-FileHash $pdf).Hash}
        if ($job.Docx) {
            $pin.OfficeView=$view; $pin.ExportSettings=$settings; $pin.ExportLogSha256=(Get-FileHash $job.Log).Hash
        }
        $pin | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $job.Identity -Encoding utf8
        Write-Output ($document.Id+' '+$mode+' independent reference retained')
}
