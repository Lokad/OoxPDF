function Get-CompositionWordView([string] $Mode) {
    if ($Mode -ceq 'presentation') { return }
    if ($Mode -cnotin @('final-preserve', 'all-word-compatible')) { throw ('Unknown portfolio mode: ' + $Mode) }
    $review = $Mode -ceq 'all-word-compatible'
    $markupMode = if ($review) { 2 } else { 0 }
    [ordered]@{
        ShowRevisions=$review; RevisionsView=0; MarkupMode=$markupMode; RevisionsMode=$markupMode
        ShowComments=$review; ShowFormatChanges=$review; ShowInsertionsAndDeletions=$review; RejectAll=$false
    }
}

function Read-CompositionWordExport([string] $Log, $View, [string] $Case) {
    if (!(Test-Path -LiteralPath $Log -PathType Leaf)) { throw ('Reference Office view is unverified: ' + $Case) }
    $bytes = [IO.File]::ReadAllBytes((Resolve-Path -LiteralPath $Log).ProviderPath)
    $reader = [IO.StreamReader]::new([IO.MemoryStream]::new($bytes))
    try { $lines = @($reader.ReadToEnd() -split '\r?\n' | Where-Object { $_ -match '^ExportSettings: Word/' }) }
    finally { $reader.Dispose() }
    if ($lines.Count -ne 1) { throw ('Reference Office view is unverified: ' + $Case) }
    foreach ($key in $View.Keys) {
        $value = [regex]::Match($lines[0], '(?:^| )' + $key + '=([^ ]+)')
        if (!$value.Success -or $value.Groups[1].Value -cne [string]$View[$key]) {
            throw ('Reference Office view differs: ' + $Case + ' / ' + $key)
        }
    }
    [pscustomobject]@{ Settings=$lines[0]; Sha256=[Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes)) }
}

function Read-CompositionReference([string] $Directory, [string] $Id, [string] $Mode, [string] $InputSha256, [string] $Kind) {
    $case = $Id + '-' + $Mode
    $view = Get-CompositionWordView $Mode
    $identity = Join-Path $Directory 'source-identity.json'
    if (!(Test-Path -LiteralPath $identity -PathType Leaf)) { throw ('Reference identity missing: ' + $case) }
    $identityHash = (Get-FileHash -LiteralPath $identity).Hash
    $pin = Get-Content -LiteralPath $identity -Raw | ConvertFrom-Json -AsHashtable
    $referenceKind = if ($pin.ReferenceKind) { $pin.ReferenceKind } else { 'OfficePdf' }
    if ($referenceKind -ne $Kind) { throw ('Cached reference kind differs: ' + $case) }
    $pdf = Join-Path $Directory 'reference.pdf'
    if (!(Test-Path -LiteralPath $pdf -PathType Leaf)) { throw ('Reference PDF missing: ' + $case) }
    $pdfHash = (Get-FileHash -LiteralPath $pdf).Hash
    if ($pin.Id -cne $Id -or $pin.Mode -cne $Mode -or $pin.InputSha256 -ne $InputSha256 -or $pin.PdfSha256 -ne $pdfHash) {
        throw ('Reference identity differs: ' + $case)
    }

    $log = Join-Path $Directory 'export.log'
    $logHash = $null
    if ($view) {
        $export = Read-CompositionWordExport $log $view $case
        $logHash = $export.Sha256
        $hasViewPin = $pin.ContainsKey('OfficeView')
        if ($hasViewPin) {
            foreach ($key in $view.Keys) {
                if ($pin.OfficeView -isnot [Collections.IDictionary] -or [string]$pin.OfficeView[$key] -cne [string]$view[$key]) {
                    throw ('Cached reference Office view differs: ' + $case + ' / ' + $key)
                }
            }
        }
        if (($hasViewPin -or $pin.ContainsKey('ExportSettings')) -and $pin.ExportSettings -cne $export.Settings -or
            $hasViewPin -and !$pin.ExportLogSha256) {
            throw ('Reference export provenance differs: ' + $case)
        }
    }
    elseif ($pin.ExportLogSha256 -and (Test-Path -LiteralPath $log -PathType Leaf)) {
        $logHash = (Get-FileHash -LiteralPath $log).Hash
    }
    if ($pin.ExportLogSha256 -and $pin.ExportLogSha256 -ne $logHash) {
        throw ('Reference export provenance differs: ' + $case)
    }
    if ((Get-FileHash -LiteralPath $identity).Hash -ne $identityHash) { throw ('Reference identity changed: ' + $case) }

    [pscustomobject]@{ Pdf=$pdf; Hash=$pdfHash; Identity=$identity; IdentityHash=$identityHash;
        ExportLog=$log; ExportLogHash=$logHash; Pin=$pin }
}
