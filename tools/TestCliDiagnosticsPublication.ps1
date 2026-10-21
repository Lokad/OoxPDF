# Exercise real library/CLI publication without a build or Office; aliases are owned.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)] [string] $CliPath,
    [string] $OutputDirectory
)
$ErrorActionPreference='Stop'
$cli=(Resolve-Path -LiteralPath $CliPath).Path
$library=Join-Path (Split-Path $cli) 'Lokad.OoxPdf.dll'
$cliHash=(Get-FileHash $cli).Hash;$libraryHash=(Get-FileHash $library).Hash
if (!$OutputDirectory) { $OutputDirectory=Join-Path (Split-Path $PSScriptRoot) ('artifacts/cli-publication-'+[guid]::NewGuid().ToString('N')) }
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Use a fresh test directory' }
$root=New-Item -ItemType Directory -Path $OutputDirectory
$results=@()
foreach ($kind in @('HardLink','SymbolicLink','Separate','Directory')) {
    foreach ($target in $(if ($kind -in @('HardLink','SymbolicLink')) { @('input','output') } else { @('none') })) {
        $run=New-Item -ItemType Directory -Path (Join-Path $root.FullName ($kind+'-'+$target))
        $inputFile=Join-Path $run.FullName 'source.txt';$outputFile=Join-Path $run.FullName 'existing.pdf'
        $diagnostics=Join-Path $run.FullName 'diagnostics.json'
        [IO.File]::WriteAllText($inputFile,'Owned source sentinel')
        [IO.File]::WriteAllText($outputFile,'Owned destination sentinel')
        $before=@((Get-FileHash $inputFile).Hash,(Get-FileHash $outputFile).Hash)
        if ($kind -in @('HardLink','SymbolicLink')) {
            try { New-Item -ItemType $kind -Path $diagnostics -Target $(if ($target -eq 'input') {$inputFile} else {$outputFile}) | Out-Null }
            catch {
                if ($kind -ne 'SymbolicLink' -or !$IsWindows -or $_.Exception.Message -notmatch 'privilege|permission|access|1314') { throw }
                $results+=@{Kind=$kind;Target=$target;Outcome='skipped';Reason='Windows symbolic-link privilege unavailable'}
                continue
            }
        }
        if ($kind -eq 'Directory') {
            New-Item -ItemType Directory -Path $diagnostics | Out-Null
            [IO.File]::WriteAllText((Join-Path $diagnostics 'sentinel'),'Retain this directory')
        }
        & dotnet $cli convert $inputFile $outputFile --diagnostics $diagnostics *> (Join-Path $run.FullName 'cli.log')
        if ($LASTEXITCODE -ne 1) { throw ('Expected conversion-failure exit1: '+$kind+'/'+$target) }
        if ((Get-FileHash $inputFile).Hash -ne $before[0] -or (Get-FileHash $outputFile).Hash -ne $before[1]) {
            throw ('Diagnostics corrupted an aliased source/destination: '+$kind+'/'+$target)
        }
        if ($kind -eq 'Directory') {
            if ([IO.File]::ReadAllText((Join-Path $diagnostics 'sentinel')) -ne 'Retain this directory') { throw 'Failed publication changed its destination' }
        }
        else {
            $items=@(Get-Content $diagnostics -Raw | ConvertFrom-Json)
            if (@($items | Where-Object Id -EQ 'OOXML_CONVERSION_FAILED').Count -ne 1) { throw 'Separate diagnostics were not published' }
            if ((Get-Item -LiteralPath $diagnostics).LinkType -eq 'SymbolicLink') { throw 'Publication followed the symlink instead of replacing it' }
        }
        if (@(Get-ChildItem -LiteralPath $run.FullName -Filter '*.tmp-*').Count) { throw 'Diagnostics staging file leaked' }
        $results+=@{Kind=$kind;Target=$target;Outcome='passed'}
    }
}
[Reflection.Assembly]::LoadFrom($library) | Out-Null
$fixture=Join-Path (Split-Path $PSScriptRoot) 'tests/Lokad.OoxPdf.Tests/Cases/docx-basic-paragraphs.docx'
$fixtureHash=(Get-FileHash $fixture).Hash
$source=[IO.File]::ReadAllBytes($fixture)
foreach ($kind in @('LibraryAlias','LibraryAncestorAlias','LibraryTargetAncestorAlias','CliInputAlias','CliOutputAlias','CliUncreatedOutputAlias','DistinctLinkedDestination','LibraryFileLink')) {
    $run=New-Item -ItemType Directory -Path (Join-Path $root.FullName $kind)
    $real=New-Item -ItemType Directory -Path (Join-Path $run.FullName 'real')
    $alias=Join-Path $run.FullName 'alias'
    $linkKind=if ($IsWindows) {'Junction'} else {'SymbolicLink'}
    $container=if ($kind -in @('LibraryAncestorAlias','LibraryTargetAncestorAlias')) {'nested'} else {''}
    $realDirectory=Join-Path $real.FullName $container
    New-Item -ItemType Directory -Path $realDirectory -Force | Out-Null
    $targetAlias=$null
    if ($kind -eq 'LibraryTargetAncestorAlias') {
        $targetAlias=Join-Path $run.FullName 'target-alias'
        New-Item -ItemType $linkKind -Path $targetAlias -Target $real.FullName | Out-Null
        New-Item -ItemType $linkKind -Path $alias -Target (Join-Path $targetAlias $container) | Out-Null
        $aliasDirectory=$alias
    }
    else {
        New-Item -ItemType $linkKind -Path $alias -Target $real.FullName | Out-Null
        $aliasDirectory=Join-Path $alias $container
    }
    $inputFile=Join-Path $realDirectory 'source.docx'
    [IO.File]::WriteAllBytes($inputFile,$source)
    $outputFile=Join-Path $realDirectory 'result.pdf'
    $existing='Owned PDF sentinel'
    if ($kind -ne 'CliUncreatedOutputAlias') { [IO.File]::WriteAllText($outputFile,$existing) }
    try {
        if ($kind -in @('LibraryAlias','LibraryAncestorAlias','LibraryTargetAncestorAlias')) {
            $failure=$null
            try { [Lokad.OoxPdf.OoxPdfConverter]::Convert($inputFile,(Join-Path $aliasDirectory 'source.docx')) }
            catch { $failure=$_.Exception.GetBaseException() }
            if ($failure -isnot [ArgumentException]) { throw ('Directory alias was not rejected: '+$kind) }
        }
        elseif ($kind -like 'Cli*') {
            $diagnostics=Join-Path $aliasDirectory $(if ($kind -eq 'CliInputAlias') {'source.docx'} else {'result.pdf'})
            & dotnet $cli convert $inputFile $outputFile --diagnostics $diagnostics *> (Join-Path $run.FullName 'cli.log')
            if ($LASTEXITCODE -ne 2) { throw ('CLI directory alias must refuse before conversion: '+$kind) }
            if ($kind -eq 'CliUncreatedOutputAlias') {
                if (Test-Path -LiteralPath $outputFile) { throw 'Refused CLI created output' }
            }
        }
        else {
            if ($kind -eq 'LibraryFileLink') {
                Remove-Item -LiteralPath $outputFile
                New-Item -ItemType HardLink -Path $outputFile -Target $inputFile | Out-Null
            }
            [Lokad.OoxPdf.OoxPdfConverter]::Convert($inputFile,(Join-Path $aliasDirectory 'result.pdf'))
            if (![Text.Encoding]::ASCII.GetString([IO.File]::ReadAllBytes($outputFile),0,5).StartsWith('%PDF-')) { throw 'Distinct linked destination is not PDF' }
        }
        if ((Get-FileHash $inputFile).Hash -ne $fixtureHash) { throw ('Publication replaced its input: '+$kind) }
        if ($kind -in @('LibraryAlias','LibraryAncestorAlias','LibraryTargetAncestorAlias','CliInputAlias','CliOutputAlias') -and [IO.File]::ReadAllText($outputFile) -ne $existing) {
            throw ('Refused publication replaced its PDF: '+$kind)
        }
        if (@(Get-ChildItem -LiteralPath $realDirectory -Filter '*.tmp-*').Count) { throw 'Directory-alias staging file leaked' }
        $results+=@{Kind=$kind;Target='directory-entry';Outcome='passed'}
    }
    finally {
        # Remove only this owned link, so later recursive test cleanup cannot follow it.
        [IO.Directory]::Delete($alias)
        if ($targetAlias) { [IO.Directory]::Delete($targetAlias) }
    }
}
foreach ($kind in @('LibraryInputLink','LibraryInputLinkSameEntry','CliInputLinkSameEntry','LibraryInputLinkChain','LibraryInputLinkTargetAncestor','CliInputLinkOutput','CliInputLinkDiagnostics','LibraryOutputLinkControl','LibraryInputHardLinkControl')) {
    $run=New-Item -ItemType Directory -Path (Join-Path $root.FullName $kind)
    $inputFile=Join-Path $run.FullName 'source.docx'
    [IO.File]::WriteAllBytes($inputFile,$source)
    $outputFile=Join-Path $run.FullName 'result.pdf'
    $alias=Join-Path $run.FullName 'input-link.docx'
    $directoryAlias=$null
    try {
        if ($kind -eq 'LibraryOutputLinkControl') {
            New-Item -ItemType SymbolicLink -Path $outputFile -Target $inputFile | Out-Null
        }
        else {
            $target=$inputFile
            if ($kind -eq 'LibraryInputLinkTargetAncestor') {
                $directoryAlias=Join-Path $run.FullName 'parent-link'
                New-Item -ItemType $(if ($IsWindows) {'Junction'} else {'SymbolicLink'}) -Path $directoryAlias -Target $run.FullName | Out-Null
                $target=Join-Path $directoryAlias 'source.docx'
            }
            if ($kind -eq 'LibraryInputLinkChain') {
                $middle=Join-Path $run.FullName 'middle.docx'
                New-Item -ItemType SymbolicLink -Path $middle -Target $target | Out-Null
                $target=$middle
            }
            New-Item -ItemType $(if ($kind -eq 'LibraryInputHardLinkControl') {'HardLink'} else {'SymbolicLink'}) -Path $alias -Target $target | Out-Null
        }
    }
    catch {
        if ($directoryAlias) { [IO.Directory]::Delete($directoryAlias) }
        if (!$IsWindows -or $_.Exception.Message -notmatch 'privilege|permission|access|1314') { throw }
        $results+=@{Kind=$kind;Target='input-referent';Outcome='skipped';Reason='Windows symbolic-link privilege unavailable'}
        continue
    }
    try {
        if ($kind -like 'Cli*') {
            $diagnostics=if ($kind -eq 'CliInputLinkSameEntry') {$alias} elseif ($kind -eq 'CliInputLinkDiagnostics') {$inputFile} else {Join-Path $run.FullName 'diagnostics.json'}
            $destination=if ($kind -eq 'CliInputLinkOutput') {$inputFile} else {$outputFile}
            & dotnet $cli convert $alias $destination --diagnostics $diagnostics *> (Join-Path $run.FullName 'cli.log')
            $expected=if ($kind -in @('CliInputLinkDiagnostics','CliInputLinkSameEntry')) {2} else {1}
            if ($LASTEXITCODE -ne $expected) { throw ('CLI input referent was not protected: '+$kind) }
        }
        elseif ($kind -in @('LibraryOutputLinkControl','LibraryInputHardLinkControl')) {
            $readPath=if ($kind -eq 'LibraryInputHardLinkControl') {$alias} else {$inputFile}
            $destination=if ($kind -eq 'LibraryInputHardLinkControl') {$inputFile} else {$outputFile}
            [Lokad.OoxPdf.OoxPdfConverter]::Convert($readPath,$destination)
            if ((Get-FileHash $readPath).Hash -ne $fixtureHash) { throw ('Separate file-link replacement changed input: '+$kind) }
            if (![Text.Encoding]::ASCII.GetString([IO.File]::ReadAllBytes($destination),0,5).StartsWith('%PDF-')) { throw 'Allowed link destination is not PDF' }
        }
        else {
            $failure=$null
            try { [Lokad.OoxPdf.OoxPdfConverter]::Convert($alias,$(if ($kind -eq 'LibraryInputLinkSameEntry') {$alias} else {$inputFile})) }
            catch { $failure=$_.Exception.GetBaseException() }
            if ($failure -isnot [ArgumentException]) { throw ('Input file referent was not rejected: '+$kind) }
        }
        if ($kind -ne 'LibraryInputHardLinkControl' -and (Get-FileHash $inputFile).Hash -ne $fixtureHash) { throw ('Input file referent changed: '+$kind) }
        if (@(Get-ChildItem -LiteralPath $run.FullName -Filter '*.tmp-*').Count) { throw 'Input-link staging file leaked' }
        $results+=@{Kind=$kind;Target='input-referent';Outcome='passed'}
    }
    finally { if ($directoryAlias) { [IO.Directory]::Delete($directoryAlias) } }
}
if ((Get-FileHash $fixture).Hash -ne $fixtureHash) { throw 'Original public fixture changed' }
if ((Get-FileHash $cli).Hash -ne $cliHash -or (Get-FileHash $library).Hash -ne $libraryHash) { throw 'CLI runtime changed' }
@{CliSha256=$cliHash;LibrarySha256=$libraryHash;Builds=0;OfficeExports=0;Results=$results} |
    ConvertTo-Json -Depth 5 | Set-Content (Join-Path $root.FullName 'evidence.json') -Encoding utf8
Write-Output ('CLI diagnostics publication: '+@($results | Where-Object Outcome -EQ 'passed').Count+' pass, '+@($results | Where-Object Outcome -EQ 'skipped').Count+' skip; no build.')
$global:LASTEXITCODE=0
