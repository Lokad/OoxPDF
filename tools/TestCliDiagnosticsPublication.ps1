# Exercise the real CLI without a build or Office; each alias names owned inputs.
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
if ((Get-FileHash $cli).Hash -ne $cliHash -or (Get-FileHash $library).Hash -ne $libraryHash) { throw 'CLI runtime changed' }
@{CliSha256=$cliHash;LibrarySha256=$libraryHash;Builds=0;OfficeExports=0;Results=$results} |
    ConvertTo-Json -Depth 5 | Set-Content (Join-Path $root.FullName 'evidence.json') -Encoding utf8
Write-Output ('CLI diagnostics publication: '+@($results | Where-Object Outcome -EQ 'passed').Count+' pass, '+@($results | Where-Object Outcome -EQ 'skipped').Count+' skip; no build.')
