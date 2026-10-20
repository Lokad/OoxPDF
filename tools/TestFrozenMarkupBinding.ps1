# Real runtime-binding branches with fake executables; no library/Office execution.
[CmdletBinding()]
param([string] $OutputDirectory)
$ErrorActionPreference='Stop'
if (!$OutputDirectory) { $OutputDirectory=Join-Path (Split-Path $PSScriptRoot) ('artifacts/frozen-markup-test-'+[Guid]::NewGuid().ToString('N')) }
if (Test-Path $OutputDirectory) { throw 'Use a fresh test directory' }
$repoRoot=Join-Path $OutputDirectory 'repository'
New-Item -ItemType Directory -Force (Join-Path $repoRoot 'tools') | Out-Null
$repoRoot=(Resolve-Path $repoRoot).Path
foreach ($name in @('InspectDocx.ps1','InspectPptxText.ps1','InspectPdf.ps1')) { Copy-Item (Join-Path $PSScriptRoot $name) (Join-Path $repoRoot 'tools') }
@'
param($Project,$OutputDll,$Description,$RecordPath)
throw 'Frozen binding requested an unexpected build'
'@ | Set-Content (Join-Path $repoRoot 'tools/EnsureDotnetBuild.ps1')
$inputFile=Join-Path $repoRoot 'input.docx'
[IO.File]::WriteAllText($inputFile,'fake public input')
$runtime=@{}
foreach ($name in @('cli','docx','pptx','pdf')) {
    $directory=Join-Path $repoRoot ('frozen/'+$name)
    New-Item -ItemType Directory -Force $directory | Out-Null
    $runtime[$name]=Join-Path $directory 'probe.dll'
    [IO.File]::WriteAllText($runtime[$name],'fake inspector')
    [IO.File]::WriteAllText((Join-Path $directory 'Lokad.OoxPdf.dll'),'same fake library')
}
$libraryHash=(Get-FileHash (Join-Path (Split-Path $runtime.cli) 'Lokad.OoxPdf.dll')).Hash
$executionState=@{Count=0;Mutate=$false}
function dotnet {
    if ($args[0] -eq 'build') { throw 'A frozen inspector tried to build' }
    $executionState.Count++
    if ($executionState.Mutate) { [IO.File]::WriteAllText([string]$args[0],'changed during inspection') }
    $global:LASTEXITCODE=0
}
try {
    foreach ($name in @('docx','pptx','pdf')) {
        $output=Join-Path $repoRoot ('results/'+$name)
        $arguments=@{OutputDirectory=$output;InspectorPath=$runtime[$name];SkipBuild=$true}
        if ($name -eq 'docx') { $arguments.InputDocx=$inputFile;$arguments.ExpectedLibrarySha256=$libraryHash;$scriptName='InspectDocx.ps1' }
        elseif ($name -eq 'pptx') { $arguments.InputPptx=$inputFile;$arguments.ExpectedLibrarySha256=$libraryHash;$scriptName='InspectPptxText.ps1' }
        else { $arguments.InputPdf=$inputFile;$scriptName='InspectPdf.ps1' }
        & (Join-Path $repoRoot ('tools/'+$scriptName)) @arguments
        $record=Get-ChildItem $output -Filter '*build-info.json' | Select-Object -First 1
        $pin=Get-Content $record.FullName -Raw | ConvertFrom-Json
        if ($pin.BuildRequested -ne $false -or $pin.OutputSha256 -ne (Get-FileHash $runtime[$name]).Hash) { throw 'Inspector provenance does not identify the supplied output' }
    }
    if ($executionState.Count -ne 3) { throw 'Frozen inspectors did not run exactly once each' }
    $rejected=$false
    try { & (Join-Path $repoRoot 'tools/InspectDocx.ps1') -InputDocx $inputFile -OutputDirectory (Join-Path $repoRoot 'mismatch') -InspectorPath $runtime.docx -ExpectedLibrarySha256 ('0'*64) -SkipBuild }
    catch { if (!$_.Exception.Message.Contains('differs from the qualified converter')) { throw };$rejected=$true }
    if (!$rejected -or $executionState.Count -ne 3) { throw 'A mismatched library executed' }
    $executionState.Mutate=$true
    $rejected=$false
    try { & (Join-Path $repoRoot 'tools/InspectPdf.ps1') -InputPdf $inputFile -OutputDirectory (Join-Path $repoRoot 'changed') -InspectorPath $runtime.pdf -SkipBuild }
    catch { if (!$_.Exception.Message.Contains('changed during execution')) { throw };$rejected=$true }
    if (!$rejected) { throw 'A changed inspector was accepted' }
}
finally { Remove-Item Function:\dotnet }

# Execute the cached comparison's real pre-conversion binding statements, without
# entering document parsing, comparison or Office. No test hooks in production.
$comparison=Join-Path $PSScriptRoot 'CompareCachedDocxMarkupReference.ps1'
$source=Get-Content $comparison -Raw
$start=$source.IndexOf('$cliProject = Join-Path $repoRoot')
$end=$source.IndexOf('$candidatePdf = Join-Path $candidateDir',$start)
if ($start -lt 0 -or $end -lt $start) { throw 'Cached runtime binding is missing' }
$binding=[scriptblock]::Create($source.Substring($start,$end-$start))
$CliPath=$runtime.cli;$DocxInspectorPath=$runtime.docx;$SkipBuild=$true
$runRoot=Join-Path $repoRoot 'cached'
New-Item -ItemType Directory -Force $runRoot | Out-Null
. $binding
$pin=Get-Content (Join-Path $runRoot 'cli-build-info.json') -Raw | ConvertFrom-Json
if ($pin.LibrarySha256 -ne $libraryHash -or $pin.BuildRequested -ne $false) { throw 'Cached comparison lost frozen CLI identity' }
$DocxInspectorPath=$null
$rejected=$false
try { . $binding } catch { if (!$_.Exception.Message.Contains('matching DOCX inspector together')) { throw };$rejected=$true }
if (!$rejected) { throw 'An unpaired CLI was accepted' }
$DocxInspectorPath=$runtime.docx
[IO.File]::WriteAllText((Join-Path (Split-Path $runtime.docx) 'Lokad.OoxPdf.dll'),'different fake library')
$rejected=$false
try { . $binding } catch { if (!$_.Exception.Message.Contains('differs from the qualified converter')) { throw };$rejected=$true }
if (!$rejected) { throw 'Cached comparison accepted a different library' }
foreach ($script in @('CompareCachedDocxMarkupReference.ps1','RunDocxMarkupReferenceGate.ps1')) {
    $parameters=(Get-Command (Join-Path $PSScriptRoot $script)).Parameters
    foreach ($name in @('CliPath','DocxInspectorPath','PdfInspectorPath','SkipBuild')) {
        if (!$parameters.ContainsKey($name)) { throw ('Missing forwarded runtime option: '+$script+'/'+$name) }
    }
}
@{Checks=@('Frozen DOCX inspection','Frozen PPTX inspection','Frozen PDF inspection','Inspector identity records',
    'Wrong library refused before execution','Changed inspector refused','Cached CLI identity',
    'Unpaired CLI refused','Cached library mismatch refused','Gate runtime parameters');Builds=0;LibraryExecutions=0;OfficeExports=0} |
    ConvertTo-Json | Set-Content (Join-Path $OutputDirectory 'evidence.json')
'Ten frozen markup/inspection checks passed; no library build, real conversion or Office export.'
