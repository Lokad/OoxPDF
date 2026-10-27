# Exercise the real family orchestration with fake builds/conversions, no Office or library execution.
[CmdletBinding()]
param([string] $OutputDirectory)
$ErrorActionPreference = 'Stop'
if (!$OutputDirectory) { $OutputDirectory = Join-Path (Split-Path $PSScriptRoot) ('artifacts/visual-family-test-' + [Guid]::NewGuid().ToString('N')) }
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Use a fresh test directory' }
$repo = Join-Path $OutputDirectory 'repository'
New-Item -ItemType Directory -Path (Join-Path $repo 'tools') -Force | Out-Null
$repo = (Resolve-Path -LiteralPath $repo).Path
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'CheckVisualFamily.ps1') -Destination (Join-Path $repo 'tools')
$buildLog = Join-Path $repo 'builds.jsonl'
$caseLog = Join-Path $repo 'cases.jsonl'
@'
param($Project, $OutputDll, $Description, $RecordPath)
@{Project=$Project;Description=$Description} | ConvertTo-Json -Compress | Add-Content (Join-Path (Split-Path $PSScriptRoot) 'builds.jsonl')
New-Item -ItemType Directory -Path (Split-Path $OutputDll) -Force | Out-Null
[IO.File]::WriteAllText($OutputDll, 'fake CLI or validation tool')
if ($Description -eq 'CLI') { [IO.File]::WriteAllText((Join-Path (Split-Path $OutputDll) 'Lokad.OoxPdf.dll'), 'fake library') }
'@ | Set-Content (Join-Path $repo 'tools/EnsureDotnetBuild.ps1')
@'
param($Case, $RunId, [switch]$SkipBuild, $CliPath, [switch]$CacheOnlyReference)
@{Case=$Case;RunId=$RunId;SkipBuild=[bool]$SkipBuild;CliPath=$CliPath;CacheOnlyReference=[bool]$CacheOnlyReference} |
    ConvertTo-Json -Compress | Add-Content (Join-Path (Split-Path $PSScriptRoot) 'cases.jsonl')
if (!$SkipBuild -or !(Test-Path -LiteralPath $CliPath)) { throw 'Family did not bind an existing CLI without per-case builds' }
'@ | Set-Content (Join-Path $repo 'tools/CheckVisualCase.ps1')
New-Item -ItemType Directory -Path (Join-Path $repo 'visual-cases/families') -Force | Out-Null
'{"casePatterns":["probe-*"]}' | Set-Content (Join-Path $repo 'visual-cases/families/probe.json')
foreach ($id in @('probe-a','probe-b')) {
    $directory = Join-Path $repo "visual-cases/cases/$id"
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
    @{id=$id;kind='pptx';tags=@('approximate')} | ConvertTo-Json | Set-Content (Join-Path $directory 'case.json')
}
$frozenCli = Join-Path $repo 'frozen/Lokad.OoxPdf.Cli.dll'
New-Item -ItemType Directory -Path (Split-Path $frozenCli) -Force | Out-Null
[IO.File]::WriteAllText($frozenCli, 'frozen fake CLI')
$frozenLibrary = Join-Path (Split-Path $frozenCli) 'Lokad.OoxPdf.dll'
[IO.File]::WriteAllText($frozenLibrary, 'frozen fake library')
$script = Join-Path $repo 'tools/CheckVisualFamily.ps1'
$listed = @(& $script -Family probe -List -CliPath 'missing.dll')
if ($listed.Count -ne 2 -or (Test-Path $buildLog) -or (Test-Path $caseLog)) { throw 'Listing performed builds or required a runtime' }
& $script -Family probe -CliPath $frozenCli -CacheOnlyReference | Out-Null
$builds = @(Get-Content $buildLog | ConvertFrom-Json)
$cases = @(Get-Content $caseLog | ConvertFrom-Json)
if ($builds.Count -ne 2 -or @($builds | Where-Object Description -eq 'CLI').Count -or
    $cases.Count -ne 2 -or @($cases | Where-Object { $_.CliPath -ne $frozenCli -or !$_.CacheOnlyReference }).Count) {
    throw 'Frozen family binding rebuilt the CLI or lost case arguments'
}
$report = Get-Content (Join-Path $repo 'artifacts/visual/reports/probe.json') -Raw | ConvertFrom-Json
if ($report.cliSha256 -ne (Get-FileHash $frozenCli).Hash -or $report.librarySha256 -ne (Get-FileHash $frozenLibrary).Hash) {
    throw 'Family report omitted the actual runtime identity'
}
& $script -Family probe -CliPath $frozenCli -SkipBuild | Out-Null
if (@(Get-Content $buildLog).Count -ne 2 -or @(Get-Content $caseLog).Count -ne 4) { throw 'SkipBuild rebuilt tools or missed cases' }
& $script -Family probe | Out-Null
$builds = @(Get-Content $buildLog | ConvertFrom-Json)
$cases = @(Get-Content $caseLog | ConvertFrom-Json)
if ($builds.Count -ne 5 -or @($builds | Where-Object Description -eq 'CLI').Count -ne 1 -or $cases.Count -ne 6) {
    throw 'Default family execution no longer builds each required project once'
}
$defaultCli = $cases[-1].CliPath
& $script -Family probe -SkipBuild | Out-Null
$cases = @(Get-Content $caseLog | ConvertFrom-Json)
if (@(Get-Content $buildLog).Count -ne 5 -or $cases.Count -ne 8 -or $cases[-1].CliPath -ne $defaultCli) { throw 'Prebuilt default binding failed' }
# Every inspection capability prepares the tool once for both cases. False font
# flags leave raster-only families free of unused inspection builds.
$inspectionKeys=@('maxTextOperationPositionDelta','maxTextLineStartDelta','maxGraphicsOperationBoundsDelta',
    'maxChartGraphicsStructureBoundsDelta','maxChartGraphicsStructureBoundsDeltaByKind',
    'maxChartTextStructurePositionDelta','maxChartTextStructurePositionDeltaByKind',
    'candidateFontsMustBeEmbedded','candidateFontsMustHaveToUnicode','candidateFontsMustBeSubset')
foreach($key in $inspectionKeys) {
    foreach($id in @('probe-a','probe-b')) {
        $expected=@{}; $expected[$key]=if($key.StartsWith('candidateFonts')) {$true} else {0}
        @{id=$id;kind='docx';tags=@('approximate');expected=$expected} | ConvertTo-Json -Depth 4 |
            Set-Content (Join-Path $repo "visual-cases/cases/$id/case.json")
    }
    $before=@(Get-Content $buildLog).Count
    & $script -Family probe -CliPath $frozenCli | Out-Null
    $added=@(Get-Content $buildLog | Select-Object -Skip $before | ConvertFrom-Json)
    if($added.Count -ne 3 -or @($added | Where-Object Description -eq 'PDF inspect').Count -ne 1) {throw ('Inspection tool was not prepared exactly once: '+$key)}
}
$before=@(Get-Content $buildLog).Count
& $script -Family probe -CliPath $frozenCli -SkipBuild | Out-Null
if(@(Get-Content $buildLog).Count -ne $before) {throw 'An inspecting prebuilt family requested a build'}
$inspector=Join-Path $repo 'tools/Lokad.OoxPdf.PdfInspect/bin/Debug/net10.0/Lokad.OoxPdf.PdfInspect.dll'
if(!$inspector.StartsWith($repo+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) {throw 'Unexpected synthetic inspector path'}
Remove-Item -LiteralPath $inspector
$caseCount=@(Get-Content $caseLog).Count
$failure=$null
try { & $script -Family probe -CliPath $frozenCli -SkipBuild | Out-Null } catch {$failure=$_.Exception.Message}
if($failure -notlike '*PDF inspector is missing*' -or @(Get-Content $buildLog).Count -ne $before -or @(Get-Content $caseLog).Count -ne $caseCount) {throw 'Missing prebuilt inspector did not fail before cases/builds'}
[IO.File]::WriteAllText($inspector,'prebuilt synthetic PDF inspector')
foreach($id in @('probe-a','probe-b')) {
    @{id=$id;kind='docx';tags=@('approximate');expected=@{candidateFontsMustBeEmbedded=$false}} | ConvertTo-Json -Depth 4 |
        Set-Content (Join-Path $repo "visual-cases/cases/$id/case.json")
}
& $script -Family probe -CliPath $frozenCli | Out-Null
if(@(Get-Content $buildLog).Count -ne $before+2) {throw 'A false font flag caused an unused inspection build'}

# Cross the real case -> inspector boundary. Only native executions and Office/
# raster helpers are stubbed. The real scripts decide preparation and forwarding.
foreach($name in @('CheckVisualCase.ps1','InspectPdf.ps1','JsonArray.ps1','CompareMath.ps1','ComparePdfTextLineStarts.ps1')) {
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot $name) -Destination (Join-Path $repo 'tools') -Force
}
@'
param($InputPath,$OutputDirectory,$Dpi,$CaseId,[switch]$SkipBuild)
[IO.File]::WriteAllText((Join-Path $OutputDirectory 'reference.pdf'),'%PDF-fake reference')
[IO.File]::WriteAllText((Join-Path $OutputDirectory 'page-001.png'),'synthetic raster')
'@ | Set-Content (Join-Path $repo 'tools/RenderCachedReference.ps1')
@'
param($InputPdf,$OutputDirectory,$Dpi,[switch]$SkipBuild)
[IO.File]::WriteAllText((Join-Path $OutputDirectory 'page-001.png'),'synthetic raster')
'@ | Set-Content (Join-Path $repo 'tools/RasterizePdf.ps1')
foreach($id in @('probe-a','probe-b')) {
    $directory=Join-Path $repo "visual-cases/cases/$id"
    [IO.File]::WriteAllText((Join-Path $directory 'input.docx'),'synthetic orchestration input')
    @{id=$id;kind='docx';input='input.docx';tags=@('approximate');expected=@{
        maxTextLineStartDelta=0.15;maxTextLineStartYDelta=0.5;candidateFontsMustBeEmbedded=$true}} | ConvertTo-Json -Depth 4 |
        Set-Content (Join-Path $directory 'case.json')
}
$inspectionLog=Join-Path $repo 'inspection-calls.jsonl'
function dotnet {
    $values=@($args)
    if($values[0] -like '*Lokad.OoxPdf.Cli.dll') {
        [IO.File]::WriteAllText($values[3],'%PDF-fake candidate')
        [IO.File]::WriteAllText($values[5],'[]')
    } elseif($values[0] -like '*Lokad.OoxPdf.VisualDiff.dll') {
        '[{"Page":1,"DimensionsMatch":true,"MeanAbsoluteError":0}]' | Set-Content (Join-Path $values[3] 'metrics.json')
    } elseif($values[0] -like '*Lokad.OoxPdf.PdfInspect.dll') {
        $output=$values[2]
        $baseline=if($values[1] -like '*output.pdf' -and $output -like '*real-y-shift*') {500} else {700}
        @(@{PageNumber=1;X=72;Y=$baseline;DecodedText='Visible content'}) | ConvertTo-Json -AsArray | Set-Content (Join-Path $output 'text-operations.json')
        '[{"PageNumber":1,"ResourceName":"F1","Embedded":true,"HasToUnicode":true,"IsSubset":true}]' | Set-Content (Join-Path $output 'font-resources.json')
        @{Inspector=$values[0];InspectorSha256=(Get-FileHash $values[0]).Hash;Output=$output} | ConvertTo-Json -Compress |
            Add-Content $inspectionLog
    } else {throw 'Unexpected native execution request'}
    $global:LASTEXITCODE=0
}
$realCase=Join-Path $repo 'tools/CheckVisualCase.ps1'
$realManifest=Join-Path $repo 'visual-cases/cases/probe-a/case.json'
try {
    $before=@(Get-Content $buildLog).Count
    & $realCase -Case $realManifest -CliPath $frozenCli -SkipBuild -RunId 'real-prebuilt' | Out-Null
    if(@(Get-Content $buildLog).Count -ne $before) {throw 'Real prebuilt case requested a build'}
    & $realCase -Case $realManifest -CliPath $frozenCli -RunId 'real-prepare-once' | Out-Null
    $added=@(Get-Content $buildLog | Select-Object -Skip $before | ConvertFrom-Json)
    if($added.Count -ne 2 -or @($added | Where-Object Description -eq 'PDF inspect').Count -ne 1) {throw 'Real direct case did not prepare inspector once'}
    $before=@(Get-Content $buildLog).Count
    & $script -Family probe -CliPath $frozenCli | Out-Null
    $added=@(Get-Content $buildLog | Select-Object -Skip $before | ConvertFrom-Json)
    if($added.Count -ne 3 -or @($added | Where-Object Description -eq 'PDF inspect').Count -ne 1) {throw 'Real family repeated per-case inspector builds'}
    $before=@(Get-Content $buildLog).Count
    & $script -Family probe -CliPath $frozenCli -SkipBuild | Out-Null
    if(@(Get-Content $buildLog).Count -ne $before) {throw 'Real prebuilt family requested a build'}
    $failure=$null
    try { & $realCase -Case $realManifest -CliPath $frozenCli -SkipBuild -RunId 'real-y-shift' | Out-Null } catch {$failure=$_.Exception.Message}
    if($failure -ne 'PDF text line-start gate failed.' -or @(Get-Content $buildLog).Count -ne $before) {throw 'The real case did not enforce its manifest Y bound without building'}
    Remove-Item -LiteralPath $inspector
    $failure=$null
    try { & $realCase -Case $realManifest -CliPath $frozenCli -SkipBuild -RunId 'real-missing-tool' | Out-Null } catch {$failure=$_.Exception.Message}
    if($failure -notlike '*PDF inspect DLL is missing*' -or @(Get-Content $buildLog).Count -ne $before) {throw 'Real missing-tool case requested a build or ignored refusal'}
}
finally {Remove-Item Function:/dotnet}
$inspections=@(Get-Content (Join-Path $repo 'inspection-calls.jsonl') | ConvertFrom-Json)
if($inspections.Count -ne 20) {throw 'Real inspection coverage is incomplete'}
foreach($inspection in $inspections) {
    $record=Get-Content (Join-Path $inspection.Output 'pdf-inspect-build-info.json') -Raw | ConvertFrom-Json
    if($record.BuildRequested -ne $false -or $record.OutputSha256 -ne $inspection.InspectorSha256) {throw 'Inspector provenance or explicit prebuilt binding was lost'}
}
[ordered]@{Checks=@('List without runtime','Frozen CLI forwarding','Validation-only builds','Runtime identity',
    'Skip all builds','Default builds once','Prebuilt default CLI','All ten inspection capabilities',
    'No unused inspector','Missing prebuilt tool preflight','Real case/inspector forwarding',
    'Real once-only preparation','Real family reuse','Real missing-tool refusal','Manifest Y bound forwarding','Inspector provenance');
    InspectionCapabilities=$inspectionKeys.Count;RealInspectionCalls=$inspections.Count;LibraryBuilds=0;LibraryExecutions=0;OfficeExports=0} |
    ConvertTo-Json | Set-Content (Join-Path $OutputDirectory 'evidence.json')
Write-Output 'Visual-family and real case/inspector binding checks pass; no actual build, library execution or Office export.'
