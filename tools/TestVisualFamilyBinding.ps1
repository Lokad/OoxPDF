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
param($Project, $OutputDll, $Description)
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
[ordered]@{Checks=@('List without runtime','Frozen CLI forwarding','Validation-only builds','Runtime identity',
    'Skip all builds','Default builds once','Prebuilt default CLI');LibraryBuilds=0;LibraryExecutions=0} |
    ConvertTo-Json | Set-Content (Join-Path $OutputDirectory 'evidence.json')
Write-Output 'Seven visual-family binding checks passed; no real build, library execution or Office export.'
