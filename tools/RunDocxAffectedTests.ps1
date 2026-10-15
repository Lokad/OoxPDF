# Qualify capability-specific tests once per method for one frozen test assembly.
[CmdletBinding(DefaultParameterSetName = 'Affected')]
param(
    [Parameter(Mandatory = $true)] [string] $RunnerPath,
    [Parameter(Mandatory = $true)] [string] $OutputDirectory,
    [Parameter(Mandatory = $true, ParameterSetName = 'Affected')]
    [ValidateNotNullOrEmpty()] [string[]] $TestFilter,
    [Parameter(Mandatory = $true, ParameterSetName = 'Capabilities')]
    [ValidateNotNullOrEmpty()] [string[]] $Capability,
    [Parameter(ParameterSetName = 'Capabilities')]
    [string] $SelectionManifest = (Join-Path $PSScriptRoot 'qualification-capabilities.json'),
    [Parameter(Mandatory = $true, ParameterSetName = 'Milestone')] [switch] $AllDocx,
    [switch] $ListOnly
)
$ErrorActionPreference = 'Stop'
$runner = (Resolve-Path -LiteralPath $RunnerPath).Path
$library = Join-Path (Split-Path $runner) 'Lokad.OoxPdf.dll'
if (!(Test-Path -LiteralPath $library)) { throw 'Runner library is missing' }
$runnerHash = (Get-FileHash -LiteralPath $runner).Hash
$libraryHash = (Get-FileHash -LiteralPath $library).Hash
foreach ($name in @('affected-tests.json', 'evidence.json', 'catalog.txt')) {
    if (Test-Path -LiteralPath (Join-Path $OutputDirectory $name)) { throw 'Output directory already contains qualification reports' }
}
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$output = (Resolve-Path -LiteralPath $OutputDirectory).Path
$temporaryDirectory = Join-Path $output 'temporary-files'
function Assert-FrozenAssembly {
    if ((Get-FileHash -LiteralPath $runner).Hash -ne $runnerHash -or
        (Get-FileHash -LiteralPath $library).Hash -ne $libraryHash) {
        throw 'The test assembly or library changed during qualification'
    }
}
function Invoke-FrozenTests([string[]] $Arguments, [string] $Log) {
    New-Item -ItemType Directory -Force -Path $temporaryDirectory | Out-Null
    $previousTmp = [Environment]::GetEnvironmentVariable('TMP')
    $previousTemp = [Environment]::GetEnvironmentVariable('TEMP')
    try {
        [Environment]::SetEnvironmentVariable('TMP', $temporaryDirectory)
        [Environment]::SetEnvironmentVariable('TEMP', $temporaryDirectory)
        dotnet $runner @Arguments *> $Log
        return $LASTEXITCODE
    }
    finally {
        [Environment]::SetEnvironmentVariable('TMP', $previousTmp)
        [Environment]::SetEnvironmentVariable('TEMP', $previousTemp)
    }
}
$catalogPath = Join-Path $output 'catalog.txt'
Assert-FrozenAssembly
dotnet $runner --list *> $catalogPath
if ($LASTEXITCODE -ne 0) { throw 'Test catalogue failed' }
Assert-FrozenAssembly
$catalog = @(foreach ($line in Get-Content -LiteralPath $catalogPath) {
    if ($line -match '^(FAST|SLOW) ([^ ]+) ([^ ]+)$') {
        [pscustomobject]@{ Group = $Matches[2]; Name = $Matches[3] }
    }
})
if (!$catalog.Count -or @($catalog.Name | Sort-Object -Unique).Count -ne $catalog.Count) {
    throw 'Test catalogue is empty or ambiguous'
}
$filters = if ($AllDocx) { @('Docx', 'Balloon') }
elseif ($Capability) {
    $manifest = Get-Content -LiteralPath $SelectionManifest -Raw | ConvertFrom-Json -AsHashtable
    @(foreach ($key in $Capability) {
        if (!$manifest.ContainsKey($key)) { throw ('Unknown capability: ' + $key) }
        if (!$manifest[$key].Count) { throw ('Empty capability: ' + $key) }
        $manifest[$key]
    }) | Select-Object -Unique
}
else { @($TestFilter | Select-Object -Unique) }
if (!$AllDocx -and @($filters | Where-Object { $_ -in @('Docx', 'Balloon', '') }).Count) {
    throw 'Use capability-specific filters, or explicit -AllDocx for a broader milestone'
}
$selected = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
foreach ($filter in $filters) {
    $matches = @($catalog | Where-Object { $_.Name.Contains($filter, [StringComparison]::OrdinalIgnoreCase) })
    if (!$matches.Count) { throw ('Empty capability selector: ' + $filter) }
    foreach ($test in $matches) { $null = $selected.Add($test.Name) }
}
$selectionPath = Join-Path $output 'selected-methods.json'
ConvertTo-Json -InputObject @($selected | Sort-Object) | Set-Content -LiteralPath $selectionPath -Encoding utf8
$inventory = [ordered]@{
    Mode = if ($AllDocx) { 'Broader milestone' } else { 'Affected capabilities' }
    TestFilters = $filters
    SelectedMethods = @($selected | Sort-Object)
    UniqueSelectedMethods = $selected.Count
    RunnerInvocations = 1
    LibrarySha256 = $libraryHash
    TestAssemblySha256 = $runnerHash
    CatalogueSha256 = (Get-FileHash -LiteralPath $catalogPath).Hash
    SelectionSha256 = (Get-FileHash -LiteralPath $selectionPath).Hash
    TemporaryDirectory = $temporaryDirectory
}
if ($Capability) {
    $inventory.Capabilities = $Capability
    $inventory.SelectionManifestSha256 = (Get-FileHash -LiteralPath $SelectionManifest).Hash
}
if ($ListOnly) { $inventory | ConvertTo-Json -Depth 5; return }
$sources = [Collections.Generic.List[object]]::new()
$rows = [Collections.Generic.List[object]]::new()
$path = Join-Path $output ('selection-{0:D3}.json' -f $sources.Count)
Assert-FrozenAssembly
$exitCode = Invoke-FrozenTests @('--test-set', $selectionPath, '--report', $path) ($path + '.log')
if ($exitCode -ne 0) { throw 'Capability qualification failed; see the selection log' }
$report = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
$names = @($report.tests.name)
if ($report.passed -ne $selected.Count -or $report.failed -ne 0 -or
    $report.skipped -ne 0 -or $names.Count -ne $selected.Count -or
    @(Compare-Object ($names | Sort-Object) ($selected | Sort-Object)).Count -or
    @($report.tests | Where-Object { $_.outcome -ne 'passed' }).Count) {
    throw ('Incomplete affected report: ' + $path)
}
Assert-FrozenAssembly
foreach ($row in $report.tests) { $rows.Add($row) }
$sources.Add([pscustomobject]@{ Selection = $selectionPath; Report = $path; Sha256 = (Get-FileHash -LiteralPath $path).Hash })

if ($rows.Count -ne $selected.Count -or @($rows.name | Sort-Object -Unique).Count -ne $selected.Count) {
    throw 'Affected reports do not cover each selected method exactly once'
}
$inventory.CompletedUtc = [DateTime]::UtcNow.ToString('o')
$inventory.UniqueExecutedMethods = $rows.Count
$inventory.SourceReports = @($sources)
$reportPath = Join-Path $output 'affected-tests.json'
[ordered]@{
    passed = $rows.Count; failed = 0; skipped = 0; tests = @($rows)
    provenance = $inventory
} | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $reportPath -Encoding utf8
$inventory.AffectedReportSha256 = (Get-FileHash -LiteralPath $reportPath).Hash
$inventory | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $output 'evidence.json') -Encoding utf8
Write-Output "$($rows.Count) unique capability checks passed in 1 runner invocation."
