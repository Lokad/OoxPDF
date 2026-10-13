# Qualify capability-specific tests once per method for one frozen test assembly.
[CmdletBinding(DefaultParameterSetName = 'Affected')]
param(
    [Parameter(Mandatory = $true)] [string] $RunnerPath,
    [Parameter(Mandatory = $true)] [string] $OutputDirectory,
    [Parameter(Mandatory = $true, ParameterSetName = 'Affected')]
    [ValidateNotNullOrEmpty()] [string[]] $TestFilter,
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
$filters = if ($AllDocx) { @('Docx', 'Balloon') } else { @($TestFilter | Select-Object -Unique) }
if (!$AllDocx -and @($filters | Where-Object { $_ -in @('Docx', 'Balloon', '') }).Count) {
    throw 'Use capability-specific filters, or explicit -AllDocx for a broader milestone'
}
$selected = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
$invocations = [Collections.Generic.List[object]]::new()
foreach ($filter in $filters) {
    $matches = @($catalog | Where-Object { $_.Name.Contains($filter, [StringComparison]::OrdinalIgnoreCase) })
    if (!$matches.Count) { throw ('Empty capability selector: ' + $filter) }
    $new = @($matches | Where-Object { !$selected.Contains($_.Name) })
    if ($new.Count -eq $matches.Count) {
        $invocations.Add([pscustomobject]@{ Filter = $filter; Names = @($new.Name) })
    }
    elseif ($new.Count) {
        # Partially overlapping substring filters need exact, unambiguous calls.
        foreach ($test in $new) {
            if (@($catalog | Where-Object { $_.Name.Contains($test.Name, [StringComparison]::OrdinalIgnoreCase) }).Count -ne 1) {
                throw ('Ambiguous additional test selector: ' + $test.Name)
            }
            $invocations.Add([pscustomobject]@{ Filter = $test.Name; Names = @($test.Name) })
        }
    }
    foreach ($test in $new) { $null = $selected.Add($test.Name) }
}
$inventory = [ordered]@{
    Mode = if ($AllDocx) { 'Broader milestone' } else { 'Affected capabilities' }
    TestFilters = $filters
    SelectedMethods = @($selected | Sort-Object)
    UniqueSelectedMethods = $selected.Count
    RunnerInvocations = $invocations.Count
    LibrarySha256 = $libraryHash
    TestAssemblySha256 = $runnerHash
    CatalogueSha256 = (Get-FileHash -LiteralPath $catalogPath).Hash
    TemporaryDirectory = $temporaryDirectory
}
if ($ListOnly) { $inventory | ConvertTo-Json -Depth 5; return }
$sources = [Collections.Generic.List[object]]::new()
$rows = [Collections.Generic.List[object]]::new()
foreach ($invocation in $invocations) {
    $path = Join-Path $output ('selection-{0:D3}.json' -f $sources.Count)
    Assert-FrozenAssembly
    $exitCode = Invoke-FrozenTests @('--test', $invocation.Filter, '--report', $path) ($path + '.log')
    if ($exitCode -ne 0) { throw ('Capability qualification failed: ' + $invocation.Filter) }
    $report = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
    $names = @($report.tests.name)
    if ($report.passed -ne $invocation.Names.Count -or $report.failed -ne 0 -or
        $report.skipped -ne 0 -or $names.Count -ne $invocation.Names.Count -or
        @(Compare-Object ($names | Sort-Object) ($invocation.Names | Sort-Object)).Count -or
        @($report.tests | Where-Object { $_.outcome -ne 'passed' }).Count) {
        throw ('Incomplete affected report: ' + $path)
    }
    Assert-FrozenAssembly
    foreach ($row in $report.tests) { $rows.Add($row) }
    $sources.Add([pscustomobject]@{ Filter = $invocation.Filter; Report = $path; Sha256 = (Get-FileHash -LiteralPath $path).Hash })
}
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
Write-Output "$($rows.Count) unique capability checks passed in $($invocations.Count) runner invocations."
