# Qualify capabilities and optional fast groups once per method on a frozen assembly.
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
    [Parameter(ParameterSetName = 'Capabilities')]
    [ValidateNotNullOrEmpty()] [string[]] $FastGroup = @(),
    [Parameter(Mandatory = $true, ParameterSetName = 'Milestone')] [switch] $AllDocx,
    [string[]] $ExpectedSkip = @(),
    [switch] $ListOnly
)
$ErrorActionPreference = 'Stop'
$runner = (Resolve-Path -LiteralPath $RunnerPath).Path
$library = Join-Path (Split-Path $runner) 'Lokad.OoxPdf.dll'
if (!(Test-Path -LiteralPath $library)) { throw 'Runner library is missing' }
$runnerHash = (Get-FileHash -LiteralPath $runner).Hash
$libraryHash = (Get-FileHash -LiteralPath $library).Hash
$cliPath = $null
$cliHash = $null
foreach ($name in @('affected-tests.json', 'evidence.json', 'catalog.txt')) {
    if (Test-Path -LiteralPath (Join-Path $OutputDirectory $name)) { throw 'Output directory already contains qualification reports' }
}
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$output = (Resolve-Path -LiteralPath $OutputDirectory).Path
$temporaryDirectory = Join-Path $output 'temporary-files'
function Assert-PhysicalDirectory([string] $Path) {
    $directory = Get-Item -LiteralPath $Path -Force
    while ($null -ne $directory) {
        if ($directory -isnot [IO.DirectoryInfo] -or ($directory.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
            throw 'Qualification temporary storage requires physical directories.'
        }
        $directory = $directory.Parent
    }
}
function Assert-FrozenAssembly {
    if ((Get-FileHash -LiteralPath $runner).Hash -ne $runnerHash -or
        (Get-FileHash -LiteralPath $library).Hash -ne $libraryHash) {
        throw 'The test assembly or library changed during qualification'
    }
    if ($cliPath -and (Get-FileHash -LiteralPath $cliPath).Hash -ne $cliHash) {
        throw 'The frozen CLI changed during qualification'
    }
}
function Invoke-FrozenTests([string[]] $Arguments, [string] $Log) {
    Assert-PhysicalDirectory $output
    # One test-set invocation owns this fresh leaf; existing files are never ours.
    New-Item -ItemType Directory -Path $temporaryDirectory | Out-Null
    $previousTmp = [Environment]::GetEnvironmentVariable('TMP')
    $previousTemp = [Environment]::GetEnvironmentVariable('TEMP')
    $previousTmpDir = [Environment]::GetEnvironmentVariable('TMPDIR')
    try {
        [Environment]::SetEnvironmentVariable('TMP', $temporaryDirectory)
        [Environment]::SetEnvironmentVariable('TEMP', $temporaryDirectory)
        [Environment]::SetEnvironmentVariable('TMPDIR', $temporaryDirectory)
        dotnet $runner @Arguments *> $Log
        return $LASTEXITCODE
    }
    finally {
        [Environment]::SetEnvironmentVariable('TMP', ($null -eq $previousTmp ? [NullString]::Value : $previousTmp))
        [Environment]::SetEnvironmentVariable('TEMP', ($null -eq $previousTemp ? [NullString]::Value : $previousTemp))
        [Environment]::SetEnvironmentVariable('TMPDIR', ($null -eq $previousTmpDir ? [NullString]::Value : $previousTmpDir))
    }
}
$catalogPath = Join-Path $output 'catalog.txt'
Assert-FrozenAssembly
dotnet $runner --list *> $catalogPath
if ($LASTEXITCODE -ne 0) { throw 'Test catalogue failed' }
Assert-FrozenAssembly
$catalog = @(foreach ($line in Get-Content -LiteralPath $catalogPath) {
    if ($line -match '^(FAST|SLOW) ([^ ]+) ([^ ]+)$') {
        [pscustomobject]@{ Slow = $Matches[1] -eq 'SLOW'; Group = $Matches[2]; Name = $Matches[3] }
    }
})
if (!$catalog.Count -or @($catalog.Name | Sort-Object -Unique).Count -ne $catalog.Count) {
    throw 'Test catalogue is empty or ambiguous'
}
$filters = if ($AllDocx) { @('Docx', 'Balloon', 'PortableRendering') }
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
if ($AllDocx) {
    # Retain the name selectors for neighboring tests in shared groups.
    $docxMembers = @($catalog | Where-Object { $_.Group.StartsWith('docx-', [StringComparison]::Ordinal) })
    foreach ($test in $docxMembers) { $null = $selected.Add($test.Name) }
}
$exactNames = [Collections.Generic.HashSet[string]]::new([string[]]$catalog.Name, [StringComparer]::Ordinal)
# Search the catalogue in native string code instead of a PowerShell loop over
# every selector/test pair. Newlines delimit names and cannot occur in a name.
$catalogNames = [string]::Join("`n", [string[]]$catalog.Name) + "`n"
foreach ($filter in $filters) {
    if ($filter.StartsWith('=', [StringComparison]::Ordinal)) {
        $name = $filter.Substring(1)
        if (!$exactNames.Contains($name)) { throw ('Empty capability selector: ' + $filter) }
        $null = $selected.Add($name)
        continue
    }
    $found = $false
    $offset = 0
    while (!$filter.Contains("`n") -and
        ($match = $catalogNames.IndexOf($filter, $offset, [StringComparison]::OrdinalIgnoreCase)) -ge 0) {
        $first = $catalogNames.LastIndexOf([char]10, $match) + 1
        $last = $catalogNames.IndexOf([char]10, $match)
        $null = $selected.Add($catalogNames.Substring($first, $last - $first))
        $found = $true
        $offset = $last + 1
    }
    if (!$found) { throw ('Empty capability selector: ' + $filter) }
}
# Capabilities retain their strict skip contract. Added fast groups preserve
# ordinary --group/--skip-slow behavior; their environment skips stay reported.
$skipValidated = $selected
if ($FastGroup.Count) {
    $skipValidated = [Collections.Generic.HashSet[string]]::new($selected, [StringComparer]::Ordinal)
    foreach ($group in $FastGroup) {
        $members = @($catalog | Where-Object { !$_.Slow -and $_.Group -ceq $group })
        if (!$members.Count) { throw ('Unknown or empty fast group: ' + $group) }
        foreach ($test in $members) { $null = $selected.Add($test.Name) }
    }
}
$selectionPath = Join-Path $output 'selected-methods.json'
if (@($ExpectedSkip | Where-Object { !$selected.Contains($_) }).Count) { throw 'Expected skips must name selected exact methods' }
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
    ExpectedSkips = @($ExpectedSkip | Sort-Object -Unique)
}
if ($AllDocx) {
    $inventory.DocxGroups = @($docxMembers.Group | Sort-Object -Unique)
}
if ($Capability) {
    $inventory.Capabilities = $Capability
    $inventory.SelectionManifestSha256 = (Get-FileHash -LiteralPath $SelectionManifest).Hash
}
if ($FastGroup.Count) {
    $inventory.FastGroups = @($FastGroup | Sort-Object -Unique -CaseSensitive)
    $inventory.SkipValidationMethods = @($skipValidated | Sort-Object)
}
if ($ListOnly) { $inventory | ConvertTo-Json -Depth 5; return }
if (@($catalog | Where-Object { $_.Group -eq 'cli' -and $selected.Contains($_.Name) }).Count) {
    if ([string]::IsNullOrWhiteSpace($env:OOXPDF_TEST_CLI_ASSEMBLY)) {
        throw 'CLI tests require OOXPDF_TEST_CLI_ASSEMBLY bound to the matching frozen CLI; implicit CLI builds are not qualification.'
    }
    $cliPath = (Resolve-Path -LiteralPath $env:OOXPDF_TEST_CLI_ASSEMBLY).Path
    $cliLibrary = Join-Path (Split-Path -Parent $cliPath) 'Lokad.OoxPdf.dll'
    if (!(Test-Path -LiteralPath $cliLibrary -PathType Leaf) -or
        (Get-FileHash -LiteralPath $cliLibrary).Hash -ne $libraryHash) {
        throw 'The frozen CLI and test runner must use the same qualified library.'
    }
    $cliHash = (Get-FileHash -LiteralPath $cliPath).Hash
    $inventory.CliPath = $cliPath
    $inventory.CliSha256 = $cliHash
}
$path = Join-Path $output 'selection-000.json'
Assert-FrozenAssembly
$exitCode = Invoke-FrozenTests @('--test-set', $selectionPath, '--report', $path) ($path + '.log')
$report = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
$rows = @($report.tests)
$seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
$counts = @{ passed = 0; failed = 0; skipped = 0 }
foreach ($row in $rows) {
    if ($row.name -isnot [string] -or !$selected.Contains($row.name) -or !$seen.Add($row.name) -or
        $row.outcome -isnot [string] -or !$counts.ContainsKey($row.outcome)) {
        throw ('Incomplete affected report: ' + $path)
    }
    $counts[$row.outcome]++
}
if ($seen.Count -ne $selected.Count -or $counts.passed -ne $report.passed -or
    $counts.failed -ne $report.failed -or $counts.skipped -ne $report.skipped) {
    throw ('Incomplete affected report: ' + $path)
}
Assert-FrozenAssembly

$inventory.CompletedUtc = [DateTime]::UtcNow.ToString('o')
$inventory.UniqueExecutedMethods = $rows.Count
$inventory.SourceReports = @([pscustomobject]@{ Selection = $selectionPath; Report = $path; Sha256 = (Get-FileHash -LiteralPath $path).Hash })
$inventory.RunnerExitCode = $exitCode
$inventory.UnexpectedSkips = @($rows | Where-Object {
    $_.outcome -eq 'skipped' -and $skipValidated.Contains($_.name) -and $_.name -cnotin $ExpectedSkip
} | ForEach-Object name)
if ($FastGroup.Count) {
    $inventory.GroupOnlySkips = @($rows | Where-Object {
        $_.outcome -eq 'skipped' -and !$skipValidated.Contains($_.name)
    } | ForEach-Object name)
}
$inventory.GatePassed = $exitCode -eq 0 -and $report.failed -eq 0 -and !$inventory.UnexpectedSkips.Count
$inventory.TemporaryFilesRetained = Test-Path -LiteralPath $temporaryDirectory
if ($inventory.GatePassed -and $inventory.TemporaryFilesRetained) {
    try {
        $owned = [IO.Path]::GetFullPath($temporaryDirectory)
        if ([IO.Path]::GetDirectoryName($owned) -cne $output) { throw 'Temporary storage escaped its declared output directory.' }
        Assert-PhysicalDirectory $owned
        # The console runner owns test files and may deliberately retain them.
        [IO.Directory]::Delete($owned, $false)
        $inventory.TemporaryFilesRetained = $false
    }
    catch {
        # Cleanup does not reclassify semantic results or discard their reports.
        $inventory.TemporaryCleanupError = $_.Exception.Message
        Write-Warning ('Passing test temporary files were retained: ' + $_.Exception.Message)
    }
}
$reportPath = Join-Path $output 'affected-tests.json'
[ordered]@{
    passed = $report.passed; failed = $report.failed; skipped = $report.skipped; tests = @($rows)
    provenance = $inventory
} | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $reportPath -Encoding utf8
$inventory.AffectedReportSha256 = (Get-FileHash -LiteralPath $reportPath).Hash
$inventory | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $output 'evidence.json') -Encoding utf8
if (!$inventory.GatePassed) { throw 'Capability qualification failed or had unexpected skips; see retained affected-tests.json and selection log' }
Write-Output "$($rows.Count) unique qualification checks: $($report.passed) pass / $($report.skipped) skip in 1 runner invocation."
