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
$rawReportName = 'selection-000.json'
foreach ($name in @('affected-tests.json', 'evidence.json', 'catalog.txt', $rawReportName)) {
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
function Read-QualificationJsonSource([string] $Path) {
    $bytes = [IO.File]::ReadAllBytes($Path)
    $reader = [IO.StreamReader]::new([IO.MemoryStream]::new($bytes))
    try {
        [pscustomobject]@{
            Text = $reader.ReadToEnd()
            Hash = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes))
        }
    }
    finally { $reader.Dispose() }
}
function Assert-SelectionManifest {
    if ($Capability -and (Get-FileHash -LiteralPath $SelectionManifest).Hash -cne $manifestSource.Hash) {
        throw 'The capability selection manifest changed during qualification'
    }
}
function Invoke-FrozenTests([string[]] $Arguments, [string] $Log) {
    Assert-PhysicalDirectory $output
    # One test-set invocation owns this fresh leaf; existing files are never ours.
    New-Item -ItemType Directory -Path $temporaryDirectory | Out-Null
    $previousTmp = [Environment]::GetEnvironmentVariable('TMP')
    $previousTemp = [Environment]::GetEnvironmentVariable('TEMP')
    $previousTmpDir = [Environment]::GetEnvironmentVariable('TMPDIR')
    $previousToolBinding = [Environment]::GetEnvironmentVariable('OOXPDF_TEST_REQUIRE_BOUND_TOOLS')
    try {
        [Environment]::SetEnvironmentVariable('TMP', $temporaryDirectory)
        [Environment]::SetEnvironmentVariable('TEMP', $temporaryDirectory)
        [Environment]::SetEnvironmentVariable('TMPDIR', $temporaryDirectory)
        [Environment]::SetEnvironmentVariable('OOXPDF_TEST_REQUIRE_BOUND_TOOLS', '1')
        dotnet $runner @Arguments *> $Log
        return $LASTEXITCODE
    }
    finally {
        [Environment]::SetEnvironmentVariable('TMP', ($null -eq $previousTmp ? [NullString]::Value : $previousTmp))
        [Environment]::SetEnvironmentVariable('TEMP', ($null -eq $previousTemp ? [NullString]::Value : $previousTemp))
        [Environment]::SetEnvironmentVariable('TMPDIR', ($null -eq $previousTmpDir ? [NullString]::Value : $previousTmpDir))
        [Environment]::SetEnvironmentVariable('OOXPDF_TEST_REQUIRE_BOUND_TOOLS', ($null -eq $previousToolBinding ? [NullString]::Value : $previousToolBinding))
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
$filters = if ($AllDocx) { @('Docx', 'Balloon') }
elseif ($Capability) {
    $SelectionManifest = (Resolve-Path -LiteralPath $SelectionManifest).Path
    $manifestSource = Read-QualificationJsonSource $SelectionManifest
    $manifest = ConvertFrom-Json -InputObject $manifestSource.Text -AsHashtable
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
    # DOCX-specific portable tests carry their format in the registered group.
    # Retain the legacy selectors for neighboring tests in shared groups.
    $docxMembers = @($catalog | Where-Object { $_.Group.StartsWith('docx-', [StringComparison]::Ordinal) })
    foreach ($test in $docxMembers) { $null = $selected.Add($test.Name) }
}
$exactNames = [Collections.Generic.HashSet[string]]::new([string[]]$catalog.Name, [StringComparer]::Ordinal)
$methodsByFilter = [Collections.Generic.Dictionary[string,string[]]]::new([StringComparer]::Ordinal)
# Search the catalogue in native string code instead of a PowerShell loop over
# every selector/test pair. Newlines delimit names and cannot occur in a name.
$catalogNames = [string]::Join("`n", [string[]]$catalog.Name) + "`n"
foreach ($filter in $filters) {
    if ($filter.StartsWith('=', [StringComparison]::Ordinal)) {
        $name = $filter.Substring(1)
        if (!$exactNames.Contains($name)) { throw ('Empty capability selector: ' + $filter) }
        $null = $selected.Add($name)
        $methodsByFilter[$filter] = @($name)
        continue
    }
    $methodNames = [Collections.Generic.List[string]]::new()
    $offset = 0
    while (!$filter.Contains("`n") -and
        ($match = $catalogNames.IndexOf($filter, $offset, [StringComparison]::OrdinalIgnoreCase)) -ge 0) {
        $first = $catalogNames.LastIndexOf([char]10, $match) + 1
        $last = $catalogNames.IndexOf([char]10, $match)
        $name = $catalogNames.Substring($first, $last - $first)
        $null = $selected.Add($name)
        $methodNames.Add($name)
        $offset = $last + 1
    }
    if (!$methodNames.Count) { throw ('Empty capability selector: ' + $filter) }
    $methodsByFilter[$filter] = $methodNames.ToArray()
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
    RunnerInvocations = 0
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
    $inventory.SelectionManifestSha256 = $manifestSource.Hash
}
if ($FastGroup.Count) {
    $inventory.FastGroups = @($FastGroup | Sort-Object -Unique -CaseSensitive)
    $inventory.SkipValidationMethods = @($skipValidated | Sort-Object)
}
if ($ListOnly) {
    Assert-SelectionManifest
    $inventory | ConvertTo-Json -Depth 5
    return
}
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
$path = Join-Path $output $rawReportName
Assert-FrozenAssembly
$exitCode = Invoke-FrozenTests @('--test-set', $selectionPath, '--report', $path) ($path + '.log')
$inventory.RunnerInvocations++
if (!(Test-Path -LiteralPath $path -PathType Leaf)) { throw 'Test runner did not produce its report' }
$reportSource = Read-QualificationJsonSource $path
$report = ConvertFrom-Json -InputObject $reportSource.Text
$rows = @($report.tests)
$counts = @{ passed = 0; failed = 0; skipped = 0 }
$outcomes = [Collections.Generic.Dictionary[string,string]]::new([StringComparer]::Ordinal)
$primaryPassed = 0
foreach ($row in $rows) {
    if ($row.name -isnot [string] -or !$selected.Contains($row.name) -or
        $row.outcome -isnot [string] -or !$counts.ContainsKey($row.outcome) -or !$outcomes.TryAdd($row.name, $row.outcome)) {
        throw ('Incomplete affected report: ' + $path)
    }
    $counts[$row.outcome]++
    if ($row.outcome -eq 'passed' -and $skipValidated.Contains($row.name)) { $primaryPassed++ }
}
if ($outcomes.Count -ne $selected.Count -or $counts.passed -ne $report.passed -or
    $counts.failed -ne $report.failed -or $counts.skipped -ne $report.skipped) {
    throw ('Incomplete affected report: ' + $path)
}
Assert-FrozenAssembly

if ($null -ne $report.validationTools) {
    $validationTools = [ordered]@{}
    foreach ($entry in $report.validationTools.PSObject.Properties) {
        $bound = [Environment]::GetEnvironmentVariable($entry.Name)
        $tool = $entry.Value
        if (!$bound -or $tool.Path -isnot [string] -or $tool.Sha256 -isnot [string] -or
            $tool.Sha256 -cnotmatch '^[0-9A-F]{64}$' -or
            [IO.Path]::GetFullPath($bound) -cne $tool.Path) {
            throw ('Validation tool report differs from the supplied binding: ' + $entry.Name)
        }
        if (!(Test-Path -LiteralPath $tool.Path -PathType Leaf) -or (Get-FileHash -LiteralPath $tool.Path).Hash -cne $tool.Sha256) {
            throw ('Validation tool changed during qualification: ' + $entry.Name)
        }
        $validationTools[$entry.Name] = $tool
    }
    $inventory.ValidationTools = $validationTools
}

$inventory.CompletedUtc = [DateTime]::UtcNow.ToString('o')
# Keep the legacy row count, including skipped methods, alongside actual passes.
$inventory.UniqueExecutedMethods = $rows.Count
$inventory.UniquePassedMethods = $report.passed
$inventory.UniquePassedPrimaryMethods = $primaryPassed
$inventory.SourceReports = @([pscustomobject]@{ Selection = $selectionPath; Report = $path; Sha256 = $reportSource.Hash })
$inventory.RunnerExitCode = $exitCode
$inventory.UnexpectedSkips = @($rows | Where-Object {
    $_.outcome -eq 'skipped' -and $skipValidated.Contains($_.name) -and $_.name -cnotin $ExpectedSkip
} | ForEach-Object name)
if ($FastGroup.Count) {
    $inventory.GroupOnlySkips = @($rows | Where-Object {
        $_.outcome -eq 'skipped' -and !$skipValidated.Contains($_.name)
    } | ForEach-Object name)
}
$runnerAccepted = $exitCode -eq 0 -and $report.failed -eq 0 -and !$inventory.UnexpectedSkips.Count
$inventory.GatePassed = $runnerAccepted -and $primaryPassed -gt 0
if ($Capability) {
    $unexpected = [Collections.Generic.HashSet[string]]::new([string[]]$inventory.UnexpectedSkips, [StringComparer]::Ordinal)
    $inventory.CapabilityCoverage = @(foreach ($key in ($Capability | Select-Object -Unique)) {
        $methods = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
        foreach ($filter in $manifest[$key]) { $methods.UnionWith($methodsByFilter[$filter]) }
        $capabilityCounts = @{ passed = 0; failed = 0; skipped = 0 }
        $hasUnexpectedSkip = $methods.Overlaps($unexpected)
        foreach ($name in $methods) { $capabilityCounts[$outcomes[$name]]++ }
        [ordered]@{
            Name = $key
            Selected = $methods.Count
            Passed = $capabilityCounts.passed
            Failed = $capabilityCounts.failed
            Skipped = $capabilityCounts.skipped
            Qualified = $exitCode -eq 0 -and $capabilityCounts.passed -gt 0 -and
                !$capabilityCounts.failed -and !$hasUnexpectedSkip
        }
    })
    $fullySkipped = @($inventory.CapabilityCoverage | Where-Object { $_.Skipped -eq $_.Selected } | ForEach-Object Name)
    if ($fullySkipped.Count) { Write-Warning ('All selected methods skipped for: ' + ($fullySkipped -join ', ')) }
}
Assert-SelectionManifest
if ((Get-FileHash -LiteralPath $path).Hash -cne $reportSource.Hash) {
    throw 'The raw test report changed during qualification'
}
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
if (!$runnerAccepted) { throw 'Capability qualification failed or had unexpected skips; see retained affected-tests.json and selection log' }
if (!$report.passed) { throw 'No selected qualification test passed. Expected skips are retained in affected-tests.json.' }
if (!$inventory.GatePassed) { throw 'No primary qualification test passed. Fast-group-only passes and expected skips are retained in affected-tests.json.' }
Write-Output "$($rows.Count) unique qualification checks: $($report.passed) pass / $($report.skipped) skip in 1 runner invocation."
