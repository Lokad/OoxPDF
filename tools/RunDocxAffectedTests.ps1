# Qualify DOCX and balloon tests once per method for one frozen test assembly.
param(
    [Parameter(Mandatory = $true)] [string] $RunnerPath,
    [Parameter(Mandatory = $true)] [string] $OutputDirectory,
    [switch] $ListOnly
)
$ErrorActionPreference = 'Stop'
$runner = (Resolve-Path -LiteralPath $RunnerPath).Path
$library = Join-Path (Split-Path $runner) 'Lokad.OoxPdf.dll'
if (!(Test-Path -LiteralPath $library)) { throw 'Runner library is missing' }
$runnerHash = (Get-FileHash -LiteralPath $runner).Hash
$libraryHash = (Get-FileHash -LiteralPath $library).Hash
foreach ($name in @('docx-tests.json', 'balloon-tests.json', 'evidence.json')) {
    if (Test-Path -LiteralPath (Join-Path $OutputDirectory $name)) { throw 'Output directory already contains qualification reports' }
}
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$output = (Resolve-Path -LiteralPath $OutputDirectory).Path
$temporaryDirectory = Join-Path $output 'temporary-files'
# Crowded Windows temp directories can make fixture path creation dominate tests.
# Restrict this environment change to each runner invocation and its children.
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
dotnet $runner --list *> $catalogPath
if ($LASTEXITCODE -ne 0) { throw 'Test catalogue failed' }
$catalog = @(foreach ($line in Get-Content -LiteralPath $catalogPath) {
    if ($line -match '^(FAST|SLOW) ([^ ]+) ([^ ]+)$') {
        [pscustomobject]@{ Group = $Matches[2]; Name = $Matches[3] }
    }
})
if (!$catalog.Count -or @($catalog.Name | Sort-Object -Unique).Count -ne $catalog.Count) {
    throw 'Test catalogue is empty or ambiguous'
}
$docx = @($catalog | Where-Object { $_.Name.Contains('Docx', [StringComparison]::OrdinalIgnoreCase) })
$balloon = @($catalog | Where-Object { $_.Name.Contains('Balloon', [StringComparison]::OrdinalIgnoreCase) })
$missing = @($balloon | Where-Object { $_.Name -notin $docx.Name })
if (!$docx.Count -or !$balloon.Count) { throw 'Affected catalogue is empty' }
$inventory = [ordered]@{
    DocxMethods = $docx.Count
    BalloonMethods = $balloon.Count
    BalloonMethodsCoveredByDocx = $balloon.Count - $missing.Count
    AdditionalMethods = @($missing.Name)
    LibrarySha256 = $libraryHash
    TestAssemblySha256 = $runnerHash
    TemporaryDirectory = $temporaryDirectory
}
if ($ListOnly) { $inventory | ConvertTo-Json -Depth 5; return }

function Assert-FrozenAssembly {
    if ((Get-FileHash -LiteralPath $runner).Hash -ne $runnerHash -or
        (Get-FileHash -LiteralPath $library).Hash -ne $libraryHash) {
        throw 'The test assembly or library changed during qualification'
    }
}
function Read-CompleteReport([string] $Path, [string[]] $ExpectedNames) {
    $report = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
    $names = @($report.tests.name)
    if ($report.passed -ne $ExpectedNames.Count -or $report.failed -ne 0 -or
        $report.skipped -ne 0 -or $names.Count -ne $ExpectedNames.Count -or
        @(Compare-Object ($names | Sort-Object) ($ExpectedNames | Sort-Object)).Count -or
        @($report.tests | Where-Object { $_.outcome -ne 'passed' }).Count) {
        throw ('Incomplete affected report: ' + $Path)
    }
    Assert-FrozenAssembly
    return $report
}
$docxPath = Join-Path $output 'docx-tests.json'
$exitCode = Invoke-FrozenTests @('--test', 'Docx', '--report', $docxPath) (Join-Path $output 'docx-tests.log')
if ($exitCode -ne 0) { throw 'DOCX qualification failed' }
$docxReport = Read-CompleteReport $docxPath @($docx.Name)
$sources = [Collections.Generic.List[object]]::new()
$sources.Add([pscustomobject]@{ Report = $docxPath; Sha256 = (Get-FileHash $docxPath).Hash })
$balloonRows = [Collections.Generic.List[object]]::new()
foreach ($row in @($docxReport.tests | Where-Object { $_.name -in $balloon.Name })) {
    $balloonRows.Add($row)
}
foreach ($test in $missing) {
    # The console runner supports substring filters. Refuse ambiguous selectors.
    if (@($catalog | Where-Object { $_.Name.Contains($test.Name, [StringComparison]::OrdinalIgnoreCase) }).Count -ne 1) {
        throw ('Ambiguous additional test selector: ' + $test.Name)
    }
    $path = Join-Path $output ($test.Name + '.json')
    $exitCode = Invoke-FrozenTests @('--test', $test.Name, '--report', $path) (Join-Path $output ($test.Name + '.log'))
    if ($exitCode -ne 0) { throw ('Additional qualification failed: ' + $test.Name) }
    $report = Read-CompleteReport $path @($test.Name)
    $balloonRows.Add($report.tests[0])
    $sources.Add([pscustomobject]@{ Report = $path; Sha256 = (Get-FileHash $path).Hash })
}
Assert-FrozenAssembly
if ($balloonRows.Count -ne $balloon.Count -or
    @($balloonRows.name | Sort-Object -Unique).Count -ne $balloon.Count) {
    throw 'Balloon partition does not cover its catalogue'
}
[ordered]@{
    passed = $balloonRows.Count; failed = 0; skipped = 0; tests = @($balloonRows)
    provenance = [ordered]@{
        Method = 'Selected completed DOCX checks plus freshly executed additional methods'
        LibrarySha256 = $libraryHash; TestAssemblySha256 = $runnerHash
        SourceReports = @($sources)
    }
} | ConvertTo-Json -Depth 10 | Set-Content (Join-Path $output 'balloon-tests.json') -Encoding utf8
$inventory.CompletedUtc = [DateTime]::UtcNow.ToString('o')
$inventory.UniqueExecutedMethods = $docx.Count + $missing.Count
$inventory.SourceReports = @($sources)
$inventory | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $output 'evidence.json') -Encoding utf8
Write-Output "$($docx.Count) DOCX and $($balloon.Count) balloon checks passed; $($docx.Count + $missing.Count) unique methods executed."
