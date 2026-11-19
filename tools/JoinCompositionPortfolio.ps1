# Join already measured outcomes; copy only independent PDFs/logs/inspection summaries.
# Raster and comparison evidence stays at its original path and is reused by Compare.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)] [string[]] $SourceDirectory,
    [Parameter(Mandatory=$true)] [string] $Manifest,
    [Parameter(Mandatory=$true)] [string] $OutputDirectory,
    [ValidateSet('Both','Final','Review')] [string] $DocxPreset = 'Both'
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'CompositionOutcomes.ps1')
$output = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $output) { throw 'A fresh joined portfolio directory is required' }
$manifestBytes = [IO.File]::ReadAllBytes((Resolve-Path -LiteralPath $Manifest).Path)
$documents = ConvertFrom-Json -AsHashtable -InputObject ([Text.Encoding]::UTF8.GetString($manifestBytes).TrimStart([char]0xFEFF))
$cases = @{}
foreach ($case in @($documents)) {
    if ($case.Id -cnotmatch '^[a-z0-9]+(?:-[a-z0-9]+)*$' -or $cases.ContainsKey($case.Id)) { throw 'Invalid or duplicate manifest ID' }
    $cases[$case.Id] = $case
}
if (!$cases.Count) { throw 'An explicit nonempty manifest is required' }
$reports = @()
$seenReports = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
$jobs = [ordered]@{}
$runtime = $null
$files = @('candidate.pdf','diagnostics.json','convert.log','inspect.log',
    'inspect/font-resources.json','inspect/graphics-operations.json','inspect/text-operations.json')
foreach ($directory in $SourceDirectory) {
    $source = (Resolve-Path -LiteralPath $directory).Path
    $reportPath = Join-Path $source 'outcomes.json'
    # Repeating a source selection is harmless. Distinct reports must not claim
    # the same outcome: equal PDF bytes do not establish equal classifications.
    if (!$seenReports.Add($reportPath)) { continue }
    $reportBytes = [IO.File]::ReadAllBytes($reportPath)
    $report = ConvertFrom-Json -AsHashtable -InputObject ([Text.Encoding]::UTF8.GetString($reportBytes).TrimStart([char]0xFEFF))
    $reportHash = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($reportBytes))
    $used = $false
    foreach ($row in $report.Rows) {
        if (!$cases.ContainsKey($row.Id)) { continue }
        $case = $cases[$row.Id]
        $key = $row.Id + '-' + $row.Mode
        if ($row.Mode -cnotin (Get-PortfolioModes $case 'Both')) { throw ('Outcome mode does not match the input format: ' + $key) }
        if ($row.Mode -cnotin (Get-PortfolioModes $case $DocxPreset)) { continue }
        if ($row.Id -cne $case.Id -or $row.InputSha256 -ne $case.Sha256) { throw ('Input identity differs: ' + $row.Id) }
        $identity = @($report.LibrarySha256,$report.CliSha256,$report.InspectorSha256)
        if (@($identity | Where-Object { !$_ }).Count) { throw 'Source runtime identities are incomplete' }
        if ($null -eq $runtime) { $runtime = $identity }
        elseif (@(Compare-Object $runtime $identity -SyncWindow 0).Count) { throw 'Join outcomes from one frozen runtime and inspector' }
        $run = Join-Path $source $key
        if ($jobs.Contains($key)) {
            throw ('Conflicting duplicate outcome; select one source report: ' + $key)
        }
        $textPath = Join-Path $run 'inspect/text-operations.json'
        if (!$row.TextOperationsSha256) {
            throw ('Text inspection identity is missing; reinspect the candidate PDF: ' + $key)
        }
        if (!(Test-Path -LiteralPath $textPath -PathType Leaf)) { throw ('Candidate text inspection is missing: ' + $key) }
        $textInspection = Read-PortfolioTextOperations $textPath $key
        if ($textInspection.Hash -cne $row.TextOperationsSha256) { throw ('Candidate text inspection differs: ' + $key) }
        $expectations = Get-PortfolioContentExpectations $case $row.Mode
        $missing = @(Get-MissingPortfolioText $textInspection.Operations $expectations.Text)
        if ($row.MissingExpectedText -isnot [Array] -or
            ![Linq.Enumerable]::SequenceEqual[string]([string[]]$row.MissingExpectedText, [string[]]$missing, [StringComparer]::Ordinal) -or
            $row.ExpectedSourceEquations -ne $expectations.Equations -or $row.HeldOut -ne $case.HeldOut) {
            throw ('Content expectations differ; reinspect the candidate under the requested manifest: ' + $key)
        }
        $diagnosticsPath = Join-Path $run 'diagnostics.json'
        if (!(Test-Path -LiteralPath $diagnosticsPath -PathType Leaf)) { throw ('Candidate diagnostics are missing: ' + $key) }
        $diagnosticsHash = Get-VerifiedDiagnosticsHash $diagnosticsPath $row.Diagnostics $key
        if ($row.IgnoredEquationDiagnostics -ne (Get-PortfolioIgnoredEquationCount $row.Diagnostics)) {
            throw ('Retained equation diagnostic count differs: ' + $key)
        }
        $copies = @(foreach ($name in $files) {
            $path = Join-Path $run $name
            if (Test-Path -LiteralPath $path -PathType Leaf) {
                $hash = switch ($name) {
                    'inspect/text-operations.json' { $textInspection.Hash }
                    'diagnostics.json' { $diagnosticsHash }
                    default { (Get-FileHash -LiteralPath $path).Hash }
                }
                [pscustomobject]@{Name=$name;Path=$path;Hash=$hash}
            }
        })
        $pdf = $copies | Where-Object Name -CEQ 'candidate.pdf'
        if (!$pdf -or $pdf.Hash -cne $row.PdfSha256) { throw ('Candidate identity differs: ' + $key) }
        if (!($copies.Name -contains 'inspect.log')) { throw ('Inspection log is missing: ' + $key) }
        if ('inspect/text-operations.json' -cnotin $copies.Name -or 'diagnostics.json' -cnotin $copies.Name) {
            throw ('Candidate inspection or diagnostics disappeared: ' + $key)
        }
        $jobs[$key] = [pscustomobject]@{Row=$row;Source=$source;Copies=$copies;Report=$reportPath;ReportHash=$reportHash}
        $used = $true
    }
    if ($used) { $reports += $reportPath }
}
foreach ($case in $cases.Values) {
    foreach ($mode in (Get-PortfolioModes $case $DocxPreset)) {
        if (!$jobs.Contains($case.Id+'-'+$mode)) { throw ('Measured outcome is missing: ' + $case.Id+'/'+$mode) }
    }
}
# Finish the complete preflight before publishing any joined file. Copies have
# independent ownership; neither comparisons nor later edits can mutate sources.
New-Item -ItemType Directory -Path $output | Out-Null
$rows = @()
foreach ($key in $jobs.Keys) {
    $job = $jobs[$key]
    if ((Get-FileHash -LiteralPath $job.Report).Hash -ne $job.ReportHash) { throw 'Source outcome report changed' }
    foreach ($file in $job.Copies) {
        $target = Join-Path (Join-Path $output $key) $file.Name
        New-Item -ItemType Directory -Path (Split-Path $target) -Force | Out-Null
        Copy-Item -LiteralPath $file.Path -Destination $target
        if ((Get-FileHash -LiteralPath $target).Hash -ne $file.Hash -or
            (Get-FileHash -LiteralPath $file.Path).Hash -ne $file.Hash) { throw 'A copied outcome changed' }
    }
    $row = $job.Row.Clone()
    $row.JoinedFrom = @{Report=$job.Report;ReportSha256=$job.ReportHash;RunDirectory=(Join-Path $job.Source $key)}
    $rows += $row
}
[ordered]@{
    LibrarySha256=$runtime[0];CliSha256=$runtime[1];InspectorSha256=$runtime[2]
    ManifestSha256=[Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($manifestBytes));Rebuilt=$false;DocxPreset=$DocxPreset
    InspectionMode='SummariesOnly';NewConversions=0;SourceReports=$reports
    MeasurementScope='Joined prior measured outcomes; original classifications/provenance retained. No new conversion or content/visual pass.'
    Rows=$rows
} | ConvertTo-Json -Depth 18 | Set-Content -LiteralPath (Join-Path $output 'outcomes.json') -Encoding utf8
Write-Output "$($rows.Count) outcomes joined with independent PDFs/summaries; no build, conversion, raster or comparison replay."
