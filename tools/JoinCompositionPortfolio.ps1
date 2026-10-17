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
$output = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $output) { throw 'A fresh joined portfolio directory is required' }
$cases = @{}
foreach ($case in @(Get-Content -LiteralPath $Manifest -Raw | ConvertFrom-Json)) {
    if ($case.Id -cnotmatch '^[a-z0-9]+(?:-[a-z0-9]+)*$' -or $cases.ContainsKey($case.Id)) { throw 'Invalid or duplicate manifest ID' }
    $cases[$case.Id] = $case
}
if (!$cases.Count) { throw 'An explicit nonempty manifest is required' }
$reports = @()
$jobs = [ordered]@{}
$runtime = $null
$files = @('candidate.pdf','diagnostics.json','convert.log','inspect.log',
    'inspect/font-resources.json','inspect/graphics-operations.json','inspect/text-operations.json')
foreach ($directory in $SourceDirectory) {
    $source = (Resolve-Path -LiteralPath $directory).Path
    $reportPath = Join-Path $source 'outcomes.json'
    $report = Get-Content -LiteralPath $reportPath -Raw | ConvertFrom-Json -AsHashtable
    $used = $false
    foreach ($row in $report.Rows) {
        if (!$cases.ContainsKey($row.Id)) { continue }
        if ($row.Mode -notin @('final-preserve','all-word-compatible','presentation')) { throw 'Unknown outcome mode' }
        if ($row.Mode -ne 'presentation' -and $DocxPreset -ne 'Both' -and
            $row.Mode -ne $(if ($DocxPreset -eq 'Final') {'final-preserve'} else {'all-word-compatible'})) { continue }
        if ($row.InputSha256 -ne $cases[$row.Id].Sha256) { throw ('Input identity differs: ' + $row.Id) }
        $identity = @($report.LibrarySha256,$report.CliSha256,$report.InspectorSha256)
        if (@($identity | Where-Object { !$_ }).Count) { throw 'Source runtime identities are incomplete' }
        if ($null -eq $runtime) { $runtime = $identity }
        elseif (@(Compare-Object $runtime $identity -SyncWindow 0).Count) { throw 'Join outcomes from one frozen runtime and inspector' }
        $key = $row.Id + '-' + $row.Mode
        $run = Join-Path $source $key
        if ((Get-FileHash -LiteralPath (Join-Path $run 'candidate.pdf')).Hash -ne $row.PdfSha256) { throw ('Candidate identity differs: ' + $key) }
        if ($jobs.Contains($key)) {
            if ($jobs[$key].Row.PdfSha256 -ne $row.PdfSha256) { throw ('Conflicting duplicate outcome: ' + $key) }
            continue
        }
        $copies = @()
        foreach ($name in $files) {
            $path = Join-Path $run $name
            if (Test-Path -LiteralPath $path -PathType Leaf) {
                $copies += [pscustomobject]@{Name=$name;Path=$path;Hash=(Get-FileHash -LiteralPath $path).Hash}
            }
        }
        if (!($copies.Name -contains 'inspect.log')) { throw ('Inspection log is missing: ' + $key) }
        $jobs[$key] = [pscustomobject]@{Row=$row;Source=$source;Copies=$copies;Report=$reportPath;ReportHash=(Get-FileHash -LiteralPath $reportPath).Hash}
        $used = $true
    }
    if ($used) { $reports += $reportPath }
}
foreach ($case in $cases.Values) {
    $modes = if ($case.Input.EndsWith('.docx',[StringComparison]::OrdinalIgnoreCase)) {
        if ($DocxPreset -eq 'Both') { @('final-preserve','all-word-compatible') }
        elseif ($DocxPreset -eq 'Final') { @('final-preserve') } else { @('all-word-compatible') }
    } else { @('presentation') }
    foreach ($mode in $modes) {
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
    ManifestSha256=(Get-FileHash -LiteralPath $Manifest).Hash;Rebuilt=$false;DocxPreset=$DocxPreset
    InspectionMode='SummariesOnly';NewConversions=0;SourceReports=$reports
    MeasurementScope='Joined prior measured outcomes; original classifications/provenance retained. No new conversion or content/visual pass.'
    Rows=$rows
} | ConvertTo-Json -Depth 18 | Set-Content -LiteralPath (Join-Path $output 'outcomes.json') -Encoding utf8
Write-Output "$($rows.Count) outcomes joined with independent PDFs/summaries; no build, conversion, raster or comparison replay."
