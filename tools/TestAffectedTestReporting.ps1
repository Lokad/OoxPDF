# Small runner fixtures exercise reporting without compiling or replaying tests.
[CmdletBinding()]
param([string] $OutputDirectory)
$ErrorActionPreference='Stop'
if (!$OutputDirectory) {$OutputDirectory=Join-Path (Split-Path $PSScriptRoot) ('artifacts/affected-report-test-'+[guid]::NewGuid().ToString('N'))}
if(Test-Path -LiteralPath $OutputDirectory) {throw 'Use a fresh test directory'}
$root=(New-Item -ItemType Directory -Path $OutputDirectory).FullName
$runner=Join-Path $root 'Lokad.OoxPdf.Tests.dll'
[IO.File]::WriteAllText($runner,'Fake runner identity')
[IO.File]::WriteAllText((Join-Path $root 'Lokad.OoxPdf.dll'),'Fake library identity')
$global:ooxPdfAffectedReportFault=''
function dotnet {
    if($args -contains '--list') {
        'FAST fixture ReportAlpha';'FAST fixture ReportNeighbor';$global:LASTEXITCODE=0;return
    }
    $report=$args[[Array]::IndexOf([object[]]$args,'--report')+1]
    $outcome=if($global:ooxPdfAffectedReportFault -eq 'failed') {'failed'} elseif($global:ooxPdfAffectedReportFault -in @('expected-skip','unexpected-skip')) {'skipped'} else {'passed'}
    $rows=@(@{name='ReportAlpha';outcome=$outcome},@{name='ReportNeighbor';outcome='passed'})
    if($global:ooxPdfAffectedReportFault -eq 'incomplete') {$rows=@($rows[0])}
    if($global:ooxPdfAffectedReportFault -eq 'duplicate') {$rows=@($rows[0],$rows[0])}
    $passed=@($rows | Where-Object outcome -EQ 'passed').Count
    $failed=@($rows | Where-Object outcome -EQ 'failed').Count
    $skipped=@($rows | Where-Object outcome -EQ 'skipped').Count
    if($global:ooxPdfAffectedReportFault -eq 'miscount') {$passed=999}
    @{passed=$passed;failed=$failed;skipped=$skipped;tests=$rows} | ConvertTo-Json -Depth 5 | Set-Content $report -Encoding utf8
    $global:LASTEXITCODE=if($failed) {1} else {0}
}
$checks=@()
try {
    foreach($case in @('pass','expected-skip','unexpected-skip','failed','unknown-expected','incomplete','duplicate','miscount')) {
        $global:ooxPdfAffectedReportFault=$case
        $output=Join-Path $root $case
        $expected=@()
        if($case -in @('expected-skip','failed')) {$expected=@('ReportAlpha')}
        if($case -eq 'unknown-expected') {$expected=@('NotSelected')}
        $caught=$null
        try {& (Join-Path $PSScriptRoot 'RunDocxAffectedTests.ps1') -RunnerPath $runner -OutputDirectory $output -TestFilter @('ReportAlpha','ReportNeighbor') -ExpectedSkip $expected *> (Join-Path $root ($case+'.log'))}
        catch {$caught=$_.Exception.Message}
        $shouldPass=$case -in @('pass','expected-skip')
        if($shouldPass -eq ($null -ne $caught)) {throw ($case+': unexpected gate outcome '+$caught)}
        $retained=Test-Path (Join-Path $output 'affected-tests.json')
        if($retained -ne ($case -in @('pass','expected-skip','unexpected-skip','failed'))) {throw ($case+': incorrect complete-report retention')}
        if($retained) {
            $report=Get-Content (Join-Path $output 'affected-tests.json') -Raw | ConvertFrom-Json
            if($report.provenance.GatePassed -ne $shouldPass -or $report.tests.Count -ne 2 -or
                $report.failed -ne $(if($case -eq 'failed') {1} else {0}) -or
                $report.skipped -ne $(if($case -in @('expected-skip','unexpected-skip')) {1} else {0})) {throw 'Reported outcomes were reclassified'}
        }
        $checks+=@{Case=$case;GatePassed=$shouldPass;Error=$caught;ReportRetained=$retained}
    }
}
finally {Remove-Item Function:dotnet; Remove-Variable ooxPdfAffectedReportFault -Scope Global}
@{Checks=$checks;Builds=0;Conversions=0;RealTestsReplayed=0} | ConvertTo-Json -Depth 5 | Set-Content (Join-Path $root 'evidence.json') -Encoding utf8
"Affected reporting: $($checks.Count) checks pass; no builds or test replay."
