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
$global:ooxPdfAffectedReportCalls=0
function dotnet {
    if($args -contains '--list') {
        'FAST fixture ReportAlpha';'FAST fixture ReportNeighbor';'FAST auxiliary ReportGroupOnly'
        'SLOW auxiliary ReportGroupSlow';'SLOW fixture ReportExplicitSlow';'SLOW only-slow ReportSlowOnly'
        'FAST selection SelectionAlpha';'FAST selection SelectionAlphaNeighbor';'FAST selection Outer.SelectionAlpha'
        'FAST selection SelectionOther';'FAST selection Selection[Literal]';'FAST selection SelectionL'
        'FAST selection item';'FAST selection İtemNeighbor'
        $global:LASTEXITCODE=0;return
    }
    $global:ooxPdfAffectedReportCalls++
    $report=$args[[Array]::IndexOf([object[]]$args,'--report')+1]
    if([IO.Path]::GetFileName($env:TMPDIR) -ne 'temporary-files') {throw 'Runner temp path escaped its output directory'}
    $selection=$args[[Array]::IndexOf([object[]]$args,'--test-set')+1]
    $methods=@(Get-Content -LiteralPath $selection -Raw | ConvertFrom-Json)
    $rows=@(foreach($method in $methods) {
        $outcome=if($method -eq 'ReportAlpha' -and $global:ooxPdfAffectedReportFault -eq 'failed') {'failed'}
            elseif(($method -eq 'ReportAlpha' -and $global:ooxPdfAffectedReportFault -in @('expected-skip','unexpected-skip','capability-skip','group-and-capability-skip')) -or
                ($method -eq 'ReportGroupOnly' -and $global:ooxPdfAffectedReportFault -in @('group-skip','group-and-capability-skip'))) {'skipped'}
            else {'passed'}
        @{name=$method;outcome=$outcome}
    })
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
        $global:ooxPdfAffectedReportCalls=0
        $output=Join-Path $root $case
        $expected=@()
        if($case -in @('expected-skip','failed')) {$expected=@('ReportAlpha')}
        if($case -eq 'unknown-expected') {$expected=@('NotSelected')}
        $caught=$null
        $beforeTmpDir=[Environment]::GetEnvironmentVariable('TMPDIR')
        try {& (Join-Path $PSScriptRoot 'RunDocxAffectedTests.ps1') -RunnerPath $runner -OutputDirectory $output -TestFilter @('ReportAlpha','ReportNeighbor') -ExpectedSkip $expected *> (Join-Path $root ($case+'.log'))}
        catch {$caught=$_.Exception.Message}
        if([Environment]::GetEnvironmentVariable('TMPDIR') -cne $beforeTmpDir) {throw 'Runner did not restore TMPDIR'}
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

    $manifest=Join-Path $root 'union-capabilities.json'
    @{union=@('ReportAlpha','ReportExplicitSlow')} | ConvertTo-Json | Set-Content -LiteralPath $manifest -Encoding utf8
    foreach($case in @('union','group-skip','capability-skip','group-and-capability-skip','unknown-group','slow-only-group','list-only')) {
        $global:ooxPdfAffectedReportFault=$case
        $global:ooxPdfAffectedReportCalls=0
        $output=Join-Path $root $case
        $parameters=@{RunnerPath=$runner;OutputDirectory=$output;Capability='union';SelectionManifest=$manifest;FastGroup=@('fixture','auxiliary','fixture')}
        if($case -eq 'group-and-capability-skip') {$parameters.ExpectedSkip=@('ReportAlpha')}
        if($case -eq 'unknown-group') {$parameters.FastGroup=@('missing')}
        if($case -eq 'slow-only-group') {$parameters.FastGroup=@('only-slow')}
        if($case -eq 'list-only') {$parameters.ListOnly=$true}
        $caught=$null
        $beforeTmpDir=[Environment]::GetEnvironmentVariable('TMPDIR')
        try {& (Join-Path $PSScriptRoot 'RunDocxAffectedTests.ps1') @parameters *> (Join-Path $root ($case+'.log'))}
        catch {$caught=$_.Exception.Message}
        if([Environment]::GetEnvironmentVariable('TMPDIR') -cne $beforeTmpDir) {throw 'Union did not restore TMPDIR'}
        $shouldPass=$case -in @('union','group-skip','group-and-capability-skip','list-only')
        if($shouldPass -eq ($null -ne $caught)) {throw ($case+': unexpected union gate outcome '+$caught)}
        $preflight=$case -in @('unknown-group','slow-only-group')
        $calls=if($preflight -or $case -eq 'list-only') {0} else {1}
        if($global:ooxPdfAffectedReportCalls -ne $calls) {throw ($case+': wrong runner invocation count')}
        if(!$preflight) {
            $actual=@(Get-Content -LiteralPath (Join-Path $output 'selected-methods.json') -Raw | ConvertFrom-Json)
            $wanted=@('ReportAlpha','ReportExplicitSlow','ReportGroupOnly','ReportNeighbor')
            if($actual.Count -ne 4 -or @(Compare-Object $wanted $actual).Count) {throw 'Union lost, repeated or admitted an unwanted slow method'}
            if($case -ne 'list-only') {
                $report=Get-Content -LiteralPath (Join-Path $output 'affected-tests.json') -Raw | ConvertFrom-Json
                $groupSkips=if($case -in @('group-skip','group-and-capability-skip')) {@('ReportGroupOnly')} else {@()}
                if($report.provenance.GatePassed -ne $shouldPass -or $report.tests.Count -ne 4 -or
                    @($report.provenance.GroupOnlySkips).Count -ne $groupSkips.Count -or
                    @($report.provenance.GroupOnlySkips | Where-Object { $_ -cnotin $groupSkips }).Count -or
                    $report.provenance.UnexpectedSkips.Count -ne $(if($case -eq 'capability-skip') {1} else {0})) {
                    throw 'Union skip contracts or outcomes changed'
                }
            }
        }
        $checks+=@{Case=$case;GatePassed=$shouldPass;Error=$caught;RunnerInvocations=$global:ooxPdfAffectedReportCalls}
    }

    # An exact method name is still a substring selector. Preserve its neighbors,
    # literal punctuation and ordinal casing, and never match across catalogue rows.
    $selectorCases=@(
        @{Name='selector-overlap';Filters=@('SelectionAlpha','ectionOther','SelectionAlpha');Wanted=@('SelectionAlpha','SelectionAlphaNeighbor','Outer.SelectionAlpha','SelectionOther')},
        @{Name='selector-literal-case';Filters=@('selectionalpha','ITEM','Selection[Literal]');Wanted=@('SelectionAlpha','SelectionAlphaNeighbor','Outer.SelectionAlpha','item','Selection[Literal]')},
        @{Name='selector-missing';Filters=@('SelectionAlpha','MissingSelection');Wanted=@()},
        @{Name='selector-cross-row';Filters=@("SelectionAlpha`nSelectionAlphaNeighbor");Wanted=@()}
    )
    $previousCulture=[Globalization.CultureInfo]::CurrentCulture
    try {
        [Globalization.CultureInfo]::CurrentCulture=[Globalization.CultureInfo]::GetCultureInfo('tr-TR')
        foreach($case in $selectorCases) {
            $global:ooxPdfAffectedReportFault=''
            $global:ooxPdfAffectedReportCalls=0
            $output=Join-Path $root $case.Name
            $caught=$null
            try {& (Join-Path $PSScriptRoot 'RunDocxAffectedTests.ps1') -RunnerPath $runner -OutputDirectory $output -TestFilter $case.Filters *> (Join-Path $root ($case.Name+'.log'))}
            catch {$caught=$_.Exception.Message}
            $shouldPass=$case.Wanted.Count -gt 0
            if($shouldPass -eq ($null -ne $caught) -or $global:ooxPdfAffectedReportCalls -ne [int]$shouldPass) {throw ($case.Name+': unexpected selector outcome '+$caught)}
            if($shouldPass) {
                $actual=@(Get-Content -LiteralPath (Join-Path $output 'selected-methods.json') -Raw | ConvertFrom-Json)
                if($actual.Count -ne $case.Wanted.Count -or @(Compare-Object $case.Wanted $actual -CaseSensitive).Count) {throw 'Selector matching changed'}
            }
            elseif($caught -notlike 'Empty capability selector:*') {throw 'Missing selector was not refused explicitly'}
            $checks+=@{Case=$case.Name;GatePassed=$shouldPass;Error=$caught;RunnerInvocations=$global:ooxPdfAffectedReportCalls}
        }
    }
    finally {[Globalization.CultureInfo]::CurrentCulture=$previousCulture}
}
finally {Remove-Item Function:dotnet; Remove-Variable ooxPdfAffectedReportFault,ooxPdfAffectedReportCalls -Scope Global}
@{Checks=$checks;Builds=0;Conversions=0;RealTestsReplayed=0} | ConvertTo-Json -Depth 5 | Set-Content (Join-Path $root 'evidence.json') -Encoding utf8
"Affected reporting: $($checks.Count) checks pass; no builds or test replay."
