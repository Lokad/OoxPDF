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
function Get-FileHash {
    param([string] $LiteralPath)
    $name = [IO.Path]::GetFileName($LiteralPath)
    # Replace only probe inputs when the wrapper verifies their current digest.
    if ($name -ceq 'mutable-capabilities.json' -and $global:ooxPdfAffectedReportFault -in @('changed-manifest','changed-list-manifest')) {
        [IO.File]::WriteAllText($LiteralPath, '{"union":["=SelectionOther"]}')
    }
    if ($name -ceq 'selection-000.json' -and $global:ooxPdfAffectedReportFault -eq 'changed-raw-report') {
        [IO.File]::WriteAllText($LiteralPath, '{"passed":0,"failed":1,"skipped":0,"tests":[{"name":"ReportAlpha","outcome":"failed"}]}')
    }
    Microsoft.PowerShell.Utility\Get-FileHash -LiteralPath $LiteralPath
}
function dotnet {
    if($args -contains '--list') {
        'FAST fixture ReportAlpha';'FAST fixture ReportNeighbor';'FAST auxiliary ReportGroupOnly'
        'SLOW auxiliary ReportGroupSlow';'SLOW fixture ReportExplicitSlow';'SLOW only-slow ReportSlowOnly'
        'FAST selection SelectionAlpha';'FAST selection SelectionAlphaNeighbor';'FAST selection Outer.SelectionAlpha'
        'FAST selection SelectionOther';'FAST selection Selection[Literal]';'FAST selection SelectionL'
        'FAST selection item';'FAST selection İtemNeighbor'
        'FAST cli CliFixture'
        'FAST docx-tables ReplayCellMemo';'SLOW docx-text WrapEmergencyToken'
        'FAST docx-tables DocxTableControl';'FAST docx-composition-portable PortableRenderingSharedControl'
        'FAST docx-math-portable MathFontSharedControl';'FAST rendering-portable PortableRenderingPptxControl'
        'FAST rendering-portable BalloonSharedControl';'FAST pptx-typography LayoutSlide'
        $global:LASTEXITCODE=0;return
    }
    $global:ooxPdfAffectedReportCalls++
    $report=$args[[Array]::IndexOf([object[]]$args,'--report')+1]
    if([IO.Path]::GetFileName($env:TMPDIR) -ne 'temporary-files') {throw 'Runner temp path escaped its output directory'}
    if($env:OOXPDF_TEST_REQUIRE_BOUND_TOOLS -cne '1') {throw 'Runner did not require bound tools'}
    if ($global:ooxPdfAffectedReportFault -eq 'missing-report') { $global:LASTEXITCODE=0; return }
    if ($global:ooxPdfAffectedReportFault -in @('failed','cleanup-residual','cleanup-link','cleanup-child-link')) {
        $scratch=New-Item -ItemType Directory -Path (Join-Path $env:TMPDIR 'nested')
        [IO.File]::WriteAllText((Join-Path $scratch.FullName 'retained.bin'),'Retained runner data')
    }
    $selection=$args[[Array]::IndexOf([object[]]$args,'--test-set')+1]
    $methods=@(Get-Content -LiteralPath $selection -Raw | ConvertFrom-Json)
    $rows=@(foreach($method in $methods) {
        $outcome=if($method -eq 'ReportAlpha' -and $global:ooxPdfAffectedReportFault -eq 'failed') {'failed'}
            elseif($global:ooxPdfAffectedReportFault -eq 'all-expected-skips' -or
                ($global:ooxPdfAffectedReportFault -eq 'all-primary-skips' -and $method -in @('ReportAlpha','ReportExplicitSlow')) -or
                ($method -eq 'ReportAlpha' -and $global:ooxPdfAffectedReportFault -in @('expected-skip','unexpected-skip','capability-skip','group-and-capability-skip')) -or
                ($method -eq 'ReportGroupOnly' -and $global:ooxPdfAffectedReportFault -in @('group-skip','group-and-capability-skip'))) {'skipped'}
            else {'passed'}
        @{name=$method;outcome=$outcome}
    })
    if($global:ooxPdfAffectedReportFault -eq 'incomplete') {$rows=@($rows[0])}
    if($global:ooxPdfAffectedReportFault -eq 'duplicate') {$rows=@($rows[0],$rows[0])}
    if($global:ooxPdfAffectedReportFault -eq 'wrong-name-case') {$rows[0].name=$rows[0].name.ToLowerInvariant()}
    $passed=@($rows | Where-Object outcome -EQ 'passed').Count
    $failed=@($rows | Where-Object outcome -EQ 'failed').Count
    $skipped=@($rows | Where-Object outcome -EQ 'skipped').Count
    if($global:ooxPdfAffectedReportFault -eq 'miscount') {$passed=999}
    $raw = @{passed=$passed;failed=$failed;skipped=$skipped;tests=$rows}
    if ($global:ooxPdfAffectedReportFault -in @('inspector-bound','inspector-changed','inspector-unbound')) {
        $tool = Join-Path (Split-Path (Split-Path $report)) 'validation tools/PdfInspect.dll'
        $raw.validationTools = @{
            OOXPDF_TEST_PDF_INSPECTOR = @{ Path=$tool; Sha256=(Get-FileHash -LiteralPath $tool).Hash }
        }
        if ($global:ooxPdfAffectedReportFault -eq 'inspector-changed') {
            [IO.File]::WriteAllText($tool, 'Modified validation tool')
        }
    }
    $raw | ConvertTo-Json -Depth 5 | Set-Content $report -Encoding utf8
    if($global:ooxPdfAffectedReportFault -in @('cleanup-link','cleanup-child-link')) {
        $target=Join-Path (Split-Path $env:TMPDIR) 'unowned-target'
        New-Item -ItemType Directory -Path $target | Out-Null
        [IO.File]::WriteAllText((Join-Path $target 'keep.txt'),'Keep this target')
        $link=Join-Path $env:TMPDIR 'nested/link'
        if($global:ooxPdfAffectedReportFault -eq 'cleanup-link') {
            $retained=$env:TMPDIR+'.retained'
            if([IO.Path]::GetDirectoryName($retained) -ne [IO.Path]::GetDirectoryName($env:TMPDIR)) {throw 'Fixture paths escaped their owner'}
            Move-Item -LiteralPath $env:TMPDIR -Destination $retained
            $link=$env:TMPDIR
        }
        New-Item -ItemType $(if($IsWindows){'Junction'}else{'SymbolicLink'}) -Path $link -Target $target | Out-Null
    }
    $global:LASTEXITCODE=if($failed) {1} else {0}
}
$checks=@()
try {
    foreach($case in @('pass','expected-skip','all-expected-skips','unexpected-skip','failed','unknown-expected','incomplete','duplicate','wrong-name-case','miscount','missing-report')) {
        $global:ooxPdfAffectedReportFault=$case
        $global:ooxPdfAffectedReportCalls=0
        $output=Join-Path $root $case
        $expected=@()
        if($case -in @('expected-skip','failed')) {$expected=@('ReportAlpha')}
        if($case -eq 'all-expected-skips') {$expected=@('ReportAlpha','ReportNeighbor')}
        if($case -eq 'unknown-expected') {$expected=@('NotSelected')}
        $caught=$null
        $beforeTmpDir=[Environment]::GetEnvironmentVariable('TMPDIR')
        $beforeToolBinding=[Environment]::GetEnvironmentVariable('OOXPDF_TEST_REQUIRE_BOUND_TOOLS')
        try {& (Join-Path $PSScriptRoot 'RunDocxAffectedTests.ps1') -RunnerPath $runner -OutputDirectory $output -TestFilter @('ReportAlpha','ReportNeighbor') -ExpectedSkip $expected *> (Join-Path $root ($case+'.log'))}
        catch {$caught=$_.Exception.Message}
        if([Environment]::GetEnvironmentVariable('TMPDIR') -cne $beforeTmpDir) {throw 'Runner did not restore TMPDIR'}
        if([Environment]::GetEnvironmentVariable('OOXPDF_TEST_REQUIRE_BOUND_TOOLS') -cne $beforeToolBinding) {
            throw 'Runner did not restore the caller tool policy'
        }
        $shouldPass=$case -in @('pass','expected-skip')
        if($shouldPass -eq ($null -ne $caught)) {throw ($case+': unexpected gate outcome '+$caught)}
        if($case -eq 'wrong-name-case' -and ($caught -notlike 'Incomplete affected report:*' -or
            !(Test-Path -LiteralPath (Join-Path $output 'selection-000.json') -PathType Leaf))) {
            throw 'A case-only method mismatch must be rejected and retain its raw report'
        }
        if ($case -eq 'missing-report' -and $caught -ne 'Test runner did not produce its report') {
            throw 'Successful exit with a missing raw report must fail qualification'
        }
        if ($case -eq 'all-expected-skips' -and $caught -notlike 'No selected qualification test passed*') {
            throw 'A successful runner exit with only expected skips must report no passing witness'
        }
        $retained=Test-Path (Join-Path $output 'affected-tests.json')
        if($retained -ne ($case -in @('pass','expected-skip','all-expected-skips','unexpected-skip','failed'))) {throw ($case+': incorrect complete-report retention')}
        if($retained) {
            $report=Get-Content (Join-Path $output 'affected-tests.json') -Raw | ConvertFrom-Json
            if($report.provenance.GatePassed -ne $shouldPass -or $report.tests.Count -ne 2 -or
                $report.provenance.UniqueExecutedMethods -ne 2 -or $report.provenance.UniquePassedMethods -ne $report.passed -or
                $report.failed -ne $(if($case -eq 'failed') {1} else {0}) -or
                $report.skipped -ne $(if($case -eq 'all-expected-skips') {2} elseif($case -in @('expected-skip','unexpected-skip')) {1} else {0})) {
                throw 'Reported outcomes were reclassified'
            }
            if($report.provenance.TemporaryFilesRetained -eq $shouldPass) {throw 'Passing and failed temporary ownership changed'}
        }
        if($global:ooxPdfAffectedReportCalls -gt 0 -and
            (Test-Path -LiteralPath (Join-Path $output 'temporary-files')) -eq $shouldPass) {
            throw 'The empty outer directory cleanup did not follow the complete passing gate'
        }
        if($case -eq 'failed' -and
            !(Test-Path -LiteralPath (Join-Path $output 'temporary-files/nested/retained.bin') -PathType Leaf)) {
            throw 'Failed runner scratch did not survive qualification'
        }
        $checks+=@{Case=$case;GatePassed=$shouldPass;Error=$caught;ReportRetained=$retained}
    }

    $case='preexisting-raw-report'
    $output=(New-Item -ItemType Directory -Path (Join-Path $root $case)).FullName
    $rawReport=Join-Path $output 'selection-000.json'
    [IO.File]::Copy((Join-Path $root 'pass/selection-000.json'), $rawReport)
    $rawHash=(Get-FileHash -LiteralPath $rawReport).Hash
    $global:ooxPdfAffectedReportFault='missing-report'
    $global:ooxPdfAffectedReportCalls=0
    $caught=$null
    try {& (Join-Path $PSScriptRoot 'RunDocxAffectedTests.ps1') -RunnerPath $runner -OutputDirectory $output -TestFilter @('ReportAlpha','ReportNeighbor') *> (Join-Path $root ($case+'.log'))}
    catch {$caught=$_.Exception.Message}
    if ($caught -ne 'Output directory already contains qualification reports' -or
        $global:ooxPdfAffectedReportCalls -ne 0 -or (Get-FileHash -LiteralPath $rawReport).Hash -ne $rawHash -or
        (Test-Path -LiteralPath (Join-Path $output 'affected-tests.json'))) {
        throw 'A preexisting raw report must be preserved and refused before execution'
    }
    $checks+=@{Case=$case;GatePassed=$false;Error=$caught;RawReportSha256=$rawHash;RunnerInvocations=$global:ooxPdfAffectedReportCalls}

    $manifest=Join-Path $root 'union-capabilities.json'
    @{union=@('=ReportAlpha','ReportExplicitSlow'); 'skip-only'=@('=ReportAlpha')} |
        ConvertTo-Json | Set-Content -LiteralPath $manifest -Encoding unicode
    $manifestHash=(Get-FileHash -LiteralPath $manifest).Hash
    foreach($case in @('union','group-skip','capability-skip','group-and-capability-skip','partial-capability-coverage','all-primary-skips','unknown-group','slow-only-group','list-only')) {
        $global:ooxPdfAffectedReportFault=$case
        $global:ooxPdfAffectedReportCalls=0
        $output=Join-Path $root $case
        $parameters=@{RunnerPath=$runner;OutputDirectory=$output;Capability='union';SelectionManifest=$manifest;FastGroup=@('fixture','auxiliary','fixture')}
        if($case -eq 'partial-capability-coverage') {
            $global:ooxPdfAffectedReportFault='group-and-capability-skip'
            $parameters.Capability=@('skip-only','union')
            $parameters.ExpectedSkip=@('ReportAlpha')
        }
        if($case -eq 'group-and-capability-skip') {$parameters.ExpectedSkip=@('ReportAlpha')}
        if($case -eq 'all-primary-skips') {$parameters.ExpectedSkip=@('ReportAlpha','ReportExplicitSlow')}
        if($case -eq 'unknown-group') {$parameters.FastGroup=@('missing')}
        if($case -eq 'slow-only-group') {$parameters.FastGroup=@('only-slow')}
        if($case -eq 'list-only') {$parameters.ListOnly=$true}
        $caught=$null
        $beforeTmpDir=[Environment]::GetEnvironmentVariable('TMPDIR')
        try {& (Join-Path $PSScriptRoot 'RunDocxAffectedTests.ps1') @parameters *> (Join-Path $root ($case+'.log'))}
        catch {$caught=$_.Exception.Message}
        if([Environment]::GetEnvironmentVariable('TMPDIR') -cne $beforeTmpDir) {throw 'Union did not restore TMPDIR'}
        $shouldPass=$case -in @('union','group-skip','group-and-capability-skip','partial-capability-coverage','list-only')
        if($shouldPass -eq ($null -ne $caught)) {throw ($case+': unexpected union gate outcome '+$caught)}
        $preflight=$case -in @('unknown-group','slow-only-group')
        $calls=if($preflight -or $case -eq 'list-only') {0} else {1}
        if($global:ooxPdfAffectedReportCalls -ne $calls) {throw ($case+': wrong runner invocation count')}
        if(!$preflight) {
            $actual=@(Get-Content -LiteralPath (Join-Path $output 'selected-methods.json') -Raw | ConvertFrom-Json)
            $wanted=@('ReportAlpha','ReportExplicitSlow','ReportGroupOnly','ReportNeighbor')
            if($actual.Count -ne 4 -or @(Compare-Object $wanted $actual).Count) {throw 'Union lost, repeated or admitted an unwanted slow method'}
            if($case -eq 'list-only') {
                $inventory=Get-Content -LiteralPath (Join-Path $root ($case+'.log')) -Raw | ConvertFrom-Json
                if($inventory.RunnerInvocations -ne 0 -or $inventory.SelectionManifestSha256 -cne $manifestHash -or
                    (Test-Path -LiteralPath (Join-Path $output 'affected-tests.json'))) {
                    throw 'List-only inventory must report zero test runs and publish no execution report'
                }
            }
            else {
                $report=Get-Content -LiteralPath (Join-Path $output 'affected-tests.json') -Raw | ConvertFrom-Json
                $groupSkips=if($case -in @('group-skip','group-and-capability-skip','partial-capability-coverage')) {@('ReportGroupOnly')} else {@()}
                $primaryPasses=switch ($case) {
                    'all-primary-skips' {0}
                    {$_ -in @('capability-skip','group-and-capability-skip','partial-capability-coverage')} {1}
                    default {2}
                }
                if ($case -eq 'all-primary-skips' -and ($report.passed -ne 2 -or $caught -notlike 'No primary qualification test passed.*')) {
                    throw 'Fast-group-only passes must be reported and cannot qualify an all-skipped primary selection'
                }
                if($report.provenance.GatePassed -ne $shouldPass -or $report.tests.Count -ne 4 -or
                    $report.provenance.UniquePassedPrimaryMethods -ne $primaryPasses -or
                    $report.provenance.SelectionManifestSha256 -cne $manifestHash -or
                    $report.provenance.SourceReports[0].Sha256 -cne (Get-FileHash -LiteralPath (Join-Path $output 'selection-000.json')).Hash -or
                    $report.provenance.RunnerInvocations -ne 1 -or
                    @($report.provenance.GroupOnlySkips).Count -ne $groupSkips.Count -or
                    @($report.provenance.GroupOnlySkips | Where-Object { $_ -cnotin $groupSkips }).Count -or
                    $report.provenance.UnexpectedSkips.Count -ne $(if($case -eq 'capability-skip') {1} else {0})) {
                    throw 'Union skip contracts or outcomes changed'
                }
                if($case -eq 'partial-capability-coverage') {
                    $coverage=$report.provenance.CapabilityCoverage
                    $skipped=@($coverage | Where-Object Name -CEQ 'skip-only')
                    $union=@($coverage | Where-Object Name -CEQ 'union')
                    if($coverage.Count -ne 2 -or $skipped.Count -ne 1 -or $union.Count -ne 1 -or
                        $skipped[0].Selected -ne 1 -or $skipped[0].Skipped -ne 1 -or $skipped[0].Qualified -or
                        $union[0].Selected -ne 2 -or $union[0].Passed -ne 1 -or $union[0].Skipped -ne 1 -or
                        !$union[0].Qualified -or
                        !(Select-String -LiteralPath (Join-Path $root ($case+'.log')) -SimpleMatch 'All selected methods skipped for: skip-only' -Quiet)) {
                        throw 'A capability pass from another selection must not qualify a fully skipped capability'
                    }
                }
            }
        }
        $checks+=@{Case=$case;GatePassed=$shouldPass;Error=$caught;RunnerInvocations=$global:ooxPdfAffectedReportCalls}
    }

    foreach ($case in @('changed-manifest','changed-list-manifest','changed-raw-report')) {
        $global:ooxPdfAffectedReportCalls=0
        $output=Join-Path $root $case
        $parameters=@{RunnerPath=$runner;OutputDirectory=$output}
        $listOnly=$case -eq 'changed-list-manifest'
        $rawChanges=$case -eq 'changed-raw-report'
        if ($rawChanges) { $parameters.TestFilter='=ReportAlpha' }
        else {
            $mutable=Join-Path $root 'mutable-capabilities.json'
            [IO.File]::Copy($manifest,$mutable,$true)
            $parameters.Capability='union'
            $parameters.SelectionManifest=$mutable
            $parameters.ListOnly=$listOnly
        }
        $expected=if ($rawChanges) { 'The raw test report changed during qualification' }
            else { 'The capability selection manifest changed during qualification' }
        $global:ooxPdfAffectedReportFault=$case
        $caught=$null
        try { & (Join-Path $PSScriptRoot 'RunDocxAffectedTests.ps1') @parameters *> (Join-Path $root ($case+'.log')) }
        catch { $caught=$_.Exception.Message }
        $raw=Join-Path $output 'selection-000.json'
        if ($caught -cne $expected -or $global:ooxPdfAffectedReportCalls -ne [int](!$listOnly) -or
            (Test-Path -LiteralPath (Join-Path $output 'affected-tests.json')) -or
            (Test-Path -LiteralPath (Join-Path $output 'evidence.json')) -or
            (Test-Path -LiteralPath $raw -PathType Leaf) -eq $listOnly) {
            throw ('Changed qualification input was accepted or its raw evidence was lost: ' + $case)
        }
        $checks+=@{Case=$case;GatePassed=$false;Error=$caught;RunnerInvocations=$global:ooxPdfAffectedReportCalls}
    }

    foreach($case in @('milestone-list','milestone-run')) {
        $global:ooxPdfAffectedReportFault=''
        $global:ooxPdfAffectedReportCalls=0
        $output=Join-Path $root $case
        $listOnly=$case -eq 'milestone-list'
        & (Join-Path $PSScriptRoot 'RunDocxAffectedTests.ps1') -RunnerPath $runner -OutputDirectory $output -AllDocx -ListOnly:$listOnly *> (Join-Path $root ($case+'.log'))
        $actual=@(Get-Content -LiteralPath (Join-Path $output 'selected-methods.json') -Raw | ConvertFrom-Json)
        $wanted=@('ReplayCellMemo','WrapEmergencyToken','DocxTableControl','PortableRenderingSharedControl','MathFontSharedControl','BalloonSharedControl')
        if ($actual.Count -ne $wanted.Count -or @(Compare-Object $wanted $actual -CaseSensitive).Count -or
            $global:ooxPdfAffectedReportCalls -ne [int]!$listOnly) {
            throw 'DOCX milestones must include the registered groups and shared neighbors once'
        }
        $checks+=@{Case=$case;SelectedMethods=$actual;RunnerInvocations=$global:ooxPdfAffectedReportCalls}
    }

    # Bare selectors keep literal substring matching. The equals prefix requires
    # ordinal identity with one registered name, even in a mixed selection.
    $selectorCases=@(
        @{Name='selector-overlap';Filters=@('SelectionAlpha','ectionOther','SelectionAlpha');Wanted=@('SelectionAlpha','SelectionAlphaNeighbor','Outer.SelectionAlpha','SelectionOther')},
        @{Name='selector-literal-case';Filters=@('selectionalpha','ITEM','Selection[Literal]');Wanted=@('SelectionAlpha','SelectionAlphaNeighbor','Outer.SelectionAlpha','item','Selection[Literal]')},
        @{Name='selector-exact-mixed';Filters=@('=SelectionAlpha','ectionOther','=SelectionAlpha','=Selection[Literal]');Wanted=@('SelectionAlpha','SelectionOther','Selection[Literal]')},
        @{Name='selector-exact-case';Filters=@('=selectionalpha');Wanted=@()},
        @{Name='selector-exact-prefix';Filters=@('=SelectionAlph');Wanted=@()},
        @{Name='selector-exact-empty';Filters=@('=');Wanted=@()},
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

    $priorCli = [Environment]::GetEnvironmentVariable('OOXPDF_TEST_CLI_ASSEMBLY')
    try {
        foreach ($case in @('cli-unbound','cli-mismatched','cli-bound','cli-list-only')) {
            $global:ooxPdfAffectedReportFault=''
            $global:ooxPdfAffectedReportCalls=0
            $output=Join-Path $root $case
            $cliDirectory=(New-Item -ItemType Directory -Path (Join-Path $root ($case+'-cli'))).FullName
            $cli=Join-Path $cliDirectory 'Lokad.OoxPdf.Cli.dll'
            [IO.File]::WriteAllText($cli,'Fake frozen CLI')
            [IO.File]::WriteAllText((Join-Path $cliDirectory 'Lokad.OoxPdf.dll'),
                $(if ($case -eq 'cli-mismatched') {'Another library'} else {'Fake library identity'}))
            $env:OOXPDF_TEST_CLI_ASSEMBLY=if ($case -in @('cli-unbound','cli-list-only')) {$null} else {$cli}
            $caught=$null
            try {& (Join-Path $PSScriptRoot 'RunDocxAffectedTests.ps1') -RunnerPath $runner -OutputDirectory $output -TestFilter CliFixture -ListOnly:($case -eq 'cli-list-only') *> (Join-Path $root ($case+'.log'))}
            catch {$caught=$_.Exception.Message}
            $shouldPass=$case -in @('cli-bound','cli-list-only')
            $calls=[int]($case -eq 'cli-bound')
            if ($shouldPass -eq ($null -ne $caught) -or $global:ooxPdfAffectedReportCalls -ne $calls) {
                throw ($case+': incorrect frozen CLI admission '+$caught)
            }
            if ($case -eq 'cli-bound') {
                $report=Get-Content -LiteralPath (Join-Path $output 'affected-tests.json') -Raw | ConvertFrom-Json
                if ($report.provenance.CliSha256 -ne (Get-FileHash -LiteralPath $cli).Hash) {throw 'CLI identity is missing from qualification'}
            }
            $checks+=@{Case=$case;GatePassed=$shouldPass;Error=$caught;RunnerInvocations=$global:ooxPdfAffectedReportCalls}
        }
    }
    finally {[Environment]::SetEnvironmentVariable('OOXPDF_TEST_CLI_ASSEMBLY', $(if ($null -eq $priorCli) {[NullString]::Value} else {$priorCli}))}

    $inspector = Join-Path (New-Item -ItemType Directory (Join-Path $root 'validation tools')).FullName 'PdfInspect.dll'
    $priorInspector = [Environment]::GetEnvironmentVariable('OOXPDF_TEST_PDF_INSPECTOR')
    try {
        foreach ($case in @('inspector-bound','inspector-changed','inspector-unbound','tool-free-selection')) {
            [IO.File]::WriteAllText($inspector, 'Frozen validation tool')
            $env:OOXPDF_TEST_PDF_INSPECTOR = switch ($case) {
                'inspector-unbound' { $null }
                'tool-free-selection' { Join-Path $root 'absent.dll' }
                default { $inspector }
            }
            $global:ooxPdfAffectedReportFault=$case
            $output=Join-Path $root $case
            $caught=$null
            try {& (Join-Path $PSScriptRoot 'RunDocxAffectedTests.ps1') -RunnerPath $runner -OutputDirectory $output -TestFilter ReportAlpha *> (Join-Path $root ($case+'.log'))}
            catch {$caught=$_.Exception.Message}
            $expected = switch ($case) {
                'inspector-changed' { 'Validation tool changed during qualification: OOXPDF_TEST_PDF_INSPECTOR' }
                'inspector-unbound' { 'Validation tool report differs from the supplied binding: OOXPDF_TEST_PDF_INSPECTOR' }
            }
            if ($caught -cne $expected) { throw ($case+': validation tool identity result differs '+$caught) }
            if (!$caught) {
                $affected=Get-Content -LiteralPath (Join-Path $output 'affected-tests.json') -Raw | ConvertFrom-Json
                $binding=$affected.provenance.ValidationTools.OOXPDF_TEST_PDF_INSPECTOR
                if ($case -eq 'inspector-bound' -and ($binding.Path -cne $inspector -or
                    $binding.Sha256 -cne (Get-FileHash -LiteralPath $inspector).Hash)) { throw 'Inspector identity is missing from qualification' }
                if ($case -eq 'tool-free-selection' -and $affected.provenance.ValidationTools) { throw 'Tool-free selection published an unused binding' }
            }
            $checks+=@{Case=$case;GatePassed=(!$caught);Error=$caught}
        }
    }
    finally {
        [Environment]::SetEnvironmentVariable('OOXPDF_TEST_PDF_INSPECTOR', $(if ($null -eq $priorInspector) {[NullString]::Value} else {$priorInspector}))
    }

    foreach($case in @('temp-existing','temp-linked-output','cleanup-link','cleanup-child-link','cleanup-residual')) {
        $global:ooxPdfAffectedReportFault=$case
        $global:ooxPdfAffectedReportCalls=0
        $directory=Join-Path $root $case
        New-Item -ItemType Directory -Path $directory | Out-Null
        $output=Join-Path $directory 'output'
        if($case -eq 'temp-existing') {
            New-Item -ItemType Directory -Path (Join-Path $output 'temporary-files') | Out-Null
            [IO.File]::WriteAllText((Join-Path $output 'temporary-files/keep.txt'),'Keep existing storage')
        }
        if($case -eq 'temp-linked-output') {
            $target=Join-Path $directory 'physical-target'
            New-Item -ItemType Directory -Path $target | Out-Null
            [IO.File]::WriteAllText((Join-Path $target 'keep.txt'),'Keep linked output')
            New-Item -ItemType $(if($IsWindows){'Junction'}else{'SymbolicLink'}) -Path $output -Target $target | Out-Null
        }
        $caught=$null
        try {& (Join-Path $PSScriptRoot 'RunDocxAffectedTests.ps1') -RunnerPath $runner -OutputDirectory $output -TestFilter ReportAlpha *> (Join-Path $directory 'run.log')}
        catch {$caught=$_.Exception.Message}
        if($case -in @('cleanup-link','cleanup-child-link','cleanup-residual')) {
            $report=Get-Content -LiteralPath (Join-Path $output 'affected-tests.json') -Raw | ConvertFrom-Json
            $rootLink=$case -eq 'cleanup-link'
            $retained=Join-Path $output $(if ($rootLink) {'temporary-files.retained'} else {'temporary-files'})
            if($caught -or !$report.provenance.GatePassed -or !$report.provenance.TemporaryFilesRetained -or
                !$report.provenance.TemporaryCleanupError -or $global:ooxPdfAffectedReportCalls -ne 1 -or
                [IO.File]::ReadAllText((Join-Path $retained 'nested/retained.bin')) -ne 'Retained runner data') {
                throw 'The wrapper lost runner scratch or changed the passing result'
            }
            if ($case -ne 'cleanup-residual' -and
                [IO.File]::ReadAllText((Join-Path $output 'unowned-target/keep.txt')) -ne 'Keep this target') {
                throw 'The wrapper changed the unowned link target'
            }
        }
        else {
            $keep=if($case -eq 'temp-existing'){Join-Path $output 'temporary-files/keep.txt'}else{Join-Path $target 'keep.txt'}
            if(!$caught -or $global:ooxPdfAffectedReportCalls -ne 0 -or !(Test-Path -LiteralPath $keep)) {
                throw 'Preexisting or linked temporary storage was not refused before test execution'
            }
        }
        $checks+=@{Case=$case;Error=$caught;RunnerInvocations=$global:ooxPdfAffectedReportCalls;UnownedTargetRetained=$true}
    }
}
finally {Remove-Item Function:dotnet,Function:Get-FileHash; Remove-Variable ooxPdfAffectedReportFault,ooxPdfAffectedReportCalls -Scope Global}
@{Checks=$checks;Builds=0;Conversions=0;RealTestsReplayed=0} | ConvertTo-Json -Depth 5 | Set-Content (Join-Path $root 'evidence.json') -Encoding utf8
"Affected reporting: $($checks.Count) checks pass; no builds or test replay."
