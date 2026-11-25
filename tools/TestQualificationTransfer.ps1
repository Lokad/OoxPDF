# Check transfer ownership and refusal with byte fixtures, never a real CLI/build.
[CmdletBinding()]
param(
    [ValidateSet('net10.0','net11.0')] [string] $Framework = 'net11.0',
    [string] $OutputDirectory,
    [string] $TransferScript = (Join-Path $PSScriptRoot 'TransferQualificationArtifact.ps1')
)
$ErrorActionPreference='Stop'
if (!$OutputDirectory) { $OutputDirectory=Join-Path (Split-Path $PSScriptRoot) ('artifacts/transfer-test-'+[Guid]::NewGuid().ToString('N')) }
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Use a fresh transfer self-test directory' }
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
$root=(Resolve-Path -LiteralPath $OutputDirectory).Path
$asset="src/Lokad.OoxPdf/bin/Release/$Framework/Lokad.OoxPdf.dll"
$rows=@()
function Write-Text([string] $Path,[string] $Value) {
    New-Item -ItemType Directory -Path (Split-Path -Parent $Path) -Force | Out-Null
    [IO.File]::WriteAllText($Path,$Value)
}
function Write-Json([string] $Path,$Value) { Write-Text $Path (ConvertTo-Json -InputObject $Value -Depth 8) }
function Invoke-Git([string] $Directory,[string[]] $Arguments) {
    $output=@(& git -C $Directory @Arguments 2>&1)
    if ($LASTEXITCODE -ne 0) { throw ($output -join "`n") }
    return $output
}
function dotnet {
    if ($args.Count -lt 4 -or $args[1] -ne 'convert') { throw 'Only a synthetic transfer smoke is allowed' }
    $kind = [IO.File]::ReadAllText($args[2])
    if ($kind -eq 'smoke-failure') { throw 'Synthetic smoke failure' }
    $exit = switch ($kind) {
        { $_ -in @('expected-warning','unexpected-warning','wrong-warning-details') } { 3 }
        'fatal-exit' { 2 }
        default { 0 }
    }
    if ($kind -in @('expected-warning','wrong-warning-details','fatal-exit')) {
        if ($args.Count -ne 7 -or $args[4] -ne '--strict' -or $args[5] -ne '--diagnostics') {
            throw 'The warning smoke must bind its diagnostics sidecar'
        }
        [IO.File]::WriteAllText($args[6], $(if ($kind -eq 'wrong-warning-details') { 'other warning' } else { 'expected warning' }))
    }
    elseif ($args.Count -ne 4 -and ($args.Count -ne 5 -or $args[4] -ne '--strict')) { throw 'Unexpected smoke arguments' }
    [IO.File]::WriteAllText($args[3], $(if ($kind -eq 'wrong-pdf') { 'different PDF' } else { 'synthetic PDF result' }))
    Set-Variable -Name LASTEXITCODE -Value $exit -Scope 1
}
try {
    foreach ($kind in @('independent','hard-linked','smoke-failure','publication-failure','directory-link','source-mismatch','input-mismatch','untracked-import',
        'expected-warning','unexpected-warning','wrong-warning-details','fatal-exit','wrong-pdf','missing-warning-pin','unsupported-exit')) {
        $checkout=Join-Path $root $kind
        New-Item -ItemType Directory -Path (Join-Path $checkout tools) -Force | Out-Null
        Copy-Item -LiteralPath $TransferScript -Destination (Join-Path $checkout tools/TransferQualificationArtifact.ps1)
        Copy-Item -LiteralPath (Join-Path (Split-Path -Parent $TransferScript) 'QualificationBuildIdentity.ps1') -Destination (Join-Path $checkout tools)
        Write-Text (Join-Path $checkout .gitignore) "artifacts/`n**/bin`n"
        Write-Text (Join-Path $checkout src/source.txt) 'unchanged source'
        $null=Invoke-Git $checkout @('init','--quiet')
        $null=Invoke-Git $checkout @('add','.')
        $null=Invoke-Git $checkout @('-c','user.name=Transfer self-test','-c','user.email=transfer@example.invalid','-c','commit.gpgsign=false','commit','--quiet','-m','source')
        $revision=[string](Invoke-Git $checkout @('rev-parse','HEAD'))
        $artifact=Join-Path $checkout artifacts/candidate
        $newFile=Join-Path $artifact $asset
        Write-Text $newFile 'new frozen bytes'
        $newHash=(Get-FileHash -LiteralPath $newFile).Hash
        $oldFile=Join-Path $checkout "artifacts/previous/Release/$Framework/Lokad.OoxPdf.dll"
        Write-Text $oldFile 'previous retained bytes'
        $oldHash=(Get-FileHash -LiteralPath $oldFile).Hash
        $target=Join-Path $checkout $asset
        if ($kind -eq 'directory-link') {
            $bin=Join-Path $checkout src/Lokad.OoxPdf/bin
            New-Item -ItemType Directory -Path (Split-Path $bin) -Force | Out-Null
            $linkKind=if ($IsWindows) {'Junction'} else {'SymbolicLink'}
            New-Item -ItemType $linkKind -Path $bin -Target (Join-Path $checkout artifacts/previous) | Out-Null
        }
        elseif ($kind -eq 'publication-failure') { Write-Text (Join-Path $target retained.txt) 'retained directory' }
        else {
            New-Item -ItemType Directory -Path (Split-Path $target) -Force | Out-Null
            if ($kind -eq 'independent') { Copy-Item -LiteralPath $oldFile -Destination $target }
            else { New-Item -ItemType HardLink -Path $target -Target $oldFile | Out-Null }
        }
        Write-Json (Join-Path $artifact identity.json) @{Revision=$revision;Framework=$Framework;LibrarySha256=$newHash;TestAssemblySha256=('B'*64);Assets=@{$asset=$newHash}}
        Write-Json (Join-Path $checkout artifacts/qualification.json) @{LibrarySha256=$newHash;TestAssemblySha256=('B'*64)}
        $inputFile=Join-Path $checkout artifacts/input.docx
        Write-Text $inputFile $kind
        $expectedPdf=Join-Path $checkout artifacts/expected.pdf
        Write-Text $expectedPdf 'synthetic PDF result'
        $job=@{Input=$inputFile;InputSha256=(Get-FileHash -LiteralPath $inputFile).Hash;PdfSha256=(Get-FileHash -LiteralPath $expectedPdf).Hash;Arguments=@()}
        if ($kind -in @('expected-warning','unexpected-warning','wrong-warning-details','fatal-exit','missing-warning-pin')) { $job.Arguments=@('--strict') }
        if ($kind -in @('expected-warning','wrong-warning-details','fatal-exit','missing-warning-pin')) { $job.ExpectedExitCode=3 }
        if ($kind -eq 'unsupported-exit') { $job.ExpectedExitCode=2 }
        if ($kind -in @('expected-warning','wrong-warning-details','fatal-exit')) {
            $expectedDiagnostics=Join-Path $checkout artifacts/expected.json
            Write-Text $expectedDiagnostics 'expected warning'
            $job.DiagnosticsSha256=(Get-FileHash -LiteralPath $expectedDiagnostics).Hash
        }
        Write-Json (Join-Path $checkout artifacts/smoke.json) @($job)
        if ($kind -eq 'source-mismatch') { Write-Text $newFile 'changed before preflight' }
        if ($kind -eq 'input-mismatch') { Write-Text $inputFile 'changed before preflight' }
        if ($kind -eq 'untracked-import') { Write-Text (Join-Path $checkout 'shared/build/Local.props') '<Project/>' }
        $sourceHashBefore=(Get-FileHash -LiteralPath $newFile).Hash
        $output=Join-Path $checkout artifacts/transfer
        $caught=$null; $cause=$null
        try {
            & (Join-Path $checkout tools/TransferQualificationArtifact.ps1) -ArtifactDirectory $artifact `
                -QualificationEvidence (Join-Path $checkout artifacts/qualification.json) `
                -SmokeManifest (Join-Path $checkout artifacts/smoke.json) -OutputDirectory $output | Out-Null
        } catch { $cause=$_.Exception.GetBaseException(); $caught=$cause.Message }
        $success=$kind -in @('independent','hard-linked','expected-warning')
        if ($success -and $caught) { throw ('Unexpected transfer failure: '+$kind+' / '+$caught) }
        if (!$success -and !$caught) { throw ('Expected transfer refusal: '+$kind) }
        $expected=@{'smoke-failure'='Synthetic smoke failure';'directory-link'='Root asset directories must not be linked';
            'source-mismatch'='Frozen asset differs';'input-mismatch'='Smoke input differs';
            'untracked-import'='Source/input/build-setting transfer identity differs';
            'unexpected-warning'='Transfer smoke exit differs';'wrong-warning-details'='Transfer smoke diagnostics differ';
            'fatal-exit'='Transfer smoke exit differs';'wrong-pdf'='Transfer smoke PDF differs';
            'missing-warning-pin'='Smoke diagnostics require a SHA-256 pin';'unsupported-exit'='Smoke expected exit must be 0 or 3'}
        if ($expected.ContainsKey($kind) -and $caught -ne $expected[$kind]) { throw ('Unexpected refusal: '+$kind+' / '+$caught) }
        if ($kind -eq 'publication-failure' -and $cause -isnot [IO.IOException] -and $cause -isnot [UnauthorizedAccessException]) { throw ('Expected a publication I/O failure: '+$caught) }
        if ((Get-FileHash -LiteralPath $oldFile).Hash -ne $oldHash) { throw ('Previous artifact changed: '+$kind) }
        if ((Get-FileHash -LiteralPath $newFile).Hash -ne $sourceHashBefore) { throw ('Frozen source changed: '+$kind) }
        if (@(Get-ChildItem -LiteralPath (Split-Path $target) -Filter '*.staging-*' -File).Count) { throw ('Owned staging file leaked: '+$kind) }
        $preflight=$kind -in @('directory-link','source-mismatch','input-mismatch','untracked-import','missing-warning-pin','unsupported-exit')
        if ($preflight -and ((Test-Path -LiteralPath $output) -or (Get-FileHash -LiteralPath $target).Hash -ne $oldHash)) { throw ('Preflight changed output: '+$kind) }
        if ($kind -eq 'publication-failure' -and [IO.File]::ReadAllText((Join-Path $target retained.txt)) -ne 'retained directory') { throw 'Publication failure changed the existing directory' }
        if (!$preflight -and $kind -ne 'publication-failure' -and (Get-FileHash -LiteralPath $target).Hash -ne $newHash) { throw ('Root asset differs: '+$kind) }
        $evidence=Join-Path $output evidence.json
        if ($success) {
            $report=Get-Content -LiteralPath $evidence -Raw | ConvertFrom-Json
            if ($report.AssetsVerified -ne 1 -or $report.Rebuilt -or $report.RepeatedSuite -or @($report.Smoke).Count -ne 1) { throw 'Transfer evidence differs' }
            if ($kind -eq 'expected-warning' -and ($report.Smoke[0].ExpectedExitCode -ne 3 -or
                (Get-FileHash -LiteralPath (Join-Path $output '0.diagnostics.json')).Hash -cne $job.DiagnosticsSha256)) {
                throw 'Expected strict warning evidence differs'
            }
        }
        elseif (Test-Path -LiteralPath $evidence) { throw 'Failed transfer published completed evidence' }
        $rows+=@{Case=$kind;Succeeded=$success;Error=$caught;PreviousArtifactPreserved=$true;FrozenSourcePreserved=$true;OwnedStagingCleaned=$true;Preflight=$preflight}
    }
} finally { Remove-Item Function:\dotnet }
Write-Json (Join-Path $root evidence.json) @{Checks=$rows;Builds=0;RealConversions=0;RunnerInvocations=0;ToolSha256=(Get-FileHash -LiteralPath $TransferScript).Hash}
"Transfer checks pass for $($rows.Count) ownership, exit, PDF and diagnostics controls."
