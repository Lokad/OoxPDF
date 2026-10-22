# Check transfer ownership and refusal with byte fixtures, never a real CLI/build.
[CmdletBinding()]
param(
    [string] $OutputDirectory,
    [string] $TransferScript = (Join-Path $PSScriptRoot 'TransferQualificationArtifact.ps1')
)
$ErrorActionPreference='Stop'
if (!$OutputDirectory) { $OutputDirectory=Join-Path (Split-Path $PSScriptRoot) ('artifacts/transfer-test-'+[Guid]::NewGuid().ToString('N')) }
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Use a fresh transfer self-test directory' }
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
$root=(Resolve-Path -LiteralPath $OutputDirectory).Path
$asset='src/Lokad.OoxPdf/bin/Release/net10.0/Lokad.OoxPdf.dll'
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
    # The native CLI flattens the trailing empty argument array; a function sees it.
    if ($args.Count -lt 4 -or $args.Count -gt 5 -or $args[1] -ne 'convert' -or
        ($args.Count -eq 5 -and $args[4].Length -ne 0)) { throw 'Only a synthetic transfer smoke is allowed' }
    if ([IO.File]::ReadAllText($args[2]) -eq 'smoke-failure') { throw 'Synthetic smoke failure' }
    [IO.File]::WriteAllText($args[3],'synthetic PDF result')
    Set-Variable -Name LASTEXITCODE -Value 0 -Scope 1
}
try {
    foreach ($kind in @('independent','hard-linked','smoke-failure','publication-failure','directory-link','source-mismatch','input-mismatch')) {
        $checkout=Join-Path $root $kind
        New-Item -ItemType Directory -Path (Join-Path $checkout tools) -Force | Out-Null
        Copy-Item -LiteralPath $TransferScript -Destination (Join-Path $checkout tools/TransferQualificationArtifact.ps1)
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
        $oldFile=Join-Path $checkout artifacts/previous/Release/net10.0/Lokad.OoxPdf.dll
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
        Write-Json (Join-Path $artifact identity.json) @{Revision=$revision;LibrarySha256=$newHash;TestAssemblySha256=('B'*64);Assets=@{$asset=$newHash}}
        Write-Json (Join-Path $checkout artifacts/qualification.json) @{LibrarySha256=$newHash;TestAssemblySha256=('B'*64)}
        $inputFile=Join-Path $checkout artifacts/input.docx
        Write-Text $inputFile $(if ($kind -eq 'smoke-failure') {'smoke-failure'} else {'source input'})
        $expectedPdf=Join-Path $checkout artifacts/expected.pdf
        Write-Text $expectedPdf 'synthetic PDF result'
        Write-Json (Join-Path $checkout artifacts/smoke.json) @(@{Input=$inputFile;InputSha256=(Get-FileHash -LiteralPath $inputFile).Hash;PdfSha256=(Get-FileHash -LiteralPath $expectedPdf).Hash;Arguments=@()})
        if ($kind -eq 'source-mismatch') { Write-Text $newFile 'changed before preflight' }
        if ($kind -eq 'input-mismatch') { Write-Text $inputFile 'changed before preflight' }
        $sourceHashBefore=(Get-FileHash -LiteralPath $newFile).Hash
        $output=Join-Path $checkout artifacts/transfer
        $caught=$null; $cause=$null
        try {
            & (Join-Path $checkout tools/TransferQualificationArtifact.ps1) -ArtifactDirectory $artifact `
                -QualificationEvidence (Join-Path $checkout artifacts/qualification.json) `
                -SmokeManifest (Join-Path $checkout artifacts/smoke.json) -OutputDirectory $output | Out-Null
        } catch { $cause=$_.Exception.GetBaseException(); $caught=$cause.Message }
        $success=$kind -in @('independent','hard-linked')
        if ($success -and $caught) { throw ('Unexpected transfer failure: '+$kind+' / '+$caught) }
        if (!$success -and !$caught) { throw ('Expected transfer refusal: '+$kind) }
        $expected=@{'smoke-failure'='Synthetic smoke failure';'directory-link'='Root asset directories must not be linked';
            'source-mismatch'='Frozen asset differs';'input-mismatch'='Smoke input differs'}
        if ($expected.ContainsKey($kind) -and $caught -ne $expected[$kind]) { throw ('Unexpected refusal: '+$kind+' / '+$caught) }
        if ($kind -eq 'publication-failure' -and $cause -isnot [IO.IOException] -and $cause -isnot [UnauthorizedAccessException]) { throw ('Expected a publication I/O failure: '+$caught) }
        if ((Get-FileHash -LiteralPath $oldFile).Hash -ne $oldHash) { throw ('Previous artifact changed: '+$kind) }
        if ((Get-FileHash -LiteralPath $newFile).Hash -ne $sourceHashBefore) { throw ('Frozen source changed: '+$kind) }
        if (@(Get-ChildItem -LiteralPath (Split-Path $target) -Filter '*.staging-*' -File).Count) { throw ('Owned staging file leaked: '+$kind) }
        $preflight=$kind -in @('directory-link','source-mismatch','input-mismatch')
        if ($preflight -and ((Test-Path -LiteralPath $output) -or (Get-FileHash -LiteralPath $target).Hash -ne $oldHash)) { throw ('Preflight changed output: '+$kind) }
        if ($kind -eq 'publication-failure' -and [IO.File]::ReadAllText((Join-Path $target retained.txt)) -ne 'retained directory') { throw 'Publication failure changed the existing directory' }
        if (($success -or $kind -eq 'smoke-failure') -and (Get-FileHash -LiteralPath $target).Hash -ne $newHash) { throw ('Root asset differs: '+$kind) }
        $evidence=Join-Path $output evidence.json
        if ($success) {
            $report=Get-Content -LiteralPath $evidence -Raw | ConvertFrom-Json
            if ($report.AssetsVerified -ne 1 -or $report.Rebuilt -or $report.RepeatedSuite -or @($report.Smoke).Count -ne 1) { throw 'Transfer evidence differs' }
        }
        elseif (Test-Path -LiteralPath $evidence) { throw 'Failed transfer published completed evidence' }
        $rows+=@{Case=$kind;Succeeded=$success;Error=$caught;PreviousArtifactPreserved=$true;FrozenSourcePreserved=$true;OwnedStagingCleaned=$true;Preflight=$preflight}
    }
} finally { Remove-Item Function:\dotnet }
Write-Json (Join-Path $root evidence.json) @{Checks=$rows;Builds=0;RealConversions=0;RunnerInvocations=0;ToolSha256=(Get-FileHash -LiteralPath $TransferScript).Hash}
'Transfer checks pass: two successful ownership cases, five publication/smoke/preflight refusals; no builds or real conversions.'
