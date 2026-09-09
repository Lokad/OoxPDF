# Acceptance checks for R01 (supervised Office rendering), R02 (trusted
# reference reuse), and the T03 build/run binding used by visual workflows.
# Runs without Office and without network: Office COM paths are exercised via
# fake workers, and reference PDFs stand in for Office exports for cache
# mechanics (structure, hashing, atomicity), never as fidelity references.

param(
    [int] $RenderTimeoutSeconds = 600
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $repoRoot "tools/ReferenceCache.ps1")

$scratch = Join-Path $repoRoot "artifacts/workflow-test"
if (Test-Path -LiteralPath $scratch) {
    Remove-Item -LiteralPath $scratch -Recurse -Force
}

New-Item -ItemType Directory -Force -Path $scratch | Out-Null
$failures = @()
function Assert-True([bool] $Condition, [string] $Message) {
    if (-not $Condition) {
        $script:failures += $Message
        Write-Host ("FAIL: {0}" -f $Message)
    }
    else {
        Write-Host ("PASS: {0}" -f $Message)
    }
}

function New-ScratchDocx([string] $Name) {
    # Per-run unique bytes so cache identities never collide across runs.
    $Name = [System.IO.Path]::GetFileNameWithoutExtension($Name) + "-p$PID" + [System.IO.Path]::GetExtension($Name)
    $source = Join-Path $repoRoot "tests/Lokad.OoxPdf.Tests/Cases/docx-blank.docx"
    $dest = Join-Path $scratch $Name
    Copy-Item -LiteralPath $source -Destination $dest -Force
    Add-Content -LiteralPath $dest -Value ("#" + $Name) -Encoding Ascii -NoNewline
    return $dest
}

function New-FakeWorker([string] $Name, [string[]] $BodyLines) {
    $path = Join-Path $scratch $Name
    $header = @(
        'param([string] $InputPath, [string] $WorkDirectory, [int] $Dpi = 144, [string] $ProgressLog, [string] $StatusPath)',
        '$ErrorActionPreference = "Stop"',
        'function Write-Stage([string] $Name) {',
        '    Add-Content -LiteralPath $ProgressLog -Value ("stage:{0}" -f $Name)',
        '}'
    )
    Set-Content -LiteralPath $path -Value (($header + $BodyLines) -join "`r`n") -NoNewline
    return $path
}

# T03: the shared build gate records identities instead of trusting mtimes.
$cliDll = Join-Path $repoRoot "src/Lokad.OoxPdf.Cli/bin/Debug/net10.0/Lokad.OoxPdf.Cli.dll"
& (Join-Path $repoRoot "tools/EnsureDotnetBuild.ps1") -Project (Join-Path $repoRoot "src/Lokad.OoxPdf.Cli/Lokad.OoxPdf.Cli.csproj") -OutputDll $cliDll -Description "CLI" -RecordPath (Join-Path $scratch "cli-build-info.json")
$buildInfo = Get-Content -Raw -LiteralPath (Join-Path $scratch "cli-build-info.json") | ConvertFrom-Json
Assert-True ((Test-Path -LiteralPath $cliDll) -and $buildInfo.OutputSha256 -eq (Get-FileHash -LiteralPath $cliDll -Algorithm SHA256).Hash.ToLowerInvariant()) "T03 shared build records the real CLI binary hash"
Assert-True (-not [string]::IsNullOrWhiteSpace([string]$buildInfo.GitCommit)) "T03 shared build records the source commit"

# A real (locally rendered, non-Office) PDF stands in for Office exports below.
$standinDocx = New-ScratchDocx "standin.docx"
$standinPdf = Join-Path $scratch "standin-reference.pdf"
& dotnet $cliDll convert $standinDocx $standinPdf | Out-Null
Assert-True ((Test-Path -LiteralPath $standinPdf) -and (Get-Item -LiteralPath $standinPdf).Length -gt 1000) "stand-in reference PDF renders locally"

# R02: cache miss in CacheOnly mode is actionable, not silent.
$missDocx = New-ScratchDocx "miss.docx"
$missed = $false
try {
    & (Join-Path $repoRoot "tools/RenderCachedReference.ps1") -InputPath $missDocx -OutputDirectory (Join-Path $scratch "miss-out") -Dpi 144 -CacheOnly | Out-Null
}
catch {
    $missed = ($_.Exception.Message -like "*cache miss*") -and ($_.Exception.Message -like "*RenderCachedReference.ps1*")
}

Assert-True $missed "R02 cache miss throws an actionable populate command"

# R02: trusted import publishes a verified v2 entry.
& (Join-Path $repoRoot "tools/ImportOfficeReferenceCache.ps1") -InputPath $missDocx -ReferencePdf $standinPdf -Dpi 144 -CaseId "workflow-test" | Out-Null
$identityKey = Get-ReferenceIdentityKey $missDocx ""
$entryDir = Join-Path (Get-ReferenceCacheRoot) $identityKey
Assert-True ((Test-Path -LiteralPath (Join-Path $entryDir "complete.txt")) -and (Test-Path -LiteralPath (Join-Path $entryDir "reference-metadata.json"))) "R02 import publishes identity metadata plus completion marker"
$entryCheck = Test-ReferenceIdentityEntry $entryDir ((Get-FileHash -LiteralPath $missDocx -Algorithm SHA256).Hash.ToLowerInvariant()) ""
Assert-True ($entryCheck.State -eq "Complete") "R02 published entry verifies as Complete"

# R02: hit serves the PDF identity plus DPI derivatives with provenance.
& (Join-Path $repoRoot "tools/RenderCachedReference.ps1") -InputPath $missDocx -OutputDirectory (Join-Path $scratch "hit-out") -Dpi 144 -CacheOnly | Out-Null
$servedPdf = Join-Path $scratch "hit-out/reference.pdf"
$servedPages = @(Get-ChildItem -LiteralPath (Join-Path $scratch "hit-out") -Filter "page-*.png")
Assert-True ((Test-Path -LiteralPath $servedPdf) -and ((Get-FileHash -LiteralPath $servedPdf -Algorithm SHA256).Hash -eq (Get-FileHash -LiteralPath (Join-Path $entryDir "reference.pdf") -Algorithm SHA256).Hash)) "R02 hit serves the identical reference PDF"
Assert-True ($servedPages.Count -ge 1 -and (Test-Path -LiteralPath (Join-Path $scratch "hit-out/reference-metadata.json"))) "R02 hit serves raster derivatives plus provenance metadata"

# R02: variant mismatch is a miss, not a false hit.
$variantMissed = $false
try {
    & (Join-Path $repoRoot "tools/RenderCachedReference.ps1") -InputPath $missDocx -OutputDirectory (Join-Path $scratch "variant-out") -Dpi 144 -CacheOnly -CacheVariant "docxMarkup=all;docxMarkupGeometry=preserve" | Out-Null
}
catch {
    $variantMissed = ($_.Exception.Message -like "*cache miss*")
}

Assert-True $variantMissed "R02 variant mismatch misses instead of sharing an entry"

# R02: a second DPI rasterizes offline from the cached PDF identity.
& (Join-Path $repoRoot "tools/RenderCachedReference.ps1") -InputPath $missDocx -OutputDirectory (Join-Path $scratch "hit72-out") -Dpi 72 -CacheOnly | Out-Null
Assert-True (@(Get-ChildItem -LiteralPath (Join-Path $scratch "hit72-out") -Filter "page-*.png").Count -ge 1) "R02 missing derivative rasterizes offline from the cached PDF"

# R02: corrupt entries fail loudly instead of serving partial output.
Add-Content -LiteralPath (Join-Path $entryDir "reference.pdf") -Value "corruption" -Encoding Ascii
$corruptFailed = $false
try {
    & (Join-Path $repoRoot "tools/RenderCachedReference.ps1") -InputPath $missDocx -OutputDirectory (Join-Path $scratch "corrupt-out") -Dpi 144 -CacheOnly | Out-Null
}
catch {
    $corruptFailed = ($_.Exception.Message -like "*corrupt*")
}

Assert-True $corruptFailed "R02 corrupt entry throws instead of serving"
& (Join-Path $repoRoot "tools/ImportOfficeReferenceCache.ps1") -InputPath $missDocx -ReferencePdf $standinPdf -Dpi 144 -Force | Out-Null

# R02: pre-v2 flat entries are adopted, not orphaned.
$legacyDocx = New-ScratchDocx "legacy.docx"
$legacyKey = Get-ReferenceCacheKey $legacyDocx 144 ""
$legacyDir = Join-Path (Get-ReferenceCacheRoot) $legacyKey
New-Item -ItemType Directory -Force -Path $legacyDir | Out-Null
Copy-Item -LiteralPath $standinPdf -Destination (Join-Path $legacyDir "reference.pdf") -Force
Copy-Item -LiteralPath (Join-Path $scratch "hit-out/page-001.png") -Destination (Join-Path $legacyDir "page-001.png") -Force
Set-Content -LiteralPath (Join-Path $legacyDir "complete.txt") -Value ("input={0}`ndpi=144`nvariant=`n" -f (Resolve-Path -LiteralPath $legacyDocx).Path) -NoNewline
& (Join-Path $repoRoot "tools/RenderCachedReference.ps1") -InputPath $legacyDocx -OutputDirectory (Join-Path $scratch "legacy-out") -Dpi 144 -CacheOnly | Out-Null
$newIdentityDir = Join-Path (Get-ReferenceCacheRoot) (Get-ReferenceIdentityKey $legacyDocx "")
Assert-True ((Test-Path -LiteralPath (Join-Path $scratch "legacy-out/reference.pdf")) -and (Test-Path -LiteralPath (Join-Path $newIdentityDir "reference-metadata.json")) -and -not (Test-Path -LiteralPath $legacyDir)) "R02 legacy flat entry is adopted into the v2 layout and served"

# R02: concurrent publishers converge on one valid entry.
$raceDocx = New-ScratchDocx "race.docx"
$raceScript = Join-Path $scratch "race-import.ps1"
Set-Content -LiteralPath $raceScript -Value ("& '" + (Join-Path $repoRoot "tools/ImportOfficeReferenceCache.ps1") + "' -InputPath '" + $raceDocx + "' -ReferencePdf '" + $standinPdf + "' -Dpi 144 -Force") -NoNewline
$race1 = Start-Process -FilePath "pwsh" -ArgumentList @("-NoProfile", "-File", $raceScript) -PassThru -WindowStyle Hidden
$race2 = Start-Process -FilePath "pwsh" -ArgumentList @("-NoProfile", "-File", $raceScript) -PassThru -WindowStyle Hidden
$race1.WaitForExit(120000)
$race2.WaitForExit(120000)
$raceDir = Join-Path (Get-ReferenceCacheRoot) (Get-ReferenceIdentityKey $raceDocx "")
$raceCheck = Test-ReferenceIdentityEntry $raceDir ((Get-FileHash -LiteralPath $raceDocx -Algorithm SHA256).Hash.ToLowerInvariant()) ""
Assert-True (($race1.ExitCode -eq 0) -and ($race2.ExitCode -eq 0) -and ($raceCheck.State -eq "Complete")) "R02 concurrent publishers converge on a valid entry"

# R01: a worker hung in export dies on the deadline; strangers survive.
$hangWorker = New-FakeWorker "fake-hang.ps1" @(
    'Write-Stage "activation"',
    'Write-Stage "open"',
    'Write-Stage "export"',
    'Start-Sleep -Seconds 120'
)
$sentinel = Start-Process -FilePath "pwsh" -ArgumentList @("-NoProfile", "-Command", "Start-Sleep -Seconds 120") -PassThru -WindowStyle Hidden
$hangOut = Join-Path $scratch "hang-out"
$sw = [System.Diagnostics.Stopwatch]::StartNew()
$timedOut = $false
try {
    & (Join-Path $repoRoot "tools/RenderReference.ps1") -InputPath $standinDocx -OutputDirectory $hangOut -Dpi 144 -TimeoutSeconds 25 -WorkerScript $hangWorker | Out-Null
}
catch {
    $timedOut = ($_.Exception.Message -like "*timed out*export*")
}

$sw.Stop()
$sentinelAlive = -not $sentinel.HasExited
Stop-Process -Id $sentinel.Id -Force -ErrorAction SilentlyContinue
$hangStatus = Get-Content -Raw -LiteralPath (Join-Path $hangOut "reference-status.json") | ConvertFrom-Json
Assert-True (($timedOut) -and (($sw.Elapsed.TotalSeconds -ge 24) -and ($sw.Elapsed.TotalSeconds -lt 90))) "R01 hung export terminates on the deadline"
Assert-True ($sentinelAlive -and [string]$hangStatus.Status -eq "timeout" -and [string]$hangStatus.Stage -eq "export") "R01 timeout leaves unrelated processes untouched with staged status"

# R01: activation failure surfaces as unavailable without hanging.
$deadWorker = New-FakeWorker "fake-dead.ps1" @(
    'Write-Stage "activation"',
    '[ordered]@{ Status = "unavailable"; Stage = "activation"; OfficeApp = "Word"; OfficeVersion = ""; ExportSettings = ""; Error = "Office COM activation failed (Word.Application): fake" } | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $StatusPath -Encoding UTF8',
    'exit 1'
)
$unavailable = $false
$sw2 = [System.Diagnostics.Stopwatch]::StartNew()
try {
    & (Join-Path $repoRoot "tools/RenderReference.ps1") -InputPath $standinDocx -OutputDirectory (Join-Path $scratch "dead-out") -Dpi 144 -TimeoutSeconds 60 -WorkerScript $deadWorker | Out-Null
}
catch {
    $unavailable = ($_.Exception.Message -like "*unavailable*")
}

$sw2.Stop()
Assert-True ($unavailable -and ($sw2.Elapsed.TotalSeconds -lt 60)) "R01 COM failure surfaces as structured unavailable"

# R01: a healthy worker publishes output plus status.
$okWorker = New-FakeWorker "fake-ok.ps1" @(
    'Write-Stage "activation"',
    'Write-Stage "open"',
    'Write-Stage "export"',
    'Copy-Item -LiteralPath $InputPath -Destination (Join-Path $WorkDirectory "reference.pdf") -Force',
    'Write-Stage "rasterize"',
    'Write-Stage "done"',
    '[ordered]@{ Status = "ok"; Stage = "done"; OfficeApp = "fake"; OfficeVersion = "0"; ExportSettings = "fake"; Error = "" } | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $StatusPath -Encoding UTF8'
)
Copy-Item -LiteralPath $standinPdf -Destination (Join-Path $scratch "ok-input.pdf") -Force
& (Join-Path $repoRoot "tools/RenderReference.ps1") -InputPath (Join-Path $scratch "ok-input.pdf") -OutputDirectory (Join-Path $scratch "ok-out") -Dpi 144 -TimeoutSeconds 60 -WorkerScript $okWorker | Out-Null
$okStatus = Get-Content -Raw -LiteralPath (Join-Path $scratch "ok-out/reference-status.json") | ConvertFrom-Json
Assert-True (([string]$okStatus.Status -eq "ok") -and (Test-Path -LiteralPath (Join-Path $scratch "ok-out/reference.pdf"))) "R01 healthy worker output is published with ok status"



# R01: a worker hung before its first stage still dies on the deadline.
$earlyHangWorker = New-FakeWorker "fake-early-hang.ps1" @(
    'Start-Sleep -Seconds 120'
)
$earlyOut = Join-Path $scratch "early-hang-out"
$sw3 = [System.Diagnostics.Stopwatch]::StartNew()
$earlyTimedOut = $false
try {
    & (Join-Path $repoRoot "tools/RenderReference.ps1") -InputPath $standinDocx -OutputDirectory $earlyOut -Dpi 144 -TimeoutSeconds 8 -WorkerScript $earlyHangWorker | Out-Null
}
catch {
    $earlyTimedOut = ($_.Exception.Message -like "*timed out*starting*")
}

$sw3.Stop()
$earlyStatus = Get-Content -Raw -LiteralPath (Join-Path $earlyOut "reference-status.json") | ConvertFrom-Json
Assert-True (($earlyTimedOut) -and (($sw3.Elapsed.TotalSeconds -ge 8) -and ($sw3.Elapsed.TotalSeconds -lt 45)) -and ([string]$earlyStatus.Stage -eq "starting")) "R01 pre-stage hang terminates on the deadline as starting"

# R02: an entry directory without a completion marker is a miss, not a hit.
$incompleteDocx = New-ScratchDocx "incomplete.docx"
$incompleteKey = Get-ReferenceIdentityKey $incompleteDocx ""
New-Item -ItemType Directory -Force -Path (Join-Path (Get-ReferenceCacheRoot) $incompleteKey) | Out-Null
Set-Content -LiteralPath (Join-Path (Join-Path (Get-ReferenceCacheRoot) $incompleteKey) "junk.txt") -Value "partial" -NoNewline
$incompleteMissed = $false
try {
    & (Join-Path $repoRoot "tools/RenderCachedReference.ps1") -InputPath $incompleteDocx -OutputDirectory (Join-Path $scratch "incomplete-out") -Dpi 144 -CacheOnly | Out-Null
}
catch {
    $incompleteMissed = ($_.Exception.Message -like "*cache miss*")
}

Assert-True $incompleteMissed "R02 incomplete entry without a completion marker misses"

# R02: the markup import wrapper delegates to the shared verified path.
$markupCase = Join-Path $repoRoot "visual-cases/cases/docx-markup-all/case.json"
& (Join-Path $repoRoot "tools/ImportDocxMarkupReferenceCache.ps1") -Case $markupCase -ReferencePdf $standinPdf -Dpi 144 -Force | Out-Null
$markupManifest = Get-Content -Raw -LiteralPath $markupCase | ConvertFrom-Json
$markupInput = (Resolve-Path -LiteralPath (Join-Path (Split-Path -Parent $markupCase) $markupManifest.input)).Path
$markupVariant = "docxMarkup=all;docxMarkupGeometry=preserve"
$markupKey = Get-ReferenceIdentityKey $markupInput $markupVariant
$markupDir = Join-Path (Get-ReferenceCacheRoot) $markupKey
$markupCheck = Test-ReferenceIdentityEntry $markupDir ((Get-FileHash -LiteralPath $markupInput -Algorithm SHA256).Hash.ToLowerInvariant()) $markupVariant
Assert-True (($markupCheck.State -eq "Complete") -and ([string]$markupCheck.Metadata.CaseId -eq "docx-markup-all")) "R02 markup import wrapper publishes a verified entry with case provenance"
& (Join-Path $repoRoot "tools/RenderCachedReference.ps1") -InputPath $markupInput -OutputDirectory (Join-Path $scratch "markup-hit-out") -Dpi 144 -CacheOnly -CacheVariant $markupVariant | Out-Null
Assert-True ((Test-Path -LiteralPath (Join-Path $scratch "markup-hit-out/reference.pdf")) -and (@(Get-ChildItem -LiteralPath (Join-Path $scratch "markup-hit-out") -Filter "page-*.png").Count -ge 1)) "R02 markup variant hit serves PDF plus derivatives"

# The reference-request bundle tracks the v2 identity layout.
$requestDir = Join-Path $scratch "request-out"
& (Join-Path $repoRoot "tools/NewDocxMarkupReferenceRequest.ps1") -Case $markupCase -OutputDirectory $requestDir | Out-Null
$requestSummary = Get-Content -Raw -LiteralPath (Join-Path $requestDir "summary.json") | ConvertFrom-Json
Assert-True (([int]$requestSummary.SelectedCaseCount -eq 1) -and ([string]$requestSummary.Items[0].ExpectedCacheDirectory -like "*-v2*") -and ([string]$requestSummary.Items[0].CacheKey -eq [string](Get-ReferenceIdentityKey $markupInput $markupVariant))) "Reference request bundle tracks the v2 identity key"

if ($script:failures.Count -ne 0) {
    throw ("Reference workflow checks failed: " + ($script:failures -join " | "))
}

Write-Host "All reference workflow checks passed."
