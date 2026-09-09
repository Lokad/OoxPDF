# Supervised Office reference renderer (R01).
#
# Office COM calls can hang inside activation/open/export/close/quit, where no
# in-process finally can recover. This supervisor therefore runs the real work
# in an isolated worker process (RenderReferenceWorker.ps1, overridable via
# -WorkerScript for fake-worker tests) under a total deadline, with staged
# progress, separate logs, and structured status results.
#
# Ownership rule: the supervisor kills ONLY the worker process it started
# (owned PID). It never snapshots or kills WINWORD/POWERPNT processes, so
# unrelated applications launched during the run are left untouched.
# Concurrent supervised renders are serialized with a machine-wide mutex.
#
# Output layout on success (unchanged): reference.pdf plus raster page-*.png
# in -OutputDirectory. A reference-status.json is always written there.

param(
    [Parameter(Mandatory = $true)]
    [string] $InputPath,

    [Parameter(Mandatory = $true)]
    [string] $OutputDirectory,

    [int] $Dpi = 144,

    [int] $TimeoutSeconds = 600,

    [string] $WorkerScript
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($WorkerScript)) {
    $WorkerScript = Join-Path $PSScriptRoot "RenderReferenceWorker.ps1"
}

$inputFull = (Resolve-Path -LiteralPath $InputPath).Path
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$outputFull = (Resolve-Path -LiteralPath $OutputDirectory).Path
$statusPath = Join-Path $outputFull "reference-status.json"
$supervisorLog = Join-Path $outputFull "reference-supervisor.log"
$workerLog = Join-Path $outputFull "reference-worker.log"
$progressLog = Join-Path $outputFull "reference-progress.log"

function Write-SupervisorLog([string] $Message) {
    Add-Content -LiteralPath $supervisorLog -Value ("[{0}] {1}" -f ([DateTime]::UtcNow.ToString("O", [Globalization.CultureInfo]::InvariantCulture)), $Message)
}

function Write-OutputStatus([string] $Status, [string] $Stage, [string] $ErrorMessage, [bool] $OfficeOrphanPossible) {
    [ordered]@{
        Status = $Status
        Stage = $Stage
        Error = $ErrorMessage
        OfficeOrphanPossible = $OfficeOrphanPossible
        TimeoutSeconds = $TimeoutSeconds
    } | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $statusPath -Encoding UTF8
}

function Get-LastWorkerStage([string] $LogPath) {
    if (-not (Test-Path -LiteralPath $LogPath)) {
        return "starting"
    }

    $last = Get-Content -LiteralPath $LogPath | Select-Object -Last 1
    $match = [regex]::Match([string]$last, '^stage:(\S+)')
    if ($match.Success) {
        return $match.Groups[1].Value
    }

    return "starting"
}

function Get-WorkerStatus {
    $workerStatusPath = Join-Path $workerWorkDir "reference-status.json"
    if (-not (Test-Path -LiteralPath $workerStatusPath)) {
        return $null
    }

    try {
        return Get-Content -Raw -LiteralPath $workerStatusPath | ConvertFrom-Json
    }
    catch {
        return $null
    }
}

"supervisor-start input=$inputFull dpi=$Dpi timeoutSeconds=$TimeoutSeconds" | Set-Content -LiteralPath $supervisorLog
$deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
$mutex = New-Object System.Threading.Mutex($false, "Global\OoxPdfOfficeRender")
$mutexAcquired = $false
$workerWorkDir = Join-Path $outputFull ("_worker-" + [System.Guid]::NewGuid().ToString("N"))
try {
    $mutexWaitMs = [Math]::Max(0, [int](($deadline - [DateTime]::UtcNow).TotalMilliseconds))
    Write-SupervisorLog "waiting for office-render mutex"
    $mutexAcquired = $mutex.WaitOne($mutexWaitMs)
    if (-not $mutexAcquired) {
        Write-OutputStatus "unavailable" "mutex" "Timed out waiting for another supervised Office render to finish." $false
        throw "Office reference rendering is unavailable: another supervised render holds the worker lock past the deadline."
    }

    if ([DateTime]::UtcNow -ge $deadline) {
        Write-OutputStatus "timeout" "mutex" "Deadline expired while waiting for the render lock." $false
        throw "Office reference rendering timed out waiting for the render lock."
    }

    New-Item -ItemType Directory -Force -Path $workerWorkDir | Out-Null
    $workerStatusPath = Join-Path $workerWorkDir "reference-status.json"
    $workerProgressLog = Join-Path $workerWorkDir "reference-progress.log"
    $shell = (Get-Command pwsh -ErrorAction SilentlyContinue | Select-Object -First 1)
    $shellPath = if ($null -ne $shell) { $shell.Source } else { "powershell" }
    $workerArgs = @("-NoProfile", "-File", $WorkerScript, "-InputPath", $inputFull, "-WorkDirectory", $workerWorkDir, "-Dpi", "$Dpi", "-ProgressLog", $workerProgressLog, "-StatusPath", $workerStatusPath)
    Write-SupervisorLog "starting owned worker: $shellPath"
    $worker = Start-Process -FilePath $shellPath -ArgumentList $workerArgs -PassThru -WindowStyle Hidden -RedirectStandardOutput (Join-Path $workerWorkDir "worker-stdout.log") -RedirectStandardError (Join-Path $workerWorkDir "worker-stderr.log")
    Write-SupervisorLog ("owned worker pid={0}" -f $worker.Id)
    try {
        while (-not $worker.HasExited -and [DateTime]::UtcNow -lt $deadline) {
            Start-Sleep -Milliseconds 250
        }

        if (-not $worker.HasExited) {
            Stop-Process -Id $worker.Id -Force -ErrorAction SilentlyContinue
            $worker.WaitForExit(10000)
            $stage = Get-LastWorkerStage $workerProgressLog
            # Post-mortem grace poll: the kill can land while final stage writes are
            # still in flight, so re-read briefly before concluding "starting".
            $stageWaitMs = 0
            while ($stage -eq "starting" -and $stageWaitMs -lt 5000) {
                Start-Sleep -Milliseconds 250
                $stageWaitMs += 250
                $stage = Get-LastWorkerStage $workerProgressLog
            }
            Copy-Item -LiteralPath $workerProgressLog -Destination $progressLog -Force -ErrorAction SilentlyContinue
            Write-OutputStatus "timeout" $stage ("Office reference worker (owned pid {0}) exceeded the {1}s deadline during stage '{2}' and was terminated. An Office process started by the hung worker may have been orphaned; no other processes were touched." -f $worker.Id, $TimeoutSeconds, $stage) $true
            throw "Office reference rendering timed out during stage '$stage' after ${TimeoutSeconds}s; owned worker terminated."
        }
    }
    finally {
        if (-not $worker.HasExited) {
            Stop-Process -Id $worker.Id -Force -ErrorAction SilentlyContinue
        }
    }

    Copy-Item -LiteralPath $workerProgressLog -Destination $progressLog -Force -ErrorAction SilentlyContinue
    $workerStatus = Get-WorkerStatus
    if ($worker.ExitCode -ne 0 -or $null -eq $workerStatus -or [string]$workerStatus.Status -ne "ok") {
        $stage = if ($null -ne $workerStatus -and -not [string]::IsNullOrWhiteSpace([string]$workerStatus.Stage)) { [string]$workerStatus.Stage } else { Get-LastWorkerStage $progressLog }
        $detail = if ($null -ne $workerStatus -and -not [string]::IsNullOrWhiteSpace([string]$workerStatus.Error)) { [string]$workerStatus.Error } else { "worker exit code $($worker.ExitCode)" }
        $status = if ($null -ne $workerStatus -and [string]$workerStatus.Status -eq "unavailable") { "unavailable" } else { "export-failed" }
        Write-OutputStatus $status $stage $detail $false
        throw "Office reference rendering $status during stage '$stage': $detail"
    }

    $stagedPdf = Join-Path $workerWorkDir "reference.pdf"
    if (-not (Test-Path -LiteralPath $stagedPdf) -or (Get-Item -LiteralPath $stagedPdf).Length -lt 8) {
        Write-OutputStatus "export-failed" "verify" "Worker reported success but reference.pdf is missing or truncated." $false
        throw "Office reference rendering failed verification: reference.pdf is missing or truncated."
    }

    foreach ($item in @(Get-ChildItem -LiteralPath $workerWorkDir -Force | Where-Object { $_.Name -notlike "_*" -and $_.Name -ne "worker-stdout.log" -and $_.Name -ne "worker-stderr.log" -and $_.Name -ne "reference-progress.log" -and $_.Name -ne "reference-status.json" })) {
        Copy-Item -LiteralPath $item.FullName -Destination (Join-Path $outputFull $item.Name) -Recurse -Force
    }

    Copy-Item -LiteralPath $workerStatusPath -Destination (Join-Path $outputFull "reference-worker-status.json") -Force
    Write-OutputStatus "ok" "done" "" $false
    Write-SupervisorLog "success"
}
finally {
    if ($mutexAcquired) {
        $mutex.ReleaseMutex()
    }

    $mutex.Dispose()
    if (Test-Path -LiteralPath $workerWorkDir) {
        Remove-Item -LiteralPath $workerWorkDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}
