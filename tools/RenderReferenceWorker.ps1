# Isolated Office reference-render worker. Invoked ONLY by the
# RenderReference.ps1 supervisor, which enforces the total deadline and owns
# this process. This worker must never kill processes: if COM hangs inside
# cleanup, the supervisor terminates this (owned) process instead.
#
# Contract (all paths):
#   - appends "stage:<name>" lines to -ProgressLog before each operation,
#   - always writes -StatusPath JSON: ok | export-failed | unavailable,
#   - on success leaves reference.pdf plus DPI raster page-*.png in -WorkDirectory.

param(
    [ValidateSet('net10.0','net11.0')] [string] $Framework = $(if ($env:OOXPDF_TARGET_FRAMEWORK) {$env:OOXPDF_TARGET_FRAMEWORK} else {'net11.0'}),
    [Parameter(Mandatory = $true)]
    [string] $InputPath,

    [Parameter(Mandatory = $true)]
    [string] $WorkDirectory,

    [int] $Dpi = 144,

    [Parameter(Mandatory = $true)]
    [string] $ProgressLog,

    [Parameter(Mandatory = $true)]
    [string] $StatusPath,

    [switch] $SkipBuild
)

$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot 'OfficeReferenceOwnership.ps1')

function Write-Stage([string] $Name) {
    Add-Content -LiteralPath $ProgressLog -Value ("stage:{0} {1}" -f $Name, ([DateTime]::UtcNow.ToString("O", [Globalization.CultureInfo]::InvariantCulture)))
}

function Write-Status([string] $Status, [string] $Stage, [string] $OfficeApp, [string] $OfficeVersion, [string] $ExportSettings, [string] $ErrorMessage) {
    [ordered]@{
        Status = $Status
        Stage = $Stage
        OfficeApp = $OfficeApp
        OfficeVersion = $OfficeVersion
        ExportSettings = $ExportSettings
        Ownership = $script:ownership
        Error = $ErrorMessage
    } | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $StatusPath -Encoding UTF8
}

function Release-ComObject($value) {
    if ($null -ne $value -and [System.Runtime.InteropServices.Marshal]::IsComObject($value)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)
    }
}

function Complete-ComCleanup() {
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$inputFull = (Resolve-Path -LiteralPath $InputPath).Path
New-Item -ItemType Directory -Force -Path $WorkDirectory | Out-Null
$workFull = (Resolve-Path -LiteralPath $WorkDirectory).Path
$extension = [System.IO.Path]::GetExtension($inputFull).ToLowerInvariant()
$stage = "starting"
$officeApp = $null
$officeVersion = $null
$exportSettings = $null
$ownership = [ordered]@{InputCopy=$null;DocumentOwned=$false;ProcessId=$null;ProcessOwned=$false;QuitCalled=$false}
Write-Stage $stage

try {
    # Opening a private copy establishes document ownership even when Office
    # activation attaches to an existing interactive application.
    if ($extension -in @('.pptx','.docx')) {
        $copyDirectory = New-Item -ItemType Directory -Path (Join-Path $workFull 'reference-input')
        $referenceInput = Join-Path $copyDirectory.FullName ([IO.Path]::GetFileName($inputFull))
        Copy-Item -LiteralPath $inputFull -Destination $referenceInput
        if ((Get-FileHash $inputFull).Hash -ne (Get-FileHash $referenceInput).Hash) { throw 'Office input copy differs' }
        $ownership.InputCopy = $referenceInput
    }
    if ($extension -eq ".pptx") {
        $officeApp = "PowerPoint"
        $exportSettings = "powerpoint-saveas-ppSaveAsPDF(32)"
        $powerPoint = $null
        $presentation = $null
        $referencePdf = Join-Path $workFull "reference.pdf"
        $activationStarted = [DateTime]::UtcNow
        $existingProcesses = @(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue | ForEach-Object Id)
        try {
            $stage = "activation"
            Write-Stage $stage
            try {
                $powerPoint = New-Object -ComObject PowerPoint.Application
            }
            catch {
                throw "Office COM activation failed (PowerPoint.Application): $($_.Exception.Message)"
            }

            $officeVersion = [string]$powerPoint.Version
            $officeProcess = Get-OfficeWindowProcess $powerPoint.HWND
            $ownership.ProcessId = $officeProcess.Id
            $ownership.ProcessOwned = Test-OfficeProcessOwned $officeProcess $existingProcesses $activationStarted 'POWERPNT'
            if ($ownership.ProcessOwned) { $powerPoint.DisplayAlerts = 1 }
            $stage = "open"
            Write-Stage $stage
            $presentation = $powerPoint.Presentations.Open($referenceInput, $true, $false, $false)
            $ownership.DocumentOwned = Test-OfficeDocumentOwned $presentation $referenceInput
            if (!$ownership.DocumentOwned) { throw 'Opened PowerPoint document is not the owned input copy; leaving it open' }
            if (!$officeProcess) {
                $officeProcess = Get-OfficeWindowProcess $powerPoint.HWND
                $ownership.ProcessId = $officeProcess.Id
                $ownership.ProcessOwned = Test-OfficeProcessOwned $officeProcess $existingProcesses $activationStarted 'POWERPNT'
            }
            $stage = "export"
            Write-Stage $stage
            $presentation.SaveAs($referencePdf, 32)
        }
        finally {
            $operationStage = $stage
            $stage = "cleanup"
            Write-Stage $stage
            try {
                if ($ownership.DocumentOwned) { $presentation.Close() }
            }
            finally {
                try {
                    if ($ownership.ProcessOwned -and $powerPoint.Presentations.Count -eq 0) {
                        $ownership.QuitCalled = $true
                        $powerPoint.Quit()
                    }
                }
                finally {
                    Release-ComObject $presentation
                    Release-ComObject $powerPoint
                    Complete-ComCleanup
                }
            }

            # A completed cleanup must not replace the stage of a pending failure.
            # If cleanup itself throws, execution never restores the operation stage.
            $stage = $operationStage
        }

        $stage = "rasterize"
        Write-Stage $stage
        & (Join-Path $PSScriptRoot "RasterizePdf.ps1") -Framework $Framework -InputPdf $referencePdf -OutputDirectory $workFull -Dpi $Dpi -SkipBuild:$SkipBuild -RasterizerPath $env:OOXPDF_REFERENCE_RASTERIZER
        $stage = "done"
        Write-Stage $stage
        Write-Status "ok" $stage $officeApp $officeVersion $exportSettings ""
        return
    }

    if ($extension -eq ".docx") {
        $officeApp = "Word"
        $exportSettings = "word-saveas2-wdFormatPDF(17)"
        $word = $null
        $document = $null
        $referencePdf = Join-Path $workFull "reference.pdf"
        $activationStarted = [DateTime]::UtcNow
        $existingProcesses = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
        try {
            $stage = "activation"
            Write-Stage $stage
            try {
                $word = New-Object -ComObject Word.Application
            }
            catch {
                throw "Office COM activation failed (Word.Application): $($_.Exception.Message)"
            }

            $officeVersion = [string]$word.Version
            $stage = "open"
            Write-Stage $stage
            $missing = [Type]::Missing
            $document = $word.Documents.OpenNoRepairDialog($referenceInput, $false, $true, $false,
                $missing, $missing, $false, $missing, $missing, $missing, $missing, $false)
            $ownership.DocumentOwned = Test-OfficeDocumentOwned $document $referenceInput
            if (!$ownership.DocumentOwned) { throw 'Opened Word document is not the owned input copy; leaving it open' }
            $officeProcess = Get-OfficeWindowProcess $document.ActiveWindow.Hwnd
            $ownership.ProcessId = $officeProcess.Id
            $ownership.ProcessOwned = Test-OfficeProcessOwned $officeProcess $existingProcesses $activationStarted 'WINWORD'
            if ($ownership.ProcessOwned) { $word.Visible = $false; $word.DisplayAlerts = 0 }
            $stage = "export"
            Write-Stage $stage
            $document.SaveAs2($referencePdf, 17)
        }
        finally {
            $operationStage = $stage
            $stage = "cleanup"
            Write-Stage $stage
            try {
                if ($ownership.DocumentOwned) { $document.Close($false) }
            }
            finally {
                try {
                    if ($ownership.ProcessOwned -and $word.Documents.Count -eq 0) {
                        $ownership.QuitCalled = $true
                        $word.Quit()
                    }
                }
                finally {
                    Release-ComObject $document
                    Release-ComObject $word
                    Complete-ComCleanup
                }
            }

            $stage = $operationStage
        }

        $stage = "rasterize"
        Write-Stage $stage
        & (Join-Path $PSScriptRoot "RasterizePdf.ps1") -Framework $Framework -InputPdf $referencePdf -OutputDirectory $workFull -Dpi $Dpi -SkipBuild:$SkipBuild -RasterizerPath $env:OOXPDF_REFERENCE_RASTERIZER
        $stage = "done"
        Write-Stage $stage
        Write-Status "ok" $stage $officeApp $officeVersion $exportSettings ""
        return
    }

    throw "Unsupported reference input extension '$extension'. Expected .pptx or .docx."
}
catch {
    $message = $_.Exception.Message
    if ($message -like "Office COM activation failed*") {
        Write-Status "unavailable" $stage $officeApp $officeVersion $exportSettings $message
    }
    else {
        Write-Status "export-failed" $stage $officeApp $officeVersion $exportSettings $message
    }

    throw
}
