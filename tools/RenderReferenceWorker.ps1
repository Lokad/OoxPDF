# Isolated Office reference-render worker. Invoked ONLY by the
# RenderReference.ps1 supervisor, which enforces the total deadline and owns
# this process. This worker must never kill processes: if COM hangs inside
# cleanup, the supervisor terminates this (owned) process instead.
#
# Contract (all paths):
#   - appends "stage:<name>" lines to -ProgressLog as stages complete,
#   - always writes -StatusPath JSON: ok | export-failed | unavailable,
#   - on success leaves reference.pdf plus DPI raster page-*.png in -WorkDirectory.

param(
    [Parameter(Mandatory = $true)]
    [string] $InputPath,

    [Parameter(Mandatory = $true)]
    [string] $WorkDirectory,

    [int] $Dpi = 144,

    [Parameter(Mandatory = $true)]
    [string] $ProgressLog,

    [Parameter(Mandatory = $true)]
    [string] $StatusPath
)

$ErrorActionPreference = "Stop"

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
Write-Stage $stage

try {
    if ($extension -eq ".pptx") {
        $officeApp = "PowerPoint"
        $exportSettings = "powerpoint-saveas-ppSaveAsPDF(32)"
        $powerPoint = $null
        $presentation = $null
        $referencePdf = Join-Path $workFull "reference.pdf"
        try {
            $stage = "activation"
            try {
                $powerPoint = New-Object -ComObject PowerPoint.Application
            }
            catch {
                throw "Office COM activation failed (PowerPoint.Application): $($_.Exception.Message)"
            }

            Write-Stage $stage
            $officeVersion = [string]$powerPoint.Version
            $powerPoint.DisplayAlerts = 1
            $stage = "open"
            $presentation = $powerPoint.Presentations.Open($inputFull, $true, $true, $false)
            Write-Stage $stage
            $stage = "export"
            $presentation.SaveAs($referencePdf, 32)
            Write-Stage $stage
        }
        finally {
            $stage = "cleanup"
            try {
                if ($presentation -ne $null) { $presentation.Close() }
            }
            finally {
                try {
                    if ($powerPoint -ne $null) { $powerPoint.Quit() }
                }
                finally {
                    Release-ComObject $presentation
                    Release-ComObject $powerPoint
                    Complete-ComCleanup
                }
            }

            Write-Stage $stage
        }

        $stage = "rasterize"
        & (Join-Path $PSScriptRoot "RasterizePdf.ps1") -InputPdf $referencePdf -OutputDirectory $workFull -Dpi $Dpi
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
        try {
            $stage = "activation"
            try {
                $word = New-Object -ComObject Word.Application
            }
            catch {
                throw "Office COM activation failed (Word.Application): $($_.Exception.Message)"
            }

            Write-Stage $stage
            $officeVersion = [string]$word.Version
            $word.Visible = $false
            $word.DisplayAlerts = 0
            $stage = "open"
            $document = $word.Documents.OpenNoRepairDialog($inputFull, $false, $true, $false)
            Write-Stage $stage
            $stage = "export"
            $document.SaveAs2($referencePdf, 17)
            Write-Stage $stage
        }
        finally {
            $stage = "cleanup"
            try {
                if ($document -ne $null) { $document.Close($false) }
            }
            finally {
                try {
                    if ($word -ne $null) { $word.Quit() }
                }
                finally {
                    Release-ComObject $document
                    Release-ComObject $word
                    Complete-ComCleanup
                }
            }

            Write-Stage $stage
        }

        $stage = "rasterize"
        & (Join-Path $PSScriptRoot "RasterizePdf.ps1") -InputPdf $referencePdf -OutputDirectory $workFull -Dpi $Dpi
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
