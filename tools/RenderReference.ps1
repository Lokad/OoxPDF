param(
    [Parameter(Mandatory = $true)]
    [string] $InputPath,

    [Parameter(Mandatory = $true)]
    [string] $OutputDirectory,

    [int] $Dpi = 144
)

$ErrorActionPreference = "Stop"

function Release-ComObject($value) {
    if ($null -ne $value -and [System.Runtime.InteropServices.Marshal]::IsComObject($value)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)
    }
}

function Get-OfficeProcessIds([string] $Name) {
    @(Get-Process -Name $Name -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
}

function Stop-NewOfficeProcesses([string] $Name, [int[]] $BeforeIds) {
    $known = @{}
    foreach ($id in $BeforeIds) {
        $known[$id] = $true
    }

    Get-Process -Name $Name -ErrorAction SilentlyContinue |
        Where-Object { -not $known.ContainsKey($_.Id) } |
        ForEach-Object {
            Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
        }
}

function Complete-ComCleanup() {
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$inputFull = (Resolve-Path -LiteralPath $InputPath).Path
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$outputFull = (Resolve-Path -LiteralPath $OutputDirectory).Path
$extension = [System.IO.Path]::GetExtension($inputFull).ToLowerInvariant()

if ($extension -eq ".pptx") {
    $beforePowerPointIds = Get-OfficeProcessIds "POWERPNT"
    $powerPoint = $null
    $presentation = $null
    $referencePdf = Join-Path $outputFull "reference.pdf"
    try {
        $powerPoint = New-Object -ComObject PowerPoint.Application
        $powerPoint.DisplayAlerts = 1
        $presentation = $powerPoint.Presentations.Open($inputFull, $true, $true, $false)
        $presentation.SaveAs($referencePdf, 32)
    }
    finally {
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
                Stop-NewOfficeProcesses "POWERPNT" $beforePowerPointIds
            }
        }
    }

    & (Join-Path $PSScriptRoot "RasterizePdf.ps1") -InputPdf $referencePdf -OutputDirectory $outputFull -Dpi $Dpi
    return
}

if ($extension -eq ".docx") {
    $beforeWordIds = Get-OfficeProcessIds "WINWORD"
    $word = $null
    $document = $null
    $referencePdf = Join-Path $outputFull "reference.pdf"
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $document = $word.Documents.OpenNoRepairDialog($inputFull, $false, $true, $false)
        $document.SaveAs2($referencePdf, 17)
    }
    finally {
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
                Stop-NewOfficeProcesses "WINWORD" $beforeWordIds
            }
        }
    }

    & (Join-Path $PSScriptRoot "RasterizePdf.ps1") -InputPdf $referencePdf -OutputDirectory $outputFull -Dpi $Dpi
    return
}

throw "Unsupported reference input extension '$extension'. Expected .pptx or .docx."
