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

$inputFull = (Resolve-Path -LiteralPath $InputPath).Path
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$outputFull = (Resolve-Path -LiteralPath $OutputDirectory).Path
$extension = [System.IO.Path]::GetExtension($inputFull).ToLowerInvariant()

if ($extension -eq ".pptx") {
    $powerPoint = $null
    $presentation = $null
    $referencePdf = Join-Path $outputFull "reference.pdf"
    try {
        $powerPoint = New-Object -ComObject PowerPoint.Application
        $presentation = $powerPoint.Presentations.Open($inputFull, $true, $true, $false)
        $presentation.SaveAs($referencePdf, 32)
    }
    finally {
        if ($presentation -ne $null) { $presentation.Close() }
        if ($powerPoint -ne $null) { $powerPoint.Quit() }
        Release-ComObject $presentation
        Release-ComObject $powerPoint
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }

    & (Join-Path $PSScriptRoot "RasterizePdf.ps1") -InputPdf $referencePdf -OutputDirectory $outputFull -Dpi $Dpi
    return
}

if ($extension -eq ".docx") {
    $word = $null
    $document = $null
    $referencePdf = Join-Path $outputFull "reference.pdf"
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $document = $word.Documents.Open($inputFull, $false, $true)
        $document.ExportAsFixedFormat($referencePdf, 17)
    }
    finally {
        if ($document -ne $null) { $document.Close($false) }
        if ($word -ne $null) { $word.Quit() }
        Release-ComObject $document
        Release-ComObject $word
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }

    & (Join-Path $PSScriptRoot "RasterizePdf.ps1") -InputPdf $referencePdf -OutputDirectory $outputFull -Dpi $Dpi
    return
}

throw "Unsupported reference input extension '$extension'. Expected .pptx or .docx."
