# Content-tokenizer checks for PdfInspect (T05): comments and inline-image bytes
# must never tokenize as operators, and nested/escaped literal strings must
# decode exactly. Builds adversarial PDFs, inspects them, and asserts the
# extracted operations. No Office/COM or reference cache needed.

param([string] $OutputDirectory, [string] $InspectorPath)
$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$scratch = if ($OutputDirectory) { [IO.Path]::GetFullPath($OutputDirectory) }
    else { Join-Path $repoRoot ('artifacts/content-tokenize-probe-' + [Guid]::NewGuid().ToString('N')) }
if (Test-Path -LiteralPath $scratch) { throw 'Use a fresh tokenizer-check directory' }
New-Item -ItemType Directory -Force -Path $scratch | Out-Null

$pdfLines = @(
'%PDF-1.7'
'1 0 obj'
'<< /Type /Catalog /Pages 2 0 R >>'
'endobj'
'2 0 obj'
'<< /Type /Pages /Kids [3 0 R] /Count 1 >>'
'endobj'
'3 0 obj'
'<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>'
'endobj'
'4 0 obj'
'<< /Length 1000 >>'
'stream'
'% probe comment: 999 888 l Q q fakeop'
'q'
'1 0 0 1 10 20 cm'
'0 0 100 50 re'
'f'
'Q'
'BI'
'/W 1'
'/H 1'
'/CS /G'
'/BPC 1'
'ID'
'10 20 l'
'EI'
'20 20 40 30 re'
'f*'
'/F1 12 Tf'
'% 1 0 0 rg is not an operator here'
'BT'
'(Hello) Tj'
'ET'
'q'
'0.2 g'
'BT'
'(a(b)c) Tj'
'ET'
'q'
'0 0 1 rg'
'BT'
'(a\(b) Tj'
'ET'
'Q'
'BT'
'(gray again) Tj'
'ET'
'Q'
'BT'
'(default again) Tj'
'ET'
'0 1 1 0 k'
'BT'
'(CMYK) Tj'
'(0 0 1 rg) Tj'
'ET'
'% (phantom) Tj'
'endstream'
'endobj'
'trailer'
'<< /Root 1 0 R >>'
)
$pdfPath = Join-Path $scratch "adversarial.pdf"
[IO.File]::WriteAllLines($pdfPath, $pdfLines)
$inspectDir = Join-Path $scratch "inspect"
if ($InspectorPath) { dotnet $InspectorPath $pdfPath $inspectDir | Out-Null }
else { & (Join-Path $repoRoot "tools/InspectPdf.ps1") -InputPdf $pdfPath -OutputDirectory $inspectDir | Out-Null }
if ($LASTEXITCODE -ne 0) {
    throw "PdfInspect failed with exit code $LASTEXITCODE."
}

$failures = @()
function Assert-Near([string] $Name, [double] $Actual, [double] $Expected) {
    if ([Math]::Abs($Actual - $Expected) -gt 0.001d) {
        $script:failures += ("{0}: got {1}, want {2}" -f $Name, $Actual, $Expected)
        Write-Host ("FAIL: {0}" -f $script:failures[-1])
    }
    else {
        Write-Host ("PASS: {0} = {1}" -f $Name, $Actual)
    }
}
$graphics = Get-Content -Raw -LiteralPath (Join-Path $inspectDir "graphics-operations.json") | ConvertFrom-Json
Assert-Near "graphics-count" $graphics.Count 2
if ($graphics.Count -eq 2) {
    Assert-Near "op0-minx" $graphics[0].MinX 10
    Assert-Near "op0-miny" $graphics[0].MinY 20
    Assert-Near "op0-maxx" $graphics[0].MaxX 110
    Assert-Near "op0-maxy" $graphics[0].MaxY 70
    Assert-Near "op1-minx" $graphics[1].MinX 20
    Assert-Near "op1-miny" $graphics[1].MinY 20
    Assert-Near "op1-maxx" $graphics[1].MaxX 60
    Assert-Near "op1-maxy" $graphics[1].MaxY 50
    if ($graphics[0].Kind -cne "Fill" -or $graphics[0].Operator -cne "f") { $script:failures += "op0 kind/operator"; Write-Host "FAIL: op0 kind/operator" }
    if ($graphics[1].Kind -cne "Fill" -or $graphics[1].Operator -cne "f*") { $script:failures += "op1 kind/operator"; Write-Host "FAIL: op1 kind/operator" }
}
$textOperations = @(Get-Content -Raw -LiteralPath (Join-Path $inspectDir "text-operations.json") | ConvertFrom-Json)
$texts = @($textOperations.DecodedText)
$wantTexts = @("Hello", "a(b)c", "a(b", "gray again", "default again", "CMYK", "0 0 1 rg")
if (@(Compare-Object $texts $wantTexts).Count -ne 0) {
    $script:failures += ("texts: got [" + ($texts -join "|") + "]")
    Write-Host ("FAIL: texts: got [" + ($texts -join "|") + "]")
}
else {
    Write-Host 'PASS: nested/escaped text and color-operator payloads'
}
$wantColors = @('g:0', 'g:0.2', 'rg:0,0,1', 'g:0.2', 'g:0', 'k:0,1,1,0', 'k:0,1,1,0')
if (($textOperations.FillColor -join '|') -cne ($wantColors -join '|')) {
    $failures += ('Fill-color state: ' + ($textOperations.FillColor -join '|'))
}
else { Write-Host 'PASS: default/gray/RGB/CMYK fill, nested save/restore and literal isolation' }

if ($failures.Count -ne 0) {
    throw ("Content tokenization checks failed: " + ($failures -join "; "))
}
Write-Host "Content tokenization checks passed."
