param([Parameter(Mandatory = $true)] [string] $InspectorPath)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$id = 'pptx-inventory-test-' + [guid]::NewGuid().ToString('N')
$inputs = (New-Item -ItemType Directory -Path (Join-Path $repo "private-cases/$id")).FullName
$output = Join-Path $repo "artifacts/private-visual/$id/inventory"
$fixtures = Join-Path $repo 'tests/Lokad.OoxPdf.Tests/Cases'
$original = Join-Path $inputs 'baseline.pptx'
$relocated = Join-Path $inputs 'relocated.pptx'
Copy-Item -LiteralPath (Join-Path $fixtures 'pptx-ladder-04-typography-repeated-spaces.pptx') -Destination $original
Copy-Item -LiteralPath $original -Destination $relocated
[IO.File]::SetAttributes($relocated, [IO.FileAttributes]::Normal)

$zip = [IO.Compression.ZipFile]::Open($relocated, [IO.Compression.ZipArchiveMode]::Update)
try {
    function Replace-Entry([string] $Name, [string] $NewName, [string] $OldText, [string] $NewText) {
        $entry = $zip.GetEntry($Name)
        $reader = [IO.StreamReader]::new($entry.Open())
        try { $text = $reader.ReadToEnd() }
        finally { $reader.Dispose() }
        if ($OldText) {
            if (!$text.Contains($OldText)) { throw ('Missing expected source token in ' + $Name) }
            $text = $text.Replace($OldText, $NewText)
        }
        $entry.Delete()
        $writer = [IO.StreamWriter]::new($zip.CreateEntry($NewName).Open())
        try { $writer.Write($text) }
        finally { $writer.Dispose() }
    }
    Replace-Entry 'ppt/presentation.xml' 'ppt/moved.xml'
    Replace-Entry 'ppt/_rels/presentation.xml.rels' 'ppt/_rels/moved.xml.rels'
    Replace-Entry '_rels/.rels' '_rels/.rels' 'ppt/presentation.xml' 'ppt/moved.xml'
    Replace-Entry '[Content_Types].xml' '[Content_Types].xml' '/ppt/presentation.xml' '/ppt/moved.xml'
}
finally { $zip.Dispose() }

Copy-Item -LiteralPath (Join-Path $fixtures 'SlideMembership/listed-hidden-unlisted-visible.pptx') -Destination (Join-Path $inputs 'hidden.pptx')
$reports = @{}
foreach ($variant in @('baseline', 'relocated', 'hidden')) {
    $manifest = Join-Path $inputs ($variant + '.json')
    @{ input = $variant + '.pptx'; kind = 'pptx'; id = $id } | ConvertTo-Json | Set-Content -LiteralPath $manifest
    $report = Join-Path $output ($variant + '.json')
    & (Join-Path $PSScriptRoot 'InventoryPptxSlides.ps1') -Case $manifest -OutputPath $report -InspectorPath $InspectorPath -SkipBuild | Out-Null
    $reports[$variant] = Get-Content -LiteralPath $report -Raw | ConvertFrom-Json -AsHashtable
    $reports[$variant].Summary.Remove('GeneratedAt')
}
if ($reports.baseline.Summary.SlideCount -ne 1 -or
    (ConvertTo-Json -InputObject $reports.baseline -Depth 8 -Compress) -cne
        (ConvertTo-Json -InputObject $reports.relocated -Depth 8 -Compress)) {
    throw 'Relocating the presentation part changed the source inventory.'
}
$slide = $reports.relocated.Slides[0]
if ($slide.SlideContent.Shapes -ne 1 -or $slide.SlideContent.TextBodies -ne 1 -or
    $slide.Layout.Shapes -ne 3 -or $slide.Master.Shapes -ne 5 -or $slide.LayoutRelationships.SlideMasters -ne 1) {
    throw 'The relocated presentation must inventory its slide, layout and master content.'
}
if ($reports.hidden.Summary.SlideCount -ne 1 -or $reports.hidden.Slides[0].SlidePart -cne '/ppt/slides/slide1.xml') {
    throw 'Inventory must keep the authored hidden slide and exclude the unlisted slide.'
}
Write-Output 'PPTX inventory preserves relocated package ownership and authored hidden-slide membership.'
