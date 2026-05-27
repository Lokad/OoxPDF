param(
    [Parameter(Mandatory = $true)]
    [string] $InputPptx,

    [Parameter(Mandatory = $true)]
    [string] $ReferenceTextOperations,

    [int] $Slide = 1,

    [string] $OutputJson
)

$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.IO.Compression.FileSystem

function Read-JsonArray([string] $Path) {
    $items = Get-Content -Raw -LiteralPath (Resolve-Path -LiteralPath $Path).Path | ConvertFrom-Json
    if ($null -eq $items) {
        return ,@()
    }

    if ($items -is [array]) {
        return ,$items
    }

    return ,@($items)
}

function Read-ZipXml([System.IO.Compression.ZipArchive] $Zip, [string] $PartName) {
    $entry = $Zip.GetEntry($PartName.TrimStart("/"))
    if ($null -eq $entry) {
        throw "Missing PPTX part: $PartName"
    }

    $stream = $entry.Open()
    try {
        $reader = [System.IO.StreamReader]::new($stream)
        try {
            return [xml]$reader.ReadToEnd()
        }
        finally {
            $reader.Dispose()
        }
    }
    finally {
        $stream.Dispose()
    }
}

function Convert-EmuToPoint($Value) {
    if ($null -eq $Value -or [string]$Value -eq "") {
        return $null
    }

    return [Math]::Round([double]$Value / 12700d, 6)
}

function Get-OptionalAttribute($Node, [string] $Name) {
    if ($null -eq $Node) {
        return $null
    }

    $attribute = $Node.Attributes[$Name]
    if ($null -eq $attribute) {
        return $null
    }

    return $attribute.Value
}

function Get-OfficeGridFontSize([double] $FontSize) {
    $deviceUnits = [Math]::Round($FontSize * 600d / 72d, [MidpointRounding]::AwayFromZero)
    return [Math]::Round($deviceUnits * 72d / 600d, 6)
}

function Select-One($Node, [System.Xml.XmlNamespaceManager] $NamespaceManager, [string] $XPath) {
    return $Node.SelectSingleNode($XPath, $NamespaceManager)
}

function Get-SlideTextBoxes([string] $PptxPath, [int] $SlideNumber) {
    $zip = [System.IO.Compression.ZipFile]::OpenRead((Resolve-Path -LiteralPath $PptxPath).Path)
    try {
        $slideXml = Read-ZipXml $zip "ppt/slides/slide$SlideNumber.xml"
        $namespaceManager = [System.Xml.XmlNamespaceManager]::new($slideXml.NameTable)
        $namespaceManager.AddNamespace("p", "http://schemas.openxmlformats.org/presentationml/2006/main")
        $namespaceManager.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")

        $rows = New-Object System.Collections.Generic.List[object]
        $shapeIndex = 0
        foreach ($shape in $slideXml.SelectNodes("//p:sp", $namespaceManager)) {
            $textNodes = $shape.SelectNodes(".//a:t", $namespaceManager)
            if ($textNodes.Count -eq 0) {
                continue
            }

            $textParts = foreach ($textNode in $textNodes) { $textNode.InnerText }
            $text = $textParts -join ""
            $runProperties = Select-One $shape $namespaceManager ".//a:rPr[@sz]"
            $sourceFontSizeExplicit = $null -ne $runProperties
            $sourceFontSize = if ($sourceFontSizeExplicit) { [Math]::Round([double](Get-OptionalAttribute $runProperties "sz") / 100d, 6) } else { 18d }
            $off = Select-One $shape $namespaceManager "./p:spPr/a:xfrm/a:off"
            $ext = Select-One $shape $namespaceManager "./p:spPr/a:xfrm/a:ext"
            $bodyPr = Select-One $shape $namespaceManager "./p:txBody/a:bodyPr"
            $autofit = if ($null -ne (Select-One $shape $namespaceManager "./p:txBody/a:bodyPr/a:noAutofit")) {
                "noAutofit"
            }
            elseif ($null -ne (Select-One $shape $namespaceManager "./p:txBody/a:bodyPr/a:spAutoFit")) {
                "spAutoFit"
            }
            elseif ($null -ne (Select-One $shape $namespaceManager "./p:txBody/a:bodyPr/a:normAutofit")) {
                "normAutofit"
            }
            else {
                ""
            }

            $rows.Add([pscustomobject]@{
                ShapeIndex = $shapeIndex
                Text = $text
                SourceFontSize = $sourceFontSize
                SourceFontSizeExplicit = $sourceFontSizeExplicit
                X = Convert-EmuToPoint (Get-OptionalAttribute $off "x")
                Y = Convert-EmuToPoint (Get-OptionalAttribute $off "y")
                Width = Convert-EmuToPoint (Get-OptionalAttribute $ext "cx")
                Height = Convert-EmuToPoint (Get-OptionalAttribute $ext "cy")
                Wrap = Get-OptionalAttribute $bodyPr "wrap"
                Autofit = $autofit
                LeftInset = Convert-EmuToPoint (Get-OptionalAttribute $bodyPr "lIns")
                RightInset = Convert-EmuToPoint (Get-OptionalAttribute $bodyPr "rIns")
                TopInset = Convert-EmuToPoint (Get-OptionalAttribute $bodyPr "tIns")
                BottomInset = Convert-EmuToPoint (Get-OptionalAttribute $bodyPr "bIns")
            })
            $shapeIndex++
        }

        return ,$rows.ToArray()
    }
    finally {
        $zip.Dispose()
    }
}

$textBoxes = Get-SlideTextBoxes $InputPptx $Slide
$referenceOps = Read-JsonArray $ReferenceTextOperations |
    Where-Object { $null -ne $_.DecodedText -and [string]$_.DecodedText -ne "" }

$rows = New-Object System.Collections.Generic.List[object]
$count = [Math]::Min($textBoxes.Count, $referenceOps.Count)
for ($i = 0; $i -lt $count; $i++) {
    $textBox = $textBoxes[$i]
    $operation = $referenceOps[$i]
    $sourceFontSize = $textBox.SourceFontSize
    $firstGridFontSize = if ($null -eq $sourceFontSize) { $null } else { Get-OfficeGridFontSize $sourceFontSize }
    $officeFontSize = [Math]::Round([double]$operation.FontSize, 6)

    $rows.Add([pscustomobject]@{
        Index = $i
        ShapeIndex = $textBox.ShapeIndex
        Text = $textBox.Text
        ReferenceText = [string]$operation.DecodedText
        TextMatches = $textBox.Text -eq [string]$operation.DecodedText
        SourceFontSize = $sourceFontSize
        SourceFontSizeExplicit = $textBox.SourceFontSizeExplicit
        OfficeFontSize = $officeFontSize
        FirstGridFontSize = $firstGridFontSize
        SecondaryDelta = if ($null -eq $firstGridFontSize) { $null } else { [Math]::Round($officeFontSize - $firstGridFontSize, 6) }
        X = $textBox.X
        Y = $textBox.Y
        Width = $textBox.Width
        Height = $textBox.Height
        Wrap = $textBox.Wrap
        Autofit = $textBox.Autofit
        LeftInset = $textBox.LeftInset
        RightInset = $textBox.RightInset
        TopInset = $textBox.TopInset
        BottomInset = $textBox.BottomInset
        RefX = if ($null -eq $operation.EffectiveX) { [Math]::Round([double]$operation.X, 6) } else { [Math]::Round([double]$operation.EffectiveX, 6) }
        RefBaselineY = if ($null -eq $operation.EffectiveY) { [Math]::Round([double]$operation.Y, 6) } else { [Math]::Round([double]$operation.EffectiveY, 6) }
        RefObjectNumber = $operation.ObjectNumber
    })
}

$result = $rows.ToArray()
if ($OutputJson) {
    $outputPath = [System.IO.Path]::GetFullPath($OutputJson)
    $outputDirectory = [System.IO.Path]::GetDirectoryName($outputPath)
    if (-not [string]::IsNullOrWhiteSpace($outputDirectory)) {
        New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null
    }

    $result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $outputPath -Encoding UTF8
}
else {
    $result
}
