param(
    [Parameter(Mandatory = $true)]
    [string] $Case,

    [string] $OutputPath
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$privateRoot = Join-Path $repoRoot "private-cases"
$artifactRoot = Join-Path $repoRoot "artifacts/private-visual"

function Test-UnderDirectory([string] $Path, [string] $Directory) {
    $fullPath = [System.IO.Path]::GetFullPath($Path).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    $fullDirectory = [System.IO.Path]::GetFullPath($Directory).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    return $fullPath.Equals($fullDirectory, [System.StringComparison]::OrdinalIgnoreCase) -or
        $fullPath.StartsWith($fullDirectory + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase) -or
        $fullPath.StartsWith($fullDirectory + [System.IO.Path]::AltDirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)
}

function Test-GitTracked([string] $Path) {
    $relative = [System.IO.Path]::GetRelativePath($repoRoot, $Path)
    git -C $repoRoot ls-files --error-unmatch -- $relative *> $null
    return $LASTEXITCODE -eq 0
}

function Assert-PrivateUntracked([string] $Path, [string] $Label) {
    if (-not (Test-UnderDirectory $Path $privateRoot)) {
        throw "$Label must be under $privateRoot."
    }

    if (Test-GitTracked $Path) {
        throw "$Label is tracked by git and must not be used as a private case: $Path"
    }
}

function Read-ZipXml([System.IO.Compression.ZipArchive] $Zip, [string] $PartName) {
    $entryName = $PartName.TrimStart("/")
    $entry = $Zip.GetEntry($entryName)
    if ($entry -eq $null) {
        return $null
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

function Get-RelationshipPartName([string] $PartName) {
    $trimmed = $PartName.TrimStart("/")
    $directory = [System.IO.Path]::GetDirectoryName($trimmed).Replace("\", "/")
    $file = [System.IO.Path]::GetFileName($trimmed)
    if ([string]::IsNullOrWhiteSpace($directory)) {
        return "/_rels/$file.rels"
    }

    return "/$directory/_rels/$file.rels"
}

function Resolve-PartTarget([string] $BasePartName, [string] $Target) {
    if ($Target.StartsWith("/", [System.StringComparison]::Ordinal)) {
        return $Target
    }

    $baseDirectory = [System.IO.Path]::GetDirectoryName($BasePartName.TrimStart("/")).Replace("\", "/")
    $combined = if ([string]::IsNullOrWhiteSpace($baseDirectory)) { $Target } else { "$baseDirectory/$Target" }
    $segments = New-Object System.Collections.Generic.List[string]
    foreach ($segment in $combined -split "/") {
        if ($segment.Length -eq 0 -or $segment -eq ".") {
            continue
        }

        if ($segment -eq "..") {
            if ($segments.Count -gt 0) {
                $segments.RemoveAt($segments.Count - 1)
            }
            continue
        }

        $segments.Add($segment)
    }

    return "/" + ($segments -join "/")
}

function Get-Relationships([System.IO.Compression.ZipArchive] $Zip, [string] $PartName) {
    $relsXml = Read-ZipXml $Zip (Get-RelationshipPartName $PartName)
    if ($relsXml -eq $null) {
        return @()
    }

    $relationships = @()
    foreach ($relationship in $relsXml.Relationships.Relationship) {
        $target = [string]$relationship.Target
        $relationships += [pscustomobject]@{
            Id = [string]$relationship.Id
            Type = [string]$relationship.Type
            Target = $target
            TargetMode = [string]$relationship.TargetMode
            ResolvedTarget = if ([string]$relationship.TargetMode -eq "External") { $target } else { Resolve-PartTarget $PartName $target }
        }
    }

    return $relationships
}

function Count-XPath([xml] $Xml, [string] $XPath) {
    if ($Xml -eq $null) {
        return 0
    }

    return $Xml.SelectNodes($XPath).Count
}

function Test-XPath([xml] $Xml, [string] $XPath) {
    return (Count-XPath $Xml $XPath) -gt 0
}

function Get-PartInventory([System.IO.Compression.ZipArchive] $Zip, [string] $PartName) {
    $xml = Read-ZipXml $Zip $PartName
    if ($xml -eq $null) {
        return $null
    }

    [pscustomobject]@{
        PartName = $PartName
        Shapes = Count-XPath $xml "//*[local-name()='sp']"
        GroupShapes = Count-XPath $xml "//*[local-name()='grpSp']"
        Pictures = Count-XPath $xml "//*[local-name()='pic']"
        GraphicFrames = Count-XPath $xml "//*[local-name()='graphicFrame']"
        Tables = Count-XPath $xml "//*[local-name()='tbl']"
        Charts = Count-XPath $xml "//*[local-name()='chart']"
        TextBodies = Count-XPath $xml "//*[local-name()='txBody']"
        Placeholders = Count-XPath $xml "//*[local-name()='ph']"
        Transforms = Count-XPath $xml "//*[local-name()='xfrm']"
        RotatedTransforms = Count-XPath $xml "//*[local-name()='xfrm' and @rot]"
        FlippedTransforms = Count-XPath $xml "//*[local-name()='xfrm' and (@flipH or @flipV)]"
        SolidFills = Count-XPath $xml "//*[local-name()='solidFill']"
        GradientFills = Count-XPath $xml "//*[local-name()='gradFill']"
        PatternFills = Count-XPath $xml "//*[local-name()='pattFill']"
        PictureFills = Count-XPath $xml "//*[local-name()='blipFill']"
        Transparency = Count-XPath $xml "//*[local-name()='alpha']"
        Effects = Count-XPath $xml "//*[local-name()='effectLst' or local-name()='effectDag']"
        Clips = Count-XPath $xml "//*[local-name()='srcRect']"
        SmartArtSignals = Count-XPath $xml "//*[namespace-uri()='http://schemas.openxmlformats.org/drawingml/2006/diagram']"
        MediaSignals = Count-XPath $xml "//*[local-name()='video' or local-name()='audio' or local-name()='videoFile' or local-name()='audioFile']"
        OleSignals = Count-XPath $xml "//*[local-name()='oleObj']"
        Notes = @{
            HasBackground = Test-XPath $xml "//*[local-name()='bg']"
            HasHiddenPlaceholder = Test-XPath $xml "//*[local-name()='ph' and (@hidden='1' or @hidden='true')]"
            HasGroupedContent = Test-XPath $xml "//*[local-name()='grpSp']"
            HasEffects = Test-XPath $xml "//*[local-name()='effectLst' or local-name()='effectDag']"
            HasTransparency = Test-XPath $xml "//*[local-name()='alpha']"
        }
    }
}

function Get-RelationshipSummary($Relationships) {
    $known = [ordered]@{
        Images = 0
        Charts = 0
        SlideLayouts = 0
        SlideMasters = 0
        Themes = 0
        Hyperlinks = 0
        OleObjects = 0
        Media = 0
        Other = 0
        External = 0
    }

    foreach ($relationship in $Relationships) {
        if ($relationship.TargetMode -eq "External") {
            $known.External++
        }

        if ($relationship.Type.EndsWith("/image", [System.StringComparison]::OrdinalIgnoreCase)) {
            $known.Images++
        }
        elseif ($relationship.Type.EndsWith("/chart", [System.StringComparison]::OrdinalIgnoreCase)) {
            $known.Charts++
        }
        elseif ($relationship.Type.EndsWith("/slideLayout", [System.StringComparison]::OrdinalIgnoreCase)) {
            $known.SlideLayouts++
        }
        elseif ($relationship.Type.EndsWith("/slideMaster", [System.StringComparison]::OrdinalIgnoreCase)) {
            $known.SlideMasters++
        }
        elseif ($relationship.Type.EndsWith("/theme", [System.StringComparison]::OrdinalIgnoreCase)) {
            $known.Themes++
        }
        elseif ($relationship.Type.EndsWith("/hyperlink", [System.StringComparison]::OrdinalIgnoreCase)) {
            $known.Hyperlinks++
        }
        elseif ($relationship.Type.EndsWith("/oleObject", [System.StringComparison]::OrdinalIgnoreCase)) {
            $known.OleObjects++
        }
        elseif ($relationship.Type.EndsWith("/video", [System.StringComparison]::OrdinalIgnoreCase) -or
            $relationship.Type.EndsWith("/audio", [System.StringComparison]::OrdinalIgnoreCase) -or
            $relationship.Type.EndsWith("/media", [System.StringComparison]::OrdinalIgnoreCase)) {
            $known.Media++
        }
        else {
            $known.Other++
        }
    }

    [pscustomobject]$known
}

Add-Type -AssemblyName System.IO.Compression.FileSystem

$caseFull = (Resolve-Path -LiteralPath $Case).Path
Assert-PrivateUntracked $caseFull "Private case manifest"
$caseDirectory = Split-Path -Parent $caseFull
$manifest = Get-Content -Raw -LiteralPath $caseFull | ConvertFrom-Json

if ([string]::IsNullOrWhiteSpace($manifest.id)) {
    throw "Private case manifest must contain an id."
}

if ($manifest.kind -ne $null -and -not ([string]$manifest.kind).Equals("pptx", [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Private case manifest kind must be pptx."
}

$caseId = [string]$manifest.id
if ($caseId.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars()) -ge 0 -or $caseId.Contains("/") -or $caseId.Contains("\")) {
    throw "Private case id must be a single filename-safe path segment."
}

if ([string]::IsNullOrWhiteSpace($manifest.input)) {
    throw "Private case manifest must contain an input path."
}

$inputFull = (Resolve-Path -LiteralPath (Join-Path $caseDirectory $manifest.input)).Path
Assert-PrivateUntracked $inputFull "Private case input"

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $inventoryRoot = Join-Path $artifactRoot "$caseId/inventory"
    New-Item -ItemType Directory -Force -Path $inventoryRoot | Out-Null
    $OutputPath = Join-Path $inventoryRoot ((Get-Date -Format "yyyyMMdd-HHmmss") + ".json")
}

$outputFull = [System.IO.Path]::GetFullPath($OutputPath)
$outputDirectory = [System.IO.Path]::GetDirectoryName($outputFull)
if (-not (Test-UnderDirectory $outputFull $artifactRoot)) {
    throw "Inventory output must be under $artifactRoot."
}

New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null

$zip = [System.IO.Compression.ZipFile]::OpenRead($inputFull)
try {
    $presentationPart = "/ppt/presentation.xml"
    $presentation = Read-ZipXml $zip $presentationPart
    if ($presentation -eq $null) {
        throw "PPTX package does not contain ppt/presentation.xml."
    }

    $presentationRels = Get-Relationships $zip $presentationPart
    $slides = @()
    $slideIndex = 1
    foreach ($slideId in $presentation.SelectNodes("//*[local-name()='sldId']")) {
        $relationshipId = [string]$slideId.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
        $slideRelationship = $presentationRels | Where-Object { $_.Id -eq $relationshipId } | Select-Object -First 1
        if ($slideRelationship -eq $null) {
            continue
        }

        $slidePart = $slideRelationship.ResolvedTarget
        $slideRels = @(Get-Relationships $zip $slidePart)
        $layoutRel = $slideRels | Where-Object { $_.Type.EndsWith("/slideLayout", [System.StringComparison]::OrdinalIgnoreCase) } | Select-Object -First 1
        $layoutPart = if ($layoutRel -ne $null) { $layoutRel.ResolvedTarget } else { $null }
        $layoutRels = if ($layoutPart -ne $null) { @(Get-Relationships $zip $layoutPart) } else { @() }
        $masterRel = $layoutRels | Where-Object { $_.Type.EndsWith("/slideMaster", [System.StringComparison]::OrdinalIgnoreCase) } | Select-Object -First 1
        $masterPart = if ($masterRel -ne $null) { $masterRel.ResolvedTarget } else { $null }

        $slideInventory = Get-PartInventory $zip $slidePart
        $layoutInventory = if ($layoutPart -ne $null) { Get-PartInventory $zip $layoutPart } else { $null }
        $masterInventory = if ($masterPart -ne $null) { Get-PartInventory $zip $masterPart } else { $null }

        $slides += [pscustomobject]@{
            SlideNumber = $slideIndex
            SlidePart = $slidePart
            LayoutPart = $layoutPart
            MasterPart = $masterPart
            SlideContent = $slideInventory
            Layout = $layoutInventory
            Master = $masterInventory
            Relationships = Get-RelationshipSummary $slideRels
            LayoutRelationships = Get-RelationshipSummary $layoutRels
        }

        $slideIndex++
    }

    $summary = [pscustomobject]@{
        CaseId = $caseId
        GeneratedAt = (Get-Date).ToString("o")
        SlideCount = $slides.Count
        SlidesWithCharts = @($slides | Where-Object { $_.SlideContent.Charts -gt 0 -or $_.Relationships.Charts -gt 0 }).Count
        SlidesWithTables = @($slides | Where-Object { $_.SlideContent.Tables -gt 0 }).Count
        SlidesWithPictures = @($slides | Where-Object { $_.SlideContent.Pictures -gt 0 -or $_.Relationships.Images -gt 0 }).Count
        SlidesWithGroups = @($slides | Where-Object { $_.SlideContent.GroupShapes -gt 0 }).Count
        SlidesWithEffects = @($slides | Where-Object { $_.SlideContent.Effects -gt 0 -or $_.Layout.Effects -gt 0 -or $_.Master.Effects -gt 0 }).Count
        SlidesWithTransparency = @($slides | Where-Object { $_.SlideContent.Transparency -gt 0 -or $_.Layout.Transparency -gt 0 -or $_.Master.Transparency -gt 0 }).Count
        SlidesWithInheritedContent = @($slides | Where-Object { ($_.Layout.Shapes + $_.Layout.Pictures + $_.Master.Shapes + $_.Master.Pictures) -gt 0 }).Count
        TotalSlideShapes = ($slides | ForEach-Object { $_.SlideContent.Shapes } | Measure-Object -Sum).Sum
        TotalSlidePictures = ($slides | ForEach-Object { $_.SlideContent.Pictures } | Measure-Object -Sum).Sum
        TotalSlideTables = ($slides | ForEach-Object { $_.SlideContent.Tables } | Measure-Object -Sum).Sum
        TotalSlideCharts = ($slides | ForEach-Object { $_.SlideContent.Charts } | Measure-Object -Sum).Sum
    }

    $inventory = [pscustomobject]@{
        Summary = $summary
        Slides = $slides
    }

    $inventory | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $outputFull
    Write-Host "PPTX slide inventory: $outputFull"
    $summary | ConvertTo-Json -Depth 4
}
finally {
    $zip.Dispose()
}
