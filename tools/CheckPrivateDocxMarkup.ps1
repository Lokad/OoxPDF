param(
    [Parameter(Mandatory = $true)]
    [string] $Case,

    [ValidateSet("final", "original", "simple", "all", "simple-markup", "all-markup")]
    [string] $MarkupMode = "simple",

    [ValidateSet("preserve", "preserve-layout", "preserve-document-layout", "reserve", "reserve-margin", "markup-margin", "reserve-markup-margin", "word", "word-compatible", "word-compatible-all-markup", "office", "office-compatible", "office-compatible-all-markup")]
    [string] $MarkupGeometry = "preserve",

    [switch] $SkipDocxInspect,

    [switch] $SkipPdfInspect,

    [switch] $ValidateOnly
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$privateRoot = Join-Path $repoRoot "private-cases"

function Test-UnderDirectory([string] $Path, [string] $Directory) {
    $fullPath = [System.IO.Path]::GetFullPath($Path).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    $fullDirectory = [System.IO.Path]::GetFullPath($Directory).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    return $fullPath.Equals($fullDirectory, [System.StringComparison]::OrdinalIgnoreCase) -or
        $fullPath.StartsWith($fullDirectory + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase) -or
        $fullPath.StartsWith($fullDirectory + [System.IO.Path]::AltDirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)
}

function Test-GitTracked([string] $Path) {
    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $fullRoot = [System.IO.Path]::GetFullPath($repoRoot).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    $relative = $fullPath.Substring($fullRoot.Length).TrimStart([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    $previousErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        git -C $repoRoot ls-files --error-unmatch -- $relative *> $null
        return $LASTEXITCODE -eq 0
    }
    finally {
        $ErrorActionPreference = $previousErrorActionPreference
    }
}

function Assert-PrivateUntracked([string] $Path, [string] $Label) {
    if (-not (Test-UnderDirectory $Path $privateRoot)) {
        throw "$Label must be under $privateRoot."
    }

    if (Test-GitTracked $Path) {
        throw "$Label is tracked by git and must not be used as a private case: $Path"
    }
}

function ConvertTo-RepoPath([string] $Path) {
    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $fullRoot = [System.IO.Path]::GetFullPath($repoRoot).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    if ($fullPath.StartsWith($fullRoot + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $fullPath.Substring($fullRoot.Length + 1).Replace([System.IO.Path]::DirectorySeparatorChar, "/")
    }

    return $fullPath
}

function Get-PdfPageCount([string] $Path) {
    $pdf = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::Latin1)
    return [regex]::Matches($pdf, "/Type\s*/Page(?!s)").Count
}

function Get-PdfMediaBoxSummary([string] $Path) {
    $pdf = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::Latin1)
    $matches = [regex]::Matches($pdf, "/MediaBox\s*\[\s*(?<x0>-?\d+(?:\.\d+)?)\s+(?<y0>-?\d+(?:\.\d+)?)\s+(?<x1>-?\d+(?:\.\d+)?)\s+(?<y1>-?\d+(?:\.\d+)?)\s*\]")
    $boxes = foreach ($match in $matches) {
        $x0 = [double]::Parse($match.Groups["x0"].Value, [Globalization.CultureInfo]::InvariantCulture)
        $y0 = [double]::Parse($match.Groups["y0"].Value, [Globalization.CultureInfo]::InvariantCulture)
        $x1 = [double]::Parse($match.Groups["x1"].Value, [Globalization.CultureInfo]::InvariantCulture)
        $y1 = [double]::Parse($match.Groups["y1"].Value, [Globalization.CultureInfo]::InvariantCulture)
        [pscustomobject]@{
            Width = [Math]::Round($x1 - $x0, 3)
            Height = [Math]::Round($y1 - $y0, 3)
        }
    }

    return @($boxes |
        Group-Object Width, Height |
        ForEach-Object {
            [pscustomobject]@{
                Width = $_.Group[0].Width
                Height = $_.Group[0].Height
                Count = $_.Count
            }
        } |
        Sort-Object Width, Height)
}

function Add-Count([System.Collections.IDictionary] $Counts, [string] $Name, [int] $Value) {
    if (-not $Counts.Contains($Name)) {
        $Counts[$Name] = 0
    }

    $Counts[$Name] = [int]$Counts[$Name] + $Value
}

function Get-PrivateSafeFieldOpcode([string] $Instruction) {
    if ([string]::IsNullOrWhiteSpace($Instruction)) {
        return "UNKNOWN"
    }

    $match = [regex]::Match($Instruction.TrimStart(), "^[A-Za-z]+")
    if (-not $match.Success) {
        return "UNKNOWN"
    }

    $opcode = $match.Value.ToUpperInvariant()
    $knownOpcodes = @(
        "ADVANCE",
        "ASK",
        "AUTHOR",
        "COMMENTS",
        "CREATEDATE",
        "DATE",
        "DOCPROPERTY",
        "EDITTIME",
        "FILENAME",
        "FILESIZE",
        "FORMCHECKBOX",
        "FORMDROPDOWN",
        "FORMTEXT",
        "HYPERLINK",
        "IF",
        "INCLUDEPICTURE",
        "KEYWORDS",
        "LASTSAVEDBY",
        "MERGEFIELD",
        "NOTEREF",
        "NUMPAGES",
        "PAGE",
        "PAGEREF",
        "PRINTDATE",
        "REF",
        "REVNUM",
        "SAVEDATE",
        "SEQ",
        "SET",
        "STYLEREF",
        "SUBJECT",
        "SYMBOL",
        "TIME",
        "TITLE",
        "TOA",
        "TOC",
        "USERADDRESS",
        "USERINITIALS",
        "USERNAME",
        "XE"
    )

    if ($knownOpcodes -contains $opcode) {
        return $opcode
    }

    return "OTHER"
}

function Get-PrivateSafeComplexFieldContainer([System.Xml.Linq.XElement] $Element) {
    $run = $Element.Parent
    $container = if ($null -eq $run) { $null } else { $run.Parent }
    if ($null -eq $container) {
        return "none"
    }

    $localName = $container.Name.LocalName
    $knownContainers = @("p", "hyperlink", "fldSimple", "ins", "del", "moveFrom", "moveTo", "sdtContent")
    if ($knownContainers -contains $localName) {
        return $localName
    }

    return "other"
}

function New-PrivateSafeComplexFieldState([string] $Container) {
    return [pscustomobject]@{
        Instruction = [System.Text.StringBuilder]::new()
        Container = $Container
        HasSeparate = $false
        InResult = $false
        HasCachedResult = $false
    }
}

function Add-PrivateSafeComplexFieldSummary(
    [System.Collections.IDictionary] $Summary,
    [object] $Field) {
    $instruction = $Field.Instruction.ToString()
    $opcode = Get-PrivateSafeFieldOpcode $instruction
    $hasCachedResult = [bool]$Field.HasCachedResult
    $hasPlaceholder = $opcode -eq "PAGE" -or $opcode -eq "NUMPAGES"
    $container = [string]$Field.Container
    $supportedContainers = @("p", "hyperlink", "fldSimple", "ins", "del", "moveFrom", "moveTo", "sdtContent")
    $isSupportedContainer = $supportedContainers -contains $container
    $isUnsupportedShape = (-not $isSupportedContainer) -or
        ((-not $hasPlaceholder) -and ((-not [bool]$Field.HasSeparate) -or (-not $hasCachedResult)))

    $Summary.Total = [int]$Summary.Total + 1
    Add-Count $Summary.ByOpcode $opcode 1
    Add-Count $Summary.ByContainer $container 1

    if ($hasCachedResult) {
        $Summary.WithCachedResult = [int]$Summary.WithCachedResult + 1
    }
    else {
        $Summary.WithoutCachedResult = [int]$Summary.WithoutCachedResult + 1
        Add-Count $Summary.WithoutCachedResultByOpcode $opcode 1
    }

    if ($hasPlaceholder) {
        $Summary.WithDynamicPlaceholder = [int]$Summary.WithDynamicPlaceholder + 1
    }

    if (-not $isSupportedContainer) {
        $Summary.UnsupportedContainer = [int]$Summary.UnsupportedContainer + 1
    }

    if ($isUnsupportedShape) {
        $Summary.UnsupportedShape = [int]$Summary.UnsupportedShape + 1
        Add-Count $Summary.UnsupportedShapeByOpcode $opcode 1
    }
}

function Add-PrivateSafeSimpleFieldSummary(
    [System.Collections.IDictionary] $Summary,
    [System.Xml.Linq.XElement] $Element,
    [string] $WordprocessingNs) {
    $instructionAttribute = $Element.Attribute([System.Xml.Linq.XName]::Get("instr", $WordprocessingNs))
    $instruction = if ($null -eq $instructionAttribute) { "" } else { [string]$instructionAttribute.Value }
    $opcode = Get-PrivateSafeFieldOpcode $instruction
    $container = if ($null -eq $Element.Parent) { "none" } else { $Element.Parent.Name.LocalName }
    $hasCachedResult = @($Element.Descendants() |
        Where-Object {
            ($_.Name.LocalName -eq "t" -or $_.Name.LocalName -eq "delText") -and
                -not [string]::IsNullOrEmpty([string]$_.Value)
        }).Count -ne 0
    $hasPlaceholder = $opcode -eq "PAGE" -or $opcode -eq "NUMPAGES"

    $Summary.Total = [int]$Summary.Total + 1
    Add-Count $Summary.ByOpcode $opcode 1
    Add-Count $Summary.ByContainer $container 1
    if ($hasCachedResult) {
        $Summary.WithCachedResult = [int]$Summary.WithCachedResult + 1
    }
    else {
        $Summary.WithoutCachedResult = [int]$Summary.WithoutCachedResult + 1
    }

    if ($hasPlaceholder) {
        $Summary.WithDynamicPlaceholder = [int]$Summary.WithDynamicPlaceholder + 1
        Add-Count $Summary.DynamicPlaceholderByOpcode $opcode 1
    }
}

function Get-PrivateSafeFormattingRevisionFamily([string] $ElementName) {
    switch ($ElementName) {
        "rPrChange" { return "Run" }
        "pPrChange" { return "Paragraph" }
        "tblPrChange" { return "Table" }
        "trPrChange" { return "Row" }
        "tcPrChange" { return "Cell" }
        "sectPrChange" { return "Section" }
        default { return $null }
    }
}

function Get-PrivateSafeFormattingRevisionPropertyNames(
    [System.Xml.Linq.XElement] $Element,
    [string] $WordprocessingNs) {
    $propertyNames = [System.Collections.Generic.List[string]]::new()
    foreach ($child in @($Element.Elements() | Where-Object { $_.Name.NamespaceName -eq $WordprocessingNs })) {
        $grandchildren = @($child.Elements() | Where-Object { $_.Name.NamespaceName -eq $WordprocessingNs })
        if ($grandchildren.Count -eq 0) {
            $propertyNames.Add($child.Name.LocalName)
            continue
        }

        foreach ($grandchild in $grandchildren) {
            $propertyNames.Add($grandchild.Name.LocalName)
        }
    }

    return @($propertyNames | Sort-Object -Unique)
}

function Add-PrivateSafeFormattingRevisionSummary(
    [System.Collections.IDictionary] $Summary,
    [System.Xml.Linq.XElement] $Element,
    [string] $WordprocessingNs) {
    $family = Get-PrivateSafeFormattingRevisionFamily $Element.Name.LocalName
    if ($null -eq $family) {
        return
    }

    $Summary.Total = [int]$Summary.Total + 1
    Add-Count $Summary.ByFamily $family 1

    $propertyNames = @(Get-PrivateSafeFormattingRevisionPropertyNames $Element $WordprocessingNs)
    if ($propertyNames.Count -eq 0) {
        $Summary.WithoutPropertyName = [int]$Summary.WithoutPropertyName + 1
        return
    }

    foreach ($propertyName in $propertyNames) {
        Add-Count $Summary.ByFamilyAndProperty ("{0}:{1}" -f $family, $propertyName) 1
    }
}

function Test-OoxBooleanTrue([string] $Value) {
    $normalized = $Value.Trim().ToLowerInvariant()
    return $normalized -eq "1" -or $normalized -eq "true" -or $normalized -eq "on" -or $normalized -eq "yes"
}

function Test-OoxBooleanFalse([string] $Value) {
    $normalized = $Value.Trim().ToLowerInvariant()
    return $normalized -eq "0" -or $normalized -eq "false" -or $normalized -eq "off" -or $normalized -eq "no"
}

function Get-DocxMarkupFeatureCounts([string] $Path) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $wordprocessingNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    $office2010WordNs = "http://schemas.microsoft.com/office/word/2010/wordml"
    $office2012WordNs = "http://schemas.microsoft.com/office/word/2012/wordml"
    $vmlNs = "urn:schemas-microsoft-com:vml"
    $revisionNames = @(
        "ins",
        "del",
        "delText",
        "moveFrom",
        "moveTo",
        "moveFromRangeStart",
        "moveFromRangeEnd",
        "moveToRangeStart",
        "moveToRangeEnd",
        "rPrChange",
        "pPrChange",
        "tblPrChange",
        "trPrChange",
        "tcPrChange",
        "sectPrChange"
    )
    $commentMarkerNames = @("commentRangeStart", "commentRangeEnd", "commentReference")
    $revisionCounts = [ordered]@{}
    $commentMarkerCounts = [ordered]@{}
    $formattingRevisionSummary = [ordered]@{
        Total = 0
        WithoutPropertyName = 0
        ByFamily = [ordered]@{
            Run = 0
            Paragraph = 0
            Table = 0
            Row = 0
            Cell = 0
            Section = 0
        }
        ByFamilyAndProperty = [ordered]@{}
    }
    foreach ($name in $revisionNames) {
        $revisionCounts[$name] = 0
    }

    foreach ($name in $commentMarkerNames) {
        $commentMarkerCounts[$name] = 0
    }

    $commentBodyCount = 0
    $modernCommentElementCount = 0
    $modernCommentSummary = [ordered]@{
        CommentParagraphIdCount = 0
        CommentExCount = 0
        WithParagraphId = 0
        WithParentParagraphId = 0
        ResolvedCount = 0
        OpenCount = 0
        UnknownResolvedStateCount = 0
        MetadataLinkedToKnownCommentCount = 0
        MetadataWithoutKnownCommentCount = 0
        ParentLinkedToKnownCommentCount = 0
        ParentWithoutKnownCommentCount = 0
    }
    $commentParagraphIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $commentMetadataParagraphIds = [System.Collections.Generic.List[string]]::new()
    $commentMetadataParentParagraphIds = [System.Collections.Generic.List[string]]::new()
    $xmlPartCount = 0
    $vmlElementCount = 0
    $vmlElementCounts = [ordered]@{}
    $vmlShapeCount = 0
    $vmlShapeWithImageDataCount = 0
    $vmlShapeWithTextboxOnlyCount = 0
    $vmlTextBoxCount = 0
    $vmlTextBoxWithContentCount = 0
    $vmlTextBoxParagraphCount = 0
    $vmlTextBoxTableCount = 0
    $vmlShapeChildSignatures = [ordered]@{}
    $complexFieldSummary = [ordered]@{
        Total = 0
        WithCachedResult = 0
        WithoutCachedResult = 0
        WithDynamicPlaceholder = 0
        MalformedOrUnclosed = 0
        UnsupportedContainer = 0
        UnsupportedShape = 0
        ByOpcode = [ordered]@{}
        WithoutCachedResultByOpcode = [ordered]@{}
        UnsupportedShapeByOpcode = [ordered]@{}
        ByContainer = [ordered]@{}
    }
    $simpleFieldSummary = [ordered]@{
        Total = 0
        WithCachedResult = 0
        WithoutCachedResult = 0
        WithDynamicPlaceholder = 0
        ByOpcode = [ordered]@{}
        DynamicPlaceholderByOpcode = [ordered]@{}
        ByContainer = [ordered]@{}
    }
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        foreach ($entry in $zip.Entries) {
            if (-not $entry.FullName.StartsWith("word/", [System.StringComparison]::OrdinalIgnoreCase) -or
                -not $entry.FullName.EndsWith(".xml", [System.StringComparison]::OrdinalIgnoreCase)) {
                continue
            }

            $stream = $entry.Open()
            try {
                $xml = [System.Xml.Linq.XDocument]::Load($stream)
            }
            catch {
                continue
            }
            finally {
                $stream.Dispose()
            }

            $xmlPartCount++
            $wordElements = @($xml.Descendants() | Where-Object { $_.Name.NamespaceName -eq $wordprocessingNs })
            foreach ($name in $revisionNames) {
                $count = @($wordElements | Where-Object { $_.Name.LocalName -eq $name }).Count
                Add-Count $revisionCounts $name $count
            }

            foreach ($element in $wordElements) {
                Add-PrivateSafeFormattingRevisionSummary $formattingRevisionSummary $element $wordprocessingNs
            }

            foreach ($simpleField in @($wordElements | Where-Object { $_.Name.LocalName -eq "fldSimple" })) {
                Add-PrivateSafeSimpleFieldSummary $simpleFieldSummary $simpleField $wordprocessingNs
            }

            foreach ($name in $commentMarkerNames) {
                $count = @($wordElements | Where-Object { $_.Name.LocalName -eq $name }).Count
                Add-Count $commentMarkerCounts $name $count
            }

            if ($entry.FullName.Equals("word/comments.xml", [System.StringComparison]::OrdinalIgnoreCase)) {
                $comments = @($wordElements | Where-Object { $_.Name.LocalName -eq "comment" })
                $commentBodyCount += $comments.Count
                foreach ($comment in $comments) {
                    $paragraphId = @($comment.Descendants([System.Xml.Linq.XName]::Get("p", $wordprocessingNs)) |
                        ForEach-Object {
                            $attribute = $_.Attribute([System.Xml.Linq.XName]::Get("paraId", $office2010WordNs))
                            if ($null -ne $attribute -and -not [string]::IsNullOrWhiteSpace($attribute.Value)) {
                                [string]$attribute.Value
                            }
                        } |
                        Select-Object -First 1)
                    if ($paragraphId.Count -ne 0) {
                        [void]$commentParagraphIds.Add($paragraphId[0])
                    }
                }
            }

            if ($entry.FullName.Contains("comments", [System.StringComparison]::OrdinalIgnoreCase)) {
                $modernCommentElementCount += @($xml.Descendants() | Where-Object { $_.Name.LocalName -like "commentEx*" }).Count
            }

            foreach ($commentEx in @($xml.Descendants([System.Xml.Linq.XName]::Get("commentEx", $office2012WordNs)))) {
                $modernCommentSummary.CommentExCount = [int]$modernCommentSummary.CommentExCount + 1
                $paragraphIdAttribute = $commentEx.Attribute([System.Xml.Linq.XName]::Get("paraId", $office2012WordNs))
                if ($null -ne $paragraphIdAttribute -and -not [string]::IsNullOrWhiteSpace($paragraphIdAttribute.Value)) {
                    $modernCommentSummary.WithParagraphId = [int]$modernCommentSummary.WithParagraphId + 1
                    $commentMetadataParagraphIds.Add([string]$paragraphIdAttribute.Value)
                }

                $parentParagraphIdAttribute = $commentEx.Attribute([System.Xml.Linq.XName]::Get("paraIdParent", $office2012WordNs))
                if ($null -ne $parentParagraphIdAttribute -and -not [string]::IsNullOrWhiteSpace($parentParagraphIdAttribute.Value)) {
                    $modernCommentSummary.WithParentParagraphId = [int]$modernCommentSummary.WithParentParagraphId + 1
                    $commentMetadataParentParagraphIds.Add([string]$parentParagraphIdAttribute.Value)
                }

                $doneAttribute = $commentEx.Attribute([System.Xml.Linq.XName]::Get("done", $office2012WordNs))
                if ($null -eq $doneAttribute -or [string]::IsNullOrWhiteSpace($doneAttribute.Value)) {
                    $modernCommentSummary.UnknownResolvedStateCount = [int]$modernCommentSummary.UnknownResolvedStateCount + 1
                }
                elseif (Test-OoxBooleanTrue $doneAttribute.Value) {
                    $modernCommentSummary.ResolvedCount = [int]$modernCommentSummary.ResolvedCount + 1
                }
                elseif (Test-OoxBooleanFalse $doneAttribute.Value) {
                    $modernCommentSummary.OpenCount = [int]$modernCommentSummary.OpenCount + 1
                }
                else {
                    $modernCommentSummary.UnknownResolvedStateCount = [int]$modernCommentSummary.UnknownResolvedStateCount + 1
                }
            }

            $vmlElements = @($xml.Descendants() | Where-Object { $_.Name.NamespaceName -eq $vmlNs })
            $vmlElementCount += $vmlElements.Count
            foreach ($element in $vmlElements) {
                Add-Count $vmlElementCounts $element.Name.LocalName 1
            }

            foreach ($shape in @($xml.Descendants([System.Xml.Linq.XName]::Get("shape", $vmlNs)))) {
                $vmlShapeCount++
                if (@($shape.Descendants([System.Xml.Linq.XName]::Get("imagedata", $vmlNs))).Count -gt 0) {
                    $vmlShapeWithImageDataCount++
                }

                $childNames = @($shape.Elements() |
                    Where-Object { $_.Name.NamespaceName -eq $vmlNs } |
                    ForEach-Object { $_.Name.LocalName } |
                    Sort-Object -Unique)
                $signature = if ($childNames.Count -eq 0) { "(none)" } else { [string]::Join("+", $childNames) }
                Add-Count $vmlShapeChildSignatures $signature 1
                if ($childNames.Count -eq 1 -and $childNames[0] -eq "textbox") {
                    $vmlShapeWithTextboxOnlyCount++
                }
            }

            foreach ($textBox in @($xml.Descendants([System.Xml.Linq.XName]::Get("textbox", $vmlNs)))) {
                $vmlTextBoxCount++
                if (@($textBox.Descendants([System.Xml.Linq.XName]::Get("txbxContent", $wordprocessingNs))).Count -gt 0) {
                    $vmlTextBoxWithContentCount++
                }

                $vmlTextBoxParagraphCount += @($textBox.Descendants([System.Xml.Linq.XName]::Get("p", $wordprocessingNs))).Count
                $vmlTextBoxTableCount += @($textBox.Descendants([System.Xml.Linq.XName]::Get("tbl", $wordprocessingNs))).Count
            }

            $complexFieldStack = [System.Collections.Generic.List[object]]::new()
            foreach ($element in $wordElements) {
                $localName = $element.Name.LocalName
                if ($localName -eq "fldChar") {
                    $fieldCharTypeAttribute = $element.Attribute([System.Xml.Linq.XName]::Get("fldCharType", $wordprocessingNs))
                    $fieldCharType = if ($null -eq $fieldCharTypeAttribute) { "" } else { [string]$fieldCharTypeAttribute.Value }
                    if ($fieldCharType.Equals("begin", [System.StringComparison]::OrdinalIgnoreCase)) {
                        if ($complexFieldStack.Count -ne 0 -and -not [bool]$complexFieldStack[$complexFieldStack.Count - 1].InResult) {
                            $complexFieldSummary.MalformedOrUnclosed = [int]$complexFieldSummary.MalformedOrUnclosed + 1
                        }

                        $complexFieldStack.Add((New-PrivateSafeComplexFieldState (Get-PrivateSafeComplexFieldContainer $element)))
                        continue
                    }

                    if ($fieldCharType.Equals("separate", [System.StringComparison]::OrdinalIgnoreCase)) {
                        if ($complexFieldStack.Count -eq 0) {
                            $complexFieldSummary.MalformedOrUnclosed = [int]$complexFieldSummary.MalformedOrUnclosed + 1
                            continue
                        }

                        $field = $complexFieldStack[$complexFieldStack.Count - 1]
                        $field.HasSeparate = $true
                        $field.InResult = $true
                        continue
                    }

                    if ($fieldCharType.Equals("end", [System.StringComparison]::OrdinalIgnoreCase)) {
                        if ($complexFieldStack.Count -eq 0) {
                            $complexFieldSummary.MalformedOrUnclosed = [int]$complexFieldSummary.MalformedOrUnclosed + 1
                            continue
                        }

                        $field = $complexFieldStack[$complexFieldStack.Count - 1]
                        $complexFieldStack.RemoveAt($complexFieldStack.Count - 1)
                        Add-PrivateSafeComplexFieldSummary $complexFieldSummary $field
                        continue
                    }

                    $complexFieldSummary.MalformedOrUnclosed = [int]$complexFieldSummary.MalformedOrUnclosed + 1
                    continue
                }

                if ($localName -eq "instrText") {
                    if ($complexFieldStack.Count -eq 0 -or [bool]$complexFieldStack[$complexFieldStack.Count - 1].InResult) {
                        $complexFieldSummary.MalformedOrUnclosed = [int]$complexFieldSummary.MalformedOrUnclosed + 1
                        continue
                    }

                    [void]$complexFieldStack[$complexFieldStack.Count - 1].Instruction.Append([string]$element.Value)
                    continue
                }

                if (($localName -eq "t" -or $localName -eq "delText") -and -not [string]::IsNullOrEmpty([string]$element.Value)) {
                    foreach ($field in $complexFieldStack) {
                        if ([bool]$field.InResult) {
                            $field.HasCachedResult = $true
                        }
                    }
                }
            }

            if ($complexFieldStack.Count -ne 0) {
                $complexFieldSummary.MalformedOrUnclosed = [int]$complexFieldSummary.MalformedOrUnclosed + $complexFieldStack.Count
            }
        }
    }
    finally {
        $zip.Dispose()
    }

    $revisionElementCount = 0
    foreach ($value in $revisionCounts.Values) {
        $revisionElementCount += [int]$value
    }

    $commentMarkerCount = 0
    foreach ($value in $commentMarkerCounts.Values) {
        $commentMarkerCount += [int]$value
    }

    $modernCommentSummary.CommentParagraphIdCount = $commentParagraphIds.Count
    foreach ($paragraphId in $commentMetadataParagraphIds) {
        if ($commentParagraphIds.Contains($paragraphId)) {
            $modernCommentSummary.MetadataLinkedToKnownCommentCount = [int]$modernCommentSummary.MetadataLinkedToKnownCommentCount + 1
        }
        else {
            $modernCommentSummary.MetadataWithoutKnownCommentCount = [int]$modernCommentSummary.MetadataWithoutKnownCommentCount + 1
        }
    }

    foreach ($parentParagraphId in $commentMetadataParentParagraphIds) {
        if ($commentParagraphIds.Contains($parentParagraphId)) {
            $modernCommentSummary.ParentLinkedToKnownCommentCount = [int]$modernCommentSummary.ParentLinkedToKnownCommentCount + 1
        }
        else {
            $modernCommentSummary.ParentWithoutKnownCommentCount = [int]$modernCommentSummary.ParentWithoutKnownCommentCount + 1
        }
    }

    return [ordered]@{
        XmlPartCount = $xmlPartCount
        RevisionElementCount = $revisionElementCount
        RevisionElements = $revisionCounts
        FormattingRevisionCount = $formattingRevisionSummary.Total
        FormattingRevisions = $formattingRevisionSummary
        CommentMarkerCount = $commentMarkerCount
        CommentMarkers = $commentMarkerCounts
        CommentBodyCount = $commentBodyCount
        ModernCommentElementCount = $modernCommentElementCount
        ModernComments = $modernCommentSummary
        SimpleFieldCount = $simpleFieldSummary.Total
        SimpleFields = $simpleFieldSummary
        ComplexFieldCount = $complexFieldSummary.Total
        ComplexFields = $complexFieldSummary
        VmlElementCount = $vmlElementCount
        VmlElements = $vmlElementCounts
        VmlShapeCount = $vmlShapeCount
        VmlShapeWithImageDataCount = $vmlShapeWithImageDataCount
        VmlShapeWithTextboxOnlyCount = $vmlShapeWithTextboxOnlyCount
        VmlShapeChildSignatures = $vmlShapeChildSignatures
        VmlTextBoxCount = $vmlTextBoxCount
        VmlTextBoxWithContentCount = $vmlTextBoxWithContentCount
        VmlTextBoxParagraphCount = $vmlTextBoxParagraphCount
        VmlTextBoxTableCount = $vmlTextBoxTableCount
    }
}

function Read-JsonArray([string] $Path) {
    if (-not (Test-Path -LiteralPath $Path)) {
        return @()
    }

    $json = Get-Content -Raw -LiteralPath $Path
    if ([string]::IsNullOrWhiteSpace($json) -or $json.Trim() -eq "[]") {
        return @()
    }

    return @(ConvertFrom-Json -InputObject $json)
}

$caseFull = (Resolve-Path -LiteralPath $Case).Path
Assert-PrivateUntracked $caseFull "Private case manifest"

$caseDirectory = Split-Path -Parent $caseFull
$manifest = Get-Content -Raw -LiteralPath $caseFull | ConvertFrom-Json
if ([string]::IsNullOrWhiteSpace($manifest.id)) {
    throw "Private case manifest must contain an id."
}

$caseId = [string]$manifest.id
if ($caseId.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars()) -ge 0 -or $caseId.Contains("/") -or $caseId.Contains("\")) {
    throw "Private case id must be a single filename-safe path segment."
}

if ($manifest.kind -ne $null -and -not [string]::Equals([string]$manifest.kind, "docx", [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Private DOCX markup check only supports kind 'docx'."
}

if ([string]::IsNullOrWhiteSpace($manifest.input)) {
    throw "Private case manifest must contain an input path."
}

$inputPath = Join-Path $caseDirectory $manifest.input
$inputFull = (Resolve-Path -LiteralPath $inputPath).Path
Assert-PrivateUntracked $inputFull "Private case input"
if (-not [string]::Equals([System.IO.Path]::GetExtension($inputFull), ".docx", [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Private DOCX markup check requires a .docx input."
}

if ($ValidateOnly) {
    Write-Host "Private DOCX markup case validation passed: $caseFull"
    return
}

$runId = "candidate-only-" + (Get-Date -Format "yyyyMMdd-HHmmss")
$runRoot = Join-Path $repoRoot ("artifacts/private-visual/{0}/{1}" -f $caseId, $runId)
$candidateDir = Join-Path $runRoot "candidate"
$inspectDir = Join-Path $runRoot "docx-inspect"
$inventoryDir = Join-Path $runRoot "inventory"
$pdfInspectDir = Join-Path $runRoot "pdf-inspect"
New-Item -ItemType Directory -Force -Path $candidateDir, $inventoryDir | Out-Null
if (-not $SkipDocxInspect) {
    New-Item -ItemType Directory -Force -Path $inspectDir | Out-Null
}

if (-not $SkipPdfInspect) {
    New-Item -ItemType Directory -Force -Path $pdfInspectDir | Out-Null
}

$candidatePdf = Join-Path $candidateDir "output.pdf"
$diagnosticsPath = Join-Path $candidateDir "diagnostics.json"
$summaryPath = Join-Path $runRoot "summary.json"
$featureCountsPath = Join-Path $inventoryDir "markup-feature-counts.json"

dotnet build (Join-Path $repoRoot "src/Lokad.OoxPdf.Cli/Lokad.OoxPdf.Cli.csproj") --nologo
if ($LASTEXITCODE -ne 0) {
    throw "CLI build failed with exit code $LASTEXITCODE."
}

$cliDll = Join-Path $repoRoot "src/Lokad.OoxPdf.Cli/bin/Debug/net10.0/Lokad.OoxPdf.Cli.dll"
dotnet $cliDll convert $inputFull $candidatePdf --diagnostics $diagnosticsPath --docx-markup $MarkupMode --docx-markup-geometry $MarkupGeometry
if ($LASTEXITCODE -ne 0) {
    throw "Candidate conversion failed with exit code $LASTEXITCODE."
}

if (-not $SkipDocxInspect) {
    & (Join-Path $PSScriptRoot "InspectDocx.ps1") -InputDocx $inputFull -OutputDirectory $inspectDir -DocxMarkup $MarkupMode -DocxMarkupGeometry $MarkupGeometry | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "DOCX inspection failed with exit code $LASTEXITCODE."
    }
}

if (-not $SkipPdfInspect) {
    & (Join-Path $PSScriptRoot "InspectPdf.ps1") -InputPdf $candidatePdf -OutputDirectory $pdfInspectDir -TextOnly | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "PDF inspection failed with exit code $LASTEXITCODE."
    }
}

$featureCounts = Get-DocxMarkupFeatureCounts $inputFull
$featureCounts | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $featureCountsPath -Encoding UTF8

$diagnostics = Read-JsonArray $diagnosticsPath
$diagnosticIds = @($diagnostics | ForEach-Object { $_.Id } | Sort-Object -Unique)
$markupSummaryPath = Join-Path $inspectDir "markup-summary.json"
$markupSummary = if (-not $SkipDocxInspect -and (Test-Path -LiteralPath $markupSummaryPath)) {
    Get-Content -Raw -LiteralPath $markupSummaryPath | ConvertFrom-Json
}
else {
    $null
}

$summary = [ordered]@{
    CaseId = $caseId
    RunId = $runId
    MarkupMode = $MarkupMode
    Kind = "docx"
    CandidatePdf = ConvertTo-RepoPath $candidatePdf
    Diagnostics = ConvertTo-RepoPath $diagnosticsPath
    DocxInspect = if ($SkipDocxInspect) { $null } else { ConvertTo-RepoPath $inspectDir }
    PdfInspect = if ($SkipPdfInspect) { $null } else { ConvertTo-RepoPath $pdfInspectDir }
    DocxInspectSkipped = [bool]$SkipDocxInspect
    PdfInspectSkipped = [bool]$SkipPdfInspect
    Inventory = ConvertTo-RepoPath $featureCountsPath
    PdfPageCount = Get-PdfPageCount $candidatePdf
    PdfMediaBoxes = Get-PdfMediaBoxSummary $candidatePdf
    DiagnosticCount = $diagnostics.Count
    DiagnosticIds = $diagnosticIds
    MarkupFeatureCounts = $featureCounts
    MarkupInspectionSummary = $markupSummary
    MarkupQualityCounters = if ($null -ne $markupSummary -and $markupSummary.PSObject.Properties.Name -contains "QualityCounters") {
        $markupSummary.QualityCounters
    }
    else {
        $null
    }
}
$summary | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $summaryPath -Encoding UTF8

Write-Host "Candidate-only private DOCX markup artifacts: $runRoot"
