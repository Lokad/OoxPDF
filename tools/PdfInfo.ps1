# Page facts for the plain PDF objects used by the DOCX markup tools.

function Read-PdfLatin1Text([string] $PdfPath) {
    return [Text.Encoding]::Latin1.GetString([IO.File]::ReadAllBytes($PdfPath))
}

function Get-PdfObjects([string] $PdfText) {
    $matches = [regex]::Matches($PdfText, '(?s)(?<number>\d+)\s+(?<generation>\d+)\s+obj\s*(?<body>.*?)\s*endobj')
    foreach ($match in $matches) {
        [pscustomobject]@{
            Number = [int]$match.Groups['number'].Value
            Generation = [int]$match.Groups['generation'].Value
            Body = [string]$match.Groups['body'].Value
        }
    }
}

function Read-PdfMediaBox([string] $Body, [hashtable] $ByReference, [string] $PageReference) {
    $declared = [regex]::Match($Body, '/MediaBox\b\s*(?:\[(?<box>[^\]]*)\]|(?<number>\d+)\s+(?<generation>\d+)\s+R\b)?')
    if (!$declared.Success) { return }
    $box = $declared.Groups['box'].Value
    if ($declared.Groups['number'].Success) {
        $key = "$($declared.Groups['number'].Value) $($declared.Groups['generation'].Value)"
        $box = [regex]::Match([string]$ByReference[$key].Body, '^\s*\[(?<box>[^\]]*)\]\s*$').Groups['box'].Value
    }
    $numbers = @($box -split '\s+' | Where-Object Length -gt 0)
    if ($numbers.Count -ne 4 -or @($numbers | Where-Object { $_ -cnotmatch '^[-+]?(?:\d+(?:\.\d*)?|\.\d+)$' }).Count) {
        throw "Unsupported PDF media box at page-tree reference $PageReference."
    }
    $x0, $y0, $x1, $y1 = @($numbers | ForEach-Object {
        [double]::Parse($_, [Globalization.CultureInfo]::InvariantCulture)
    })
    [pscustomobject]@{
        MinX = [Math]::Round($x0, 6)
        MinY = [Math]::Round($y0, 6)
        MaxX = [Math]::Round($x1, 6)
        MaxY = [Math]::Round($y1, 6)
        Width = [Math]::Round($x1 - $x0, 6)
        Height = [Math]::Round($y1 - $y0, 6)
    }
}

function Get-PdfPageFacts($Objects) {
    $byReference = @{}
    foreach ($object in $Objects) {
        $byReference["$($object.Number) $($object.Generation)"] = $object
    }
    $catalog = $Objects | Where-Object { $_.Body -match '/Type\s*/Catalog\b' } | Select-Object -Last 1
    $root = [regex]::Match([string]$catalog.Body, '/Pages\s+(?<number>\d+)\s+(?<generation>\d+)\s+R')
    if (!$root.Success) { throw 'PDF page tree is unavailable to the annotation inspector.' }

    # Carry inherited geometry down the same /Kids walk that assigns page numbers.
    $pageByReference = @{}
    $mediaBoxes = [Collections.Generic.List[object]]::new()
    $pending = [Collections.Generic.Stack[object]]::new()
    $pending.Push(@("$($root.Groups['number'].Value) $($root.Groups['generation'].Value)", $null))
    $visited = [Collections.Generic.HashSet[string]]::new()
    while ($pending.Count -gt 0) {
        $key, $inheritedBox = $pending.Pop()
        if (!$visited.Add($key) -or !$byReference.ContainsKey($key)) {
            throw "Unresolvable or repeated PDF page-tree reference: $key"
        }
        $body = $byReference[$key].Body
        $box = (Read-PdfMediaBox $body $byReference $key) ?? $inheritedBox
        if ($body -match '/Type\s*/Page\b') {
            if ($null -eq $box) { throw "Missing effective PDF media box at page reference $key." }
            $pageByReference[$key] = $pageByReference.Count + 1
            $mediaBoxes.Add($box)
            continue
        }
        $kids = [regex]::Match($body, '(?s)/Kids\s*\[(?<items>.*?)\]')
        if ($body -notmatch '/Type\s*/Pages\b' -or !$kids.Success) {
            throw "Unsupported PDF page-tree node: $key"
        }
        $references = [regex]::Matches($kids.Groups['items'].Value, '(?<number>\d+)\s+(?<generation>\d+)\s+R')
        for ($i = $references.Count - 1; $i -ge 0; $i--) {
            $pending.Push(@("$($references[$i].Groups['number'].Value) $($references[$i].Groups['generation'].Value)", $box))
        }
    }
    [pscustomobject]@{
        PageByReference = $pageByReference
        PageCount = $pageByReference.Count
        MediaBoxes = $mediaBoxes.ToArray()
    }
}
