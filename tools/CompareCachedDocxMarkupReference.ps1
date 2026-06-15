param(
    [Parameter(Mandatory = $true)]
    [string] $Case,

    [string] $OutputDirectory,

    [string] $ReferenceDirectory,

    [string] $ReferencePdf,

    [ValidateSet("final", "original", "simple", "all", "simple-markup", "all-markup")]
    [string] $DocxMarkup,

    [ValidateSet("preserve", "preserve-layout", "preserve-document-layout", "reserve", "reserve-margin", "markup-margin", "reserve-markup-margin", "word", "word-compatible", "word-compatible-all-markup", "office", "office-compatible", "office-compatible-all-markup")]
    [string] $DocxMarkupGeometry,

    [string] $CaseId,

    [switch] $PrivateSafeSummary,

    [int] $Dpi = 144,

    [double] $TextPositionTolerance = 0.5,

    [double] $TextFontSizeTolerance = 0.1,

    [double] $TextCharacterSpacingTolerance = 0.01,

    [double] $TextStateSpacingTolerance = 0.25,

    [double] $GraphicsBoundsTolerance = 1.0,

    [double] $AnnotationBoundsTolerance = 1.0,

    [double] $BalloonBoundsTolerance = 8.0,

    [int] $MaxPageCountDelta = 0,

    [double] $MediaBoxTolerance = 0.5,

    [double] $BodyFrameTolerance = 1.5,

    [double] $MarkupMarginTolerance = 1.5,

    [double] $TextBaselineTolerance = 0.75,

    [double] $TextXPositionTolerance = 0.75,

    [double] $GlyphAdvanceTolerance = 1.0,

    [double] $TableGridBoundsTolerance = 1.0,

    [double] $ConnectorBoundsTolerance = 2.0,

    [int] $MaxAnnotationDeltaCount = 0,

    [int] $MaxBalloonGeometryDeltaCount = 0,

    [double] $MaxChangedPixelRatioAtThreshold16 = 0.02,

    [double] $MaxMeanAbsoluteError = 3.0,

    [switch] $SkipRasterDiff,

    [switch] $FailOnDeltas,

    [switch] $ValidateOnly
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot

function Invoke-DotnetBuildIfStale {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Project,

        [Parameter(Mandatory = $true)]
        [string] $OutputDll,

        [Parameter(Mandatory = $true)]
        [string] $Description,

        [string[]] $AdditionalSourceDirectories = @()
    )

    $projectDirectory = Split-Path -Parent $Project
    $sourceDirectories = @($projectDirectory) + $AdditionalSourceDirectories
    $sourceNewest = $sourceDirectories | ForEach-Object {
        Get-ChildItem -LiteralPath $_ -Recurse -Include *.cs,*.csproj
    } |
        Where-Object { $_.FullName -notmatch '[\\/](bin|obj)[\\/]' } |
        Sort-Object LastWriteTimeUtc -Descending |
        Select-Object -First 1
    if ((Test-Path -LiteralPath $OutputDll) -and $sourceNewest.LastWriteTimeUtc -le (Get-Item -LiteralPath $OutputDll).LastWriteTimeUtc) {
        return
    }

    dotnet build $Project --nologo
    if ($LASTEXITCODE -ne 0) {
        throw "$Description build failed with exit code $LASTEXITCODE."
    }
}

function Read-JsonArray([string] $Path) {
    if (-not (Test-Path -LiteralPath $Path)) {
        return ,@()
    }

    $items = Get-Content -Raw -LiteralPath $Path | ConvertFrom-Json
    if ($null -eq $items) {
        return @()
    }

    if ($items -is [array]) {
        return $items
    }

    return @($items)
}

function Read-JsonObject([string] $Path) {
    if (-not (Test-Path -LiteralPath $Path)) {
        return $null
    }

    return Get-Content -Raw -LiteralPath $Path | ConvertFrom-Json
}

function Get-PdfObjects([string] $PdfPath) {
    $bytes = [System.IO.File]::ReadAllBytes($PdfPath)
    $text = [System.Text.Encoding]::Latin1.GetString($bytes)
    $matches = [regex]::Matches($text, '(?s)(?<number>\d+)\s+(?<generation>\d+)\s+obj\s*(?<body>.*?)\s*endobj')
    foreach ($match in $matches) {
        [pscustomobject]@{
            Number = [int]$match.Groups["number"].Value
            Generation = [int]$match.Groups["generation"].Value
            Body = [string]$match.Groups["body"].Value
        }
    }
}

function Get-Sha256Hex([string] $Value) {
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Value)
        return ([System.BitConverter]::ToString($sha256.ComputeHash($bytes)) -replace "-", "").ToLowerInvariant()
    }
    finally {
        $sha256.Dispose()
    }
}

function Decode-PdfLiteralString([string] $Value) {
    $builder = [System.Text.StringBuilder]::new()
    for ($i = 0; $i -lt $Value.Length; $i++) {
        $character = $Value[$i]
        if ($character -ne '\') {
            [void]$builder.Append($character)
            continue
        }

        if ($i + 1 -ge $Value.Length) {
            break
        }

        $i++
        $escaped = $Value[$i]
        switch ($escaped) {
            'n' { [void]$builder.Append("`n"); continue }
            'r' { [void]$builder.Append("`r"); continue }
            't' { [void]$builder.Append("`t"); continue }
            'b' { [void]$builder.Append("`b"); continue }
            'f' { [void]$builder.Append("`f"); continue }
            "`r" {
                if ($i + 1 -lt $Value.Length -and $Value[$i + 1] -eq "`n") {
                    $i++
                }
                continue
            }
            "`n" { continue }
        }

        if ($escaped -ge '0' -and $escaped -le '7') {
            $octal = [string]$escaped
            for ($j = 0; $j -lt 2 -and $i + 1 -lt $Value.Length; $j++) {
                $next = $Value[$i + 1]
                if ($next -lt '0' -or $next -gt '7') {
                    break
                }

                $octal += [string]$next
                $i++
            }

            [void]$builder.Append([char][Convert]::ToInt32($octal, 8))
            continue
        }

        [void]$builder.Append($escaped)
    }

    return $builder.ToString()
}

function Get-PdfAnnotationTargetInfo([string] $Body) {
    $uriMatch = [regex]::Match($Body, '(?s)/A\s*<<.*?/S\s*/URI\b.*?/URI\s*\((?<uri>(?:\\.|[^\\)])*)\)')
    $destinationMatch = [regex]::Match($Body, '(?s)/Dest\s*(?<dest>\[[^\]]*\]|\((?:\\.|[^\\)])*\)|/[A-Za-z0-9_.-]+)')
    $hasAction = [regex]::IsMatch($Body, '/A\s*<<')
    $hasDestination = $destinationMatch.Success
    $uriSha256 = $null
    $uriLength = $null
    if ($uriMatch.Success) {
        $uri = Decode-PdfLiteralString $uriMatch.Groups["uri"].Value
        $uriLength = $uri.Length
        $uriSha256 = Get-Sha256Hex $uri
    }

    $destinationSha256 = if ($destinationMatch.Success) {
        Get-Sha256Hex $destinationMatch.Groups["dest"].Value
    }
    else {
        $null
    }
    $targetKind = if ($uriMatch.Success) {
        "ExternalUri"
    }
    elseif ($destinationMatch.Success) {
        "InternalDestination"
    }
    elseif ($hasAction) {
        "Action"
    }
    else {
        "None"
    }

    [pscustomobject]@{
        TargetKind = $targetKind
        HasAction = $hasAction
        HasDestination = $hasDestination
        UriLength = $uriLength
        UriSha256 = $uriSha256
        DestinationSha256 = $destinationSha256
        TargetSha256 = if ($null -ne $uriSha256) { $uriSha256 } else { $destinationSha256 }
    }
}

function Get-PdfAnnotations([string] $PdfPath) {
    $objects = @(Get-PdfObjects $PdfPath)
    $pageByAnnotationObject = @{}
    $pageNumber = 0
    foreach ($object in $objects) {
        if ($object.Body -match '/Type\s*/Page\b' -and $object.Body -notmatch '/Type\s*/Pages\b') {
            $pageNumber++
            $annotsMatch = [regex]::Match($object.Body, '(?s)/Annots\s*\[(?<items>.*?)\]')
            if ($annotsMatch.Success) {
                foreach ($ref in [regex]::Matches($annotsMatch.Groups["items"].Value, '(?<number>\d+)\s+\d+\s+R')) {
                    $annotationNumber = [int]$ref.Groups["number"].Value
                    if (-not $pageByAnnotationObject.ContainsKey($annotationNumber)) {
                        $pageByAnnotationObject[$annotationNumber] = $pageNumber
                    }
                }
            }
        }
    }

    foreach ($object in $objects) {
        $subtypeMatch = [regex]::Match($object.Body, '/Subtype\s*/(?<subtype>[A-Za-z0-9]+)')
        $rectMatch = [regex]::Match($object.Body, '(?s)/Rect\s*\[(?<rect>.*?)\]')
        if (-not $subtypeMatch.Success -or -not $rectMatch.Success) {
            continue
        }

        $numbers = @([regex]::Matches($rectMatch.Groups["rect"].Value, '[-+]?(?:\d+(?:\.\d*)?|\.\d+)') | ForEach-Object {
            [double]::Parse($_.Value, [Globalization.CultureInfo]::InvariantCulture)
        })
        if ($numbers.Count -lt 4) {
            continue
        }

        $minX = [Math]::Min($numbers[0], $numbers[2])
        $maxX = [Math]::Max($numbers[0], $numbers[2])
        $minY = [Math]::Min($numbers[1], $numbers[3])
        $maxY = [Math]::Max($numbers[1], $numbers[3])
        $targetInfo = Get-PdfAnnotationTargetInfo $object.Body
        [pscustomobject]@{
            ObjectNumber = $object.Number
            Page = if ($pageByAnnotationObject.ContainsKey($object.Number)) { $pageByAnnotationObject[$object.Number] } else { $null }
            Subtype = $subtypeMatch.Groups["subtype"].Value
            TargetKind = $targetInfo.TargetKind
            HasAction = $targetInfo.HasAction
            HasDestination = $targetInfo.HasDestination
            UriLength = $targetInfo.UriLength
            UriSha256 = $targetInfo.UriSha256
            DestinationSha256 = $targetInfo.DestinationSha256
            TargetSha256 = $targetInfo.TargetSha256
            MinX = [Math]::Round($minX, 6)
            MinY = [Math]::Round($minY, 6)
            MaxX = [Math]::Round($maxX, 6)
            MaxY = [Math]::Round($maxY, 6)
            Width = [Math]::Round($maxX - $minX, 6)
            Height = [Math]::Round($maxY - $minY, 6)
        }
    }
}

function Get-CenterX($item) {
    return ([double]$item.MinX + [double]$item.MaxX) / 2d
}

function Get-CenterY($item) {
    return ([double]$item.MinY + [double]$item.MaxY) / 2d
}

function Get-OptionalPropertyValue($Item, [string] $Name) {
    if ($null -eq $Item -or $Item.PSObject.Properties.Name -notcontains $Name) {
        return $null
    }

    return $Item.PSObject.Properties[$Name].Value
}

function Get-TargetKindPenalty($Reference, $Candidate) {
    $referenceTargetKind = Get-OptionalPropertyValue $Reference "TargetKind"
    $candidateTargetKind = Get-OptionalPropertyValue $Candidate "TargetKind"
    if ($null -eq $referenceTargetKind -and $null -eq $candidateTargetKind) {
        return 0d
    }

    if ([string]$referenceTargetKind -eq [string]$candidateTargetKind) {
        return 0d
    }

    return 10000d
}

function Get-TargetHashPenalty($Reference, $Candidate) {
    $referenceTargetHash = Get-OptionalPropertyValue $Reference "TargetSha256"
    $candidateTargetHash = Get-OptionalPropertyValue $Candidate "TargetSha256"
    if ($null -eq $referenceTargetHash -and $null -eq $candidateTargetHash) {
        return 0d
    }

    if ([string]$referenceTargetHash -eq [string]$candidateTargetHash) {
        return 0d
    }

    return 1000d
}

function New-RectComparisonRow(
    [string] $Status,
    $Candidate,
    $Reference,
    $MaxDelta)
{
    $candidateTargetKind = Get-OptionalPropertyValue $Candidate "TargetKind"
    $referenceTargetKind = Get-OptionalPropertyValue $Reference "TargetKind"
    $candidateTargetHash = Get-OptionalPropertyValue $Candidate "TargetSha256"
    $referenceTargetHash = Get-OptionalPropertyValue $Reference "TargetSha256"
    $hasTargetInfo = $null -ne $candidateTargetKind -or $null -ne $referenceTargetKind -or $null -ne $candidateTargetHash -or $null -ne $referenceTargetHash
    $targetKindDelta = if ($hasTargetInfo) { [string]$candidateTargetKind -ne [string]$referenceTargetKind } else { $false }
    $targetHashDelta = if ($hasTargetInfo -and ($null -ne $candidateTargetHash -or $null -ne $referenceTargetHash)) { [string]$candidateTargetHash -ne [string]$referenceTargetHash } else { $false }
    $targetDelta = $targetKindDelta -or $targetHashDelta
    $rowStatus = if ($Status -eq "ok" -and $targetDelta) { "delta" } else { $Status }

    [pscustomobject]@{
        Status = $rowStatus
        Candidate = $Candidate
        Reference = $Reference
        MaxDelta = $MaxDelta
        CandidateTargetKind = $candidateTargetKind
        ReferenceTargetKind = $referenceTargetKind
        TargetKindDelta = $targetKindDelta
        CandidateTargetSha256 = $candidateTargetHash
        ReferenceTargetSha256 = $referenceTargetHash
        TargetSha256Delta = $targetHashDelta
        TargetDelta = $targetDelta
    }
}

function Compare-RectLists($ReferenceItems, $CandidateItems, [double] $Tolerance) {
    $unmatched = New-Object System.Collections.Generic.List[object]
    foreach ($item in @($ReferenceItems)) {
        $unmatched.Add($item)
    }

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($candidate in @($CandidateItems)) {
        $bestIndex = -1
        $bestScore = [double]::PositiveInfinity
        for ($i = 0; $i -lt $unmatched.Count; $i++) {
            $reference = $unmatched[$i]
            $pagePenalty = if ($candidate.Page -eq $reference.Page) { 0d } else { 1000000d }
            $subtypePenalty = if ($candidate.Subtype -eq $reference.Subtype) { 0d } else { 100000d }
            $score = $pagePenalty +
                $subtypePenalty +
                (Get-TargetKindPenalty $reference $candidate) +
                (Get-TargetHashPenalty $reference $candidate) +
                [Math]::Abs((Get-CenterX $candidate) - (Get-CenterX $reference)) +
                [Math]::Abs((Get-CenterY $candidate) - (Get-CenterY $reference))
            if ($score -lt $bestScore) {
                $bestScore = $score
                $bestIndex = $i
            }
        }

        $referenceItem = if ($bestIndex -ge 0) { $unmatched[$bestIndex] } else { $null }
        if ($bestIndex -ge 0) {
            $unmatched.RemoveAt($bestIndex)
        }

        if ($null -eq $referenceItem) {
            $rows.Add((New-RectComparisonRow "missing-reference" $candidate $null $null))
            continue
        }

        $deltas = @(
            [Math]::Abs([double]$candidate.MinX - [double]$referenceItem.MinX),
            [Math]::Abs([double]$candidate.MinY - [double]$referenceItem.MinY),
            [Math]::Abs([double]$candidate.MaxX - [double]$referenceItem.MaxX),
            [Math]::Abs([double]$candidate.MaxY - [double]$referenceItem.MaxY)
        )
        $maxDelta = ($deltas | Measure-Object -Maximum).Maximum
        $status = if ($maxDelta -le $Tolerance) { "ok" } else { "delta" }
        $rows.Add((New-RectComparisonRow $status $candidate $referenceItem ([Math]::Round($maxDelta, 6))))
    }

    foreach ($reference in $unmatched) {
        $rows.Add((New-RectComparisonRow "missing-candidate" $null $reference $null))
    }

    return $rows.ToArray()
}

function Get-OperationPathPoints($Operation) {
    $points = New-Object System.Collections.Generic.List[object]
    if ($null -eq $Operation -or $Operation.PSObject.Properties.Name -notcontains "PathCommands") {
        return $points.ToArray()
    }

    foreach ($command in @($Operation.PathCommands)) {
        if ($null -eq $command -or $command.PSObject.Properties.Name -notcontains "Values") {
            continue
        }

        $values = @($command.Values)
        if ($values.Count -lt 2 -or $command.Operator -notin @("m", "l")) {
            continue
        }

        $points.Add([pscustomobject]@{ X = [double]$values[0]; Y = [double]$values[1] }) | Out-Null
    }

    return $points.ToArray()
}

function Convert-ConnectorOperationToEndpoints($Operation) {
    if ($null -eq $Operation.MinX -or $null -eq $Operation.MinY -or $null -eq $Operation.MaxX -or $null -eq $Operation.MaxY) {
        return $null
    }

    $points = @(Get-OperationPathPoints $Operation)
    if ($points.Count -ge 2) {
        $start = $points[0]
        $end = $points[$points.Count - 1]
    }
    else {
        $start = [pscustomobject]@{ X = [double]$Operation.MinX; Y = [double]$Operation.MinY }
        $end = [pscustomobject]@{ X = [double]$Operation.MaxX; Y = [double]$Operation.MaxY }
    }

    [pscustomobject]@{
        Page = Get-OperationPage $Operation
        Subtype = "connector-endpoints"
        MinX = [double]$Operation.MinX
        MinY = [double]$Operation.MinY
        MaxX = [double]$Operation.MaxX
        MaxY = [double]$Operation.MaxY
        Width = [double]$Operation.MaxX - [double]$Operation.MinX
        Height = [double]$Operation.MaxY - [double]$Operation.MinY
        StartX = [double]$start.X
        StartY = [double]$start.Y
        EndX = [double]$end.X
        EndY = [double]$end.Y
        SegmentCount = if ($Operation.SegmentCount -ne $null) { [int]$Operation.SegmentCount } else { $null }
        Dash = if ($Operation.Dash -ne $null) { [string]$Operation.Dash } else { $null }
    }
}

function Select-ConnectorEndpointItems($Operations) {
    foreach ($operation in @($Operations)) {
        if (-not (Test-ConnectorLikeGraphic $operation)) {
            continue
        }

        $item = Convert-ConnectorOperationToEndpoints $operation
        if ($null -ne $item) {
            $item
        }
    }
}

function Test-ConnectorLaneEndpointItem($Item) {
    if ($null -eq $Item -or $Item.Width -eq $null -or $Item.Height -eq $null) {
        return $false
    }

    return [double]$Item.Width -le 48d -and [double]$Item.Height -gt 0.1d
}

function Get-ConnectorEndpointPairDelta($Reference, $Candidate, [bool] $Reverse) {
    if ($Reverse) {
        $deltas = @(
            [Math]::Abs([double]$Candidate.StartX - [double]$Reference.EndX),
            [Math]::Abs([double]$Candidate.StartY - [double]$Reference.EndY),
            [Math]::Abs([double]$Candidate.EndX - [double]$Reference.StartX),
            [Math]::Abs([double]$Candidate.EndY - [double]$Reference.StartY)
        )
    }
    else {
        $deltas = @(
            [Math]::Abs([double]$Candidate.StartX - [double]$Reference.StartX),
            [Math]::Abs([double]$Candidate.StartY - [double]$Reference.StartY),
            [Math]::Abs([double]$Candidate.EndX - [double]$Reference.EndX),
            [Math]::Abs([double]$Candidate.EndY - [double]$Reference.EndY)
        )
    }

    return [double](($deltas | Measure-Object -Maximum).Maximum)
}

function Compare-ConnectorEndpointLists($ReferenceItems, $CandidateItems, [double] $Tolerance) {
    $unmatched = New-Object System.Collections.Generic.List[object]
    foreach ($item in @($ReferenceItems)) {
        $unmatched.Add($item)
    }

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($candidate in @($CandidateItems)) {
        $bestIndex = -1
        $bestScore = [double]::PositiveInfinity
        $bestReversed = $false
        for ($i = 0; $i -lt $unmatched.Count; $i++) {
            $reference = $unmatched[$i]
            $pagePenalty = if ($candidate.Page -eq $reference.Page) { 0d } else { 1000000d }
            $directDelta = Get-ConnectorEndpointPairDelta $reference $candidate $false
            $reversedDelta = Get-ConnectorEndpointPairDelta $reference $candidate $true
            $isReversed = $reversedDelta -lt $directDelta
            $score = $pagePenalty + [Math]::Min($directDelta, $reversedDelta)
            if ($score -lt $bestScore) {
                $bestScore = $score
                $bestIndex = $i
                $bestReversed = $isReversed
            }
        }

        $referenceItem = if ($bestIndex -ge 0) { $unmatched[$bestIndex] } else { $null }
        if ($bestIndex -ge 0) {
            $unmatched.RemoveAt($bestIndex)
        }

        if ($null -eq $referenceItem) {
            $rows.Add([pscustomobject]@{ Status = "missing-reference"; Candidate = $candidate; Reference = $null; MaxDelta = $null; Reversed = $null })
            continue
        }

        $maxDelta = Get-ConnectorEndpointPairDelta $referenceItem $candidate $bestReversed
        $rows.Add([pscustomobject]@{
            Status = if ($maxDelta -le $Tolerance) { "ok" } else { "delta" }
            Candidate = $candidate
            Reference = $referenceItem
            MaxDelta = [Math]::Round($maxDelta, 6)
            Reversed = $bestReversed
        })
    }

    foreach ($reference in $unmatched) {
        $rows.Add([pscustomobject]@{ Status = "missing-candidate"; Candidate = $null; Reference = $reference; MaxDelta = $null; Reversed = $null })
    }

    return $rows.ToArray()
}

function Convert-BalloonToRect($balloon) {
    $laneBandIndex = if ($balloon.PSObject.Properties.Name -contains "LaneBandIndex" -and $balloon.LaneBandIndex -ne $null) { [int]$balloon.LaneBandIndex } else { $null }
    $laneBandCandidateCount = if ($balloon.PSObject.Properties.Name -contains "LaneBandCandidateCount" -and $balloon.LaneBandCandidateCount -ne $null) { [int]$balloon.LaneBandCandidateCount } else { $null }
    [pscustomobject]@{
        Page = [int]$balloon.PageIndex + 1
        Subtype = [string]$balloon.Kind
        MinX = [double]$balloon.X
        MinY = [double]$balloon.Y
        MaxX = [double]$balloon.X + [double]$balloon.Width
        MaxY = [double]$balloon.Y + [double]$balloon.Height
        Width = [double]$balloon.Width
        Height = [double]$balloon.Height
        Side = [string]$balloon.Side
        IsOverflowSummary = [bool]$balloon.IsOverflowSummary
        AnchorConnectorX = [double]$balloon.AnchorConnectorX
        BalloonConnectorX = [double]$balloon.BalloonConnectorX
        AnchorConnectorClamped = [bool]$balloon.AnchorConnectorClamped
        CandidateCount = [int]$balloon.CandidateCount
        LaneBandIndex = $laneBandIndex
        LaneBandCandidateCount = $laneBandCandidateCount
    }
}

function Select-ReferenceBalloonGraphicCandidates($graphicsOperations) {
    foreach ($op in @($graphicsOperations)) {
        if ($op.Kind -notin @("Stroke", "Fill", "FillStroke")) {
            continue
        }

        $width = [double]$op.MaxX - [double]$op.MinX
        $height = [double]$op.MaxY - [double]$op.MinY
        if ($width -lt 12d -or $height -lt 6d) {
            continue
        }

        [pscustomobject]@{
            Page = if ($op.PageNumber -ne $null) { [int]$op.PageNumber } else { $null }
            Subtype = [string]$op.Kind
            MinX = [double]$op.MinX
            MinY = [double]$op.MinY
            MaxX = [double]$op.MaxX
            MaxY = [double]$op.MaxY
            Width = $width
            Height = $height
            SourceOperator = $op.SourceOperator
        }
    }
}

function Test-RectOverlapsMarkupLane($Rect, $MarkupLaneRects) {
    $lanes = @($MarkupLaneRects | Where-Object { $null -ne $_ })
    if ($lanes.Count -eq 0) {
        return $true
    }

    foreach ($lane in $lanes) {
        if ($Rect.Page -ne $null -and $lane.Page -ne $null -and [int]$Rect.Page -ne [int]$lane.Page) {
            continue
        }

        $overlapX = [Math]::Min([double]$Rect.MaxX, [double]$lane.MaxX) - [Math]::Max([double]$Rect.MinX, [double]$lane.MinX)
        $overlapY = [Math]::Min([double]$Rect.MaxY, [double]$lane.MaxY) - [Math]::Max([double]$Rect.MinY, [double]$lane.MinY)
        if ($overlapX -gt 0d -and $overlapY -gt 0d) {
            return $true
        }
    }

    return $false
}

function Get-RectDedupeKey($Rect) {
    $page = if ($Rect.Page -ne $null) { [string][int]$Rect.Page } else { "null" }
    return "{0}|{1:F2}|{2:F2}|{3:F2}|{4:F2}" -f $page, [double]$Rect.MinX, [double]$Rect.MinY, [double]$Rect.MaxX, [double]$Rect.MaxY
}

function Select-DistinctRects($Rects) {
    $seen = @{}
    foreach ($rect in @($Rects | Where-Object { $null -ne $_ })) {
        $key = Get-RectDedupeKey $rect
        if (-not $seen.ContainsKey($key)) {
            $seen[$key] = $rect
        }
    }

    foreach ($key in @($seen.Keys | Sort-Object)) {
        $seen[$key]
    }
}

function Select-BalloonBodyGraphicCandidates($graphicsOperations, $MarkupLaneRects) {
    $items = New-Object System.Collections.Generic.List[object]
    $lanes = @($MarkupLaneRects | Where-Object { $null -ne $_ })
    foreach ($op in @($graphicsOperations)) {
        if ($op.Kind -notin @("Stroke", "Fill", "FillStroke")) {
            continue
        }

        $width = Get-OperationWidth $op
        $height = Get-OperationHeight $op
        if ($null -eq $width -or $null -eq $height) {
            continue
        }

        $laneHeightLimit = $null
        if ($lanes.Count -ne 0 -and $op.PageNumber -ne $null) {
            $pageLanes = @($lanes | Where-Object { $_.Page -eq $op.PageNumber -and $_.Height -ne $null })
            if ($pageLanes.Count -ne 0) {
                $laneHeightLimit = [double](($pageLanes | ForEach-Object { [double]$_.Height * 0.7d } | Measure-Object -Maximum).Maximum)
            }
        }

        $maxBodyHeight = if ($null -ne $laneHeightLimit) { [Math]::Max(96d, $laneHeightLimit) } else { 240d }
        if ([double]$width -lt 72d -or [double]$height -lt 8d -or [double]$height -gt $maxBodyHeight) {
            continue
        }

        $rect = Convert-GraphicOperationToRect $op "balloon-body"
        if ($null -eq $rect -or -not (Test-RectOverlapsMarkupLane $rect $lanes)) {
            continue
        }

        $items.Add([pscustomobject]@{
                Page = $rect.Page
                Subtype = "balloon-body"
                MinX = $rect.MinX
                MinY = $rect.MinY
                MaxX = $rect.MaxX
                MaxY = $rect.MaxY
                Width = $rect.Width
                Height = $rect.Height
                SourceKind = [string]$op.Kind
                SourceOperator = $op.SourceOperator
            }) | Out-Null
    }

    Select-DistinctRects $items
}

function Write-JsonFile([string] $Path, $Value, [int] $Depth = 12) {
    $Value | ConvertTo-Json -Depth $Depth | Set-Content -LiteralPath $Path -Encoding UTF8
}

function New-FontResourceSummary($fontResources) {
    $items = @($fontResources)
    $decodedLengths = @($items | Where-Object { $_.FontFileDecodedLength -ne $null } | ForEach-Object { [int]$_.FontFileDecodedLength })
    [pscustomobject]@{
        FontResourceCount = $items.Count
        DistinctFontObjectCount = @($items | ForEach-Object { $_.FontObjectNumber } | Sort-Object -Unique).Count
        EmbeddedFontResourceCount = @($items | Where-Object { $_.Embedded -eq $true }).Count
        MissingEmbeddedFontResourceCount = @($items | Where-Object { $_.Embedded -ne $true }).Count
        ToUnicodeFontResourceCount = @($items | Where-Object { $_.HasToUnicode -eq $true }).Count
        MissingToUnicodeFontResourceCount = @($items | Where-Object { $_.HasToUnicode -ne $true }).Count
        SubsetFontResourceCount = @($items | Where-Object { $_.IsSubset -eq $true }).Count
        NonSubsetFontResourceCount = @($items | Where-Object { $_.IsSubset -ne $true }).Count
        MaxFontFileDecodedLength = if ($decodedLengths.Count -eq 0) { $null } else { ($decodedLengths | Measure-Object -Maximum).Maximum }
    }
}

function Get-PdfPageCount([string] $PdfPath) {
    $bytes = [System.IO.File]::ReadAllBytes($PdfPath)
    $text = [System.Text.Encoding]::Latin1.GetString($bytes)
    return ([regex]::Matches($text, '/Type\s*/Page\b') | Where-Object { $_.Value -notmatch '/Pages' }).Count
}

function Get-PdfMediaBoxes([string] $PdfPath) {
    $bytes = [System.IO.File]::ReadAllBytes($PdfPath)
    $text = [System.Text.Encoding]::Latin1.GetString($bytes)
    $matches = [regex]::Matches($text, '(?s)/MediaBox\s*\[(?<box>.*?)\]')
    foreach ($match in $matches) {
        $numbers = @([regex]::Matches($match.Groups["box"].Value, '[-+]?(?:\d+(?:\.\d*)?|\.\d+)') | ForEach-Object {
            [double]::Parse($_.Value, [Globalization.CultureInfo]::InvariantCulture)
        })
        if ($numbers.Count -lt 4) {
            continue
        }

        [pscustomobject]@{
            MinX = [Math]::Round($numbers[0], 6)
            MinY = [Math]::Round($numbers[1], 6)
            MaxX = [Math]::Round($numbers[2], 6)
            MaxY = [Math]::Round($numbers[3], 6)
            Width = [Math]::Round($numbers[2] - $numbers[0], 6)
            Height = [Math]::Round($numbers[3] - $numbers[1], 6)
        }
    }
}

function Get-PageOperationCounts($Operations, [string] $PageProperty) {
    $counts = @{}
    foreach ($operation in @($Operations)) {
        $pageValue = $operation.$PageProperty
        if ($null -eq $pageValue) {
            continue
        }

        $page = [int]$pageValue
        if (-not $counts.ContainsKey($page)) {
            $counts[$page] = 0
        }

        $counts[$page]++
    }

    return $counts
}

function Select-VisibleGraphicsOperations($Operations) {
    @($Operations | Where-Object { $_.Kind -in @("Stroke", "Fill", "FillStroke") })
}

function Get-RectDeltaPageCounts($Rows) {
    $counts = @{}
    foreach ($row in @($Rows | Where-Object { $_.Status -ne "ok" })) {
        $page = $null
        if ($null -ne $row.Candidate -and $null -ne $row.Candidate.Page) {
            $page = [int]$row.Candidate.Page
        }
        elseif ($null -ne $row.Reference -and $null -ne $row.Reference.Page) {
            $page = [int]$row.Reference.Page
        }

        if ($null -eq $page) {
            continue
        }

        if (-not $counts.ContainsKey($page)) {
            $counts[$page] = 0
        }

        $counts[$page]++
    }

    return $counts
}

function Get-CountValue($Counts, [int] $Page) {
    if ($Counts.ContainsKey($Page)) {
        return [int]$Counts[$Page]
    }

    return 0
}

function New-PageDeltaSummary(
    [int] $ReferencePageCount,
    [int] $CandidatePageCount,
    $VisualMetrics,
    $ReferenceTextOperations,
    $CandidateTextOperations,
    $ReferenceGraphicsOperations,
    $CandidateGraphicsOperations,
    $AnnotationComparison,
    $BalloonComparison,
    $CandidateMarkupDensity)
{
    $referenceTextCounts = Get-PageOperationCounts $ReferenceTextOperations "PageNumber"
    $candidateTextCounts = Get-PageOperationCounts $CandidateTextOperations "PageNumber"
    $referenceGraphicsCounts = Get-PageOperationCounts $ReferenceGraphicsOperations "PageNumber"
    $candidateGraphicsCounts = Get-PageOperationCounts $CandidateGraphicsOperations "PageNumber"
    $annotationDeltaCounts = Get-RectDeltaPageCounts $AnnotationComparison
    $balloonDeltaCounts = Get-RectDeltaPageCounts $BalloonComparison
    $densityByPage = @{}
    foreach ($density in @($CandidateMarkupDensity)) {
        if ($density.Page -ne $null) {
            $densityByPage[[int]$density.Page] = $density
        }
    }

    $metricsByPage = @{}
    foreach ($metric in @($VisualMetrics)) {
        if ($metric.Page -ne $null) {
            $metricsByPage[[int]$metric.Page] = $metric
        }
    }

    $pageCount = [Math]::Max($ReferencePageCount, $CandidatePageCount)
    for ($page = 1; $page -le $pageCount; $page++) {
        $metric = if ($metricsByPage.ContainsKey($page)) { $metricsByPage[$page] } else { $null }
        $density = if ($densityByPage.ContainsKey($page)) { $densityByPage[$page] } else { $null }
        $textDrift = [Math]::Abs((Get-CountValue $candidateTextCounts $page) - (Get-CountValue $referenceTextCounts $page))
        $graphicsDrift = [Math]::Abs((Get-CountValue $candidateGraphicsCounts $page) - (Get-CountValue $referenceGraphicsCounts $page))
        $annotationDrift = Get-CountValue $annotationDeltaCounts $page
        $balloonDrift = Get-CountValue $balloonDeltaCounts $page
        $changedRatio = if ($null -ne $metric -and $metric.ChangedPixelRatioAtThreshold16 -ne $null) { [double]$metric.ChangedPixelRatioAtThreshold16 } else { 0d }
        $mae = if ($null -ne $metric -and $metric.MeanAbsoluteError -ne $null) { [double]$metric.MeanAbsoluteError } else { 0d }
        $markupDensity = if ($null -ne $density -and $density.MarkupSignalCount -ne $null) { [int]$density.MarkupSignalCount } else { 0 }
        $score = [Math]::Round(
            $changedRatio * 1000d +
            $mae +
            $textDrift * 0.05d +
            $graphicsDrift * 0.02d +
            $annotationDrift * 2d +
            $balloonDrift * 2d +
            $markupDensity * 0.01d,
            6)

        [pscustomobject]@{
            Page = $page
            PriorityScore = $score
            ReferencePageExists = $page -le $ReferencePageCount
            CandidatePageExists = $page -le $CandidatePageCount
            DimensionsMatch = if ($null -eq $metric) { $null } else { $metric.DimensionsMatch }
            MeanAbsoluteError = if ($null -eq $metric) { $null } else { $metric.MeanAbsoluteError }
            ChangedPixelRatioAtThreshold16 = if ($null -eq $metric) { $null } else { $metric.ChangedPixelRatioAtThreshold16 }
            StructuralSimilarity = if ($null -eq $metric) { $null } else { $metric.StructuralSimilarity }
            ReferenceTextOperationCount = Get-CountValue $referenceTextCounts $page
            CandidateTextOperationCount = Get-CountValue $candidateTextCounts $page
            TextOperationCountDelta = $textDrift
            ReferenceGraphicsOperationCount = Get-CountValue $referenceGraphicsCounts $page
            CandidateGraphicsOperationCount = Get-CountValue $candidateGraphicsCounts $page
            GraphicsOperationCountDelta = $graphicsDrift
            AnnotationDeltaCount = $annotationDrift
            BalloonGeometryDeltaCount = $balloonDrift
            MarkupSignalCount = $markupDensity
            RevisionItemCount = if ($null -eq $density) { $null } else { $density.RevisionItemCount }
            CommentReferenceItemCount = if ($null -eq $density) { $null } else { $density.CommentReferenceItemCount }
            BalloonPlacementCount = if ($null -eq $density) { $null } else { $density.BalloonPlacementCount }
            OverflowBalloonPlacementCount = if ($null -eq $density) { $null } else { $density.OverflowBalloonPlacementCount }
        }
    }
}

function New-PixelDiffTriageReport(
    $PageDeltaSummary,
    [bool] $SkipRasterDiff,
    [double] $MaxChangedPixelRatio,
    [double] $MaxMeanAbsoluteError)
{
    $items = @()
    foreach ($page in @($PageDeltaSummary)) {
        $signals = @()
        $changedRatio = if ($page.ChangedPixelRatioAtThreshold16 -ne $null) { [double]$page.ChangedPixelRatioAtThreshold16 } else { $null }
        $meanAbsoluteError = if ($page.MeanAbsoluteError -ne $null) { [double]$page.MeanAbsoluteError } else { $null }
        $structuralSimilarity = if ($page.StructuralSimilarity -ne $null) { [double]$page.StructuralSimilarity } else { $null }
        $textDelta = [int]$page.TextOperationCountDelta
        $graphicsDelta = [int]$page.GraphicsOperationCountDelta
        $annotationDelta = [int]$page.AnnotationDeltaCount
        $balloonDelta = [int]$page.BalloonGeometryDeltaCount
        $category = "clean"

        if ($SkipRasterDiff) {
            $category = "not-computed"
            $signals += "raster-skipped"
        }
        elseif ($page.ReferencePageExists -eq $true -and $page.CandidatePageExists -ne $true) {
            $category = "missing-content"
            $signals += "missing-candidate-page"
        }
        elseif ($page.ReferencePageExists -ne $true -and $page.CandidatePageExists -eq $true) {
            $category = "extra-content"
            $signals += "extra-candidate-page"
        }
        elseif ($page.DimensionsMatch -ne $true) {
            $category = "layout"
            $signals += "page-dimensions"
        }
        else {
            $hasNoRasterDelta =
                $changedRatio -ne $null -and
                $meanAbsoluteError -ne $null -and
                $changedRatio -eq 0d -and
                $meanAbsoluteError -eq 0d

            if (-not $hasNoRasterDelta) {
                if ($textDelta -ne 0) {
                    $signals += "text-operations"
                }
                if ($graphicsDelta -ne 0) {
                    $signals += "graphics-operations"
                }
                if ($annotationDelta -ne 0) {
                    $signals += "annotations"
                }
                if ($balloonDelta -ne 0) {
                    $signals += "balloons"
                }

                $isAntialiasingOnly =
                    $signals.Count -eq 0 -and
                    $changedRatio -ne $null -and
                    $meanAbsoluteError -ne $null -and
                    $changedRatio -le 0.002d -and
                    $meanAbsoluteError -le 1.0d -and
                    ($structuralSimilarity -eq $null -or $structuralSimilarity -ge 0.995d)

                if ($isAntialiasingOnly) {
                    $category = "antialiasing-only"
                    $signals += "low-raster-noise"
                }
                elseif ($annotationDelta -ne 0 -or $balloonDelta -ne 0) {
                    $category = "markup-ui"
                }
                elseif ($textDelta -ne 0 -and $graphicsDelta -eq 0) {
                    $category = "text"
                }
                elseif ($graphicsDelta -ne 0 -and $textDelta -eq 0) {
                    $category = "graphics"
                }
                elseif ($textDelta -ne 0 -and $graphicsDelta -ne 0) {
                    $category = "layout"
                }
                elseif (($changedRatio -ne $null -and $changedRatio -gt $MaxChangedPixelRatio) -or
                    ($meanAbsoluteError -ne $null -and $meanAbsoluteError -gt $MaxMeanAbsoluteError)) {
                    $category = "raster-only"
                    $signals += "unexplained-raster-delta"
                }
            }
        }

        $severity = if ($category -in @("missing-content", "extra-content")) {
            "blocking"
        }
        elseif ($changedRatio -ne $null -and $changedRatio -gt [Math]::Max(0.05d, $MaxChangedPixelRatio * 5d)) {
            "high"
        }
        elseif ($changedRatio -ne $null -and $changedRatio -gt $MaxChangedPixelRatio) {
            "medium"
        }
        elseif ($category -notin @("clean", "antialiasing-only", "not-computed")) {
            "low"
        }
        else {
            "none"
        }

        $items += [pscustomobject]@{
            Page = [int]$page.Page
            Category = $category
            Severity = $severity
            Signals = $signals
            PriorityScore = $page.PriorityScore
            MeanAbsoluteError = $page.MeanAbsoluteError
            ChangedPixelRatioAtThreshold16 = $page.ChangedPixelRatioAtThreshold16
            StructuralSimilarity = $page.StructuralSimilarity
            DimensionsMatch = $page.DimensionsMatch
            TextOperationCountDelta = $textDelta
            GraphicsOperationCountDelta = $graphicsDelta
            AnnotationDeltaCount = $annotationDelta
            BalloonGeometryDeltaCount = $balloonDelta
            MarkupSignalCount = $page.MarkupSignalCount
        }
    }

    $categoryCounts = @{}
    foreach ($item in $items) {
        if (-not $categoryCounts.ContainsKey($item.Category)) {
            $categoryCounts[$item.Category] = 0
        }

        $categoryCounts[$item.Category]++
    }

    [pscustomobject]@{
        Summary = [pscustomobject]@{
            RasterDiffSkipped = $SkipRasterDiff
            PageCount = $items.Count
            CategoryCounts = $categoryCounts
            HighOrBlockingCount = @($items | Where-Object { $_.Severity -in @("high", "blocking") }).Count
            AntialiasingOnlyCount = @($items | Where-Object { $_.Category -eq "antialiasing-only" }).Count
            RasterOnlyCount = @($items | Where-Object { $_.Category -eq "raster-only" }).Count
            TextCount = @($items | Where-Object { $_.Category -eq "text" }).Count
            GraphicsCount = @($items | Where-Object { $_.Category -eq "graphics" }).Count
            MarkupUiCount = @($items | Where-Object { $_.Category -eq "markup-ui" }).Count
            LayoutCount = @($items | Where-Object { $_.Category -eq "layout" }).Count
            MissingContentCount = @($items | Where-Object { $_.Category -eq "missing-content" }).Count
        }
        Pages = @($items | Sort-Object -Property @{ Expression = "PriorityScore"; Descending = $true }, Page)
        TopPages = @($items | Sort-Object -Property @{ Expression = "PriorityScore"; Descending = $true }, Page | Select-Object -First 10)
    }
}

function New-RegionDeltaSummary(
    $ReferenceGraphicsOperations,
    $CandidateGraphicsOperations,
    $ReferenceTextOperations,
    $CandidateTextOperations,
    $CandidateMarkupDensity,
    $ReferenceBalloonGraphics,
    $CandidateBalloons)
{
    $script:markupDensePages = @($CandidateMarkupDensity |
        Where-Object { $_.MarkupSignalCount -ne $null -and [int]$_.MarkupSignalCount -gt 0 } |
        ForEach-Object { [int]$_.Page })

    @(
        New-OperationRegionSummary `
            -Name "body-text" `
            -ReferenceTextOperations $ReferenceTextOperations `
            -CandidateTextOperations $CandidateTextOperations `
            -ReferenceGraphicsOperations @() `
            -CandidateGraphicsOperations @() `
            -Predicate { param($op) (Test-BodyTextRegion $op) }
        New-OperationRegionSummary `
            -Name "headers-footers" `
            -ReferenceTextOperations $ReferenceTextOperations `
            -CandidateTextOperations $CandidateTextOperations `
            -ReferenceGraphicsOperations $ReferenceGraphicsOperations `
            -CandidateGraphicsOperations $CandidateGraphicsOperations `
            -Predicate { param($op) (Test-HeaderFooterRegion $op) }
        New-OperationRegionSummary `
            -Name "footnotes-endnotes" `
            -ReferenceTextOperations $ReferenceTextOperations `
            -CandidateTextOperations $CandidateTextOperations `
            -ReferenceGraphicsOperations $ReferenceGraphicsOperations `
            -CandidateGraphicsOperations $CandidateGraphicsOperations `
            -Predicate { param($op) (Test-FootnoteEndnoteRegion $op) }
        New-OperationRegionSummary `
            -Name "tables" `
            -ReferenceTextOperations @() `
            -CandidateTextOperations @() `
            -ReferenceGraphicsOperations $ReferenceGraphicsOperations `
            -CandidateGraphicsOperations $CandidateGraphicsOperations `
            -Predicate { param($op) (Test-TableLikeGraphic $op) }
        New-OperationRegionSummary `
            -Name "drawings" `
            -ReferenceTextOperations @() `
            -CandidateTextOperations @() `
            -ReferenceGraphicsOperations $ReferenceGraphicsOperations `
            -CandidateGraphicsOperations $CandidateGraphicsOperations `
            -Predicate { param($op) (Test-DrawingLikeGraphic $op) }
        New-BalloonRegionSummary `
            -Name "comment-balloons" `
            -ReferenceBalloonGraphics $ReferenceBalloonGraphics `
            -CandidateBalloons @($CandidateBalloons | Where-Object { $_.Subtype -eq "Comment" })
        New-BalloonRegionSummary `
            -Name "revision-balloons" `
            -ReferenceBalloonGraphics $ReferenceBalloonGraphics `
            -CandidateBalloons @($CandidateBalloons | Where-Object { $_.Subtype -eq "Revision" })
        New-OperationRegionSummary `
            -Name "connectors" `
            -ReferenceTextOperations @() `
            -CandidateTextOperations @() `
            -ReferenceGraphicsOperations $ReferenceGraphicsOperations `
            -CandidateGraphicsOperations $CandidateGraphicsOperations `
            -Predicate { param($op) (Test-ConnectorLikeGraphic $op) }
        New-OperationRegionSummary `
            -Name "markup-dense-pages" `
            -ReferenceTextOperations $ReferenceTextOperations `
            -CandidateTextOperations $CandidateTextOperations `
            -ReferenceGraphicsOperations $ReferenceGraphicsOperations `
            -CandidateGraphicsOperations $CandidateGraphicsOperations `
            -Predicate { param($op) (Get-OperationPage $op) -ne $null -and $script:markupDensePages.Contains((Get-OperationPage $op)) }
        New-OperationRegionSummary `
            -Name "right-markup-lane" `
            -ReferenceTextOperations $ReferenceTextOperations `
            -CandidateTextOperations $CandidateTextOperations `
            -ReferenceGraphicsOperations $ReferenceGraphicsOperations `
            -CandidateGraphicsOperations $CandidateGraphicsOperations `
            -Predicate { param($op) (Get-OperationX $op) -ne $null -and (Get-OperationX $op) -gt 420d }
        New-OperationRegionSummary `
            -Name "left-markup-lane" `
            -ReferenceTextOperations $ReferenceTextOperations `
            -CandidateTextOperations $CandidateTextOperations `
            -ReferenceGraphicsOperations $ReferenceGraphicsOperations `
            -CandidateGraphicsOperations $CandidateGraphicsOperations `
            -Predicate { param($op) (Get-OperationX $op) -ne $null -and (Get-OperationX $op) -lt 120d }
    )
}

function New-OperationRegionSummary(
    [string] $Name,
    $ReferenceTextOperations,
    $CandidateTextOperations,
    $ReferenceGraphicsOperations,
    $CandidateGraphicsOperations,
    [scriptblock] $Predicate)
{
    $referenceText = @($ReferenceTextOperations | Where-Object { & $Predicate $_ })
    $candidateText = @($CandidateTextOperations | Where-Object { & $Predicate $_ })
    $referenceGraphics = @($ReferenceGraphicsOperations | Where-Object { & $Predicate $_ })
    $candidateGraphics = @($CandidateGraphicsOperations | Where-Object { & $Predicate $_ })
    [pscustomobject]@{
        Region = $Name
        ReferenceTextOperationCount = $referenceText.Count
        CandidateTextOperationCount = $candidateText.Count
        TextOperationCountDelta = [Math]::Abs($candidateText.Count - $referenceText.Count)
        TextSpacingDelta = New-TextGateDeltaSummary $referenceText $candidateText
        ReferenceGraphicsOperationCount = $referenceGraphics.Count
        CandidateGraphicsOperationCount = $candidateGraphics.Count
        GraphicsOperationCountDelta = [Math]::Abs($candidateGraphics.Count - $referenceGraphics.Count)
        ReferenceBalloonCandidateCount = $null
        CandidateBalloonCount = $null
        BalloonCountDelta = $null
    }
}

function New-BalloonRegionSummary(
    [string] $Name,
    $ReferenceBalloonGraphics,
    $CandidateBalloons)
{
    [pscustomobject]@{
        Region = $Name
        ReferenceTextOperationCount = 0
        CandidateTextOperationCount = 0
        TextOperationCountDelta = 0
        ReferenceGraphicsOperationCount = 0
        CandidateGraphicsOperationCount = 0
        GraphicsOperationCountDelta = 0
        ReferenceBalloonCandidateCount = @($ReferenceBalloonGraphics).Count
        CandidateBalloonCount = @($CandidateBalloons).Count
        BalloonCountDelta = [Math]::Abs(@($CandidateBalloons).Count - @($ReferenceBalloonGraphics).Count)
    }
}

function Get-OperationPage($Operation) {
    if ($Operation.PageNumber -ne $null) {
        return [int]$Operation.PageNumber
    }

    if ($Operation.Page -ne $null) {
        return [int]$Operation.Page
    }

    return $null
}

function Get-OperationX($Operation) {
    if ($Operation.X -ne $null) {
        return [double]$Operation.X
    }

    if ($Operation.MinX -ne $null -and $Operation.MaxX -ne $null) {
        return ([double]$Operation.MinX + [double]$Operation.MaxX) / 2d
    }

    return $null
}

function Get-OperationY($Operation) {
    if ($Operation.Y -ne $null) {
        return [double]$Operation.Y
    }

    if ($Operation.MinY -ne $null -and $Operation.MaxY -ne $null) {
        return ([double]$Operation.MinY + [double]$Operation.MaxY) / 2d
    }

    return $null
}

function Get-OperationWidth($Operation) {
    if ($Operation.Width -ne $null) {
        return [double]$Operation.Width
    }

    if ($Operation.MinX -ne $null -and $Operation.MaxX -ne $null) {
        return [Math]::Abs([double]$Operation.MaxX - [double]$Operation.MinX)
    }

    return $null
}

function Get-OperationHeight($Operation) {
    if ($Operation.Height -ne $null) {
        return [double]$Operation.Height
    }

    if ($Operation.MinY -ne $null -and $Operation.MaxY -ne $null) {
        return [Math]::Abs([double]$Operation.MaxY - [double]$Operation.MinY)
    }

    return $null
}

function Test-BodyTextRegion($Operation) {
    $x = Get-OperationX $Operation
    $y = Get-OperationY $Operation
    return $null -ne $x -and $null -ne $y -and $x -ge 120d -and $x -le 460d -and $y -ge 96d -and $y -le 720d
}

function Test-HeaderFooterRegion($Operation) {
    $y = Get-OperationY $Operation
    return $null -ne $y -and ($y -lt 72d -or $y -gt 720d)
}

function Test-FootnoteEndnoteRegion($Operation) {
    $y = Get-OperationY $Operation
    $x = Get-OperationX $Operation
    return $null -ne $x -and $null -ne $y -and $x -ge 120d -and $x -le 460d -and $y -ge 72d -and $y -lt 180d
}

function Test-TableLikeGraphic($Operation) {
    $width = Get-OperationWidth $Operation
    $height = Get-OperationHeight $Operation
    return $Operation.Kind -in @("Stroke", "FillStroke", "Clip") -and
        $null -ne $width -and $null -ne $height -and
        ($width -gt 24d -or $height -gt 8d) -and
        [int]$Operation.SegmentCount -ge 1
}

function Test-DrawingLikeGraphic($Operation) {
    $width = Get-OperationWidth $Operation
    $height = Get-OperationHeight $Operation
    return $null -ne $width -and $null -ne $height -and
        $width -gt 24d -and $height -gt 24d -and
        $Operation.Kind -in @("Fill", "FillStroke", "Clip")
}

function Test-ConnectorLikeGraphic($Operation) {
    $width = Get-OperationWidth $Operation
    $height = Get-OperationHeight $Operation
    return $Operation.Kind -eq "Stroke" -and
        $null -ne $width -and $null -ne $height -and
        [int]$Operation.SegmentCount -le 2 -and
        (($width -gt 18d -and $height -lt 18d) -or ($height -gt 18d -and $width -lt 18d) -or ($width -gt 12d -and $height -gt 12d))
}

function Get-DoubleMetric($Item, [string[]] $Names) {
    if ($null -eq $Item) {
        return $null
    }

    foreach ($name in $Names) {
        if ($Item.PSObject.Properties.Name -contains $name -and $null -ne $Item.$name) {
            return [double]$Item.$name
        }
    }

    return $null
}

function Update-MaxDelta([ref] $Maximum, $Reference, $Candidate) {
    if ($null -eq $Reference -or $null -eq $Candidate) {
        return $false
    }

    $delta = [Math]::Abs([double]$Candidate - [double]$Reference)
    if ($null -eq $Maximum.Value -or $delta -gt [double]$Maximum.Value) {
        $Maximum.Value = $delta
    }

    return $true
}

function Get-MaxOrZero($Values) {
    $items = @($Values | Where-Object { $null -ne $_ })
    if ($items.Count -eq 0) {
        return 0d
    }

    return [double](($items | Measure-Object -Maximum).Maximum)
}

function Get-MedianOrNull($Values) {
    $items = @($Values |
        Where-Object { $null -ne $_ } |
        ForEach-Object { [double]$_ } |
        Sort-Object)
    if ($items.Count -eq 0) {
        return $null
    }

    $middle = [int][Math]::Floor($items.Count / 2)
    if ($items.Count % 2 -eq 1) {
        return [double]$items[$middle]
    }

    return ([double]$items[$middle - 1] + [double]$items[$middle]) / 2d
}

function Get-MaxAbsResidualOrNull($Values, $Center) {
    if ($null -eq $Center) {
        return $null
    }

    $residuals = @($Values |
        Where-Object { $null -ne $_ } |
        ForEach-Object { [Math]::Abs([double]$_ - [double]$Center) })
    if ($residuals.Count -eq 0) {
        return $null
    }

    return [double](($residuals | Measure-Object -Maximum).Maximum)
}

function Get-RectComparisonMaxDelta($Rows) {
    return Get-MaxOrZero (@($Rows | ForEach-Object { $_.MaxDelta }))
}

function Convert-GraphicOperationToRect($Operation, [string] $Subtype) {
    if ($null -eq $Operation.MinX -or $null -eq $Operation.MinY -or $null -eq $Operation.MaxX -or $null -eq $Operation.MaxY) {
        return $null
    }

    [pscustomobject]@{
        Page = Get-OperationPage $Operation
        Subtype = $Subtype
        MinX = [double]$Operation.MinX
        MinY = [double]$Operation.MinY
        MaxX = [double]$Operation.MaxX
        MaxY = [double]$Operation.MaxY
        Width = [double]$Operation.MaxX - [double]$Operation.MinX
        Height = [double]$Operation.MaxY - [double]$Operation.MinY
    }
}

function Select-GraphicRects($Operations, [scriptblock] $Predicate, [string] $Subtype) {
    foreach ($operation in @($Operations)) {
        if (-not (& $Predicate $operation)) {
            continue
        }

        $rect = Convert-GraphicOperationToRect $operation $Subtype
        if ($null -ne $rect) {
            $rect
        }
    }
}

function New-RectDeltaSummary([string] $Name, $ReferenceRects, $CandidateRects, $ComparisonRows) {
    $deltaRows = @($ComparisonRows | Where-Object { $_.Status -ne "ok" })
    $targetDeltaRows = @($ComparisonRows | Where-Object { $_.TargetDelta -eq $true })
    [pscustomobject]@{
        Name = $Name
        ReferenceCount = @($ReferenceRects).Count
        CandidateCount = @($CandidateRects).Count
        CountDelta = [Math]::Abs(@($CandidateRects).Count - @($ReferenceRects).Count)
        DeltaCount = $deltaRows.Count
        MaxBoundsDelta = [Math]::Round((Get-RectComparisonMaxDelta $ComparisonRows), 6)
        TargetDeltaCount = $targetDeltaRows.Count
        TargetKindDeltaCount = @($ComparisonRows | Where-Object { $_.TargetKindDelta -eq $true }).Count
        TargetSha256DeltaCount = @($ComparisonRows | Where-Object { $_.TargetSha256Delta -eq $true }).Count
    }
}

function New-ConnectorEndpointDeltaSummary($ReferenceItems, $CandidateItems, $ComparisonRows, [string] $Name = "connector-endpoints") {
    $deltaRows = @($ComparisonRows | Where-Object { $_.Status -ne "ok" })
    [pscustomobject]@{
        Name = $Name
        ReferenceCount = @($ReferenceItems).Count
        CandidateCount = @($CandidateItems).Count
        CountDelta = [Math]::Abs(@($CandidateItems).Count - @($ReferenceItems).Count)
        DeltaCount = $deltaRows.Count
        MissingReferenceCount = @($ComparisonRows | Where-Object { $_.Status -eq "missing-reference" }).Count
        MissingCandidateCount = @($ComparisonRows | Where-Object { $_.Status -eq "missing-candidate" }).Count
        MaxEndpointDelta = [Math]::Round((Get-RectComparisonMaxDelta $ComparisonRows), 6)
    }
}

function Get-LayoutSnapshotPages($LayoutSnapshot) {
    if ($null -ne $LayoutSnapshot -and $LayoutSnapshot.PSObject.Properties.Name -contains "Pages") {
        return @($LayoutSnapshot.Pages | Where-Object { $null -ne $_ })
    }

    return @()
}

function Get-LayoutPageNumber($LayoutPage, [int] $Index) {
    if ($null -ne $LayoutPage -and
        $LayoutPage.PSObject.Properties.Name -contains "PageNumber" -and
        $null -ne $LayoutPage.PageNumber) {
        return [int]$LayoutPage.PageNumber
    }

    return $Index + 1
}

function Get-LayoutColumnFrames($LayoutPage) {
    if ($null -ne $LayoutPage -and $LayoutPage.PSObject.Properties.Name -contains "ColumnFrames") {
        return @($LayoutPage.ColumnFrames | Where-Object { $null -ne $_ })
    }

    return @()
}

function New-LayoutRect(
    [int] $Page,
    [string] $Subtype,
    [double] $MinX,
    [double] $MinY,
    [double] $MaxX,
    [double] $MaxY,
    [string] $Side = "",
    [int] $ColumnIndex = -1)
{
    [pscustomobject]@{
        Page = $Page
        Subtype = $Subtype
        MinX = [Math]::Round($MinX, 6)
        MinY = [Math]::Round($MinY, 6)
        MaxX = [Math]::Round($MaxX, 6)
        MaxY = [Math]::Round($MaxY, 6)
        Width = [Math]::Round([Math]::Max(0d, $MaxX - $MinX), 6)
        Height = [Math]::Round([Math]::Max(0d, $MaxY - $MinY), 6)
        Side = if ([string]::IsNullOrWhiteSpace($Side)) { $null } else { $Side }
        ColumnIndex = if ($ColumnIndex -lt 0) { $null } else { $ColumnIndex }
    }
}

function New-LayoutBodyFrameRect($LayoutPage, [int] $PageNumber) {
    $width = Get-DoubleMetric $LayoutPage @("Width")
    $height = Get-DoubleMetric $LayoutPage @("Height")
    $marginLeft = Get-DoubleMetric $LayoutPage @("MarginLeft")
    $marginRight = Get-DoubleMetric $LayoutPage @("MarginRight")
    $marginTop = Get-DoubleMetric $LayoutPage @("MarginTop")
    $marginBottom = Get-DoubleMetric $LayoutPage @("MarginBottom")
    if ($null -eq $width -or $null -eq $height -or $null -eq $marginLeft -or
        $null -eq $marginRight -or $null -eq $marginTop -or $null -eq $marginBottom) {
        return $null
    }

    $frames = @(Get-LayoutColumnFrames $LayoutPage)
    if ($frames.Count -eq 0) {
        $minX = [double]$marginLeft
        $maxX = [double]$width - [double]$marginRight
    }
    else {
        $minX = [double](($frames | Measure-Object -Property X -Minimum).Minimum)
        $maxX = [double](($frames | ForEach-Object { [double]$_.X + [double]$_.Width } | Measure-Object -Maximum).Maximum)
    }

    $minY = [double]$marginBottom
    $maxY = [double]$height - [double]$marginTop
    if ($maxX -le $minX -or $maxY -le $minY) {
        return $null
    }

    New-LayoutRect $PageNumber "candidate-layout-body-frame" $minX $minY $maxX $maxY
}

function Select-LayoutBodyFrameRects($LayoutSnapshot) {
    $pages = @(Get-LayoutSnapshotPages $LayoutSnapshot)
    for ($i = 0; $i -lt $pages.Count; $i++) {
        $rect = New-LayoutBodyFrameRect $pages[$i] (Get-LayoutPageNumber $pages[$i] $i)
        if ($null -ne $rect) {
            $rect
        }
    }
}

function Select-LayoutColumnFrameRects($LayoutSnapshot) {
    $pages = @(Get-LayoutSnapshotPages $LayoutSnapshot)
    for ($i = 0; $i -lt $pages.Count; $i++) {
        $page = $pages[$i]
        $pageNumber = Get-LayoutPageNumber $page $i
        $height = Get-DoubleMetric $page @("Height")
        $marginTop = Get-DoubleMetric $page @("MarginTop")
        $marginBottom = Get-DoubleMetric $page @("MarginBottom")
        if ($null -eq $height -or $null -eq $marginTop -or $null -eq $marginBottom) {
            continue
        }

        $minY = [double]$marginBottom
        $maxY = [double]$height - [double]$marginTop
        foreach ($frame in @(Get-LayoutColumnFrames $page)) {
            if ($frame.X -eq $null -or $frame.Width -eq $null) {
                continue
            }

            $minX = [double]$frame.X
            $maxX = $minX + [double]$frame.Width
            $columnIndex = if ($frame.Index -ne $null) { [int]$frame.Index } else { -1 }
            if ($maxX -gt $minX -and $maxY -gt $minY) {
                New-LayoutRect $pageNumber "candidate-layout-column-frame" $minX $minY $maxX $maxY "" $columnIndex
            }
        }
    }
}

function New-LayoutMarkupLaneRect($LayoutPage, [int] $PageNumber) {
    $width = Get-DoubleMetric $LayoutPage @("Width")
    $height = Get-DoubleMetric $LayoutPage @("Height")
    $marginLeft = Get-DoubleMetric $LayoutPage @("MarginLeft")
    $marginRight = Get-DoubleMetric $LayoutPage @("MarginRight")
    $marginTop = Get-DoubleMetric $LayoutPage @("MarginTop")
    $marginBottom = Get-DoubleMetric $LayoutPage @("MarginBottom")
    if ($null -eq $width -or $null -eq $height -or $null -eq $marginLeft -or
        $null -eq $marginRight -or $null -eq $marginTop -or $null -eq $marginBottom) {
        return $null
    }

    $leftAvailable = [Math]::Max(0d, [double]$marginLeft - 8d)
    $rightAvailable = [Math]::Max(0d, [double]$marginRight - 8d)
    $useLeft = $leftAvailable -gt $rightAvailable -and $leftAvailable -ge 24d
    $laneWidth = [Math]::Max(24d, $(if ($useLeft) { $leftAvailable } else { $rightAvailable }))
    if ($useLeft) {
        $laneX = [Math]::Max(2d, [double]$marginLeft - $laneWidth - 4d)
        $side = "Left"
    }
    else {
        $laneX = [Math]::Min([double]$width - $laneWidth - 2d, [double]$width - [double]$marginRight + 4d)
        $side = "Right"
    }

    $minY = [double]$marginBottom
    $maxY = [double]$height - [double]$marginTop
    if ($laneWidth -le 0d -or $maxY -le $minY) {
        return $null
    }

    New-LayoutRect $PageNumber "candidate-layout-markup-lane" $laneX $minY ($laneX + $laneWidth) $maxY $side
}

function Select-LayoutMarkupLaneRects($LayoutSnapshot) {
    $pages = @(Get-LayoutSnapshotPages $LayoutSnapshot)
    for ($i = 0; $i -lt $pages.Count; $i++) {
        $rect = New-LayoutMarkupLaneRect $pages[$i] (Get-LayoutPageNumber $pages[$i] $i)
        if ($null -ne $rect) {
            $rect
        }
    }
}

function Select-ReferenceMarkupLaneRects($ReferenceBalloonGraphics, $ReferenceMediaBoxes) {
    $mediaByPage = @{}
    $mediaBoxes = @($ReferenceMediaBoxes)
    for ($i = 0; $i -lt $mediaBoxes.Count; $i++) {
        $mediaByPage[$i + 1] = $mediaBoxes[$i]
    }

    foreach ($group in @($ReferenceBalloonGraphics | Where-Object { $_.Page -ne $null } | Group-Object -Property Page)) {
        $page = [int]$group.Name
        $pageWidth = if ($mediaByPage.ContainsKey($page) -and $mediaByPage[$page].Width -ne $null) { [double]$mediaByPage[$page].Width } else { 0d }
        $items = @($group.Group | Where-Object {
                if ($pageWidth -le 0d) {
                    $true
                }
                else {
                    $itemWidth = [double]$_.MaxX - [double]$_.MinX
                    $centerX = ([double]$_.MinX + [double]$_.MaxX) / 2d
                    if ($itemWidth -gt $pageWidth * 0.45d) {
                        $false
                    }
                    elseif ($centerX -ge $pageWidth / 2d) {
                        [double]$_.MinX -ge $pageWidth * 0.45d
                    }
                    else {
                        [double]$_.MaxX -le $pageWidth * 0.55d
                    }
                }
            })
        if ($items.Count -eq 0) {
            continue
        }

        if ($pageWidth -gt 0d) {
            $rightItems = @($items | Where-Object { (([double]$_.MinX + [double]$_.MaxX) / 2d) -ge $pageWidth / 2d })
            $leftItems = @($items | Where-Object { (([double]$_.MinX + [double]$_.MaxX) / 2d) -lt $pageWidth / 2d })
            if ($rightItems.Count -ne 0 -or $leftItems.Count -ne 0) {
                $rightArea = ($rightItems | ForEach-Object { ([double]$_.MaxX - [double]$_.MinX) * ([double]$_.MaxY - [double]$_.MinY) } | Measure-Object -Sum).Sum
                $leftArea = ($leftItems | ForEach-Object { ([double]$_.MaxX - [double]$_.MinX) * ([double]$_.MaxY - [double]$_.MinY) } | Measure-Object -Sum).Sum
                $items = if ($rightItems.Count -gt $leftItems.Count -or
                    ($rightItems.Count -eq $leftItems.Count -and [double]$rightArea -ge [double]$leftArea)) {
                    $rightItems
                }
                else {
                    $leftItems
                }
            }
        }

        $minX = [double](($items | Measure-Object -Property MinX -Minimum).Minimum)
        $maxX = [double](($items | Measure-Object -Property MaxX -Maximum).Maximum)
        $minY = [double](($items | Measure-Object -Property MinY -Minimum).Minimum)
        $maxY = [double](($items | Measure-Object -Property MaxY -Maximum).Maximum)
        $side = if ($pageWidth -le 0d) {
            $null
        }
        elseif (($minX + $maxX) / 2d -ge $pageWidth / 2d) {
            "Right"
        }
        else {
            "Left"
        }

        New-LayoutRect $page "reference-markup-lane-occupied" $minX $minY $maxX $maxY $side
    }
}

function Compare-MarkupLaneHorizontalRects($ReferenceLaneRects, $CandidateLaneRects, [double] $Tolerance) {
    $candidateByPage = @{}
    foreach ($candidate in @($CandidateLaneRects)) {
        if ($candidate.Page -ne $null -and -not $candidateByPage.ContainsKey([int]$candidate.Page)) {
            $candidateByPage[[int]$candidate.Page] = $candidate
        }
    }

    $referenceByPage = @{}
    foreach ($reference in @($ReferenceLaneRects)) {
        if ($reference.Page -ne $null -and -not $referenceByPage.ContainsKey([int]$reference.Page)) {
            $referenceByPage[[int]$reference.Page] = $reference
        }
    }

    $pages = New-Object System.Collections.Generic.HashSet[int]
    foreach ($page in $candidateByPage.Keys) { [void]$pages.Add([int]$page) }
    foreach ($page in $referenceByPage.Keys) { [void]$pages.Add([int]$page) }

    foreach ($page in @($pages | Sort-Object)) {
        $reference = if ($referenceByPage.ContainsKey($page)) { $referenceByPage[$page] } else { $null }
        $candidate = if ($candidateByPage.ContainsKey($page)) { $candidateByPage[$page] } else { $null }
        if ($null -eq $reference) {
            [pscustomobject]@{
                Page = $page
                Status = "missing-reference"
                Reference = $null
                Candidate = $candidate
                SideMatches = $null
                MaxHorizontalDelta = $null
            }
            continue
        }

        if ($null -eq $candidate) {
            [pscustomobject]@{
                Page = $page
                Status = "missing-candidate"
                Reference = $reference
                Candidate = $null
                SideMatches = $null
                MaxHorizontalDelta = $null
            }
            continue
        }

        $sideMatches = [string]$reference.Side -eq [string]$candidate.Side
        $bodyEdgeDelta = if ($sideMatches -and [string]$reference.Side -eq "Right") {
            [Math]::Abs([double]$reference.MinX - [double]$candidate.MinX)
        }
        elseif ($sideMatches -and [string]$reference.Side -eq "Left") {
            [Math]::Abs([double]$reference.MaxX - [double]$candidate.MaxX)
        }
        else {
            @(
                [Math]::Abs([double]$reference.MinX - [double]$candidate.MinX),
                [Math]::Abs([double]$reference.MaxX - [double]$candidate.MaxX)
            ) | Measure-Object -Minimum | Select-Object -ExpandProperty Minimum
        }
        $widthDelta = [Math]::Abs([double]$reference.Width - [double]$candidate.Width)
        [pscustomobject]@{
            Page = $page
            Status = if ($sideMatches -and $bodyEdgeDelta -le $Tolerance) { "ok" } else { "delta" }
            Reference = $reference
            Candidate = $candidate
            SideMatches = $sideMatches
            BodyEdgeDelta = [Math]::Round([double]$bodyEdgeDelta, 6)
            WidthDelta = [Math]::Round([double]$widthDelta, 6)
            MaxHorizontalDelta = [Math]::Round([double]$bodyEdgeDelta, 6)
        }
    }
}

function New-MarkupLaneHorizontalDeltaSummary($ReferenceLaneRects, $CandidateLaneRects, $ComparisonRows) {
    $rows = @($ComparisonRows)
    $comparedPageCount = @($rows | Where-Object { $_.Reference -ne $null -and $_.Candidate -ne $null }).Count
    $sideMismatchCount = @($rows | Where-Object { $_.Reference -ne $null -and $_.Candidate -ne $null -and $_.SideMatches -ne $true }).Count
    [pscustomobject]@{
        Name = "markup-lane-horizontal"
        ReferenceCount = @($ReferenceLaneRects).Count
        CandidateCount = @($CandidateLaneRects).Count
        ComparedPageCount = $comparedPageCount
        MissingReferenceCount = @($rows | Where-Object { $_.Status -eq "missing-reference" }).Count
        MissingCandidateCount = @($rows | Where-Object { $_.Status -eq "missing-candidate" }).Count
        SideMismatchCount = if ($comparedPageCount -eq 0) { $null } else { $sideMismatchCount }
        DeltaCount = @($rows | Where-Object { $_.Status -eq "delta" }).Count
        MaxBodyEdgeDelta = Get-MaxOrZero ($rows | ForEach-Object { $_.BodyEdgeDelta })
        MaxOccupiedWidthDelta = Get-MaxOrZero ($rows | ForEach-Object { $_.WidthDelta })
        MaxHorizontalDelta = Get-MaxOrZero ($rows | ForEach-Object { $_.MaxHorizontalDelta })
    }
}

function Compare-BodyFrameLaneEdgeRects($ReferenceLaneRects, $CandidateBodyFrameRects, $CandidateLaneRects, [double] $Tolerance) {
    $referenceByPage = @{}
    foreach ($reference in @($ReferenceLaneRects)) {
        if ($reference.Page -ne $null -and -not $referenceByPage.ContainsKey([int]$reference.Page)) {
            $referenceByPage[[int]$reference.Page] = $reference
        }
    }

    $bodyByPage = @{}
    foreach ($body in @($CandidateBodyFrameRects)) {
        if ($body.Page -ne $null -and -not $bodyByPage.ContainsKey([int]$body.Page)) {
            $bodyByPage[[int]$body.Page] = $body
        }
    }

    $laneByPage = @{}
    foreach ($lane in @($CandidateLaneRects)) {
        if ($lane.Page -ne $null -and -not $laneByPage.ContainsKey([int]$lane.Page)) {
            $laneByPage[[int]$lane.Page] = $lane
        }
    }

    $pages = New-Object System.Collections.Generic.HashSet[int]
    foreach ($page in $referenceByPage.Keys) { [void]$pages.Add([int]$page) }
    foreach ($page in $bodyByPage.Keys) { [void]$pages.Add([int]$page) }

    foreach ($page in @($pages | Sort-Object)) {
        $reference = if ($referenceByPage.ContainsKey($page)) { $referenceByPage[$page] } else { $null }
        $body = if ($bodyByPage.ContainsKey($page)) { $bodyByPage[$page] } else { $null }
        $lane = if ($laneByPage.ContainsKey($page)) { $laneByPage[$page] } else { $null }
        if ($null -eq $reference) {
            [pscustomobject]@{
                Page = $page
                Status = "missing-reference"
                Reference = $null
                CandidateBodyFrame = $body
                CandidateMarkupLane = $lane
                SideMatches = $null
                BodyEdgeDelta = $null
                CandidateLaneGap = $null
            }
            continue
        }

        if ($null -eq $body -or $null -eq $lane) {
            [pscustomobject]@{
                Page = $page
                Status = "missing-candidate"
                Reference = $reference
                CandidateBodyFrame = $body
                CandidateMarkupLane = $lane
                SideMatches = $null
                BodyEdgeDelta = $null
                CandidateLaneGap = $null
            }
            continue
        }

        $sideMatches = [string]$reference.Side -eq [string]$lane.Side
        if ($sideMatches -and [string]$reference.Side -eq "Right") {
            $gap = [double]$lane.MinX - [double]$body.MaxX
            $referenceBodyEdge = [double]$reference.MinX - $gap
            $candidateBodyEdge = [double]$body.MaxX
        }
        elseif ($sideMatches -and [string]$reference.Side -eq "Left") {
            $gap = [double]$body.MinX - [double]$lane.MaxX
            $referenceBodyEdge = [double]$reference.MaxX + $gap
            $candidateBodyEdge = [double]$body.MinX
        }
        else {
            $gap = $null
            $referenceBodyEdge = if ([Math]::Abs([double]$reference.MinX - [double]$body.MinX) -le [Math]::Abs([double]$reference.MaxX - [double]$body.MaxX)) {
                [double]$reference.MinX
            }
            else {
                [double]$reference.MaxX
            }
            $candidateBodyEdge = if ([Math]::Abs($referenceBodyEdge - [double]$body.MinX) -le [Math]::Abs($referenceBodyEdge - [double]$body.MaxX)) {
                [double]$body.MinX
            }
            else {
                [double]$body.MaxX
            }
        }

        $bodyEdgeDelta = [Math]::Abs($referenceBodyEdge - $candidateBodyEdge)
        [pscustomobject]@{
            Page = $page
            Status = if ($sideMatches -and $bodyEdgeDelta -le $Tolerance) { "ok" } else { "delta" }
            Reference = $reference
            CandidateBodyFrame = $body
            CandidateMarkupLane = $lane
            SideMatches = $sideMatches
            BodyEdgeDelta = [Math]::Round([double]$bodyEdgeDelta, 6)
            CandidateLaneGap = if ($null -eq $gap) { $null } else { [Math]::Round([double]$gap, 6) }
            ReferenceBodyEdge = [Math]::Round([double]$referenceBodyEdge, 6)
            CandidateBodyEdge = [Math]::Round([double]$candidateBodyEdge, 6)
        }
    }
}

function New-BodyFrameLaneEdgeDeltaSummary($ReferenceLaneRects, $CandidateBodyFrameRects, $CandidateLaneRects, $ComparisonRows) {
    $rows = @($ComparisonRows)
    $comparedPageCount = @($rows | Where-Object { $_.Reference -ne $null -and $_.CandidateBodyFrame -ne $null -and $_.CandidateMarkupLane -ne $null }).Count
    $sideMismatchCount = @($rows | Where-Object { $_.Reference -ne $null -and $_.CandidateBodyFrame -ne $null -and $_.CandidateMarkupLane -ne $null -and $_.SideMatches -ne $true }).Count
    [pscustomobject]@{
        Name = "body-frame-lane-edge"
        ReferenceLaneCount = @($ReferenceLaneRects).Count
        CandidateBodyFrameCount = @($CandidateBodyFrameRects).Count
        CandidateMarkupLaneCount = @($CandidateLaneRects).Count
        ComparedPageCount = $comparedPageCount
        MissingReferenceCount = @($rows | Where-Object { $_.Status -eq "missing-reference" }).Count
        MissingCandidateCount = @($rows | Where-Object { $_.Status -eq "missing-candidate" }).Count
        SideMismatchCount = if ($comparedPageCount -eq 0) { $null } else { $sideMismatchCount }
        DeltaCount = @($rows | Where-Object { $_.Status -eq "delta" }).Count
        MaxBodyEdgeDelta = Get-MaxOrZero ($rows | ForEach-Object { $_.BodyEdgeDelta })
        MaxCandidateLaneGap = Get-MaxOrZero ($rows | ForEach-Object { if ($_.CandidateLaneGap -ne $null) { [Math]::Abs([double]$_.CandidateLaneGap) } })
    }
}

function Get-PointBounds($Operations, [scriptblock] $Predicate) {
    $points = New-Object System.Collections.Generic.List[object]
    foreach ($operation in @($Operations)) {
        if (-not (& $Predicate $operation)) {
            continue
        }

        $x = Get-OperationX $operation
        $y = Get-OperationY $operation
        if ($null -ne $x -and $null -ne $y) {
            $points.Add([pscustomobject]@{ X = [double]$x; Y = [double]$y }) | Out-Null
        }
    }

    if ($points.Count -eq 0) {
        return [pscustomobject]@{
            Count = 0
            MinX = $null
            MinY = $null
            MaxX = $null
            MaxY = $null
        }
    }

    [pscustomobject]@{
        Count = $points.Count
        MinX = [Math]::Round((($points | Measure-Object -Property X -Minimum).Minimum), 6)
        MinY = [Math]::Round((($points | Measure-Object -Property Y -Minimum).Minimum), 6)
        MaxX = [Math]::Round((($points | Measure-Object -Property X -Maximum).Maximum), 6)
        MaxY = [Math]::Round((($points | Measure-Object -Property Y -Maximum).Maximum), 6)
    }
}

function New-PointBoundsDeltaSummary([string] $Name, $ReferenceOperations, $CandidateOperations, [scriptblock] $Predicate) {
    $referenceBounds = Get-PointBounds $ReferenceOperations $Predicate
    $candidateBounds = Get-PointBounds $CandidateOperations $Predicate
    $maxDelta = $null
    foreach ($property in @("MinX", "MinY", "MaxX", "MaxY")) {
        [void](Update-MaxDelta ([ref]$maxDelta) $referenceBounds.$property $candidateBounds.$property)
    }

    [pscustomobject]@{
        Name = $Name
        ReferenceBounds = $referenceBounds
        CandidateBounds = $candidateBounds
        CountDelta = [Math]::Abs([int]$candidateBounds.Count - [int]$referenceBounds.Count)
        MaxBoundsDelta = if ($null -eq $maxDelta) { $null } else { [Math]::Round([double]$maxDelta, 6) }
    }
}

function Sort-TextOperationsForGate($Operations) {
    @($Operations | Sort-Object `
            @{ Expression = { if ($_.PageNumber -ne $null) { [int]$_.PageNumber } else { 0 } } }, `
            @{ Expression = { if ($_.EffectiveY -ne $null) { -[double]$_.EffectiveY } elseif ($_.Y -ne $null) { -[double]$_.Y } else { 0d } } }, `
            @{ Expression = { if ($_.EffectiveX -ne $null) { [double]$_.EffectiveX } elseif ($_.X -ne $null) { [double]$_.X } else { 0d } } }, `
            @{ Expression = { if ($_.DecodedRuneCount -ne $null) { [int]$_.DecodedRuneCount } else { 0 } } })
}

function New-TextGateDeltaSummary($ReferenceTextOperations, $CandidateTextOperations) {
    $reference = @(Sort-TextOperationsForGate $ReferenceTextOperations)
    $candidate = @(Sort-TextOperationsForGate $CandidateTextOperations)
    $pairCount = [Math]::Min($reference.Count, $candidate.Count)
    $maxBaselineDelta = $null
    $maxXDelta = $null
    $maxFontSizeDelta = $null
    $maxCharacterSpacingDelta = $null
    $maxCharacterSpacingGapTotalDelta = $null
    $maxAdjustmentTotalDelta = $null
    $maxNetSpacingGapTotalDelta = $null
    $maxNetAverageCharacterSpacingDelta = $null
    $maxNaturalWidthDelta = $null
    $maxEmittedAdvanceDelta = $null
    $maxGlyphAdvanceDelta = $null
    $characterSpacingPairCount = 0
    $characterSpacingGapTotalPairCount = 0
    $adjustmentTotalPairCount = 0
    $netSpacingGapTotalPairCount = 0
    $netAverageCharacterSpacingPairCount = 0
    $naturalWidthPairCount = 0
    $emittedAdvancePairCount = 0
    $glyphAdvancePairCount = 0
    $fontSizeRatios = New-Object System.Collections.Generic.List[double]
    $xPairs = New-Object System.Collections.Generic.List[object]
    $baselinePairs = New-Object System.Collections.Generic.List[object]

    for ($i = 0; $i -lt $pairCount; $i++) {
        $referenceItem = $reference[$i]
        $candidateItem = $candidate[$i]
        $referenceBaseline = Get-DoubleMetric $referenceItem @("EffectiveY", "Y")
        $candidateBaseline = Get-DoubleMetric $candidateItem @("EffectiveY", "Y")
        $referenceX = Get-DoubleMetric $referenceItem @("EffectiveX", "X")
        $candidateX = Get-DoubleMetric $candidateItem @("EffectiveX", "X")
        $referenceFontSize = Get-DoubleMetric $referenceItem @("FontSize")
        $candidateFontSize = Get-DoubleMetric $candidateItem @("FontSize")
        [void](Update-MaxDelta ([ref]$maxBaselineDelta) $referenceBaseline $candidateBaseline)
        [void](Update-MaxDelta ([ref]$maxXDelta) $referenceX $candidateX)
        [void](Update-MaxDelta ([ref]$maxFontSizeDelta) $referenceFontSize $candidateFontSize)
        if ($null -ne $referenceFontSize -and $null -ne $candidateFontSize -and [double]$candidateFontSize -gt 0d) {
            $fontSizeRatios.Add([double]$referenceFontSize / [double]$candidateFontSize) | Out-Null
        }
        if ($null -ne $referenceX -and $null -ne $candidateX) {
            $xPairs.Add([pscustomobject]@{ Reference = [double]$referenceX; Candidate = [double]$candidateX }) | Out-Null
        }
        if ($null -ne $referenceBaseline -and $null -ne $candidateBaseline) {
            $baselinePairs.Add([pscustomobject]@{ Reference = [double]$referenceBaseline; Candidate = [double]$candidateBaseline }) | Out-Null
        }
        if (Update-MaxDelta ([ref]$maxCharacterSpacingDelta) (Get-DoubleMetric $referenceItem @("CharacterSpacing")) (Get-DoubleMetric $candidateItem @("CharacterSpacing"))) {
            $characterSpacingPairCount++
        }
        if (Update-MaxDelta ([ref]$maxCharacterSpacingGapTotalDelta) (Get-DoubleMetric $referenceItem @("CharacterSpacingGapTotalPoints")) (Get-DoubleMetric $candidateItem @("CharacterSpacingGapTotalPoints"))) {
            $characterSpacingGapTotalPairCount++
        }
        if (Update-MaxDelta ([ref]$maxAdjustmentTotalDelta) (Get-DoubleMetric $referenceItem @("AdjustmentTotalPoints")) (Get-DoubleMetric $candidateItem @("AdjustmentTotalPoints"))) {
            $adjustmentTotalPairCount++
        }
        if (Update-MaxDelta ([ref]$maxNetSpacingGapTotalDelta) (Get-DoubleMetric $referenceItem @("NetSpacingGapTotalPoints")) (Get-DoubleMetric $candidateItem @("NetSpacingGapTotalPoints"))) {
            $netSpacingGapTotalPairCount++
        }
        if (Update-MaxDelta ([ref]$maxNetAverageCharacterSpacingDelta) (Get-DoubleMetric $referenceItem @("NetAverageCharacterSpacing")) (Get-DoubleMetric $candidateItem @("NetAverageCharacterSpacing"))) {
            $netAverageCharacterSpacingPairCount++
        }
        if (Update-MaxDelta ([ref]$maxNaturalWidthDelta) (Get-DoubleMetric $referenceItem @("NaturalWidthPoints")) (Get-DoubleMetric $candidateItem @("NaturalWidthPoints"))) {
            $naturalWidthPairCount++
        }
        if (Update-MaxDelta ([ref]$maxEmittedAdvanceDelta) (Get-DoubleMetric $referenceItem @("EmittedAdvancePoints")) (Get-DoubleMetric $candidateItem @("EmittedAdvancePoints"))) {
            $emittedAdvancePairCount++
        }

        $beforeAdvance = $maxGlyphAdvanceDelta
        $matchedAdvance = (Update-MaxDelta ([ref]$maxGlyphAdvanceDelta) (Get-DoubleMetric $referenceItem @("EmittedAdvancePoints", "NaturalWidthPoints")) (Get-DoubleMetric $candidateItem @("EmittedAdvancePoints", "NaturalWidthPoints")))
        if ($matchedAdvance -or $beforeAdvance -ne $maxGlyphAdvanceDelta) {
            $glyphAdvancePairCount++
        }
    }

    $medianFontSizeRatio = Get-MedianOrNull $fontSizeRatios
    $fontSizeRatioDeviation = if ($null -eq $medianFontSizeRatio) {
        $null
    }
    else {
        Get-MaxAbsResidualOrNull $fontSizeRatios $medianFontSizeRatio
    }
    $xOffsetsAtMedianFontScale = if ($null -eq $medianFontSizeRatio) {
        @()
    }
    else {
        @($xPairs | ForEach-Object { [double]$_.Reference - ([double]$_.Candidate * [double]$medianFontSizeRatio) })
    }
    $baselineOffsetsAtMedianFontScale = if ($null -eq $medianFontSizeRatio) {
        @()
    }
    else {
        @($baselinePairs | ForEach-Object { [double]$_.Reference - ([double]$_.Candidate * [double]$medianFontSizeRatio) })
    }
    $medianXOffsetAtMedianFontScale = Get-MedianOrNull $xOffsetsAtMedianFontScale
    $medianBaselineOffsetAtMedianFontScale = Get-MedianOrNull $baselineOffsetsAtMedianFontScale
    $maxXResidualAtMedianFontScale = Get-MaxAbsResidualOrNull $xOffsetsAtMedianFontScale $medianXOffsetAtMedianFontScale
    $maxBaselineResidualAtMedianFontScale = Get-MaxAbsResidualOrNull $baselineOffsetsAtMedianFontScale $medianBaselineOffsetAtMedianFontScale

    [pscustomobject]@{
        ReferenceOperationCount = $reference.Count
        CandidateOperationCount = $candidate.Count
        OperationCountDelta = [Math]::Abs($candidate.Count - $reference.Count)
        OperationPairCount = $pairCount
        MaxBaselineDelta = if ($null -eq $maxBaselineDelta) { $null } else { [Math]::Round([double]$maxBaselineDelta, 6) }
        MaxXDelta = if ($null -eq $maxXDelta) { $null } else { [Math]::Round([double]$maxXDelta, 6) }
        MaxFontSizeDelta = if ($null -eq $maxFontSizeDelta) { $null } else { [Math]::Round([double]$maxFontSizeDelta, 6) }
        MaxCharacterSpacingDelta = if ($null -eq $maxCharacterSpacingDelta) { $null } else { [Math]::Round([double]$maxCharacterSpacingDelta, 6) }
        MaxCharacterSpacingGapTotalDelta = if ($null -eq $maxCharacterSpacingGapTotalDelta) { $null } else { [Math]::Round([double]$maxCharacterSpacingGapTotalDelta, 6) }
        MaxAdjustmentTotalDelta = if ($null -eq $maxAdjustmentTotalDelta) { $null } else { [Math]::Round([double]$maxAdjustmentTotalDelta, 6) }
        MaxNetSpacingGapTotalDelta = if ($null -eq $maxNetSpacingGapTotalDelta) { $null } else { [Math]::Round([double]$maxNetSpacingGapTotalDelta, 6) }
        MaxNetAverageCharacterSpacingDelta = if ($null -eq $maxNetAverageCharacterSpacingDelta) { $null } else { [Math]::Round([double]$maxNetAverageCharacterSpacingDelta, 6) }
        MaxNaturalWidthDelta = if ($null -eq $maxNaturalWidthDelta) { $null } else { [Math]::Round([double]$maxNaturalWidthDelta, 6) }
        MaxEmittedAdvanceDelta = if ($null -eq $maxEmittedAdvanceDelta) { $null } else { [Math]::Round([double]$maxEmittedAdvanceDelta, 6) }
        MaxGlyphAdvanceDelta = if ($null -eq $maxGlyphAdvanceDelta) { $null } else { [Math]::Round([double]$maxGlyphAdvanceDelta, 6) }
        FontSizeRatioPairCount = $fontSizeRatios.Count
        MedianReferenceToCandidateFontSizeRatio = if ($null -eq $medianFontSizeRatio) { $null } else { [Math]::Round([double]$medianFontSizeRatio, 6) }
        MinReferenceToCandidateFontSizeRatio = if ($fontSizeRatios.Count -eq 0) { $null } else { [Math]::Round([double](($fontSizeRatios | Measure-Object -Minimum).Minimum), 6) }
        MaxReferenceToCandidateFontSizeRatio = if ($fontSizeRatios.Count -eq 0) { $null } else { [Math]::Round([double](($fontSizeRatios | Measure-Object -Maximum).Maximum), 6) }
        MaxFontSizeRatioDeviationFromMedian = if ($null -eq $fontSizeRatioDeviation) { $null } else { [Math]::Round([double]$fontSizeRatioDeviation, 6) }
        MedianXOffsetAtMedianFontScale = if ($null -eq $medianXOffsetAtMedianFontScale) { $null } else { [Math]::Round([double]$medianXOffsetAtMedianFontScale, 6) }
        MaxXResidualAtMedianFontScale = if ($null -eq $maxXResidualAtMedianFontScale) { $null } else { [Math]::Round([double]$maxXResidualAtMedianFontScale, 6) }
        MedianBaselineOffsetAtMedianFontScale = if ($null -eq $medianBaselineOffsetAtMedianFontScale) { $null } else { [Math]::Round([double]$medianBaselineOffsetAtMedianFontScale, 6) }
        MaxBaselineResidualAtMedianFontScale = if ($null -eq $maxBaselineResidualAtMedianFontScale) { $null } else { [Math]::Round([double]$maxBaselineResidualAtMedianFontScale, 6) }
        CharacterSpacingPairCount = $characterSpacingPairCount
        CharacterSpacingGapTotalPairCount = $characterSpacingGapTotalPairCount
        AdjustmentTotalPairCount = $adjustmentTotalPairCount
        NetSpacingGapTotalPairCount = $netSpacingGapTotalPairCount
        NetAverageCharacterSpacingPairCount = $netAverageCharacterSpacingPairCount
        NaturalWidthPairCount = $naturalWidthPairCount
        EmittedAdvancePairCount = $emittedAdvancePairCount
        GlyphAdvancePairCount = $glyphAdvancePairCount
    }
}

function Get-TextOperationCharacterProfile($Operation) {
    $digitCount = 0
    $letterCount = 0
    $whitespaceCount = 0
    $punctuationCount = 0
    $symbolCount = 0
    $otherCount = 0
    $text = if ($null -ne $Operation -and
        $Operation.PSObject.Properties.Name -contains "DecodedText" -and
        $null -ne $Operation.DecodedText) {
        [string]$Operation.DecodedText
    }
    else {
        ""
    }

    foreach ($character in $text.ToCharArray()) {
        if ([char]::IsDigit($character)) {
            $digitCount++
        }
        elseif ([char]::IsLetter($character)) {
            $letterCount++
        }
        elseif ([char]::IsWhiteSpace($character)) {
            $whitespaceCount++
        }
        elseif ([char]::IsPunctuation($character)) {
            $punctuationCount++
        }
        elseif ([char]::IsSymbol($character)) {
            $symbolCount++
        }
        else {
            $otherCount++
        }
    }

    [pscustomobject]@{
        DigitCount = $digitCount
        LetterCount = $letterCount
        WhitespaceCount = $whitespaceCount
        PunctuationCount = $punctuationCount
        SymbolCount = $symbolCount
        OtherCount = $otherCount
    }
}

function New-TextOperationMetricSnapshot($Operation) {
    if ($null -eq $Operation) {
        return $null
    }

    [pscustomobject]@{
        Page = Get-OperationPage $Operation
        X = if ($null -eq (Get-DoubleMetric $Operation @("EffectiveX", "X"))) { $null } else { [Math]::Round((Get-DoubleMetric $Operation @("EffectiveX", "X")), 6) }
        Y = if ($null -eq (Get-DoubleMetric $Operation @("EffectiveY", "Y"))) { $null } else { [Math]::Round((Get-DoubleMetric $Operation @("EffectiveY", "Y")), 6) }
        FontSize = if ($null -eq (Get-DoubleMetric $Operation @("FontSize"))) { $null } else { [Math]::Round((Get-DoubleMetric $Operation @("FontSize")), 6) }
        CharacterSpacing = if ($null -eq (Get-DoubleMetric $Operation @("CharacterSpacing"))) { $null } else { [Math]::Round((Get-DoubleMetric $Operation @("CharacterSpacing")), 6) }
        DecodedRuneCount = if ($Operation.PSObject.Properties.Name -contains "DecodedRuneCount" -and $null -ne $Operation.DecodedRuneCount) { [int]$Operation.DecodedRuneCount } else { $null }
        TextChunkCount = if ($Operation.PSObject.Properties.Name -contains "TextChunkCount" -and $null -ne $Operation.TextChunkCount) { [int]$Operation.TextChunkCount } else { $null }
        AdjustmentCount = if ($Operation.PSObject.Properties.Name -contains "AdjustmentCount" -and $null -ne $Operation.AdjustmentCount) { [int]$Operation.AdjustmentCount } else { $null }
        AdjustmentTotalPoints = if ($null -eq (Get-DoubleMetric $Operation @("AdjustmentTotalPoints"))) { $null } else { [Math]::Round((Get-DoubleMetric $Operation @("AdjustmentTotalPoints")), 6) }
        NetSpacingGapTotalPoints = if ($null -eq (Get-DoubleMetric $Operation @("NetSpacingGapTotalPoints"))) { $null } else { [Math]::Round((Get-DoubleMetric $Operation @("NetSpacingGapTotalPoints")), 6) }
        NetAverageCharacterSpacing = if ($null -eq (Get-DoubleMetric $Operation @("NetAverageCharacterSpacing"))) { $null } else { [Math]::Round((Get-DoubleMetric $Operation @("NetAverageCharacterSpacing")), 6) }
        NaturalWidthPoints = if ($null -eq (Get-DoubleMetric $Operation @("NaturalWidthPoints"))) { $null } else { [Math]::Round((Get-DoubleMetric $Operation @("NaturalWidthPoints")), 6) }
        EmittedAdvancePoints = if ($null -eq (Get-DoubleMetric $Operation @("EmittedAdvancePoints"))) { $null } else { [Math]::Round((Get-DoubleMetric $Operation @("EmittedAdvancePoints")), 6) }
        Operator = if ($Operation.PSObject.Properties.Name -contains "Operator") { $Operation.Operator } else { $null }
        CharacterProfile = Get-TextOperationCharacterProfile $Operation
    }
}

function New-TextOperationDriftPair(
    [int] $Index,
    [string] $Partition,
    $ReferenceOperation,
    $CandidateOperation)
{
    $referenceX = Get-DoubleMetric $ReferenceOperation @("EffectiveX", "X")
    $candidateX = Get-DoubleMetric $CandidateOperation @("EffectiveX", "X")
    $referenceY = Get-DoubleMetric $ReferenceOperation @("EffectiveY", "Y")
    $candidateY = Get-DoubleMetric $CandidateOperation @("EffectiveY", "Y")
    $referenceFontSize = Get-DoubleMetric $ReferenceOperation @("FontSize")
    $candidateFontSize = Get-DoubleMetric $CandidateOperation @("FontSize")
    $referenceAdjustmentTotal = Get-DoubleMetric $ReferenceOperation @("AdjustmentTotalPoints")
    $candidateAdjustmentTotal = Get-DoubleMetric $CandidateOperation @("AdjustmentTotalPoints")
    $referenceNetSpacingGapTotal = Get-DoubleMetric $ReferenceOperation @("NetSpacingGapTotalPoints")
    $candidateNetSpacingGapTotal = Get-DoubleMetric $CandidateOperation @("NetSpacingGapTotalPoints")
    $referenceNetAverageCharacterSpacing = Get-DoubleMetric $ReferenceOperation @("NetAverageCharacterSpacing")
    $candidateNetAverageCharacterSpacing = Get-DoubleMetric $CandidateOperation @("NetAverageCharacterSpacing")

    [pscustomobject]@{
        Index = $Index
        Partition = $Partition
        Page = if ($null -ne $ReferenceOperation) { Get-OperationPage $ReferenceOperation } elseif ($null -ne $CandidateOperation) { Get-OperationPage $CandidateOperation } else { $null }
        XDelta = if ($null -eq $referenceX -or $null -eq $candidateX) { $null } else { [Math]::Round([double]$candidateX - [double]$referenceX, 6) }
        AbsXDelta = if ($null -eq $referenceX -or $null -eq $candidateX) { $null } else { [Math]::Round([Math]::Abs([double]$candidateX - [double]$referenceX), 6) }
        BaselineDelta = if ($null -eq $referenceY -or $null -eq $candidateY) { $null } else { [Math]::Round([double]$candidateY - [double]$referenceY, 6) }
        AbsBaselineDelta = if ($null -eq $referenceY -or $null -eq $candidateY) { $null } else { [Math]::Round([Math]::Abs([double]$candidateY - [double]$referenceY), 6) }
        FontSizeDelta = if ($null -eq $referenceFontSize -or $null -eq $candidateFontSize) { $null } else { [Math]::Round([double]$candidateFontSize - [double]$referenceFontSize, 6) }
        AdjustmentTotalDelta = if ($null -eq $referenceAdjustmentTotal -or $null -eq $candidateAdjustmentTotal) { $null } else { [Math]::Round([double]$candidateAdjustmentTotal - [double]$referenceAdjustmentTotal, 6) }
        AbsAdjustmentTotalDelta = if ($null -eq $referenceAdjustmentTotal -or $null -eq $candidateAdjustmentTotal) { $null } else { [Math]::Round([Math]::Abs([double]$candidateAdjustmentTotal - [double]$referenceAdjustmentTotal), 6) }
        NetSpacingGapTotalDelta = if ($null -eq $referenceNetSpacingGapTotal -or $null -eq $candidateNetSpacingGapTotal) { $null } else { [Math]::Round([double]$candidateNetSpacingGapTotal - [double]$referenceNetSpacingGapTotal, 6) }
        AbsNetSpacingGapTotalDelta = if ($null -eq $referenceNetSpacingGapTotal -or $null -eq $candidateNetSpacingGapTotal) { $null } else { [Math]::Round([Math]::Abs([double]$candidateNetSpacingGapTotal - [double]$referenceNetSpacingGapTotal), 6) }
        NetAverageCharacterSpacingDelta = if ($null -eq $referenceNetAverageCharacterSpacing -or $null -eq $candidateNetAverageCharacterSpacing) { $null } else { [Math]::Round([double]$candidateNetAverageCharacterSpacing - [double]$referenceNetAverageCharacterSpacing, 6) }
        AbsNetAverageCharacterSpacingDelta = if ($null -eq $referenceNetAverageCharacterSpacing -or $null -eq $candidateNetAverageCharacterSpacing) { $null } else { [Math]::Round([Math]::Abs([double]$candidateNetAverageCharacterSpacing - [double]$referenceNetAverageCharacterSpacing), 6) }
        Reference = New-TextOperationMetricSnapshot $ReferenceOperation
        Candidate = New-TextOperationMetricSnapshot $CandidateOperation
    }
}

function Select-TopTextDriftRows($Rows, [string] $Metric, [int] $Limit) {
    @($Rows |
        Where-Object { $null -ne $_.$Metric } |
        Sort-Object -Property @{ Expression = $Metric; Descending = $true }, Index |
        Select-Object -First $Limit)
}

function New-TextOperationDriftSummary(
    $ReferenceTextOperations,
    $CandidateTextOperations,
    $ReferenceMarkupLaneRects,
    $CandidateMarkupLaneRects,
    [int] $Limit = 10)
{
    $referenceRows = if (@($ReferenceMarkupLaneRects).Count -eq 0 -and @($CandidateMarkupLaneRects).Count -eq 0) {
        @(Sort-TextOperationsForGate $ReferenceTextOperations | ForEach-Object { [pscustomobject]@{ Key = "all"; Operation = $_ } })
    }
    else {
        @(New-TextOperationPartitionRows $ReferenceTextOperations $ReferenceMarkupLaneRects)
    }
    $candidateRows = if (@($ReferenceMarkupLaneRects).Count -eq 0 -and @($CandidateMarkupLaneRects).Count -eq 0) {
        @(Sort-TextOperationsForGate $CandidateTextOperations | ForEach-Object { [pscustomobject]@{ Key = "all"; Operation = $_ } })
    }
    else {
        @(New-TextOperationPartitionRows $CandidateTextOperations $CandidateMarkupLaneRects)
    }
    $keys = @(@($referenceRows | ForEach-Object { $_.Key }) + @($candidateRows | ForEach-Object { $_.Key }) | Sort-Object -Unique)
    $pairs = New-Object System.Collections.Generic.List[object]
    $globalIndex = 0
    foreach ($rawKey in $keys) {
        $key = [string]$rawKey
        $referenceItems = @(Sort-TextOperationsForGate (@($referenceRows | Where-Object { $_.Key -eq $key } | ForEach-Object { $_.Operation })))
        $candidateItems = @(Sort-TextOperationsForGate (@($candidateRows | Where-Object { $_.Key -eq $key } | ForEach-Object { $_.Operation })))
        $pairCount = [Math]::Min($referenceItems.Count, $candidateItems.Count)
        for ($i = 0; $i -lt $pairCount; $i++) {
            $pairs.Add((New-TextOperationDriftPair $globalIndex $key $referenceItems[$i] $candidateItems[$i])) | Out-Null
            $globalIndex++
        }
    }

    [pscustomobject]@{
        ReferenceOperationCount = @($ReferenceTextOperations).Count
        CandidateOperationCount = @($CandidateTextOperations).Count
        OperationCountDelta = [Math]::Abs(@($CandidateTextOperations).Count - @($ReferenceTextOperations).Count)
        OperationPairCount = $pairs.Count
        TopXDeltaPairs = Select-TopTextDriftRows $pairs "AbsXDelta" $Limit
        TopBaselineDeltaPairs = Select-TopTextDriftRows $pairs "AbsBaselineDelta" $Limit
        TopAdjustmentTotalDeltaPairs = Select-TopTextDriftRows $pairs "AbsAdjustmentTotalDelta" $Limit
        TopNetSpacingGapTotalDeltaPairs = Select-TopTextDriftRows $pairs "AbsNetSpacingGapTotalDelta" $Limit
        TopNetAverageCharacterSpacingDeltaPairs = Select-TopTextDriftRows $pairs "AbsNetAverageCharacterSpacingDelta" $Limit
    }
}

function Test-TextOperationInsideMarkupLane($Operation, $MarkupLaneRects) {
    $page = Get-OperationPage $Operation
    $x = Get-DoubleMetric $Operation @("EffectiveX", "X")
    $y = Get-DoubleMetric $Operation @("EffectiveY", "Y")
    if ($null -eq $page -or $null -eq $x -or $null -eq $y) {
        return $false
    }

    foreach ($rect in @($MarkupLaneRects | Where-Object { $_.Page -eq $page })) {
        if ($rect.MinX -eq $null -or $rect.MaxX -eq $null -or $rect.MinY -eq $null -or $rect.MaxY -eq $null) {
            continue
        }

        if ([double]$x -ge [double]$rect.MinX - 2d -and
            [double]$x -le [double]$rect.MaxX + 2d -and
            [double]$y -ge [double]$rect.MinY - 2d -and
            [double]$y -le [double]$rect.MaxY + 2d) {
            return $true
        }
    }

    return $false
}

function Get-TextGatePartitionKey($Operation, $MarkupLaneRects) {
    $page = Get-OperationPage $Operation
    if ($null -eq $page) {
        $page = 0
    }

    if (Test-TextOperationInsideMarkupLane $Operation $MarkupLaneRects) {
        return "markup-lane:$page"
    }

    return "body:$page"
}

function New-TextOperationPartitionRows($Operations, $MarkupLaneRects) {
    foreach ($operation in @($Operations)) {
        [pscustomobject]@{
            Key = Get-TextGatePartitionKey $operation $MarkupLaneRects
            Operation = $operation
        }
    }
}

function Get-SummaryMaxMetric($Summaries, [string] $Name) {
    $values = @($Summaries | ForEach-Object { $_.$Name } | Where-Object { $null -ne $_ })
    if ($values.Count -eq 0) {
        return $null
    }

    return [Math]::Round([double](($values | Measure-Object -Maximum).Maximum), 6)
}

function Get-SummaryMinMetric($Summaries, [string] $Name) {
    $values = @($Summaries | ForEach-Object { $_.$Name } | Where-Object { $null -ne $_ })
    if ($values.Count -eq 0) {
        return $null
    }

    return [Math]::Round([double](($values | Measure-Object -Minimum).Minimum), 6)
}

function Get-SummarySumMetric($Summaries, [string] $Name) {
    $values = @($Summaries | ForEach-Object { $_.$Name } | Where-Object { $null -ne $_ })
    if ($values.Count -eq 0) {
        return 0
    }

    return [int](($values | Measure-Object -Sum).Sum)
}

function New-PartitionedTextGateDeltaSummary(
    $ReferenceTextOperations,
    $CandidateTextOperations,
    $ReferenceMarkupLaneRects,
    $CandidateMarkupLaneRects)
{
    if (@($ReferenceMarkupLaneRects).Count -eq 0 -and @($CandidateMarkupLaneRects).Count -eq 0) {
        return New-TextGateDeltaSummary $ReferenceTextOperations $CandidateTextOperations
    }

    $referenceRows = @(New-TextOperationPartitionRows $ReferenceTextOperations $ReferenceMarkupLaneRects)
    $candidateRows = @(New-TextOperationPartitionRows $CandidateTextOperations $CandidateMarkupLaneRects)
    $keys = @(@($referenceRows | ForEach-Object { $_.Key }) + @($candidateRows | ForEach-Object { $_.Key }) | Sort-Object -Unique)
    $partitionSummaries = @(
        foreach ($rawKey in $keys) {
            $key = [string]$rawKey
            $referenceItems = @($referenceRows | Where-Object { $_.Key -eq $key } | ForEach-Object { $_.Operation })
            $candidateItems = @($candidateRows | Where-Object { $_.Key -eq $key } | ForEach-Object { $_.Operation })
            $summary = New-TextGateDeltaSummary $referenceItems $candidateItems
            $summary | Add-Member -NotePropertyName Partition -NotePropertyValue $key -Force
            $summary
        }
    )

    [pscustomobject]@{
        ReferenceOperationCount = @($ReferenceTextOperations).Count
        CandidateOperationCount = @($CandidateTextOperations).Count
        OperationCountDelta = [Math]::Abs(@($CandidateTextOperations).Count - @($ReferenceTextOperations).Count)
        OperationPairCount = Get-SummarySumMetric $partitionSummaries "OperationPairCount"
        MaxBaselineDelta = Get-SummaryMaxMetric $partitionSummaries "MaxBaselineDelta"
        MaxXDelta = Get-SummaryMaxMetric $partitionSummaries "MaxXDelta"
        MaxFontSizeDelta = Get-SummaryMaxMetric $partitionSummaries "MaxFontSizeDelta"
        MaxCharacterSpacingDelta = Get-SummaryMaxMetric $partitionSummaries "MaxCharacterSpacingDelta"
        MaxCharacterSpacingGapTotalDelta = Get-SummaryMaxMetric $partitionSummaries "MaxCharacterSpacingGapTotalDelta"
        MaxAdjustmentTotalDelta = Get-SummaryMaxMetric $partitionSummaries "MaxAdjustmentTotalDelta"
        MaxNetSpacingGapTotalDelta = Get-SummaryMaxMetric $partitionSummaries "MaxNetSpacingGapTotalDelta"
        MaxNetAverageCharacterSpacingDelta = Get-SummaryMaxMetric $partitionSummaries "MaxNetAverageCharacterSpacingDelta"
        MaxNaturalWidthDelta = Get-SummaryMaxMetric $partitionSummaries "MaxNaturalWidthDelta"
        MaxEmittedAdvanceDelta = Get-SummaryMaxMetric $partitionSummaries "MaxEmittedAdvanceDelta"
        MaxGlyphAdvanceDelta = Get-SummaryMaxMetric $partitionSummaries "MaxGlyphAdvanceDelta"
        FontSizeRatioPairCount = Get-SummarySumMetric $partitionSummaries "FontSizeRatioPairCount"
        MedianReferenceToCandidateFontSizeRatio = $null
        MinReferenceToCandidateFontSizeRatio = Get-SummaryMinMetric $partitionSummaries "MinReferenceToCandidateFontSizeRatio"
        MaxReferenceToCandidateFontSizeRatio = Get-SummaryMaxMetric $partitionSummaries "MaxReferenceToCandidateFontSizeRatio"
        MaxFontSizeRatioDeviationFromMedian = Get-SummaryMaxMetric $partitionSummaries "MaxFontSizeRatioDeviationFromMedian"
        MedianXOffsetAtMedianFontScale = $null
        MaxXResidualAtMedianFontScale = Get-SummaryMaxMetric $partitionSummaries "MaxXResidualAtMedianFontScale"
        MedianBaselineOffsetAtMedianFontScale = $null
        MaxBaselineResidualAtMedianFontScale = Get-SummaryMaxMetric $partitionSummaries "MaxBaselineResidualAtMedianFontScale"
        CharacterSpacingPairCount = Get-SummarySumMetric $partitionSummaries "CharacterSpacingPairCount"
        CharacterSpacingGapTotalPairCount = Get-SummarySumMetric $partitionSummaries "CharacterSpacingGapTotalPairCount"
        AdjustmentTotalPairCount = Get-SummarySumMetric $partitionSummaries "AdjustmentTotalPairCount"
        NetSpacingGapTotalPairCount = Get-SummarySumMetric $partitionSummaries "NetSpacingGapTotalPairCount"
        NetAverageCharacterSpacingPairCount = Get-SummarySumMetric $partitionSummaries "NetAverageCharacterSpacingPairCount"
        NaturalWidthPairCount = Get-SummarySumMetric $partitionSummaries "NaturalWidthPairCount"
        EmittedAdvancePairCount = Get-SummarySumMetric $partitionSummaries "EmittedAdvancePairCount"
        GlyphAdvancePairCount = Get-SummarySumMetric $partitionSummaries "GlyphAdvancePairCount"
        PartitionCount = $partitionSummaries.Count
        Partitions = $partitionSummaries
    }
}

function New-MediaBoxDeltaSummary($ReferenceMediaBoxes, $CandidateMediaBoxes) {
    $reference = @($ReferenceMediaBoxes)
    $candidate = @($CandidateMediaBoxes)
    $pairCount = [Math]::Min($reference.Count, $candidate.Count)
    $maxDelta = $null
    for ($i = 0; $i -lt $pairCount; $i++) {
        foreach ($property in @("MinX", "MinY", "MaxX", "MaxY", "Width", "Height")) {
            [void](Update-MaxDelta ([ref]$maxDelta) $reference[$i].$property $candidate[$i].$property)
        }
    }

    [pscustomobject]@{
        ReferenceCount = $reference.Count
        CandidateCount = $candidate.Count
        CountDelta = [Math]::Abs($candidate.Count - $reference.Count)
        PairCount = $pairCount
        MaxDelta = if ($null -eq $maxDelta) { $null } else { [Math]::Round([double]$maxDelta, 6) }
    }
}

function Get-TextBaselinePageSummary($TextOperations, [int] $PageNumber) {
    $pageOperations = @($TextOperations | Where-Object { $_.PageNumber -ne $null -and [int]$_.PageNumber -eq $PageNumber })
    $baselines = @($pageOperations |
        ForEach-Object { Get-DoubleMetric $_ @("EffectiveY", "Y") } |
        Where-Object { $null -ne $_ } |
        ForEach-Object { [double]$_ })
    if ($baselines.Count -eq 0) {
        return [pscustomobject]@{
            TextOperationCount = $pageOperations.Count
            FirstBaselineY = $null
            LastBaselineY = $null
            BodyFrameHeightUsedPoints = $null
        }
    }

    $first = [double](($baselines | Measure-Object -Maximum).Maximum)
    $last = [double](($baselines | Measure-Object -Minimum).Minimum)
    [pscustomobject]@{
        TextOperationCount = $pageOperations.Count
        FirstBaselineY = [Math]::Round($first, 6)
        LastBaselineY = [Math]::Round($last, 6)
        BodyFrameHeightUsedPoints = [Math]::Round([Math]::Abs($first - $last), 6)
    }
}

function Get-LayoutPageByPageNumber($LayoutPages, [int] $PageNumber) {
    if ($LayoutPages.Count -eq 0) {
        return $null
    }

    foreach ($page in $LayoutPages) {
        if ($page.PSObject.Properties.Name -contains "PageNumber" -and $null -ne $page.PageNumber -and [int]$page.PageNumber -eq $PageNumber) {
            return $page
        }
    }

    if ($PageNumber -ge 1 -and $PageNumber -le $LayoutPages.Count) {
        return $LayoutPages[$PageNumber - 1]
    }

    return $null
}

function Get-TextEmissionPageByPageNumber($TextEmissionPages, [int] $PageNumber) {
    foreach ($page in $TextEmissionPages) {
        if ($page.PSObject.Properties.Name -contains "PageIndex" -and $null -ne $page.PageIndex -and ([int]$page.PageIndex + 1) -eq $PageNumber) {
            return $page
        }
    }

    return $null
}

function Get-CarriedSourceBlockCount($SourceBlocks, [int] $PageNumber) {
    $pageIndex = $PageNumber - 1
    return @($SourceBlocks | Where-Object {
            $_.FirstPageIndex -ne $null -and
                $_.LastPageIndex -ne $null -and
                [int]$_.FirstPageIndex -le $pageIndex -and
                [int]$_.LastPageIndex -ge $pageIndex -and
                [int]$_.FirstPageIndex -ne [int]$_.LastPageIndex
        }).Count
}

function New-PageFlowDeltaSummary(
    $ReferenceTextOperations,
    $CandidateTextOperations,
    $CandidateLayoutSnapshot,
    $CandidateTextEmissionSummary,
    $CandidateSourceBlocks,
    [int] $ReferencePageCount,
    [int] $CandidatePageCount)
{
    $layoutPages = if ($null -ne $CandidateLayoutSnapshot -and $CandidateLayoutSnapshot.PSObject.Properties.Name -contains "Pages") {
        @($CandidateLayoutSnapshot.Pages | Where-Object { $null -ne $_ })
    }
    else {
        @()
    }
    $textEmissionPages = if ($null -ne $CandidateTextEmissionSummary -and $CandidateTextEmissionSummary.PSObject.Properties.Name -contains "LinesByPage") {
        @($CandidateTextEmissionSummary.LinesByPage | Where-Object { $null -ne $_ })
    }
    else {
        @()
    }
    $sourceBlocks = @($CandidateSourceBlocks | Where-Object { $null -ne $_ })
    $pageCount = [Math]::Max([Math]::Max($ReferencePageCount, $CandidatePageCount), $layoutPages.Count)

    $pages = @(
        for ($pageNumber = 1; $pageNumber -le $pageCount; $pageNumber++) {
            $referenceBaseline = Get-TextBaselinePageSummary $ReferenceTextOperations $pageNumber
            $candidateBaseline = Get-TextBaselinePageSummary $CandidateTextOperations $pageNumber
            $layoutPage = Get-LayoutPageByPageNumber $layoutPages $pageNumber
            $emissionPage = Get-TextEmissionPageByPageNumber $textEmissionPages $pageNumber
            [pscustomobject]@{
                Page = $pageNumber
                ReferenceTextOperationCount = $referenceBaseline.TextOperationCount
                CandidateTextOperationCount = $candidateBaseline.TextOperationCount
                TextOperationCountDelta = [Math]::Abs([int]$candidateBaseline.TextOperationCount - [int]$referenceBaseline.TextOperationCount)
                ReferenceFirstBaselineY = $referenceBaseline.FirstBaselineY
                CandidateFirstBaselineY = $candidateBaseline.FirstBaselineY
                FirstBaselineDelta = if ($null -eq $referenceBaseline.FirstBaselineY -or $null -eq $candidateBaseline.FirstBaselineY) { $null } else { [Math]::Round([Math]::Abs([double]$candidateBaseline.FirstBaselineY - [double]$referenceBaseline.FirstBaselineY), 6) }
                ReferenceLastBaselineY = $referenceBaseline.LastBaselineY
                CandidateLastBaselineY = $candidateBaseline.LastBaselineY
                LastBaselineDelta = if ($null -eq $referenceBaseline.LastBaselineY -or $null -eq $candidateBaseline.LastBaselineY) { $null } else { [Math]::Round([Math]::Abs([double]$candidateBaseline.LastBaselineY - [double]$referenceBaseline.LastBaselineY), 6) }
                ReferenceBodyFrameHeightUsedPoints = $referenceBaseline.BodyFrameHeightUsedPoints
                CandidateBodyFrameHeightUsedPoints = $candidateBaseline.BodyFrameHeightUsedPoints
                BodyFrameHeightUsedDelta = if ($null -eq $referenceBaseline.BodyFrameHeightUsedPoints -or $null -eq $candidateBaseline.BodyFrameHeightUsedPoints) { $null } else { [Math]::Round([Math]::Abs([double]$candidateBaseline.BodyFrameHeightUsedPoints - [double]$referenceBaseline.BodyFrameHeightUsedPoints), 6) }
                CandidateSourceBlockCount = if ($null -eq $layoutPage -or $layoutPage.SourceBlockCount -eq $null) { $null } else { [int]$layoutPage.SourceBlockCount }
                CandidateFirstSourceBlockIndex = if ($null -eq $layoutPage -or $layoutPage.FirstSourceBlockIndex -eq $null) { $null } else { [int]$layoutPage.FirstSourceBlockIndex }
                CandidateLastSourceBlockIndex = if ($null -eq $layoutPage -or $layoutPage.LastSourceBlockIndex -eq $null) { $null } else { [int]$layoutPage.LastSourceBlockIndex }
                CandidateEmissionLineCount = if ($null -eq $emissionPage -or $emissionPage.LineCount -eq $null) { $null } else { [int]$emissionPage.LineCount }
                CandidateEmissionFirstSourceBlockIndex = if ($null -eq $emissionPage -or $emissionPage.FirstSourceBlockIndex -eq $null) { $null } else { [int]$emissionPage.FirstSourceBlockIndex }
                CandidateEmissionLastSourceBlockIndex = if ($null -eq $emissionPage -or $emissionPage.LastSourceBlockIndex -eq $null) { $null } else { [int]$emissionPage.LastSourceBlockIndex }
                CandidateCarriedSourceBlockCount = Get-CarriedSourceBlockCount $sourceBlocks $pageNumber
            }
        }
    )

    [pscustomobject]@{
        ReferencePageCount = $ReferencePageCount
        CandidatePageCount = $CandidatePageCount
        CandidateLayoutPageCount = $layoutPages.Count
        MaxFirstBaselineDelta = Get-MaxOrZero ($pages | ForEach-Object { $_.FirstBaselineDelta })
        MaxLastBaselineDelta = Get-MaxOrZero ($pages | ForEach-Object { $_.LastBaselineDelta })
        MaxBodyFrameHeightUsedDelta = Get-MaxOrZero ($pages | ForEach-Object { $_.BodyFrameHeightUsedDelta })
        CandidateCarriedSourceBlockPageCount = @($pages | Where-Object { $_.CandidateCarriedSourceBlockCount -gt 0 }).Count
        Pages = $pages
    }
}

function Get-MaxVisualMetric($VisualMetrics, [string] $Property) {
    $values = @($VisualMetrics |
        Where-Object { $_.PSObject.Properties.Name -contains $Property -and $null -ne $_.$Property } |
        ForEach-Object { [double]$_.$Property })
    if ($values.Count -eq 0) {
        return $null
    }

    return [Math]::Round([double](($values | Measure-Object -Maximum).Maximum), 6)
}

function New-GateResult([string] $Name, [string] $Category, $Actual, $Limit, [string] $Unit) {
    if ($null -eq $Actual -or $null -eq $Limit) {
        return [pscustomobject]@{
            Name = $Name
            Category = $Category
            Status = "skipped"
            Passed = $null
            Actual = $Actual
            Limit = $Limit
            Unit = $Unit
        }
    }

    $passed = [double]$Actual -le [double]$Limit
    [pscustomobject]@{
        Name = $Name
        Category = $Category
        Status = if ($passed) { "passed" } else { "failed" }
        Passed = $passed
        Actual = [Math]::Round([double]$Actual, 6)
        Limit = [Math]::Round([double]$Limit, 6)
        Unit = $Unit
    }
}

function Get-GatePhase($Gate) {
    switch -Regex ($Gate.Name) {
        '^(page-count-delta|media-box-count-delta|media-box-max-delta)$' {
            return [pscustomobject]@{ Order = 1; Id = "page-media-box-parity"; Name = "Page count and media box parity" }
        }
        '^(body-frame-lane-edge-max-delta|body-frame-lane-side-mismatch-count|markup-lane-side-mismatch-count|markup-lane-body-edge-max-delta|page-flow-first-baseline-max-delta|page-flow-last-baseline-max-delta|page-flow-body-height-used-max-delta|table-grid-delta-count|table-grid-max-bounds-delta|connector-lane-endpoint-delta-count|connector-lane-endpoint-max-delta|connector-side-inconsistent-count|annotation-delta-count|annotation-target-delta-count|annotation-max-bounds-delta|balloon-geometry-delta-count|balloon-max-bounds-delta)$' {
            return [pscustomobject]@{ Order = 2; Id = "structural-region-parity"; Name = "Structural region parity" }
        }
        '^(text-|graphics-(visible-)?operation-comparison-exit-code)' {
            return [pscustomobject]@{ Order = 3; Id = "text-graphics-operation-parity"; Name = "Text and graphics operation parity" }
        }
        '^raster-' {
            return [pscustomobject]@{ Order = 4; Id = "raster-parity"; Name = "Raster parity" }
        }
        default {
            return [pscustomobject]@{ Order = 2; Id = "structural-region-parity"; Name = "Structural region parity" }
        }
    }
}

function New-GateSummary($Gates) {
    $gateItems = @($Gates | ForEach-Object {
        $phase = Get-GatePhase $_
        $_ | Add-Member -NotePropertyName PhaseOrder -NotePropertyValue $phase.Order -Force -PassThru |
            Add-Member -NotePropertyName PhaseId -NotePropertyValue $phase.Id -Force -PassThru |
            Add-Member -NotePropertyName PhaseName -NotePropertyValue $phase.Name -Force -PassThru
    })
    $failures = @($gateItems | Where-Object { $_.Status -eq "failed" })
    $phases = @($gateItems |
        Group-Object PhaseOrder, PhaseId, PhaseName |
        ForEach-Object {
            $parts = $_.Name -split ', '
            $phaseGates = @($_.Group)
            $phaseFailures = @($phaseGates | Where-Object { $_.Status -eq "failed" })
            $phaseSkipped = @($phaseGates | Where-Object { $_.Status -eq "skipped" })
            [pscustomobject]@{
                Order = [int]$parts[0]
                Id = $parts[1]
                Name = $parts[2]
                Status = if ($phaseFailures.Count -ne 0) { "failed" } elseif ($phaseSkipped.Count -eq $phaseGates.Count) { "skipped" } else { "passed" }
                GateCount = $phaseGates.Count
                FailureCount = $phaseFailures.Count
                SkippedCount = $phaseSkipped.Count
                Failures = $phaseFailures
            }
        } |
        Sort-Object Order)
    $blockingPhase = $phases | Where-Object { $_.Status -eq "failed" } | Select-Object -First 1
    [pscustomobject]@{
        FailureCount = $failures.Count
        Failures = $failures
        CurrentBlockingPhase = $blockingPhase
        Phases = $phases
        Gates = $gateItems
    }
}

function New-RasterRegionSpec([string] $Region, [int] $Page, [double] $XRatio, [double] $YRatio, [double] $WidthRatio, [double] $HeightRatio) {
    [pscustomobject]@{
        page = $Page
        region = $Region
        xRatio = [Math]::Round((Clamp-Ratio $XRatio), 6)
        yRatio = [Math]::Round((Clamp-Ratio $YRatio), 6)
        widthRatio = [Math]::Round((Clamp-Ratio $WidthRatio), 6)
        heightRatio = [Math]::Round((Clamp-Ratio $HeightRatio), 6)
    }
}

function Clamp-Ratio([double] $Value) {
    return [Math]::Min(1d, [Math]::Max(0d, $Value))
}

function Get-MediaBoxForPage($MediaBoxes, [int] $Page) {
    $boxes = @($MediaBoxes)
    if ($Page -ge 1 -and $Page -le $boxes.Count) {
        return $boxes[$Page - 1]
    }

    if ($boxes.Count -ne 0) {
        return $boxes[0]
    }

    return [pscustomobject]@{
        MinX = 0d
        MinY = 0d
        MaxX = 612d
        MaxY = 792d
        Width = 612d
        Height = 792d
    }
}

function New-PdfBoundsRasterRegionSpec([string] $Region, [int] $Page, $MediaBoxes, [double] $MinX, [double] $MinY, [double] $MaxX, [double] $MaxY, [double] $PaddingPoints = 4d) {
    $box = Get-MediaBoxForPage $MediaBoxes $Page
    $pageWidth = [Math]::Max(1d, [double]$box.Width)
    $pageHeight = [Math]::Max(1d, [double]$box.Height)
    $left = [Math]::Max(0d, $MinX - $PaddingPoints)
    $bottom = [Math]::Max(0d, $MinY - $PaddingPoints)
    $right = [Math]::Min($pageWidth, $MaxX + $PaddingPoints)
    $top = [Math]::Min($pageHeight, $MaxY + $PaddingPoints)
    if ($right -le $left -or $top -le $bottom) {
        return $null
    }

    return New-RasterRegionSpec `
        -Region $Region `
        -Page $Page `
        -XRatio ($left / $pageWidth) `
        -YRatio (($pageHeight - $top) / $pageHeight) `
        -WidthRatio (($right - $left) / $pageWidth) `
        -HeightRatio (($top - $bottom) / $pageHeight)
}

function New-OperationBoundsRasterRegionSpec(
    [string] $Region,
    [int] $Page,
    $ReferenceOperations,
    $CandidateOperations,
    [scriptblock] $Predicate,
    $MediaBoxes)
{
    $items = @($ReferenceOperations + $CandidateOperations | Where-Object {
            (Get-OperationPage $_) -eq $Page -and (& $Predicate $_) -and
            $_.MinX -ne $null -and $_.MinY -ne $null -and $_.MaxX -ne $null -and $_.MaxY -ne $null
        })
    if ($items.Count -eq 0) {
        return $null
    }

    return New-PdfBoundsRasterRegionSpec `
        -Region $Region `
        -Page $Page `
        -MediaBoxes $MediaBoxes `
        -MinX (($items | Measure-Object -Property MinX -Minimum).Minimum) `
        -MinY (($items | Measure-Object -Property MinY -Minimum).Minimum) `
        -MaxX (($items | Measure-Object -Property MaxX -Maximum).Maximum) `
        -MaxY (($items | Measure-Object -Property MaxY -Maximum).Maximum)
}

function New-BalloonBoundsRasterRegionSpec([string] $Region, [int] $Page, $Balloons, $MediaBoxes) {
    $items = @($Balloons | Where-Object { $_.Page -eq $Page -and $_.MinX -ne $null -and $_.MinY -ne $null -and $_.MaxX -ne $null -and $_.MaxY -ne $null })
    if ($items.Count -eq 0) {
        return $null
    }

    return New-PdfBoundsRasterRegionSpec `
        -Region $Region `
        -Page $Page `
        -MediaBoxes $MediaBoxes `
        -MinX (($items | Measure-Object -Property MinX -Minimum).Minimum) `
        -MinY (($items | Measure-Object -Property MinY -Minimum).Minimum) `
        -MaxX (($items | Measure-Object -Property MaxX -Maximum).Maximum) `
        -MaxY (($items | Measure-Object -Property MaxY -Maximum).Maximum)
}

function New-RasterRegionSpecs(
    [int] $ReferencePageCount,
    [int] $CandidatePageCount,
    $MediaBoxes,
    $ReferenceGraphicsOperations,
    $CandidateGraphicsOperations,
    $CandidateBalloons)
{
    $pageCount = [Math]::Max($ReferencePageCount, $CandidatePageCount)
    for ($page = 1; $page -le $pageCount; $page++) {
        New-RasterRegionSpec "body-text" $page 0.16 0.12 0.58 0.76
        New-RasterRegionSpec "headers-footers" $page 0.00 0.00 1.00 0.12
        New-RasterRegionSpec "headers-footers" $page 0.00 0.88 1.00 0.12
        New-RasterRegionSpec "footnotes-endnotes" $page 0.12 0.72 0.64 0.18
        New-RasterRegionSpec "left-markup-lane" $page 0.00 0.10 0.18 0.80
        New-RasterRegionSpec "right-markup-lane" $page 0.72 0.10 0.28 0.80

        $tableSpec = New-OperationBoundsRasterRegionSpec "tables" $page $ReferenceGraphicsOperations $CandidateGraphicsOperations { param($op) Test-TableLikeGraphic $op } $MediaBoxes
        if ($null -ne $tableSpec) { $tableSpec }
        $drawingSpec = New-OperationBoundsRasterRegionSpec "drawings" $page $ReferenceGraphicsOperations $CandidateGraphicsOperations { param($op) Test-DrawingLikeGraphic $op } $MediaBoxes
        if ($null -ne $drawingSpec) { $drawingSpec }
        $connectorSpec = New-OperationBoundsRasterRegionSpec "connectors" $page $ReferenceGraphicsOperations $CandidateGraphicsOperations { param($op) Test-ConnectorLikeGraphic $op } $MediaBoxes
        if ($null -ne $connectorSpec) { $connectorSpec }
        $commentSpec = New-BalloonBoundsRasterRegionSpec "comment-balloons" $page (@($CandidateBalloons | Where-Object { $_.Subtype -eq "Comment" })) $MediaBoxes
        if ($null -ne $commentSpec) { $commentSpec }
        $revisionSpec = New-BalloonBoundsRasterRegionSpec "revision-balloons" $page (@($CandidateBalloons | Where-Object { $_.Subtype -eq "Revision" })) $MediaBoxes
        if ($null -ne $revisionSpec) { $revisionSpec }
    }
}

function New-RasterRegionMetricSummary($Metrics) {
    foreach ($group in @($Metrics | Group-Object -Property Region | Sort-Object Name)) {
        $items = @($group.Group)
        [pscustomobject]@{
            Region = $group.Name
            RegionRectCount = $items.Count
            PixelCount = ($items | Measure-Object -Property PixelCount -Sum).Sum
            DimensionMismatchCount = @($items | Where-Object { $_.DimensionsMatch -ne $true }).Count
            MaxMeanAbsoluteError = Get-MaxOrZero ($items | ForEach-Object { $_.MeanAbsoluteError })
            MaxChangedPixelRatioAtThreshold16 = Get-MaxOrZero ($items | ForEach-Object { $_.ChangedPixelRatioAtThreshold16 })
            MinStructuralSimilarity = if (@($items | Where-Object { $_.StructuralSimilarity -ne $null }).Count -eq 0) { $null } else { ($items | Measure-Object -Property StructuralSimilarity -Minimum).Minimum }
        }
    }
}

function Get-RasterRegionDeltaCount($RegionSummaries, [double] $MaxMeanAbsoluteErrorLimit, [double] $MaxChangedPixelRatioLimit) {
    @($RegionSummaries | Where-Object {
            ($_.DimensionMismatchCount -ne $null -and [int]$_.DimensionMismatchCount -ne 0) -or
                ($_.MaxMeanAbsoluteError -ne $null -and [double]$_.MaxMeanAbsoluteError -gt $MaxMeanAbsoluteErrorLimit) -or
                ($_.MaxChangedPixelRatioAtThreshold16 -ne $null -and [double]$_.MaxChangedPixelRatioAtThreshold16 -gt $MaxChangedPixelRatioLimit)
        }).Count
}

function Get-RectComparisonRowPage($Row) {
    if ($null -ne $Row.Candidate -and $null -ne $Row.Candidate.Page) {
        return [int]$Row.Candidate.Page
    }

    if ($null -ne $Row.Reference -and $null -ne $Row.Reference.Page) {
        return [int]$Row.Reference.Page
    }

    return $null
}

function New-RectComparisonPageCounter([string] $Name, $Rows, [int] $Page) {
    $pageRows = @($Rows | Where-Object { (Get-RectComparisonRowPage $_) -eq $Page })
    $deltaRows = @($pageRows | Where-Object { $_.Status -ne "ok" })
    [pscustomobject]@{
        Name = $Name
        ComparisonCount = $pageRows.Count
        DeltaCount = $deltaRows.Count
        MissingCandidateCount = @($pageRows | Where-Object { $_.Status -eq "missing-candidate" }).Count
        MissingReferenceCount = @($pageRows | Where-Object { $_.Status -eq "missing-reference" }).Count
        MaxBoundsDelta = Get-MaxOrZero ($pageRows | ForEach-Object { $_.MaxDelta })
    }
}

function Resolve-BalloonSide($Balloons, [double] $PageWidth) {
    $sides = @($Balloons |
        Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.Side) } |
        ForEach-Object { [string]$_.Side } |
        Sort-Object -Unique)
    if ($sides.Count -eq 1) {
        return $sides[0]
    }

    if ($sides.Count -gt 1) {
        return "Mixed"
    }

    if ($Balloons.Count -eq 0 -or $PageWidth -le 0d) {
        return "None"
    }

    $rightCount = @($Balloons | Where-Object { $_.MinX -ne $null -and [double]$_.MinX -ge ($PageWidth / 2d) }).Count
    if ($rightCount -eq $Balloons.Count) {
        return "Right"
    }

    if ($rightCount -eq 0) {
        return "Left"
    }

    return "Mixed"
}

function Test-BalloonConnectorSideConsistent($Balloon) {
    if ($null -eq $Balloon -or $Balloon.IsOverflowSummary -eq $true) {
        return $true
    }

    $side = if ($Balloon.PSObject.Properties.Name -contains "Side") { [string]$Balloon.Side } else { "" }
    $x = Get-DoubleMetric $Balloon @("MinX", "X")
    $width = Get-DoubleMetric $Balloon @("Width")
    $connectorX = Get-DoubleMetric $Balloon @("BalloonConnectorX")
    if ($null -eq $x -or $null -eq $width -or $null -eq $connectorX) {
        return $true
    }

    if ($side -eq "Left") {
        return [double]$connectorX -ge ([double]$x + [double]$width - 0.001d)
    }

    if ($side -eq "Right") {
        return [double]$connectorX -le ([double]$x + 0.001d)
    }

    return $true
}

function New-MarkupGeometryPageCounter(
    [int] $PageNumber,
    $LayoutPage,
    $CandidateBalloons,
    $TableGridComparison,
    $ConnectorComparison,
    $BalloonGraphicComparison,
    $AnnotationComparison)
{
    $width = Get-DoubleMetric $LayoutPage @("Width")
    $height = Get-DoubleMetric $LayoutPage @("Height")
    $marginLeft = Get-DoubleMetric $LayoutPage @("MarginLeft")
    $marginRight = Get-DoubleMetric $LayoutPage @("MarginRight")
    $columnFrameWidthSum = Get-DoubleMetric $LayoutPage @("ColumnFrameWidthSum")
    $markupReserve = Get-DoubleMetric $LayoutPage @("MarkupMarginReservePoints")
    $frames = @()
    if ($null -ne $LayoutPage -and $LayoutPage.PSObject.Properties.Name -contains "ColumnFrames") {
        $frames = @($LayoutPage.ColumnFrames | Where-Object { $null -ne $_ })
    }

    $minFrameX = if ($frames.Count -eq 0) { $null } else { [Math]::Round([double](($frames | Measure-Object -Property X -Minimum).Minimum), 6) }
    $maxFrameRight = if ($frames.Count -eq 0) {
        $null
    }
    else {
        [Math]::Round([double](($frames | ForEach-Object { [double]$_.X + [double]$_.Width } | Measure-Object -Maximum).Maximum), 6)
    }
    $authoredBodyWidth = if ($null -eq $width -or $null -eq $marginLeft -or $null -eq $marginRight) {
        $null
    }
    else {
        [Math]::Round([double]$width - [double]$marginLeft - [double]$marginRight, 6)
    }
    $bodyFrameWidthDeficit = if ($null -eq $authoredBodyWidth -or $null -eq $columnFrameWidthSum) {
        $null
    }
    else {
        [Math]::Round([Math]::Max(0d, [double]$authoredBodyWidth - [double]$columnFrameWidthSum), 6)
    }
    $leftShift = if ($null -eq $minFrameX -or $null -eq $marginLeft) {
        $null
    }
    else {
        [Math]::Round([double]$minFrameX - [double]$marginLeft, 6)
    }
    $rightInset = if ($null -eq $maxFrameRight -or $null -eq $width -or $null -eq $marginRight) {
        $null
    }
    else {
        [Math]::Round(([double]$width - [double]$marginRight) - [double]$maxFrameRight, 6)
    }

    $pageBalloons = @($CandidateBalloons | Where-Object { $_.Page -eq $PageNumber })
    $commentBalloons = @($pageBalloons | Where-Object { $_.Subtype -eq "Comment" })
    $revisionBalloons = @($pageBalloons | Where-Object { $_.Subtype -eq "Revision" })
    $overflowBalloons = @($pageBalloons | Where-Object { $_.IsOverflowSummary -eq $true })
    $connectorBalloons = @($pageBalloons | Where-Object { $_.IsOverflowSummary -ne $true })
    $clampedConnectorBalloons = @($connectorBalloons | Where-Object { $_.AnchorConnectorClamped -eq $true })
    $sideInconsistentConnectorBalloons = @($connectorBalloons | Where-Object { -not (Test-BalloonConnectorSideConsistent $_) })
    $pageWidthForSide = if ($null -eq $width) { 0d } else { [double]$width }
    $side = Resolve-BalloonSide $pageBalloons $pageWidthForSide

    [pscustomobject]@{
        Page = $PageNumber
        Width = if ($null -eq $width) { $null } else { [Math]::Round([double]$width, 6) }
        Height = if ($null -eq $height) { $null } else { [Math]::Round([double]$height, 6) }
        MarginLeft = if ($null -eq $marginLeft) { $null } else { [Math]::Round([double]$marginLeft, 6) }
        MarginRight = if ($null -eq $marginRight) { $null } else { [Math]::Round([double]$marginRight, 6) }
        MarkupMarginReservePoints = if ($null -eq $markupReserve) { $null } else { [Math]::Round([double]$markupReserve, 6) }
        ColumnFrameCount = if ($null -eq $LayoutPage -or $LayoutPage.ColumnFrameCount -eq $null) { $frames.Count } else { [int]$LayoutPage.ColumnFrameCount }
        ColumnFrameWidthSum = if ($null -eq $columnFrameWidthSum) { $null } else { [Math]::Round([double]$columnFrameWidthSum, 6) }
        MinColumnFrameX = $minFrameX
        MaxColumnFrameRight = $maxFrameRight
        AuthoredBodyWidthPoints = $authoredBodyWidth
        BodyFrameWidthDeficitPoints = $bodyFrameWidthDeficit
        BodyFrameLeftShiftPoints = $leftShift
        BodyFrameRightInsetPoints = $rightInset
        MarkupSide = $side
        BalloonLaneWidthPoints = Get-MaxOrZero ($pageBalloons | ForEach-Object { $_.Width })
        BalloonCount = $pageBalloons.Count
        CommentBalloonCount = $commentBalloons.Count
        RevisionBalloonCount = $revisionBalloons.Count
        OverflowBalloonCount = $overflowBalloons.Count
        ConnectorPlacementCount = $connectorBalloons.Count
        ClampedConnectorPlacementCount = $clampedConnectorBalloons.Count
        SideInconsistentConnectorPlacementCount = $sideInconsistentConnectorBalloons.Count
        RectangleDeltas = @(
            New-RectComparisonPageCounter "table-grid" $TableGridComparison $PageNumber
            New-RectComparisonPageCounter "connectors" $ConnectorComparison $PageNumber
            New-RectComparisonPageCounter "balloon-graphics" $BalloonGraphicComparison $PageNumber
            New-RectComparisonPageCounter "annotations" $AnnotationComparison $PageNumber
        )
    }
}

function New-MarkupGeometryCounterSummary(
    $LayoutSnapshot,
    [int] $ReferencePageCount,
    [int] $CandidatePageCount,
    $CandidateBalloons,
    $BodyFrameDeltaSummary,
    $MarkupMarginDeltaSummary,
    $TableGridDeltaSummary,
    $ConnectorDeltaSummary,
    $BalloonGraphicDeltaSummary,
    $AnnotationComparison,
    $TableGridComparison,
    $ConnectorComparison,
    $BalloonGraphicComparison)
{
    $layoutPages = if ($null -ne $LayoutSnapshot -and $LayoutSnapshot.PSObject.Properties.Name -contains "Pages") {
        @($LayoutSnapshot.Pages | Where-Object { $null -ne $_ })
    }
    else {
        @()
    }
    $layoutByPage = @{}
    for ($i = 0; $i -lt $layoutPages.Count; $i++) {
        $page = $layoutPages[$i]
        $pageNumber = if ($page.PSObject.Properties.Name -contains "PageNumber" -and $null -ne $page.PageNumber) { [int]$page.PageNumber } else { $i + 1 }
        $layoutByPage[$pageNumber] = $page
    }

    $pageCount = [Math]::Max([Math]::Max($ReferencePageCount, $CandidatePageCount), $layoutPages.Count)
    $pages = @(
        for ($pageNumber = 1; $pageNumber -le $pageCount; $pageNumber++) {
            New-MarkupGeometryPageCounter `
                -PageNumber $pageNumber `
                -LayoutPage $(if ($layoutByPage.ContainsKey($pageNumber)) { $layoutByPage[$pageNumber] } else { $null }) `
                -CandidateBalloons $CandidateBalloons `
                -TableGridComparison $TableGridComparison `
                -ConnectorComparison $ConnectorComparison `
                -BalloonGraphicComparison $BalloonGraphicComparison `
                -AnnotationComparison $AnnotationComparison
        }
    )

    [pscustomobject]@{
        MarkupGeometryMode = if ($null -eq $LayoutSnapshot -or $LayoutSnapshot.MarkupGeometryMode -eq $null) { $null } else { [string]$LayoutSnapshot.MarkupGeometryMode }
        ReferencePageCount = $ReferencePageCount
        CandidatePageCount = $CandidatePageCount
        CandidateLayoutPageCount = $layoutPages.Count
        ReserveMarginPageCount = @($pages | Where-Object { $_.MarkupMarginReservePoints -ne $null -and [double]$_.MarkupMarginReservePoints -gt 0d }).Count
        MaxMarkupMarginReservePoints = Get-MaxOrZero ($pages | ForEach-Object { $_.MarkupMarginReservePoints })
        MaxBodyFrameWidthDeficitPoints = Get-MaxOrZero ($pages | ForEach-Object { $_.BodyFrameWidthDeficitPoints })
        MaxBodyFrameLeftShiftPoints = Get-MaxOrZero ($pages | ForEach-Object { if ($_.BodyFrameLeftShiftPoints -ne $null) { [Math]::Abs([double]$_.BodyFrameLeftShiftPoints) } })
        MaxBodyFrameRightInsetPoints = Get-MaxOrZero ($pages | ForEach-Object { $_.BodyFrameRightInsetPoints })
        MaxBalloonLaneWidthPoints = Get-MaxOrZero ($pages | ForEach-Object { $_.BalloonLaneWidthPoints })
        MarkupSideBuckets = @($pages | Group-Object -Property MarkupSide | Sort-Object Name | ForEach-Object { [pscustomobject]@{ Side = $_.Name; PageCount = $_.Count } })
        ConnectorPlacementCount = Get-SummarySumMetric $pages "ConnectorPlacementCount"
        ClampedConnectorPlacementCount = Get-SummarySumMetric $pages "ClampedConnectorPlacementCount"
        ClampedConnectorPageCount = @($pages | Where-Object { $_.ClampedConnectorPlacementCount -gt 0 }).Count
        SideInconsistentConnectorPlacementCount = Get-SummarySumMetric $pages "SideInconsistentConnectorPlacementCount"
        SideInconsistentConnectorPageCount = @($pages | Where-Object { $_.SideInconsistentConnectorPlacementCount -gt 0 }).Count
        BodyFrameDelta = $BodyFrameDeltaSummary
        MarkupMarginDelta = $MarkupMarginDeltaSummary
        TableGridDelta = $TableGridDeltaSummary
        ConnectorDelta = $ConnectorDeltaSummary
        BalloonGraphicDelta = $BalloonGraphicDeltaSummary
        AnnotationRectangleDeltaCount = @($AnnotationComparison | Where-Object { $_.Status -ne "ok" }).Count
        AnnotationTargetDeltaCount = @($AnnotationComparison | Where-Object { $_.TargetDelta -eq $true }).Count
        AnnotationTargetKindDeltaCount = @($AnnotationComparison | Where-Object { $_.TargetKindDelta -eq $true }).Count
        AnnotationTargetSha256DeltaCount = @($AnnotationComparison | Where-Object { $_.TargetSha256Delta -eq $true }).Count
        AnnotationMaxBoundsDelta = Get-RectComparisonMaxDelta $AnnotationComparison
        Pages = $pages
    }
}

function Get-RectItemPage($Item) {
    if ($null -ne $Item -and $Item.Page -ne $null) {
        return [int]$Item.Page
    }

    return $null
}

function Get-ComparisonStatusCount($Rows, [string] $Status) {
    return @($Rows | Where-Object { $_.Status -eq $Status }).Count
}

function Get-CandidateLaneBandCount($Balloons) {
    @(
        $Balloons |
            Where-Object { $_.LaneBandIndex -ne $null } |
            ForEach-Object { "{0}:{1}" -f $_.Page, $_.LaneBandIndex } |
            Sort-Object -Unique
    ).Count
}

function Get-MaxCandidateLaneBandCandidateCount($Balloons) {
    $counts = @(
        $Balloons |
            Where-Object { $_.LaneBandCandidateCount -ne $null } |
            ForEach-Object { [int]$_.LaneBandCandidateCount }
    )
    if ($counts.Count -eq 0) {
        return $null
    }

    [int](($counts | Measure-Object -Maximum).Maximum)
}

function New-BalloonClassificationSummary(
    $ReferenceBalloonGraphics,
    $CandidateBalloonGraphics,
    $CandidateMarkupBalloons,
    $BalloonGraphicComparison,
    $BalloonMarkupComparison,
    $ConnectorComparison)
{
    $pageNumbers = New-Object System.Collections.Generic.HashSet[int]
    foreach ($item in @($ReferenceBalloonGraphics) + @($CandidateBalloonGraphics) + @($CandidateMarkupBalloons)) {
        $page = Get-RectItemPage $item
        if ($null -ne $page) {
            [void]$pageNumbers.Add($page)
        }
    }

    foreach ($row in @($BalloonGraphicComparison) + @($BalloonMarkupComparison) + @($ConnectorComparison)) {
        $page = Get-RectComparisonRowPage $row
        if ($null -ne $page) {
            [void]$pageNumbers.Add($page)
        }
    }

    $pages = @(
        foreach ($page in @($pageNumbers | Sort-Object)) {
            $pageCandidateMarkupBalloons = @($CandidateMarkupBalloons | Where-Object { $_.Page -eq $page })
            [pscustomobject]@{
                Page = $page
                ReferenceBalloonGraphicCount = @($ReferenceBalloonGraphics | Where-Object { $_.Page -eq $page }).Count
                CandidateBalloonGraphicCount = @($CandidateBalloonGraphics | Where-Object { $_.Page -eq $page }).Count
                CandidateMarkupBalloonCount = $pageCandidateMarkupBalloons.Count
                CandidateOverflowBalloonCount = @($pageCandidateMarkupBalloons | Where-Object { $_.IsOverflowSummary -eq $true }).Count
                CandidateLaneBandCount = Get-CandidateLaneBandCount $pageCandidateMarkupBalloons
                MaxCandidateLaneBandCandidateCount = Get-MaxCandidateLaneBandCandidateCount $pageCandidateMarkupBalloons
                GraphicRectangleDeltas = New-RectComparisonPageCounter "balloon-body-graphics" $BalloonGraphicComparison $page
                MarkupRectangleDeltas = New-RectComparisonPageCounter "candidate-markup-balloons" $BalloonMarkupComparison $page
                ConnectorDeltas = New-RectComparisonPageCounter "connectors" $ConnectorComparison $page
            }
        }
    )

    [pscustomobject]@{
        ReferenceBalloonGraphicCount = @($ReferenceBalloonGraphics).Count
        CandidateBalloonGraphicCount = @($CandidateBalloonGraphics).Count
        CandidateMarkupBalloonCount = @($CandidateMarkupBalloons).Count
        CandidateOverflowBalloonCount = @($CandidateMarkupBalloons | Where-Object { $_.IsOverflowSummary -eq $true }).Count
        CandidateLaneBandCount = Get-CandidateLaneBandCount $CandidateMarkupBalloons
        MaxCandidateLaneBandCandidateCount = Get-MaxCandidateLaneBandCandidateCount $CandidateMarkupBalloons
        UnmatchedReferenceBalloonGraphicCount = Get-ComparisonStatusCount $BalloonGraphicComparison "missing-candidate"
        UnmatchedCandidateBalloonGraphicCount = Get-ComparisonStatusCount $BalloonGraphicComparison "missing-reference"
        UnmatchedReferenceForMarkupBalloonCount = Get-ComparisonStatusCount $BalloonMarkupComparison "missing-candidate"
        UnmatchedCandidateMarkupBalloonCount = Get-ComparisonStatusCount $BalloonMarkupComparison "missing-reference"
        BalloonGraphicRectangleDeltaCount = Get-ComparisonStatusCount $BalloonGraphicComparison "delta"
        BalloonMarkupRectangleDeltaCount = Get-ComparisonStatusCount $BalloonMarkupComparison "delta"
        MaxBalloonGraphicRectangleDelta = Get-RectComparisonMaxDelta $BalloonGraphicComparison
        MaxBalloonMarkupRectangleDelta = Get-RectComparisonMaxDelta $BalloonMarkupComparison
        ConnectorDeltaCount = @($ConnectorComparison | Where-Object { $_.Status -ne "ok" }).Count
        ConnectorEndpointBoundsMaxDelta = Get-RectComparisonMaxDelta $ConnectorComparison
        ReferenceOverflowBalloonCount = $null
        OverflowCountDelta = $null
        Pages = $pages
    }
}

function ConvertTo-CanonicalDocxMarkup([string] $Value) {
    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    switch ($Value.ToLowerInvariant()) {
        "simple-markup" { return "simple" }
        "all-markup" { return "all" }
        default { return $Value.ToLowerInvariant() }
    }
}

function ConvertTo-CanonicalDocxMarkupGeometry([string] $Value) {
    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    switch ($Value.ToLowerInvariant()) {
        { $_ -in @("preserve-layout", "preserve-document-layout") } { return "preserve" }
        { $_ -in @("reserve", "reserve-margin", "markup-margin", "reserve-markup-margin") } { return "reserve-margin" }
        { $_ -in @("word", "word-compatible-all-markup", "office", "office-compatible", "office-compatible-all-markup") } { return "word-compatible" }
        default { return $Value.ToLowerInvariant() }
    }
}

$caseFull = (Resolve-Path -LiteralPath $Case).Path
$caseDirectory = Split-Path -Parent $caseFull
$manifest = Get-Content -Raw -LiteralPath $caseFull | ConvertFrom-Json
if ($manifest.kind -ne "docx") {
    throw "Cached DOCX markup comparison only supports DOCX cases."
}

$inputFull = (Resolve-Path -LiteralPath (Join-Path $caseDirectory $manifest.input)).Path
$docxMarkup = ConvertTo-CanonicalDocxMarkup $(if (-not [string]::IsNullOrWhiteSpace($DocxMarkup)) {
        $DocxMarkup
    }
    elseif ($manifest.PSObject.Properties.Name -contains "docxMarkup") {
        [string]$manifest.docxMarkup
    }
    else {
        $null
    })
if ([string]::IsNullOrWhiteSpace($docxMarkup)) {
    throw "Cached DOCX markup comparison requires docxMarkup in the case or -DocxMarkup."
}

$docxMarkupGeometry = ConvertTo-CanonicalDocxMarkupGeometry $(if (-not [string]::IsNullOrWhiteSpace($DocxMarkupGeometry)) {
        $DocxMarkupGeometry
    }
    elseif ($manifest.PSObject.Properties.Name -contains "docxMarkupGeometry") {
        [string]$manifest.docxMarkupGeometry
    }
    else {
        $null
    })
if ([string]::IsNullOrWhiteSpace($docxMarkupGeometry)) {
    $docxMarkupGeometry = "preserve"
}

$resolvedCaseId = if (-not [string]::IsNullOrWhiteSpace($CaseId)) {
    $CaseId
}
elseif (-not [string]::IsNullOrWhiteSpace([string]$manifest.id)) {
    [string]$manifest.id
}
else {
    "docx-markup-reference"
}
$runId = Get-Date -Format "yyyyMMdd-HHmmss"
$runRoot = if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    Join-Path $repoRoot ("artifacts/docx-markup-reference/{0}/{1}" -f $resolvedCaseId, $runId)
}
else {
    $OutputDirectory
}
$runRoot = [System.IO.Path]::GetFullPath($runRoot)
$referenceDir = Join-Path $runRoot "reference"
$candidateDir = Join-Path $runRoot "candidate"
$comparisonDir = Join-Path $runRoot "comparison"
$candidateDocxInspect = Join-Path $comparisonDir "candidate-docx-inspect"
$referencePdfInspect = Join-Path $comparisonDir "pdf-reference"
$candidatePdfInspect = Join-Path $comparisonDir "pdf-candidate"
$annotationsDir = Join-Path $comparisonDir "annotations"
$balloonsDir = Join-Path $comparisonDir "balloons"
$geometryDir = Join-Path $comparisonDir "geometry"
$rasterDir = Join-Path $comparisonDir "raster"
$referenceRasterDir = Join-Path $rasterDir "reference"
$candidateRasterDir = Join-Path $rasterDir "candidate"
$visualDiffDir = Join-Path $rasterDir "diff"
New-Item -ItemType Directory -Force -Path $referenceDir, $candidateDir, $comparisonDir, $annotationsDir, $balloonsDir, $geometryDir, $rasterDir | Out-Null

if (-not [string]::IsNullOrWhiteSpace($ReferencePdf) -and -not [string]::IsNullOrWhiteSpace($ReferenceDirectory)) {
    throw "Use either -ReferencePdf or -ReferenceDirectory, not both."
}

if (-not [string]::IsNullOrWhiteSpace($ReferencePdf)) {
    Copy-Item -LiteralPath (Resolve-Path -LiteralPath $ReferencePdf).Path -Destination (Join-Path $referenceDir "reference.pdf") -Force
}
elseif (-not [string]::IsNullOrWhiteSpace($ReferenceDirectory)) {
    $referenceDirectoryFull = (Resolve-Path -LiteralPath $ReferenceDirectory).Path
    if (-not (Test-Path -LiteralPath (Join-Path $referenceDirectoryFull "reference.pdf"))) {
        throw "Reference directory does not contain reference.pdf: $referenceDirectoryFull"
    }

    Copy-Item -Path (Join-Path $referenceDirectoryFull "*") -Destination $referenceDir -Recurse -Force
}
else {
    $referenceCacheVariant = "docxMarkup={0};docxMarkupGeometry={1}" -f $docxMarkup, $docxMarkupGeometry
    & (Join-Path $PSScriptRoot "RenderCachedReference.ps1") -InputPath $inputFull -OutputDirectory $referenceDir -Dpi $Dpi -CacheOnly -CacheVariant $referenceCacheVariant
}

$referencePdfPath = Join-Path $referenceDir "reference.pdf"
if (-not (Test-Path -LiteralPath $referencePdfPath)) {
    throw "Cached reference did not provide reference.pdf."
}

if ($ValidateOnly) {
    if ($PrivateSafeSummary) {
        Write-Host "Cached DOCX markup comparison validation passed for private-safe case: $resolvedCaseId"
    }
    else {
        Write-Host "Cached DOCX markup comparison validation passed: $caseFull"
    }
    Write-Host "Reference PDF: $referencePdfPath"
    return
}

$cliProject = Join-Path $repoRoot "src/Lokad.OoxPdf.Cli/Lokad.OoxPdf.Cli.csproj"
$cliDll = Join-Path $repoRoot "src/Lokad.OoxPdf.Cli/bin/Debug/net10.0/Lokad.OoxPdf.Cli.dll"
$librarySourceDirectory = Join-Path $repoRoot "src/Lokad.OoxPdf"
Invoke-DotnetBuildIfStale -Project $cliProject -OutputDll $cliDll -Description "CLI" -AdditionalSourceDirectories @($librarySourceDirectory)

$candidatePdf = Join-Path $candidateDir "output.pdf"
$diagnostics = Join-Path $candidateDir "diagnostics.json"
dotnet $cliDll convert $inputFull $candidatePdf --diagnostics $diagnostics --docx-markup $docxMarkup --docx-markup-geometry $docxMarkupGeometry
if ($LASTEXITCODE -ne 0) {
    throw "Candidate conversion failed with exit code $LASTEXITCODE."
}

& (Join-Path $PSScriptRoot "InspectDocx.ps1") -InputDocx $inputFull -OutputDirectory $candidateDocxInspect -DocxMarkup $docxMarkup -DocxMarkupGeometry $docxMarkupGeometry | Out-Null
if ($LASTEXITCODE -ne 0) {
    throw "DOCX inspection failed with exit code $LASTEXITCODE."
}

& (Join-Path $PSScriptRoot "InspectPdf.ps1") -InputPdf $referencePdfPath -OutputDirectory $referencePdfInspect | Out-Null
if ($LASTEXITCODE -ne 0) {
    throw "Reference PDF inspection failed with exit code $LASTEXITCODE."
}

& (Join-Path $PSScriptRoot "InspectPdf.ps1") -InputPdf $candidatePdf -OutputDirectory $candidatePdfInspect | Out-Null
if ($LASTEXITCODE -ne 0) {
    throw "Candidate PDF inspection failed with exit code $LASTEXITCODE."
}

$visualMetrics = @()
if (-not $SkipRasterDiff) {
    & (Join-Path $PSScriptRoot "RasterizePdf.ps1") -InputPdf $referencePdfPath -OutputDirectory $referenceRasterDir -Dpi $Dpi | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "Reference PDF rasterization failed with exit code $LASTEXITCODE."
    }

    & (Join-Path $PSScriptRoot "RasterizePdf.ps1") -InputPdf $candidatePdf -OutputDirectory $candidateRasterDir -Dpi $Dpi | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "Candidate PDF rasterization failed with exit code $LASTEXITCODE."
    }

    $visualDiffProject = Join-Path $repoRoot "tools/Lokad.OoxPdf.VisualDiff/Lokad.OoxPdf.VisualDiff.csproj"
    $visualDiffDll = Join-Path $repoRoot "tools/Lokad.OoxPdf.VisualDiff/bin/Debug/net10.0/Lokad.OoxPdf.VisualDiff.dll"
    Invoke-DotnetBuildIfStale -Project $visualDiffProject -OutputDll $visualDiffDll -Description "VisualDiff"
    dotnet $visualDiffDll $referenceRasterDir $candidateRasterDir $visualDiffDir | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "VisualDiff failed with exit code $LASTEXITCODE."
    }

    $visualMetrics = @(Read-JsonArray (Join-Path $visualDiffDir "metrics.json"))
}

$referenceTextOperations = Join-Path $referencePdfInspect "text-operations.json"
$candidateTextOperations = Join-Path $candidatePdfInspect "text-operations.json"
$referenceTextOperationItems = @(Read-JsonArray $referenceTextOperations)
$candidateTextOperationItems = @(Read-JsonArray $candidateTextOperations)
$referenceFontResources = @(Read-JsonArray (Join-Path $referencePdfInspect "font-resources.json"))
$candidateFontResources = @(Read-JsonArray (Join-Path $candidatePdfInspect "font-resources.json"))
$fontResourceSummary = [ordered]@{
    Reference = New-FontResourceSummary $referenceFontResources
    Candidate = New-FontResourceSummary $candidateFontResources
    CountDelta = [Math]::Abs($candidateFontResources.Count - $referenceFontResources.Count)
    MissingEmbeddedDelta = [Math]::Abs(
        @($candidateFontResources | Where-Object { $_.Embedded -ne $true }).Count -
        @($referenceFontResources | Where-Object { $_.Embedded -ne $true }).Count)
    MissingToUnicodeDelta = [Math]::Abs(
        @($candidateFontResources | Where-Object { $_.HasToUnicode -ne $true }).Count -
        @($referenceFontResources | Where-Object { $_.HasToUnicode -ne $true }).Count)
}
$textComparisonExitCode = $null
if ((Test-Path -LiteralPath $referenceTextOperations) -and (Test-Path -LiteralPath $candidateTextOperations)) {
    $textLog = Join-Path $comparisonDir "text-operations.txt"
    & (Join-Path $PSScriptRoot "ComparePdfTextOperations.ps1") `
        -Reference $referenceTextOperations `
        -Candidate $candidateTextOperations `
        -PositionTolerance $TextPositionTolerance `
        -FontSizeTolerance $TextFontSizeTolerance `
        -CharacterSpacingTolerance $TextCharacterSpacingTolerance `
        -MatchByPosition `
        -UseEffectiveMatrix *> $textLog
    $textComparisonExitCode = $LASTEXITCODE
}

$referenceGraphicsOperations = Join-Path $referencePdfInspect "graphics-operations.json"
$candidateGraphicsOperations = Join-Path $candidatePdfInspect "graphics-operations.json"
$referenceGraphicsOperationItems = @(Read-JsonArray $referenceGraphicsOperations)
$candidateGraphicsOperationItems = @(Read-JsonArray $candidateGraphicsOperations)
$referenceVisibleGraphicsOperationItems = @(Select-VisibleGraphicsOperations $referenceGraphicsOperationItems)
$candidateVisibleGraphicsOperationItems = @(Select-VisibleGraphicsOperations $candidateGraphicsOperationItems)
$graphicsComparisonExitCode = $null
$graphicsVisibleComparisonExitCode = $null
if ((Test-Path -LiteralPath $referenceGraphicsOperations) -and (Test-Path -LiteralPath $candidateGraphicsOperations)) {
    $graphicsLog = Join-Path $comparisonDir "graphics-operations.txt"
    & (Join-Path $PSScriptRoot "ComparePdfGraphicsOperations.ps1") `
        -Reference $referenceGraphicsOperations `
        -Candidate $candidateGraphicsOperations `
        -BoundsTolerance $GraphicsBoundsTolerance `
        -MatchByBounds *> $graphicsLog
    $graphicsComparisonExitCode = $LASTEXITCODE

    $graphicsVisibleLog = Join-Path $comparisonDir "graphics-visible-operations.txt"
    & (Join-Path $PSScriptRoot "ComparePdfGraphicsOperations.ps1") `
        -Reference $referenceGraphicsOperations `
        -Candidate $candidateGraphicsOperations `
        -BoundsTolerance $GraphicsBoundsTolerance `
        -Kinds "Stroke", "Fill", "FillStroke" `
        -MatchByBounds *> $graphicsVisibleLog
    $graphicsVisibleComparisonExitCode = $LASTEXITCODE
}

$referenceAnnotations = @(Get-PdfAnnotations $referencePdfPath)
$candidateAnnotations = @(Get-PdfAnnotations $candidatePdf)
$annotationComparison = @(Compare-RectLists $referenceAnnotations $candidateAnnotations $AnnotationBoundsTolerance)
$annotationDeltaSummary = New-RectDeltaSummary "annotations" $referenceAnnotations $candidateAnnotations $annotationComparison
Write-JsonFile (Join-Path $annotationsDir "reference.json") $referenceAnnotations
Write-JsonFile (Join-Path $annotationsDir "candidate.json") $candidateAnnotations
Write-JsonFile (Join-Path $annotationsDir "comparison.json") $annotationComparison
Write-JsonFile (Join-Path $annotationsDir "summary.json") $annotationDeltaSummary

$candidateBalloonsPath = Join-Path $candidateDocxInspect "markup-balloons.json"
$candidateBalloons = @(Read-JsonArray $candidateBalloonsPath | ForEach-Object { Convert-BalloonToRect $_ })
$referenceBalloonGraphics = @(Select-ReferenceBalloonGraphicCandidates $referenceGraphicsOperationItems)
$candidateBalloonGraphics = @(Select-ReferenceBalloonGraphicCandidates $candidateGraphicsOperationItems)
$broadBalloonGraphicComparison = @(Compare-RectLists $referenceBalloonGraphics $candidateBalloonGraphics $BalloonBoundsTolerance)
$balloonComparison = @()
$balloonGraphicComparison = @()
$balloonClassificationSummary = $null
Write-JsonFile (Join-Path $balloonsDir "candidate-markup-balloons.json") $candidateBalloons
Write-JsonFile (Join-Path $balloonsDir "reference-graphic-candidates.json") $referenceBalloonGraphics
Write-JsonFile (Join-Path $balloonsDir "candidate-graphic-candidates.json") $candidateBalloonGraphics
Write-JsonFile (Join-Path $balloonsDir "broad-graphic-comparison.json") $broadBalloonGraphicComparison

$candidateMarkupSummaryPath = Join-Path $candidateDocxInspect "markup-summary.json"
$candidateMarkupSummary = if (Test-Path -LiteralPath $candidateMarkupSummaryPath) {
    Get-Content -Raw -LiteralPath $candidateMarkupSummaryPath | ConvertFrom-Json
}
else {
    $null
}
$candidateMarkupDensity = if ($null -ne $candidateMarkupSummary -and
    $candidateMarkupSummary.PSObject.Properties.Name -contains "QualityCounters" -and
    $candidateMarkupSummary.QualityCounters.PSObject.Properties.Name -contains "PageLocalMarkupDensity") {
    @($candidateMarkupSummary.QualityCounters.PageLocalMarkupDensity)
}
else {
    @()
}
$candidateLayoutSnapshot = Read-JsonObject (Join-Path $candidateDocxInspect "layout-snapshot.json")
$candidateTextEmissionSummary = Read-JsonObject (Join-Path $candidateDocxInspect "text-emission-summary.json")
$candidateSourceBlocks = @(Read-JsonArray (Join-Path $candidateDocxInspect "source-block-summary.json"))
$candidateLayoutTableCount = if ($null -ne $candidateLayoutSnapshot -and $candidateLayoutSnapshot.PSObject.Properties.Name -contains "Tables") {
    @($candidateLayoutSnapshot.Tables | Where-Object { $null -ne $_ }).Count
}
else {
    0
}

$referencePageCount = Get-PdfPageCount $referencePdfPath
$candidatePageCount = Get-PdfPageCount $candidatePdf
$referenceMediaBoxes = @(Get-PdfMediaBoxes $referencePdfPath)
$candidateMediaBoxes = @(Get-PdfMediaBoxes $candidatePdf)
$candidateLayoutBodyFrameRects = @(Select-LayoutBodyFrameRects $candidateLayoutSnapshot)
$candidateLayoutColumnFrameRects = @(Select-LayoutColumnFrameRects $candidateLayoutSnapshot)
$candidateLayoutMarkupLaneRects = @(Select-LayoutMarkupLaneRects $candidateLayoutSnapshot)
$referenceMarkupLaneRects = @(Select-ReferenceMarkupLaneRects $referenceBalloonGraphics $referenceMediaBoxes)
$markupLaneHorizontalComparison = @(Compare-MarkupLaneHorizontalRects $referenceMarkupLaneRects $candidateLayoutMarkupLaneRects $MarkupMarginTolerance)
$markupLaneHorizontalDeltaSummary = New-MarkupLaneHorizontalDeltaSummary $referenceMarkupLaneRects $candidateLayoutMarkupLaneRects $markupLaneHorizontalComparison
$bodyFrameLaneEdgeComparison = @(Compare-BodyFrameLaneEdgeRects $referenceMarkupLaneRects $candidateLayoutBodyFrameRects $candidateLayoutMarkupLaneRects $BodyFrameTolerance)
$bodyFrameLaneEdgeDeltaSummary = New-BodyFrameLaneEdgeDeltaSummary $referenceMarkupLaneRects $candidateLayoutBodyFrameRects $candidateLayoutMarkupLaneRects $bodyFrameLaneEdgeComparison
$referenceBalloonBodyGraphics = @(Select-BalloonBodyGraphicCandidates $referenceGraphicsOperationItems $referenceMarkupLaneRects)
$candidateBalloonBodyGraphics = @(Select-BalloonBodyGraphicCandidates $candidateGraphicsOperationItems $candidateLayoutMarkupLaneRects)
$balloonComparison = @(Compare-RectLists $referenceBalloonBodyGraphics $candidateBalloons $BalloonBoundsTolerance)
$balloonGraphicComparison = @(Compare-RectLists $referenceBalloonBodyGraphics $candidateBalloonBodyGraphics $BalloonBoundsTolerance)
Write-JsonFile (Join-Path $balloonsDir "reference-body-graphic-candidates.json") $referenceBalloonBodyGraphics
Write-JsonFile (Join-Path $balloonsDir "candidate-body-graphic-candidates.json") $candidateBalloonBodyGraphics
Write-JsonFile (Join-Path $balloonsDir "comparison.json") $balloonComparison
Write-JsonFile (Join-Path $balloonsDir "graphic-comparison.json") $balloonGraphicComparison
$pageDeltaSummary = @(New-PageDeltaSummary `
        -ReferencePageCount $referencePageCount `
        -CandidatePageCount $candidatePageCount `
        -VisualMetrics $visualMetrics `
        -ReferenceTextOperations $referenceTextOperationItems `
        -CandidateTextOperations $candidateTextOperationItems `
        -ReferenceGraphicsOperations $referenceGraphicsOperationItems `
        -CandidateGraphicsOperations $candidateGraphicsOperationItems `
        -AnnotationComparison $annotationComparison `
        -BalloonComparison $balloonComparison `
        -CandidateMarkupDensity $candidateMarkupDensity |
    Sort-Object -Property PriorityScore -Descending)
$pixelDiffTriage = New-PixelDiffTriageReport `
    -PageDeltaSummary $pageDeltaSummary `
    -SkipRasterDiff:$SkipRasterDiff `
    -MaxChangedPixelRatio $MaxChangedPixelRatioAtThreshold16 `
    -MaxMeanAbsoluteError $MaxMeanAbsoluteError
$regionDeltaSummary = @(New-RegionDeltaSummary `
        -ReferenceGraphicsOperations $referenceGraphicsOperationItems `
        -CandidateGraphicsOperations $candidateGraphicsOperationItems `
        -ReferenceTextOperations $referenceTextOperationItems `
        -CandidateTextOperations $candidateTextOperationItems `
        -CandidateMarkupDensity $candidateMarkupDensity `
        -ReferenceBalloonGraphics $referenceBalloonBodyGraphics `
        -CandidateBalloons $candidateBalloons)

$referenceTableRects = @(Select-GraphicRects $referenceGraphicsOperationItems -Predicate { param($op) Test-TableLikeGraphic $op } -Subtype "table-grid")
$candidateTableRects = @(Select-GraphicRects $candidateGraphicsOperationItems -Predicate { param($op) Test-TableLikeGraphic $op } -Subtype "table-grid")
$tableGridComparison = @(Compare-RectLists $referenceTableRects $candidateTableRects $TableGridBoundsTolerance)
$tableGridDeltaSummary = New-RectDeltaSummary "table-grid" $referenceTableRects $candidateTableRects $tableGridComparison
$tableGridDeltaCountForGate = if ($candidateLayoutTableCount -eq 0) { $null } else { $tableGridDeltaSummary.DeltaCount }
$tableGridMaxBoundsDeltaForGate = if ($candidateLayoutTableCount -eq 0) { $null } else { $tableGridDeltaSummary.MaxBoundsDelta }

$referenceConnectorRects = @(Select-GraphicRects $referenceGraphicsOperationItems -Predicate { param($op) Test-ConnectorLikeGraphic $op } -Subtype "connector")
$candidateConnectorRects = @(Select-GraphicRects $candidateGraphicsOperationItems -Predicate { param($op) Test-ConnectorLikeGraphic $op } -Subtype "connector")
$connectorComparison = @(Compare-RectLists $referenceConnectorRects $candidateConnectorRects $ConnectorBoundsTolerance)
$connectorDeltaSummary = New-RectDeltaSummary "connectors" $referenceConnectorRects $candidateConnectorRects $connectorComparison
$referenceConnectorEndpoints = @(Select-ConnectorEndpointItems $referenceGraphicsOperationItems)
$candidateConnectorEndpoints = @(Select-ConnectorEndpointItems $candidateGraphicsOperationItems)
$connectorEndpointComparison = @(Compare-ConnectorEndpointLists $referenceConnectorEndpoints $candidateConnectorEndpoints $ConnectorBoundsTolerance)
$connectorEndpointDeltaSummary = New-ConnectorEndpointDeltaSummary $referenceConnectorEndpoints $candidateConnectorEndpoints $connectorEndpointComparison
$referenceConnectorLaneEndpoints = @($referenceConnectorEndpoints | Where-Object { Test-ConnectorLaneEndpointItem $_ })
$candidateConnectorLaneEndpoints = @($candidateConnectorEndpoints | Where-Object { Test-ConnectorLaneEndpointItem $_ })
$connectorLaneEndpointComparison = @(Compare-ConnectorEndpointLists $referenceConnectorLaneEndpoints $candidateConnectorLaneEndpoints $ConnectorBoundsTolerance)
$connectorLaneEndpointDeltaSummary = New-ConnectorEndpointDeltaSummary $referenceConnectorLaneEndpoints $candidateConnectorLaneEndpoints $connectorLaneEndpointComparison "connector-lane-endpoints"
$referenceConnectorBodyAnchorEndpoints = @($referenceConnectorEndpoints | Where-Object { -not (Test-ConnectorLaneEndpointItem $_) })
$candidateConnectorBodyAnchorEndpoints = @($candidateConnectorEndpoints | Where-Object { -not (Test-ConnectorLaneEndpointItem $_) })
$connectorBodyAnchorEndpointComparison = @(Compare-ConnectorEndpointLists $referenceConnectorBodyAnchorEndpoints $candidateConnectorBodyAnchorEndpoints $ConnectorBoundsTolerance)
$connectorBodyAnchorEndpointDeltaSummary = New-ConnectorEndpointDeltaSummary $referenceConnectorBodyAnchorEndpoints $candidateConnectorBodyAnchorEndpoints $connectorBodyAnchorEndpointComparison "connector-body-anchor-endpoints"
$broadBalloonGraphicDeltaSummary = New-RectDeltaSummary "broad-balloon-graphics" $referenceBalloonGraphics $candidateBalloonGraphics $broadBalloonGraphicComparison
$balloonGraphicDeltaSummary = New-RectDeltaSummary "balloon-body-graphics" $referenceBalloonBodyGraphics $candidateBalloonBodyGraphics $balloonGraphicComparison
$balloonClassificationSummary = New-BalloonClassificationSummary `
    -ReferenceBalloonGraphics $referenceBalloonBodyGraphics `
    -CandidateBalloonGraphics $candidateBalloonBodyGraphics `
    -CandidateMarkupBalloons $candidateBalloons `
    -BalloonGraphicComparison $balloonGraphicComparison `
    -BalloonMarkupComparison $balloonComparison `
    -ConnectorComparison $connectorComparison

$bodyFrameDeltaSummary = New-PointBoundsDeltaSummary "body-frame" $referenceTextOperationItems $candidateTextOperationItems -Predicate { param($op) Test-BodyTextRegion $op }
$markupMarginDeltaSummary = New-PointBoundsDeltaSummary "markup-margin" $referenceTextOperationItems $candidateTextOperationItems -Predicate {
    param($op)
    $x = Get-OperationX $op
    $null -ne $x -and ($x -gt 420d -or $x -lt 120d)
}
$textGateDeltaSummary = New-PartitionedTextGateDeltaSummary `
    -ReferenceTextOperations $referenceTextOperationItems `
    -CandidateTextOperations $candidateTextOperationItems `
    -ReferenceMarkupLaneRects $referenceMarkupLaneRects `
    -CandidateMarkupLaneRects $candidateLayoutMarkupLaneRects
$textOperationDriftSummary = New-TextOperationDriftSummary `
    -ReferenceTextOperations $referenceTextOperationItems `
    -CandidateTextOperations $candidateTextOperationItems `
    -ReferenceMarkupLaneRects $referenceMarkupLaneRects `
    -CandidateMarkupLaneRects $candidateLayoutMarkupLaneRects
$mediaBoxDeltaSummary = New-MediaBoxDeltaSummary $referenceMediaBoxes $candidateMediaBoxes
$pageFlowDeltaSummary = New-PageFlowDeltaSummary `
    -ReferenceTextOperations $referenceTextOperationItems `
    -CandidateTextOperations $candidateTextOperationItems `
    -CandidateLayoutSnapshot $candidateLayoutSnapshot `
    -CandidateTextEmissionSummary $candidateTextEmissionSummary `
    -CandidateSourceBlocks $candidateSourceBlocks `
    -ReferencePageCount $referencePageCount `
    -CandidatePageCount $candidatePageCount
$markupGeometryCounterSummary = New-MarkupGeometryCounterSummary `
    -LayoutSnapshot $candidateLayoutSnapshot `
    -ReferencePageCount $referencePageCount `
    -CandidatePageCount $candidatePageCount `
    -CandidateBalloons $candidateBalloons `
    -BodyFrameDeltaSummary $bodyFrameDeltaSummary `
    -MarkupMarginDeltaSummary $markupMarginDeltaSummary `
    -TableGridDeltaSummary $tableGridDeltaSummary `
    -ConnectorDeltaSummary $connectorDeltaSummary `
    -BalloonGraphicDeltaSummary $balloonGraphicDeltaSummary `
    -AnnotationComparison $annotationComparison `
    -TableGridComparison $tableGridComparison `
    -ConnectorComparison $connectorComparison `
    -BalloonGraphicComparison $balloonGraphicComparison

Write-JsonFile (Join-Path $comparisonDir "page-delta-summary.json") $pageDeltaSummary
Write-JsonFile (Join-Path $comparisonDir "pixel-diff-triage.json") $pixelDiffTriage
Write-JsonFile (Join-Path $comparisonDir "region-delta-summary.json") $regionDeltaSummary
Write-JsonFile (Join-Path $comparisonDir "page-flow-delta-summary.json") $pageFlowDeltaSummary
Write-JsonFile (Join-Path $comparisonDir "media-boxes.json") ([ordered]@{
        Reference = $referenceMediaBoxes
        Candidate = $candidateMediaBoxes
    })
Write-JsonFile (Join-Path $geometryDir "table-grid-comparison.json") $tableGridComparison
Write-JsonFile (Join-Path $geometryDir "connector-comparison.json") $connectorComparison
Write-JsonFile (Join-Path $geometryDir "connector-endpoint-comparison.json") $connectorEndpointComparison
Write-JsonFile (Join-Path $geometryDir "connector-endpoint-summary.json") $connectorEndpointDeltaSummary
Write-JsonFile (Join-Path $geometryDir "connector-lane-endpoint-comparison.json") $connectorLaneEndpointComparison
Write-JsonFile (Join-Path $geometryDir "connector-lane-endpoint-summary.json") $connectorLaneEndpointDeltaSummary
Write-JsonFile (Join-Path $geometryDir "connector-body-anchor-endpoint-comparison.json") $connectorBodyAnchorEndpointComparison
Write-JsonFile (Join-Path $geometryDir "connector-body-anchor-endpoint-summary.json") $connectorBodyAnchorEndpointDeltaSummary
Write-JsonFile (Join-Path $geometryDir "geometry-frame-summary.json") @($bodyFrameDeltaSummary, $markupMarginDeltaSummary)
Write-JsonFile (Join-Path $geometryDir "markup-geometry-summary.json") $markupGeometryCounterSummary
Write-JsonFile (Join-Path $geometryDir "text-gate-delta-summary.json") $textGateDeltaSummary
Write-JsonFile (Join-Path $geometryDir "text-operation-drift-summary.json") $textOperationDriftSummary
Write-JsonFile (Join-Path $geometryDir "rect-delta-summary.json") @($tableGridDeltaSummary, $connectorDeltaSummary, $annotationDeltaSummary, $balloonGraphicDeltaSummary, $broadBalloonGraphicDeltaSummary)
Write-JsonFile (Join-Path $geometryDir "layout-body-frame-rects.json") $candidateLayoutBodyFrameRects
Write-JsonFile (Join-Path $geometryDir "layout-column-frame-rects.json") $candidateLayoutColumnFrameRects
Write-JsonFile (Join-Path $geometryDir "layout-markup-lane-rects.json") $candidateLayoutMarkupLaneRects
Write-JsonFile (Join-Path $geometryDir "reference-markup-lane-rects.json") $referenceMarkupLaneRects
Write-JsonFile (Join-Path $geometryDir "markup-lane-horizontal-comparison.json") $markupLaneHorizontalComparison
Write-JsonFile (Join-Path $geometryDir "markup-lane-horizontal-summary.json") $markupLaneHorizontalDeltaSummary
Write-JsonFile (Join-Path $geometryDir "body-frame-lane-edge-comparison.json") $bodyFrameLaneEdgeComparison
Write-JsonFile (Join-Path $geometryDir "body-frame-lane-edge-summary.json") $bodyFrameLaneEdgeDeltaSummary
Write-JsonFile (Join-Path $balloonsDir "classification-summary.json") $balloonClassificationSummary
Write-JsonFile (Join-Path $balloonsDir "collision-summary.json") $balloonClassificationSummary
Write-JsonFile (Join-Path $comparisonDir "font-resource-summary.json") $fontResourceSummary
Write-JsonFile (Join-Path $comparisonDir "graphics-operation-count-summary.json") ([pscustomobject]@{
        ReferenceGraphicsOperationCount = $referenceGraphicsOperationItems.Count
        CandidateGraphicsOperationCount = $candidateGraphicsOperationItems.Count
        GraphicsOperationCountDelta = [Math]::Abs($candidateGraphicsOperationItems.Count - $referenceGraphicsOperationItems.Count)
        ReferenceVisibleGraphicsOperationCount = $referenceVisibleGraphicsOperationItems.Count
        CandidateVisibleGraphicsOperationCount = $candidateVisibleGraphicsOperationItems.Count
        VisibleGraphicsOperationCountDelta = [Math]::Abs($candidateVisibleGraphicsOperationItems.Count - $referenceVisibleGraphicsOperationItems.Count)
        ReferenceClipGraphicsOperationCount = @($referenceGraphicsOperationItems | Where-Object { $_.Kind -eq "Clip" }).Count
        CandidateClipGraphicsOperationCount = @($candidateGraphicsOperationItems | Where-Object { $_.Kind -eq "Clip" }).Count
        ClipGraphicsOperationCountDelta = [Math]::Abs(
            @($candidateGraphicsOperationItems | Where-Object { $_.Kind -eq "Clip" }).Count -
            @($referenceGraphicsOperationItems | Where-Object { $_.Kind -eq "Clip" }).Count)
        GraphicsComparisonExitCode = $graphicsComparisonExitCode
        GraphicsVisibleComparisonExitCode = $graphicsVisibleComparisonExitCode
    })

$rasterRegionSpecs = @()
$rasterRegionMetrics = @()
$rasterRegionSummary = @()
if (-not $SkipRasterDiff) {
    $rasterRegionSpecs = @(New-RasterRegionSpecs `
            -ReferencePageCount $referencePageCount `
            -CandidatePageCount $candidatePageCount `
            -MediaBoxes $referenceMediaBoxes `
            -ReferenceGraphicsOperations $referenceGraphicsOperationItems `
            -CandidateGraphicsOperations $candidateGraphicsOperationItems `
            -CandidateBalloons $candidateBalloons)
    $rasterRegionSpecPath = Join-Path $comparisonDir "raster-region-specs.json"
    Write-JsonFile $rasterRegionSpecPath $rasterRegionSpecs
    if ($rasterRegionSpecs.Count -ne 0) {
        dotnet $visualDiffDll $referenceRasterDir $candidateRasterDir $visualDiffDir $rasterRegionSpecPath | Out-Null
        if ($LASTEXITCODE -ne 0) {
            throw "VisualDiff region comparison failed with exit code $LASTEXITCODE."
        }

        $rasterRegionMetrics = @(Read-JsonArray (Join-Path $visualDiffDir "region-metrics.json"))
        $rasterRegionSummary = @(New-RasterRegionMetricSummary $rasterRegionMetrics)
    }
}
Write-JsonFile (Join-Path $comparisonDir "raster-region-summary.json") $rasterRegionSummary

$annotationDeltaCount = @($annotationComparison | Where-Object { $_.Status -ne "ok" }).Count
$balloonDeltaCount = @($balloonComparison | Where-Object { $_.Status -ne "ok" }).Count
$visualDeltaCount = @($visualMetrics | Where-Object {
        $_.DimensionsMatch -ne $true -or
            ($_.MeanAbsoluteError -ne $null -and [double]$_.MeanAbsoluteError -gt $MaxMeanAbsoluteError) -or
            ($_.ChangedPixelRatioAtThreshold16 -ne $null -and [double]$_.ChangedPixelRatioAtThreshold16 -gt $MaxChangedPixelRatioAtThreshold16)
    }).Count
$rasterVisualDeltaCount = if ($SkipRasterDiff) { $null } else { $visualDeltaCount }
$maxVisualMeanAbsoluteError = if ($SkipRasterDiff) { $null } else { Get-MaxVisualMetric $visualMetrics "MeanAbsoluteError" }
$maxVisualChangedPixelRatioAtThreshold16 = if ($SkipRasterDiff) { $null } else { Get-MaxVisualMetric $visualMetrics "ChangedPixelRatioAtThreshold16" }
$rasterRegionDeltaCount = if ($SkipRasterDiff) { $null } else { Get-RasterRegionDeltaCount $rasterRegionSummary $MaxMeanAbsoluteError $MaxChangedPixelRatioAtThreshold16 }
$rasterRegionDimensionMismatchCount = if ($SkipRasterDiff) { $null } else { Get-SummarySumMetric $rasterRegionSummary "DimensionMismatchCount" }
$maxRasterRegionMeanAbsoluteError = if ($SkipRasterDiff) { $null } else { Get-SummaryMaxMetric $rasterRegionSummary "MaxMeanAbsoluteError" }
$maxRasterRegionChangedPixelRatioAtThreshold16 = if ($SkipRasterDiff) { $null } else { Get-SummaryMaxMetric $rasterRegionSummary "MaxChangedPixelRatioAtThreshold16" }
$gateSummary = New-GateSummary @(
    New-GateResult "page-count-delta" "pagination" ([Math]::Abs($candidatePageCount - $referencePageCount)) $MaxPageCountDelta "pages"
    New-GateResult "media-box-count-delta" "page-geometry" $mediaBoxDeltaSummary.CountDelta 0 "boxes"
    New-GateResult "media-box-max-delta" "page-geometry" $mediaBoxDeltaSummary.MaxDelta $MediaBoxTolerance "points"
    New-GateResult "body-frame-lane-edge-max-delta" "page-geometry" $bodyFrameLaneEdgeDeltaSummary.MaxBodyEdgeDelta $BodyFrameTolerance "points"
    New-GateResult "body-frame-lane-side-mismatch-count" "page-geometry" $bodyFrameLaneEdgeDeltaSummary.SideMismatchCount 0 "pages"
    New-GateResult "markup-lane-side-mismatch-count" "markup-geometry" $markupLaneHorizontalDeltaSummary.SideMismatchCount 0 "pages"
    New-GateResult "markup-lane-body-edge-max-delta" "markup-geometry" $markupLaneHorizontalDeltaSummary.MaxBodyEdgeDelta $MarkupMarginTolerance "points"
    New-GateResult "page-flow-first-baseline-max-delta" "pagination" $pageFlowDeltaSummary.MaxFirstBaselineDelta $TextBaselineTolerance "points"
    New-GateResult "page-flow-last-baseline-max-delta" "pagination" $pageFlowDeltaSummary.MaxLastBaselineDelta $TextBaselineTolerance "points"
    New-GateResult "page-flow-body-height-used-max-delta" "pagination" $pageFlowDeltaSummary.MaxBodyFrameHeightUsedDelta $BodyFrameTolerance "points"
    New-GateResult "text-baseline-max-delta" "text" $textGateDeltaSummary.MaxBaselineDelta $TextBaselineTolerance "points"
    New-GateResult "text-x-position-max-delta" "text" $textGateDeltaSummary.MaxXDelta $TextXPositionTolerance "points"
    New-GateResult "text-font-size-max-delta" "text" $textGateDeltaSummary.MaxFontSizeDelta $TextFontSizeTolerance "points"
    New-GateResult "text-character-spacing-max-delta" "text" $textGateDeltaSummary.MaxCharacterSpacingDelta $TextCharacterSpacingTolerance "text-units"
    New-GateResult "text-character-spacing-gap-total-max-delta" "text" $textGateDeltaSummary.MaxCharacterSpacingGapTotalDelta $TextStateSpacingTolerance "points"
    New-GateResult "text-adjustment-total-max-delta" "text" $textGateDeltaSummary.MaxAdjustmentTotalDelta $TextStateSpacingTolerance "points"
    New-GateResult "text-net-spacing-gap-total-max-delta" "text" $textGateDeltaSummary.MaxNetSpacingGapTotalDelta $TextStateSpacingTolerance "points"
    New-GateResult "text-net-average-character-spacing-max-delta" "text" $textGateDeltaSummary.MaxNetAverageCharacterSpacingDelta $TextStateSpacingTolerance "points"
    New-GateResult "text-natural-width-max-delta" "text" $textGateDeltaSummary.MaxNaturalWidthDelta $GlyphAdvanceTolerance "points"
    New-GateResult "text-emitted-advance-max-delta" "text" $textGateDeltaSummary.MaxEmittedAdvanceDelta $GlyphAdvanceTolerance "points"
    New-GateResult "text-glyph-advance-max-delta" "text" $textGateDeltaSummary.MaxGlyphAdvanceDelta $GlyphAdvanceTolerance "points"
    New-GateResult "text-operation-comparison-exit-code" "text" $textComparisonExitCode 0 "exit-code"
    New-GateResult "graphics-visible-operation-comparison-exit-code" "graphics" $graphicsVisibleComparisonExitCode 0 "exit-code"
    New-GateResult "table-grid-delta-count" "tables" $tableGridDeltaCountForGate 0 "rects"
    New-GateResult "table-grid-max-bounds-delta" "tables" $tableGridMaxBoundsDeltaForGate $TableGridBoundsTolerance "points"
    New-GateResult "connector-lane-endpoint-delta-count" "markup-geometry" $connectorLaneEndpointDeltaSummary.DeltaCount 0 "segments"
    New-GateResult "connector-lane-endpoint-max-delta" "markup-geometry" $connectorLaneEndpointDeltaSummary.MaxEndpointDelta $ConnectorBoundsTolerance "points"
    New-GateResult "connector-side-inconsistent-count" "markup-geometry" $markupGeometryCounterSummary.SideInconsistentConnectorPlacementCount 0 "placements"
    New-GateResult "annotation-delta-count" "annotations" $annotationDeltaSummary.DeltaCount $MaxAnnotationDeltaCount "rects"
    New-GateResult "annotation-target-delta-count" "annotations" $annotationDeltaSummary.TargetDeltaCount 0 "targets"
    New-GateResult "annotation-max-bounds-delta" "annotations" $annotationDeltaSummary.MaxBoundsDelta $AnnotationBoundsTolerance "points"
    New-GateResult "balloon-geometry-delta-count" "balloons" $balloonGraphicDeltaSummary.DeltaCount $MaxBalloonGeometryDeltaCount "rects"
    New-GateResult "balloon-max-bounds-delta" "balloons" $balloonGraphicDeltaSummary.MaxBoundsDelta $BalloonBoundsTolerance "points"
    New-GateResult "raster-visual-delta-count" "raster" $rasterVisualDeltaCount 0 "pages"
    New-GateResult "raster-max-mean-absolute-error" "raster" $maxVisualMeanAbsoluteError $MaxMeanAbsoluteError "rgb-levels"
    New-GateResult "raster-max-changed-pixel-ratio-threshold16" "raster" $maxVisualChangedPixelRatioAtThreshold16 $MaxChangedPixelRatioAtThreshold16 "ratio"
    New-GateResult "raster-region-delta-count" "raster" $rasterRegionDeltaCount 0 "regions"
    New-GateResult "raster-region-dimension-mismatch-count" "raster" $rasterRegionDimensionMismatchCount 0 "regions"
    New-GateResult "raster-region-max-mean-absolute-error" "raster" $maxRasterRegionMeanAbsoluteError $MaxMeanAbsoluteError "rgb-levels"
    New-GateResult "raster-region-max-changed-pixel-ratio-threshold16" "raster" $maxRasterRegionChangedPixelRatioAtThreshold16 $MaxChangedPixelRatioAtThreshold16 "ratio"
)
Write-JsonFile (Join-Path $comparisonDir "gate-summary.json") $gateSummary
$summary = [ordered]@{
    CaseId = $resolvedCaseId
    RunId = $runId
    Input = if ($PrivateSafeSummary) { $null } else { $inputFull }
    InputSha256 = (Get-FileHash -LiteralPath $inputFull -Algorithm SHA256).Hash.ToLowerInvariant()
    MarkupMode = $docxMarkup
    MarkupGeometry = $docxMarkupGeometry
    ReferencePdf = $referencePdfPath
    CandidatePdf = $candidatePdf
    Diagnostics = $diagnostics
    ReferencePdfInspect = $referencePdfInspect
    CandidatePdfInspect = $candidatePdfInspect
    CandidateDocxInspect = $candidateDocxInspect
    ReferencePageCount = $referencePageCount
    CandidatePageCount = $candidatePageCount
    PageCountDelta = [Math]::Abs($candidatePageCount - $referencePageCount)
    RasterDiff = if ($SkipRasterDiff) { $null } else { $visualDiffDir }
    VisualDeltaCount = $visualDeltaCount
    RasterRegionDeltaCount = $rasterRegionDeltaCount
    RasterRegionDimensionMismatchCount = $rasterRegionDimensionMismatchCount
    MaxRasterRegionMeanAbsoluteError = $maxRasterRegionMeanAbsoluteError
    MaxRasterRegionChangedPixelRatioAtThreshold16 = $maxRasterRegionChangedPixelRatioAtThreshold16
    PageFlowDeltas = $pageFlowDeltaSummary
    RasterRegionMetricCount = $rasterRegionMetrics.Count
    RasterRegionDeltas = $rasterRegionSummary
    PixelDiffTriage = $pixelDiffTriage.Summary
    MaxMeanAbsoluteError = $MaxMeanAbsoluteError
    MaxChangedPixelRatioAtThreshold16 = $MaxChangedPixelRatioAtThreshold16
    TextComparisonExitCode = $textComparisonExitCode
    GraphicsComparisonExitCode = $graphicsComparisonExitCode
    GraphicsVisibleComparisonExitCode = $graphicsVisibleComparisonExitCode
    ReferenceGraphicsOperationCount = $referenceGraphicsOperationItems.Count
    CandidateGraphicsOperationCount = $candidateGraphicsOperationItems.Count
    GraphicsOperationCountDelta = [Math]::Abs($candidateGraphicsOperationItems.Count - $referenceGraphicsOperationItems.Count)
    ReferenceVisibleGraphicsOperationCount = $referenceVisibleGraphicsOperationItems.Count
    CandidateVisibleGraphicsOperationCount = $candidateVisibleGraphicsOperationItems.Count
    VisibleGraphicsOperationCountDelta = [Math]::Abs($candidateVisibleGraphicsOperationItems.Count - $referenceVisibleGraphicsOperationItems.Count)
    ReferenceClipGraphicsOperationCount = @($referenceGraphicsOperationItems | Where-Object { $_.Kind -eq "Clip" }).Count
    CandidateClipGraphicsOperationCount = @($candidateGraphicsOperationItems | Where-Object { $_.Kind -eq "Clip" }).Count
    ClipGraphicsOperationCountDelta = [Math]::Abs(
        @($candidateGraphicsOperationItems | Where-Object { $_.Kind -eq "Clip" }).Count -
        @($referenceGraphicsOperationItems | Where-Object { $_.Kind -eq "Clip" }).Count)
    GateFailureCount = $gateSummary.FailureCount
    GateFailures = $gateSummary.Failures
    CurrentBlockingGatePhase = $gateSummary.CurrentBlockingPhase
    GatePhases = $gateSummary.Phases
    MediaBoxDelta = $mediaBoxDeltaSummary
    BodyFrameDelta = $bodyFrameDeltaSummary
    MarkupMarginDelta = $markupMarginDeltaSummary
    MarkupGeometryCounters = $markupGeometryCounterSummary
    LayoutBodyFrameRectCount = $candidateLayoutBodyFrameRects.Count
    LayoutColumnFrameRectCount = $candidateLayoutColumnFrameRects.Count
    LayoutMarkupLaneRectCount = $candidateLayoutMarkupLaneRects.Count
    ReferenceMarkupLaneRectCount = $referenceMarkupLaneRects.Count
    BodyFrameLaneEdgeDelta = $bodyFrameLaneEdgeDeltaSummary
    MarkupLaneHorizontalDelta = $markupLaneHorizontalDeltaSummary
    TextGateDeltas = $textGateDeltaSummary
    TextOperationDrift = [pscustomobject]@{
        OperationPairCount = $textOperationDriftSummary.OperationPairCount
        TopXDeltaCount = @($textOperationDriftSummary.TopXDeltaPairs).Count
        TopAdjustmentTotalDeltaCount = @($textOperationDriftSummary.TopAdjustmentTotalDeltaPairs).Count
        MaxReportedXDelta = if (@($textOperationDriftSummary.TopXDeltaPairs).Count -eq 0) { $null } else { (@($textOperationDriftSummary.TopXDeltaPairs))[0].AbsXDelta }
        MaxReportedAdjustmentTotalDelta = if (@($textOperationDriftSummary.TopAdjustmentTotalDeltaPairs).Count -eq 0) { $null } else { (@($textOperationDriftSummary.TopAdjustmentTotalDeltaPairs))[0].AbsAdjustmentTotalDelta }
    }
    TableGridDelta = $tableGridDeltaSummary
    TableGridGateSourceTableCount = $candidateLayoutTableCount
    ConnectorDelta = $connectorDeltaSummary
    ConnectorEndpointDelta = $connectorEndpointDeltaSummary
    ConnectorLaneEndpointDelta = $connectorLaneEndpointDeltaSummary
    ConnectorBodyAnchorEndpointDelta = $connectorBodyAnchorEndpointDeltaSummary
    BalloonGraphicDelta = $balloonGraphicDeltaSummary
    BroadBalloonGraphicDelta = $broadBalloonGraphicDeltaSummary
    ReferenceAnnotationCount = $referenceAnnotations.Count
    CandidateAnnotationCount = $candidateAnnotations.Count
    AnnotationDeltaCount = $annotationDeltaCount
    AnnotationDelta = $annotationDeltaSummary
    AnnotationTargetDeltaCount = $annotationDeltaSummary.TargetDeltaCount
    AnnotationTargetKindDeltaCount = $annotationDeltaSummary.TargetKindDeltaCount
    AnnotationTargetSha256DeltaCount = $annotationDeltaSummary.TargetSha256DeltaCount
    CandidateBalloonCount = $candidateBalloons.Count
    ReferenceBalloonGraphicCandidateCount = $referenceBalloonBodyGraphics.Count
    CandidateBalloonGraphicCandidateCount = $candidateBalloonBodyGraphics.Count
    ReferenceBroadBalloonGraphicCandidateCount = $referenceBalloonGraphics.Count
    CandidateBroadBalloonGraphicCandidateCount = $candidateBalloonGraphics.Count
    BalloonGeometryDeltaCount = $balloonDeltaCount
    BalloonGraphicGeometryDeltaCount = $balloonGraphicDeltaSummary.DeltaCount
    BalloonClassification = $balloonClassificationSummary
    FontResources = $fontResourceSummary
    WorstPages = @($pageDeltaSummary | Select-Object -First 10)
    RegionDeltas = $regionDeltaSummary
}
Write-JsonFile (Join-Path $runRoot "summary.json") $summary

if ($FailOnDeltas) {
    if ($gateSummary.FailureCount -ne 0) {
        throw "Cached DOCX markup comparison found deltas. See $runRoot"
    }
}

Write-Host "Cached DOCX markup comparison artifacts: $runRoot"
