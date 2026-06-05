param(
    [Parameter(Mandatory = $true)]
    [string[]] $CompareJson,

    [string] $OutputJson
)

$ErrorActionPreference = "Stop"

function Expand-PathList([string[]] $Values) {
    $expanded = New-Object System.Collections.Generic.List[string]
    foreach ($value in $Values) {
        if ([string]::IsNullOrWhiteSpace($value)) {
            continue
        }

        foreach ($part in ($value -split "[,;]")) {
            if (-not [string]::IsNullOrWhiteSpace($part)) {
                $expanded.Add($part.Trim())
            }
        }
    }

    return ,$expanded.ToArray()
}

function Read-JsonArray([string] $Path) {
    $items = Get-Content -Raw -LiteralPath (Resolve-Path -LiteralPath $Path).Path | ConvertFrom-Json
    if ($null -eq $items) {
        return ,@()
    }

    if ($items -is [array]) {
        return $items
    }

    return ,@($items)
}

function OptionalValue($Item, [string] $Name) {
    if ($null -eq $Item) {
        return $null
    }

    $property = $Item.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $null
    }

    return $property.Value
}

function OptionalDouble($Item, [string] $Name) {
    $value = OptionalValue $Item $Name
    if ($null -eq $value -or [string]$value -eq "") {
        return $null
    }

    return [double]$value
}

function RoundedKey($Value, [int] $Digits) {
    if ($null -eq $Value -or [string]$Value -eq "") {
        return "(missing)"
    }

    return ([Math]::Round([double]$Value, $Digits)).ToString("0.######", [Globalization.CultureInfo]::InvariantCulture)
}

function StringKey($Value) {
    if ($null -eq $Value -or [string]$Value -eq "") {
        return "(missing)"
    }

    return [string]$Value
}

function BoolKey($Value) {
    if ($null -eq $Value -or [string]$Value -eq "") {
        return "(missing)"
    }

    return [string][bool]$Value
}

function Group-Count($Items, [scriptblock] $KeySelector) {
    $groups = @{}
    foreach ($item in $Items) {
        $key = & $KeySelector $item
        if ($null -eq $key -or [string]$key -eq "") {
            $key = "(missing)"
        }

        $key = [string]$key
        if ($groups.ContainsKey($key)) {
            $groups[$key]++
        }
        else {
            $groups[$key] = 1
        }
    }

    return @(
        foreach ($key in ($groups.Keys | Sort-Object)) {
            [pscustomobject]@{
                Key = $key
                Count = $groups[$key]
            }
        }
    )
}

function Group-Stats($Items, [scriptblock] $KeySelector, [string] $ValueName) {
    $groups = @{}
    foreach ($item in $Items) {
        $key = & $KeySelector $item
        if ($null -eq $key -or [string]$key -eq "") {
            $key = "(missing)"
        }

        $value = OptionalDouble $item $ValueName
        if ($null -eq $value) {
            continue
        }

        $key = [string]$key
        if (-not $groups.ContainsKey($key)) {
            $groups[$key] = New-Object System.Collections.Generic.List[double]
        }

        $groups[$key].Add($value)
    }

    return @(
        foreach ($key in ($groups.Keys | Sort-Object)) {
            $values = @($groups[$key])
            [pscustomobject]@{
                Key = $key
                Count = $values.Count
                Average = if ($values.Count -eq 0) { $null } else { [Math]::Round(($values | Measure-Object -Average).Average, 6) }
                Minimum = if ($values.Count -eq 0) { $null } else { [Math]::Round(($values | Measure-Object -Minimum).Minimum, 6) }
                Maximum = if ($values.Count -eq 0) { $null } else { [Math]::Round(($values | Measure-Object -Maximum).Maximum, 6) }
            }
        }
    )
}

function IsReliablePositionMatch($Row) {
    if ((OptionalValue $Row "Status") -eq "missing") {
        return $false
    }

    $dx = OptionalDouble $Row "DeltaX"
    $dy = OptionalDouble $Row "DeltaBaselineY"
    return $null -ne $dx -and
        $null -ne $dy -and
        [Math]::Abs($dx) -lt 2d -and
        [Math]::Abs($dy) -lt 2d
}

function HasNonzeroRefTc($Row) {
    $tc = OptionalDouble $Row "RefCharacterSpacing"
    return $null -ne $tc -and [Math]::Abs($tc) -gt 0.001d
}

$paths = Expand-PathList $CompareJson
if ($paths.Count -eq 0) {
    throw "No compare JSON paths were provided."
}

$summaries = foreach ($path in $paths) {
    $resolved = (Resolve-Path -LiteralPath $path).Path
    $rows = @(Read-JsonArray $resolved)
    $matchedRows = @($rows | Where-Object { (OptionalValue $_ "Status") -ne "missing" })
    $reliableRows = @($matchedRows | Where-Object { IsReliablePositionMatch $_ })
    $nonzeroRefRows = @($matchedRows | Where-Object { HasNonzeroRefTc $_ })
    $zeroRefRows = @($matchedRows | Where-Object { -not (HasNonzeroRefTc $_) })

    [pscustomobject]@{
        Case = [IO.Path]::GetFileNameWithoutExtension($resolved)
        Source = $resolved
        Total = $rows.Count
        Matched = $matchedRows.Count
        Missing = @($rows | Where-Object { (OptionalValue $_ "Status") -eq "missing" }).Count
        ReliablePositionMatches = $reliableRows.Count
        NonzeroReferenceTc = $nonzeroRefRows.Count
        ZeroReferenceTc = $zeroRefRows.Count
        ReferenceTcBuckets = Group-Count $matchedRows { param($row) RoundedKey (OptionalValue $row "RefCharacterSpacing") 6 }
        CandidateTcBuckets = Group-Count $matchedRows { param($row) RoundedKey (OptionalValue $row "CandPdfCharacterSpacing") 6 }
        NonzeroReferenceByTable = Group-Count $nonzeroRefRows { param($row) "table=" + [string]((OptionalValue $row "CandTableRowIndex") -ne $null) }
        NonzeroReferenceByFontSizeAndTc = Group-Count $nonzeroRefRows { param($row) (RoundedKey (OptionalValue $row "RefFontSize") 3) + "|tc=" + (RoundedKey (OptionalValue $row "RefCharacterSpacing") 6) }
        NonzeroReferenceByCandidateShape = Group-Count $nonzeroRefRows {
            param($row)
            "table=" + [string]((OptionalValue $row "CandTableRowIndex") -ne $null) +
                "|bold=" + (BoolKey (OptionalValue $row "CandBold")) +
                "|italic=" + (BoolKey (OptionalValue $row "CandItalic")) +
                "|lineSpans=" + (RoundedKey (OptionalValue $row "CandLineSpanCount") 0) +
                "|run=" + (RoundedKey (OptionalValue $row "CandSourceRunIndex") 0)
        }
        NonzeroReferenceByCandidateLetterCase = Group-Count $nonzeroRefRows {
            param($row)
            "upper=" + (RoundedKey (OptionalValue $row "CandUppercaseLetterCount") 0) +
                "|lower=" + (RoundedKey (OptionalValue $row "CandLowercaseLetterCount") 0) +
                "|title=" + (RoundedKey (OptionalValue $row "CandTitlecaseLetterCount") 0) +
                "|letters=" + (RoundedKey (OptionalValue $row "CandLetterCount") 0) +
                "|spaces=" + (RoundedKey (OptionalValue $row "CandSpaceCount") 0) +
                "|tc=" + (RoundedKey (OptionalValue $row "RefCharacterSpacing") 6)
        }
        NonzeroReferenceByCandidateAdjustments = Group-Count $nonzeroRefRows {
            param($row)
            "avg=" + (RoundedKey (OptionalValue $row "CandInterGlyphAdjustmentAverage") 6) +
                "|count=" + (RoundedKey (OptionalValue $row "CandInterGlyphAdjustmentCount") 0) +
                "|first=" + (RoundedKey (OptionalValue $row "CandFirstAdjustmentAfterOrigin") 6)
        }
        ReliableBaselineDeltaByFrameShape = Group-Stats $reliableRows {
            param($row)
            "frame=" + (RoundedKey (OptionalValue $row "CandFrameIndex") 0) +
                "|tableRow=" + (RoundedKey (OptionalValue $row "CandTableRowIndex") 0) +
                "|font=" + (RoundedKey (OptionalValue $row "CandPdfFontSize") 3) +
                "|cols=" + (RoundedKey (OptionalValue $row "CandFrameColumnCount") 0) +
                "|autofit=" + (StringKey (OptionalValue $row "CandFrameAutofitMode")) +
                "|wrap=" + (StringKey (OptionalValue $row "CandFrameWrapMode"))
        } "DeltaBaselineY"
        ReliableBaselineDeltaByTableCell = Group-Stats (@($reliableRows | Where-Object { $null -ne (OptionalValue $_ "CandTableRowIndex") })) {
            param($row)
            "row=" + (RoundedKey (OptionalValue $row "CandTableRowIndex") 0) +
                "|col=" + (RoundedKey (OptionalValue $row "CandTableColumnIndex") 0) +
                "|font=" + (RoundedKey (OptionalValue $row "CandPdfFontSize") 3) +
                "|line=" + (RoundedKey (OptionalValue $row "CandLineIndex") 0) +
                "|spanCount=" + (RoundedKey (OptionalValue $row "CandLineSpanCount") 0)
        } "DeltaBaselineY"
        ReliableXDeltaByTextState = Group-Stats $reliableRows {
            param($row)
            "refTc=" + (RoundedKey (OptionalValue $row "RefCharacterSpacing") 6) +
                "|candTc=" + (RoundedKey (OptionalValue $row "CandPdfCharacterSpacing") 6) +
                "|font=" + (RoundedKey (OptionalValue $row "CandPdfFontSize") 3) +
                "|glyphs=" + (RoundedKey (OptionalValue $row "CandGlyphCount") 0) +
                "|spaces=" + (RoundedKey (OptionalValue $row "CandSpaceCount") 0)
        } "DeltaX"
        MissingReferenceByTextState = Group-Count (@($rows | Where-Object { (OptionalValue $_ "Status") -eq "missing" })) {
            param($row)
            "font=" + (RoundedKey (OptionalValue $row "RefFontSize") 3) +
                "|tc=" + (RoundedKey (OptionalValue $row "RefCharacterSpacing") 6)
        }
        ZeroReferenceWithCandidateResiduals = Group-Count (@($zeroRefRows | Where-Object {
            $avg = OptionalDouble $_ "CandInterGlyphAdjustmentAverage"
            $null -ne $avg -and [Math]::Abs($avg) -gt 0.001d
        })) {
            param($row)
            "table=" + [string]((OptionalValue $row "CandTableRowIndex") -ne $null) +
                "|bold=" + (BoolKey (OptionalValue $row "CandBold")) +
                "|italic=" + (BoolKey (OptionalValue $row "CandItalic")) +
                "|lineSpans=" + (RoundedKey (OptionalValue $row "CandLineSpanCount") 0) +
                "|run=" + (RoundedKey (OptionalValue $row "CandSourceRunIndex") 0) +
                "|avg=" + (RoundedKey (OptionalValue $row "CandInterGlyphAdjustmentAverage") 6)
        }
        ZeroReferenceByCandidateLetterCaseWithResiduals = Group-Count (@($zeroRefRows | Where-Object {
            $avg = OptionalDouble $_ "CandInterGlyphAdjustmentAverage"
            $null -ne $avg -and [Math]::Abs($avg) -gt 0.001d
        })) {
            param($row)
            "upper=" + (RoundedKey (OptionalValue $row "CandUppercaseLetterCount") 0) +
                "|lower=" + (RoundedKey (OptionalValue $row "CandLowercaseLetterCount") 0) +
                "|title=" + (RoundedKey (OptionalValue $row "CandTitlecaseLetterCount") 0) +
                "|letters=" + (RoundedKey (OptionalValue $row "CandLetterCount") 0) +
                "|spaces=" + (RoundedKey (OptionalValue $row "CandSpaceCount") 0) +
                "|avg=" + (RoundedKey (OptionalValue $row "CandInterGlyphAdjustmentAverage") 6)
        }
    }
}

if (-not [string]::IsNullOrWhiteSpace($OutputJson)) {
    $outputFull = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputJson)
    $outputDirectory = Split-Path -Parent $outputFull
    if (-not [string]::IsNullOrWhiteSpace($outputDirectory)) {
        New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null
    }

    $summaries | ConvertTo-Json -Depth 10 | Set-Content -Encoding UTF8 -LiteralPath $outputFull
}

$summaries |
    Select-Object Case, Total, Matched, Missing, ReliablePositionMatches, NonzeroReferenceTc, ZeroReferenceTc |
    Format-Table -AutoSize
