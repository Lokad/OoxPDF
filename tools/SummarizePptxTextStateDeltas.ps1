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
    $nonzeroRefRows = @($matchedRows | Where-Object { HasNonzeroRefTc $_ })
    $zeroRefRows = @($matchedRows | Where-Object { -not (HasNonzeroRefTc $_) })

    [pscustomobject]@{
        Case = [IO.Path]::GetFileNameWithoutExtension($resolved)
        Source = $resolved
        Total = $rows.Count
        Matched = $matchedRows.Count
        Missing = @($rows | Where-Object { (OptionalValue $_ "Status") -eq "missing" }).Count
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
        NonzeroReferenceByCandidateAdjustments = Group-Count $nonzeroRefRows {
            param($row)
            "avg=" + (RoundedKey (OptionalValue $row "CandInterGlyphAdjustmentAverage") 6) +
                "|count=" + (RoundedKey (OptionalValue $row "CandInterGlyphAdjustmentCount") 0) +
                "|first=" + (RoundedKey (OptionalValue $row "CandFirstAdjustmentAfterOrigin") 6)
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
    Select-Object Case, Total, Matched, Missing, NonzeroReferenceTc, ZeroReferenceTc |
    Format-Table -AutoSize
