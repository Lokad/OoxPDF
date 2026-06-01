param(
    [Parameter(Mandatory = $true)]
    [string[]] $RunDirectory,

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
        return ,$items
    }

    return ,@($items)
}

function RoundedKey($Value, [int] $Digits) {
    if ($null -eq $Value -or [string]$Value -eq "") {
        return "(missing)"
    }

    return ([Math]::Round([double]$Value, $Digits)).ToString("0.######", [Globalization.CultureInfo]::InvariantCulture)
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

function TextLength($Operation) {
    if ($null -eq $Operation.DecodedText) {
        return 0
    }

    return ([string]$Operation.DecodedText).Length
}

function TextClass($Operation) {
    if ($null -eq $Operation.DecodedText) {
        return "(missing)"
    }

    $text = [string]$Operation.DecodedText
    if ($text.Length -eq 0) {
        return "empty"
    }

    if ($text -match '^\s+$') {
        return "whitespace"
    }

    $letters = [regex]::IsMatch($text, '\p{L}')
    $digits = [regex]::IsMatch($text, '\p{Nd}')
    $other = [regex]::IsMatch($text, '[^\p{L}\p{Nd}\s]')
    if ($letters -and -not $digits -and -not $other) {
        return "letters"
    }

    if ($digits -and -not $letters -and -not $other) {
        return "digits"
    }

    if ($letters -and $digits -and -not $other) {
        return "alphanumeric"
    }

    if ($other -and -not $letters -and -not $digits) {
        return "punctuation"
    }

    return "mixed"
}

function Ensure-TextOperations([string] $Run, [string] $Side) {
    $pdf = if ($Side -eq "reference") {
        Join-Path $Run "reference\reference.pdf"
    }
    else {
        Join-Path $Run "candidate\output.pdf"
    }

    if (-not (Test-Path -LiteralPath $pdf)) {
        throw "Missing $Side PDF for run '$Run': $pdf"
    }

    $runInfo = Get-Item -LiteralPath $Run
    $caseId = Split-Path -Leaf (Split-Path -Parent $runInfo.FullName)
    $runId = Split-Path -Leaf $runInfo.FullName
    $safeId = ($caseId + "-" + $runId + "-" + $Side) -replace "[^A-Za-z0-9_.-]", "_"
    $out = Join-Path "artifacts\docx-text-state" $safeId
    if (-not (Test-Path -LiteralPath $out)) {
        New-Item -ItemType Directory -Path $out | Out-Null
    }

    $json = Join-Path $out "text-operations.json"
    if (-not (Test-Path -LiteralPath $json)) {
        pwsh tools\InspectPdf.ps1 -InputPdf $pdf -OutputDirectory $out -TextOnly | Out-Null
    }

    return $json
}

function Summarize-Operations($Operations) {
    $nonzeroTc = @($Operations | Where-Object { [Math]::Abs([double]$_.CharacterSpacing) -gt 0.001d })
    return [pscustomobject]@{
        OperationCount = $Operations.Count
        NonzeroTcCount = $nonzeroTc.Count
        OperatorBuckets = @(Group-Count $Operations { param($op) [string]$op.Operator })
        TcBuckets = @(Group-Count $Operations { param($op) RoundedKey $op.CharacterSpacing 6 })
        FontSizeBuckets = @(Group-Count $Operations { param($op) RoundedKey $op.FontSize 3 })
        FontSizeByTc = @(Group-Count $Operations {
            param($op)
            (RoundedKey $op.FontSize 3) + "|tc=" + (RoundedKey $op.CharacterSpacing 6)
        })
        TextLengthByTc = @(Group-Count $Operations {
            param($op)
            "len=" + (TextLength $op) + "|tc=" + (RoundedKey $op.CharacterSpacing 6)
        })
        TextClassByTc = @(Group-Count $Operations {
            param($op)
            (TextClass $op) + "|tc=" + (RoundedKey $op.CharacterSpacing 6)
        })
        NonzeroTcByFontSize = @(Group-Count $nonzeroTc {
            param($op)
            (RoundedKey $op.FontSize 3) + "|tc=" + (RoundedKey $op.CharacterSpacing 6)
        })
        NonzeroTcByTextClass = @(Group-Count $nonzeroTc {
            param($op)
            TextClass $op
        })
        NetAverageSpacingByTc = @(Group-Count $Operations {
            param($op)
            "net=" + (RoundedKey $op.NetAverageCharacterSpacing 6) + "|tc=" + (RoundedKey $op.CharacterSpacing 6)
        })
        TextChunkCountByTc = @(Group-Count $Operations {
            param($op)
            "chunks=" + (RoundedKey $op.TextChunkCount 0) + "|tc=" + (RoundedKey $op.CharacterSpacing 6)
        })
        AdjustmentCountByTc = @(Group-Count $Operations {
            param($op)
            "adj=" + (RoundedKey $op.AdjustmentCount 0) + "|tc=" + (RoundedKey $op.CharacterSpacing 6)
        })
        AverageAdjustmentByTc = @(Group-Count $Operations {
            param($op)
            "avgAdj=" + (RoundedKey $op.AverageAdjustmentPoints 6) + "|tc=" + (RoundedKey $op.CharacterSpacing 6)
        })
    }
}

$runs = Expand-PathList $RunDirectory
if ($runs.Count -eq 0) {
    throw "No run directories were provided."
}

$summaries = foreach ($run in $runs) {
    $resolvedRun = (Resolve-Path -LiteralPath $run).Path
    $referenceOps = Read-JsonArray (Ensure-TextOperations $resolvedRun "reference")
    $candidateOps = Read-JsonArray (Ensure-TextOperations $resolvedRun "candidate")
    $metricsPath = Join-Path $resolvedRun "comparison\metrics.json"
    $metrics = if (Test-Path -LiteralPath $metricsPath) {
        $pages = Read-JsonArray $metricsPath
        [pscustomobject]@{
            PageCount = $pages.Count
            MeanAbsoluteError = if ($pages.Count -eq 0) { $null } else { [Math]::Round(($pages | Measure-Object -Property MeanAbsoluteError -Average).Average, 6) }
            Changed16 = if ($pages.Count -eq 0) { $null } else { [Math]::Round(($pages | Measure-Object -Property ChangedPixelRatioAtThreshold16 -Average).Average, 6) }
            StructuralSimilarity = if ($pages.Count -eq 0) { $null } else { [Math]::Round(($pages | Measure-Object -Property StructuralSimilarity -Average).Average, 6) }
        }
    }
    else {
        $null
    }

    [pscustomobject]@{
        RunDirectory = $resolvedRun
        CaseId = Split-Path -Leaf (Split-Path -Parent $resolvedRun)
        RunId = Split-Path -Leaf $resolvedRun
        Metrics = $metrics
        Reference = Summarize-Operations $referenceOps
        Candidate = Summarize-Operations $candidateOps
    }
}

if ($OutputJson) {
    $resolvedOutput = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputJson)
    $parent = Split-Path -Parent $resolvedOutput
    if ($parent -and -not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Path $parent | Out-Null
    }

    $summaries | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $resolvedOutput -Encoding utf8
}
else {
    $summaries | ConvertTo-Json -Depth 20
}
