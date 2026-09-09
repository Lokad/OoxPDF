# Summarizes PDF text page deltas across runs.

param(
    [Parameter(Mandatory = $true)]
    [string] $RunDirectory,

    [string] $OutputJson
)

$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "JsonArray.ps1")

function Rounded($Value, [int] $Digits) {
    if ($null -eq $Value -or [string]$Value -eq "") {
        return $null
    }

    return [Math]::Round([double]$Value, $Digits)
}

. (Join-Path $PSScriptRoot "CompareMath.ps1")


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
    $out = Join-Path "artifacts\pdf-text-pages" $safeId
    if (-not (Test-Path -LiteralPath $out)) {
        New-Item -ItemType Directory -Path $out | Out-Null
    }

    $json = Join-Path $out "text-operations.json"
    if (-not (Test-Path -LiteralPath $json)) {
        pwsh tools\InspectPdf.ps1 -InputPdf $pdf -OutputDirectory $out -TextOnly | Out-Null
    }

    return $json
}

function Select-Coordinate($Operation, [string] $Name) {
    $effectiveName = "Effective$Name"
    if ($Operation.PSObject.Properties.Name -contains $effectiveName -and $null -ne $Operation.$effectiveName) {
        return [double]$Operation.$effectiveName
    }

    return [double]$Operation.$Name
}

function Summarize-Page($Operations) {
    $ops = @($Operations)
    $nonzeroTc = @($ops | Where-Object { [Math]::Abs([double]$_.CharacterSpacing) -gt 0.001d })
    $xs = @($ops | ForEach-Object { Select-Coordinate $_ "X" })
    $ys = @($ops | ForEach-Object { Select-Coordinate $_ "Y" })

    return [pscustomobject]@{
        OperationCount = $ops.Count
        NonzeroTcCount = $nonzeroTc.Count
        XMin = if ($xs.Count -eq 0) { $null } else { Rounded ($xs | Measure-Object -Minimum).Minimum 3 }
        XMax = if ($xs.Count -eq 0) { $null } else { Rounded ($xs | Measure-Object -Maximum).Maximum 3 }
        YMin = if ($ys.Count -eq 0) { $null } else { Rounded ($ys | Measure-Object -Minimum).Minimum 3 }
        YMax = if ($ys.Count -eq 0) { $null } else { Rounded ($ys | Measure-Object -Maximum).Maximum 3 }
        TcBuckets = @(Group-Count $ops { param($op) RoundedKey $op.CharacterSpacing 6 })
        FontSizeBuckets = @(Group-Count $ops { param($op) RoundedKey $op.FontSize 3 })
        TextClassBuckets = @(Group-Count $ops { param($op) TextClass $op })
        TextClassByTc = @(Group-Count $ops {
            param($op)
            (TextClass $op) + "|tc=" + (RoundedKey $op.CharacterSpacing 6)
        })
        ChunkCountBuckets = @(Group-Count $ops { param($op) RoundedKey $op.TextChunkCount 0 })
        AdjustmentCountBuckets = @(Group-Count $ops { param($op) RoundedKey $op.AdjustmentCount 0 })
    }
}

function Get-PageSummary($Operations, [int] $Page) {
    Summarize-Page @($Operations | Where-Object { [int]$_.PageNumber -eq $Page })
}

$resolvedRun = (Resolve-Path -LiteralPath $RunDirectory).Path
$referenceOps = Read-JsonArray (Ensure-TextOperations $resolvedRun "reference")
$candidateOps = Read-JsonArray (Ensure-TextOperations $resolvedRun "candidate")
$metricsPath = Join-Path $resolvedRun "comparison\metrics.json"
$metrics = if (Test-Path -LiteralPath $metricsPath) { Read-JsonArray $metricsPath } else { @() }

$pageNumbers = @(
    $referenceOps | ForEach-Object { [int]$_.PageNumber }
    $candidateOps | ForEach-Object { [int]$_.PageNumber }
    $metrics | ForEach-Object { [int]$_.Page }
) | Sort-Object -Unique

$rows = foreach ($page in $pageNumbers) {
    $ref = Get-PageSummary $referenceOps $page
    $cand = Get-PageSummary $candidateOps $page
    $pageMetrics = @($metrics | Where-Object { [int]$_.Page -eq $page } | Select-Object -First 1)

    [pscustomobject]@{
        Page = $page
        MeanAbsoluteError = if ($pageMetrics.Count -eq 0) { $null } else { Rounded $pageMetrics[0].MeanAbsoluteError 6 }
        Changed16 = if ($pageMetrics.Count -eq 0) { $null } else { Rounded $pageMetrics[0].ChangedPixelRatioAtThreshold16 6 }
        StructuralSimilarity = if ($pageMetrics.Count -eq 0) { $null } else { Rounded $pageMetrics[0].StructuralSimilarity 6 }
        Reference = $ref
        Candidate = $cand
        Delta = [pscustomobject]@{
            OperationCount = $cand.OperationCount - $ref.OperationCount
            NonzeroTcCount = $cand.NonzeroTcCount - $ref.NonzeroTcCount
            XMin = if ($null -eq $ref.XMin -or $null -eq $cand.XMin) { $null } else { Rounded ($cand.XMin - $ref.XMin) 3 }
            XMax = if ($null -eq $ref.XMax -or $null -eq $cand.XMax) { $null } else { Rounded ($cand.XMax - $ref.XMax) 3 }
            YMin = if ($null -eq $ref.YMin -or $null -eq $cand.YMin) { $null } else { Rounded ($cand.YMin - $ref.YMin) 3 }
            YMax = if ($null -eq $ref.YMax -or $null -eq $cand.YMax) { $null } else { Rounded ($cand.YMax - $ref.YMax) 3 }
        }
    }
}

if ($OutputJson) {
    $resolvedOutput = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputJson)
    $parent = Split-Path -Parent $resolvedOutput
    if ($parent -and -not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Path $parent | Out-Null
    }

    $rows | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $resolvedOutput -Encoding utf8
}
else {
    $rows | ConvertTo-Json -Depth 20
}
