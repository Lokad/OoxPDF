param(
    [Parameter(Mandatory = $true)]
    [string] $InputJson,

    [string[]] $Kinds = @("Stroke", "Fill", "FillStroke", "Clip"),

    [string[]] $Operators = @(),

    [int] $PageNumber = 0,

    [double] $BoundsPrecision = 0.001,

    [int] $Top = 0
)

$ErrorActionPreference = "Stop"

function Read-JsonArray([string] $Path) {
    $items = Get-Content -Raw -LiteralPath (Resolve-Path -LiteralPath $Path).Path | ConvertFrom-Json
    if ($null -eq $items) {
        return @()
    }

    if ($items -is [array]) {
        return @($items)
    }

    return @($items)
}

function OperationSourceOperator($Operation) {
    if ($null -eq $Operation) {
        return ""
    }

    if ($null -ne $Operation.SourceOperator -and [string]$Operation.SourceOperator -ne "") {
        return [string]$Operation.SourceOperator
    }

    if ($null -ne $Operation.Operator -and [string]$Operation.Operator -ne "") {
        return [string]$Operation.Operator
    }

    return ""
}

function IntValue($Value) {
    if ($null -eq $Value) {
        return 0
    }

    return [int]$Value
}

function Round-Bounds([double] $Value) {
    if ($BoundsPrecision -le 0d) {
        return $Value
    }

    $rounded = [Math]::Round([Math]::Round($Value / $BoundsPrecision) * $BoundsPrecision, 6)
    if ([Math]::Abs($rounded) -lt ($BoundsPrecision / 2d)) {
        return 0d
    }

    return $rounded
}

$selectedKinds = @($Kinds | ForEach-Object { [string]$_ -split "," } | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })
$selectedOperators = @($Operators | ForEach-Object { [string]$_ -split "," } | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })

$operations = @(Read-JsonArray $InputJson | Where-Object { $selectedKinds -contains $_.Kind })
if ($selectedOperators.Count -gt 0) {
    $operations = @($operations | Where-Object { $selectedOperators -contains (OperationSourceOperator $_) })
}

if ($PageNumber -gt 0) {
    $operations = @($operations | Where-Object { [int]$_.PageNumber -eq $PageNumber })
}

$buckets = @{}
foreach ($operation in $operations) {
    $operator = OperationSourceOperator $operation
    $minX = Round-Bounds ([double]$operation.MinX)
    $minY = Round-Bounds ([double]$operation.MinY)
    $maxX = Round-Bounds ([double]$operation.MaxX)
    $maxY = Round-Bounds ([double]$operation.MaxY)
    $key = @(
        [int]$operation.PageNumber,
        [string]$operation.Kind,
        $operator,
        (IntValue $operation.SegmentCount),
        (IntValue $operation.MoveCount),
        (IntValue $operation.LineCount),
        (IntValue $operation.CurveCount),
        (IntValue $operation.CloseCount),
        $minX,
        $minY,
        $maxX,
        $maxY
    ) -join "|"

    if (-not $buckets.ContainsKey($key)) {
        $buckets[$key] = [pscustomobject]@{
            PageNumber = [int]$operation.PageNumber
            Kind = [string]$operation.Kind
            Operator = $operator
            SegmentCount = IntValue $operation.SegmentCount
            MoveCount = IntValue $operation.MoveCount
            LineCount = IntValue $operation.LineCount
            CurveCount = IntValue $operation.CurveCount
            CloseCount = IntValue $operation.CloseCount
            MinX = $minX
            MinY = $minY
            MaxX = $maxX
            MaxY = $maxY
            Count = 0
        }
    }

    $buckets[$key].Count++
}

$rows = @($buckets.Values | Sort-Object PageNumber, Kind, Operator, MinY, MinX, MaxY, MaxX, SegmentCount, CloseCount)
if ($Top -gt 0) {
    $rows = @($rows | Sort-Object Count -Descending | Select-Object -First $Top)
}

[pscustomobject]@{
    InputJson = (Resolve-Path -LiteralPath $InputJson).Path
    PageNumber = $PageNumber
    OperationCount = $operations.Count
    BucketCount = $rows.Count
    Buckets = $rows
} | ConvertTo-Json -Depth 5
