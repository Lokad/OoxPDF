function Test-VisualChartGraphicsGate($Expected) {
    $null -ne $Expected.maxChartGraphicsStructureBoundsDelta -or
        $null -ne $Expected.maxChartGraphicsStructureBoundsDeltaByKind
}

function Get-VisualChartGraphicsGateIssue($Expected, [string] $CaseId) {
    if ($null -eq $Expected -or (Test-VisualChartGraphicsGate $Expected)) { return }

    # These options are consumed only inside the primary chart graphics gate.
    $secondary = @($Expected.PSObject.Properties | Where-Object {
        $_.Name -match '^(chartGraphics|compareChartGraphicsStructure|maxChartGraphicsStructure)' -and
        $null -ne $_.Value -and !($_.Value -is [bool] -and !$_.Value)
    } | ForEach-Object Name)
    if (@($Expected.additionalChartGraphicsStructureComparisons | Where-Object { $null -ne $_ }).Count) {
        $secondary += 'additionalChartGraphicsStructureComparisons'
    }
    if ($secondary.Count) {
        "Visual case '$CaseId' requires maxChartGraphicsStructureBoundsDelta or maxChartGraphicsStructureBoundsDeltaByKind for $($secondary -join ', ')."
    }
}
