function Read-VisualCaseManifest([string] $Path) {
    $bytes = [IO.File]::ReadAllBytes($Path)
    $reader = [IO.StreamReader]::new([IO.MemoryStream]::new($bytes, $false))
    try { $manifest = $reader.ReadToEnd() | ConvertFrom-Json }
    finally { $reader.Dispose() }
    [pscustomobject]@{
        Path = $Path
        Manifest = $manifest
        Sha256 = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes))
    }
}

function Get-ManifestDouble($object, [string] $name, $defaultValue) {
    if ($null -ne $object -and $object.PSObject.Properties.Name -contains $name -and $null -ne $object.$name) {
        return [double]$object.$name
    }
    return $defaultValue
}

function Get-VisualSupportThresholds($CaseSupport, $FamilySupport) {
    [pscustomobject]@{
        StructuralSimilarity = Get-ManifestDouble $CaseSupport "minStructuralSimilarity" (Get-ManifestDouble $FamilySupport "minStructuralSimilarity" 0.995)
        ForegroundColorHistogramCorrelation = Get-ManifestDouble $CaseSupport "minForegroundColorHistogramCorrelation" (Get-ManifestDouble $FamilySupport "minForegroundColorHistogramCorrelation" 0.98)
        ForegroundRecall = Get-ManifestDouble $CaseSupport "minForegroundRecall" (Get-ManifestDouble $FamilySupport "minForegroundRecall" $null)
    }
}

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
