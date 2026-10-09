# Supplemental source-backed cell-clip triage; existing visual gates remain authoritative.
param(
    [Parameter(Mandatory = $true)] [string] $LayoutSnapshot,
    [Parameter(Mandatory = $true)] [string] $ReferenceGraphics,
    [Parameter(Mandatory = $true)] [string] $CandidateGraphics,
    [Parameter(Mandatory = $true)] [string] $OutputPath,
    [double] $CandidateMappingTolerance = 0.01,
    [double] $ReferenceSearchTolerance = 8.0
)
$ErrorActionPreference = 'Stop'
if (![double]::IsFinite($CandidateMappingTolerance) -or ![double]::IsFinite($ReferenceSearchTolerance) -or
    $CandidateMappingTolerance -le 0 -or $ReferenceSearchTolerance -le 0) { throw 'Tolerances must be finite and positive' }
$layout = Get-Content -LiteralPath $LayoutSnapshot -Raw | ConvertFrom-Json
function Read-CellClips([string] $Path) {
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($op in @(Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json)) {
        if ($op.Kind -ne 'Clip' -or $op.CurveCount -ne 0 -or $op.SegmentCount -ne 4 -or
            $op.LineCount -ne 3 -or $op.CloseCount -ne 1 -or $op.MaxX -le $op.MinX -or $op.MaxY -le $op.MinY) { continue }
        $key = '{0}|{1:F3}|{2:F3}|{3:F3}|{4:F3}' -f $op.PageNumber,$op.MinX,$op.MinY,$op.MaxX,$op.MaxY
        if ($seen.Add($key)) {
            [pscustomobject]@{ Key=$key; Page=$op.PageNumber; X=[double]$op.MinX; Y=[double]$op.MinY;
                Width=[double]($op.MaxX-$op.MinX); Height=[double]($op.MaxY-$op.MinY) }
        }
    }
}
$reference = @(Read-CellClips $ReferenceGraphics)
$candidate = @(Read-CellClips $CandidateGraphics)
$cells = @(for ($pageIndex=0; $pageIndex -lt $layout.Pages.Count; $pageIndex++) {
    foreach ($row in @($layout.Pages[$pageIndex].TableRows)) {
        foreach ($cell in @($row.Cells)) {
            [pscustomobject]@{ Page=$pageIndex+1; Table=$row.TableIndex; Row=$row.RowIndex; Cell=$cell.CellIndex;
                X=[double]$cell.X; Y=[double]$cell.Y; Width=[double]$cell.Width; Height=[double]$cell.Height;
                Ownership=$cell.VisualOwnership; FragmentCount=$row.FragmentCount }
        }
    }
})
$results = [Collections.Generic.List[object]]::new()
$usedReference = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
$usedCandidate = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
foreach ($page in @($cells | Group-Object Page)) {
    $pageCells = @($page.Group)
    $pageClips = @($candidate | Where-Object { $_.Page -eq [int]$page.Name })
    # Rendering applies one translation to cell geometry. Infer it from exact
    # source dimensions and require unique clips at every admitted coordinate.
    $offsets = @(foreach ($cell in $pageCells) {
        foreach ($clip in $pageClips) {
            if ([Math]::Abs($cell.Width-$clip.Width) -le $CandidateMappingTolerance -and
                [Math]::Abs($cell.Height-$clip.Height) -le $CandidateMappingTolerance) {
                [pscustomobject]@{ X=$clip.X-$cell.X; Y=$clip.Y-$cell.Y;
                    Key=('{0:F1}|{1:F1}' -f ($clip.X-$cell.X),($clip.Y-$cell.Y)) }
            }
        }
    })
    $groups = @($offsets | Group-Object Key | Sort-Object Count -Descending)
    $uniqueTranslation = $groups.Count -gt 0 -and ($groups.Count -eq 1 -or $groups[0].Count -gt $groups[1].Count)
    $offsetX = if ($uniqueTranslation) { ($groups[0].Group.X | Measure-Object -Average).Average } else { 0d }
    $offsetY = if ($uniqueTranslation) { ($groups[0].Group.Y | Measure-Object -Average).Average } else { 0d }
    foreach ($cell in $pageCells) {
        $row = [ordered]@{ Page=$cell.Page; Table=$cell.Table; Row=$cell.Row; Cell=$cell.Cell;
            SourceOwnership=$cell.Ownership; SourceFragmentCount=$cell.FragmentCount; Status='UnresolvedSourceMapping';
            CandidateClip=$null; ReferenceClip=$null; MaxBoundsDelta=$null }
        $matches = @(if ($uniqueTranslation) { $pageClips | Where-Object {
            [Math]::Abs($_.X-($cell.X+$offsetX)) -le $CandidateMappingTolerance -and
            [Math]::Abs($_.Y-($cell.Y+$offsetY)) -le $CandidateMappingTolerance -and
            [Math]::Abs($_.Width-$cell.Width) -le $CandidateMappingTolerance -and
            [Math]::Abs($_.Height-$cell.Height) -le $CandidateMappingTolerance } })
        if ($matches.Count -eq 1 -and $usedCandidate.Add($matches[0].Key)) {
            $clip=$matches[0]
            $row.CandidateClip=$clip
            $row.Status='MissingOrUnresolvedReferenceClip'
            $nearby = @(foreach ($ref in $reference | Where-Object { $_.Page -eq $cell.Page -and !$usedReference.Contains($_.Key) }) {
                $deltas = @([Math]::Abs($ref.X-$clip.X),[Math]::Abs($ref.Y-$clip.Y),
                    [Math]::Abs(($ref.X+$ref.Width)-($clip.X+$clip.Width)),[Math]::Abs(($ref.Y+$ref.Height)-($clip.Y+$clip.Height)))
                $maximum=($deltas|Measure-Object -Maximum).Maximum
                if ($maximum -le $ReferenceSearchTolerance) {
                    [pscustomobject]@{ Clip=$ref; Maximum=$maximum; Score=($deltas|Measure-Object -Sum).Sum }
                }
            })
            $nearby = @($nearby | Sort-Object Score)
            if ($nearby.Count -gt 0 -and ($nearby.Count -eq 1 -or [Math]::Abs($nearby[0].Score-$nearby[1].Score) -gt $CandidateMappingTolerance)) {
                $row.Status='MatchedSourceCellClip'
                $row.ReferenceClip=$nearby[0].Clip
                $row.MaxBoundsDelta=$nearby[0].Maximum
                [void]$usedReference.Add($nearby[0].Clip.Key)
            }
        }
        $results.Add([pscustomobject]$row)
    }
}
$matched=@($results|Where-Object Status -eq 'MatchedSourceCellClip')
[ordered]@{
    Method='Supplemental source-dimension and translation mapping of rectangular cell clips; does not alter gates'
    SourceCellCount=$cells.Count; MatchedCellCount=$matched.Count; UnresolvedCellCount=$cells.Count-$matched.Count
    CandidateRectangularClipCount=$candidate.Count; ReferenceRectangularClipCount=$reference.Count
    UnclassifiedCandidateClipCount=$candidate.Count-$usedCandidate.Count
    UnclassifiedReferenceClipCount=$reference.Count-$usedReference.Count
    MaxMatchedBoundsDelta=if($matched.Count){($matched.MaxBoundsDelta|Measure-Object -Maximum).Maximum}else{$null}
    CandidateMappingTolerance=$CandidateMappingTolerance; ReferenceSearchTolerance=$ReferenceSearchTolerance
    LayoutSha256=(Get-FileHash -LiteralPath $LayoutSnapshot).Hash
    ReferenceGraphicsSha256=(Get-FileHash -LiteralPath $ReferenceGraphics).Hash
    CandidateGraphicsSha256=(Get-FileHash -LiteralPath $CandidateGraphics).Hash
    Cells=@($results)
}|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $OutputPath -Encoding utf8
