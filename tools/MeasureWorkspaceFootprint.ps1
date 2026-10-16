param(
    [string]$Path = (Join-Path $PSScriptRoot '..'),
    [string]$ReportPath
)
$ErrorActionPreference = 'Stop'
$pythonName = if ($IsWindows) { 'python' } else { 'python3' }
$python = Get-Command $pythonName -CommandType Application -ErrorAction Stop | Select-Object -First 1
$payload = & $python.Source (Join-Path $PSScriptRoot 'MeasureWorkspaceFootprint.py') $Path
if ($LASTEXITCODE -ne 0) { throw 'Workspace measurement failed' }
$report = $payload | ConvertFrom-Json
if ($ReportPath) { $report | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $ReportPath -Encoding utf8 }
$report
if ($report.Errors.Count) { Write-Warning 'Some entries could not be measured; the reported size is incomplete.' }
if ($report.Bytes -gt 5e9) { Write-Warning 'Workspace exceeds the loose 5 GB target. Review obsolete generated runs and worktrees.' }
