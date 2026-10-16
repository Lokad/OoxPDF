# Check shared-file accounting and loop rejection without scanning the workspace.
[CmdletBinding()]
param([string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
if (!$OutputDirectory) { $OutputDirectory = Join-Path (Split-Path $PSScriptRoot) ('artifacts/footprint-test-' + [Guid]::NewGuid().ToString('N')) }
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Use a fresh footprint test directory' }
$inputs = Join-Path $OutputDirectory 'inputs'
New-Item -ItemType Directory -Path (Join-Path $inputs nested) -Force | Out-Null
$inputs = (Resolve-Path -LiteralPath $inputs).Path
[IO.File]::WriteAllBytes((Join-Path $inputs a.bin), [byte[]](1,2,3,4,5))
[IO.File]::WriteAllBytes((Join-Path $inputs nested/c.bin), [byte[]](6,7,8))
New-Item -ItemType HardLink -Path (Join-Path $inputs nested/b.bin) -Target (Join-Path $inputs a.bin) | Out-Null
$linkType = if ($IsWindows) { 'Junction' } else { 'SymbolicLink' }
New-Item -ItemType $linkType -Path (Join-Path $inputs loop) -Target $inputs | Out-Null
$longDirectory = Join-Path (Join-Path $inputs ('a' * 160)) ('b' * 160)
[IO.Directory]::CreateDirectory($longDirectory) | Out-Null
[IO.File]::WriteAllBytes((Join-Path $longDirectory 'long.bin'), [byte[]](9))
$path = Join-Path $OutputDirectory report.json
$report = & (Join-Path $PSScriptRoot 'MeasureWorkspaceFootprint.ps1') -Path $inputs -ReportPath $path
if ($report.Bytes -ne 9 -or $report.LogicalBytes -ne 14 -or $report.SharedHardLinkBytes -ne 5) { throw 'Shared or long-path content was counted incorrectly' }
if ($report.Files -ne 4 -or $report.Directories -ne 3 -or $report.SkippedLinks -ne 1 -or $report.Errors.Count) { throw 'Input/link/long-path traversal differs' }
$stored = Get-Content $path -Raw | ConvertFrom-Json
if ($stored.Bytes -ne $report.Bytes -or $stored.Root -ne $inputs) { throw 'Persisted footprint report differs' }
Write-Output 'Unique bytes, logical bytes, hard-link aliases, directory loops and long paths passed.'
