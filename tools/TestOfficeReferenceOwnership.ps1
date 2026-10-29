# Ownership policy checks with managed stubs; no Office, native calls or builds.
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'OfficeReferenceOwnership.ps1')
$start = [DateTime]::UtcNow
$checks = 0
foreach ($case in @(
    @{Name='new mapped process';Process=[pscustomobject]@{Id=123;ProcessName='WINWORD';StartTime=$start.AddSeconds(1)};Existing=@(122);Expected=$true},
    @{Name='preexisting process';Process=[pscustomobject]@{Id=123;ProcessName='WINWORD';StartTime=$start.AddSeconds(1)};Existing=@(123);Expected=$false},
    @{Name='older unlisted process';Process=[pscustomobject]@{Id=123;ProcessName='WINWORD';StartTime=$start.AddSeconds(-1)};Existing=@();Expected=$false},
    @{Name='other application';Process=[pscustomobject]@{Id=123;ProcessName='POWERPNT';StartTime=$start.AddSeconds(1)};Existing=@();Expected=$false},
    @{Name='unknown process';Process=$null;Existing=@();Expected=$false}
)) {
    if ((Test-OfficeProcessOwned $case.Process $case.Existing $start 'WINWORD') -ne $case.Expected) { throw $case.Name }
    $checks++
}
$copy = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../artifacts/ownership-control/copy/input.docx'))
$foreign = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../artifacts/ownership-control/foreign/input.docx'))
foreach ($case in @(
    @{Document=[pscustomobject]@{FullName=$copy};Expected=$true},
    @{Document=[pscustomobject]@{FullName=$foreign};Expected=$false},
    @{Document=[pscustomobject]@{};Expected=$false},
    @{Document=$null;Expected=$false}
)) {
    if ((Test-OfficeDocumentOwned $case.Document $copy) -ne $case.Expected) { throw 'Document ownership differs' }
    $checks++
}
if ($null -ne (Get-OfficeWindowProcess 0)) { throw 'Unknown window acquired a process' }
$checks++
Write-Output "$checks Office document/process ownership checks passed; no Office or native calls."
