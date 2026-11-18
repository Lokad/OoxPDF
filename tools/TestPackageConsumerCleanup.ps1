# Exercise the live cleanup function with owned files; no builds or consumers.
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$tokens = $null; $parseErrors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile(
    (Join-Path $PSScriptRoot 'TestPackageSmoke.ps1'), [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw 'Package smoke script did not parse.' }
$helper = $ast.Find({ param($node)
    $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Remove-CompletedPackageConsumerCaches'
}, $false)
if ($null -eq $helper) { throw 'Package consumer cleanup is missing.' }
. ([scriptblock]::Create($helper.Extent.Text))
$artifactRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../artifacts/package-consumer-cleanup'))
$temporary = Join-Path $artifactRoot ([guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temporary -Force | Out-Null
$links = [Collections.Generic.List[string]]::new()
$checks = 0; $passed = $false
function Check([bool]$Condition, [string]$Name) {
    if (!$Condition) { throw $Name }
    $script:checks++
}
function New-Consumer([string]$Name) {
    $root = Join-Path $temporary "$Name/artifacts/package-smoke"
    $directory = Join-Path $root ([guid]::NewGuid().ToString('N'))
    foreach ($role in @('bin', 'obj', 'packages', 'staging')) {
        New-Item -ItemType Directory -Path (Join-Path $directory $role) -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $directory "$role/retained.txt") -Value $role
    }
    foreach ($name in @('input.docx', 'output.pdf', 'summary.json')) {
        Set-Content -LiteralPath (Join-Path $directory $name) -Value $name
    }
    return [pscustomobject]@{Root=$root;Directory=$directory}
}
function Caches-Present($Consumer) {
    return @('bin', 'obj', 'packages' | Where-Object {
        Test-Path -LiteralPath (Join-Path $Consumer.Directory "$_/retained.txt")
    }).Count -eq 3
}
function Link-Directory([string]$Path, [string]$Target) {
    $kind = if ($IsWindows) { 'Junction' } else { 'SymbolicLink' }
    New-Item -ItemType $kind -Path $Path -Target $Target | Out-Null
    $links.Add($Path)
}
try {
    $success = New-Consumer 'passing consumer with spaces'
    Check (Remove-CompletedPackageConsumerCaches $success.Directory $success.Root) 'Passing cleanup refused'
    Check (@('bin', 'obj', 'packages' | Where-Object { Test-Path -LiteralPath (Join-Path $success.Directory $_) }).Count -eq 0) 'Passing cache remains'
    foreach ($name in @('input.docx', 'output.pdf', 'summary.json', 'staging/retained.txt')) {
        $expected = if ($name.StartsWith('staging/')) { 'staging' } else { $name }
        Check ((Get-Content -LiteralPath (Join-Path $success.Directory $name) -Raw).Trim() -ceq $expected) 'Retained consumer evidence changed'
    }
    Check (Remove-CompletedPackageConsumerCaches $success.Directory $success.Root) 'Missing caches prevent idempotent cleanup'

    $failure = New-Consumer 'failed consumer'
    Remove-Item -LiteralPath (Join-Path $failure.Directory 'summary.json')
    Check (!(Remove-CompletedPackageConsumerCaches $failure.Directory $failure.Root 3>$null)) 'Uncompleted consumer admitted'
    Check (Caches-Present $failure) 'Failed consumer caches changed'
    Check (!(Remove-CompletedPackageConsumerCaches $failure.Root $failure.Root 3>$null)) 'Parent directory admitted'

    $child = New-Consumer 'linked child'
    Link-Directory (Join-Path $child.Directory 'obj/alias') (Join-Path $failure.Directory 'staging')
    Check (!(Remove-CompletedPackageConsumerCaches $child.Directory $child.Root 3>$null)) 'Linked child admitted'
    Check (Caches-Present $child) 'Linked child refusal partially removed caches'
    $aliasRoot = Join-Path $temporary 'aliased consumer root'
    Link-Directory $aliasRoot $child.Root
    Check (!(Remove-CompletedPackageConsumerCaches (Join-Path $aliasRoot (Split-Path $child.Directory -Leaf)) $aliasRoot 3>$null)) 'Linked parent admitted'

    $hard = New-Consumer 'shared cache file'
    $foreign = Join-Path $failure.Directory 'staging/retained.txt'
    New-Item -ItemType HardLink -Path (Join-Path $hard.Directory 'packages/shared.txt') -Target $foreign | Out-Null
    Check (!(Remove-CompletedPackageConsumerCaches $hard.Directory $hard.Root 3>$null)) 'Shared file admitted'
    Check (Caches-Present $hard) 'Shared file refusal partially removed caches'
    Check ((Get-Content -LiteralPath $foreign -Raw).Trim() -ceq 'staging') 'External target changed'

    $refusal = New-Consumer 'removal refusal'
    $result = & {
        function Remove-Item { throw [IO.IOException]::new('Fixture deletion refusal') }
        Remove-CompletedPackageConsumerCaches $refusal.Directory $refusal.Root 3>$null
    }
    Check (!$result) 'Removal failure was not reported'
    Check (Caches-Present $refusal) 'Removal refusal lost cache evidence'
    $passed = $true
}
finally {
    if ($passed) {
        foreach ($link in $links) { [IO.Directory]::Delete($link) }
        $target = [IO.Path]::GetFullPath($temporary)
        if (![IO.Path]::GetDirectoryName($target).Equals($artifactRoot, [StringComparison]::Ordinal)) { throw 'Test cleanup escaped owned artifact root' }
        Remove-Item -LiteralPath $target -Recurse -Force
    }
    else { Write-Warning "Cleanup control failed; fixtures retained at $temporary" }
}
"$checks package-consumer cleanup checks passed; no build, restore or conversion."
