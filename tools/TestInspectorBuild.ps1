# Exercise the inspectors' production build calls with isolated projects.
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
foreach ($scriptName in @('InspectPdf.ps1', 'InspectPptxText.ps1')) {
    $tokens = $null
    $parseErrors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile(
        (Join-Path $PSScriptRoot $scriptName), [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count) { throw "Inspector script did not parse: $scriptName" }
    $call = $ast.FindAll({param($node)
        $node -is [Management.Automation.Language.CommandAst] -and
        $node.Extent.Text.StartsWith('& (Join-Path $repoRoot') -and
        $node.Extent.Text.Contains('EnsureDotnetBuild.ps1')
    }, $false) | Select-Object -First 1
    if (!$call) { throw "Shared inspector build call is missing: $scriptName" }
    $build = [scriptblock]::Create($call.Extent.Text)
    $scratch = Join-Path $repoRoot ('artifacts/inspector-build-test/'+[guid]::NewGuid().ToString('N'))
    $projectDirectory = Join-Path $scratch 'probe'
    New-Item -ItemType Directory -Path $projectDirectory -Force | Out-Null
    $project = Join-Path $projectDirectory 'probe.csproj'
    $source = Join-Path $projectDirectory 'Program.cs'
    $properties = Join-Path $scratch 'Directory.Build.props'
    $dll = Join-Path $projectDirectory 'bin/Debug/net10.0/probe.dll'
    $buildRecordPath = Join-Path $scratch 'build-info.json'
    Set-Content -LiteralPath $project -Value '<Project Sdk="Microsoft.NET.Sdk"><PropertyGroup><TargetFramework>net10.0</TargetFramework><OutputType>Exe</OutputType></PropertyGroup></Project>'
    Set-Content -LiteralPath $source -Value @'
#if UPDATED_BUILD
System.Console.WriteLine("updated");
#else
System.Console.WriteLine("initial");
#endif
'@
    Set-Content -LiteralPath $properties -Value '<Project />'
    . $build
    if ((& dotnet $dll) -ne 'initial') { throw "Initial inspector build failed: $scriptName" }
    Write-Host "PASS $scriptName initial build executes"
    $sourceTimestamp = (Get-Item -LiteralPath $source).LastWriteTimeUtc
    $projectTimestamp = (Get-Item -LiteralPath $project).LastWriteTimeUtc
    $initialHash = (Get-FileHash -LiteralPath $dll).Hash
    Set-Content -LiteralPath $properties -Value '<Project><PropertyGroup><DefineConstants>UPDATED_BUILD</DefineConstants></PropertyGroup></Project>'
    . $build
    if ((& dotnet $dll) -ne 'updated' -or (Get-FileHash -LiteralPath $dll).Hash -eq $initialHash) {
        throw "Changed parent build properties reused stale inspector bytes: $scriptName"
    }
    if ((Get-Item -LiteralPath $source).LastWriteTimeUtc -ne $sourceTimestamp -or
        (Get-Item -LiteralPath $project).LastWriteTimeUtc -ne $projectTimestamp) {
        throw "The inspector probe must leave source/project timestamps unchanged: $scriptName"
    }
    Write-Host "PASS $scriptName parent properties refresh unchanged source"
    $updatedHash = (Get-FileHash -LiteralPath $dll).Hash
    . $build
    $record = Get-Content -LiteralPath $buildRecordPath -Raw | ConvertFrom-Json
    if ((Get-FileHash -LiteralPath $dll).Hash -ne $updatedHash -or
        $record.OutputSha256 -ne $updatedHash.ToLowerInvariant() -or
        $record.Configuration -ne 'Debug') {
        throw "Unchanged inspector bytes/provenance disagree: $scriptName"
    }
    Write-Host "PASS $scriptName unchanged bytes and provenance agree"
    Add-Content -LiteralPath $source -Value 'This is invalid C#;'
    $rejected = $false
    try { . $build } catch {
        if ($_.Exception.Message -notlike '*build failed*') { throw }
        $rejected = $true
    }
    if (!$rejected -or !(Test-Path -LiteralPath $dll)) {
        throw "Failed inspector build must reject an existing stale DLL: $scriptName"
    }
    Write-Host "PASS $scriptName failed build rejects stale DLL"
}
Write-Host '8 inspector build checks passed.'
exit 0
