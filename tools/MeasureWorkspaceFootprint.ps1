param(
    [string]$Path = (Join-Path $PSScriptRoot '..'),
    [string]$ReportPath
)

$ErrorActionPreference = 'Stop'
$root = (Get-Item -LiteralPath $Path -Force).FullName
if (-not (Test-Path -LiteralPath $root -PathType Container)) { throw "Not a directory: $root" }
$pending = [Collections.Generic.Stack[IO.DirectoryInfo]]::new()
$pending.Push([IO.DirectoryInfo]::new($root))
$groups = @{}
[long]$bytes = 0
[long]$files = 0
[long]$directories = 0
[long]$skippedLinks = 0
$errors = [Collections.Generic.List[string]]::new()
$timer = [Diagnostics.Stopwatch]::StartNew()
while ($pending.Count) {
    $directory = $pending.Pop()
    try {
        foreach ($entry in $directory.EnumerateFileSystemInfos()) {
            if ($entry.Attributes -band [IO.FileAttributes]::ReparsePoint) { $skippedLinks++; continue }
            if ($entry -is [IO.DirectoryInfo]) { $directories++; $pending.Push($entry); continue }
            $length = $entry.Length
            $bytes += $length
            $files++
            $relative = [IO.Path]::GetRelativePath($root, $entry.FullName)
            $separator = $relative.IndexOf([IO.Path]::DirectorySeparatorChar)
            $group = if ($separator -lt 0) { '.' } else { $relative.Substring(0, $separator) }
            if (-not $groups.ContainsKey($group)) { $groups[$group] = @{ Files = 0L; Bytes = 0L } }
            $groups[$group].Files++
            $groups[$group].Bytes += $length
        }
    } catch { $errors.Add("$($directory.FullName): $($_.Exception.Message)") }
}
$report = [ordered]@{
    Root = $root
    CheckedUtc = [DateTime]::UtcNow.ToString('o')
    Bytes = $bytes
    GB = [Math]::Round($bytes / 1e9, 3)
    Files = $files
    Directories = $directories
    SkippedLinks = $skippedLinks
    Errors = @($errors.ToArray())
    ElapsedSeconds = [Math]::Round($timer.Elapsed.TotalSeconds, 2)
    Groups = @($groups.GetEnumerator() | ForEach-Object {
        [pscustomobject][ordered]@{ Path = $_.Key; Bytes = $_.Value.Bytes; GB = [Math]::Round($_.Value.Bytes / 1e9, 3); Files = $_.Value.Files }
    } | Sort-Object { $_.Bytes } -Descending)
}
if ($ReportPath) { $report | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $ReportPath -Encoding utf8 }
[pscustomobject]$report
if ($errors.Count) { Write-Warning 'Some entries could not be measured; the reported size is incomplete.' }
if ($bytes -gt 5e9) { Write-Warning 'Workspace exceeds the loose 5 GB target. Review obsolete generated runs and worktrees.' }
