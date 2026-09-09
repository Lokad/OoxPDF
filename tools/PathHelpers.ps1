# Shared path and value primitives (T04): ConvertTo-RepoPath renders
# repository-relative display paths, Expand-PathList splits comma- or
# semicolon-separated path lists, and IntValue coerces possibly-missing
# values to integers. ConvertTo-RepoPath resolves $repoRoot in caller scope;
# every consumer sets it identically. Dot-source this file instead of
# copying the functions into workflow tools.

function ConvertTo-RepoPath([string] $Path) {
    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $fullRoot = [System.IO.Path]::GetFullPath($repoRoot).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    if ($fullPath.StartsWith($fullRoot + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $fullPath.Substring($fullRoot.Length + 1).Replace([System.IO.Path]::DirectorySeparatorChar, "/")
    }

    return $fullPath

}
function IntValue($value) { if ($null -eq $value) { return 0 } return [int]$value }

function Expand-PathList([string[]] $Values) {
    $expanded = New-Object System.Collections.Generic.List[string]
    foreach ($value in $Values) {
        if ([string]::IsNullOrWhiteSpace($value)) {
            continue
        }

        foreach ($part in ($value -split "[,;]")) {
            if (-not [string]::IsNullOrWhiteSpace($part)) {
                $expanded.Add($part.Trim())
            }
        }
    }

    return ,$expanded.ToArray()
}
