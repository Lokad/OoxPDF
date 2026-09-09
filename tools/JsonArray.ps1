# Shared JSON-array reading (T04): Read-JsonArray resolves a JSON document
# into an array. Missing files throw (callers over optional inputs guard
# with Test-Path first); null or scalar documents become empty or
# single-element arrays without enumerating. Dot-source this file instead
# of copying the function into comparison and summary tools.

function Read-JsonArray([string] $Path) {
    $items = Get-Content -Raw -LiteralPath (Resolve-Path -LiteralPath $Path).Path | ConvertFrom-Json
    if ($null -eq $items) {
        return ,@()
    }

    if ($items -is [array]) {
        return ,$items
    }

    return ,@($items)
}

function Read-JsonArrayIfExists([string] $Path) {
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
        return @()
    }

    return Read-JsonArray $Path
}
