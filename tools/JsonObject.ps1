# Shared JSON-object reading (T04): Read-JsonObject resolves one JSON
# document, throwing on missing files. Read-JsonObjectIfExists returns
# null for blank or missing paths and otherwise delegates to the strict
# reader. Dot-source this file instead of copying the functions into
# comparison, summary, and gate tools.

function Read-JsonObject([string] $Path) {
    return Get-Content -Raw -LiteralPath (Resolve-Path -LiteralPath $Path).Path | ConvertFrom-Json
}

function Read-JsonObjectIfExists([string] $Path) {
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
        return $null
    }

    return Read-JsonObject $Path
}
