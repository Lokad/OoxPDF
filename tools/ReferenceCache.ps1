# Shared Office-reference cache key helpers. Dot-source this file instead of
# copying key logic: cache keys carry an explicit rendering-protocol version so
# cosmetic script edits no longer orphan every cached reference. Bump
# $ReferenceCacheProtocol only when the rendered reference bytes meaningfully
# change (Office export flags, raster DPI handling, variant semantics).

$ReferenceCacheProtocol = "v1"

function Get-ShortSha256([byte[]] $Bytes, [int] $Length = 12) {
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        $hash = ([System.BitConverter]::ToString($sha256.ComputeHash($Bytes)) -replace "-", "").ToLowerInvariant()
        return $hash.Substring(0, [Math]::Min($Length, $hash.Length))
    }
    finally {
        $sha256.Dispose()
    }
}

function Get-ReferenceCacheKey(
    [string] $InputPath,
    [int] $DpiValue,
    [string] $CacheVariant)
{
    $inputHash = (Get-FileHash -LiteralPath $InputPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $extension = [System.IO.Path]::GetExtension($InputPath).TrimStart(".").ToLowerInvariant()
    $variantKeyPart = ""
    if (-not [string]::IsNullOrWhiteSpace($CacheVariant)) {
        $variantHash = Get-ShortSha256 ([System.Text.Encoding]::UTF8.GetBytes($CacheVariant.Trim().ToLowerInvariant())) 12
        $variantKeyPart = "-variant" + $variantHash
    }

    "{0}-{1}-{2}{3}-dpi{4}" -f $extension, $inputHash.Substring(0, 24), $ReferenceCacheProtocol, $variantKeyPart, $DpiValue
}