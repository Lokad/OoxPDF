# Shared OOXML fixture packaging (T04): New-ZipPackage writes a hashtable of
# UTF-8 entries with ordinal names and a fixed ZIP epoch. Dot-source this file instead of
# copying the function into fixture generators.

function New-ZipPackage {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path,

        [Parameter(Mandatory = $true)]
        [hashtable] $Entries
    )

    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Force
    }

    $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::CreateNew)
    try {
        $archive = [System.IO.Compression.ZipArchive]::new($stream, [System.IO.Compression.ZipArchiveMode]::Create, $true)
        try {
            [string[]] $names = @($Entries.Keys)
            [Array]::Sort($names, [StringComparer]::Ordinal)
            foreach ($name in $names) {
                $entry = $archive.CreateEntry($name)
                $entry.LastWriteTime = [DateTimeOffset]::new(1980, 1, 1, 0, 0, 0, [TimeSpan]::Zero)
                $entryStream = $entry.Open()
                try {
                    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Entries[$name])
                    $entryStream.Write($bytes, 0, $bytes.Length)
                }
                finally {
                    $entryStream.Dispose()
                }
            }
        }
        finally {
            $archive.Dispose()
        }
    }
    finally {
        $stream.Dispose()
    }
}
