namespace Lokad.OoxPdf.Fonts;

public sealed class FileFontProgramSource : IFontProgramSource
{
    // PLAN M09: local files had no size guard while HTTP pack fonts are capped at
    // 64 MiB per file. Parity cap here: installed fonts peak at 35.4 MiB on the
    // reference machine (mingliub.ttc, measured 2026-09-21), so 64 MiB keeps ~1.8x
    // headroom. Oversized files fail with InvalidDataException and fall back to a
    // substitute face through FontProgramLoader, like any malformed font.
    internal const long MaxLocalFontBytes = 64L * 1024L * 1024L;

    private readonly string path;
    private readonly object sync = new();
    private ReadOnlyMemory<byte>? cachedBytes;

    public FileFontProgramSource(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        this.path = System.IO.Path.GetFullPath(path);
    }

    internal static void CheckLocalFontSize(long byteCount, string path)
    {
        if (byteCount > MaxLocalFontBytes)
        {
            throw new InvalidDataException($"Local font file '{path}' exceeds the maximum supported size of {MaxLocalFontBytes} bytes.");
        }
    }

    public string Path => path;

    public string StableId => "file:" + path;

    public ValueTask<ReadOnlyMemory<byte>> GetBytesAsync(CancellationToken ct = default)
    {
        if (cachedBytes is ReadOnlyMemory<byte> bytes)
        {
            return ValueTask.FromResult(bytes);
        }

        ct.ThrowIfCancellationRequested();
        lock (sync)
        {
            if (cachedBytes is ReadOnlyMemory<byte> cached)
            {
                return ValueTask.FromResult(cached);
            }

            ct.ThrowIfCancellationRequested();
            CheckLocalFontSize(new FileInfo(path).Length, path);
            byte[] loaded = File.ReadAllBytes(path);
            cachedBytes = loaded;
            return ValueTask.FromResult((ReadOnlyMemory<byte>)loaded);
        }
    }
}
