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
    private readonly WindowsFontResolver.RetainedFileTracker? retention;
    private readonly object sync = new();
    private ReadOnlyMemory<byte>? cachedBytes;

    public FileFontProgramSource(string path)
        : this(path, retention: null)
    {
    }

    // R13: snapshot-tracked sources report loads and accesses so the shared
    // snapshot can cap aggregate retained bytes with LRU eviction. Untracked
    // sources (including all direct public uses) behave exactly as before.
    internal FileFontProgramSource(string path, WindowsFontResolver.RetainedFileTracker? retention)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        this.path = System.IO.Path.GetFullPath(path);
        this.retention = retention;
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
            retention?.NoteAccessed(path);
            return ValueTask.FromResult(bytes);
        }

        ct.ThrowIfCancellationRequested();
        lock (sync)
        {
            if (cachedBytes is ReadOnlyMemory<byte> cached)
            {
                retention?.NoteAccessed(path);
                return ValueTask.FromResult(cached);
            }

            ct.ThrowIfCancellationRequested();
            CheckLocalFontSize(new FileInfo(path).Length, path);
            byte[] loaded = File.ReadAllBytes(path);
            cachedBytes = loaded;
            retention?.NoteLoaded(path, loaded.LongLength);
            return ValueTask.FromResult((ReadOnlyMemory<byte>)loaded);
        }
    }

    // R13: clears retained bytes so the owning snapshot stays bounded. Arrays
    // already handed out stay alive (and valid: files are immutable program
    // sources) via GC; the next access simply re-reads through the per-source
    // lock. Takes no lock itself so snapshot eviction under its own lock cannot
    // deadlock against in-progress loads.
    internal void EvictCachedBytes()
    {
        cachedBytes = null;
    }
}
