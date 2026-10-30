namespace Lokad.OoxPdf.Fonts;

/// <summary>Acquires immutable local font bytes using bounded, optionally funded caching.</summary>
/// <remarks>
/// Keep files unchanged for this source's lifetime. A cache budget is borrowed
/// until this source and outstanding leases release their storage. Without one,
/// governed operations fund private leases instead of cache retention. Disposal
/// refuses new acquisition and ends retention; outstanding leases stay valid.
/// </remarks>
public sealed class FileFontProgramSource : IFontProgramSource, IDisposable
{
    internal const long MaxLocalFontBytes = 64L * 1024L * 1024L;
    private readonly string path;
    private readonly FontProgramCache cache;
    private readonly bool ownsCache;
    private readonly SemaphoreSlim gate = new(1, 1);
    private int disposed;

    public FileFontProgramSource(string path, IMemoryBudget? cacheBudget = null)
        : this(path, new FontProgramCache(MaxLocalFontBytes, StringComparer.OrdinalIgnoreCase, cacheBudget))
    {
        ownsCache = true;
    }

    internal FileFontProgramSource(string path, FontProgramCache cache)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        this.path = System.IO.Path.GetFullPath(path);
        this.cache = cache;
    }

    internal static void CheckLocalFontSize(long byteCount, string path)
    {
        if (byteCount > MaxLocalFontBytes)
            throw new InvalidDataException($"Local font file '{path}' exceeds the maximum supported size of {MaxLocalFontBytes} bytes.");
    }

    public string Path => path;
    public string StableId => "file:" + path;

    public async ValueTask<IOoxPdfFontProgramLease> AcquireAsync(IMemoryBudget? operationBudget, CancellationToken ct = default)
    {
        ct.ThrowIfCancellationRequested();
        ObjectDisposedException.ThrowIf(Volatile.Read(ref disposed) != 0, this);
        if (cache.TryAcquire(path, operationBudget, out var cached)) return cached!;
        await gate.WaitAsync(ct).ConfigureAwait(false);
        FontProgramBuffer? program = null;
        try
        {
            ObjectDisposedException.ThrowIf(Volatile.Read(ref disposed) != 0, this);
            if (cache.TryAcquire(path, operationBudget, out cached)) return cached!;
            {
                await using var input = new FileStream(path, new FileStreamOptions { Mode = FileMode.Open,
                    Access = FileAccess.Read, Share = FileShare.Read, BufferSize = 1, Options = FileOptions.Asynchronous | FileOptions.SequentialScan });
                long length = input.Length;
                CheckLocalFontSize(length, path);
                IMemoryReservation? reservation = OoxMemory.Reserve(cache.Budget ?? operationBudget, length, cache.Budget is null ? OoxPdfLimitScope.Conversion : OoxPdfLimitScope.SharedResource);
                try
                {
                    ct.ThrowIfCancellationRequested();
                    byte[] bytes = new byte[checked((int)length)];
                    await input.ReadExactlyAsync(bytes, ct).ConfigureAwait(false);
                    ct.ThrowIfCancellationRequested();
                    program = new FontProgramBuffer(bytes, reservation);
                    reservation = null;
                }
                finally { reservation?.Dispose(); }
            }
            ct.ThrowIfCancellationRequested();
            cache.Store(path, program, operationBudget);
            return program.Acquire();
        }
        finally { program?.Dispose(); gate.Release(); }
    }

    public void Dispose()
    {
        if (Interlocked.Exchange(ref disposed, 1) == 0 && ownsCache) cache.Dispose();
    }
}
