using Lokad.OoxPdf.Diagnostics;

namespace Lokad.OoxPdf.Pdf;

// One resident window, retained vector descriptors and encoding/readback buffer. Beyond the window,
// an explicit provider owns temporary backing. Reads stay bounded, records preserve
// page/group order, and every I/O borrow completes before disposal or refund.
internal sealed class PdfPageContentStaging : IDisposable, IAsyncDisposable
{
    internal const int ChunkByteCount = 65536;
    private static readonly byte[] LengthPlaceholder = new byte[8];
    private readonly List<MemoryPage> memory = new();
    private readonly List<long> spillOffsets = new();
    private readonly long windowBytes;
    private readonly Action<OoxPdfDiagnostic>? diagnosticSink;
    private readonly IOoxPdfTemporaryStorage? temporaryStorage;
    private readonly IMemoryReservation? storage;
    private byte[]? scratch;
    private long residentBytes, bufferedBytes, retainedResourceBytes;
    private Stream? spillStream;
    private bool disposed;
    private MemoryPage? pendingPage;
    private long pendingLength, pendingOffset = -1;
    private bool pendingRecordStarted;

    private sealed class MemoryPage
    {
        public readonly List<byte[]> Segments = new();
        public int Length;
    }

    public PdfPageContentStaging(long windowBytes, Action<OoxPdfDiagnostic>? diagnosticSink,
        IOoxPdfTemporaryStorage? temporaryStorage = null)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(windowBytes);
        this.windowBytes = windowBytes;
        this.diagnosticSink = diagnosticSink;
        this.temporaryStorage = temporaryStorage;
        storage = OoxMemory.Reserve(OoxConversionBudget.Current?.Memory, 0);
    }

    internal byte[] Scratch
    {
        get
        {
            ObjectDisposedException.ThrowIf(disposed, this);
            if (scratch is null)
            {
                OoxMemory.Resize(storage, checked(retainedResourceBytes + residentBytes + ChunkByteCount));
                scratch = new byte[ChunkByteCount];
            }
            return scratch;
        }
    }

    // Descriptors share this existing owner through staged-document emission.
    // Spilling drops content segments, never their still-live vector resources.
    internal void RetainResourceCapacity(long bytes)
    {
        ObjectDisposedException.ThrowIf(disposed, this);
        long retained = checked(retainedResourceBytes + bytes);
        OoxMemory.Resize(storage, checked(retained + residentBytes + (scratch?.Length ?? 0)));
        retainedResourceBytes = retained;
    }

    public int Count => memory.Count + spillOffsets.Count;
    public long SpilledBytes { get; private set; }
    public bool UsesTemporaryStorage => spillStream is not null;

    public void BeginPage(CancellationToken cancellationToken)
    {
        ObjectDisposedException.ThrowIf(disposed, this);
        cancellationToken.ThrowIfCancellationRequested();
        if (pendingPage is not null) throw new InvalidOperationException("A staged page is already open.");
        pendingPage = new MemoryPage();
        pendingLength = 0;
        pendingOffset = -1;
        pendingRecordStarted = false;
    }

    public async ValueTask AppendPageBytesAsync(byte[] chunk, int offset, int count, CancellationToken cancellationToken)
    {
        ObjectDisposedException.ThrowIf(disposed, this);
        ArgumentNullException.ThrowIfNull(chunk);
        ArgumentOutOfRangeException.ThrowIfNegative(offset);
        ArgumentOutOfRangeException.ThrowIfNegative(count);
        if (offset > chunk.Length - count) throw new ArgumentOutOfRangeException(nameof(count));
        cancellationToken.ThrowIfCancellationRequested();
        if (pendingPage is null) throw new InvalidOperationException("No staged page is open.");
        if (count == 0) return;
        if (spillStream is null && bufferedBytes + pendingLength + count > windowBytes)
            await EscalateAsync(cancellationToken).ConfigureAwait(false);
        if (spillStream is { } stream)
        {
            if (!pendingRecordStarted)
            {
                pendingOffset = stream.Position;
                await stream.WriteAsync(LengthPlaceholder, cancellationToken).ConfigureAwait(false);
                pendingRecordStarted = true;
            }
            await stream.WriteAsync(chunk.AsMemory(offset, count), cancellationToken).ConfigureAwait(false);
            pendingLength += count;
            SpilledBytes += count;
            cancellationToken.ThrowIfCancellationRequested();
        }
        else
        {
            OoxMemory.Resize(storage, checked(retainedResourceBytes + residentBytes + count + (scratch?.Length ?? 0)));
            var copy = new byte[count];
            Buffer.BlockCopy(chunk, offset, copy, 0, count);
            pendingPage.Segments.Add(copy);
            residentBytes += count;
            pendingLength += count;
        }
    }

    public async ValueTask EndPageAsync(CancellationToken cancellationToken)
    {
        ObjectDisposedException.ThrowIf(disposed, this);
        cancellationToken.ThrowIfCancellationRequested();
        if (pendingPage is null) throw new InvalidOperationException("No staged page is open.");
        if (spillStream is { } stream)
        {
            if (!pendingRecordStarted)
            {
                pendingOffset = stream.Position;
                await stream.WriteAsync(LengthPlaceholder, cancellationToken).ConfigureAwait(false);
            }
            long end = stream.Position;
            stream.Seek(pendingOffset, SeekOrigin.Begin);
            await stream.WriteAsync(BitConverter.GetBytes(pendingLength), cancellationToken).ConfigureAwait(false);
            stream.Seek(end, SeekOrigin.Begin);
            cancellationToken.ThrowIfCancellationRequested();
            spillOffsets.Add(pendingOffset);
        }
        else
        {
            pendingPage.Length = checked((int)pendingLength);
            memory.Add(pendingPage);
            bufferedBytes += pendingLength;
        }
        pendingPage = null;
        pendingLength = 0;
        pendingOffset = -1;
        pendingRecordStarted = false;
    }

    public async ValueTask<int> GetPageLengthAsync(int pageIndex, CancellationToken cancellationToken)
    {
        ObjectDisposedException.ThrowIf(disposed, this);
        cancellationToken.ThrowIfCancellationRequested();
        if (pendingPage is not null) throw new InvalidOperationException("Page production is still open.");
        return spillStream is { } stream
            ? await ReadRecordLengthAsync(stream, pageIndex, cancellationToken).ConfigureAwait(false)
            : memory[pageIndex].Length;
    }

    public async ValueTask<int> ReadPageBytesAsync(int pageIndex, byte[] destination, int destinationOffset,
        int sourceOffset, int count, CancellationToken cancellationToken)
    {
        ObjectDisposedException.ThrowIf(disposed, this);
        ArgumentNullException.ThrowIfNull(destination);
        ArgumentOutOfRangeException.ThrowIfNegative(destinationOffset);
        ArgumentOutOfRangeException.ThrowIfNegative(sourceOffset);
        ArgumentOutOfRangeException.ThrowIfNegative(count);
        if (destinationOffset > destination.Length - count) throw new ArgumentOutOfRangeException(nameof(count));
        cancellationToken.ThrowIfCancellationRequested();
        if (pendingPage is not null) throw new InvalidOperationException("Page production is still open.");
        if (count == 0) return 0;
        if (spillStream is { } stream)
        {
            int length = await ReadRecordLengthAsync(stream, pageIndex, cancellationToken).ConfigureAwait(false);
            if (sourceOffset >= length) return 0;
            int toRead = Math.Min(count, length - sourceOffset);
            stream.Seek(spillOffsets[pageIndex] + 8 + sourceOffset, SeekOrigin.Begin);
            int total = 0;
            while (total < toRead)
            {
                int read = await stream.ReadAsync(destination.AsMemory(destinationOffset + total, toRead - total), cancellationToken).ConfigureAwait(false);
                if (read == 0) throw new InvalidDataException("Temporary page content ended unexpectedly.");
                total += read;
                cancellationToken.ThrowIfCancellationRequested();
            }
            return total;
        }
        MemoryPage page = memory[pageIndex];
        if (sourceOffset >= page.Length) return 0;
        int remaining = Math.Min(count, page.Length - sourceOffset), copied = 0, segmentOffset = 0;
        foreach (byte[] segment in page.Segments)
        {
            int segmentEnd = segmentOffset + segment.Length;
            if (segmentEnd > sourceOffset && copied < remaining)
            {
                int fromSegment = Math.Max(0, sourceOffset - segmentOffset);
                int take = Math.Min(segment.Length - fromSegment, remaining - copied);
                Buffer.BlockCopy(segment, fromSegment, destination, destinationOffset + copied, take);
                copied += take;
            }
            segmentOffset = segmentEnd;
        }
        return copied;
    }

    private async ValueTask<int> ReadRecordLengthAsync(Stream stream, int pageIndex, CancellationToken token)
    {
        stream.Seek(spillOffsets[pageIndex], SeekOrigin.Begin);
        byte[] prefix = new byte[8];
        int read = 0;
        while (read < prefix.Length)
        {
            int count = await stream.ReadAsync(prefix.AsMemory(read), token).ConfigureAwait(false);
            if (count == 0) throw new InvalidDataException("Temporary page content ended unexpectedly.");
            read += count;
            token.ThrowIfCancellationRequested();
        }
        long length = BitConverter.ToInt64(prefix);
        if (length < 0 || length > int.MaxValue) throw new InvalidDataException("Temporary page content length is invalid.");
        return (int)length;
    }

    private async ValueTask EscalateAsync(CancellationToken token)
    {
        if (temporaryStorage is null)
            throw new OoxPdfLimitExceededException("PDF content exceeded its resident window without a temporary-storage provider.");
        token.ThrowIfCancellationRequested();
        Stream created = await temporaryStorage.CreateAsync(OoxConversionBudget.Current?.Memory, token).ConfigureAwait(false)
            ?? throw new InvalidOperationException("The temporary-storage provider returned no stream.");
        try
        {
            if (!created.CanRead || !created.CanWrite || !created.CanSeek || created.Length != 0 || created.Position != 0)
                throw new InvalidOperationException("Temporary storage must be an empty readable/writable/seekable stream positioned at zero.");
            token.ThrowIfCancellationRequested();
        }
        catch
        {
            await created.DisposeAsync().ConfigureAwait(false);
            throw;
        }
        spillStream = created;
        foreach (MemoryPage buffered in memory)
        {
            token.ThrowIfCancellationRequested();
            spillOffsets.Add(created.Position);
            await created.WriteAsync(BitConverter.GetBytes((long)buffered.Length), token).ConfigureAwait(false);
            foreach (byte[] segment in buffered.Segments)
                await created.WriteAsync(segment, token).ConfigureAwait(false);
            SpilledBytes += buffered.Length;
            buffered.Segments.Clear();
            residentBytes -= buffered.Length;
            OoxMemory.Resize(storage, checked(retainedResourceBytes + residentBytes + (scratch?.Length ?? 0)));
        }
        memory.Clear();
        bufferedBytes = 0;
        if (pendingPage is { Segments.Count: > 0 })
        {
            pendingOffset = created.Position;
            await created.WriteAsync(LengthPlaceholder, token).ConfigureAwait(false);
            foreach (byte[] segment in pendingPage.Segments)
            {
                token.ThrowIfCancellationRequested();
                await created.WriteAsync(segment, token).ConfigureAwait(false);
                SpilledBytes += segment.Length;
            }
            pendingPage.Segments.Clear();
            residentBytes -= pendingLength;
            pendingRecordStarted = true;
        }
        OoxMemory.Resize(storage, checked(retainedResourceBytes + residentBytes + (scratch?.Length ?? 0)));
        token.ThrowIfCancellationRequested();
        diagnosticSink?.Invoke(new OoxPdfDiagnostic("PDF_PAGE_CONTENT_SPILLED", OoxPdfSeverity.Info,
            "Page content exceeded the resident window and moved to temporary storage.", PartName: null, SlideIndex: null,
            PageIndex: null, Feature: "page-content-spill", Fallback: "temporary-storage"));
    }

    public void Dispose() => DisposeAsync().GetAwaiter().GetResult();

    public async ValueTask DisposeAsync()
    {
        if (disposed) return;
        disposed = true;
        memory.Clear();
        pendingPage?.Segments.Clear(); pendingPage = null;
        spillOffsets.Clear(); bufferedBytes = residentBytes = retainedResourceBytes = 0;
        scratch = null;
        try
        {
            if (spillStream is { } stream) await stream.DisposeAsync().ConfigureAwait(false);
        }
        finally
        {
            spillStream = null;
            storage?.Dispose();
        }
    }
}
