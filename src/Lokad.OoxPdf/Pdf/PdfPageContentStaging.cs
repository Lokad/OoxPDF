using Lokad.OoxPdf.Diagnostics;

namespace Lokad.OoxPdf.Pdf;

// R06.3: page-content spill store for staged emission. Content bytes stay in
// memory while they fit the resident window; past the window they escalate to a
// length-prefixed temp file (8-byte little-endian lengths) and memory is released.
// Page order is preserved in both modes, so emission stays byte-identical.
// The temp file uses delete-on-close plus handle disposal, so no registry or
// awkward lifetime plumbing is needed: crashing or cancelling releases the file.
//
// RV11-P1 scratch policy: page payloads are admission-capped before production
// (content bytes are charged per page upstream), so the resident set is at most
// the admitted current page plus bounded scratch. Production encodes and appends
// in fixed chunks (see ChunkByteCount) through one pooled buffer; emission copies
// spill back to the destination through one pooled buffer. Spilled pages are never
// fully re-read: reads address explicit (page, offset, count) ranges. Memory mode
// retains at most the window in page segments; file mode retains only the running
// page being appended.
internal sealed class PdfPageContentStaging : IDisposable
{
    // RV11-P1: single scratch size for chunked encode/append/readback. Large
    // enough to keep file copies efficient, small enough that transient scratch
    // stays flat while page payloads scale.
    internal const int ChunkByteCount = 65536;

    private static readonly byte[] LengthPlaceholder = new byte[8];

    private readonly List<MemoryPage> memory = new();
    private readonly List<long> fileOffsets = new();
    private readonly long windowBytes;
    private readonly Action<OoxPdfDiagnostic>? diagnosticSink;
    private long bufferedBytes;
    private FileStream? spillFile;
    private bool spillNotified;
    private bool disposed;
    private MemoryPage? pendingPage;
    private long pendingLength;
    private long pendingFileOffset = -1;
    private bool pendingRecordStarted;

    private sealed class MemoryPage
    {
        public readonly List<byte[]> Segments = new();
        public int Length;
    }

    public PdfPageContentStaging(long windowBytes, Action<OoxPdfDiagnostic>? diagnosticSink)
    {
        if (windowBytes < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(windowBytes));
        }

        this.windowBytes = windowBytes;
        this.diagnosticSink = diagnosticSink;
    }

    public int Count => memory.Count + fileOffsets.Count;

    public long SpilledBytes { get; private set; }

    public bool IsFileBacked => spillFile is not null;

    internal string? SpillPathForTests { get; private set; }

    public void AddPage(byte[] contentBytes, CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(contentBytes);
        BeginPage(cancellationToken);
        AppendPageBytes(contentBytes, 0, contentBytes.Length, cancellationToken);
        EndPage(cancellationToken);
    }

    // RV11-P1: chunked page production. A page is opened, fed arbitrary-sized
    // chunks, then closed; closing commits it to the current backing mode and
    // makes it visible to reads. Calls must pair exactly: appending or closing
    // without an open page, or opening twice, fails loudly.
    public void BeginPage(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (pendingPage is not null)
        {
            throw new InvalidOperationException("A staged page is already open.");
        }

        pendingPage = new MemoryPage();
        pendingLength = 0;
        pendingFileOffset = -1;
        pendingRecordStarted = false;
    }

    public void AppendPageBytes(byte[] chunk, int offset, int count, CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(chunk);
        ArgumentOutOfRangeException.ThrowIfNegative(offset);
        ArgumentOutOfRangeException.ThrowIfNegative(count);
        if (offset > chunk.Length - count)
        {
            throw new ArgumentOutOfRangeException(nameof(count));
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (pendingPage is null)
        {
            throw new InvalidOperationException("No staged page is open.");
        }

        if (count == 0)
        {
            return;
        }

        if (spillFile is null && bufferedBytes + pendingLength + count > windowBytes)
        {
            Escalate(cancellationToken);
        }

        if (spillFile is { } file)
        {
            if (!pendingRecordStarted)
            {
                pendingFileOffset = file.Position;
                file.Write(LengthPlaceholder, 0, LengthPlaceholder.Length);
                pendingRecordStarted = true;
            }

            file.Write(chunk, offset, count);
            pendingLength += count;
            SpilledBytes += count;
        }
        else
        {
            var copy = new byte[count];
            Buffer.BlockCopy(chunk, offset, copy, 0, count);
            pendingPage.Segments.Add(copy);
            pendingLength += count;
        }
    }

    public void EndPage(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (pendingPage is null)
        {
            throw new InvalidOperationException("No staged page is open.");
        }

        if (spillFile is { } file)
        {
            if (!pendingRecordStarted)
            {
                // A file-backed page fed no bytes still needs its record so
                // page indexes stay aligned across modes.
                pendingFileOffset = file.Position;
                file.Write(LengthPlaceholder, 0, LengthPlaceholder.Length);
                pendingRecordStarted = true;
            }

            fileOffsets.Add(pendingFileOffset);
            long end = file.Position;
            file.Seek(pendingFileOffset, SeekOrigin.Begin);
            file.Write(BitConverter.GetBytes(pendingLength));
            file.Seek(end, SeekOrigin.Begin);
        }
        else
        {
            pendingPage.Length = checked((int)pendingLength);
            memory.Add(pendingPage);
            bufferedBytes += pendingLength;
        }

        pendingPage = null;
        pendingLength = 0;
        pendingFileOffset = -1;
        pendingRecordStarted = false;
    }

    public int GetPageLength(int pageIndex, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (pendingPage is not null)
        {
            throw new InvalidOperationException("Page production is still open.");
        }

        if (spillFile is { } file)
        {
            return ReadRecordLength(file, pageIndex);
        }

        return memory[pageIndex].Length;
    }

    // RV11-P1: bounded partial read. Copies at most count bytes at sourceOffset
    // into destination, returning the bytes copied (0 past the end). Emission
    // re-reads spilled pages through this in fixed chunks, so a tiny window
    // never triggers a whole-page allocation.
    public int ReadPageBytes(int pageIndex, byte[] destination, int destinationOffset, int sourceOffset, int count, CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(destination);
        ArgumentOutOfRangeException.ThrowIfNegative(destinationOffset);
        ArgumentOutOfRangeException.ThrowIfNegative(sourceOffset);
        ArgumentOutOfRangeException.ThrowIfNegative(count);
        if (destinationOffset > destination.Length - count)
        {
            throw new ArgumentOutOfRangeException(nameof(count));
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (pendingPage is not null)
        {
            throw new InvalidOperationException("Page production is still open.");
        }

        if (count == 0)
        {
            return 0;
        }

        if (spillFile is { } file)
        {
            int length = ReadRecordLength(file, pageIndex, cancellationToken);
            if (sourceOffset >= length)
            {
                return 0;
            }

            int toRead = Math.Min(count, Math.Min(destination.Length - destinationOffset, length - sourceOffset));
            if (toRead <= 0)
            {
                return 0;
            }

            file.Seek(fileOffsets[pageIndex] + 8 + sourceOffset, SeekOrigin.Begin);
            int total = 0;
            while (total < toRead)
            {
                cancellationToken.ThrowIfCancellationRequested();
                int read = file.Read(destination, destinationOffset + total, toRead - total);
                if (read == 0)
                {
                    throw new InvalidDataException("Page content spill file ended unexpectedly.");
                }

                total += read;
            }

            return total;
        }

        MemoryPage page = memory[pageIndex];
        if (sourceOffset >= page.Length)
        {
            return 0;
        }

        int remaining = Math.Min(count, Math.Min(destination.Length - destinationOffset, page.Length - sourceOffset));
        int copied = 0;
        int segmentOffset = 0;
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

    public byte[] GetPageBytes(int pageIndex, CancellationToken cancellationToken)
    {
        int length = GetPageLength(pageIndex, cancellationToken);
        byte[] bytes = new byte[length];
        int offset = 0;
        while (offset < length)
        {
            int read = ReadPageBytes(pageIndex, bytes, offset, offset, length - offset, cancellationToken);
            if (read == 0)
            {
                throw new InvalidDataException("Page content spill store ended unexpectedly.");
            }

            offset += read;
        }

        return bytes;
    }

    private int ReadRecordLength(FileStream file, int pageIndex, CancellationToken cancellationToken = default)
    {
        file.Seek(fileOffsets[pageIndex], SeekOrigin.Begin);
        Span<byte> prefix = stackalloc byte[8];
        ReadExactly(file, prefix, cancellationToken);
        return checked((int)BitConverter.ToInt64(prefix));
    }

    private void Escalate(CancellationToken cancellationToken)
    {
        string path = Path.Combine(Path.GetTempPath(), "OoxPdfPages-" + Guid.NewGuid().ToString("N") + ".tmp");
        FileStream file = new(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, 4096, FileOptions.SequentialScan | FileOptions.DeleteOnClose);
        spillFile = file;
        SpillPathForTests = path;
        foreach (MemoryPage buffered in memory)
        {
            cancellationToken.ThrowIfCancellationRequested();
            fileOffsets.Add(file.Position);
            file.Write(BitConverter.GetBytes((long)buffered.Length));
            foreach (byte[] segment in buffered.Segments)
            {
                file.Write(segment, 0, segment.Length);
            }

            SpilledBytes += buffered.Length;
        }

        memory.Clear();
        bufferedBytes = 0;
        if (pendingPage is not null && pendingPage.Segments.Count > 0)
        {
            // A page open across escalation keeps one record: flush its
            // buffered segments, then keep appending to the file.
            pendingFileOffset = file.Position;
            file.Write(LengthPlaceholder, 0, LengthPlaceholder.Length);
            foreach (byte[] segment in pendingPage.Segments)
            {
                cancellationToken.ThrowIfCancellationRequested();
                file.Write(segment, 0, segment.Length);
                SpilledBytes += segment.Length;
            }

            pendingPage.Segments.Clear();
            pendingRecordStarted = true;
        }
        if (!spillNotified)
        {
            spillNotified = true;
            diagnosticSink?.Invoke(new OoxPdfDiagnostic("PDF_PAGE_CONTENT_SPILLED", OoxPdfSeverity.Info, "Page content exceeded the resident window and spilled to a temp file.", PartName: null, SlideIndex: null, PageIndex: null, Feature: "page-content-spill", Fallback: "temp-file"));
        }
    }

    private static void ReadExactly(Stream stream, Span<byte> destination, CancellationToken cancellationToken = default)
    {
        int totalRead = 0;
        while (totalRead < destination.Length)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int read = stream.Read(destination.Slice(totalRead));
            if (read == 0)
            {
                throw new InvalidDataException("Page content spill file ended unexpectedly.");
            }

            totalRead += read;
        }
    }

    public void Dispose()
    {
        if (disposed)
        {
            return;
        }

        disposed = true;
        spillFile?.Dispose();
        spillFile = null;
    }
}
