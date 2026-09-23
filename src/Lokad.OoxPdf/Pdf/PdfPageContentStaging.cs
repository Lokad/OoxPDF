using Lokad.OoxPdf.Diagnostics;

namespace Lokad.OoxPdf.Pdf;

// R06.3: page-content spill store for staged emission. Content bytes stay in
// memory while they fit the resident window; past the window they escalate to a
// length-prefixed temp file (8-byte little-endian lengths) and memory is released.
// Page order is preserved in both modes, so emission stays byte-identical.
// The temp file uses delete-on-close plus handle disposal, so no registry or
// awkward lifetime plumbing is needed: crashing or cancelling releases the file.
internal sealed class PdfPageContentStaging : IDisposable
{
    private readonly List<byte[]> memory = new();
    private readonly List<long> fileOffsets = new();
    private readonly long windowBytes;
    private readonly Action<OoxPdfDiagnostic>? diagnosticSink;
    private long bufferedBytes;
    private FileStream? spillFile;
    private bool spillNotified;
    private bool disposed;

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
        cancellationToken.ThrowIfCancellationRequested();
        if (spillFile is null && bufferedBytes + contentBytes.Length > windowBytes)
        {
            Escalate(cancellationToken);
        }

        if (spillFile is { } file)
        {
            fileOffsets.Add(file.Position);
            file.Write(BitConverter.GetBytes((long)contentBytes.Length));
            file.Write(contentBytes, 0, contentBytes.Length);
            SpilledBytes += contentBytes.Length;
        }
        else
        {
            memory.Add(contentBytes);
            bufferedBytes += contentBytes.Length;
        }
    }

    public byte[] GetPageBytes(int pageIndex, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (spillFile is { } file)
        {
            file.Seek(fileOffsets[pageIndex], SeekOrigin.Begin);
            Span<byte> prefix = stackalloc byte[8];
            ReadExactly(file, prefix);
            int length = checked((int)BitConverter.ToInt64(prefix));
            byte[] bytes = new byte[length];
            ReadExactly(file, bytes);
            return bytes;
        }

        return memory[pageIndex];
    }

    private void Escalate(CancellationToken cancellationToken)
    {
        string path = Path.Combine(Path.GetTempPath(), "OoxPdfPages-" + Guid.NewGuid().ToString("N") + ".tmp");
        FileStream file = new(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, 4096, FileOptions.SequentialScan | FileOptions.DeleteOnClose);
        spillFile = file;
        SpillPathForTests = path;
        foreach (byte[] buffered in memory)
        {
            cancellationToken.ThrowIfCancellationRequested();
            fileOffsets.Add(file.Position);
            file.Write(BitConverter.GetBytes((long)buffered.Length));
            file.Write(buffered, 0, buffered.Length);
            SpilledBytes += buffered.Length;
        }

        memory.Clear();
        bufferedBytes = 0;
        if (!spillNotified)
        {
            spillNotified = true;
            diagnosticSink?.Invoke(new OoxPdfDiagnostic("PDF_PAGE_CONTENT_SPILLED", OoxPdfSeverity.Info, "Page content exceeded the resident window and spilled to a temp file.", PartName: null, SlideIndex: null, PageIndex: null, Feature: "page-content-spill", Fallback: "temp-file"));
        }
    }

    private static void ReadExactly(Stream stream, Span<byte> destination)
    {
        int totalRead = 0;
        while (totalRead < destination.Length)
        {
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
