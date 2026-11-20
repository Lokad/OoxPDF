namespace Lokad.OoxPdf.Pdf;

// A staged document owns its page descriptors and backing store through emission.
internal sealed class PdfStagedDocument : IAsyncDisposable
{
    private readonly PdfPageContentStaging staging;

    private bool disposed;

    internal PdfStagedDocument(IReadOnlyList<PdfPage> pages, PdfPageContentStaging staging, IReadOnlyList<int> pageEntries, IReadOnlyDictionary<PdfTransparencyGroup, int> groupEntries)
    {
        this.pages = pages;
        this.staging = staging;
        PageEntries = pageEntries;
        GroupEntries = groupEntries;
    }

    private IReadOnlyList<PdfPage> pages;
    public IReadOnlyList<PdfPage> Pages
    {
        get
        {
            ObjectDisposedException.ThrowIf(disposed, this);
            return pages;
        }
    }

    internal IReadOnlyList<int> PageEntries { get; private set; }

    internal IReadOnlyDictionary<PdfTransparencyGroup, int> GroupEntries { get; private set; }

    internal PdfPageContentStaging Staging
    {
        get
        {
            ObjectDisposedException.ThrowIf(disposed, this);
            return staging;
        }
    }

    public async ValueTask DisposeAsync()
    {
        if (disposed)
        {
            return;
        }

        disposed = true;
        pages = [];
        PageEntries = [];
        GroupEntries = new Dictionary<PdfTransparencyGroup, int>();
        await staging.DisposeAsync().ConfigureAwait(false);
    }
}
