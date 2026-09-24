namespace Lokad.OoxPdf.Pdf;

// RV19: owning staged-document contract (RV11 seam). Descriptors and their
// content store travel as one value so emission cannot pair pages with the
// wrong store, and disposed documents fail loudly instead of reading a
// disposed store. Only ProduceStagedPages creates instances.
internal sealed class PdfStagedDocument : IDisposable
{
    private readonly PdfPageContentStaging staging;

    private bool disposed;

    internal PdfStagedDocument(IReadOnlyList<PdfPage> pages, PdfPageContentStaging staging)
    {
        Pages = pages;
        this.staging = staging;
    }

    public IReadOnlyList<PdfPage> Pages { get; }

    internal PdfPageContentStaging Staging
    {
        get
        {
            ObjectDisposedException.ThrowIf(disposed, this);
            return staging;
        }
    }

    public void ThrowIfDisposed()
    {
        ObjectDisposedException.ThrowIf(disposed, this);
    }

    public void Dispose()
    {
        if (disposed)
        {
            return;
        }

        disposed = true;
        staging.Dispose();
    }
}
