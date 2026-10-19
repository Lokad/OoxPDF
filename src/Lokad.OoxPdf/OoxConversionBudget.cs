using Lokad.OoxPdf.Diagnostics;

namespace Lokad.OoxPdf;

// One conversion's output counters. Per-site bounds remain independent. The
// AsyncLocal isolates concurrent flows; nesting restores the previous scope.
// Rendering stays single-flow. No cumulative work or retained-memory fuel.
internal sealed class OoxConversionBudget
{
    private static readonly AsyncLocal<OoxConversionBudget?> Storage = new();
    public static OoxConversionBudget? Current => Storage.Value;
    private readonly OoxConversionLimits limits;

    internal OoxConversionBudget(OoxConversionLimits? limits) => this.limits = limits ?? new();
    private long pdfPages;
    public long PdfPages => pdfPages;
    private long pdfContentBytes;
    public long PdfContentBytes => pdfContentBytes;
    private long pdfOutputBytes;
    public long PdfOutputBytes => pdfOutputBytes;
    private long pdfFontBytes;
    public long PdfFontBytes => pdfFontBytes;
    private long pdfImageBytes;
    public long PdfImageBytes => pdfImageBytes;
    public long PageContentSpilledBytes { get; private set; }
    public OoxConversionTotals Totals => new(PdfPages, PdfContentBytes, PdfOutputBytes, PdfFontBytes, PdfImageBytes, PageContentSpilledBytes);

    public void NotePageContentSpilled(long bytes)
    {
        if (bytes < 0) { throw new ArgumentOutOfRangeException(nameof(bytes)); }
        PageContentSpilledBytes = checked(PageContentSpilledBytes + bytes);
    }

    public static Scope BeginScope(OoxConversionLimits? limits)
    {
        var budget = new OoxConversionBudget(limits);
        OoxConversionBudget? prior = Storage.Value;
        Storage.Value = budget;
        return new Scope(budget, prior);
    }

    public void ChargePdfPages(long count) => Charge(count, limits.MaxPagesPerConversion, ref pdfPages, "page budget", "pages");
    public void ChargePdfContentBytes(long count) => Charge(count, limits.MaxPdfContentBytesPerConversion, ref pdfContentBytes, "PDF content byte budget", "bytes");
    public void ChargePdfOutputBytes(long count) => Charge(count, limits.MaxOutputBytesPerConversion, ref pdfOutputBytes, "output byte budget", "bytes");
    public void ChargePdfFontBytes(long count) => Charge(count, limits.MaxPdfFontBytesPerConversion, ref pdfFontBytes, "PDF font byte budget", "bytes");
    public void ChargePdfImageBytes(long count) => Charge(count, limits.MaxPdfImageBytesPerConversion, ref pdfImageBytes, "PDF image byte budget", "bytes");

    private static void Charge(long count, long maximum, ref long total, string resource, string units)
    {
        if (count < 0) { throw new ArgumentOutOfRangeException(nameof(count)); }
        if (count > maximum - total)
        {
            throw new OoxPdfLimitExceededException($"Conversion exceeds the {resource} of {maximum} {units}.");
        }
        total += count;
    }

    internal void ThrowIfLayoutPagesExceedBudget(int constructedPages)
    {
        if (constructedPages > limits.MaxPagesPerConversion - PdfPages)
        {
            throw new OoxPdfLimitExceededException("Conversion exceeds the page budget of " + limits.MaxPagesPerConversion.ToString(System.Globalization.CultureInfo.InvariantCulture) + " pages during layout pagination.");
        }
    }

    internal sealed class Scope(OoxConversionBudget budget, OoxConversionBudget? prior) : IDisposable
    {
        private bool disposed;
        public OoxConversionBudget Budget { get; } = budget;
        public void Dispose()
        {
            if (disposed) { return; }
            disposed = true;
            Storage.Value = prior;
        }
    }
}

internal readonly record struct OoxConversionTotals(long PdfPages, long PdfContentBytes, long PdfOutputBytes,
    long PdfFontBytes, long PdfImageBytes, long PageContentSpilledBytes)
{
    public OoxPdfDiagnostic ToSummaryDiagnostic(int pageCount) => new(
        "CONVERSION_RESOURCE_SUMMARY", OoxPdfSeverity.Info, FormattableString.Invariant(
            $"Conversion output totals: pages={pageCount}; pdfPages={PdfPages}; pdfContentBytes={PdfContentBytes}; pdfOutputBytes={PdfOutputBytes}; pdfFontBytes={PdfFontBytes}; pdfImageBytes={PdfImageBytes}; pageContentSpilledBytes={PageContentSpilledBytes}."),
        PartName: null, SlideIndex: null, PageIndex: null, Feature: null, Fallback: null);
}
