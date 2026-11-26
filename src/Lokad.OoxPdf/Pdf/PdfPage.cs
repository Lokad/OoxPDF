namespace Lokad.OoxPdf.Pdf;

internal readonly record struct PdfPage
{
    public PdfPage(
        double width,
        double height,
        string content = "",
        IReadOnlyList<PdfFontResource>? fonts = null,
        IReadOnlyList<PdfImageResource>? images = null,
        IReadOnlyList<PdfExtGStateResource>? extGStates = null,
        IReadOnlyList<PdfShadingResource>? shadings = null,
        IReadOnlyList<PdfTilingPatternResource>? patterns = null,
        IReadOnlyList<PdfLinkAnnotation>? annotations = null,
        IReadOnlyList<PdfFallbackFontResource>? fallbackFonts = null,
        IReadOnlyList<PdfTransparencyGroupResource>? groups = null,
        IReadOnlyList<PdfShadingPatternResource>? shadingPatterns = null)
    {
        Width = width;
        Height = height;
        Content = content;
        Fonts = fonts ?? [];
        Images = images ?? [];
        ExtGStates = extGStates ?? [];
        Shadings = shadings ?? [];
        Patterns = patterns ?? [];
        Annotations = annotations ?? [];
        FallbackFonts = fallbackFonts ?? [];
        Groups = groups ?? [];
        ShadingPatterns = shadingPatterns ?? [];
    }

    public double Width { get; }

    public double Height { get; }

    public string Content { get; init; }

    public IReadOnlyList<PdfFontResource> Fonts { get; }

    public IReadOnlyList<PdfFallbackFontResource> FallbackFonts { get; }

    public IReadOnlyList<PdfImageResource> Images { get; }

    public IReadOnlyList<PdfExtGStateResource> ExtGStates { get; }

    public IReadOnlyList<PdfShadingResource> Shadings { get; }

    public IReadOnlyList<PdfTilingPatternResource> Patterns { get; }

    public IReadOnlyList<PdfShadingPatternResource> ShadingPatterns { get; }

    public IReadOnlyList<PdfTransparencyGroupResource> Groups { get; init; }

    public IReadOnlyList<PdfLinkAnnotation> Annotations { get; }
}

internal readonly record struct PdfLinkDestination(int PageIndex, double? Left, double? Top, double? Zoom);

// The factories set exactly one target. Serialization also rejects the default
// value, which has no target.
internal readonly record struct PdfLinkAnnotation
{
    public double X { get; }

    public double Y { get; }

    public double Width { get; }

    public double Height { get; }

    public string? Uri { get; }

    public PdfLinkDestination? Destination { get; }

    private PdfLinkAnnotation(double x, double y, double width, double height, string? uri, PdfLinkDestination? destination)
    {
        X = x;
        Y = y;
        Width = width;
        Height = height;
        Uri = uri;
        Destination = destination;
    }

    internal static PdfLinkAnnotation ToUri(double x, double y, double width, double height, string uri)
    {
        ArgumentException.ThrowIfNullOrEmpty(uri);
        return new PdfLinkAnnotation(x, y, width, height, uri, null);
    }

    internal static PdfLinkAnnotation ToDestination(double x, double y, double width, double height, PdfLinkDestination destination)
    {
        return new PdfLinkAnnotation(x, y, width, height, null, destination);
    }
}
