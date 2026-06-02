namespace Lokad.OoxPdf.Pdf;

internal readonly record struct PdfPage
{
    public PdfPage(double width, double height)
        : this(width, height, string.Empty)
    {
    }

    public PdfPage(double width, double height, string content)
        : this(width, height, content, [])
    {
    }

    public PdfPage(double width, double height, string content, IReadOnlyList<PdfFontResource> fonts)
        : this(width, height, content, fonts, [])
    {
    }

    public PdfPage(double width, double height, string content, IReadOnlyList<PdfFontResource> fonts, IReadOnlyList<PdfImageResource> images)
        : this(width, height, content, fonts, images, [])
    {
    }

    public PdfPage(double width, double height, string content, IReadOnlyList<PdfFontResource> fonts, IReadOnlyList<PdfImageResource> images, IReadOnlyList<PdfExtGStateResource> extGStates)
        : this(width, height, content, fonts, images, extGStates, [])
    {
    }

    public PdfPage(double width, double height, string content, IReadOnlyList<PdfFontResource> fonts, IReadOnlyList<PdfImageResource> images, IReadOnlyList<PdfExtGStateResource> extGStates, IReadOnlyList<PdfShadingResource> shadings)
        : this(width, height, content, fonts, images, extGStates, shadings, [])
    {
    }

    public PdfPage(double width, double height, string content, IReadOnlyList<PdfFontResource> fonts, IReadOnlyList<PdfImageResource> images, IReadOnlyList<PdfExtGStateResource> extGStates, IReadOnlyList<PdfShadingResource> shadings, IReadOnlyList<PdfTilingPatternResource> patterns)
        : this(width, height, content, fonts, images, extGStates, shadings, patterns, [])
    {
    }

    public PdfPage(double width, double height, string content, IReadOnlyList<PdfFontResource> fonts, IReadOnlyList<PdfImageResource> images, IReadOnlyList<PdfExtGStateResource> extGStates, IReadOnlyList<PdfShadingResource> shadings, IReadOnlyList<PdfTilingPatternResource> patterns, IReadOnlyList<PdfLinkAnnotation> annotations)
    {
        Width = width;
        Height = height;
        Content = content;
        Fonts = fonts;
        Images = images;
        ExtGStates = extGStates;
        Shadings = shadings;
        Patterns = patterns;
        Annotations = annotations;
    }

    public double Width { get; }

    public double Height { get; }

    public string Content { get; }

    public IReadOnlyList<PdfFontResource> Fonts { get; }

    public IReadOnlyList<PdfImageResource> Images { get; }

    public IReadOnlyList<PdfExtGStateResource> ExtGStates { get; }

    public IReadOnlyList<PdfShadingResource> Shadings { get; }

    public IReadOnlyList<PdfTilingPatternResource> Patterns { get; }

    public IReadOnlyList<PdfLinkAnnotation> Annotations { get; }
}

internal readonly record struct PdfLinkDestination(int PageIndex, double? Left, double? Top, double? Zoom);

internal readonly record struct PdfLinkAnnotation(
    double X,
    double Y,
    double Width,
    double Height,
    string? Uri,
    PdfLinkDestination? Destination = null)
{
    internal static PdfLinkAnnotation ToUri(double x, double y, double width, double height, string uri)
    {
        return new PdfLinkAnnotation(x, y, width, height, uri, null);
    }

    internal static PdfLinkAnnotation ToDestination(double x, double y, double width, double height, PdfLinkDestination destination)
    {
        return new PdfLinkAnnotation(x, y, width, height, null, destination);
    }
}
