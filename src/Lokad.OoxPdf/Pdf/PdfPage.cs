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
    {
        Width = width;
        Height = height;
        Content = content;
        Fonts = fonts;
        Images = images;
        ExtGStates = extGStates;
        Shadings = shadings;
        Patterns = patterns;
    }

    public double Width { get; }

    public double Height { get; }

    public string Content { get; }

    public IReadOnlyList<PdfFontResource> Fonts { get; }

    public IReadOnlyList<PdfImageResource> Images { get; }

    public IReadOnlyList<PdfExtGStateResource> ExtGStates { get; }

    public IReadOnlyList<PdfShadingResource> Shadings { get; }

    public IReadOnlyList<PdfTilingPatternResource> Patterns { get; }
}
