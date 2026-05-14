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
    {
        Width = width;
        Height = height;
        Content = content;
        Fonts = fonts;
    }

    public double Width { get; }

    public double Height { get; }

    public string Content { get; }

    public IReadOnlyList<PdfFontResource> Fonts { get; }
}
