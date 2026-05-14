namespace Lokad.OoxPdf.Pdf;

internal readonly record struct PdfPage
{
    public PdfPage(double width, double height)
        : this(width, height, string.Empty)
    {
    }

    public PdfPage(double width, double height, string content)
    {
        Width = width;
        Height = height;
        Content = content;
    }

    public double Width { get; }

    public double Height { get; }

    public string Content { get; }
}
