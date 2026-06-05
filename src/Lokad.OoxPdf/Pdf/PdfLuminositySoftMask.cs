namespace Lokad.OoxPdf.Pdf;

internal sealed class PdfLuminositySoftMask
{
    public PdfLuminositySoftMask(PdfImageXObject image, double x, double y, double width, double height, double cropLeft, double cropTop, double cropRight, double cropBottom)
    {
        Image = image;
        X = x;
        Y = y;
        Width = width;
        Height = height;
        CropLeft = cropLeft;
        CropTop = cropTop;
        CropRight = cropRight;
        CropBottom = cropBottom;
    }

    public PdfImageXObject Image { get; }

    public double X { get; }

    public double Y { get; }

    public double Width { get; }

    public double Height { get; }

    public double CropLeft { get; }

    public double CropTop { get; }

    public double CropRight { get; }

    public double CropBottom { get; }

    public string ResourceKey => FormattableString.Invariant(
        $"{Image.ResourceKey}:luminosity:{X:0.###}:{Y:0.###}:{Width:0.###}:{Height:0.###}:{CropLeft:0.#####}:{CropTop:0.#####}:{CropRight:0.#####}:{CropBottom:0.#####}");
}
