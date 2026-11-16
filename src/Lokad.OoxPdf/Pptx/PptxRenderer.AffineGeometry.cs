using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static (double A, double B, double C, double D, double E, double F) CenteredTransformMatrix(
        double rotationDegrees, double centerX, double centerY, bool flipHorizontal, bool flipVertical)
    {
        double radians = -rotationDegrees * Math.PI / 180d;
        double sx = flipHorizontal ? -1d : 1d;
        double sy = flipVertical ? -1d : 1d;
        double cos = Math.Cos(radians);
        double sin = Math.Sin(radians);
        double a = cos * sx;
        double b = sin * sx;
        double c = -sin * sy;
        double d = cos * sy;
        double e = centerX - a * centerX - c * centerY;
        double f = centerY - b * centerX - d * centerY;
        return (a, b, c, d, e, f);
    }

    private static PdfRectangle TransformedRectangleBounds(PdfRectangle rectangle,
        (double A, double B, double C, double D, double E, double F) matrix)
    {
        var (a, b, c, d, e, f) = matrix;
        double x0 = rectangle.X, x1 = rectangle.X + rectangle.Width;
        double y0 = rectangle.Y, y1 = rectangle.Y + rectangle.Height;
        double left = Math.Min(Math.Min(a * x0 + c * y0 + e, a * x1 + c * y0 + e), Math.Min(a * x0 + c * y1 + e, a * x1 + c * y1 + e));
        double right = Math.Max(Math.Max(a * x0 + c * y0 + e, a * x1 + c * y0 + e), Math.Max(a * x0 + c * y1 + e, a * x1 + c * y1 + e));
        double bottom = Math.Min(Math.Min(b * x0 + d * y0 + f, b * x1 + d * y0 + f), Math.Min(b * x0 + d * y1 + f, b * x1 + d * y1 + f));
        double top = Math.Max(Math.Max(b * x0 + d * y0 + f, b * x1 + d * y0 + f), Math.Max(b * x0 + d * y1 + f, b * x1 + d * y1 + f));
        return new PdfRectangle(left, bottom, right - left, top - bottom);
    }
}
