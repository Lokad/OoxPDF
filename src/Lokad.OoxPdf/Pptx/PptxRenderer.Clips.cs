using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static void ClipSlideBoundsEvenOdd(PptxDocument document, PdfGraphicsBuilder graphics)
    {
        graphics.ClipRectangleEvenOdd(0d, 0d, document.SlideWidthPoints, document.SlideHeightPoints);
    }

    private static bool TryIntersectWithSlideBounds(
        double x,
        double y,
        double width,
        double height,
        PptxDocument document,
        out double clipX,
        out double clipY,
        out double clipWidth,
        out double clipHeight)
    {
        double minX = Math.Max(0d, x);
        double minY = Math.Max(0d, y);
        double maxX = Math.Min(document.SlideWidthPoints, x + Math.Max(0d, width));
        double maxY = Math.Min(document.SlideHeightPoints, y + Math.Max(0d, height));
        clipX = minX;
        clipY = minY;
        clipWidth = Math.Max(0d, maxX - minX);
        clipHeight = Math.Max(0d, maxY - minY);
        return clipWidth > 0d && clipHeight > 0d;
    }
}
