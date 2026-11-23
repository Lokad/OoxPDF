using System.Xml.Linq;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static GroupTransform ReadGroupTransform(XElement group)
    {
        return ToGroupTransform(PptxSceneBuilder.ReadGroupTransform(group));
    }

    private static GroupTransform ToGroupTransform(PptxSceneGroupTransform group)
    {
        return new GroupTransform(
            group.OffsetX,
            group.OffsetY,
            group.Width,
            group.Height,
            group.ChildOffsetX,
            group.ChildOffsetY,
            group.ScaleX,
            group.ScaleY,
            group.RotationDegrees,
            group.FlipHorizontal,
            group.FlipVertical);
    }

    private static void ApplyShapeTransform(PdfGraphicsBuilder graphics, double x, double y, double width, double height, ShapeBounds bounds)
    {
        var (a, b, c, d, e, f) = ShapeTransformMatrix(x, y, width, height, bounds);
        graphics.Transform(a, b, c, d, e, f);
    }

    private static (double A, double B, double C, double D, double E, double F) ShapeTransformMatrix(
        double x, double y, double width, double height, ShapeBounds bounds) =>
        CenteredTransformMatrix(bounds.RotationDegrees, x + width / 2d, y + height / 2d, bounds.FlipHorizontal, bounds.FlipVertical);

    private static void StrokeShapePatternFill(
        PdfGraphicsBuilder graphics,
        string preset,
        double x,
        double y,
        double width,
        double height,
        PptxScenePatternFill fill,
        IReadOnlyDictionary<string, double>? presetAdjustments)
    {
        if (width <= 0d || height <= 0d)
        {
            return;
        }

        graphics.SaveState();
        ClipToPresetShape(graphics, preset, x, y, width, height, presetAdjustments);
        graphics.SetStrokeRgb(fill.Foreground.Red, fill.Foreground.Green, fill.Foreground.Blue);
        graphics.SetLineWidth(IsDarkDiagonalPatternFill(fill.Preset) ? 1.0d : 0.5d);
        double spacing = IsDarkDiagonalPatternFill(fill.Preset) ? 4d : 5d;
        bool up = fill.Preset.Contains("UpDiag", StringComparison.OrdinalIgnoreCase);
        for (double offset = -height; offset <= width + height; offset += spacing)
        {
            if (up)
            {
                graphics.StrokeLine(x + offset, y, x + offset + height, y + height);
            }
            else
            {
                graphics.StrokeLine(x + offset, y + height, x + offset + height, y);
            }
        }

        graphics.RestoreState();
    }

    private static bool IsDarkDiagonalPatternFill(string preset)
    {
        return preset.StartsWith("dk", StringComparison.OrdinalIgnoreCase);
    }

    private static (RgbColor Color, double Alpha)? ReadShapeFontPaint(XElement shape, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? fontRef = shape
            .Element(PresentationNamespace + "style")
            ?.Element(DrawingNamespace + "fontRef");
        return PptxColorResolver.TryReadSolidColorWithAlpha(fontRef, theme, colorMap, out RgbColor color, out double alpha)
            ? (color, alpha) : null;
    }
}
