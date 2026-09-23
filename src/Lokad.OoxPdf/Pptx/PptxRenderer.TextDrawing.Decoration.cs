using System.Globalization;
using System.Text;
using System.Xml.Linq;

using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static void DrawHighlightRectangle(PdfGraphicsBuilder graphics, PdfEmbeddedFont embedded, TextRun run, RgbColor highlight, double baselineY, double lineWidth)
    {
        if (!TryGetHighlightRectangle(embedded, run, baselineY, lineWidth, out TextHighlightRectangle rectangle))
        {
            return;
        }

        graphics.SaveState();
        if (HasTextTransform(run))
        {
            ApplyTextTransform(graphics, run);
        }

        graphics.ClipRectangleEvenOdd(run.ClipX, run.ClipY, run.ClipWidth, run.ClipHeight);
        graphics.SetFillRgb(highlight.Red, highlight.Green, highlight.Blue);
        graphics.FillRectangleEvenOdd(rectangle.X, rectangle.Y, rectangle.Width, rectangle.Height);
        graphics.RestoreState();
    }

    private static bool TryGetHighlightRectangle(PdfEmbeddedFont embedded, TextRun run, double baselineY, double lineWidth, out TextHighlightRectangle rectangle)
    {
        rectangle = default;
        if (!BaselineIntersectsClip(run, baselineY))
        {
            return false;
        }

        double fontScale = run.FontSize / embedded.Font.UnitsPerEm;
        double highlightDescent = PptxTextMetricRules.HighlightDescent(embedded, run.FontSize, fontScale);
        double highlightHeight = PptxTextMetricRules.HighlightHeight(embedded, run.FontSize, fontScale);
        double highlightY = baselineY - highlightDescent;
        rectangle = new TextHighlightRectangle(run.X, highlightY, lineWidth, highlightHeight);
        return true;
    }

    private static bool TryGetUnderlineRectangle(PdfEmbeddedFont embedded, TextGlyphRun glyphRun, out TextDecorationRectangle rectangle)
    {
        TextRun run = glyphRun.Source;
        double underlineScale = run.FontSize / embedded.Font.UnitsPerEm;
        double underlineThickness = PptxTextMetricRules.UnderlineThickness(embedded, run.FontSize);
        double underlineTopY = glyphRun.BaselineY + embedded.Font.Post.UnderlinePosition * underlineScale;
        double underlineY = underlineTopY - underlineThickness;
        rectangle = new TextDecorationRectangle(glyphRun.X, underlineY, glyphRun.Width, underlineThickness);
        return glyphRun.Width > PptxTextMetricRules.TextStateTolerance && underlineThickness > 0d;
    }

    private static bool TryGetStrikeRectangle(PdfEmbeddedFont embedded, TextGlyphRun glyphRun, out TextDecorationRectangle rectangle)
    {
        TextRun run = glyphRun.Source;
        rectangle = new TextDecorationRectangle(
            glyphRun.X,
            PptxTextMetricRules.StrikeY(embedded, glyphRun.BaselineY, run.FontSize),
            glyphRun.Width,
            PptxTextMetricRules.StrikeThickness(embedded, run.FontSize));
        return glyphRun.Width > PptxTextMetricRules.TextStateTolerance && rectangle.Height > 0d;
    }

    private static void FillTextDecorationRectangleEvenOdd(PdfGraphicsBuilder graphics, TextDecorationRectangle rectangle)
    {
        double x = rectangle.X;
        double midX = rectangle.X + rectangle.Width / 2d;
        double right = rectangle.X + rectangle.Width;
        double bottom = rectangle.Y;
        double top = rectangle.Y + rectangle.Height;

        graphics.MoveTo(x, top);
        graphics.LineTo(midX, top);
        graphics.LineTo(right, top);
        graphics.LineTo(right, bottom);
        graphics.LineTo(midX, bottom);
        graphics.LineTo(x, bottom);
        graphics.ClosePath();
        graphics.FillCurrentPathEvenOdd();
    }

    private static bool BaselineInsideVerticalClip(TextRun run, double baselineY)
    {
        if (HasTextTransform(run))
        {
            return true;
        }

        return baselineY >= run.ClipY - PptxTextMetricRules.TextStateTolerance &&
            baselineY <= run.ClipY + run.ClipHeight + PptxTextMetricRules.TextStateTolerance;
    }

    private static bool BaselineIntersectsClip(TextRun run, double baselineY)
    {
        if (HasTextTransform(run))
        {
            return true;
        }

        if (!run.StrictClip)
        {
            return true;
        }

        return baselineY + run.FontSize >= run.ClipY - PptxTextMetricRules.TextStateTolerance &&
            baselineY - run.FontSize <= run.ClipY + run.ClipHeight + PptxTextMetricRules.TextStateTolerance;
    }

    private static bool HasTextTransform(TextRun run)
    {
        return Math.Abs(run.RotationDegrees) > PptxTextMetricRules.TextStateTolerance ||
            run.FlipHorizontal ||
            run.FlipVertical;
    }

    private static void ApplyTextTransform(PdfGraphicsBuilder graphics, TextRun run)
    {
        (double a, double b, double c, double d, double e, double f) = TextTransformMatrix(run);
        graphics.Transform(a, b, c, d, e, f);
    }

    // Affine matrix matching ApplyTextTransform, reused so link areas cover the
    // same transformed glyph positions the emitter paints (S08).
    // Clip rects are authored in device space; inside a rotated text transform
    // they must be mapped back through the inverse rotation, otherwise rotated
    // text near the frame edge is cut (right-side vertical axis titles lost all
    // but their first glyphs). Flip-only runs keep legacy behavior.
    private static (double X, double Y, double Width, double Height) InverseTransformClip(TextRun run, double clipX, double clipY, double clipWidth, double clipHeight)
    {
        if (Math.Abs(run.RotationDegrees) <= PptxTextMetricRules.TextStateTolerance)
        {
            return (clipX, clipY, clipWidth, clipHeight);
        }

        (double a, double b, double c, double d, double e, double f) = TextTransformMatrix(run);
        double det = a * d - b * c;
        double MapX(double x, double y) => (d * (x - e) - c * (y - f)) / det;
        double MapY(double x, double y) => (a * (y - f) - b * (x - e)) / det;
        double x0 = MapX(clipX, clipY);
        double y0 = MapY(clipX, clipY);
        double x1 = MapX(clipX + clipWidth, clipY);
        double y1 = MapY(clipX + clipWidth, clipY);
        double x2 = MapX(clipX, clipY + clipHeight);
        double y2 = MapY(clipX, clipY + clipHeight);
        double x3 = MapX(clipX + clipWidth, clipY + clipHeight);
        double y3 = MapY(clipX + clipWidth, clipY + clipHeight);
        double minX = Math.Min(Math.Min(x0, x1), Math.Min(x2, x3));
        double maxX = Math.Max(Math.Max(x0, x1), Math.Max(x2, x3));
        double minY = Math.Min(Math.Min(y0, y1), Math.Min(y2, y3));
        double maxY = Math.Max(Math.Max(y0, y1), Math.Max(y2, y3));
        return (minX, minY, maxX - minX, maxY - minY);
    }

    private static (double A, double B, double C, double D, double E, double F) TextTransformMatrix(TextRun run)
    {
        double radians = -run.RotationDegrees * Math.PI / 180d;
        double sx = run.FlipHorizontal ? -1d : 1d;
        double sy = run.FlipVertical ? -1d : 1d;
        double cos = Math.Cos(radians);
        double sin = Math.Sin(radians);
        double a = cos * sx;
        double b = sin * sx;
        double c = -sin * sy;
        double d = cos * sy;
        double e = run.RotationCenterX - a * run.RotationCenterX - c * run.RotationCenterY;
        double f = run.RotationCenterY - b * run.RotationCenterX - d * run.RotationCenterY;
        return (a, b, c, d, e, f);
    }
}
