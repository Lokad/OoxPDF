using System.Globalization;
using System.Text;
using System.Xml.Linq;

using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static void DrawHighlightRectangle(PdfGraphicsBuilder graphics, PdfEmbeddedFont embedded, TextRun run,
        (RgbColor Color, double Alpha) highlight, double baselineY, double lineWidth)
    {
        if (!TryGetHighlightRectangle(embedded, run, baselineY, lineWidth, out TextHighlightRectangle rectangle))
        {
            return;
        }

        graphics.SaveState();
        PrepareTextPaint(graphics, run);
        SetHighlightPaint(graphics, highlight);
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
            PptxTextMetricRules.StrikeY(glyphRun.BaselineY, run.FontSize),
            glyphRun.Width,
            PptxTextMetricRules.StrikeThickness(embedded, run.FontSize));
        return glyphRun.Width > PptxTextMetricRules.TextStateTolerance && rectangle.Height > 0d;
    }

    // RV01: fallback decoration and highlight rectangles from diagnosed constants
    // (no font tables available). Widths come from measured layout geometry.
    private static void FillFallbackTextDecoration(PdfGraphicsBuilder graphics, double x, double baselineY, double width, double fontSize, bool underline, bool strike)
    {
        if (width <= 0d || (!underline && !strike))
        {
            return;
        }

        if (underline)
        {
            double thickness = fontSize * PdfFallbackFont.UnderlineThicknessEm;
            graphics.FillRectangle(x, baselineY + fontSize * PdfFallbackFont.UnderlinePositionEm - thickness / 2d, width, thickness);
        }

        if (strike)
        {
            double thickness = fontSize * PdfFallbackFont.StrikeoutThicknessEm;
            graphics.FillRectangle(x, baselineY + fontSize * PdfFallbackFont.StrikeoutPositionEm - thickness / 2d, width, thickness);
        }
    }

    private static void FillFallbackHighlight(PdfGraphicsBuilder graphics, TextRun run,
        (RgbColor Color, double Alpha) highlight, double baselineY, double lineWidth)
    {
        if (!BaselineIntersectsClip(run, baselineY))
        {
            return;
        }

        double descent = run.FontSize * PdfFallbackFont.DescentEm;
        double height = run.FontSize * (PdfFallbackFont.AscentEm + PdfFallbackFont.DescentEm);
        graphics.SaveState();
        PrepareTextPaint(graphics, run);
        SetHighlightPaint(graphics, highlight);
        graphics.FillRectangleEvenOdd(run.X, baselineY - descent, lineWidth, height);
        graphics.RestoreState();
    }

    private static void SetHighlightPaint(PdfGraphicsBuilder graphics, (RgbColor Color, double Alpha) paint)
    {
        if (paint.Alpha < 1d) graphics.SetAlpha(paint.Alpha, 1d);
        graphics.SetFillRgb(paint.Color.Red, paint.Color.Green, paint.Color.Blue);
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

    private static void PrepareTextPaint(PdfGraphicsBuilder graphics, TextRun run)
    {
        // Rotated runs carry a device-space clip. Establish it before changing
        // coordinates so an inverse bounding box cannot enlarge its footprint.
        // Preserve the existing local-clip convention for unrotated flips.
        bool deviceClip = TextClipUsesPhysicalCoordinates(run);
        if (deviceClip) graphics.ClipRectangleEvenOdd(run.ClipX, run.ClipY, run.ClipWidth, run.ClipHeight);
        if (HasTextTransform(run)) ApplyTextTransform(graphics, run);
        if (!deviceClip) graphics.ClipRectangleEvenOdd(run.ClipX, run.ClipY, run.ClipWidth, run.ClipHeight);
    }

    private static bool TextClipUsesPhysicalCoordinates(TextRun run) =>
        Math.Abs(run.RotationDegrees) > PptxTextMetricRules.TextStateTolerance;

    private static (double A, double B, double C, double D, double E, double F) TextTransformMatrix(TextRun run) =>
        CenteredTransformMatrix(run.RotationDegrees, run.RotationCenterX, run.RotationCenterY, run.FlipHorizontal, run.FlipVertical);
}
