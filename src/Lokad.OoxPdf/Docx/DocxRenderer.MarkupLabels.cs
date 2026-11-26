using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    private static string? FirstNonEmpty(params string?[] values)
    {
        foreach (string? value in values)
        {
            if (!string.IsNullOrWhiteSpace(value))
            {
                return value.Trim();
            }
        }

        return null;
    }

    internal static string? FormatCommentDate(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        return DateTimeOffset.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out DateTimeOffset date)
            ? date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
            : value.Trim();
    }

    private static string TrimBalloonText(string text, double marginRight)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return string.Empty;
        }

        int maxLength = Math.Max(8, (int)Math.Floor(Math.Max(42d, marginRight - 8d) / 2.8d));
        string normalized = text.Trim();
        if (normalized.Length <= maxLength) return normalized;

        int prefixLimit = maxLength - 3;
        int end = 0;
        while (end < prefixLimit)
        {
            int next = end + StringInfo.GetNextTextElementLength(normalized.AsSpan(end));
            if (next > prefixLimit) break;
            end = next;
        }

        return normalized[..end] + "...";
    }

    // Single source for the Word-compatible balloon body wrap shared by text
    // rendering and row-derived balloon heights: the title plus the first body
    // words share the top row, and an overflowing body continues below it.
    private static void ComputeWordCompatibleBalloonWrapWidths(
        double titleWidth,
        double balloonWidth,
        double printScale,
        out double firstLineWidth,
        out double continuationWidth)
    {
        double bodyFirstLineX = ResolveWordCompatibleBalloonTextInsetX(printScale) + titleWidth;
        double rightEdge = balloonWidth - 0.5d;
        firstLineWidth = Math.Max(0d, rightEdge - bodyFirstLineX - WordCompatibleAllMarkupBalloonBodyFirstLineTrailingPadPoints);
        continuationWidth = Math.Max(0d, balloonWidth - 2d * WordCompatibleAllMarkupBalloonTextInsetDesignPoints * printScale);
    }

    private static void DrawBalloonText(
        PdfGraphicsBuilder graphics,
        DocxRunFontResource? resource,
        string text,
        double x,
        double baselineY,
        double fontSize,
        byte red,
        byte green,
        byte blue,
        PdfFallbackFontResource? fallback = null)
    {
        if (resource is null)
        {
            // RV06 (RV01 residual): balloon text without any embeddable face renders
            // with the diagnosed standard-14 fallback instead of vanishing.
            if (fallback is not null && text.Length != 0)
            {
                DrawFallbackRunGlyphText(graphics, fallback, text, x, baselineY, new RgbColor(red, green, blue), fontSize, 0d);
            }

            return;
        }
        // Office balloon text carries no tracking: the shared encoder applies pair
        // kerning like the body path and yields plain glyph shows when the face has
        // no kerns for the line.
        var (operand, isPositioned) = resource.Embedded.EncodeGlyphText(
            text,
            0d,
            fontSize,
            forcePositioningArray: false,
            kerningEnabled: true);
        if (operand is null) return;
        graphics.DrawGlyphText(resource.Name, fontSize, x, baselineY, red, green, blue, operand,
            italic: false, characterSpacing: 0d, textRenderingMode: 0, strokeRed: 0, strokeGreen: 0,
            strokeBlue: 0, strokeWidth: 0d, isPositioned: isPositioned);
    }

    private static double GetSegmentFontSize(DocxTextSegmentLayout segment, double lineFontSize)
    {
        return segment.FontSize ?? lineFontSize;
    }

    private static double GetSegmentBaselineY(DocxTextSegmentLayout segment, double lineBaselineY)
    {
        return lineBaselineY + segment.BaselineOffsetY;
    }
}
