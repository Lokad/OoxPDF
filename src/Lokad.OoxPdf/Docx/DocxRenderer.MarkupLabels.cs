using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
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

    private static string? FormatCommentDate(string? value)
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
        return normalized.Length <= maxLength
            ? normalized
            : normalized[..Math.Max(0, maxLength - 3)] + "...";
    }

    private static string[] WrapWordCompatibleBalloonBody(
        string text,
        PdfEmbeddedFont embedded,
        double fontSize,
        double firstLineWidth,
        double continuationWidth)
    {
        string normalized = Regex.Replace(text.Trim(), @"\s+", " ");
        if (normalized.Length == 0)
        {
            return [];
        }

        string[] words = normalized.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        var lines = new List<string>();
        string firstLine = ConsumeBalloonWords(0, firstLineWidth, out int nextWordIndex);
        if (nextWordIndex >= words.Length)
        {
            return [firstLine];
        }

        // RV06: Word-compatible balloons render full comment text across as many
        // continuation rows as needed instead of truncating after the second row.
        lines.Add(firstLine + " ");
        while (nextWordIndex < words.Length)
        {
            lines.Add(ConsumeBalloonWords(nextWordIndex, continuationWidth, out nextWordIndex));
        }

        return lines.ToArray();

        string ConsumeBalloonWords(int startIndex, double maxWidth, out int nextWordIndex)
        {
            string line = string.Empty;
            int index = startIndex;
            for (; index < words.Length; index++)
            {
                string candidate = line.Length == 0
                    ? words[index]
                    : line + " " + words[index];
                if (line.Length != 0 &&
                    embedded.MeasureTextPoints(candidate, fontSize) > maxWidth)
                {
                    break;
                }

                line = candidate;
            }

            if (line.Length == 0 && startIndex < words.Length)
            {
                line = FitWordCompatibleBalloonLine(words[startIndex], maxWidth, embedded, fontSize);
                index = startIndex + 1;
            }

            nextWordIndex = index;
            return line;
        }
    }

    // Single source for the Word-compatible balloon body wrap shared by text
    // rendering and row-derived balloon heights: the title plus the first body
    // words share the top row, an overflowing body wraps once onto a second row.
    private static void ComputeWordCompatibleBalloonWrapWidths(
        double titleWidth,
        double balloonWidth,
        out double firstLineWidth,
        out double continuationWidth)
    {
        double bodyFirstLineX = WordCompatibleAllMarkupBalloonTextInsetXPoints +
            titleWidth +
            WordCompatibleAllMarkupBalloonBodyFirstLineXOffsetPoints;
        double rightEdge = balloonWidth - 0.5d;
        firstLineWidth = Math.Max(0d, rightEdge - bodyFirstLineX - WordCompatibleAllMarkupBalloonBodyFirstLineTrailingPadPoints);
        continuationWidth = Math.Max(0d, rightEdge - WordCompatibleAllMarkupBalloonTextInsetXPoints);
    }

    private static int CountWordCompatibleBalloonTextRows(
        string body,
        PdfEmbeddedFont bodyEmbedded,
        double fontSize,
        double firstLineWidth,
        double continuationWidth)
    {
        if (string.IsNullOrWhiteSpace(body))
        {
            return 1;
        }

        string[] lines = WrapWordCompatibleBalloonBody(body, bodyEmbedded, fontSize, firstLineWidth, continuationWidth);
        return lines.Length == 0 ? 1 : lines.Length;
    }

    private static string FitWordCompatibleBalloonLine(
        string text,
        double maxWidth,
        PdfEmbeddedFont embedded,
        double fontSize)
    {
        if (text.Length == 0 ||
            maxWidth <= 0d ||
            embedded.MeasureTextPoints(text, fontSize) <= maxWidth)
        {
            return text;
        }

        const string suffix = "...";
        for (int length = Math.Max(0, text.Length - 1); length >= 0; length--)
        {
            string candidate = text[..length].TrimEnd() + suffix;
            if (embedded.MeasureTextPoints(candidate, fontSize) <= maxWidth)
            {
                return candidate;
            }
        }

        return suffix;
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
        DrawBalloonText(graphics, resource, text, x, baselineY, fontSize, red, green, blue, positioningCharacterSpacing: 0d, fallback);
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
        double positioningCharacterSpacing,
        PdfFallbackFontResource? fallback = null)
    {
        if (resource is null)
        {
            // RV06 (RV01 residual): balloon text without any embeddable face renders
            // with the diagnosed standard-14 fallback instead of vanishing.
            if (fallback is not null && text.Length != 0)
            {
                DrawFallbackRunGlyphText(graphics, fallback, text, x, baselineY, new RgbColor(red, green, blue), fontSize, positioningCharacterSpacing);
            }

            return;
        }
        if (Math.Abs(positioningCharacterSpacing) > 0.001d)
        {
            string? positioningArray = resource.Embedded.EncodeGlyphPositioningArray(
                text,
                positioningCharacterSpacing,
                fontSize,
                forcePositioningArray: true,
                kerningEnabled: false);
            if (positioningArray is not null)
            {
                graphics.DrawGlyphPositionedText(resource.Name, fontSize, x, baselineY, red, green, blue, positioningArray, italic: false, characterSpacing: 0d, textRenderingMode: 0, strokeRed: 0, strokeGreen: 0, strokeBlue: 0, strokeWidth: 0d);
            }

            return;
        }

        string glyphHex = resource.Embedded.EncodeGlyphHex(text);
        if (glyphHex.Length != 0)
        {
            graphics.DrawGlyphText(resource.Name, fontSize, x, baselineY, red, green, blue, glyphHex, italic: false, characterSpacing: 0d, textRenderingMode: 0, strokeRed: 0, strokeGreen: 0, strokeBlue: 0, strokeWidth: 0d);
        }
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
