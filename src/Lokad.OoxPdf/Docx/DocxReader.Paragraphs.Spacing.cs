using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxReader
{
    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static double ResolveSpacingBeforePoints(DocxResolvedParagraphProperties paragraph, double lineHeight)
    {
        if (paragraph.SpacingBeforePoints is { } points)
        {
            return points;
        }

        if (OoxBoolean.IsTrue(paragraph.Spacing.BeforeAutoSpacingValue))
        {
            return WordAutomaticParagraphSpacingPoints;
        }

        return TryReadLineBasedSpacing(paragraph.Spacing.BeforeLinesValue, lineHeight, out double linePoints)
            ? linePoints
            : 0d;
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static double ResolveDefaultAutoLineSpacingFactor(DocxResolvedParagraphProperties paragraph)
    {
        return DocxParagraphSpacing.HasBeforeSpacingSide(paragraph.Spacing) || DocxParagraphSpacing.HasAfterSpacingSide(paragraph.Spacing)
            ? WordSpacingTokenAutoLineSpacingFactor
            : WordUntokenedAutoLineSpacingFactor;
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static double ResolveSpacingAfterPoints(DocxResolvedParagraphProperties paragraph, double lineHeight)
    {
        if (paragraph.SpacingAfterPoints is { } points)
        {
            return points;
        }

        if (OoxBoolean.IsTrue(paragraph.Spacing.AfterAutoSpacingValue))
        {
            return WordAutomaticParagraphSpacingPoints;
        }

        return TryReadLineBasedSpacing(paragraph.Spacing.AfterLinesValue, lineHeight, out double linePoints)
            ? linePoints
            : WordDefaultSpacingAfterPoints;
    }

    private static bool TryReadLineBasedSpacing(string? value, double lineHeight, out double points)
    {
        if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int hundredthsOfLine))
        {
            points = lineHeight * hundredthsOfLine / 100d;
            return true;
        }

        points = 0d;
        return false;
    }
}
