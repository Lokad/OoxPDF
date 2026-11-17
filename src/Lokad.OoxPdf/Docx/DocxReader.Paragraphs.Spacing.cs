using System.Globalization;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxReader
{
    private static double ResolveParagraphSpacingPoints(DocxResolvedParagraphProperties paragraph,
        double lineHeight, bool before)
    {
        var spacing = paragraph.Spacing;
        // Select one effective unit after inheritance. Enabled automatic spacing
        // supersedes line and absolute values; line units supersede absolute ones.
        if (OoxBoolean.IsTrue(before ? spacing.BeforeAutoSpacingValue : spacing.AfterAutoSpacingValue))
        {
            return WordAutomaticParagraphSpacingPoints;
        }

        if (int.TryParse(before ? spacing.BeforeLinesValue : spacing.AfterLinesValue,
            NumberStyles.Integer, CultureInfo.InvariantCulture, out int hundredthsOfLine))
        {
            return lineHeight * hundredthsOfLine / 100d;
        }

        return (before ? paragraph.SpacingBeforePoints : paragraph.SpacingAfterPoints) ??
            (before ? 0d : DocxDefaults.DefaultParagraphAfterSpacingPoints);
    }
}
