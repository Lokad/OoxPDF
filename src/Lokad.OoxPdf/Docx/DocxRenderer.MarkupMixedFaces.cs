using System.Text;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    private static IReadOnlyList<DocxMarkupBalloonBodyPart>? ResolveMixedCommentBalloonBodyParts(
        DocxRelatedStoryLayout? storyLayout,
        string body,
        DocxFontResources? fontResources,
        CancellationToken cancellationToken)
    {
        if (fontResources is null || storyLayout is null ||
            storyLayout.Story.BodyElements.Count != 1 ||
            storyLayout.Story.BodyElements[0] is not DocxParagraphElement element ||
            storyLayout.InlineImages.Count != 0 || storyLayout.FloatingDrawings.Count != 0)
        {
            return null;
        }
        DocxParagraph paragraph = element.Paragraph;
        if (paragraph.Images.Count != 0 || paragraph.InlineTextBoxes.Count != 0 ||
            paragraph.FieldReferences.Count != 0 || paragraph.Revisions.Count != 0)
        {
            return null;
        }
        var parts = new List<DocxMarkupBalloonBodyPart>();
        foreach (DocxTextRun run in paragraph.Runs)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (run.Text.Length == 0) { continue; }
            if (run.Bold || run.Italic || run.Underline || run.Strike || run.DoubleStrike ||
                run.AllCaps || run.SmallCaps || run.Hidden || run.VerticalAlignmentValue is not null ||
                run.HighlightValue is not null || run.ShadingFillHex is not null ||
                run.CharacterSpacingPoints != 0d || run.FieldKind is not null ||
                (run.ColorHex is not null && run.ColorHex != "000000") ||
                !fontResources.RunResources.TryGetValue(run, out DocxRunFontResource? resource) ||
                resource.Resolution.Bold || resource.Resolution.Italic || resource.Resolution.IsFallback)
            {
                return null;
            }
            int index = 0;
            foreach (Rune rune in run.Text.EnumerateRunes())
            {
                if ((index++ & 255) == 0) { cancellationToken.ThrowIfCancellationRequested(); }
                ushort glyph = resource.Embedded.Font.MapCodePoint(rune.Value);
                if ((Rune.IsWhiteSpace(rune) && rune.Value != ' ') || glyph == 0 ||
                    !resource.Embedded.TryGetEncodedCid(glyph, out _)) { return null; }
            }
            parts.Add(new(run.Text, resource));
        }
        if (parts.Count < 2 || parts.All(part => ReferenceEquals(part.Resource.Embedded.Font, parts[0].Resource.Embedded.Font)))
        {
            return null;
        }
        parts[0] = parts[0] with { Text = parts[0].Text.TrimStart(' ') };
        parts[^1] = parts[^1] with { Text = parts[^1].Text.TrimEnd(' ') };
        int offset = 0;
        foreach (DocxMarkupBalloonBodyPart part in parts)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (part.Text.Length > body.Length - offset || !body.AsSpan(offset, part.Text.Length).SequenceEqual(part.Text)) { return null; }
            offset += part.Text.Length;
        }
        if (offset != body.Length || body.Contains("  ", StringComparison.Ordinal)) { return null; }
        return parts.ToArray();
    }

    private static bool TryMeasureWordCompatibleSingleRowParts(
        IReadOnlyList<DocxMarkupBalloonBodyPart>? parts,
        double fontSize,
        double availableWidth,
        out double width)
    {
        width = 0d;
        if (parts is null || parts.Count == 0) { return false; }
        foreach (DocxMarkupBalloonBodyPart part in parts)
        {
            width += part.Resource.Embedded.MeasureTextPoints(part.Text, fontSize);
        }
        return double.IsFinite(width) && width > 0d && width <= availableWidth;
    }

    private sealed record DocxMarkupTwoFaceRows(
        DocxMarkupBalloonBodyPart Prefix,
        DocxRunFontResource TailResource,
        DocxUniformBalloonRow[] TailLines,
        double? TailLineHeightEm);

    private static DocxMarkupTwoFaceRows? ResolveWordCompatibleTwoFaceRows(
        IReadOnlyList<DocxMarkupBalloonBodyPart>? parts,
        double fontSize,
        double firstLineWidth,
        double continuationWidth,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (parts is not { Count: 2 } || parts[0].Text.Length == 0 || parts[1].Text.Length == 0 ||
            parts[1].Text[0] == ' ') { return null; }
        double tailFirstWidth = firstLineWidth - parts[0].Resource.Embedded.MeasureTextPoints(parts[0].Text, fontSize);
        if (!double.IsFinite(tailFirstWidth) || tailFirstWidth <= 0d) { return null; }
        DocxRunFontResource tail = parts[1].Resource;
        ushort space = tail.Embedded.Font.MapCodePoint(' ');
        if (space == 0 || !tail.Embedded.TryGetEncodedCid(space, out _)) { return null; }
        string[] words = parts[1].Text.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (words.Length < 2) { return null; }
        // A fitting ASCII word prefix can share its first word with the tail.
        // Prefixes with spaces or other break opportunities still require the
        // first tail word to fit; their row composition needs separate evidence.
        double firstWordWidth = tail.Embedded.MeasureTextPoints(words[0] + " ", fontSize);
        if (!double.IsFinite(firstWordWidth) ||
            firstWordWidth > tailFirstWidth && !parts[0].Text.All(char.IsAsciiLetterOrDigit)) { return null; }
        var lines = WrapUniformBalloonWords(parts[1].Text, tail.Embedded, fontSize, tailFirstWidth, continuationWidth, cancellationToken);
        if (lines.Length < 2) { return null; }
        double? lineHeightEm = tail.Embedded.Font.UnitsPerEm > 0
            ? DocxLineMetrics.MeasureHheaLineHeight(tail.Embedded.Font, 1d) : null;
        return new(parts[0], tail, lines, lineHeightEm);
    }

    private static void RenderWordCompatibleTwoFaceRows(
        DocxMarkupTwoFaceRows rows,
        DocxMarkupBalloonPlacement placement,
        PdfGraphicsBuilder graphics,
        DocxRunFontResource terminalResource,
        double textX,
        double bodyFirstLineX,
        double firstBaselineY,
        double fontSize,
        CancellationToken cancellationToken)
    {
        DrawBalloonText(graphics, rows.Prefix.Resource, rows.Prefix.Text, bodyFirstLineX, firstBaselineY, fontSize,
            placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
        double tailX = bodyFirstLineX + rows.Prefix.Resource.Embedded.MeasureTextPoints(rows.Prefix.Text, fontSize);
        double gap = ResolveWordCompatibleBalloonLineGap(rows.TailLineHeightEm, fontSize);
        RenderUniformBalloonRows(rows.TailLines, placement, graphics, rows.TailResource, terminalResource,
            tailX, textX, firstBaselineY, gap, fontSize, cancellationToken);
    }
}
