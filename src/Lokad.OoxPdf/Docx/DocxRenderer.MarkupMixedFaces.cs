using System.Text;
using Lokad.OoxPdf.Fonts;
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
        double? TailLineHeightEm,
        double? FirstContinuationGapEm,
        DocxMarkupBalloonBodyPart? ContinuationPrefix = null,
        double? SecondContinuationGapEm = null)
    {
        public int PrintedRows => TailLines.Length + (ContinuationPrefix is null ? 0 : 1);
    }

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
        double firstWordWidth = tail.Embedded.MeasureTextPoints(words[0] + " ", fontSize);
        if (!double.IsFinite(firstWordWidth)) { return null; }
        double? firstGapEm = null;
        if (firstWordWidth > tailFirstWidth && !parts[0].Text.All(char.IsAsciiLetterOrDigit))
        {
            string prefix = parts[0].Text;
            if (!prefix.EndsWith(' ') || !prefix[..^1].All(char.IsAsciiLetterOrDigit))
            {
                return ResolveWordCompatibleComposedPrefixRows(parts[0], parts[1], fontSize,
                    continuationWidth, cancellationToken);
            }
            OpenTypeFont prefixFont = parts[0].Resource.Embedded.Font;
            OpenTypeFont tailFont = tail.Embedded.Font;
            if (prefixFont.UnitsPerEm <= 0 || tailFont.UnitsPerEm <= 0) { return null; }
            double transition = -prefixFont.Hhea.HorizontalDescender / (double)prefixFont.UnitsPerEm +
                (tailFont.Hhea.HorizontalAscender + tailFont.Hhea.HorizontalLineGap) / (double)tailFont.UnitsPerEm;
            if (!double.IsFinite(transition) || transition <= 0d) { return null; }
            firstGapEm = transition;
        }
        // A joined ASCII word can split beside its prefix. A separated word
        // starts below the prefix, even if its first glyph has zero advance.
        var lines = WrapUniformBalloonWords(parts[1].Text, tail.Embedded, fontSize,
            firstGapEm is null ? tailFirstWidth : -1d, continuationWidth, cancellationToken);
        if (lines.Length < 2) { return null; }
        double? lineHeightEm = tail.Embedded.Font.UnitsPerEm > 0
            ? DocxLineMetrics.MeasureHheaLineHeight(tail.Embedded.Font, 1d) : null;
        return new(parts[0], tail, lines, lineHeightEm, firstGapEm);
    }

    private static DocxMarkupTwoFaceRows? ResolveWordCompatibleComposedPrefixRows(
        DocxMarkupBalloonBodyPart prefix, DocxMarkupBalloonBodyPart tail,
        double fontSize, double continuationWidth, CancellationToken cancellationToken)
    {
        int split = prefix.Text.LastIndexOf(' ');
        if (split <= 0 || split == prefix.Text.Length - 1 || split != prefix.Text.IndexOf(' ')) { return null; }
        for (int i = 0; i < prefix.Text.Length; i++)
        {
            if ((i & 255) == 0) { cancellationToken.ThrowIfCancellationRequested(); }
            if (prefix.Text[i] != ' ' && !char.IsAsciiLetterOrDigit(prefix.Text[i])) { return null; }
        }
        DocxMarkupBalloonBodyPart residual = prefix with { Text = prefix.Text[(split + 1)..] };
        double available = continuationWidth - residual.Resource.Embedded.MeasureTextPoints(residual.Text, fontSize);
        if (!double.IsFinite(available) || available <= 0d) { return null; }
        OpenTypeFont prefixFont = prefix.Resource.Embedded.Font;
        OpenTypeFont tailFont = tail.Resource.Embedded.Font;
        if (prefixFont.UnitsPerEm <= 0 || tailFont.UnitsPerEm <= 0) { return null; }
        double prefixDescent = -prefixFont.Hhea.HorizontalDescender / (double)prefixFont.UnitsPerEm;
        double tailDescent = -tailFont.Hhea.HorizontalDescender / (double)tailFont.UnitsPerEm;
        double prefixAscentGap = (prefixFont.Hhea.HorizontalAscender + prefixFont.Hhea.HorizontalLineGap) / (double)prefixFont.UnitsPerEm;
        double tailAscentGap = (tailFont.Hhea.HorizontalAscender + tailFont.Hhea.HorizontalLineGap) / (double)tailFont.UnitsPerEm;
        double firstGap = prefixDescent + Math.Max(prefixAscentGap, tailAscentGap);
        double secondGap = Math.Max(prefixDescent, tailDescent) + tailAscentGap;
        if (!double.IsFinite(firstGap) || firstGap <= 0d || !double.IsFinite(secondGap) || secondGap <= 0d) { return null; }
        // The first tail row is already a continuation of the prefix row:
        // its visible words fit without reserving the emitted break separator.
        var lines = WrapUniformBalloonWords(tail.Text, tail.Resource.Embedded, fontSize,
            available, continuationWidth, cancellationToken, reserveFirstRowBreakSpace: false);
        if (lines.Length == 0 || lines[0].Text.Length == 0) { return null; }
        return new(prefix with { Text = prefix.Text[..(split + 1)] }, tail.Resource, lines,
            DocxLineMetrics.MeasureHheaLineHeight(tailFont, 1d), firstGap, residual, secondGap);
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
        if (rows.ContinuationPrefix is DocxMarkupBalloonBodyPart residual)
        {
            double mixedY = firstBaselineY - rows.FirstContinuationGapEm!.Value * fontSize;
            DrawBalloonText(graphics, residual.Resource, residual.Text, textX, mixedY, fontSize,
                placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
            tailX = textX + residual.Resource.Embedded.MeasureTextPoints(residual.Text, fontSize);
            RenderUniformBalloonRows(rows.TailLines, placement, graphics, rows.TailResource, terminalResource,
                tailX, textX, mixedY, gap, fontSize, cancellationToken,
                rows.SecondContinuationGapEm * fontSize);
            return;
        }
        RenderUniformBalloonRows(rows.TailLines, placement, graphics, rows.TailResource, terminalResource,
            tailX, textX, firstBaselineY, gap, fontSize, cancellationToken,
            rows.FirstContinuationGapEm * fontSize);
    }
}
