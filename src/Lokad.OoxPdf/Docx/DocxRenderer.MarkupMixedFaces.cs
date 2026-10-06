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
        if (offset != body.Length) { return null; }
        if (body.Contains("  ", StringComparison.Ordinal))
        {
            if (parts.Count != 2 || !parts[1].Text.StartsWith("  ", StringComparison.Ordinal) ||
                parts[1].Text.Length <= 2 || parts[1].Text[2] == ' ' ||
                parts[1].Text[2..].Contains("  ", StringComparison.Ordinal)) { return null; }
            string prefix = parts[0].Text;
            int split = prefix.IndexOf(' ');
            if (split <= 0 || split != prefix.LastIndexOf(' ') || split == prefix.Length - 1) { return null; }
            for (int index = 0; index < prefix.Length; index++)
            {
                if ((index & 255) == 0) { cancellationToken.ThrowIfCancellationRequested(); }
                if (prefix[index] != ' ' && !char.IsAsciiLetterOrDigit(prefix[index])) { return null; }
            }
        }
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
        double? SecondContinuationGapEm = null,
        DocxUniformBalloonRow[]? PrefixLines = null,
        double? PrefixLineHeightEm = null)
    {
        public int PrintedRows => TailLines.Length + (PrefixLines is null
            ? ContinuationPrefix is null ? 0 : 1 : PrefixLines.Length - 1);
    }

    private static DocxMarkupTwoFaceRows? ResolveWordCompatibleTwoFaceRows(
        IReadOnlyList<DocxMarkupBalloonBodyPart>? parts,
        double fontSize,
        double firstLineWidth,
        double continuationWidth,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (parts is not { Count: 2 } || parts[0].Text.Length == 0 || parts[1].Text.Length == 0) { return null; }
        double tailFirstWidth = firstLineWidth - parts[0].Resource.Embedded.MeasureTextPoints(parts[0].Text, fontSize);
        if (!double.IsFinite(tailFirstWidth)) { return null; }
        if (tailFirstWidth <= 0d)
        {
            return ResolveWordCompatibleWidePrefixRows(parts[0], parts[1], fontSize,
                firstLineWidth, continuationWidth, cancellationToken);
        }
        if (parts[1].Text[0] == ' ')
        {
            return ResolveWordCompatibleLeadingTailRows(parts[0], parts[1], fontSize,
                tailFirstWidth, continuationWidth, cancellationToken);
        }
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

    private static DocxMarkupTwoFaceRows? ResolveWordCompatibleWidePrefixRows(
        DocxMarkupBalloonBodyPart prefix, DocxMarkupBalloonBodyPart tail,
        double fontSize, double firstLineWidth, double continuationWidth, CancellationToken cancellationToken)
    {
        bool singlePrefixWord = prefix.Text.IndexOf(' ') < 0;
        if ((!singlePrefixWord && prefix.Text.IndexOf(' ') <= 0) || prefix.Text.EndsWith(' ') || tail.Text[0] == ' ' ||
            !double.IsFinite(firstLineWidth) || firstLineWidth <= 0d) { return null; }
        for (int i = 0; i < prefix.Text.Length; i++)
        {
            if ((i & 255) == 0) { cancellationToken.ThrowIfCancellationRequested(); }
            if (prefix.Text[i] != ' ' && !char.IsAsciiLetterOrDigit(prefix.Text[i])) { return null; }
        }
        OpenTypeFont prefixFont = prefix.Resource.Embedded.Font;
        OpenTypeFont tailFont = tail.Resource.Embedded.Font;
        if (prefixFont.UnitsPerEm <= 0 || tailFont.UnitsPerEm <= 0) { return null; }
        if (singlePrefixWord)
        {
            for (int index = 0; index < prefix.Text.Length; index++)
            {
                if ((index & 255) == 0) { cancellationToken.ThrowIfCancellationRequested(); }
                double advance = prefixFont.GetAdvanceWidth(prefixFont.MapCodePoint(prefix.Text[index])) /
                    (double)prefixFont.UnitsPerEm * fontSize;
                if (!double.IsFinite(advance) || advance > continuationWidth ||
                    (index == 0 && (advance <= 0d || advance > firstLineWidth))) { return null; }
            }
        }
        else
        {
            foreach (string word in prefix.Text.Split(' ', StringSplitOptions.RemoveEmptyEntries))
            {
                cancellationToken.ThrowIfCancellationRequested();
                double width = prefix.Resource.Embedded.MeasureTextPoints(word, fontSize);
                if (!double.IsFinite(width) || width <= 0d || width > continuationWidth) { return null; }
            }
            string firstPrefixWord = prefix.Text[..prefix.Text.IndexOf(' ')];
            if (prefix.Resource.Embedded.MeasureTextPoints(firstPrefixWord + " ", fontSize) > firstLineWidth) { return null; }
        }
        ushort space = tailFont.MapCodePoint(' ');
        if (space == 0 || !tail.Resource.Embedded.TryGetEncodedCid(space, out _)) { return null; }
        string[] tailWords = tail.Text.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (tailWords.Length < 2) { return null; }
        var prefixLines = WrapUniformBalloonWords(prefix.Text, prefix.Resource.Embedded, fontSize,
            firstLineWidth, continuationWidth, cancellationToken, reserveFirstRowBreakSpace: false);
        if (prefixLines.Length < 2 || prefixLines[0].Text.Length == 0) { return null; }
        double available = continuationWidth - prefix.Resource.Embedded.MeasureTextPoints(prefixLines[^1].Text, fontSize);
        double firstTailWordWidth = tail.Resource.Embedded.MeasureTextPoints(tailWords[0] + " ", fontSize);
        if (!double.IsFinite(available) || available <= 0d || !double.IsFinite(firstTailWordWidth)) { return null; }
        if (firstTailWordWidth > available)
        {
            DocxUniformBalloonRow lastPrefixRow = prefixLines[^1];
            int split = lastPrefixRow.Text.LastIndexOf(' ');
            if (split > 0)
            {
                // The last prefix word joins the first tail word. Move that
                // whole joined word to a fresh row before splitting its tail.
                prefixLines[^1] = lastPrefixRow with { Text = lastPrefixRow.Text[..split], SpaceAfter = true };
                prefixLines = [.. prefixLines, new(lastPrefixRow.Text[(split + 1)..], SpaceAfter: false)];
                available = continuationWidth - prefix.Resource.Embedded.MeasureTextPoints(prefixLines[^1].Text, fontSize);
                if (!double.IsFinite(available) || available <= 0d) { return null; }
            }
            int scalarIndex = 0;
            foreach (Rune rune in tailWords[0].EnumerateRunes())
            {
                if ((scalarIndex & 255) == 0) { cancellationToken.ThrowIfCancellationRequested(); }
                double advance = tailFont.GetAdvanceWidth(tailFont.MapCodePoint(rune.Value)) /
                    (double)tailFont.UnitsPerEm * fontSize;
                if (!double.IsFinite(advance) || advance > continuationWidth ||
                    (scalarIndex == 0 && (advance <= 0d || advance > available))) { return null; }
                scalarIndex++;
            }
        }
        double prefixHeight = DocxLineMetrics.MeasureHheaLineHeight(prefixFont, 1d);
        double prefixDescent = -prefixFont.Hhea.HorizontalDescender / (double)prefixFont.UnitsPerEm;
        double tailDescent = -tailFont.Hhea.HorizontalDescender / (double)tailFont.UnitsPerEm;
        double prefixAscentGap = (prefixFont.Hhea.HorizontalAscender + prefixFont.Hhea.HorizontalLineGap) / (double)prefixFont.UnitsPerEm;
        double tailAscentGap = (tailFont.Hhea.HorizontalAscender + tailFont.Hhea.HorizontalLineGap) / (double)tailFont.UnitsPerEm;
        double incomingGap = prefixDescent + Math.Max(prefixAscentGap, tailAscentGap);
        double outgoingGap = Math.Max(prefixDescent, tailDescent) + tailAscentGap;
        if (!double.IsFinite(prefixHeight) || prefixHeight <= 0d || !double.IsFinite(incomingGap) || incomingGap <= 0d ||
            !double.IsFinite(outgoingGap) || outgoingGap <= 0d) { return null; }
        var tailLines = WrapUniformBalloonWords(tail.Text, tail.Resource.Embedded, fontSize,
            available, continuationWidth, cancellationToken, reserveFirstRowBreakSpace: false);
        if (tailLines.Length == 0 || tailLines[0].Text.Length == 0) { return null; }
        prefixLines[^1] = prefixLines[^1] with { SpaceAfter = false };
        return new(prefix, tail.Resource, tailLines, DocxLineMetrics.MeasureHheaLineHeight(tailFont, 1d), incomingGap,
            SecondContinuationGapEm: outgoingGap, PrefixLines: prefixLines, PrefixLineHeightEm: prefixHeight);
    }

    private static DocxMarkupTwoFaceRows? ResolveWordCompatibleLeadingTailRows(
        DocxMarkupBalloonBodyPart prefix, DocxMarkupBalloonBodyPart tail,
        double fontSize, double firstTailWidth, double continuationWidth, CancellationToken cancellationToken)
    {
        int split = prefix.Text.IndexOf(' ');
        int separatorLength = tail.Text.Length > 1 && tail.Text[1] == ' ' ? 2 : 1;
        // A terminal comma stays in the prefix face; other punctuation keeps fallback.
        int plainPrefixLength = prefix.Text.EndsWith(',') ? prefix.Text.Length - 1 : prefix.Text.Length;
        if (split <= 0 || split != prefix.Text.LastIndexOf(' ') || split >= plainPrefixLength - 1 ||
            tail.Text.Length <= separatorLength || tail.Text[separatorLength] == ' ') { return null; }
        for (int i = 0; i < plainPrefixLength; i++)
        {
            if ((i & 255) == 0) { cancellationToken.ThrowIfCancellationRequested(); }
            if (prefix.Text[i] != ' ' && !char.IsAsciiLetterOrDigit(prefix.Text[i])) { return null; }
        }
        OpenTypeFont prefixFont = prefix.Resource.Embedded.Font;
        OpenTypeFont tailFont = tail.Resource.Embedded.Font;
        ushort space = tailFont.MapCodePoint(' ');
        if (space == 0 || !tail.Resource.Embedded.TryGetEncodedCid(space, out _) ||
            prefixFont.UnitsPerEm <= 0 || tailFont.UnitsPerEm <= 0) { return null; }
        string separator = tail.Text[..separatorLength];
        double available = firstTailWidth - tail.Resource.Embedded.MeasureTextPoints(separator, fontSize);
        string text = tail.Text[separatorLength..];
        string[] words = text.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (!double.IsFinite(available) || available <= 0d || words.Length < 2) { return null; }
        double firstWordWidth = tail.Resource.Embedded.MeasureTextPoints(words[0] + " ", fontSize);
        if (!double.IsFinite(firstWordWidth)) { return null; }
        bool startsBelowPrefix = firstWordWidth > available;
        double prefixDescent = -prefixFont.Hhea.HorizontalDescender / (double)prefixFont.UnitsPerEm;
        double descent = startsBelowPrefix ? prefixDescent : Math.Max(prefixDescent,
            -tailFont.Hhea.HorizontalDescender / (double)tailFont.UnitsPerEm);
        double transition = descent + (tailFont.Hhea.HorizontalAscender + tailFont.Hhea.HorizontalLineGap) / (double)tailFont.UnitsPerEm;
        if (!double.IsFinite(transition) || transition <= 0d) { return null; }
        var lines = WrapUniformBalloonWords(text, tail.Resource.Embedded, fontSize,
            startsBelowPrefix ? -1d : available, continuationWidth, cancellationToken);
        if (lines.Length < 2 || (!startsBelowPrefix && lines[0].Text.Length == 0)) { return null; }
        // Keep the authored separators in their tail resource and reserve their
        // advance beside the complete prefix. A separated first word that
        // does not fit starts below it, including zero-advance first glyphs.
        lines[0] = lines[0] with { Text = separator + lines[0].Text };
        return new(prefix, tail.Resource, lines, DocxLineMetrics.MeasureHheaLineHeight(tailFont, 1d), transition);
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
        if (rows.PrefixLines is { } prefixLines)
        {
            double prefixGap = ResolveWordCompatibleBalloonLineGap(rows.PrefixLineHeightEm, fontSize);
            double mixedBaselineY = firstBaselineY - (prefixLines.Length - 2) * prefixGap -
                rows.FirstContinuationGapEm!.Value * fontSize;
            for (int i = 0; i < prefixLines.Length; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                string text = prefixLines[i].Text + (prefixLines[i].SpaceAfter ? " " : string.Empty);
                double x = i == 0 ? bodyFirstLineX : textX;
                double y = i == prefixLines.Length - 1 ? mixedBaselineY : firstBaselineY - i * prefixGap;
                DrawBalloonText(graphics, rows.Prefix.Resource, text, x, y, fontSize,
                    placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
            }
            double mixedTailX = textX + rows.Prefix.Resource.Embedded.MeasureTextPoints(prefixLines[^1].Text, fontSize);
            double tailGap = ResolveWordCompatibleBalloonLineGap(rows.TailLineHeightEm, fontSize);
            RenderUniformBalloonRows(rows.TailLines, placement, graphics, rows.TailResource, terminalResource,
                mixedTailX, textX, mixedBaselineY, tailGap, fontSize, cancellationToken,
                rows.SecondContinuationGapEm * fontSize);
            return;
        }
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
