using System.Globalization;
using System.Text;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    private static IReadOnlyList<DocxMarkupBalloonParagraph>? ResolveCommentBalloonParagraphs(
        DocxRelatedStoryLayout? storyLayout,
        string preview,
        DocxFontResources? fonts,
        CancellationToken cancellationToken)
    {
        // Preserve each plain paragraph's prepared body and mark faces. The
        // measured rendering paths separately admit fitting and wrapped bodies.
        if (fonts is null || storyLayout is null || storyLayout.Story.BodyElements.Count is not (2 or 3 or 4) ||
            storyLayout.InlineImages.Count != 0 || storyLayout.FloatingDrawings.Count != 0) { return null; }
        var paragraphs = new List<DocxMarkupBalloonParagraph>(storyLayout.Story.BodyElements.Count);
        foreach (DocxBodyElement item in storyLayout.Story.BodyElements)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (item is not DocxParagraphElement element) { return null; }
            DocxParagraph paragraph = element.Paragraph;
            if (paragraph.Images.Count != 0 || paragraph.InlineTextBoxes.Count != 0 ||
                paragraph.FieldReferences.Count != 0 || paragraph.Revisions.Count != 0 ||
                paragraph.ListLabel is not null || paragraph.Alignment != DocxTextAlignment.Left ||
                paragraph.Indent != DocxParagraphIndent.Empty || paragraph.TabStops.Count != 0 ||
                !IsWordCompatibleBalloonParagraphSpacing(paragraph.Spacing) || paragraph.LineSpacingPoints is not null ||
                paragraph.ParagraphMarkRun is not { } mark || !IsPlainBalloonParagraphRun(mark) ||
                !fonts.RunResources.TryGetValue(mark, out DocxRunFontResource? markResource) ||
                markResource.Resolution.Bold || markResource.Resolution.Italic || markResource.Resolution.IsFallback ||
                !HasEncodedBalloonParagraphText(markResource, " ", cancellationToken)) { return null; }
            DocxRunFontResource? bodyResource = null;
            var text = new StringBuilder();
            foreach (DocxTextRun run in paragraph.Runs)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (run.Text.Length == 0) { continue; }
                if (!IsPlainBalloonParagraphRun(run) ||
                    !fonts.RunResources.TryGetValue(run, out DocxRunFontResource? resource) ||
                    resource.Resolution.Bold || resource.Resolution.Italic || resource.Resolution.IsFallback ||
                    (bodyResource is not null && !ReferenceEquals(bodyResource.Embedded.Font, resource.Embedded.Font)) ||
                    !HasEncodedBalloonParagraphText(resource, run.Text, cancellationToken)) { return null; }
                bodyResource = resource;
                text.Append(run.Text);
            }
            string body = text.ToString().Trim(' ');
            if (bodyResource is null || body.Length == 0 || body.Contains("  ", StringComparison.Ordinal)) { return null; }
            paragraphs.Add(new(new(body, bodyResource), markResource));
        }
        return string.Join(" ", paragraphs.Select(paragraph => paragraph.Body.Text)) == preview ? paragraphs : null;
    }

    private static bool IsPlainBalloonParagraphRun(DocxTextRun run) =>
        !run.Bold && !run.Italic && !run.Underline && !run.Strike && !run.DoubleStrike &&
        !run.AllCaps && !run.SmallCaps && !run.Hidden && run.VerticalAlignmentValue is null &&
        run.HighlightValue is null && run.ShadingFillHex is null && run.CharacterSpacingPoints == 0d &&
        run.FieldKind is null && (run.ColorHex is null || run.ColorHex == "000000");

    private static bool IsWordCompatibleBalloonParagraphSpacing(DocxParagraphSpacing spacing) =>
        // Office normalizes explicit twips, canonical before/afterLines=100 and afterAutospacing=1 in balloons.
        // Keep unqualified line/automatic tokens and contextual spacing on the existing path.
        (spacing.BeforeLinesValue is null or "100") && (spacing.AfterLinesValue is null or "100") &&
        spacing.BeforeAutoSpacingValue is null && (spacing.AfterAutoSpacingValue is null or "1") &&
        spacing.LineValue is null && spacing.LineRuleValue is null && spacing.ContextualSpacing is null &&
        IsBalloonParagraphTwipsToken(spacing.BeforeValue) && IsBalloonParagraphTwipsToken(spacing.AfterValue);

    private static bool IsBalloonParagraphTwipsToken(string? value) =>
        value is null || uint.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out _);

    private static bool HasEncodedBalloonParagraphText(DocxRunFontResource resource, string text, CancellationToken cancellationToken)
    {
        int index = 0;
        foreach (Rune rune in text.EnumerateRunes())
        {
            if ((index++ & 255) == 0) { cancellationToken.ThrowIfCancellationRequested(); }
            ushort glyph = resource.Embedded.Font.MapCodePoint(rune.Value);
            if ((Rune.IsWhiteSpace(rune) && rune.Value != ' ') || glyph == 0 ||
                !resource.Embedded.TryGetEncodedCid(glyph, out _)) { return false; }
        }
        return true;
    }

    private static double? ResolveWordCompatibleParagraphGap(
        IReadOnlyList<DocxMarkupBalloonParagraph>? paragraphs, double fontSize,
        double firstWidth, double continuationWidth)
    {
        if (paragraphs is not { Count: 2 }) { return null; }
        for (int index = 0; index < 2; index++)
        {
            DocxMarkupBalloonBodyPart body = paragraphs[index].Body;
            double width = body.Resource.Embedded.MeasureTextPoints(body.Text, fontSize);
            if (!double.IsFinite(width) || width <= 0d || width > (index == 0 ? firstWidth : continuationWidth)) { return null; }
        }
        return MeasureWordCompatibleParagraphTransition(paragraphs, fontSize);
    }

    private static double? MeasureWordCompatibleParagraphTransition(IReadOnlyList<DocxMarkupBalloonParagraph> paragraphs, double fontSize)
        => MeasureWordCompatibleParagraphTransition(paragraphs[0], paragraphs[1], fontSize);

    private static double? MeasureWordCompatibleParagraphTransition(
        DocxMarkupBalloonParagraph previous, DocxMarkupBalloonParagraph next, double fontSize)
    {
        var first = previous.Body.Resource.Embedded.Font;
        var second = next.Body.Resource.Embedded.Font;
        if (first.UnitsPerEm <= 0 || second.UnitsPerEm <= 0) { return null; }
        double gap = (-first.Hhea.HorizontalDescender / (double)first.UnitsPerEm +
            (second.Hhea.HorizontalAscender + second.Hhea.HorizontalLineGap) / (double)second.UnitsPerEm) * fontSize;
        return double.IsFinite(gap) && gap > 0d ? gap : null;
    }

    private static double[]? ResolveWordCompatibleFittingParagraphGaps(
        IReadOnlyList<DocxMarkupBalloonParagraph>? paragraphs, double fontSize,
        double firstWidth, double continuationWidth)
    {
        if (paragraphs is null || paragraphs.Count is not (3 or 4)) { return null; }
        for (int index = 0; index < paragraphs.Count; index++)
        {
            DocxMarkupBalloonBodyPart body = paragraphs[index].Body;
            double width = body.Resource.Embedded.MeasureTextPoints(body.Text, fontSize);
            if (!double.IsFinite(width) || width <= 0d || width > (index == 0 ? firstWidth : continuationWidth)) { return null; }
        }
        var gaps = new double[paragraphs.Count - 1];
        for (int index = 0; index < gaps.Length; index++)
        {
            if (MeasureWordCompatibleParagraphTransition(paragraphs[index], paragraphs[index + 1], fontSize) is not double gap) { return null; }
            gaps[index] = gap;
        }
        return gaps;
    }

    private static void RenderWordCompatibleFittingParagraphs(
        IReadOnlyList<DocxMarkupBalloonParagraph> paragraphs, double[] gaps,
        DocxMarkupBalloonPlacement placement, PdfGraphicsBuilder graphics, double firstX,
        double continuationX, double firstY, double fontSize, CancellationToken cancellationToken)
    {
        double y = firstY;
        for (int index = 0; index < paragraphs.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxMarkupBalloonParagraph paragraph = paragraphs[index];
            double x = index == 0 ? firstX : continuationX;
            DrawBalloonText(graphics, paragraph.Body.Resource, paragraph.Body.Text, x, y, fontSize,
                placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
            DrawBalloonText(graphics, paragraph.Mark, " ", x + paragraph.Body.Resource.Embedded.MeasureTextPoints(paragraph.Body.Text, fontSize),
                y, fontSize, placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
            if (index < gaps.Length) { y -= gaps[index]; }
        }
    }

    private sealed record DocxMarkupLastWrappedParagraphRows(
        DocxMarkupBalloonParagraph[] Leading, double[] LeadingGaps,
        DocxMarkupBalloonParagraph Last, DocxUniformBalloonRow[] LastRows,
        double LastGap, double LastPitch)
    {
        public double LastStartHeight => LeadingGaps.Sum() + LastGap;
        public double ContinuationsHeight => LastStartHeight + (LastRows.Length - 1) * LastPitch;
    }

    private static DocxMarkupLastWrappedParagraphRows? ResolveWordCompatibleLastWrappedParagraphRows(
        IReadOnlyList<DocxMarkupBalloonParagraph>? paragraphs, double fontSize,
        double firstWidth, double continuationWidth, CancellationToken cancellationToken)
    {
        if (paragraphs is null || paragraphs.Count != 4 || !double.IsFinite(firstWidth) || firstWidth <= 0d ||
            !double.IsFinite(continuationWidth) || continuationWidth <= 0d) { return null; }
        cancellationToken.ThrowIfCancellationRequested();
        var leading = paragraphs.Take(3).ToArray();
        double[]? gaps = ResolveWordCompatibleFittingParagraphGaps(leading, fontSize, firstWidth, continuationWidth);
        if (gaps is null) { return null; }
        DocxMarkupBalloonParagraph last = paragraphs[3];
        double width = last.Body.Resource.Embedded.MeasureTextPoints(last.Body.Text, fontSize);
        if (!double.IsFinite(width) || width <= continuationWidth) { return null; }
        double? lastGap = MeasureWordCompatibleParagraphTransition(leading[^1], last, fontSize);
        var font = last.Body.Resource.Embedded.Font;
        double pitch = DocxLineMetrics.MeasureHheaLineHeight(font, fontSize);
        if (lastGap is null || !double.IsFinite(pitch) || pitch <= 0d) { return null; }
        int index = 0;
        foreach (Rune rune in last.Body.Text.EnumerateRunes())
        {
            if ((index++ & 255) == 0) { cancellationToken.ThrowIfCancellationRequested(); }
            double scalarWidth = font.GetAdvanceWidth(font.MapCodePoint(rune.Value)) * fontSize / font.UnitsPerEm;
            if (!double.IsFinite(scalarWidth) || scalarWidth > continuationWidth) { return null; }
        }
        var rows = WrapUniformBalloonWords(last.Body.Text, last.Body.Resource.Embedded, fontSize,
            continuationWidth, continuationWidth, cancellationToken, reserveFirstRowBreakSpace: false);
        return rows.Length >= 2 && rows[0].Text.Length != 0 ? new(leading, gaps, last, rows, lastGap.Value, pitch) : null;
    }

    private static void RenderWordCompatibleLastWrappedParagraphRows(
        DocxMarkupLastWrappedParagraphRows rows, DocxMarkupBalloonPlacement placement,
        PdfGraphicsBuilder graphics, double firstX, double continuationX, double firstY,
        double fontSize, CancellationToken cancellationToken)
    {
        RenderWordCompatibleFittingParagraphs(rows.Leading, rows.LeadingGaps, placement, graphics,
            firstX, continuationX, firstY, fontSize, cancellationToken);
        RenderUniformBalloonRows(rows.LastRows, placement, graphics, rows.Last.Body.Resource, rows.Last.Mark,
            continuationX, continuationX, firstY - rows.LastStartHeight, rows.LastPitch, fontSize, cancellationToken);
    }

    private sealed record DocxMarkupWrappedParagraphRows(
        DocxMarkupBalloonParagraph First, DocxMarkupBalloonParagraph Second,
        DocxUniformBalloonRow[] TailRows, double FirstGap, double TailGap,
        DocxMarkupBalloonParagraph? Third = null, double? ThirdGap = null,
        DocxUniformBalloonRow[]? ThirdRows = null, double ThirdTailGap = 0d,
        DocxUniformBalloonRow[]? FirstRows = null, double FirstTailGap = 0d,
        DocxMarkupBalloonParagraph? Fourth = null, double? FourthGap = null,
        DocxUniformBalloonRow[]? FourthRows = null, double FourthTailGap = 0d)
    {
        public double FirstContinuationsHeight => FirstRows is null ? 0d : (FirstRows.Length - 1) * FirstTailGap;
        public double FourthContinuationsHeight => FourthRows is null ? 0d : (FourthRows.Length - 1) * FourthTailGap;
        public double ContinuationsHeight => FirstContinuationsHeight + FirstGap + (TailRows.Length - 1) * TailGap + (ThirdGap ?? 0d) +
            (ThirdRows is null ? 0d : (ThirdRows.Length - 1) * ThirdTailGap) + (FourthGap ?? 0d) + FourthContinuationsHeight;
    }

    private static DocxMarkupWrappedParagraphRows? ResolveWordCompatibleWrappedParagraphRows(
        IReadOnlyList<DocxMarkupBalloonParagraph>? paragraphs, double fontSize,
        double firstWidth, double continuationWidth, CancellationToken cancellationToken)
    {
        if (paragraphs is null || paragraphs.Count is not (2 or 3 or 4) || !double.IsFinite(continuationWidth) || continuationWidth <= 0d) { return null; }
        DocxMarkupBalloonParagraph first = paragraphs[0], second = paragraphs[1];
        double prefixWidth = first.Body.Resource.Embedded.MeasureTextPoints(first.Body.Text, fontSize);
        double tailWidth = second.Body.Resource.Embedded.MeasureTextPoints(second.Body.Text, fontSize);
        if (!double.IsFinite(prefixWidth) || prefixWidth <= 0d ||
            !double.IsFinite(tailWidth) || tailWidth <= 0d) { return null; }
        if (paragraphs.Count == 4)
        {
            double thirdWidth = paragraphs[2].Body.Resource.Embedded.MeasureTextPoints(paragraphs[2].Body.Text, fontSize);
            double fourthWidth = paragraphs[3].Body.Resource.Embedded.MeasureTextPoints(paragraphs[3].Body.Text, fontSize);
            bool firstWrapped = prefixWidth > firstWidth, secondWrapped = tailWidth > continuationWidth;
            bool thirdWrapped = thirdWidth > continuationWidth;
            if (!double.IsFinite(firstWidth) || firstWidth <= 0d || !double.IsFinite(thirdWidth) || thirdWidth <= 0d ||
                !double.IsFinite(fourthWidth) || fourthWidth <= 0d ||
                !(firstWrapped || secondWrapped || thirdWrapped)) { return null; }
        }
        DocxUniformBalloonRow[]? firstRows = null;
        double firstTailGap = 0d;
        if (prefixWidth > firstWidth)
        {
            if (!double.IsFinite(firstWidth) || firstWidth <= 0d) { return null; }
            var firstFont = first.Body.Resource.Embedded.Font;
            int firstIndex = 0;
            foreach (Rune rune in first.Body.Text.EnumerateRunes())
            {
                if ((firstIndex++ & 255) == 0) { cancellationToken.ThrowIfCancellationRequested(); }
                double width = firstFont.GetAdvanceWidth(firstFont.MapCodePoint(rune.Value)) * fontSize / firstFont.UnitsPerEm;
                if (!double.IsFinite(width) || width > continuationWidth) { return null; }
            }
            firstTailGap = DocxLineMetrics.MeasureHheaLineHeight(firstFont, fontSize);
            if (!double.IsFinite(firstTailGap) || firstTailGap <= 0d) { return null; }
            firstRows = WrapUniformBalloonWords(first.Body.Text, first.Body.Resource.Embedded, fontSize,
                firstWidth, continuationWidth, cancellationToken, reserveFirstRowBreakSpace: false);
            if (firstRows.Length < 2 || firstRows[0].Text.Length == 0) { return null; }
        }
        DocxMarkupBalloonParagraph? third = paragraphs.Count >= 3 ? paragraphs[2] : null;
        double? thirdGap = null;
        DocxUniformBalloonRow[]? thirdRows = null;
        double thirdTailGap = 0d;
        if (third is not null)
        {
            double width = third.Body.Resource.Embedded.MeasureTextPoints(third.Body.Text, fontSize);
            if (!double.IsFinite(width) || width <= 0d) { return null; }
            thirdGap = MeasureWordCompatibleParagraphTransition(second, third, fontSize);
            if (thirdGap is null) { return null; }
            if (width > continuationWidth)
            {
                var thirdFont = third.Body.Resource.Embedded.Font;
                int thirdIndex = 0;
                foreach (Rune rune in third.Body.Text.EnumerateRunes())
                {
                    if ((thirdIndex++ & 255) == 0) { cancellationToken.ThrowIfCancellationRequested(); }
                    double scalarWidth = thirdFont.GetAdvanceWidth(thirdFont.MapCodePoint(rune.Value)) * fontSize / thirdFont.UnitsPerEm;
                    if (!double.IsFinite(scalarWidth) || scalarWidth > continuationWidth) { return null; }
                }
                thirdTailGap = DocxLineMetrics.MeasureHheaLineHeight(thirdFont, fontSize);
                if (!double.IsFinite(thirdTailGap) || thirdTailGap <= 0d) { return null; }
                thirdRows = WrapUniformBalloonWords(third.Body.Text, third.Body.Resource.Embedded, fontSize,
                    continuationWidth, continuationWidth, cancellationToken, reserveFirstRowBreakSpace: false);
                if (thirdRows.Length < 2 || thirdRows[0].Text.Length == 0) { return null; }
            }
        }
        DocxMarkupBalloonParagraph? fourth = paragraphs.Count == 4 ? paragraphs[3] : null;
        double? fourthGap = fourth is null ? null : MeasureWordCompatibleParagraphTransition(third!, fourth, fontSize);
        if (fourth is not null && fourthGap is null) { return null; }
        DocxUniformBalloonRow[]? fourthRows = null;
        double fourthTailGap = 0d;
        if (fourth is not null && fourth.Body.Resource.Embedded.MeasureTextPoints(fourth.Body.Text, fontSize) > continuationWidth)
        {
            var fourthFont = fourth.Body.Resource.Embedded.Font;
            int fourthIndex = 0;
            foreach (Rune rune in fourth.Body.Text.EnumerateRunes())
            {
                if ((fourthIndex++ & 255) == 0) { cancellationToken.ThrowIfCancellationRequested(); }
                double scalarWidth = fourthFont.GetAdvanceWidth(fourthFont.MapCodePoint(rune.Value)) * fontSize / fourthFont.UnitsPerEm;
                if (!double.IsFinite(scalarWidth) || scalarWidth > continuationWidth) { return null; }
            }
            fourthTailGap = DocxLineMetrics.MeasureHheaLineHeight(fourthFont, fontSize);
            if (!double.IsFinite(fourthTailGap) || fourthTailGap <= 0d) { return null; }
            fourthRows = WrapUniformBalloonWords(fourth.Body.Text, fourth.Body.Resource.Embedded, fontSize,
                continuationWidth, continuationWidth, cancellationToken, reserveFirstRowBreakSpace: false);
            if (fourthRows.Length < 2 || fourthRows[0].Text.Length == 0) { return null; }
        }
        if (firstRows is null && thirdRows is null && tailWidth <= continuationWidth) { return null; }
        var tailFont = second.Body.Resource.Embedded.Font;
        int index = 0;
        foreach (Rune rune in second.Body.Text.EnumerateRunes())
        {
            if ((index++ & 255) == 0) { cancellationToken.ThrowIfCancellationRequested(); }
            double width = tailFont.GetAdvanceWidth(tailFont.MapCodePoint(rune.Value)) * fontSize / tailFont.UnitsPerEm;
            if (!double.IsFinite(width) || width > continuationWidth) { return null; }
        }
        double? firstGap = MeasureWordCompatibleParagraphTransition(paragraphs, fontSize);
        double tailGap = DocxLineMetrics.MeasureHheaLineHeight(tailFont, fontSize);
        if (firstGap is null || !double.IsFinite(tailGap) || tailGap <= 0d) { return null; }
        var lines = WrapUniformBalloonWords(second.Body.Text, second.Body.Resource.Embedded, fontSize,
            continuationWidth, continuationWidth, cancellationToken, reserveFirstRowBreakSpace: false);
        return lines.Length >= (firstRows is null && thirdRows is null ? 2 : 1) && lines[0].Text.Length != 0
            ? new(first, second, lines, firstGap.Value, tailGap, third, thirdGap, thirdRows, thirdTailGap, firstRows, firstTailGap,
                fourth, fourthGap, fourthRows, fourthTailGap) : null;
    }

    private static void RenderWordCompatibleWrappedParagraphRows(
        DocxMarkupWrappedParagraphRows rows, DocxMarkupBalloonPlacement placement, PdfGraphicsBuilder graphics,
        double firstX, double continuationX, double firstY, double fontSize, CancellationToken cancellationToken)
    {
        if (rows.FirstRows is { } firstRows)
        {
            RenderUniformBalloonRows(firstRows, placement, graphics, rows.First.Body.Resource, rows.First.Mark,
                firstX, continuationX, firstY, rows.FirstTailGap, fontSize, cancellationToken);
        }
        else
        {
            DrawBalloonText(graphics, rows.First.Body.Resource, rows.First.Body.Text, firstX, firstY, fontSize,
                placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
            DrawBalloonText(graphics, rows.First.Mark, " ", firstX + rows.First.Body.Resource.Embedded.MeasureTextPoints(rows.First.Body.Text, fontSize),
                firstY, fontSize, placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
        }
        RenderUniformBalloonRows(rows.TailRows, placement, graphics, rows.Second.Body.Resource, rows.Second.Mark,
            continuationX, continuationX, firstY - rows.FirstContinuationsHeight - rows.FirstGap, rows.TailGap, fontSize, cancellationToken);
        if (rows.Third is { } third)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (rows.ThirdRows is { } thirdRows)
            {
                double thirdY = firstY - rows.FirstContinuationsHeight - rows.FirstGap - (rows.TailRows.Length - 1) * rows.TailGap - rows.ThirdGap!.Value;
                RenderUniformBalloonRows(thirdRows, placement, graphics, third.Body.Resource, third.Mark,
                    continuationX, continuationX, thirdY, rows.ThirdTailGap, fontSize, cancellationToken);
            }
            else
            {
                double y = firstY - rows.ContinuationsHeight + (rows.FourthGap ?? 0d) + rows.FourthContinuationsHeight;
                DrawBalloonText(graphics, third.Body.Resource, third.Body.Text, continuationX, y, fontSize,
                    placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
                DrawBalloonText(graphics, third.Mark, " ", continuationX + third.Body.Resource.Embedded.MeasureTextPoints(third.Body.Text, fontSize),
                    y, fontSize, placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
            }
        }
        if (rows.Fourth is { } fourth)
        {
            cancellationToken.ThrowIfCancellationRequested();
            double y = firstY - rows.ContinuationsHeight + rows.FourthContinuationsHeight;
            if (rows.FourthRows is { } fourthRows)
            {
                RenderUniformBalloonRows(fourthRows, placement, graphics, fourth.Body.Resource, fourth.Mark,
                    continuationX, continuationX, y, rows.FourthTailGap, fontSize, cancellationToken);
            }
            else
            {
                DrawBalloonText(graphics, fourth.Body.Resource, fourth.Body.Text, continuationX, y, fontSize,
                    placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
                DrawBalloonText(graphics, fourth.Mark, " ", continuationX + fourth.Body.Resource.Embedded.MeasureTextPoints(fourth.Body.Text, fontSize),
                    y, fontSize, placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
            }
        }
    }

    private static void RenderWordCompatibleParagraphs(
        IReadOnlyList<DocxMarkupBalloonParagraph> paragraphs, DocxMarkupBalloonPlacement placement,
        PdfGraphicsBuilder graphics, double firstX, double continuationX,
        double firstY, double gap, double fontSize, CancellationToken cancellationToken)
    {
        for (int index = 0; index < 2; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxMarkupBalloonParagraph paragraph = paragraphs[index];
            double x = index == 0 ? firstX : continuationX;
            double y = firstY - index * gap;
            DrawBalloonText(graphics, paragraph.Body.Resource, paragraph.Body.Text, x, y, fontSize,
                placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
            DrawBalloonText(graphics, paragraph.Mark, " ", x + paragraph.Body.Resource.Embedded.MeasureTextPoints(paragraph.Body.Text, fontSize),
                y, fontSize, placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
        }
    }
}
