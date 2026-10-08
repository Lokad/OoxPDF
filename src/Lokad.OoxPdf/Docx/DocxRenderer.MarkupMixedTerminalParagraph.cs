using System.Text;
using Lokad.OoxPdf.Pdf;
using static Lokad.OoxPdf.Docx.DocxBalloonParagraphPolicy;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    private sealed record DocxMarkupMixedTerminalParagraph(
        IReadOnlyList<DocxMarkupBalloonParagraph> Leading,
        DocxMarkupBalloonBodyPart[] Parts, DocxRunFontResource Mark);

    private static DocxMarkupMixedTerminalParagraph? ResolveCommentMixedTerminalParagraph(
        DocxRelatedStoryLayout? layout, string preview, DocxFontResources? fonts, CancellationToken cancellationToken)
    {
        if (layout is null || fonts is null || layout.Story.BodyElements.Count is not (2 or 3 or 4) ||
            layout.InlineImages.Count != 0 || layout.FloatingDrawings.Count != 0 ||
            layout.Story.BodyElements[^1] is not DocxParagraphElement last ||
            !HasPlainBalloonParagraphShape(last.Paragraph) ||
            last.Paragraph.ParagraphMarkRun is not { } mark || !IsPlainBalloonParagraphRun(mark) ||
            !fonts.RunResources.TryGetValue(mark, out DocxRunFontResource? markResource) ||
            markResource.Resolution.Bold || markResource.Resolution.Italic || markResource.Resolution.IsFallback ||
            !HasEncodedBalloonParagraphText(markResource, " ", cancellationToken)) { return null; }
        var elements = layout.Story.BodyElements.Take(layout.Story.BodyElements.Count - 1).ToArray();
        if (elements.Any(element => element is not DocxParagraphElement)) { return null; }
        string leadingPreview = string.Join(" ", elements.Cast<DocxParagraphElement>()
            .Select(element => string.Concat(element.Paragraph.Runs.Select(run => run.Text)).Trim(' ')));
        var leading = ResolveCommentBalloonParagraphs(layout with { Story = layout.Story with { BodyElements = elements } },
            leadingPreview, fonts, cancellationToken, allowSingleParagraph: elements.Length == 1);
        if (leading is null) { return null; }
        DocxTextRun[] runs = last.Paragraph.Runs.Where(run => run.Text.Length != 0).ToArray();
        if (runs.Length != 2) { return null; }
        var parts = new DocxMarkupBalloonBodyPart[2];
        for (int index = 0; index < 2; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxTextRun run = runs[index];
            string text = index == 0 ? run.Text.TrimStart(' ') : run.Text.TrimEnd(' ');
            if (text.Length == 0 || !IsPlainBalloonParagraphRun(run) ||
                !fonts.RunResources.TryGetValue(run, out DocxRunFontResource? resource) ||
                resource.Resolution.Bold || resource.Resolution.Italic || resource.Resolution.IsFallback ||
                !HasEncodedBalloonParagraphText(resource, text, cancellationToken)) { return null; }
            parts[index] = new(text, resource);
        }
        string terminalText = parts[0].Text + parts[1].Text;
        if (ReferenceEquals(parts[0].Resource.Embedded.Font, parts[1].Resource.Embedded.Font) ||
            terminalText.Contains("  ", StringComparison.Ordinal) || leadingPreview + " " + terminalText != preview) { return null; }
        return new(leading, parts, markResource);
    }

    private sealed record DocxMarkupFittingMixedTerminalRows(
        DocxMarkupMixedTerminalParagraph Paragraphs,
        DocxUniformBalloonRow[][] LeadingRows, double[] Pitches, double[] Transitions)
    {
        public double ContinuationsHeight => Enumerable.Range(0, Paragraphs.Leading.Count)
            .Sum(index => (LeadingRows[index].Length - 1) * Pitches[index]) + Transitions.Sum();
    }

    private static DocxMarkupFittingMixedTerminalRows? ResolveWordCompatibleFittingMixedTerminalRows(
        DocxMarkupMixedTerminalParagraph? paragraphs, double size, double firstWidth,
        double continuationWidth, CancellationToken cancellationToken)
    {
        if (paragraphs is null || !double.IsFinite(firstWidth) || firstWidth <= 0d ||
            !double.IsFinite(continuationWidth) || continuationWidth <= 0d) { return null; }
        int leadingCount = paragraphs.Leading.Count;
        if (leadingCount is not (2 or 3)) { return null; }
        var rows = new DocxUniformBalloonRow[leadingCount][];
        var pitches = new double[leadingCount];
        var transitions = new double[leadingCount];
        for (int index = 0; index < leadingCount; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var body = paragraphs.Leading[index].Body;
            var font = body.Resource.Embedded.Font;
            double width = body.Resource.Embedded.MeasureTextPoints(body.Text, size);
            double pitch = DocxLineMetrics.MeasureHheaLineHeight(font, size);
            if (font.UnitsPerEm <= 0 || !double.IsFinite(width) || width <= 0d || !double.IsFinite(pitch) || pitch <= 0d) { return null; }
            int scalarIndex = 0;
            foreach (Rune rune in body.Text.EnumerateRunes())
            {
                if ((scalarIndex++ & 255) == 0) { cancellationToken.ThrowIfCancellationRequested(); }
                double scalarWidth = font.GetAdvanceWidth(font.MapCodePoint(rune.Value)) * size / font.UnitsPerEm;
                if (!double.IsFinite(scalarWidth) || scalarWidth > continuationWidth) { return null; }
            }
            rows[index] = WrapUniformBalloonWords(body.Text, body.Resource.Embedded, size,
                index == 0 ? firstWidth : continuationWidth, continuationWidth, cancellationToken,
                reserveFirstRowBreakSpace: false);
            if (rows[index].Length == 0 || rows[index][0].Text.Length == 0) { return null; }
            pitches[index] = pitch;
            if (index > 0)
            {
                double? step = MeasureWordCompatibleParagraphTransition(paragraphs.Leading[index - 1], paragraphs.Leading[index], size);
                if (step is null) { return null; }
                transitions[index - 1] = step.Value;
            }
        }
        double totalWidth = 0d;
        double nextAscent = double.NegativeInfinity;
        foreach (var part in paragraphs.Parts)
        {
            var font = part.Resource.Embedded.Font;
            double width = part.Resource.Embedded.MeasureTextPoints(part.Text, size);
            double pitch = DocxLineMetrics.MeasureHheaLineHeight(font, size);
            if (font.UnitsPerEm <= 0 || !double.IsFinite(width) || width <= 0d ||
                !double.IsFinite(pitch) || pitch <= 0d) { return null; }
            totalWidth += width;
            nextAscent = Math.Max(nextAscent, (font.Hhea.HorizontalAscender + font.Hhea.HorizontalLineGap) / (double)font.UnitsPerEm);
        }
        if (!double.IsFinite(totalWidth) || totalWidth > continuationWidth) { return null; }
        var previous = paragraphs.Leading[leadingCount - 1].Body.Resource.Embedded.Font;
        double finalStep = (-previous.Hhea.HorizontalDescender / (double)previous.UnitsPerEm + nextAscent) * size;
        if (!double.IsFinite(finalStep) || finalStep <= 0d) { return null; }
        transitions[leadingCount - 1] = finalStep;
        var result = new DocxMarkupFittingMixedTerminalRows(paragraphs, rows, pitches, transitions);
        return double.IsFinite(result.ContinuationsHeight) ? result : null;
    }

    private static void RenderWordCompatibleFittingMixedTerminalRows(
        DocxMarkupFittingMixedTerminalRows rows, DocxMarkupBalloonPlacement placement, PdfGraphicsBuilder graphics,
        double firstX, double continuationX, double firstY, double size, CancellationToken cancellationToken)
    {
        double y = firstY;
        for (int index = 0; index < rows.Paragraphs.Leading.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var paragraph = rows.Paragraphs.Leading[index];
            RenderUniformBalloonRows(rows.LeadingRows[index], placement, graphics, paragraph.Body.Resource, paragraph.Mark,
                index == 0 ? firstX : continuationX, continuationX, y, rows.Pitches[index], size, cancellationToken);
            y -= (rows.LeadingRows[index].Length - 1) * rows.Pitches[index] + rows.Transitions[index];
        }
        double x = continuationX;
        foreach (var part in rows.Paragraphs.Parts)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DrawBalloonText(graphics, part.Resource, part.Text, x, y, size,
                placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
            x += part.Resource.Embedded.MeasureTextPoints(part.Text, size);
        }
        DrawBalloonText(graphics, rows.Paragraphs.Mark, " ", x, y, size,
            placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
    }
}
