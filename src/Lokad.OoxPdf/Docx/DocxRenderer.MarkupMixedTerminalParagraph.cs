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
        if (layout is null || fonts is null || layout.Story.BodyElements.Count < 2 ||
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

    private sealed record DocxMarkupMixedTerminalRows(
        DocxMarkupMixedTerminalParagraph Paragraphs,
        DocxMarkupParagraphRows Leading, DocxMarkupWrappedTerminalRows? Wrapped)
    {
        public double ContinuationsHeight
        {
            get
            {
                double height = Leading.ContinuationsHeight;
                return Wrapped is { } wrapped
                    ? height + (wrapped.TerminalRows.Length - (wrapped.HasClosingRow ? 1 : 2)) * wrapped.TerminalPitch +
                        wrapped.FinalPitch + (wrapped.PartialClosingSuffix is null ? 0d : wrapped.MixedTerminalPitch - wrapped.TerminalPitch)
                    : height;
            }
        }
    }

    private static DocxMarkupMixedTerminalRows? ResolveWordCompatibleMixedTerminalRows(
        DocxMarkupMixedTerminalParagraph? paragraphs, double size, double firstWidth,
        double continuationWidth, CancellationToken cancellationToken)
    {
        if (paragraphs is null ||
            ResolveWordCompatibleParagraphRows(paragraphs.Leading, size, firstWidth, continuationWidth,
                cancellationToken, hasMixedTerminal: true) is not { } leading) { return null; }
        if (MeasureFittingTerminalTransition(paragraphs, size, continuationWidth) is double finalStep)
        {
            leading.Transitions[^1] = finalStep;
            var fitting = new DocxMarkupMixedTerminalRows(paragraphs, leading, null);
            if (double.IsFinite(fitting.ContinuationsHeight)) { return fitting; }
        }
        var wrapped = ResolveWordCompatibleWrappedTerminalRows(paragraphs, size, continuationWidth, cancellationToken,
            out double firstStep);
        if (wrapped is null) { return null; }
        leading.Transitions[^1] = firstStep;
        var result = new DocxMarkupMixedTerminalRows(paragraphs, leading, wrapped);
        return double.IsFinite(result.ContinuationsHeight) ? result : null;
    }

    private static double? MeasureFittingTerminalTransition(
        DocxMarkupMixedTerminalParagraph paragraphs, double size, double continuationWidth)
    {
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
        var previous = paragraphs.Leading[^1].Body.Resource.Embedded.Font;
        double finalStep = (-previous.Hhea.HorizontalDescender / (double)previous.UnitsPerEm + nextAscent) * size;
        return double.IsFinite(finalStep) && finalStep > 0d ? finalStep : null;
    }

    private static void RenderWordCompatibleMixedTerminalRows(
        DocxMarkupMixedTerminalRows rows, DocxMarkupBalloonPlacement placement, PdfGraphicsBuilder graphics,
        double firstX, double continuationX, double firstY, double size, CancellationToken cancellationToken)
    {
        double y = firstY;
        var leading = rows.Leading;
        for (int index = 0; index < leading.Paragraphs.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var paragraph = leading.Paragraphs[index];
            RenderUniformBalloonRows(leading.BodyRows[index], placement, graphics, paragraph.Body.Resource, paragraph.Mark,
                index == 0 ? firstX : continuationX, continuationX, y, leading.Pitches[index], size, cancellationToken);
            y -= (leading.BodyRows[index].Length - 1) * leading.Pitches[index] + leading.Transitions[index];
        }
        if (rows.Wrapped is { } wrapped)
        {
            RenderWordCompatibleWrappedTerminalRows(rows.Paragraphs, wrapped, placement, graphics,
                continuationX, y, size, cancellationToken);
            return;
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
