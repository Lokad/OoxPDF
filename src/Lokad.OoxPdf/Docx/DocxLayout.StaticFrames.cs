namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    private static DocxStaticStoryLayoutResult CreatePageStaticStoryLayout(
        DocxLayoutPage page, DocxSelectedStaticStory story, double x, double width, double startY,
        bool isHeader, int pageNumber, int pageCount, IDocxTextMeasurer textMeasurer,
        IDocxStaticTextMetricsProvider staticMetrics, double defaultTabStopPoints, double scale,
        IDocxLineMetricsProvider? unscaledLineMetrics, CancellationToken token,
        IReadOnlyDictionary<DocxParagraph, DocxPreparedMathParagraph>? mathParagraphs,
        IDocxTextMeasurer? unscaledTextMeasurer)
    {
        bool canMap = scale is > 0d and < 1d && double.IsFinite(width / scale) &&
            page.ColumnFrames.Count <= 1 &&
            page.Items.FirstOrDefault() is DocxTextLineLayout &&
            unscaledTextMeasurer is IDocxStaticTextMetricsProvider &&
            page.PageSettings.HeaderFloatingDrawingsByType.Values.All(drawings => drawings.Count == 0) &&
            page.PageSettings.FooterFloatingDrawingsByType.Values.All(drawings => drawings.Count == 0) &&
            IsNaturalStaticStory(story, width / scale, mathParagraphs, token) &&
            IsNaturalStaticStory(SelectStaticHeaderFooter(
                isHeader ? page.PageSettings.FooterBodyElementsByType : page.PageSettings.HeaderBodyElementsByType,
                isHeader ? page.PageSettings.FooterParagraphsByType : page.PageSettings.HeaderParagraphsByType,
                page.PageSettings, pageNumber), width / scale, mathParagraphs, token);
        if (!canMap)
            return CreateStaticStoryLayout(story, x, width, startY, isHeader, pageNumber, pageCount,
                textMeasurer, staticMetrics, defaultTabStopPoints, scale, unscaledLineMetrics, token, mathParagraphs);

        // Lay out one source block using the existing natural metrics. Mapping
        // its complete frame keeps the first inset, pitches and footer anchor
        // in the same coordinate system, including prepared mathematical boxes.
        var source = CreateStaticStoryLayout(story, x, width / scale, startY, isHeader, pageNumber, pageCount,
            unscaledTextMeasurer!, (IDocxStaticTextMetricsProvider)unscaledTextMeasurer!,
            defaultTabStopPoints, 1d, unscaledLineMetrics, token, mathParagraphs);
        if (!HasVisibleStaticContent(source)) { return source; }
        double center = page.Height / 2d;
        double frameAnchor = isHeader ? page.Height - page.MarginTop : page.MarginBottom;
        var lines = new List<DocxTextLineLayout>(source.TextLines.Count);
        foreach (var line in source.TextLines)
        {
            token.ThrowIfCancellationRequested();
            lines.Add(line with
            {
                X = x + (line.X - x) * scale,
                BaselineY = center + (line.BaselineY - center) * scale,
                Width = line.Width * scale,
                Segments = line.Segments.Select(segment =>
                {
                    token.ThrowIfCancellationRequested();
                    return segment with { X = x + (segment.X - x) * scale, Width = segment.Width * scale };
                }).ToArray(),
                LineHeight = line.LineHeight * scale,
                AppliedBeforeSpacing = line.AppliedBeforeSpacing * scale,
                PendingAfterSpacing = line.PendingAfterSpacing * scale,
                ParagraphBeforeSpacing = line.ParagraphBeforeSpacing * scale,
                ParagraphAfterSpacing = line.ParagraphAfterSpacing * scale,
                MathScale = line.MathScale * scale,
                StaticStoryUsesPhysicalY = true
            });
        }
        return source with
        {
            TextLines = lines,
            EndCursorY = frameAnchor + (source.EndCursorY - frameAnchor) * scale,
            EndPendingAfterSpacing = source.EndPendingAfterSpacing * scale,
            MappedFooterContentTop = isHeader ? null : frameAnchor +
                (ResolveStaticFooterContentTop(source, unscaledLineMetrics) - frameAnchor) * scale
        };
    }

    private static bool IsNaturalStaticStory(DocxSelectedStaticStory story, double width,
        IReadOnlyDictionary<DocxParagraph, DocxPreparedMathParagraph>? mathParagraphs, CancellationToken token)
    {
        foreach (var element in story.BodyElements)
        {
            token.ThrowIfCancellationRequested();
            if (element is not DocxParagraphElement { Revisions.Count: 0, Paragraph: var paragraph }) { return false; }
            if (paragraph.Equations.Count != 0)
            {
                if (!TryGetStandaloneMath(paragraph, width, 1d, mathParagraphs, out _)) { return false; }
            }
            else if (paragraph.Alignment is not (DocxTextAlignment.Left or DocxTextAlignment.Center or DocxTextAlignment.Right) ||
                !HasNaturalStaticParagraphContent(paragraph)) { return false; }
        }
        return true;
    }
}
