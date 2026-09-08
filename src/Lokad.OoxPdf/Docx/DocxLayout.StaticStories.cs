using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    private static IReadOnlyList<DocxLayoutPage> AddStaticContent(
        IReadOnlyList<DocxLayoutPage> pages,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        double paragraphSpacingScale,
        CancellationToken cancellationToken)
    {
        if (textMeasurer is not IDocxStaticTextMetricsProvider staticMetrics)
        {
            return pages;
        }

        var pagesWithStaticText = new DocxLayoutPage[pages.Count];
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxLayoutPage page = pages[pageIndex];
            int pageNumber = pageIndex + 1;
            double bodyWidth = Math.Max(1d, page.Width - page.MarginLeft - page.MarginRight);
            DocxSelectedStaticStory selectedHeader = SelectStaticHeaderFooter(
                page.PageSettings.HeaderBodyElementsByType,
                page.PageSettings.HeaderParagraphsByType,
                page.PageSettings,
                pageNumber);
            DocxSelectedStaticStory selectedFooter = SelectStaticHeaderFooter(
                page.PageSettings.FooterBodyElementsByType,
                page.PageSettings.FooterParagraphsByType,
                page.PageSettings,
                pageNumber);
            DocxStaticStoryLayoutResult headerLayout = CreateStaticStoryLayout(
                    selectedHeader,
                    page.MarginLeft,
                    bodyWidth,
                    page.Height - ResolveHeaderDistance(page),
                    true,
                    pageNumber,
                    pages.Count,
                    textMeasurer,
                    staticMetrics,
                    defaultTabStopPoints,
                    paragraphSpacingScale,
                    cancellationToken);
            DocxStaticStoryLayoutResult footerLayout = CreateStaticStoryLayout(
                    selectedFooter,
                    page.MarginLeft,
                    bodyWidth,
                    ResolveFooterDistance(page),
                    false,
                    pageNumber,
                    pages.Count,
                    textMeasurer,
                    staticMetrics,
                    defaultTabStopPoints,
                    paragraphSpacingScale,
                    cancellationToken);
            pagesWithStaticText[pageIndex] = page with
            {
                StaticTextLines = headerLayout.TextLines.Concat(footerLayout.TextLines).ToArray(),
                StaticInlineImages = headerLayout.InlineImages.Concat(footerLayout.InlineImages).ToArray(),
                StaticTableRows = headerLayout.TableRows.Concat(footerLayout.TableRows).ToArray(),
                StaticInlineTextBoxes = headerLayout.InlineTextBoxes.Concat(footerLayout.InlineTextBoxes).ToArray()
            };
        }

        return pagesWithStaticText;
    }

    private static DocxStaticStoryLayoutResult CreateStaticStoryLayout(
        DocxSelectedStaticStory story,
        double x,
        double width,
        double startY,
        bool isHeader,
        int pageNumber,
        int pageCount,
        IDocxTextMeasurer textMeasurer,
        IDocxStaticTextMetricsProvider staticMetrics,
        double defaultTabStopPoints,
        double paragraphSpacingScale,
        CancellationToken cancellationToken)
    {
        var lines = new List<DocxTextLineLayout>();
        var images = new List<DocxInlineImageLayout>();
        var tableRows = new List<DocxTableRowLayout>();
        var boxes = new List<DocxInlineTextBoxLayout>();
        double cursorY = startY;
        double pendingSpacingAfter = 0d;
        DocxParagraph? previousParagraph = null;
        int paragraphIndex = 0;
        int tableIndex = 0;
        for (int elementIndex = 0; elementIndex < story.BodyElements.Count; elementIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (story.BodyElements[elementIndex] is DocxTableElement tableElement)
            {
                cursorY -= pendingSpacingAfter;
                pendingSpacingAfter = 0d;
                previousParagraph = null;
                DocxTableLayoutFrame frame = CreateTableLayoutFrame(
                    tableElement.Table,
                    tableIndex++,
                    elementIndex,
                    x,
                    width,
                    UnpagedRelatedStoryCanvasHeightPoints,
                    textMeasurer,
                    defaultTabStopPoints,
                    cancellationToken,
                    pageNumber,
                    pageCount: null,
                    paragraphSpacingScale: paragraphSpacingScale);
                for (int rowIndex = 0; rowIndex < tableElement.Table.Rows.Count; rowIndex++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    double rowHeight = frame.RowHeights[rowIndex];
                    tableRows.Add(CreateTableRowLayout(
                        tableElement.Table,
                        frame.Context,
                        tableElement.Table.Rows[rowIndex],
                        rowIndex,
                        frame.RowHeights,
                        frame.EffectiveColumns,
                        frame.Scale,
                        textMeasurer,
                        defaultTabStopPoints,
                        () => pageNumber,
                        cursorY,
                        rowHeight,
                        cursorY,
                        FragmentIndex: 0,
                        FragmentCount: 1,
                        FragmentReason: "None",
                        StoryKind: isHeader ? "Header" : "Footer",
                        StoryVariantType: story.VariantType,
                        pageCount: pageCount,
                        paragraphSpacingScale: paragraphSpacingScale));
                    cursorY -= rowHeight;
                }

                continue;
            }

            if (story.BodyElements[elementIndex] is not DocxParagraphElement paragraphElement)
            {
                pendingSpacingAfter = 0d;
                previousParagraph = null;
                continue;
            }

            DocxParagraph paragraph = paragraphElement.Paragraph;
            DocxParagraphSpacingProfile spacingProfile = ResolveParagraphSpacingProfile(previousParagraph, paragraph, pendingSpacingAfter, paragraphSpacingScale);
            cursorY -= spacingProfile.AppliedBeforeSpacing;
            pendingSpacingAfter = 0d;
            DocxTextSpan[] spans = CreateStaticTextSpans(paragraph.Runs);
            int sourceLineIndex = 0;
            if (spans.Length != 0)
            {
                foreach (DocxWrappedTextLine line in WrapStaticTextLines(spans, width, textMeasurer))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (line.Spans.Count == 0)
                    {
                        continue;
                    }

                    double lineWidth = MeasureStaticTextSpans(line.Spans, textMeasurer);
                    double lineX = paragraph.EffectiveProperties.Alignment switch
                    {
                        DocxTextAlignment.Center => x + Math.Max(0d, width - lineWidth) / 2d,
                        DocxTextAlignment.Right => x + Math.Max(0d, width - lineWidth),
                        _ => x
                    };
                    double ascender = line.Spans.Max(span => staticMetrics.MeasureWindowsAscender(span.StyleRun, span.StyleRun.EffectiveProperties.FontSize));
                    double descender = line.Spans.Max(span => staticMetrics.MeasureWindowsDescender(span.StyleRun, span.StyleRun.EffectiveProperties.FontSize));
                    double baselineY = isHeader ? cursorY - ascender : cursorY + descender;
                    IReadOnlyList<DocxTextSegmentLayout> segments = CreateStaticTextSegments(line.Spans, lineX);
                    lines.Add(new DocxTextLineLayout(
                        line.Text,
                        line.Spans[0].StyleRun,
                        line.Spans.Max(span => span.StyleRun.EffectiveProperties.FontSize),
                        lineX,
                        baselineY,
                        lineWidth,
                        segments,
                        LineHeight: ascender + descender,
                        AppliedBeforeSpacing: sourceLineIndex == 0 ? spacingProfile.AppliedBeforeSpacing : 0d,
                        IsFirstParagraphLine: sourceLineIndex == 0,
                        SourceLineIndex: sourceLineIndex,
                        PendingAfterSpacing: sourceLineIndex == 0 ? spacingProfile.PendingAfterSpacing : null,
                        ParagraphBeforeSpacing: sourceLineIndex == 0 ? spacingProfile.ParagraphBeforeSpacing : null,
                        ParagraphAfterSpacing: sourceLineIndex == 0 ? spacingProfile.ParagraphAfterSpacing : null,
                        ContextualSpacingSuppressed: sourceLineIndex == 0 ? spacingProfile.ContextualSpacingSuppressed : null,
                        SourceParagraph: paragraph,
                        SourceParagraphIndex: paragraphIndex,
                        StoryKind: isHeader ? "Header" : "Footer",
                        StoryVariantType: story.VariantType,
                        LineHeightSource: DocxLineHeightSource.StaticWindowsExtents, SourceBlockIndex: null, EndsWithIntraTokenBreak: false, SingleLineHeight: null, ListLabelSingleLineHeight: null, BodyWindowsLineHeight: null, ListLabelWindowsLineHeight: null, EffectiveLineSpacingFactor: null, LineSpacingFactorFloorApplied: null, EmitsTerminalParagraphMark: false));
                    sourceLineIndex++;
                    cursorY -= ascender + descender;
                }
            }

            foreach (DocxInlineImage image in paragraph.Images)
            {
                cancellationToken.ThrowIfCancellationRequested();
                double imageWidth = Math.Min(width, image.WidthPoints);
                double imageHeight = image.HeightPoints * imageWidth / Math.Max(1d, image.WidthPoints);
                double imageX = paragraph.EffectiveProperties.Alignment switch
                {
                    DocxTextAlignment.Center => x + Math.Max(0d, width - imageWidth) / 2d,
                    DocxTextAlignment.Right => x + Math.Max(0d, width - imageWidth),
                    _ => x
                };
                images.Add(new DocxInlineImageLayout(
                    image,
                    imageX,
                    cursorY - imageHeight,
                    imageWidth,
                    imageHeight,
                    pageNumber,
                    SourceBlockIndex: null,
                    SourceParagraphIndex: paragraphIndex,
                    StoryKind: isHeader ? "Header" : "Footer",
                    StoryVariantType: story.VariantType));
                cursorY -= imageHeight + InlineImageParagraphGapPoints;
            }

            foreach (DocxInlineTextBox textBox in paragraph.InlineTextBoxes)
            {
                cancellationToken.ThrowIfCancellationRequested();
                DocxInlineTextBoxLayout? textBoxLayout = CreateInlineTextBoxLayout(
                    textBox,
                    sourceBlockIndex: null,
                    x,
                    width,
                    cursorY,
                    paragraph.EffectiveProperties.Alignment,
                    textMeasurer,
                    defaultTabStopPoints,
                    paragraphSpacingScale,
                    pageNumber,
                    cancellationToken,
                    sourceParagraphIndex: paragraphIndex);
                if (textBoxLayout is null)
                {
                    continue;
                }

                boxes.Add(textBoxLayout with { StoryKind = isHeader ? "Header" : "Footer", StoryVariantType = story.VariantType });
                cursorY -= textBoxLayout.BoxHeight + InlineImageParagraphGapPoints;
            }

            pendingSpacingAfter = spacingProfile.ParagraphAfterSpacing;
            previousParagraph = paragraph;
            paragraphIndex++;
        }

        return new DocxStaticStoryLayoutResult(lines.ToArray(), images.ToArray(), tableRows.ToArray(), boxes.ToArray());

        DocxTextSpan[] CreateStaticTextSpans(IReadOnlyList<DocxTextRun> runs)
        {
            if (runs.Count != 0 && runs.All(run => run.Text.Length == 0 || run.EffectiveProperties.Hidden))
            {
                for (int i = 0; i < runs.Count; i++)
                {
                    if (!runs[i].EffectiveProperties.Hidden)
                    {
                        return [CreateTextSpan(" ", runs[i], i)];
                    }
                }

                return [];
            }

            return runs
                .Select((run, index) => (run, index))
                .Where(item => !item.run.EffectiveProperties.Hidden)
                .Select(item => CreateTextSpan(ResolveStaticFieldPlaceholders(item.run.Text, pageNumber, pageCount), item.run, item.index))
                .Where(span => span.Text.Length != 0)
                .ToArray();
        }

        IReadOnlyList<DocxTextSegmentLayout> CreateStaticTextSegments(IReadOnlyList<DocxTextSpan> spans, double lineX)
        {
            var segments = new List<DocxTextSegmentLayout>(spans.Count);
            double segmentX = lineX;
            for (int i = 0; i < spans.Count; i++)
            {
                DocxTextSpan span = spans[i];
                double nominalFontSize = span.StyleRun.EffectiveProperties.FontSize;
                double layoutFontSize = DocxVerticalAlignMetrics.ResolveFontSize(nominalFontSize, span.StyleRun);
                double baselineOffset = DocxVerticalAlignMetrics.ResolveBaselineOffset(nominalFontSize, layoutFontSize, span.StyleRun);
                double segmentWidth = textMeasurer.MeasureText(span.StyleRun, span.Text, layoutFontSize);
                segments.Add(new DocxTextSegmentLayout(
                    span.Text,
                    span.StyleRun,
                    segmentX,
                    segmentWidth,
                    layoutFontSize,
                    baselineOffset,
                    SourceTextRunIndex: span.SourceTextRunIndex,
                    SourceTextOffsetInRun: span.SourceTextOffsetInRun, PdfCharacterSpacing: 0d, PdfCharacterSpacingSource: DocxTextStateCharacterSpacingSource.None, CompensatePdfCharacterSpacing: true, Role: DocxTextSegmentRole.Text));
                segmentX += segmentWidth;
                if (i + 1 < spans.Count)
                {
                    segmentX += DocxTextSpacing.BoundarySpacing(span.StyleRun, span.Text, spans[i + 1].Text);
                }
            }

            return segments;
        }
    }

    private static IEnumerable<DocxWrappedTextLine> WrapStaticTextLines(
        IReadOnlyList<DocxTextSpan> spans,
        double maxWidth,
        IDocxTextMeasurer textMeasurer)
    {
        string text = string.Concat(spans.Select(span => span.Text));
        int segmentStart = 0;
        while (segmentStart <= text.Length)
        {
            int breakIndex = text.IndexOf('\n', segmentStart);
            int segmentLength = breakIndex < 0 ? text.Length - segmentStart : breakIndex - segmentStart;
            bool yielded = false;
            foreach (DocxWrappedTextLine line in WrapStaticWords(text, spans, segmentStart, segmentLength, maxWidth, textMeasurer))
            {
                yielded = true;
                yield return line;
            }

            if (!yielded && segmentLength == 0)
            {
                yield return new DocxWrappedTextLine(string.Empty, [], false);
            }

            if (breakIndex < 0)
            {
                yield break;
            }

            segmentStart = breakIndex + 1;
        }
    }

    private static IEnumerable<DocxWrappedTextLine> WrapStaticWords(
        string text,
        IReadOnlyList<DocxTextSpan> spans,
        int segmentStart,
        int segmentLength,
        double maxWidth,
        IDocxTextMeasurer textMeasurer)
    {
        IReadOnlyList<TextToken> tokens = TokenizeSpaces(text, segmentStart, segmentLength);
        if (tokens.Count == 0)
        {
            yield break;
        }

        int lineStart = tokens[0].Start;
        int lineLength = 0;
        foreach (TextToken token in tokens)
        {
            int candidateLength = token.Start + token.Length - lineStart;
            bool lineHasNonWhitespace = HasNonWhitespace(text, lineStart, lineLength);
            if (lineLength > 0 &&
                lineHasNonWhitespace &&
                !token.IsBreakableWhitespace &&
                MeasureStaticTextSpansForWrapping(SliceTextSpans(spans, lineStart, candidateLength), textMeasurer) > maxWidth)
            {
                yield return CreateWrappedTextLine(text, spans, lineStart, lineLength, false);
                lineStart = token.Start;
                lineLength = token.Length;
            }
            else
            {
                lineLength = candidateLength;
            }
        }

        if (lineLength > 0)
        {
            yield return CreateWrappedTextLine(text, spans, lineStart, lineLength, false);
        }
    }

    private static double MeasureStaticTextSpans(IReadOnlyList<DocxTextSpan> spans, IDocxTextMeasurer textMeasurer)
    {
        double width = 0d;
        for (int i = 0; i < spans.Count; i++)
        {
            DocxTextSpan span = spans[i];
            double layoutFontSize = DocxVerticalAlignMetrics.ResolveFontSize(span.StyleRun.EffectiveProperties.FontSize, span.StyleRun);
            width += textMeasurer.MeasureText(span.StyleRun, span.Text, layoutFontSize);
            if (i + 1 < spans.Count)
            {
                width += DocxTextSpacing.BoundarySpacing(span.StyleRun, span.Text, spans[i + 1].Text);
            }
        }

        return width;
    }

    private static double MeasureStaticTextSpansForWrapping(IReadOnlyList<DocxTextSpan> spans, IDocxTextMeasurer textMeasurer)
    {
        return MeasureStaticTextSpans(NormalizeHiddenBreakSpans(spans, preserveTerminalSoftHyphen: false), textMeasurer);
    }

    private static string ResolveStaticFieldPlaceholders(string text, int pageNumber, int pageCount)
    {
        return text
            .Replace("{NUMPAGES}", pageCount.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal)
            .Replace("{PAGE}", pageNumber.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal);
    }
}
