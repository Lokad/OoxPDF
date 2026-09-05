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
                StaticTableRows = headerLayout.TableRows.Concat(footerLayout.TableRows).ToArray()
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
            DocxTextSpan[] spans = CreateStaticTextSpans(paragraph.Runs, pageNumber, pageCount);
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
                    IReadOnlyList<DocxTextSegmentLayout> segments = CreateStaticTextSegments(line.Spans, lineX, textMeasurer);
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

            pendingSpacingAfter = spacingProfile.ParagraphAfterSpacing;
            previousParagraph = paragraph;
            paragraphIndex++;
        }

        return new DocxStaticStoryLayoutResult(lines.ToArray(), images.ToArray(), tableRows.ToArray());
    }

    private static DocxTextSpan[] CreateStaticTextSpans(IReadOnlyList<DocxTextRun> runs, int pageNumber, int pageCount)
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
                yield return new DocxWrappedTextLine(string.Empty, []);
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

    private static IReadOnlyList<DocxTextSegmentLayout> CreateStaticTextSegments(
        IReadOnlyList<DocxTextSpan> spans,
        double lineX,
        IDocxTextMeasurer textMeasurer)
    {
        var segments = new List<DocxTextSegmentLayout>(spans.Count);
        double segmentX = lineX;
        for (int i = 0; i < spans.Count; i++)
        {
            DocxTextSpan span = spans[i];
            double nominalFontSize = span.StyleRun.EffectiveProperties.FontSize;
            double layoutFontSize = DocxVerticalAlignMetrics.ResolveFontSize(nominalFontSize, span.StyleRun);
            double baselineOffset = DocxVerticalAlignMetrics.ResolveBaselineOffset(nominalFontSize, layoutFontSize, span.StyleRun);
            double width = textMeasurer.MeasureText(span.StyleRun, span.Text, layoutFontSize);
            segments.Add(new DocxTextSegmentLayout(
                span.Text,
                span.StyleRun,
                segmentX,
                width,
                layoutFontSize,
                baselineOffset,
                SourceTextRunIndex: span.SourceTextRunIndex,
                SourceTextOffsetInRun: span.SourceTextOffsetInRun, PdfCharacterSpacing: 0d, PdfCharacterSpacingSource: DocxTextStateCharacterSpacingSource.None, CompensatePdfCharacterSpacing: true, Role: DocxTextSegmentRole.Text));
            segmentX += width;
            if (i + 1 < spans.Count)
            {
                segmentX += DocxTextSpacing.BoundarySpacing(span.StyleRun, span.Text, spans[i + 1].Text);
            }
        }

        return segments;
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

    private static DocxSelectedStaticStory SelectStaticHeaderFooter(
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> bodyElementsByType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> paragraphsByType,
        DocxPageSettings settings,
        int pageNumber)
    {
        if (settings.TitlePage == true &&
            pageNumber == 1 &&
            DocxBlockTraversal.TryGetStaticStoryBodyElements("first", bodyElementsByType, paragraphsByType, out IReadOnlyList<DocxBodyElement>? first))
        {
            return new DocxSelectedStaticStory(first, "first");
        }

        if (settings.EvenAndOddHeaders == true &&
            pageNumber % 2 == 0 &&
            DocxBlockTraversal.TryGetStaticStoryBodyElements("even", bodyElementsByType, paragraphsByType, out IReadOnlyList<DocxBodyElement>? even))
        {
            return new DocxSelectedStaticStory(even, "even");
        }

        return DocxBlockTraversal.TryGetStaticStoryBodyElements("default", bodyElementsByType, paragraphsByType, out IReadOnlyList<DocxBodyElement>? defaults)
            ? new DocxSelectedStaticStory(defaults, "default")
            : new DocxSelectedStaticStory([], null);
    }

    private sealed record DocxSelectedStaticStory(IReadOnlyList<DocxBodyElement> BodyElements, string? VariantType);

    private static DocxSelectedStaticDrawings SelectStaticHeaderFooterDrawings(
        IReadOnlyDictionary<string, IReadOnlyList<DocxFloatingDrawing>> drawingsByType,
        DocxPageSettings settings,
        int pageNumber)
    {
        if (settings.TitlePage == true &&
            pageNumber == 1 &&
            drawingsByType.TryGetValue("first", out IReadOnlyList<DocxFloatingDrawing>? first))
        {
            return new DocxSelectedStaticDrawings(first, "first");
        }

        if (settings.EvenAndOddHeaders == true &&
            pageNumber % 2 == 0 &&
            drawingsByType.TryGetValue("even", out IReadOnlyList<DocxFloatingDrawing>? even))
        {
            return new DocxSelectedStaticDrawings(even, "even");
        }

        return drawingsByType.TryGetValue("default", out IReadOnlyList<DocxFloatingDrawing>? defaults)
            ? new DocxSelectedStaticDrawings(defaults, "default")
            : new DocxSelectedStaticDrawings([], null);
    }

    private sealed record DocxSelectedStaticDrawings(IReadOnlyList<DocxFloatingDrawing> Drawings, string? VariantType);

    private static double ResolveHeaderDistance(DocxLayoutPage page)
    {
        return page.PageSettings.HeaderDistancePoints ?? Math.Max(18d, page.MarginTop / 2d);
    }

    private static double ResolveFooterDistance(DocxLayoutPage page)
    {
        return page.PageSettings.FooterDistancePoints ?? Math.Max(18d, page.MarginBottom / 2d);
    }

    private static string ResolveStaticFieldPlaceholders(string text, int pageNumber, int pageCount)
    {
        return text
            .Replace("{NUMPAGES}", pageCount.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal)
            .Replace("{PAGE}", pageNumber.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal);
    }

    private static IReadOnlyDictionary<int, DocxEffectiveSectionSettings> BuildEffectiveSectionSettings(DocxDocument document, out DocxEffectiveSectionSettings finalSectionSettings)
    {
        var sectionSettingsByElementIndex = new Dictionary<int, DocxEffectiveSectionSettings>();
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> inheritedHeadersByType =
            new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase);
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> inheritedFootersByType =
            new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase);
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> inheritedHeaderBodyElementsByType =
            new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase);
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> inheritedFooterBodyElementsByType =
            new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase);

        for (int elementIndex = 0; elementIndex < document.BodyElements.Count; elementIndex++)
        {
            if (document.BodyElements[elementIndex] is not DocxSectionBreakElement sectionBreak)
            {
                continue;
            }

            DocxPageSettings effectiveSettings = ResolveEffectiveSectionSettings(
                sectionBreak.PageSettings,
                inheritedHeadersByType,
                inheritedFootersByType,
                inheritedHeaderBodyElementsByType,
                inheritedFooterBodyElementsByType);
            sectionSettingsByElementIndex[elementIndex] = new DocxEffectiveSectionSettings(
                effectiveSettings,
                CreateSectionLayoutProperties(sectionBreak));
            inheritedHeadersByType = effectiveSettings.HeaderParagraphsByType;
            inheritedFootersByType = effectiveSettings.FooterParagraphsByType;
            inheritedHeaderBodyElementsByType = effectiveSettings.HeaderBodyElementsByType;
            inheritedFooterBodyElementsByType = effectiveSettings.FooterBodyElementsByType;
        }

        finalSectionSettings = new DocxEffectiveSectionSettings(
            ResolveEffectiveSectionSettings(
                BuildFinalSectionSettings(document),
                inheritedHeadersByType,
                inheritedFootersByType,
                inheritedHeaderBodyElementsByType,
                inheritedFooterBodyElementsByType),
            document.FinalSectionBreak is null
                ? new DocxSectionLayoutProperties(null, null, null, null, null, null, [])
                : CreateSectionLayoutProperties(document.FinalSectionBreak));
        return sectionSettingsByElementIndex;
    }

    private static DocxSectionLayoutProperties CreateSectionLayoutProperties(DocxSectionBreakElement sectionBreak)
    {
        return new DocxSectionLayoutProperties(
            sectionBreak.TypeValue?.ToValueString(),
            sectionBreak.ColumnCountValue,
            sectionBreak.ColumnEqualWidthValue,
            sectionBreak.ColumnSpaceValue,
            ReadOptionalInt32Value(sectionBreak.ColumnCountValue),
            ReadOptionalTwipsValue(sectionBreak.ColumnSpaceValue),
            sectionBreak.ColumnDefinitions
                .Select(column => new DocxSectionColumnLayoutProperties(
                    column.WidthValue,
                    column.SpaceValue,
                    ReadOptionalTwipsValue(column.WidthValue),
                    ReadOptionalTwipsValue(column.SpaceValue)))
                .ToArray());
    }

    private static DocxPageSettings BuildFinalSectionSettings(DocxDocument document)
    {
        DocxPageSettings settings = document.PageSettings;
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> headersByType = settings.HeaderParagraphsByType.Count == 0 && document.HeaderParagraphsByType.Count > 0
            ? document.HeaderParagraphsByType
            : settings.HeaderParagraphsByType;
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> footersByType = settings.FooterParagraphsByType.Count == 0 && document.FooterParagraphsByType.Count > 0
            ? document.FooterParagraphsByType
            : settings.FooterParagraphsByType;
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> headerBodyElementsByType = settings.HeaderBodyElementsByType.Count == 0 && document.HeaderBodyElementsByType.Count > 0
            ? document.HeaderBodyElementsByType
            : settings.HeaderBodyElementsByType;
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> footerBodyElementsByType = settings.FooterBodyElementsByType.Count == 0 && document.FooterBodyElementsByType.Count > 0
            ? document.FooterBodyElementsByType
            : settings.FooterBodyElementsByType;

        if (headersByType.Count == 0 && document.HeaderParagraphs.Count > 0)
        {
            headersByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = document.HeaderParagraphs
            };
            headerBodyElementsByType = ToStaticBodyElementsByType(headersByType);
        }

        if (footersByType.Count == 0 && document.FooterParagraphs.Count > 0)
        {
            footersByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = document.FooterParagraphs
            };
            footerBodyElementsByType = ToStaticBodyElementsByType(footersByType);
        }

        return settings with
        {
            HeaderParagraphsByType = headersByType,
            FooterParagraphsByType = footersByType,
            HeaderBodyElementsByType = headerBodyElementsByType.Count == 0 ? ToStaticBodyElementsByType(headersByType) : headerBodyElementsByType,
            FooterBodyElementsByType = footerBodyElementsByType.Count == 0 ? ToStaticBodyElementsByType(footersByType) : footerBodyElementsByType
        };
    }

    private static DocxPageSettings ResolveEffectiveSectionSettings(
        DocxPageSettings settings,
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> inheritedHeadersByType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> inheritedFootersByType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> inheritedHeaderBodyElementsByType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> inheritedFooterBodyElementsByType)
    {
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> headersByType = MergeInheritedStaticParagraphs(inheritedHeadersByType, settings.HeaderParagraphsByType);
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> footersByType = MergeInheritedStaticParagraphs(inheritedFootersByType, settings.FooterParagraphsByType);
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> headerBodyElementsByType = MergeInheritedStaticBodyElements(inheritedHeaderBodyElementsByType, settings.HeaderBodyElementsByType);
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> footerBodyElementsByType = MergeInheritedStaticBodyElements(inheritedFooterBodyElementsByType, settings.FooterBodyElementsByType);

        return settings with
        {
            HeaderParagraphsByType = headersByType,
            FooterParagraphsByType = footersByType,
            HeaderBodyElementsByType = headerBodyElementsByType.Count == 0 ? ToStaticBodyElementsByType(headersByType) : headerBodyElementsByType,
            FooterBodyElementsByType = footerBodyElementsByType.Count == 0 ? ToStaticBodyElementsByType(footersByType) : footerBodyElementsByType
        };
    }

    private static IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> ToStaticBodyElementsByType(
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> paragraphsByType)
    {
        return paragraphsByType.ToDictionary(
            pair => pair.Key,
            pair => (IReadOnlyList<DocxBodyElement>)pair.Value.Select(DocxBodyElementFactory.CreateParagraph).Cast<DocxBodyElement>().ToArray(),
            StringComparer.OrdinalIgnoreCase);
    }

    private static IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> MergeInheritedStaticBodyElements(
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> inheritedBodyElementsByType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> localBodyElementsByType)
    {
        if (inheritedBodyElementsByType.Count == 0)
        {
            return localBodyElementsByType;
        }

        if (localBodyElementsByType.Count == 0)
        {
            return inheritedBodyElementsByType;
        }

        var merged = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(inheritedBodyElementsByType, StringComparer.OrdinalIgnoreCase);
        foreach ((string type, IReadOnlyList<DocxBodyElement> bodyElements) in localBodyElementsByType)
        {
            merged[type] = bodyElements;
        }

        return merged;
    }

    private static IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> MergeInheritedStaticParagraphs(
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> inheritedParagraphsByType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> localParagraphsByType)
    {
        if (inheritedParagraphsByType.Count == 0)
        {
            return localParagraphsByType;
        }

        if (localParagraphsByType.Count == 0)
        {
            return inheritedParagraphsByType;
        }

        var merged = new Dictionary<string, IReadOnlyList<DocxParagraph>>(inheritedParagraphsByType, StringComparer.OrdinalIgnoreCase);
        foreach ((string type, IReadOnlyList<DocxParagraph> paragraphs) in localParagraphsByType)
        {
            merged[type] = paragraphs;
        }

        return merged;
    }

    private static DocxEffectiveSectionSettings? FindSectionSettingsAtOrAfter(
        IReadOnlyList<DocxBodyElement> elements,
        int startIndex,
        IReadOnlyDictionary<int, DocxEffectiveSectionSettings> sectionSettingsByElementIndex)
    {
        for (int i = Math.Max(0, startIndex); i < elements.Count; i++)
        {
            if (elements[i] is DocxSectionBreakElement && sectionSettingsByElementIndex.TryGetValue(i, out DocxEffectiveSectionSettings? settings))
            {
                return settings;
            }
        }

        return null;
    }

    private static DocxPageGeometry ResolveSectionGeometry(
        DocxDocument document,
        DocxEffectiveSectionSettings section,
        bool reserveMarkupMargin,
        int pageNumber)
    {
        DocxPageSettings effectiveSettings = section.PageSettings;
        double width = ReadTwipsValue(effectiveSettings.WidthValue, document.PageWidthPoints);
        double height = ReadTwipsValue(effectiveSettings.HeightValue, document.PageHeightPoints);
        (width, height) = NormalizePageSize(width, height);
        if (effectiveSettings.OrientationValue?.Equals("landscape", StringComparison.OrdinalIgnoreCase) == true && height > width)
        {
            (width, height) = (height, width);
        }

        double marginLeft = ReadTwipsValue(effectiveSettings.MarginLeftValue, document.MarginLeftPoints);
        double marginRight = ReadTwipsValue(effectiveSettings.MarginRightValue, document.MarginRightPoints);
        double gutter = Math.Max(0d, ReadTwipsValue(effectiveSettings.GutterDistanceValue, effectiveSettings.GutterDistancePoints ?? 0d));
        if (gutter > 0d)
        {
            if (ShouldApplyGutterToRightMargin(document, pageNumber))
            {
                marginRight += gutter;
            }
            else
            {
                marginLeft += gutter;
            }
        }

        double authoredMarginLeft = marginLeft;
        double authoredMarginRight = marginRight;
        if (reserveMarkupMargin)
        {
            if (ShouldReserveLeftMarkupMargin(document, pageNumber))
            {
                marginLeft = ResolveReservedMarkupLeftMargin(width, marginLeft, marginRight);
            }
            else
            {
                marginRight = ResolveReservedMarkupRightMargin(width, marginLeft, marginRight);
            }
        }

        double markupMarginReservePoints = Math.Max(0d, Math.Max(marginLeft - authoredMarginLeft, marginRight - authoredMarginRight));
        double marginTop = ReadTwipsValue(effectiveSettings.MarginTopValue, document.MarginTopPoints);
        double marginBottom = ReadTwipsValue(effectiveSettings.MarginBottomValue, document.MarginBottomPoints);

        return new DocxPageGeometry(
            width,
            height,
            marginLeft,
            marginRight,
            markupMarginReservePoints,
            marginTop,
            marginBottom,
            effectiveSettings,
            section.SectionProperties,
            CreateColumnFrames(
                width,
                marginLeft,
                marginRight,
                section.SectionProperties));
    }

    private static double ResolveReservedMarkupRightMargin(double pageWidth, double marginLeft, double marginRight)
    {
        double bodyWidth = Math.Max(1d, pageWidth - marginLeft - marginRight);
        if (bodyWidth <= MinimumMarkupBodyWidthPoints)
        {
            return marginRight;
        }

        double maxRightMargin = Math.Max(marginRight, pageWidth - marginLeft - MinimumMarkupBodyWidthPoints);
        return Math.Max(marginRight, Math.Min(PreferredMarkupMarginPoints, maxRightMargin));
    }

    private static double ResolveReservedMarkupLeftMargin(double pageWidth, double marginLeft, double marginRight)
    {
        double bodyWidth = Math.Max(1d, pageWidth - marginLeft - marginRight);
        if (bodyWidth <= MinimumMarkupBodyWidthPoints)
        {
            return marginLeft;
        }

        double maxLeftMargin = Math.Max(marginLeft, pageWidth - marginRight - MinimumMarkupBodyWidthPoints);
        return Math.Max(marginLeft, Math.Min(PreferredMarkupMarginPoints, maxLeftMargin));
    }

    private static bool ShouldReserveLeftMarkupMargin(DocxDocument document, int pageNumber)
    {
        return IsEvenMirroredPage(document, pageNumber);
    }

    private static bool ShouldApplyGutterToRightMargin(DocxDocument document, int pageNumber)
    {
        return IsEvenMirroredPage(document, pageNumber);
    }

    private static bool IsEvenMirroredPage(DocxDocument document, int pageNumber)
    {
        return document.Settings.MirrorMargins == true && pageNumber % 2 == 0;
    }

    private static IReadOnlyList<DocxLayoutColumnFrame> CreateColumnFrames(
        double pageWidth,
        double marginLeft,
        double marginRight,
        DocxSectionLayoutProperties section)
    {
        double bodyWidth = Math.Max(1d, pageWidth - marginLeft - marginRight);
        int columnCount = Math.Max(1, section.ColumnCount ?? 1);
        if (columnCount == 1)
        {
            return [new DocxLayoutColumnFrame(0, marginLeft, bodyWidth, null)];
        }

        if (string.Equals(section.ColumnEqualWidthValue, "0", StringComparison.OrdinalIgnoreCase))
        {
            if (section.ColumnDefinitions.Count == 0)
            {
                return [];
            }

            double x = marginLeft;
            var frames = new List<DocxLayoutColumnFrame>();
            for (int index = 0; index < section.ColumnDefinitions.Count; index++)
            {
                DocxSectionColumnLayoutProperties column = section.ColumnDefinitions[index];
                double customColumnWidth = Math.Max(1d, column.WidthPoints ?? 0d);
                double? customGutter = index + 1 < section.ColumnDefinitions.Count
                    ? Math.Max(0d, column.SpacePoints ?? 0d)
                    : null;
                frames.Add(new DocxLayoutColumnFrame(index, x, customColumnWidth, customGutter));
                x += customColumnWidth + (customGutter ?? 0d);
            }

            return frames;
        }

        double gutter = Math.Max(0d, section.ColumnSpacePoints ?? 0d);
        double columnWidth = Math.Max(1d, (bodyWidth - gutter * (columnCount - 1)) / columnCount);
        return Enumerable.Range(0, columnCount)
            .Select(index => new DocxLayoutColumnFrame(
                index,
                marginLeft + index * (columnWidth + gutter),
                columnWidth,
                index + 1 < columnCount ? gutter : null))
            .ToArray();
    }

    private static DocxLayoutColumnFrame ResolveActiveColumnFrame(DocxPageGeometry page, int activeColumnIndex)
    {
        if (page.ColumnFrames.Count == 0)
        {
            return new DocxLayoutColumnFrame(0, page.MarginLeft, page.BodyWidth, null);
        }

        int index = Math.Clamp(activeColumnIndex, 0, page.ColumnFrames.Count - 1);
        return page.ColumnFrames[index];
    }

    private static double ReadTwipsValue(string? value, double fallback)
    {
        return long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long twips)
            ? OoxUnits.TwipsToPoints(twips)
            : fallback;
    }

    private static double? ReadOptionalTwipsValue(string? value)
    {
        return long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long twips)
            ? OoxUnits.TwipsToPoints(twips)
            : null;
    }

    private static int? ReadOptionalInt32Value(string? value)
    {
        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int result)
            ? result
            : null;
    }

    private static (double Width, double Height) NormalizePageSize(double width, double height)
    {
        if (Math.Abs(width - 595d) < 0.01d && Math.Abs(height - 842d) < 0.01d)
        {
            return (594.96d, 842.04d);
        }

        return (width, height);
    }

    private static bool ShouldStartNewPageForSectionBreak(DocxSectionBreakElement sectionBreak)
    {
        return sectionBreak.TypeValue is null ||
            sectionBreak.TypeValue is DocxSectionBreakType.NextPage or DocxSectionBreakType.OddPage or DocxSectionBreakType.EvenPage;
    }

    private static bool IsContinuousSectionBreak(DocxSectionBreakElement sectionBreak)
    {
        return sectionBreak.TypeValue == DocxSectionBreakType.Continuous;
    }

    private static bool ShouldInsertParityBlankPage(DocxSectionBreakElement sectionBreak, int nextPageNumber)
    {
        if (sectionBreak.TypeValue == DocxSectionBreakType.OddPage)
        {
            return nextPageNumber % 2 == 0;
        }

        if (sectionBreak.TypeValue == DocxSectionBreakType.EvenPage)
        {
            return nextPageNumber % 2 != 0;
        }

        return false;
    }
}
