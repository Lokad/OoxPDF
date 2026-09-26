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
    private static (IReadOnlyList<DocxLayoutPage> Pages, IReadOnlyDictionary<int, double> HeaderContentBottomByPage, IReadOnlyDictionary<int, double> FooterContentTopByPage) AddStaticContent(
        IReadOnlyList<DocxLayoutPage> pages,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        double paragraphSpacingScale,
        IDocxTextMeasurer? unscaledTextMeasurer,
        CancellationToken cancellationToken)
    {
        if (textMeasurer is not IDocxStaticTextMetricsProvider staticMetrics)
        {
            return (pages, new Dictionary<int, double>(), new Dictionary<int, double>());
        }

        IDocxLineMetricsProvider? unscaledLineMetrics =
            (unscaledTextMeasurer as IDocxLineMetricsProvider) ?? (textMeasurer as IDocxLineMetricsProvider);

        var pagesWithStaticText = new DocxLayoutPage[pages.Count];
        var headerBottomCursors = new double[pages.Count];
        var footerTopEdges = new double[pages.Count];
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
                    unscaledLineMetrics,
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
                    unscaledLineMetrics,
                    cancellationToken);
            // Stories that laid out no visible content (absent headers, empty selections)
            // displace nothing: their untouched start cursor sits inside the header zone,
            // which would otherwise read as overflow whenever the header distance is
            // smaller than the top margin (table-fragment regression 2026-09-08: empty
            // headers on a 100pt page displaced body by 8).
            bool headerHasVisibleContent = HasVisibleStaticContent(headerLayout);

            static bool HasVisibleStaticContent(DocxStaticStoryLayoutResult layout)
            {
                return layout.TextLines.Count != 0 ||
                    layout.InlineImages.Count != 0 ||
                    layout.TableRows.Count != 0 ||
                    layout.InlineTextBoxes.Count != 0;
            }
            headerBottomCursors[pageIndex] = headerHasVisibleContent
                ? headerLayout.EndCursorY - headerLayout.EndPendingAfterSpacing
                : double.PositiveInfinity;
            // Footer content top bounds the body frame from below (w48 multi-line footer
            // plus long body, Word-COM rendered: body breaks before footer content instead
            // of overlapping it). Text-only stories expose their first line top under the
            // body baseline rule; stories without text lines constrain nothing.
            double footerContentTop = double.NegativeInfinity;
            DocxTextLineLayout? firstFooterLine = footerLayout.TextLines.FirstOrDefault();
            if (HasVisibleStaticContent(footerLayout) && firstFooterLine?.SourceParagraph is DocxParagraph firstFooterParagraph)
            {
                double firstFooterSize = GetParagraphFontSize(firstFooterParagraph);
                double firstFooterOffset = DocxLineMetrics.ResolveBodyBaselineOffset(firstFooterSize, firstFooterLine.LineHeight ?? firstFooterSize, hasExplicitLineSpacing: false);
                if (HasNoSpacingElement(firstFooterParagraph.EffectiveProperties) && Math.Abs(firstFooterSize - 11d) < 0.000000001d)
                {
                    firstFooterOffset += UntokenedParagraphBaselineExtraPoints;
                }

                footerContentTop = firstFooterLine.BaselineY + firstFooterOffset;
            }

            footerTopEdges[pageIndex] = footerContentTop;
            pagesWithStaticText[pageIndex] = page with
            {
                StaticTextLines = headerLayout.TextLines.Concat(footerLayout.TextLines).ToArray(),
                StaticInlineImages = headerLayout.InlineImages.Concat(footerLayout.InlineImages).ToArray(),
                StaticTableRows = headerLayout.TableRows.Concat(footerLayout.TableRows).ToArray(),
                StaticInlineTextBoxes = headerLayout.InlineTextBoxes.Concat(footerLayout.InlineTextBoxes).ToArray()
            };
        }

        var headerContentBottomByPage = new Dictionary<int, double>();
        var footerContentTopByPage = new Dictionary<int, double>();
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++)
        {
            headerContentBottomByPage[pageIndex] = headerBottomCursors[pageIndex];
            footerContentTopByPage[pageIndex] = footerTopEdges[pageIndex];
        }

        return (pagesWithStaticText, headerContentBottomByPage, footerContentTopByPage);
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
        IDocxLineMetricsProvider? unscaledLineMetrics,
        CancellationToken cancellationToken)
    {
        var lines = new List<DocxTextLineLayout>();
        var images = new List<DocxInlineImageLayout>();
        var tableRows = new List<DocxTableRowLayout>();
        var boxes = new List<DocxInlineTextBoxLayout>();
        double cursorY = startY;
        double pendingSpacingAfter = 0d;
        DocxParagraph? previousParagraph = null;
        double? firstStaticLineBaselineOffset = null;
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
                firstStaticLineBaselineOffset = null;
                var cellMemo = new DocxTableCellTextLinesMemo();
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
                        Story: DocxStoryId.HeaderOrFooter(isHeader, story.VariantType),
                        pageCount: pageCount,
                        paragraphSpacingScale: paragraphSpacingScale,
                        cellMemo: cellMemo));
                    cursorY -= rowHeight;
                }

                continue;
            }

            if (story.BodyElements[elementIndex] is not DocxParagraphElement paragraphElement)
            {
                pendingSpacingAfter = 0d;
                previousParagraph = null;
                firstStaticLineBaselineOffset = null;
                continue;
            }

            DocxParagraph paragraph = paragraphElement.Paragraph;
            DocxParagraphSpacingProfile spacingProfile = ResolveParagraphSpacingProfile(previousParagraph, paragraph, pendingSpacingAfter, paragraphSpacingScale);
            cursorY -= spacingProfile.AppliedBeforeSpacing;
            pendingSpacingAfter = 0d;
            DocxTextSpan[] spans = CreateStaticTextSpans(paragraph.Runs);
            int sourceLineIndex = 0;
            var placedStaticImages = new List<DocxInlineImage>();
            if (spans.Length != 0)
            {
                DocxWrappedTextLine[] staticLines = WrapStaticTextLines(spans, width, textMeasurer, ResolveInlineImageWrapWidths(paragraph, spans)).ToArray();
                // RV05: ordered inline atoms (static header/footer path). Affined images in
                // text-mixed paragraphs attach to wrapped lines at run position.
                DocxMidLinePlan? staticMidLinePlan = CreateStaticMidLinePlan(paragraph, spans, staticLines, width, isHeader, staticMetrics, unscaledLineMetrics, paragraphSpacingScale, cancellationToken);
                for (int staticLineIndex = 0; staticLineIndex < staticLines.Length; staticLineIndex++)
                {
                    DocxWrappedTextLine line = staticLines[staticLineIndex];
                    cancellationToken.ThrowIfCancellationRequested();
                    if (line.Spans.Count == 0)
                    {
                        continue;
                    }

                    double lineWidth = MeasureStaticTextSpans(line.Spans, textMeasurer) + (staticMidLinePlan?.LineImageWidths[staticLineIndex] ?? 0d);
                    // RV06 header-align probe: Office centers/rights drawable header and
                    // footer text too, letting trailing spaces overflow past the edge.
                    double lineAlignWidth = paragraph.EffectiveProperties.Alignment is DocxTextAlignment.Center or DocxTextAlignment.Right
                        ? MeasureStaticTextSpans(SliceTextSpans(line.Spans, 0, FindDrawableTextLength(line.Spans)), textMeasurer) + (staticMidLinePlan?.LineImageWidths[staticLineIndex] ?? 0d)
                        : lineWidth;
                    double lineX = paragraph.EffectiveProperties.Alignment switch
                    {
                        DocxTextAlignment.Center => x + Math.Max(0d, width - lineAlignWidth) / 2d,
                        DocxTextAlignment.Right => x + Math.Max(0d, width - lineAlignWidth),
                        _ => x
                    };
                    // RV05 calibration (Word 16.0): image top pins to the natural line top.
                    double extraAbove = IsExactLineSpacing(paragraph.EffectiveProperties) ? 0d : (staticMidLinePlan?.ShiftAboveHeights[staticLineIndex] ?? 0d);
                    cursorY -= extraAbove;

                    double ascender = line.Spans.Max(span => staticMetrics.MeasureWindowsAscender(span.StyleRun, span.StyleRun.EffectiveProperties.FontSize));
                    double descender = line.Spans.Max(span => staticMetrics.MeasureWindowsDescender(span.StyleRun, span.StyleRun.EffectiveProperties.FontSize));
                    DocxEffectiveParagraphProperties staticEffective = paragraph.EffectiveProperties;
                    double staticFontSize = GetParagraphFontSize(paragraph);
                    double staticLineHeight = ascender + descender;
                    double staticBaselineY = isHeader ? cursorY - ascender : cursorY + descender;
                    double? staticSingleLineHeight = null;
                    if (unscaledLineMetrics is not null &&
                        staticEffective.LineSpacingPoints is null)
                    {
                        // Office A/B (w37-staticfree/w38-footer-free/w43-solo Final-mode
                        // probes at s=1 plus w36 WC balloon re-baselines, the w39 bare-11
                        // header probe, the w44 explicit-factor probe, and locked
                        // header2/footer2, all Word-COM rendered): static lineHeight is
                        // single-height times the spacing factor whenever no exact or
                        // atLeast height is pinned (auto Arial10 pitch 13.32 vs 13.321,
                        // explicit-100 Arial12 pitch 13.80 vs 13.80, explicit-115 pitch
                        // 15.86 vs 15.871, all 16 within 0.08), not windows extents; Word
                        // scales static advances uniformly with the lane scale
                        // (after-steps agree to 0.03), not the 0.7936 line-metric
                        // compromise, so the raw single height advances by the spacing
                        // scale here. Origin follows the body baseline rule including the
                        // bare-11 bump (w39 header inset 10.44 vs body-bare 10.46, locked
                        // explicit-100 header2 inset 9.36 vs 9.40); exact/atLeast statics
                        // keep legacy behavior (unprobed).
                        double rawSingleLineHeight = unscaledLineMetrics.MeasureSingleLineHeight(paragraph.Runs.FirstOrDefault(), staticFontSize);
                        double staticAutoFactor = ResolveAutoLineSpacingFactor(paragraph, out _);
                        double staticDesignLineHeight = rawSingleLineHeight * staticAutoFactor;
                        staticSingleLineHeight = rawSingleLineHeight;
                        staticLineHeight = staticDesignLineHeight * paragraphSpacingScale;
                        double staticBaselineOffset = DocxLineMetrics.ResolveBodyBaselineOffset(staticFontSize, staticDesignLineHeight, hasExplicitLineSpacing: false);
                        if (HasNoSpacingElement(staticEffective) && Math.Abs(staticFontSize - 11d) < 0.000000001d)
                        {
                            staticBaselineOffset += UntokenedParagraphBaselineExtraPoints;
                        }

                        // RV06 pagination probe (edge-mixedstory-wc, Word 16.0): under a
                        // word-compatible print scale, Office scales whole pitches
                        // including the font-size offset transition. Each line
                        // corrects against the story-block-first offset; same-size
                        // lines self-zero and scale 1.0 keeps legacy exactly.
                        double rawStaticBaselineOffset = staticBaselineOffset;
                        firstStaticLineBaselineOffset ??= rawStaticBaselineOffset;
                        if (firstStaticLineBaselineOffset is { } firstStaticOffset)
                        {
                            staticBaselineOffset -= (staticBaselineOffset - firstStaticOffset) * (1d - paragraphSpacingScale);
                        }

                        staticBaselineY = cursorY - staticBaselineOffset;
                    }
                    IReadOnlyList<DocxTextSegmentLayout> segments = CreateStaticTextSegments(line.Spans, lineX);
                    IReadOnlyList<DocxTextSegmentLayout> emissionSegments = segments;
                    if (staticMidLinePlan is not null && staticMidLinePlan.ImagesByLine[staticLineIndex].Count != 0)
                    {
                        var imageShifts = new List<(double BoundaryX, double Shift)>();
                        foreach (DocxMidLineImage placed in staticMidLinePlan.ImagesByLine[staticLineIndex])
                        {
                            double shiftBeforeWidth = MeasureStaticTextSpans(SliceTextSpans(line.Spans, 0, placed.LineCharOffset), textMeasurer);
                            imageShifts.Add((shiftBeforeWidth, placed.Width));
                        }

                        emissionSegments = ShiftSegmentsPastMidLineImages(segments, line.Spans, paragraph, sourceLineIndex == 0, lineX, imageShifts);
                    }
                    lines.Add(new DocxTextLineLayout(
                        line.Text,
                        line.Spans[0].StyleRun,
                        line.Spans.Max(span => span.StyleRun.EffectiveProperties.FontSize),
                        lineX,
                        staticBaselineY,
                        lineWidth,
                        emissionSegments,
                        LineHeight: staticLineHeight,
                        AppliedBeforeSpacing: sourceLineIndex == 0 ? spacingProfile.AppliedBeforeSpacing : 0d,
                        IsFirstParagraphLine: sourceLineIndex == 0,
                        SourceLineIndex: sourceLineIndex,
                        PendingAfterSpacing: sourceLineIndex == 0 ? spacingProfile.PendingAfterSpacing : null,
                        ParagraphBeforeSpacing: sourceLineIndex == 0 ? spacingProfile.ParagraphBeforeSpacing : null,
                        ParagraphAfterSpacing: sourceLineIndex == 0 ? spacingProfile.ParagraphAfterSpacing : null,
                        ContextualSpacingSuppressed: sourceLineIndex == 0 ? spacingProfile.ContextualSpacingSuppressed : null,
                        SourceParagraph: paragraph,
                        SourceParagraphIndex: paragraphIndex,
                        Story: DocxStoryId.HeaderOrFooter(isHeader, story.VariantType),
                        LineHeightSource: DocxLineHeightSource.StaticWindowsExtents, SourceBlockIndex: null, EndsWithIntraTokenBreak: false, SingleLineHeight: staticSingleLineHeight, ListLabelSingleLineHeight: null, BodyWindowsLineHeight: null, ListLabelWindowsLineHeight: null, EffectiveLineSpacingFactor: null, LineSpacingFactorFloorApplied: null, EmitsTerminalParagraphMark: false));
                    // RV06 header-align probe: headers and footers keep one row-end space
                    // beyond authored trailing on the last wrapped line too.
                    if (paragraph.Images.Count == 0 &&
                        paragraph.InlineTextBoxes.Count == 0 &&
                        staticLineIndex == staticLines.Length - 1 &&
                        line.Text.EndsWith(' ') &&
                        spans.Any(static span => span.Text.Any(static character => !char.IsWhiteSpace(character))))
                    {
                        for (int headerLineIndex = lines.Count - 1; headerLineIndex >= 0; headerLineIndex--)
                        {
                            if (lines[headerLineIndex] is not DocxTextLineLayout staticLastLine ||
                                !ReferenceEquals(staticLastLine.SourceParagraph, paragraph) ||
                                staticLastLine.Text.Length == 0)
                            {
                                continue;
                            }

                            DocxTextRun staticStyleRun = line.Spans[0].StyleRun;
                            double staticSpaceFontSize = line.Spans.Max(span => span.StyleRun.EffectiveProperties.FontSize);
                            double staticSpaceWidth = textMeasurer.MeasureText(staticStyleRun, " ", staticSpaceFontSize);
                            lines[headerLineIndex] = staticLastLine with
                            {
                                Text = staticLastLine.Text + " ",
                                Width = staticLastLine.Width + staticSpaceWidth,
                                Segments =
                                [
                                    .. staticLastLine.Segments,
                                    new DocxTextSegmentLayout(
                                        " ",
                                        staticStyleRun,
                                        staticLastLine.X + staticLastLine.Width,
                                        staticSpaceWidth,
                                        staticSpaceFontSize,
                                        0d,
                                        0d,
                                        DocxTextStateCharacterSpacingSource.None,
                                        true,
                                        -1,
                                        0,
                                        DocxTextSegmentRole.BreakSpill),
                                ],
                            };
                            break;
                        }
                    }
                    if (staticMidLinePlan is not null)
                    {
                        foreach (DocxMidLineImage placed in staticMidLinePlan.ImagesByLine[staticLineIndex])
                        {
                            double beforeWidth = MeasureStaticTextSpans(SliceTextSpans(line.Spans, 0, placed.LineCharOffset), textMeasurer);
                            images.Add(new DocxInlineImageLayout(
                                placed.Image,
                                lineX + beforeWidth,
                                staticBaselineY - placed.Height,
                                placed.Width,
                                placed.Height,
                                pageNumber,
                                SourceBlockIndex: null,
                                SourceParagraphIndex: paragraphIndex,
                                Story: DocxStoryId.HeaderOrFooter(isHeader, story.VariantType)));
                            placedStaticImages.Add(placed.Image);
                        }
                    }
                    sourceLineIndex++;
                    cursorY -= staticLineHeight;
                }
            }

            foreach (DocxInlineImage image in paragraph.Images)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (placedStaticImages.Any(placed => ReferenceEquals(placed, image)))
                {
                    continue;
                }

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
                    Story: DocxStoryId.HeaderOrFooter(isHeader, story.VariantType)));
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

                boxes.Add(textBoxLayout with { Story = DocxStoryId.HeaderOrFooter(isHeader, story.VariantType) });
                cursorY -= textBoxLayout.BoxHeight + InlineImageParagraphGapPoints;
            }

            pendingSpacingAfter = spacingProfile.ParagraphAfterSpacing;
            previousParagraph = paragraph;
            paragraphIndex++;
        }

        if (!isHeader &&
            images.Count == 0 &&
            boxes.Count == 0 &&
            tableRows.Count == 0 &&
            lines.Any(line => line.SingleLineHeight.HasValue))
        {
            // Office A/B (w38 default multi-line footer at distance 36, w39 tokened
            // multi-line footer at 72, w40 bare multi-line footer at 54, w43-solo
            // single-line bare footer, all Word-COM rendered): a text-only footer
            // block bottom-anchors its trailing cursor (content plus trailing after
            // spacing) at the footer distance (residuals minus 0.11, minus 0.15,
            // minus 0.11, minus 0.02). The shift preserves every internal pitch (all
            // proven above) and only translates the block; stories with images,
            // textboxes, tables, or exact/atLeast lines keep legacy placement
            // (unprobed there), as do explicit-rule single lines by construction of
            // the gate below combined with the unchanged legacy path above.
            double footerShift = startY - (cursorY - pendingSpacingAfter);
            for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++)
            {
                lines[lineIndex] = lines[lineIndex] with { BaselineY = lines[lineIndex].BaselineY + footerShift };
            }
        }

        return new DocxStaticStoryLayoutResult(lines.ToArray(), images.ToArray(), tableRows.ToArray(), boxes.ToArray(), cursorY, pendingSpacingAfter);

        // RV05: ordered inline atoms (static header/footer path). Growth uses the
        // largest per-line advance and baseline displacement, matching the scalar
        // convention of the body path; positioning stays per-line exact.
        DocxMidLinePlan? CreateStaticMidLinePlan(DocxParagraph planParagraph, DocxTextSpan[] planSpans, DocxWrappedTextLine[] planLines, double planWidth, bool planIsHeader, IDocxStaticTextMetricsProvider planMetrics, IDocxLineMetricsProvider? planUnscaledMetrics, double planScale, CancellationToken planCancellationToken)
        {
            bool hasAffined = false;
            foreach (DocxInlineImage planImage in planParagraph.Images)
            {
                planCancellationToken.ThrowIfCancellationRequested();
                if (planImage.SourceRunIndex >= 0)
                {
                    hasAffined = true;
                    break;
                }
            }

            if (!hasAffined || planLines.Length == 0)
            {
                return null;
            }

            double planFontSize = GetParagraphFontSize(planParagraph);
            bool useSingleHeight = planUnscaledMetrics is not null && planParagraph.EffectiveProperties.LineSpacingPoints is null;
            double maxHeight = 0d;
            double maxDisplacement = 0d;
            if (planUnscaledMetrics is not null && useSingleHeight)
            {
                double rawSingle = planUnscaledMetrics.MeasureSingleLineHeight(planParagraph.Runs.FirstOrDefault(), planFontSize);
                double autoFactor = ResolveAutoLineSpacingFactor(planParagraph, out _);
                maxHeight = rawSingle * autoFactor * planScale;
                maxDisplacement = DocxLineMetrics.ResolveBodyBaselineOffset(planFontSize, rawSingle * autoFactor, hasExplicitLineSpacing: false);
                if (HasNoSpacingElement(planParagraph.EffectiveProperties) && Math.Abs(planFontSize - 11d) < 0.000000001d)
                {
                    maxDisplacement += UntokenedParagraphBaselineExtraPoints;
                }
            }

            foreach (DocxWrappedTextLine planLine in planLines)
            {
                planCancellationToken.ThrowIfCancellationRequested();
                if (planLine.Spans.Count == 0)
                {
                    continue;
                }

                double ascender = planLine.Spans.Max(span => planMetrics.MeasureWindowsAscender(span.StyleRun, span.StyleRun.EffectiveProperties.FontSize));
                double descender = planLine.Spans.Max(span => planMetrics.MeasureWindowsDescender(span.StyleRun, span.StyleRun.EffectiveProperties.FontSize));
                if (!useSingleHeight)
                {
                    maxHeight = Math.Max(maxHeight, ascender + descender);
                    maxDisplacement = Math.Max(maxDisplacement, planIsHeader ? ascender : -descender);
                }
            }

            return CreateMidLinePlan(planParagraph, planSpans, planLines, planWidth, planWidth, maxDisplacement, maxHeight);
        }

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
                .Select(item => CreateTextSpan(ResolveStaticFieldPlaceholders(item.run, pageNumber, pageCount), item.run, item.index))
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
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<(int CharOffset, double Width)>? inlineImageWidths = null)
    {
        string text = string.Concat(spans.Select(span => span.Text));
        int segmentStart = 0;
        while (segmentStart <= text.Length)
        {
            int breakIndex = text.IndexOf('\n', segmentStart);
            int segmentLength = breakIndex < 0 ? text.Length - segmentStart : breakIndex - segmentStart;
            bool yielded = false;
            foreach (DocxWrappedTextLine line in WrapStaticWords(text, spans, segmentStart, segmentLength, maxWidth, textMeasurer, inlineImageWidths))
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
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<(int CharOffset, double Width)>? inlineImageWidths = null)
    {
        IReadOnlyList<TextToken> tokens = TokenizeSpaces(text, segmentStart, segmentLength);
        if (tokens.Count == 0)
        {
            yield break;
        }

        // RV13: index span starts once so slice lookups seek instead of
        // restarting from the first span on every measure and line build.
        int[] spanStarts = new int[spans.Count + 1];
        for (int spanIndex = 0; spanIndex < spans.Count; spanIndex++)
        {
            spanStarts[spanIndex + 1] = spanStarts[spanIndex] + spans[spanIndex].Text.Length;
        }

        // RV13: memoize static slice widths by text coordinates; repeated
        // measures of the same slice shape once.
        var measureMemo = new Dictionary<(int Start, int Length), double>();
        var imageWrapIndex = BuildInlineImageWrapIndex(inlineImageWidths);
        int[] imageOffsets = imageWrapIndex.Offsets;
        double[] imagePrefixSums = imageWrapIndex.PrefixSums;
        int lineStart = tokens[0].Start;
        int lineLength = 0;
        // RV13: maintain whitespace presence incrementally instead of
        // rescanning the whole line per token. Added extents chain
        // contiguously from lineStart, so the running OR equals a fresh scan.
        bool lineHasNonWhitespace = false;
        for (int tokenIndex = 0; tokenIndex < tokens.Count; tokenIndex++)
        {
            TextToken token = tokens[tokenIndex];
            int candidateLength = token.Start + token.Length - lineStart;
            if (lineLength > 0 &&
                lineHasNonWhitespace &&
                !token.IsBreakableWhitespace &&
                MeasureStaticSlice(measureMemo, spans, lineStart, candidateLength, textMeasurer, spanStarts) + ImageWidthInRange(imageWrapIndex, lineStart, lineStart + candidateLength) > maxWidth)
            {
                int imageBreakOffset = -1;
                for (int imageBreakIndex = 0; imageBreakIndex < imageOffsets.Length; imageBreakIndex++)
                {
                    int imageOffset = imageOffsets[imageBreakIndex];
                    if (imageOffset <= lineStart)
                    {
                        continue;
                    }

                    if (imageOffset >= lineStart + candidateLength)
                    {
                        break;
                    }

                    double overflowingImageWidth = imagePrefixSums[imageBreakIndex + 1] - imagePrefixSums[imageBreakIndex];
                    double beforeImageWidth = MeasureStaticSlice(measureMemo, spans, lineStart, imageOffset - lineStart, textMeasurer, spanStarts) + ImageWidthInRange(imageWrapIndex, lineStart, imageOffset);
                    if (beforeImageWidth <= maxWidth && beforeImageWidth + overflowingImageWidth > maxWidth && overflowingImageWidth <= maxWidth)
                    {
                        imageBreakOffset = imageOffset;
                        break;
                    }
                }

                if (imageBreakOffset > lineStart)
                {
                    yield return CreateWrappedTextLine(text, spans, lineStart, imageBreakOffset - lineStart, false, spanStarts);
                    lineStart = imageBreakOffset;
                    lineLength = 0;
                    lineHasNonWhitespace = false;
                    tokenIndex--;
                    continue;
                }

                yield return CreateWrappedTextLine(text, spans, lineStart, lineLength, false, spanStarts);
                lineStart = token.Start;
                lineLength = token.Length;
                lineHasNonWhitespace = HasNonWhitespace(text, lineStart, lineLength);
            }
            else
            {
                int addedStart = Math.Max(token.Start, lineStart);
                lineHasNonWhitespace = lineHasNonWhitespace || HasNonWhitespace(text, addedStart, token.Start + token.Length - addedStart);
                lineLength = candidateLength;
            }
        }

        if (lineLength > 0)
        {
            yield return CreateWrappedTextLine(text, spans, lineStart, lineLength, false, spanStarts);
        }
    }

    // RV13: memoized static slice measurement by text coordinates for one
    // segment; identical slices measure once.
    private static double MeasureStaticSlice(
        Dictionary<(int Start, int Length), double> measureMemo,
        IReadOnlyList<DocxTextSpan> spans,
        int start,
        int length,
        IDocxTextMeasurer textMeasurer,
        int[] spanStarts)
    {
        if (measureMemo.TryGetValue((start, length), out double cached))
        {
            return cached;
        }

        double width = MeasureStaticTextSpansForWrapping(SliceTextSpans(spans, start, length, spanStarts), textMeasurer);
        measureMemo[(start, length)] = width;
        return width;
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

    private static string ResolveStaticFieldPlaceholders(DocxTextRun run, int pageNumber, int pageCount)
    {
        string text = run.Text;
        if (run.FieldKind == DocxFieldKind.Page)
        {
            text = text.Replace("{PAGE}", pageNumber.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal);
        }

        if (run.FieldKind == DocxFieldKind.NumPages)
        {
            text = text.Replace("{NUMPAGES}", pageCount.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal);
        }

        return text;
    }
}
