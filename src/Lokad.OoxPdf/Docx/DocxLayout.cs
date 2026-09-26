using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed record DocxLayout(
    IReadOnlyList<DocxLayoutPage> Pages,
    IReadOnlyList<DocxFloatingDrawingLayout> FloatingDrawings,
    IReadOnlyList<DocxFloatingDrawingLayout> StaticFloatingDrawings,
    IReadOnlyList<DocxRelatedStoryLayout> RelatedStories,
    IReadOnlyDictionary<int, double> HeaderContentBottomByPage,
    IReadOnlyDictionary<int, double> FooterContentTopByPage);

internal sealed record DocxRelatedStoryLayout(
    DocxRelatedStory Story,
    int StoryIndex,
    IReadOnlyList<DocxTextLineLayout> TextLines,
    IReadOnlyList<DocxInlineImageLayout> InlineImages,
    IReadOnlyList<DocxFloatingDrawingLayout> FloatingDrawings,
    IReadOnlyList<DocxTableRowLayout> TableRows,
    double ContentHeight);

internal sealed record DocxPlacedRelatedStoryLayout(
    DocxRelatedStoryLayout StoryLayout,
    int StoryIndex,
    int SourceBlockIndex,
    double X,
    double TopY,
    double Width,
    double Height,
    double ContentTopOffset,
    double ContentHeight,
    double? SeparatorY,
    double SeparatorWidth,
    double SeparatorThickness,
    IReadOnlyList<DocxTextLineLayout> TextLines,
    IReadOnlyList<DocxInlineImageLayout> InlineImages,
    IReadOnlyList<DocxFloatingDrawingLayout> FloatingDrawings,
    IReadOnlyList<DocxTableRowLayout> TableRows);

internal sealed record DocxFloatingDrawingLayout(
    DocxFloatingDrawing Drawing,
    int? PageStartIndex,
    int? PageEndIndex,
    int? AnchorPageIndex,
    int? AnchorColumnIndex,
    double? AnchorBlockVerticalTop,
    double? AnchorBlockVerticalBottom,
    double? ExtentWidthPoints,
    double? ExtentHeightPoints,
    double? HorizontalOffsetPoints,
    double? VerticalOffsetPoints,
    double? DistanceTopPoints,
    double? DistanceBottomPoints,
    double? DistanceLeftPoints,
    double? DistanceRightPoints,
    double? HorizontalReferenceX,
    double? HorizontalReferenceWidth,
    double? VerticalReferenceTop,
    double? VerticalReferenceBottom,
    double? PlacedX,
    double? PlacedTop,
    DocxAnchorPlacementSource? HorizontalPlacementSource,
    DocxAnchorPlacementSource? VerticalPlacementSource,
    double? WrapExclusionX,
    double? WrapExclusionTop,
    double? WrapExclusionWidth,
    double? WrapExclusionHeight,
    DocxStoryId? Story,
    DocxRelatedStoryLayout? TextBoxLayout);

internal sealed record DocxWrapExclusionFrame(
    double X,
    double Top,
    double Width,
    double Height);

internal enum DocxAnchorPlacementSource
{
    Align,
    Offset,
    Unsupported,
    MissingReferenceOrExtent
}

internal sealed record DocxAnchorPlacement(double? Position, DocxAnchorPlacementSource Source);

internal sealed record DocxLineHeightProfile(
    double LineHeight,
    double? SingleLineHeight,
    double? ListLabelSingleLineHeight,
    double? BodyWindowsLineHeight,
    double? ListLabelWindowsLineHeight,
    double? EffectiveLineSpacingFactor,
    bool LineSpacingFactorFloorApplied,
    DocxLineHeightSource Source);

internal enum DocxLineHeightSource
{
    ExactLineSpacing,
    AtLeastLineSpacing,
    BodySingleLineAuto,
    StaticWindowsExtents,
    TerminalParagraphMark
}

internal sealed record DocxParagraphLineShape(
    string Text,
    double X,
    double Width,
    IReadOnlyList<DocxTextSegmentLayout> Segments);

internal sealed record DocxParagraphSpacingProfile(
    double PendingAfterSpacing,
    double ParagraphBeforeSpacing,
    double ParagraphAfterSpacing,
    double AppliedBeforeSpacing,
    bool ContextualSpacingSuppressed);

internal sealed record DocxSectionLayoutProperties(
    string? BreakTypeValue,
    string? ColumnCountValue,
    string? ColumnEqualWidthValue,
    string? ColumnSpaceValue,
    int? ColumnCount,
    double? ColumnSpacePoints,
    IReadOnlyList<DocxSectionColumnLayoutProperties> ColumnDefinitions);

internal sealed record DocxSectionColumnLayoutProperties(
    string? WidthValue,
    string? SpaceValue,
    double? WidthPoints,
    double? SpacePoints);

internal sealed record DocxLayoutColumnFrame(
    int Index,
    double X,
    double Width,
    double? GutterAfterPoints);

internal static class DocxLayoutColumnOwnership
{
    public static int? ResolveColumnIndex(IReadOnlyList<DocxLayoutColumnFrame> frames, double x, double width)
    {
        if (frames.Count == 0)
        {
            return null;
        }

        double center = x + Math.Max(0d, width) / 2d;
        DocxLayoutColumnFrame? containingFrame = frames.FirstOrDefault(frame => center >= frame.X && center <= frame.X + frame.Width);
        if (containingFrame is not null)
        {
            return containingFrame.Index;
        }

        return frames
            .OrderBy(frame => Math.Abs(center - (frame.X + frame.Width / 2d)))
            .First()
            .Index;
    }
}

internal sealed partial class DocxLayoutEngine
{
    private const double WordDefaultTabStopPoints = 36d;
    private const double InlineImageParagraphGapPoints = 6d;
    private const double WordListMinimumAutoLineSpacingFactor = 1.16d;
    private const double FootnoteSeparatorGapPoints = 3d;
    // Word draws the footnote separator rule 144 points wide regardless of page width or separator story content
    // (footnote Office probes 2026-09-07: identical rects with pBdr, borderless, thick-pBdr, absent-story, and narrow-page variants).
    private const double FootnoteSeparatorWidthPoints = 144d;
    // Word rule thickness measures 0.72 to 0.84 across probes (likely export-grid snapping around 0.75); midpoint taken.
    private const double FootnoteSeparatorThicknessPoints = 0.75d;
    // Word rule bottom sits this far above the separator space baseline (2.04 and 2.16 across probes).
    private const double FootnoteSeparatorRuleBottomOffsetPoints = 2.1d;
    // RV06 endnote probes (3.76 and 3.72 across two probes): the endnote separator
    // rule sits ~3.74 above the mark baseline versus 2.1 for footnotes.
    private const double EndnoteSeparatorRuleBottomOffsetPoints = 3.74d;
    // Word separator space baselines sit within 0.15 of the separator block bottom on both probes; midpoint taken.
    private const double FootnoteSeparatorBaselineOffsetPoints = 0.15d;
    private const double UnpagedRelatedStoryCanvasHeightPoints = 100000d;
    private const double PreferredMarkupMarginPoints = 207d;
    private const double MinimumMarkupBodyWidthPoints = 216d;
    private const double WordCompatibleAllMarkupParagraphSpacingScale = 0.842391d;
    private const double TableCellNoWrapLineWidthPoints = 1_000_000d;
    private const double UntokenedParagraphBaselineExtraPoints = 0.12d;

    // RV06: break-adjacent spacing rows spill from the paragraph's last non-empty
    // body line. Comment-range paragraphs may carry trailing marker lines, so scan
    // back for the last non-empty line owned by the paragraph itself.
    private static int FindBreakSpillLineIndex(IReadOnlyList<DocxLayoutItem> currentItems, DocxParagraph paragraph)
    {
        for (int itemIndex = currentItems.Count - 1; itemIndex >= 0; itemIndex--)
        {
            if (currentItems[itemIndex] is DocxTextLineLayout candidate &&
                ReferenceEquals(candidate.SourceParagraph, paragraph) &&
                candidate.Text.Length > 0)
            {
                return itemIndex;
            }
        }

        return -1;
    }

    private static bool HasNoSpacingElement(DocxEffectiveParagraphProperties effective)
    {
        return !DocxParagraphSpacing.HasBeforeSpacingSide(effective.Spacing) &&
            !DocxParagraphSpacing.HasAfterSpacingSide(effective.Spacing) &&
            effective.Spacing.LineValue is null &&
            effective.Spacing.LineRuleValue is null &&
            effective.LineSpacingPoints is null;
    }
    private readonly bool reserveMarkupMargin;
    private readonly double paragraphSpacingScale;
    private readonly bool retuneReserveToPrintScale;
    private readonly double reservePrintScale;

    private sealed record DocxPageGeometry(
        double Width,
        double Height,
        double MarginLeft,
        double MarginRight,
        double MarkupMarginReservePoints,
        double MarginTop,
        double MarginBottom,
        DocxPageSettings PageSettings,
        DocxSectionLayoutProperties SectionProperties,
        IReadOnlyList<DocxLayoutColumnFrame> ColumnFrames)
    {
        public double BodyWidth => Math.Max(1d, Width - MarginLeft - MarginRight);
    }

    private sealed record DocxEffectiveSectionSettings(
        DocxPageSettings PageSettings,
        DocxSectionLayoutProperties SectionProperties);

    private sealed record DocxTableLayoutFrame(
        DocxTableLayoutContext Context,
        IReadOnlyList<double> EffectiveColumns,
        double Scale,
        IReadOnlyList<double> RowHeights,
        double PageContentHeight,
        double TableX);

    private sealed record DocxResolvedTableGrid(
        double TableX,
        double TableAvailableWidth,
        double TargetTableWidth,
        IReadOnlyList<double> EffectiveColumns,
        double Scale,
        IReadOnlyList<double> ResolvedColumnWidths);

    public DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode markupGeometryMode, double wordCompatiblePrintScale = WordCompatibleAllMarkupParagraphSpacingScale)
    {
        reserveMarkupMargin = markupGeometryMode is OoxPdfDocxMarkupGeometryMode.ReserveMarkupMargin or OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
        paragraphSpacingScale = markupGeometryMode == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
            ? wordCompatiblePrintScale
            : 1d;
        retuneReserveToPrintScale = markupGeometryMode == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
        reservePrintScale = wordCompatiblePrintScale;
    }

    public DocxLayout Create(DocxDocument document, PdfEmbeddedFont? embedded, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        IDocxTextMeasurer? textMeasurer = embedded is null ? null : new DocxEmbeddedTextMeasurer(embedded);
        return Create(document, textMeasurer, cancellationToken);
    }

    internal DocxLayout Create(DocxDocument document, IDocxTextMeasurer? textMeasurer, CancellationToken cancellationToken, IDocxTextMeasurer? unscaledTextMeasurer = null, IReadOnlyDictionary<int, double>? headerOverflowDisplacementByPage = null, IReadOnlyDictionary<int, double>? footerFrameBottomDisplacementByPage = null)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var pages = new List<DocxLayoutPage>();
        var currentItems = new List<DocxLayoutItem>();
        IReadOnlyDictionary<int, DocxEffectiveSectionSettings> sectionSettingsByElementIndex = BuildEffectiveSectionSettings(document, out DocxEffectiveSectionSettings finalSectionSettings);
        DocxEffectiveSectionSettings activeSectionSettings = FindSectionSettingsAtOrAfter(document.BodyElements, 0, sectionSettingsByElementIndex) ?? finalSectionSettings;
        DocxPageGeometry page = ResolveSectionGeometry(document, activeSectionSettings, reserveMarkupMargin, retuneReserveToPrintScale, reservePrintScale, pageNumber: 1);
        int activeColumnIndex = 0;
        double x = ResolveActiveColumnFrame(page, activeColumnIndex).X;
        double width = ResolveActiveColumnFrame(page, activeColumnIndex).Width;
        double cursorY = ResolvePageStartCursor();
        double pendingSpacingAfter = 0d;
        DocxParagraph? previousParagraph = null;
        bool activeColumnHasContent = false;
        int tableIndex = 0;
        double defaultTabStopPoints = document.Settings.DefaultTabStopPoints ?? WordDefaultTabStopPoints;
        var relatedStoryLayoutsByBodyWidth = new Dictionary<double, IReadOnlyList<DocxRelatedStoryLayout>>();
        var footnoteReserveHeightByBodyWidth = new Dictionary<double, IReadOnlyDictionary<int, double>>();

        IReadOnlyList<DocxRelatedStoryLayout> GetRelatedStoryLayouts(double bodyWidth)
        {
            double key = Math.Round(Math.Max(1d, bodyWidth), 3);
            if (!relatedStoryLayoutsByBodyWidth.TryGetValue(key, out IReadOnlyList<DocxRelatedStoryLayout>? layouts))
            {
                layouts = CreateRelatedStoryLayouts(document.RelatedStories, key, textMeasurer, defaultTabStopPoints, paragraphSpacingScale, cancellationToken, unscaledTextMeasurer);
                relatedStoryLayoutsByBodyWidth[key] = layouts;
            }

            return layouts;
        }

        IReadOnlyDictionary<int, double> GetFootnoteReserveHeightBySourceBlock(double bodyWidth)
        {
            double key = Math.Round(Math.Max(1d, bodyWidth), 3);
            if (!footnoteReserveHeightByBodyWidth.TryGetValue(key, out IReadOnlyDictionary<int, double>? reserveHeights))
            {
                reserveHeights = CreateFootnoteReserveHeightBySourceBlock(document, GetRelatedStoryLayouts(key), cancellationToken);
                footnoteReserveHeightByBodyWidth[key] = reserveHeights;
            }

            return reserveHeights;
        }

        IReadOnlyList<DocxRelatedStoryLayout> relatedStoryLayouts = GetRelatedStoryLayouts(page.BodyWidth);
        double currentPageFootnoteReserveHeight = 0d;

        void ApplyActiveColumnFrame()
        {
            DocxLayoutColumnFrame frame = ResolveActiveColumnFrame(page, activeColumnIndex);
            x = frame.X;
            width = frame.Width;
        }

        void FinishPage()
        {
            pages.Add(new DocxLayoutPage(
                page.Width,
                page.Height,
                page.MarginLeft,
                page.MarginRight,
                page.MarkupMarginReservePoints,
                page.MarginTop,
                page.MarginBottom,
                page.PageSettings,
                page.SectionProperties,
                page.ColumnFrames,
                [],
                [],
                [],
                [],
                currentItems.ToArray()));
            // RV12: guard the page budget while paginating without consuming it, so
            // a tiny budget trips before the full layout is retained; emission still
            // charges once per final page and repagination never double-charges.
            OoxConversionBudget.Current?.ThrowIfLayoutPagesExceedBudget(pages.Count);
            currentItems = [];
            activeColumnIndex = 0;
            page = ResolveSectionGeometry(document, activeSectionSettings, reserveMarkupMargin, retuneReserveToPrintScale, reservePrintScale, pages.Count + 1);
            ApplyActiveColumnFrame();
            cursorY = ResolvePageStartCursor();
            pendingSpacingAfter = 0d;
            previousParagraph = null;
            activeColumnHasContent = false;
            currentPageFootnoteReserveHeight = 0d;
        }

        double CurrentFrameBottom()
        {
            // Office A/B (w48 multi-line footer plus long body, Word-COM rendered): body
            // breaks before footer content instead of overlapping it, so an overflowing
            // footer top raises the body frame just like a footnote reserve does; the
            // higher of the two keep-out zones binds.
            return page.MarginBottom + Math.Max(currentPageFootnoteReserveHeight, ResolvePagedDisplacement(footerFrameBottomDisplacementByPage));
        }

        double ResolvePageStartCursor()
        {
            // Office A/B (w37/w39 multi-paragraph headers plus w47 cross-story spacing,
            // Word-COM rendered): body starts below overflowing header content. The
            // displacement map carries per-page overflow from a previous full layout;
            // without a map (or on fitting pages) this is exactly the legacy page top.
            // Pages beyond a short map reuse its last entry (extra pages arise from the
            // displacement itself under a repeated header).
            return page.Height - page.MarginTop - ResolvePagedDisplacement(headerOverflowDisplacementByPage);
        }

        double ResolvePagedDisplacement(IReadOnlyDictionary<int, double>? displacementByPage)
        {
            // Pages beyond a short map reuse its last entry (extra pages arise from the
            // displacement itself under repeated headers and footers).
            double displacement = 0d;
            if (displacementByPage is not null)
            {
                if (!displacementByPage.TryGetValue(pages.Count, out displacement) &&
                    displacementByPage.Count != 0)
                {
                    displacementByPage.TryGetValue(displacementByPage.Count - 1, out displacement);
                }
            }

            return displacement;
        }

        void EnsureFootnoteReserveForSourceBlock(int sourceBlockIndex)
        {
            IReadOnlyDictionary<int, double> footnoteReserveHeightBySourceBlock = GetFootnoteReserveHeightBySourceBlock(page.BodyWidth);
            if (footnoteReserveHeightBySourceBlock.TryGetValue(sourceBlockIndex, out double reserveHeight))
            {
                currentPageFootnoteReserveHeight = Math.Max(currentPageFootnoteReserveHeight, reserveHeight);
            }
        }

        void AdvanceColumnOrPage()
        {
            if (activeColumnIndex + 1 < page.ColumnFrames.Count)
            {
                activeColumnIndex++;
                ApplyActiveColumnFrame();
                cursorY = ResolvePageStartCursor();
                pendingSpacingAfter = 0d;
                previousParagraph = null;
                activeColumnHasContent = false;
                currentPageFootnoteReserveHeight = 0d;
                return;
            }

            FinishPage();
        }

        void ApplySectionAfterBreak(int elementIndex)
        {
            activeSectionSettings = FindSectionSettingsAtOrAfter(document.BodyElements, elementIndex + 1, sectionSettingsByElementIndex) ?? finalSectionSettings;
            page = ResolveSectionGeometry(document, activeSectionSettings, reserveMarkupMargin, retuneReserveToPrintScale, reservePrintScale, pages.Count + 1);
            activeColumnIndex = 0;
            ApplyActiveColumnFrame();
            activeColumnHasContent = false;
            if (!HasPageContent())
            {
                cursorY = ResolvePageStartCursor();
            }
        }

        bool HasPageContent() => currentItems.Count > 0;
        bool HasCurrentColumnContent() => activeColumnHasContent;

        for (int elementIndex = 0; elementIndex < document.BodyElements.Count; elementIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxBodyElement element = document.BodyElements[elementIndex];
            if (element is DocxPageBreakElement pageBreak)
            {
                if (pageBreak.BreakParagraph is { } breakParagraph)
                {
                    DocxParagraphSpacingProfile breakSpacingProfile = ResolveParagraphSpacingProfile(previousParagraph, breakParagraph, pendingSpacingAfter, paragraphSpacingScale);
                    double breakFontSize = GetParagraphFontSize(breakParagraph);
                    double breakLineHeight = ResolveLineHeight(breakParagraph, breakFontSize, textMeasurer);
                    double paragraphAdvance = breakSpacingProfile.AppliedBeforeSpacing + breakLineHeight;
                    if (cursorY - paragraphAdvance < CurrentFrameBottom() && HasCurrentColumnContent())
                    {
                        AdvanceColumnOrPage();
                    }

                    cursorY -= paragraphAdvance;
                    activeColumnHasContent = true;
                }

                if (HasPageContent() || pageBreak.BreakParagraph is not null)
                {
                    FinishPage();
                }

                pendingSpacingAfter = 0d;
                previousParagraph = null;
                continue;
            }

            if (element is DocxManualBreakElement manualBreak)
            {
                if (manualBreak.Value?.Equals("column", StringComparison.OrdinalIgnoreCase) == true)
                {
                    AdvanceColumnOrPage();
                    continue;
                }

                pendingSpacingAfter = 0d;
                previousParagraph = null;
                continue;
            }

            if (element is DocxSectionBreakElement sectionBreak)
            {
                bool startsNewPage = ShouldStartNewPageForSectionBreak(sectionBreak);
                if (startsNewPage && HasPageContent())
                {
                    FinishPage();
                }

                if (ShouldInsertParityBlankPage(sectionBreak, pages.Count + 1))
                {
                    FinishPage();
                }

                if (startsNewPage || (IsContinuousSectionBreak(sectionBreak) && !HasPageContent()))
                {
                    ApplySectionAfterBreak(elementIndex);
                }

                pendingSpacingAfter = 0d;
                previousParagraph = null;
                continue;
            }

            if (element is DocxTableElement tableElement)
            {
                EnsureFootnoteReserveForSourceBlock(elementIndex);
                cursorY -= pendingSpacingAfter;
                pendingSpacingAfter = 0d;
                previousParagraph = null;
                int itemCountBeforeTable = currentItems.Count;
                int currentTableIndex = tableIndex++;
                bool tableActiveFrameHasContent = false;
                void MarkTableBoundaryContent()
                {
                    tableActiveFrameHasContent = true;
                }

                Action advanceTableBoundary = page.ColumnFrames.Count > 1
                    ? () =>
                    {
                        AdvanceColumnOrPage();
                        tableActiveFrameHasContent = false;
                    }
                    : FinishPage;
                Func<bool> hasTableBoundaryContent = page.ColumnFrames.Count > 1
                    ? () => HasCurrentColumnContent() || tableActiveFrameHasContent
                    : HasPageContent;
                DocxTableLayoutFrame ResolveCurrentTableFrame()
                {
                    return CreateTableLayoutFrame(
                        tableElement.Table,
                        currentTableIndex,
                        elementIndex,
                        x,
                        width,
                        page.Height - page.MarginTop - CurrentFrameBottom(),
                        textMeasurer,
                        defaultTabStopPoints,
                        cancellationToken,
                        pageNumber: pages.Count + 1,
                        pageCount: null,
                        paragraphSpacingScale: paragraphSpacingScale);
                }

                LayoutTable(tableElement.Table, CurrentFrameBottom(), textMeasurer, defaultTabStopPoints, () => pages.Count + 1, ref currentItems, ref cursorY, ResolveCurrentTableFrame, advanceTableBoundary, hasTableBoundaryContent, MarkTableBoundaryContent, cancellationToken, paragraphSpacingScale, new DocxTableCellTextLinesMemo());
                if (currentItems.Count > itemCountBeforeTable)
                {
                    activeColumnHasContent = true;
                }

                continue;
            }

            if (element is DocxImplicitParagraphElement implicitParagraph &&
                implicitParagraph.SourceKind == DocxBreakSourceKind.TerminalTable)
            {
                DocxTextRun markRun = DocxImplicitParagraphElement.CreateParagraphMarkRun();
                double markFontSize = markRun.EffectiveProperties.FontSize;
                double baselineOffset = DocxLineMetrics.ResolveBodyBaselineOffset(markFontSize, markFontSize, hasExplicitLineSpacing: false);
                currentItems.Add(new DocxTextLineLayout(
                    string.Empty,
                    markRun,
                    markFontSize,
                    x,
                    cursorY - baselineOffset,
                    0d,
                    [new DocxTextSegmentLayout(string.Empty, markRun, x, 0d, null, 0d, 0d, DocxTextStateCharacterSpacingSource.None, true, -1, 0, DocxTextSegmentRole.Text)],
                    SourceBlockIndex: elementIndex,
                    SourceParagraphIndex: 0,
                    SourceLineIndex: 0,
                    Story: DocxStoryId.Body(),
                    LineHeight: markFontSize,
                    IsFirstParagraphLine: true,
                    AppliedBeforeSpacing: null, EndsWithIntraTokenBreak: false, SingleLineHeight: null, ListLabelSingleLineHeight: null, BodyWindowsLineHeight: null, ListLabelWindowsLineHeight: null, EffectiveLineSpacingFactor: null, LineSpacingFactorFloorApplied: null, PendingAfterSpacing: null, ParagraphBeforeSpacing: null, ParagraphAfterSpacing: null, ContextualSpacingSuppressed: null, SourceParagraph: null, LineHeightSource: DocxLineHeightSource.TerminalParagraphMark,
                    EmitsTerminalParagraphMark: true));
                activeColumnHasContent = true;
                previousParagraph = null;
                pendingSpacingAfter = 0d;
                continue;
            }

            if (element is not DocxParagraphElement paragraphElement)
            {
                continue;
            }

            DocxParagraph paragraph = paragraphElement.Paragraph;
            DocxEffectiveParagraphProperties effective = paragraph.EffectiveProperties;
            EnsureFootnoteReserveForSourceBlock(elementIndex);
            DocxParagraphSpacingProfile spacingProfile = ResolveParagraphSpacingProfile(previousParagraph, paragraph, pendingSpacingAfter, paragraphSpacingScale);
            cursorY -= spacingProfile.AppliedBeforeSpacing;
            pendingSpacingAfter = 0d;
            double paragraphFontSize = GetParagraphFontSize(paragraph);
            DocxLineHeightProfile lineHeightProfile = ResolveLineHeightProfile(paragraph, paragraphFontSize, textMeasurer);
            double lineHeight = lineHeightProfile.LineHeight;
            if (textMeasurer is not null &&
                HasPageContent() &&
                ShouldKeepParagraphBlockTogether(paragraph) &&
                cursorY - EstimateKeptParagraphBlock(document.BodyElements, elementIndex, width, textMeasurer, defaultTabStopPoints, pages.Count + 1, paragraphSpacingScale).Height <= CurrentFrameBottom())
            {
                AdvanceColumnOrPage();
                EnsureFootnoteReserveForSourceBlock(elementIndex);
            }

            IReadOnlyList<DocxTextSpan> textSpans = textMeasurer is null ? [] : CreateTextSpans(paragraph.Runs, pages.Count + 1, null);
            DocxMidLinePlan? midLinePlan = null;
            if (textMeasurer is not null && textSpans.Count > 0)
            {
                double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, paragraphFontSize, textMeasurer, paragraphSpacingScale);
                double continuationTextStartOffset = GetParagraphTextStartOffset(paragraph, paragraphSpacingScale);
                double labelStartOffset = GetParagraphLabelStartOffset(paragraph, paragraphSpacingScale);
                double paragraphX = x + textStartOffset;
                double paragraphWidth = Math.Max(1d, width - textStartOffset - GetParagraphRightInset(paragraph, paragraphSpacingScale));
                DocxTextRun firstRun = paragraph.Runs[0];
                bool firstLine = true;
                double continuationParagraphWidth = Math.Max(1d, width - continuationTextStartOffset - GetParagraphRightInset(paragraph, paragraphSpacingScale));
                DocxWrappedTextLine[] lines = WrapTextLines(textSpans, paragraphWidth, continuationParagraphWidth, paragraphFontSize, textMeasurer, ScaleTabStopPositions(effective.TabStops, paragraphSpacingScale), defaultTabStopPoints * paragraphSpacingScale, allowOverwideTokenBreaks: ShouldAllowCharacterLevelWordWrap(paragraph), dynamicFieldPageNumber: pages.Count + 1, cancellationToken, inlineImageWidths: ResolveInlineImageWrapWidths(paragraph, textSpans)).ToArray();
                if (ShouldMoveParagraphForWidowControl(paragraph, lines.Length, cursorY, lineHeight, CurrentFrameBottom(), HasCurrentColumnContent()))
                {
                    AdvanceColumnOrPage();
                    EnsureFootnoteReserveForSourceBlock(elementIndex);
                }

                // RV05: ordered inline atoms (body path). Affined images in text-mixed
                // paragraphs attach to wrapped lines at run position; wrapping is untouched.
                midLinePlan = CreateMidLinePlan(paragraph, textSpans, lines, paragraphWidth, continuationParagraphWidth, DocxLineMetrics.ResolveBodyBaselineOffset(paragraphFontSize, lineHeight, IsExactLineSpacing(effective)), lineHeight);
                for (int lineIndex = 0; lineIndex < lines.Length; lineIndex++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    DocxWrappedTextLine line = lines[lineIndex];
                    // RV05 calibration (Word 16.0): the image top pins to the natural line top,
                    // so an auto-spaced line shifts down by image height minus ascent here; the
                    // advance below never grows. Exact spacing stays rigid (no shift).
                    double extraAbove = IsExactLineSpacing(effective) ? 0d : (midLinePlan?.ShiftAboveHeights[lineIndex] ?? 0d);
                    if (firstLine)
                    {
                        cursorY -= ResolveListLabelFirstLineExtraLeading(paragraph, paragraphFontSize, textMeasurer);
                    }

                    if (cursorY - extraAbove - lineHeight < CurrentFrameBottom() && HasCurrentColumnContent())
                    {
                        AdvanceColumnOrPage();
                        EnsureFootnoteReserveForSourceBlock(elementIndex);
                    }

                    cursorY -= extraAbove;

                    double lineWidth = MeasureTextSpansForLayout(line.Spans, paragraphFontSize, textMeasurer, ScaleTabStopPositions(effective.TabStops, paragraphSpacingScale), defaultTabStopPoints * paragraphSpacingScale, pages.Count + 1) + (midLinePlan?.LineImageWidths[lineIndex] ?? 0d);
                    // RV06 align matrix: Office centers/rights the drawable text, letting
                    // authored and added trailing spaces overflow past the edge.
                    double lineAlignWidth = effective.Alignment is DocxTextAlignment.Center or DocxTextAlignment.Right
                        ? MeasureDrawableTextSpansForLayout(line.Spans, paragraphFontSize, textMeasurer, ScaleTabStopPositions(effective.TabStops, paragraphSpacingScale), defaultTabStopPoints * paragraphSpacingScale, pages.Count + 1) + (midLinePlan?.LineImageWidths[lineIndex] ?? 0d)
                        : lineWidth;
                    double lineX = effective.Alignment switch
                    {
                        DocxTextAlignment.Center => paragraphX + Math.Max(0, paragraphWidth - lineAlignWidth) / 2d,
                        DocxTextAlignment.Right => paragraphX + Math.Max(0, paragraphWidth - lineAlignWidth),
                        _ => paragraphX
                    };
                    double baselineOffset = DocxLineMetrics.ResolveBodyBaselineOffset(paragraphFontSize, lineHeight, IsExactLineSpacing(effective));
                    if (HasNoSpacingElement(effective) && Math.Abs(paragraphFontSize - 11d) < 0.000000001d)
                    {
                        // Office A/B (w18/w20/w21/w26/w29/w32 untokened probes, Word-COM
                        // rendered): bare 11pt paragraphs (no w:spacing element, max
                        // nominal size 11 including footnote marks) place baselines 0.12
                        // deeper with identical boxes (doc-start first baselines minus
                        // 0.12 at top 72 and 144; mixed mid-doc edges plus 0.09 and
                        // minus 0.09 with page totals preserved; phantom-before refuted
                        // by the shrunken following gap). Other sizes keep legacy
                        // placement (10/12/14/20pt bare match the tokened model to
                        // 0.06); w:line-only paragraphs keep legacy placement
                        // (explicit-rule first-gap anomaly persists); serif residuals
                        // (Times exact, Georgia minus 0.24) await the font-by-size
                        // inset matrix program.
                        baselineOffset += UntokenedParagraphBaselineExtraPoints;
                    }
                    DocxParagraphLineShape lineShape = CreateParagraphLineShape(
                        paragraph,
                        line,
                        firstRun,
                        firstLine,
                        lineIndex == lines.Length - 1,
                        x + labelStartOffset,
                        lineX,
                        paragraphWidth,
                        paragraphFontSize,
                        textMeasurer,
                        ScaleTabStopPositions(effective.TabStops, paragraphSpacingScale),
                        defaultTabStopPoints * paragraphSpacingScale,
                        pages.Count + 1);
                    IReadOnlyList<DocxTextSegmentLayout> emissionSegments = lineShape.Segments;
                    List<DocxInlineImageLayout>? lineImages = null;
                    List<(double BoundaryX, double Shift)>? imageShifts = null;
                    if (midLinePlan is not null && midLinePlan.ImagesByLine[lineIndex].Count != 0)
                    {
                        double textBaselineY = cursorY - baselineOffset;
                        imageShifts = new List<(double BoundaryX, double Shift)>();
                        lineImages = new List<DocxInlineImageLayout>();
                        foreach (DocxMidLineImage placed in midLinePlan.ImagesByLine[lineIndex])
                        {
                            double beforeWidth = MeasureMidLineBeforeWidth(line.Spans, placed.LineCharOffset, paragraph, firstLine, lineIndex == lines.Length - 1, paragraphWidth, paragraphFontSize, textMeasurer, ScaleTabStopPositions(effective.TabStops, paragraphSpacingScale), defaultTabStopPoints * paragraphSpacingScale, pages.Count + 1);
                            imageShifts.Add((beforeWidth, placed.Width));
                            lineImages.Add(new DocxInlineImageLayout(
                                placed.Image,
                                lineX + beforeWidth,
                                textBaselineY - placed.Height,
                                placed.Width,
                                placed.Height,
                                pages.Count + 1,
                                SourceBlockIndex: elementIndex,
                                SourceParagraphIndex: 0, Story: null));
                        }

                        emissionSegments = ShiftSegmentsPastMidLineImages(lineShape.Segments, line.Spans, paragraph, firstLine, lineX, imageShifts);
                    }

                    currentItems.Add(new DocxTextLineLayout(
                        lineShape.Text,
                        firstRun,
                        paragraphFontSize,
                        lineShape.X,
                        cursorY - baselineOffset,
                        lineShape.Width,
                        emissionSegments,
                        SourceBlockIndex: elementIndex,
                        SourceParagraphIndex: 0,
                        SourceLineIndex: lineIndex,
                        LineHeight: lineHeight,
                        AppliedBeforeSpacing: firstLine ? spacingProfile.AppliedBeforeSpacing : 0d,
                        IsFirstParagraphLine: firstLine,
                        EndsWithIntraTokenBreak: line.EndsWithIntraTokenBreak,
                        SingleLineHeight: lineHeightProfile.SingleLineHeight,
                        ListLabelSingleLineHeight: lineHeightProfile.ListLabelSingleLineHeight,
                        BodyWindowsLineHeight: lineHeightProfile.BodyWindowsLineHeight,
                        ListLabelWindowsLineHeight: lineHeightProfile.ListLabelWindowsLineHeight,
                        EffectiveLineSpacingFactor: lineHeightProfile.EffectiveLineSpacingFactor,
                        LineSpacingFactorFloorApplied: lineHeightProfile.LineSpacingFactorFloorApplied,
                        LineHeightSource: lineHeightProfile.Source,
                        PendingAfterSpacing: firstLine ? spacingProfile.PendingAfterSpacing : null,
                        ParagraphBeforeSpacing: firstLine ? spacingProfile.ParagraphBeforeSpacing : null,
                        ParagraphAfterSpacing: firstLine ? spacingProfile.ParagraphAfterSpacing : null,
                        ContextualSpacingSuppressed: firstLine ? spacingProfile.ContextualSpacingSuppressed : null,
                        SourceParagraph: paragraph,
                        Story: DocxStoryId.Body(), EmitsTerminalParagraphMark: false));
                    if (lineImages is not null)
                    {
                        currentItems.AddRange(lineImages);
                    }
                    activeColumnHasContent = true;
                    firstLine = false;
                    paragraphX = x + continuationTextStartOffset;
                    paragraphWidth = Math.Max(1d, width - continuationTextStartOffset - GetParagraphRightInset(paragraph, paragraphSpacingScale));
                    cursorY -= lineHeight;
                }

                // RV06: break-adjacent spacing rows (Word 16.0 spill-matrix and
                // trailing-space probes, plus the row-end matrix): every non-empty
                // text-only body paragraph keeps exactly one row-end space beyond
                // authored trailing on its last wrapped line (clean-ending lines
                // already carry it from the wrap pipeline, break or not), and a
                // paragraph immediately followed by a break-only page-break
                // paragraph additionally carries a two-space spill row on the same
                // page past the after-spacing gap (second space at body-left plus
                // 144 design points). Content-independent across clean, trailing
                // (one or three spaces) and comment-range paragraphs. Scoped to
                // text-only body paragraphs (hand-built null breaks and inline
                // splits stay open for spill, as do centered/right and table-cell
                // patterns).
                int breakSpillLineIndex = FindBreakSpillLineIndex(currentItems, paragraph);
                bool breakParagraphHasVisibleText = textSpans.Any(static span => span.Text.Any(static character => !char.IsWhiteSpace(character)));
                double breakSpaceWidth = textMeasurer.MeasureText(firstRun, " ", paragraphFontSize);
                if (paragraph.Images.Count == 0 &&
                    paragraph.InlineTextBoxes.Count == 0 &&
                    lines.Length > 0 &&
                    lines[^1].Text.EndsWith(' ') &&
                    breakParagraphHasVisibleText &&
                    breakSpillLineIndex >= 0)
                {
                    DocxTextLineLayout breakLastLine = (DocxTextLineLayout)currentItems[breakSpillLineIndex];
                    var breakRowEndSegment = new DocxTextSegmentLayout(
                        " ",
                        firstRun,
                        breakLastLine.X + breakLastLine.Width,
                        breakSpaceWidth,
                        paragraphFontSize,
                        0d,
                        0d,
                        DocxTextStateCharacterSpacingSource.None,
                        true,
                        -1,
                        0,
                        DocxTextSegmentRole.BreakSpill);
                    currentItems[breakSpillLineIndex] = breakLastLine with
                    {
                        Text = breakLastLine.Text + " ",
                        Width = breakLastLine.Width + breakSpaceWidth,
                        Segments = [.. breakLastLine.Segments, breakRowEndSegment],
                    };
                }

                if (paragraph.Images.Count == 0 &&
                    paragraph.InlineTextBoxes.Count == 0 &&
                    lines.Length > 0 &&
                    elementIndex + 1 < document.BodyElements.Count &&
                    document.BodyElements[elementIndex + 1] is DocxPageBreakElement nextBreak &&
                    nextBreak.BreakParagraph is { } nextBreakParagraph &&
                    nextBreakParagraph.Images.Count == 0 &&
                    nextBreakParagraph.InlineTextBoxes.Count == 0 &&
                    nextBreakParagraph.Runs.All(static run => run.Text.Length == 0) &&
                    breakSpillLineIndex >= 0)
                {
                    // Office sets the spill row past the paragraph after-spacing gap,
                    // not on the immediate next line slot.
                    double breakSpillAfterSpacing = spacingProfile.ParagraphAfterSpacing;
                    double breakSpillX = x + continuationTextStartOffset;
                    double breakSpillTabX = breakSpillX + 144d * paragraphSpacingScale;
                    double breakSpillBaseline = DocxLineMetrics.ResolveBodyBaselineOffset(paragraphFontSize, lineHeight, IsExactLineSpacing(effective));
                    if (HasNoSpacingElement(effective) && Math.Abs(paragraphFontSize - 11d) < 0.000000001d)
                    {
                        breakSpillBaseline += UntokenedParagraphBaselineExtraPoints;
                    }

                    currentItems.Add(new DocxTextLineLayout(
                        "  ",
                        firstRun,
                        paragraphFontSize,
                        breakSpillX,
                        cursorY - breakSpillAfterSpacing - breakSpillBaseline,
                        (breakSpillTabX - breakSpillX) + breakSpaceWidth,
                        [
                            new DocxTextSegmentLayout(
                                " ",
                                firstRun,
                                breakSpillX,
                                breakSpaceWidth,
                                paragraphFontSize,
                                0d,
                                0d,
                                DocxTextStateCharacterSpacingSource.None,
                                true,
                                -1,
                                0,
                                DocxTextSegmentRole.BreakSpill),
                            new DocxTextSegmentLayout(
                                " ",
                                firstRun,
                                breakSpillTabX,
                                breakSpaceWidth,
                                paragraphFontSize,
                                0d,
                                0d,
                                DocxTextStateCharacterSpacingSource.None,
                                true,
                                -1,
                                0,
                                DocxTextSegmentRole.BreakSpill),
                        ],
                        SourceBlockIndex: elementIndex,
                        SourceParagraphIndex: 0,
                        SourceLineIndex: lines.Length,
                        LineHeight: lineHeight,
                        AppliedBeforeSpacing: 0d,
                        IsFirstParagraphLine: false,
                        EndsWithIntraTokenBreak: lines[^1].EndsWithIntraTokenBreak,
                        SingleLineHeight: lineHeightProfile.SingleLineHeight,
                        ListLabelSingleLineHeight: lineHeightProfile.ListLabelSingleLineHeight,
                        BodyWindowsLineHeight: lineHeightProfile.BodyWindowsLineHeight,
                        ListLabelWindowsLineHeight: lineHeightProfile.ListLabelWindowsLineHeight,
                        EffectiveLineSpacingFactor: lineHeightProfile.EffectiveLineSpacingFactor,
                        LineSpacingFactorFloorApplied: lineHeightProfile.LineSpacingFactorFloorApplied,
                        LineHeightSource: lineHeightProfile.Source,
                        PendingAfterSpacing: null,
                        ParagraphBeforeSpacing: null,
                        ParagraphAfterSpacing: null,
                        ContextualSpacingSuppressed: null,
                        SourceParagraph: paragraph,
                        Story: DocxStoryId.Body(), EmitsTerminalParagraphMark: false));
                    cursorY -= breakSpillAfterSpacing + lineHeight;
                }
            }
            else if (paragraph.Images.Count == 0 && paragraph.InlineTextBoxes.Count == 0)
            {
                if (cursorY - lineHeight < CurrentFrameBottom() && HasCurrentColumnContent())
                {
                    AdvanceColumnOrPage();
                    EnsureFootnoteReserveForSourceBlock(elementIndex);
                }

                cursorY -= ResolveListLabelFirstLineExtraLeading(paragraph, paragraphFontSize, textMeasurer);
                cursorY -= lineHeight;
                activeColumnHasContent = true;
            }

            for (int imageIndex = 0; imageIndex < paragraph.Images.Count; imageIndex++)
            {
                DocxInlineImage image = paragraph.Images[imageIndex];
                if (midLinePlan?.PlacedMask[imageIndex] == true)
                {
                    continue;
                }
                cancellationToken.ThrowIfCancellationRequested();
                double imageWidth = Math.Min(width, image.WidthPoints);
                double imageHeight = image.HeightPoints * imageWidth / Math.Max(1d, image.WidthPoints);
                if (cursorY - imageHeight < CurrentFrameBottom() && HasCurrentColumnContent())
                {
                    AdvanceColumnOrPage();
                    EnsureFootnoteReserveForSourceBlock(elementIndex);
                }

                double imageX = effective.Alignment switch
                {
                    DocxTextAlignment.Center => x + Math.Max(0, width - imageWidth) / 2d,
                    DocxTextAlignment.Right => x + Math.Max(0, width - imageWidth),
                    _ => x
                };
                currentItems.Add(new DocxInlineImageLayout(
                    image,
                    imageX,
                    cursorY - imageHeight,
                    imageWidth,
                    imageHeight,
                    pages.Count + 1,
                    SourceBlockIndex: elementIndex,
                    SourceParagraphIndex: 0, Story: null));
                activeColumnHasContent = true;
                cursorY -= imageHeight + InlineImageParagraphGapPoints;
            }

            foreach (DocxInlineTextBox textBox in paragraph.InlineTextBoxes)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (cursorY - EstimateInlineTextBoxHeight(textBox, paragraphSpacingScale) < CurrentFrameBottom() && HasCurrentColumnContent())
                {
                    AdvanceColumnOrPage();
                    EnsureFootnoteReserveForSourceBlock(elementIndex);
                }

                DocxInlineTextBoxLayout? textBoxLayout = CreateInlineTextBoxLayout(
                    textBox,
                    elementIndex,
                    x,
                    width,
                    cursorY,
                    effective.Alignment,
                    textMeasurer,
                    defaultTabStopPoints,
                    paragraphSpacingScale,
                    pages.Count + 1,
                    cancellationToken);
                if (textBoxLayout is null)
                {
                    continue;
                }

                currentItems.Add(textBoxLayout);
                activeColumnHasContent = true;
                cursorY -= textBoxLayout.BoxHeight + InlineImageParagraphGapPoints;
            }

            pendingSpacingAfter = spacingProfile.ParagraphAfterSpacing;
            previousParagraph = paragraph;
        }

        if (HasPageContent() || pages.Count == 0)
        {
            FinishPage();
        }

        DocxLayoutPage[] pagesWithRelatedStories = AddPlacedRelatedStories(document, pages, GetRelatedStoryLayouts, cancellationToken).ToArray();
        var staticContent = AddStaticContent(pagesWithRelatedStories, textMeasurer, defaultTabStopPoints, paragraphSpacingScale, unscaledTextMeasurer, cancellationToken);
        DocxLayoutPage[] pagesWithStaticText = staticContent.Pages.ToArray();
        IReadOnlyDictionary<int, double> footerContentTopByPage = staticContent.FooterContentTopByPage;
        return new DocxLayout(
            pagesWithStaticText,
            CreateFloatingDrawingLayouts(document.FloatingDrawings, pagesWithStaticText, textMeasurer, defaultTabStopPoints, paragraphSpacingScale, cancellationToken, unscaledTextMeasurer),
            CreateStaticFloatingDrawingLayouts(pagesWithStaticText, textMeasurer, defaultTabStopPoints, paragraphSpacingScale, cancellationToken, unscaledTextMeasurer),
            relatedStoryLayouts,
            staticContent.HeaderContentBottomByPage,
            footerContentTopByPage);
    }

    private static IReadOnlyList<DocxTextLineLayout> ShiftTextLines(IReadOnlyList<DocxTextLineLayout> lines, double deltaY, double deltaX)
    {
        return lines
            .Select(line => line with
            {
                X = line.X + deltaX,
                BaselineY = line.BaselineY + deltaY,
                Segments = line.Segments
                    .Select(segment => segment with { X = segment.X + deltaX })
                    .ToArray()
            })
            .ToArray();
    }

    private static IReadOnlyList<DocxInlineImageLayout> ShiftInlineImages(IReadOnlyList<DocxInlineImageLayout> images, double deltaY, double deltaX)
    {
        return images
            .Select(image => image with { X = image.X + deltaX, Y = image.Y + deltaY })
            .ToArray();
    }

    private static IReadOnlyList<DocxFloatingDrawingLayout> ShiftFloatingDrawings(
        IReadOnlyList<DocxFloatingDrawingLayout> drawings,
        int pageIndex,
        double deltaY,
        double deltaX)
    {
        return drawings
            .Select(drawing => drawing with
            {
                PageStartIndex = pageIndex,
                PageEndIndex = pageIndex,
                AnchorPageIndex = pageIndex,
                HorizontalReferenceX = ShiftNullable(drawing.HorizontalReferenceX, deltaX),
                AnchorBlockVerticalTop = ShiftNullable(drawing.AnchorBlockVerticalTop, deltaY),
                AnchorBlockVerticalBottom = ShiftNullable(drawing.AnchorBlockVerticalBottom, deltaY),
                VerticalReferenceTop = ShiftNullable(drawing.VerticalReferenceTop, deltaY),
                VerticalReferenceBottom = ShiftNullable(drawing.VerticalReferenceBottom, deltaY),
                PlacedX = ShiftNullable(drawing.PlacedX, deltaX),
                PlacedTop = ShiftNullable(drawing.PlacedTop, deltaY),
                WrapExclusionX = ShiftNullable(drawing.WrapExclusionX, deltaX),
                WrapExclusionTop = ShiftNullable(drawing.WrapExclusionTop, deltaY)
            })
            .ToArray();
    }

    private static double? ShiftNullable(double? value, double delta)
    {
        return value is null ? null : value.Value + delta;
    }

    private static DocxInlineTextBoxLayout ShiftInlineTextBox(DocxInlineTextBoxLayout box, double deltaY, double deltaX)
    {
        return box with
        {
            BoxX = box.BoxX + deltaX,
            BoxTop = box.BoxTop + deltaY,
            TextLines = ShiftTextLines(box.TextLines, deltaY, deltaX),
            InlineImages = ShiftInlineImages(box.InlineImages, deltaY, deltaX),
            TableRows = ShiftTableRows(box.TableRows, deltaY, deltaX)
        };
    }

    private static IReadOnlyList<DocxTableRowLayout> ShiftTableRows(IReadOnlyList<DocxTableRowLayout> rows, double deltaY, double deltaX)
    {
        return rows
            .Select(row => row with
            {
                Table = row.Table with
                {
                    TableX = row.Table.TableX + deltaX
                },
                Y = row.Y + deltaY,
                Cells = row.Cells.Select(cell => ShiftTableCell(cell, deltaY, deltaX)).ToArray()
            })
            .ToArray();
    }

    private static DocxTableCellLayout ShiftTableCell(DocxTableCellLayout cell, double deltaY, double deltaX)
    {
        return cell with
        {
            X = cell.X + deltaX,
            Y = cell.Y + deltaY,
            TextLines = ShiftTextLines(cell.TextLines, deltaY, deltaX),
            InlineImages = ShiftInlineImages(cell.InlineImages, deltaY, deltaX),
            NestedTableRows = ShiftTableRows(cell.NestedRows, deltaY, deltaX)
        };
    }

    // W6-a1: fixed table geometry joins scaled space; fixedScale reuses the layout
    // spacing scale (identical in every mode).
    private static double ResolveTableCellHorizontalPadding(double? points, double fixedScale)
    {
        return Math.Max(0d, points ?? 0d) * fixedScale;
    }

    // Office A/B (w63-w68 table border/margin probes, Word-COM rendered plus
    // PdfInspect): cell text starts at the grid plus max(borderHalf, margin).
    // Unset margins default to 0.48pt (nil-border text sits at grid + 0.48);
    // explicit margins replace the default and swallow border halves; explicit-0
    // restores content anchoring. W6-a1: the default joins scaled space through
    // fixedScale like the other fixed table insets.
    private const double DefaultTableCellHorizontalMarginPoints = 0.48d;

    private static double ResolveTableCellHorizontalEdgeInset(DocxTableCell cell, string edge, double? marginPoints, double fixedScale)
    {
        double borderHalf = DocxTableBorderGeometry.ResolveVisibleWidth(DocxTableBorderGeometry.Find(cell.Borders, edge)) / 2d;
        double margin = Math.Max(0d, marginPoints ?? DefaultTableCellHorizontalMarginPoints);
        if (!cell.PinTextToMargin)
        {
            return Math.Max(borderHalf, margin) * fixedScale;
        }

        // Office A/B (w72/m12 pin probes): pinned text clears the style margin box
        // (borders swallowed); direct margins still apply by cascade (w74).
        double style = string.Equals(edge, "left", StringComparison.OrdinalIgnoreCase)
            ? cell.StyleMargins?.LeftPoints ?? 0d
            : cell.StyleMargins?.RightPoints ?? 0d;
        return Math.Max(style, margin) * fixedScale;
    }

    private static double ResolveTableCellBorderContentInset(DocxTableCell cell, string edge, double fixedScale)
    {
        return DocxTableBorderGeometry.ResolveVisibleWidth(DocxTableBorderGeometry.Find(cell.Borders, edge)) / 2d * fixedScale;
    }

    private static double ResolveTableCellVerticalPadding(double? points, double fixedScale)
    {
        return Math.Max(0d, points ?? 0d) * fixedScale;
    }

    private static double ResolveTableRowTopPadding(DocxTableRow row, double fixedScale)
    {
        return row.Cells
            .Select(cell => ResolveTableCellVerticalPadding(cell.Margins.TopPoints, fixedScale))
            .DefaultIfEmpty(0d)
            .Max();
    }

    private static double ResolveTableCellFirstBaselineInset(IReadOnlyList<DocxParagraph> paragraphs)
    {
        return DocxLineMetrics.ResolveTableCellFirstBaselineInset(paragraphs);
    }

}
