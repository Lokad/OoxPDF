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
    private readonly bool scaleBaselineOffsetTransitions;
    private readonly bool retuneReserveToPrintScale;
    private readonly double reservePrintScale;
    private readonly IReadOnlyDictionary<string, string>? commentMarkerLabels;

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
        public double? MarkupLaneDesignBodyEnd { get; init; }
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

    public DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode markupGeometryMode, double wordCompatiblePrintScale = WordCompatibleAllMarkupParagraphSpacingScale, IReadOnlyDictionary<string, string>? commentMarkerLabels = null)
    {
        reserveMarkupMargin = markupGeometryMode is OoxPdfDocxMarkupGeometryMode.ReserveMarkupMargin or OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
        paragraphSpacingScale = markupGeometryMode == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
            ? wordCompatiblePrintScale
            : 1d;
        retuneReserveToPrintScale = markupGeometryMode == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
        scaleBaselineOffsetTransitions = markupGeometryMode == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
        reservePrintScale = wordCompatiblePrintScale;
        this.commentMarkerLabels = commentMarkerLabels;
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
        var inFlightNotes = new Dictionary<(DocxRelatedStoryKind Kind, string Id), InFlightRelatedStory>();
        var placedStoryKeys = new HashSet<(DocxRelatedStoryKind Kind, string Id)>();
        Dictionary<int, List<DocxInlineReferenceLocation>> locationsByBlock = BuildInlineReferenceLocationsByBlock(document, cancellationToken);
        var storyLookupByWidth = new Dictionary<double, Dictionary<(DocxRelatedStoryKind Kind, string Id), DocxRelatedStoryLayout>>();
        var keepLookaheadYielded = new HashSet<(DocxRelatedStoryKind Kind, string Id)>();
        var footnotePlacedLinesAtLastYield = new Dictionary<(DocxRelatedStoryKind Kind, string Id), int>();
        var continuationHeaderByPage = new Dictionary<(DocxPageSettings, double, int, int), DocxStaticStoryLayoutResult?>();
        double? bodyFirstBaselineMemo = null;
        bool bodyFirstBaselineComputed = false;

        IReadOnlyDictionary<int, DocxEffectiveSectionSettings> sectionSettingsByElementIndex = BuildEffectiveSectionSettings(document, out DocxEffectiveSectionSettings finalSectionSettings);
        DocxEffectiveSectionSettings activeSectionSettings = FindSectionSettingsAtOrAfter(document.BodyElements, 0, sectionSettingsByElementIndex) ?? finalSectionSettings;
        DocxPageGeometry page = ResolveSectionGeometry(document, activeSectionSettings, reserveMarkupMargin, retuneReserveToPrintScale, reservePrintScale, pageNumber: 1);
        int activeColumnIndex = 0;
        double x = ResolveActiveColumnFrame(page, activeColumnIndex).X;
        double width = ResolveActiveColumnFrame(page, activeColumnIndex).Width;
        double cursorY = ResolvePageStartCursor();
        double pendingSpacingAfter = 0d;
        DocxParagraph? previousParagraph = null;
        double? firstBodyLineBaselineOffset = null;
        double? firstDocumentBodyBaselineOffset = null;
        bool activeColumnHasContent = false;
        int tableIndex = 0;
        double defaultTabStopPoints = document.Settings.DefaultTabStopPoints ?? WordDefaultTabStopPoints;
        bool hasSimpleScaledBodyFrame =
            scaleBaselineOffsetTransitions && paragraphSpacingScale < 1d &&
            document.HeaderParagraphs.Count == 0 && document.FooterParagraphs.Count == 0 &&
            document.HeaderParagraphsByType.Count == 0 && document.FooterParagraphsByType.Count == 0 &&
            document.HeaderBodyElementsByType.Count == 0 && document.FooterBodyElementsByType.Count == 0 &&
            document.HeaderFloatingDrawingsByType.Count == 0 && document.FooterFloatingDrawingsByType.Count == 0 &&
            document.PageSettings.HeaderParagraphsByType.Count == 0 &&
            document.PageSettings.FooterParagraphsByType.Count == 0 &&
            document.PageSettings.HeaderBodyElementsByType.Count == 0 &&
            document.PageSettings.FooterBodyElementsByType.Count == 0 &&
            document.PageSettings.HeaderFloatingDrawingsByType.Count == 0 &&
            document.PageSettings.FooterFloatingDrawingsByType.Count == 0 &&
            !document.RelatedStories.Any(story => story.Kind is DocxRelatedStoryKind.Footnote or DocxRelatedStoryKind.Endnote) &&
            !document.BodyElements.Any(element => element is DocxTableElement or DocxSectionBreakElement ||
                element is DocxParagraphElement bodyParagraph &&
                (bodyParagraph.Paragraph.Images.Count != 0 || bodyParagraph.Paragraph.InlineTextBoxes.Count != 0));
        var relatedStoryLayoutsByBodyWidth = new Dictionary<double, IReadOnlyList<DocxRelatedStoryLayout>>();
        var footnoteReserveHeightByBodyWidth = new Dictionary<double, IReadOnlyDictionary<int, double>>();

        IReadOnlyList<DocxRelatedStoryLayout> GetFootnoteStoryLayouts(DocxPageGeometry geometry)
        {
            return GetRelatedStoryLayouts(FootnoteStoryLayoutWidth(geometry));
        }

        // RV06 interleaving (Word 16.0, edge-footlong-mixed WC): Office keeps footnote
        // text at the full body width while the review lane shrinks the body, so
        // footnote takes match Final instead of wrapping to the narrowed lane.

        static double FootnoteStoryLayoutWidth(DocxPageGeometry geometry)
        {
            return Math.Max(1d, geometry.BodyWidth + geometry.MarkupMarginReservePoints);
        }

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
                reserveHeights = CreateFootnoteReserveHeightBySourceBlock(document, GetRelatedStoryLayouts(key), cancellationToken, unscaledTextMeasurer ?? textMeasurer);
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
                currentItems.ToArray())
            {
                MarkupLaneDesignBodyEnd = page.MarkupLaneDesignBodyEnd
            });
            PlaceInFlightFootnotesOnCompletingPage();
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
            firstBodyLineBaselineOffset = null;
            activeColumnHasContent = false;
            currentPageFootnoteReserveHeight = FootnoteRemainderTotal();
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

        IDocxTextMeasurer? separatorMeasurerForNotes = unscaledTextMeasurer ?? textMeasurer;

        bool FootnoteReserveYieldsPageToDrain(double itemHeight)
        {
            if (HasCurrentColumnContent() || currentPageFootnoteReserveHeight <= 0d || cursorY - itemHeight >= CurrentFrameBottom())
            {
                return false;
            }

            bool drainableProgress = false;
            foreach (InFlightRelatedStory inFlight in inFlightNotes.Values)
            {
                if (inFlight.RemainingLineCount <= 0 || inFlight.StoryLayout.TextLines.Count == 0)
                {
                    continue;
                }

                if (inFlight.Location.Reference.Id is null)
                {
                    continue;
                }

                var drainableKey = (inFlight.Location.Reference.Kind, inFlight.Location.Reference.Id);
                if (!placedStoryKeys.Contains(drainableKey))
                {
                    continue;
                }

                int placedBefore = footnotePlacedLinesAtLastYield.TryGetValue(drainableKey, out int before) ? before : -1;
                if (inFlight.PlacedLineCount > placedBefore)
                {
                    drainableProgress = true;
                    footnotePlacedLinesAtLastYield[drainableKey] = inFlight.PlacedLineCount;
                }
            }

            return drainableProgress;
        }

        void AdvanceForOverflowingItem(double itemHeight, int sourceBlockIndex)
        {
            AdvanceColumnOrPage();
            RegisterInFlightFootnotesForSourceBlock(sourceBlockIndex);
            while (!HasCurrentColumnContent() && cursorY - itemHeight < CurrentFrameBottom() && FootnoteReserveYieldsPageToDrain(itemHeight))
            {
                cancellationToken.ThrowIfCancellationRequested();
                AdvanceColumnOrPage();
                RegisterInFlightFootnotesForSourceBlock(sourceBlockIndex);
            }
        }

        // Keep decisions look ahead: the chain about to move must reserve the notes
        // its not-yet-laid follower blocks will register, measured exactly like the
        // body reserve. Measured notes are consumed from future lookahead so follower
        // checks on later pages do not re-advance the same chain; in-flight notes live
        // in the body reserve and the checking block itself stays exempt like Register.
        double KeepChainFootnoteReserve(int sourceBlockIndex)
        {
            double reserve = 0d;
            IReadOnlyList<DocxRelatedStoryLayout> pageStoryLayouts = GetFootnoteStoryLayouts(page);
            double widthKey = Math.Round(Math.Max(1d, page.BodyWidth), 3);
            if (!storyLookupByWidth.TryGetValue(widthKey, out Dictionary<(DocxRelatedStoryKind Kind, string Id), DocxRelatedStoryLayout>? storyByKey))
            {
                storyByKey = CreateRelatedStoryLookup(pageStoryLayouts);
                storyLookupByWidth[widthKey] = storyByKey;
            }

            DocxRelatedStoryLayout? separatorLayout = FindSpecialRelatedStoryLayout(pageStoryLayouts, DocxRelatedStoryKind.Footnote, DocxRelatedStoryType.Separator);
            DocxRelatedStoryLayout? continuationLayout = FindSpecialRelatedStoryLayout(pageStoryLayouts, DocxRelatedStoryKind.Footnote, DocxRelatedStoryType.ContinuationSeparator) ?? separatorLayout;
            foreach (int chainBlock in KeepChainBlockIndexes(document.BodyElements, sourceBlockIndex))
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (chainBlock == sourceBlockIndex)
                {
                    continue;
                }

                if (!locationsByBlock.TryGetValue(chainBlock, out List<DocxInlineReferenceLocation>? blockLocations))
                {
                    continue;
                }

                foreach (DocxInlineReferenceLocation location in blockLocations)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    DocxInlineReference reference = location.Reference;
                    if (reference.Kind != DocxRelatedStoryKind.Footnote || reference.Id is null)
                    {
                        continue;
                    }

                    var key = (reference.Kind, reference.Id);
                    if (placedStoryKeys.Contains(key) || keepLookaheadYielded.Contains(key))
                    {
                        continue;
                    }

                    if (!storyByKey.TryGetValue(key, out DocxRelatedStoryLayout? storyLayout) || storyLayout.ContentHeight <= 0d)
                    {
                        continue;
                    }

                    if (storyLayout.TextLines.Count == 0)
                    {
                        continue;
                    }

                    keepLookaheadYielded.Add(key);
                    var transient = new InFlightRelatedStory(storyLayout, separatorLayout, continuationLayout, chainBlock, location);
                    reserve += FootnoteRemainderHeight(transient, GetStoryTextLineBoxHeights(storyLayout), separatorMeasurerForNotes);
                }
            }

            return Math.Max(0d, reserve);
        }

        double FootnoteRemainderTotal()
        {
            double total = 0d;
            foreach (InFlightRelatedStory inFlight in inFlightNotes.Values)
            {
                total += FootnoteRemainderHeight(inFlight, GetStoryTextLineBoxHeights(inFlight.StoryLayout), separatorMeasurerForNotes);
            }

            return Math.Max(0d, total);
        }

        void RegisterInFlightFootnotesForSourceBlock(int sourceBlockIndex)
        {
            if (!locationsByBlock.TryGetValue(sourceBlockIndex, out List<DocxInlineReferenceLocation>? blockLocations))
            {
                currentPageFootnoteReserveHeight = Math.Max(currentPageFootnoteReserveHeight, FootnoteRemainderTotal());
                return;
            }
            // A block with page content already placed never reserves against notes
            // it registers itself, so the ref-bearing paragraph shares its page with
            // the note head instead of evicting itself; later blocks see the full
            // remainder below. Blocks starting a fresh page keep the full bump, so
            // small-note marker pagination is unchanged.

            IReadOnlyList<DocxRelatedStoryLayout> pageStoryLayouts = GetFootnoteStoryLayouts(page);
            double widthKey = Math.Round(Math.Max(1d, page.BodyWidth), 3);
            if (!storyLookupByWidth.TryGetValue(widthKey, out Dictionary<(DocxRelatedStoryKind Kind, string Id), DocxRelatedStoryLayout>? storyByKey))
            {
                storyByKey = CreateRelatedStoryLookup(pageStoryLayouts);
                storyLookupByWidth[widthKey] = storyByKey;
            }

            foreach (DocxInlineReferenceLocation location in blockLocations)
            {
                cancellationToken.ThrowIfCancellationRequested();
                DocxInlineReference reference = location.Reference;
                if (reference.Kind != DocxRelatedStoryKind.Footnote || reference.Id is null)
                {
                    continue;
                }

                var key = (reference.Kind, reference.Id);
                if (placedStoryKeys.Contains(key) || inFlightNotes.ContainsKey(key))
                {
                    continue;
                }

                if (!storyByKey.TryGetValue(key, out DocxRelatedStoryLayout? storyLayout) || storyLayout.ContentHeight <= 0d)
                {
                    continue;
                }

                DocxRelatedStoryLayout? separatorLayout = FindSpecialRelatedStoryLayout(pageStoryLayouts, DocxRelatedStoryKind.Footnote, DocxRelatedStoryType.Separator);
                DocxRelatedStoryLayout? continuationLayout = FindSpecialRelatedStoryLayout(pageStoryLayouts, DocxRelatedStoryKind.Footnote, DocxRelatedStoryType.ContinuationSeparator) ?? separatorLayout;
                inFlightNotes[key] = new InFlightRelatedStory(storyLayout, separatorLayout, continuationLayout, sourceBlockIndex, location);
            }

            double registeringBlockRemainder = 0d;
            if (HasCurrentColumnContent())
            {
                foreach (InFlightRelatedStory inFlight in inFlightNotes.Values)
                {
                    if (inFlight.SourceBlockIndex != sourceBlockIndex)
                    {
                        continue;
                    }

                    registeringBlockRemainder += FootnoteRemainderHeight(inFlight, GetStoryTextLineBoxHeights(inFlight.StoryLayout), separatorMeasurerForNotes);
                }
            }

            currentPageFootnoteReserveHeight = Math.Max(currentPageFootnoteReserveHeight, Math.Max(0d, FootnoteRemainderTotal() - registeringBlockRemainder));
        }

        void PlaceInFlightFootnotesOnCompletingPage()
        {
            if (inFlightNotes.Count == 0)
            {
                return;
            }

            var blocksOnPage = new HashSet<int>();
            var pageSegments = new List<(DocxTextLineLayout Line, int RunIndex, int Start, int End)>();
            foreach (DocxLayoutItem item in currentItems)
            {
                foreach (DocxPageTextLineOwner owner in EnumerateTextLineOwners(item, null))
                {
                    DocxTextLineLayout line = owner.Line;
                    if (owner.SourceBlockIndex is { } blockIndex)
                    {
                        blocksOnPage.Add(blockIndex);
                    }

                    foreach (DocxTextSegmentLayout segment in line.Segments)
                    {
                        int start = Math.Max(0, segment.SourceTextOffsetInRun);
                        pageSegments.Add((line, segment.SourceTextRunIndex, start, start + segment.Text.Length));
                    }
                }
            }

            IReadOnlyList<DocxRelatedStoryLayout> pageStoryLayouts = GetFootnoteStoryLayouts(page);
            double widthKey = Math.Round(Math.Max(1d, page.BodyWidth), 3);
            if (!storyLookupByWidth.TryGetValue(widthKey, out Dictionary<(DocxRelatedStoryKind Kind, string Id), DocxRelatedStoryLayout>? storyByKey))
            {
                storyByKey = CreateRelatedStoryLookup(pageStoryLayouts);
                storyLookupByWidth[widthKey] = storyByKey;
            }

            var pageFootnoteStories = new List<DocxReferencedRelatedStoryLayout>();
            foreach (int blockIndex in blocksOnPage.OrderBy(index => index))
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (blockIndex < 0 || blockIndex >= document.BodyElements.Count)
                {
                    continue;
                }

                if (!locationsByBlock.TryGetValue(blockIndex, out List<DocxInlineReferenceLocation>? blockLocations))
                {
                    continue;
                }

                foreach (DocxInlineReferenceLocation location in blockLocations)
                {
                    DocxInlineReference reference = location.Reference;
                    if (reference.Kind != DocxRelatedStoryKind.Footnote || reference.Id is null)
                    {
                        continue;
                    }

                    var key = (reference.Kind, reference.Id);
                    if (placedStoryKeys.Contains(key))
                    {
                        continue;
                    }

                    if (!storyByKey.TryGetValue(key, out DocxRelatedStoryLayout? storyLayout) || storyLayout.ContentHeight <= 0d)
                    {
                        continue;
                    }

                    if (storyLayout.TextLines.Count == 0)
                    {
                        continue;
                    }

                    if (location.Reference.SourceRunIndex < 0)
                    {
                        continue;
                    }

                    bool matched = false;
                    foreach ((DocxTextLineLayout segmentLine, int runIndex, int start, int end) in pageSegments)
                    {
                        if (runIndex == location.Reference.SourceRunIndex &&
                            start <= location.Reference.TextOffsetInRun &&
                            location.Reference.TextOffsetInRun < end &&
                            (segmentLine.SourceParagraph is null || ReferenceEquals(segmentLine.SourceParagraph, location.SourceParagraph)))
                        {
                            matched = true;
                            break;
                        }
                    }

                    if (!matched)
                    {
                        continue;
                    }

                    pageFootnoteStories.Add(new DocxReferencedRelatedStoryLayout(location, storyLayout));
                }
            }

            var freshKeys = new HashSet<(DocxRelatedStoryKind Kind, string Id)>();
            foreach (DocxReferencedRelatedStoryLayout fresh in pageFootnoteStories)
            {
                if (fresh.Location.Reference.Id is not null)
                {
                    freshKeys.Add((fresh.Location.Reference.Kind, fresh.Location.Reference.Id));
                }
            }

            var continuedStories = new List<DocxReferencedRelatedStoryLayout>();
            foreach (InFlightRelatedStory inFlight in inFlightNotes.Values)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (inFlight.RemainingLineCount <= 0 || inFlight.StoryLayout.TextLines.Count == 0)
                {
                    continue;
                }

                var continuedKey = (inFlight.Location.Reference.Kind, inFlight.Location.Reference.Id ?? string.Empty);
                if (freshKeys.Contains(continuedKey))
                {
                    continue;
                }

                if (!placedStoryKeys.Contains(continuedKey))
                {
                    continue;
                }

                DocxRelatedStoryLayout narrowedLayout = NarrowStoryTextLinesForOffset(inFlight.StoryLayout, inFlight.PlacedLineCount);
                continuedStories.Add(new DocxReferencedRelatedStoryLayout(inFlight.Location, narrowedLayout));
            }

            if (pageFootnoteStories.Count == 0 && continuedStories.Count == 0)
            {
                return;
            }

            DocxLayoutPage templatePage = pages[^1];
            int templatePageNumber = pages.Count;
            var sharedBatch = new List<DocxReferencedRelatedStoryLayout>(continuedStories.Count + pageFootnoteStories.Count);
            sharedBatch.AddRange(continuedStories);
            sharedBatch.AddRange(pageFootnoteStories);
            var sharedPlacedBefore = new Dictionary<(DocxRelatedStoryKind Kind, string Id), int>();
            foreach (DocxReferencedRelatedStoryLayout batched in sharedBatch)
            {
                if (batched.Location.Reference.Id is not null)
                {
                    var batchedKey = (batched.Location.Reference.Kind, batched.Location.Reference.Id);
                    if (inFlightNotes.TryGetValue(batchedKey, out InFlightRelatedStory? batchedFlight))
                    {
                        sharedPlacedBefore[batchedKey] = batchedFlight.PlacedLineCount;
                    }
                }
            }

            var headScratchPages = new List<DocxLayoutPage> { templatePage };
            DocxRelatedStoryLayout? separatorLayout = FindSpecialRelatedStoryLayout(pageStoryLayouts, DocxRelatedStoryKind.Footnote, DocxRelatedStoryType.Separator);
            DocxRelatedStoryLayout? continuationLayout = FindSpecialRelatedStoryLayout(pageStoryLayouts, DocxRelatedStoryKind.Footnote, DocxRelatedStoryType.ContinuationSeparator) ?? separatorLayout;
            IDocxTextMeasurer? separatorMeasurer = unscaledTextMeasurer ?? textMeasurer;
            if (pageFootnoteStories.Count == 0 && templatePage.Items.Count == 0)
            {
                DocxLayoutPage slicePage = headScratchPages[0];
                int slicePageIndex = 0;
                var slicePlacedStories = templatePage.PlacedRelatedStories.ToList();
                double sliceCursorTop = Math.Min(templatePage.Height - templatePage.MarginTop, ResolveHeaderKeepOut(templatePage, templatePageNumber, templatePageNumber));
                foreach (DocxReferencedRelatedStoryLayout continuedStory in sharedBatch)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    slicePage = headScratchPages[slicePageIndex];
                    InFlightRelatedStory? continuedFlight = continuedStory.Location.Reference.Id is not null && inFlightNotes.TryGetValue((continuedStory.Location.Reference.Kind, continuedStory.Location.Reference.Id), out InFlightRelatedStory? continuedTracked) ? continuedTracked : null;
                    sliceCursorTop = PlaceContinuationSeparatorIfNeeded(headScratchPages, ref slicePageIndex, ref slicePage, ref slicePlacedStories, sliceCursorTop, continuedStory.Location.SourceBlockIndex, continuedStory.StoryLayout, 0, GetStoryTextLineBoxHeights(continuedStory.StoryLayout), continuationLayout, separatorMeasurer, paragraphSpacingScale);
                    PlaceRelatedStorySlices(headScratchPages, ref slicePageIndex, ref slicePage, ref slicePlacedStories, ref sliceCursorTop, continuedStory.StoryLayout, continuedStory.Location.SourceBlockIndex, insertContinuationAfterActivePage: true, continuationLayout, separatorMeasurer, paragraphSpacingScale, ResolveHeaderKeepOut, stopAfterCurrentPage: true, inFlight: continuedFlight);
                }
            }
            else if (separatorLayout is null)
            {
                PlaceFootnoteStoriesWithoutSeparator(
                    headScratchPages,
                    0,
                    templatePage,
                    sharedBatch,
                    continuationLayout,
                    separatorMeasurer,
                    paragraphSpacingScale,
                    ResolveHeaderKeepOut,
                    sharedSlicePlacement: true,
                    inFlightNotes: inFlightNotes,
                    pageNumberOverride: templatePageNumber,
                    pageCountOverride: templatePageNumber,
                    sharedContentTop: cursorY - Math.Max(0d, pendingSpacingAfter));
            }
            else
            {
                PlaceFootnoteStories(
                    headScratchPages,
                    0,
                    templatePage,
                    sharedBatch,
                    separatorLayout,
                    continuationLayout,
                    separatorMeasurer,
                    paragraphSpacingScale,
                    ResolveHeaderKeepOut,
                    sharedSlicePlacement: true,
                    inFlightNotes: inFlightNotes,
                    pageNumberOverride: templatePageNumber,
                    pageCountOverride: templatePageNumber,
                    sharedContentTop: cursorY - Math.Max(0d, pendingSpacingAfter));
            }

            var harvestedStories = new List<DocxPlacedRelatedStoryLayout>();
            foreach (DocxPlacedRelatedStoryLayout placed in headScratchPages[0].PlacedRelatedStories)
            {
                harvestedStories.Add(placed);
            }

            if (harvestedStories.Count != 0)
            {
                pages[^1] = templatePage with { PlacedRelatedStories = harvestedStories.ToArray() };
            }

            foreach (KeyValuePair<(DocxRelatedStoryKind Kind, string Id), int> placedBefore in sharedPlacedBefore)
            {
                if (inFlightNotes.TryGetValue(placedBefore.Key, out InFlightRelatedStory? placedFlight) && placedFlight.PlacedLineCount > placedBefore.Value)
                {
                    placedStoryKeys.Add(placedBefore.Key);
                }
            }
        }

        void DrainRemainingInFlightFootnotes()
        {
            if (inFlightNotes.Count == 0 || pages.Count == 0)
            {
                return;
            }

            bool drainStarted = false;
            int tailPageIndex = pages.Count - 1;
            DocxLayoutPage tailPage = pages[^1];
            var tailPlacedStories = new List<DocxPlacedRelatedStoryLayout>();
            double tailCursorTop = 0d;
            foreach (InFlightRelatedStory inFlight in inFlightNotes.Values)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (inFlight.RemainingLineCount <= 0 || inFlight.StoryLayout.TextLines.Count == 0)
                {
                    continue;
                }

                if (inFlight.Location.Reference.Id is null || !placedStoryKeys.Contains((inFlight.Location.Reference.Kind, inFlight.Location.Reference.Id)))
                {
                    continue;
                }

                if (!drainStarted)
                {
                    drainStarted = true;
                    MoveToRelatedStoryContinuationPage(pages, ref tailPageIndex, ref tailPage, ref tailPlacedStories, ref tailCursorTop, insertContinuationAfterActivePage: false, ResolveHeaderKeepOut);
                }

                DocxRelatedStoryLayout narrowedLayout = NarrowStoryTextLinesForOffset(inFlight.StoryLayout, inFlight.PlacedLineCount);
                double[] narrowedLineBoxes = GetStoryTextLineBoxHeights(narrowedLayout);
                tailCursorTop = PlaceContinuationSeparatorIfNeeded(pages, ref tailPageIndex, ref tailPage, ref tailPlacedStories, tailCursorTop, inFlight.SourceBlockIndex, narrowedLayout, 0, narrowedLineBoxes, inFlight.ContinuationLayout, separatorMeasurerForNotes, paragraphSpacingScale);
                PlaceRelatedStorySlices(pages, ref tailPageIndex, ref tailPage, ref tailPlacedStories, ref tailCursorTop, narrowedLayout, inFlight.SourceBlockIndex, insertContinuationAfterActivePage: true, inFlight.ContinuationLayout, separatorMeasurerForNotes, paragraphSpacingScale, ResolveHeaderKeepOut, stopAfterCurrentPage: false, inFlight: inFlight);
            }
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
                firstBodyLineBaselineOffset = null;
                activeColumnHasContent = false;
                currentPageFootnoteReserveHeight = FootnoteRemainderTotal();
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
                    if (cursorY - paragraphAdvance < CurrentFrameBottom() && (HasCurrentColumnContent() || FootnoteReserveYieldsPageToDrain(paragraphAdvance)))
                    {
                        AdvanceForOverflowingItem(paragraphAdvance, elementIndex);
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
                firstBodyLineBaselineOffset = null;
                continue;
            }

            if (element is DocxManualBreakElement manualBreak)
            {
                if (manualBreak.Value?.Equals("column", StringComparison.OrdinalIgnoreCase) == true)
                {
                    DocxParagraph? breakParagraph = manualBreak.BreakParagraph;
                    DocxParagraphSpacingProfile? breakSpacing = breakParagraph is null ? null
                        : ResolveParagraphSpacingProfile(previousParagraph, breakParagraph, pendingSpacingAfter, paragraphSpacingScale);
                    AdvanceColumnOrPage();
                    if (breakParagraph is not null && breakSpacing is not null)
                    {
                        // Word 16 column controls: the paragraph mark resumes in the
                        // next frame, carrying only before-spacing beyond the old after-gap.
                        double extraBefore = Math.Max(0d, breakSpacing.AppliedBeforeSpacing - breakSpacing.PendingAfterSpacing);
                        cursorY -= extraBefore + ResolveLineHeight(breakParagraph, GetParagraphFontSize(breakParagraph), textMeasurer);
                        pendingSpacingAfter = breakSpacing.ParagraphAfterSpacing;
                        previousParagraph = breakParagraph;
                        activeColumnHasContent = true;
                    }
                    continue;
                }

                pendingSpacingAfter = 0d;
                previousParagraph = null;
                firstBodyLineBaselineOffset = null;
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
                firstBodyLineBaselineOffset = null;
                continue;
            }

            if (element is DocxTableElement tableElement)
            {
                double? precedingBodyBaselineInset = firstBodyLineBaselineOffset;
                int completedPagesBeforeTable = pages.Count;
                EnsureFootnoteReserveForSourceBlock(elementIndex);
                RegisterInFlightFootnotesForSourceBlock(elementIndex);
                cursorY -= pendingSpacingAfter;
                pendingSpacingAfter = 0d;
                previousParagraph = null;
                firstBodyLineBaselineOffset = null;
                int itemCountBeforeTable = currentItems.Count;
                int currentTableIndex = tableIndex++;
                bool tableActiveFrameHasContent = false;
                double reviewPaginationScale = scaleBaselineOffsetTransitions && page.ColumnFrames.Count <= 1 &&
                    page.PageSettings.HeaderParagraphsByType.Values.All(paragraphs => paragraphs.Count == 0) &&
                    page.PageSettings.FooterParagraphsByType.Values.All(paragraphs => paragraphs.Count == 0) &&
                    page.PageSettings.HeaderBodyElementsByType.Values.All(elements => elements.Count == 0) &&
                    page.PageSettings.FooterBodyElementsByType.Values.All(elements => elements.Count == 0) &&
                    page.PageSettings.HeaderFloatingDrawingsByType.Values.All(drawings => drawings.Count == 0) &&
                    page.PageSettings.FooterFloatingDrawingsByType.Values.All(drawings => drawings.Count == 0) &&
                    document.RelatedStories.All(story => story.Kind == DocxRelatedStoryKind.Comment) &&
                    CanUseReviewTablePageCapacity(tableElement.Table) ? paragraphSpacingScale : 1d;
                void MarkTableBoundaryContent()
                {
                    tableActiveFrameHasContent = true;
                }

                Action advanceTableBoundaryCore = page.ColumnFrames.Count > 1
                    ? () =>
                    {
                        AdvanceColumnOrPage();
                        tableActiveFrameHasContent = false;
                    }
                    : FinishPage;
                Action advanceTableBoundary = reviewPaginationScale < 1d && textMeasurer is not null
                    ? () =>
                    {
                        ApplyReviewTableBaselineGeometry(currentItems, itemCountBeforeTable, precedingBodyBaselineInset,
                            textMeasurer, paragraphSpacingScale, ref cursorY, cancellationToken, allowPageBoundaryFragments: true);
                        advanceTableBoundaryCore();
                        itemCountBeforeTable = 0;
                    }
                    : advanceTableBoundaryCore;
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
                        paragraphSpacingScale: paragraphSpacingScale,
                        commentMarkerLabels: commentMarkerLabels);
                }

                LayoutTable(tableElement.Table, CurrentFrameBottom(), textMeasurer, defaultTabStopPoints, () => pages.Count + 1, ref currentItems, ref cursorY, ResolveCurrentTableFrame, advanceTableBoundary, hasTableBoundaryContent, MarkTableBoundaryContent, cancellationToken, paragraphSpacingScale, new DocxTableCellTextLinesMemo(),
                    reviewPaginationScale: reviewPaginationScale);
                if (scaleBaselineOffsetTransitions &&
                    (completedPagesBeforeTable == pages.Count || reviewPaginationScale < 1d) && textMeasurer is not null)
                {
                    ApplyReviewTableBaselineGeometry(currentItems, itemCountBeforeTable, precedingBodyBaselineInset,
                        textMeasurer, paragraphSpacingScale, ref cursorY, cancellationToken,
                        allowPageBoundaryFragments: reviewPaginationScale < 1d);
                }
                if (currentItems.Count > itemCountBeforeTable)
                {
                    activeColumnHasContent = true;
                }

                continue;
            }

            if (element is DocxImplicitParagraphElement implicitParagraph &&
                implicitParagraph.SourceKind == DocxBreakSourceKind.TerminalTable)
            {
                DocxTextRun markRun = DocxImplicitParagraphElement.CreateParagraphMarkRun(implicitParagraph.MarkFontSizePoints);
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
                firstBodyLineBaselineOffset = null;
                pendingSpacingAfter = 0d;
                continue;
            }

            if (element is not DocxParagraphElement paragraphElement)
            {
                continue;
            }

            DocxParagraph paragraph = paragraphElement.Paragraph;
            DocxEffectiveParagraphProperties effective = paragraph.EffectiveProperties;
            RegisterInFlightFootnotesForSourceBlock(elementIndex);
            DocxParagraphSpacingProfile spacingProfile = ResolveParagraphSpacingProfile(previousParagraph, paragraph, pendingSpacingAfter, paragraphSpacingScale);
            cursorY -= spacingProfile.AppliedBeforeSpacing;
            pendingSpacingAfter = 0d;
            double paragraphFontSize = GetParagraphFontSize(paragraph);
            DocxLineHeightProfile lineHeightProfile = ResolveLineHeightProfile(paragraph, paragraphFontSize, textMeasurer);
            double baselineLineHeight = lineHeightProfile.LineHeight;
            bool scalesExactBodyAdvance = scaleBaselineOffsetTransitions && page.ColumnFrames.Count <= 1 &&
                ShouldScaleExactBodyAdvance(paragraph, lineHeightProfile, paragraphFontSize);
            // The authored exact box already receives its baseline-offset scale below.
            // Consume its printed height independently; scaling that baseline box again
            // would displace the exact paragraph while fixing its following text.
            double lineHeight = scalesExactBodyAdvance ? baselineLineHeight * paragraphSpacingScale : baselineLineHeight;
            double ParagraphFrameBottom()
            {
                double bottom = CurrentFrameBottom();
                if (!scalesExactBodyAdvance || !hasSimpleScaledBodyFrame)
                {
                    return bottom;
                }

                // Advances occupy printed points below the authored top cursor.
                // Reserve only the printed body height, matching Word pagination.
                double printedBottom = page.Height - page.MarginTop -
                    (page.Height - page.MarginTop - page.MarginBottom) * paragraphSpacingScale;
                return Math.Max(bottom, printedBottom);
            }
            if (textMeasurer is not null &&
                HasPageContent() &&
                ShouldKeepParagraphBlockTogether(paragraph) &&
                cursorY - EstimateKeptParagraphBlock(document.BodyElements, elementIndex, width, textMeasurer, defaultTabStopPoints, pages.Count + 1, paragraphSpacingScale, scaleBaselineOffsetTransitions && page.ColumnFrames.Count <= 1).Height <= Math.Max(ParagraphFrameBottom(), page.MarginBottom + KeepChainFootnoteReserve(elementIndex)))
            {
                AdvanceColumnOrPage();
                RegisterInFlightFootnotesForSourceBlock(elementIndex);
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
                if (ShouldMoveParagraphForWidowControl(paragraph, lines.Length, cursorY, lineHeight, ParagraphFrameBottom(), HasCurrentColumnContent()))
                {
                    AdvanceColumnOrPage();
                    RegisterInFlightFootnotesForSourceBlock(elementIndex);
                }

                // RV05: ordered inline atoms (body path). Affined images in text-mixed
                // paragraphs attach to wrapped lines at run position; wrapping is untouched.
                midLinePlan = CreateMidLinePlan(paragraph, textSpans, lines, paragraphWidth, continuationParagraphWidth, DocxLineMetrics.ResolveBodyBaselineOffset(paragraphFontSize, baselineLineHeight, IsExactLineSpacing(effective)), lineHeight);
                double? typographicBaselineInset = lineHeightProfile.Source == DocxLineHeightSource.BodySingleLineAuto
                    ? DocxLineMetrics.ResolveUniformBodyTypographicBaselineInset(paragraph, paragraphFontSize, textMeasurer) : null;
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

                    if (cursorY - extraAbove - lineHeight < ParagraphFrameBottom() && (HasCurrentColumnContent() || FootnoteReserveYieldsPageToDrain(extraAbove + lineHeight)))
                    {
                        AdvanceForOverflowingItem(extraAbove + lineHeight, elementIndex);
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
                    double? bodyHheaAscender = DocxLineMetrics.ResolveHheaAscenderPoints(paragraph, paragraphFontSize, textMeasurer);
                    double? bodyTierAMax = DocxLineMetrics.ResolveBodyTierAMaxPoints(paragraph, textMeasurer);
                    double baselineOffset = DocxLineMetrics.ResolveBodyBaselineOffset(paragraphFontSize, baselineLineHeight, IsExactLineSpacing(effective), bodyHheaAscender, bodyTierAMax);
                    baselineOffset = typographicBaselineInset ?? baselineOffset;
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
                    double rawBaselineOffset = baselineOffset;
                    // An exact paragraph starting a printed continuation page
                    // retains the document's body origin; choosing its tall box
                    // as a new anchor would leave the first line too low.
                    firstBodyLineBaselineOffset ??= hasSimpleScaledBodyFrame && scalesExactBodyAdvance
                        ? firstDocumentBodyBaselineOffset ?? rawBaselineOffset : rawBaselineOffset;
                    firstDocumentBodyBaselineOffset ??= rawBaselineOffset;
                    if (scaleBaselineOffsetTransitions && firstBodyLineBaselineOffset is { } firstOffset)
                    {
                        // RV06 pagination probe (edge-mixed15, Word 16.0): under a
                        // word-compatible print scale, Office scales whole pitches
                        // including the font-size offset transition, while the
                        // renderer leaves transitions unscaled (12->15 exceeds by
                        // 0.75, 15->12 undershoots by 0.76 at scale 0.758). The
                        // block-first inset stays unscaled on both sides, so each
                        // line corrects against the block-first offset; per-line
                        // chaining would leak second differences across runs.
                        // Same-size lines self-zero and other modes keep legacy.
                        baselineOffset -= (baselineOffset - firstOffset) * (1d - paragraphSpacingScale);
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
                        pages.Count + 1, midLinePlan?.LineImageWidths[lineIndex] ?? 0d);
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
                            double beforeWidth = MeasureMidLineBeforeWidth(line.Spans, placed.LineCharOffset, paragraph, firstLine, lineIndex == lines.Length - 1, paragraphWidth, paragraphFontSize, textMeasurer, ScaleTabStopPositions(effective.TabStops, paragraphSpacingScale), defaultTabStopPoints * paragraphSpacingScale, pages.Count + 1, midLinePlan?.LineImageWidths[lineIndex] ?? 0d);
                            imageShifts.Add((beforeWidth, placed.Width));
                            // RV05 calibration (Word 16.0 fnimg probe): Office image bottoms sit
                            // at the text baseline, so midline images anchor by the line baseline
                            // instead of hanging a full height below it (cell path precedent).
                            lineImages.Add(new DocxInlineImageLayout(
                                placed.Image,
                                lineX + beforeWidth,
                                textBaselineY,
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
                        Story: DocxStoryId.Body(), EmitsTerminalParagraphMark: false,
                        // RV05 floatbox probe (Word COM reference edge-floatbox):
                        // paragraph-relative boxes anchor to the paragraph top, which is
                        // the first baseline plus the first-line inset, not plus the font
                        // size. Only body first lines carry it; other paths keep legacy
                        // bounds until separately probed.
                        FirstLineInsetPoints: firstLine ? (double?)baselineOffset : null)
                    {
                        BodyColumnOriginOffsetX = x - lineShape.X,
                        UsesUniformBodyTypographicBaseline = typographicBaselineInset is not null,
                        BodyLineBoxBaselineInsetPoints = lineHeightProfile.Source is DocxLineHeightSource.ExactLineSpacing or DocxLineHeightSource.BodySingleLineAuto
                            ? rawBaselineOffset * paragraphSpacingScale : null,
                        BodyLineBoxHeightPoints = lineHeightProfile.Source switch
                        {
                            DocxLineHeightSource.ExactLineSpacing => baselineLineHeight * paragraphSpacingScale,
                            DocxLineHeightSource.BodySingleLineAuto => lineHeight,
                            _ => null
                        }
                    });
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
                // Spill sizing probes (Word COM references edge-spillsrc-break14/docdef16):
                // row-end and spill spaces resolve at pilcrow size through the style cascade
                // (a 9pt marker carries 12pt spills; the tabbed second spill space follows
                // the break pilcrow at 14), not at direct run size. Null pilcrows keep the
                // legacy run sizes (synthetic paragraphs).
                double markerPilcrowSize = paragraph.ParagraphMarkFontSize ?? paragraphFontSize;
                double markerSpaceWidth = textMeasurer.MeasureText(firstRun, " ", markerPilcrowSize);
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
                        markerSpaceWidth,
                        markerPilcrowSize,
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
                        Width = breakLastLine.Width + markerSpaceWidth,
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
                    double breakPilcrowSize = nextBreakParagraph.ParagraphMarkFontSize ?? GetParagraphFontSize(nextBreakParagraph);
                    double breakSpaceWidth = textMeasurer.MeasureText(firstRun, " ", breakPilcrowSize);
                    double breakSpillBaseline = DocxLineMetrics.ResolveBodyBaselineOffset(breakPilcrowSize, lineHeight, IsExactLineSpacing(effective));
                    if (HasNoSpacingElement(effective) && Math.Abs(breakPilcrowSize - 11d) < 0.000000001d)
                    {
                        breakSpillBaseline += UntokenedParagraphBaselineExtraPoints;
                    }

                    // Fragment-threshold case (Word COM reference on
                    // docx-ladder-03-table-row-fragment-threshold): markerB spill paints
                    // below the bottom margin on our p2 while Office carries it to the
                    // p3 fragment top. An overflowing spill turns the page like any
                    // other overflowing item instead of painting below the margin; fit
                    // accounting stays zero-consumption so the following break still
                    // fits on the fresh page.
                    bool breakSpillTurnedPage = cursorY - breakSpillAfterSpacing - lineHeight < ParagraphFrameBottom() && (HasCurrentColumnContent() || FootnoteReserveYieldsPageToDrain(lineHeight));
                    if (breakSpillTurnedPage)
                    {
                        AdvanceForOverflowingItem(lineHeight, elementIndex);
                    }
                    // Turned spills open the fresh page with a plain first-line inset:
                    // uniform 11.3pt at 12pt across auto/exact probes and Aptos/Arial faces
                    // (face-independent 0.94em: both spill spaces share one Y across faces with
                    // different hhea ascenders, and exact-24 continuations would sit 7.9 deeper).
                    double breakSpillBaselineY = breakSpillTurnedPage
                        ? cursorY - breakPilcrowSize * DocxLineMetrics.WordAutoLineBaselineOffsetEm
                        : cursorY - breakSpillAfterSpacing - breakSpillBaseline;

                    currentItems.Add(new DocxTextLineLayout(
                        "  ",
                        firstRun,
                        breakPilcrowSize,
                        breakSpillX,
                        breakSpillBaselineY,
                        (breakSpillTabX - breakSpillX) + breakSpaceWidth,
                        [
                            new DocxTextSegmentLayout(
                                " ",
                                firstRun,
                                breakSpillX,
                                markerSpaceWidth,
                                markerPilcrowSize,
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
                                breakPilcrowSize,
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
                    // Ladder-03 row-fragment case (Word COM reference): break-adjacent
                    // spacing rows emit without consuming page fit. Consuming a line box
                    // here shrinks the sliver below the following break advance and turns
                    // a phantom empty page, while Office absorbs the spill row (probe B
                    // starts p2 with no phantom). The row is already placed above, so
                    // only the cursor advance is skipped.
                }
            }
            else if (paragraph.Images.Count == 0 && paragraph.InlineTextBoxes.Count == 0)
            {
                if (cursorY - lineHeight < ParagraphFrameBottom() && (HasCurrentColumnContent() || FootnoteReserveYieldsPageToDrain(lineHeight)))
                {
                    AdvanceForOverflowingItem(lineHeight, elementIndex);
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
                if (cursorY - imageHeight < CurrentFrameBottom() && (HasCurrentColumnContent() || FootnoteReserveYieldsPageToDrain(imageHeight)))
                {
                    AdvanceForOverflowingItem(imageHeight, elementIndex);
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
                if (cursorY - EstimateInlineTextBoxHeight(textBox, paragraphSpacingScale) < CurrentFrameBottom() && (HasCurrentColumnContent() || FootnoteReserveYieldsPageToDrain(EstimateInlineTextBoxHeight(textBox, paragraphSpacingScale))))
                {
                    AdvanceForOverflowingItem(EstimateInlineTextBoxHeight(textBox, paragraphSpacingScale), elementIndex);
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

        DrainRemainingInFlightFootnotes();

        DocxStaticStoryLayoutResult? GetContinuationHeader(DocxLayoutPage templatePage, int pageNumber, int pageCount)
        {
            var headerKey = (templatePage.PageSettings, Math.Max(1d, templatePage.Width - templatePage.MarginLeft - templatePage.MarginRight), pageNumber, pageCount);
            if (!continuationHeaderByPage.TryGetValue(headerKey, out DocxStaticStoryLayoutResult? headerLayout))
            {
                headerLayout = LayoutContinuationHeader(templatePage, pageNumber, pageCount, textMeasurer, defaultTabStopPoints, paragraphSpacingScale, unscaledTextMeasurer, cancellationToken);
                continuationHeaderByPage[headerKey] = headerLayout;
            }

            return headerLayout;
        }

        static bool ContinuationHeaderHasVisibleContent(DocxStaticStoryLayoutResult headerLayout)
        {
            return headerLayout.TextLines.Count != 0 ||
                headerLayout.InlineImages.Count != 0 ||
                headerLayout.TableRows.Count != 0 ||
                headerLayout.InlineTextBoxes.Count != 0;
        }

        double ResolveHeaderKeepOut(DocxLayoutPage templatePage, int pageNumber, int pageCount)
        {
            DocxStaticStoryLayoutResult? headerLayout = GetContinuationHeader(templatePage, pageNumber, pageCount);
            if (headerLayout is null || !ContinuationHeaderHasVisibleContent(headerLayout))
            {
                return double.PositiveInfinity;
            }

            double keepOutLayout = headerLayout.EndCursorY - headerLayout.EndPendingAfterSpacing;
            // The take loop consumes design line boxes while static layout stacks scaled
            // advances, so the keep-out crosses into take space through the same maps as
            // emission: static lines shift by the first-pin offset while footnote stories
            // map uniformly about the page center. The page height cancels, leaving the
            // first-pin baseline and the print scale (identity at unit scale).
            if (!bodyFirstBaselineComputed && pages.Count != 0)
            {
                bodyFirstBaselineComputed = true;
                foreach (DocxTextLineLayout bodyLine in DocxRenderer.EnumerateBodyTextLines(pages[0]))
                {
                    bodyFirstBaselineMemo = bodyFirstBaselineMemo is null ? bodyLine.BaselineY : Math.Max(bodyFirstBaselineMemo.Value, bodyLine.BaselineY);
                }
            }

            double? firstPinBaseline = bodyFirstBaselineMemo;
            if (pages.Count != 0)
            {
                DocxStaticStoryLayoutResult? firstHeaderLayout = GetContinuationHeader(pages[0], 1, pageCount);
                if (firstHeaderLayout is not null && firstHeaderLayout.TextLines.Count != 0)
                {
                    double firstHeaderBaseline = firstHeaderLayout.TextLines.Max(staticLine => staticLine.BaselineY);
                    firstPinBaseline = firstPinBaseline is null ? firstHeaderBaseline : Math.Max(firstPinBaseline.Value, firstHeaderBaseline);
                }
            }

            if (firstPinBaseline is null)
            {
                return keepOutLayout;
            }

            double takeScale = paragraphSpacingScale;
            return (keepOutLayout - firstPinBaseline.Value * (1d - takeScale)) / takeScale;
        }
        DocxLayoutPage[] pagesWithRelatedStories = AddPlacedRelatedStories(document, pages, GetRelatedStoryLayouts, cancellationToken, paragraphSpacingScale, unscaledTextMeasurer ?? textMeasurer, ResolveHeaderKeepOut, placedStoryKeys).ToArray();
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

    private static double ResolveTableCellFirstBaselineInset(IReadOnlyList<DocxParagraph> paragraphs, IDocxTextMeasurer? measurer = null)
    {
        return DocxLineMetrics.ResolveTableCellFirstBaselineInset(paragraphs, measurer);
    }

}
