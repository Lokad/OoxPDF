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
    // Parsed pictures own their source position and use authored paragraph
    // spacing. Unpositioned legacy cell objects retain their block fallback.
    private static double ResolveCellImageBlockAdvance(DocxInlineImage image, double height) =>
        height + (image.SourceRunIndex < 0 ? InlineImageParagraphGapPoints : 0d);

    // R16: small immutable cell-layout environment shared by the cell measure and
    // layout passes. Metrics (width-independent scales), dynamic field state (page
    // inputs), and shared services (measurer, memo) travel as one value instead of a
    // long positional tail where a swapped width/scale silently mislays out; per-call
    // geometry (origins, paddings, memo-independent widths) stays positional because it
    // differs per pass. Construction uses named arguments so same-typed page inputs
    // cannot swap.
    internal readonly record struct DocxTableCellLayoutContext(
        IDocxTextMeasurer? TextMeasurer,
        double DefaultTabStopPoints,
        double ParagraphSpacingScale,
        DocxPageNumber? PageNumber,
        int? PageCount,
        DocxTableCellTextLinesMemo? CellMemo,
        IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph>? MathEquations = null,
        CancellationToken CancellationToken = default);

    // Word hangs the full top-border width above cell content (border probes 2026-09-06: top-term deltas 0.96/1.08/0.58 vs widths 1.0/1.0/0.5); horizontal edges keep the half-width content inset.
    // W6-a1: fixed table geometry joins scaled space; fixedScale reuses the layout
    // spacing scale (identical in every mode).
    private static double ResolveTableCellTopBorderContentInset(DocxTableCell cell, double fixedScale)
    {
        return DocxTableBorderGeometry.ResolveVisibleWidth(DocxTableBorderGeometry.Find(cell.Borders, "top")) * fixedScale;
    }

    // Office A/B (w6-celltucksize probes at two lane-fit scales plus w6-celltuckspacing
    // spacing/box-order probes, Word-COM rendered): a box-only cell paragraph after a
    // text paragraph tucks its top below the last text baseline by a font-size affine
    // law, honoring only spacing excess over the 8pt default. Same-paragraph boxes keep
    // the legacy block rule, as do boxes after empty, image-bearing, or nested content.
    private const double CellInlineTextBoxTuckSlope = 0.470d;
    private const double CellInlineTextBoxTuckInterceptPoints = 3.65d;
    private const double CellInlineTextBoxTuckDefaultAfterPoints = 8d;

    private static bool HasVisibleTextSpans(IReadOnlyList<DocxTextSpan> textSpans)
    {
        foreach (DocxTextSpan span in textSpans)
        {
            if (!string.IsNullOrWhiteSpace(span.Text))
            {
                return true;
            }
        }

        return false;
    }

    private static double ResolveCellInlineTextBoxTuckDistance(
        double previousFontSize,
        double spacingBefore,
        double paragraphSpacingScale)
    {
        double tuck = paragraphSpacingScale * (CellInlineTextBoxTuckSlope * previousFontSize + CellInlineTextBoxTuckInterceptPoints);
        double excess = Math.Max(0d, spacingBefore - CellInlineTextBoxTuckDefaultAfterPoints * paragraphSpacingScale);
        return tuck + excess;
    }

    private readonly record struct DocxTableCellContentHeight(double FullHeight, double MinimumPictureOverlap)
    {
        public double TrailingAfterSpacing { get; init; }
        public bool ContainsMath { get; init; }
        public double Resolve(bool minimumPicturesOwned) => minimumPicturesOwned &&
            double.IsFinite(FullHeight) && double.IsFinite(MinimumPictureOverlap) &&
            MinimumPictureOverlap > 0d && MinimumPictureOverlap <= FullHeight
                ? FullHeight - MinimumPictureOverlap : FullHeight;
    }

    private static DocxTableCellContentHeight MeasureTableCellContentHeight(
        DocxTableCell cell,
        double cellWidth,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        double? rowTopPadding,
        DocxPageNumber? pageNumber,
        int? pageCount,
        double paragraphSpacingScale,
        IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph>? mathEquations = null,
        CancellationToken cancellationToken = default)
    {
        var measureContext = new DocxTableCellLayoutContext(
            TextMeasurer: textMeasurer,
            DefaultTabStopPoints: defaultTabStopPoints,
            ParagraphSpacingScale: paragraphSpacingScale,
            PageNumber: pageNumber,
            PageCount: pageCount,
            CellMemo: null,
            MathEquations: mathEquations,
            CancellationToken: cancellationToken);
        IReadOnlyList<DocxBodyElement> bodyElements = DocxTableCellContent.GetBodyElements(cell);
        if (bodyElements.Count == 0 ||
            (textMeasurer is null && !bodyElements.Any(element => element is DocxTableElement
                or DocxParagraphElement { Paragraph.Images.Count: > 0 }
                or DocxParagraphElement { Paragraph.InlineTextBoxes.Count: > 0 })))
        {
            return default;
        }

        IReadOnlyList<DocxParagraph> paragraphs = DocxBlockTraversal.EnumerateDirectParagraphs(bodyElements).ToArray();

        double paddingLeft = ResolveTableCellHorizontalEdgeInset(cell, "left", cell.Margins.LeftPoints, paragraphSpacingScale);
        double paddingRight = ResolveTableCellHorizontalEdgeInset(cell, "right", cell.Margins.RightPoints, paragraphSpacingScale);
        double paddingTop = rowTopPadding ?? ResolveTableCellVerticalPadding(cell.Margins.TopPoints, paragraphSpacingScale);
        double paddingBottom = ResolveTableCellVerticalPadding(cell.Margins.BottomPoints, paragraphSpacingScale);
        double textWidth = Math.Max(1d, cellWidth - paddingLeft - paddingRight);
        double contentHeight = paddingTop + paddingBottom;
        double minimumPictureOverlap = 0d;
        bool containsMath = false;
        double pendingSpacingAfter = 0d;
        DocxParagraph? previousParagraph = null;
        bool previousParagraphHasText = false;
        bool previousParagraphHasTextLines = false;
        double previousTextFontSize = 0d;
        double previousTextLineHeight = 0d;
        for (int elementIndex = 0; elementIndex < bodyElements.Count; elementIndex++)
        {
            DocxBodyElement bodyElement = bodyElements[elementIndex];
            if (bodyElement is DocxTableElement tableElement)
            {
                contentHeight += pendingSpacingAfter;
                pendingSpacingAfter = 0d;
                previousParagraph = null;
                previousParagraphHasText = false;
                previousParagraphHasTextLines = false;
                contentHeight += MeasureNestedTableHeight(tableElement.Table, textWidth, measureContext);
                continue;
            }

            if (bodyElement is DocxPageBreakElement estimateBreak &&
                estimateBreak.SourceKind == DocxBreakSourceKind.RunBreak)
            {
                // RV06 cellbreak probe: the spill row occupies one laid-out line, so the
                // estimate must reserve it exactly like layout does. The text walk emits
                // a line for every paragraph with text spans (even all-empty ones
                // materialize as a space line that the spill search finds), so any
                // spans-carrying previous paragraph counts, including mixed paragraphs
                // with block images and empty paragraphs. Spacing flows as for an empty
                // default paragraph: pending is consumed, then the default after applies.
                contentHeight += Math.Max(0d, pendingSpacingAfter);
                pendingSpacingAfter = 0d;
                if (previousParagraphHasTextLines)
                {
                    contentHeight += previousTextLineHeight;
                }

                pendingSpacingAfter = DocxDefaults.DefaultParagraphAfterSpacingPoints * paragraphSpacingScale;
                previousParagraph = null;
                previousParagraphHasText = false;
                previousParagraphHasTextLines = false;
                continue;
            }

            if (bodyElement is not DocxParagraphElement paragraphElement)
            {
                continue;
            }

            DocxParagraph paragraph = paragraphElement.Paragraph;
            cancellationToken.ThrowIfCancellationRequested();
            if (DocxTableCellContent.IsTerminalNestedTableMark(bodyElements, elementIndex, cancellationToken))
                continue;
            DocxParagraphSpacingProfile spacingProfile = ResolveParagraphSpacingProfile(previousParagraph, paragraph, pendingSpacingAfter, paragraphSpacingScale);
            contentHeight += spacingProfile.AppliedBeforeSpacing;
            pendingSpacingAfter = 0d;
            IReadOnlyList<DocxTextSpan> textSpans = [];
            DocxMidLinePlan? estimatePlan = null;
            // The cell cursor still places owned drawings when no text can be measured.
            if (textMeasurer is not null)
            {
                if (TryGetStandaloneMath(paragraph, textWidth, paragraphSpacingScale, mathEquations, out var math))
                {
                    contentHeight += math.LineHeight * paragraphSpacingScale;
                    containsMath = true;
                    pendingSpacingAfter = spacingProfile.ParagraphAfterSpacing;
                    previousParagraph = paragraph;
                    previousParagraphHasText = false;
                    previousParagraphHasTextLines = false;
                    continue;
                }
                double fontSize = GetParagraphFontSize(paragraph);
                DocxLineHeightProfile lineHeightProfile = ResolveLineHeightProfile(paragraph, fontSize, textMeasurer);
                double lineHeight = lineHeightProfile.LineHeight;
                if (TryMeasureOrderedCellContent(cell, paragraph, paragraphs, measureContext, textWidth, fontSize, lineHeight,
                    out var orderedLines))
                {
                    contentHeight += orderedLines.Sum(line => line.OccupiedHeight);
                    containsMath |= orderedLines.Any(line => line.Fragments.Any(fragment => fragment.Math is not null));
                    pendingSpacingAfter = spacingProfile.ParagraphAfterSpacing;
                    previousParagraph = paragraph;
                    previousParagraphHasText = true;
                    previousParagraphHasTextLines = true;
                    previousTextLineHeight = orderedLines[^1].OccupiedHeight;
                    continue;
                }
                textSpans = CreateTextSpans(paragraph.Runs, pageNumber, pageCount);
                if (textSpans.Count != 0)
                {
                    double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, fontSize, textMeasurer, paragraphSpacingScale, defaultTabStopPoints);
                    double firstParagraphWidth = ResolveTableCellTextWrapWidth(cell, textWidth - textStartOffset - GetParagraphRightInset(paragraph, paragraphSpacingScale));
                    double continuationParagraphWidth = ResolveTableCellTextWrapWidth(cell, textWidth - GetParagraphTextStartOffset(paragraph, paragraphSpacingScale) - GetParagraphRightInset(paragraph, paragraphSpacingScale));
                    DocxWrappedTextLine[] estimateLines = WrapTextLines(textSpans, firstParagraphWidth, continuationParagraphWidth, fontSize, textMeasurer, ScaleTabStopPositions(paragraph.EffectiveProperties.TabStops, paragraphSpacingScale), defaultTabStopPoints * paragraphSpacingScale, allowOverwideTokenBreaks: true, dynamicFieldPageNumber: pageNumber, inlineImageWidths: ResolveInlineImageWrapWidths(paragraph, textSpans, paragraphSpacingScale), sourceParagraph: paragraph).ToArray();
                    DocxMidLineMath? inlineMath = WrapTableInlineMath(paragraph, textSpans, ref estimateLines,
                        firstParagraphWidth, continuationParagraphWidth, fontSize, measureContext);
                    containsMath |= inlineMath is not null;
                    int lineCount = estimateLines.Length;
                    lineHeight = QuantizeTableCellWrappedLineHeight(lineHeight, lineCount);
                    bool exactSpacing = IsExactLineSpacing(paragraph.EffectiveProperties);
                    double baselineOffset = DocxLineMetrics.ResolveTableCellBodyBaselineOffset(fontSize, lineHeight,
                        exactSpacing, paragraphs, paragraph, textMeasurer);
                    estimatePlan = CreateMidLinePlan(paragraph, textSpans, estimateLines, firstParagraphWidth,
                        continuationParagraphWidth, baselineOffset, paragraphSpacingScale, textMeasurer);
                    previousTextFontSize = fontSize;
                    previousTextLineHeight = lineHeight;
                    contentHeight += ResolveListLabelFirstLineExtraLeading(paragraph, fontSize, textMeasurer);
                    double minimumExcess = 0d;
                    if (lineHeightProfile.Source == DocxLineHeightSource.AtLeastLineSpacing &&
                        estimatePlan is not null && estimatePlan.PlacedMask.All(placed => placed) &&
                        textMeasurer is IDocxLineMetricsProvider metrics && paragraph.Runs.Count > 0)
                    {
                        double natural = metrics.MeasureSingleLineHeight(paragraph.Runs[0], fontSize);
                        if (double.IsFinite(natural) && natural > 0d)
                            minimumExcess = Math.Max(0d, lineHeight - natural);
                    }
                    for (int estimateLineIndex = 0; estimateLineIndex < estimateLines.Length; estimateLineIndex++)
                    {
                        double estimateShift = exactSpacing ? 0d : (estimatePlan?.ShiftAboveHeights[estimateLineIndex] ?? 0d);
                        var expansion = inlineMath?.LineIndex == estimateLineIndex
                            ? ResolveTableInlineMathExpansion(inlineMath.Math, fontSize, lineHeight, exactSpacing,
                                ResolveTableCellFirstBaselineInset(paragraphs, textMeasurer), baselineOffset, paragraphSpacingScale)
                            : (Above: 0d, Below: 0d);
                        contentHeight += lineHeight + Math.Max(estimateShift, expansion.Above) +
                            Math.Max(expansion.Below, exactSpacing ? 0d : estimatePlan?.ExtraBelow(estimateLineIndex) ?? 0d);
                        minimumPictureOverlap += ResolveMinimumPictureOverlap(minimumExcess, estimateShift);
                    }
                }
                else if (paragraph.Images.Count == 0 && paragraph.InlineTextBoxes.Count == 0)
                {
                    contentHeight += ResolveListLabelFirstLineExtraLeading(paragraph, fontSize, textMeasurer);
                    contentHeight += lineHeight;
                }
            }

            for (int estimateImageIndex = 0; estimateImageIndex < paragraph.Images.Count; estimateImageIndex++)
            {
                if (estimatePlan?.PlacedMask[estimateImageIndex] == true)
                {
                    continue;
                }

                DocxInlineImage image = paragraph.Images[estimateImageIndex];
                var (imageWidth, imageHeight) = ResolveInlineImageSize(image, textWidth, paragraphSpacingScale);

                contentHeight += ResolveCellImageBlockAdvance(image, imageHeight);
            }

            bool currentHasVisibleText = HasVisibleTextSpans(textSpans);
            bool tuckApplied = false;
            foreach (DocxInlineTextBox textBox in paragraph.InlineTextBoxes)
            {
                double boxParagraphWidth = Math.Max(1d, textWidth - GetParagraphStartOffset(paragraph, paragraphSpacingScale) - GetParagraphRightInset(paragraph, paragraphSpacingScale));
                DocxInlineTextBoxLayout? measuredBox = CreateInlineTextBoxLayout(
                    textBox, sourceBlockIndex: null, 0d, boxParagraphWidth, 0d,
                    paragraph.EffectiveProperties.Alignment, textMeasurer, defaultTabStopPoints,
                    paragraphSpacingScale, pageNumber ?? 1, cancellationToken,
                    mathEquations: mathEquations, pageCount: pageCount);
                if (measuredBox is null)
                {
                    continue;
                }
                double boxContribution = measuredBox.BoxHeight;
                if (!tuckApplied &&
                    previousParagraphHasText &&
                    paragraph.Images.Count == 0 &&
                    !currentHasVisibleText)
                {
                    double tuckDistance = ResolveCellInlineTextBoxTuckDistance(previousTextFontSize,
                        spacingProfile.AppliedBeforeSpacing, paragraphSpacingScale);
                    boxContribution += tuckDistance - (previousTextLineHeight + spacingProfile.AppliedBeforeSpacing);
                    tuckApplied = true;
                }

                contentHeight += boxContribution;
            }

            pendingSpacingAfter = spacingProfile.ParagraphAfterSpacing;
            previousParagraph = paragraph;
            previousParagraphHasText = currentHasVisibleText && paragraph.Images.Count == 0 && paragraph.InlineTextBoxes.Count == 0;
            previousParagraphHasTextLines = textSpans.Count != 0;
        }

        contentHeight += pendingSpacingAfter;
        return new(contentHeight, minimumPictureOverlap)
        { TrailingAfterSpacing = Math.Max(0d, pendingSpacingAfter), ContainsMath = containsMath };
    }

    private static double MeasureNestedTableHeight(
        DocxTable table,
        double availableWidth,
        DocxTableCellLayoutContext context)
    {
        DocxTableLayoutFrame frame = CreateTableLayoutFrame(
            table,
            tableIndex: -1,
            sourceBlockIndex: -1,
            x: 0d,
            availableWidth: availableWidth,
            pageContentHeight: double.MaxValue / 4d,
            context.TextMeasurer,
            context.DefaultTabStopPoints,
            pageNumber: context.PageNumber,
            pageCount: context.PageCount,
            paragraphSpacingScale: context.ParagraphSpacingScale,
            mathEquations: context.MathEquations,
            cancellationToken: context.CancellationToken,
            anchorAuthoredCellText: false);
        return frame.RowHeights.Sum();
    }

    internal static (IReadOnlyList<DocxTextLineLayout> Lines, IReadOnlyList<DocxInlineImageLayout> PlacedImages,
        IReadOnlyList<DocxInlineTextBoxLayout> PlacedBoxes, IReadOnlyList<DocxTableRowLayout> NestedRows,
        double UsedHeight, IReadOnlyList<DocxPlacedDrawingAnchor> DrawingAnchors) LayoutTableCellContent(
        DocxTableCell cell,
        double cellX,
        double cellY,
        double cellWidth,
        double cellHeight,
        double rowTopPadding,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        DocxPageNumber? pageNumber,
        int? pageCount,
        double paragraphSpacingScale,
        DocxTableCellTextLinesMemo? cellMemo = null,
        int pageIndex = 0,
        IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph>? mathEquations = null,
        CancellationToken cancellationToken = default)
    {
        var context = new DocxTableCellLayoutContext(
            TextMeasurer: textMeasurer,
            DefaultTabStopPoints: defaultTabStopPoints,
            ParagraphSpacingScale: paragraphSpacingScale,
            PageNumber: pageNumber,
            PageCount: pageCount,
            CellMemo: cellMemo,
            MathEquations: mathEquations,
            CancellationToken: cancellationToken);
        IReadOnlyList<DocxBodyElement> bodyElements = DocxTableCellContent.GetBodyElements(cell);
        IReadOnlyList<DocxParagraph> paragraphs = DocxBlockTraversal.EnumerateDirectParagraphs(bodyElements).ToArray();
        if (bodyElements.Count == 0)
        {
            return ([], [], [], [], 0d, []);
        }

        double paddingLeft = ResolveTableCellHorizontalEdgeInset(cell, "left", cell.Margins.LeftPoints, paragraphSpacingScale);
        double paddingRight = ResolveTableCellHorizontalEdgeInset(cell, "right", cell.Margins.RightPoints, paragraphSpacingScale);
        double paddingTop = rowTopPadding + ResolveTableCellTopBorderContentInset(cell, paragraphSpacingScale);
        double paddingBottom = ResolveTableCellVerticalPadding(cell.Margins.BottomPoints, paragraphSpacingScale);
        double baselineInset = ResolveTableCellFirstBaselineInset(paragraphs, textMeasurer);
        double textWidth = Math.Max(1d, cellWidth - paddingLeft - paddingRight);
        double startBaselineY = cellY + cellHeight - baselineInset - paddingTop;
        var dependencies = DocxFieldText.GetDependencies(bodyElements, context.CancellationToken);
        var (lines, usedHeight, placedImages, placedBoxes, nestedRows, anchors) = ComputeTableCellContent(cell, cellX, cellWidth, cellHeight, context, bodyElements, paragraphs, paddingLeft, textWidth, startBaselineY, rowTopPadding, cellX, cellY + cellHeight, dependencies, pageIndex);
        double extra = ResolveTableCellExtraVerticalSpace(paragraphs, cellHeight, paddingTop, paddingBottom, usedHeight, baselineInset);
        double verticalOffset = ResolveVerticalContentOffset(cell, extra);
        return verticalOffset == 0d
            ? (lines, placedImages, placedBoxes, nestedRows, usedHeight, anchors)
            : (ShiftTextLines(lines, -verticalOffset, 0d), ShiftInlineImages(placedImages, -verticalOffset, 0d),
                placedBoxes.Select(box => ShiftInlineTextBox(box, -verticalOffset, 0d)).ToArray(),
                ShiftTableRows(nestedRows, -verticalOffset, 0d), usedHeight, DocxPlacedDrawingAnchor.TranslateY(anchors, -verticalOffset));
    }

    private static (IReadOnlyList<DocxTextLineLayout> Lines, double UsedHeight, IReadOnlyList<DocxInlineImageLayout> PlacedImages,
        IReadOnlyList<DocxInlineTextBoxLayout> PlacedBoxes, IReadOnlyList<DocxTableRowLayout> NestedRows,
        IReadOnlyList<DocxPlacedDrawingAnchor> DrawingAnchors) ComputeTableCellContent(
        DocxTableCell cell,
        double cellX,
        double cellWidth,
        double cellHeight,
        DocxTableCellLayoutContext context,
        IReadOnlyList<DocxBodyElement> bodyElements,
        IReadOnlyList<DocxParagraph> paragraphs,
        double paddingLeft,
        double textWidth,
        double startBaselineY,
        double rowTopPadding,
        double originX,
        double originY,
        DocxCounterDependencies dependencies,
        int pageIndex)
    {
        context.CancellationToken.ThrowIfCancellationRequested();
        DocxTableCellTextLinesMemo? cellMemo = context.CellMemo;
        // Geometry-only callers have no text measurer, but authored picture
        // and box dimensions still participate in the same cell cursor.
        IDocxTextMeasurer? measurer = context.TextMeasurer;

        // Nested frame admission depends on the available height. Ordinary
        // paragraph measurements remain reusable when only cell height changes.
        double nestedPageHeight = bodyElements.Any(element => element is DocxTableElement) ? cellHeight : 0d;
        List<DocxPlacedDrawingAnchor>? anchors = paragraphs.Any(paragraph => paragraph.DrawingOwners.Count > 0) ? [] : null;
        if (cellMemo is not null &&
            cellMemo.TryGetRelativeLines(cell, cellWidth, measurer, context.DefaultTabStopPoints, rowTopPadding, context.ParagraphSpacingScale, context.PageNumber, context.PageCount, dependencies,
                out var cachedRelative, out var cachedPlaced, out var cachedBoxes, out double cachedUsedHeight,
                out var cachedRows, nestedPageHeight, anchors))
        {
            return (ShiftTextLines(cachedRelative, originY, originX), cachedUsedHeight,
                ShiftInlineImages(cachedPlaced, originY, originX, pageIndex),
                cachedBoxes.Select(box => ShiftInlineTextBox(box, originY, originX, pageIndex)).ToArray(),
                ShiftTableRows(cachedRows, originY, originX, pageIndex), DocxPlacedDrawingAnchor.TranslateY(anchors ?? [], originY));
        }

        double cursorY = startBaselineY;
        var lines = new List<DocxTextLineLayout>();
        var placedImages = new List<DocxInlineImageLayout>();
        var placedBoxes = new List<DocxInlineTextBoxLayout>();
        var nestedRows = new List<DocxTableRowLayout>();
        double pendingSpacingAfter = 0d;
        DocxParagraph? previousParagraph = null;
        bool previousParagraphHasText = false;
        double previousTextBaselineY = 0d;
        double previousTextFontSize = 0d;
        double cellFirstBaselineInset = ResolveTableCellFirstBaselineInset(paragraphs, measurer);
        int paragraphIndex = 0;
        int nestedTableIndex = 0;
        for (int elementIndex = 0; elementIndex < bodyElements.Count; elementIndex++)
        {
            DocxBodyElement bodyElement = bodyElements[elementIndex];
            if (bodyElement is DocxTableElement tableElement)
            {
                cursorY -= pendingSpacingAfter;
                pendingSpacingAfter = 0d;
                previousParagraph = null;
                previousParagraphHasText = false;
                if (measurer is null) continue;
                var frame = CreateTableLayoutFrame(tableElement.Table, nestedTableIndex++, sourceBlockIndex: -1,
                    cellX + paddingLeft, textWidth, cellHeight, measurer, context.DefaultTabStopPoints,
                    context.CancellationToken, context.PageNumber, context.PageCount,
                    context.ParagraphSpacingScale, mathEquations: context.MathEquations, anchorAuthoredCellText: false);
                for (int rowIndex = 0; rowIndex < tableElement.Table.Rows.Count; rowIndex++)
                {
                    context.CancellationToken.ThrowIfCancellationRequested();
                    double rowHeight = frame.RowHeights[rowIndex];
                    double rowTop = cursorY + cellFirstBaselineInset;
                    nestedRows.Add(CreateTableRowLayout(tableElement.Table, frame.Context, tableElement.Table.Rows[rowIndex],
                        rowIndex, frame.RowHeights, frame.EffectiveColumns, frame.Scale, measurer,
                        context.DefaultTabStopPoints, () => pageIndex + 1, rowTop, rowHeight, rowTop,
                        FragmentIndex: 0, FragmentCount: 1, FragmentReason: "None", Story: null,
                        pageCount: context.PageCount, paragraphSpacingScale: context.ParagraphSpacingScale,
                        cellMemo: cellMemo, fieldPageNumber: context.PageNumber, previousRow: nestedRows.LastOrDefault()));
                    cursorY -= rowHeight;
                }
                continue;
            }

            if (bodyElement is DocxPageBreakElement cellPageBreak &&
                cellPageBreak.SourceKind == DocxBreakSourceKind.RunBreak)
            {
                // RV06 cellbreak probe: explicit page breaks inside table cells do not
                // turn pages in Office; one spill space follows the previous line on the
                // same page and content flows on. The break behaves as an empty paragraph
                // with document defaults (edge-t0/t1: 25pt pitches = line plus 8pt after),
                // so pending after-spacing is consumed, not reset.
                cursorY -= Math.Max(0d, pendingSpacingAfter);
                pendingSpacingAfter = 0d;
                if (previousParagraph is not null && measurer is not null)
                {
                    for (int cellLineIndex = lines.Count - 1; cellLineIndex >= 0; cellLineIndex--)
                    {
                        if (lines[cellLineIndex] is not DocxTextLineLayout cellLastLine ||
                            !ReferenceEquals(cellLastLine.SourceParagraph, previousParagraph) ||
                            cellLastLine.Text.Length == 0)
                        {
                            continue;
                        }

                        double spillFontSize = GetParagraphFontSize(previousParagraph);
                        DocxLineHeightProfile spillProfile = ResolveLineHeightProfile(previousParagraph, spillFontSize, measurer);
                        double spillLineHeight = cellLastLine.LineHeight ?? spillProfile.LineHeight;
                        double spillX = cellX + paddingLeft + GetParagraphTextStartOffset(previousParagraph, context.ParagraphSpacingScale);
                        double spillWidth = measurer.MeasureText(cellLastLine.StyleRun, " ", cellLastLine.FontSize);
                        lines.Add(new DocxTextLineLayout(
                            " ",
                            cellLastLine.StyleRun,
                            cellLastLine.FontSize,
                            spillX,
                            cursorY,
                            spillWidth,
                            [
                                new DocxTextSegmentLayout(
                                    " ",
                                    cellLastLine.StyleRun,
                                    spillX,
                                    spillWidth,
                                    cellLastLine.FontSize,
                                    0d,
                                    0d,
                                    DocxTextStateCharacterSpacingSource.None,
                                    true,
                                    -1,
                                    0,
                                    DocxTextSegmentRole.BreakSpill),
                            ],
                            SourceBlockIndex: null,
                            SourceParagraphIndex: cellLastLine.SourceParagraphIndex,
                            SourceLineIndex: lines.Count,
                            Story: DocxStoryId.TableCell(),
                            LineHeight: spillLineHeight,
                            AppliedBeforeSpacing: 0d,
                            IsFirstParagraphLine: false,
                            EndsWithIntraTokenBreak: false,
                            SingleLineHeight: spillProfile.SingleLineHeight,
                            ListLabelSingleLineHeight: spillProfile.ListLabelSingleLineHeight,
                            BodyWindowsLineHeight: spillProfile.BodyWindowsLineHeight,
                            ListLabelWindowsLineHeight: spillProfile.ListLabelWindowsLineHeight,
                            EffectiveLineSpacingFactor: spillProfile.EffectiveLineSpacingFactor,
                            LineSpacingFactorFloorApplied: spillProfile.LineSpacingFactorFloorApplied,
                            LineHeightSource: spillProfile.Source,
                            PendingAfterSpacing: null,
                            ParagraphBeforeSpacing: null,
                            ParagraphAfterSpacing: null,
                            ContextualSpacingSuppressed: null,
                            SourceParagraph: previousParagraph,
                            EmitsTerminalParagraphMark: false));
                        cursorY -= spillLineHeight;
                        break;
                    }
                }

                pendingSpacingAfter = DocxDefaults.DefaultParagraphAfterSpacingPoints * context.ParagraphSpacingScale;
                previousParagraph = null;
                previousParagraphHasText = false;
                continue;
            }

            if (bodyElement is not DocxParagraphElement paragraphElement)
            {
                continue;
            }

            DocxParagraph paragraph = paragraphElement.Paragraph;
            context.CancellationToken.ThrowIfCancellationRequested();
            bool terminalMark = DocxTableCellContent.IsTerminalNestedTableMark(bodyElements, elementIndex, context.CancellationToken);
            double paragraphCursorY = cursorY;
            int paragraphLineStart = lines.Count;
            DocxParagraphSpacingProfile spacingProfile = ResolveParagraphSpacingProfile(previousParagraph, paragraph, pendingSpacingAfter, context.ParagraphSpacingScale);
            cursorY -= spacingProfile.AppliedBeforeSpacing;
            pendingSpacingAfter = 0d;
            double paragraphTop = cursorY + cellFirstBaselineInset;
            IReadOnlyList<DocxTextSpan> textSpans = [];
            DocxMidLinePlan? midLinePlan = null;
            if (measurer is not null)
            {
                if (TryGetStandaloneMath(paragraph, textWidth, context.ParagraphSpacingScale, context.MathEquations, out var math))
                {
                    // Preserve a nominal paragraph fallback when the model has no
                    // recorded cell anchor; parsed empty runs normally carry it.
                    double mathBaselineAnchor = cellFirstBaselineInset > 0d ? cellFirstBaselineInset
                        : DocxLineMetrics.ResolveBodyBaselineOffset(math.Equation.ParentFontSize ?? GetParagraphFontSize(paragraph),
                            math.LineHeight, false);
                    lines.Add(CreateStandaloneMathLine(paragraph, math, cellX + paddingLeft, textWidth,
                        cursorY + cellFirstBaselineInset, context.ParagraphSpacingScale, paragraphIndex, spacingProfile, HasFollowingParagraphShading(bodyElements, elementIndex),
                        ResolveParagraphShadingBeforeSpacing(previousParagraph, paragraph, spacingProfile), mathBaselineAnchor));
                    cursorY -= math.LineHeight * context.ParagraphSpacingScale;
                    RecordDrawingAnchors();
                    pendingSpacingAfter = spacingProfile.ParagraphAfterSpacing;
                    previousParagraph = paragraph;
                    previousParagraphHasText = false;
                    paragraphIndex++;
                    continue;
                }
                double fontSize = GetParagraphFontSize(paragraph);
                DocxLineHeightProfile lineHeightProfile = ResolveLineHeightProfile(paragraph, fontSize, measurer);
                double lineHeight = lineHeightProfile.LineHeight;
                if (TryMeasureOrderedCellContent(cell, paragraph, paragraphs, context, textWidth, fontSize, lineHeight,
                    out var orderedLines))
                {
                    for (int index = 0; index < orderedLines.Count; index++)
                    {
                        context.CancellationToken.ThrowIfCancellationRequested();
                        var line = orderedLines[index];
                        double inset = index == 0
                            ? GetParagraphFirstLineTextStartOffset(paragraph, fontSize, measurer, context.ParagraphSpacingScale, context.DefaultTabStopPoints)
                            : GetParagraphTextStartOffset(paragraph, context.ParagraphSpacingScale);
                        double available = Math.Max(1d, textWidth - inset - GetParagraphRightInset(paragraph, context.ParagraphSpacingScale));
                        double align = ResolveOrderedInlineAlignment(line, paragraph.Alignment, available);
                        var placed = PlaceOrderedInlineLine(paragraph, line, cellX + paddingLeft + inset + align, cellX + paddingLeft,
                            cursorY + cellFirstBaselineInset, fontSize, context.ParagraphSpacingScale, pageIndex,
                            null, paragraphIndex, index, DocxStoryId.TableCell());
                        lines.Add(placed.Line with { SingleLineHeight = lineHeightProfile.SingleLineHeight,
                            ParagraphShadingFrame = CreateParagraphShadingFrame(paragraph, HasFollowingParagraphShading(bodyElements, elementIndex), cellX + paddingLeft, textWidth, placed.Line.X,
                                line.BaselineInset, line.OccupiedHeight, context.ParagraphSpacingScale,
                                index == orderedLines.Count - 1 ? spacingProfile.ParagraphAfterSpacing : 0d,
                                index == 0 ? ResolveParagraphShadingBeforeSpacing(previousParagraph, paragraph, spacingProfile) : 0d),
                            AppliedBeforeSpacing = index == 0 ? spacingProfile.AppliedBeforeSpacing : 0d,
                            ParagraphAfterSpacing = index == orderedLines.Count - 1 ? spacingProfile.ParagraphAfterSpacing : 0d });
                        placedImages.AddRange(placed.Images);
                        placedBoxes.AddRange(placed.TextBoxes);
                        cursorY -= line.OccupiedHeight;
                    }
                    pendingSpacingAfter = spacingProfile.ParagraphAfterSpacing;
                    previousParagraph = paragraph;
                    previousParagraphHasText = paragraph.Runs.Any(run => !string.IsNullOrWhiteSpace(run.Text)) &&
                        paragraph.Images.Count == 0 && paragraph.InlineTextBoxes.Count == 0;
                    previousTextBaselineY = lines[^1].BaselineY;
                    previousTextFontSize = lines[^1].FontSize;
                    paragraphIndex++;
                    RecordDrawingAnchors();
                    continue;
                }
                textSpans = CreateTextSpans(paragraph.Runs, context.PageNumber, context.PageCount);
                if (textSpans.Count == 0)
                {
                    if (paragraph.Images.Count == 0 && paragraph.InlineTextBoxes.Count == 0)
                    {
                        cursorY -= ResolveListLabelFirstLineExtraLeading(paragraph, fontSize, measurer);
                        cursorY -= lineHeight;
                    }
                }
                else
                {
                    DocxTextRun firstRun = paragraph.Runs[0];
                    double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, fontSize, measurer, context.ParagraphSpacingScale, context.DefaultTabStopPoints);
                    double continuationTextStartOffset = GetParagraphTextStartOffset(paragraph, context.ParagraphSpacingScale);
                    double labelStartOffset = GetParagraphLabelStartOffset(paragraph, context.ParagraphSpacingScale);
                    double paragraphX = cellX + paddingLeft + textStartOffset;
                    double paragraphWidth = Math.Max(1d, textWidth - textStartOffset - GetParagraphRightInset(paragraph, context.ParagraphSpacingScale));
                    double continuationParagraphWidth = Math.Max(1d, textWidth - continuationTextStartOffset - GetParagraphRightInset(paragraph, context.ParagraphSpacingScale));
                    var tabStops = ScaleTabStopPositions(paragraph.EffectiveProperties.TabStops, context.ParagraphSpacingScale);
                    double defaultTabStop = context.DefaultTabStopPoints * context.ParagraphSpacingScale;
                    bool firstLine = true;
                    DocxWrappedTextLine[] wrappedLines = WrapTextLines(
                        textSpans,
                        ResolveTableCellTextWrapWidth(cell, paragraphWidth),
                        ResolveTableCellTextWrapWidth(cell, continuationParagraphWidth),
                        fontSize,
                        measurer,
                        tabStops,
                        defaultTabStop,
                        allowOverwideTokenBreaks: true,
                        dynamicFieldPageNumber: context.PageNumber, inlineImageWidths: ResolveInlineImageWrapWidths(paragraph, textSpans, context.ParagraphSpacingScale), sourceParagraph: paragraph).ToArray();
                    DocxMidLineMath? inlineMath = WrapTableInlineMath(paragraph, textSpans, ref wrappedLines,
                        ResolveTableCellTextWrapWidth(cell, paragraphWidth), ResolveTableCellTextWrapWidth(cell, continuationParagraphWidth), fontSize, context);
                    double? leadingDecimalAdvance = paragraph.Images.Count == 0 && paragraph.Equations.Count == 0 &&
                        paragraph.InlineTextBoxes.Count == 0 &&
                        ResolveImplicitCellDecimalStop(cell, paragraph, context.ParagraphSpacingScale, context.CancellationToken) is { } decimalStop
                            ? MeasureImplicitCellDecimalAdvance(wrappedLines, textSpans, decimalStop, paragraphWidth, fontSize, measurer)
                            : null;
                    lineHeight = QuantizeTableCellWrappedLineHeight(lineHeight, wrappedLines.Length);
                    // RV05: ordered inline atoms (table-cell path). Affined images in
                    // text-mixed paragraphs attach to wrapped lines at run position.
                    // RV06 cellmidline size matrix (Tahoma small slash medium slash large images): first-line
                    // totals follow max(hhea-aware inset, image height), so the attach shift rides the
                    // hhea-aware inset like the transition; estimate, image and single-line paths
                    // below use the same form for coherence.
                    bool exactSpacing = IsExactLineSpacing(paragraph.EffectiveProperties);
                    double baselineOffset = DocxLineMetrics.ResolveTableCellBodyBaselineOffset(fontSize, lineHeight,
                        exactSpacing, paragraphs, paragraph, measurer);
                    midLinePlan = CreateMidLinePlan(paragraph, textSpans, wrappedLines,
                        ResolveTableCellTextWrapWidth(cell, paragraphWidth), ResolveTableCellTextWrapWidth(cell, continuationParagraphWidth),
                        baselineOffset, context.ParagraphSpacingScale, measurer);
                    // RV06 pagination probe (edge-mixedcell, Word 16.0): Office stacks
                    // cell baselines with per-line font-size offsets (12->15 carries
                    // +2.82 at scale 1.0), while the renderer stacks pure line heights
                    // after the first-line inset. Each line corrects against the
                    // cell-first inset, scaled like the pitch; single-size cells
                    // self-zero.
                    // Cell first baselines hang off this transition offset (the cell insets
                    // above cancel out of it), so it joins the hhea-ascender opt-in while the
                    // midline image-attach offset below keeps legacy (unprobed for images).
                    double lineBaselineTransition = (baselineOffset - cellFirstBaselineInset) * context.ParagraphSpacingScale;
                    for (int lineIndex = 0; lineIndex < wrappedLines.Length; lineIndex++)
                    {
                        DocxWrappedTextLine line = wrappedLines[lineIndex];
                        if (firstLine)
                        {
                            cursorY -= ResolveListLabelFirstLineExtraLeading(paragraph, fontSize, measurer);
                        }

                        // RV05 calibration (Word 16.0): image top pins to the natural line top.
                        double extraAbove = exactSpacing ? 0d : (midLinePlan?.ShiftAboveHeights[lineIndex] ?? 0d);
                        context.CancellationToken.ThrowIfCancellationRequested();
                        var mathExpansion = inlineMath?.LineIndex == lineIndex
                            ? ResolveTableInlineMathExpansion(inlineMath.Math, fontSize, lineHeight, exactSpacing,
                                cellFirstBaselineInset, baselineOffset, context.ParagraphSpacingScale)
                            : (Above: 0d, Below: 0d);
                        extraAbove = Math.Max(extraAbove, mathExpansion.Above);
                        double extraBelow = Math.Max(mathExpansion.Below,
                            exactSpacing ? 0d : midLinePlan?.ExtraBelow(lineIndex) ?? 0d);
                        cursorY -= extraAbove;
                        double mathWidth = inlineMath?.LineIndex == lineIndex ? inlineMath.Width : 0d;
                        double inlineWidth = (midLinePlan?.LineImageWidths[lineIndex] ?? 0d) + mathWidth;
                        double? leadingTabAdvance = firstLine ? leadingDecimalAdvance : null;
                        double lineWidth = MeasureTextSpansForLayout(line.Spans, fontSize, measurer, tabStops, defaultTabStop, context.PageNumber) + inlineWidth;
                        // RV06 table-align probe: Office centers/rights the drawable cell
                        // text too, letting trailing spaces overflow past the edge.
                        double lineAlignWidth = paragraph.EffectiveProperties.Alignment is DocxTextAlignment.Center or DocxTextAlignment.Right
                            ? MeasureDrawableTextSpansForLayout(line.Spans, fontSize, measurer, tabStops, defaultTabStop, context.PageNumber) + inlineWidth
                            : lineWidth;
                        lineAlignWidth += leadingTabAdvance ?? 0d;
                        double lineX = paragraph.EffectiveProperties.Alignment switch
                        {
                            DocxTextAlignment.Center => paragraphX + Math.Max(0, paragraphWidth - lineAlignWidth) / 2d,
                            DocxTextAlignment.Right => paragraphX + Math.Max(0, paragraphWidth - lineAlignWidth),
                            _ => paragraphX
                        };
                        DocxParagraphLineShape lineShape = CreateParagraphLineShape(
                            paragraph,
                            line,
                            firstRun,
                            firstLine,
                            lineIndex == wrappedLines.Length - 1,
                            cellX + paddingLeft + labelStartOffset,
                            lineX,
                            paragraphWidth,
                            fontSize,
                            measurer,
                            tabStops,
                            defaultTabStop,
                            context.PageNumber, inlineWidth, leadingTabAdvance);
                        // A fitted text-only line occupies its full inset frame. Labels
                        // and inline objects retain their separately placed origins.
                        double? fittedOrigin = inlineWidth == 0d && (!firstLine || paragraph.ListLabel is null)
                            ? paragraphX
                            : null;
                        lineShape = FitTableCellLineText(lineShape, paragraphWidth, fittedOrigin);
                        IReadOnlyList<DocxTextSegmentLayout> emissionSegments = lineShape.Segments;
                        double mathOffset = 0d;
                        if (inlineMath?.LineIndex == lineIndex)
                        {
                            double before = MeasureMidLineBeforeWidth(line.Spans, inlineMath.LineCharOffset, paragraph,
                                firstLine, lineIndex == wrappedLines.Length - 1, paragraphWidth, fontSize, measurer,
                                tabStops, defaultTabStop, context.PageNumber, inlineMath.Width);
                            emissionSegments = ShiftSegmentsPastMidLineImages(emissionSegments, line.Spans, paragraph,
                                firstLine, lineX, [(before, inlineMath.Width)]);
                            mathOffset = lineX + before - lineShape.X;
                        }
                        if (midLinePlan is not null && midLinePlan.ImagesByLine[lineIndex].Count != 0)
                        {
                            var imageShifts = new List<(double BoundaryX, double Shift)>();
                            foreach (DocxMidLineImage placed in midLinePlan.ImagesByLine[lineIndex])
                            {
                                double beforeWidth = MeasureMidLineBeforeWidth(line.Spans, placed.LineCharOffset, paragraph, firstLine, lineIndex == wrappedLines.Length - 1, paragraphWidth, fontSize, measurer, tabStops, defaultTabStop, context.PageNumber, midLinePlan.LineImageWidths[lineIndex]);
                                imageShifts.Add((beforeWidth, placed.Width));
                                // RV06 cellmidline size matrix (Word 16.0, blue-bounds raster measurement):
                                // Office cell image bottoms sit at the text baseline (medium bounds 679.0
                                // against baseline 678.82 both families, all sizes within 0.5), so midline
                                // images anchor by the line baseline instead of hanging a full height below it.
                                placedImages.Add(new DocxInlineImageLayout(
                                    placed.Image,
                                    lineX + beforeWidth + placed.AdvanceBefore,
                                    cursorY - lineBaselineTransition,
                                    placed.Width,
                                    placed.Height,
                                    pageIndex,
                                    SourceBlockIndex: null,
                                    SourceParagraphIndex: paragraphIndex,
                                    Story: null)
                                // Cell cursors originate at the first baseline; the
                                // drawable line frame originates at the line top.
                                { LineFrame = DocxPictureLineFrame.FromLine(cursorY + cellFirstBaselineInset, lineHeight, extraAbove, lineIndex == wrappedLines.Length - 1 ? spacingProfile.ParagraphAfterSpacing : 0d) });
                            }

                            emissionSegments = ShiftSegmentsPastMidLineImages(lineShape.Segments, line.Spans, paragraph, firstLine, lineX, imageShifts);
                        }
                        lines.Add(new DocxTextLineLayout(
                            lineShape.Text,
                            firstRun,
                            fontSize,
                            lineShape.X,
                            cursorY - lineBaselineTransition,
                            lineShape.Width,
                            emissionSegments,
                            SourceBlockIndex: null,
                            SourceParagraphIndex: paragraphIndex,
                            SourceLineIndex: lineIndex,
                            Story: DocxStoryId.TableCell(),
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
                            SourceParagraph: paragraph, EmitsTerminalParagraphMark: false)
                        { FrameOriginOffsetX = cellX + paddingLeft - lineShape.X,
                            ParagraphShadingFrame = CreateParagraphShadingFrame(paragraph, HasFollowingParagraphShading(bodyElements, elementIndex), cellX + paddingLeft, textWidth, lineShape.X,
                                cellFirstBaselineInset * context.ParagraphSpacingScale + lineBaselineTransition + extraAbove,
                                lineHeight + extraAbove + extraBelow, context.ParagraphSpacingScale,
                                lineIndex == wrappedLines.Length - 1 ? spacingProfile.ParagraphAfterSpacing : 0d,
                                firstLine ? ResolveParagraphShadingBeforeSpacing(previousParagraph, paragraph, spacingProfile) : 0d),
                            SourceBaselineOffset = exactSpacing || midLinePlan is null ? 0d :
                                Math.Max(0d, midLinePlan.SourceShiftAboveHeights[lineIndex] - extraAbove),
                            Math = mathWidth > 0d ? inlineMath?.Math : null, IsInlineMath = mathWidth > 0d,
                            InlineMathOffsetX = mathOffset, MathScale = context.ParagraphSpacingScale,
                            // The legacy cursor seat stays nominal. The occupied
                            // strut uses physical metrics relative to that baseline,
                            // just like the scaled line height and transition.
                            OccupiedLineBox = new(lineHeight + extraAbove + extraBelow,
                                cellFirstBaselineInset * context.ParagraphSpacingScale + lineBaselineTransition + extraAbove) });
                        firstLine = false;
                        paragraphX = cellX + paddingLeft + continuationTextStartOffset;
                        paragraphWidth = Math.Max(1d, textWidth - continuationTextStartOffset - GetParagraphRightInset(paragraph, context.ParagraphSpacingScale));
                        cursorY -= lineHeight + extraBelow;
                    }

                    // RV06 table-align probe: cells keep one row-end space beyond authored
                    // trailing too. No spill here (breaks inside cells stay queued); skip
                    // shrink-to-fit cells whose width contract is explicit.
                    if (paragraph.Images.Count == 0 &&
                        paragraph.InlineTextBoxes.Count == 0 &&
                        !cell.FitText &&
                        wrappedLines.Length > 0 &&
                        wrappedLines[^1].Text.EndsWith(' ') &&
                        HasVisibleTextSpans(textSpans) &&
                        context.TextMeasurer is not null)
                    {
                        AppendRowEndSpace(lines, paragraph, firstRun, fontSize, measurer);
                    }
                }

            }
            bool currentHasVisibleText = HasVisibleTextSpans(textSpans);
            for (int imageIndex = 0; imageIndex < paragraph.Images.Count; imageIndex++)
            {
                if (midLinePlan?.PlacedMask[imageIndex] == true)
                {
                    continue;
                }

                DocxInlineImage image = paragraph.Images[imageIndex];
                double paragraphX = cellX + paddingLeft + GetParagraphStartOffset(paragraph, context.ParagraphSpacingScale);
                double paragraphWidth = Math.Max(1d, textWidth - GetParagraphStartOffset(paragraph, context.ParagraphSpacingScale) - GetParagraphRightInset(paragraph, context.ParagraphSpacingScale));
                var (imageWidth, imageHeight) = ResolveInlineImageSize(image, paragraphWidth, context.ParagraphSpacingScale);

                double imageX = paragraph.EffectiveProperties.Alignment switch
                {
                    DocxTextAlignment.Center => paragraphX + Math.Max(0, paragraphWidth - imageWidth) / 2d,
                    DocxTextAlignment.Right => paragraphX + Math.Max(0, paragraphWidth - imageWidth),
                    _ => paragraphX
                };
                placedImages.Add(new DocxInlineImageLayout(image, imageX, cursorY - imageHeight, imageWidth, imageHeight, pageIndex, null, SourceParagraphIndex: paragraphIndex, Story: null)
                { SourceParagraph = paragraph,
                    ParagraphShadingFrame = CreateParagraphShadingFrame(paragraph, HasFollowingParagraphShading(bodyElements, elementIndex), cellX + paddingLeft, textWidth, imageX,
                        imageHeight, imageHeight, context.ParagraphSpacingScale, spacingProfile.ParagraphAfterSpacing,
                        textSpans.Count == 0 && imageIndex == 0 ? ResolveParagraphShadingBeforeSpacing(previousParagraph, paragraph, spacingProfile) : 0d) });
                cursorY -= ResolveCellImageBlockAdvance(image, imageHeight);
            }

            bool tuckApplied = false;
            foreach (DocxInlineTextBox textBox in paragraph.InlineTextBoxes)
            {
                double boxParagraphX = cellX + paddingLeft + GetParagraphStartOffset(paragraph, context.ParagraphSpacingScale);
                double boxParagraphWidth = Math.Max(1d, textWidth - GetParagraphStartOffset(paragraph, context.ParagraphSpacingScale) - GetParagraphRightInset(paragraph, context.ParagraphSpacingScale));
                double boxTop = cursorY;
                bool tuckThisBox = !tuckApplied &&
                    previousParagraphHasText &&
                    paragraph.Images.Count == 0 &&
                    !currentHasVisibleText;
                if (tuckThisBox)
                {
                    double tuckDistance = ResolveCellInlineTextBoxTuckDistance(previousTextFontSize,
                        spacingProfile.AppliedBeforeSpacing, context.ParagraphSpacingScale);
                    boxTop = previousTextBaselineY - tuckDistance;
                }

                DocxInlineTextBoxLayout? textBoxLayout = CreateInlineTextBoxLayout(
                    textBox,
                    sourceBlockIndex: null,
                    boxParagraphX,
                    boxParagraphWidth,
                    boxTop,
                    paragraph.EffectiveProperties.Alignment,
                    measurer,
                    context.DefaultTabStopPoints,
                    context.ParagraphSpacingScale,
                    context.PageNumber ?? pageIndex,
                    context.CancellationToken,
                    sourceParagraphIndex: paragraphIndex,
                    mathEquations: context.MathEquations, pageCount: context.PageCount);
                if (textBoxLayout is null)
                {
                    continue;
                }

                tuckApplied |= tuckThisBox;
                placedBoxes.Add(textBoxLayout);
                cursorY = boxTop - textBoxLayout.BoxHeight;
            }

            pendingSpacingAfter = spacingProfile.ParagraphAfterSpacing;
            previousParagraph = paragraph;
            previousParagraphHasText = currentHasVisibleText && paragraph.Images.Count == 0 && paragraph.InlineTextBoxes.Count == 0;
            if (lines.Count > paragraphLineStart)
            {
                previousTextBaselineY = lines[^1].BaselineY;
                previousTextFontSize = lines[^1].FontSize;
            }
            paragraphIndex++;
            if (terminalMark)
            {
                // Keep the placed insertion frame relative to the nominal
                // baseline, including spacing that owns no occupied height.
                cursorY = paragraphCursorY;
                paragraphTop = cursorY + cellFirstBaselineInset;
                pendingSpacingAfter = 0d;
                for (int i = paragraphLineStart; i < lines.Count; i++)
                    lines[i] = lines[i] with { FirstLineInsetPoints = paragraphTop - lines[i].BaselineY,
                        OccupiedLineBox = new(0d, 0d),
                        AppliedBeforeSpacing = 0d, PendingAfterSpacing = 0d,
                        ParagraphBeforeSpacing = 0d, ParagraphAfterSpacing = 0d };
            }
            RecordDrawingAnchors();

            void RecordDrawingAnchors()
            {
                if (anchors is null) return;
                foreach (var owner in paragraph.DrawingOwners)
                    anchors.Add(new(owner, paragraphTop, cursorY + cellFirstBaselineInset));
            }
        }

        cursorY -= pendingSpacingAfter;
        double usedHeight = Math.Max(0d, startBaselineY - cursorY);

        if (cellMemo is not null)
        {
            cellMemo.StoreRelativeLines(cell, cellWidth, measurer, context.DefaultTabStopPoints, rowTopPadding, context.ParagraphSpacingScale, context.PageNumber, context.PageCount, dependencies, lines, placedImages, placedBoxes, originX, originY, usedHeight, nestedRows, nestedPageHeight, anchors);
        }

        return (lines, usedHeight, placedImages, placedBoxes, nestedRows, anchors ?? []);

        DocxParagraphLineShape FitTableCellLineText(DocxParagraphLineShape lineShape, double targetWidth, double? fittedOrigin)
        {
            if (!cell.FitText ||
                lineShape.Segments.Count == 0 ||
                lineShape.Text.IndexOf('\t') >= 0 ||
                targetWidth <= 0d)
            {
                return lineShape;
            }

            int gapCount = lineShape.Segments.Sum(CountFitTextCharacterSpacingGaps);
            if (gapCount == 0)
            {
                return lineShape with { Width = Math.Max(0d, targetWidth) };
            }

            double fitSpacing = (targetWidth - lineShape.Width) / gapCount;
            var segments = new List<DocxTextSegmentLayout>(lineShape.Segments.Count);
            double origin = fittedOrigin ?? lineShape.X;
            double segmentX = lineShape.Segments[0].X + (origin - lineShape.X);
            for (int index = 0; index < lineShape.Segments.Count; index++)
            {
                DocxTextSegmentLayout segment = lineShape.Segments[index];
                int segmentGapCount = CountFitTextCharacterSpacingGaps(segment);
                double fittedWidth = segment.Width + (fitSpacing * segmentGapCount);
                segments.Add(segmentGapCount == 0
                    ? segment with { X = segmentX }
                    : segment with
                    {
                        X = segmentX,
                        Width = fittedWidth,
                        PdfCharacterSpacing = segment.PdfCharacterSpacing + fitSpacing,
                        PdfCharacterSpacingSource = DocxTextStateCharacterSpacingSource.AdvanceTarget,
                        CompensatePdfCharacterSpacing = false
                    });

                double boundaryAdvance = index + 1 < lineShape.Segments.Count
                    ? lineShape.Segments[index + 1].X - (segment.X + segment.Width)
                    : 0d;
                segmentX += fittedWidth + boundaryAdvance;
            }

            return lineShape with
            {
                X = origin,
                Width = targetWidth,
                Segments = segments
            };
        }

    }

    // All cell content uses the same occupied height and vertical offset;
    // unknown alignment spellings use the parsed Top fallback.
    private static double ResolveTableCellExtraVerticalSpace(IReadOnlyList<DocxParagraph> paragraphs,
        double cellHeight, double paddingTop, double paddingBottom, double usedHeight, double baselineInset)
    {
        // Natural/exact slots include their baseline inset in used height.
        // Minimum-leading content retains the separately reserved initial headroom.
        bool reservesHeadroom = paragraphs.Any(paragraph => paragraph.EffectiveProperties.LineSpacingPoints is > 0d &&
            !IsExactLineSpacing(paragraph.EffectiveProperties));
        double headroom = reservesHeadroom ? baselineInset : 0d;
        return Math.Max(0d, cellHeight - paddingTop - paddingBottom - headroom - usedHeight);
    }

    private static double ResolveVerticalContentOffset(DocxTableCell cell, double extra)
    {
        return cell.VerticalAlignment switch
        {
            DocxTableCellVerticalAlignment.Bottom => extra,
            DocxTableCellVerticalAlignment.Center => extra / 2d,
            _ => 0d,
        };
    }

    private static double ResolveTableCellTextWrapWidth(DocxTableCell cell, double width)
    {
        return cell.NoWrap || cell.FitText ? TableCellNoWrapLineWidthPoints : Math.Max(1d, width);
    }

    private static double? MeasureImplicitCellDecimalAdvance(IReadOnlyList<DocxWrappedTextLine> lines,
        IReadOnlyList<DocxTextSpan> source, double position, double width, double fontSize, IDocxTextMeasurer measurer)
    {
        if (lines.Count == 0 || string.IsNullOrWhiteSpace(lines[0].Text)) return null;
        // Moving a complete fitting first field cannot change its row admission.
        // A soft-wrapped field keeps its existing text flow.
        if (lines.Count > 1)
        {
            int sourceOffset = 0;
            foreach (var span in source)
            {
                int lineBreak = span.Text.IndexOf('\n');
                if (lineBreak >= 0)
                {
                    sourceOffset += lineBreak + 1;
                    break;
                }
                sourceOffset += span.Text.Length;
            }
            if (lines[1].SourceTextStart < sourceOffset) return null;
        }

        var spans = lines[0].Spans;
        var stop = new DocxResolvedTabStop(position, DocxTabStopKind.Decimal);
        var (fieldWidth, anchor) = MeasureTextFieldSpans(spans, fontSize, measurer, alignDecimal: true);
        return TryFitAlignedTabField(stop, fieldWidth, anchor, 0d, width, out double start) ? start : null;
    }

    private static int CountFitTextCharacterSpacingGaps(DocxTextSegmentLayout segment)
    {
        if (segment.Role != DocxTextSegmentRole.Text)
        {
            return 0;
        }

        int runeCount = segment.Text.EnumerateRunes().Count();
        return Math.Max(0, runeCount - 1);
    }
}
