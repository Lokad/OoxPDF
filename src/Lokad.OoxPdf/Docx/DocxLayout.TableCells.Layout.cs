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
    // R16: small immutable cell-layout environment shared by the cell measure and
    // layout passes. Metrics (width-independent scales), dynamic field state (page
    // inputs), and shared services (measurer, memo) travel as one value instead of a
    // long positional tail where a swapped width/scale silently mislays out; per-call
    // geometry (origins, paddings, memo-independent widths) stays positional because it
    // differs per pass. Construction uses named arguments so same-typed page inputs
    // cannot swap.
    internal readonly record struct DocxTableCellLayoutContext(
        IDocxTextMeasurer TextMeasurer,
        double DefaultTabStopPoints,
        double ParagraphSpacingScale,
        int? PageNumber,
        int? PageCount,
        DocxTableCellTextLinesMemo? CellMemo);

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

    private static double ResolveCellInlineTextBoxTuckTop(
        double previousBaselineY,
        double previousFontSize,
        double spacingBefore,
        double paragraphSpacingScale)
    {
        double tuck = paragraphSpacingScale * (CellInlineTextBoxTuckSlope * previousFontSize + CellInlineTextBoxTuckInterceptPoints);
        double excess = Math.Max(0d, spacingBefore - CellInlineTextBoxTuckDefaultAfterPoints * paragraphSpacingScale);
        return previousBaselineY - tuck - excess;
    }

    private static double MeasureTableCellContentHeight(
        DocxTableCell cell,
        double cellWidth,
        IDocxTextMeasurer textMeasurer,
        double defaultTabStopPoints,
        double? rowTopPadding,
        int? pageNumber,
        int? pageCount,
        double paragraphSpacingScale)
    {
        var measureContext = new DocxTableCellLayoutContext(
            TextMeasurer: textMeasurer,
            DefaultTabStopPoints: defaultTabStopPoints,
            ParagraphSpacingScale: paragraphSpacingScale,
            PageNumber: pageNumber,
            PageCount: pageCount,
            CellMemo: null);
        IReadOnlyList<DocxBodyElement> bodyElements = GetTableCellLayoutBodyElements(cell);
        if (bodyElements.Count == 0)
        {
            return 0d;
        }

        double paddingLeft = ResolveTableCellHorizontalEdgeInset(cell, "left", cell.Margins.LeftPoints, paragraphSpacingScale);
        double paddingRight = ResolveTableCellHorizontalEdgeInset(cell, "right", cell.Margins.RightPoints, paragraphSpacingScale);
        double paddingTop = rowTopPadding ?? ResolveTableCellVerticalPadding(cell.Margins.TopPoints, paragraphSpacingScale);
        double paddingBottom = ResolveTableCellVerticalPadding(cell.Margins.BottomPoints, paragraphSpacingScale);
        double textWidth = Math.Max(1d, cellWidth - paddingLeft - paddingRight);
        double contentHeight = paddingTop + paddingBottom;
        double pendingSpacingAfter = 0d;
        DocxParagraph? previousParagraph = null;
        bool previousParagraphHasText = false;
        bool previousParagraphHasTextLines = false;
        double previousTextFontSize = 0d;
        double previousTextLineHeight = 0d;
        foreach (DocxBodyElement bodyElement in bodyElements)
        {
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
            DocxParagraphSpacingProfile spacingProfile = ResolveParagraphSpacingProfile(previousParagraph, paragraph, pendingSpacingAfter, paragraphSpacingScale);
            contentHeight += spacingProfile.AppliedBeforeSpacing;
            pendingSpacingAfter = 0d;
            double fontSize = GetParagraphFontSize(paragraph);
            DocxLineHeightProfile lineHeightProfile = ResolveLineHeightProfile(paragraph, fontSize, textMeasurer);
            double lineHeight = lineHeightProfile.LineHeight;
            IReadOnlyList<DocxTextSpan> textSpans = CreateTextSpans(paragraph.Runs, pageNumber, pageCount);
            DocxMidLinePlan? estimatePlan = null;
            if (textSpans.Count != 0)
            {
                double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, fontSize, textMeasurer, paragraphSpacingScale);
                double firstParagraphWidth = ResolveTableCellTextWrapWidth(cell, textWidth - textStartOffset - GetParagraphRightInset(paragraph, paragraphSpacingScale));
                double continuationParagraphWidth = ResolveTableCellTextWrapWidth(cell, textWidth - GetParagraphTextStartOffset(paragraph, paragraphSpacingScale) - GetParagraphRightInset(paragraph, paragraphSpacingScale));
                DocxWrappedTextLine[] estimateLines = WrapTextLines(textSpans, firstParagraphWidth, continuationParagraphWidth, fontSize, textMeasurer, ScaleTabStopPositions(paragraph.EffectiveProperties.TabStops, paragraphSpacingScale), defaultTabStopPoints * paragraphSpacingScale, allowOverwideTokenBreaks: true, dynamicFieldPageNumber: pageNumber, inlineImageWidths: ResolveInlineImageWrapWidths(paragraph, textSpans)).ToArray();
                int lineCount = estimateLines.Length;
                lineHeight = QuantizeTableCellWrappedLineHeight(lineHeight, lineCount);
                estimatePlan = CreateMidLinePlan(paragraph, textSpans, estimateLines, firstParagraphWidth, continuationParagraphWidth, DocxLineMetrics.ResolveBodyBaselineOffset(fontSize, lineHeight, IsExactLineSpacing(paragraph.EffectiveProperties)), lineHeight);
                previousTextFontSize = fontSize;
                previousTextLineHeight = lineHeight;
                contentHeight += ResolveListLabelFirstLineExtraLeading(paragraph, fontSize, textMeasurer);
                for (int estimateLineIndex = 0; estimateLineIndex < estimateLines.Length; estimateLineIndex++)
                {
                    double estimateShift = IsExactLineSpacing(paragraph.EffectiveProperties) ? 0d : (estimatePlan?.ShiftAboveHeights[estimateLineIndex] ?? 0d);
                    contentHeight += lineHeight + estimateShift;
                }
            }
            else if (paragraph.Images.Count == 0 && paragraph.InlineTextBoxes.Count == 0)
            {
                contentHeight += ResolveListLabelFirstLineExtraLeading(paragraph, fontSize, textMeasurer);
                contentHeight += lineHeight;
            }

            for (int estimateImageIndex = 0; estimateImageIndex < paragraph.Images.Count; estimateImageIndex++)
            {
                if (estimatePlan?.PlacedMask[estimateImageIndex] == true)
                {
                    continue;
                }

                DocxInlineImage image = paragraph.Images[estimateImageIndex];
                double imageWidth = Math.Min(textWidth, image.WidthPoints);
                double imageHeight = image.HeightPoints * imageWidth / Math.Max(1d, image.WidthPoints);
                contentHeight += imageHeight + InlineImageParagraphGapPoints;
            }

            bool currentHasVisibleText = HasVisibleTextSpans(textSpans);
            bool tuckApplied = false;
            foreach (DocxInlineTextBox textBox in paragraph.InlineTextBoxes)
            {
                double boxContribution = EstimateInlineTextBoxHeight(textBox, paragraphSpacingScale) + InlineImageParagraphGapPoints;
                if (!tuckApplied &&
                    previousParagraphHasText &&
                    paragraph.Images.Count == 0 &&
                    !currentHasVisibleText)
                {
                    double tuck = paragraphSpacingScale * (CellInlineTextBoxTuckSlope * previousTextFontSize + CellInlineTextBoxTuckInterceptPoints);
                    double excess = Math.Max(0d, spacingProfile.AppliedBeforeSpacing - CellInlineTextBoxTuckDefaultAfterPoints * paragraphSpacingScale);
                    boxContribution += tuck + excess - (previousTextLineHeight + spacingProfile.AppliedBeforeSpacing);
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
        return contentHeight;
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
            cancellationToken: CancellationToken.None);
        return frame.RowHeights.Sum();
    }

    private static (IReadOnlyList<DocxTextLineLayout> Lines, IReadOnlyList<DocxInlineImageLayout> PlacedImages) LayoutTableCellTextLines(
        DocxTableCell cell,
        double cellX,
        double cellY,
        double cellWidth,
        double cellHeight,
        double rowTopPadding,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        int? pageNumber,
        int? pageCount,
        double paragraphSpacingScale,
        DocxTableCellTextLinesMemo? cellMemo = null,
        int pageIndex = 0)
    {
        if (textMeasurer is null)
        {
            return (Array.Empty<DocxTextLineLayout>(), Array.Empty<DocxInlineImageLayout>());
        }

        var context = new DocxTableCellLayoutContext(
            TextMeasurer: textMeasurer,
            DefaultTabStopPoints: defaultTabStopPoints,
            ParagraphSpacingScale: paragraphSpacingScale,
            PageNumber: pageNumber,
            PageCount: pageCount,
            CellMemo: cellMemo);
        IReadOnlyList<DocxBodyElement> bodyElements = GetTableCellLayoutBodyElements(cell);
        IReadOnlyList<DocxParagraph> paragraphs = GetParagraphsFromBodyElements(bodyElements);
        if (paragraphs.Count == 0)
        {
            return (Array.Empty<DocxTextLineLayout>(), Array.Empty<DocxInlineImageLayout>());
        }

        double paddingLeft = ResolveTableCellHorizontalEdgeInset(cell, "left", cell.Margins.LeftPoints, paragraphSpacingScale);
        double paddingRight = ResolveTableCellHorizontalEdgeInset(cell, "right", cell.Margins.RightPoints, paragraphSpacingScale);
        double paddingTop = rowTopPadding + ResolveTableCellTopBorderContentInset(cell, paragraphSpacingScale);
        double paddingBottom = ResolveTableCellVerticalPadding(cell.Margins.BottomPoints, paragraphSpacingScale);
        double baselineInset = ResolveTableCellFirstBaselineInset(paragraphs);
        double textWidth = Math.Max(1d, cellWidth - paddingLeft - paddingRight);
        double startBaselineY = cellY + cellHeight - baselineInset - paddingTop;
        bool pageStatic = !HasPageDynamicFields(bodyElements);
        (List<DocxTextLineLayout> lines, double usedHeight, List<DocxInlineImageLayout> placedImages) = ComputeTableCellTextLines(cell, cellX, cellWidth, context, bodyElements, paddingLeft, textWidth, startBaselineY, rowTopPadding, cellX, cellY + cellHeight, pageStatic, pageIndex);
        if (lines.Count == 0)
        {
            return (lines, placedImages);
        }
        double availableHeight = Math.Max(0d, cellHeight - paddingTop - paddingBottom - baselineInset);
        double extra = Math.Max(0d, availableHeight - usedHeight);
        double verticalOffset = ResolveVerticalContentOffset(cell, extra);
        return verticalOffset == 0d
            ? (lines, placedImages)
            : (ShiftTextLines(lines, -verticalOffset, 0d), ShiftInlineImages(placedImages, -verticalOffset, 0d));
    }

    private static (List<DocxTextLineLayout> Lines, double UsedHeight, List<DocxInlineImageLayout> PlacedImages) ComputeTableCellTextLines(
        DocxTableCell cell,
        double cellX,
        double cellWidth,
        DocxTableCellLayoutContext context,
        IReadOnlyList<DocxBodyElement> bodyElements,
        double paddingLeft,
        double textWidth,
        double startBaselineY,
        double rowTopPadding,
        double originX,
        double originY,
        bool pageStatic,
        int pageIndex)
    {
        DocxTableCellTextLinesMemo? cellMemo = context.CellMemo;
        // Callers pass a null measurer only through the empty early-return path;
        // bind the proven face once so text measurement below is total.
        if (context.TextMeasurer is not IDocxTextMeasurer measurer)
        {
            return (new List<DocxTextLineLayout>(), 0d, new List<DocxInlineImageLayout>());
        }

        if (cellMemo is not null &&
            cellMemo.TryGetRelativeLines(cell, cellWidth, measurer, context.DefaultTabStopPoints, rowTopPadding, context.ParagraphSpacingScale, context.PageNumber, context.PageCount, pageStatic, out IReadOnlyList<DocxTextLineLayout> cachedRelative, out IReadOnlyList<DocxInlineImageLayout> cachedPlaced, out double cachedUsedHeight))
        {
            return (new List<DocxTextLineLayout>(DocxTableCellTextLinesMemo.ShiftLines(cachedRelative, originX, originY)), cachedUsedHeight, new List<DocxInlineImageLayout>(ShiftInlineImages(cachedPlaced, originY, originX)));
        }

        double cursorY = startBaselineY;
        var lines = new List<DocxTextLineLayout>();
        var placedImages = new List<DocxInlineImageLayout>();
        double pendingSpacingAfter = 0d;
        DocxParagraph? previousParagraph = null;
        int paragraphIndex = 0;
        foreach (DocxBodyElement bodyElement in bodyElements)
        {
            if (bodyElement is DocxTableElement tableElement)
            {
                cursorY -= pendingSpacingAfter;
                pendingSpacingAfter = 0d;
                previousParagraph = null;
                cursorY -= MeasureNestedTableHeight(tableElement.Table, textWidth, context);
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
                if (previousParagraph is not null)
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
                continue;
            }

            if (bodyElement is not DocxParagraphElement paragraphElement)
            {
                continue;
            }

            DocxParagraph paragraph = paragraphElement.Paragraph;
            DocxParagraphSpacingProfile spacingProfile = ResolveParagraphSpacingProfile(previousParagraph, paragraph, pendingSpacingAfter, context.ParagraphSpacingScale);
            cursorY -= spacingProfile.AppliedBeforeSpacing;
            pendingSpacingAfter = 0d;
            double fontSize = GetParagraphFontSize(paragraph);
            DocxLineHeightProfile lineHeightProfile = ResolveLineHeightProfile(paragraph, fontSize, measurer);
            double lineHeight = lineHeightProfile.LineHeight;
            IReadOnlyList<DocxTextSpan> textSpans = CreateTextSpans(paragraph.Runs, context.PageNumber, context.PageCount);
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
                double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, fontSize, measurer, context.ParagraphSpacingScale);
                double continuationTextStartOffset = GetParagraphTextStartOffset(paragraph, context.ParagraphSpacingScale);
                double labelStartOffset = GetParagraphLabelStartOffset(paragraph, context.ParagraphSpacingScale);
                double paragraphX = cellX + paddingLeft + textStartOffset;
                double paragraphWidth = Math.Max(1d, textWidth - textStartOffset - GetParagraphRightInset(paragraph, context.ParagraphSpacingScale));
                double continuationParagraphWidth = Math.Max(1d, textWidth - continuationTextStartOffset - GetParagraphRightInset(paragraph, context.ParagraphSpacingScale));
                bool firstLine = true;
                DocxWrappedTextLine[] wrappedLines = WrapTextLines(
                    textSpans,
                    ResolveTableCellTextWrapWidth(cell, paragraphWidth),
                    ResolveTableCellTextWrapWidth(cell, continuationParagraphWidth),
                    fontSize,
                    measurer,
                    ScaleTabStopPositions(paragraph.EffectiveProperties.TabStops, context.ParagraphSpacingScale),
                    context.DefaultTabStopPoints * context.ParagraphSpacingScale,
                    allowOverwideTokenBreaks: true,
                    dynamicFieldPageNumber: context.PageNumber, inlineImageWidths: ResolveInlineImageWrapWidths(paragraph, textSpans)).ToArray();
                lineHeight = QuantizeTableCellWrappedLineHeight(lineHeight, wrappedLines.Length);
                // RV05: ordered inline atoms (table-cell path). Affined images in
                // text-mixed paragraphs attach to wrapped lines at run position.
                DocxMidLinePlan? midLinePlan = CreateMidLinePlan(paragraph, textSpans, wrappedLines, ResolveTableCellTextWrapWidth(cell, paragraphWidth), ResolveTableCellTextWrapWidth(cell, continuationParagraphWidth), DocxLineMetrics.ResolveBodyBaselineOffset(fontSize, lineHeight, IsExactLineSpacing(paragraph.EffectiveProperties)), lineHeight);
                for (int lineIndex = 0; lineIndex < wrappedLines.Length; lineIndex++)
                {
                    DocxWrappedTextLine line = wrappedLines[lineIndex];
                    if (firstLine)
                    {
                        cursorY -= ResolveListLabelFirstLineExtraLeading(paragraph, fontSize, measurer);
                    }

                    // RV05 calibration (Word 16.0): image top pins to the natural line top.
                    double extraAbove = IsExactLineSpacing(paragraph.EffectiveProperties) ? 0d : (midLinePlan?.ShiftAboveHeights[lineIndex] ?? 0d);
                    cursorY -= extraAbove;
                    double lineWidth = MeasureTextSpansForLayout(line.Spans, fontSize, measurer, ScaleTabStopPositions(paragraph.EffectiveProperties.TabStops, context.ParagraphSpacingScale), context.DefaultTabStopPoints * context.ParagraphSpacingScale, context.PageNumber) + (midLinePlan?.LineImageWidths[lineIndex] ?? 0d);
                    // RV06 table-align probe: Office centers/rights the drawable cell
                    // text too, letting trailing spaces overflow past the edge.
                    double lineAlignWidth = paragraph.EffectiveProperties.Alignment is DocxTextAlignment.Center or DocxTextAlignment.Right
                        ? MeasureDrawableTextSpansForLayout(line.Spans, fontSize, measurer, ScaleTabStopPositions(paragraph.EffectiveProperties.TabStops, context.ParagraphSpacingScale), context.DefaultTabStopPoints * context.ParagraphSpacingScale, context.PageNumber) + (midLinePlan?.LineImageWidths[lineIndex] ?? 0d)
                        : lineWidth;
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
                        ScaleTabStopPositions(paragraph.EffectiveProperties.TabStops, context.ParagraphSpacingScale),
                        context.DefaultTabStopPoints * context.ParagraphSpacingScale,
                        context.PageNumber);
                    lineShape = FitTableCellLineText(lineShape, paragraphWidth);
                    IReadOnlyList<DocxTextSegmentLayout> emissionSegments = lineShape.Segments;
                    if (midLinePlan is not null && midLinePlan.ImagesByLine[lineIndex].Count != 0)
                    {
                        var imageShifts = new List<(double BoundaryX, double Shift)>();
                        foreach (DocxMidLineImage placed in midLinePlan.ImagesByLine[lineIndex])
                        {
                            double shiftBeforeWidth = MeasureMidLineBeforeWidth(line.Spans, placed.LineCharOffset, paragraph, firstLine, lineIndex == wrappedLines.Length - 1, paragraphWidth, fontSize, measurer, ScaleTabStopPositions(paragraph.EffectiveProperties.TabStops, context.ParagraphSpacingScale), context.DefaultTabStopPoints * context.ParagraphSpacingScale, context.PageNumber);
                            imageShifts.Add((shiftBeforeWidth, placed.Width));
                        }

                        emissionSegments = ShiftSegmentsPastMidLineImages(lineShape.Segments, line.Spans, paragraph, firstLine, lineX, imageShifts);
                    }
                    lines.Add(new DocxTextLineLayout(
                        lineShape.Text,
                        firstRun,
                        fontSize,
                        lineShape.X,
                        cursorY,
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
                        SourceParagraph: paragraph, EmitsTerminalParagraphMark: false));
                    if (midLinePlan is not null)
                    {
                        foreach (DocxMidLineImage placed in midLinePlan.ImagesByLine[lineIndex])
                        {
                            double beforeWidth = MeasureMidLineBeforeWidth(line.Spans, placed.LineCharOffset, paragraph, firstLine, lineIndex == wrappedLines.Length - 1, paragraphWidth, fontSize, measurer, ScaleTabStopPositions(paragraph.EffectiveProperties.TabStops, context.ParagraphSpacingScale), context.DefaultTabStopPoints * context.ParagraphSpacingScale, context.PageNumber);
                            placedImages.Add(new DocxInlineImageLayout(
                                placed.Image,
                                lineX + beforeWidth,
                                cursorY - placed.Height,
                                placed.Width,
                                placed.Height,
                                pageIndex,
                                SourceBlockIndex: null,
                                SourceParagraphIndex: paragraphIndex,
                                Story: null));
                        }
                    }
                    firstLine = false;
                    paragraphX = cellX + paddingLeft + continuationTextStartOffset;
                    paragraphWidth = Math.Max(1d, textWidth - continuationTextStartOffset - GetParagraphRightInset(paragraph, context.ParagraphSpacingScale));
                    cursorY -= lineHeight;
                }

                // RV06 table-align probe: cells keep one row-end space beyond authored
                // trailing too. No spill here (breaks inside cells stay queued); skip
                // shrink-to-fit cells whose width contract is explicit.
                if (paragraph.Images.Count == 0 &&
                    paragraph.InlineTextBoxes.Count == 0 &&
                    !cell.FitText &&
                    wrappedLines.Length > 0 &&
                    wrappedLines[^1].Text.EndsWith(' ') &&
                    textSpans.Any(static span => span.Text.Any(static character => !char.IsWhiteSpace(character))) &&
                    context.TextMeasurer is not null)
                {
                    for (int cellLineIndex = lines.Count - 1; cellLineIndex >= 0; cellLineIndex--)
                    {
                        if (lines[cellLineIndex] is not DocxTextLineLayout cellLastLine ||
                            !ReferenceEquals(cellLastLine.SourceParagraph, paragraph) ||
                            cellLastLine.Text.Length == 0)
                        {
                            continue;
                        }

                        double cellSpaceWidth = measurer.MeasureText(firstRun, " ", fontSize);
                        lines[cellLineIndex] = cellLastLine with
                        {
                            Text = cellLastLine.Text + " ",
                            Width = cellLastLine.Width + cellSpaceWidth,
                            Segments =
                            [
                                .. cellLastLine.Segments,
                                new DocxTextSegmentLayout(
                                    " ",
                                    firstRun,
                                    cellLastLine.X + cellLastLine.Width,
                                    cellSpaceWidth,
                                    fontSize,
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
            }

            pendingSpacingAfter = spacingProfile.ParagraphAfterSpacing;
            previousParagraph = paragraph;
            paragraphIndex++;
        }

        cursorY -= pendingSpacingAfter;
        if (lines.Count == 0)
        {
            return (lines, 0d, placedImages);
        }

        double usedHeight = Math.Max(0d, startBaselineY - cursorY);

        if (cellMemo is not null)
        {
            cellMemo.StoreRelativeLines(cell, cellWidth, measurer, context.DefaultTabStopPoints, rowTopPadding, context.ParagraphSpacingScale, context.PageNumber, context.PageCount, pageStatic, lines, placedImages, originX, originY, usedHeight);
        }

        return (lines, usedHeight, placedImages);

        DocxParagraphLineShape FitTableCellLineText(DocxParagraphLineShape lineShape, double targetWidth)
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
            double segmentX = lineShape.Segments[0].X;
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
                Width = targetWidth,
                Segments = segments
            };
        }

    }

    private static (IReadOnlyList<DocxInlineImageLayout> Images, IReadOnlyList<DocxInlineTextBoxLayout> TextBoxes) LayoutTableCellInlineImages(
        DocxTableCell cell,
        double cellX,
        double cellY,
        double cellWidth,
        double cellHeight,
        double rowTopPadding,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        int pageIndex,
        int? pageNumber,
        int? pageCount,
        double paragraphSpacingScale)
    {
        IReadOnlyList<DocxBodyElement> bodyElements = GetTableCellLayoutBodyElements(cell);
        IReadOnlyList<DocxParagraph> paragraphs = GetParagraphsFromBodyElements(bodyElements);
        if (paragraphs.Count == 0 || !paragraphs.Any(paragraph => paragraph.Images.Count != 0 || paragraph.InlineTextBoxes.Count != 0))
        {
            return ([], []);
        }

        double paddingLeft = ResolveTableCellHorizontalEdgeInset(cell, "left", cell.Margins.LeftPoints, paragraphSpacingScale);
        double paddingRight = ResolveTableCellHorizontalEdgeInset(cell, "right", cell.Margins.RightPoints, paragraphSpacingScale);
        double paddingTop = rowTopPadding + ResolveTableCellTopBorderContentInset(cell, paragraphSpacingScale);
        double paddingBottom = ResolveTableCellVerticalPadding(cell.Margins.BottomPoints, paragraphSpacingScale);
        double baselineInset = ResolveTableCellFirstBaselineInset(paragraphs);
        double textWidth = Math.Max(1d, cellWidth - paddingLeft - paddingRight);
        double startBaselineY = cellY + cellHeight - baselineInset - paddingTop;
        double cursorY = startBaselineY;
        var images = new List<DocxInlineImageLayout>();
        var boxes = new List<DocxInlineTextBoxLayout>();
        double pendingSpacingAfter = 0d;
        DocxParagraph? previousParagraph = null;
        bool previousParagraphHasText = false;
        double previousTextBaselineY = 0d;
        double previousTextFontSize = 0d;
        double previousTextLineHeight = 0d;
        int paragraphIndex = 0;
        foreach (DocxBodyElement bodyElement in bodyElements)
        {
            if (bodyElement is DocxTableElement tableElement)
            {
                cursorY -= pendingSpacingAfter;
                pendingSpacingAfter = 0d;
                previousParagraph = null;
                previousParagraphHasText = false;
                if (textMeasurer is not null)
                {
                    var nestedContext = new DocxTableCellLayoutContext(
                        TextMeasurer: textMeasurer,
                        DefaultTabStopPoints: defaultTabStopPoints,
                        ParagraphSpacingScale: paragraphSpacingScale,
                        PageNumber: pageNumber,
                        PageCount: pageCount,
                        CellMemo: null);
                    cursorY -= MeasureNestedTableHeight(tableElement.Table, textWidth, nestedContext);
                }

                continue;
            }

            if (bodyElement is DocxPageBreakElement inlineCellBreak &&
                inlineCellBreak.SourceKind == DocxBreakSourceKind.RunBreak)
            {
                // RV06 cellbreak probe: mirror the text walk so block content after the
                // break sits below the previous line plus one spill row, with spacing
                // flowing as for an empty default paragraph. The spill height follows
                // this walk's line-height convention (profile height).
                cursorY -= Math.Max(0d, pendingSpacingAfter);
                pendingSpacingAfter = 0d;
                if (previousParagraph is not null &&
                    CreateTextSpans(previousParagraph.Runs, pageNumber, pageCount).Count != 0)
                {
                    cursorY -= previousTextLineHeight;
                }

                pendingSpacingAfter = DocxDefaults.DefaultParagraphAfterSpacingPoints * paragraphSpacingScale;
                previousParagraph = null;
                previousParagraphHasText = false;
                continue;
            }

            if (bodyElement is not DocxParagraphElement paragraphElement)
            {
                continue;
            }

            DocxParagraph paragraph = paragraphElement.Paragraph;
            DocxParagraphSpacingProfile spacingProfile = ResolveParagraphSpacingProfile(previousParagraph, paragraph, pendingSpacingAfter, paragraphSpacingScale);
            cursorY -= spacingProfile.AppliedBeforeSpacing;
            pendingSpacingAfter = 0d;
            bool currentHasVisibleText = false;
            DocxMidLinePlan? midLinePlan = null;
            if (textMeasurer is not null)
            {
                double fontSize = GetParagraphFontSize(paragraph);
                double lineHeight = ResolveLineHeight(paragraph, fontSize, textMeasurer);
                previousTextLineHeight = lineHeight;
                IReadOnlyList<DocxTextSpan> textSpans = CreateTextSpans(paragraph.Runs, pageNumber, pageCount);
                if (textSpans.Count != 0)
                {
                    double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, fontSize, textMeasurer, paragraphSpacingScale);
                    double firstParagraphWidth = ResolveTableCellTextWrapWidth(cell, textWidth - textStartOffset - GetParagraphRightInset(paragraph, paragraphSpacingScale));
                    double continuationParagraphWidth = ResolveTableCellTextWrapWidth(cell, textWidth - GetParagraphTextStartOffset(paragraph, paragraphSpacingScale) - GetParagraphRightInset(paragraph, paragraphSpacingScale));
                    DocxWrappedTextLine[] wrappedLines = WrapTextLines(textSpans, firstParagraphWidth, continuationParagraphWidth, fontSize, textMeasurer, ScaleTabStopPositions(paragraph.EffectiveProperties.TabStops, paragraphSpacingScale), defaultTabStopPoints * paragraphSpacingScale, allowOverwideTokenBreaks: true, dynamicFieldPageNumber: pageNumber, inlineImageWidths: ResolveInlineImageWrapWidths(paragraph, textSpans)).ToArray();
                    midLinePlan = CreateMidLinePlan(paragraph, textSpans, wrappedLines, firstParagraphWidth, continuationParagraphWidth, DocxLineMetrics.ResolveBodyBaselineOffset(fontSize, lineHeight, IsExactLineSpacing(paragraph.EffectiveProperties)), lineHeight);
                    int wrappedLineCount = wrappedLines.Length;
                    if (wrappedLineCount != 0)
                    {
                        previousTextBaselineY = cursorY - (wrappedLineCount - 1) * lineHeight;
                        previousTextFontSize = fontSize;
                    }

                    currentHasVisibleText = HasVisibleTextSpans(textSpans);
                    cursorY -= wrappedLineCount * lineHeight;
                }
                else if (paragraph.Images.Count == 0 && paragraph.InlineTextBoxes.Count == 0)
                {
                    cursorY -= lineHeight;
                }
            }

            for (int imageIndex = 0; imageIndex < paragraph.Images.Count; imageIndex++)
            {
                if (midLinePlan?.PlacedMask[imageIndex] == true)
                {
                    continue;
                }

                DocxInlineImage image = paragraph.Images[imageIndex];
                double paragraphX = cellX + paddingLeft + GetParagraphStartOffset(paragraph, paragraphSpacingScale);
                double paragraphWidth = Math.Max(1d, textWidth - GetParagraphStartOffset(paragraph, paragraphSpacingScale) - GetParagraphRightInset(paragraph, paragraphSpacingScale));
                double imageWidth = Math.Min(paragraphWidth, image.WidthPoints);
                double imageHeight = image.HeightPoints * imageWidth / Math.Max(1d, image.WidthPoints);
                double imageX = paragraph.EffectiveProperties.Alignment switch
                {
                    DocxTextAlignment.Center => paragraphX + Math.Max(0, paragraphWidth - imageWidth) / 2d,
                    DocxTextAlignment.Right => paragraphX + Math.Max(0, paragraphWidth - imageWidth),
                    _ => paragraphX
                };
                images.Add(new DocxInlineImageLayout(image, imageX, cursorY - imageHeight, imageWidth, imageHeight, pageIndex, null, SourceParagraphIndex: paragraphIndex, Story: null));
                cursorY -= imageHeight + InlineImageParagraphGapPoints;
            }

            bool tuckApplied = false;
            foreach (DocxInlineTextBox textBox in paragraph.InlineTextBoxes)
            {
                double boxParagraphX = cellX + paddingLeft + GetParagraphStartOffset(paragraph, paragraphSpacingScale);
                double boxParagraphWidth = Math.Max(1d, textWidth - GetParagraphStartOffset(paragraph, paragraphSpacingScale) - GetParagraphRightInset(paragraph, paragraphSpacingScale));
                double boxTop = cursorY;
                if (!tuckApplied &&
                    previousParagraphHasText &&
                    paragraph.Images.Count == 0 &&
                    !currentHasVisibleText)
                {
                    boxTop = ResolveCellInlineTextBoxTuckTop(previousTextBaselineY, previousTextFontSize, spacingProfile.AppliedBeforeSpacing, paragraphSpacingScale);
                    tuckApplied = true;
                }

                DocxInlineTextBoxLayout? textBoxLayout = CreateInlineTextBoxLayout(
                    textBox,
                    sourceBlockIndex: null,
                    boxParagraphX,
                    boxParagraphWidth,
                    boxTop,
                    paragraph.EffectiveProperties.Alignment,
                    textMeasurer,
                    defaultTabStopPoints,
                    paragraphSpacingScale,
                    pageNumber ?? pageIndex,
                    CancellationToken.None,
                    sourceParagraphIndex: paragraphIndex);
                if (textBoxLayout is null)
                {
                    continue;
                }

                boxes.Add(textBoxLayout);
                cursorY = boxTop - textBoxLayout.BoxHeight - InlineImageParagraphGapPoints;
            }

            pendingSpacingAfter = spacingProfile.ParagraphAfterSpacing;
            previousParagraph = paragraph;
            previousParagraphHasText = currentHasVisibleText && paragraph.Images.Count == 0 && paragraph.InlineTextBoxes.Count == 0;
            paragraphIndex++;
        }

        cursorY -= pendingSpacingAfter;
        if (images.Count == 0 && boxes.Count == 0)
        {
            return ([], []);
        }

        double usedHeight = Math.Max(0d, startBaselineY - cursorY);
        double availableHeight = Math.Max(0d, cellHeight - paddingTop - paddingBottom - baselineInset);
        double extra = Math.Max(0d, availableHeight - usedHeight);
        double verticalOffset = ResolveVerticalContentOffset(cell, extra);
        return verticalOffset == 0d
            ? (images, boxes)
            : (images.Select(image => image with { Y = image.Y - verticalOffset }).ToArray(),
                boxes.Select(box => ShiftInlineTextBox(box, -verticalOffset, 0d)).ToArray());
    }

    private static IReadOnlyList<DocxTableRowLayout> LayoutTableCellNestedTables(
        DocxTableCell cell,
        double cellX,
        double cellY,
        double cellWidth,
        double cellHeight,
        double rowTopPadding,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        int pageIndex,
        int? pageNumber,
        int? pageCount,
        double paragraphSpacingScale,
        DocxTableCellTextLinesMemo? cellMemo = null)
    {
        IReadOnlyList<DocxBodyElement> bodyElements = GetTableCellLayoutBodyElements(cell);
        if (textMeasurer is null || !bodyElements.OfType<DocxTableElement>().Any())
        {
            return [];
        }

        var context = new DocxTableCellLayoutContext(
            TextMeasurer: textMeasurer,
            DefaultTabStopPoints: defaultTabStopPoints,
            ParagraphSpacingScale: paragraphSpacingScale,
            PageNumber: pageNumber,
            PageCount: pageCount,
            CellMemo: cellMemo);

        double paddingLeft = ResolveTableCellHorizontalEdgeInset(cell, "left", cell.Margins.LeftPoints, paragraphSpacingScale);
        double paddingRight = ResolveTableCellHorizontalEdgeInset(cell, "right", cell.Margins.RightPoints, paragraphSpacingScale);
        double textWidth = Math.Max(1d, cellWidth - paddingLeft - paddingRight);
        double cursorY = cellY + cellHeight - rowTopPadding - ResolveTableCellTopBorderContentInset(cell, paragraphSpacingScale);
        var nestedRows = new List<DocxTableRowLayout>();
        double pendingSpacingAfter = 0d;
        DocxParagraph? previousParagraph = null;
        int nestedTableIndex = 0;
        foreach (DocxBodyElement bodyElement in bodyElements)
        {
            if (bodyElement is DocxTableElement tableElement)
            {
                cursorY -= pendingSpacingAfter;
                pendingSpacingAfter = 0d;
                previousParagraph = null;
                DocxTableLayoutFrame frame = CreateTableLayoutFrame(
                    tableElement.Table,
                    nestedTableIndex,
                    sourceBlockIndex: -1,
                    cellX + paddingLeft,
                    textWidth,
                    cellHeight,
                    textMeasurer,
                    defaultTabStopPoints,
                    pageNumber: pageNumber,
                    pageCount: pageCount,
                    paragraphSpacingScale: paragraphSpacingScale,
                    cancellationToken: CancellationToken.None);
                int nestedPageNumber = pageNumber ?? pageIndex + 1;
                for (int rowIndex = 0; rowIndex < tableElement.Table.Rows.Count; rowIndex++)
                {
                    double rowHeight = frame.RowHeights[rowIndex];
                    nestedRows.Add(CreateTableRowLayout(
                        tableElement.Table,
                        frame.Context,
                        tableElement.Table.Rows[rowIndex],
                        rowIndex,
                        frame.RowHeights,
                        frame.EffectiveColumns,
                        frame.Scale,
                        textMeasurer,
                        defaultTabStopPoints,
                        () => nestedPageNumber,
                        cursorY,
                        rowHeight,
                        cursorY,
                        FragmentIndex: 0,
                        FragmentCount: 1,
                        FragmentReason: "None",
                        Story: null,
                        pageCount: pageCount,
                        paragraphSpacingScale: paragraphSpacingScale,
                        cellMemo: cellMemo));
                    cursorY -= rowHeight;
                }

                nestedTableIndex++;
                continue;
            }

            if (bodyElement is DocxPageBreakElement nestedCellBreak &&
                nestedCellBreak.SourceKind == DocxBreakSourceKind.RunBreak)
            {
                // RV06 cellbreak probe: explicit page breaks inside table cells do not
                // turn pages in Office; content after the break sits below the previous
                // line plus one spill row, exactly like the text walk lays out. Spacing
                // flows as for an empty default paragraph: pending is consumed, then the
                // default after applies. The spill height follows this walk's line-height
                // convention (profile height).
                cursorY -= Math.Max(0d, pendingSpacingAfter);
                pendingSpacingAfter = 0d;
                if (previousParagraph is not null &&
                    CreateTextSpans(previousParagraph.Runs, pageNumber, pageCount).Count != 0)
                {
                    double spillFontSize = GetParagraphFontSize(previousParagraph);
                    cursorY -= ResolveLineHeightProfile(previousParagraph, spillFontSize, textMeasurer).LineHeight;
                }

                pendingSpacingAfter = DocxDefaults.DefaultParagraphAfterSpacingPoints * paragraphSpacingScale;
                previousParagraph = null;
                continue;
            }

            if (bodyElement is not DocxParagraphElement paragraphElement)
            {
                continue;
            }

            DocxParagraph paragraph = paragraphElement.Paragraph;
            DocxParagraphSpacingProfile spacingProfile = ResolveParagraphSpacingProfile(previousParagraph, paragraph, pendingSpacingAfter, paragraphSpacingScale);
            cursorY -= spacingProfile.AppliedBeforeSpacing;
            pendingSpacingAfter = 0d;
            cursorY -= MeasureTableCellParagraphContentHeight(cell, paragraph, textWidth, context);
            pendingSpacingAfter = spacingProfile.ParagraphAfterSpacing;
            previousParagraph = paragraph;
        }

        return nestedRows;
    }

    // R17: single vertical-alignment execution point. The parsed enum keeps the
    // two content shifters (text lines, images/boxes) consistent; unknown spellings
    // ride the Top fallback from parsing.
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

    private static int CountFitTextCharacterSpacingGaps(DocxTextSegmentLayout segment)
    {
        if (segment.Role != DocxTextSegmentRole.Text)
        {
            return 0;
        }

        int runeCount = segment.Text.EnumerateRunes().Count();
        return Math.Max(0, runeCount - 1);
    }

    private static double MeasureTableCellParagraphContentHeight(
        DocxTableCell cell,
        DocxParagraph paragraph,
        double textWidth,
        DocxTableCellLayoutContext context)
    {
        double height = 0d;
        double fontSize = GetParagraphFontSize(paragraph);
        double lineHeight = ResolveLineHeight(paragraph, fontSize, context.TextMeasurer);
        IReadOnlyList<DocxTextSpan> textSpans = CreateTextSpans(paragraph.Runs, context.PageNumber, context.PageCount);
        DocxMidLinePlan? singlePlan = null;
        if (textSpans.Count != 0)
        {
            double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, fontSize, context.TextMeasurer, context.ParagraphSpacingScale);
            double firstParagraphWidth = ResolveTableCellTextWrapWidth(cell, textWidth - textStartOffset - GetParagraphRightInset(paragraph, context.ParagraphSpacingScale));
            double continuationParagraphWidth = ResolveTableCellTextWrapWidth(cell, textWidth - GetParagraphTextStartOffset(paragraph, context.ParagraphSpacingScale) - GetParagraphRightInset(paragraph, context.ParagraphSpacingScale));
            height += ResolveListLabelFirstLineExtraLeading(paragraph, fontSize, context.TextMeasurer);
            DocxWrappedTextLine[] singleLines = WrapTextLines(textSpans, firstParagraphWidth, continuationParagraphWidth, fontSize, context.TextMeasurer, ScaleTabStopPositions(paragraph.EffectiveProperties.TabStops, context.ParagraphSpacingScale), context.DefaultTabStopPoints * context.ParagraphSpacingScale, allowOverwideTokenBreaks: true, dynamicFieldPageNumber: context.PageNumber, inlineImageWidths: ResolveInlineImageWrapWidths(paragraph, textSpans)).ToArray();
            singlePlan = CreateMidLinePlan(paragraph, textSpans, singleLines, firstParagraphWidth, continuationParagraphWidth, DocxLineMetrics.ResolveBodyBaselineOffset(fontSize, lineHeight, IsExactLineSpacing(paragraph.EffectiveProperties)), lineHeight);
            for (int singleLineIndex = 0; singleLineIndex < singleLines.Length; singleLineIndex++)
            {
                double singleShift = IsExactLineSpacing(paragraph.EffectiveProperties) ? 0d : (singlePlan?.ShiftAboveHeights[singleLineIndex] ?? 0d);
                height += lineHeight + singleShift;
            }
        }
        else if (paragraph.Images.Count == 0)
        {
            height += ResolveListLabelFirstLineExtraLeading(paragraph, fontSize, context.TextMeasurer);
            height += lineHeight;
        }

        for (int singleImageIndex = 0; singleImageIndex < paragraph.Images.Count; singleImageIndex++)
        {
            if (singlePlan?.PlacedMask[singleImageIndex] == true)
            {
                continue;
            }

            DocxInlineImage image = paragraph.Images[singleImageIndex];
            double imageWidth = Math.Min(textWidth, image.WidthPoints);
            double imageHeight = image.HeightPoints * imageWidth / Math.Max(1d, image.WidthPoints);
            height += imageHeight + InlineImageParagraphGapPoints;
        }

        return height;
    }
}
