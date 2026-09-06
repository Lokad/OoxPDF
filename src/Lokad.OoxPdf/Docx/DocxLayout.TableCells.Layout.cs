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
    // Word hangs the full top-border width above cell content (border probes 2026-09-06: top-term deltas 0.96/1.08/0.58 vs widths 1.0/1.0/0.5); horizontal edges keep the half-width content inset.
    private static double ResolveTableCellTopBorderContentInset(DocxTableCell cell)
    {
        return DocxTableBorderGeometry.ResolveVisibleWidth(DocxTableBorderGeometry.Find(cell.Borders, "top"));
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
        IReadOnlyList<DocxBodyElement> bodyElements = GetTableCellLayoutBodyElements(cell);
        if (bodyElements.Count == 0)
        {
            return 0d;
        }

        double paddingLeft = ResolveTableCellHorizontalPadding(cell.Margins.LeftPoints) + ResolveTableCellBorderContentInset(cell, "left");
        double paddingRight = ResolveTableCellHorizontalPadding(cell.Margins.RightPoints) + ResolveTableCellBorderContentInset(cell, "right");
        double paddingTop = rowTopPadding ?? ResolveTableCellVerticalPadding(cell.Margins.TopPoints);
        double paddingBottom = ResolveTableCellVerticalPadding(cell.Margins.BottomPoints);
        double textWidth = Math.Max(1d, cellWidth - paddingLeft - paddingRight);
        double contentHeight = paddingTop + paddingBottom;
        double pendingSpacingAfter = 0d;
        DocxParagraph? previousParagraph = null;
        foreach (DocxBodyElement bodyElement in bodyElements)
        {
            if (bodyElement is DocxTableElement tableElement)
            {
                contentHeight += pendingSpacingAfter;
                pendingSpacingAfter = 0d;
                previousParagraph = null;
                contentHeight += MeasureNestedTableHeight(tableElement.Table, textWidth, textMeasurer, defaultTabStopPoints, pageNumber, pageCount, paragraphSpacingScale);
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
            if (textSpans.Count != 0)
            {
                double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, fontSize, textMeasurer);
                double firstParagraphWidth = ResolveTableCellTextWrapWidth(cell, textWidth - textStartOffset - GetParagraphRightInset(paragraph));
                double continuationParagraphWidth = ResolveTableCellTextWrapWidth(cell, textWidth - GetParagraphTextStartOffset(paragraph) - GetParagraphRightInset(paragraph));
                int lineCount = WrapTextLines(textSpans, firstParagraphWidth, continuationParagraphWidth, fontSize, textMeasurer, paragraph.EffectiveProperties.TabStops, defaultTabStopPoints, allowOverwideTokenBreaks: true, dynamicFieldPageNumber: pageNumber).Count();
                lineHeight = QuantizeTableCellWrappedLineHeight(lineHeight, lineCount);
                contentHeight += lineCount * lineHeight;
            }
            else if (paragraph.Images.Count == 0)
            {
                contentHeight += lineHeight;
            }

            foreach (DocxInlineImage image in paragraph.Images)
            {
                double imageWidth = Math.Min(textWidth, image.WidthPoints);
                double imageHeight = image.HeightPoints * imageWidth / Math.Max(1d, image.WidthPoints);
                contentHeight += imageHeight + InlineImageParagraphGapPoints;
            }

            pendingSpacingAfter = spacingProfile.ParagraphAfterSpacing;
            previousParagraph = paragraph;
        }

        contentHeight += pendingSpacingAfter;
        return contentHeight;
    }

    private static double MeasureNestedTableHeight(
        DocxTable table,
        double availableWidth,
        IDocxTextMeasurer textMeasurer,
        double defaultTabStopPoints,
        int? pageNumber,
        int? pageCount,
        double paragraphSpacingScale)
    {
        DocxTableLayoutFrame frame = CreateTableLayoutFrame(
            table,
            tableIndex: -1,
            sourceBlockIndex: -1,
            x: 0d,
            availableWidth: availableWidth,
            pageContentHeight: double.MaxValue / 4d,
            textMeasurer,
            defaultTabStopPoints,
            pageNumber: pageNumber,
            pageCount: pageCount,
            paragraphSpacingScale: paragraphSpacingScale,
            cancellationToken: CancellationToken.None);
        return frame.RowHeights.Sum();
    }

    private static IReadOnlyList<DocxTextLineLayout> LayoutTableCellTextLines(
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
        double paragraphSpacingScale)
    {
        if (textMeasurer is null)
        {
            return [];
        }

        IReadOnlyList<DocxBodyElement> bodyElements = GetTableCellLayoutBodyElements(cell);
        IReadOnlyList<DocxParagraph> paragraphs = GetParagraphsFromBodyElements(bodyElements);
        if (paragraphs.Count == 0)
        {
            return [];
        }

        double paddingLeft = ResolveTableCellHorizontalPadding(cell.Margins.LeftPoints) + ResolveTableCellBorderContentInset(cell, "left");
        double paddingRight = ResolveTableCellHorizontalPadding(cell.Margins.RightPoints) + ResolveTableCellBorderContentInset(cell, "right");
        double paddingTop = rowTopPadding + ResolveTableCellTopBorderContentInset(cell);
        double paddingBottom = ResolveTableCellVerticalPadding(cell.Margins.BottomPoints);
        double baselineInset = ResolveTableCellFirstBaselineInset(paragraphs);
        double textWidth = Math.Max(1d, cellWidth - paddingLeft - paddingRight);
        double startBaselineY = cellY + cellHeight - baselineInset - paddingTop;
        double cursorY = startBaselineY;
        var lines = new List<DocxTextLineLayout>();
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
                cursorY -= MeasureNestedTableHeight(tableElement.Table, textWidth, textMeasurer, defaultTabStopPoints, pageNumber, pageCount, paragraphSpacingScale);
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
            double fontSize = GetParagraphFontSize(paragraph);
            DocxLineHeightProfile lineHeightProfile = ResolveLineHeightProfile(paragraph, fontSize, textMeasurer);
            double lineHeight = lineHeightProfile.LineHeight;
            IReadOnlyList<DocxTextSpan> textSpans = CreateTextSpans(paragraph.Runs, pageNumber, pageCount);
            if (textSpans.Count == 0)
            {
                if (paragraph.Images.Count == 0)
                {
                    cursorY -= ResolveListLabelFirstLineExtraLeading(paragraph, fontSize, textMeasurer);
                    cursorY -= lineHeight;
                }
            }
            else
            {
                DocxTextRun firstRun = paragraph.Runs[0];
                double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, fontSize, textMeasurer);
                double continuationTextStartOffset = GetParagraphTextStartOffset(paragraph);
                double labelStartOffset = GetParagraphLabelStartOffset(paragraph);
                double paragraphX = cellX + paddingLeft + textStartOffset;
                double paragraphWidth = Math.Max(1d, textWidth - textStartOffset - GetParagraphRightInset(paragraph));
                double continuationParagraphWidth = Math.Max(1d, textWidth - continuationTextStartOffset - GetParagraphRightInset(paragraph));
                bool firstLine = true;
                DocxWrappedTextLine[] wrappedLines = WrapTextLines(
                    textSpans,
                    ResolveTableCellTextWrapWidth(cell, paragraphWidth),
                    ResolveTableCellTextWrapWidth(cell, continuationParagraphWidth),
                    fontSize,
                    textMeasurer,
                    paragraph.EffectiveProperties.TabStops,
                    defaultTabStopPoints,
                    allowOverwideTokenBreaks: true,
                    dynamicFieldPageNumber: pageNumber).ToArray();
                lineHeight = QuantizeTableCellWrappedLineHeight(lineHeight, wrappedLines.Length);
                for (int lineIndex = 0; lineIndex < wrappedLines.Length; lineIndex++)
                {
                    DocxWrappedTextLine line = wrappedLines[lineIndex];
                    if (firstLine)
                    {
                        cursorY -= ResolveListLabelFirstLineExtraLeading(paragraph, fontSize, textMeasurer);
                    }

                    double lineWidth = MeasureTextSpansForLayout(line.Spans, fontSize, textMeasurer, paragraph.EffectiveProperties.TabStops, defaultTabStopPoints, pageNumber);
                    double lineX = paragraph.EffectiveProperties.Alignment switch
                    {
                        DocxTextAlignment.Center => paragraphX + Math.Max(0, paragraphWidth - lineWidth) / 2d,
                        DocxTextAlignment.Right => paragraphX + Math.Max(0, paragraphWidth - lineWidth),
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
                        textMeasurer,
                        paragraph.EffectiveProperties.TabStops,
                        defaultTabStopPoints,
                        pageNumber);
                    lineShape = FitTableCellLineText(lineShape, paragraphWidth);
                    lines.Add(new DocxTextLineLayout(
                        lineShape.Text,
                        firstRun,
                        fontSize,
                        lineShape.X,
                        cursorY,
                        lineShape.Width,
                        lineShape.Segments,
                        SourceBlockIndex: null,
                        SourceParagraphIndex: paragraphIndex,
                        SourceLineIndex: lineIndex,
                        StoryKind: "TableCell",
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
                        SourceParagraph: paragraph, StoryVariantType: null, EmitsTerminalParagraphMark: false));
                    firstLine = false;
                    paragraphX = cellX + paddingLeft + continuationTextStartOffset;
                    paragraphWidth = Math.Max(1d, textWidth - continuationTextStartOffset - GetParagraphRightInset(paragraph));
                    cursorY -= lineHeight;
                }
            }

            pendingSpacingAfter = spacingProfile.ParagraphAfterSpacing;
            previousParagraph = paragraph;
            paragraphIndex++;
        }

        cursorY -= pendingSpacingAfter;
        if (lines.Count == 0)
        {
            return lines;
        }

        double usedHeight = Math.Max(0d, startBaselineY - cursorY);
        double availableHeight = Math.Max(0d, cellHeight - paddingTop - paddingBottom - baselineInset);
        double extra = Math.Max(0d, availableHeight - usedHeight);
        double verticalOffset = cell.VerticalAlignmentValue?.Equals("bottom", StringComparison.OrdinalIgnoreCase) == true
            ? extra
            : cell.VerticalAlignmentValue?.Equals("center", StringComparison.OrdinalIgnoreCase) == true
                ? extra / 2d
                : 0d;
        return verticalOffset == 0d ? lines : ShiftTextLines(lines, -verticalOffset, 0d);

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

    private static IReadOnlyList<DocxInlineImageLayout> LayoutTableCellInlineImages(
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
        if (paragraphs.Count == 0 || !paragraphs.Any(paragraph => paragraph.Images.Count != 0))
        {
            return [];
        }

        double paddingLeft = ResolveTableCellHorizontalPadding(cell.Margins.LeftPoints) + ResolveTableCellBorderContentInset(cell, "left");
        double paddingRight = ResolveTableCellHorizontalPadding(cell.Margins.RightPoints) + ResolveTableCellBorderContentInset(cell, "right");
        double paddingTop = rowTopPadding + ResolveTableCellTopBorderContentInset(cell);
        double paddingBottom = ResolveTableCellVerticalPadding(cell.Margins.BottomPoints);
        double baselineInset = ResolveTableCellFirstBaselineInset(paragraphs);
        double textWidth = Math.Max(1d, cellWidth - paddingLeft - paddingRight);
        double startBaselineY = cellY + cellHeight - baselineInset - paddingTop;
        double cursorY = startBaselineY;
        var images = new List<DocxInlineImageLayout>();
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
                if (textMeasurer is not null)
                {
                    cursorY -= MeasureNestedTableHeight(tableElement.Table, textWidth, textMeasurer, defaultTabStopPoints, pageNumber, pageCount, paragraphSpacingScale);
                }

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
            if (textMeasurer is not null)
            {
                double fontSize = GetParagraphFontSize(paragraph);
                double lineHeight = ResolveLineHeight(paragraph, fontSize, textMeasurer);
                IReadOnlyList<DocxTextSpan> textSpans = CreateTextSpans(paragraph.Runs, pageNumber, pageCount);
                if (textSpans.Count != 0)
                {
                    double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, fontSize, textMeasurer);
                    double firstParagraphWidth = ResolveTableCellTextWrapWidth(cell, textWidth - textStartOffset - GetParagraphRightInset(paragraph));
                    double continuationParagraphWidth = ResolveTableCellTextWrapWidth(cell, textWidth - GetParagraphTextStartOffset(paragraph) - GetParagraphRightInset(paragraph));
                    cursorY -= WrapTextLines(textSpans, firstParagraphWidth, continuationParagraphWidth, fontSize, textMeasurer, paragraph.EffectiveProperties.TabStops, defaultTabStopPoints, allowOverwideTokenBreaks: true, dynamicFieldPageNumber: pageNumber).Count() * lineHeight;
                }
                else if (paragraph.Images.Count == 0)
                {
                    cursorY -= lineHeight;
                }
            }

            foreach (DocxInlineImage image in paragraph.Images)
            {
                double paragraphX = cellX + paddingLeft + GetParagraphStartOffset(paragraph);
                double paragraphWidth = Math.Max(1d, textWidth - GetParagraphStartOffset(paragraph) - GetParagraphRightInset(paragraph));
                double imageWidth = Math.Min(paragraphWidth, image.WidthPoints);
                double imageHeight = image.HeightPoints * imageWidth / Math.Max(1d, image.WidthPoints);
                double imageX = paragraph.EffectiveProperties.Alignment switch
                {
                    DocxTextAlignment.Center => paragraphX + Math.Max(0, paragraphWidth - imageWidth) / 2d,
                    DocxTextAlignment.Right => paragraphX + Math.Max(0, paragraphWidth - imageWidth),
                    _ => paragraphX
                };
                images.Add(new DocxInlineImageLayout(image, imageX, cursorY - imageHeight, imageWidth, imageHeight, pageIndex, null, SourceParagraphIndex: paragraphIndex, StoryKind: null, StoryVariantType: null));
                cursorY -= imageHeight + InlineImageParagraphGapPoints;
            }

            pendingSpacingAfter = spacingProfile.ParagraphAfterSpacing;
            previousParagraph = paragraph;
            paragraphIndex++;
        }

        cursorY -= pendingSpacingAfter;
        if (images.Count == 0)
        {
            return images;
        }

        double usedHeight = Math.Max(0d, startBaselineY - cursorY);
        double availableHeight = Math.Max(0d, cellHeight - paddingTop - paddingBottom - baselineInset);
        double extra = Math.Max(0d, availableHeight - usedHeight);
        double verticalOffset = cell.VerticalAlignmentValue?.Equals("bottom", StringComparison.OrdinalIgnoreCase) == true
            ? extra
            : cell.VerticalAlignmentValue?.Equals("center", StringComparison.OrdinalIgnoreCase) == true
                ? extra / 2d
                : 0d;
        return verticalOffset == 0d
            ? images
            : images.Select(image => image with { Y = image.Y - verticalOffset }).ToArray();
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
        double paragraphSpacingScale)
    {
        IReadOnlyList<DocxBodyElement> bodyElements = GetTableCellLayoutBodyElements(cell);
        if (textMeasurer is null || !bodyElements.OfType<DocxTableElement>().Any())
        {
            return [];
        }

        double paddingLeft = ResolveTableCellHorizontalPadding(cell.Margins.LeftPoints) + ResolveTableCellBorderContentInset(cell, "left");
        double paddingRight = ResolveTableCellHorizontalPadding(cell.Margins.RightPoints) + ResolveTableCellBorderContentInset(cell, "right");
        double textWidth = Math.Max(1d, cellWidth - paddingLeft - paddingRight);
        double cursorY = cellY + cellHeight - rowTopPadding - ResolveTableCellTopBorderContentInset(cell);
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
                        StoryKind: null,
                        StoryVariantType: null,
                        pageCount: pageCount,
                        paragraphSpacingScale: paragraphSpacingScale));
                    cursorY -= rowHeight;
                }

                nestedTableIndex++;
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
            cursorY -= MeasureTableCellParagraphContentHeight(cell, paragraph, textWidth, textMeasurer, defaultTabStopPoints, pageNumber, pageCount);
            pendingSpacingAfter = spacingProfile.ParagraphAfterSpacing;
            previousParagraph = paragraph;
        }

        return nestedRows;
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
        IDocxTextMeasurer textMeasurer,
        double defaultTabStopPoints,
        int? pageNumber,
        int? pageCount)
    {
        double height = 0d;
        double fontSize = GetParagraphFontSize(paragraph);
        double lineHeight = ResolveLineHeight(paragraph, fontSize, textMeasurer);
        IReadOnlyList<DocxTextSpan> textSpans = CreateTextSpans(paragraph.Runs, pageNumber, pageCount);
        if (textSpans.Count != 0)
        {
            double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, fontSize, textMeasurer);
            double firstParagraphWidth = ResolveTableCellTextWrapWidth(cell, textWidth - textStartOffset - GetParagraphRightInset(paragraph));
            double continuationParagraphWidth = ResolveTableCellTextWrapWidth(cell, textWidth - GetParagraphTextStartOffset(paragraph) - GetParagraphRightInset(paragraph));
            height += WrapTextLines(textSpans, firstParagraphWidth, continuationParagraphWidth, fontSize, textMeasurer, paragraph.EffectiveProperties.TabStops, defaultTabStopPoints, allowOverwideTokenBreaks: true, dynamicFieldPageNumber: pageNumber).Count() * lineHeight;
        }
        else if (paragraph.Images.Count == 0)
        {
            height += lineHeight;
        }

        foreach (DocxInlineImage image in paragraph.Images)
        {
            double imageWidth = Math.Min(textWidth, image.WidthPoints);
            double imageHeight = image.HeightPoints * imageWidth / Math.Max(1d, image.WidthPoints);
            height += imageHeight + InlineImageParagraphGapPoints;
        }

        return height;
    }
}
