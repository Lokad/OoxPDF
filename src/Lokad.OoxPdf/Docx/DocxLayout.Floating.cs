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
    private static IReadOnlyList<DocxFloatingDrawingLayout> CreateRelatedStoryFloatingDrawingLayouts(
        DocxRelatedStory story,
        double bodyWidth,
        IReadOnlyList<DocxTextLineLayout> textLines,
        IReadOnlyList<DocxInlineImageLayout> inlineImages,
        IReadOnlyList<DocxTableRowLayout> tableRows,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        double paragraphSpacingScale,
        CancellationToken cancellationToken,
        int? pageNumber,
        int? pageCount)
    {
        if (story.FloatingDrawings.Count == 0)
        {
            return [];
        }

        DocxLayoutPage storyCanvas = CreateRelatedStoryLayoutCanvas();
        return story.FloatingDrawings
            .Select(drawing =>
            {
                DocxLayoutSourceBlockBounds? sourceBlock = drawing.SourceBlockIndex is null
                    ? null
                    : FindSourceBlockBounds([storyCanvas], drawing.SourceBlockIndex.Value);
                return CreateFloatingDrawingLayout(
                    drawing,
                    storyCanvas,
                    pageStartIndex: null,
                    pageEndIndex: null,
                    anchorPageIndex: null,
                    anchorColumnIndex: sourceBlock?.FirstColumnIndex,
                    sourceBlock,
                    storyKind: story.Kind.ToValueString(),
                    storyVariantType: null,
                    pageCount: pageCount,
                    textMeasurer: textMeasurer,
                    defaultTabStopPoints: defaultTabStopPoints,
                    paragraphSpacingScale: paragraphSpacingScale,
                    cancellationToken: cancellationToken);
            })
            .ToArray();

        DocxLayoutPage CreateRelatedStoryLayoutCanvas()
        {
            IReadOnlyList<DocxLayoutItem> items = textLines
                .Cast<DocxLayoutItem>()
                .Concat(inlineImages)
                .Concat(tableRows)
                .ToArray();
            return new DocxLayoutPage(
                bodyWidth,
                UnpagedRelatedStoryCanvasHeightPoints,
                0d,
                0d,
                0d,
                0d,
                0d,
                DocxPageSettings.Empty,
                new DocxSectionLayoutProperties(null, null, null, null, null, null, []),
                [new DocxLayoutColumnFrame(0, 0d, bodyWidth, null)],
                [],
                [],
                [],
                [],
                items);
        }
    }

    private static IReadOnlyList<DocxTextLineLayout> LayoutRelatedStoryParagraphTextLines(
        DocxParagraph paragraph,
        int sourceBlockIndex,
        int sourceParagraphIndex,
        string storyKind,
        double bodyWidth,
        double cursorY,
        DocxParagraphSpacingProfile spacingProfile,
        IDocxTextMeasurer textMeasurer,
        double defaultTabStopPoints,
        int? pageNumber,
        int? pageCount)
    {
        IReadOnlyList<DocxTextSpan> textSpans = CreateTextSpans(paragraph.Runs, pageNumber, pageCount);
        if (textSpans.Count == 0)
        {
            return [];
        }

        double fontSize = GetParagraphFontSize(paragraph);
        DocxLineHeightProfile lineHeightProfile = ResolveLineHeightProfile(paragraph, fontSize, textMeasurer);
        double lineHeight = lineHeightProfile.LineHeight;
        DocxEffectiveParagraphProperties effective = paragraph.EffectiveProperties;
        double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, fontSize, textMeasurer);
        double continuationTextStartOffset = GetParagraphTextStartOffset(paragraph);
        double labelStartOffset = GetParagraphLabelStartOffset(paragraph);
        double paragraphX = textStartOffset;
        double paragraphWidth = Math.Max(1d, bodyWidth - textStartOffset - GetParagraphRightInset(paragraph));
        double continuationParagraphWidth = Math.Max(1d, bodyWidth - continuationTextStartOffset - GetParagraphRightInset(paragraph));
        DocxTextRun firstRun = paragraph.Runs[0];
        DocxWrappedTextLine[] lines = WrapTextLines(textSpans, paragraphWidth, continuationParagraphWidth, fontSize, textMeasurer, effective.TabStops, defaultTabStopPoints, allowOverwideTokenBreaks: ShouldAllowCharacterLevelWordWrap(paragraph), dynamicFieldPageNumber: pageNumber).ToArray();
        var layouts = new List<DocxTextLineLayout>(lines.Length);
        bool firstLine = true;
        for (int lineIndex = 0; lineIndex < lines.Length; lineIndex++)
        {
            DocxWrappedTextLine line = lines[lineIndex];
            double lineWidth = MeasureTextSpansForLayout(line.Spans, fontSize, textMeasurer, effective.TabStops, defaultTabStopPoints, pageNumber);
            double lineX = effective.Alignment switch
            {
                DocxTextAlignment.Center => paragraphX + Math.Max(0, paragraphWidth - lineWidth) / 2d,
                DocxTextAlignment.Right => paragraphX + Math.Max(0, paragraphWidth - lineWidth),
                _ => paragraphX
            };
            double baselineOffset = DocxLineMetrics.ResolveBodyBaselineOffset(fontSize, lineHeight, IsExactLineSpacing(effective));
            DocxParagraphLineShape lineShape = CreateParagraphLineShape(
                paragraph,
                line,
                firstRun,
                firstLine,
                lineIndex == lines.Length - 1,
                labelStartOffset,
                lineX,
                paragraphWidth,
                fontSize,
                textMeasurer,
                effective.TabStops,
                defaultTabStopPoints,
                pageNumber);
            layouts.Add(new DocxTextLineLayout(
                lineShape.Text,
                firstRun,
                fontSize,
                lineShape.X,
                cursorY - baselineOffset,
                lineShape.Width,
                lineShape.Segments,
                SourceBlockIndex: sourceBlockIndex,
                SourceParagraphIndex: sourceParagraphIndex,
                SourceLineIndex: lineIndex,
                StoryKind: storyKind,
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
            paragraphX = continuationTextStartOffset;
            paragraphWidth = Math.Max(1d, bodyWidth - continuationTextStartOffset - GetParagraphRightInset(paragraph));
            cursorY -= lineHeight;
        }

        return layouts;
    }

    private static IReadOnlyList<DocxFloatingDrawingLayout> CreateFloatingDrawingLayouts(
        IReadOnlyList<DocxFloatingDrawing> drawings,
        IReadOnlyList<DocxLayoutPage> pages,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        double paragraphSpacingScale,
        CancellationToken cancellationToken)
    {
        var layouts = new DocxFloatingDrawingLayout[drawings.Count];
        for (int i = 0; i < drawings.Count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxFloatingDrawing drawing = drawings[i];
            DocxLayoutSourceBlockBounds? sourceBlock = drawing.SourceBlockIndex is null
                ? null
                : FindSourceBlockBounds(pages, drawing.SourceBlockIndex.Value);
            DocxLayoutPage? anchorPage = sourceBlock is null
                ? pages.FirstOrDefault()
                : pages[sourceBlock.FirstPageIndex];
            layouts[i] = CreateFloatingDrawingLayout(
                drawing,
                anchorPage,
                sourceBlock?.FirstPageIndex,
                sourceBlock?.LastPageIndex,
                sourceBlock?.FirstPageIndex,
                sourceBlock?.FirstColumnIndex,
                sourceBlock,
                storyKind: null,
                storyVariantType: null,
                pageCount: null,
                textMeasurer: textMeasurer,
                defaultTabStopPoints: defaultTabStopPoints,
                paragraphSpacingScale: paragraphSpacingScale,
                cancellationToken: cancellationToken);
        }

        return layouts;
    }

    private static IReadOnlyList<DocxFloatingDrawingLayout> CreateStaticFloatingDrawingLayouts(
        IReadOnlyList<DocxLayoutPage> pages,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        double paragraphSpacingScale,
        CancellationToken cancellationToken)
    {
        var layouts = new List<DocxFloatingDrawingLayout>();
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxLayoutPage page = pages[pageIndex];
            int pageNumber = pageIndex + 1;
            DocxSelectedStaticDrawings selectedHeader = SelectStaticHeaderFooterDrawings(
                page.PageSettings.HeaderFloatingDrawingsByType,
                page.PageSettings,
                pageNumber);
            DocxSelectedStaticDrawings selectedFooter = SelectStaticHeaderFooterDrawings(
                page.PageSettings.FooterFloatingDrawingsByType,
                page.PageSettings,
                pageNumber);
            layouts.AddRange(CreateStaticFloatingDrawingLayouts(selectedHeader, "Header", page, pageIndex, pages.Count, textMeasurer, defaultTabStopPoints, paragraphSpacingScale, cancellationToken));
            layouts.AddRange(CreateStaticFloatingDrawingLayouts(selectedFooter, "Footer", page, pageIndex, pages.Count, textMeasurer, defaultTabStopPoints, paragraphSpacingScale, cancellationToken));
        }

        return layouts;
    }

    private static IEnumerable<DocxFloatingDrawingLayout> CreateStaticFloatingDrawingLayouts(
        DocxSelectedStaticDrawings selectedDrawings,
        string storyKind,
        DocxLayoutPage page,
        int pageIndex,
        int pageCount,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        double paragraphSpacingScale,
        CancellationToken cancellationToken)
    {
        foreach (DocxFloatingDrawing drawing in selectedDrawings.Drawings)
        {
            cancellationToken.ThrowIfCancellationRequested();
            yield return CreateFloatingDrawingLayout(
                drawing,
                page,
                pageStartIndex: pageIndex,
                pageEndIndex: pageIndex,
                anchorPageIndex: pageIndex,
                anchorColumnIndex: null,
                sourceBlock: null,
                storyKind: storyKind,
                storyVariantType: selectedDrawings.VariantType,
                pageCount: pageCount,
                textMeasurer: textMeasurer,
                defaultTabStopPoints: defaultTabStopPoints,
                paragraphSpacingScale: paragraphSpacingScale,
                cancellationToken: cancellationToken);
        }
    }

    private static DocxFloatingDrawingLayout CreateFloatingDrawingLayout(
        DocxFloatingDrawing drawing,
        DocxLayoutPage? anchorPage,
        int? pageStartIndex,
        int? pageEndIndex,
        int? anchorPageIndex,
        int? anchorColumnIndex,
        DocxLayoutSourceBlockBounds? sourceBlock,
        string? storyKind,
        string? storyVariantType,
        int? pageCount,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        double paragraphSpacingScale,
        CancellationToken cancellationToken)
    {
        DocxAnchorReferenceFrame? horizontalReference = ResolveHorizontalReferenceFrame(drawing, anchorPage, sourceBlock);
        DocxAnchorReferenceFrame? verticalReference = ResolveVerticalReferenceFrame(drawing, anchorPage, sourceBlock);
        double? extentWidth = ReadEmuPoints(drawing.ExtentCxValue);
        double? extentHeight = ReadEmuPoints(drawing.ExtentCyValue);
        double? horizontalOffset = ReadEmuPoints(drawing.HorizontalOffsetValue);
        double? verticalOffset = ReadEmuPoints(drawing.VerticalOffsetValue);
        double? distanceTop = ReadEmuPoints(drawing.DistanceTopValue);
        double? distanceBottom = ReadEmuPoints(drawing.DistanceBottomValue);
        double? distanceLeft = ReadEmuPoints(drawing.DistanceLeftValue);
        double? distanceRight = ReadEmuPoints(drawing.DistanceRightValue);
        DocxAnchorPlacement horizontalPlacement = ResolveHorizontalPlacement(horizontalReference, extentWidth, drawing.HorizontalAlignValue, horizontalOffset);
        DocxAnchorPlacement verticalPlacement = ResolveVerticalPlacement(verticalReference, extentHeight, drawing.VerticalAlignValue, verticalOffset);
        double? placedX = horizontalPlacement.Position;
        double? placedTop = verticalPlacement.Position;
        DocxWrapExclusionFrame? wrapExclusion = CreateWrapExclusionFrame();
        DocxRelatedStoryLayout? textBoxLayout = CreateFloatingTextBoxLayout(
            anchorPageIndex is null ? null : anchorPageIndex.Value + 1);
        return new DocxFloatingDrawingLayout(
            drawing,
            pageStartIndex,
            pageEndIndex,
            anchorPageIndex,
            anchorColumnIndex,
            sourceBlock?.VerticalTop,
            sourceBlock?.VerticalBottom,
            extentWidth,
            extentHeight,
            horizontalOffset,
            verticalOffset,
            distanceTop,
            distanceBottom,
            distanceLeft,
            distanceRight,
            horizontalReference?.Start,
            horizontalReference?.Size,
            verticalReference?.Start,
            verticalReference?.End,
            placedX,
            placedTop,
            horizontalPlacement.Source,
            verticalPlacement.Source,
            wrapExclusion?.X,
            wrapExclusion?.Top,
            wrapExclusion?.Width,
            wrapExclusion?.Height,
            storyKind,
            storyVariantType,
            textBoxLayout);

        DocxRelatedStoryLayout? CreateFloatingTextBoxLayout(int? pageNumber)
        {
            if (textMeasurer is null ||
                drawing.TextBoxBodyElements.Count == 0 ||
                extentWidth is not { } width ||
                width <= 0d)
            {
                return null;
            }

            var story = new DocxRelatedStory(
                DocxRelatedStoryKind.TextBox,
                "floating-drawing",
                drawing.ImageRelationshipId,
                drawing.TextBoxBodyElements,
                [],
                [], null);
            return CreateRelatedStoryLayout(
                story,
                storyIndex: -1,
                Math.Max(1d, width),
                textMeasurer,
                defaultTabStopPoints,
                paragraphSpacingScale,
                pageNumber: pageNumber,
                pageCount: pageCount,
                cancellationToken: cancellationToken);
        }

        DocxWrapExclusionFrame? CreateWrapExclusionFrame()
        {
            bool IsWrapExclusionKind()
            {
                return drawing.WrapKind is DocxFloatingWrapKind.Square or DocxFloatingWrapKind.Tight or DocxFloatingWrapKind.Through or DocxFloatingWrapKind.TopAndBottom;
            }

            if (!IsWrapExclusionKind() ||
                placedX is not { } x ||
                placedTop is not { } top ||
                extentWidth is not { } width ||
                extentHeight is not { } height)
            {
                return null;
            }

            double leftDistance = distanceLeft ?? 0d;
            double rightDistance = distanceRight ?? 0d;
            double topDistance = distanceTop ?? 0d;
            double bottomDistance = distanceBottom ?? 0d;
            return new DocxWrapExclusionFrame(
                x - leftDistance,
                top + topDistance,
                width + leftDistance + rightDistance,
                height + topDistance + bottomDistance);
        }
    }

    private static DocxAnchorPlacement ResolveHorizontalPlacement(
        DocxAnchorReferenceFrame? reference,
        double? extentWidth,
        string? alignValue,
        double? offset)
    {
        if (reference is null || extentWidth is null)
        {
            return new DocxAnchorPlacement(null, DocxAnchorPlacementSource.MissingReferenceOrExtent);
        }

        return alignValue?.ToLowerInvariant() switch
        {
            "left" => new DocxAnchorPlacement(reference.Start, DocxAnchorPlacementSource.Align),
            "center" => new DocxAnchorPlacement(reference.Start + Math.Max(0d, reference.Size - extentWidth.Value) / 2d, DocxAnchorPlacementSource.Align),
            "right" => new DocxAnchorPlacement(reference.End - extentWidth.Value, DocxAnchorPlacementSource.Align),
            null when offset is not null => new DocxAnchorPlacement(reference.Start + offset.Value, DocxAnchorPlacementSource.Offset),
            "" when offset is not null => new DocxAnchorPlacement(reference.Start + offset.Value, DocxAnchorPlacementSource.Offset),
            _ => new DocxAnchorPlacement(null, DocxAnchorPlacementSource.Unsupported)
        };
    }

    private static DocxAnchorPlacement ResolveVerticalPlacement(
        DocxAnchorReferenceFrame? reference,
        double? extentHeight,
        string? alignValue,
        double? offset)
    {
        if (reference is null || extentHeight is null)
        {
            return new DocxAnchorPlacement(null, DocxAnchorPlacementSource.MissingReferenceOrExtent);
        }

        return alignValue?.ToLowerInvariant() switch
        {
            "top" => new DocxAnchorPlacement(reference.Start, DocxAnchorPlacementSource.Align),
            "center" => new DocxAnchorPlacement(reference.End + (reference.Size + extentHeight.Value) / 2d, DocxAnchorPlacementSource.Align),
            "bottom" => new DocxAnchorPlacement(reference.End + extentHeight.Value, DocxAnchorPlacementSource.Align),
            null when offset is not null => new DocxAnchorPlacement(reference.Start - offset.Value, DocxAnchorPlacementSource.Offset),
            "" when offset is not null => new DocxAnchorPlacement(reference.Start - offset.Value, DocxAnchorPlacementSource.Offset),
            _ => new DocxAnchorPlacement(null, DocxAnchorPlacementSource.Unsupported)
        };
    }

    private static DocxAnchorReferenceFrame? ResolveHorizontalReferenceFrame(
        DocxFloatingDrawing drawing,
        DocxLayoutPage? page,
        DocxLayoutSourceBlockBounds? sourceBlock)
    {
        if (page is null)
        {
            return null;
        }

        return drawing.HorizontalRelativeFromValue?.ToLowerInvariant() switch
        {
            "page" => new DocxAnchorReferenceFrame(0d, page.Width),
            "margin" => new DocxAnchorReferenceFrame(page.MarginLeft, page.Width - page.MarginRight),
            "column" when page.ColumnFrames.Count == 1 => new DocxAnchorReferenceFrame(page.ColumnFrames[0].X, page.ColumnFrames[0].X + page.ColumnFrames[0].Width),
            "column" when sourceBlock?.FirstColumnIndex is { } columnIndex && columnIndex >= 0 && columnIndex < page.ColumnFrames.Count =>
                new DocxAnchorReferenceFrame(page.ColumnFrames[columnIndex].X, page.ColumnFrames[columnIndex].X + page.ColumnFrames[columnIndex].Width),
            _ => null
        };
    }

    private static DocxAnchorReferenceFrame? ResolveVerticalReferenceFrame(
        DocxFloatingDrawing drawing,
        DocxLayoutPage? page,
        DocxLayoutSourceBlockBounds? sourceBlock)
    {
        if (page is null)
        {
            return null;
        }

        return drawing.VerticalRelativeFromValue?.ToLowerInvariant() switch
        {
            "page" => new DocxAnchorReferenceFrame(page.Height, 0d),
            "margin" => new DocxAnchorReferenceFrame(page.Height - page.MarginTop, page.MarginBottom),
            "paragraph" when sourceBlock is not null => new DocxAnchorReferenceFrame(sourceBlock.VerticalTop, sourceBlock.VerticalBottom),
            _ => null
        };
    }

    private static double? ReadEmuPoints(string? value)
    {
        return long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long emu)
            ? OoxUnits.EmuToPoints(emu)
            : null;
    }

    private static DocxLayoutSourceBlockBounds? FindSourceBlockBounds(IReadOnlyList<DocxLayoutPage> pages, int sourceBlockIndex)
    {
        int? firstPageIndex = null;
        int? lastPageIndex = null;
        int? firstColumnIndex = null;
        int? lastColumnIndex = null;
        double verticalTop = double.NegativeInfinity;
        double verticalBottom = double.PositiveInfinity;
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++)
        {
            DocxLayoutPage page = pages[pageIndex];
            foreach (DocxLayoutItem item in pages[pageIndex].Items)
            {
                if (GetSourceBlockIndex(item) != sourceBlockIndex)
                {
                    continue;
                }

                firstPageIndex ??= pageIndex;
                lastPageIndex = pageIndex;
                if (ResolveItemColumnIndex(page, item) is { } columnIndex)
                {
                    firstColumnIndex = firstColumnIndex is null ? columnIndex : Math.Min(firstColumnIndex.Value, columnIndex);
                    lastColumnIndex = lastColumnIndex is null ? columnIndex : Math.Max(lastColumnIndex.Value, columnIndex);
                }

                (double y, double height) = GetVerticalBounds(item);
                verticalTop = Math.Max(verticalTop, y + height);
                verticalBottom = Math.Min(verticalBottom, y);
            }
        }

        return firstPageIndex is null || lastPageIndex is null
            ? null
            : new DocxLayoutSourceBlockBounds(firstPageIndex.Value, lastPageIndex.Value, firstColumnIndex, lastColumnIndex, verticalTop, verticalBottom);
    }

    private static int? GetSourceBlockIndex(DocxLayoutItem item)
    {
        return item switch
        {
            DocxTextLineLayout text => text.SourceBlockIndex,
            DocxInlineImageLayout image => image.SourceBlockIndex,
            DocxTableRowLayout row => row.Table.SourceBlockIndex,
            _ => null
        };
    }

    private static (double Y, double Height) GetVerticalBounds(DocxLayoutItem item)
    {
        return item switch
        {
            DocxTextLineLayout text => (text.BaselineY, text.FontSize),
            DocxInlineImageLayout image => (image.Y, image.Height),
            DocxTableRowLayout row => (row.Y, row.Height),
            _ => (0d, 0d)
        };
    }

    private static int? ResolveItemColumnIndex(DocxLayoutPage page, DocxLayoutItem item)
    {
        return ResolveItemColumnIndex(page.ColumnFrames, item);
    }

    private static int? ResolveItemColumnIndex(IReadOnlyList<DocxLayoutColumnFrame> frames, DocxLayoutItem item)
    {
        (double x, double width) = item switch
        {
            DocxTextLineLayout text => (text.X, text.Width),
            DocxInlineImageLayout image => (image.X, image.Width),
            DocxTableRowLayout row => (row.Table.TableX, row.Table.ResolvedTableWidth),
            _ => (0d, 0d)
        };
        return DocxLayoutColumnOwnership.ResolveColumnIndex(frames, x, width);
    }

    private sealed record DocxLayoutSourceBlockBounds(
        int FirstPageIndex,
        int LastPageIndex,
        int? FirstColumnIndex,
        int? LastColumnIndex,
        double VerticalTop,
        double VerticalBottom);

    private sealed record DocxAnchorReferenceFrame(double Start, double End)
    {
        public double Size => Math.Abs(Start - End);
    }

    private sealed record DocxStaticStoryLayoutResult(
        IReadOnlyList<DocxTextLineLayout> TextLines,
        IReadOnlyList<DocxInlineImageLayout> InlineImages,
        IReadOnlyList<DocxTableRowLayout> TableRows);
}
