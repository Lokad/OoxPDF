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
        int? pageCount,
        IDocxTextMeasurer? unscaledTextMeasurer = null)
    {
        if (story.FloatingDrawings.Count == 0)
        {
            return [];
        }

        DocxLayoutPage storyCanvas = CreateRelatedStoryLayoutCanvas();
        IReadOnlyDictionary<int, DocxLayoutSourceBlockBounds> sourceBlocks = BuildSourceBlockIndex([storyCanvas], cancellationToken);
        return story.FloatingDrawings
            .Select(drawing =>
            {
                DocxLayoutSourceBlockBounds? sourceBlock = drawing.SourceBlockIndex is { } storyBlockIndex
                    && sourceBlocks.TryGetValue(storyBlockIndex, out DocxLayoutSourceBlockBounds? storyFound)
                    ? storyFound
                    : null;
                return CreateFloatingDrawingLayout(
                    drawing,
                    storyCanvas,
                    pageStartIndex: null,
                    pageEndIndex: null,
                    anchorPageIndex: null,
                    anchorColumnIndex: sourceBlock?.FirstColumnIndex,
                    sourceBlock,
                    story: DocxStoryId.Related(story.Kind),
                    pageCount: pageCount,
                    textMeasurer: textMeasurer,
                    defaultTabStopPoints: defaultTabStopPoints,
                    paragraphSpacingScale: paragraphSpacingScale,
                    cancellationToken: cancellationToken,
                    unscaledTextMeasurer: unscaledTextMeasurer);
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

    private static double? ResolveRelatedStoryFirstInsetPoints(DocxParagraph paragraph, IDocxTextMeasurer textMeasurer, IDocxHheaLineGapProvider gapProvider)
    {
        if (textMeasurer is not IDocxLineMetricsProvider lineMetrics || paragraph.Runs.Count == 0)
        {
            return null;
        }

        // RV06 take battery (Word 16.0, a12 21/21/18): mixed notes keep the legacy
        // floor so take capacities hold, while uniform notes resolve their own summed
        // inset; a single distinct family-plus-size keeps the new path order-blind.
        string? insetFamily = null;
        double insetSize = 0d;
        bool uniformRuns = true;
        foreach (DocxTextRun run in paragraph.Runs)
        {
            if (string.IsNullOrWhiteSpace(run.Text) || run.EffectiveProperties.Hidden)
            {
                continue;
            }

            double size = run.EffectiveProperties.FontSize;
            if (insetFamily is null)
            {
                insetFamily = run.FontFamily ?? string.Empty;
                insetSize = size;
            }
            else if (!string.Equals(insetFamily, run.FontFamily ?? string.Empty, StringComparison.Ordinal) || insetSize != size)
            {
                uniformRuns = false;
                break;
            }
        }

        if (!uniformRuns || insetFamily is null)
        {
            return null;
        }

        foreach (DocxTextRun run in paragraph.Runs)
        {
            if (string.IsNullOrWhiteSpace(run.Text) || run.EffectiveProperties.Hidden)
            {
                continue;
            }

            // RV06 gap slice (Palatino/Algerian first baselines) plus Calibri-class tie
            // (content single-em 1.220703125 against reference 1.2207): gap-overflowing notes
            // keep the legacy floor, so the new path applies only while content single-em
            // stays at or below the mark reference within float representation.
            if (insetSize > 0d && lineMetrics.MeasureSingleLineHeight(run, insetSize) / insetSize > EndnoteSeparatorReferenceSingleLineEm + 1e-4)
            {
                return null;
            }

            if (textMeasurer is IDocxStaticTextMetricsProvider staticMetrics &&
                textMeasurer is IDocxHheaDescenderProvider)
            {
                // RV06 descender probes (Word 16.0, Informal Roman 10/12/14pt size matrix
                // plus Consolas, Magneto and Calibri resolved sheets): Office bottom bearing
                // follows the Windows descender even when the hhea descender runs deeper,
                // so the inset is the hhea box minus the Windows descender; the hhea
                // provider gate stays since full-metric measurers serve both faces together.
                return lineMetrics.MeasureHheaLineHeight(run, insetSize) - staticMetrics.MeasureWindowsDescender(run, insetSize);
            }

            return lineMetrics.MeasureHheaAscender(run, insetSize) + gapProvider.MeasureHheaLineGap(run, insetSize);
        }

        return null;
    }

    private static (IReadOnlyList<DocxTextLineLayout> Lines, IReadOnlyList<DocxInlineImageLayout> PlacedImages, double UsedHeight, double BaselineOffset) LayoutRelatedStoryParagraphTextLines(
        DocxParagraph paragraph,
        double fixedScale,
        int sourceBlockIndex,
        int sourceParagraphIndex,
        DocxStoryId? story,
        double bodyWidth,
        double cursorY,
        DocxParagraphSpacingProfile spacingProfile,
        IDocxTextMeasurer textMeasurer,
        double defaultTabStopPoints,
        int? pageNumber,
        int? pageCount,
        double? firstStoryBaselineOffset = null,
        double transitionScale = 1d)
    {
        IReadOnlyList<DocxTextSpan> textSpans = CreateTextSpans(paragraph.Runs, pageNumber, pageCount);
        if (textSpans.Count == 0)
        {
            return (Array.Empty<DocxTextLineLayout>(), Array.Empty<DocxInlineImageLayout>(), 0d, 0d);
        }

        double fontSize = GetParagraphFontSize(paragraph);
        DocxLineHeightProfile lineHeightProfile = ResolveLineHeightProfile(paragraph, fontSize, textMeasurer, selectMaxAcrossRuns: story?.Kind is DocxStoryKind.Footnote or DocxStoryKind.Endnote);
        double lineHeight = lineHeightProfile.LineHeight;
        DocxEffectiveParagraphProperties effective = paragraph.EffectiveProperties;
        double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, fontSize, textMeasurer, fixedScale);
        double continuationTextStartOffset = GetParagraphTextStartOffset(paragraph, fixedScale);
        double labelStartOffset = GetParagraphLabelStartOffset(paragraph, fixedScale);
        double paragraphX = textStartOffset;
        double paragraphWidth = Math.Max(1d, bodyWidth - textStartOffset - GetParagraphRightInset(paragraph, fixedScale));
        double continuationParagraphWidth = Math.Max(1d, bodyWidth - continuationTextStartOffset - GetParagraphRightInset(paragraph, fixedScale));
        DocxTextRun firstRun = paragraph.Runs[0];
        DocxWrappedTextLine[] lines = WrapTextLines(textSpans, paragraphWidth, continuationParagraphWidth, fontSize, textMeasurer, ScaleTabStopPositions(effective.TabStops, fixedScale), defaultTabStopPoints * fixedScale, allowOverwideTokenBreaks: ShouldAllowCharacterLevelWordWrap(paragraph), dynamicFieldPageNumber: pageNumber, inlineImageWidths: ResolveInlineImageWrapWidths(paragraph, textSpans)).ToArray();
        // Auto insets take max(hheaAscender, 0.94em) for footnote/endnote stories
        // (Word 16.0 endnote grids); comment/textbox stories keep legacy (unprobed).
        double? storyHheaAscender = story?.Kind is DocxStoryKind.Footnote or DocxStoryKind.Endnote
            ? DocxLineMetrics.ResolveHheaAscenderPoints(paragraph, fontSize, textMeasurer, selectMaxHhea: false)
            : null;
        // RV06 shape probes (Word 16.0, eighteen families) plus take battery (a12
        // 21/21/18): uniform-note first insets follow hhea ascender plus hhea gap with
        // no 0.94em floor, while mixed notes keep the legacy floor so take capacities
        // hold; gap-aware measurers resolve uniform notes directly and every other path
        // stays byte-identical.
        double storyBaselineOffset = story?.Kind is DocxStoryKind.Footnote or DocxStoryKind.Endnote &&
            textMeasurer is IDocxHheaLineGapProvider gapProvider &&
            ResolveRelatedStoryFirstInsetPoints(paragraph, textMeasurer, gapProvider) is double relatedFirstInset
            ? relatedFirstInset
            : DocxLineMetrics.ResolveBodyBaselineOffset(fontSize, lineHeight, IsExactLineSpacing(effective), storyHheaAscender);
        // RV05: ordered inline atoms (related-story path). Affined images in
        // text-mixed paragraphs attach to wrapped lines at run position.
        DocxMidLinePlan? storyMidLinePlan = CreateMidLinePlan(paragraph, textSpans, lines, paragraphWidth, continuationParagraphWidth, storyBaselineOffset, lineHeight);
        var layouts = new List<DocxTextLineLayout>(lines.Length);
        var placedImages = new List<DocxInlineImageLayout>();
        double startCursorY = cursorY;
        bool firstLine = true;
        for (int lineIndex = 0; lineIndex < lines.Length; lineIndex++)
        {
            DocxWrappedTextLine line = lines[lineIndex];
            double lineWidth = MeasureTextSpansForLayout(line.Spans, fontSize, textMeasurer, ScaleTabStopPositions(effective.TabStops, fixedScale), defaultTabStopPoints * fixedScale, pageNumber) + (storyMidLinePlan?.LineImageWidths[lineIndex] ?? 0d);
            // RV05 calibration (Word 16.0): image top pins to the natural line top.
            double extraAbove = IsExactLineSpacing(paragraph.EffectiveProperties) ? 0d : (storyMidLinePlan?.ShiftAboveHeights[lineIndex] ?? 0d);
            cursorY -= extraAbove;
            // RV06 footnote-align probe: Office centers/rights drawable related-story
            // text too, letting trailing spaces overflow past the edge.
            double lineAlignWidth = effective.Alignment is DocxTextAlignment.Center or DocxTextAlignment.Right
                ? MeasureDrawableTextSpansForLayout(line.Spans, fontSize, textMeasurer, ScaleTabStopPositions(effective.TabStops, fixedScale), defaultTabStopPoints * fixedScale, pageNumber) + (storyMidLinePlan?.LineImageWidths[lineIndex] ?? 0d)
                : lineWidth;
            double lineX = effective.Alignment switch
            {
                DocxTextAlignment.Center => paragraphX + Math.Max(0, paragraphWidth - lineAlignWidth) / 2d,
                DocxTextAlignment.Right => paragraphX + Math.Max(0, paragraphWidth - lineAlignWidth),
                _ => paragraphX
            };
            double baselineOffset = storyBaselineOffset;
            if (firstStoryBaselineOffset is { } firstStoryOffset)
            {
                // RV06 pagination probe (edge-mixedstory-wc, Word 16.0): under a
                // word-compatible print scale, Office scales whole pitches
                // including the font-size offset transition. Each line corrects
                // against the story-block-first offset; same-size lines self-zero
                // and scale 1.0 keeps legacy exactly.
                baselineOffset -= (baselineOffset - firstStoryOffset) * (1d - transitionScale);
            }
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
                ScaleTabStopPositions(effective.TabStops, fixedScale),
                defaultTabStopPoints * fixedScale,
                pageNumber);
            IReadOnlyList<DocxTextSegmentLayout> emissionSegments = lineShape.Segments;
            if (storyMidLinePlan is not null && storyMidLinePlan.ImagesByLine[lineIndex].Count != 0)
            {
                var imageShifts = new List<(double BoundaryX, double Shift)>();
                foreach (DocxMidLineImage placed in storyMidLinePlan.ImagesByLine[lineIndex])
                {
                    double shiftBeforeWidth = MeasureMidLineBeforeWidth(line.Spans, placed.LineCharOffset, paragraph, firstLine, lineIndex == lines.Length - 1, paragraphWidth, fontSize, textMeasurer, ScaleTabStopPositions(effective.TabStops, fixedScale), defaultTabStopPoints * fixedScale, pageNumber);
                    imageShifts.Add((shiftBeforeWidth, placed.Width));
                }

                emissionSegments = ShiftSegmentsPastMidLineImages(lineShape.Segments, line.Spans, paragraph, firstLine, lineX, imageShifts);
            }

            layouts.Add(new DocxTextLineLayout(
                lineShape.Text,
                firstRun,
                fontSize,
                lineShape.X,
                cursorY - baselineOffset,
                lineShape.Width,
                emissionSegments,
                SourceBlockIndex: sourceBlockIndex,
                SourceParagraphIndex: sourceParagraphIndex,
                SourceLineIndex: lineIndex,
                Story: story,
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
            if (storyMidLinePlan is not null)
            {
                foreach (DocxMidLineImage placed in storyMidLinePlan.ImagesByLine[lineIndex])
                {
                    double beforeWidth = MeasureMidLineBeforeWidth(line.Spans, placed.LineCharOffset, paragraph, firstLine, lineIndex == lines.Length - 1, paragraphWidth, fontSize, textMeasurer, ScaleTabStopPositions(effective.TabStops, fixedScale), defaultTabStopPoints * fixedScale, pageNumber);
                    placedImages.Add(new DocxInlineImageLayout(
                        placed.Image,
                        lineX + beforeWidth,
                        cursorY - baselineOffset - placed.Height,
                        placed.Width,
                        placed.Height,
                        PageIndex: 0,
                        SourceBlockIndex: sourceBlockIndex,
                        SourceParagraphIndex: sourceParagraphIndex,
                        Story: null));
                }
            }
            firstLine = false;
            paragraphX = continuationTextStartOffset;
            paragraphWidth = Math.Max(1d, bodyWidth - continuationTextStartOffset - GetParagraphRightInset(paragraph, fixedScale));
            cursorY -= lineHeight;
        }

        // RV06 footnote-align probe: related stories keep one row-end space beyond
        // authored trailing too.
        if (paragraph.Images.Count == 0 &&
            paragraph.InlineTextBoxes.Count == 0 &&
            lines.Length > 0 &&
            lines[^1].Text.EndsWith(' ') &&
            textSpans.Any(static span => span.Text.Any(static character => !char.IsWhiteSpace(character))))
        {
            for (int storyLineIndex = layouts.Count - 1; storyLineIndex >= 0; storyLineIndex--)
            {
                if (layouts[storyLineIndex] is not DocxTextLineLayout storyLastLine ||
                    !ReferenceEquals(storyLastLine.SourceParagraph, paragraph) ||
                    storyLastLine.Text.Length == 0)
                {
                    continue;
                }

                double storySpaceWidth = textMeasurer.MeasureText(firstRun, " ", fontSize);
                layouts[storyLineIndex] = storyLastLine with
                {
                    Text = storyLastLine.Text + " ",
                    Width = storyLastLine.Width + storySpaceWidth,
                    Segments =
                    [
                        .. storyLastLine.Segments,
                        new DocxTextSegmentLayout(
                            " ",
                            firstRun,
                            storyLastLine.X + storyLastLine.Width,
                            storySpaceWidth,
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

        return (layouts, placedImages, startCursorY - cursorY, storyBaselineOffset);
    }

    private static IReadOnlyList<DocxFloatingDrawingLayout> CreateFloatingDrawingLayouts(
        IReadOnlyList<DocxFloatingDrawing> drawings,
        IReadOnlyList<DocxLayoutPage> pages,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        double paragraphSpacingScale,
        CancellationToken cancellationToken,
        IDocxTextMeasurer? unscaledTextMeasurer = null)
    {
        var layouts = new DocxFloatingDrawingLayout[drawings.Count];
        IReadOnlyDictionary<int, DocxLayoutSourceBlockBounds>? sourceBlocks = null;
        for (int i = 0; i < drawings.Count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxFloatingDrawing drawing = drawings[i];
            DocxLayoutSourceBlockBounds? sourceBlock = null;
            if (drawing.SourceBlockIndex is { } blockIndex)
            {
                sourceBlocks ??= BuildSourceBlockIndex(pages, cancellationToken);
                sourceBlocks.TryGetValue(blockIndex, out sourceBlock);
            }
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
                story: null,
                pageCount: null,
                textMeasurer: textMeasurer,
                defaultTabStopPoints: defaultTabStopPoints,
                paragraphSpacingScale: paragraphSpacingScale,
                cancellationToken: cancellationToken,
                unscaledTextMeasurer: unscaledTextMeasurer);
        }

        return layouts;
    }

    private static IReadOnlyList<DocxFloatingDrawingLayout> CreateStaticFloatingDrawingLayouts(
        IReadOnlyList<DocxLayoutPage> pages,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        double paragraphSpacingScale,
        CancellationToken cancellationToken,
        IDocxTextMeasurer? unscaledTextMeasurer = null)
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
            layouts.AddRange(CreateStaticFloatingDrawingLayouts(selectedHeader, DocxStoryId.Header(selectedHeader.VariantType), page, pageIndex, pages.Count, textMeasurer, defaultTabStopPoints, paragraphSpacingScale, cancellationToken));
            layouts.AddRange(CreateStaticFloatingDrawingLayouts(selectedFooter, DocxStoryId.Footer(selectedFooter.VariantType), page, pageIndex, pages.Count, textMeasurer, defaultTabStopPoints, paragraphSpacingScale, cancellationToken));
        }

        return layouts;
    }

    private static IEnumerable<DocxFloatingDrawingLayout> CreateStaticFloatingDrawingLayouts(
        DocxSelectedStaticDrawings selectedDrawings,
        DocxStoryId story,
        DocxLayoutPage page,
        int pageIndex,
        int pageCount,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        double paragraphSpacingScale,
        CancellationToken cancellationToken,
        IDocxTextMeasurer? unscaledTextMeasurer = null)
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
                story: story,
                pageCount: pageCount,
                textMeasurer: textMeasurer,
                defaultTabStopPoints: defaultTabStopPoints,
                paragraphSpacingScale: paragraphSpacingScale,
                cancellationToken: cancellationToken,
                unscaledTextMeasurer: unscaledTextMeasurer);
        }
    }

    // Office defaults when bodyPr carries no explicit insets: 0.1in sides, 0.05in
    // top and bottom (w6a1d-tbx reference measures exactly these).
    private const long TextBoxDefaultHorizontalInsetEmu = 91440L;
    private const long TextBoxDefaultVerticalInsetEmu = 45720L;

    internal static void ResolveTextBoxContentInsets(DocxFloatingDrawing drawing, out double insetLeft, out double insetTop, out double insetRight, out double insetBottom)
    {
        // RV05 floatbox probe (Word COM reference edge-floatbox): stroked textbox
        // content starts inside the border, so every inset grows by half the outline
        // width (0.5pt stroke shifts text by 0.25 on both axes). Un-outlined shapes
        // resolve zero and keep legacy insets bit-identically.
        double halfStrokePoints = (ReadEmuPoints(drawing.TextBoxStrokeWidthEmuValue) ?? 0d) / 2d;
        insetLeft = ReadInsetEmuPoints(drawing.TextBoxInsetLeftValue, TextBoxDefaultHorizontalInsetEmu) + halfStrokePoints;
        insetTop = ReadInsetEmuPoints(drawing.TextBoxInsetTopValue, TextBoxDefaultVerticalInsetEmu) + halfStrokePoints;
        insetRight = ReadInsetEmuPoints(drawing.TextBoxInsetRightValue, TextBoxDefaultHorizontalInsetEmu) + halfStrokePoints;
        insetBottom = ReadInsetEmuPoints(drawing.TextBoxInsetBottomValue, TextBoxDefaultVerticalInsetEmu) + halfStrokePoints;
    }

    internal static void ResolveInlineTextBoxContentInsets(DocxInlineTextBox textBox, out double insetLeft, out double insetTop, out double insetRight, out double insetBottom)
    {
        // RV05 inline-box probe (Word COM reference edge-inlinebox): stroked inline
        // content starts inside the border like floating boxes. Un-outlined shapes
        // resolve zero and keep legacy insets bit-identically.
        double halfStrokePoints = (ReadEmuPoints(textBox.TextBoxStrokeWidthEmuValue) ?? 0d) / 2d;
        insetLeft = ReadInsetEmuPoints(textBox.TextBoxInsetLeftValue, TextBoxDefaultHorizontalInsetEmu) + halfStrokePoints;
        insetTop = ReadInsetEmuPoints(textBox.TextBoxInsetTopValue, TextBoxDefaultVerticalInsetEmu) + halfStrokePoints;
        insetRight = ReadInsetEmuPoints(textBox.TextBoxInsetRightValue, TextBoxDefaultHorizontalInsetEmu) + halfStrokePoints;
        insetBottom = ReadInsetEmuPoints(textBox.TextBoxInsetBottomValue, TextBoxDefaultVerticalInsetEmu) + halfStrokePoints;
    }

    // Inline boxes join the scaled body flow uniformly (Office: the w6-inline probe shows
    // the box 191.1 wide equals 252 times s with 9.1pt content), so extents and insets
    // scale with the layout spacing scale like every other fixed design length (W6-a1).
    // Content measures with the ambient (scaled on WC pages) measurer; no raw threading
    // and no emission map apply: coordinates are absolute flow space from birth.
    internal static double EstimateInlineTextBoxHeight(DocxInlineTextBox textBox, double paragraphSpacingScale)
    {
        return (ReadEmuPoints(textBox.ExtentCyValue) ?? 0d) * paragraphSpacingScale;
    }

    internal static DocxInlineTextBoxLayout? CreateInlineTextBoxLayout(
        DocxInlineTextBox textBox,
        int? sourceBlockIndex,
        double x,
        double width,
        double boxTop,
        DocxTextAlignment alignment,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        double paragraphSpacingScale,
        int pageNumber,
        CancellationToken cancellationToken,
        int sourceParagraphIndex = 0)
    {
        double? extentWidth = ReadEmuPoints(textBox.ExtentCxValue);
        double? extentHeight = ReadEmuPoints(textBox.ExtentCyValue);
        if (textMeasurer is null ||
            textBox.BodyElements.Count == 0 ||
            extentWidth is not { } fileWidth ||
            fileWidth <= 0d ||
            extentHeight is not { } fileHeight ||
            fileHeight <= 0d)
        {
            return null;
        }

        ResolveInlineTextBoxContentInsets(textBox, out double insetLeft, out double insetTop, out double insetRight, out _);
        double boxWidth = Math.Min(width, fileWidth * paragraphSpacingScale);
        double contentWidth = Math.Max(1d, boxWidth - (insetLeft + insetRight) * paragraphSpacingScale);
        var story = new DocxRelatedStory(
            DocxRelatedStoryKind.TextBox,
            "inline-textbox",
            null,
            textBox.BodyElements,
            [],
            [], null);
        DocxRelatedStoryLayout storyLayout = CreateRelatedStoryLayout(
            story,
            storyIndex: -1,
            contentWidth,
            textMeasurer,
            defaultTabStopPoints,
            paragraphSpacingScale,
            pageNumber: pageNumber,
            pageCount: null,
            cancellationToken: cancellationToken);
        double boxX = alignment switch
        {
            DocxTextAlignment.Center => x + Math.Max(0, width - boxWidth) / 2d,
            DocxTextAlignment.Right => x + Math.Max(0, width - boxWidth),
            _ => x
        };
        double contentX = boxX + insetLeft * paragraphSpacingScale;
        double contentTop = boxTop - insetTop * paragraphSpacingScale;
        // Autofit-grow: the block grows when content exceeds the declared extent (Word
        // default); shrinking behavior stays queued under autofit W6-b.
        double boxHeight = Math.Max(fileHeight * paragraphSpacingScale, storyLayout.ContentHeight);
        return new DocxInlineTextBoxLayout(
            textBox,
            boxX,
            boxTop,
            boxWidth,
            boxHeight,
            ShiftTextLines(storyLayout.TextLines, contentTop, contentX),
            storyLayout.InlineImages
                .Select(image => image with
                {
                    X = contentX + image.X,
                    Y = contentTop + image.Y,
                    PageIndex = pageNumber
                })
                .ToArray(),
            ShiftTableRows(storyLayout.TableRows, contentTop, contentX),
            sourceBlockIndex,
            SourceParagraphIndex: sourceParagraphIndex,
            Story: null);
    }
    private static double ReadInsetEmuPoints(string? value, long defaultEmu)
    {
        return ReadEmuPoints(value) ?? OoxUnits.EmuToPoints(defaultEmu);
    }

    private static DocxFloatingDrawingLayout CreateFloatingDrawingLayout(
        DocxFloatingDrawing drawing,
        DocxLayoutPage? anchorPage,
        int? pageStartIndex,
        int? pageEndIndex,
        int? anchorPageIndex,
        int? anchorColumnIndex,
        DocxLayoutSourceBlockBounds? sourceBlock,
        DocxStoryId? story,
        int? pageCount,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        double paragraphSpacingScale,
        CancellationToken cancellationToken,
        IDocxTextMeasurer? unscaledTextMeasurer = null)
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
            story,
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

            // Textbox content lays out in design space on scaled pages (Office: the tbxrev
            // probe shows Word center-scales design content to 9.1pt beside scaled body, so
            // breaks must come from raw 12pt advances in the unscaled content box). The
            // renderer maps design coordinates uniformly to emission space
            // (FloatingTextBoxEmissionMap); unit-scale pages keep the single-measurer path.
            // Frames stay unscaled in every mode.
            IDocxTextMeasurer contentMeasurer = textMeasurer;
            double contentSpacingScale = paragraphSpacingScale;
            if (unscaledTextMeasurer is not null &&
                Math.Abs(paragraphSpacingScale - 1d) >= 0.000000001d)
            {
                contentMeasurer = unscaledTextMeasurer;
                contentSpacingScale = 1d;
            }

            ResolveTextBoxContentInsets(drawing, out double insetLeft, out _, out double insetRight, out _);
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
                Math.Max(1d, width - insetLeft - insetRight),
                contentMeasurer,
                defaultTabStopPoints,
                contentSpacingScale,
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

    // RV15: source-block bounds index built once per stable page list. The
    // single pass reproduces FindSourceBlockBounds exactly (first/last page,
    // min/max column, topmost top and bottommost bottom across all matching
    // items); lookups are O(1) instead of O(pages x items) per drawing.
    // The index is built from the pages handed to each placement call, so
    // repagination/displacement can never reuse page-sensitive geometry.
    // FindSourceBlockBounds stays as the independently pinned reference.
    private static IReadOnlyDictionary<int, DocxLayoutSourceBlockBounds> BuildSourceBlockIndex(
        IReadOnlyList<DocxLayoutPage> pages,
        CancellationToken cancellationToken)
    {
        var index = new Dictionary<int, DocxLayoutSourceBlockBounds>();
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxLayoutPage page = pages[pageIndex];
            foreach (DocxLayoutItem item in page.Items)
            {
                if (GetSourceBlockIndex(item) is not { } sourceBlockIndex)
                {
                    continue;
                }

                if (!index.TryGetValue(sourceBlockIndex, out DocxLayoutSourceBlockBounds? bounds))
                {
                    bounds = new DocxLayoutSourceBlockBounds(pageIndex, pageIndex, null, null, double.NegativeInfinity, double.PositiveInfinity);
                }
                else
                {
                    bounds = bounds with { LastPageIndex = pageIndex };
                }

                if (ResolveItemColumnIndex(page, item) is { } columnIndex)
                {
                    bounds = bounds with
                    {
                        FirstColumnIndex = bounds.FirstColumnIndex is null ? columnIndex : Math.Min(bounds.FirstColumnIndex.Value, columnIndex),
                        LastColumnIndex = bounds.LastColumnIndex is null ? columnIndex : Math.Max(bounds.LastColumnIndex.Value, columnIndex)
                    };
                }

                (double y, double height) = GetVerticalBounds(item);
                bounds = bounds with
                {
                    VerticalTop = Math.Max(bounds.VerticalTop, y + height),
                    VerticalBottom = Math.Min(bounds.VerticalBottom, y)
                };
                index[sourceBlockIndex] = bounds;
            }
        }

        return index;
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
            // RV05 floatbox probe: paragraph tops hang the first baseline plus the
            // first-line inset. Lines without a stored inset keep the legacy font-size
            // bound; the shared bound keeps both index builders in agreement.
            DocxTextLineLayout text => (text.BaselineY, text.IsFirstParagraphLine == true && text.FirstLineInsetPoints is double inset ? inset : text.FontSize),
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
        IReadOnlyList<DocxTableRowLayout> TableRows,
        IReadOnlyList<DocxInlineTextBoxLayout> InlineTextBoxes,
        double EndCursorY,
        double EndPendingAfterSpacing);
}
