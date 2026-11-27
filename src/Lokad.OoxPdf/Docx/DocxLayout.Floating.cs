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
        int? pageCount,
        IDocxTextMeasurer? unscaledTextMeasurer,
        IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph>? mathEquations)
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
                DocxLayoutSourceBlockBounds? sourceBlock = (drawing.LayoutBlockIndex ?? drawing.SourceBlockIndex) is { } storyBlockIndex
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
                    unscaledTextMeasurer: unscaledTextMeasurer,
                    mathEquations: mathEquations);
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

    private static double? ResolveRelatedStoryFirstInsetPoints(DocxParagraph paragraph, IDocxTextMeasurer textMeasurer, IDocxHheaLineGapProvider gapProvider, bool selectMaxHheaRun = false)
    {
        if (textMeasurer is not IDocxLineMetricsProvider lineMetrics || paragraph.Runs.Count == 0)
        {
            return null;
        }

        // RV06 take battery (Word 16.0, a12 21/21/18): footnote slash endnote mixed notes
        // keep the legacy floor so take capacities hold, while uniform notes resolve their
        // own summed inset; a single distinct family-plus-size keeps the new path order-blind.
        // Textbox mixed notes resolve the max-hhea run instead (take-neutral: textbox takes
        // never drain pages).
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

        DocxTextRun? insetRun = null;
        double insetRunSize = 0d;
        if (uniformRuns && insetFamily is not null)
        {
            foreach (DocxTextRun run in paragraph.Runs)
            {
                if (string.IsNullOrWhiteSpace(run.Text) || run.EffectiveProperties.Hidden)
                {
                    continue;
                }

                insetRun = run;
                insetRunSize = insetSize;
                break;
            }
        }
        else if (selectMaxHheaRun)
        {
            // RV06 floatmixsz probes (Word 16.0, Tahoma 10/14pt mixed with Calibri 12pt, both orders): Office
            // mixed-textbox insets take the max tier across runs (Tah10 reads Cal-tier, Tah14 and Tah12
            // read Tah-tier), unlike footnote mixed notes which keep the legacy floor so take capacities hold.
            double maxTierPoints = double.NegativeInfinity;
            foreach (DocxTextRun run in paragraph.Runs)
            {
                if (string.IsNullOrWhiteSpace(run.Text) || run.EffectiveProperties.Hidden)
                {
                    continue;
                }
                double size = run.EffectiveProperties.FontSize;
                double tier;
                if (textMeasurer is IDocxStaticTextMetricsProvider staticSel && textMeasurer is IDocxHheaDescenderProvider)
                {
                    tier = lineMetrics.MeasureHheaLineHeight(run, size) - staticSel.MeasureWindowsDescender(run, size);
                }
                else
                {
                    tier = lineMetrics.MeasureHheaAscender(run, size) + gapProvider.MeasureHheaLineGap(run, size);
                }
                if (tier > maxTierPoints)
                {
                    maxTierPoints = tier;
                    insetRun = run;
                    insetRunSize = size;
                }
            }
        }

        if (insetRun is null)
        {
            return null;
        }

        // RV06 gap slice (Palatino/Algerian first baselines) plus Calibri-class tie
        // (content single-em 1.220703125 against reference 1.2207): gap-overflowing notes
        // keep the legacy floor, so the new path applies only while content single-em
        // stays at or below the mark reference within float representation.
        if (insetRunSize > 0d && lineMetrics.MeasureSingleLineHeight(insetRun, insetRunSize) / insetRunSize > EndnoteSeparatorReferenceSingleLineEm + 1e-4)
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
            return lineMetrics.MeasureHheaLineHeight(insetRun, insetRunSize) - staticMetrics.MeasureWindowsDescender(insetRun, insetRunSize);
        }

        return lineMetrics.MeasureHheaAscender(insetRun, insetRunSize) + gapProvider.MeasureHheaLineGap(insetRun, insetRunSize);
    }

    private static DocxParagraph? ResolveUnmarkedNoteMetricParagraph(DocxParagraph paragraph)
    {
        double contentSize = paragraph.Runs.Where(r => !IsAutomaticSuperscriptNoteMarker(r) &&
            !string.IsNullOrWhiteSpace(r.Text) && !r.EffectiveProperties.Hidden)
            .Select(r => r.EffectiveProperties.FontSize).DefaultIfEmpty(0d).Max();
        // Retain the independently qualified mixed inset when the source marker
        // does not dominate the content's nominal metric box.
        if (contentSize <= 0d || paragraph.Runs[0].EffectiveProperties.FontSize <= contentSize) return null;
        return paragraph with { Runs = paragraph.Runs.Where(r => !IsAutomaticSuperscriptNoteMarker(r)).ToArray() };
    }

    private static bool IsAutomaticSuperscriptNoteMarker(DocxTextRun run) =>
        run.IsAutomaticNoteStoryMarker && run.VerticalAlignment == DocxRunVerticalAlignment.Superscript;

    private static double ResolveRelatedStoryBaselineOffset(DocxParagraph paragraph, double fontSize,
        double lineHeight, IDocxTextMeasurer textMeasurer, DocxStoryKind? kind)
    {
        double? hheaAscender = kind is DocxStoryKind.Footnote or DocxStoryKind.Endnote
            ? DocxLineMetrics.ResolveHheaAscenderPoints(paragraph, textMeasurer as IDocxLineMetricsProvider, selectMaxHhea: false) : null;
        if (kind is DocxStoryKind.Footnote or DocxStoryKind.Endnote or DocxStoryKind.TextBox &&
            textMeasurer is IDocxHheaLineGapProvider gapProvider &&
            ResolveRelatedStoryFirstInsetPoints(paragraph, textMeasurer, gapProvider, kind is DocxStoryKind.TextBox) is double inset &&
            double.IsFinite(inset) && inset >= 0d) return inset;
        // Global font descenders may exceed the note metric frame. Such a hint
        // cannot own a physical baseline; retain the existing measured fallback
        // instead of making ordered inline content fail its extent invariants.
        return DocxLineMetrics.ResolveBodyBaselineOffset(fontSize, lineHeight,
            IsExactLineSpacing(paragraph.EffectiveProperties), hheaAscender);
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
        bool continuesShading,
        double beforeSpacing,
        IDocxTextMeasurer textMeasurer,
        double defaultTabStopPoints,
        DocxPageNumber? pageNumber,
        int? pageCount,
        double? firstStoryBaselineOffset = null,
        double transitionScale = 1d,
        double sourceStoryScale = 1d,
        DocxTextRun? customNoteMarker = null)
    {
        IReadOnlyList<DocxTextSpan> textSpans = CreateTextSpans(paragraph.Runs, pageNumber, pageCount);
        if (textSpans.Count == 0)
        {
            return (Array.Empty<DocxTextLineLayout>(), Array.Empty<DocxInlineImageLayout>(), 0d, 0d);
        }

        double fontSize = GetParagraphFontSize(paragraph);
        DocxLineHeightProfile lineHeightProfile = ResolveLineHeightProfile(paragraph, fontSize, textMeasurer, selectMaxAcrossRuns: story?.Kind is DocxStoryKind.Footnote or DocxStoryKind.Endnote);
        double baseLineHeight = lineHeightProfile.LineHeight;
        DocxEffectiveParagraphProperties effective = paragraph.EffectiveProperties;
        double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, fontSize, textMeasurer, fixedScale, defaultTabStopPoints);
        double continuationTextStartOffset = GetParagraphTextStartOffset(paragraph, fixedScale);
        double labelStartOffset = GetParagraphLabelStartOffset(paragraph, fixedScale);
        double paragraphX = textStartOffset;
        double paragraphWidth = Math.Max(1d, bodyWidth - textStartOffset - GetParagraphRightInset(paragraph, fixedScale));
        double continuationParagraphWidth = Math.Max(1d, bodyWidth - continuationTextStartOffset - GetParagraphRightInset(paragraph, fixedScale));
        DocxTextRun firstRun = paragraph.Runs[0];
        var storyTabStops = ScaleTabStopPositions(effective.TabStops, fixedScale);
        DocxWrappedTextLine[] lines = WrapTextLines(textSpans, paragraphWidth, continuationParagraphWidth, fontSize, textMeasurer, storyTabStops, defaultTabStopPoints * fixedScale, allowOverwideTokenBreaks: ShouldAllowCharacterLevelWordWrap(paragraph), dynamicFieldPageNumber: pageNumber, inlineImageWidths: ResolveInlineImageWrapWidths(paragraph, textSpans), sourceParagraph: paragraph).ToArray();
        // Auto insets take max(hheaAscender, 0.94em) for footnote/endnote stories
        // (Word 16.0 endnote grids); comment stories keep legacy (unprobed).
        // RV06 shape probes (Word 16.0, eighteen families) plus take battery (a12
        // 21/21/18) plus floatinset probes (Tahoma slash Calibri slash Informal textboxes):
        // uniform-note first insets follow hhea box minus Windows descender with
        // no 0.94em floor, while mixed notes keep the legacy floor so take capacities
        // hold; gap-aware measurers resolve uniform notes directly and every other path
        // stays byte-identical.
        double storyBaselineOffset = ResolveRelatedStoryBaselineOffset(paragraph, fontSize, baseLineHeight, textMeasurer, story?.Kind);
        DocxLineHeightProfile? unmarkedProfile = null;
        double unmarkedBaselineOffset = storyBaselineOffset;
        bool plainNoteParagraph = fixedScale == 1d && transitionScale == 1d &&
            story?.Kind is DocxStoryKind.Footnote or DocxStoryKind.Endnote &&
            DocxPlainBodyFrame.IsPlainParagraph(paragraph) &&
            paragraph.Runs.Count >= 2 && IsAutomaticSuperscriptNoteMarker(paragraph.Runs[0]);
        bool plainNoteMetrics = plainNoteParagraph && DocxPlainBodyFrame.HasContentDrivenLineFlow(paragraph);
        if (plainNoteParagraph && HasNoteParagraphFlowProperties(paragraph) && sourceStoryScale == 1d &&
            lineHeightProfile.Source == DocxLineHeightSource.ExactLineSpacing &&
            double.IsFinite(baseLineHeight) && baseLineHeight > 0d)
        {
            // Plain notes share the Office-qualified body/table fixed-line inset.
            // Their fixed boxes retain clipping; they do not use unmarked natural metrics.
            storyBaselineOffset = DocxLineMetrics.ResolveBodyBaselineOffset(fontSize, baseLineHeight, hasExplicitLineSpacing: true);
            unmarkedBaselineOffset = storyBaselineOffset;
        }
        if (lines.Length > 1 && plainNoteMetrics &&
            ResolveUnmarkedNoteMetricParagraph(paragraph) is DocxParagraph metricParagraph)
        {
            double metricSize = GetParagraphFontSize(metricParagraph);
            unmarkedProfile = ResolveNoteContentLineHeightProfile(metricParagraph, metricSize, textMeasurer);
            unmarkedBaselineOffset = ResolveRelatedStoryBaselineOffset(metricParagraph, metricSize,
                unmarkedProfile.LineHeight, textMeasurer, story?.Kind);
        }
        bool customNoteMetrics = customNoteMarker is not null && sourceStoryScale == 1d && fixedScale == 1d &&
            transitionScale == 1d && DocxPlainBodyFrame.IsPlainParagraph(paragraph) &&
            HasNoteParagraphFlowProperties(paragraph) && DocxPlainBodyFrame.HasContentDrivenLineFlow(paragraph);
        if (sourceStoryScale == 1d && (customNoteMetrics ||
            plainNoteMetrics && HasNoteParagraphFlowProperties(paragraph)))
        {
            DocxLineHeightProfile natural = ResolveNoteContentLineHeightProfile(paragraph, fontSize,
                textMeasurer, out double? naturalInset);
            if (lineHeightProfile.SingleLineHeight is { } sourceBox &&
                natural.SingleLineHeight is { } naturalBox && naturalBox > sourceBox && naturalInset is { } inset)
            {
                // A larger natural first-line box owns its measured baseline
                // bearing as well as its flow charge. Changing either alone
                // moves the rule and glyphs by different amounts.
                lineHeightProfile = natural;
                storyBaselineOffset = Math.Max(storyBaselineOffset, inset);
            }
        }
        // RV05: ordered inline atoms (related-story path). Affined images in
        // text-mixed paragraphs attach to wrapped lines at run position.
        DocxMidLinePlan? storyMidLinePlan = CreateMidLinePlan(paragraph, textSpans, lines, paragraphWidth, continuationParagraphWidth, storyBaselineOffset);
        var layouts = new List<DocxTextLineLayout>(lines.Length);
        var placedImages = new List<DocxInlineImageLayout>();
        double startCursorY = cursorY;
        bool firstLine = true;
        for (int lineIndex = 0; lineIndex < lines.Length; lineIndex++)
        {
            DocxWrappedTextLine line = lines[lineIndex];
            bool unmarked = unmarkedProfile is not null && !line.Spans.Any(s => IsAutomaticSuperscriptNoteMarker(s.StyleRun));
            DocxLineHeightProfile activeProfile = unmarked ? unmarkedProfile! : lineHeightProfile;
            double customBaselineOffset = storyBaselineOffset;
            if (customNoteMetrics && !line.Spans.Any(span => ReferenceEquals(span.StyleRun, customNoteMarker)))
            {
                // A wrapped line owns only the source runs actually present on it.
                // An absent custom prefix does not inflate its natural metric box.
                DocxParagraph presentMetricParagraph = paragraph with
                {
                    Runs = line.Spans.Select(span => span.StyleRun).DistinctBy(run => run, ReferenceEqualityComparer.Instance).ToArray(),
                };
                double metricSize = GetParagraphFontSize(presentMetricParagraph);
                DocxLineHeightProfile presentProfile = ResolveNoteContentLineHeightProfile(presentMetricParagraph, metricSize, textMeasurer);
                // Preserve the qualified mixed inset when the prefix did not
                // enlarge the natural box. Only an absent dominating run needs
                // a new metric view, including inside an authored minimum box.
                if (presentProfile.SingleLineHeight is { } presentBox &&
                    lineHeightProfile.SingleLineHeight is { } sourceBox && presentBox < sourceBox)
                {
                    activeProfile = presentProfile;
                    customBaselineOffset = ResolveRelatedStoryBaselineOffset(presentMetricParagraph, metricSize,
                        activeProfile.LineHeight, textMeasurer, story?.Kind);
                }
            }
            double lineHeight = activeProfile.LineHeight;
            double lineWidth = MeasureTextSpansForLayout(line.Spans, fontSize, textMeasurer, storyTabStops, defaultTabStopPoints * fixedScale, pageNumber) + (storyMidLinePlan?.LineImageWidths[lineIndex] ?? 0d);
            // RV05 calibration (Word 16.0): image top pins to the natural line top.
            double extraAbove = IsExactLineSpacing(paragraph.EffectiveProperties) ? 0d : (storyMidLinePlan?.ShiftAboveHeights[lineIndex] ?? 0d);
            cursorY -= extraAbove;
            // RV06 footnote-align probe: Office centers/rights drawable related-story
            // text too, letting trailing spaces overflow past the edge.
            double lineAlignWidth = effective.Alignment is DocxTextAlignment.Center or DocxTextAlignment.Right
                ? MeasureDrawableTextSpansForLayout(line.Spans, fontSize, textMeasurer, storyTabStops, defaultTabStopPoints * fixedScale, pageNumber) + (storyMidLinePlan?.LineImageWidths[lineIndex] ?? 0d)
                : lineWidth;
            double lineX = effective.Alignment switch
            {
                DocxTextAlignment.Center => paragraphX + Math.Max(0, paragraphWidth - lineAlignWidth) / 2d,
                DocxTextAlignment.Right => paragraphX + Math.Max(0, paragraphWidth - lineAlignWidth),
                _ => paragraphX
            };
            double baselineOffset = customNoteMetrics ? customBaselineOffset : unmarked ? unmarkedBaselineOffset : storyBaselineOffset;
            if ((unmarkedProfile is not null || customNoteMetrics || plainNoteMetrics && sourceStoryScale == 1d) &&
                activeProfile.Source == DocxLineHeightSource.AtLeastLineSpacing &&
                activeProfile.SingleLineHeight is { } naturalLineHeight)
            {
                // A minimum line box grows above the same source-owned bottom bearing.
                baselineOffset += Math.Max(0d, lineHeight - naturalLineHeight);
            }
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
                storyTabStops,
                defaultTabStopPoints * fixedScale,
                pageNumber, storyMidLinePlan?.LineImageWidths[lineIndex] ?? 0d);
            IReadOnlyList<DocxTextSegmentLayout> emissionSegments = lineShape.Segments;
            if (storyMidLinePlan is not null && storyMidLinePlan.ImagesByLine[lineIndex].Count != 0)
            {
                var imageShifts = new List<(double BoundaryX, double Shift)>();
                foreach (DocxMidLineImage placed in storyMidLinePlan.ImagesByLine[lineIndex])
                {
                    double beforeWidth = MeasureMidLineBeforeWidth(line.Spans, placed.LineCharOffset, paragraph, firstLine, lineIndex == lines.Length - 1, paragraphWidth, fontSize, textMeasurer, storyTabStops, defaultTabStopPoints * fixedScale, pageNumber, storyMidLinePlan.LineImageWidths[lineIndex]);
                    imageShifts.Add((beforeWidth, placed.Width));
                    placedImages.Add(new DocxInlineImageLayout(
                        placed.Image,
                        lineX + beforeWidth + placed.AdvanceBefore,
                        // RV05 calibration (Word 16.0 fnimg probe): Office footnote image bottoms sit
                        // at the text baseline (image cm bottom 85.7 against baseline 85.584), so
                        // related-story midline images anchor by the line baseline like cells.
                        cursorY - baselineOffset,
                        placed.Width,
                        placed.Height,
                        PageIndex: 0,
                        SourceBlockIndex: sourceBlockIndex,
                        SourceParagraphIndex: sourceParagraphIndex,
                        Story: null)
                    { SourceLineIndex = lineIndex,
                        LineFrame = DocxPictureLineFrame.FromLine(cursorY, lineHeight, extraAbove, lineIndex == lines.Length - 1 ? spacingProfile.ParagraphAfterSpacing : 0d) });
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
                SingleLineHeight: activeProfile.SingleLineHeight,
                ListLabelSingleLineHeight: activeProfile.ListLabelSingleLineHeight,
                BodyWindowsLineHeight: activeProfile.BodyWindowsLineHeight,
                ListLabelWindowsLineHeight: activeProfile.ListLabelWindowsLineHeight,
                EffectiveLineSpacingFactor: activeProfile.EffectiveLineSpacingFactor,
                LineSpacingFactorFloorApplied: activeProfile.LineSpacingFactorFloorApplied,
                LineHeightSource: activeProfile.Source,
                PendingAfterSpacing: firstLine ? spacingProfile.PendingAfterSpacing : null,
                ParagraphBeforeSpacing: firstLine ? spacingProfile.ParagraphBeforeSpacing : null,
                ParagraphAfterSpacing: firstLine ? spacingProfile.ParagraphAfterSpacing : null,
                ContextualSpacingSuppressed: firstLine ? spacingProfile.ContextualSpacingSuppressed : null,
                SourceParagraph: paragraph, EmitsTerminalParagraphMark: false)
            { ParagraphShadingFrame = CreateParagraphShadingFrame(paragraph, continuesShading, 0d, bodyWidth, lineShape.X,
                baselineOffset + extraAbove, lineHeight + extraAbove, fixedScale,
                lineIndex == lines.Length - 1 ? spacingProfile.ParagraphAfterSpacing : 0d,
                lineIndex == 0 ? beforeSpacing : 0d) });
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
            HasVisibleTextSpans(textSpans))
        {
            AppendRowEndSpace(layouts, paragraph, firstRun, fontSize, textMeasurer);
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
        IDocxTextMeasurer? unscaledTextMeasurer,
        IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph>? mathEquations)
    {
        var layouts = new DocxFloatingDrawingLayout[drawings.Count];
        IReadOnlyDictionary<int, DocxLayoutSourceBlockBounds>? sourceBlocks = null;
        Dictionary<DocxParagraphDrawingOwner, DocxLayoutSourceBlockBounds>? cellAnchors = null;
        for (int i = 0; i < drawings.Count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxFloatingDrawing drawing = drawings[i];
            DocxLayoutSourceBlockBounds? sourceBlock = null;
            if (drawing.ParagraphOwner is { } owner)
            {
                cellAnchors ??= IndexBodyCellDrawingAnchors(pages, paragraphSpacingScale < 1d, cancellationToken);
                cellAnchors.TryGetValue(owner, out sourceBlock);
            }
            if (sourceBlock is null && (drawing.LayoutBlockIndex ?? drawing.SourceBlockIndex) is { } blockIndex)
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
                unscaledTextMeasurer: unscaledTextMeasurer,
                mathEquations: mathEquations);
        }

        return layouts;
    }

    private static Dictionary<DocxParagraphDrawingOwner, DocxLayoutSourceBlockBounds> IndexBodyCellDrawingAnchors(
        IReadOnlyList<DocxLayoutPage> pages, bool usePrintedLines, CancellationToken cancellationToken)
    {
        var anchors = new Dictionary<DocxParagraphDrawingOwner, DocxLayoutSourceBlockBounds>();
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var page = pages[pageIndex];
            foreach (var item in page.Items)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (item is not DocxTableRowLayout row) continue;
                int? column = ResolveItemColumnIndex(page, row);
                bool admitted = false;
                foreach (var (anchor, _, _) in EnumerateCellDrawingAnchors(row, null, cancellationToken))
                    admitted |= anchors.TryAdd(anchor.Owner, new(pageIndex, pageIndex, column, column, anchor.Top, anchor.Bottom));
                if (!usePrintedLines || !admitted) continue;

                // The row's finalized lines include its Review projection and
                // any movement inherited from nested tables or page fragments.
                foreach (var (line, _) in DocxTableCellClipGeometry.EnumerateCellLines(row, null, cancellationToken))
                {
                    if (line is not { IsFirstParagraphLine: true, SourceParagraph: { } paragraph }) continue;
                    foreach (var owner in paragraph.DrawingOwners)
                    {
                        if (anchors.TryGetValue(owner, out var bounds) && bounds.FirstPageIndex == pageIndex && bounds.AnchorTextLine is null)
                            anchors[owner] = bounds with { AnchorTextLine = line };
                    }
                }
            }
        }
        return anchors;
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
        DocxStaticStoryLayoutResult? placedStory,
        IDocxTextMeasurer? unscaledTextMeasurer,
        IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph>? mathEquations)
    {
        Dictionary<DocxParagraphDrawingOwner, (DocxPlacedDrawingAnchor Anchor, DocxStaticAnchorY Frame, PdfRectangle Clip)>? cellAnchors = null;
        foreach (DocxFloatingDrawing drawing in selectedDrawings.Drawings)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxLayoutSourceBlockBounds? sourceBlock = null;
            DocxStaticAnchorY? anchorY = null;
            PdfRectangle? cellClip = null;
            if (drawing.ParagraphOwner is { } owner && placedStory is not null)
            {
                if (placedStory.DrawingAnchors is { } anchors && anchors.TryGetValue(owner, out var anchor))
                {
                    sourceBlock = new(pageIndex, pageIndex, null, null, anchor.Top, anchor.Bottom);
                    anchorY = placedStory.DrawingAnchorsUseDesignY ? DocxStaticAnchorY.Design : DocxStaticAnchorY.Text;
                }
                else if (placedStory.TableRows.Count > 0)
                {
                    if (cellAnchors is null)
                    {
                        cellAnchors = new();
                        foreach (var row in placedStory.TableRows)
                        foreach (var (cellAnchor, clip, physicalY) in EnumerateCellDrawingAnchors(row, null, cancellationToken))
                            cellAnchors.TryAdd(cellAnchor.Owner, (cellAnchor,
                                physicalY ? DocxStaticAnchorY.Physical : DocxStaticAnchorY.Text,
                                new(clip.X, clip.Y, clip.Width, clip.Height)));
                    }
                    if (cellAnchors.TryGetValue(owner, out var cell))
                    {
                        sourceBlock = new(pageIndex, pageIndex, null, null, cell.Anchor.Top, cell.Anchor.Bottom);
                        anchorY = cell.Frame;
                        cellClip = cell.Clip;
                    }
                }
            }
            yield return CreateFloatingDrawingLayout(
                drawing,
                page,
                pageStartIndex: pageIndex,
                pageEndIndex: pageIndex,
                anchorPageIndex: pageIndex,
                anchorColumnIndex: null,
                sourceBlock: sourceBlock,
                story: story,
                pageCount: pageCount,
                textMeasurer: textMeasurer,
                defaultTabStopPoints: defaultTabStopPoints,
                paragraphSpacingScale: paragraphSpacingScale,
                cancellationToken: cancellationToken,
                unscaledTextMeasurer: unscaledTextMeasurer,
                mathEquations: mathEquations) with { StaticAnchorY = anchorY, StaticCellClip = cellClip };
        }
    }

    private static IEnumerable<(DocxPlacedDrawingAnchor Anchor, DocxTableCellClipSnapshot Clip, bool PhysicalY)> EnumerateCellDrawingAnchors(
        DocxTableRowLayout row, DocxTableCellClipSnapshot? parent, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        foreach (var cell in row.Cells)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (cell.VisualOwnership != DocxTableCellVisualOwnership.OwnCell) continue;
            var clip = DocxTableCellClipGeometry.Resolve(row, cell, parent);
            if (!(clip.Width > 0d && clip.Height > 0d)) continue;
            foreach (var anchor in cell.DrawingAnchors)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (double.IsFinite(anchor.Top) && double.IsFinite(anchor.Bottom) && anchor.Top >= anchor.Bottom &&
                    anchor.Top >= clip.Y - .000001d && anchor.Bottom <= clip.Y + clip.Height + .000001d)
                    yield return (anchor, clip, row.StaticStoryUsesPhysicalY);
            }
            foreach (var nested in cell.NestedRows)
                foreach (var anchor in EnumerateCellDrawingAnchors(nested, clip, cancellationToken))
                    yield return anchor;
        }
    }

    internal static DocxFloatingDrawingLayout PlaceStaticAnchorInDesignFrame(
        DocxFloatingDrawingLayout drawing, double top, double bottom)
    {
        var placement = ResolveVerticalPlacement(new(top, bottom), drawing.ExtentHeightPoints,
            drawing.Drawing.VerticalAlignValue, drawing.VerticalOffsetPoints);
        return drawing with
        {
            AnchorBlockVerticalTop = top, AnchorBlockVerticalBottom = bottom,
            VerticalReferenceTop = top, VerticalReferenceBottom = bottom,
            PlacedTop = placement.Position,
            WrapExclusionTop = drawing.WrapExclusionTop + (placement.Position - drawing.PlacedTop),
            StaticAnchorY = DocxStaticAnchorY.Design
        };
    }

    // Office defaults when bodyPr carries no explicit insets: 0.1in sides, 0.05in
    // top and bottom (w6a1d-tbx reference measures exactly these).
    private const long TextBoxDefaultHorizontalInsetEmu = 91440L;
    private const long TextBoxDefaultVerticalInsetEmu = 45720L;

    internal static (double Left, double Top, double Right, double Bottom) ResolveTextBoxContentInsets(DocxFloatingDrawing drawing) =>
        ResolveTextBoxContentInsets(drawing.TextBoxInsetLeftValue, drawing.TextBoxInsetTopValue,
            drawing.TextBoxInsetRightValue, drawing.TextBoxInsetBottomValue, drawing.TextBoxStrokeWidthEmuValue);

    internal static (double Left, double Top, double Right, double Bottom) ResolveTextBoxContentInsets(DocxInlineTextBox box) =>
        ResolveTextBoxContentInsets(box.TextBoxInsetLeftValue, box.TextBoxInsetTopValue,
            box.TextBoxInsetRightValue, box.TextBoxInsetBottomValue, box.TextBoxStrokeWidthEmuValue);

    private static (double Left, double Top, double Right, double Bottom) ResolveTextBoxContentInsets(
        string? left, string? top, string? right, string? bottom, string? strokeWidth)
    {
        // The Word inline and floating probes put the content inside the outline.
        double halfStrokePoints = (ReadEmuPoints(strokeWidth) ?? 0d) / 2d;
        return (ReadInsetEmuPoints(left, TextBoxDefaultHorizontalInsetEmu) + halfStrokePoints,
            ReadInsetEmuPoints(top, TextBoxDefaultVerticalInsetEmu) + halfStrokePoints,
            ReadInsetEmuPoints(right, TextBoxDefaultHorizontalInsetEmu) + halfStrokePoints,
            ReadInsetEmuPoints(bottom, TextBoxDefaultVerticalInsetEmu) + halfStrokePoints);
    }

    // Inline boxes join the scaled body flow uniformly (Office: the w6-inline probe shows
    // the box 191.1 wide equals 252 times s with 9.1pt content), so extents and insets
    // scale with the layout spacing scale like every other fixed design length (W6-a1).
    // Content measures with the ambient (scaled on WC pages) measurer; no raw threading
    // and no emission map apply: coordinates are absolute flow space from birth.
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
        DocxPageNumber pageNumber,
        CancellationToken cancellationToken,
        int sourceParagraphIndex = 0,
        IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph>? mathEquations = null,
        int? pageCount = null)
    {
        if (!TryReadInlineBoxExtent(textBox, out double fileWidth, out double fileHeight))
        {
            return null;
        }

        var (insetLeft, insetTop, insetRight, insetBottom) = ResolveTextBoxContentInsets(textBox);
        double boxWidth = Math.Min(width, fileWidth * paragraphSpacingScale);
        double contentWidth = Math.Max(1d, boxWidth - (insetLeft + insetRight) * paragraphSpacingScale);
        var story = new DocxRelatedStory(
            DocxRelatedStoryKind.TextBox,
            "inline-textbox",
            null,
            textBox.BodyElements,
            null);
        DocxRelatedStoryLayout? storyLayout = CreateTextBoxStoryLayout(
            story,
            contentWidth,
            textMeasurer,
            defaultTabStopPoints,
            paragraphSpacingScale,
            pageNumber: pageNumber,
            pageCount: pageCount,
            cancellationToken: cancellationToken,
            mathEquations: mathEquations);
        if (storyLayout is null)
        {
            return null;
        }

        double boxX = alignment switch
        {
            DocxTextAlignment.Center => x + Math.Max(0, width - boxWidth) / 2d,
            DocxTextAlignment.Right => x + Math.Max(0, width - boxWidth),
            _ => x
        };
        double contentX = boxX + insetLeft * paragraphSpacingScale;
        double contentTop = boxTop - insetTop * paragraphSpacingScale;
        // The existing growth path keeps the declared extent as a floor.
        // Children start below the top inset; their outer frame also owns the
        // bottom inset, rather than clipping its final source rows.
        double boxHeight = Math.Max(fileHeight * paragraphSpacingScale,
            storyLayout.ContentHeight + (insetTop + insetBottom) * paragraphSpacingScale);
        return new DocxInlineTextBoxLayout(
            textBox,
            boxX,
            boxTop,
            boxWidth,
            boxHeight,
            ShiftTextLines(storyLayout.TextLines, contentTop, contentX),
            storyLayout.InlineImages
                .Select(image => image.Relocate(contentX + image.X, contentTop + image.Y, pageNumber.Physical))
                .ToArray(),
            ShiftTableRows(storyLayout.TableRows, contentTop, contentX),
            sourceBlockIndex,
            SourceParagraphIndex: sourceParagraphIndex,
            Story: null) { PrintScale = paragraphSpacingScale };
    }
    private static bool TryReadInlineBoxExtent(DocxInlineTextBox box, out double width, out double height)
    {
        width = ReadEmuPoints(box.ExtentCxValue) ?? 0d;
        height = ReadEmuPoints(box.ExtentCyValue) ?? 0d;
        return width > 0d && height > 0d && box.BodyElements.Count != 0;
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
        IDocxTextMeasurer? unscaledTextMeasurer = null,
        IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph>? mathEquations = null)
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
            anchorPageIndex is { } index ? anchorPage?.GetPageNumber(index + 1) ?? new DocxPageNumber(index + 1) : null);
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
            textBoxLayout) { AnchorTextLine = sourceBlock?.AnchorTextLine };

        DocxRelatedStoryLayout? CreateFloatingTextBoxLayout(DocxPageNumber? pageNumber)
        {
            if (drawing.TextBoxBodyElements.Count == 0 ||
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
            IDocxTextMeasurer? contentMeasurer = textMeasurer;
            double contentSpacingScale = textMeasurer is null ? 1d : paragraphSpacingScale;
            if (unscaledTextMeasurer is not null &&
                Math.Abs(paragraphSpacingScale - 1d) >= 0.000000001d)
            {
                contentMeasurer = unscaledTextMeasurer;
                contentSpacingScale = 1d;
            }

            var (insetLeft, _, insetRight, _) = ResolveTextBoxContentInsets(drawing);
            var story = new DocxRelatedStory(
                DocxRelatedStoryKind.TextBox,
                "floating-drawing",
                drawing.ImageRelationshipId,
                drawing.TextBoxBodyElements,
                null);
            return CreateTextBoxStoryLayout(
                story,
                Math.Max(1d, width - insetLeft - insetRight),
                contentMeasurer,
                defaultTabStopPoints,
                contentSpacingScale,
                pageNumber: pageNumber,
                pageCount: pageCount,
                cancellationToken: cancellationToken,
                mathEquations: mathEquations);
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

    // Index source-block geometry once per stable page list: first/last page,
    // min/max column and full vertical extent. Rebuild after placement changes
    // so pagination/displacement cannot reuse page-sensitive geometry.
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
                    VerticalBottom = Math.Min(bounds.VerticalBottom, y),
                    AnchorTextLine = y + height > bounds.VerticalTop
                        ? item is DocxTextLineLayout { IsFirstParagraphLine: true } firstLine ? firstLine : null
                        : bounds.AnchorTextLine
                };
                index[sourceBlockIndex] = bounds;
            }
        }

        return index;
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
        double VerticalBottom)
    {
        public DocxTextLineLayout? AnchorTextLine { get; init; }
    }

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
        double EndPendingAfterSpacing)
    {
        public double? BodyFooterContentTop { get; init; }
        public double? PhysicalFooterFrameTop { get; init; }
        public double? PhysicalHeaderFrameBottom { get; init; }
        public IReadOnlyDictionary<DocxParagraphDrawingOwner, DocxPlacedDrawingAnchor>? DrawingAnchors { get; init; }
        public bool DrawingAnchorsUseDesignY { get; init; }
    }

}
