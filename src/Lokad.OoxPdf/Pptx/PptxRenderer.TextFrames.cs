using System.Globalization;
using System.Text;
using System.Xml.Linq;

using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static async ValueTask<PptxTextFrameLayout> BuildTextFrameLayoutAsync(PptxTextFlowFrame flowFrame, PptxDocument document, TextAdvanceEstimator advanceEstimator, bool allowWrapping)
    {
        PptxTextFrameLayout layout = (await BuildTextFrameLayoutAsync(flowFrame, document, advanceEstimator, allowWrapping, PptxTextColumnBreakMode.StrictFit, 0, 0).ConfigureAwait(false));
        var (mode, lineBalanceTarget, lineBalanceStartColumn) = ResolveOverflowColumnBalance(layout);
        return mode == PptxTextColumnBreakMode.StrictFit
            ? layout
            : await BuildTextFrameLayoutAsync(flowFrame, document, advanceEstimator, allowWrapping, mode, lineBalanceTarget, lineBalanceStartColumn).ConfigureAwait(false);
    }

    // RV03: word glue across run seams. A segment continues the current word
    // when the previous advance ends in a word char and the segment starts
    // with one. Letters, digits, apostrophes and no-break spaces glue; spaces,
    // tabs, punctuation and breaks start new words. Astral code points glue
    // conservatively so surrogate pairs are never split.
    private static bool WordContinuesOntoLine(int? previousCodePoint, string advanceText)
    {
        if (previousCodePoint is not { } previous || advanceText.Length == 0)
        {
            return false;
        }

        return IsWordGlueCodePoint(previous) && StartsWithWordGlue(advanceText);
    }

    private static bool IsWordGlueCodePoint(int codePoint)
    {
        if (codePoint == 0xA0 || codePoint == 0x27 || codePoint == 0x2019 || codePoint > 0xFFFF)
        {
            return true;
        }

        return Rune.IsLetterOrDigit(new Rune(codePoint));
    }

    private static bool StartsWithWordGlue(string advanceText)
    {
        foreach (Rune rune in advanceText.EnumerateRunes())
        {
            int value = rune.Value;
            return value == 0xA0 || value == 0x27 || value == 0x2019 || value > 0xFFFF || Rune.IsLetterOrDigit(rune);
        }
        return false;
    }
    // RV03: backward word scan over emitted line spans. Locates the trailing
    // word start (span index plus char offset within it) for pullback. Only
    // space-led splits proceed: the word must start after spaces, so the head
    // stays drawn with unchanged neighbors while the tail moves whole.
    private static bool TryFindTrailingWordStart(
        IReadOnlyList<PptxTextSpanLayout> spans,
        out int spanIndex,
        out int charOffset)
    {
        spanIndex = -1;
        charOffset = 0;
        bool inWord = false;
        for (int i = 0; i < spans.Count; i++)
        {
            string text = spans[i].Run.Text;
            int offset = 0;
            foreach (Rune rune in text.EnumerateRunes())
            {
                if (IsWordGlueCodePoint(rune.Value))
                {
                    if (!inWord)
                    {
                        spanIndex = i;
                        charOffset = offset;
                    }
                    inWord = true;
                }
                else
                {
                    inWord = false;
                }
                offset += rune.Utf16SequenceLength;
            }
        }
        if (!inWord || spanIndex < 0)
        {
            return false;
        }
        string head = spans[spanIndex].Run.Text[..charOffset];
        foreach (char c in head)
        {
            if (c != (char)32)
            {
                return false;
            }
        }
        return true;
    }
    // RV03: detaches the trailing-word spans from the line for re-emission
    // after the break. A space-only head stays drawn at the old line end
    // (mirroring leading-space emission); tails return without leading spaces
    // for the caller to re-add on the fresh line.
    private static List<PptxTextSpanLayout>? DetachTrailingWordSpans(
        TextLayoutLine line,
        int wordStartSpan,
        int wordLeadingSpaces,
        ref double cursorX,
        ref int? previousAdvanceCodePoint,
        TextAdvanceEstimator advanceEstimator)
    {
        if (wordStartSpan <= 0 || wordStartSpan >= line.Spans.Count)
        {
            return null;
        }
        var tails = line.RemoveTail(wordStartSpan);
        PptxTextSpanLayout first = tails[0];
        TextRun firstRun = first.Run;
        string tailText = firstRun.Text[wordLeadingSpaces..];
        if (wordLeadingSpaces > 0)
        {
            string headText = firstRun.Text[..wordLeadingSpaces];
            int? headPrevious = LastCodePoint(line.Spans[wordStartSpan - 1].Run.Text);
            double headIntrinsic = advanceEstimator.Measure(headText, firstRun.FontSize, firstRun.FontFamily, firstRun.Bold, firstRun.Italic, firstRun.CharacterSpacing, firstRun.KerningEnabled);
            double headBoundary = MeasureFlowSegmentBoundaryAdjustment(advanceEstimator, headText, headPrevious, firstRun.FontSize, new TextAdvanceOptions(firstRun.FontFamily, firstRun.Bold, firstRun.Italic, firstRun.CharacterSpacing, firstRun.KerningEnabled));
            double headWidth = Math.Max(0d, headIntrinsic + headBoundary);
            TextRun headRun = firstRun with { Text = headText, X = line.EndX, Width = headWidth };
            line.Add(first.SourceRun, headRun, line.EndX + headWidth, BuildTextAtoms(headRun, advanceEstimator, PptxTextAtomKind.Space), BuildGlyphSpan(headRun, advanceEstimator, 0d));
            cursorX = line.EndX;
            previousAdvanceCodePoint = LastCodePoint(headText);
        }
        tails[0] = first with { Run = firstRun with { Text = tailText } };
        return tails;
    }

    private static async ValueTask<PptxTextFrameLayout> BuildTextFrameLayoutAsync(PptxTextFlowFrame flowFrame, PptxDocument document, TextAdvanceEstimator advanceEstimator, bool allowWrapping, PptxTextColumnBreakMode columnBreakMode, int lineBalanceTarget, int lineBalanceStartColumn)
    {
        PptxTextFrameModel frame = flowFrame.Model;
        allowWrapping &= TextBodyAllowsWrapping(frame.BodyProperties);
        double cursorLineTop = flowFrame.Box.CursorTop;
        int columnIndex = 0;
        int linesInCurrentColumn = 0;
        double totalColumnSpacing = frame.ColumnSpacing * (frame.ColumnCount - 1);
        double columnWidth = frame.ColumnCount <= 1
            ? frame.TextWidth
            : Math.Max(1d, (frame.TextWidth - totalColumnSpacing) / frame.ColumnCount);
        double columnWrapWidth = frame.ColumnCount <= 1
            ? frame.TextWrapWidth
            : Math.Max(1d, (frame.TextWrapWidth - totalColumnSpacing) / frame.ColumnCount);
        double columnStartX = frame.TextX;
        bool clipsColumnsIndividually = !frame.TableRowIndex.HasValue && (frame.TextClipX != 0d ||
            frame.TextClipWidth < document.SlideWidthPoints - PptxTextMetricRules.CoordinateTolerance);
        bool strictClip = ClipsTextVerticalOverflow(frame.BodyProperties.VerticalOverflow);
        bool cullOutOfFrameLines = frame.BodyProperties.VerticalOverflow == PptxTextVerticalOverflow.Clip;
        int autoNumberValue = 1;
        var autoNumberSequences = new Dictionary<int, (string Type, int Start, int Next)>();
        bool hasPlacedParagraph = false;
        var paragraphLayouts = new List<PptxTextParagraphLayout>();

        foreach (PptxTextFlowParagraph flowParagraph in flowFrame.Paragraphs)
        {
            PptxTextParagraphModel paragraph = flowParagraph.Model;
            var lineLayouts = new List<PptxTextLineLayout>();
            ResolvedParagraphTextStyle paragraphStyle = flowParagraph.Style;
            bool suppressNoWrapDistribution = UsesNoWrapAsciiDistribution(frame, paragraph);
            bool useNumberSequence = !frame.TableRowIndex.HasValue && UsesOfficeAutoNumberSequence(paragraph.Bullet, paragraph.Level);
            bool emptyNumber = useNumberSequence && !paragraph.Runs.Any(run => run.Kind == PptxTextRunKind.Break || run.Text.Length > 0);
            if (emptyNumber)
            {
                ReadBulletText(paragraph.Bullet, paragraph.Level, ref autoNumberValue, autoNumberSequences, useNumberSequence, advanceSequence: false);
            }
            if (!paragraph.HasVisibleContent && paragraph.Bullet.Kind != PptxParagraphBulletKind.AutoNumber)
            {
                autoNumberSequences.Clear();
            }
            if (!paragraph.HasVisibleContent || emptyNumber)
            {
                if (paragraph.HasLayoutContent)
                {
                    double emptyFontSize = paragraph.EndParagraphStyle.FontSize;
                    cursorLineTop -= (hasPlacedParagraph ? paragraph.EmptySpacingBefore : 0d) + ReadParagraphAdvance(paragraphStyle.LineSpacing, emptyFontSize) + paragraph.EmptySpacingAfter;
                    hasPlacedParagraph = true;
                }

                paragraphLayouts.Add(new PptxTextParagraphLayout(paragraph, lineLayouts));
                continue;
            }

            if (paragraph.Bullet.Kind != PptxParagraphBulletKind.AutoNumber)
            {
                autoNumberValue = 1;
                autoNumberSequences.Clear();
            }

            string? bulletText = ReadBulletText(paragraph.Bullet, paragraph.Level, ref autoNumberValue, autoNumberSequences, useNumberSequence, advanceSequence: true);
            bool bulletPending = bulletText is not null;
            double effectiveTextWidth = columnWrapWidth;
            double bulletInset = PptxTextMetricRules.ClampNonNegative(paragraphStyle.Indent.MarginLeft + paragraphStyle.Indent.Hanging);
            double paragraphInset = bulletText is null ? bulletInset : PptxTextMetricRules.ClampNonNegative(paragraphStyle.Indent.MarginLeft);
            if (hasPlacedParagraph)
            {
                cursorLineTop -= paragraphStyle.SpacingBefore;
                double nextLineAdvance = ReadLineAdvance(paragraphStyle.LineSpacing, paragraphStyle.FontSize);
                MoveColumn(nextLineAdvance, linePlaced: false);
            }
            double bulletX = columnStartX + bulletInset;
            double paragraphTextX = columnStartX + paragraphInset;
            (double columnClipX, double columnClipWidth) = ReadColumnClip();

            bool afterManualLineBreak = false;
            bool afterLeadingManualLineBreak = false;
            bool shapeAutoFit = HasShapeAutoFit(frame.BodyProperties);
            bool useExplicitMultipleBaselineOffset = ShouldUseExplicitMultipleBaselineOffset(frame, paragraphStyle.LineSpacing);
            double cursorY = cursorLineTop - ReadFirstLineBaselineOffset(paragraph, paragraphStyle.LineSpacing, advanceEstimator, frame.UseOfficeBaselineFloor, useExplicitMultipleBaselineOffset);
            double cursorX = paragraphTextX;
            double maxFontSize = 0d;
            var line = new TextLayoutLine(paragraphTextX);
            int? previousAdvanceCodePoint = null;
            double pendingVisibleLeadingAdjustment = 0d;
            int trailingStartSpan = 0;
            bool trailingHasNoBreakSpace = false;
            int remainingDrawableSegments = CountDrawableTextSegments(flowParagraph);
            bool useNumberingWordWrap = false;
            foreach (PptxTextFlowRun flowRun in flowParagraph.Runs)
            {
                PptxTextRunModel modelRun = flowRun.Source;
                if (modelRun.Kind == PptxTextRunKind.Break)
                {
                    CommitManualBreak(flowRun.Style.FontSize);
                    continue;
                }

                ResolvedRunTextStyle runStyle = flowRun.Style;
                ResumeAfterManualBreak(runStyle);

                if (bulletPending && bulletText is not null)
                {
                    BulletStyle bulletStyle = ReadBulletStyle(paragraph.Bullet, runStyle.FontSize, runStyle.Color, runStyle.Typeface);
                    if (!useNumberSequence)
                    {
                        maxFontSize = Math.Max(maxFontSize, bulletStyle.FontSize);
                    }
                    double bulletWidth = PptxTextMetricRules.MinimumWidth(effectiveTextWidth - (bulletX - columnStartX));
                    bool useNumberingLabelLayout = useNumberSequence &&
                        paragraphStyle.Indent.Hanging <= PptxTextMetricRules.CoordinateTolerance &&
                        paragraphStyle.Alignment == TextAlignment.Left &&
                        paragraphStyle.TabStops.Count == 0 &&
                        frame.Orientation == PptxTextOrientation.Horizontal &&
                        Math.Abs(frame.TextRotationDegrees) <= PptxTextMetricRules.CoordinateTolerance &&
                        frame.ColumnCount == 1 && !frame.TableRowIndex.HasValue &&
                        frame.BodyProperties.AutofitMode != PptxTextAutofitMode.Normal;
                    if (useNumberingLabelLayout)
                        await advanceEstimator.PrepareAsync(runStyle.Typeface, runStyle.Bold, runStyle.Italic, bulletText).ConfigureAwait(false);
                    double numberingEndX = useNumberingLabelLayout
                        ? bulletX + advanceEstimator.Measure(bulletText, bulletStyle.FontSize, runStyle.Typeface, runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing, false) + runStyle.CharacterSpacing
                        : 0d;
                    useNumberingLabelLayout &= numberingEndX <= columnStartX + effectiveTextWidth;
                    if (useNumberingLabelLayout &&
                        (paragraph.Bullet.SizeKind != PptxParagraphBulletSizeKind.Text || paragraph.Bullet.FontTypeface is not null))
                    {
                        // Keep narrow override frames on their existing fallback
                        // when the first body fragment cannot fit after the label.
                        PptxTextFlowSegment firstBodyFragment = flowRun.Segments.FirstOrDefault(segment =>
                            segment.Draw && !string.IsNullOrWhiteSpace(segment.AdvanceText));
                        double fragmentWidth = advanceEstimator.Measure(firstBodyFragment.AdvanceText ?? string.Empty,
                            runStyle.FontSize * firstBodyFragment.FontScale, runStyle.Typeface,
                            runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing, runStyle.KerningEnabled);
                        useNumberingLabelLayout &= numberingEndX + fragmentWidth <= columnStartX + effectiveTextWidth;
                    }
                    if (useNumberingLabelLayout)
                    {
                        // Office automatic numbering follows the body face even
                        // when a bullet font is authored. The label keeps its size.
                        bulletStyle = bulletStyle with { Typeface = runStyle.Typeface };
                    }
                    useNumberingWordWrap = useNumberingLabelLayout && allowWrapping && HasNoAutoFit(frame.BodyProperties) &&
                        paragraph.Runs.All(run => run.Kind == PptxTextRunKind.Text &&
                            run.Text.All(character => character is >= ' ' and <= '~'));
                    if (useNumberingWordWrap)
                    {
                        // Strict word fitting must not introduce an empty body
                        // line when the first word already exceeds the frame.
                        PptxTextFlowSegment firstBodyFragment = flowRun.Segments.FirstOrDefault(segment =>
                            segment.Draw && !string.IsNullOrWhiteSpace(segment.AdvanceText));
                        double fragmentWidth = advanceEstimator.Measure(firstBodyFragment.AdvanceText ?? string.Empty,
                            runStyle.FontSize * firstBodyFragment.FontScale, runStyle.Typeface,
                            runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing, runStyle.KerningEnabled);
                        useNumberingWordWrap &= Math.Max(cursorX, numberingEndX) + fragmentWidth <= columnStartX + effectiveTextWidth;
                    }
                    // The authored bullet face can be overridden by numbering
                    // placement. Acquire only after that existing choice is final.
                    await advanceEstimator.PrepareAsync(bulletStyle.Typeface, runStyle.Bold, runStyle.Italic, bulletText).ConfigureAwait(false);
                    double bulletEndX = useNumberingLabelLayout ? numberingEndX : bulletX + advanceEstimator.Measure(bulletText, bulletStyle.FontSize, bulletStyle.Typeface, runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing, true);
                    TextRun bulletRun = new(bulletText, bulletX, cursorY, bulletWidth, frame.TextHeight, columnClipX, frame.TextClipY, columnClipWidth, frame.TextClipHeight, bulletStyle.FontSize, runStyle.CharacterSpacing, 0d, bulletStyle.Color, bulletStyle.Alpha, null, runStyle.Bold, runStyle.Italic, runStyle.Underline, runStyle.Strike, useNumberingLabelLayout ? false : runStyle.KerningEnabled, paragraphStyle.Alignment, bulletStyle.Typeface, frame.TextRotationDegrees, frame.RotationCenterX, frame.RotationCenterY, FlipHorizontal: false, FlipVertical: false, PreventCoalesce: false, Outline: null, StrictClip: strictClip);
                    line.Add(modelRun, bulletRun, bulletEndX, BuildTextAtoms(bulletRun, advanceEstimator, PptxTextAtomKind.Word), BuildGlyphSpan(bulletRun, advanceEstimator, 0d));
                    if (useNumberingLabelLayout)
                    {
                        // Office keeps the continuation indent, but the first body
                        // fragment cannot start inside a wide numbering label.
                        cursorX = Math.Max(cursorX, bulletEndX);
                    }
                    bulletPending = false;
                    EndWrapTail();
                }

                foreach (PptxTextFlowSegment flowSegment in flowRun.Segments)
                {
                    if (flowSegment.Kind == PptxTextFlowSegmentKind.Break)
                    {
                        CommitManualBreak(runStyle.FontSize);
                        continue;
                    }

                    ResumeAfterManualBreak(runStyle);

                    if (flowSegment.Kind == PptxTextFlowSegmentKind.Tab)
                    {
                        double tabSpaceWidth = advanceEstimator.Measure(" ", runStyle.FontSize, runStyle.Typeface, runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing, runStyle.KerningEnabled);
                        TextRun tabRun = new(" ", cursorX, cursorY, PptxTextMetricRules.MinimumWidth(tabSpaceWidth), frame.TextHeight, columnClipX, frame.TextClipY, columnClipWidth, frame.TextClipHeight, runStyle.FontSize, runStyle.CharacterSpacing, runStyle.BaselineOffset, runStyle.Color, runStyle.Alpha, runStyle.Highlight, runStyle.Bold, runStyle.Italic, runStyle.Underline, runStyle.Strike, runStyle.KerningEnabled, paragraphStyle.Alignment, runStyle.Typeface, frame.TextRotationDegrees, frame.RotationCenterX, frame.RotationCenterY, FlipHorizontal: false, FlipVertical: false, PreventCoalesce: true, Outline: runStyle.Outline, StrictClip: strictClip);
                        line.Add(modelRun, tabRun, cursorX + tabSpaceWidth, BuildTextAtoms(tabRun, advanceEstimator, PptxTextAtomKind.Tab), BuildGlyphSpan(tabRun, advanceEstimator, 0d));
                        cursorX = ResolveNextTabX(cursorX, paragraphTextX, paragraphStyle.TabStops);
                        line.AdvanceTo(cursorX);
                        previousAdvanceCodePoint = null;
                        pendingVisibleLeadingAdjustment = 0d;
                        EndWrapTail();
                        continue;
                    }

                    double fragmentFontSize = runStyle.FontSize * flowSegment.FontScale;
                    string currentSegment = flowSegment.Text;
                    string currentAdvanceText = flowSegment.AdvanceText;
                    bool isNoBreakHiddenAdvance = flowSegment.Kind == PptxTextFlowSegmentKind.NoBreakHiddenAdvance;
                    if (currentAdvanceText.StartsWith(' ') || (previousAdvanceCodePoint == '-' && !isNoBreakHiddenAdvance))
                    {
                        EndWrapTail();
                    }
                    bool isDrawableTextSegment = flowSegment.Draw && currentAdvanceText.TrimStart().Length > 0;
                    bool isFinalShortWordSegment = isDrawableTextSegment &&
                        remainingDrawableSegments == 1 &&
                        lineLayouts.Count == 0 &&
                        IsShortWordSegment(currentAdvanceText);
                    double segmentIntrinsicWidth = advanceEstimator.Measure(currentAdvanceText, fragmentFontSize, runStyle.Typeface, runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing, runStyle.KerningEnabled);
                    double segmentBoundaryAdjustment = MeasureFlowSegmentBoundaryAdjustment(advanceEstimator, currentAdvanceText, previousAdvanceCodePoint, fragmentFontSize, new TextAdvanceOptions(runStyle.Typeface, runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing, runStyle.KerningEnabled));
                    double segmentWidth = Math.Max(0d, segmentIntrinsicWidth + segmentBoundaryAdjustment);
                    bool splitOverwideFirstSegment = allowWrapping &&
                        flowSegment.Kind == PptxTextFlowSegmentKind.Text &&
                        flowSegment.Draw &&
                        currentSegment == currentAdvanceText &&
                        currentSegment.Length > 1 &&
                        line.Spans.Count == 0 &&
                        segmentWidth > effectiveTextWidth + PptxTextMetricRules.CoordinateTolerance;
                    if ((frame.Orientation != PptxTextOrientation.Horizontal &&
                            flowSegment.Kind == PptxTextFlowSegmentKind.Text &&
                            flowSegment.Draw &&
                            currentSegment == currentAdvanceText &&
                            currentSegment.Length > 1 &&
                            segmentWidth > frame.TextWidth) ||
                        splitOverwideFirstSegment)
                    {
                        // Office starts every vertical chunk run on a fresh pitch row instead
                        // of gluing it after the previous segment on one line. This mirrors the
                        // manual-break path (no leading-space shuffling: observed vertical chunk
                        // runs carry trailing spaces, never leading ones).
                        if (frame.Orientation == PptxTextOrientation.Vertical && line.Spans.Count > 0)
                        {
                            double freshLineFontSize = ResolveLineFontSize(maxFontSize, paragraphStyle.FontSize);
                            AddClippedParagraphLine(lineLayouts, line, CreateLineBox(cursorLineTop, cursorY, paragraphStyle.LineSpacing, freshLineFontSize, line, advanceEstimator, frame.UseOfficeBaselineFloor), paragraphStyle.Alignment, columnStartX, effectiveTextWidth, justify: false, distribute: paragraphStyle.Alignment == TextAlignment.Distributed, advanceEstimator, cullOutOfFrameLines, cursorY, frame.TextClipY, frame.TextClipHeight, suppressNoWrapDistribution);
                            double freshLineAdvance = ReadLineAdvance(paragraphStyle.LineSpacing, freshLineFontSize);
                            AdvanceTextLine(freshLineAdvance);
                            cursorY = cursorLineTop - LineBaselineOffset(fragmentFontSize, paragraphStyle.LineSpacing, runStyle, advanceEstimator, frame.UseOfficeBaselineFloor, useExplicitMultipleBaselineOffset);
                            cursorX = paragraphTextX;
                            line.Reset(paragraphTextX);
                            maxFontSize = 0d;
                            previousAdvanceCodePoint = null;
                            pendingVisibleLeadingAdjustment = 0d;
                            EndWrapTail();
                        }

                        double chunkMaxWidth = frame.Orientation == PptxTextOrientation.Horizontal
                            ? effectiveTextWidth
                            : frame.TextWidth;
                        // Vertical and qualified two/three-column chunks split at the column edge.
                        // A font-sized allowance changes the character boundary of unbroken
                        // column text and displaces its following lines.
                        double chunkFitTolerance = frame.Orientation == PptxTextOrientation.Vertical ||
                            (frame.Orientation == PptxTextOrientation.Horizontal && frame.ColumnCount is 2 or 3)
                            ? PptxTextMetricRules.CoordinateTolerance
                            : PptxTextMetricRules.WrapFitTolerance(fragmentFontSize);
                        // Office glues trailing spaces to the final vertical chunk ('al ')
                        // instead of stranding them on their own pitch step.
                        string chunkSourceText = currentSegment;
                        string chunkTrailingSpaces = string.Empty;
                        if (frame.Orientation == PptxTextOrientation.Vertical)
                        {
                            int chunkContentEnd = chunkSourceText.Length;
                            while (chunkContentEnd > 0 && chunkSourceText[chunkContentEnd - 1] == ' ')
                            {
                                chunkContentEnd--;
                            }

                            if (chunkContentEnd > 0 && chunkContentEnd < chunkSourceText.Length)
                            {
                                chunkTrailingSpaces = chunkSourceText[chunkContentEnd..];
                                chunkSourceText = chunkSourceText[..chunkContentEnd];
                            }
                        }

                        string[] chunks = SplitTextIntoFittingChunks(chunkSourceText, chunkMaxWidth + chunkFitTolerance, fragmentFontSize, runStyle, advanceEstimator);
                        if (chunkTrailingSpaces.Length != 0 && chunks.Length != 0)
                        {
                            chunks[^1] += chunkTrailingSpaces;
                        }
                        for (int chunkIndex = 0; chunkIndex < chunks.Length; chunkIndex++)
                        {
                            string chunk = chunks[chunkIndex];
                            double chunkWidth = advanceEstimator.Measure(chunk, fragmentFontSize, runStyle.Typeface, runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing, runStyle.KerningEnabled);
                            double chunkBoundaryAdjustment = chunkIndex == 0
                                ? MeasureFlowSegmentBoundaryAdjustment(advanceEstimator, chunk, previousAdvanceCodePoint, fragmentFontSize, new TextAdvanceOptions(runStyle.Typeface, runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing, runStyle.KerningEnabled))
                                : 0d;
                            double chunkTotalWidth = Math.Max(0d, chunkWidth + chunkBoundaryAdjustment);
                            maxFontSize = Math.Max(maxFontSize, fragmentFontSize);
                            double chunkX = cursorX + chunkBoundaryAdjustment;
                            double chunkClipX = frame.Orientation == PptxTextOrientation.Horizontal ? columnClipX : frame.TextClipX;
                            double chunkClipWidth = frame.Orientation == PptxTextOrientation.Horizontal ? columnClipWidth : frame.TextClipWidth;
                            TextRun textRun = new(chunk, chunkX, cursorY, PptxTextMetricRules.MinimumWidth(chunkWidth), frame.TextHeight, chunkClipX, frame.TextClipY, chunkClipWidth, frame.TextClipHeight, fragmentFontSize, runStyle.CharacterSpacing, runStyle.BaselineOffset, runStyle.Color, runStyle.Alpha, runStyle.Highlight, runStyle.Bold, runStyle.Italic, runStyle.Underline, runStyle.Strike, runStyle.KerningEnabled, paragraphStyle.Alignment, runStyle.Typeface, frame.TextRotationDegrees, frame.RotationCenterX, frame.RotationCenterY, FlipHorizontal: false, FlipVertical: false, flowSegment.PreventCoalesce, Outline: runStyle.Outline, StrictClip: strictClip);
                            double chunkLeadingAdjustment = pendingVisibleLeadingAdjustment + chunkBoundaryAdjustment;
                            EndWrapTail();
                            line.Add(modelRun, textRun, cursorX + chunkTotalWidth, BuildTextAtoms(textRun, advanceEstimator, null), BuildGlyphSpan(textRun, advanceEstimator, chunkLeadingAdjustment));
                            pendingVisibleLeadingAdjustment = 0d;
                            cursorX += chunkTotalWidth;
                            line.AdvanceTo(cursorX);
                            if (chunkIndex < chunks.Length - 1)
                            {
                                double lineFontSize = ResolveLineFontSize(maxFontSize, paragraphStyle.FontSize);
                                double lineTextX = frame.Orientation == PptxTextOrientation.Horizontal ? columnStartX : frame.TextX;
                                double lineTextWidth = frame.Orientation == PptxTextOrientation.Horizontal ? effectiveTextWidth : frame.TextWidth;
                                AddClippedParagraphLine(lineLayouts, line, CreateLineBox(cursorLineTop, cursorY, paragraphStyle.LineSpacing, lineFontSize, line, advanceEstimator, frame.UseOfficeBaselineFloor), paragraphStyle.Alignment, lineTextX, lineTextWidth, justify: false, distribute: paragraphStyle.Alignment == TextAlignment.Distributed, advanceEstimator, cullOutOfFrameLines, cursorY, frame.TextClipY, frame.TextClipHeight, suppressNoWrapDistribution);
                                double lineAdvance = ReadLineAdvance(paragraphStyle.LineSpacing, lineFontSize);
                                AdvanceTextLine(lineAdvance);
                                cursorY = cursorLineTop - LineBaselineOffset(fragmentFontSize, paragraphStyle.LineSpacing, runStyle, advanceEstimator, frame.UseOfficeBaselineFloor, useExplicitMultipleBaselineOffset);
                                cursorX = paragraphTextX;
                                line.Reset(paragraphTextX);
                                maxFontSize = 0d;
                                previousAdvanceCodePoint = null;
                                pendingVisibleLeadingAdjustment = 0d;
                            }
                        }

                        previousAdvanceCodePoint = LastCodePoint(currentAdvanceText);
                        if (currentAdvanceText.EndsWith(' '))
                        {
                            EndWrapTail();
                        }
                        if (isDrawableTextSegment)
                        {
                            remainingDrawableSegments--;
                        }

                        continue;
                    }

                    bool usesCenteredShapeAutoFit = shapeAutoFit &&
                        paragraphStyle.Alignment == TextAlignment.Center;
                    double wrapTolerance = usesCenteredShapeAutoFit
                        ? PptxTextMetricRules.CoordinateTolerance
                        : IsCenteredTableCellText(frame, paragraphStyle)
                        ? PptxTextMetricRules.CenteredTableCellWrapTolerance(fragmentFontSize, effectiveTextWidth)
                        : useNumberingWordWrap
                        ? PptxTextMetricRules.CoordinateTolerance
                        : bulletText is not null
                        ? PptxTextMetricRules.BulletWrapFitTolerance(fragmentFontSize)
                        : shapeAutoFit
                        ? PptxTextMetricRules.ShapeAutoFitWrapTolerance(fragmentFontSize, effectiveTextWidth)
                        : PptxTextMetricRules.CoordinateTolerance;
                    if (isFinalShortWordSegment && shapeAutoFit && !usesCenteredShapeAutoFit)
                    {
                        wrapTolerance = PptxTextMetricRules.FinalWordWrapTolerance(fragmentFontSize, effectiveTextWidth);
                    }

                    // Check a no-break space with the following fragment so a
                    // fitting phrase can move together.
                    bool overflowsLine = allowWrapping && !isNoBreakHiddenAdvance &&
                        cursorX > paragraphTextX &&
                        cursorX + segmentWidth > columnStartX + effectiveTextWidth + wrapTolerance;
                    if (overflowsLine)
                    {
                        List<PptxTextSpanLayout>? movedNoBreakSpans = null;
                        double movedNoBreakAdvance = 0d;
                        if (isDrawableTextSegment && trailingHasNoBreakSpace &&
                            trailingStartSpan > 0 && trailingStartSpan < line.Spans.Count)
                        {
                            movedNoBreakAdvance = cursorX - line.Spans[trailingStartSpan].Run.X;
                            double continuationWidth = columnStartX + effectiveTextWidth - paragraphTextX;
                            if (movedNoBreakAdvance + segmentWidth <= continuationWidth + wrapTolerance)
                            {
                                movedNoBreakSpans = line.RemoveTail(trailingStartSpan);
                                cursorX = line.EndX;
                            }
                        }

                        if (flowSegment.Draw)
                        {
                            int leadingSpaceCount = CountLeadingSpaces(currentSegment);
                            if (leadingSpaceCount > 0)
                            {
                                string lineEndSpaces = currentSegment[..leadingSpaceCount];
                                double lineEndSpaceWidth = advanceEstimator.Measure(lineEndSpaces, fragmentFontSize, runStyle.Typeface, runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing, runStyle.KerningEnabled);
                                TextRun spaceRun = new(lineEndSpaces, cursorX, cursorY, PptxTextMetricRules.MinimumWidth(lineEndSpaceWidth), frame.TextHeight, columnClipX, frame.TextClipY, columnClipWidth, frame.TextClipHeight, fragmentFontSize, runStyle.CharacterSpacing, runStyle.BaselineOffset, runStyle.Color, runStyle.Alpha, runStyle.Highlight, runStyle.Bold, runStyle.Italic, runStyle.Underline, runStyle.Strike, runStyle.KerningEnabled, paragraphStyle.Alignment, runStyle.Typeface, frame.TextRotationDegrees, frame.RotationCenterX, frame.RotationCenterY, FlipHorizontal: false, FlipVertical: false, PreventCoalesce: false, Outline: runStyle.Outline, StrictClip: strictClip);
                                line.Add(modelRun, spaceRun, cursorX + lineEndSpaceWidth, BuildTextAtoms(spaceRun, advanceEstimator, PptxTextAtomKind.Space), BuildGlyphSpan(spaceRun, advanceEstimator, 0d));
                                cursorX += lineEndSpaceWidth;
                                line.AdvanceTo(cursorX);
                            }
                        }

                        // RV03: words split across runs break at word boundaries, not run
                        // seams. When the overflowing segment continues a word begun on this
                        // line and the whole word fits the line, detach the word tail so the
                        // standard break below emits without it; re-added after the break.
                        List<PptxTextSpanLayout>? pulledWordSpans = null;
                        // The no-break phrase owns the trailing spans, even when
                        // it is too wide to fit on the next line.
                        if (!trailingHasNoBreakSpace && flowSegment.Draw &&
                            currentAdvanceText.Length != 0 &&
                            WordContinuesOntoLine(previousAdvanceCodePoint, currentAdvanceText) &&
                            TryFindTrailingWordStart(line.Spans, out int wordStartSpan, out int wordLeadingSpaces) &&
                            wordStartSpan > 0 &&
                            cursorX - line.Spans[wordStartSpan].Run.X + segmentWidth <= effectiveTextWidth + wrapTolerance)
                        {
                            pulledWordSpans = DetachTrailingWordSpans(line, wordStartSpan, wordLeadingSpaces, ref cursorX, ref previousAdvanceCodePoint, advanceEstimator);
                        }

                        double lineFontSize = ResolveLineFontSize(maxFontSize, paragraphStyle.FontSize);
                        AddClippedParagraphLine(lineLayouts, line, CreateLineBox(cursorLineTop, cursorY, paragraphStyle.LineSpacing, lineFontSize, line, advanceEstimator, frame.UseOfficeBaselineFloor), paragraphStyle.Alignment, columnStartX, effectiveTextWidth, justify: IsWordJustifiedAlignment(paragraphStyle.Alignment), distribute: paragraphStyle.Alignment == TextAlignment.Distributed, advanceEstimator, cullOutOfFrameLines, cursorY, frame.TextClipY, frame.TextClipHeight, suppressNoWrapDistribution);
                        double lineAdvance = ReadLineAdvance(paragraphStyle.LineSpacing, lineFontSize);
                        AdvanceTextLine(lineAdvance);
                        cursorY = cursorLineTop - LineBaselineOffset(fragmentFontSize, paragraphStyle.LineSpacing, runStyle, advanceEstimator, frame.UseOfficeBaselineFloor, useExplicitMultipleBaselineOffset);
                        cursorX = paragraphTextX;
                        line.Reset(paragraphTextX);
                        maxFontSize = 0d;
                        int? movedPreviousCodePoint = previousAdvanceCodePoint;
                        double movedLeadingAdjustment = pendingVisibleLeadingAdjustment;
                        previousAdvanceCodePoint = null;
                        pendingVisibleLeadingAdjustment = 0d;
                        EndWrapTail();
                        if (movedNoBreakSpans is not null)
                        {
                            // Translate the measured spans and the gaps between
                            // them, including advances owned by other source runs.
                            double shiftX = cursorX - movedNoBreakSpans[0].Run.X;
                            for (int index = 0; index < movedNoBreakSpans.Count; index++)
                            {
                                PptxTextSpanLayout span = movedNoBreakSpans[index];
                                TextRun movedRun = span.Run with { X = span.Run.X + shiftX, Y = cursorY, ClipX = columnClipX, ClipWidth = columnClipWidth };
                                var atoms = span.Atoms.Select(atom => atom with { X = atom.X + shiftX }).ToArray();
                                var glyphs = index == 0 ? span.GlyphSpan with { LeadingAdjustment = 0d } : span.GlyphSpan;
                                line.Add(span.SourceRun, movedRun, span.EndX + shiftX, atoms, glyphs);
                                maxFontSize = Math.Max(maxFontSize, movedRun.FontSize);
                            }

                            cursorX += movedNoBreakAdvance;
                            line.AdvanceTo(cursorX);
                            previousAdvanceCodePoint = movedPreviousCodePoint;
                            pendingVisibleLeadingAdjustment = movedLeadingAdjustment;
                            trailingHasNoBreakSpace = true;
                        }

                        if (pulledWordSpans is not null)
                        {
                            foreach (PptxTextSpanLayout wordSpan in pulledWordSpans)
                            {
                                TextRun wordRun = wordSpan.Run;
                                double wordBoundary = MeasureFlowSegmentBoundaryAdjustment(advanceEstimator, wordRun.Text, previousAdvanceCodePoint, wordRun.FontSize, new TextAdvanceOptions(wordRun.FontFamily, wordRun.Bold, wordRun.Italic, wordRun.CharacterSpacing, wordRun.KerningEnabled));
                                double wordIntrinsic = advanceEstimator.Measure(wordRun.Text, wordRun.FontSize, wordRun.FontFamily, wordRun.Bold, wordRun.Italic, wordRun.CharacterSpacing, wordRun.KerningEnabled);
                                double wordWidth = Math.Max(0d, wordIntrinsic + wordBoundary);
                                double wordLeading = pendingVisibleLeadingAdjustment + wordBoundary;
                                TextRun placedRun = wordRun with { X = cursorX + wordBoundary, Y = cursorY, Width = wordWidth, ClipX = columnClipX, ClipWidth = columnClipWidth };
                                double wordEndX = cursorX + wordWidth;
                                line.Add(wordSpan.SourceRun, placedRun, wordEndX, BuildTextAtoms(placedRun, advanceEstimator, null), BuildGlyphSpan(placedRun, advanceEstimator, wordLeading));
                                maxFontSize = Math.Max(maxFontSize, placedRun.FontSize);
                                cursorX = wordEndX;
                                line.AdvanceTo(cursorX);
                                pendingVisibleLeadingAdjustment = 0d;
                                previousAdvanceCodePoint = LastCodePoint(wordRun.Text);
                            }
                        }

                        currentSegment = currentSegment.TrimStart();
                        currentAdvanceText = currentAdvanceText.TrimStart();
                        segmentIntrinsicWidth = advanceEstimator.Measure(currentAdvanceText, fragmentFontSize, runStyle.Typeface, runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing, runStyle.KerningEnabled);
                        segmentBoundaryAdjustment = MeasureFlowSegmentBoundaryAdjustment(advanceEstimator, currentAdvanceText, previousAdvanceCodePoint, fragmentFontSize, new TextAdvanceOptions(runStyle.Typeface, runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing, runStyle.KerningEnabled));
                        segmentWidth = Math.Max(0d, segmentIntrinsicWidth + segmentBoundaryAdjustment);
                    }

                    if (currentAdvanceText.Length == 0)
                    {
                        continue;
                    }

                    if (flowSegment.Draw && currentSegment.Length != 0)
                    {
                        maxFontSize = Math.Max(maxFontSize, fragmentFontSize);
                        double textRunX = cursorX + segmentBoundaryAdjustment;
                        TextRun textRun = new(currentSegment, textRunX, cursorY, PptxTextMetricRules.MinimumWidth(segmentIntrinsicWidth), frame.TextHeight, columnClipX, frame.TextClipY, columnClipWidth, frame.TextClipHeight, fragmentFontSize, runStyle.CharacterSpacing, runStyle.BaselineOffset, runStyle.Color, runStyle.Alpha, runStyle.Highlight, runStyle.Bold, runStyle.Italic, runStyle.Underline, runStyle.Strike, runStyle.KerningEnabled, paragraphStyle.Alignment, runStyle.Typeface, frame.TextRotationDegrees, frame.RotationCenterX, frame.RotationCenterY, FlipHorizontal: false, FlipVertical: false, flowSegment.PreventCoalesce, Outline: runStyle.Outline, StrictClip: strictClip);
                        double leadingAdjustment = pendingVisibleLeadingAdjustment + segmentBoundaryAdjustment;
                        line.Add(modelRun, textRun, cursorX + segmentWidth, BuildTextAtoms(textRun, advanceEstimator, null), BuildGlyphSpan(textRun, advanceEstimator, leadingAdjustment));
                        pendingVisibleLeadingAdjustment = 0d;
                    }
                    else
                    {
                        pendingVisibleLeadingAdjustment += segmentBoundaryAdjustment;
                    }

                    cursorX += segmentWidth;
                    line.AdvanceTo(cursorX);
                    previousAdvanceCodePoint = LastCodePoint(currentAdvanceText);
                    if (isNoBreakHiddenAdvance && trailingStartSpan < line.Spans.Count)
                    {
                        trailingHasNoBreakSpace = true;
                    }
                    if (currentAdvanceText.EndsWith(' '))
                    {
                        EndWrapTail();
                    }
                    if (isDrawableTextSegment)
                    {
                        remainingDrawableSegments--;
                    }
                }
            }

            double paragraphLineFontSize = ResolveLineFontSize(maxFontSize, paragraphStyle.FontSize);
            if (afterManualLineBreak && line.Spans.Count == 0)
            {
                paragraphLineFontSize = paragraph.EndParagraphProperties is null
                    ? paragraphLineFontSize
                    : paragraph.EndParagraphStyle.FontSize;
            }

            if (double.IsNaN(cursorY))
            {
                cursorY = cursorLineTop - (afterManualLineBreak
                    ? ManualBreakBaselineOffset(paragraphLineFontSize, paragraphStyle.LineSpacing, useExplicitMultipleBaselineOffset)
                    : LineBaselineOffset(paragraphLineFontSize, paragraphStyle.LineSpacing, useExplicitMultipleBaselineOffset));
            }

            AddClippedParagraphLine(lineLayouts, line, CreateLineBox(cursorLineTop, cursorY, paragraphStyle.LineSpacing, paragraphLineFontSize, line, advanceEstimator, frame.UseOfficeBaselineFloor), paragraphStyle.Alignment, columnStartX, effectiveTextWidth, justify: false, distribute: paragraphStyle.Alignment == TextAlignment.Distributed, advanceEstimator, cullOutOfFrameLines, cursorY, frame.TextClipY, frame.TextClipHeight, suppressNoWrapDistribution);
            double paragraphAdvance = ReadParagraphAdvance(paragraphStyle.LineSpacing, paragraphLineFontSize);
            cursorLineTop -= paragraphAdvance + paragraphStyle.SpacingAfter;
            MoveColumn(paragraphAdvance, linePlaced: true);
            hasPlacedParagraph = true;
            paragraphLayouts.Add(new PptxTextParagraphLayout(paragraph, lineLayouts));

            void AdvanceTextLine(double advance)
            {
                cursorLineTop -= advance;
                MoveColumn(advance, linePlaced: true);
                (columnClipX, columnClipWidth) = ReadColumnClip();
                paragraphTextX = columnStartX + paragraphInset;
            }

            void CommitManualBreak(double breakFontSize)
            {
                double lineFontSize = ResolveLineFontSize(maxFontSize, breakFontSize);
                bool leadingManualBreak = line.Spans.Count == 0;
                bool useManualBreakFallback = (leadingManualBreak || !shapeAutoFit) && !frame.BodyProperties.CompatibleLineSpacing;
                AddClippedParagraphLine(lineLayouts, line,
                    CreateLineBox(cursorLineTop, cursorY, paragraphStyle.LineSpacing, lineFontSize, line, advanceEstimator, frame.UseOfficeBaselineFloor),
                    paragraphStyle.Alignment, columnStartX, effectiveTextWidth, justify: false, distribute: false,
                    advanceEstimator, cullOutOfFrameLines, cursorY, frame.TextClipY, frame.TextClipHeight, suppressNoWrapDistribution);
                double lineAdvance = useManualBreakFallback
                    ? ReadManualBreakLineAdvance(paragraphStyle.LineSpacing, lineFontSize)
                    : ReadLineAdvance(paragraphStyle.LineSpacing, lineFontSize);
                AdvanceTextLine(lineAdvance);
                cursorY = double.NaN;
                afterManualLineBreak = true;
                afterLeadingManualLineBreak = useManualBreakFallback;
                cursorX = paragraphTextX;
                line.Reset(paragraphTextX);
                maxFontSize = 0d;
                previousAdvanceCodePoint = null;
                pendingVisibleLeadingAdjustment = 0d;
                EndWrapTail();
            }

            void EndWrapTail()
            {
                trailingStartSpan = line.Spans.Count;
                trailingHasNoBreakSpace = false;
            }

            void ResumeAfterManualBreak(ResolvedRunTextStyle runStyle)
            {
                if (!double.IsNaN(cursorY)) return;

                cursorY = cursorLineTop - (afterLeadingManualLineBreak
                    ? ManualBreakBaselineOffset(runStyle.NominalFontSize, paragraphStyle.LineSpacing, useExplicitMultipleBaselineOffset)
                    : LineBaselineOffset(runStyle.NominalFontSize, paragraphStyle.LineSpacing, runStyle, advanceEstimator, frame.UseOfficeBaselineFloor, useExplicitMultipleBaselineOffset));
                afterManualLineBreak = false;
                afterLeadingManualLineBreak = false;
            }
        }

        return new PptxTextFrameLayout(frame, paragraphLayouts);

        (double X, double Width) ReadColumnClip() => clipsColumnsIndividually
            ? (columnStartX, columnWidth)
            : (frame.TextClipX, frame.TextClipWidth);

        void MoveColumn(double advance, bool linePlaced) => MoveToNextColumnIfNeeded(
            ref cursorLineTop, ref columnIndex, ref columnStartX, ref linesInCurrentColumn,
            flowFrame.Box.CursorTop, frame.TextX, columnWidth, frame.ColumnSpacing, frame.ColumnCount,
            flowFrame.Box, frame.BodyProperties.VerticalOverflow, columnBreakMode, advance,
            lineBalanceTarget, lineBalanceStartColumn, linePlaced);
    }
}
