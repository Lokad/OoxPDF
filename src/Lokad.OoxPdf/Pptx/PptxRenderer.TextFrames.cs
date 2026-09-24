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
    private static PptxTextFrameLayout BuildTextFrameLayout(PptxTextFlowFrame flowFrame, PptxDocument document, TextAdvanceEstimator advanceEstimator, bool allowWrapping)
    {
        PptxTextFrameLayout layout = BuildTextFrameLayout(flowFrame, document, advanceEstimator, allowWrapping, PptxTextColumnBreakMode.StrictFit, 0, 0);
        if (TryResolveOfficeOverflowColumnLineBalance(layout, out int lineBalanceTarget, out int lineBalanceStartColumn))
        {
            return BuildTextFrameLayout(flowFrame, document, advanceEstimator, allowWrapping, PptxTextColumnBreakMode.LineCountBalance, lineBalanceTarget, lineBalanceStartColumn);
        }

        return ShouldUseOfficeOverflowColumnBalance(layout)
            ? BuildTextFrameLayout(flowFrame, document, advanceEstimator, allowWrapping, PptxTextColumnBreakMode.OverflowBalance, 0, 0)
            : layout;
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
    private static List<(PptxTextRunModel? SourceRun, TextRun Run)>? DetachTrailingWordSpans(
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
        var tails = new List<(PptxTextRunModel? SourceRun, TextRun Run)>();
        while (line.Spans.Count > wordStartSpan && line.TryRemoveLastSpan(out PptxTextSpanLayout? removed))
        {
            tails.Add((removed.SourceRun, removed.Run));
        }
        tails.Reverse();
        if (tails.Count == 0)
        {
            return null;
        }
        (PptxTextRunModel? firstSource, TextRun firstRun) = tails[0];
        string tailText = firstRun.Text[wordLeadingSpaces..];
        if (wordLeadingSpaces > 0)
        {
            string headText = firstRun.Text[..wordLeadingSpaces];
            int? headPrevious = LastCodePoint(line.Spans[wordStartSpan - 1].Run.Text);
            double headIntrinsic = advanceEstimator.Measure(headText, firstRun.FontSize, firstRun.FontFamily, firstRun.Bold, firstRun.Italic, firstRun.CharacterSpacing, firstRun.KerningEnabled);
            double headBoundary = MeasureFlowSegmentBoundaryAdjustment(advanceEstimator, headText, headPrevious, firstRun.FontSize, new TextAdvanceOptions(firstRun.FontFamily, firstRun.Bold, firstRun.Italic, firstRun.CharacterSpacing, firstRun.KerningEnabled));
            double headWidth = Math.Max(0d, headIntrinsic + headBoundary);
            TextRun headRun = firstRun with { Text = headText, X = line.EndX, Width = headWidth };
            line.Add(firstSource, headRun, line.EndX + headWidth, BuildTextAtoms(headRun, advanceEstimator, PptxTextAtomKind.Space), BuildGlyphSpan(headRun, advanceEstimator, 0d));
            cursorX = line.EndX;
            previousAdvanceCodePoint = LastCodePoint(headText);
        }
        tails[0] = (firstSource, firstRun with { Text = tailText });
        return tails;
    }

    private static PptxTextFrameLayout BuildTextFrameLayout(PptxTextFlowFrame flowFrame, PptxDocument document, TextAdvanceEstimator advanceEstimator, bool allowWrapping, PptxTextColumnBreakMode columnBreakMode, int lineBalanceTarget, int lineBalanceStartColumn)
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
        bool strictClip = ClipsTextVerticalOverflow(frame.BodyProperties.VerticalOverflow);
        bool cullOutOfFrameLines = frame.BodyProperties.VerticalOverflow == PptxTextVerticalOverflow.Clip;
        int autoNumberValue = 1;
        bool hasPlacedParagraph = false;
        var paragraphLayouts = new List<PptxTextParagraphLayout>();

        foreach (PptxTextFlowParagraph flowParagraph in flowFrame.Paragraphs)
        {
            PptxTextParagraphModel paragraph = flowParagraph.Model;
            var lineLayouts = new List<PptxTextLineLayout>();
            ResolvedParagraphTextStyle paragraphStyle = flowParagraph.Style;
            if (!paragraph.HasVisibleContent)
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
            }

            string? bulletText = ReadBulletText(paragraph.Bullet, ref autoNumberValue);
            bool bulletPending = bulletText is not null;
            double effectiveTextWidth = columnWrapWidth;
            double bulletX = columnStartX + PptxTextMetricRules.ClampNonNegative(paragraphStyle.Indent.MarginLeft + paragraphStyle.Indent.Hanging);
            double paragraphTextX = bulletText is null
                ? columnStartX + PptxTextMetricRules.ClampNonNegative(paragraphStyle.Indent.MarginLeft + paragraphStyle.Indent.Hanging)
                : columnStartX + PptxTextMetricRules.ClampNonNegative(paragraphStyle.Indent.MarginLeft);
            bool clipsLocally = frame.TextClipX != 0d ||
                frame.TextClipWidth < document.SlideWidthPoints - PptxTextMetricRules.CoordinateTolerance;
            bool clipsColumnsIndividually = clipsLocally && !frame.TableRowIndex.HasValue;
            double columnClipX = clipsColumnsIndividually ? columnStartX : frame.TextClipX;
            double columnClipWidth = clipsColumnsIndividually ? columnWidth : frame.TextClipWidth;
            if (hasPlacedParagraph)
            {
                cursorLineTop -= paragraphStyle.SpacingBefore;
                double nextLineAdvance = ReadLineAdvance(paragraphStyle.LineSpacing, paragraphStyle.FontSize);
                MoveToNextColumnIfNeeded(ref cursorLineTop, ref columnIndex, ref columnStartX, ref linesInCurrentColumn, flowFrame.Box.CursorTop, frame.TextX, columnWidth, frame.ColumnSpacing, frame.ColumnCount, flowFrame.Box, frame.BodyProperties.VerticalOverflow, columnBreakMode, nextLineAdvance, lineBalanceTarget, lineBalanceStartColumn, linePlaced: false);
                columnClipX = clipsColumnsIndividually ? columnStartX : frame.TextClipX;
                columnClipWidth = clipsColumnsIndividually ? columnWidth : frame.TextClipWidth;
                bulletX = columnStartX + PptxTextMetricRules.ClampNonNegative(paragraphStyle.Indent.MarginLeft + paragraphStyle.Indent.Hanging);
                paragraphTextX = bulletText is null
                    ? columnStartX + PptxTextMetricRules.ClampNonNegative(paragraphStyle.Indent.MarginLeft + paragraphStyle.Indent.Hanging)
                    : columnStartX + PptxTextMetricRules.ClampNonNegative(paragraphStyle.Indent.MarginLeft);
            }

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
            PptxTextSpanLayout? noBreakAnchorSpan = null;
            string pendingNoBreakAdvanceText = string.Empty;
            int remainingDrawableSegments = CountDrawableTextSegments(flowParagraph);
            foreach (PptxTextFlowRun flowRun in flowParagraph.Runs)
            {
                PptxTextRunModel modelRun = flowRun.Source;
                if (modelRun.Kind == PptxTextRunKind.Break)
                {
                    double lineFontSize = ResolveLineFontSize(maxFontSize, flowRun.Style.FontSize);
                    bool leadingManualBreak = line.Spans.Count == 0;
                    bool useManualBreakFallback = (leadingManualBreak || !shapeAutoFit) && !frame.BodyProperties.CompatibleLineSpacing;
                    AddClippedParagraphLine(lineLayouts, line, CreateLineBox(cursorLineTop, cursorY, paragraphStyle.LineSpacing, lineFontSize, line, advanceEstimator, frame.UseOfficeBaselineFloor), paragraphStyle.Alignment, columnStartX, effectiveTextWidth, justify: false, distribute: false, advanceEstimator, cullOutOfFrameLines, cursorY, frame.TextClipY, frame.TextClipHeight);
                    double lineAdvance = useManualBreakFallback ? ReadManualBreakLineAdvance(paragraphStyle.LineSpacing, lineFontSize) : ReadLineAdvance(paragraphStyle.LineSpacing, lineFontSize);
                    cursorLineTop -= lineAdvance;
                    MoveToNextColumnIfNeeded(ref cursorLineTop, ref columnIndex, ref columnStartX, ref linesInCurrentColumn, flowFrame.Box.CursorTop, frame.TextX, columnWidth, frame.ColumnSpacing, frame.ColumnCount, flowFrame.Box, frame.BodyProperties.VerticalOverflow, columnBreakMode, lineAdvance, lineBalanceTarget, lineBalanceStartColumn, linePlaced: true);
                    columnClipX = clipsColumnsIndividually ? columnStartX : frame.TextClipX;
                    columnClipWidth = clipsColumnsIndividually ? columnWidth : frame.TextClipWidth;
                    paragraphTextX = bulletText is null
                        ? columnStartX + PptxTextMetricRules.ClampNonNegative(paragraphStyle.Indent.MarginLeft + paragraphStyle.Indent.Hanging)
                        : columnStartX + PptxTextMetricRules.ClampNonNegative(paragraphStyle.Indent.MarginLeft);
                    cursorY = double.NaN;
                    afterManualLineBreak = true;
                    afterLeadingManualLineBreak = useManualBreakFallback;
                    cursorX = paragraphTextX;
                    line.Reset(paragraphTextX);
                    maxFontSize = 0d;
                    previousAdvanceCodePoint = null;
                    pendingVisibleLeadingAdjustment = 0d;
                    noBreakAnchorSpan = null;
                    pendingNoBreakAdvanceText = string.Empty;
                    continue;
                }

                ResolvedRunTextStyle runStyle = flowRun.Style;
                if (double.IsNaN(cursorY))
                {
                    cursorY = cursorLineTop - (afterLeadingManualLineBreak
                        ? ManualBreakBaselineOffset(runStyle.NominalFontSize, paragraphStyle.LineSpacing, frame.UseOfficeBaselineFloor, useExplicitMultipleBaselineOffset)
                        : LineBaselineOffset(runStyle.NominalFontSize, paragraphStyle.LineSpacing, runStyle, advanceEstimator, frame.UseOfficeBaselineFloor, useExplicitMultipleBaselineOffset));
                    afterManualLineBreak = false;
                    afterLeadingManualLineBreak = false;
                }

                if (bulletPending && bulletText is not null)
                {
                    BulletStyle bulletStyle = ReadBulletStyle(paragraph.Bullet, runStyle.FontSize, runStyle.Color, runStyle.Typeface);
                    maxFontSize = Math.Max(maxFontSize, bulletStyle.FontSize);
                    double bulletWidth = PptxTextMetricRules.MinimumWidth(effectiveTextWidth - (bulletX - columnStartX));
                    double bulletEndX = bulletX + advanceEstimator.Measure(bulletText, bulletStyle.FontSize, bulletStyle.Typeface, runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing, true);
                    TextRun bulletRun = new(bulletText, bulletX, cursorY, bulletWidth, frame.TextHeight, columnClipX, frame.TextClipY, columnClipWidth, frame.TextClipHeight, bulletStyle.FontSize, runStyle.CharacterSpacing, 0d, bulletStyle.Color, 1d, null, runStyle.Bold, runStyle.Italic, runStyle.Underline, runStyle.Strike, runStyle.KerningEnabled, paragraphStyle.Alignment, bulletStyle.Typeface, frame.TextRotationDegrees, frame.RotationCenterX, frame.RotationCenterY, frame.TextFlipHorizontal, frame.TextFlipVertical, PreventCoalesce: false, Outline: null, StrictClip: strictClip);
                    line.Add(modelRun, bulletRun, bulletEndX, BuildTextAtoms(bulletRun, advanceEstimator, PptxTextAtomKind.Word), BuildGlyphSpan(bulletRun, advanceEstimator, 0d));
                    bulletPending = false;
                }

                foreach (PptxTextFlowSegment flowSegment in flowRun.Segments)
                {
                    if (flowSegment.Kind == PptxTextFlowSegmentKind.Break)
                    {
                        double lineFontSize = ResolveLineFontSize(maxFontSize, runStyle.FontSize);
                        bool leadingManualBreak = line.Spans.Count == 0;
                        bool useManualBreakFallback = (leadingManualBreak || !shapeAutoFit) && !frame.BodyProperties.CompatibleLineSpacing;
                        AddClippedParagraphLine(lineLayouts, line, CreateLineBox(cursorLineTop, cursorY, paragraphStyle.LineSpacing, lineFontSize, line, advanceEstimator, frame.UseOfficeBaselineFloor), paragraphStyle.Alignment, columnStartX, effectiveTextWidth, justify: false, distribute: false, advanceEstimator, cullOutOfFrameLines, cursorY, frame.TextClipY, frame.TextClipHeight);
                        double lineAdvance = useManualBreakFallback ? ReadManualBreakLineAdvance(paragraphStyle.LineSpacing, lineFontSize) : ReadLineAdvance(paragraphStyle.LineSpacing, lineFontSize);
                        cursorLineTop -= lineAdvance;
                        MoveToNextColumnIfNeeded(ref cursorLineTop, ref columnIndex, ref columnStartX, ref linesInCurrentColumn, flowFrame.Box.CursorTop, frame.TextX, columnWidth, frame.ColumnSpacing, frame.ColumnCount, flowFrame.Box, frame.BodyProperties.VerticalOverflow, columnBreakMode, lineAdvance, lineBalanceTarget, lineBalanceStartColumn, linePlaced: true);
                        columnClipX = clipsColumnsIndividually ? columnStartX : frame.TextClipX;
                        columnClipWidth = clipsColumnsIndividually ? columnWidth : frame.TextClipWidth;
                        paragraphTextX = bulletText is null
                            ? columnStartX + PptxTextMetricRules.ClampNonNegative(paragraphStyle.Indent.MarginLeft + paragraphStyle.Indent.Hanging)
                            : columnStartX + PptxTextMetricRules.ClampNonNegative(paragraphStyle.Indent.MarginLeft);
                        cursorY = double.NaN;
                        afterManualLineBreak = true;
                        afterLeadingManualLineBreak = useManualBreakFallback;
                        cursorX = paragraphTextX;
                        line.Reset(paragraphTextX);
                        maxFontSize = 0d;
                        previousAdvanceCodePoint = null;
                        pendingVisibleLeadingAdjustment = 0d;
                        noBreakAnchorSpan = null;
                        pendingNoBreakAdvanceText = string.Empty;
                        continue;
                    }

                    if (double.IsNaN(cursorY))
                    {
                        cursorY = cursorLineTop - (afterLeadingManualLineBreak
                            ? ManualBreakBaselineOffset(runStyle.NominalFontSize, paragraphStyle.LineSpacing, frame.UseOfficeBaselineFloor, useExplicitMultipleBaselineOffset)
                            : LineBaselineOffset(runStyle.NominalFontSize, paragraphStyle.LineSpacing, runStyle, advanceEstimator, frame.UseOfficeBaselineFloor, useExplicitMultipleBaselineOffset));
                        afterManualLineBreak = false;
                        afterLeadingManualLineBreak = false;
                    }

                    if (flowSegment.Kind == PptxTextFlowSegmentKind.Tab)
                    {
                        double tabSpaceWidth = advanceEstimator.Measure(" ", runStyle.FontSize, runStyle.Typeface, runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing, runStyle.KerningEnabled);
                        TextRun tabRun = new(" ", cursorX, cursorY, PptxTextMetricRules.MinimumWidth(tabSpaceWidth), frame.TextHeight, columnClipX, frame.TextClipY, columnClipWidth, frame.TextClipHeight, runStyle.FontSize, runStyle.CharacterSpacing, runStyle.BaselineOffset, runStyle.Color, runStyle.Alpha, runStyle.Highlight, runStyle.Bold, runStyle.Italic, runStyle.Underline, runStyle.Strike, runStyle.KerningEnabled, paragraphStyle.Alignment, runStyle.Typeface, frame.TextRotationDegrees, frame.RotationCenterX, frame.RotationCenterY, frame.TextFlipHorizontal, frame.TextFlipVertical, PreventCoalesce: true, Outline: runStyle.Outline, StrictClip: strictClip);
                        line.Add(modelRun, tabRun, cursorX + tabSpaceWidth, BuildTextAtoms(tabRun, advanceEstimator, PptxTextAtomKind.Tab), BuildGlyphSpan(tabRun, advanceEstimator, 0d));
                        cursorX = ResolveNextTabX(cursorX, paragraphTextX, paragraphStyle.TabStops);
                        line.AdvanceTo(cursorX);
                        previousAdvanceCodePoint = null;
                        pendingVisibleLeadingAdjustment = 0d;
                        noBreakAnchorSpan = null;
                        pendingNoBreakAdvanceText = string.Empty;
                        continue;
                    }

                    double fragmentFontSize = runStyle.FontSize * flowSegment.FontScale;
                    string currentSegment = flowSegment.Text;
                    string currentAdvanceText = flowSegment.AdvanceText;
                    bool isNoBreakHiddenAdvance = flowSegment.Kind == PptxTextFlowSegmentKind.NoBreakHiddenAdvance;
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
                            AddClippedParagraphLine(lineLayouts, line, CreateLineBox(cursorLineTop, cursorY, paragraphStyle.LineSpacing, freshLineFontSize, line, advanceEstimator, frame.UseOfficeBaselineFloor), paragraphStyle.Alignment, columnStartX, effectiveTextWidth, justify: false, distribute: paragraphStyle.Alignment == TextAlignment.Distributed, advanceEstimator, cullOutOfFrameLines, cursorY, frame.TextClipY, frame.TextClipHeight);
                            double freshLineAdvance = ReadLineAdvance(paragraphStyle.LineSpacing, freshLineFontSize);
                            cursorLineTop -= freshLineAdvance;
                            MoveToNextColumnIfNeeded(ref cursorLineTop, ref columnIndex, ref columnStartX, ref linesInCurrentColumn, flowFrame.Box.CursorTop, frame.TextX, columnWidth, frame.ColumnSpacing, frame.ColumnCount, flowFrame.Box, frame.BodyProperties.VerticalOverflow, columnBreakMode, freshLineAdvance, lineBalanceTarget, lineBalanceStartColumn, linePlaced: true);
                            columnClipX = clipsColumnsIndividually ? columnStartX : frame.TextClipX;
                            columnClipWidth = clipsColumnsIndividually ? columnWidth : frame.TextClipWidth;
                            paragraphTextX = bulletText is null
                                ? columnStartX + PptxTextMetricRules.ClampNonNegative(paragraphStyle.Indent.MarginLeft + paragraphStyle.Indent.Hanging)
                                : columnStartX + PptxTextMetricRules.ClampNonNegative(paragraphStyle.Indent.MarginLeft);
                            cursorY = cursorLineTop - LineBaselineOffset(fragmentFontSize, paragraphStyle.LineSpacing, runStyle, advanceEstimator, frame.UseOfficeBaselineFloor, useExplicitMultipleBaselineOffset);
                            cursorX = paragraphTextX;
                            line.Reset(paragraphTextX);
                            maxFontSize = 0d;
                            previousAdvanceCodePoint = null;
                            pendingVisibleLeadingAdjustment = 0d;
                            noBreakAnchorSpan = null;
                            pendingNoBreakAdvanceText = string.Empty;
                        }

                        double chunkMaxWidth = frame.Orientation == PptxTextOrientation.Horizontal
                            ? effectiveTextWidth
                            : frame.TextWidth;
                        // Vertical chunks split strictly at the column edge: Office breaks them
                        // tighter than the wrap tolerance ('tic' splits at +2.19 over in a 21.81 column).
                        double chunkFitTolerance = frame.Orientation == PptxTextOrientation.Vertical
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
                            TextRun textRun = new(chunk, chunkX, cursorY, PptxTextMetricRules.MinimumWidth(chunkWidth), frame.TextHeight, chunkClipX, frame.TextClipY, chunkClipWidth, frame.TextClipHeight, fragmentFontSize, runStyle.CharacterSpacing, runStyle.BaselineOffset, runStyle.Color, runStyle.Alpha, runStyle.Highlight, runStyle.Bold, runStyle.Italic, runStyle.Underline, runStyle.Strike, runStyle.KerningEnabled, paragraphStyle.Alignment, runStyle.Typeface, frame.TextRotationDegrees, frame.RotationCenterX, frame.RotationCenterY, frame.TextFlipHorizontal, frame.TextFlipVertical, flowSegment.PreventCoalesce, Outline: runStyle.Outline, StrictClip: strictClip);
                            double chunkLeadingAdjustment = pendingVisibleLeadingAdjustment + chunkBoundaryAdjustment;
                            line.Add(modelRun, textRun, cursorX + chunkTotalWidth, BuildTextAtoms(textRun, advanceEstimator, null), BuildGlyphSpan(textRun, advanceEstimator, chunkLeadingAdjustment));
                            pendingVisibleLeadingAdjustment = 0d;
                            cursorX += chunkTotalWidth;
                            line.AdvanceTo(cursorX);
                            if (chunkIndex < chunks.Length - 1)
                            {
                                double lineFontSize = ResolveLineFontSize(maxFontSize, paragraphStyle.FontSize);
                                double lineTextX = frame.Orientation == PptxTextOrientation.Horizontal ? columnStartX : frame.TextX;
                                double lineTextWidth = frame.Orientation == PptxTextOrientation.Horizontal ? effectiveTextWidth : frame.TextWidth;
                                AddClippedParagraphLine(lineLayouts, line, CreateLineBox(cursorLineTop, cursorY, paragraphStyle.LineSpacing, lineFontSize, line, advanceEstimator, frame.UseOfficeBaselineFloor), paragraphStyle.Alignment, lineTextX, lineTextWidth, justify: false, distribute: paragraphStyle.Alignment == TextAlignment.Distributed, advanceEstimator, cullOutOfFrameLines, cursorY, frame.TextClipY, frame.TextClipHeight);
                                double lineAdvance = ReadLineAdvance(paragraphStyle.LineSpacing, lineFontSize);
                                cursorLineTop -= lineAdvance;
                                MoveToNextColumnIfNeeded(ref cursorLineTop, ref columnIndex, ref columnStartX, ref linesInCurrentColumn, flowFrame.Box.CursorTop, frame.TextX, columnWidth, frame.ColumnSpacing, frame.ColumnCount, flowFrame.Box, frame.BodyProperties.VerticalOverflow, columnBreakMode, lineAdvance, lineBalanceTarget, lineBalanceStartColumn, linePlaced: true);
                                columnClipX = clipsColumnsIndividually ? columnStartX : frame.TextClipX;
                                columnClipWidth = clipsColumnsIndividually ? columnWidth : frame.TextClipWidth;
                                paragraphTextX = bulletText is null
                                    ? columnStartX + PptxTextMetricRules.ClampNonNegative(paragraphStyle.Indent.MarginLeft + paragraphStyle.Indent.Hanging)
                                    : columnStartX + PptxTextMetricRules.ClampNonNegative(paragraphStyle.Indent.MarginLeft);
                                cursorY = cursorLineTop - LineBaselineOffset(fragmentFontSize, paragraphStyle.LineSpacing, runStyle, advanceEstimator, frame.UseOfficeBaselineFloor, useExplicitMultipleBaselineOffset);
                                cursorX = paragraphTextX;
                                line.Reset(paragraphTextX);
                                maxFontSize = 0d;
                                previousAdvanceCodePoint = null;
                                pendingVisibleLeadingAdjustment = 0d;
                                noBreakAnchorSpan = null;
                                pendingNoBreakAdvanceText = string.Empty;
                            }
                        }

                        previousAdvanceCodePoint = LastCodePoint(currentAdvanceText);
                        noBreakAnchorSpan = null;
                        pendingNoBreakAdvanceText = string.Empty;
                        if (isDrawableTextSegment)
                        {
                            remainingDrawableSegments--;
                        }

                        continue;
                    }

                    bool usesCenteredShapeAutoFit = HasShapeAutoFit(frame.BodyProperties) &&
                        paragraphStyle.Alignment == TextAlignment.Center;
                    double wrapTolerance = usesCenteredShapeAutoFit
                        ? PptxTextMetricRules.CoordinateTolerance
                        : IsCenteredTableCellText(frame, paragraphStyle)
                        ? PptxTextMetricRules.CenteredTableCellWrapTolerance(fragmentFontSize, effectiveTextWidth)
                        : bulletText is not null
                        ? PptxTextMetricRules.BulletWrapFitTolerance(fragmentFontSize)
                        : HasShapeAutoFit(frame.BodyProperties)
                        ? PptxTextMetricRules.ShapeAutoFitWrapTolerance(fragmentFontSize, effectiveTextWidth)
                        : !HasNoAutoFit(frame.BodyProperties) ||
                        IsWordJustifiedAlignment(paragraphStyle.Alignment) ||
                        paragraphStyle.Alignment == TextAlignment.Distributed
                        ? PptxTextMetricRules.CoordinateTolerance
                        : PptxTextMetricRules.CoordinateTolerance;
                    if (isFinalShortWordSegment &&
                        HasShapeAutoFit(frame.BodyProperties) &&
                        !usesCenteredShapeAutoFit)
                    {
                        wrapTolerance = PptxTextMetricRules.FinalWordWrapTolerance(fragmentFontSize, effectiveTextWidth);
                    }

                    bool overflowsLine = allowWrapping &&
                        cursorX > paragraphTextX &&
                        cursorX + segmentWidth > columnStartX + effectiveTextWidth + wrapTolerance;
                    if (overflowsLine)
                    {
                        bool movedNoBreakCluster = false;
                        PptxTextSpanLayout? movedNoBreakSpan = null;
                        string movedNoBreakAdvanceText = string.Empty;
                        if (flowSegment.Draw &&
                            noBreakAnchorSpan is not null &&
                            pendingNoBreakAdvanceText.Length != 0 &&
                            line.Spans.Count > 0 &&
                            line.Spans[^1].Equals(noBreakAnchorSpan) &&
                            line.TryRemoveLastSpan(out PptxTextSpanLayout? removedNoBreakSpan))
                        {
                            movedNoBreakCluster = true;
                            movedNoBreakSpan = removedNoBreakSpan;
                            movedNoBreakAdvanceText = pendingNoBreakAdvanceText;
                            cursorX = line.EndX;
                        }

                        if (flowSegment.Draw)
                        {
                            int leadingSpaceCount = CountLeadingSpaces(currentSegment);
                            if (leadingSpaceCount > 0)
                            {
                                string lineEndSpaces = currentSegment[..leadingSpaceCount];
                                double lineEndSpaceWidth = advanceEstimator.Measure(lineEndSpaces, fragmentFontSize, runStyle.Typeface, runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing, runStyle.KerningEnabled);
                                TextRun spaceRun = new(lineEndSpaces, cursorX, cursorY, PptxTextMetricRules.MinimumWidth(lineEndSpaceWidth), frame.TextHeight, columnClipX, frame.TextClipY, columnClipWidth, frame.TextClipHeight, fragmentFontSize, runStyle.CharacterSpacing, runStyle.BaselineOffset, runStyle.Color, runStyle.Alpha, runStyle.Highlight, runStyle.Bold, runStyle.Italic, runStyle.Underline, runStyle.Strike, runStyle.KerningEnabled, paragraphStyle.Alignment, runStyle.Typeface, frame.TextRotationDegrees, frame.RotationCenterX, frame.RotationCenterY, frame.TextFlipHorizontal, frame.TextFlipVertical, PreventCoalesce: false, Outline: runStyle.Outline, StrictClip: strictClip);
                                line.Add(modelRun, spaceRun, cursorX + lineEndSpaceWidth, BuildTextAtoms(spaceRun, advanceEstimator, PptxTextAtomKind.Space), BuildGlyphSpan(spaceRun, advanceEstimator, 0d));
                                cursorX += lineEndSpaceWidth;
                                line.AdvanceTo(cursorX);
                            }
                        }

                        // RV03: words split across runs break at word boundaries, not run
                        // seams. When the overflowing segment continues a word begun on this
                        // line and the whole word fits the line, detach the word tail so the
                        // standard break below emits without it; re-added after the break.
                        List<(PptxTextRunModel? SourceRun, TextRun Run)>? pulledWordSpans = null;
                        if (flowSegment.Draw &&
                            currentAdvanceText.Length != 0 &&
                            WordContinuesOntoLine(previousAdvanceCodePoint, currentAdvanceText) &&
                            TryFindTrailingWordStart(line.Spans, out int wordStartSpan, out int wordLeadingSpaces) &&
                            wordStartSpan > 0 &&
                            cursorX - line.Spans[wordStartSpan].Run.X + segmentWidth <= effectiveTextWidth + wrapTolerance)
                        {
                            pulledWordSpans = DetachTrailingWordSpans(line, wordStartSpan, wordLeadingSpaces, ref cursorX, ref previousAdvanceCodePoint, advanceEstimator);
                        }

                        double lineFontSize = ResolveLineFontSize(maxFontSize, paragraphStyle.FontSize);
                        AddClippedParagraphLine(lineLayouts, line, CreateLineBox(cursorLineTop, cursorY, paragraphStyle.LineSpacing, lineFontSize, line, advanceEstimator, frame.UseOfficeBaselineFloor), paragraphStyle.Alignment, columnStartX, effectiveTextWidth, justify: IsWordJustifiedAlignment(paragraphStyle.Alignment), distribute: paragraphStyle.Alignment == TextAlignment.Distributed, advanceEstimator, cullOutOfFrameLines, cursorY, frame.TextClipY, frame.TextClipHeight);
                        double lineAdvance = ReadLineAdvance(paragraphStyle.LineSpacing, lineFontSize);
                        cursorLineTop -= lineAdvance;
                        MoveToNextColumnIfNeeded(ref cursorLineTop, ref columnIndex, ref columnStartX, ref linesInCurrentColumn, flowFrame.Box.CursorTop, frame.TextX, columnWidth, frame.ColumnSpacing, frame.ColumnCount, flowFrame.Box, frame.BodyProperties.VerticalOverflow, columnBreakMode, lineAdvance, lineBalanceTarget, lineBalanceStartColumn, linePlaced: true);
                        columnClipX = clipsColumnsIndividually ? columnStartX : frame.TextClipX;
                        columnClipWidth = clipsColumnsIndividually ? columnWidth : frame.TextClipWidth;
                        paragraphTextX = bulletText is null
                            ? columnStartX + PptxTextMetricRules.ClampNonNegative(paragraphStyle.Indent.MarginLeft + paragraphStyle.Indent.Hanging)
                            : columnStartX + PptxTextMetricRules.ClampNonNegative(paragraphStyle.Indent.MarginLeft);
                        cursorY = cursorLineTop - LineBaselineOffset(fragmentFontSize, paragraphStyle.LineSpacing, runStyle, advanceEstimator, frame.UseOfficeBaselineFloor, useExplicitMultipleBaselineOffset);
                        cursorX = paragraphTextX;
                        line.Reset(paragraphTextX);
                        maxFontSize = 0d;
                        previousAdvanceCodePoint = null;
                        pendingVisibleLeadingAdjustment = 0d;
                        noBreakAnchorSpan = null;
                        pendingNoBreakAdvanceText = string.Empty;
                        if (movedNoBreakCluster && movedNoBreakSpan is not null)
                        {
                            TextRun movedRun = movedNoBreakSpan.Run with { X = cursorX, Y = cursorY };
                            double movedEndX = cursorX + movedRun.Width;
                            maxFontSize = Math.Max(maxFontSize, movedRun.FontSize);
                            line.Add(
                                movedNoBreakSpan.SourceRun,
                                movedRun,
                                movedEndX,
                                BuildTextAtoms(movedRun, advanceEstimator, null),
                                BuildGlyphSpan(movedRun, advanceEstimator, 0d));
                            cursorX = movedEndX;
                            previousAdvanceCodePoint = LastCodePoint(movedRun.Text);

                            double hiddenIntrinsicWidth = advanceEstimator.Measure(movedNoBreakAdvanceText, fragmentFontSize, runStyle.Typeface, runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing, runStyle.KerningEnabled);
                            double hiddenBoundaryAdjustment = MeasureFlowSegmentBoundaryAdjustment(advanceEstimator, movedNoBreakAdvanceText, previousAdvanceCodePoint, fragmentFontSize, new TextAdvanceOptions(runStyle.Typeface, runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing, runStyle.KerningEnabled));
                            double hiddenWidth = Math.Max(0d, hiddenIntrinsicWidth + hiddenBoundaryAdjustment);
                            pendingVisibleLeadingAdjustment = hiddenBoundaryAdjustment;
                            cursorX += hiddenWidth;
                            line.AdvanceTo(cursorX);
                            previousAdvanceCodePoint = LastCodePoint(movedNoBreakAdvanceText);
                        }

                                                if (pulledWordSpans is not null)
                        {
                            foreach ((PptxTextRunModel? wordSource, TextRun wordRun) in pulledWordSpans)
                            {
                                double wordBoundary = MeasureFlowSegmentBoundaryAdjustment(advanceEstimator, wordRun.Text, previousAdvanceCodePoint, wordRun.FontSize, new TextAdvanceOptions(wordRun.FontFamily, wordRun.Bold, wordRun.Italic, wordRun.CharacterSpacing, wordRun.KerningEnabled));
                                double wordIntrinsic = advanceEstimator.Measure(wordRun.Text, wordRun.FontSize, wordRun.FontFamily, wordRun.Bold, wordRun.Italic, wordRun.CharacterSpacing, wordRun.KerningEnabled);
                                double wordWidth = Math.Max(0d, wordIntrinsic + wordBoundary);
                                double wordLeading = pendingVisibleLeadingAdjustment + wordBoundary;
                                TextRun placedRun = wordRun with { X = cursorX + wordBoundary, Y = cursorY, Width = wordWidth, ClipX = columnClipX, ClipWidth = columnClipWidth };
                                double wordEndX = cursorX + wordWidth;
                                line.Add(wordSource, placedRun, wordEndX, BuildTextAtoms(placedRun, advanceEstimator, null), BuildGlyphSpan(placedRun, advanceEstimator, wordLeading));
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
                        TextRun textRun = new(currentSegment, textRunX, cursorY, PptxTextMetricRules.MinimumWidth(segmentIntrinsicWidth), frame.TextHeight, columnClipX, frame.TextClipY, columnClipWidth, frame.TextClipHeight, fragmentFontSize, runStyle.CharacterSpacing, runStyle.BaselineOffset, runStyle.Color, runStyle.Alpha, runStyle.Highlight, runStyle.Bold, runStyle.Italic, runStyle.Underline, runStyle.Strike, runStyle.KerningEnabled, paragraphStyle.Alignment, runStyle.Typeface, frame.TextRotationDegrees, frame.RotationCenterX, frame.RotationCenterY, frame.TextFlipHorizontal, frame.TextFlipVertical, flowSegment.PreventCoalesce, Outline: runStyle.Outline, StrictClip: strictClip);
                        double leadingAdjustment = pendingVisibleLeadingAdjustment + segmentBoundaryAdjustment;
                        line.Add(modelRun, textRun, cursorX + segmentWidth, BuildTextAtoms(textRun, advanceEstimator, null), BuildGlyphSpan(textRun, advanceEstimator, leadingAdjustment));
                        pendingVisibleLeadingAdjustment = 0d;
                        noBreakAnchorSpan = null;
                        pendingNoBreakAdvanceText = string.Empty;
                    }
                    else
                    {
                        pendingVisibleLeadingAdjustment += segmentBoundaryAdjustment;
                    }

                    cursorX += segmentWidth;
                    line.AdvanceTo(cursorX);
                    previousAdvanceCodePoint = LastCodePoint(currentAdvanceText);
                    if (isNoBreakHiddenAdvance)
                    {
                        noBreakAnchorSpan = line.Spans.LastOrDefault();
                        pendingNoBreakAdvanceText = currentAdvanceText;
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
                    ? ManualBreakBaselineOffset(paragraphLineFontSize, paragraphStyle.LineSpacing, frame.UseOfficeBaselineFloor, useExplicitMultipleBaselineOffset)
                    : LineBaselineOffset(paragraphLineFontSize, paragraphStyle.LineSpacing, frame.UseOfficeBaselineFloor, useExplicitMultipleBaselineOffset));
            }

            AddClippedParagraphLine(lineLayouts, line, CreateLineBox(cursorLineTop, cursorY, paragraphStyle.LineSpacing, paragraphLineFontSize, line, advanceEstimator, frame.UseOfficeBaselineFloor), paragraphStyle.Alignment, columnStartX, effectiveTextWidth, justify: false, distribute: paragraphStyle.Alignment == TextAlignment.Distributed, advanceEstimator, cullOutOfFrameLines, cursorY, frame.TextClipY, frame.TextClipHeight);
            double paragraphAdvance = ReadParagraphAdvance(paragraphStyle.LineSpacing, paragraphLineFontSize);
            cursorLineTop -= paragraphAdvance + paragraphStyle.SpacingAfter;
            MoveToNextColumnIfNeeded(ref cursorLineTop, ref columnIndex, ref columnStartX, ref linesInCurrentColumn, flowFrame.Box.CursorTop, frame.TextX, columnWidth, frame.ColumnSpacing, frame.ColumnCount, flowFrame.Box, frame.BodyProperties.VerticalOverflow, columnBreakMode, paragraphAdvance, lineBalanceTarget, lineBalanceStartColumn, linePlaced: true);
            hasPlacedParagraph = true;
            paragraphLayouts.Add(new PptxTextParagraphLayout(paragraph, lineLayouts));
        }

        return new PptxTextFrameLayout(frame, paragraphLayouts);
    }
}
