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
            double cursorY = cursorLineTop - ReadFirstLineBaselineOffset(paragraph, paragraphStyle.LineSpacing, advanceEstimator, frame.UseOfficeBaselineFloor, shapeAutoFit, useExplicitMultipleBaselineOffset);
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
                    bool useManualBreakFallback = leadingManualBreak || !shapeAutoFit;
                    AddAlignedParagraphLine(lineLayouts, line, CreateLineBox(cursorLineTop, cursorY, paragraphStyle.LineSpacing, lineFontSize, line, advanceEstimator, frame.UseOfficeBaselineFloor), paragraphStyle.Alignment, columnStartX, effectiveTextWidth, justify: false, distribute: false, advanceEstimator);
                    double lineAdvance = useManualBreakFallback
                        ? ReadManualBreakLineAdvance(paragraphStyle.LineSpacing, lineFontSize)
                        : ReadLineAdvance(paragraphStyle.LineSpacing, lineFontSize);
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
                        bool useManualBreakFallback = leadingManualBreak || !shapeAutoFit;
                        AddAlignedParagraphLine(lineLayouts, line, CreateLineBox(cursorLineTop, cursorY, paragraphStyle.LineSpacing, lineFontSize, line, advanceEstimator, frame.UseOfficeBaselineFloor), paragraphStyle.Alignment, columnStartX, effectiveTextWidth, justify: false, distribute: false, advanceEstimator);
                        double lineAdvance = useManualBreakFallback
                            ? ReadManualBreakLineAdvance(paragraphStyle.LineSpacing, lineFontSize)
                            : ReadLineAdvance(paragraphStyle.LineSpacing, lineFontSize);
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
                    double segmentBoundaryAdjustment = MeasureFlowSegmentBoundaryAdjustment(advanceEstimator, currentAdvanceText, previousAdvanceCodePoint, fragmentFontSize, runStyle.Typeface, runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing, runStyle.KerningEnabled);
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
                        double chunkMaxWidth = frame.Orientation == PptxTextOrientation.Horizontal
                            ? effectiveTextWidth
                            : frame.TextWidth;
                        double chunkFitTolerance = PptxTextMetricRules.WrapFitTolerance(fragmentFontSize);
                        string[] chunks = SplitTextIntoFittingChunks(currentSegment, chunkMaxWidth + chunkFitTolerance, fragmentFontSize, runStyle, advanceEstimator);
                        for (int chunkIndex = 0; chunkIndex < chunks.Length; chunkIndex++)
                        {
                            string chunk = chunks[chunkIndex];
                            double chunkWidth = advanceEstimator.Measure(chunk, fragmentFontSize, runStyle.Typeface, runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing, runStyle.KerningEnabled);
                            double chunkBoundaryAdjustment = chunkIndex == 0
                                ? MeasureFlowSegmentBoundaryAdjustment(advanceEstimator, chunk, previousAdvanceCodePoint, fragmentFontSize, runStyle.Typeface, runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing, runStyle.KerningEnabled)
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
                                AddAlignedParagraphLine(lineLayouts, line, CreateLineBox(cursorLineTop, cursorY, paragraphStyle.LineSpacing, lineFontSize, line, advanceEstimator, frame.UseOfficeBaselineFloor), paragraphStyle.Alignment, lineTextX, lineTextWidth, justify: false, distribute: paragraphStyle.Alignment == TextAlignment.Distributed, advanceEstimator);
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

                        double lineFontSize = ResolveLineFontSize(maxFontSize, paragraphStyle.FontSize);
                        AddAlignedParagraphLine(lineLayouts, line, CreateLineBox(cursorLineTop, cursorY, paragraphStyle.LineSpacing, lineFontSize, line, advanceEstimator, frame.UseOfficeBaselineFloor), paragraphStyle.Alignment, columnStartX, effectiveTextWidth, justify: IsWordJustifiedAlignment(paragraphStyle.Alignment), distribute: paragraphStyle.Alignment == TextAlignment.Distributed, advanceEstimator);
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
                            double hiddenBoundaryAdjustment = MeasureFlowSegmentBoundaryAdjustment(advanceEstimator, movedNoBreakAdvanceText, previousAdvanceCodePoint, fragmentFontSize, runStyle.Typeface, runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing, runStyle.KerningEnabled);
                            double hiddenWidth = Math.Max(0d, hiddenIntrinsicWidth + hiddenBoundaryAdjustment);
                            pendingVisibleLeadingAdjustment = hiddenBoundaryAdjustment;
                            cursorX += hiddenWidth;
                            line.AdvanceTo(cursorX);
                            previousAdvanceCodePoint = LastCodePoint(movedNoBreakAdvanceText);
                        }

                        currentSegment = currentSegment.TrimStart();
                        currentAdvanceText = currentAdvanceText.TrimStart();
                        segmentIntrinsicWidth = advanceEstimator.Measure(currentAdvanceText, fragmentFontSize, runStyle.Typeface, runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing, runStyle.KerningEnabled);
                        segmentBoundaryAdjustment = MeasureFlowSegmentBoundaryAdjustment(advanceEstimator, currentAdvanceText, previousAdvanceCodePoint, fragmentFontSize, runStyle.Typeface, runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing, runStyle.KerningEnabled);
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

            AddAlignedParagraphLine(lineLayouts, line, CreateLineBox(cursorLineTop, cursorY, paragraphStyle.LineSpacing, paragraphLineFontSize, line, advanceEstimator, frame.UseOfficeBaselineFloor), paragraphStyle.Alignment, columnStartX, effectiveTextWidth, justify: false, distribute: paragraphStyle.Alignment == TextAlignment.Distributed, advanceEstimator);
            double paragraphAdvance = ReadParagraphAdvance(paragraphStyle.LineSpacing, paragraphLineFontSize);
            cursorLineTop -= paragraphAdvance + paragraphStyle.SpacingAfter;
            MoveToNextColumnIfNeeded(ref cursorLineTop, ref columnIndex, ref columnStartX, ref linesInCurrentColumn, flowFrame.Box.CursorTop, frame.TextX, columnWidth, frame.ColumnSpacing, frame.ColumnCount, flowFrame.Box, frame.BodyProperties.VerticalOverflow, columnBreakMode, paragraphAdvance, lineBalanceTarget, lineBalanceStartColumn, linePlaced: true);
            hasPlacedParagraph = true;
            paragraphLayouts.Add(new PptxTextParagraphLayout(paragraph, lineLayouts));
        }

        return new PptxTextFrameLayout(frame, paragraphLayouts);
    }
}
