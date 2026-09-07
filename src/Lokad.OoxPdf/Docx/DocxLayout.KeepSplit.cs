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
    private static bool ShouldKeepParagraphBlockTogether(DocxParagraph paragraph)
    {
        DocxParagraphKeepRules keepRules = paragraph.EffectiveProperties.KeepRules;
        return keepRules.KeepLines == true || keepRules.KeepNext == true;
    }

    private static double ResolveLineHeight(DocxParagraph paragraph, double fontSize, IDocxTextMeasurer? textMeasurer)
    {
        return ResolveLineHeightProfile(paragraph, fontSize, textMeasurer).LineHeight;
    }

    private static DocxLineHeightProfile ResolveLineHeightProfile(DocxParagraph paragraph, double fontSize, IDocxTextMeasurer? textMeasurer)
    {
        DocxEffectiveParagraphProperties effective = paragraph.EffectiveProperties;
        if (effective.LineSpacingPoints is { } exactLineHeight && IsExactLineSpacing(effective))
        {
            return new DocxLineHeightProfile(
                exactLineHeight,
                SingleLineHeight: null,
                ListLabelSingleLineHeight: null,
                BodyWindowsLineHeight: null,
                ListLabelWindowsLineHeight: null,
                EffectiveLineSpacingFactor: null,
                LineSpacingFactorFloorApplied: false,
                Source: DocxLineHeightSource.ExactLineSpacing);
        }

        DocxTextRun? bodyRun = paragraph.Runs.FirstOrDefault();
        DocxTextRun? listLabelRun = paragraph.ListLabel is null
            ? null
            : CreateListLabelRun(paragraph.ListLabel, bodyRun, fontSize);
        IDocxLineMetricsProvider? metricsProvider = textMeasurer as IDocxLineMetricsProvider;
        IDocxStaticTextMetricsProvider? staticMetrics = textMeasurer as IDocxStaticTextMetricsProvider;
        double singleLineHeight = metricsProvider is not null
            ? metricsProvider.MeasureSingleLineHeight(bodyRun, fontSize)
            : fontSize;
        double? listLabelSingleLineHeight = metricsProvider is not null && listLabelRun is not null
            ? metricsProvider.MeasureSingleLineHeight(listLabelRun, listLabelRun.EffectiveProperties.FontSize)
            : null;
        double? bodyWindowsLineHeight = staticMetrics is not null
            ? staticMetrics.MeasureWindowsAscender(bodyRun, fontSize) + staticMetrics.MeasureWindowsDescender(bodyRun, fontSize)
            : null;
        double? listLabelWindowsLineHeight = staticMetrics is not null && listLabelRun is not null
            ? staticMetrics.MeasureWindowsAscender(listLabelRun, listLabelRun.EffectiveProperties.FontSize) +
                staticMetrics.MeasureWindowsDescender(listLabelRun, listLabelRun.EffectiveProperties.FontSize)
            : null;
        if (effective.LineSpacingPoints is { } minimumLineHeight &&
            effective.Spacing.LineRuleValue?.Equals("atLeast", StringComparison.OrdinalIgnoreCase) == true)
        {
            double naturalLineHeight = Math.Max(singleLineHeight, listLabelSingleLineHeight ?? 0d);
            return new DocxLineHeightProfile(
                Math.Max(minimumLineHeight, naturalLineHeight),
                singleLineHeight,
                listLabelSingleLineHeight,
                bodyWindowsLineHeight,
                listLabelWindowsLineHeight,
                EffectiveLineSpacingFactor: null,
                LineSpacingFactorFloorApplied: false,
                DocxLineHeightSource.AtLeastLineSpacing);
        }

        double effectiveLineSpacingFactor = ResolveAutoLineSpacingFactor(paragraph, out bool floorApplied);
        return new DocxLineHeightProfile(
            singleLineHeight * effectiveLineSpacingFactor,
            singleLineHeight,
            listLabelSingleLineHeight,
            bodyWindowsLineHeight,
            listLabelWindowsLineHeight,
            effectiveLineSpacingFactor,
            floorApplied,
            DocxLineHeightSource.BodySingleLineAuto);
    }

    private static bool IsExactLineSpacing(DocxEffectiveParagraphProperties effective)
    {
        return effective.LineSpacingPoints is not null &&
            !string.Equals(effective.Spacing.LineRuleValue, "atLeast", StringComparison.OrdinalIgnoreCase);
    }

    private static double ResolveAutoLineSpacingFactor(DocxParagraph paragraph, out bool floorApplied)
    {
        DocxEffectiveParagraphProperties effective = paragraph.EffectiveProperties;
        // Office floors default-auto lists (no w:line) to the list minimum but honors explicit w:line factors as-authored
        // (line-height probe 2026-09-06: default bullets pitch 14.16 at 10pt; explicit-115 probe: explicit 1.15 pitches 14.04, explicit 1.0 pitches 12.12).
        if (paragraph.ListLabel is not null &&
            effective.LineSpacingPoints is null &&
            effective.Spacing.LineValue is null)
        {
            double effectiveFactor = Math.Max(effective.LineSpacingFactor, WordListMinimumAutoLineSpacingFactor);
            floorApplied = effectiveFactor > effective.LineSpacingFactor;
            return effectiveFactor;
        }

        floorApplied = false;
        return effective.LineSpacingFactor;
    }

    // Word reserves extra top leading on the first line of a list item when the label ascender exceeds the body ascender
    // (style-line probe 2026-09-06: Symbol-bullet first-line gaps pitch 16.32 vs 15.84 body-predicted at 10pt;
    // Calibri-bullet first-line gaps match body prediction, so the excess is label-specific, not per-list).
    private static double ResolveListLabelFirstLineExtraLeading(DocxParagraph paragraph, double fontSize, IDocxTextMeasurer? textMeasurer)
    {
        if (paragraph.ListLabel is null ||
            textMeasurer is not IDocxStaticTextMetricsProvider staticMetrics)
        {
            return 0d;
        }

        if (paragraph.EffectiveProperties.LineSpacingPoints is not null)
        {
            return 0d;
        }

        DocxTextRun? bodyRun = paragraph.Runs.FirstOrDefault();
        DocxTextRun labelRun = CreateListLabelRun(paragraph.ListLabel, bodyRun, fontSize);
        double bodyAscender = staticMetrics.MeasureWindowsAscender(bodyRun, fontSize);
        double labelAscender = staticMetrics.MeasureWindowsAscender(labelRun, labelRun.EffectiveProperties.FontSize);
        return Math.Max(0d, labelAscender - bodyAscender);
    }

    private static double QuantizeTableCellWrappedLineHeight(double lineHeight, int wrappedLineCount)
    {
        double RoundToTwips()
        {
            return Math.Round(lineHeight * 20d, MidpointRounding.AwayFromZero) / 20d;
        }

        return wrappedLineCount > 1 ? RoundToTwips() : lineHeight;
    }

    private static bool ShouldMoveParagraphForWidowControl(
        DocxParagraph paragraph,
        int lineCount,
        double cursorY,
        double lineHeight,
        double marginBottom,
        bool hasPageContent)
    {
        if (paragraph.EffectiveProperties.KeepRules.WidowControl == false ||
            lineCount <= 1 ||
            !hasPageContent)
        {
            return false;
        }

        int fittingLineCount = (int)Math.Floor(Math.Max(0d, cursorY - marginBottom) / lineHeight);
        return fittingLineCount > 0 &&
            fittingLineCount < lineCount &&
            (fittingLineCount == 1 || lineCount - fittingLineCount == 1);
    }

    private static string NormalizeContextualSpacingStyleId(string? styleId)
    {
        return string.IsNullOrWhiteSpace(styleId) ? string.Empty : styleId;
    }

    private static DocxParagraphSpacingProfile ResolveParagraphSpacingProfile(
        DocxParagraph? previousParagraph,
        DocxParagraph paragraph,
        double pendingAfterSpacing,
        double spacingScale)
    {
        DocxEffectiveParagraphProperties effective = paragraph.EffectiveProperties;
        bool ShouldSuppressContextualSpacing()
        {
            DocxEffectiveParagraphProperties currentEffective = paragraph.EffectiveProperties;
            DocxEffectiveParagraphProperties? previousEffective = previousParagraph?.EffectiveProperties;
            return currentEffective.Spacing.ContextualSpacing == true &&
                previousEffective is not null &&
                string.Equals(
                    NormalizeContextualSpacingStyleId(previousEffective.StyleId),
                    NormalizeContextualSpacingStyleId(currentEffective.StyleId),
                    StringComparison.Ordinal);
        }

        bool suppress = ShouldSuppressContextualSpacing();
        double spacingBefore = effective.SpacingBeforePoints * spacingScale;
        double spacingAfter = effective.SpacingAfterPoints * spacingScale;
        double appliedBefore = suppress
            ? 0d
            : Math.Max(pendingAfterSpacing, spacingBefore);
        return new DocxParagraphSpacingProfile(
            pendingAfterSpacing,
            spacingBefore,
            spacingAfter,
            appliedBefore,
            suppress);
    }

    private static DocxKeepBlockEstimate EstimateKeptParagraphBlock(
        IReadOnlyList<DocxBodyElement> elements,
        int elementIndex,
        double availableWidth,
        IDocxTextMeasurer textMeasurer,
        double defaultTabStopPoints,
        int? pageNumber,
        double paragraphSpacingScale)
    {
        if (elements[elementIndex] is not DocxParagraphElement paragraphElement)
        {
            return new DocxKeepBlockEstimate(0d, 0, 0);
        }

        DocxParagraph paragraph = paragraphElement.Paragraph;
        double height = EstimateParagraphContentHeight(paragraph, availableWidth, textMeasurer, defaultTabStopPoints, pageNumber, paragraphSpacingScale);
        int paragraphCount = 1;
        int firstTableRowCount = 0;
        int nextSearchIndex = elementIndex + 1;
        while (paragraph.EffectiveProperties.KeepRules.KeepNext == true &&
            TryFindNextKeepTarget(elements, nextSearchIndex, out int nextIndex, out DocxBodyElement? next))
        {
            if (next is DocxParagraphElement nextParagraph)
            {
                DocxParagraphSpacingProfile spacingProfile = ResolveParagraphSpacingProfile(
                    paragraph,
                    nextParagraph.Paragraph,
                    paragraph.EffectiveProperties.SpacingAfterPoints * paragraphSpacingScale,
                    paragraphSpacingScale);
                height += spacingProfile.AppliedBeforeSpacing;
                height += EstimateParagraphContentHeight(nextParagraph.Paragraph, availableWidth, textMeasurer, defaultTabStopPoints, pageNumber, paragraphSpacingScale);
                paragraphCount++;
                paragraph = nextParagraph.Paragraph;
                nextSearchIndex = nextIndex + 1;
                continue;
            }

            if (next is DocxTableElement nextTable)
            {
                height += paragraph.EffectiveProperties.SpacingAfterPoints * paragraphSpacingScale;
                height += EstimateFirstTableRowHeight(nextTable.Table);
                firstTableRowCount++;
            }

            break;
        }

        return new DocxKeepBlockEstimate(height, paragraphCount, firstTableRowCount);

        double EstimateFirstTableRowHeight(DocxTable table)
        {
            DocxTableRow? row = table.Rows.FirstOrDefault();
            if (row is null)
            {
                return 0d;
            }

            DocxResolvedTableGrid grid = ResolveTableGrid(table, x: 0d, availableWidth, paragraphSpacingScale);
            double[] cellWidths = GetTableRowCellWidths(row, grid.EffectiveColumns, grid.Scale);
            double rowTopPadding = ResolveTableRowTopPadding(row, paragraphSpacingScale);
            double contentHeight = row.Cells
                .Select((cell, columnIndex) => MeasureTableCellContentHeight(cell, cellWidths[columnIndex], textMeasurer, defaultTabStopPoints, rowTopPadding, pageNumber, null, paragraphSpacingScale: paragraphSpacingScale))
                .DefaultIfEmpty(0d)
                .Max();
            return ResolveTableRowHeight(row, contentHeight, paragraphSpacingScale);
        }
    }

    private static bool TryFindNextKeepTarget(IReadOnlyList<DocxBodyElement> elements, int startIndex, out DocxBodyElement? target)
    {
        bool found = TryFindNextKeepTarget(elements, startIndex, out _, out DocxBodyElement? indexedTarget);
        target = indexedTarget;
        return found;
    }

    private static bool TryFindNextKeepTarget(IReadOnlyList<DocxBodyElement> elements, int startIndex, out int targetIndex, out DocxBodyElement? target)
    {
        for (int i = startIndex; i < elements.Count; i++)
        {
            if (elements[i] is DocxParagraphElement or DocxTableElement)
            {
                targetIndex = i;
                target = elements[i];
                return true;
            }

            if (elements[i] is DocxPageBreakElement or DocxManualBreakElement or DocxSectionBreakElement)
            {
                break;
            }
        }

        targetIndex = -1;
        target = null;
        return false;
    }

    private static double EstimateParagraphContentHeight(DocxParagraph paragraph, double availableWidth, IDocxTextMeasurer textMeasurer, double defaultTabStopPoints, int? pageNumber, double fixedScale)
    {
        double height = 0d;
        double fontSize = GetParagraphFontSize(paragraph);
        double lineHeight = ResolveLineHeight(paragraph, fontSize, textMeasurer);
        IReadOnlyList<DocxTextSpan> textSpans = CreateTextSpans(paragraph.Runs, pageNumber, null);
        if (textSpans.Count != 0)
        {
            double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, fontSize, textMeasurer, fixedScale);
            double firstParagraphWidth = Math.Max(1d, availableWidth - textStartOffset - GetParagraphRightInset(paragraph, fixedScale));
            double continuationParagraphWidth = Math.Max(1d, availableWidth - GetParagraphTextStartOffset(paragraph, fixedScale) - GetParagraphRightInset(paragraph, fixedScale));
            height += WrapTextLines(textSpans, firstParagraphWidth, continuationParagraphWidth, fontSize, textMeasurer, ScaleTabStopPositions(paragraph.EffectiveProperties.TabStops, fixedScale), defaultTabStopPoints * fixedScale, allowOverwideTokenBreaks: ShouldAllowCharacterLevelWordWrap(paragraph), dynamicFieldPageNumber: pageNumber).Count() * lineHeight;
        }
        else if (paragraph.Images.Count == 0)
        {
            height += lineHeight;
        }

        foreach (DocxInlineImage image in paragraph.Images)
        {
            double imageWidth = Math.Min(availableWidth, image.WidthPoints);
            double imageHeight = image.HeightPoints * imageWidth / Math.Max(1d, image.WidthPoints);
            height += imageHeight + InlineImageParagraphGapPoints;
        }

        return height;
    }

    // W6-a1: fixed horizontal text offsets join scaled space; fixedScale reuses the
    // layout spacing scale (identical in every mode).
    private static double GetParagraphTextStartOffset(DocxParagraph paragraph, double fixedScale)
    {
        if (paragraph.ListLabel is null)
        {
            return Math.Max(0d, paragraph.EffectiveProperties.Indent.LeftPoints ?? 0d) * fixedScale;
        }

        DocxNumberingIndent indent = ResolveListIndent(paragraph);
        double left = indent.LeftPoints ?? 0d;
        double firstLine = indent.FirstLinePoints ?? 0d;
        return Math.Max(0d, left + firstLine) * fixedScale;
    }

    private static double GetParagraphFirstLineTextStartOffset(DocxParagraph paragraph, double fontSize, IDocxTextMeasurer textMeasurer, double fixedScale)
    {
        if (paragraph.ListLabel is null)
        {
            return GetParagraphFirstLineIndentOffset(paragraph, fixedScale);
        }

        bool IsNumberingTabSuffix(DocxListLabel label)
        {
            return string.IsNullOrEmpty(label.SuffixValue) ||
                label.SuffixValue.Equals("tab", StringComparison.OrdinalIgnoreCase);
        }

        bool IsNumberingSpaceSuffix(DocxListLabel label)
        {
            return label.SuffixValue.Equals("space", StringComparison.OrdinalIgnoreCase);
        }

        if (IsNumberingTabSuffix(paragraph.ListLabel))
        {
            double textStart = GetParagraphTextStartOffset(paragraph, fixedScale);
            return Math.Max(textStart, GetNumberingTabPosition() ?? 0d);
        }

        double gap = IsNumberingSpaceSuffix(paragraph.ListLabel)
            ? textMeasurer.MeasureText(paragraph.Runs.FirstOrDefault(), " ", fontSize)
            : 0d;
        DocxTextRun labelRun = CreateListLabelRun(paragraph.ListLabel, paragraph.Runs.FirstOrDefault(), fontSize);
        return Math.Max(
            0d,
            GetParagraphLabelStartOffset(paragraph, fixedScale) +
                textMeasurer.MeasureText(labelRun, paragraph.ListLabel.Text, labelRun.EffectiveProperties.FontSize) +
                gap);

        double? GetNumberingTabPosition()
        {
            double? paragraphNumberingTab = paragraph.EffectiveProperties.TabStops
                .Where(tab => string.Equals(tab.Value, "num", StringComparison.OrdinalIgnoreCase))
                .Select(tab => tab.PositionPoints)
                .FirstOrDefault(position => position is not null);
            return (paragraphNumberingTab ?? paragraph.ListLabel?.Indent.NumberingTabPositionPoints) * fixedScale;
        }
    }

    private static double GetParagraphLabelStartOffset(DocxParagraph paragraph, double fixedScale)
    {
        if (paragraph.ListLabel is null)
        {
            return 0d;
        }

        DocxNumberingIndent indent = ResolveListIndent(paragraph);
        double left = indent.LeftPoints ?? 0d;
        double hanging = indent.HangingPoints ?? 0d;
        double firstLine = indent.FirstLinePoints ?? 0d;
        return Math.Max(0d, left - hanging + firstLine) * fixedScale;
    }

    private static DocxNumberingIndent ResolveListIndent(DocxParagraph paragraph)
    {
        DocxNumberingIndent listIndent = paragraph.ListLabel?.Indent ?? DocxNumberingIndent.Empty;
        DocxParagraphIndent paragraphIndent = paragraph.EffectiveProperties.Indent;
        if (!HasParagraphIndentOverride(paragraphIndent))
        {
            return listIndent;
        }

        bool hasParagraphFirstLineSide =
            paragraphIndent.FirstLinePoints is not null ||
            paragraphIndent.HangingPoints is not null ||
            paragraphIndent.FirstLineValue is not null ||
            paragraphIndent.HangingValue is not null;
        return new DocxNumberingIndent(
            paragraphIndent.LeftPoints ?? listIndent.LeftPoints,
            paragraphIndent.RightPoints ?? listIndent.RightPoints,
            hasParagraphFirstLineSide ? paragraphIndent.FirstLinePoints : listIndent.FirstLinePoints,
            hasParagraphFirstLineSide ? paragraphIndent.HangingPoints : listIndent.HangingPoints,
            listIndent.NumberingTabPositionPoints,
            paragraphIndent.LeftValue ?? listIndent.LeftValue,
            paragraphIndent.RightValue ?? listIndent.RightValue,
            hasParagraphFirstLineSide ? paragraphIndent.FirstLineValue : listIndent.FirstLineValue,
            hasParagraphFirstLineSide ? paragraphIndent.HangingValue : listIndent.HangingValue,
            listIndent.NumberingTabValue,
            listIndent.NumberingTabPositionValue);
    }

    private static bool HasParagraphIndentOverride(DocxParagraphIndent indent)
    {
        return indent.LeftPoints is not null ||
            indent.RightPoints is not null ||
            indent.FirstLinePoints is not null ||
            indent.HangingPoints is not null ||
            indent.LeftValue is not null ||
            indent.RightValue is not null ||
            indent.FirstLineValue is not null ||
            indent.HangingValue is not null;
    }

    private static double GetParagraphStartOffset(DocxParagraph paragraph, double fixedScale)
    {
        return paragraph.ListLabel is null
            ? GetParagraphFirstLineIndentOffset(paragraph, fixedScale)
            : GetParagraphLabelStartOffset(paragraph, fixedScale);
    }

    private static double GetParagraphRightInset(DocxParagraph paragraph, double fixedScale)
    {
        return (paragraph.ListLabel?.Indent.RightPoints ?? paragraph.EffectiveProperties.Indent.RightPoints ?? 0d) * fixedScale;
    }

    private static double GetParagraphFirstLineIndentOffset(DocxParagraph paragraph, double fixedScale)
    {
        DocxParagraphIndent indent = paragraph.EffectiveProperties.Indent;
        double left = indent.LeftPoints ?? 0d;
        double firstLine = indent.FirstLinePoints ?? 0d;
        double hanging = indent.HangingPoints ?? 0d;
        return Math.Max(0d, left + firstLine - hanging) * fixedScale;
    }

    private static DocxParagraphLineShape CreateParagraphLineShape(
        DocxParagraph paragraph,
        DocxWrappedTextLine line,
        DocxTextRun firstRun,
        bool firstLine,
        bool finalWrappedLine,
        double labelX,
        double lineX,
        double paragraphWidth,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints,
        int? dynamicFieldPageNumber)
    {
        double lineWidth = MeasureTextSpansForLayout(line.Spans, fontSize, textMeasurer, tabStops, defaultTabStopPoints, dynamicFieldPageNumber);
        double drawableLineWidth = MeasureDrawableTextSpansForLayout(line.Spans, fontSize, textMeasurer, tabStops, defaultTabStopPoints, dynamicFieldPageNumber);
        bool justifyLine = (paragraph.ListLabel is null || !firstLine) &&
            ShouldJustifyTextLine(paragraph.EffectiveProperties.Alignment, finalWrappedLine, drawableLineWidth, paragraphWidth, line.Spans);
        IReadOnlyList<DocxTextSegmentLayout> segments = firstLine && paragraph.ListLabel is not null
            ? CreateNumberedLineSegments(paragraph.ListLabel, line.Spans, firstRun)
            : justifyLine
                ? CreateJustifiedTextSegments(line.Spans, lineX, drawableLineWidth, paragraphWidth, fontSize, textMeasurer, tabStops, defaultTabStopPoints)
                : CreateTextSegments(line.Spans, lineX, fontSize, textMeasurer, tabStops, defaultTabStopPoints);
        double effectiveX = firstLine && paragraph.ListLabel is not null ? labelX : lineX;
        double effectiveWidth = firstLine && paragraph.ListLabel is not null
            ? Math.Max(lineX + lineWidth, labelX + MeasureListLabel(paragraph.ListLabel, firstRun, fontSize, textMeasurer)) - labelX
            : justifyLine
                ? paragraphWidth
                : lineWidth;
        string GetListLabelTextSeparator(DocxListLabel label)
        {
            return label.SuffixValue switch
            {
                "nothing" => string.Empty,
                "space" => " ",
                _ => "\t"
            };
        }

        string text = firstLine && paragraph.ListLabel is not null
            ? paragraph.ListLabel.Text + GetListLabelTextSeparator(paragraph.ListLabel) + line.Text
            : line.Text;
        return new DocxParagraphLineShape(text, effectiveX, effectiveWidth, segments);

        IReadOnlyList<DocxTextSegmentLayout> CreateNumberedLineSegments(DocxListLabel label, IReadOnlyList<DocxTextSpan> lineSpans, DocxTextRun styleRun)
        {
            DocxTextRun labelRun = CreateListLabelRun(label, styleRun, fontSize);
            double labelFontSize = labelRun.EffectiveProperties.FontSize;
            double labelWidth = textMeasurer.MeasureText(labelRun, label.Text, labelFontSize);
            DocxTextEmissionPlan labelPlan = DocxTextEmissionPlanner.CreateForListLabel(labelRun, label);
            var numberedSegments = new List<DocxTextSegmentLayout>
            {
                new(
                    label.Text,
                    labelRun,
                    labelX,
                    labelWidth,
                    labelFontSize,
                    PdfCharacterSpacing: labelPlan.PdfCharacterSpacing,
                    PdfCharacterSpacingSource: labelPlan.PdfCharacterSpacingSource,
                    CompensatePdfCharacterSpacing: labelPlan.CompensatePdfCharacterSpacing,
                    SourceTextRunIndex: -1,
                    BaselineOffsetY: 0d,
                    SourceTextOffsetInRun: 0,
                    Role: DocxTextSegmentRole.ListLabel)
            };

            string GetListLabelPdfSeparator(DocxListLabel label)
            {
                return label.SuffixValue.Equals("nothing", StringComparison.OrdinalIgnoreCase) ? string.Empty : " ";
            }

            string separator = GetListLabelPdfSeparator(label);
            if (separator.Length != 0)
            {
                DocxTextRun separatorRun = styleRun;
                double separatorFontSize = separatorRun.EffectiveProperties.FontSize;
                double separatorX = labelX + labelWidth;
                double separatorWidth = textMeasurer.MeasureText(separatorRun, separator, separatorFontSize);
                numberedSegments.Add(new DocxTextSegmentLayout(separator, separatorRun, separatorX, separatorWidth, separatorFontSize, SourceTextRunIndex: -1, BaselineOffsetY: 0d, PdfCharacterSpacing: 0d, PdfCharacterSpacingSource: DocxTextStateCharacterSpacingSource.None, CompensatePdfCharacterSpacing: true, SourceTextOffsetInRun: 0, Role: DocxTextSegmentRole.ListSeparator));
            }

            numberedSegments.AddRange(CreateTextSegments(lineSpans, lineX, fontSize, textMeasurer, tabStops, defaultTabStopPoints));
            return numberedSegments;
        }
    }

    private static double MeasureListLabel(DocxListLabel label, DocxTextRun? baseRun, double fontSize, IDocxTextMeasurer textMeasurer)
    {
        DocxTextRun labelRun = CreateListLabelRun(label, baseRun, fontSize);
        return textMeasurer.MeasureText(labelRun, label.Text, labelRun.EffectiveProperties.FontSize);
    }

    internal static DocxTextRun CreateListLabelRun(DocxListLabel label, DocxTextRun? baseRun, double fontSize)
    {
        return label.Style.ApplyTo(baseRun, label.Text, fontSize);
    }

}
