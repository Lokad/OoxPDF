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
    private static IReadOnlyList<DocxTextSpan> CreateTextSpans(IReadOnlyList<DocxTextRun> runs, int? pageNumber, int? pageCount)
    {
        if (runs.Count != 0 && runs.All(run => run.Text.Length == 0 || run.EffectiveProperties.Hidden))
        {
            for (int i = 0; i < runs.Count; i++)
            {
                if (!runs[i].EffectiveProperties.Hidden)
                {
                    return [CreateTextSpan(" ", runs[i], i)];
                }
            }

            return [];
        }

        return runs
            .Select((run, index) => (run, index))
            .Where(item => item.run.Text.Length != 0 && !item.run.EffectiveProperties.Hidden)
            .Select(item => CreateTextSpan(ResolveLayoutFieldPlaceholders(item.run.Text, pageNumber, pageCount), item.run, item.index))
            .ToArray();
    }

    private static string ResolveLayoutFieldPlaceholders(string text, int? pageNumber, int? pageCount)
    {
        if (pageNumber is not null)
        {
            text = text.Replace("{PAGE}", pageNumber.Value.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal);
        }

        return pageCount is null
            ? text
            : text.Replace("{NUMPAGES}", pageCount.Value.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal);
    }

    private static DocxTextSpan CreateTextSpan(string text, DocxTextRun run, int fallbackSourceRunIndex)
    {
        return new DocxTextSpan(
            text,
            run,
            run.SourceRunIndex >= 0 ? run.SourceRunIndex : fallbackSourceRunIndex,
            Math.Max(0, run.SourceTextOffsetInRun));
    }

    private static double GetParagraphFontSize(DocxParagraph paragraph)
    {
        return paragraph.Runs.Count == 0 ? DocxDefaults.FontSizePoints : paragraph.Runs.Max(run => run.EffectiveProperties.FontSize);
    }

    private static IReadOnlyList<DocxTextSegmentLayout> CreateTextSegments(
        IReadOnlyList<DocxTextSpan> spans,
        double lineX,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints)
    {
        IReadOnlyList<DocxTabField> fields = SplitTabFields(spans);
        var segments = new List<DocxTextSegmentLayout>(spans.Count);
        double currentWidth = 0d;
        foreach (DocxTabField field in fields)
        {
            if (!field.StartsAfterTab)
            {
                currentWidth = AddTextFieldSegments(segments, field.Spans, lineX + currentWidth, fontSize, textMeasurer) - lineX;
                continue;
            }

            DocxResolvedTabStop tabStop = ResolveNextTabStop(currentWidth, tabStops, defaultTabStopPoints);
            double fieldWidth = MeasureTextFieldSpans(field.Spans, fontSize, textMeasurer);
            double? decimalAlignmentWidth = GetDecimalTabAlignmentWidth(tabStop, field.Spans, fontSize, textMeasurer);
            double fieldStart = GetAlignedTabFieldStart(tabStop, fieldWidth, decimalAlignmentWidth);
            AddTextFieldSegments(segments, field.Spans, lineX + fieldStart, fontSize, textMeasurer);
            currentWidth = GetAlignedTabFieldEnd(tabStop, fieldWidth, decimalAlignmentWidth);
        }

        return segments;

        double GetAlignedTabFieldStart(DocxResolvedTabStop tabStop, double fieldWidth, double? decimalAlignmentWidth)
        {
            if (string.Equals(tabStop.Value, "right", StringComparison.OrdinalIgnoreCase))
            {
                return tabStop.PositionPoints - fieldWidth;
            }

            if (decimalAlignmentWidth is not null)
            {
                return tabStop.PositionPoints - decimalAlignmentWidth.Value;
            }

            return string.Equals(tabStop.Value, "center", StringComparison.OrdinalIgnoreCase)
                ? tabStop.PositionPoints - (fieldWidth / 2d)
                : tabStop.PositionPoints;
        }
    }

    private static double MeasureTextSpans(
        IReadOnlyList<DocxTextSpan> spans,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints)
    {
        IReadOnlyList<DocxTabField> fields = SplitTabFields(spans);
        double currentWidth = 0d;
        foreach (DocxTabField field in fields)
        {
            double fieldWidth = MeasureTextFieldSpans(field.Spans, fontSize, textMeasurer);
            if (!field.StartsAfterTab)
            {
                currentWidth += fieldWidth;
                continue;
            }

            DocxResolvedTabStop tabStop = ResolveNextTabStop(currentWidth, tabStops, defaultTabStopPoints);
            double? decimalAlignmentWidth = GetDecimalTabAlignmentWidth(tabStop, field.Spans, fontSize, textMeasurer);
            currentWidth = GetAlignedTabFieldEnd(tabStop, fieldWidth, decimalAlignmentWidth);
        }

        return currentWidth;
    }

    private static double MeasureTextSpansForLayout(
        IReadOnlyList<DocxTextSpan> spans,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints,
        int? dynamicFieldPageNumber)
    {
        return MeasureTextSpans(
            NormalizeDynamicFieldMeasurementSpans(spans, dynamicFieldPageNumber),
            fontSize,
            textMeasurer,
            tabStops,
            defaultTabStopPoints);
    }

    private static double MeasureDrawableTextSpans(
        IReadOnlyList<DocxTextSpan> spans,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints)
    {
        int length = spans.Sum(span => span.Text.Length);
        int drawableLength = FindDrawableTextLength(spans);
        return drawableLength == length
            ? MeasureTextSpans(spans, fontSize, textMeasurer, tabStops, defaultTabStopPoints)
            : MeasureTextSpans(SliceTextSpans(spans, 0, drawableLength), fontSize, textMeasurer, tabStops, defaultTabStopPoints);
    }

    private static double MeasureDrawableTextSpansForLayout(
        IReadOnlyList<DocxTextSpan> spans,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints,
        int? dynamicFieldPageNumber)
    {
        int length = spans.Sum(span => span.Text.Length);
        int drawableLength = FindDrawableTextLength(spans);
        IReadOnlyList<DocxTextSpan> drawableSpans = drawableLength == length
            ? spans
            : SliceTextSpans(spans, 0, drawableLength);
        return MeasureTextSpansForLayout(drawableSpans, fontSize, textMeasurer, tabStops, defaultTabStopPoints, dynamicFieldPageNumber);
    }

    private static bool ShouldJustifyTextLine(
        DocxTextAlignment alignment,
        bool isLastLine,
        double drawableLineWidth,
        double paragraphWidth,
        IReadOnlyList<DocxTextSpan> spans)
    {
        return alignment == DocxTextAlignment.Justified &&
            !isLastLine &&
            paragraphWidth - drawableLineWidth > 0.001d &&
            CountStretchableJustificationSpaces(spans) > 0 &&
            !spans.Any(span => span.Text.IndexOf('\t') >= 0);
    }

    private static IReadOnlyList<DocxTextSegmentLayout> CreateJustifiedTextSegments(
        IReadOnlyList<DocxTextSpan> spans,
        double lineX,
        double drawableLineWidth,
        double paragraphWidth,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints)
    {
        int stretchableSpaces = CountStretchableJustificationSpaces(spans);
        if (stretchableSpaces == 0 || spans.Any(span => span.Text.IndexOf('\t') >= 0))
        {
            return CreateTextSegments(spans, lineX, fontSize, textMeasurer, tabStops, defaultTabStopPoints);
        }

        double extraPerSpace = Math.Max(0d, paragraphWidth - drawableLineWidth) / stretchableSpaces;
        int drawableLength = FindDrawableTextLength(spans);
        IReadOnlyList<DocxTextSpan> drawableSpans = SliceTextSpans(spans, 0, drawableLength);
        var segments = new List<DocxTextSegmentLayout>(drawableSpans.Count);
        double segmentX = lineX;
        for (int i = 0; i < drawableSpans.Count; i++)
        {
            DocxTextSpan span = drawableSpans[i];
            segmentX = AddJustifiedSpanSegments(span, segmentX);
            if (i + 1 < drawableSpans.Count)
            {
                segmentX += DocxTextSpacing.BoundarySpacing(span.StyleRun, span.Text, drawableSpans[i + 1].Text);
            }
        }

        return segments;

        double AddJustifiedSpanSegments(DocxTextSpan span, double segmentX)
        {
            double spanFontSize = GetTextSpanFontSize(span, fontSize);
            double baselineOffset = GetTextSpanBaselineOffset(span, fontSize);
            int start = 0;
            while (start < span.Text.Length)
            {
                bool isSpace = IsJustificationSpace(span.Text[start]);
                int end = start + 1;
                while (end < span.Text.Length && IsJustificationSpace(span.Text[end]) == isSpace)
                {
                    end++;
                }

                string text = span.Text[start..end];
                double width = textMeasurer.MeasureText(span.StyleRun, text, spanFontSize);
                if (isSpace)
                {
                    segmentX += width + text.Length * extraPerSpace;
                }
                else
                {
                    segments.Add(new DocxTextSegmentLayout(
                        text,
                        span.StyleRun,
                        segmentX,
                        width,
                        spanFontSize,
                        baselineOffset,
                        SourceTextRunIndex: span.SourceTextRunIndex,
                        SourceTextOffsetInRun: span.SourceTextOffsetInRun + start, PdfCharacterSpacing: 0d, PdfCharacterSpacingSource: DocxTextStateCharacterSpacingSource.None, CompensatePdfCharacterSpacing: true, Role: DocxTextSegmentRole.Text));
                    segmentX += width;
                }

                start = end;
            }

            return segmentX;
        }
    }

    private static int CountStretchableJustificationSpaces(IReadOnlyList<DocxTextSpan> spans)
    {
        int drawableLength = FindDrawableTextLength(spans);
        int seen = 0;
        int count = 0;
        foreach (DocxTextSpan span in spans)
        {
            foreach (char c in span.Text)
            {
                if (seen++ >= drawableLength)
                {
                    return count;
                }

                if (IsJustificationSpace(c))
                {
                    count++;
                }
            }
        }

        return count;
    }

    private static int FindDrawableTextLength(IReadOnlyList<DocxTextSpan> spans)
    {
        int length = spans.Sum(span => span.Text.Length);
        int index = length;
        for (int spanIndex = spans.Count - 1; spanIndex >= 0; spanIndex--)
        {
            string text = spans[spanIndex].Text;
            for (int i = text.Length - 1; i >= 0; i--)
            {
                index--;
                if (!IsBreakableWhitespaceChar(text[i]))
                {
                    return index + 1;
                }
            }
        }

        return 0;
    }

    private static bool IsJustificationSpace(char c)
    {
        return c == ' ';
    }

    private sealed class DocxTabField
    {
        public DocxTabField(bool startsAfterTab)
        {
            StartsAfterTab = startsAfterTab;
        }

        public bool StartsAfterTab { get; }

        public List<DocxTextSpan> Spans { get; } = [];
    }

    private readonly record struct DocxResolvedTabStop(double PositionPoints, string? Value);

    private static IReadOnlyList<DocxTabField> SplitTabFields(IReadOnlyList<DocxTextSpan> spans)
    {
        var fields = new List<DocxTabField> { new(startsAfterTab: false) };
        foreach (DocxTextSpan span in spans)
        {
            int start = 0;
            for (int i = 0; i < span.Text.Length; i++)
            {
                if (span.Text[i] != '\t')
                {
                    continue;
                }

                AddTabFieldSpan(fields[^1], span, start, i);
                fields.Add(new DocxTabField(startsAfterTab: true));
                start = i + 1;
            }

            AddTabFieldSpan(fields[^1], span, start, span.Text.Length);
        }

        return fields;
    }

    private static void AddTabFieldSpan(DocxTabField field, DocxTextSpan source, int start, int end)
    {
        if (end <= start)
        {
            return;
        }

        field.Spans.Add(new DocxTextSpan(
            source.Text[start..end],
            source.StyleRun,
            source.SourceTextRunIndex,
            source.SourceTextOffsetInRun + start));
    }

    private static double AddTextFieldSegments(
        List<DocxTextSegmentLayout> segments,
        IReadOnlyList<DocxTextSpan> spans,
        double segmentX,
        double fontSize,
        IDocxTextMeasurer textMeasurer)
    {
        for (int i = 0; i < spans.Count; i++)
        {
            DocxTextSpan span = spans[i];
            double spanFontSize = GetTextSpanFontSize(span, fontSize);
            double baselineOffset = GetTextSpanBaselineOffset(span, fontSize);
            segmentX = AddTextSegment(
                span.StyleRun,
                span.SourceTextRunIndex,
                span.SourceTextOffsetInRun,
                span.Text,
                segmentX,
                spanFontSize,
                baselineOffset);
            if (i + 1 < spans.Count)
            {
                segmentX += DocxTextSpacing.BoundarySpacing(span.StyleRun, span.Text, spans[i + 1].Text);
            }
        }

        return segmentX;

        double AddTextSegment(DocxTextRun styleRun, int sourceTextRunIndex, int sourceTextOffsetInRun, string text, double segmentX, double fontSize, double baselineOffset)
        {
            int leadingSpaces = CountLeadingOfficeSeparatedSpaces();
            if (leadingSpaces == 0 || leadingSpaces == text.Length)
            {
                double width = textMeasurer.MeasureText(styleRun, text, fontSize);
                segments.Add(new DocxTextSegmentLayout(
                    text,
                    styleRun,
                    segmentX,
                    width,
                    fontSize,
                    baselineOffset,
                    SourceTextRunIndex: sourceTextRunIndex,
                    SourceTextOffsetInRun: sourceTextOffsetInRun, PdfCharacterSpacing: 0d, PdfCharacterSpacingSource: DocxTextStateCharacterSpacingSource.None, CompensatePdfCharacterSpacing: true, Role: DocxTextSegmentRole.Text));
                return segmentX + width;
            }

            string spaceText = text[..leadingSpaces];
            double spaceWidth = textMeasurer.MeasureText(styleRun, spaceText, fontSize);
            segments.Add(new DocxTextSegmentLayout(
                spaceText,
                styleRun,
                segmentX,
                spaceWidth,
                fontSize,
                baselineOffset,
                SourceTextRunIndex: sourceTextRunIndex,
                SourceTextOffsetInRun: sourceTextOffsetInRun, PdfCharacterSpacing: 0d, PdfCharacterSpacingSource: DocxTextStateCharacterSpacingSource.None, CompensatePdfCharacterSpacing: true, Role: DocxTextSegmentRole.Text));
            segmentX += spaceWidth + DocxTextSpacing.BoundarySpacing(styleRun, spaceText, text[leadingSpaces..]);

            string bodyText = text[leadingSpaces..];
            double bodyWidth = textMeasurer.MeasureText(styleRun, bodyText, fontSize);
            segments.Add(new DocxTextSegmentLayout(
                bodyText,
                styleRun,
                segmentX,
                bodyWidth,
                fontSize,
                baselineOffset,
                SourceTextRunIndex: sourceTextRunIndex,
                SourceTextOffsetInRun: sourceTextOffsetInRun + leadingSpaces, PdfCharacterSpacing: 0d, PdfCharacterSpacingSource: DocxTextStateCharacterSpacingSource.None, CompensatePdfCharacterSpacing: true, Role: DocxTextSegmentRole.Text));
            return segmentX + bodyWidth;

            int CountLeadingOfficeSeparatedSpaces()
            {
                int count = 0;
                while (count < text.Length && text[count] == ' ')
                {
                    count++;
                }

                return count;
            }
        }
    }

    private static double MeasureTextFieldSpans(
        IReadOnlyList<DocxTextSpan> spans,
        double fontSize,
        IDocxTextMeasurer textMeasurer)
    {
        double width = 0d;
        for (int i = 0; i < spans.Count; i++)
        {
            DocxTextSpan span = spans[i];
            double spanFontSize = GetTextSpanFontSize(span, fontSize);
            width += textMeasurer.MeasureText(span.StyleRun, span.Text, spanFontSize);
            if (i + 1 < spans.Count)
            {
                width += DocxTextSpacing.BoundarySpacing(span.StyleRun, span.Text, spans[i + 1].Text);
            }
        }

        return width;
    }

    private static double? GetDecimalTabAlignmentWidth(
        DocxResolvedTabStop tabStop,
        IReadOnlyList<DocxTextSpan> spans,
        double fontSize,
        IDocxTextMeasurer textMeasurer)
    {
        if (!string.Equals(tabStop.Value, "decimal", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        double width = 0d;
        for (int i = 0; i < spans.Count; i++)
        {
            DocxTextSpan span = spans[i];
            int decimalIndex = IndexOfDecimalSeparator(span.Text);
            double spanFontSize = GetTextSpanFontSize(span, fontSize);
            if (decimalIndex >= 0)
            {
                width += textMeasurer.MeasureText(span.StyleRun, span.Text[..decimalIndex], spanFontSize);
                return width;
            }

            width += textMeasurer.MeasureText(span.StyleRun, span.Text, spanFontSize);
            if (i + 1 < spans.Count)
            {
                width += DocxTextSpacing.BoundarySpacing(span.StyleRun, span.Text, spans[i + 1].Text);
            }
        }

        return width;

        int IndexOfDecimalSeparator(string text)
        {
            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] is '.' or ',')
                {
                    return i;
                }
            }

            return -1;
        }
    }

    private static double GetAlignedTabFieldEnd(DocxResolvedTabStop tabStop, double fieldWidth, double? decimalAlignmentWidth)
    {
        if (string.Equals(tabStop.Value, "right", StringComparison.OrdinalIgnoreCase))
        {
            return tabStop.PositionPoints;
        }

        if (decimalAlignmentWidth is not null)
        {
            return tabStop.PositionPoints - decimalAlignmentWidth.Value + fieldWidth;
        }

        return string.Equals(tabStop.Value, "center", StringComparison.OrdinalIgnoreCase)
            ? tabStop.PositionPoints + (fieldWidth / 2d)
            : tabStop.PositionPoints + fieldWidth;
    }

    private static double GetTextSpanFontSize(DocxTextSpan span, double fallbackFontSize)
    {
        double fontSize = span.StyleRun.EffectiveProperties.FontSize;
        double nominalFontSize = fontSize > 0d ? fontSize : fallbackFontSize;
        return DocxVerticalAlignMetrics.ResolveFontSize(nominalFontSize, span.StyleRun);
    }

    private static double GetTextSpanBaselineOffset(DocxTextSpan span, double fallbackFontSize)
    {
        double fontSize = span.StyleRun.EffectiveProperties.FontSize;
        double nominalFontSize = fontSize > 0d ? fontSize : fallbackFontSize;
        double layoutFontSize = DocxVerticalAlignMetrics.ResolveFontSize(nominalFontSize, span.StyleRun);
        return DocxVerticalAlignMetrics.ResolveBaselineOffset(nominalFontSize, layoutFontSize, span.StyleRun);
    }

    private static DocxResolvedTabStop ResolveNextTabStop(double width, IReadOnlyList<DocxTabStop> tabStops, double defaultTabStopPoints)
    {
        foreach (DocxTabStop tabStop in tabStops
            .Where(tabStop => tabStop.PositionPoints is not null && IsPositioningTabStop(tabStop))
            .OrderBy(tabStop => tabStop.PositionPoints ?? 0d))
        {
            if ((tabStop.PositionPoints ?? 0d) > width + 0.001d)
            {
                return new DocxResolvedTabStop(tabStop.PositionPoints ?? 0d, tabStop.Value);
            }
        }

        return new DocxResolvedTabStop(AdvanceToNextDefaultTabStop(), null);

        double AdvanceToNextDefaultTabStop()
        {
            double tabStop = defaultTabStopPoints > 0d ? defaultTabStopPoints : WordDefaultTabStopPoints;
            return (Math.Floor(width / tabStop) + 1d) * tabStop;
        }

        bool IsPositioningTabStop(DocxTabStop tabStop)
        {
            return !string.Equals(tabStop.Value, "bar", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(tabStop.Value, "clear", StringComparison.OrdinalIgnoreCase);
        }
    }

    private static IEnumerable<DocxWrappedTextLine> WrapTextLines(
        IReadOnlyList<DocxTextSpan> spans,
        double firstLineMaxWidth,
        double continuationLineMaxWidth,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints,
        bool allowOverwideTokenBreaks,
        int? dynamicFieldPageNumber)
    {
        string text = string.Concat(spans.Select(span => span.Text));
        int lineIndex = 0;
        int segmentStart = 0;
        while (segmentStart <= text.Length)
        {
            int breakIndex = text.IndexOf('\n', segmentStart);
            int segmentLength = breakIndex < 0 ? text.Length - segmentStart : breakIndex - segmentStart;
            bool yielded = false;
            foreach (DocxWrappedTextLine line in WrapWords(text, spans, segmentStart, segmentLength, index => index == 0 && lineIndex == 0 ? firstLineMaxWidth : continuationLineMaxWidth, fontSize, textMeasurer, tabStops, defaultTabStopPoints, allowOverwideTokenBreaks, dynamicFieldPageNumber))
            {
                yielded = true;
                yield return line;
                lineIndex++;
            }

            if (!yielded && segmentLength == 0)
            {
                yield return new DocxWrappedTextLine(string.Empty, [], false);
                lineIndex++;
            }

            if (breakIndex < 0)
            {
                yield break;
            }

            segmentStart = breakIndex + 1;
        }
    }

    private static bool ShouldAllowCharacterLevelWordWrap(DocxParagraph paragraph)
    {
        return paragraph.EffectiveProperties.WordWrap == false;
    }

    private static IEnumerable<DocxWrappedTextLine> WrapWords(
        string text,
        IReadOnlyList<DocxTextSpan> spans,
        int segmentStart,
        int segmentLength,
        Func<int, double> maxWidth,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints,
        bool allowOverwideTokenBreaks,
        int? dynamicFieldPageNumber)
    {
        IReadOnlyList<TextToken> tokens = TokenizeSpaces(text, segmentStart, segmentLength);
        if (tokens.Count == 0)
        {
            yield break;
        }

        int lineStart = tokens[0].Start;
        int lineLength = 0;
        int lineIndex = 0;
        for (int tokenIndex = 0; tokenIndex < tokens.Count; tokenIndex++)
        {
            TextToken token = tokens[tokenIndex];
            int candidateLength = token.Start + token.Length - lineStart;
            bool lineHasNonWhitespace = HasNonWhitespace(text, lineStart, lineLength);
            if (lineLength > 0 &&
                lineHasNonWhitespace &&
                !token.IsBreakableWhitespace &&
                MeasureTextSpansForWrapping(SliceTextSpans(spans, lineStart, candidateLength), fontSize, textMeasurer, tabStops, defaultTabStopPoints, preserveTerminalSoftHyphen: false, dynamicFieldPageNumber) > maxWidth(lineIndex))
            {
                yield return CreateWrappedTextLine(text, spans, lineStart, lineLength, false);
                lineIndex++;
                lineStart = token.Start;
                lineLength = 0;
                tokenIndex--;
                continue;
            }

            if (lineLength == 0 &&
                !token.IsBreakableWhitespace &&
                TryFindPreferredTokenBreak(text, spans, token, maxWidth(lineIndex), fontSize, textMeasurer, tabStops, defaultTabStopPoints, dynamicFieldPageNumber, out int breakLength))
            {
                yield return CreateWrappedTextLine(text, spans, token.Start, breakLength, endsWithIntraTokenBreak: true);
                lineIndex++;
                lineStart = token.Start + breakLength;
                lineLength = 0;
                tokens = ReplaceToken(tokens, tokenIndex, new TextToken(text.Substring(lineStart, token.Length - breakLength), lineStart, token.Length - breakLength));
                tokenIndex--;
            }
            else if (lineLength == 0 &&
                allowOverwideTokenBreaks &&
                !token.IsBreakableWhitespace &&
                TryFindOverwideTokenBreak(text, spans, token, maxWidth(lineIndex), fontSize, textMeasurer, tabStops, defaultTabStopPoints, dynamicFieldPageNumber, out breakLength))
            {
                yield return CreateWrappedTextLine(text, spans, token.Start, breakLength, endsWithIntraTokenBreak: true);
                lineIndex++;
                lineStart = token.Start + breakLength;
                lineLength = 0;
                tokens = ReplaceToken(tokens, tokenIndex, new TextToken(text.Substring(lineStart, token.Length - breakLength), lineStart, token.Length - breakLength));
                tokenIndex--;
            }
            else
            {
                lineLength = candidateLength;
            }
        }

        if (lineLength > 0)
        {
            yield return CreateWrappedTextLine(text, spans, lineStart, lineLength, false);
        }
    }

    private static bool TryFindOverwideTokenBreak(
        string text,
        IReadOnlyList<DocxTextSpan> spans,
        TextToken token,
        double maxWidth,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints,
        int? dynamicFieldPageNumber,
        out int breakLength)
    {
        if (TryFindPreferredTokenBreak(text, spans, token, maxWidth, fontSize, textMeasurer, tabStops, defaultTabStopPoints, dynamicFieldPageNumber, out breakLength))
        {
            return true;
        }

        double tokenWidth = MeasureTextSpansForWrapping(SliceTextSpans(spans, token.Start, token.Length), fontSize, textMeasurer, tabStops, defaultTabStopPoints, preserveTerminalSoftHyphen: false, dynamicFieldPageNumber);
        if (tokenWidth <= maxWidth)
        {
            breakLength = 0;
            return false;
        }

        for (int length = token.Length - 1; length > 0; length--)
        {
            if (!IsSafeEmergencyTokenBreak(length))
            {
                continue;
            }

            bool preserveTerminalSoftHyphen = text[token.Start + length - 1] == '\u00AD';
            double prefixWidth = MeasureTextSpansForWrapping(SliceTextSpans(spans, token.Start, length), fontSize, textMeasurer, tabStops, defaultTabStopPoints, preserveTerminalSoftHyphen, dynamicFieldPageNumber);
            if (prefixWidth <= maxWidth)
            {
                breakLength = length;
                return true;
            }
        }

        return false;

        bool IsSafeEmergencyTokenBreak(int length)
        {
            if (length <= 0 || length >= token.Length)
            {
                return false;
            }

            int breakIndex = token.Start + length;
            char before = text[breakIndex - 1];
            char after = text[breakIndex];
            if (char.IsHighSurrogate(before) && char.IsLowSurrogate(after))
            {
                return false;
            }

            if (before == '\u2011' ||
                after == '\u2011' ||
                IsNoBreakWhitespaceChar(before) ||
                IsNoBreakWhitespaceChar(after))
            {
                return false;
            }

            UnicodeCategory afterCategory = char.GetUnicodeCategory(after);
            return afterCategory is not UnicodeCategory.NonSpacingMark and
                not UnicodeCategory.SpacingCombiningMark and
                not UnicodeCategory.EnclosingMark;
        }
    }

    private static bool TryFindPreferredTokenBreak(
        string text,
        IReadOnlyList<DocxTextSpan> spans,
        TextToken token,
        double maxWidth,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints,
        int? dynamicFieldPageNumber,
        out int breakLength)
    {
        breakLength = 0;
        double tokenWidth = MeasureTextSpansForWrapping(SliceTextSpans(spans, token.Start, token.Length), fontSize, textMeasurer, tabStops, defaultTabStopPoints, preserveTerminalSoftHyphen: false, dynamicFieldPageNumber);
        if (tokenWidth <= maxWidth)
        {
            return false;
        }

        for (int length = token.Length - 1; length > 0; length--)
        {
            int absoluteIndex = token.Start + length - 1;
            if (!DocxLineBreakOpportunities.IsOpportunityAfter(text[absoluteIndex]))
            {
                continue;
            }

            bool preserveTerminalSoftHyphen = text[token.Start + length - 1] == '\u00AD';
            double prefixWidth = MeasureTextSpansForWrapping(SliceTextSpans(spans, token.Start, length), fontSize, textMeasurer, tabStops, defaultTabStopPoints, preserveTerminalSoftHyphen, dynamicFieldPageNumber);
            if (prefixWidth <= maxWidth)
            {
                breakLength = length;
                return true;
            }
        }

        return false;
    }

    private static IReadOnlyList<TextToken> ReplaceToken(IReadOnlyList<TextToken> tokens, int index, TextToken replacement)
    {
        var result = new List<TextToken>(tokens.Count);
        for (int i = 0; i < tokens.Count; i++)
        {
            result.Add(i == index ? replacement : tokens[i]);
        }

        return result;
    }

    private static DocxWrappedTextLine CreateWrappedTextLine(
        string text,
        IReadOnlyList<DocxTextSpan> spans,
        int start,
        int length,
        bool endsWithIntraTokenBreak)
    {
        string lineText = text.Substring(start, length);
        bool preserveTerminalSoftHyphen = endsWithIntraTokenBreak && lineText.EndsWith('\u00AD');
        return new DocxWrappedTextLine(
            RemoveHiddenBreakCharacters(lineText, preserveTerminalSoftHyphen),
            NormalizeHiddenBreakSpans(SliceTextSpans(spans, start, length), preserveTerminalSoftHyphen),
            endsWithIntraTokenBreak);
    }

    private static IReadOnlyList<DocxTextSpan> SliceTextSpans(
        IReadOnlyList<DocxTextSpan> spans,
        int start,
        int length)
    {
        if (length == 0)
        {
            return [];
        }

        var sliced = new List<DocxTextSpan>();
        int spanStart = 0;
        int end = start + length;
        foreach (DocxTextSpan span in spans)
        {
            int spanEnd = spanStart + span.Text.Length;
            int sliceStart = Math.Max(start, spanStart);
            int sliceEnd = Math.Min(end, spanEnd);
            if (sliceStart < sliceEnd)
            {
                int spanOffset = sliceStart - spanStart;
                sliced.Add(new DocxTextSpan(
                    span.Text[spanOffset..(sliceEnd - spanStart)],
                    span.StyleRun,
                    span.SourceTextRunIndex,
                    span.SourceTextOffsetInRun + spanOffset));
            }

            if (spanEnd >= end)
            {
                break;
            }

            spanStart = spanEnd;
        }

        return sliced;
    }

    private static double MeasureTextSpansForWrapping(
        IReadOnlyList<DocxTextSpan> spans,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints,
        bool preserveTerminalSoftHyphen,
        int? dynamicFieldPageNumber)
    {
        return MeasureTextSpans(
            NormalizeDynamicFieldMeasurementSpans(NormalizeHiddenBreakSpans(spans, preserveTerminalSoftHyphen), dynamicFieldPageNumber),
            fontSize,
            textMeasurer,
            tabStops,
            defaultTabStopPoints);
    }

    private static IReadOnlyList<DocxTextSpan> NormalizeDynamicFieldMeasurementSpans(
        IReadOnlyList<DocxTextSpan> spans,
        int? pageNumber)
    {
        if (pageNumber is null ||
            !spans.Any(span => span.Text.Contains("{NUMPAGES}", StringComparison.Ordinal)))
        {
            return spans;
        }

        string proxy = pageNumber.Value.ToString(CultureInfo.InvariantCulture);
        return spans
            .Select(span => span.Text.Contains("{NUMPAGES}", StringComparison.Ordinal)
                ? span with { Text = span.Text.Replace("{NUMPAGES}", proxy, StringComparison.Ordinal) }
                : span)
            .ToArray();
    }

    private static IReadOnlyList<DocxTextSpan> NormalizeHiddenBreakSpans(
        IReadOnlyList<DocxTextSpan> spans,
        bool preserveTerminalSoftHyphen)
    {
        if (!spans.Any(span => span.Text.IndexOf('\u00AD') >= 0 || span.Text.IndexOf('\u200B') >= 0))
        {
            return spans;
        }

        var normalized = new List<DocxTextSpan>(spans.Count);
        for (int index = 0; index < spans.Count; index++)
        {
            DocxTextSpan span = spans[index];
            AddHiddenBreakNormalizedSpans(span, preserveTerminalSoftHyphen && index == spans.Count - 1 && span.Text.EndsWith('\u00AD'));
        }

        return normalized;

        void AddHiddenBreakNormalizedSpans(DocxTextSpan span, bool preserveTerminalSoftHyphen)
        {
            int chunkStart = 0;
            for (int index = 0; index < span.Text.Length; index++)
            {
                if (!IsHiddenBreakCharacter(span.Text[index], preserveTerminalSoftHyphen && index == span.Text.Length - 1))
                {
                    continue;
                }

                AddHiddenBreakNormalizedSpan(normalized, span, chunkStart, index);
                chunkStart = index + 1;
            }

            AddHiddenBreakNormalizedSpan(normalized, span, chunkStart, span.Text.Length);
        }
    }

    private static void AddHiddenBreakNormalizedSpan(
        List<DocxTextSpan> normalized,
        DocxTextSpan span,
        int start,
        int end)
    {
        if (end <= start)
        {
            return;
        }

        normalized.Add(span with
        {
            Text = span.Text[start..end],
            SourceTextOffsetInRun = span.SourceTextOffsetInRun + start
        });
    }

    private static string RemoveHiddenBreakCharacters(string text, bool preserveTerminalSoftHyphen)
    {
        if (text.IndexOf('\u00AD') < 0 && text.IndexOf('\u200B') < 0)
        {
            return text;
        }

        var builder = new StringBuilder(text.Length);
        for (int index = 0; index < text.Length; index++)
        {
            char value = text[index];
            if (IsHiddenBreakCharacter(value, preserveTerminalSoftHyphen && index == text.Length - 1))
            {
                continue;
            }

            builder.Append(value);
        }

        return builder.ToString();
    }

    private static bool IsHiddenBreakCharacter(char value, bool preserveTerminalSoftHyphen)
    {
        return value == '\u200B' ||
            (value == '\u00AD' && !preserveTerminalSoftHyphen);
    }

    private static IReadOnlyList<TextToken> TokenizeSpaces(string text, int start, int length)
    {
        if (length == 0)
        {
            return [];
        }

        string segment = text.Substring(start, length);
        return TokenizeSpaces(segment)
            .Select(token => new TextToken(token.Text, start + token.Start, token.Length))
            .ToArray();
    }

    private static bool HasNonWhitespace(string text, int start, int length)
    {
        for (int i = start; i < start + length; i++)
        {
            if (!char.IsWhiteSpace(text[i]))
            {
                return true;
            }
        }

        return false;
    }

    private static IReadOnlyList<TextToken> TokenizeSpaces(string text)
    {
        if (text.Length == 0)
        {
            return [];
        }

        var tokens = new List<TextToken>();
        int start = 0;
        bool inBreakableWhitespace = IsBreakableWhitespaceChar(text[0]);
        for (int i = 1; i < text.Length; i++)
        {
            bool breakableWhitespace = IsBreakableWhitespaceChar(text[i]);
            if (breakableWhitespace == inBreakableWhitespace)
            {
                continue;
            }

            tokens.Add(new TextToken(text[start..i], start, i - start));
            start = i;
            inBreakableWhitespace = breakableWhitespace;
        }

        tokens.Add(new TextToken(text[start..], start, text.Length - start));
        return tokens;
    }

    private static bool IsBreakableWhitespaceChar(char value)
    {
        return char.IsWhiteSpace(value) &&
            !IsNoBreakWhitespaceChar(value);
    }

    private static bool IsNoBreakWhitespaceChar(char value)
    {
        return value is '\u00A0' or '\u202F' or '\u2007';
    }

    private static class DocxLineBreakOpportunities
    {
        public static bool IsOpportunityAfter(char value)
        {
            return value is '-' or '/' or '\\' or '\u00AD' or '\u200B' or '\u2010' or '\u2012' or '\u2013' or '\u2014';
        }
    }

    private readonly record struct TextToken(string Text, int Start, int Length)
    {
        public bool IsBreakableWhitespace => Text.All(IsBreakableWhitespaceChar);
    }
}
