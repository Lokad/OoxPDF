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
                // Word also breaks an overlong token after a hyphen (or slash) when the prefix fits the remaining width; previously only line-leading overwide tokens used preferred breaks.
                double usedWidth = MeasureTextSpansForWrapping(SliceTextSpans(spans, lineStart, lineLength), fontSize, textMeasurer, tabStops, defaultTabStopPoints, preserveTerminalSoftHyphen: false, dynamicFieldPageNumber);
                double remainingWidth = maxWidth(lineIndex) - usedWidth;
                if (remainingWidth > 0d &&
                    TryFindPreferredTokenBreak(text, spans, token, remainingWidth, fontSize, textMeasurer, tabStops, defaultTabStopPoints, dynamicFieldPageNumber, out int preferredBreakLength))
                {
                    yield return CreateWrappedTextLine(text, spans, lineStart, lineLength + preferredBreakLength, endsWithIntraTokenBreak: true);
                    lineIndex++;
                    lineStart = token.Start + preferredBreakLength;
                    lineLength = 0;
                    tokens = ReplaceToken(tokens, tokenIndex, new TextToken(text.Substring(lineStart, token.Length - preferredBreakLength), lineStart, token.Length - preferredBreakLength));
                    tokenIndex--;
                    continue;
                }

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
                DocxTextBreakRules.IsNoBreakWhitespaceChar(before) ||
                DocxTextBreakRules.IsNoBreakWhitespaceChar(after))
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
        bool inBreakableWhitespace = DocxTextBreakRules.IsBreakableWhitespaceChar(text[0]);
        for (int i = 1; i < text.Length; i++)
        {
            bool breakableWhitespace = DocxTextBreakRules.IsBreakableWhitespaceChar(text[i]);
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

    private static class DocxLineBreakOpportunities
    {
        public static bool IsOpportunityAfter(char value)
        {
            return value is '-' or '/' or '\\' or '\u00AD' or '\u200B' or '\u2010' or '\u2012' or '\u2013' or '\u2014';
        }
    }

    private readonly record struct TextToken(string Text, int Start, int Length)
    {
        public bool IsBreakableWhitespace => Text.All(DocxTextBreakRules.IsBreakableWhitespaceChar);
    }
}
