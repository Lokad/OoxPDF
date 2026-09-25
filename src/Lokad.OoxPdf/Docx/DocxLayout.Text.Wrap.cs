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
    // R07.2: finite work limit for break-search scans. Grow stops at the first
    // overflow by construction but could still walk far on adversary-sized capacities,
    // so it emits the fitting candidate when capped. Shrink scans walk backward for a
    // fitting safe prefix and could traverse the whole token on adversarial non-monotonic
    // widths; past the cap the conversion fails defined (observable limit message)
    // instead of searching unbounded or silently changing breaks.
    private const int MaxSearchProbesPerLine = 128;

    // RV05: image-driven breaking. Character offset of an affined image in the concatenated
    // paragraph text, shared by wrapping and mid-line placement.
    private static int ResolveInlineImageCharOffset(
        DocxParagraph paragraph,
        IReadOnlyList<DocxTextSpan> textSpans,
        DocxInlineImage image)
    {
        int position = 0;
        foreach (DocxTextSpan span in textSpans)
        {
            if (span.SourceTextRunIndex >= 0 && span.SourceTextRunIndex < paragraph.Runs.Count && paragraph.Runs[span.SourceTextRunIndex].SourceRunIndex < image.SourceRunIndex)
            {
                position += span.Text.Length;
            }
        }

        return position;
    }

    // RV05: image-driven breaking. Affined image widths by character offset, stably sorted
    // (OrderBy is stable, so same-offset images keep document order). Callers pass the list
    // into wrapping; the range index requires sorted offsets.
    private static List<(int CharOffset, double Width)> ResolveInlineImageWrapWidths(
        DocxParagraph paragraph,
        IReadOnlyList<DocxTextSpan> textSpans)
    {
        var widths = new List<(int CharOffset, double Width)>();
        for (int i = 0; i < paragraph.Images.Count; i++)
        {
            if (paragraph.Images[i].SourceRunIndex >= 0)
            {
                widths.Add((ResolveInlineImageCharOffset(paragraph, textSpans, paragraph.Images[i]), paragraph.Images[i].WidthPoints));
            }
        }

        return widths.OrderBy(entry => entry.CharOffset).ToList();
    }

    // RV05: image-driven breaking. Sorted offsets with prefix sums backing range width
    // lookups shared by the body and static wrappers.
    private static (int[] Offsets, double[] PrefixSums) BuildInlineImageWrapIndex(
        IReadOnlyList<(int CharOffset, double Width)>? inlineImageWidths)
    {
        if (inlineImageWidths is null || inlineImageWidths.Count == 0)
        {
            return ([], [0d]);
        }

        var offsets = new int[inlineImageWidths.Count];
        var prefixSums = new double[inlineImageWidths.Count + 1];
        for (int i = 0; i < inlineImageWidths.Count; i++)
        {
            offsets[i] = inlineImageWidths[i].CharOffset;
            prefixSums[i + 1] = prefixSums[i] + inlineImageWidths[i].Width;
        }

        return (offsets, prefixSums);
    }

    private static double ImageWidthInRange((int[] Offsets, double[] PrefixSums) index, int rangeStart, int rangeEnd)
    {
        if (index.Offsets.Length == 0 || rangeEnd <= rangeStart)
        {
            return 0d;
        }

        int lo = Array.BinarySearch(index.Offsets, rangeStart);
        if (lo < 0)
        {
            lo = ~lo;
        }

        int hi = Array.BinarySearch(index.Offsets, rangeEnd);
        if (hi < 0)
        {
            hi = ~hi;
        }

        return index.PrefixSums[hi] - index.PrefixSums[lo];
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
        int? dynamicFieldPageNumber,
        CancellationToken cancellationToken = default,
        IReadOnlyList<(int CharOffset, double Width)>? inlineImageWidths = null)
    {
        string text = string.Concat(spans.Select(span => span.Text));
        int lineIndex = 0;
        int segmentStart = 0;
        while (segmentStart <= text.Length)
        {
            int breakIndex = text.IndexOf('\n', segmentStart);
            int segmentLength = breakIndex < 0 ? text.Length - segmentStart : breakIndex - segmentStart;
            bool yielded = false;
            foreach (DocxWrappedTextLine line in WrapWords(text, spans, segmentStart, segmentLength, index => index == 0 && lineIndex == 0 ? firstLineMaxWidth : continuationLineMaxWidth, fontSize, textMeasurer, tabStops, defaultTabStopPoints, allowOverwideTokenBreaks, dynamicFieldPageNumber, cancellationToken, inlineImageWidths))
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
        int? dynamicFieldPageNumber,
        CancellationToken cancellationToken,
        IReadOnlyList<(int CharOffset, double Width)>? inlineImageWidths = null)
    {
        // slice widths are memoized by text coordinates for this segment so
        // repeated measures of the same slice (whole-token checks, preferred/overwide
        // overlap, re-examined tokens) shape once. A slice value depends only on its
        // content here: every measure starts at width zero, and fontSize, tab stops,
        // and field state are fixed for the call. Binary search is deliberately NOT
        // used: prefix widths are not provably monotonic (kerning, negative tracking,
        // dynamic fields), and emergency tokens structurally contain no tabs (they
        // split tokens), which still does not certify shaping monotonicity.
        var measureMemo = new Dictionary<(int Start, int Length, bool PreserveTerminalSoftHyphen), double>();
        // RV13: per-token-chain average advance keyed by token end. Remainders
        // produced while wrapping one chain share its end, so the first full
        // measure of a chain lets later remainders skip their full measure.
        var chainAverages = new Dictionary<int, double>();
        // R07.2: index span start offsets once per segment so slice lookups seek
        // instead of restarting traversal from the first span on every measure.
        int[] spanStarts = new int[spans.Count + 1];
        for (int spanIndex = 0; spanIndex < spans.Count; spanIndex++)
        {
            spanStarts[spanIndex + 1] = spanStarts[spanIndex] + spans[spanIndex].Text.Length;
        }
        // per-measure normalization scans (soft-hyphen and NUMPAGES checks over
        // every slice) cost O(slice) each with zero reuse. Hoist the absence checks to the
        // segment once: slices inside a clean segment skip them exactly (both normalizers
        // return their input unchanged when there is nothing to rewrite).
        bool segmentHasHiddenBreaks = text.AsSpan(segmentStart, segmentLength).ContainsAny("\u00AD\u200B");
        bool segmentHasDynamicFields = SpansOverlapDynamicField(spans, segmentStart, segmentLength);
        List<TextToken> tokens = TokenizeSpaces(text, segmentStart, segmentLength).ToList();
        if (tokens.Count == 0)
        {
            yield break;
        }

        int lineStart = tokens[0].Start;
        int lineLength = 0;
        // HasNonWhitespace rescanned the whole line per token (quadratic
        // char scans). Maintain it incrementally instead: each token extent is scanned
        // once when added, and extents chain contiguously (lineStart always sits at a
        // token boundary or inside the re-examined remainder token), so the running OR
        // equals a fresh scan exactly.
        // RV05: image-driven breaking. Sorted image offsets with prefix sums make range
        // width lookups logarithmic; text measurement and memoization stay untouched.
        var imageWrapIndex = BuildInlineImageWrapIndex(inlineImageWidths);
        int[] imageOffsets = imageWrapIndex.Offsets;
        double[] imagePrefixSums = imageWrapIndex.PrefixSums;

        bool lineHasNonWhitespace = false;
        int lineIndex = 0;
        for (int tokenIndex = 0; tokenIndex < tokens.Count; tokenIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            TextToken token = tokens[tokenIndex];
            int candidateLength = token.Start + token.Length - lineStart;
            if (lineLength > 0 &&
                lineHasNonWhitespace &&
                !token.IsBreakableWhitespace &&
                MeasureWrapSlice(measureMemo, segmentHasHiddenBreaks, segmentHasDynamicFields, spans, lineStart, candidateLength, fontSize, textMeasurer, tabStops, defaultTabStopPoints, preserveTerminalSoftHyphen: false, dynamicFieldPageNumber, spanStarts) + ImageWidthInRange(imageWrapIndex, lineStart, lineStart + candidateLength) > maxWidth(lineIndex))
            {
                // RV05: image-driven breaking. A line-owning image that overflows the remainder
                // breaks onto the next line at its own offset (words stay whole); overwide images
                // that fit no line keep the overwide-text convention and overflow.
                int imageBreakOffset = -1;
                for (int imageBreakIndex = 0; imageBreakIndex < imageOffsets.Length; imageBreakIndex++)
                {
                    int imageOffset = imageOffsets[imageBreakIndex];
                    if (imageOffset <= lineStart)
                    {
                        continue;
                    }

                    if (imageOffset >= lineStart + candidateLength)
                    {
                        break;
                    }

                    double overflowingImageWidth = imagePrefixSums[imageBreakIndex + 1] - imagePrefixSums[imageBreakIndex];
                    double beforeImageWidth = MeasureWrapSlice(measureMemo, segmentHasHiddenBreaks, segmentHasDynamicFields, spans, lineStart, imageOffset - lineStart, fontSize, textMeasurer, tabStops, defaultTabStopPoints, preserveTerminalSoftHyphen: false, dynamicFieldPageNumber, spanStarts) + ImageWidthInRange(imageWrapIndex, lineStart, imageOffset);
                    if (beforeImageWidth <= maxWidth(lineIndex) && beforeImageWidth + overflowingImageWidth > maxWidth(lineIndex) && overflowingImageWidth <= maxWidth(lineIndex + 1))
                    {
                        imageBreakOffset = imageOffset;
                        break;
                    }
                }

                if (imageBreakOffset > lineStart)
                {
                    yield return CreateWrappedTextLine(text, spans, lineStart, imageBreakOffset - lineStart, false, spanStarts);
                    lineIndex++;
                    lineStart = imageBreakOffset;
                    lineLength = 0;
                    lineHasNonWhitespace = false;
                    tokenIndex--;
                    continue;
                }

                // Word also breaks an overlong token after a hyphen (or slash) when the prefix fits the remaining width; previously only line-leading overwide tokens used preferred breaks.
                double usedWidth = MeasureWrapSlice(measureMemo, segmentHasHiddenBreaks, segmentHasDynamicFields, spans, lineStart, lineLength, fontSize, textMeasurer, tabStops, defaultTabStopPoints, preserveTerminalSoftHyphen: false, dynamicFieldPageNumber, spanStarts) + ImageWidthInRange(imageWrapIndex, lineStart, lineStart + lineLength);
                double remainingWidth = maxWidth(lineIndex) - usedWidth;
                if (remainingWidth > 0d &&
                    TryFindPreferredTokenBreak(text, spans, token, remainingWidth, fontSize, textMeasurer, tabStops, defaultTabStopPoints, dynamicFieldPageNumber, measureMemo, segmentHasHiddenBreaks, segmentHasDynamicFields, cancellationToken, spanStarts, chainAverages, out int preferredBreakLength))
                {
                    yield return CreateWrappedTextLine(text, spans, lineStart, lineLength + preferredBreakLength, endsWithIntraTokenBreak: true, spanStarts);
                    lineIndex++;
                    lineStart = token.Start + preferredBreakLength;
                    lineLength = 0;
                    lineHasNonWhitespace = false;
                    // R07.2: remainder stays a range over the original text; no suffix copy.
                    tokens[tokenIndex] = new TextToken(lineStart, token.Length - preferredBreakLength, false);
                    tokenIndex--;
                    continue;
                }

                yield return CreateWrappedTextLine(text, spans, lineStart, lineLength, false, spanStarts);
                lineIndex++;
                lineStart = token.Start;
                lineLength = 0;
                lineHasNonWhitespace = false;
                tokenIndex--;
                continue;
            }

            if (lineLength == 0 &&
                !token.IsBreakableWhitespace &&
                TryFindPreferredTokenBreak(text, spans, token, maxWidth(lineIndex), fontSize, textMeasurer, tabStops, defaultTabStopPoints, dynamicFieldPageNumber, measureMemo, segmentHasHiddenBreaks, segmentHasDynamicFields, cancellationToken, spanStarts, chainAverages, out int breakLength))
            {
                yield return CreateWrappedTextLine(text, spans, token.Start, breakLength, endsWithIntraTokenBreak: true, spanStarts);
                lineIndex++;
                lineStart = token.Start + breakLength;
                lineLength = 0;
                lineHasNonWhitespace = false;
                // R07.2: remainder stays a range over the original text; no suffix copy.
                    tokens[tokenIndex] = new TextToken(lineStart, token.Length - breakLength, false);
                tokenIndex--;
            }
            else if (lineLength == 0 &&
                allowOverwideTokenBreaks &&
                !token.IsBreakableWhitespace &&
                TryFindOverwideTokenBreak(text, spans, token, maxWidth(lineIndex), fontSize, textMeasurer, tabStops, defaultTabStopPoints, dynamicFieldPageNumber, measureMemo, segmentHasHiddenBreaks, segmentHasDynamicFields, cancellationToken, spanStarts, chainAverages, out breakLength))
            {
                yield return CreateWrappedTextLine(text, spans, token.Start, breakLength, endsWithIntraTokenBreak: true, spanStarts);
                lineIndex++;
                lineStart = token.Start + breakLength;
                lineLength = 0;
                lineHasNonWhitespace = false;
                // R07.2: remainder stays a range over the original text; no suffix copy.
                    tokens[tokenIndex] = new TextToken(lineStart, token.Length - breakLength, false);
                tokenIndex--;
            }
            else
            {
                int addedStart = Math.Max(token.Start, lineStart);
                lineHasNonWhitespace = lineHasNonWhitespace || HasNonWhitespace(text, addedStart, token.Start + token.Length - addedStart);
                lineLength = candidateLength;
            }
        }

        if (lineLength > 0)
        {
            yield return CreateWrappedTextLine(text, spans, lineStart, lineLength, false, spanStarts);
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
        Dictionary<(int Start, int Length, bool PreserveTerminalSoftHyphen), double> measureMemo,
        bool segmentHasHiddenBreaks,
        bool segmentHasDynamicFields,
        CancellationToken cancellationToken,
        int[] spanStarts,
        Dictionary<int, double> chainAverages,
        out int breakLength)
    {
        if (TryFindPreferredTokenBreak(text, spans, token, maxWidth, fontSize, textMeasurer, tabStops, defaultTabStopPoints, dynamicFieldPageNumber, measureMemo, segmentHasHiddenBreaks, segmentHasDynamicFields, cancellationToken, spanStarts, chainAverages, out breakLength))
        {
            return true;
        }

        double averageCharWidth;
        if (!TryEstimateChainOverflow(chainAverages, token, maxWidth, out averageCharWidth))
        {
            double tokenWidth = MeasureWrapSlice(measureMemo, segmentHasHiddenBreaks, segmentHasDynamicFields, spans, token.Start, token.Length, fontSize, textMeasurer, tabStops, defaultTabStopPoints, preserveTerminalSoftHyphen: false, dynamicFieldPageNumber, spanStarts);
            RecordChainAverageWidth(chainAverages, token, tokenWidth);
            if (tokenWidth <= maxWidth)
            {
                breakLength = 0;
                return false;
            }

            averageCharWidth = token.Length > 0 ? tokenWidth / token.Length : 0d;
        }

        // R07: estimate fit via average char width to avoid O(N) prefix measures per line.
        // Prefix widths are not provably monotonic, so local grow/shrink finds the longest
        // fitting safe prefix under monotonic assumption; non-monotonic longer fits beyond
        // the first overflow are accepted as shorter safe breaks (still fitting) to bound work.
        if (token.Length <= 1)
        {
            breakLength = 0;
            return false;
        }

        int estimatedFit = averageCharWidth > 0d
            ? Math.Clamp((int)(maxWidth / averageCharWidth), 1, token.Length - 1)
            : 1;
        int candidate = estimatedFit;
        while (candidate > 1 && !IsSafeEmergencyTokenBreak(candidate))
        {
            cancellationToken.ThrowIfCancellationRequested();
            candidate--;
        }

        if (candidate <= 0 || !IsSafeEmergencyTokenBreak(candidate))
        {
            breakLength = 0;
            return false;
        }

        bool candidatePreserveTerminalSoftHyphen = text[token.Start + candidate - 1] == '\u00AD';
        double candidateWidth = MeasureWrapSlice(measureMemo, segmentHasHiddenBreaks, segmentHasDynamicFields, spans, token.Start, candidate, fontSize, textMeasurer, tabStops, defaultTabStopPoints, candidatePreserveTerminalSoftHyphen, dynamicFieldPageNumber, spanStarts);
        if (candidateWidth <= maxWidth)
        {
            int best = candidate;
            int next = candidate + 1;
            int growProbes = 0;
            while (next < token.Length)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (++growProbes > MaxSearchProbesPerLine)
                {
                    // R07.2: emit the fitting candidate instead of walking further.
                    break;
                }
                if (!IsSafeEmergencyTokenBreak(next))
                {
                    next++;
                    continue;
                }

                bool nextPreserveTerminalSoftHyphen = text[token.Start + next - 1] == '\u00AD';
                double nextWidth = MeasureWrapSlice(measureMemo, segmentHasHiddenBreaks, segmentHasDynamicFields, spans, token.Start, next, fontSize, textMeasurer, tabStops, defaultTabStopPoints, nextPreserveTerminalSoftHyphen, dynamicFieldPageNumber, spanStarts);
                if (nextWidth <= maxWidth)
                {
                    best = next;
                    next++;
                }
                else
                {
                    break;
                }
            }

            // RV13: the scan reached the end of the token with every proper
            // prefix fitting (cap and overflow exits keep next short of the
            // end). The skipped full measure may have estimated overflow on a
            // mixed-width chain, so confirm: a fitting whole remainder needs
            // no break. Uses the same unpreserved fit question as the gate.
            if (next >= token.Length)
            {
                double fullWidth = MeasureWrapSlice(measureMemo, segmentHasHiddenBreaks, segmentHasDynamicFields, spans, token.Start, token.Length, fontSize, textMeasurer, tabStops, defaultTabStopPoints, preserveTerminalSoftHyphen: false, dynamicFieldPageNumber, spanStarts);
                if (fullWidth <= maxWidth)
                {
                    breakLength = 0;
                    return false;
                }
            }

            breakLength = best;
            return true;
        }
        else
        {
            int current = candidate - 1;
            int shrinkProbes = 0;
            while (current > 0)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!IsSafeEmergencyTokenBreak(current))
                {
                    current--;
                    continue;
                }

                if (++shrinkProbes > MaxSearchProbesPerLine)
                {
                    throw new OoxPdfLimitExceededException("Emergency wrap search exceeded " + MaxSearchProbesPerLine + " shrink probes on one line; widths are non-monotonic over the feasible prefixes.");
                }

                bool currentPreserveTerminalSoftHyphen = text[token.Start + current - 1] == '\u00AD';
                double currentWidth = MeasureWrapSlice(measureMemo, segmentHasHiddenBreaks, segmentHasDynamicFields, spans, token.Start, current, fontSize, textMeasurer, tabStops, defaultTabStopPoints, currentPreserveTerminalSoftHyphen, dynamicFieldPageNumber, spanStarts);
                if (currentWidth <= maxWidth)
                {
                    breakLength = current;
                    return true;
                }

                current--;
            }

            breakLength = 0;
            return false;
        }


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
        Dictionary<(int Start, int Length, bool PreserveTerminalSoftHyphen), double> measureMemo,
        bool segmentHasHiddenBreaks,
        bool segmentHasDynamicFields,
        CancellationToken cancellationToken,
        int[] spanStarts,
        Dictionary<int, double> chainAverages,
        out int breakLength)
    {
        breakLength = 0;
        double averagePreferredCharWidth;
        if (!TryEstimateChainOverflow(chainAverages, token, maxWidth, out averagePreferredCharWidth))
        {
            double tokenWidth = MeasureWrapSlice(measureMemo, segmentHasHiddenBreaks, segmentHasDynamicFields, spans, token.Start, token.Length, fontSize, textMeasurer, tabStops, defaultTabStopPoints, preserveTerminalSoftHyphen: false, dynamicFieldPageNumber, spanStarts);
            RecordChainAverageWidth(chainAverages, token, tokenWidth);
            if (tokenWidth <= maxWidth)
            {
                return false;
            }

            averagePreferredCharWidth = token.Length > 0 ? tokenWidth / token.Length : 0d;
        }

        if (token.Length <= 1)
        {
            breakLength = 0;
            return false;
        }

        int estimatedPreferredFit = averagePreferredCharWidth > 0d
            ? Math.Clamp((int)(maxWidth / averagePreferredCharWidth), 1, token.Length - 1)
            : 1;
        int preferredCandidate = estimatedPreferredFit;
        while (preferredCandidate > 1 && !DocxLineBreakOpportunities.IsOpportunityAfter(text[token.Start + preferredCandidate - 1]))
        {
            cancellationToken.ThrowIfCancellationRequested();
            preferredCandidate--;
        }

        if (preferredCandidate <= 0 || !DocxLineBreakOpportunities.IsOpportunityAfter(text[token.Start + preferredCandidate - 1]))
        {
            int upwardPreferred = estimatedPreferredFit + 1;
            while (upwardPreferred < token.Length)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (DocxLineBreakOpportunities.IsOpportunityAfter(text[token.Start + upwardPreferred - 1]))
                {
                    break;
                }

                upwardPreferred++;
            }

            if (upwardPreferred >= token.Length || !DocxLineBreakOpportunities.IsOpportunityAfter(text[token.Start + upwardPreferred - 1]))
            {
                breakLength = 0;
                return false;
            }

            preferredCandidate = upwardPreferred;
        }

        bool preferredCandidatePreserve = text[token.Start + preferredCandidate - 1] == '\u00AD';
        double preferredCandidateWidth = MeasureWrapSlice(measureMemo, segmentHasHiddenBreaks, segmentHasDynamicFields, spans, token.Start, preferredCandidate, fontSize, textMeasurer, tabStops, defaultTabStopPoints, preferredCandidatePreserve, dynamicFieldPageNumber, spanStarts);
        if (preferredCandidateWidth <= maxWidth)
        {
            int bestPreferred = preferredCandidate;
            int nextPreferred = preferredCandidate + 1;
            int preferredGrowProbes = 0;
            while (nextPreferred < token.Length)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (++preferredGrowProbes > MaxSearchProbesPerLine)
                {
                    // R07.2: emit the fitting candidate instead of walking further.
                    break;
                }
                if (!DocxLineBreakOpportunities.IsOpportunityAfter(text[token.Start + nextPreferred - 1]))
                {
                    nextPreferred++;
                    continue;
                }

                bool nextPreferredPreserve = text[token.Start + nextPreferred - 1] == '\u00AD';
                double nextPreferredWidth = MeasureWrapSlice(measureMemo, segmentHasHiddenBreaks, segmentHasDynamicFields, spans, token.Start, nextPreferred, fontSize, textMeasurer, tabStops, defaultTabStopPoints, nextPreferredPreserve, dynamicFieldPageNumber, spanStarts);
                if (nextPreferredWidth <= maxWidth)
                {
                    bestPreferred = nextPreferred;
                    nextPreferred++;
                }
                else
                {
                    break;
                }
            }

            // RV13: same fitting-tail confirmation as the emergency search: a
            // whole remainder that fits needs no preferred break either.
            if (nextPreferred >= token.Length)
            {
                double fullPreferredWidth = MeasureWrapSlice(measureMemo, segmentHasHiddenBreaks, segmentHasDynamicFields, spans, token.Start, token.Length, fontSize, textMeasurer, tabStops, defaultTabStopPoints, preserveTerminalSoftHyphen: false, dynamicFieldPageNumber, spanStarts);
                if (fullPreferredWidth <= maxWidth)
                {
                    breakLength = 0;
                    return false;
                }
            }

            breakLength = bestPreferred;
            return true;
        }
        else
        {
            int currentPreferred = preferredCandidate - 1;
            int preferredShrinkProbes = 0;
            while (currentPreferred > 0)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!DocxLineBreakOpportunities.IsOpportunityAfter(text[token.Start + currentPreferred - 1]))
                {
                    currentPreferred--;
                    continue;
                }

                bool currentPreferredPreserve = text[token.Start + currentPreferred - 1] == '\u00AD';
                if (++preferredShrinkProbes > MaxSearchProbesPerLine)
                {
                    throw new OoxPdfLimitExceededException("Preferred wrap search exceeded " + MaxSearchProbesPerLine + " shrink probes on one line; widths are non-monotonic over the feasible prefixes.");
                }

                double currentPreferredWidth = MeasureWrapSlice(measureMemo, segmentHasHiddenBreaks, segmentHasDynamicFields, spans, token.Start, currentPreferred, fontSize, textMeasurer, tabStops, defaultTabStopPoints, currentPreferredPreserve, dynamicFieldPageNumber, spanStarts);
                if (currentPreferredWidth <= maxWidth)
                {
                    breakLength = currentPreferred;
                    return true;
                }

                currentPreferred--;
            }

            breakLength = 0;
            return false;
        }

    }

    private static double MeasureWrapSlice(
        Dictionary<(int Start, int Length, bool PreserveTerminalSoftHyphen), double> measureMemo,
        bool segmentHasHiddenBreaks,
        bool segmentHasDynamicFields,
        IReadOnlyList<DocxTextSpan> spans,
        int start,
        int length,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints,
        bool preserveTerminalSoftHyphen,
        int? dynamicFieldPageNumber,
        int[]? spanStarts = null)
    {
        var key = (start, length, preserveTerminalSoftHyphen);
        if (measureMemo.TryGetValue(key, out double cached))
        {
            return cached;
        }

        IReadOnlyList<DocxTextSpan> sliced = SliceTextSpans(spans, start, length, spanStarts);
        // Clean segments skip both normalizers exactly: each returns its input unchanged
        // when there is nothing to rewrite, so slicing straight into field measurement
        // computes the identical width without the per-slice scans.
        double width = segmentHasHiddenBreaks || segmentHasDynamicFields
            ? MeasureTextSpansForWrapping(sliced, fontSize, textMeasurer, tabStops, defaultTabStopPoints, preserveTerminalSoftHyphen, dynamicFieldPageNumber)
            : MeasureTextSpans(sliced, fontSize, textMeasurer, tabStops, defaultTabStopPoints);
        measureMemo[key] = width;
        return width;
    }

    private static bool SpansOverlapDynamicField(IReadOnlyList<DocxTextSpan> spans, int start, int length)
    {
        int spanStart = 0;
        int end = start + length;
        foreach (DocxTextSpan span in spans)
        {
            int spanEnd = spanStart + span.Text.Length;
            if (spanEnd > start && spanStart < end && span.StyleRun.FieldKind == DocxFieldKind.NumPages)
            {
                return true;
            }

            if (spanEnd >= end)
            {
                break;
            }

            spanStart = spanEnd;
        }

        return false;
    }

    // RV13: per-token-chain average advance used to skip repeat full-remainder
    // measures. Remainders produced while wrapping one token chain share the
    // chain end (token.Start + token.Length), which uniquely identifies the
    // chain within a segment, while font size, tab stops and field state are
    // fixed for the wrapping call. An estimated overflow skips the full
    // measure and proceeds to the break search; an estimated fit always falls
    // through to a confirming real measure, so break decisions match full
    // measurement. Like the slice memo, averages live for one segment only.
    private static bool TryEstimateChainOverflow(
        Dictionary<int, double> chainAverages,
        TextToken token,
        double maxWidth,
        out double averageCharWidth)
    {
        if (chainAverages.TryGetValue(token.Start + token.Length, out averageCharWidth) &&
            averageCharWidth * token.Length > maxWidth)
        {
            return true;
        }

        averageCharWidth = 0d;
        return false;
    }

    private static void RecordChainAverageWidth(
        Dictionary<int, double> chainAverages,
        TextToken token,
        double tokenWidth)
    {
        if (token.Length > 0)
        {
            chainAverages[token.Start + token.Length] = tokenWidth / token.Length;
        }
    }

    private static DocxWrappedTextLine CreateWrappedTextLine(
        string text,
        IReadOnlyList<DocxTextSpan> spans,
        int start,
        int length,
        bool endsWithIntraTokenBreak,
        int[]? spanStarts = null)
    {
        string lineText = text.Substring(start, length);
        bool preserveTerminalSoftHyphen = endsWithIntraTokenBreak && lineText.EndsWith('\u00AD');
        return new DocxWrappedTextLine(
            RemoveHiddenBreakCharacters(lineText, preserveTerminalSoftHyphen),
            NormalizeHiddenBreakSpans(SliceTextSpans(spans, start, length, spanStarts), preserveTerminalSoftHyphen),
            endsWithIntraTokenBreak);
    }

    private static int FindSpanIndex(int[] spanStarts, int position)
    {
        // Binary search for the span containing position; spanStarts holds one
        // cumulative start per span plus the total.
        int index = Array.BinarySearch(spanStarts, 0, spanStarts.Length - 1, position);
        return index >= 0 ? Math.Min(index, spanStarts.Length - 2) : Math.Max(0, ~index - 1);
    }

    private static IReadOnlyList<DocxTextSpan> SliceTextSpans(
        IReadOnlyList<DocxTextSpan> spans,
        int start,
        int length,
        int[]? spanStarts = null)
    {
        if (length == 0)
        {
            return [];
        }

        var sliced = new List<DocxTextSpan>();
        int spanIndex = 0;
        int spanStart = 0;
        if (spanStarts is not null)
        {
            // R07.2: seek the first overlapping span instead of re-walking.
            spanIndex = FindSpanIndex(spanStarts, start);
            spanStart = spanStarts[spanIndex];
        }

        int end = start + length;
        for (int i = spanIndex; i < spans.Count; i++)
        {
            DocxTextSpan span = spans[i];
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
            !spans.Any(span => span.StyleRun.FieldKind == DocxFieldKind.NumPages))
        {
            return spans;
        }

        string proxy = pageNumber.Value.ToString(CultureInfo.InvariantCulture);
        return spans
            .Select(span => span.StyleRun.FieldKind == DocxFieldKind.NumPages
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
        // R07.2: tokenize by range over the segment string: no per-token substring
        // copies, and whitespace-ness rides the token instead of rescanning text.
        if (length == 0)
        {
            return [];
        }

        var tokens = new List<TextToken>();
        int end = start + length;
        int tokenStart = start;
        bool inBreakableWhitespace = DocxTextBreakRules.IsBreakableWhitespaceChar(text[start]);
        for (int i = start + 1; i < end; i++)
        {
            bool breakableWhitespace = DocxTextBreakRules.IsBreakableWhitespaceChar(text[i]);
            if (breakableWhitespace == inBreakableWhitespace)
            {
                continue;
            }

            tokens.Add(new TextToken(tokenStart, i - tokenStart, inBreakableWhitespace));
            tokenStart = i;
            inBreakableWhitespace = breakableWhitespace;
        }

        tokens.Add(new TextToken(tokenStart, end - tokenStart, inBreakableWhitespace));
        return tokens;
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


    private static class DocxLineBreakOpportunities
    {
        public static bool IsOpportunityAfter(char value)
        {
            return value is '-' or '/' or '\\' or '\u00AD' or '\u200B' or '\u2010' or '\u2012' or '\u2013' or '\u2014';
        }
    }

    private readonly record struct TextToken(int Start, int Length, bool IsBreakableWhitespace);
}
