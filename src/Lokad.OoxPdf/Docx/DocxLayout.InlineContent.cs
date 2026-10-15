using System.Diagnostics.CodeAnalysis;

namespace Lokad.OoxPdf.Docx;

// Measurements contain no page coordinates. Containers own line admission and
// alignment; source items and their original annotation indices survive wrapping.
internal sealed record DocxMeasuredInlineFragment(DocxInlineItem Source, double Advance,
    double BaselineInset, double OccupiedHeight, IReadOnlyList<DocxTextSpan> Text,
    DocxPreparedMathParagraph? Math = null)
{
    public double LeadingAdvance { get; init; }
    public IReadOnlyList<DocxTextSegmentLayout> TextSegments { get; init; } = [];
}
internal sealed record DocxMeasuredInlineLine(IReadOnlyList<DocxMeasuredInlineFragment> Fragments,
    double Advance, double BaselineInset, double OccupiedHeight)
{
    public bool IsDisplayGroup { get; init; }
}
internal readonly record struct DocxInlineParagraphFrame(double OccupiedHeight, DocxTextLineLayout LastLine);

internal sealed partial class DocxLayoutEngine
{
    private static double ResolveOrderedInlineAlignment(DocxMeasuredInlineLine line,
        DocxTextAlignment alignment, double available) =>
        (line.IsDisplayGroup ? DocxTextAlignment.Center : alignment) switch
        {
            DocxTextAlignment.Center => Math.Max(0d, available - line.Advance) / 2d,
            DocxTextAlignment.Right => Math.Max(0d, available - line.Advance),
            _ => 0d
        };

    private static double ResolveOrderedBodyBaselineInset(DocxParagraph paragraph, double size,
        double lineHeight, IDocxTextMeasurer measurer) =>
        DocxLineMetrics.ResolveBodyBaselineOffset(size, lineHeight, false,
            DocxLineMetrics.ResolveHheaAscenderPoints(paragraph, size, measurer),
            DocxLineMetrics.ResolveBodyTierAMaxPoints(paragraph, measurer));

    private static IReadOnlyDictionary<DocxParagraph, DocxInlineParagraphFrame>? GetOrderedInlineParagraphFrames(
        IReadOnlyList<DocxTextLineLayout> lines, CancellationToken token)
    {
        Dictionary<DocxParagraph, DocxInlineParagraphFrame>? heights = null;
        foreach (var line in lines)
        {
            token.ThrowIfCancellationRequested();
            if (line.InlineMath.Count == 0 || line.SourceParagraph is not { } paragraph) continue;
            heights ??= new(ReferenceEqualityComparer.Instance);
            heights.TryAdd(paragraph, new(0d, line));
        }
        if (heights is null) return null;
        foreach (var line in lines)
        {
            token.ThrowIfCancellationRequested();
            if (line.SourceParagraph is { } paragraph && heights.TryGetValue(paragraph, out var frame))
                heights[paragraph] = new(frame.OccupiedHeight + (line.LineHeight ?? 0d), line);
        }
        return heights;
    }

    private static bool TryMeasureOrderedCellContent(DocxTableCell cell, DocxParagraph paragraph,
        IReadOnlyList<DocxParagraph> paragraphs, DocxTableCellLayoutContext context,
        double textWidth, double fontSize, double lineHeight,
        [NotNullWhen(true)] out IReadOnlyList<DocxMeasuredInlineLine>? lines)
    {
        lines = null;
        if (!DocxMathParagraphPolicy.NeedsOrderedInlineLayout(paragraph)) return false;
        double scale = context.ParagraphSpacingScale;
        double firstInset = ResolveTableCellFirstBaselineInset(paragraphs, context.TextMeasurer);
        double paragraphInset = DocxLineMetrics.ResolveTableCellBodyBaselineOffset(fontSize, lineHeight,
            IsExactLineSpacing(paragraph.EffectiveProperties), paragraphs, paragraph, context.TextMeasurer);
        double baselineInset = firstInset + (paragraphInset - firstInset) * scale;
        double right = GetParagraphRightInset(paragraph, scale);
        double firstWidth = ResolveTableCellTextWrapWidth(cell, textWidth - right -
            GetParagraphFirstLineTextStartOffset(paragraph, fontSize, context.TextMeasurer, scale));
        double nextWidth = ResolveTableCellTextWrapWidth(cell, textWidth - right - GetParagraphTextStartOffset(paragraph, scale));
        return TryMeasureOrderedInlineContent(paragraph, context.MathEquations, context.TextMeasurer, fontSize,
            firstWidth, nextWidth, baselineInset, lineHeight, scale, context.CancellationToken, out lines);
    }

    private static (DocxTextLineLayout Line, IReadOnlyList<DocxInlineImageLayout> Images) PlaceOrderedInlineLine(
        DocxParagraph paragraph, DocxMeasuredInlineLine measured,
        double x, double frameOriginX, double top, double fontSize, double scale, int pageIndex,
        int? sourceBlock, int? sourceParagraph, int sourceLine, DocxStoryId story)
    {
        var segments = new List<DocxTextSegmentLayout>();
        var images = new List<DocxInlineImageLayout>();
        var equations = new List<DocxInlineMathPlacement>();
        var text = new System.Text.StringBuilder();
        double offset = 0d;
        foreach (var fragment in measured.Fragments)
        {
            offset += fragment.LeadingAdvance;
            if (fragment.Text.Count != 0)
            {
                foreach (var segment in fragment.TextSegments)
                    segments.Add(segment with { X = segment.X + x + offset });
                foreach (var span in fragment.Text) text.Append(span.Text);
            }
            else if (fragment.Math is { } math) equations.Add(new(math, offset));
            else if (fragment.Source is DocxInlinePicture picture)
                images.Add(new(picture.Image, x + offset, top - measured.BaselineInset, fragment.Advance,
                    fragment.OccupiedHeight, pageIndex, sourceBlock, sourceParagraph, story));
            offset += fragment.Advance;
        }
        var style = paragraph.Runs.FirstOrDefault() ?? paragraph.ParagraphMarkRun ??
            new DocxTextRun(string.Empty, fontSize, null, false, false, false, null, null);
        var line = new DocxTextLineLayout(text.ToString(), style, fontSize, x, top - measured.BaselineInset,
            measured.Advance, segments, sourceBlock, sourceParagraph, sourceLine, story,
            LineHeight: measured.OccupiedHeight, AppliedBeforeSpacing: 0d, IsFirstParagraphLine: sourceLine == 0,
            EndsWithIntraTokenBreak: false, SingleLineHeight: null, ListLabelSingleLineHeight: null,
            BodyWindowsLineHeight: null, ListLabelWindowsLineHeight: null, EffectiveLineSpacingFactor: null,
            LineSpacingFactorFloorApplied: null, PendingAfterSpacing: null, ParagraphBeforeSpacing: null,
            ParagraphAfterSpacing: null, ContextualSpacingSuppressed: null, SourceParagraph: paragraph,
            LineHeightSource: null, EmitsTerminalParagraphMark: false)
        { FrameOriginOffsetX = frameOriginX - x, InlineMath = equations, MathScale = scale, InlineObjectEndOffsetX = measured.Advance,
            InlineFrame = new(measured.BaselineInset, measured.OccupiedHeight) };
        return (line, images);
    }

    internal static bool TryMeasureOrderedInlineContent(DocxParagraph paragraph,
        IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph>? equations,
        IDocxTextMeasurer measurer, double fontSize, double firstWidth, double nextWidth,
        double baselineInset, double lineHeight, double scale, CancellationToken token,
        [NotNullWhen(true)] out IReadOnlyList<DocxMeasuredInlineLine>? lines)
    {
        lines = null;
        if (paragraph.InlineContent is not { } items || equations is null ||
            paragraph.Equations.Count == 0 ||
            DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count != paragraph.Equations.Count ||
            items.Any(item => item is DocxInlineTab tab && !tab.Run.EffectiveProperties.Hidden)) return false;
        if (!double.IsFinite(firstWidth) || !double.IsFinite(nextWidth) || firstWidth <= 0d || nextWidth <= 0d ||
            !double.IsFinite(scale) || scale <= 0d) return false;
        var runIndices = new Dictionary<DocxTextRun, int>(ReferenceEqualityComparer.Instance);
        for (int index = 0; index < paragraph.Runs.Count; index++)
        {
            token.ThrowIfCancellationRequested();
            if (!runIndices.TryAdd(paragraph.Runs[index], index)) return false;
        }
        var measured = new List<DocxMeasuredInlineFragment>();
        foreach (DocxInlineItem item in items)
        {
            token.ThrowIfCancellationRequested();
            switch (item)
            {
                case DocxInlineText text when text.Run.EffectiveProperties.Hidden:
                case DocxInlineTab tab when tab.Run.EffectiveProperties.Hidden:
                case DocxInlineBreak lineBreak when lineBreak.Run.EffectiveProperties.Hidden:
                    break;
                case DocxInlineText text:
                    int start = text.Start, end = start + text.Length;
                    while (start < end)
                    {
                        token.ThrowIfCancellationRequested();
                        bool space = DocxTextBreakRules.IsBreakableWhitespaceChar(text.Run.Text[start]);
                        int stop = start + 1;
                        while (stop < end && DocxTextBreakRules.IsBreakableWhitespaceChar(text.Run.Text[stop]) == space)
                        {
                            if ((stop & 255) == 0) token.ThrowIfCancellationRequested();
                            stop++;
                        }
                        var source = new DocxInlineText(text.Run, start, stop - start);
                        string value = text.Run.Text.Substring(start, stop - start);
                        // Keep the shared OOXML source affinity and offsets,
                        // including runs split or filtered during reading.
                        if (!runIndices.TryGetValue(text.Run, out int runIndex)) return false;
                        var span = CreateTextSpan(value, text.Run, runIndex);
                        span = span with { SourceTextOffsetInRun = span.SourceTextOffsetInRun + start };
                        var segments = CreateTextSegments([span], 0d, fontSize, measurer, [], 36d);
                        double width = segments.Count == 0 ? 0d : segments[^1].X + segments[^1].Width;
                        if (!double.IsFinite(width) || width < 0d) return false;
                        measured.Add(new(source, width, baselineInset, Math.Max(lineHeight, baselineInset), [span])
                            { TextSegments = segments });
                        start = stop;
                    }
                    break;
                case DocxInlineEquation math:
                    if (!equations.TryGetValue(math.Equation, out var prepared)) return false;
                    measured.Add(new(item, prepared.Box.AdvanceWidth * scale, prepared.BaselineInset * scale,
                        prepared.LineHeight * scale, [], prepared));
                    break;
                case DocxInlinePicture picture:
                    double imageWidth = picture.Image.WidthPoints * scale, imageHeight = picture.Image.HeightPoints * scale;
                    if (!double.IsFinite(imageWidth) || !double.IsFinite(imageHeight) || imageWidth <= 0d || imageHeight <= 0d) return false;
                    measured.Add(new(item, imageWidth, imageHeight, imageHeight, []));
                    break;
                case DocxInlineBreak:
                    measured.Add(new(item, 0d, baselineInset, Math.Max(lineHeight, baselineInset), []));
                    break;
                default: return false;
            }
        }
        if (measured.Any(fragment => !double.IsFinite(fragment.Advance) || fragment.Advance < 0d ||
            !double.IsFinite(fragment.BaselineInset) || fragment.BaselineInset < 0d ||
            !double.IsFinite(fragment.OccupiedHeight) || fragment.OccupiedHeight < fragment.BaselineInset)) return false;

        var output = new List<DocxMeasuredInlineLine>();
        var current = new List<DocxMeasuredInlineFragment>();
        double advance = 0d, above = baselineInset, below = Math.Max(0d, lineHeight - baselineInset);
        for (int index = 0; index < measured.Count;)
        {
            token.ThrowIfCancellationRequested();
            if (measured[index].Source is DocxInlineBreak) { Flush(); index++; continue; }
            int end = index + 1;
            bool space = IsSpace(measured[index]);
            // A formatting/run boundary is not a word-break opportunity.
            if (!space && measured[index].Source is DocxInlineText)
                while (end < measured.Count && measured[end].Source is DocxInlineText && !IsSpace(measured[end])) end++;
            double groupWidth = 0d;
            for (int offset = index; offset < end; offset++)
            {
                if (offset > index) groupWidth += InlineTextBoundaryAdvance(measured[offset - 1], measured[offset]);
                groupWidth += measured[offset].Advance;
            }
            double width = output.Count == 0 ? firstWidth : nextWidth;
            if (groupWidth > Math.Max(firstWidth, nextWidth))
            {
                if (measured[index].Source is not DocxInlineText) return false;
                if (current.Count != 0) Flush();
                var spans = measured.Skip(index).Take(end - index).SelectMany(fragment => fragment.Text).ToArray();
                // Reuse the existing bounded break search, Unicode rules and
                // authored character-wrap policy for an overwide text word.
                var wrapped = WrapTextLines(spans, output.Count == 0 ? firstWidth : nextWidth, nextWidth,
                    fontSize, measurer, [], 36d, ShouldAllowCharacterLevelWordWrap(paragraph), dynamicFieldPageNumber: null,
                    cancellationToken: token).ToArray();
                if (wrapped.Length == 0) return false;
                for (int lineIndex = 0; lineIndex < wrapped.Length; lineIndex++)
                {
                    token.ThrowIfCancellationRequested();
                    foreach (var span in wrapped[lineIndex].Spans)
                    {
                        int offset = span.SourceTextOffsetInRun - Math.Max(0, span.StyleRun.SourceTextOffsetInRun);
                        if (offset < 0 || offset > span.StyleRun.Text.Length) return false;
                        var source = new DocxInlineText(span.StyleRun, offset, Math.Min(span.Text.Length, span.StyleRun.Text.Length - offset));
                        var segments = CreateTextSegments([span], 0d, fontSize, measurer, [], 36d);
                        double partWidth = segments.Count == 0 ? 0d : segments[^1].X + segments[^1].Width;
                        if (!double.IsFinite(partWidth) || partWidth < 0d) return false;
                        var part = new DocxMeasuredInlineFragment(source, partWidth, baselineInset, Math.Max(lineHeight, baselineInset), [span])
                            { TextSegments = segments };
                        part = part with { LeadingAdvance = current.Count == 0 ? 0d : InlineTextBoundaryAdvance(current[^1], part) };
                        current.Add(part);
                        advance += part.LeadingAdvance + partWidth;
                    }
                    if (lineIndex + 1 < wrapped.Length) Flush();
                }
                index = end;
                continue;
            }
            double boundary = current.Count == 0 ? 0d : InlineTextBoundaryAdvance(current[^1], measured[index]);
            if (current.Count != 0 && advance + boundary + groupWidth > width) Flush();
            if (current.Count == 0 && groupWidth > (output.Count == 0 ? firstWidth : nextWidth))
            {
                if (output.Count == 0 && groupWidth <= nextWidth) Flush();
                else return false;
            }
            if (space && current.Count == 0) { index = end; continue; }
            for (; index < end; index++)
            {
                var fragment = measured[index] with
                    { LeadingAdvance = current.Count == 0 ? 0d : InlineTextBoundaryAdvance(current[^1], measured[index]) };
                current.Add(fragment); advance += fragment.LeadingAdvance + fragment.Advance;
                above = Math.Max(above, fragment.BaselineInset);
                below = Math.Max(below, fragment.OccupiedHeight - fragment.BaselineInset);
            }
        }
        if (current.Count != 0 || output.Count == 0) Flush();
        lines = output; return true;

        static bool IsSpace(DocxMeasuredInlineFragment fragment) => fragment.Source is DocxInlineText { Length: > 0 } text &&
            DocxTextBreakRules.IsBreakableWhitespaceChar(text.Run.Text[text.Start]);
        double InlineTextBoundaryAdvance(DocxMeasuredInlineFragment left, DocxMeasuredInlineFragment right)
        {
            if (left.Text.Count == 0 || right.Text.Count == 0) return 0d;
            var before = left.Text[^1];
            var after = right.Text[0];
            if (!ReferenceEquals(before.StyleRun, after.StyleRun))
                return DocxTextSpacing.BoundarySpacing(before.StyleRun, before.Text, after.Text);
            // Tokenization must preserve pair kerning and character spacing
            // within an authored run. Decode complete Unicode scalars, including
            // replacement scalars for malformed input, as the text measurer does.
            System.Text.Rune.DecodeLastFromUtf16(before.Text, out var last, out _);
            System.Text.Rune.DecodeFromUtf16(after.Text, out var first, out _);
            string lastText = last.ToString(), firstText = first.ToString();
            double Measure(string value) => MeasureTextSpansForLayout([before with { Text = value }],
                fontSize, measurer, [], 36d, null);
            return Measure(lastText + firstText) - Measure(lastText) - Measure(firstText);
        }
        void Flush()
        {
            while (current.Count != 0 && IsSpace(current[^1]))
            {
                var removed = current[^1];
                advance -= removed.Advance + removed.LeadingAdvance; current.RemoveAt(current.Count - 1);
            }
            bool displayGroup = current.Any(fragment => fragment.Math is not null) &&
                current.All(fragment => fragment.Math is { UsesDisplayStyle: true } ||
                    fragment.Source is DocxInlineText text &&
                    text.Run.Text.AsSpan(text.Start, text.Length).IsWhiteSpace());
            output.Add(new(current.ToArray(), advance, above, above + below) { IsDisplayGroup = displayGroup });
            current.Clear(); advance = 0d; above = baselineInset; below = Math.Max(0d, lineHeight - baselineInset);
        }
    }
}
