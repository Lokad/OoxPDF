using System.Diagnostics.CodeAnalysis;

namespace Lokad.OoxPdf.Docx;

// Measurements contain no page coordinates. Containers own line admission and
// alignment; source items and their original annotation indices survive wrapping.
internal sealed record DocxMeasuredInlineFragment(DocxInlineItem Source, double Advance,
    double BaselineInset, double OccupiedHeight, IReadOnlyList<DocxTextSpan> Text,
    DocxPreparedMathParagraph? Math = null)
{
    public double LeadingAdvance { get; init; }
    public int SourceItemIndex { get; init; } = -1;
    public DocxInlineSourceExtent? VisibleSourceExtent { get; init; }
    public double? MathContinuationAdvance { get; init; }
    public IReadOnlyList<DocxTextSegmentLayout> TextSegments { get; init; } = [];
}
internal sealed record DocxMeasuredInlineLine(IReadOnlyList<DocxMeasuredInlineFragment> Fragments,
    double Advance, double BaselineInset, double OccupiedHeight)
{
    public bool IsDisplayGroup { get; init; }
    public double TextBaselineInset { get; init; }
    public IReadOnlyList<DocxTextSegmentLayout> ListLabelSegments { get; init; } = [];
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
        if (!DocxMathParagraphPolicy.NeedsOrderedInlineLayout(paragraph, context.MathEquations)) return false;
        double scale = context.ParagraphSpacingScale;
        double paragraphInset = DocxLineMetrics.ResolveTableCellBodyBaselineOffset(fontSize, lineHeight,
            IsExactLineSpacing(paragraph.EffectiveProperties), paragraphs, paragraph, context.TextMeasurer);
        // Ordered measurements own a physical line frame, including its text
        // strut. The nominal first-cell inset belongs to the legacy path.
        double baselineInset = paragraphInset * scale;
        double right = GetParagraphRightInset(paragraph, scale);
        double firstWidth = ResolveTableCellTextWrapWidth(cell, textWidth - right -
            GetParagraphFirstLineTextStartOffset(paragraph, fontSize, context.TextMeasurer, scale));
        double nextWidth = ResolveTableCellTextWrapWidth(cell, textWidth - right - GetParagraphTextStartOffset(paragraph, scale));
        return TryMeasureOrderedInlineContent(paragraph, context.MathEquations, context.TextMeasurer, fontSize,
            firstWidth, nextWidth, baselineInset, lineHeight, scale, context.CancellationToken, out lines,
            defaultTabStopPoints: context.DefaultTabStopPoints);
    }

    private static (DocxTextLineLayout Line, IReadOnlyList<DocxInlineImageLayout> Images) PlaceOrderedInlineLine(
        DocxParagraph paragraph, DocxMeasuredInlineLine measured,
        double x, double frameOriginX, double top, double fontSize, double scale, int pageIndex,
        int? sourceBlock, int? sourceParagraph, int sourceLine, DocxStoryId story)
    {
        var segments = new List<DocxTextSegmentLayout>();
        segments.AddRange(measured.ListLabelSegments.Select(segment => segment with { X = frameOriginX + segment.X }));
        var images = new List<DocxInlineImageLayout>();
        var equations = new List<DocxInlineMathPlacement>();
        List<DocxInlineSourceExtent>? sourceExtents = paragraph.CommentRanges.Count == 0 ? null : new();
        var text = new System.Text.StringBuilder();
        if (measured.ListLabelSegments.Count != 0 && paragraph.ListLabel is { } label)
            text.Append(label.Text).Append(label.SuffixValue switch { "nothing" => string.Empty, "space" => " ", _ => "\t" });
        double offset = 0d;
        foreach (var fragment in measured.Fragments)
        {
            offset += fragment.LeadingAdvance;
            if (fragment.Text.Count != 0)
            {
                foreach (var segment in fragment.TextSegments)
                    segments.Add(segment with { X = segment.X + x + offset,
                        BaselineOffsetY = segment.BaselineOffsetY * scale });
                foreach (var span in fragment.Text) text.Append(span.Text);
            }
            else if (fragment.Math is { } math) equations.Add(new(math, offset));
            else if (fragment.Source is DocxInlinePicture picture)
                images.Add(new(picture.Image, x + offset, top - measured.BaselineInset, fragment.Advance,
                    fragment.OccupiedHeight, pageIndex, sourceBlock, sourceParagraph, story));
            else if (fragment.Source is DocxInlineTab) text.Append('\t');
            if (sourceExtents is not null)
                sourceExtents.Add(fragment.VisibleSourceExtent is { } visible
                    ? visible with { StartOffset = offset + visible.StartOffset, EndOffset = offset + visible.EndOffset }
                    : new(fragment.SourceItemIndex, offset, offset + fragment.Advance));
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
            InlineFrame = new(measured.BaselineInset, measured.OccupiedHeight),
            InlineSourceExtents = sourceExtents is null ? [] : sourceExtents };
        return (line, images);
    }

    internal static bool TryMeasureOrderedInlineContent(DocxParagraph paragraph,
        IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph>? equations,
        IDocxTextMeasurer measurer, double fontSize, double firstWidth, double nextWidth,
        double baselineInset, double lineHeight, double scale, CancellationToken token,
        [NotNullWhen(true)] out IReadOnlyList<DocxMeasuredInlineLine>? lines,
        double? physicalTextBaselineInset = null, double defaultTabStopPoints = WordDefaultTabStopPoints,
        int? pageNumber = null, int? pageCount = null)
    {
        lines = null;
        bool hasActualFieldInputs = (pageNumber is not null || pageCount is not null) &&
            paragraph.FieldReferences.All(field => !field.UsesPlaceholder || !field.Kind.IsDynamic() ||
                DocxFieldText.CanResolve(field.Kind, pageNumber, pageCount)) &&
            paragraph.Runs.All(run => run.FieldKind is not { } kind || !kind.IsDynamic() ||
                DocxFieldText.CanResolve(kind, pageNumber, pageCount));
        if (paragraph.InlineContent is not { } items || equations is null ||
            paragraph.Equations.Count == 0 ||
            DocxMathParagraphPolicy.GetPreparationEquations(paragraph, hasActualFieldInputs).Count != paragraph.Equations.Count) return false;
        bool hasTabs = items.Any(item => item is DocxInlineTab tab && !tab.Run.EffectiveProperties.Hidden);
        // Left/default tabs depend only on the preceding advance. Aligned fields,
        // bar stops and leaders need a separate field-measurement/painting contract.
        if (hasTabs && paragraph.TabStops.Any(stop =>
            stop.Value is not null && !string.Equals(stop.Value, "left", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(stop.Value, "clear", StringComparison.OrdinalIgnoreCase) ||
            stop.LeaderValue is not null && !string.Equals(stop.LeaderValue, "none", StringComparison.OrdinalIgnoreCase))) return false;
        if (!double.IsFinite(firstWidth) || !double.IsFinite(nextWidth) || firstWidth <= 0d || nextWidth <= 0d ||
            !double.IsFinite(scale) || scale <= 0d) return false;
        var runIndices = new Dictionary<DocxTextRun, int>(ReferenceEqualityComparer.Instance);
        var tabStops = hasTabs ? ScaleTabStopPositions(paragraph.TabStops, scale) : [];
        for (int index = 0; index < paragraph.Runs.Count; index++)
        {
            token.ThrowIfCancellationRequested();
            if (!runIndices.TryAdd(paragraph.Runs[index], index)) return false;
        }
        var measured = new List<DocxMeasuredInlineFragment>();
        int sourceItemIndex = 0;
        foreach (DocxInlineItem item in items)
        {
            token.ThrowIfCancellationRequested();
            int firstFragment = measured.Count;
            switch (item)
            {
                case DocxInlineText text when text.Run.EffectiveProperties.Hidden:
                case DocxInlineTab tab when tab.Run.EffectiveProperties.Hidden:
                case DocxInlineBreak lineBreak when lineBreak.Run.EffectiveProperties.Hidden:
                    break;
                case DocxInlineText text when text.Run.FieldKind is { } kind && kind.IsDynamic():
                    // The owned placeholder is atomic. Resolve only with actual
                    // page inputs; retain the original run/item for annotations.
                    if (!runIndices.TryGetValue(text.Run, out int fieldRunIndex) ||
                        !DocxFieldText.TryResolve(text.Run, text.Run.Text.Substring(text.Start, text.Length),
                            pageNumber, pageCount, out string fieldValue)) return false;
                    var fieldSpan = CreateTextSpan(fieldValue, text.Run, fieldRunIndex) with
                    { SourceTextOffsetInRun = text.Run.SourceTextOffsetInRun + text.Start, SourceInlineItemIndex = sourceItemIndex };
                    var fieldSegments = CreateTextSegments([fieldSpan], 0d, fontSize, measurer, [], 36d);
                    double fieldWidth = fieldSegments.Count == 0 ? 0d : fieldSegments[^1].X + fieldSegments[^1].Width;
                    measured.Add(new(text, fieldWidth, baselineInset, Math.Max(lineHeight, baselineInset), [fieldSpan])
                        { TextSegments = fieldSegments });
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
                        span = span with { SourceTextOffsetInRun = span.SourceTextOffsetInRun + start,
                            SourceInlineItemIndex = sourceItemIndex };
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
                    AddMeasuredMath(item, prepared);
                    break;
                case DocxInlinePicture picture:
                    double imageWidth = picture.Image.WidthPoints * scale, imageHeight = picture.Image.HeightPoints * scale;
                    if (!double.IsFinite(imageWidth) || !double.IsFinite(imageHeight) || imageWidth <= 0d || imageHeight <= 0d) return false;
                    measured.Add(new(item, imageWidth, imageHeight, imageHeight, []));
                    break;
                case DocxInlineTab:
                case DocxInlineBreak:
                    measured.Add(new(item, 0d, baselineInset, Math.Max(lineHeight, baselineInset), []));
                    break;
                default: return false;
            }
            for (int fragmentIndex = firstFragment; paragraph.CommentRanges.Count != 0 && fragmentIndex < measured.Count; fragmentIndex++)
            {
                var fragment = measured[fragmentIndex] with { SourceItemIndex = sourceItemIndex };
                if (fragment.Math is { } math)
                {
                    double left = double.PositiveInfinity, right = double.NegativeInfinity;
                    foreach (var glyph in math.Box.Glyphs)
                    {
                        token.ThrowIfCancellationRequested();
                        left = Math.Min(left, glyph.X);
                        right = Math.Max(right, glyph.X + glyph.Advance);
                    }
                    foreach (var rule in math.Box.Rules)
                    {
                        token.ThrowIfCancellationRequested();
                        left = Math.Min(left, rule.X);
                        right = Math.Max(right, rule.X + rule.Width);
                    }
                    if (double.IsFinite(left) && double.IsFinite(right))
                        fragment = fragment with { VisibleSourceExtent = new(sourceItemIndex, left * scale, right * scale,
                            math.Box.AboveBaseline * scale, math.Box.BelowBaseline * scale, IsMath: true) };
                }
                measured[fragmentIndex] = fragment;
            }
            sourceItemIndex++;
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
            if (measured[index].Source is DocxInlineTab)
            {
                double stop = ResolveNextTabStop(advance, tabStops, defaultTabStopPoints * scale).PositionPoints;
                double available = output.Count == 0 ? firstWidth : nextWidth;
                if (stop > available && current.Count != 0)
                {
                    Flush();
                    stop = ResolveNextTabStop(0d, tabStops, defaultTabStopPoints * scale).PositionPoints;
                    available = nextWidth;
                }
                if (!double.IsFinite(stop) || stop < advance || stop > available) return false;
                current.Add(measured[index] with { Advance = stop - advance });
                advance = stop;
                index++;
                continue;
            }
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
                            { TextSegments = segments, SourceItemIndex = span.SourceInlineItemIndex };
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
            if (ReferenceEquals(left.Source, right.Source) && right.MathContinuationAdvance is { } mathGap)
                return mathGap;
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
        void AddMeasuredMath(DocxInlineItem source, DocxPreparedMathParagraph prepared)
        {
            var box = prepared.Box;
            if (prepared.UsesDisplayStyle || prepared.Equation.Display || box.InlineBreakAfterGlyphs.Count == 0)
            {
                measured.Add(new(source, box.AdvanceWidth * scale, prepared.BaselineInset * scale,
                    prepared.LineHeight * scale, [], prepared));
                return;
            }
            int start = 0;
            double previousEnd = 0d;
            foreach (int end in box.InlineBreakAfterGlyphs.Append(box.Glyphs.Count))
            {
                token.ThrowIfCancellationRequested();
                double origin = box.Glyphs[start].X;
                double endX = end == box.Glyphs.Count ? box.AdvanceWidth :
                    box.Glyphs[end - 1].X + box.Glyphs[end - 1].Advance;
                var glyphs = new DocxMathGlyphPlacement[end - start];
                double inkRight = endX - origin;
                for (int index = start; index < end; index++)
                {
                    token.ThrowIfCancellationRequested();
                    var glyph = box.Glyphs[index] with { X = box.Glyphs[index].X - origin };
                    glyphs[index - start] = glyph;
                    var font = glyph.Font.Font.Font;
                    if (font.TryGetGlyphBounds(glyph.OriginalGlyph, out var bounds))
                        inkRight = Math.Max(inkRight, glyph.X + (bounds.XMax + glyph.TextShear * bounds.YMax) * glyph.FontSize / font.UnitsPerEm);
                }
                var part = prepared with { Box = box with { Width = inkRight,
                    LogicalAdvance = endX - origin, Glyphs = glyphs, InlineBreakAfterGlyphs = [] } };
                measured.Add(new(source, part.Box.AdvanceWidth * scale, prepared.BaselineInset * scale,
                    prepared.LineHeight * scale, [], part)
                    { MathContinuationAdvance = start == 0 ? null : (origin - previousEnd) * scale });
                previousEnd = endX;
                start = end;
            }
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
            double textInset = baselineInset;
            if (physicalTextBaselineInset is { } physicalInset && !displayGroup &&
                current.Any(fragment => fragment.Math is not null) &&
                current.All(fragment => fragment.Math is { } math
                    ? !math.UsesDisplayStyle && !math.Equation.Display &&
                      math.Box.AboveBaseline * scale <= physicalInset &&
                      math.Box.BelowBaseline * scale <= lineHeight - physicalInset
                    : fragment.Source is DocxInlineText or DocxInlineTab))
            {
                // An inline object entirely inside the ordinary text strut must
                // not add a second font line frame. Taller objects keep their
                // independently qualified occupied extent.
                textInset = physicalInset;
                above = physicalInset;
                below = Math.Max(0d, lineHeight - physicalInset);
                for (int index = 0; index < current.Count; index++)
                    if (current[index].Math is { } math)
                        current[index] = current[index] with { BaselineInset = math.Box.AboveBaseline * scale,
                            OccupiedHeight = (math.Box.AboveBaseline + math.Box.BelowBaseline) * scale };
            }
            IReadOnlyList<DocxTextSegmentLayout> labelSegments = [];
            if (output.Count == 0 && paragraph.ListLabel is { } label)
            {
                var style = paragraph.Runs.FirstOrDefault() ?? paragraph.ParagraphMarkRun ??
                    new DocxTextRun(string.Empty, fontSize, null, false, false, false, null, null);
                labelSegments = CreateListLabelSegments(label, style, fontSize,
                    GetParagraphLabelStartOffset(paragraph, scale), measurer);
                above += ResolveListLabelFirstLineExtraLeading(paragraph, fontSize, measurer);
            }
            output.Add(new(current.ToArray(), advance, above, above + below)
                { IsDisplayGroup = displayGroup, ListLabelSegments = labelSegments, TextBaselineInset = textInset });
            current.Clear(); advance = 0d; above = baselineInset; below = Math.Max(0d, lineHeight - baselineInset);
        }
    }
}
