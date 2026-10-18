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
    public DocxTextAlignment? DisplayAlignment { get; init; }
    public bool IsDisplayGroup => DisplayAlignment is not null;
    public double TextBaselineInset { get; init; }
    public DocxInlineSourceCursor EndCursor { get; init; }
    public IReadOnlyList<DocxTextSegmentLayout> ListLabelSegments { get; init; } = [];
}
// A cursor commits consumed source only after its container admits the line.
// A split counter retains the value bound when its first digits were placed.
// Position Count addresses the closing paragraph mark; Count + 1 consumes it.
internal readonly record struct DocxInlineSourceCursor(int AtomIndex, int TextOffset = 0,
    string? BoundFieldValue = null, int LineIndex = 0)
{
    public bool IsComplete(DocxMeasuredInlineAtoms atoms) => AtomIndex == atoms.Fragments.Count + 1 && TextOffset == 0;
}
internal sealed record DocxMeasuredInlineAtoms(DocxParagraph Paragraph,
    IReadOnlyList<DocxMeasuredInlineFragment> Fragments, IDocxTextMeasurer Measurer,
    double FontSize, double BaselineInset, double LineHeight, double Scale, double DefaultTabStopPoints,
    IReadOnlyList<DocxTabStop> TabStops, IReadOnlyList<int> CounterFragmentIndices);
internal readonly record struct DocxInlineParagraphFrame(double OccupiedHeight, DocxTextLineLayout LastLine);

internal sealed partial class DocxLayoutEngine
{
    private static double ResolveOrderedInlineAlignment(DocxMeasuredInlineLine line,
        DocxTextAlignment alignment, double available) =>
        (line.DisplayAlignment ?? alignment) switch
        {
            DocxTextAlignment.Center => Math.Max(0d, available - line.Advance) / 2d,
            DocxTextAlignment.Right => Math.Max(0d, available - line.Advance),
            _ => 0d
        };

    private static double ResolveOrderedBodyBaselineInset(DocxParagraph paragraph, double size,
        double lineHeight, IDocxTextMeasurer measurer) =>
        !DocxPlainBodyFrame.HasFixedLineFlow(paragraph) &&
            DocxLineMetrics.ResolveUniformBodyTypographicBaselineInset(paragraph, size, measurer) is { } typographic
            ? typographic
            : DocxLineMetrics.ResolveBodyBaselineOffset(size, lineHeight, DocxPlainBodyFrame.HasFixedLineFlow(paragraph),
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
            GetParagraphFirstLineTextStartOffset(paragraph, fontSize, context.TextMeasurer, scale, context.DefaultTabStopPoints));
        double nextWidth = ResolveTableCellTextWrapWidth(cell, textWidth - right - GetParagraphTextStartOffset(paragraph, scale));
        return TryMeasureOrderedInlineContent(paragraph, context.MathEquations, context.TextMeasurer, fontSize,
            firstWidth, nextWidth, baselineInset, lineHeight, scale, context.CancellationToken, out lines,
            defaultTabStopPoints: context.DefaultTabStopPoints, pageNumber: context.PageNumber, pageCount: context.PageCount);
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
        List<DocxInlineSourceExtent>? sourceExtents = DocxInlineContent.RequiresMeasuredSourceRanges(paragraph) ? new() : null;
        var text = new System.Text.StringBuilder();
        if (measured.ListLabelSegments.Count != 0 && paragraph.ListLabel is { } label)
            text.Append(label.DisplayText).Append(label.SuffixValue switch { "nothing" => string.Empty, "space" => " ", _ => "\t" });
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
                    fragment.OccupiedHeight, pageIndex, sourceBlock, sourceParagraph, story)
                    { SourceLineIndex = sourceLine });
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
            measured.Advance, CoalesceOrderedTextSegments(segments), sourceBlock, sourceParagraph, sourceLine, story,
            LineHeight: measured.OccupiedHeight, AppliedBeforeSpacing: 0d, IsFirstParagraphLine: sourceLine == 0,
            EndsWithIntraTokenBreak: false, SingleLineHeight: null, ListLabelSingleLineHeight: null,
            BodyWindowsLineHeight: null, ListLabelWindowsLineHeight: null, EffectiveLineSpacingFactor: null,
            LineSpacingFactorFloorApplied: null, PendingAfterSpacing: null, ParagraphBeforeSpacing: null,
            ParagraphAfterSpacing: null, ContextualSpacingSuppressed: null, SourceParagraph: paragraph,
            LineHeightSource: null, EmitsTerminalParagraphMark: false)
        { FrameOriginOffsetX = frameOriginX - x, InlineMath = equations, MathScale = scale, InlineObjectEndOffsetX = measured.Advance,
            InlineFrame = new(measured.BaselineInset, measured.OccupiedHeight)
                { TextBaselineInset = measured.IsDisplayGroup ? null : measured.TextBaselineInset },
            InlineSourceExtents = sourceExtents is null ? [] : sourceExtents };
        return (line, images);
    }

    // Wrapping splits a source run into word atoms. Painting retains one glyph
    // run across adjacent atoms with the same source owner and text state.
    private static IReadOnlyList<DocxTextSegmentLayout> CoalesceOrderedTextSegments(IReadOnlyList<DocxTextSegmentLayout> segments)
    {
        if (segments.Count < 2) { return segments; }
        var result = new List<DocxTextSegmentLayout>(segments.Count);
        for (int index = 0; index < segments.Count;)
        {
            var first = segments[index++];
            double width = first.Width;
            int length = first.Text.Length;
            System.Text.StringBuilder? text = null;
            while (index < segments.Count)
            {
                var next = segments[index];
                if (first.Role != DocxTextSegmentRole.Text || next.Role != first.Role || first.SourceTextRunIndex < 0 ||
                    next.SourceTextRunIndex != first.SourceTextRunIndex ||
                    next.SourceTextOffsetInRun != first.SourceTextOffsetInRun + length ||
                    !ReferenceEquals(first.StyleRun, next.StyleRun) || !ReferenceEquals(first.FallbackFace, next.FallbackFace) ||
                    first.FontSize != next.FontSize || first.BaselineOffsetY != next.BaselineOffsetY ||
                    first.PdfCharacterSpacing != next.PdfCharacterSpacing || first.PdfCharacterSpacingSource != next.PdfCharacterSpacingSource ||
                    first.CompensatePdfCharacterSpacing != next.CompensatePdfCharacterSpacing ||
                    Math.Abs(first.X + width - next.X) > .000001d) { break; }
                (text ??= new System.Text.StringBuilder(first.Text)).Append(next.Text);
                length += next.Text.Length;
                width += next.Width;
                index++;
            }
            result.Add(text is null ? first : first with { Text = text.ToString(), Width = width });
        }
        return result;
    }

    private static bool HaveSameInlineMeasurementState(DocxTextRun before, DocxTextRun after) =>
        before.FieldKind is null && after.FieldKind is null &&
        !before.IsAutomaticNoteReferenceMarker && !after.IsAutomaticNoteReferenceMarker &&
        !before.IsAutomaticNoteStoryMarker && !after.IsAutomaticNoteStoryMarker &&
        before.ScriptBaseFontSize == after.ScriptBaseFontSize &&
        before.UseLogicalSizeForPdfAdvances == after.UseLogicalSizeForPdfAdvances &&
        before.EffectiveProperties == after.EffectiveProperties;

    internal static bool TryMeasureOrderedInlineContent(DocxParagraph paragraph,
        IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph>? equations,
        IDocxTextMeasurer measurer, double fontSize, double firstWidth, double nextWidth,
        double baselineInset, double lineHeight, double scale, CancellationToken token,
        [NotNullWhen(true)] out IReadOnlyList<DocxMeasuredInlineLine>? lines,
        double? physicalTextBaselineInset = null, double defaultTabStopPoints = WordDefaultTabStopPoints,
        int? pageNumber = null, int? pageCount = null, bool allowStaticFallback = false)
    {
        lines = null;
        return double.IsFinite(firstWidth) && double.IsFinite(nextWidth) && firstWidth > 0d && nextWidth > 0d &&
            TryMeasureOrderedInlineAtoms(paragraph, equations, measurer, fontSize, baselineInset, lineHeight, scale,
                token, out var atoms, defaultTabStopPoints, pageNumber, pageCount, allowStaticFallback) &&
            TryWrapOrderedInlineAtoms(atoms, firstWidth, nextWidth, token, out lines, physicalTextBaselineInset);
    }

    internal static bool TryMeasureOrderedInlineAtoms(DocxParagraph paragraph,
        IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph>? equations,
        IDocxTextMeasurer measurer, double fontSize, double baselineInset, double lineHeight, double scale,
        CancellationToken token, [NotNullWhen(true)] out DocxMeasuredInlineAtoms? atoms,
        double defaultTabStopPoints = WordDefaultTabStopPoints, int? pageNumber = null, int? pageCount = null,
        bool allowStaticFallback = false)
    {
        atoms = null;
        bool hasActualFieldInputs = (pageNumber is not null || pageCount is not null) &&
            paragraph.FieldReferences.All(field => !field.UsesPlaceholder || !field.Kind.IsDynamic() ||
                DocxFieldText.CanResolve(field.Kind, pageNumber, pageCount)) &&
            paragraph.Runs.All(run => run.FieldKind is not { } kind || !kind.IsDynamic() ||
                DocxFieldText.CanResolve(kind, pageNumber, pageCount));
        var items = paragraph.InlineContent;
        bool retainPreparedContent = allowStaticFallback;
        if (paragraph.Equations.Count != 0)
        {
            if (items is null || equations is null ||
                DocxMathParagraphPolicy.GetPreparationEquations(paragraph, hasActualFieldInputs, token).Count != paragraph.Equations.Count) return false;
            if (!retainPreparedContent)
                foreach (var equation in paragraph.Equations)
                {
                    token.ThrowIfCancellationRequested();
                    if (equations.ContainsKey(equation)) { retainPreparedContent = true; break; }
                }
        }
        else
        {
            if (!allowStaticFallback || paragraph.Images.Count != 0 ||
                !DocxInlineContent.HasSupportedSourceOwner(paragraph, hasActualFieldInputs, token) ||
                !DocxPlainBodyFrame.HasAutomaticLineFlow(paragraph) ||
                paragraph.Alignment is not (DocxTextAlignment.Left or DocxTextAlignment.Center or DocxTextAlignment.Right)) return false;
            // Ordinary text paragraphs need no parser-side mixed-content view.
            // Reuse the same source-run affinity without mutating their model.
            items ??= DocxInlineContent.FromPrefix(paragraph.Runs, paragraph.Images, token);
        }
        bool hasTabs = items.Any(item => item is DocxInlineTab tab && !tab.Run.EffectiveProperties.Hidden);
        // Bars can draw without consuming a tab, and unknown kinds have no
        // placement contract. Positioning stops are checked when consumed.
        if (hasTabs && paragraph.TabStops.Any(stop =>
            stop.Value is not null && !string.Equals(stop.Value, "left", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(stop.Value, "right", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(stop.Value, "center", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(stop.Value, "decimal", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(stop.Value, "clear", StringComparison.OrdinalIgnoreCase))) return false;
        if (!double.IsFinite(scale) || scale <= 0d) return false;
        if (DocxPlainBodyFrame.HasFixedLineFlow(paragraph))
        {
            // Containers share the authored physical slot; measured ink may
            // overhang it without changing advance or the text baseline.
            lineHeight = paragraph.LineSpacingPoints!.Value * scale;
            baselineInset = DocxLineMetrics.ResolveBodyBaselineOffset(fontSize, lineHeight, hasExplicitLineSpacing: true);
        }
        if (paragraph.LineSpacingPoints is not null && DocxPlainBodyFrame.HasContentDrivenLineFlow(paragraph))
        {
            // Measure intrinsic fragments first; the completed row owns the
            // paragraph minimum, including any empty leading before its baseline.
            lineHeight = Math.Min(lineHeight, ResolveLineHeightProfile(paragraph with { LineSpacingPoints = 0d },
                fontSize, measurer).LineHeight);
        }
        var runIndices = new Dictionary<DocxTextRun, int>(ReferenceEqualityComparer.Instance);
        for (int index = 0; index < paragraph.Runs.Count; index++)
        {
            token.ThrowIfCancellationRequested();
            if (!runIndices.TryAdd(paragraph.Runs[index], index)) return false;
        }
        var measured = new List<DocxMeasuredInlineFragment>();
        bool sourceRanges = DocxInlineContent.RequiresMeasuredSourceRanges(paragraph);
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
                    if (equations is null || !equations.TryGetValue(math.Equation, out var prepared))
                    {
                        // A trusted ordered paragraph keeps its prepared items
                        // when a sibling cannot be rendered. The existing
                        // unsupported diagnostic still belongs to that item;
                        // entirely unsupported body content keeps its fallback.
                        if (retainPreparedContent) break;
                        return false;
                    }
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
            for (int fragmentIndex = firstFragment; sourceRanges && fragmentIndex < measured.Count; fragmentIndex++)
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

        atoms = new(paragraph, measured, measurer, fontSize, baselineInset, lineHeight, scale, defaultTabStopPoints,
            hasTabs ? ScaleTabStopPositions(paragraph.TabStops, scale) : [],
            measured.Select((fragment, index) => (fragment, index))
                .Where(item => item.fragment.Source is DocxInlineText { Run.FieldKind: { } kind } && kind.IsDynamic())
                .Select(item => item.index).ToArray());
        return true;

        void AddMeasuredMath(DocxInlineItem source, DocxPreparedMathParagraph prepared)
        {
            var box = prepared.Box;
            if (prepared.UsesDisplayStyle || box.InlineBreakAfterGlyphs.Count == 0)
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
    }

    internal static bool TryRebindInlineCounters(DocxMeasuredInlineAtoms atoms, int? pageNumber, int? pageCount,
        CancellationToken token, [NotNullWhen(true)] out DocxMeasuredInlineAtoms? rebound)
    {
        rebound = null;
        token.ThrowIfCancellationRequested();
        DocxMeasuredInlineFragment[]? changed = null;
        foreach (int index in atoms.CounterFragmentIndices)
        {
            token.ThrowIfCancellationRequested();
            var fragment = atoms.Fragments[index];
            if (fragment.Source is not DocxInlineText { Run.FieldKind: { } kind } source || !kind.IsDynamic()) continue;
            if (fragment.Text.Count != 1 || !DocxFieldText.TryResolve(source.Run,
                source.Run.Text.Substring(source.Start, source.Length), pageNumber, pageCount, out string value)) return false;
            var span = fragment.Text[0];
            if (span.Text == value) continue;
            span = span with { Text = value };
            var segments = CreateTextSegments([span], 0d, atoms.FontSize, atoms.Measurer, [], 36d);
            double width = segments.Count == 0 ? 0d : segments[^1].X + segments[^1].Width;
            if (!double.IsFinite(width) || width < 0d) return false;
            changed ??= atoms.Fragments.ToArray();
            changed[index] = fragment with { Advance = width, Text = [span], TextSegments = segments };
        }
        rebound = changed is null ? atoms : atoms with { Fragments = changed };
        return true;
    }

    internal static bool TryWrapOrderedInlineAtoms(DocxMeasuredInlineAtoms atoms, double firstWidth, double nextWidth,
        CancellationToken token, [NotNullWhen(true)] out IReadOnlyList<DocxMeasuredInlineLine>? lines,
        double? physicalTextBaselineInset = null)
        => TryWrapOrderedInlineAtomsCore(atoms, firstWidth, nextWidth, token, out lines, physicalTextBaselineInset,
            default, stopAfterFirstLine: false);

    internal static bool TryGetNextOrderedInlineLine(DocxMeasuredInlineAtoms atoms, DocxInlineSourceCursor cursor,
        double firstWidth, double nextWidth, CancellationToken token,
        [NotNullWhen(true)] out DocxMeasuredInlineLine? line, double? physicalTextBaselineInset = null)
    {
        line = null;
        if (cursor.IsComplete(atoms)) return false;
        if (!TryWrapOrderedInlineAtomsCore(atoms, firstWidth, nextWidth, token, out var lines,
            physicalTextBaselineInset, cursor, stopAfterFirstLine: true)) return false;
        line = lines[0];
        return true;
    }

    private static bool TryWrapOrderedInlineAtomsCore(DocxMeasuredInlineAtoms atoms, double firstWidth, double nextWidth,
        CancellationToken token, [NotNullWhen(true)] out IReadOnlyList<DocxMeasuredInlineLine>? lines,
        double? physicalTextBaselineInset, DocxInlineSourceCursor cursor, bool stopAfterFirstLine)
    {
        lines = null;
        if (!double.IsFinite(firstWidth) || !double.IsFinite(nextWidth) || firstWidth <= 0d || nextWidth <= 0d) return false;
        var paragraph = atoms.Paragraph;
        var measured = atoms.Fragments;
        var measurer = atoms.Measurer;
        double fontSize = atoms.FontSize, baselineInset = atoms.BaselineInset, lineHeight = atoms.LineHeight,
            scale = atoms.Scale, defaultTabStopPoints = atoms.DefaultTabStopPoints;
        var tabStops = atoms.TabStops;
        if (cursor.AtomIndex < 0 || cursor.AtomIndex > measured.Count + 1 || cursor.TextOffset < 0 || cursor.LineIndex < 0 ||
            cursor.AtomIndex >= measured.Count && cursor.TextOffset != 0) return false;
        DocxMeasuredInlineFragment? partial = null;
        if (cursor.TextOffset != 0)
        {
            var original = measured[cursor.AtomIndex];
            if (original.Source is not DocxInlineText source || original.Text.Count != 1) return false;
            bool field = source.Run.FieldKind is { } kind && kind.IsDynamic();
            string fullValue = field ? cursor.BoundFieldValue ?? original.Text[0].Text : original.Text[0].Text;
            if (cursor.TextOffset >= fullValue.Length || !field && cursor.BoundFieldValue is not null) return false;
            var span = original.Text[0] with { Text = fullValue[cursor.TextOffset..],
                SourceTextOffsetInRun = original.Text[0].SourceTextOffsetInRun + cursor.TextOffset };
            var segments = CreateTextSegments([span], 0d, fontSize, measurer, [], 36d);
            double partWidth = segments.Count == 0 ? 0d : segments[^1].X + segments[^1].Width;
            if (!double.IsFinite(partWidth) || partWidth < 0d) return false;
            partial = original with
            {
                Source = field ? source : source with { Start = source.Start + cursor.TextOffset, Length = source.Length - cursor.TextOffset },
                Advance = partWidth, Text = [span], TextSegments = segments
            };
        }
        DocxMeasuredInlineFragment Atom(int index) => index == cursor.AtomIndex && partial is not null ? partial : measured[index];
        var output = new List<DocxMeasuredInlineLine>();
        var current = new List<DocxMeasuredInlineFragment>();
        double advance = 0d, above = baselineInset, below = Math.Max(0d, lineHeight - baselineInset);
        bool authoredLineStart = cursor.AtomIndex == 0 && cursor.LineIndex == 0 ||
            cursor.AtomIndex > 0 && measured[cursor.AtomIndex - 1].Source is DocxInlineBreak;
        for (int index = cursor.AtomIndex; index < measured.Count;)
        {
            token.ThrowIfCancellationRequested();
            if (Atom(index).Source is DocxInlineBreak)
            {
                Flush(new(index + 1));
                authoredLineStart = true;
                if (stopAfterFirstLine) { lines = output; return true; }
                index++; continue;
            }
            if (Atom(index).Source is DocxInlineTab)
            {
                var tabStop = ResolveNextTabStop(advance, tabStops, defaultTabStopPoints * scale);
                if (!CanMeasureOrderedTabStop(tabStop)) return false;
                double stop = tabStop.PositionPoints;
                double available = output.Count == 0 ? firstWidth : nextWidth;
                if (string.Equals(tabStop.Value, "right", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(tabStop.Value, "center", StringComparison.OrdinalIgnoreCase))
                {
                    double fieldWidth = 0d;
                    for (int field = index + 1; field < measured.Count && Atom(field).Source is not (DocxInlineTab or DocxInlineBreak); field++)
                    {
                        token.ThrowIfCancellationRequested();
                        if (field > index + 1) fieldWidth += InlineBoundaryAdvance(Atom(field - 1), Atom(field));
                        fieldWidth += Atom(field).Advance;
                    }
                    double start = GetAlignedTabFieldStart(tabStop, fieldWidth, null);
                    // The complete aligned field must fit without overlap. Its
                    // wrapping/leader/decimal variants remain staged fallbacks.
                    if (!double.IsFinite(start) || start < advance || start + fieldWidth > available) return false;
                    current.Add(Atom(index) with { Advance = start - advance });
                    advance = start;
                    index++;
                    continue;
                }
                if (stop > available && current.Count != 0)
                {
                    Flush(new(index));
                    if (stopAfterFirstLine) { lines = output; return true; }
                    tabStop = ResolveNextTabStop(0d, tabStops, defaultTabStopPoints * scale);
                    if (!CanMeasureOrderedTabStop(tabStop)) return false;
                    stop = tabStop.PositionPoints;
                    available = nextWidth;
                }
                if (!double.IsFinite(stop) || stop < advance || stop > available) return false;
                current.Add(Atom(index) with { Advance = stop - advance });
                advance = stop;
                index++;
                continue;
            }
            int end = index + 1;
            bool space = IsSpace(Atom(index));
            // A formatting/run boundary is not a word-break opportunity.
            if (!space && Atom(index).Source is DocxInlineText)
                while (end < measured.Count && Atom(end).Source is DocxInlineText && !IsSpace(Atom(end))) end++;
            double groupWidth = 0d;
            for (int offset = index; offset < end; offset++)
            {
                if (offset > index) groupWidth += InlineBoundaryAdvance(measured[offset - 1], Atom(offset));
                groupWidth += Atom(offset).Advance;
            }
            double width = output.Count == 0 ? firstWidth : nextWidth;
            if (groupWidth > Math.Max(firstWidth, nextWidth))
            {
                if (Atom(index).Source is not DocxInlineText) return false;
                if (current.Count != 0)
                {
                    Flush(new(index));
                    if (stopAfterFirstLine) { lines = output; return true; }
                }
                var spans = Enumerable.Range(index, end - index).SelectMany(atom => Atom(atom).Text).ToArray();
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
                        DocxInlineText source;
                        if (span.StyleRun.FieldKind is { } fieldKind && fieldKind.IsDynamic())
                        {
                            // A resolved counter may have more digits than its
                            // canonical placeholder. Its source stays canonical;
                            // the cursor carries a partially consumed value.
                            source = measured.Skip(index).Take(end - index)
                                .Select(fragment => fragment.Source).OfType<DocxInlineText>()
                                .First(item => ReferenceEquals(item.Run, span.StyleRun));
                        }
                        else
                        {
                            int offset = span.SourceTextOffsetInRun - Math.Max(0, span.StyleRun.SourceTextOffsetInRun);
                            if (offset < 0 || offset > span.StyleRun.Text.Length) return false;
                            source = new(span.StyleRun, offset, Math.Min(span.Text.Length, span.StyleRun.Text.Length - offset));
                        }
                        var segments = CreateTextSegments([span], 0d, fontSize, measurer, [], 36d);
                        double partWidth = segments.Count == 0 ? 0d : segments[^1].X + segments[^1].Width;
                        if (!double.IsFinite(partWidth) || partWidth < 0d) return false;
                        var part = new DocxMeasuredInlineFragment(source, partWidth, baselineInset, Math.Max(lineHeight, baselineInset), [span])
                            { TextSegments = segments, SourceItemIndex = span.SourceInlineItemIndex };
                        part = part with { LeadingAdvance = current.Count == 0 ? 0d : InlineBoundaryAdvance(current[^1], part) };
                        current.Add(part);
                        advance += part.LeadingAdvance + partWidth;
                    }
                    if (lineIndex + 1 < wrapped.Length)
                    {
                        var endCursor = wrapped[lineIndex].Spans.Count == 0 ? new DocxInlineSourceCursor(end) :
                            CursorAfterSpan(wrapped[lineIndex].Spans[^1], index, end);
                        if (endCursor is null) return false;
                        Flush(endCursor.Value);
                        if (stopAfterFirstLine) { lines = output; return true; }
                    }
                }
                index = end;
                continue;
            }
            double boundary = current.Count == 0 ? 0d : InlineBoundaryAdvance(current[^1], Atom(index));
            if (current.Count != 0 && advance + boundary + groupWidth > width)
            {
                if (Atom(index).Math is { Box.AllowsInlineRelocation: false }) return false;
                Flush(new(index));
                if (stopAfterFirstLine) { lines = output; return true; }
            }
            if (current.Count == 0 && groupWidth > (output.Count == 0 ? firstWidth : nextWidth))
            {
                if (output.Count == 0 && groupWidth <= nextWidth)
                {
                    if (Atom(index).Math is { Box.AllowsInlineRelocation: false }) return false;
                    Flush(new(index));
                    if (stopAfterFirstLine) { lines = output; return true; }
                }
                else return false;
            }
            if (space && current.Count == 0 && !authoredLineStart) { index = end; continue; }
            for (; index < end; index++)
            {
                var fragment = Atom(index) with
                    { LeadingAdvance = current.Count == 0 ? 0d : InlineBoundaryAdvance(current[^1], Atom(index)) };
                current.Add(fragment); advance += fragment.LeadingAdvance + fragment.Advance;
                above = Math.Max(above, fragment.BaselineInset);
                below = Math.Max(below, fragment.OccupiedHeight - fragment.BaselineInset);
            }
        }
        if (current.Count != 0 || output.Count == 0 || authoredLineStart) Flush(new(measured.Count + 1));
        lines = output; return true;

        DocxInlineSourceCursor? CursorAfterSpan(DocxTextSpan span, int begin, int end)
        {
            for (int index = begin; index < end; index++)
            {
                var original = measured[index];
                if (original.Source is not DocxInlineText source || original.Text.Count != 1 ||
                    !ReferenceEquals(source.Run, span.StyleRun)) continue;
                bool field = source.Run.FieldKind is { } kind && kind.IsDynamic();
                string value = field && index == cursor.AtomIndex && cursor.BoundFieldValue is { } bound
                    ? bound : original.Text[0].Text;
                int offset = span.SourceTextOffsetInRun - original.Text[0].SourceTextOffsetInRun;
                if (offset < 0 || offset >= value.Length) continue;
                int consumed = offset + span.Text.Length;
                return consumed >= value.Length ? new(index + 1) : new(index, consumed, field ? value : null);
            }
            return null;
        }

        static bool IsSpace(DocxMeasuredInlineFragment fragment) => fragment.Source is DocxInlineText { Length: > 0 } text &&
            DocxTextBreakRules.IsBreakableWhitespaceChar(text.Run.Text[text.Start]);
        double InlineBoundaryAdvance(DocxMeasuredInlineFragment left, DocxMeasuredInlineFragment right)
        {
            if (ReferenceEquals(left.Source, right.Source) && right.MathContinuationAdvance is { } mathGap)
                return mathGap;
            // Word merges adjacent math zones. Reuse the same atom glue as a
            // structured equation row so independent fractions keep their bars.
            if (left.Math is { } beforeMath && right.Math is { } afterMath &&
                !ReferenceEquals(beforeMath.Equation, afterMath.Equation) &&
                beforeMath.Box.Edges is { } beforeEdges && afterMath.Box.Edges is { FirstNormalText: false } afterEdges)
            {
                double spacingSize = afterMath.Equation.ParentFontSize ?? afterMath.Box.ParentMetrics?.Size ?? fontSize;
                return DocxMathRowLayout.Glue(beforeEdges.Last, afterEdges.First) * spacingSize * scale / 18d;
            }
            if (left.Text.Count == 0 || right.Text.Count == 0) return 0d;
            var before = left.Text[^1];
            var after = right.Text[0];
            if (!ReferenceEquals(before.StyleRun, after.StyleRun) &&
                !HaveSameInlineMeasurementState(before.StyleRun, after.StyleRun))
                return DocxTextSpacing.BoundarySpacing(before.StyleRun, before.Text, after.Text);
            // Source-run boundaries with unchanged measurement state do not
            // reset pair kerning. Keep their separate annotation/source owners.
            // Decode complete Unicode scalars, including
            // replacement scalars for malformed input, as the text measurer does.
            System.Text.Rune.DecodeLastFromUtf16(before.Text, out var last, out _);
            System.Text.Rune.DecodeFromUtf16(after.Text, out var first, out _);
            string lastText = last.ToString(), firstText = first.ToString();
            double Measure(string value) => MeasureTextSpansForLayout([before with { Text = value }],
                fontSize, measurer, [], 36d, null);
            return Measure(lastText + firstText) - Measure(lastText) - Measure(firstText);
        }
        void Flush(DocxInlineSourceCursor endCursor)
        {
            // Only soft wrapping discards the next line's leading spaces.
            // Authored starts retain their spaces and the closing paragraph
            // mark owns a final text strut after a trailing hard break.
            if (endCursor.TextOffset == 0 && endCursor.AtomIndex <= measured.Count &&
                (endCursor.AtomIndex == 0 || measured[endCursor.AtomIndex - 1].Source is not DocxInlineBreak))
            {
                while (endCursor.AtomIndex < measured.Count && IsSpace(measured[endCursor.AtomIndex]))
                    endCursor = endCursor with { AtomIndex = endCursor.AtomIndex + 1 };
                if (endCursor.AtomIndex == measured.Count)
                    endCursor = endCursor with { AtomIndex = measured.Count + 1 };
            }
            authoredLineStart = false;
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
                    ? !math.UsesDisplayStyle &&
                      math.Box.AboveBaseline * scale <= physicalInset &&
                      math.Box.BelowBaseline * scale <= lineHeight - physicalInset
                    : fragment.Source is DocxInlineText or DocxInlineTab))
            {
                // Contained ink shares the ordinary physical text strut, but
                // its prepared occupied frame still contributes leading once.
                textInset = physicalInset;
                above = physicalInset;
                below = Math.Max(0d, lineHeight - physicalInset);
                foreach (var fragment in current)
                {
                    token.ThrowIfCancellationRequested();
                    if (fragment.Math is null) { continue; }
                    above = Math.Max(above, fragment.BaselineInset);
                    below = Math.Max(below, fragment.OccupiedHeight - fragment.BaselineInset);
                }
            }
            IReadOnlyList<DocxTextSegmentLayout> labelSegments = [];
            if (cursor.LineIndex == 0 && output.Count == 0 && paragraph.ListLabel is { } label)
            {
                var style = paragraph.Runs.FirstOrDefault() ?? paragraph.ParagraphMarkRun ??
                    new DocxTextRun(string.Empty, fontSize, null, false, false, false, null, null);
                labelSegments = CreateListLabelSegments(label, style, fontSize,
                    GetParagraphLabelStartOffset(paragraph, scale), measurer,
                    GetParagraphFirstLineTextStartOffset(paragraph, fontSize, measurer, scale, defaultTabStopPoints));
                above += ResolveListLabelFirstLineExtraLeading(paragraph, fontSize, measurer);
            }
            if (paragraph.LineSpacingPoints is { } minimum && DocxPlainBodyFrame.HasContentDrivenLineFlow(paragraph))
                above += Math.Max(0d, minimum * scale - above - below);
            if (DocxPlainBodyFrame.HasFixedLineFlow(paragraph))
            {
                above = baselineInset;
                below = lineHeight - baselineInset;
                textInset = baselineInset;
            }
            output.Add(new(current.ToArray(), advance, above, above + below)
                { DisplayAlignment = displayGroup ? DocxMathParagraphPolicy.ResolveDisplayAlignment(
                    current.Where(fragment => fragment.Math is not null).Select(fragment => fragment.Math!.Equation), token) : null,
                    ListLabelSegments = labelSegments, TextBaselineInset = textInset,
                    EndCursor = endCursor with { LineIndex = cursor.LineIndex + output.Count + 1 } });
            current.Clear(); advance = 0d; above = baselineInset; below = Math.Max(0d, lineHeight - baselineInset);
        }
    }

    private static bool CanMeasureOrderedTabStop(DocxResolvedTabStop stop) =>
        !string.Equals(stop.Value, "decimal", StringComparison.OrdinalIgnoreCase) &&
        (stop.LeaderValue is null || string.Equals(stop.LeaderValue, "none", StringComparison.OrdinalIgnoreCase));
}
