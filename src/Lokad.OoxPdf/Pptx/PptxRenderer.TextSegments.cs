using System.Globalization;
using System.Text;
using System.Xml.Linq;

using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

// font collection and painting computed identical text layouts per shape
// node. Keyed by node identity plus the traversal context (color map, placeholder
// mode) that feeds the layout; nodes compare by reference because record value
// equality would deep-compare whole subtrees per lookup. Color maps ride the render
// context, so the same instances flow through both passes and reference equality is
// exact.
internal readonly record struct PptxTextSpanMemoKey(PptxSceneNode Node, PptxColorMap ColorMap, bool IncludePlaceholders);

internal sealed class PptxTextSpanMemoKeyComparer : IEqualityComparer<PptxTextSpanMemoKey>
{
    public static readonly PptxTextSpanMemoKeyComparer Instance = new();

    public bool Equals(PptxTextSpanMemoKey left, PptxTextSpanMemoKey right)
    {
        return ReferenceEquals(left.Node, right.Node) &&
            ReferenceEquals(left.ColorMap, right.ColorMap) &&
            left.IncludePlaceholders == right.IncludePlaceholders;
    }

    public int GetHashCode(PptxTextSpanMemoKey key)
    {
        return HashCode.Combine(
            System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(key.Node),
            System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(key.ColorMap),
            key.IncludePlaceholders);
    }
}

internal sealed partial class PptxRenderer
{
    private static (PptxTextColumnBreakMode Mode, int Target, int StartColumn) ResolveOverflowColumnBalance(PptxTextFrameLayout layout)
    {
        PptxTextFrameModel frame = layout.Model;
        if (frame.ColumnCount <= 1 || frame.BodyProperties.VerticalOverflow != PptxTextVerticalOverflow.Overflow)
        {
            return (PptxTextColumnBreakMode.StrictFit, 0, 0);
        }

        var columns = layout.Paragraphs
            .SelectMany((paragraph, paragraphIndex) => paragraph.Lines.Select(line => (line.StartX, ParagraphIndex: paragraphIndex)))
            .GroupBy(item => Math.Round(item.StartX, 2))
            .OrderBy(group => group.Key)
            .ToArray();
        if (columns.Length != frame.ColumnCount)
        {
            return (PptxTextColumnBreakMode.StrictFit, 0, 0);
        }

        int[] counts = columns.Select(column => column.Count()).ToArray();
        int lineCount = counts.Sum();
        int balancedFloor = lineCount / frame.ColumnCount;
        int balancedCeiling = (int)Math.Ceiling((double)lineCount / frame.ColumnCount);
        if (counts.Take(frame.ColumnCount - 1).All(count => count < balancedFloor) &&
            counts[^1] >= balancedCeiling + frame.ColumnCount)
        {
            return (PptxTextColumnBreakMode.LineCountBalance, balancedCeiling, 0);
        }

        if (frame.ColumnCount >= 3 && counts[0] >= balancedFloor &&
            counts.Skip(1).Take(frame.ColumnCount - 2).Any(count => count < balancedFloor) &&
            counts[^1] > balancedFloor)
        {
            return (PptxTextColumnBreakMode.LineCountBalance, balancedFloor, 1);
        }

        bool unevenCounts = lineCount % frame.ColumnCount == frame.ColumnCount - 1 &&
            counts[^1] == balancedCeiling &&
            lineCount - counts[^1] == balancedCeiling * (frame.ColumnCount - 1) - 1 &&
            counts[^2] == balancedCeiling - 1;
        bool evenCounts = frame.ColumnCount >= 3 && counts.All(count => count == balancedFloor);
        if ((unevenCounts || evenCounts) &&
            columns[^2].Last().ParagraphIndex == columns[^1].First().ParagraphIndex)
        {
            return unevenCounts
                ? (PptxTextColumnBreakMode.LineCountBalance, balancedCeiling, 0)
                : (PptxTextColumnBreakMode.LineCountBalance, balancedFloor + 1, 1);
        }

        bool overflowBalance = lineCount % frame.ColumnCount == 0 &&
            counts.Take(frame.ColumnCount - 1).All(count => count == balancedFloor - 1) &&
            counts[^1] == balancedFloor + frame.ColumnCount - 1;
        return (overflowBalance ? PptxTextColumnBreakMode.OverflowBalance : PptxTextColumnBreakMode.StrictFit, 0, 0);
    }

    private static int CountLeadingSpaces(string text)
    {
        int count = 0;
        while (count < text.Length && text[count] == ' ')
        {
            count++;
        }

        return count;
    }

    private static int CountDrawableTextSegments(PptxTextFlowParagraph paragraph)
    {
        return paragraph.Runs
            .SelectMany(run => run.Segments)
            .Count(segment => segment.Draw && segment.AdvanceText.TrimStart().Length > 0);
    }

    private static bool UsesRotatedFrameAutoFit(PptxTextOrientation orientation)
    {
        return orientation is PptxTextOrientation.Vertical or PptxTextOrientation.Vertical270;
    }

    private static bool ClipsTextVerticalOverflow(PptxTextVerticalOverflow overflow)
    {
        return overflow is PptxTextVerticalOverflow.Clip or PptxTextVerticalOverflow.Ellipsis;
    }

    private static bool IsShortWordSegment(string text)
    {
        string trimmed = text.Trim();
        return trimmed.Length is > 0 and <= 16 && !trimmed.Any(char.IsWhiteSpace);
    }

    private static double ResolveLineFontSize(double visibleMaxFontSize, double fallbackFontSize)
    {
        return visibleMaxFontSize > PptxTextMetricRules.TextStateTolerance
            ? visibleMaxFontSize
            : fallbackFontSize;
    }

    private static string[] SplitTextIntoFittingChunks(
        string text,
        double maxWidth,
        double fontSize,
        ResolvedRunTextStyle style,
        TextAdvanceEstimator advanceEstimator)
    {
        var chunks = new List<string>();
        var chunk = new StringBuilder();
        foreach (Rune rune in text.EnumerateRunes())
        {
            string candidate = chunk.ToString() + rune;
            double candidateWidth = advanceEstimator.Measure(candidate, fontSize, style.Typeface, style.Bold, style.Italic, style.CharacterSpacing, style.KerningEnabled);
            if (chunk.Length > 0 && candidateWidth > maxWidth)
            {
                chunks.Add(chunk.ToString());
                chunk.Clear();
            }

            chunk.Append(rune);
        }

        if (chunk.Length > 0)
        {
            chunks.Add(chunk.ToString());
        }

        return chunks.Count == 0 ? [text] : chunks.ToArray();
    }

    private enum PptxTextColumnBreakMode
    {
        StrictFit,
        OverflowBalance,
        LineCountBalance
    }

    private static void MoveToNextColumnIfNeeded(
        ref double cursorLineTop,
        ref int columnIndex,
        ref double columnStartX,
        ref int linesInCurrentColumn,
        double firstColumnTop,
        double frameTextX,
        double columnWidth,
        double columnSpacing,
        int columnCount,
        PptxTextFlowBox box,
        PptxTextVerticalOverflow verticalOverflow,
        PptxTextColumnBreakMode columnBreakMode,
        double lineAdvance,
        int lineBalanceTarget,
        int lineBalanceStartColumn,
        bool linePlaced)
    {
        if (columnCount <= 1 || columnIndex >= columnCount - 1)
        {
            return;
        }

        if (columnBreakMode == PptxTextColumnBreakMode.LineCountBalance &&
            columnIndex >= lineBalanceStartColumn)
        {
            if (linePlaced)
            {
                linesInCurrentColumn++;
            }

            if (lineBalanceTarget <= 0 || linesInCurrentColumn < lineBalanceTarget)
            {
                return;
            }

            columnIndex++;
            columnStartX = frameTextX + columnIndex * (columnWidth + columnSpacing);
            cursorLineTop = firstColumnTop;
            linesInCurrentColumn = 0;
            return;
        }

        double bottom = box.CursorTop - box.TextHeight;
        double nextLineThreshold = verticalOverflow == PptxTextVerticalOverflow.Overflow &&
            columnBreakMode == PptxTextColumnBreakMode.OverflowBalance
            ? cursorLineTop
            : cursorLineTop - lineAdvance;
        if (nextLineThreshold >= bottom - PptxTextMetricRules.TextStateTolerance)
        {
            return;
        }

        columnIndex++;
        columnStartX = frameTextX + columnIndex * (columnWidth + columnSpacing);
        cursorLineTop = firstColumnTop;
        linesInCurrentColumn = 0;
    }

    private static GroupTransform ReadAncestorGroupTransform(XElement shape)
    {
        GroupTransform transform = GroupTransform.Identity;
        foreach (XElement group in shape.Ancestors(PresentationNamespace + "grpSp").Reverse())
        {
            transform = transform.Combine(ReadGroupTransform(group));
        }

        return transform;
    }

    private static async ValueTask<IReadOnlyList<PptxPositionedTextSpan>> ReadTextSpansForSceneNodeAsync(PptxSceneNode node, PptxRenderContext context, PptxColorMap colorMap, bool includePlaceholders)
    {
        if (node.TextBody is null)
        {
            return [];
        }

        // font collection and painting computed this full layout per node.
        // The layout is a pure function of (node, color map, placeholder mode) within
        // a slide render, so compute once and share the instance between both passes.
        Dictionary<PptxTextSpanMemoKey, IReadOnlyList<PptxPositionedTextSpan>>? memo = context.TextSpanMemo;
        if (memo is null)
        {
            return (await ComputeTextSpansForSceneNodeAsync(node, context, colorMap, includePlaceholders).ConfigureAwait(false));
        }

        var key = new PptxTextSpanMemoKey(node, colorMap, includePlaceholders);
        if (memo.TryGetValue(key, out IReadOnlyList<PptxPositionedTextSpan>? cached))
        {
            return cached;
        }

        IReadOnlyList<PptxPositionedTextSpan> spans = (await ComputeTextSpansForSceneNodeAsync(node, context, colorMap, includePlaceholders).ConfigureAwait(false));
        memo[key] = spans;
        return spans;
    }

    private static async ValueTask<IReadOnlyList<PptxPositionedTextSpan>> ComputeTextSpansForSceneNodeAsync(PptxSceneNode node, PptxRenderContext context, PptxColorMap colorMap, bool includePlaceholders)
    {
        return (await BuildSceneFedTextSpansAsync(node, context.Document, context.Theme, colorMap, context.DisplayedSlideNumber, includePlaceholders, context.InheritedXml, context.FontResolver, context.CancellationToken).ConfigureAwait(false));
    }

    // Mongolian vertical Latin runs as a horizontal left-to-right row of clockwise-spun glyphs
    // (Office emits per-run 0 -1 1 0 text matrices). The layout engine stacks vertical text downward
    // under a shared pivot, so remap those shape spans to per-run clockwise matrices here.
    // Single-sample calibration on pptx-ladder-04-vertical-text-port: row starts at shape X plus
    // 6.15pt with a self-calibrating pitch (median nonzero run-Y step, 28.8pt on the fixture).
    private const double MongolianVerticalRowStartInset = 6.15d;

    private static bool IsMongolianVerticalShape(XElement shape)
    {
        foreach (XElement bodyPr in shape.Descendants(DrawingNamespace + "bodyPr"))
        {
            XAttribute? vert = bodyPr.Attribute("vert");
            if (vert is not null && string.Equals(vert.Value, "mongolianVert", StringComparison.Ordinal))
            {
                return true;
            }
        }

        return false;
    }

    private static IReadOnlyList<PptxPositionedTextSpan> RemapMongolianVerticalSpans(IReadOnlyList<PptxPositionedTextSpan> spans, XElement shape)
    {
        if (spans.Count == 0 || !IsMongolianVerticalShape(shape))
        {
            return spans;
        }

        double pitch = ResolveMongolianVerticalPitch(spans);
        if (pitch <= PptxTextMetricRules.CoordinateTolerance)
        {
            return spans;
        }

        PptxPositionedTextSpan first = spans[0];
        if (!HasTextTransform(first.Run))
        {
            return spans;
        }

        (double a, double b, double c, double d, double e, double f) = TextTransformMatrix(first.Run);
        double firstBaselineY = first.Run.Y + first.Run.BaselineOffset;
        double deviceY = b * first.Run.X + d * firstBaselineY + f;
        double startX = first.FrameShapeX + MongolianVerticalRowStartInset;
        var remapped = new List<PptxPositionedTextSpan>(spans.Count);
        for (int i = 0; i < spans.Count; i++)
        {
            PptxPositionedTextSpan span = spans[i];
            double newX = startX + i * pitch;
            double newY = deviceY - span.Run.BaselineOffset;
            TextRun newRun = span.Run with { X = newX, Y = newY, RotationDegrees = 0d, PreventCoalesce = true, GlyphRotationQuarterTurns = 1 };
            IReadOnlyList<PptxTextGlyphLayout> glyphs = span.GlyphSpan.Glyphs;
            if (glyphs.Count > 0 && Math.Abs(glyphs[0].AdjustmentBefore) > PptxTextMetricRules.TextStateTolerance)
            {
                var fixedGlyphs = new PptxTextGlyphLayout[glyphs.Count];
                for (int g = 0; g < glyphs.Count; g++)
                {
                    fixedGlyphs[g] = g == 0 ? glyphs[g] with { AdjustmentBefore = 0d } : glyphs[g];
                }

                glyphs = fixedGlyphs;
            }

            PptxTextGlyphSpanLayout newGlyphSpan = span.GlyphSpan with { Glyphs = glyphs };
            remapped.Add(span with { Run = newRun, EndX = newX + span.Run.Width, LineBox = null, GlyphSpan = newGlyphSpan });
        }

        return remapped;
    }

    private static double ResolveMongolianVerticalPitch(IReadOnlyList<PptxPositionedTextSpan> spans)
    {
        var gaps = new List<double>(spans.Count);
        for (int i = 1; i < spans.Count; i++)
        {
            double gap = Math.Abs(spans[i].Run.Y - spans[i - 1].Run.Y);
            if (gap > PptxTextMetricRules.CoordinateTolerance)
            {
                gaps.Add(gap);
            }
        }

        if (gaps.Count == 0)
        {
            return 0d;
        }

        gaps.Sort();
        return gaps[(gaps.Count - 1) / 2];
    }

    private static async ValueTask<(IReadOnlyList<PptxPositionedTextSpan> Spans, bool AllVisibleTextClipped)>
        BuildTableCellTextSpansAsync(PptxTextFrameModel frame, PptxRenderContext context)
    {
        PptxTextFrameLayout layout = await BuildTextFrameLayoutAsync(frame, context.Document,
            new TextAdvanceEstimator(context.FontResolver, context.CancellationToken)).ConfigureAwait(false);
        IReadOnlyList<PptxPositionedTextSpan> spans = await FlattenTextLayoutToSpansAsync(
            new PptxTextLayoutModel([layout]), context.FontResolver, context.Document).ConfigureAwait(false);
        return (spans, layout.HasCulledVisibleTableText && !spans.Any(span => IsVisibleTableText(span.Run)));
    }

    private static bool IsVisibleTableText(TextRun run) =>
        (run.Alpha > 0d || run.Outline is { Alpha: > 0d }) && !string.IsNullOrWhiteSpace(run.Text);

    private static IReadOnlyList<XElement> FindInheritedPlaceholderShapes(XElement shape, IReadOnlyList<XDocument> placeholderSources)
    {
        return PptxPlaceholderMatcher.FindInheritedPlaceholderShapes(shape, placeholderSources);
    }

    private static IEnumerable<PptxTextFlowSegment> SplitFlowSegments(string text, bool attachSpacesToFollowingWord)
    {
        if (!attachSpacesToFollowingWord)
        {
            foreach (PptxTextFlowSegment segment in SplitFlowSegmentsWithTrailingSpaces(text))
            {
                yield return segment;
            }

            yield break;
        }

        int index = 0;
        while (index < text.Length)
        {
            int start = index;
            while (index < text.Length && text[index] == ' ')
            {
                index++;
            }

            if (index >= text.Length)
            {
                if (index > start)
                {
                    foreach (PptxTextFlowSegment segment in SplitControlSegments(text[start..index]))
                    {
                        yield return segment;
                    }
                }

                yield break;
            }

            while (index < text.Length && text[index] != ' ')
            {
                index++;
                if (text[index - 1] == '-' && index < text.Length && text[index] != ' ')
                {
                    break;
                }
            }

            if (index > start)
            {
                foreach (PptxTextFlowSegment segment in SplitControlSegments(text[start..index]))
                {
                    yield return segment;
                }
            }
        }
    }

    private static IEnumerable<PptxTextFlowSegment> SplitFlowSegmentsWithTrailingSpaces(string text)
    {
        int index = 0;
        while (index < text.Length)
        {
            int start = index;
            while (index < text.Length && text[index] != ' ')
            {
                index++;
                if (text[index - 1] == '-' && index < text.Length && text[index] != ' ')
                {
                    break;
                }
            }

            while (index < text.Length && text[index] == ' ')
            {
                index++;
            }

            if (index > start)
            {
                foreach (PptxTextFlowSegment segment in SplitControlSegments(text[start..index]))
                {
                    yield return segment;
                }
            }
        }
    }

    private static IEnumerable<PptxTextFlowSegment> SplitControlSegments(string text)
    {
        var builder = new StringBuilder();
        bool nextPreventsCoalesce = false;
        bool hideLeadingSpaces = false;
        for (int i = 0; i < text.Length; i++)
        {
            char c = text[i];
            if (c == '\r' || c == '\n')
            {
                if (builder.Length > 0)
                {
                    yield return new PptxTextFlowSegment(builder.ToString(), builder.ToString(), PptxTextFlowSegmentKind.Text, Draw: true, PreventCoalesce: nextPreventsCoalesce, FontScale: 1d);
                    builder.Clear();
                }

                if (c == '\r' && i + 1 < text.Length && text[i + 1] == '\n')
                {
                    i++;
                }

                yield return new PptxTextFlowSegment("\n", "\n", PptxTextFlowSegmentKind.Break, Draw: false, PreventCoalesce: true, FontScale: 1d);
                nextPreventsCoalesce = false;
                hideLeadingSpaces = false;
                continue;
            }

            if (hideLeadingSpaces && c == ' ')
            {
                if (builder.Length > 0)
                {
                    yield return new PptxTextFlowSegment(builder.ToString(), builder.ToString(), PptxTextFlowSegmentKind.Text, Draw: true, PreventCoalesce: nextPreventsCoalesce, FontScale: 1d);
                    builder.Clear();
                }

                yield return new PptxTextFlowSegment(string.Empty, c.ToString(), PptxTextFlowSegmentKind.HiddenAdvance, Draw: false, PreventCoalesce: true, FontScale: 1d);
                nextPreventsCoalesce = true;
                continue;
            }

            hideLeadingSpaces = false;
            if (c == '\u00A0' || c == '\u202F')
            {
                if (builder.Length > 0)
                {
                    yield return new PptxTextFlowSegment(builder.ToString(), builder.ToString(), PptxTextFlowSegmentKind.Text, Draw: true, PreventCoalesce: nextPreventsCoalesce, FontScale: 1d);
                    builder.Clear();
                }

                yield return new PptxTextFlowSegment(string.Empty, c.ToString(), PptxTextFlowSegmentKind.NoBreakHiddenAdvance, Draw: false, PreventCoalesce: true, FontScale: 1d);
                nextPreventsCoalesce = true;
                continue;
            }

            if (c == '\u00AD')
            {
                if (builder.Length > 0)
                {
                    yield return new PptxTextFlowSegment(builder.ToString(), builder.ToString(), PptxTextFlowSegmentKind.Text, Draw: true, PreventCoalesce: nextPreventsCoalesce, FontScale: 1d);
                    builder.Clear();
                }

                nextPreventsCoalesce = true;
                continue;
            }

            if (OfficePdfTextEmissionProfile.IsBoundaryPunctuation(c))
            {
                if (builder.Length > 0)
                {
                    yield return new PptxTextFlowSegment(builder.ToString(), builder.ToString(), PptxTextFlowSegmentKind.Text, Draw: true, PreventCoalesce: nextPreventsCoalesce, FontScale: 1d);
                    builder.Clear();
                }

                yield return new PptxTextFlowSegment(c.ToString(), c.ToString(), PptxTextFlowSegmentKind.BoundaryPunctuation, Draw: true, PreventCoalesce: true, FontScale: 1d);
                nextPreventsCoalesce = true;
                hideLeadingSpaces = true;
                continue;
            }

            builder.Append(c);
        }

        if (builder.Length > 0)
        {
            yield return new PptxTextFlowSegment(builder.ToString(), builder.ToString(), PptxTextFlowSegmentKind.Text, Draw: true, PreventCoalesce: nextPreventsCoalesce, FontScale: 1d);
        }
    }
}
