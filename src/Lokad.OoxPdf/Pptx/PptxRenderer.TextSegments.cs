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
    private static bool ShouldUseOfficeOverflowColumnBalance(PptxTextFrameLayout layout)
    {
        if (!TryReadOverflowColumnLineCounts(layout, out PptxTextLineLayout[] lines, out int[] counts))
        {
            return false;
        }

        PptxTextFrameModel frame = layout.Model;
        if (lines.Length % frame.ColumnCount != 0)
        {
            return false;
        }

        int target = lines.Length / frame.ColumnCount;
        return counts.Take(frame.ColumnCount - 1).All(count => count == target - 1) &&
            counts[^1] == target + frame.ColumnCount - 1;
    }

    private static bool TryResolveOfficeOverflowColumnLineBalance(PptxTextFrameLayout layout, out int lineBalanceTarget, out int lineBalanceStartColumn)
    {
        lineBalanceTarget = 0;
        lineBalanceStartColumn = 0;
        if (!TryReadOverflowColumnLineCounts(layout, out PptxTextLineLayout[] lines, out int[] counts))
        {
            return false;
        }

        PptxTextFrameModel frame = layout.Model;
        int balancedFloor = lines.Length / frame.ColumnCount;
        int balancedCeiling = (int)Math.Ceiling((double)lines.Length / frame.ColumnCount);
        if (counts.Take(frame.ColumnCount - 1).All(count => count < balancedFloor) &&
            counts[^1] >= balancedCeiling + frame.ColumnCount)
        {
            lineBalanceTarget = balancedCeiling;
            return true;
        }

        if (ShouldResolveTrailingOverflowColumnBalance(layout, counts, balancedFloor, balancedCeiling))
        {
            lineBalanceTarget = balancedFloor;
            lineBalanceStartColumn = 1;
            return true;
        }

        if (ShouldResolveOverflowColumnBalanceAcrossContinuedParagraph(layout, lines.Length, counts, balancedCeiling))
        {
            lineBalanceTarget = balancedCeiling;
            return true;
        }

        if (ShouldResolveEvenOverflowColumnBalanceAcrossContinuedParagraph(layout, counts, balancedFloor))
        {
            lineBalanceTarget = balancedFloor + 1;
            lineBalanceStartColumn = 1;
            return true;
        }

        return false;
    }

    private static bool ShouldResolveTrailingOverflowColumnBalance(PptxTextFrameLayout layout, int[] counts, int balancedFloor, int balancedCeiling)
    {
        PptxTextFrameModel frame = layout.Model;
        return frame.ColumnCount >= 3 &&
            counts[0] >= balancedFloor &&
            counts.Skip(1).Take(frame.ColumnCount - 2).Any(count => count < balancedFloor) &&
            counts[^1] > balancedFloor;
    }

    private static bool ShouldResolveOverflowColumnBalanceAcrossContinuedParagraph(PptxTextFrameLayout layout, int lineCount, int[] counts, int balancedCeiling)
    {
        PptxTextFrameModel frame = layout.Model;
        if (lineCount % frame.ColumnCount != frame.ColumnCount - 1 ||
            counts[^1] != balancedCeiling ||
            counts.Take(frame.ColumnCount - 1).Sum() != balancedCeiling * (frame.ColumnCount - 1) - 1 ||
            counts[^2] != balancedCeiling - 1)
        {
            return false;
        }

        var columns = layout.Paragraphs
            .SelectMany((paragraph, paragraphIndex) => paragraph.Lines.Select(line => new
            {
                Line = line,
                ParagraphIndex = paragraphIndex
            }))
            .GroupBy(item => Math.Round(item.Line.StartX, 2))
            .OrderBy(group => group.Key)
            .Select(group => group.ToArray())
            .ToArray();
        if (columns.Length != frame.ColumnCount || columns[^2].Length == 0 || columns[^1].Length == 0)
        {
            return false;
        }

        return columns[^2][^1].ParagraphIndex == columns[^1][0].ParagraphIndex;
    }

    private static bool ShouldResolveEvenOverflowColumnBalanceAcrossContinuedParagraph(PptxTextFrameLayout layout, int[] counts, int balancedFloor)
    {
        PptxTextFrameModel frame = layout.Model;
        if (frame.ColumnCount < 3 ||
            counts.Any(count => count != balancedFloor))
        {
            return false;
        }

        var columns = layout.Paragraphs
            .SelectMany((paragraph, paragraphIndex) => paragraph.Lines.Select(line => new
            {
                Line = line,
                ParagraphIndex = paragraphIndex
            }))
            .GroupBy(item => Math.Round(item.Line.StartX, 2))
            .OrderBy(group => group.Key)
            .Select(group => group.ToArray())
            .ToArray();
        if (columns.Length != frame.ColumnCount || columns[^2].Length == 0 || columns[^1].Length == 0)
        {
            return false;
        }

        return columns[^2][^1].ParagraphIndex == columns[^1][0].ParagraphIndex;
    }

    private static bool TryReadOverflowColumnLineCounts(PptxTextFrameLayout layout, out PptxTextLineLayout[] lines, out int[] counts)
    {
        PptxTextFrameModel frame = layout.Model;
        lines = [];
        counts = [];
        if (frame.ColumnCount <= 1 || frame.BodyProperties.VerticalOverflow != PptxTextVerticalOverflow.Overflow)
        {
            return false;
        }

        lines = layout.Paragraphs.SelectMany(paragraph => paragraph.Lines).ToArray();
        if (lines.Length == 0)
        {
            return false;
        }

        counts = lines
            .GroupBy(line => Math.Round(line.StartX, 2))
            .OrderBy(group => group.Key)
            .Select(group => group.Count())
            .ToArray();
        if (counts.Length != frame.ColumnCount)
        {
            return false;
        }

        return true;
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

    private static IReadOnlyList<PptxPositionedTextSpan> ReadTextSpansForSceneNode(PptxSceneNode node, PptxRenderContext context, PptxColorMap colorMap, bool includePlaceholders)
    {
        return node.TextBody is null
            ? []
            : ReadTextSpansForShape(node.Source, context.Document, context.Theme, colorMap, context.SlideNumber, includePlaceholders, context.InheritedXml, context.FontResolver, context.CancellationToken);
    }

    private static IReadOnlyList<PptxPositionedTextSpan> ReadTextSpansForShape(
        XElement shape,
        PptxRenderContext context,
        bool includePlaceholders)
    {
        return ReadTextSpansForShape(shape, context.Document, context.Theme, context.SlideColorMap, context.SlideNumber, includePlaceholders, context.InheritedXml, context.FontResolver, context.CancellationToken);
    }

    private static IReadOnlyList<PptxPositionedTextSpan> ReadTextSpansForShape(
        XElement shape,
        PptxDocument document,
        PptxTheme theme,
        PptxColorMap colorMap,
        int slideNumber,
        bool includePlaceholders,
        IReadOnlyList<XDocument> placeholderSources,
        PresentationFontResolver? fontResolver,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        XElement current = new(shape);
        foreach (XElement group in shape.Ancestors(PresentationNamespace + "grpSp"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var groupCopy = new XElement(PresentationNamespace + "grpSp");
            if (group.Element(PresentationNamespace + "grpSpPr") is { } properties)
            {
                groupCopy.Add(new XElement(properties));
            }

            groupCopy.Add(current);
            current = groupCopy;
        }

        var slide = new XDocument(
            new XElement(PresentationNamespace + "sld",
                new XElement(PresentationNamespace + "cSld",
                    new XElement(PresentationNamespace + "spTree", current))));
        return RemapMongolianVerticalSpans(FlattenTextLayoutToSpans(BuildTextLayoutModel(slide, document, theme, colorMap, slideNumber, includePlaceholders, placeholderSources, fontResolver, cancellationToken), fontResolver), shape);
    }

    // Mongolian vertical Latin runs as a horizontal left-to-right row of clockwise-spun glyphs
    // (Office emits per-run 0 -1 1 0 text matrices). The layout engine stacks vertical text downward
    // under a shared pivot, so remap those spans to per-run clockwise matrices here. Tables never
    // reach this path (table cells use ReadTextSpansForTableCellTextFrame below).
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

    private static IReadOnlyList<PptxPositionedTextSpan> ReadTextSpansForTableCellTextFrame(PptxTableCellTextFrame tableFrame, PptxRenderContext context)
    {
        PptxTextFrameModel frameModel = BuildTextFrameModel(tableFrame, context.Document, context.Theme, context.SlideNumber, context.InheritedXml, context.FontResolver, context.CancellationToken);
        PptxTextFrameLayout layout = BuildTextFrameLayout(frameModel, context.Document, new TextAdvanceEstimator(context.FontResolver, context.CancellationToken));
        return FlattenTextLayoutToSpans(new PptxTextLayoutModel([layout]), context.FontResolver);
    }

    private static IReadOnlyList<XElement> FindInheritedPlaceholderShapes(XElement shape, IReadOnlyList<XDocument> placeholderSources)
    {
        return PptxPlaceholderMatcher.FindInheritedPlaceholderShapes(shape, placeholderSources);
    }

    private static XElement? FindInheritedTextStyle(XElement shape, IReadOnlyList<XDocument> placeholderSources, string levelName)
    {
        return PptxTextStyleInheritance.FindInheritedTextStyle(shape, placeholderSources, levelName);
    }

    private static XElement? FindDefaultTextStyle(IReadOnlyList<XDocument> placeholderSources, string levelName)
    {
        return PptxTextStyleInheritance.FindDefaultTextStyle(placeholderSources, levelName);
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

            if (IsOfficeTextOperationBoundaryPunctuation(c))
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
