using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    private static DocxTextLineLayout ResolveCommentAnchorLine(
        DocxTextLineLayout fallbackLine,
        IReadOnlyList<DocxTextLineLayout> anchorTextLines,
        DocxParagraph paragraph,
        DocxInlineReference reference)
    {
        DocxCommentRange? range = paragraph.CommentRanges.FirstOrDefault(range =>
            string.Equals(range.Id, reference.Id, StringComparison.Ordinal));
        if (range is not null &&
            TryResolveCommentRangeEndAnchor(fallbackLine, anchorTextLines, paragraph, range, out DocxTextLineLayout anchorLine, out _))
        {
            return anchorLine;
        }

        foreach (DocxTextLineLayout candidateLine in EnumerateCommentAnchorSearchLines(fallbackLine, anchorTextLines, paragraph))
        {
            if (TryResolveSourceOffsetAnchorX(candidateLine, reference.SourceRunIndex, reference.TextOffsetInRun, out _) ||
                TryResolvePreviousSourceRunEndAnchorX(candidateLine, reference.SourceRunIndex, minimumSourceRunIndex: null, out _))
            {
                return candidateLine;
            }
        }

        return fallbackLine;
    }

    private static bool TryResolveCommentRangeEndAnchor(
        DocxTextLineLayout line,
        IReadOnlyList<DocxTextLineLayout> anchorTextLines,
        DocxParagraph paragraph,
        DocxCommentRange range,
        out DocxTextLineLayout anchorLine,
        out double anchorX)
    {
        foreach (DocxTextLineLayout candidateLine in EnumerateCommentAnchorSearchLines(line, anchorTextLines, paragraph))
        {
            if (TryResolveSourceOffsetAnchorX(candidateLine, range.EndSourceRunIndex, range.EndTextOffset, out anchorX))
            {
                anchorLine = candidateLine;
                return true;
            }
        }

        foreach (DocxTextLineLayout candidateLine in EnumerateCommentAnchorSearchLines(line, anchorTextLines, paragraph))
        {
            if (TryResolvePreviousSourceRunEndAnchorX(candidateLine, range.EndSourceRunIndex, range.StartSourceRunIndex, out anchorX))
            {
                anchorLine = candidateLine;
                return true;
            }
        }

        foreach (DocxTextLineLayout candidateLine in EnumerateCommentAnchorSearchLines(line, anchorTextLines, paragraph))
        {
            if (TryResolveSourceRunEndAnchorX(candidateLine, range.StartSourceRunIndex, out anchorX))
            {
                anchorLine = candidateLine;
                return true;
            }
        }

        anchorLine = line;
        anchorX = 0d;
        return false;
    }

    private static double ResolveCommentAnchorX(
        DocxTextLineLayout line,
        IReadOnlyList<DocxTextLineLayout> anchorTextLines,
        DocxParagraph paragraph,
        DocxInlineReference reference,
        DocxMarkupContext markupContext)
    {
        DocxCommentRange? range = paragraph.CommentRanges.FirstOrDefault(range =>
            string.Equals(range.Id, reference.Id, StringComparison.Ordinal));
        if (range is not null)
        {
            if (UsesWordCompatibleAllMarkupTextProfile(markupContext) &&
                TryResolveCommentRangeEndAnchor(line, anchorTextLines, paragraph, range, out DocxTextLineLayout anchorLine, out double wordCompatibleRangeEndX))
            {
                return wordCompatibleRangeEndX + ResolveTextEmissionXOffset(markupContext) -
                    WordCompatibleAllMarkupConnectorBodyAnchorInsetPoints;
            }

            if (TryResolveSourceOffsetAnchorX(line, range.StartSourceRunIndex, range.StartTextOffset, out double startX))
            {
                return startX + ResolveTextEmissionXOffset(markupContext);
            }

            if (TryResolveSourceOffsetAnchorX(line, range.EndSourceRunIndex, range.EndTextOffset, out double endX))
            {
                return endX + ResolveTextEmissionXOffset(markupContext);
            }
        }

        return ResolveInlineReferenceAnchorX(line, reference) + ResolveTextEmissionXOffset(markupContext);
    }

    private static IEnumerable<DocxTextLineLayout> EnumerateCommentAnchorSearchLines(
        DocxTextLineLayout line,
        IReadOnlyList<DocxTextLineLayout> anchorTextLines,
        DocxParagraph paragraph)
    {
        yield return line;
        foreach (DocxTextLineLayout candidateLine in anchorTextLines)
        {
            if (!ReferenceEquals(candidateLine, line) &&
                ReferenceEquals(candidateLine.SourceParagraph, paragraph))
            {
                yield return candidateLine;
            }
        }
    }

    private static double ResolveInlineReferenceAnchorX(DocxTextLineLayout line, DocxInlineReference reference)
    {
        if (reference.SourceRunIndex < 0)
        {
            return line.X + Math.Max(0d, line.Width) * 0.5d;
        }

        if (TryResolveSourceOffsetAnchorX(line, reference.SourceRunIndex, reference.TextOffsetInRun, out double anchorX))
        {
            return anchorX;
        }

        return line.X + Math.Max(0d, line.Width) * 0.5d;
    }

    private static bool TryResolveSourceOffsetAnchorX(
        DocxTextLineLayout line,
        int? sourceRunIndex,
        int? textOffsetInRun,
        out double anchorX)
    {
        anchorX = 0d;
        if (sourceRunIndex is not { } runIndex || runIndex < 0)
        {
            return false;
        }

        int offset = Math.Max(0, textOffsetInRun ?? 0);
        foreach (DocxTextSegmentLayout segment in line.Segments)
        {
            if (segment.SourceTextRunIndex != runIndex)
            {
                continue;
            }

            int segmentStart = Math.Max(0, segment.SourceTextOffsetInRun);
            int segmentEnd = segmentStart + segment.Text.Length;
            if (offset < segmentStart || offset > segmentEnd)
            {
                continue;
            }

            if (segment.Text.Length == 0)
            {
                anchorX = segment.X;
                return true;
            }

            double ratio = Math.Clamp((offset - segmentStart) / (double)segment.Text.Length, 0d, 1d);
            anchorX = segment.X + Math.Max(0d, segment.Width) * ratio;
            return true;
        }

        return false;
    }

    private static bool TryResolveSourceRunStartAnchorX(
        DocxTextLineLayout line,
        int? sourceRunIndex,
        out double anchorX)
    {
        anchorX = 0d;
        if (sourceRunIndex is not { } runIndex)
        {
            return false;
        }

        DocxTextSegmentLayout[] segments = line.Segments
            .Where(segment => segment.SourceTextRunIndex == runIndex)
            .ToArray();
        if (segments.Length == 0)
        {
            return false;
        }

        anchorX = segments.Min(segment => segment.X);
        return true;
    }

    private static bool TryResolvePreviousSourceRunEndAnchorX(
        DocxTextLineLayout line,
        int? sourceRunIndex,
        int? minimumSourceRunIndex,
        out double anchorX)
    {
        anchorX = 0d;
        if (sourceRunIndex is not { } runIndex)
        {
            return false;
        }

        int previousRunIndex = line.Segments
            .Where(segment =>
                segment.SourceTextRunIndex < runIndex &&
                (minimumSourceRunIndex is null || segment.SourceTextRunIndex >= minimumSourceRunIndex.Value))
            .Select(segment => segment.SourceTextRunIndex)
            .DefaultIfEmpty(-1)
            .Max();
        return previousRunIndex >= 0 &&
            TryResolveSourceRunEndAnchorX(line, previousRunIndex, out anchorX);
    }

    private static bool TryResolveSourceRunEndAnchorX(
        DocxTextLineLayout line,
        int? sourceRunIndex,
        out double anchorX)
    {
        anchorX = 0d;
        if (sourceRunIndex is not { } runIndex)
        {
            return false;
        }

        DocxTextSegmentLayout[] segments = line.Segments
            .Where(segment => segment.SourceTextRunIndex == runIndex)
            .ToArray();
        if (segments.Length == 0)
        {
            return false;
        }

        anchorX = segments.Max(segment => segment.X + Math.Max(0d, segment.Width));
        return true;
    }

    private static IEnumerable<DocxTableRowLayout> EnumerateMarkupBalloonTableRows(
        DocxLayoutPage page,
        IReadOnlyList<DocxFloatingDrawingLayout> floatingDrawings)
    {
        foreach (DocxTableRowLayout row in EnumerateTableRows(page.StaticTableRows))
        {
            yield return row;
        }

        foreach (DocxInlineTextBoxLayout box in page.StaticInlineTextBoxes)
        {
            foreach (DocxTableRowLayout row in box.TableRows)
            {
                foreach (DocxTableRowLayout nested in EnumerateTableRows(row))
                {
                    yield return nested;
                }
            }
        }

        foreach (DocxTableRowLayout row in EnumerateTableRows(page.Items))
        {
            yield return row;
        }

        foreach (DocxPlacedRelatedStoryLayout story in page.PlacedRelatedStories)
        {
            foreach (DocxTableRowLayout row in EnumerateTableRows(story.TableRows))
            {
                yield return row;
            }
        }

        foreach (DocxTableRowLayout row in EnumerateFloatingDrawingTextBoxTableRows(floatingDrawings))
        {
            yield return row;
        }

        foreach (DocxLayoutItem item in page.Items)
        {
            if (item is not DocxInlineTextBoxLayout box)
            {
                continue;
            }

            foreach (DocxTableRowLayout row in box.TableRows)
            {
                foreach (DocxTableRowLayout nested in EnumerateTableRows(row))
                {
                    yield return nested;
                }
            }
        }

        foreach (DocxPlacedRelatedStoryLayout story in page.PlacedRelatedStories)
        {
            foreach (DocxTableRowLayout row in EnumerateFloatingDrawingTextBoxTableRows(story.FloatingDrawings))
            {
                yield return row;
            }
        }
    }

    private static IEnumerable<DocxTableRowLayout> EnumerateTableRows(IEnumerable<DocxLayoutItem> items)
    {
        foreach (DocxLayoutItem item in items)
        {
            if (item is DocxTableRowLayout row)
            {
                foreach (DocxTableRowLayout nested in EnumerateTableRows(row))
                {
                    yield return nested;
                }
            }
        }
    }

    private static IEnumerable<DocxTableRowLayout> EnumerateTableRows(IEnumerable<DocxTableRowLayout> rows)
    {
        foreach (DocxTableRowLayout row in rows)
        {
            foreach (DocxTableRowLayout nested in EnumerateTableRows(row))
            {
                yield return nested;
            }
        }
    }

    private static IEnumerable<DocxTableRowLayout> EnumerateTableRows(DocxTableRowLayout row)
    {
        yield return row;
        foreach (DocxTableCellLayout cell in row.Cells)
        {
            foreach (DocxTableRowLayout nestedRow in cell.NestedRows)
            {
                foreach (DocxTableRowLayout nested in EnumerateTableRows(nestedRow))
                {
                    yield return nested;
                }
            }

            foreach (DocxInlineTextBoxLayout box in cell.InlineTextBoxes)
            {
                foreach (DocxTableRowLayout boxRow in box.TableRows)
                {
                    foreach (DocxTableRowLayout nested in EnumerateTableRows(boxRow))
                    {
                        yield return nested;
                    }
                }
            }
        }
    }

    private static string TableBalloonKey(DocxTableRowLayout row)
    {
        return (row.StoryKind ?? string.Empty) +
            ":" + (row.StoryVariantType ?? string.Empty) +
            ":" + row.Table.SourceBlockIndex.ToString(CultureInfo.InvariantCulture) +
            ":" + row.Table.TableIndex.ToString(CultureInfo.InvariantCulture);
    }

    private static string TableRowBalloonKey(DocxTableRowLayout row)
    {
        return TableBalloonKey(row) +
            ":" + row.RowIndex.ToString(CultureInfo.InvariantCulture);
    }

    private static string TableCellBalloonKey(DocxTableRowLayout row, int cellIndex)
    {
        return TableRowBalloonKey(row) +
            ":" + cellIndex.ToString(CultureInfo.InvariantCulture);
    }
}
