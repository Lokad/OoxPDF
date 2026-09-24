using System.Runtime.CompilerServices;

namespace Lokad.OoxPdf.Docx;

// Cell text-line memo. One table layout pass measures each cell up to three
// times (split-feasibility, break resolution, final fragments) plus nested-table
// remeasurement; keys capture every layout input so hits reproduce the computed lines
// exactly after caller-side re-offsetting by the new origin. Stored lines are
// origin-relative (absolute minus store origin); callers reconstruct absolute lines
// by shifting relative lines by the lookup origin. Cells carrying PAGE/NUMPAGES runs
// keep their page args in the key (dynamic tier); all other cells omit them (static
// tier), which is sound because page inputs only reach text spans and wrap measurement
// through dynamic field runs. Vertical-alignment shifts stay outside the memo: every
// caller applies its own height-dependent shift to the shared unshifted lines.
// Cell identity is explicit reference identity: equal-valued distinct cell instances
// must not share entries.
internal sealed partial class DocxLayoutEngine
{
    internal static bool HasPageDynamicFields(IReadOnlyList<DocxBodyElement> bodyElements)
    {
        foreach (DocxBodyElement element in bodyElements)
        {
            if (element is DocxParagraphElement paragraphElement)
            {
                foreach (DocxTextRun run in paragraphElement.Paragraph.Runs)
                {
                    if (run.FieldKind is DocxFieldKind.Page or DocxFieldKind.NumPages)
                    {
                        return true;
                    }
                }
            }
            else if (element is DocxTableElement tableElement)
            {
                foreach (DocxTableRow row in tableElement.Table.Rows)
                {
                    foreach (DocxTableCell nested in row.Cells)
                    {
                        if (HasPageDynamicFields(GetTableCellLayoutBodyElements(nested)))
                        {
                            return true;
                        }
                    }
                }
            }
        }

        return false;
    }

    internal sealed class DocxTableCellTextLinesMemo
    {
        public long Hits { get; private set; }

        public long Misses { get; private set; }

        internal static long TotalHits { get; private set; }

        internal static long TotalMisses { get; private set; }

        internal static void ResetTotals()
        {
            TotalHits = 0;
            TotalMisses = 0;
        }

        private readonly Dictionary<MemoKey, StoredLines> entries = new();

        public bool TryGetRelativeLines(
            DocxTableCell cell,
            double cellWidth,
            IDocxTextMeasurer? measurer,
            double defaultTabStopPoints,
            double rowTopPadding,
            double paragraphSpacingScale,
            int? pageNumber,
            int? pageCount,
            bool pageStatic,
            out IReadOnlyList<DocxTextLineLayout> relativeLines,
            out IReadOnlyList<DocxInlineImageLayout> relativePlacedImages,
            out double usedHeight)
        {
            if (entries.TryGetValue(
                new MemoKey(cell, cellWidth, measurer, defaultTabStopPoints, rowTopPadding, paragraphSpacingScale, pageStatic ? null : pageNumber, pageStatic ? null : pageCount),
                out StoredLines? stored) && stored is not null)
            {
                Hits++;
                TotalHits++;
                relativeLines = stored.Lines;
                relativePlacedImages = stored.PlacedImages;
                usedHeight = stored.UsedHeight;
                return true;
            }

            Misses++;
            TotalMisses++;
            relativeLines = [];
            relativePlacedImages = [];
            usedHeight = 0d;
            return false;
        }

        public void StoreRelativeLines(
            DocxTableCell cell,
            double cellWidth,
            IDocxTextMeasurer? measurer,
            double defaultTabStopPoints,
            double rowTopPadding,
            double paragraphSpacingScale,
            int? pageNumber,
            int? pageCount,
            bool pageStatic,
            IReadOnlyList<DocxTextLineLayout> absoluteLines,
            IReadOnlyList<DocxInlineImageLayout> absolutePlacedImages,
            double originX,
            double originY,
            double usedHeight)
        {
            IReadOnlyList<DocxTextLineLayout> relative = ToRelativeLines(absoluteLines, originX, originY);
            IReadOnlyList<DocxInlineImageLayout> relativePlaced = ToRelativeImages(absolutePlacedImages, originX, originY);
            entries[new MemoKey(cell, cellWidth, measurer, defaultTabStopPoints, rowTopPadding, paragraphSpacingScale, pageStatic ? null : pageNumber, pageStatic ? null : pageCount)] = new StoredLines(relative, relativePlaced, usedHeight);
        }

        internal static IReadOnlyList<DocxTextLineLayout> ToRelativeLines(
            IReadOnlyList<DocxTextLineLayout> lines,
            double originX,
            double originY)
        {
            var relative = new List<DocxTextLineLayout>(lines.Count);
            foreach (DocxTextLineLayout line in lines)
            {
                relative.Add(ShiftLine(line, -originX, -originY));
            }

            return relative;
        }

        internal static IReadOnlyList<DocxTextLineLayout> ShiftLines(
            IReadOnlyList<DocxTextLineLayout> lines,
            double deltaX,
            double deltaY)
        {
            if (deltaX == 0d && deltaY == 0d)
            {
                return lines;
            }

            var shifted = new List<DocxTextLineLayout>(lines.Count);
            foreach (DocxTextLineLayout line in lines)
            {
                shifted.Add(ShiftLine(line, deltaX, deltaY));
            }

            return shifted;
        }

        private static DocxTextLineLayout ShiftLine(DocxTextLineLayout line, double deltaX, double deltaY)
        {
            return line with
            {
                X = line.X + deltaX,
                BaselineY = line.BaselineY + deltaY,
                Segments = ShiftSegments(line.Segments, deltaX)
            };
        }

        private static IReadOnlyList<DocxTextSegmentLayout> ShiftSegments(IReadOnlyList<DocxTextSegmentLayout> segments, double deltaX)
        {
            if (deltaX == 0d)
            {
                return segments;
            }

            var shifted = new List<DocxTextSegmentLayout>(segments.Count);
            foreach (DocxTextSegmentLayout segment in segments)
            {
                shifted.Add(segment with { X = segment.X + deltaX });
            }

            return shifted;
        }

        private readonly struct MemoKey : IEquatable<MemoKey>
        {
            private readonly DocxTableCell cell;
            private readonly double cellWidth;
            private readonly IDocxTextMeasurer? measurer;
            private readonly double defaultTabStopPoints;
            private readonly double rowTopPadding;
            private readonly double paragraphSpacingScale;
            private readonly int? pageNumber;
            private readonly int? pageCount;

            public MemoKey(
                DocxTableCell cell,
                double cellWidth,
                IDocxTextMeasurer? measurer,
                double defaultTabStopPoints,
                double rowTopPadding,
                double paragraphSpacingScale,
                int? pageNumber,
                int? pageCount)
            {
                this.cell = cell;
                this.cellWidth = cellWidth;
                this.measurer = measurer;
                this.defaultTabStopPoints = defaultTabStopPoints;
                this.rowTopPadding = rowTopPadding;
                this.paragraphSpacingScale = paragraphSpacingScale;
                this.pageNumber = pageNumber;
                this.pageCount = pageCount;
            }

            public bool Equals(MemoKey other)
            {
                return ReferenceEquals(cell, other.cell)
                    && cellWidth.Equals(other.cellWidth)
                    && ReferenceEquals(measurer, other.measurer)
                    && defaultTabStopPoints.Equals(other.defaultTabStopPoints)
                    && rowTopPadding.Equals(other.rowTopPadding)
                    && paragraphSpacingScale.Equals(other.paragraphSpacingScale)
                    && pageNumber == other.pageNumber
                    && pageCount == other.pageCount;
            }

            public override bool Equals(object? obj)
            {
                return obj is MemoKey other && Equals(other);
            }

            public override int GetHashCode()
            {
                var hash = new HashCode();
                hash.Add(RuntimeHelpers.GetHashCode(cell));
                hash.Add(cellWidth);
                hash.Add(measurer is null ? 0 : RuntimeHelpers.GetHashCode(measurer));
                hash.Add(defaultTabStopPoints);
                hash.Add(rowTopPadding);
                hash.Add(paragraphSpacingScale);
                hash.Add(pageNumber);
                hash.Add(pageCount);
                return hash.ToHashCode();
            }
        }

        private sealed class StoredLines(IReadOnlyList<DocxTextLineLayout> lines, IReadOnlyList<DocxInlineImageLayout> placedImages, double usedHeight)
        {
            public IReadOnlyList<DocxTextLineLayout> Lines { get; } = lines;

            public IReadOnlyList<DocxInlineImageLayout> PlacedImages { get; } = placedImages;

            public double UsedHeight { get; } = usedHeight;
        }

        // RV05: placed mid-line images relativize exactly like lines so memo hits
        // shift them back with the same origin.
        internal static IReadOnlyList<DocxInlineImageLayout> ToRelativeImages(
            IReadOnlyList<DocxInlineImageLayout> images,
            double originX,
            double originY)
        {
            var relative = new List<DocxInlineImageLayout>(images.Count);
            foreach (DocxInlineImageLayout image in images)
            {
                relative.Add(image with { X = image.X - originX, Y = image.Y - originY });
            }

            return relative;
        }
    }
}
