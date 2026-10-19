using System.Runtime.CompilerServices;

namespace Lokad.OoxPdf.Docx;

// Cell text-line memo. One table layout pass measures each cell up to three
// times (split-feasibility, break resolution, final fragments) plus nested-table
// remeasurement; keys capture every layout input so hits reproduce the computed lines
// exactly after caller-side re-offsetting by the new origin. Stored lines are
// origin-relative (absolute minus store origin); callers reconstruct absolute lines
// by shifting relative lines by the lookup origin. Source PAGE/NUMPAGES dependencies
// select each page input independently, including inline-box children. The memo
// stays within one table-layout pass; math/story context remains owned by that pass.
// Vertical-alignment
// shifts stay outside the memo: every
// caller applies its own height-dependent shift to the shared unshifted lines.
// Cell identity is explicit reference identity: equal-valued distinct cell instances
// must not share entries.
internal sealed partial class DocxLayoutEngine
{
    internal sealed class DocxTableCellTextLinesMemo
    {
        public long Hits { get; private set; }

        public long Misses { get; private set; }

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
            DocxCounterDependencies dependencies,
            out IReadOnlyList<DocxTextLineLayout> relativeLines,
            out IReadOnlyList<DocxInlineImageLayout> relativePlacedImages,
            out IReadOnlyList<DocxInlineTextBoxLayout> relativePlacedBoxes,
            out double usedHeight)
        {
            if (entries.TryGetValue(
                new MemoKey(cell, cellWidth, measurer, defaultTabStopPoints, rowTopPadding, paragraphSpacingScale, pageNumber, pageCount, dependencies),
                out StoredLines? stored) && stored is not null)
            {
                Hits++;
                relativeLines = stored.Lines;
                relativePlacedImages = stored.PlacedImages;
                relativePlacedBoxes = stored.PlacedBoxes;
                usedHeight = stored.UsedHeight;
                return true;
            }

            Misses++;
            relativeLines = [];
            relativePlacedImages = [];
            relativePlacedBoxes = [];
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
            DocxCounterDependencies dependencies,
            IReadOnlyList<DocxTextLineLayout> absoluteLines,
            IReadOnlyList<DocxInlineImageLayout> absolutePlacedImages,
            IReadOnlyList<DocxInlineTextBoxLayout> absolutePlacedBoxes,
            double originX,
            double originY,
            double usedHeight)
        {
            IReadOnlyList<DocxTextLineLayout> relative = ToRelativeLines(absoluteLines, originX, originY);
            IReadOnlyList<DocxInlineImageLayout> relativePlaced = ToRelativeImages(absolutePlacedImages, originX, originY);
            var relativeBoxes = absolutePlacedBoxes.Select(box => ShiftInlineTextBox(box, -originY, -originX)).ToArray();
            entries[new MemoKey(cell, cellWidth, measurer, defaultTabStopPoints, rowTopPadding, paragraphSpacingScale, pageNumber, pageCount, dependencies)] = new StoredLines(relative, relativePlaced, relativeBoxes, usedHeight);
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
                int? pageCount,
                DocxCounterDependencies dependencies)
            {
                this.cell = cell;
                this.cellWidth = cellWidth;
                this.measurer = measurer;
                this.defaultTabStopPoints = defaultTabStopPoints;
                this.rowTopPadding = rowTopPadding;
                this.paragraphSpacingScale = paragraphSpacingScale;
                this.pageNumber = (dependencies & DocxCounterDependencies.PageNumber) != 0 ? pageNumber : null;
                this.pageCount = (dependencies & DocxCounterDependencies.PageCount) != 0 ? pageCount : null;
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

        private sealed class StoredLines(IReadOnlyList<DocxTextLineLayout> lines, IReadOnlyList<DocxInlineImageLayout> placedImages,
            IReadOnlyList<DocxInlineTextBoxLayout> placedBoxes, double usedHeight)
        {
            public IReadOnlyList<DocxInlineTextBoxLayout> PlacedBoxes { get; } = placedBoxes;
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
                relative.Add(image.Translate(-originX, -originY));
            }

            return relative;
        }
    }
}
