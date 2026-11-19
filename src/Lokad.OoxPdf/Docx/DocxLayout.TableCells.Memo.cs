using System.Runtime.CompilerServices;

namespace Lokad.OoxPdf.Docx;

// Cell content memo. One table layout pass measures each cell up to three
// times (split-feasibility, break resolution, final fragments) plus nested-table
// remeasurement; keys capture every layout input so hits reproduce the computed lines
// exactly after caller-side re-offsetting by the new origin. Stored content is
// origin-relative (absolute minus store origin); callers reconstruct absolute lines
// by shifting relative lines by the lookup origin. Source PAGE/NUMPAGES dependencies
// select each page input independently, including inline-box children. The memo
// stays within one table-layout pass; math/story context remains owned by that pass.
// Each source/geometry key retains only its current counter inputs. Keeping all
// prior PAGE values would retain a full cell measurement for every fragment.
// Callers apply height-dependent vertical alignment after replay.
// Cell identity is explicit reference identity: equal-valued distinct cell instances
// must not share entries.
internal sealed partial class DocxLayoutEngine
{
    internal sealed class DocxTableCellTextLinesMemo
    {
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
            out double usedHeight,
            out IReadOnlyList<DocxTableRowLayout> relativeNestedRows,
            double nestedPageHeight = 0d,
            List<DocxPlacedDrawingAnchor>? relativeAnchors = null)
        {
            if (entries.TryGetValue(
                new MemoKey(cell, cellWidth, measurer, defaultTabStopPoints, rowTopPadding, paragraphSpacingScale, nestedPageHeight),
                out StoredLines? stored) && stored is not null &&
                stored.CounterInputs == ResolveCounterInputs(pageNumber, pageCount, dependencies))
            {
                relativeLines = stored.Lines;
                relativePlacedImages = stored.PlacedImages;
                relativePlacedBoxes = stored.PlacedBoxes;
                usedHeight = stored.UsedHeight;
                relativeNestedRows = stored.NestedRows;
                relativeAnchors?.AddRange(stored.DrawingAnchors);
                return true;
            }

            relativeLines = [];
            relativePlacedImages = [];
            relativePlacedBoxes = [];
            usedHeight = 0d;
            relativeNestedRows = [];
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
            double usedHeight,
            IReadOnlyList<DocxTableRowLayout> absoluteNestedRows,
            double nestedPageHeight = 0d,
            IReadOnlyList<DocxPlacedDrawingAnchor>? absoluteAnchors = null)
        {
            IReadOnlyList<DocxTextLineLayout> relative = ShiftTextLines(absoluteLines, -originY, -originX);
            IReadOnlyList<DocxInlineImageLayout> relativePlaced = ShiftInlineImages(absolutePlacedImages, -originY, -originX);
            var relativeBoxes = absolutePlacedBoxes.Select(box => ShiftInlineTextBox(box, -originY, -originX)).ToArray();
            var relativeRows = ShiftTableRows(absoluteNestedRows, -originY, -originX);
            entries[new MemoKey(cell, cellWidth, measurer, defaultTabStopPoints, rowTopPadding, paragraphSpacingScale, nestedPageHeight)] =
                new StoredLines(relative, relativePlaced, relativeBoxes, relativeRows, usedHeight,
                    ResolveCounterInputs(pageNumber, pageCount, dependencies),
                    DocxPlacedDrawingAnchor.TranslateY(absoluteAnchors ?? [], -originY));
        }

        private static (int? PageNumber, int? PageCount) ResolveCounterInputs(
            int? pageNumber, int? pageCount, DocxCounterDependencies dependencies) =>
            ((dependencies & DocxCounterDependencies.PageNumber) != 0 ? pageNumber : null,
             (dependencies & DocxCounterDependencies.PageCount) != 0 ? pageCount : null);

        private readonly struct MemoKey : IEquatable<MemoKey>
        {
            private readonly DocxTableCell cell;
            private readonly double cellWidth;
            private readonly IDocxTextMeasurer? measurer;
            private readonly double defaultTabStopPoints;
            private readonly double rowTopPadding;
            private readonly double paragraphSpacingScale;
            private readonly double nestedPageHeight;

            public MemoKey(
                DocxTableCell cell,
                double cellWidth,
                IDocxTextMeasurer? measurer,
                double defaultTabStopPoints,
                double rowTopPadding,
                double paragraphSpacingScale,
                double nestedPageHeight)
            {
                this.cell = cell;
                this.cellWidth = cellWidth;
                this.measurer = measurer;
                this.defaultTabStopPoints = defaultTabStopPoints;
                this.rowTopPadding = rowTopPadding;
                this.paragraphSpacingScale = paragraphSpacingScale;
                this.nestedPageHeight = nestedPageHeight;
            }

            public bool Equals(MemoKey other)
            {
                return ReferenceEquals(cell, other.cell)
                    && cellWidth.Equals(other.cellWidth)
                    && ReferenceEquals(measurer, other.measurer)
                    && defaultTabStopPoints.Equals(other.defaultTabStopPoints)
                    && rowTopPadding.Equals(other.rowTopPadding)
                    && paragraphSpacingScale.Equals(other.paragraphSpacingScale)
                    && nestedPageHeight.Equals(other.nestedPageHeight);
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
                hash.Add(nestedPageHeight);
                return hash.ToHashCode();
            }
        }

        private sealed class StoredLines(IReadOnlyList<DocxTextLineLayout> lines, IReadOnlyList<DocxInlineImageLayout> placedImages,
            IReadOnlyList<DocxInlineTextBoxLayout> placedBoxes, IReadOnlyList<DocxTableRowLayout> nestedRows, double usedHeight,
            (int? PageNumber, int? PageCount) counterInputs, IReadOnlyList<DocxPlacedDrawingAnchor> drawingAnchors)
        {
            public (int? PageNumber, int? PageCount) CounterInputs { get; } = counterInputs;

            public IReadOnlyList<DocxInlineTextBoxLayout> PlacedBoxes { get; } = placedBoxes;
            public IReadOnlyList<DocxTextLineLayout> Lines { get; } = lines;

            public IReadOnlyList<DocxInlineImageLayout> PlacedImages { get; } = placedImages;

            public double UsedHeight { get; } = usedHeight;
            public IReadOnlyList<DocxTableRowLayout> NestedRows { get; } = nestedRows;
            public IReadOnlyList<DocxPlacedDrawingAnchor> DrawingAnchors { get; } = drawingAnchors;
        }
    }
}
