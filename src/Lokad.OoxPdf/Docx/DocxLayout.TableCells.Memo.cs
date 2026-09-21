namespace Lokad.OoxPdf.Docx;

// PLAN W04: cell text-line memo. One table layout pass measures each cell up to three
// times (split-feasibility, break resolution, final fragments) plus nested-table
// remeasurement; keys capture every layout input so hits reproduce the computed lines
// exactly after caller-side re-offsetting. Cells carrying PAGE/NUMPAGES runs keep
// their page args in the key (dynamic tier); all other cells omit them (static tier),
// which is sound because page inputs only reach text spans and wrap measurement
// through dynamic field runs. Vertical-alignment shifts stay outside the memo: every
// caller applies its own height-dependent shift to the shared unshifted lines.
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

        // W01 precedent: structural counters proving the memo fires on real corpus.
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
            out IReadOnlyList<DocxTextLineLayout> lines,
            out double originX,
            out double originY,
            out double usedHeight)
        {
            if (entries.TryGetValue(
                new MemoKey(cell, cellWidth, measurer, defaultTabStopPoints, rowTopPadding, paragraphSpacingScale, pageStatic ? null : pageNumber, pageStatic ? null : pageCount),
                out StoredLines? stored) && stored is not null)
            {
                Hits++;
                System.Console.WriteLine("W04HIT w=" + cellWidth + " ox=" + stored.OriginX + " oy=" + stored.OriginY + " n=" + stored.Lines.Count);
                TotalHits++;
                lines = stored.Lines;
                originX = stored.OriginX;
                originY = stored.OriginY;
                usedHeight = stored.UsedHeight;
                Hits++;
            }

            Misses++;
            TotalMisses++;
            lines = [];
            originX = 0d;
            originY = 0d;
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
            IReadOnlyList<DocxTextLineLayout> lines,
            double originX,
            double originY,
            double usedHeight)
        {
            entries[new MemoKey(cell, cellWidth, measurer, defaultTabStopPoints, rowTopPadding, paragraphSpacingScale, pageStatic ? null : pageNumber, pageStatic ? null : pageCount)] = new StoredLines(lines, originX, originY, usedHeight);
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

        // Object-typed cells compare by reference through the default comparer
        // (records would compare by value); reference sharing is the always-sound
        // subset since every other layout input sits in the key.
        private readonly record struct MemoKey(
            object Cell,
            double CellWidth,
            object? Measurer,
            double DefaultTabStopPoints,
            double RowTopPadding,
            double ParagraphSpacingScale,
            int? PageNumber,
            int? PageCount);

        private sealed class StoredLines(IReadOnlyList<DocxTextLineLayout> lines, double originX, double originY, double usedHeight)
        {
            public IReadOnlyList<DocxTextLineLayout> Lines { get; } = lines;

            public double OriginX { get; } = originX;

            public double OriginY { get; } = originY;

        public double UsedHeight { get; } = usedHeight;
        }
    }
}
