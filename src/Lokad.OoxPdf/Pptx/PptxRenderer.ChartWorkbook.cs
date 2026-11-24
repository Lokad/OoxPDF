namespace Lokad.OoxPdf.Pptx;

// SpreadsheetML workbook data model backing chart series resolution.
internal sealed partial class PptxRenderer
{
    internal readonly record struct ChartWorkbookCell(string Text, bool HasValue, int? StyleIndex);

    internal sealed class ChartWorksheetData
    {
        private readonly bool[] hiddenColumns;

        public ChartWorksheetData(
            Dictionary<string, ChartWorkbookCell> cells,
            IReadOnlySet<int> hiddenRows,
            bool[] hiddenColumns)
        {
            Cells = cells;
            HiddenRows = hiddenRows;
            this.hiddenColumns = hiddenColumns;
        }

        public Dictionary<string, ChartWorkbookCell> Cells { get; }

        public IReadOnlySet<int> HiddenRows { get; }

        public bool IsColumnHidden(int column) =>
            column > 0 && column < hiddenColumns.Length && hiddenColumns[column];
    }

    internal readonly record struct ChartWorkbookCellFormat(string NumberFormatCode, bool? ApplyNumberFormat);

    internal readonly record struct ChartWorkbookTable(
        string Name,
        string DisplayName,
        string SheetName,
        int FirstColumn,
        int FirstRow,
        int LastColumn,
        int LastRow,
        IReadOnlyList<string> ColumnNames,
        int HeaderRowCount,
        int TotalsRowCount);

    private readonly record struct ChartWorkbookStructuredReferenceBody(
        string FirstColumnName,
        string LastColumnName,
        bool IncludeHeader,
        bool OnlyHeader,
        bool OnlyTotals,
        bool WholeTable);

    internal readonly record struct ChartWorkbookLocalName(
        string Name,
        string Formula,
        string SheetName);

    internal readonly record struct ChartWorkbookSheet(
        string Name,
        string RelationshipId);

    internal sealed class ChartWorkbookStyles
    {
        public ChartWorkbookStyles(IReadOnlyList<ChartWorkbookCellFormat> cellFormats)
        {
            CellFormats = cellFormats;
        }

        public static ChartWorkbookStyles Empty { get; } = new([]);

        private IReadOnlyList<ChartWorkbookCellFormat> CellFormats { get; }

        public ChartWorkbookCellFormat ResolveCellFormat(int? styleIndex)
        {
            return styleIndex is { } index && index >= 0 && index < CellFormats.Count
                ? CellFormats[index]
                : default;
        }
    }

    internal readonly record struct ChartWorkbookRangeCell(
        int Index,
        string Text,
        bool HasValue,
        string StyleNumberFormatCode,
        bool? StyleAppliesNumberFormat,
        bool RowHidden,
        bool ColumnHidden);

    // Conversion-owned workbook model shared by chart frames.
    internal sealed partial class ChartWorkbookData
    {
        private readonly IReadOnlyDictionary<string, ChartWorksheetData> sheets;
        private readonly ChartWorkbookStyles styles;
        private readonly IReadOnlyDictionary<string, string> definedNames;
        private readonly IReadOnlyList<ChartWorkbookLocalName> localNames;
        private readonly IReadOnlyDictionary<string, ChartWorkbookTable> tables;
        // The same conversion owns initial loading and later frame expansions.
        private readonly CancellationToken cancellationToken;

        public ChartWorkbookData(
            IReadOnlyDictionary<string, ChartWorksheetData> sheets,
            ChartWorkbookStyles styles,
            IReadOnlyDictionary<string, string> definedNames,
            IReadOnlyDictionary<string, ChartWorkbookTable> tables,
            IReadOnlyList<ChartWorkbookLocalName> localNames,
            CancellationToken cancellationToken)
        {
            this.sheets = sheets;
            this.styles = styles;
            this.definedNames = definedNames;
            this.localNames = localNames;
            this.tables = tables;
            this.cancellationToken = cancellationToken;
        }

        // numeric/category/title/label paths expand the same ranges repeatedly
        // within one chart frame. Memoize expansions per frame (keyed by the raw formula;
        // resolution is deterministic and visibility filtering happens downstream, so the
        // cached cells are valid for every caller). Memo hits perform no work, so they
        // do not consult the token; only expansions do. Callers never mutate the returned
        // arrays; RenderChartFrame clears the memo when the frame completes, so nothing
        // accumulates across frames.
        private readonly Dictionary<string, ChartWorkbookRangeCell[]> rangeMemo = new(StringComparer.Ordinal);

        // Per-frame shared dense category labels. Keyed by the originating plot
        // plus visibility, so reserve, emission, and
        // legend passes within one frame densify once; literal-only charts without a
        // workbook keep the direct path. Cleared with the range memo per frame, so
        // nothing accumulates across frames or charts.
        private readonly Dictionary<(PptxSceneChartPlot Plot, bool PlotVisibleOnly), IReadOnlyList<ChartIndexedTextPoint?>> labelMemo = new();

        private readonly Dictionary<PptxSceneChartPlot, IReadOnlyList<string>> seriesNameMemo = new();

        private readonly Dictionary<(PptxSceneChartPlot Plot, bool PlotVisibleOnly), IReadOnlyList<ChartIndexedNumberVector>> seriesVectorsMemo = new();

        // Per-frame shared bar value extents. Keyed by the originating plot
        // plus grouping and visibility, so the
        // reserve, emission, and crossing passes within one frame densify once.
        // Cleared with the range memo per frame.
        private readonly Dictionary<(PptxSceneChartPlot Plot, PptxSceneChartGrouping Grouping, bool PlotVisibleOnly), ChartValueExtents> valueExtentsMemo = new();
        private readonly Dictionary<(PptxSceneChartPlot Plot, bool Stacked, bool PercentStacked, bool PlotVisibleOnly), ChartValueExtents> lineValueExtentsMemo = new();

        internal void ClearRangeMemo()
        {
            rangeMemo.Clear();
            labelMemo.Clear();
            seriesNameMemo.Clear();
            seriesVectorsMemo.Clear();
            valueExtentsMemo.Clear();
            lineValueExtentsMemo.Clear();
        }

        internal IReadOnlyList<string> GetOrAddSeriesNames(
            PptxSceneChartPlot plot,
            Func<IReadOnlyList<string>> factory)
        {
            if (seriesNameMemo.TryGetValue(plot, out IReadOnlyList<string>? cached))
            {
                return cached;
            }

            IReadOnlyList<string> names = factory();
            seriesNameMemo[plot] = names;
            return names;
        }

        internal IReadOnlyList<ChartIndexedNumberVector> GetOrAddSeriesVectors(
            PptxSceneChartPlot plot,
            bool plotVisibleOnly,
            Func<IReadOnlyList<ChartIndexedNumberVector>> factory)
        {
            var key = (plot, plotVisibleOnly);
            if (seriesVectorsMemo.TryGetValue(key, out IReadOnlyList<ChartIndexedNumberVector>? cached))
            {
                return cached;
            }

            IReadOnlyList<ChartIndexedNumberVector> vectors = factory();
            seriesVectorsMemo[key] = vectors;
            return vectors;
        }

        internal ChartValueExtents GetOrAddValueExtents(
            PptxSceneChartPlot plot,
            PptxSceneChartGrouping grouping,
            bool plotVisibleOnly,
            Func<ChartValueExtents> factory)
        {
            var key = (plot, grouping, plotVisibleOnly);
            if (valueExtentsMemo.TryGetValue(key, out ChartValueExtents cached))
            {
                return cached;
            }

            ChartValueExtents extents = factory();
            valueExtentsMemo[key] = extents;
            return extents;
        }

        internal ChartValueExtents GetOrAddLineValueExtents(
            PptxSceneChartPlot plot,
            bool stacked,
            bool percentStacked,
            bool plotVisibleOnly,
            Func<ChartValueExtents> factory)
        {
            var key = (plot, stacked, percentStacked, plotVisibleOnly);
            if (lineValueExtentsMemo.TryGetValue(key, out ChartValueExtents cached))
            {
                return cached;
            }

            ChartValueExtents extents = factory();
            lineValueExtentsMemo[key] = extents;
            return extents;
        }

        internal IReadOnlyList<ChartIndexedTextPoint?> GetOrAddCategoryLabels(
            PptxSceneChartPlot plot,
            bool plotVisibleOnly,
            Func<IReadOnlyList<ChartIndexedTextPoint?>> factory)
        {
            var key = (plot, plotVisibleOnly);
            if (labelMemo.TryGetValue(key, out IReadOnlyList<ChartIndexedTextPoint?>? cached))
            {
                return cached;
            }

            IReadOnlyList<ChartIndexedTextPoint?> dense = factory();
            labelMemo[key] = dense;
            return dense;
        }

        public ChartWorkbookRangeCell[] ReadRangeCells(string? formula)
        {
            string key = formula ?? string.Empty;
            if (rangeMemo.TryGetValue(key, out ChartWorkbookRangeCell[]? cached))
            {
                return cached;
            }

            cancellationToken.ThrowIfCancellationRequested();
            string[] rangeAreas = SplitRangeAreas(ResolveRangeFormula(formula));
            if (rangeAreas.Length == 0)
            {
                return [];
            }

            if (rangeAreas.Length > ChartWorkbookData.MaxChartRangeAreas)
            {
                throw new OoxPdfLimitExceededException(
                    "Chart data range exceeds the maximum supported area count of " + ChartWorkbookData.MaxChartRangeAreas + ".");
            }

            var values = new List<ChartWorkbookRangeCell>();
            int index = 0;
            for (int areaIndex = 0; areaIndex < rangeAreas.Length; areaIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                AddRangeAreaCells(values, rangeAreas[areaIndex], ref index, cancellationToken);
            }

            ChartWorkbookRangeCell[] result = values.ToArray();
            rangeMemo[key] = result;
            return result;
        }

    }

}
