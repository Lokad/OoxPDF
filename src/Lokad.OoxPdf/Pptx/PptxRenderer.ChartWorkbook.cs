using System.Globalization;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

// SpreadsheetML workbook data model backing chart series resolution.
// First split from the chart renderer file: pure data types, no drawing.
internal sealed partial class PptxRenderer
{
    internal readonly record struct ChartWorkbookCell(
        string Text,
        string RawValue,
        bool HasValue,
        bool HasValueElement,
        int? SharedStringIndex,
        int SharedStringRunCount,
        bool SharedStringHasRichText,
        bool SharedStringHasPhoneticText,
        bool SharedStringPreserveSpace,
        int? StyleIndex,
        string CellType,
        ChartWorkbookCellValueKind ValueKind,
        string Formula,
        string FormulaType,
        IReadOnlyDictionary<string, string> FormulaAttributes);

    private readonly record struct ChartWorkbookSharedString(
        string Text,
        int RunCount,
        bool HasRichText,
        bool HasPhoneticText,
        bool PreserveSpace);

    internal enum ChartWorkbookCellValueKind
    {
        Blank,
        Number,
        SharedString,
        InlineString,
        FormulaString,
        Boolean,
        Error,
        Other
    }

    internal sealed class ChartWorksheetData
    {
        public ChartWorksheetData(
            Dictionary<string, ChartWorkbookCell> cells,
            IReadOnlySet<int> hiddenRows,
            IReadOnlySet<int> hiddenColumns)
        {
            Cells = cells;
            HiddenRows = hiddenRows;
            HiddenColumns = hiddenColumns;
        }

        public Dictionary<string, ChartWorkbookCell> Cells { get; }

        public IReadOnlySet<int> HiddenRows { get; }

        public IReadOnlySet<int> HiddenColumns { get; }
    }

    internal readonly record struct ChartWorkbookCellFormat(
        int? NumberFormatId,
        string NumberFormatCode,
        bool? ApplyNumberFormat,
        bool NumberFormatIsDateLike);

    internal readonly record struct ChartWorkbookTable(
        string Name,
        string DisplayName,
        string SheetName,
        int FirstColumn,
        int FirstRow,
        int LastColumn,
        int LastRow,
        IReadOnlyList<string> ColumnNames,
        IReadOnlyList<ChartWorkbookTableColumn> Columns,
        int HeaderRowCount,
        int TotalsRowCount,
        bool TotalsRowShown,
        string AutoFilterReference,
        IReadOnlyList<int> FilterColumnIds,
        IReadOnlyList<ChartWorkbookTableFilterColumn> FilterColumns);

    internal readonly record struct ChartWorkbookTableFilterColumn(
        int? ColumnId,
        bool? HiddenButton,
        bool? ShowButton,
        string FilterKind,
        IReadOnlyList<string> FilterValues,
        string DynamicFilterType,
        string Top10Value,
        bool? Top10Percent,
        bool? Top10Top,
        IReadOnlyList<ChartWorkbookTableCustomFilter> CustomFilters);

    internal readonly record struct ChartWorkbookTableCustomFilter(
        string Operator,
        string Value);

    internal readonly record struct ChartWorkbookTableColumn(
        int? Id,
        string Name,
        string TotalsRowFunction,
        string TotalsRowFormula,
        string CalculatedColumnFormula);

    private readonly record struct ChartWorkbookStructuredReferenceBody(
        string FirstColumnName,
        string LastColumnName,
        bool IncludeHeader,
        bool OnlyHeader,
        bool OnlyTotals,
        bool WholeTable);

    internal readonly record struct ChartWorkbookCalculationProperties(
        string CalculationMode,
        string CalculationId,
        bool FullCalculationOnLoad,
        bool ForceFullCalculation);

    internal readonly record struct ChartWorkbookDefinedName(
        string Name,
        string Formula,
        int? LocalSheetId,
        string SheetName);

    internal readonly record struct ChartWorkbookSheet(
        string Name,
        string SheetId,
        string RelationshipId,
        string State,
        int Index,
        string TargetPartName);

    internal sealed class ChartWorkbookStyles
    {
        public ChartWorkbookStyles(
            IReadOnlyDictionary<int, string> customNumberFormats,
            IReadOnlyList<ChartWorkbookCellFormat> cellFormats)
        {
            CustomNumberFormats = customNumberFormats;
            CellFormats = cellFormats;
        }

        public static ChartWorkbookStyles Empty { get; } = new(new Dictionary<int, string>(), []);

        public IReadOnlyDictionary<int, string> CustomNumberFormats { get; }

        public IReadOnlyList<ChartWorkbookCellFormat> CellFormats { get; }

        public ChartWorkbookCellFormat ResolveCellFormat(int? styleIndex)
        {
            return styleIndex is { } index && index >= 0 && index < CellFormats.Count
                ? CellFormats[index]
                : default;
        }
    }

    internal readonly record struct ChartWorkbookRangeCell(
        int Index,
        int RangeAreaIndex,
        int RangeAreaCount,
        int RangeRowIndex,
        int RangeColumnIndex,
        int RangeRowCount,
        int RangeColumnCount,
        int SheetRow,
        int SheetColumn,
        string Reference,
        string SheetName,
        string SourceFormula,
        string ResolvedFormula,
        ChartWorkbookRangeSourceKind SourceKind,
        string DefinedName,
        string DefinedNameSheetName,
        int? DefinedNameLocalSheetId,
        string TableName,
        string TableColumnName,
        int? TableColumnId,
        string TableFirstColumnName,
        int? TableFirstColumnId,
        string TableLastColumnName,
        int? TableLastColumnId,
        int? TableRowIndex,
        bool TableHeaderRow,
        bool TableDataRow,
        bool TableTotalsRow,
        int? TableCellColumnIndex,
        string TableCellColumnName,
        int? TableCellColumnId,
        string TableCellTotalsRowFunction,
        string TableCellTotalsRowFormula,
        string TableCellCalculatedColumnFormula,
        string Text,
        string RawValue,
        bool HasCell,
        bool HasValue,
        bool HasValueElement,
        int? SharedStringIndex,
        int SharedStringRunCount,
        bool SharedStringHasRichText,
        bool SharedStringHasPhoneticText,
        bool SharedStringPreserveSpace,
        int? StyleIndex,
        string CellType,
        ChartWorkbookCellValueKind ValueKind,
        string Formula,
        string FormulaType,
        IReadOnlyDictionary<string, string> FormulaAttributes,
        int? StyleNumberFormatId,
        string StyleNumberFormatCode,
        bool? StyleAppliesNumberFormat,
        bool StyleNumberFormatIsDateLike,
        bool RowHidden,
        bool ColumnHidden);

    internal enum ChartWorkbookRangeSourceKind
    {
        Unknown,
        DirectRange,
        DefinedName,
        StructuredReference,
        DefinedNameStructuredReference
    }

    private readonly record struct ChartWorkbookRangeResolution(
        string SourceFormula,
        string ResolvedFormula,
        ChartWorkbookRangeSourceKind SourceKind,
        string DefinedName,
        string DefinedNameSheetName,
        int? DefinedNameLocalSheetId,
        string TableName,
        string TableColumnName,
        int? TableColumnId,
        string TableFirstColumnName,
        int? TableFirstColumnId,
        string TableLastColumnName,
        int? TableLastColumnId);

    internal readonly record struct ChartWorkbookNumericValue(
        ChartWorkbookRangeCell Cell,
        double Value);

    internal readonly record struct ChartWorkbookTextValue(
        ChartWorkbookRangeCell Cell,
        string Text);

    // R16: internal so the render-context workbook cache carries the model type; members stay as-is.
    internal sealed partial class ChartWorkbookData
    {
        private readonly IReadOnlyDictionary<string, ChartWorksheetData> sheets;
        private readonly ChartWorkbookStyles styles;
        private readonly IReadOnlyDictionary<string, string> definedNames;
        private readonly IReadOnlyList<ChartWorkbookDefinedName> definedNameRecords;
        private readonly IReadOnlyList<ChartWorkbookSheet> sheetRecords;
        private readonly IReadOnlyDictionary<string, ChartWorkbookTable> tables;
        private readonly ChartWorkbookCalculationProperties calculation;

        public ChartWorkbookData(IReadOnlyDictionary<string, Dictionary<string, string>> sheets)
            : this(sheets, date1904: false)
        {
        }

        public ChartWorkbookData(IReadOnlyDictionary<string, Dictionary<string, string>> sheets, bool date1904)
        {
            this.sheets = ConvertWorkbookSheets(sheets);
            styles = ChartWorkbookStyles.Empty;
            definedNames = new Dictionary<string, string>();
            definedNameRecords = [];
            sheetRecords = ConvertWorkbookSheetRecords(this.sheets);
            tables = new Dictionary<string, ChartWorkbookTable>();
            calculation = default;
            Date1904 = date1904;
        }

        public ChartWorkbookData(IReadOnlyDictionary<string, Dictionary<string, ChartWorkbookCell>> sheets, bool date1904)
            : this(sheets, date1904, ChartWorkbookStyles.Empty)
        {
        }

        public ChartWorkbookData(IReadOnlyDictionary<string, Dictionary<string, ChartWorkbookCell>> sheets, bool date1904, ChartWorkbookStyles styles)
            : this(ConvertWorkbookSheets(sheets), date1904, styles, new Dictionary<string, string>())
        {
        }

        public ChartWorkbookData(
            IReadOnlyDictionary<string, ChartWorksheetData> sheets,
            bool date1904,
            ChartWorkbookStyles styles,
            IReadOnlyDictionary<string, string> definedNames)
            : this(sheets, date1904, styles, definedNames, new Dictionary<string, ChartWorkbookTable>())
        {
        }

        public ChartWorkbookData(
            IReadOnlyDictionary<string, ChartWorksheetData> sheets,
            bool date1904,
            ChartWorkbookStyles styles,
            IReadOnlyDictionary<string, string> definedNames,
            IReadOnlyDictionary<string, ChartWorkbookTable> tables)
            : this(sheets, date1904, styles, definedNames, tables, default)
        {
        }

        public ChartWorkbookData(
            IReadOnlyDictionary<string, ChartWorksheetData> sheets,
            bool date1904,
            ChartWorkbookStyles styles,
            IReadOnlyDictionary<string, string> definedNames,
            IReadOnlyDictionary<string, ChartWorkbookTable> tables,
            ChartWorkbookCalculationProperties calculation)
            : this(sheets, date1904, styles, definedNames, tables, calculation, [])
        {
        }

        public ChartWorkbookData(
            IReadOnlyDictionary<string, ChartWorksheetData> sheets,
            bool date1904,
            ChartWorkbookStyles styles,
            IReadOnlyDictionary<string, string> definedNames,
            IReadOnlyDictionary<string, ChartWorkbookTable> tables,
            ChartWorkbookCalculationProperties calculation,
            IReadOnlyList<ChartWorkbookDefinedName> definedNameRecords)
            : this(sheets, date1904, styles, definedNames, tables, calculation, definedNameRecords, ConvertWorkbookSheetRecords(sheets))
        {
        }

        public ChartWorkbookData(
            IReadOnlyDictionary<string, ChartWorksheetData> sheets,
            bool date1904,
            ChartWorkbookStyles styles,
            IReadOnlyDictionary<string, string> definedNames,
            IReadOnlyDictionary<string, ChartWorkbookTable> tables,
            ChartWorkbookCalculationProperties calculation,
            IReadOnlyList<ChartWorkbookDefinedName> definedNameRecords,
            IReadOnlyList<ChartWorkbookSheet> sheetRecords)
        {
            this.sheets = sheets;
            this.styles = styles;
            this.definedNames = definedNames;
            this.definedNameRecords = definedNameRecords;
            this.sheetRecords = sheetRecords;
            this.tables = tables;
            this.calculation = calculation;
            Date1904 = date1904;
        }

        public bool Date1904 { get; }

        public IReadOnlyDictionary<string, string> DefinedNames => definedNames;

        public IReadOnlyList<ChartWorkbookDefinedName> DefinedNameRecords => definedNameRecords;

        public IReadOnlyList<ChartWorkbookSheet> Sheets => sheetRecords;

        public IReadOnlyDictionary<string, ChartWorkbookTable> Tables => tables;

        public ChartWorkbookCalculationProperties Calculation => calculation;

        public ChartWorkbookNumericValue[] ReadNumericRange(string? formula)
        {
            // PLAN W05: the previous Where+Select parsed every value twice. Single pass:
            // TryParse is deterministic on identical input, so results are unchanged.
            var values = new List<ChartWorkbookNumericValue>();
            foreach (ChartWorkbookRangeCell cell in ReadRangeCells(formula))
            {
                if (double.TryParse(cell.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out double value))
                {
                    values.Add(new ChartWorkbookNumericValue(cell, value));
                }
            }

            return values.ToArray();
        }

        public ChartWorkbookTextValue[] ReadTextRange(string? formula)
        {
            return ReadRangeCells(formula)
                .Where(cell => !string.IsNullOrWhiteSpace(cell.Text))
                .Select(cell => new ChartWorkbookTextValue(cell, cell.Text))
                .ToArray();
        }

        // PLAN W05: numeric/category/title/label paths expand the same ranges repeatedly
        // within one chart frame. Memoize expansions per frame (keyed by the raw formula;
        // resolution is deterministic and visibility filtering happens downstream, so the
        // cached cells are valid for every caller). Memo hits perform no work, so they
        // do not consult the token; only expansions do. Callers never mutate the returned
        // arrays; RenderChartFrame clears the memo when the frame completes, so nothing
        // accumulates across frames.
        private readonly Dictionary<string, ChartWorkbookRangeCell[]> rangeMemo = new(StringComparer.Ordinal);

        // R15: per-frame shared dense category labels. Keyed by the originating plot
        // (or chart element when sceneless) plus visibility, so reserve, emission, and
        // legend passes within one frame densify once; literal-only charts without a
        // workbook keep the direct path. Cleared with the range memo per frame, so
        // nothing accumulates across frames or charts.
        private readonly Dictionary<(object? Source, XElement? ChartElement, bool PlotVisibleOnly), IReadOnlyList<ChartIndexedTextPoint?>> labelMemo = new();

        private readonly Dictionary<(object? Source, XElement? ChartElement), IReadOnlyList<ChartSeriesNameRecord>> seriesNameMemo = new();

        private readonly Dictionary<(object? Source, XElement? ChartElement, bool PlotVisibleOnly), IReadOnlyList<ChartIndexedNumberVector>> seriesVectorsMemo = new();

        // R15: per-frame shared bar value extents. Keyed by the originating plot
        // (or chart element when sceneless) plus grouping and visibility, so the
        // reserve, emission, and crossing passes within one frame densify once.
        // Cleared with the range memo per frame.
        private readonly Dictionary<(object? Source, XElement? ChartElement, PptxSceneChartGrouping Grouping, bool PlotVisibleOnly), ChartValueExtents> valueExtentsMemo = new();
        private readonly Dictionary<(object? Source, XElement? ChartElement, bool Stacked, bool PercentStacked, bool PlotVisibleOnly), ChartValueExtents> lineValueExtentsMemo = new();

        internal void ClearRangeMemo()
        {
            rangeMemo.Clear();
            labelMemo.Clear();
            seriesNameMemo.Clear();
            seriesVectorsMemo.Clear();
            valueExtentsMemo.Clear();
            lineValueExtentsMemo.Clear();
        }

        internal IReadOnlyList<ChartSeriesNameRecord> GetOrAddSeriesNames(
            object? source,
            XElement? chartElement,
            Func<IReadOnlyList<ChartSeriesNameRecord>> factory)
        {
            var key = (source, chartElement);
            if (seriesNameMemo.TryGetValue(key, out IReadOnlyList<ChartSeriesNameRecord>? cached))
            {
                return cached;
            }

            IReadOnlyList<ChartSeriesNameRecord> names = factory();
            seriesNameMemo[key] = names;
            return names;
        }

        internal IReadOnlyList<ChartIndexedNumberVector> GetOrAddSeriesVectors(
            object? source,
            XElement? chartElement,
            bool plotVisibleOnly,
            Func<IReadOnlyList<ChartIndexedNumberVector>> factory)
        {
            var key = (source, chartElement, plotVisibleOnly);
            if (seriesVectorsMemo.TryGetValue(key, out IReadOnlyList<ChartIndexedNumberVector>? cached))
            {
                return cached;
            }

            IReadOnlyList<ChartIndexedNumberVector> vectors = factory();
            seriesVectorsMemo[key] = vectors;
            return vectors;
        }

        internal ChartValueExtents GetOrAddValueExtents(
            object? source,
            XElement? chartElement,
            PptxSceneChartGrouping grouping,
            bool plotVisibleOnly,
            Func<ChartValueExtents> factory)
        {
            var key = (source, chartElement, grouping, plotVisibleOnly);
            if (valueExtentsMemo.TryGetValue(key, out ChartValueExtents cached))
            {
                return cached;
            }

            ChartValueExtents extents = factory();
            valueExtentsMemo[key] = extents;
            return extents;
        }

        internal ChartValueExtents GetOrAddLineValueExtents(
            object? source,
            XElement? chartElement,
            bool stacked,
            bool percentStacked,
            bool plotVisibleOnly,
            Func<ChartValueExtents> factory)
        {
            var key = (source, chartElement, stacked, percentStacked, plotVisibleOnly);
            if (lineValueExtentsMemo.TryGetValue(key, out ChartValueExtents cached))
            {
                return cached;
            }

            ChartValueExtents extents = factory();
            lineValueExtentsMemo[key] = extents;
            return extents;
        }

        internal IReadOnlyList<ChartIndexedTextPoint?> GetOrAddCategoryLabels(
            object? source,
            XElement? chartElement,
            bool plotVisibleOnly,
            Func<IReadOnlyList<ChartIndexedTextPoint?>> factory)
        {
            var key = (source, chartElement, plotVisibleOnly);
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
            return ReadRangeCells(formula, CancellationToken.None);
        }

        public ChartWorkbookRangeCell[] ReadRangeCells(string? formula, CancellationToken cancellationToken)
        {
            string key = formula ?? string.Empty;
            if (rangeMemo.TryGetValue(key, out ChartWorkbookRangeCell[]? cached))
            {
                return cached;
            }

            ChartWorkbookRangeResolution resolution = ResolveRangeFormula(formula);
            string[] rangeAreas = SplitRangeAreas(resolution.ResolvedFormula);
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
                AddRangeAreaCells(values, resolution with { ResolvedFormula = rangeAreas[areaIndex] }, areaIndex, rangeAreas.Length, ref index, cancellationToken);
            }

            ChartWorkbookRangeCell[] result = values.ToArray();
            rangeMemo[key] = result;
            return result;
        }

    }

}
