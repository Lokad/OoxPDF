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
    private readonly record struct ChartWorkbookCell(
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

    private enum ChartWorkbookCellValueKind
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

    private sealed class ChartWorksheetData
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

    private readonly record struct ChartWorkbookCellFormat(
        int? NumberFormatId,
        string NumberFormatCode,
        bool? ApplyNumberFormat,
        bool NumberFormatIsDateLike);

    private readonly record struct ChartWorkbookTable(
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

    private readonly record struct ChartWorkbookTableFilterColumn(
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

    private readonly record struct ChartWorkbookTableCustomFilter(
        string Operator,
        string Value);

    private readonly record struct ChartWorkbookTableColumn(
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

    private readonly record struct ChartWorkbookCalculationProperties(
        string CalculationMode,
        string CalculationId,
        bool FullCalculationOnLoad,
        bool ForceFullCalculation);

    private readonly record struct ChartWorkbookDefinedName(
        string Name,
        string Formula,
        int? LocalSheetId,
        string SheetName);

    private readonly record struct ChartWorkbookSheet(
        string Name,
        string SheetId,
        string RelationshipId,
        string State,
        int Index,
        string TargetPartName);

    private sealed class ChartWorkbookStyles
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

    private readonly record struct ChartWorkbookRangeCell(
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

    private enum ChartWorkbookRangeSourceKind
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

    private readonly record struct ChartWorkbookNumericValue(
        ChartWorkbookRangeCell Cell,
        double Value);

    private readonly record struct ChartWorkbookTextValue(
        ChartWorkbookRangeCell Cell,
        string Text);

    private sealed partial class ChartWorkbookData
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
            return ReadRangeCells(formula)
                .Where(cell => double.TryParse(cell.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out _))
                .Select(cell =>
                {
                    double.TryParse(cell.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out double value);
                    return new ChartWorkbookNumericValue(cell, value);
                })
                .ToArray();
        }

        public ChartWorkbookTextValue[] ReadTextRange(string? formula)
        {
            return ReadRangeCells(formula)
                .Where(cell => !string.IsNullOrWhiteSpace(cell.Text))
                .Select(cell => new ChartWorkbookTextValue(cell, cell.Text))
                .ToArray();
        }

        public ChartWorkbookRangeCell[] ReadRangeCells(string? formula)
        {
            ChartWorkbookRangeResolution resolution = ResolveRangeFormula(formula);
            string[] rangeAreas = SplitRangeAreas(resolution.ResolvedFormula);
            if (rangeAreas.Length == 0)
            {
                return [];
            }

            var values = new List<ChartWorkbookRangeCell>();
            int index = 0;
            for (int areaIndex = 0; areaIndex < rangeAreas.Length; areaIndex++)
            {
                AddRangeAreaCells(values, resolution with { ResolvedFormula = rangeAreas[areaIndex] }, areaIndex, rangeAreas.Length, ref index);
            }

            return values.ToArray();
        }

    }

}
