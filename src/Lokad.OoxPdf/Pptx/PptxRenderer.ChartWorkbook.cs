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
}
