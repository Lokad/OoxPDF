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

    private sealed class ChartWorkbookData
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

        private void AddRangeAreaCells(
            List<ChartWorkbookRangeCell> values,
            ChartWorkbookRangeResolution resolution,
            int rangeAreaIndex,
            int rangeAreaCount,
            ref int index)
        {
            if (!TryParseRange(resolution.ResolvedFormula, out string? sheetName, out int firstColumn, out int firstRow, out int lastColumn, out int lastRow) ||
                !sheets.TryGetValue(sheetName, out ChartWorksheetData? worksheet))
            {
                return;
            }

            int minColumn = Math.Min(firstColumn, lastColumn);
            int maxColumn = Math.Max(firstColumn, lastColumn);
            int minRow = Math.Min(firstRow, lastRow);
            int maxRow = Math.Max(firstRow, lastRow);
            int rangeRowCount = maxRow - minRow + 1;
            int rangeColumnCount = maxColumn - minColumn + 1;
            ChartWorkbookTable sourceTable = default;
            bool hasSourceTable = !string.IsNullOrWhiteSpace(resolution.TableName) &&
                tables.TryGetValue(resolution.TableName, out sourceTable);
            for (int row = minRow; row <= maxRow; row++)
            {
                for (int column = minColumn; column <= maxColumn; column++)
                {
                    string reference = ToCellReference(column, row);
                    bool hasCell = worksheet.Cells.TryGetValue(reference, out ChartWorkbookCell cell);
                    ChartWorkbookCellFormat format = hasCell ? styles.ResolveCellFormat(cell.StyleIndex) : default;
                    int? tableRowIndex = hasSourceTable ? row - sourceTable.FirstRow : null;
                    bool tableHeaderRow = hasSourceTable && tableRowIndex >= 0 && tableRowIndex < Math.Max(0, sourceTable.HeaderRowCount);
                    bool tableTotalsRow = hasSourceTable && sourceTable.TotalsRowCount > 0 && row > sourceTable.LastRow - sourceTable.TotalsRowCount && row <= sourceTable.LastRow;
                    bool tableDataRow = hasSourceTable && !tableHeaderRow && !tableTotalsRow && row >= sourceTable.FirstRow && row <= sourceTable.LastRow;
                    int? tableCellColumnIndex = null;
                    string tableCellColumnName = string.Empty;
                    int? tableCellColumnId = null;
                    string tableCellTotalsRowFunction = string.Empty;
                    string tableCellTotalsRowFormula = string.Empty;
                    string tableCellCalculatedColumnFormula = string.Empty;
                    if (hasSourceTable)
                    {
                        int columnIndex = column - sourceTable.FirstColumn;
                        if (columnIndex >= 0 && columnIndex < sourceTable.Columns.Count)
                        {
                            ChartWorkbookTableColumn tableCellColumn = sourceTable.Columns[columnIndex];
                            tableCellColumnIndex = columnIndex;
                            tableCellColumnName = tableCellColumn.Name;
                            tableCellColumnId = tableCellColumn.Id;
                            tableCellTotalsRowFunction = tableCellColumn.TotalsRowFunction;
                            tableCellTotalsRowFormula = tableCellColumn.TotalsRowFormula;
                            tableCellCalculatedColumnFormula = tableCellColumn.CalculatedColumnFormula;
                        }
                    }

                    values.Add(new ChartWorkbookRangeCell(
                        index,
                        rangeAreaIndex,
                        rangeAreaCount,
                        row - minRow,
                        column - minColumn,
                        rangeRowCount,
                        rangeColumnCount,
                        row,
                        column,
                        reference,
                        sheetName,
                        resolution.SourceFormula,
                        resolution.ResolvedFormula,
                        resolution.SourceKind,
                        resolution.DefinedName,
                        resolution.DefinedNameSheetName,
                        resolution.DefinedNameLocalSheetId,
                        resolution.TableName,
                        resolution.TableColumnName,
                        resolution.TableColumnId,
                        resolution.TableFirstColumnName,
                        resolution.TableFirstColumnId,
                        resolution.TableLastColumnName,
                        resolution.TableLastColumnId,
                        tableRowIndex,
                        tableHeaderRow,
                        tableDataRow,
                        tableTotalsRow,
                        tableCellColumnIndex,
                        tableCellColumnName,
                        tableCellColumnId,
                        tableCellTotalsRowFunction,
                        tableCellTotalsRowFormula,
                        tableCellCalculatedColumnFormula,
                        hasCell ? cell.Text : string.Empty,
                        hasCell ? cell.RawValue : string.Empty,
                        hasCell,
                        hasCell && cell.HasValue,
                        hasCell && cell.HasValueElement,
                        hasCell ? cell.SharedStringIndex : null,
                        hasCell ? cell.SharedStringRunCount : 0,
                        hasCell && cell.SharedStringHasRichText,
                        hasCell && cell.SharedStringHasPhoneticText,
                        hasCell && cell.SharedStringPreserveSpace,
                        hasCell ? cell.StyleIndex : null,
                        hasCell ? cell.CellType : string.Empty,
                        hasCell ? cell.ValueKind : ChartWorkbookCellValueKind.Blank,
                        hasCell ? cell.Formula : string.Empty,
                        hasCell ? cell.FormulaType : string.Empty,
                        hasCell ? cell.FormulaAttributes : new Dictionary<string, string>(),
                        format.NumberFormatId,
                        format.NumberFormatCode,
                        format.ApplyNumberFormat,
                        format.NumberFormatIsDateLike,
                        worksheet.HiddenRows.Contains(row),
                        worksheet.HiddenColumns.Contains(column)));
                    index++;
                }
            }
        }

        private ChartWorkbookRangeResolution ResolveRangeFormula(string? formula)
        {
            string sourceFormula = formula?.Trim() ?? string.Empty;
            if (sourceFormula.Length == 0)
            {
                return new ChartWorkbookRangeResolution(string.Empty, string.Empty, ChartWorkbookRangeSourceKind.Unknown, string.Empty, string.Empty, null, string.Empty, string.Empty, null, string.Empty, null, string.Empty, null);
            }

            string resolvedFormula = sourceFormula;
            string definedNameName = string.Empty;
            string definedNameSheetName = string.Empty;
            int? definedNameLocalSheetId = null;
            bool fromDefinedName = false;
            if (!sourceFormula.Contains('!', StringComparison.Ordinal) &&
                definedNames.TryGetValue(sourceFormula, out string? definedFormula))
            {
                resolvedFormula = definedFormula;
                definedNameName = sourceFormula;
                fromDefinedName = true;
            }
            else if (TryResolveQualifiedLocalDefinedName(sourceFormula, out ChartWorkbookDefinedName localDefinedName))
            {
                resolvedFormula = localDefinedName.Formula;
                definedNameName = localDefinedName.Name;
                definedNameSheetName = localDefinedName.SheetName;
                definedNameLocalSheetId = localDefinedName.LocalSheetId;
                fromDefinedName = true;
            }

            string structuredResolvedFormula = ResolveStructuredReferenceFormula(
                resolvedFormula,
                out string tableName,
                out string tableColumnName,
                out int? tableColumnId,
                out string tableFirstColumnName,
                out int? tableFirstColumnId,
                out string tableLastColumnName,
                out int? tableLastColumnId) ?? string.Empty;
            bool fromStructuredReference = !string.Equals(structuredResolvedFormula, resolvedFormula, StringComparison.Ordinal);
            resolvedFormula = structuredResolvedFormula;

            ChartWorkbookRangeSourceKind sourceKind = (fromDefinedName, fromStructuredReference) switch
            {
                (true, true) => ChartWorkbookRangeSourceKind.DefinedNameStructuredReference,
                (true, false) => ChartWorkbookRangeSourceKind.DefinedName,
                (false, true) => ChartWorkbookRangeSourceKind.StructuredReference,
                _ => ChartWorkbookRangeSourceKind.DirectRange
            };
            return new ChartWorkbookRangeResolution(
                sourceFormula,
                resolvedFormula,
                sourceKind,
                definedNameName,
                definedNameSheetName,
                definedNameLocalSheetId,
                tableName,
                tableColumnName,
                tableColumnId,
                tableFirstColumnName,
                tableFirstColumnId,
                tableLastColumnName,
                tableLastColumnId);
        }

        private bool TryResolveQualifiedLocalDefinedName(string sourceFormula, out ChartWorkbookDefinedName definedName)
        {
            definedName = default;
            int separator = sourceFormula.LastIndexOf('!');
            if (separator <= 0 || separator == sourceFormula.Length - 1)
            {
                return false;
            }

            string sheetName = NormalizeSheetName(sourceFormula[..separator]);
            string name = sourceFormula[(separator + 1)..].Trim();
            if (sheetName.Length == 0 ||
                name.Length == 0 ||
                name.Contains('$', StringComparison.Ordinal) ||
                name.Contains(':', StringComparison.Ordinal))
            {
                return false;
            }

            foreach (ChartWorkbookDefinedName candidate in definedNameRecords)
            {
                if (candidate.LocalSheetId is not null &&
                    string.Equals(candidate.SheetName, sheetName, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(candidate.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    definedName = candidate;
                    return true;
                }
            }

            return false;
        }

        private string? ResolveStructuredReferenceFormula(
            string? formula,
            out string tableName,
            out string tableColumnName,
            out int? tableColumnId,
            out string tableFirstColumnName,
            out int? tableFirstColumnId,
            out string tableLastColumnName,
            out int? tableLastColumnId)
        {
            tableName = string.Empty;
            tableColumnName = string.Empty;
            tableColumnId = null;
            tableFirstColumnName = string.Empty;
            tableFirstColumnId = null;
            tableLastColumnName = string.Empty;
            tableLastColumnId = null;
            if (string.IsNullOrWhiteSpace(formula))
            {
                return formula;
            }

            string trimmed = formula.Trim();
            int open = trimmed.IndexOf('[', StringComparison.Ordinal);
            if (open <= 0 || trimmed[^1] != ']')
            {
                return trimmed;
            }

            string candidateTableName = trimmed[..open];
            if (!tables.TryGetValue(candidateTableName, out ChartWorkbookTable table))
            {
                return trimmed;
            }

            string body = trimmed[(open + 1)..^1];
            if (!TryParseStructuredReferenceBody(
                body,
                out string? firstColumnName,
                out string? lastColumnName,
                out bool includeHeader,
                out bool onlyHeader,
                out bool onlyTotals,
                out bool wholeTable))
            {
                return trimmed;
            }

            int firstColumn = table.FirstColumn;
            int lastColumn = table.LastColumn;
            ChartWorkbookTableColumn tableColumn = default;
            ChartWorkbookTableColumn firstTableColumn = table.Columns.Count > 0 ? table.Columns[0] : default;
            ChartWorkbookTableColumn lastTableColumn = table.Columns.Count > 0 ? table.Columns[^1] : default;
            if (!wholeTable)
            {
                if (string.IsNullOrWhiteSpace(firstColumnName))
                {
                    return trimmed;
                }

                lastColumnName = string.IsNullOrWhiteSpace(lastColumnName) ? firstColumnName : lastColumnName;
                int firstColumnOffset = -1;
                int lastColumnOffset = -1;
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    if (string.Equals(table.Columns[i].Name, firstColumnName, StringComparison.OrdinalIgnoreCase))
                    {
                        firstColumnOffset = i;
                    }

                    if (string.Equals(table.Columns[i].Name, lastColumnName, StringComparison.OrdinalIgnoreCase))
                    {
                        lastColumnOffset = i;
                    }
                }

                if (firstColumnOffset < 0 || lastColumnOffset < 0)
                {
                    return trimmed;
                }

                int minColumnOffset = Math.Min(firstColumnOffset, lastColumnOffset);
                int maxColumnOffset = Math.Max(firstColumnOffset, lastColumnOffset);
                firstColumn = table.FirstColumn + minColumnOffset;
                lastColumn = table.FirstColumn + maxColumnOffset;
                firstTableColumn = table.Columns[minColumnOffset];
                lastTableColumn = table.Columns[maxColumnOffset];
                if (minColumnOffset == maxColumnOffset)
                {
                    tableColumn = firstTableColumn;
                }
            }

            int headerRowCount = Math.Max(0, table.HeaderRowCount);
            int totalsRowCount = Math.Max(0, table.TotalsRowCount);
            int firstRow = onlyHeader
                ? table.FirstRow
                : onlyTotals
                    ? table.LastRow - totalsRowCount + 1
                    : includeHeader
                        ? table.FirstRow
                        : table.FirstRow + headerRowCount;
            int lastRow = onlyHeader
                ? table.FirstRow + headerRowCount - 1
                : onlyTotals
                    ? table.LastRow
                    : includeHeader
                        ? table.LastRow
                        : table.LastRow - totalsRowCount;
            if (firstRow > lastRow)
            {
                return trimmed;
            }

            tableName = table.Name;
            bool singleColumn = !wholeTable && firstColumn == lastColumn;
            tableColumnName = singleColumn ? tableColumn.Name : string.Empty;
            tableColumnId = singleColumn ? tableColumn.Id : null;
            tableFirstColumnName = firstTableColumn.Name;
            tableFirstColumnId = firstTableColumn.Id;
            tableLastColumnName = lastTableColumn.Name;
            tableLastColumnId = lastTableColumn.Id;
            return FormattableString.Invariant($"{QuoteSheetName(table.SheetName)}!{ToCellReference(firstColumn, firstRow)}:{ToCellReference(lastColumn, lastRow)}");
        }

        private static bool TryParseStructuredReferenceBody(
            string body,
            out string firstColumnName,
            out string lastColumnName,
            out bool includeHeader,
            out bool onlyHeader,
            out bool onlyTotals,
            out bool wholeTable)
        {
            firstColumnName = string.Empty;
            lastColumnName = string.Empty;
            includeHeader = false;
            onlyHeader = false;
            onlyTotals = false;
            wholeTable = false;
            string trimmed = body.Trim();
            if (trimmed.Length == 0)
            {
                return false;
            }

            if (trimmed[0] != '[')
            {
                string segment = UnescapeStructuredReferenceText(trimmed);
                if (ApplyStructuredReferenceItem(segment, ref includeHeader, ref onlyHeader, ref onlyTotals, ref wholeTable))
                {
                    return true;
                }

                firstColumnName = segment;
                lastColumnName = segment;
                return true;
            }

            foreach (string segment in ParseStructuredReferenceSegments(trimmed))
            {
                if (ApplyStructuredReferenceItem(segment, ref includeHeader, ref onlyHeader, ref onlyTotals, ref wholeTable))
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(firstColumnName))
                {
                    firstColumnName = segment;
                }

                lastColumnName = segment;
                wholeTable = false;
            }

            return wholeTable || !string.IsNullOrWhiteSpace(firstColumnName);
        }

        private static bool ApplyStructuredReferenceItem(string segment, ref bool includeHeader, ref bool onlyHeader, ref bool onlyTotals, ref bool wholeTable)
        {
            if (string.Equals(segment, "#All", StringComparison.OrdinalIgnoreCase))
            {
                includeHeader = true;
                wholeTable = true;
                return true;
            }

            if (string.Equals(segment, "#Headers", StringComparison.OrdinalIgnoreCase))
            {
                includeHeader = true;
                onlyHeader = true;
                wholeTable = true;
                return true;
            }

            if (string.Equals(segment, "#Data", StringComparison.OrdinalIgnoreCase))
            {
                wholeTable = true;
                return true;
            }

            if (string.Equals(segment, "#Totals", StringComparison.OrdinalIgnoreCase))
            {
                onlyTotals = true;
                wholeTable = true;
                return true;
            }

            return false;
        }

        private static string[] ParseStructuredReferenceSegments(string body)
        {
            var segments = new List<string>();
            int index = 0;
            while (index < body.Length)
            {
                while (index < body.Length && (char.IsWhiteSpace(body[index]) || body[index] == ',' || body[index] == ':'))
                {
                    index++;
                }

                if (index >= body.Length)
                {
                    break;
                }

                if (body[index] != '[')
                {
                    int nextComma = body.IndexOf(',', index);
                    string segment = nextComma < 0 ? body[index..] : body[index..nextComma];
                    segments.Add(UnescapeStructuredReferenceText(segment.Trim()));
                    index = nextComma < 0 ? body.Length : nextComma + 1;
                    continue;
                }

                index++;
                var segmentBuilder = new StringBuilder();
                while (index < body.Length)
                {
                    char current = body[index];
                    if (current == '\'' && index + 1 < body.Length)
                    {
                        segmentBuilder.Append(body[index + 1]);
                        index += 2;
                        continue;
                    }

                    if (current == ']')
                    {
                        index++;
                        break;
                    }

                    segmentBuilder.Append(current);
                    index++;
                }

                segments.Add(segmentBuilder.ToString().Trim());
            }

            return segments.ToArray();
        }

        private static string UnescapeStructuredReferenceText(string text)
        {
            var builder = new StringBuilder(text.Length);
            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] == '\'' && i + 1 < text.Length)
                {
                    builder.Append(text[i + 1]);
                    i++;
                }
                else
                {
                    builder.Append(text[i]);
                }
            }

            return builder.ToString().Trim();
        }

        private static string QuoteSheetName(string sheetName)
        {
            return sheetName.Any(character => !char.IsAsciiLetterOrDigit(character) && character != '_')
                ? "'" + sheetName.Replace("'", "''", StringComparison.Ordinal) + "'"
                : sheetName;
        }

        private static IReadOnlyDictionary<string, ChartWorksheetData> ConvertWorkbookSheets(
            IReadOnlyDictionary<string, Dictionary<string, string>> sheets)
        {
            var converted = new Dictionary<string, ChartWorksheetData>(StringComparer.OrdinalIgnoreCase);
            foreach (KeyValuePair<string, Dictionary<string, string>> sheet in sheets)
            {
                var cells = new Dictionary<string, ChartWorkbookCell>(StringComparer.OrdinalIgnoreCase);
                foreach (KeyValuePair<string, string> cell in sheet.Value)
                {
                    cells[cell.Key] = new ChartWorkbookCell(cell.Value, cell.Value, true, true, null, 0, false, false, false, null, string.Empty, ReadWorkbookCellValueKind(null, cell.Value, true), string.Empty, string.Empty, new Dictionary<string, string>());
                }

                converted[sheet.Key] = new ChartWorksheetData(cells, new HashSet<int>(), new HashSet<int>());
            }

            return converted;
        }

        private static ChartWorkbookSheet[] ConvertWorkbookSheetRecords(IReadOnlyDictionary<string, ChartWorksheetData> sheets)
        {
            return sheets.Keys
                .Select((name, index) => new ChartWorkbookSheet(name, string.Empty, string.Empty, string.Empty, index, string.Empty))
                .ToArray();
        }

        private static IReadOnlyDictionary<string, ChartWorksheetData> ConvertWorkbookSheets(
            IReadOnlyDictionary<string, Dictionary<string, ChartWorkbookCell>> sheets)
        {
            var converted = new Dictionary<string, ChartWorksheetData>(StringComparer.OrdinalIgnoreCase);
            foreach (KeyValuePair<string, Dictionary<string, ChartWorkbookCell>> sheet in sheets)
            {
                converted[sheet.Key] = new ChartWorksheetData(sheet.Value, new HashSet<int>(), new HashSet<int>());
            }

            return converted;
        }

        private static bool TryParseRange(string? formula, out string sheetName, out int firstColumn, out int firstRow, out int lastColumn, out int lastRow)
        {
            sheetName = string.Empty;
            firstColumn = 0;
            firstRow = 0;
            lastColumn = 0;
            lastRow = 0;
            if (string.IsNullOrWhiteSpace(formula))
            {
                return false;
            }

            string trimmed = formula.Trim();
            int separator = trimmed.LastIndexOf('!');
            if (separator <= 0 || separator == trimmed.Length - 1)
            {
                return false;
            }

            sheetName = NormalizeSheetName(trimmed[..separator]);
            string[] references = trimmed[(separator + 1)..].Split(':', 2, StringSplitOptions.TrimEntries);
            if (!TryParseCellReference(references[0], out firstColumn, out firstRow))
            {
                return false;
            }

            if (references.Length == 1)
            {
                lastColumn = firstColumn;
                lastRow = firstRow;
                return true;
            }

            return TryParseCellReference(references[1], out lastColumn, out lastRow);
        }

        private static string[] SplitRangeAreas(string formula)
        {
            if (string.IsNullOrWhiteSpace(formula))
            {
                return [];
            }

            var areas = new List<string>();
            int areaStart = 0;
            bool inQuotedSheetName = false;
            for (int index = 0; index < formula.Length; index++)
            {
                char current = formula[index];
                if (current == '\'')
                {
                    if (inQuotedSheetName && index + 1 < formula.Length && formula[index + 1] == '\'')
                    {
                        index++;
                        continue;
                    }

                    inQuotedSheetName = !inQuotedSheetName;
                    continue;
                }

                if (current == ',' && !inQuotedSheetName)
                {
                    string area = formula[areaStart..index].Trim();
                    if (area.Length > 0)
                    {
                        areas.Add(area);
                    }

                    areaStart = index + 1;
                }
            }

            string lastArea = formula[areaStart..].Trim();
            if (lastArea.Length > 0)
            {
                areas.Add(lastArea);
            }

            string currentSheetPrefix = string.Empty;
            for (int index = 0; index < areas.Count; index++)
            {
                int separator = areas[index].LastIndexOf('!');
                if (separator > 0)
                {
                    currentSheetPrefix = areas[index][..(separator + 1)];
                }
                else if (currentSheetPrefix.Length > 0)
                {
                    areas[index] = currentSheetPrefix + areas[index];
                }
            }

            return areas.ToArray();
        }

        private static string NormalizeSheetName(string sheetName)
        {
            sheetName = sheetName.Trim();
            if (sheetName.Length >= 2 && sheetName[0] == '\'' && sheetName[^1] == '\'')
            {
                sheetName = sheetName[1..^1].Replace("''", "'", StringComparison.Ordinal);
            }

            int workbookEnd = sheetName.LastIndexOf(']');
            if (workbookEnd >= 0 && workbookEnd < sheetName.Length - 1)
            {
                sheetName = sheetName[(workbookEnd + 1)..];
            }

            return sheetName;
        }

        private static bool TryParseCellReference(string reference, out int column, out int row)
        {
            column = 0;
            row = 0;
            string normalized = reference.Replace("$", string.Empty, StringComparison.Ordinal).Trim();
            int index = 0;
            while (index < normalized.Length && char.IsAsciiLetter(normalized[index]))
            {
                column = (column * 26) + (char.ToUpperInvariant(normalized[index]) - 'A' + 1);
                index++;
            }

            if (column <= 0 || index == normalized.Length)
            {
                return false;
            }

            return int.TryParse(normalized[index..], NumberStyles.Integer, CultureInfo.InvariantCulture, out row) && row > 0;
        }

        private static string ToCellReference(int column, int row)
        {
            Span<char> buffer = stackalloc char[16];
            int position = buffer.Length;
            int value = column;
            while (value > 0)
            {
                value--;
                buffer[--position] = (char)('A' + (value % 26));
                value /= 26;
            }

            return string.Concat(new string(buffer[position..]), row.ToString(CultureInfo.InvariantCulture));
        }
    }

    private static ChartWorkbookData? ReadEmbeddedChartWorkbookData(PptxSceneChartExternalData sceneExternalData, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (!sceneExternalData.IsDefined ||
            sceneExternalData.Resource is null)
        {
            return null;
        }

        using var stream = new MemoryStream(sceneExternalData.Resource.Bytes, writable: false);
        OoxPackage workbookPackage = OoxPackage.Open(stream, cancellationToken);
        return ReadWorkbookDataCore(workbookPackage, cancellationToken);
    }

    private static ChartWorkbookData? ReadWorkbookData(OoxPackage workbookPackage)
    {
        return ReadWorkbookDataCore(workbookPackage, CancellationToken.None);
    }

    private static ChartWorkbookData? ReadWorkbookDataCore(OoxPackage workbookPackage, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        OoxPart? workbookPart = workbookPackage
            .GetRelationships("/", cancellationToken)
            .Where(relationship => !relationship.IsExternal && relationship.Type == WorkbookRelationshipType && relationship.ResolvedTarget is not null)
            .Select(relationship => workbookPackage.GetPart(relationship.ResolvedTarget ?? string.Empty))
            .FirstOrDefault(part => part is not null);
        workbookPart ??= workbookPackage.Parts.FirstOrDefault(part => part.ContentType == WorkbookContentType);
        if (workbookPart is null)
        {
            return null;
        }

        using Stream workbookStream = workbookPart.OpenRead();
        XDocument workbookXml = SafeXml.Load(workbookStream, cancellationToken);
        ChartWorkbookSharedString[] sharedStrings = ReadWorkbookSharedStrings(workbookPackage, workbookPart, cancellationToken);
        ChartWorkbookStyles styles = ReadWorkbookStyles(workbookPackage, workbookPart, cancellationToken);
        IReadOnlyDictionary<string, OoxRelationship> workbookRelationships = workbookPackage
            .GetRelationships(workbookPart.Name, cancellationToken)
            .Where(relationship => !relationship.IsExternal && relationship.ResolvedTarget is not null)
            .ToDictionary(relationship => relationship.Id, StringComparer.Ordinal);
        ChartWorkbookSheet[] workbookSheets = ReadWorkbookSheetRecords(workbookXml, workbookRelationships);
        ChartWorkbookDefinedName[] definedNameRecords = ReadWorkbookDefinedNameRecords(workbookXml, workbookSheets);
        IReadOnlyDictionary<string, string> definedNames = ReadWorkbookDefinedNames(definedNameRecords);
        ChartWorkbookCalculationProperties calculation = ReadWorkbookCalculationProperties(workbookXml);
        var tables = new Dictionary<string, ChartWorkbookTable>(StringComparer.OrdinalIgnoreCase);
        var sheets = new Dictionary<string, ChartWorksheetData>(StringComparer.OrdinalIgnoreCase);

        foreach (ChartWorkbookSheet sheet in workbookSheets)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (string.IsNullOrWhiteSpace(sheet.Name) ||
                string.IsNullOrWhiteSpace(sheet.RelationshipId) ||
                !workbookRelationships.TryGetValue(sheet.RelationshipId, out OoxRelationship? relationship) ||
                relationship.Type != WorksheetRelationshipType ||
                relationship.ResolvedTarget is null)
            {
                continue;
            }

            OoxPart? worksheetPart = workbookPackage.GetPart(relationship.ResolvedTarget);
            if (worksheetPart is null)
            {
                continue;
            }

            sheets[sheet.Name] = ReadWorksheetData(worksheetPart, sharedStrings, cancellationToken);
            foreach (ChartWorkbookTable table in ReadWorksheetTables(workbookPackage, worksheetPart, sheet.Name, cancellationToken))
            {
                cancellationToken.ThrowIfCancellationRequested();
                tables[table.Name] = table;
                if (!string.Equals(table.DisplayName, table.Name, StringComparison.OrdinalIgnoreCase))
                {
                    tables[table.DisplayName] = table;
                }
            }
        }

        return sheets.Count == 0 ? null : new ChartWorkbookData(sheets, ReadWorkbookDate1904(workbookXml), styles, definedNames, tables, calculation, definedNameRecords, workbookSheets);
    }

    private static bool ReadWorkbookDate1904(XDocument workbookXml)
    {
        return IsOoxmlTrue((string?)workbookXml
            .Root?
            .Element(SpreadsheetNamespace + "workbookPr")
            ?.Attribute("date1904"));
    }

    private static ChartWorkbookCalculationProperties ReadWorkbookCalculationProperties(XDocument workbookXml)
    {
        XElement? calculation = workbookXml.Root?.Element(SpreadsheetNamespace + "calcPr");
        return calculation is null
            ? default
            : new ChartWorkbookCalculationProperties(
                (string?)calculation.Attribute("calcMode") ?? string.Empty,
                (string?)calculation.Attribute("calcId") ?? string.Empty,
                IsOoxmlTrue((string?)calculation.Attribute("fullCalcOnLoad")),
                IsOoxmlTrue((string?)calculation.Attribute("forceFullCalc")));
    }

    private static ChartWorkbookSharedString[] ReadWorkbookSharedStrings(OoxPackage workbookPackage, OoxPart workbookPart, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        OoxPart? sharedStringsPart = workbookPackage
            .GetRelationships(workbookPart.Name, cancellationToken)
            .Where(relationship => !relationship.IsExternal && relationship.Type == SharedStringsRelationshipType && relationship.ResolvedTarget is not null)
            .Select(relationship => workbookPackage.GetPart(relationship.ResolvedTarget ?? string.Empty))
            .FirstOrDefault(part => part is not null);
        sharedStringsPart ??= workbookPackage.Parts.FirstOrDefault(part => part.ContentType == SharedStringsContentType);
        if (sharedStringsPart is null)
        {
            return [];
        }

        using Stream stream = sharedStringsPart.OpenRead();
        XDocument document = SafeXml.Load(stream, cancellationToken);
        var sharedStrings = new List<ChartWorkbookSharedString>();
        foreach (XElement item in document.Descendants(SpreadsheetNamespace + "si"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            sharedStrings.Add(ReadWorkbookSharedString(item));
        }

        return sharedStrings.ToArray();
    }

    private static ChartWorkbookSharedString ReadWorkbookSharedString(XElement item)
    {
        XElement[] textElements = item.Descendants(SpreadsheetNamespace + "t").ToArray();
        int runCount = item.Elements(SpreadsheetNamespace + "r").Count();
        return new ChartWorkbookSharedString(
            string.Concat(textElements.Select(text => text.Value)),
            runCount,
            runCount != 0,
            item.Descendants(SpreadsheetNamespace + "rPh").Any(),
            textElements.Any(text => string.Equals((string?)text.Attribute(XNamespace.Xml + "space"), "preserve", StringComparison.Ordinal)));
    }

    private static ChartWorkbookSheet[] ReadWorkbookSheetRecords(
        XDocument workbookXml,
        IReadOnlyDictionary<string, OoxRelationship> workbookRelationships)
    {
        return workbookXml.Root?
            .Element(SpreadsheetNamespace + "sheets")
            ?.Elements(SpreadsheetNamespace + "sheet")
            .Select((sheet, index) =>
            {
                string relationshipId = ((string?)sheet.Attribute(RelationshipsNamespace + "id") ?? string.Empty).Trim();
                workbookRelationships.TryGetValue(relationshipId, out OoxRelationship? relationship);
                return new ChartWorkbookSheet(
                    ((string?)sheet.Attribute("name") ?? string.Empty).Trim(),
                    ((string?)sheet.Attribute("sheetId") ?? string.Empty).Trim(),
                    relationshipId,
                    ((string?)sheet.Attribute("state") ?? string.Empty).Trim(),
                    index,
                    relationship?.ResolvedTarget ?? string.Empty);
            })
            .ToArray() ?? [];
    }

    private static ChartWorkbookDefinedName[] ReadWorkbookDefinedNameRecords(XDocument workbookXml, IReadOnlyList<ChartWorkbookSheet> workbookSheets)
    {
        var definedNames = new List<ChartWorkbookDefinedName>();
        foreach (XElement definedName in workbookXml.Root?.Element(SpreadsheetNamespace + "definedNames")?.Elements(SpreadsheetNamespace + "definedName") ?? [])
        {
            string? name = (string?)definedName.Attribute("name");
            string formula = definedName.Value.Trim();
            if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(formula))
            {
                continue;
            }

            int? localSheetId = ReadSpreadsheetIntegerAttribute(definedName, "localSheetId");
            string sheetName = localSheetId is { } index && index >= 0 && index < workbookSheets.Count
                ? workbookSheets[index].Name
                : string.Empty;
            definedNames.Add(new ChartWorkbookDefinedName(name.Trim(), formula, localSheetId, sheetName));
        }

        return definedNames.ToArray();
    }

    private static IReadOnlyDictionary<string, string> ReadWorkbookDefinedNames(IReadOnlyList<ChartWorkbookDefinedName> definedNameRecords)
    {
        var definedNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (ChartWorkbookDefinedName definedName in definedNameRecords)
        {
            if (definedName.LocalSheetId is null)
            {
                definedNames[definedName.Name] = definedName.Formula;
            }
        }

        return definedNames;
    }

    private static ChartWorkbookStyles ReadWorkbookStyles(OoxPackage workbookPackage, OoxPart workbookPart, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        OoxPart? stylesPart = workbookPackage
            .GetRelationships(workbookPart.Name, cancellationToken)
            .Where(relationship => !relationship.IsExternal && relationship.Type == SpreadsheetStylesRelationshipType && relationship.ResolvedTarget is not null)
            .Select(relationship => workbookPackage.GetPart(relationship.ResolvedTarget ?? string.Empty))
            .FirstOrDefault(part => part is not null);
        stylesPart ??= workbookPackage.Parts.FirstOrDefault(part => part.ContentType == SpreadsheetStylesContentType);
        if (stylesPart is null)
        {
            return ChartWorkbookStyles.Empty;
        }

        using Stream stream = stylesPart.OpenRead();
        XDocument document = SafeXml.Load(stream, cancellationToken);
        var customNumberFormats = new Dictionary<int, string>();
        foreach (XElement numberFormat in document.Root?.Element(SpreadsheetNamespace + "numFmts")?.Elements(SpreadsheetNamespace + "numFmt") ?? [])
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (numberFormat.Attribute("numFmtId") is { } idAttribute &&
                int.TryParse(idAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int id) &&
                id >= 0)
            {
                customNumberFormats[id] = (string?)numberFormat.Attribute("formatCode") ?? string.Empty;
            }
        }

        var cellFormats = new List<ChartWorkbookCellFormat>();
        foreach (XElement format in document.Root?.Element(SpreadsheetNamespace + "cellXfs")?.Elements(SpreadsheetNamespace + "xf") ?? [])
        {
            cancellationToken.ThrowIfCancellationRequested();
            int? numberFormatId = ReadSpreadsheetIntegerAttribute(format, "numFmtId");
            bool? applyNumberFormat = format.Attribute("applyNumberFormat") is { } applyAttribute
                ? IsOoxmlTrue(applyAttribute.Value)
                : null;
            string numberFormatCode = numberFormatId is { } id && customNumberFormats.TryGetValue(id, out string? customCode)
                ? customCode
                : string.Empty;
            cellFormats.Add(new ChartWorkbookCellFormat(
                numberFormatId,
                numberFormatCode,
                applyNumberFormat,
                IsWorkbookDateLikeNumberFormat(numberFormatId, numberFormatCode)));
        }

        return new ChartWorkbookStyles(customNumberFormats, cellFormats);
    }

    private static bool IsWorkbookDateLikeNumberFormat(int? numberFormatId, string formatCode)
    {
        if (numberFormatId is { } id && IsBuiltInWorkbookDateLikeNumberFormatId(id))
        {
            return true;
        }

        return ContainsWorkbookDateTimeFormatToken(formatCode);
    }

    private static bool IsBuiltInWorkbookDateLikeNumberFormatId(int numberFormatId)
    {
        return (numberFormatId >= 14 && numberFormatId <= 22) ||
            (numberFormatId >= 27 && numberFormatId <= 36) ||
            (numberFormatId >= 45 && numberFormatId <= 47) ||
            (numberFormatId >= 50 && numberFormatId <= 58);
    }

    private static bool ContainsWorkbookDateTimeFormatToken(string formatCode)
    {
        if (string.IsNullOrWhiteSpace(formatCode) ||
            string.Equals(formatCode, "General", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        for (int i = 0; i < formatCode.Length; i++)
        {
            char c = formatCode[i];
            if (c == '"')
            {
                i++;
                while (i < formatCode.Length && formatCode[i] != '"')
                {
                    i++;
                }

                continue;
            }

            if (c == '\\' || c == '_' || c == '*')
            {
                i++;
                continue;
            }

            if (c == '[')
            {
                int closingBracket = formatCode.IndexOf(']', i + 1);
                if (closingBracket < 0)
                {
                    closingBracket = formatCode.Length - 1;
                }

                string bracketToken = formatCode.Substring(i + 1, Math.Max(0, closingBracket - i - 1)).Trim();
                if (bracketToken.Length > 0 &&
                    bracketToken.All(c => c == 'h' || c == 'H' || c == 'm' || c == 'M' || c == 's' || c == 'S'))
                {
                    return true;
                }

                i = closingBracket;
                continue;
            }

            char lower = char.ToLowerInvariant(c);
            if (lower == 'y' || lower == 'd' || lower == 'h' || lower == 's')
            {
                return true;
            }
        }

        return false;
    }

    private static ChartWorksheetData ReadWorksheetData(OoxPart worksheetPart, IReadOnlyList<ChartWorkbookSharedString> sharedStrings, CancellationToken cancellationToken)
    {
        using Stream stream = worksheetPart.OpenRead();
        XDocument document = SafeXml.Load(stream, cancellationToken);
        var cells = new Dictionary<string, ChartWorkbookCell>(StringComparer.OrdinalIgnoreCase);
        var hiddenRows = new HashSet<int>();
        var hiddenColumns = new HashSet<int>();
        foreach (XElement row in document.Descendants(SpreadsheetNamespace + "row"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (IsOoxmlTrue((string?)row.Attribute("hidden")) &&
                ReadSpreadsheetIntegerAttribute(row, "r") is { } rowIndex)
            {
                hiddenRows.Add(rowIndex);
            }
        }

        foreach (XElement column in document.Descendants(SpreadsheetNamespace + "col"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!IsOoxmlTrue((string?)column.Attribute("hidden")) ||
                ReadSpreadsheetIntegerAttribute(column, "min") is not { } minColumn ||
                ReadSpreadsheetIntegerAttribute(column, "max") is not { } maxColumn)
            {
                continue;
            }

            for (int columnIndex = minColumn; columnIndex <= maxColumn; columnIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                hiddenColumns.Add(columnIndex);
            }
        }

        foreach (XElement cell in document.Descendants(SpreadsheetNamespace + "c"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            string? reference = (string?)cell.Attribute("r");
            if (string.IsNullOrWhiteSpace(reference))
            {
                continue;
            }

            string? cellType = (string?)cell.Attribute("t");
            XElement? valueElement = cell.Element(SpreadsheetNamespace + "v");
            XElement? formula = cell.Element(SpreadsheetNamespace + "f");
            int? styleIndex = ReadSpreadsheetCellStyleIndex(cell);
            string? value = valueElement?.Value;
            string rawValue = value ?? string.Empty;
            bool hasValue = valueElement is not null;
            if (string.Equals(cellType, "inlineStr", StringComparison.Ordinal))
            {
                value = string.Concat(cell.Descendants(SpreadsheetNamespace + "t").Select(text => text.Value));
                hasValue = cell.Element(SpreadsheetNamespace + "is") is not null;
            }

            if (!hasValue && formula is null && styleIndex is null && string.IsNullOrEmpty(cellType))
            {
                continue;
            }

            value ??= string.Empty;
            int? sharedStringIndexValue = null;
            int sharedStringRunCount = 0;
            bool sharedStringHasRichText = false;
            bool sharedStringHasPhoneticText = false;
            bool sharedStringPreserveSpace = false;
            if (string.Equals(cellType, "s", StringComparison.Ordinal) &&
                int.TryParse(rawValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int sharedStringIndex))
            {
                sharedStringIndexValue = sharedStringIndex;
                if (sharedStringIndex >= 0 && sharedStringIndex < sharedStrings.Count)
                {
                    ChartWorkbookSharedString sharedString = sharedStrings[sharedStringIndex];
                    value = sharedString.Text;
                    sharedStringRunCount = sharedString.RunCount;
                    sharedStringHasRichText = sharedString.HasRichText;
                    sharedStringHasPhoneticText = sharedString.HasPhoneticText;
                    sharedStringPreserveSpace = sharedString.PreserveSpace;
                }
            }

            cells[reference] = new ChartWorkbookCell(
                value,
                rawValue,
                hasValue,
                valueElement is not null,
                sharedStringIndexValue,
                sharedStringRunCount,
                sharedStringHasRichText,
                sharedStringHasPhoneticText,
                sharedStringPreserveSpace,
                styleIndex,
                cellType ?? string.Empty,
                ReadWorkbookCellValueKind(cellType, value, hasValue),
                formula?.Value ?? string.Empty,
                (string?)formula?.Attribute("t") ?? string.Empty,
                ReadWorkbookFormulaAttributes(formula));
        }

        return new ChartWorksheetData(cells, hiddenRows, hiddenColumns);
    }

    private static IReadOnlyDictionary<string, string> ReadWorkbookFormulaAttributes(XElement? formula)
    {
        return formula?.Attributes()
            .ToDictionary(attribute => attribute.Name.LocalName, attribute => attribute.Value, StringComparer.Ordinal) ??
            new Dictionary<string, string>();
    }

    private static ChartWorkbookCellValueKind ReadWorkbookCellValueKind(string? cellType, string value, bool hasValue)
    {
        if (!hasValue)
        {
            return ChartWorkbookCellValueKind.Blank;
        }

        return cellType switch
        {
            null or "" => double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out _)
                ? ChartWorkbookCellValueKind.Number
                : ChartWorkbookCellValueKind.Other,
            "s" => ChartWorkbookCellValueKind.SharedString,
            "inlineStr" => ChartWorkbookCellValueKind.InlineString,
            "str" => ChartWorkbookCellValueKind.FormulaString,
            "b" => ChartWorkbookCellValueKind.Boolean,
            "e" => ChartWorkbookCellValueKind.Error,
            _ => ChartWorkbookCellValueKind.Other
        };
    }

    private static IReadOnlyList<ChartWorkbookTable> ReadWorksheetTables(OoxPackage workbookPackage, OoxPart worksheetPart, string sheetName, CancellationToken cancellationToken)
    {
        var tables = new List<ChartWorkbookTable>();
        foreach (OoxRelationship relationship in workbookPackage.GetRelationships(worksheetPart.Name, cancellationToken))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (relationship.IsExternal ||
                relationship.Type != SpreadsheetTableRelationshipType ||
                relationship.ResolvedTarget is null)
            {
                continue;
            }

            OoxPart? tablePart = workbookPackage.GetPart(relationship.ResolvedTarget);
            if (tablePart is null)
            {
                continue;
            }

            using Stream stream = tablePart.OpenRead();
            XDocument tableXml = SafeXml.Load(stream, cancellationToken);
            XElement? tableElement = tableXml.Root;
            if (tableElement is null)
            {
                continue;
            }

            string name = (string?)tableElement.Attribute("name") ?? string.Empty;
            string displayName = (string?)tableElement.Attribute("displayName") ?? name;
            string? reference = (string?)tableElement.Attribute("ref");
            if (string.IsNullOrWhiteSpace(name) ||
                string.IsNullOrWhiteSpace(reference) ||
                !TryParseTableReference(reference, out int firstColumn, out int firstRow, out int lastColumn, out int lastRow))
            {
                continue;
            }

            ChartWorkbookTableColumn[] columns = tableElement
                .Element(SpreadsheetNamespace + "tableColumns")?
                .Elements(SpreadsheetNamespace + "tableColumn")
                .Select(column => new ChartWorkbookTableColumn(
                    ReadSpreadsheetIntegerAttribute(column, "id"),
                    (string?)column.Attribute("name") ?? string.Empty,
                    (string?)column.Attribute("totalsRowFunction") ?? string.Empty,
                    column.Element(SpreadsheetNamespace + "totalsRowFormula")?.Value ?? string.Empty,
                    column.Element(SpreadsheetNamespace + "calculatedColumnFormula")?.Value ?? string.Empty))
                .ToArray() ?? [];
            if (columns.Length == 0)
            {
                continue;
            }

            string[] columnNames = columns.Select(column => column.Name).ToArray();

            XElement? autoFilter = tableElement.Element(SpreadsheetNamespace + "autoFilter");
            string autoFilterReference = (string?)autoFilter?.Attribute("ref") ?? string.Empty;
            ChartWorkbookTableFilterColumn[] filterColumns = autoFilter?
                .Elements(SpreadsheetNamespace + "filterColumn")
                .Select(ReadWorkbookTableFilterColumn)
                .ToArray() ?? [];
            int[] filterColumnIds = filterColumns
                .Select(column => column.ColumnId)
                .OfType<int>()
                .ToArray();
            int headerRowCount = ReadSpreadsheetIntegerAttribute(tableElement, "headerRowCount") ?? 1;
            int totalsRowCount = ReadSpreadsheetIntegerAttribute(tableElement, "totalsRowCount") ?? 0;
            bool totalsRowShown = IsOoxmlTrue((string?)tableElement.Attribute("totalsRowShown"));

            tables.Add(new ChartWorkbookTable(
                name,
                displayName,
                sheetName,
                firstColumn,
                firstRow,
                lastColumn,
                lastRow,
                columnNames,
                columns,
                headerRowCount,
                totalsRowCount,
                totalsRowShown,
                autoFilterReference,
                filterColumnIds,
                filterColumns));
        }

        return tables;
    }

    private static ChartWorkbookTableFilterColumn ReadWorkbookTableFilterColumn(XElement column)
    {
        XElement? filters = column.Element(SpreadsheetNamespace + "filters");
        XElement? customFilters = column.Element(SpreadsheetNamespace + "customFilters");
        XElement? dynamicFilter = column.Element(SpreadsheetNamespace + "dynamicFilter");
        XElement? top10 = column.Element(SpreadsheetNamespace + "top10");
        string filterKind = column.Elements().FirstOrDefault()?.Name.LocalName ?? string.Empty;
        return new ChartWorkbookTableFilterColumn(
            ReadSpreadsheetIntegerAttribute(column, "colId"),
            ReadOoxmlBooleanAttribute(column, "hiddenButton"),
            ReadOoxmlBooleanAttribute(column, "showButton"),
            filterKind,
            filters?.Elements(SpreadsheetNamespace + "filter")
                .Select(filter => (string?)filter.Attribute("val") ?? string.Empty)
                .ToArray() ?? [],
            (string?)dynamicFilter?.Attribute("type") ?? string.Empty,
            (string?)top10?.Attribute("val") ?? string.Empty,
            ReadOoxmlBooleanAttribute(top10, "percent"),
            ReadOoxmlBooleanAttribute(top10, "top"),
            customFilters?.Elements(SpreadsheetNamespace + "customFilter")
                .Select(filter => new ChartWorkbookTableCustomFilter(
                    (string?)filter.Attribute("operator") ?? string.Empty,
                    (string?)filter.Attribute("val") ?? string.Empty))
                .ToArray() ?? []);
    }

    private static bool TryParseTableReference(string reference, out int firstColumn, out int firstRow, out int lastColumn, out int lastRow)
    {
        firstColumn = 0;
        firstRow = 0;
        lastColumn = 0;
        lastRow = 0;
        string[] references = reference.Split(':', 2, StringSplitOptions.TrimEntries);
        if (!TryParseSpreadsheetCellReference(references[0], out firstColumn, out firstRow))
        {
            return false;
        }

        if (references.Length == 1)
        {
            lastColumn = firstColumn;
            lastRow = firstRow;
            return true;
        }

        return TryParseSpreadsheetCellReference(references[1], out lastColumn, out lastRow);
    }

    private static bool TryParseSpreadsheetCellReference(string reference, out int column, out int row)
    {
        column = 0;
        row = 0;
        string normalized = reference.Replace("$", string.Empty, StringComparison.Ordinal).Trim();
        int index = 0;
        while (index < normalized.Length && char.IsAsciiLetter(normalized[index]))
        {
            column = (column * 26) + (char.ToUpperInvariant(normalized[index]) - 'A' + 1);
            index++;
        }

        if (column <= 0 || index == normalized.Length)
        {
            return false;
        }

        return int.TryParse(normalized[index..], NumberStyles.Integer, CultureInfo.InvariantCulture, out row) && row > 0;
    }

    private static int? ReadSpreadsheetCellStyleIndex(XElement cell)
    {
        return ReadSpreadsheetIntegerAttribute(cell, "s");
    }

    private static int? ReadSpreadsheetIntegerAttribute(XElement element, string attributeName)
    {
        string? value = (string?)element.Attribute(attributeName);
        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) && parsed >= 0
            ? parsed
            : null;
    }
}