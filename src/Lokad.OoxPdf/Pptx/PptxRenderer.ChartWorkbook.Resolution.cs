using System.Globalization;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private sealed partial class ChartWorkbookData
    {
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
            if (ParseStructuredReferenceBody(body) is not { } reference)
            {
                return trimmed;
            }

            string firstColumnName = reference.FirstColumnName;
            string lastColumnName = reference.LastColumnName;
            bool wholeTable = reference.WholeTable;
            bool onlyHeader = reference.OnlyHeader;
            bool onlyTotals = reference.OnlyTotals;
            bool includeHeader = reference.IncludeHeader;

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

        private static ChartWorkbookStructuredReferenceBody? ParseStructuredReferenceBody(string body)
        {
            string firstColumnName = string.Empty;
            string lastColumnName = string.Empty;
            bool includeHeader = false;
            bool onlyHeader = false;
            bool onlyTotals = false;
            bool wholeTable = false;

            bool ApplyStructuredReferenceItem(string segment)
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

            string trimmed = body.Trim();
            if (trimmed.Length == 0)
            {
                return null;
            }

            if (trimmed[0] != '[')
            {
                string segment = UnescapeStructuredReferenceText(trimmed);
                if (!ApplyStructuredReferenceItem(segment))
                {
                    firstColumnName = segment;
                    lastColumnName = segment;
                }

                return new ChartWorkbookStructuredReferenceBody(
                    firstColumnName,
                    lastColumnName,
                    includeHeader,
                    onlyHeader,
                    onlyTotals,
                    wholeTable);
            }

            foreach (string segment in ParseStructuredReferenceSegments(trimmed))
            {
                if (ApplyStructuredReferenceItem(segment))
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

            if (!wholeTable && string.IsNullOrWhiteSpace(firstColumnName))
            {
                return null;
            }

            return new ChartWorkbookStructuredReferenceBody(
                firstColumnName,
                lastColumnName,
                includeHeader,
                onlyHeader,
                onlyTotals,
                wholeTable);
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
}
