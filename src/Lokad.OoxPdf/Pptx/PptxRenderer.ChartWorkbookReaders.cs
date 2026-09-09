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
        if (workbookSheets.Length > MaxWorkbookSheets)
        {
            throw new InvalidDataException("Chart workbook declares more sheets than the maximum supported count.");
        }
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
            if (sharedStrings.Count >= MaxWorkbookSharedStrings)
            {
                throw new InvalidDataException("Chart workbook declares more shared strings than the maximum supported count.");
            }
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
            if (definedNames.Count >= MaxWorkbookDefinedNames)
            {
                throw new InvalidDataException("Chart workbook declares more defined names than the maximum supported count.");
            }

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
                if (!customNumberFormats.ContainsKey(id) && customNumberFormats.Count >= MaxWorkbookStyleRecords)
                {
                    throw new InvalidDataException("Chart workbook declares more number formats than the maximum supported count.");
                }

                customNumberFormats[id] = (string?)numberFormat.Attribute("formatCode") ?? string.Empty;
            }
        }

        var cellFormats = new List<ChartWorkbookCellFormat>();
        foreach (XElement format in document.Root?.Element(SpreadsheetNamespace + "cellXfs")?.Elements(SpreadsheetNamespace + "xf") ?? [])
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (cellFormats.Count >= MaxWorkbookStyleRecords)
            {
                throw new InvalidDataException("Chart workbook declares more cell formats than the maximum supported count.");
            }

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

        return ContainsWorkbookDateTimeFormatToken();

        bool ContainsWorkbookDateTimeFormatToken()
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

        bool IsBuiltInWorkbookDateLikeNumberFormatId(int numberFormatId)
        {
            return (numberFormatId >= 14 && numberFormatId <= 22) ||
                (numberFormatId >= 27 && numberFormatId <= 36) ||
                (numberFormatId >= 45 && numberFormatId <= 47) ||
                (numberFormatId >= 50 && numberFormatId <= 58);
        }
    }

    // Spreadsheet grid bounds (columns A..XFD, rows 1..1048576). Hidden markers
    // outside the grid can never be observed through resolved range cells, whose
    // references are grid-bounded by TryParseCellReference.
    private const int MaxSpreadsheetColumn = 16384;
    private const int MaxSpreadsheetRow = 1048576;

    // Single owner for A1-style references. Formula ranges use the Excel grid
    // bounds; worksheet table/dimension refs pass unbounded limits to preserve
    // their validation policy.
    private static bool TryParseCellReference(
        string reference,
        out int column,
        out int row,
        int maxColumn = MaxSpreadsheetColumn,
        int maxRow = MaxSpreadsheetRow,
        int maxLetters = 3)
    {
        column = 0;
        row = 0;
        string normalized = reference.Replace("$", string.Empty, StringComparison.Ordinal).Trim();
        int index = 0;
        int letters = 0;
        while (index < normalized.Length && char.IsAsciiLetter(normalized[index]))
        {
            column = column * 26 + (char.ToUpperInvariant(normalized[index]) - 'A' + 1);
            index++;
            letters++;
            if (letters > maxLetters)
            {
                return false;
            }
        }

        if (letters == 0 || letters > maxLetters || column <= 0 || column > maxColumn || index == normalized.Length)
        {
            return false;
        }

        return int.TryParse(normalized[index..], NumberStyles.Integer, CultureInfo.InvariantCulture, out row) && row >= 1 && row <= maxRow;
    }

    // Total hidden-column insertions examined across all <col> runs in one
    // worksheet. The grid holds at most MaxSpreadsheetColumn distinct columns,
    // so genuine files stay far below this budget; anything beyond signals a
    // hostile repetition of wide runs. Mirrors MaxChartRangeCells.
    private const long MaxHiddenColumnExpansions = 100_000;

    // Count budgets for embedded-workbook expansion. Chart data ranges are
    // already capped at MaxChartRangeCells cells, so genuine chart workbooks
    // stay far below these counts; anything beyond signals a hostile payload
    // smuggled in through an embedded spreadsheet.
    private const long MaxWorkbookSharedStrings = 100_000;
    private const long MaxWorksheetCells = 100_000;
    private const int MaxWorkbookSheets = 1_024;
    private const long MaxWorkbookDefinedNames = 100_000;
    private const long MaxWorkbookStyleRecords = 100_000;

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
                ReadSpreadsheetIntegerAttribute(row, "r") is { } rowIndex &&
                rowIndex >= 1 &&
                rowIndex <= MaxSpreadsheetRow)
            {
                hiddenRows.Add(rowIndex);
            }
        }

        long hiddenColumnExpansions = 0;
        foreach (XElement column in document.Descendants(SpreadsheetNamespace + "col"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (hiddenColumns.Count >= MaxSpreadsheetColumn)
            {
                break;
            }

            if (!IsOoxmlTrue((string?)column.Attribute("hidden")) ||
                ReadSpreadsheetIntegerAttribute(column, "min") is not { } minColumn ||
                ReadSpreadsheetIntegerAttribute(column, "max") is not { } maxColumn ||
                minColumn > maxColumn)
            {
                continue;
            }

            int firstHiddenColumn = Math.Max(minColumn, 1);
            int lastHiddenColumn = Math.Min(maxColumn, MaxSpreadsheetColumn);
            if (firstHiddenColumn > lastHiddenColumn)
            {
                continue;
            }

            for (int columnIndex = firstHiddenColumn; columnIndex <= lastHiddenColumn; columnIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (hiddenColumns.Count >= MaxSpreadsheetColumn)
                {
                    break;
                }

                hiddenColumns.Add(columnIndex);
                hiddenColumnExpansions++;
                if (hiddenColumnExpansions > MaxHiddenColumnExpansions)
                {
                    throw new InvalidDataException("Chart workbook hides more columns than the maximum supported count.");
                }
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

            if (!cells.ContainsKey(reference) && cells.Count >= MaxWorksheetCells)
            {
                throw new InvalidDataException("Chart worksheet declares more cells than the maximum supported count.");
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
        if (!TryParseCellReference(references[0], out firstColumn, out firstRow, maxColumn: int.MaxValue, maxRow: int.MaxValue, maxLetters: int.MaxValue))
        {
            return false;
        }

        if (references.Length == 1)
        {
            lastColumn = firstColumn;
            lastRow = firstRow;
            return true;
        }

        return TryParseCellReference(references[1], out lastColumn, out lastRow, maxColumn: int.MaxValue, maxRow: int.MaxValue, maxLetters: int.MaxValue);
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
