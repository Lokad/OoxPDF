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
        using OoxPackage workbookPackage = OoxPackage.Open(stream, cancellationToken);
        return ReadWorkbookDataCore(workbookPackage, cancellationToken);
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

        XDocument workbookXml = workbookPackage.LoadXml(workbookPart, cancellationToken);
        IMemoryReservation? modelStorage = OoxMemory.ReserveModel(workbookXml, cancellationToken);
        ChartWorkbookData? result = null;
        try
        {
            string[] sharedStrings = ReadWorkbookSharedStrings(workbookPackage, workbookPart, cancellationToken, modelStorage);
            ChartWorkbookStyles styles = ReadWorkbookStyles(workbookPackage, workbookPart, cancellationToken, modelStorage);
            IReadOnlyDictionary<string, OoxRelationship> workbookRelationships = workbookPackage
                .GetRelationships(workbookPart.Name, cancellationToken)
                .Where(relationship => !relationship.IsExternal && relationship.ResolvedTarget is not null)
                .ToDictionary(relationship => relationship.Id, StringComparer.Ordinal);
            ChartWorkbookSheet[] workbookSheets = ReadWorkbookSheetRecords(workbookXml);
            if (workbookSheets.Length > MaxWorkbookSheets)
            {
                throw new OoxPdfLimitExceededException("Chart workbook declares more sheets than the maximum supported count.");
            }
            ChartWorkbookDefinedName[] definedNameRecords = ReadWorkbookDefinedNameRecords(workbookXml, workbookSheets);
            IReadOnlyDictionary<string, string> definedNames = ReadWorkbookDefinedNames(definedNameRecords);
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

                ChartWorksheetData worksheetData = ReadWorksheetData(workbookPackage, worksheetPart, sharedStrings, cancellationToken, modelStorage);
                long totalCells = 0;
                foreach (ChartWorksheetData existing in sheets.Values)
                {
                    totalCells = checked(totalCells + existing.Cells.Count);
                }
                totalCells = checked(totalCells + worksheetData.Cells.Count);
                if (totalCells > MaxWorkbookTotalCells)
                {
                    throw new OoxPdfLimitExceededException(
                        "Chart workbook exceeds the maximum supported total cell count of " + MaxWorkbookTotalCells + ".");
                }

                sheets[sheet.Name] = worksheetData;
                foreach (ChartWorkbookTable table in ReadWorksheetTables(workbookPackage, worksheetPart, sheet.Name, cancellationToken, modelStorage))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    tables[table.Name] = table;
                    if (!string.Equals(table.DisplayName, table.Name, StringComparison.OrdinalIgnoreCase))
                    {
                        tables[table.DisplayName] = table;
                    }
                }
            }

            result = sheets.Count == 0 ? null : new ChartWorkbookData(sheets, styles, definedNames, tables, definedNameRecords, cancellationToken);
            return result;
        }
        finally
        {
            if (result is null) modelStorage?.Dispose();
            else if (modelStorage is not null) OoxMemory.Resize(modelStorage, modelStorage.Bytes / 2);
        }
    }

    private static string[] ReadWorkbookSharedStrings(OoxPackage workbookPackage, OoxPart workbookPart, CancellationToken cancellationToken, IMemoryReservation? modelStorage = null)
    {
        cancellationToken.ThrowIfCancellationRequested();
        OoxPart? sharedStringsPart = workbookPackage
            .GetRelationships(workbookPart.Name, cancellationToken)
            .Where(relationship => !relationship.IsExternal && relationship.Type == SharedStringsRelationshipType && relationship.ResolvedTarget is not null)
            .Select(relationship => workbookPackage.GetPart(relationship.ResolvedTarget ?? string.Empty))
            .FirstOrDefault(part => part is not null);
        if (sharedStringsPart is null)
        {
            return [];
        }

        XDocument document = workbookPackage.LoadXml(sharedStringsPart, cancellationToken);
        OoxMemory.GrowModel(modelStorage, document, cancellationToken);
        var sharedStrings = new List<string>();
        foreach (XElement item in document.Descendants(SpreadsheetNamespace + "si"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (sharedStrings.Count >= MaxWorkbookSharedStrings)
            {
                throw new OoxPdfLimitExceededException("Chart workbook declares more shared strings than the maximum supported count.");
            }
            sharedStrings.Add(string.Concat(ReadWorkbookBaseTextElements(item).Select(text => text.Value)));
        }

        return sharedStrings.ToArray();
    }

    // Shared and inline strings use the same base-text owners. Phonetic runs
    // annotate that text and must not be appended to the cell value.
    private static IEnumerable<XElement> ReadWorkbookBaseTextElements(XElement? item)
    {
        foreach (XElement child in item?.Elements() ?? [])
        {
            if (child.Name == SpreadsheetNamespace + "t") yield return child;
            else if (child.Name == SpreadsheetNamespace + "r")
                foreach (XElement text in child.Elements(SpreadsheetNamespace + "t")) yield return text;
        }
    }

    private static ChartWorkbookSheet[] ReadWorkbookSheetRecords(XDocument workbookXml)
    {
        return workbookXml.Root?
            .Element(SpreadsheetNamespace + "sheets")
            ?.Elements(SpreadsheetNamespace + "sheet")
            .Select(sheet =>
            {
                string relationshipId = ((string?)sheet.Attribute(RelationshipsNamespace + "id") ?? string.Empty).Trim();
                return new ChartWorkbookSheet(
                    ((string?)sheet.Attribute("name") ?? string.Empty).Trim(),
                    relationshipId);
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
                throw new OoxPdfLimitExceededException("Chart workbook declares more defined names than the maximum supported count.");
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

    private static ChartWorkbookStyles ReadWorkbookStyles(OoxPackage workbookPackage, OoxPart workbookPart, CancellationToken cancellationToken, IMemoryReservation? modelStorage = null)
    {
        cancellationToken.ThrowIfCancellationRequested();
        OoxPart? stylesPart = workbookPackage
            .GetRelationships(workbookPart.Name, cancellationToken)
            .Where(relationship => !relationship.IsExternal && relationship.Type == SpreadsheetStylesRelationshipType && relationship.ResolvedTarget is not null)
            .Select(relationship => workbookPackage.GetPart(relationship.ResolvedTarget ?? string.Empty))
            .FirstOrDefault(part => part is not null);
        if (stylesPart is null)
        {
            return ChartWorkbookStyles.Empty;
        }

        XDocument document = workbookPackage.LoadXml(stylesPart, cancellationToken);
        OoxMemory.GrowModel(modelStorage, document, cancellationToken);
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
                    throw new OoxPdfLimitExceededException("Chart workbook declares more number formats than the maximum supported count.");
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
                throw new OoxPdfLimitExceededException("Chart workbook declares more cell formats than the maximum supported count.");
            }

            int? numberFormatId = ReadSpreadsheetIntegerAttribute(format, "numFmtId");
            bool? applyNumberFormat = format.Attribute("applyNumberFormat") is { } applyAttribute
                ? IsOoxmlTrue(applyAttribute.Value)
                : null;
            string numberFormatCode = numberFormatId is { } id && customNumberFormats.TryGetValue(id, out string? customCode)
                ? customCode
                : string.Empty;
            cellFormats.Add(new ChartWorkbookCellFormat(numberFormatCode, applyNumberFormat));
        }

        return new ChartWorkbookStyles(cellFormats);
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

    // Bounds on materialized embedded-workbook records, independent of the
    // number of cells selected by an individual chart reference.
    private const long MaxWorkbookSharedStrings = 100_000;
    private const long MaxWorkbookTotalCells = 100_000;
    private const long MaxWorksheetCells = 100_000;
    private const int MaxWorkbookSheets = 1_024;
    private const long MaxWorkbookDefinedNames = 100_000;
    private const long MaxWorkbookStyleRecords = 100_000;

    private static ChartWorksheetData ReadWorksheetData(OoxPackage workbookPackage, OoxPart worksheetPart, IReadOnlyList<string> sharedStrings, CancellationToken cancellationToken, IMemoryReservation? modelStorage = null)
    {
        XDocument document = workbookPackage.LoadXml(worksheetPart, cancellationToken);
        OoxMemory.GrowModel(modelStorage, document, cancellationToken);
        var cells = new Dictionary<string, ChartWorkbookCell>(StringComparer.OrdinalIgnoreCase);
        var hiddenRows = new HashSet<int>();
        bool[] hiddenColumns = [];
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

        foreach (XElement column in document.Descendants(SpreadsheetNamespace + "col"))
        {
            cancellationToken.ThrowIfCancellationRequested();
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

            // Fixed-domain flags fit the existing model window admitted above.
            // Worksheets without a valid hidden range need no array.
            if (hiddenColumns.Length == 0)
            {
                hiddenColumns = new bool[MaxSpreadsheetColumn + 1];
            }
            hiddenColumns.AsSpan(firstHiddenColumn, lastHiddenColumn - firstHiddenColumn + 1).Fill(true);
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
                value = string.Concat(ReadWorkbookBaseTextElements(cell.Element(SpreadsheetNamespace + "is")).Select(text => text.Value));
                hasValue = cell.Element(SpreadsheetNamespace + "is") is not null;
            }

            if (!hasValue && formula is null && styleIndex is null && string.IsNullOrEmpty(cellType))
            {
                continue;
            }

            value ??= string.Empty;
            if (string.Equals(cellType, "s", StringComparison.Ordinal) &&
                int.TryParse(rawValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int sharedStringIndex) &&
                sharedStringIndex >= 0 && sharedStringIndex < sharedStrings.Count)
            {
                value = sharedStrings[sharedStringIndex];
            }

            if (!cells.ContainsKey(reference) && cells.Count >= MaxWorksheetCells)
            {
                throw new OoxPdfLimitExceededException("Chart worksheet declares more cells than the maximum supported count.");
            }

            cells[reference] = new ChartWorkbookCell(value, hasValue, styleIndex);
        }

        return new ChartWorksheetData(cells, hiddenRows, hiddenColumns);
    }

    private static IReadOnlyList<ChartWorkbookTable> ReadWorksheetTables(OoxPackage workbookPackage, OoxPart worksheetPart, string sheetName, CancellationToken cancellationToken, IMemoryReservation? modelStorage = null)
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

            XDocument tableXml = workbookPackage.LoadXml(tablePart, cancellationToken);
            OoxMemory.GrowModel(modelStorage, tableXml, cancellationToken);
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

            string[] columnNames = tableElement
                .Element(SpreadsheetNamespace + "tableColumns")?
                .Elements(SpreadsheetNamespace + "tableColumn")
                .Select(column => (string?)column.Attribute("name") ?? string.Empty)
                .ToArray() ?? [];
            if (columnNames.Length == 0)
            {
                continue;
            }

            int headerRowCount = ReadSpreadsheetIntegerAttribute(tableElement, "headerRowCount") ?? 1;
            int totalsRowCount = ReadSpreadsheetIntegerAttribute(tableElement, "totalsRowCount") ?? 0;

            tables.Add(new ChartWorkbookTable(
                name,
                displayName,
                sheetName,
                firstColumn,
                firstRow,
                lastColumn,
                lastRow,
                columnNames,
                headerRowCount,
                totalsRowCount));
        }

        return tables;
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
