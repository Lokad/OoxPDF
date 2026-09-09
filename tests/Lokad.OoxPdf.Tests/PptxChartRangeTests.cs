using System.Reflection;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxChartRangeTests
{

    public static void FullColumnFormulaThrowsPromptly()
    {
        object workbook = CreateWorkbook();
        TestAssert.Throws<InvalidDataException>(() => ReadCellCount(workbook, "Sheet1!$A$1:$XFD$1048576"));
    }

    public static void FullRowFormulaReadsWithoutExpansion()
    {
        object workbook = CreateWorkbook();
        TestAssert.Equal(16384, ReadCellCount(workbook, "Sheet1!$A$1048576:$XFD$1048576"));
    }

    public static void OutOfGridReferencesResolveEmpty()
    {
        object workbook = CreateWorkbook();
        TestAssert.Equal(0, ReadCellCount(workbook, "Sheet1!$A$1:$XFE$1"));
        TestAssert.Equal(0, ReadCellCount(workbook, "Sheet1!$A$0:$A$5"));
        TestAssert.Equal(0, ReadCellCount(workbook, "Sheet1!$AAAA$1:$AAAA$2"));
    }

    public static void SmallRangeReadsCells()
    {
        object workbook = CreateWorkbook();
        TestAssert.Equal(2, ReadCellCount(workbook, "Sheet1!$A$1:$A$2"));
        TestAssert.Equal(1, ReadCellCount(workbook, "Sheet1!$A$1"));
    }

    public static void RangeAtCellBudgetBoundaryPasses()
    {
        object workbook = CreateWorkbook();
        TestAssert.Equal(100000, ReadCellCount(workbook, "Sheet1!$A$1:$A$100000"));
    }

    public static void RangeBeyondCellBudgetThrows()
    {
        object workbook = CreateWorkbook();
        TestAssert.Throws<InvalidDataException>(() => ReadCellCount(workbook, "Sheet1!$A$1:$A$100001"));
    }

    private static object CreateWorkbook()
    {
        Type workbookType = typeof(PptxRenderer).GetNestedType(
            "ChartWorkbookData",
            BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart workbook data type.");
        var sheets = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase)
        {
            ["Sheet1"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["A1"] = "5",
                ["A2"] = "6",
            },
        };
        return Activator.CreateInstance(workbookType, [sheets]) ?? throw new InvalidOperationException("Expected workbook instance.");
    }

    private static int ReadCellCount(object workbook, string formula)
    {
        MethodInfo read = workbook.GetType().GetMethod("ReadRangeCells") ?? throw new InvalidOperationException("Expected range reader.");
        try
        {
            return ((Array)read.Invoke(workbook, [formula])!).Length;
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }

    }
    public static void HiddenColumnRunBeyondGridIsClamped()
    {
        object workbook = ReadWorkbookFromSheetXml(
            "<cols><col min=\"1\" max=\"2147483647\" hidden=\"1\"/></cols>" +
            "<sheetData><row r=\"1\"><c r=\"A1\"><v>5</v></c><c r=\"B1\"><v>6</v></c></row></sheetData>");
        Array cells = ReadRangeCells(workbook, "Sheet1!$A$1:$B$1");
        TestAssert.Equal(2, cells.Length);
        TestAssert.True(ReadColumnHidden(cells, 1, 1), "Column A must report hidden.");
        TestAssert.True(ReadColumnHidden(cells, 1, 2), "Column B must report hidden.");
    }

    public static void NarrowHiddenColumnRunPreservesSemantics()
    {
        object workbook = ReadWorkbookFromSheetXml(
            "<cols><col min=\"1\" max=\"1\" hidden=\"1\"/></cols>" +
            "<sheetData><row r=\"1\"><c r=\"A1\"><v>5</v></c><c r=\"B1\"><v>6</v></c><c r=\"C1\"><v>7</v></c></row></sheetData>");
        Array cells = ReadRangeCells(workbook, "Sheet1!$A$1:$C$1");
        TestAssert.Equal(3, cells.Length);
        TestAssert.True(ReadColumnHidden(cells, 1, 1), "Column A must report hidden.");
        TestAssert.True(!ReadColumnHidden(cells, 1, 2), "Column B must report visible.");
        TestAssert.True(!ReadColumnHidden(cells, 1, 3), "Column C must report visible.");
    }

    public static void OutOfGridHiddenRowIsIgnored()
    {
        object workbook = ReadWorkbookFromSheetXml(
            "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row>" +
            "<row r=\"2\" hidden=\"1\"><c r=\"A2\"><v>2</v></c></row>" +
            "<row r=\"99999999\" hidden=\"1\"/></sheetData>");
        Array cells = ReadRangeCells(workbook, "Sheet1!$A$1:$A$2");
        TestAssert.Equal(2, cells.Length);
        TestAssert.True(!ReadRowHidden(cells, 1, 1), "Row 1 must report visible.");
        TestAssert.True(ReadRowHidden(cells, 2, 1), "Row 2 must report hidden.");
    }

    public static void RepeatedFullWidthHiddenColumnRunsFailPromptly()
    {
        string runs = string.Concat(Enumerable.Repeat("<col min=\"1\" max=\"14000\" hidden=\"1\"/>", 8));
        TestAssert.Throws<InvalidDataException>(() => ReadWorkbookFromSheetXml(
            "<cols>" + runs + "</cols><sheetData><row r=\"1\"><c r=\"A1\"><v>5</v></c></row></sheetData>"));
    }

    public static void HiddenColumnBudgetBoundaryPasses()
    {
        string runs = string.Concat(Enumerable.Repeat("<col min=\"1\" max=\"14000\" hidden=\"1\"/>", 7));
        object workbook = ReadWorkbookFromSheetXml(
            "<cols>" + runs + "</cols><sheetData><row r=\"1\"><c r=\"A1\"><v>5</v></c></row></sheetData>");
        TestAssert.Equal(1, ReadCellCount(workbook, "Sheet1!$A$1"));
    }
    public static void WorkbookSharedStringsBeyondBudgetThrow()
    {
        var strings = new System.Text.StringBuilder();
        for (int i = 0; i < 100001; i++)
        {
            strings.Append("<si><t>x</t></si>");
        }

        TestAssert.Throws<InvalidDataException>(() => ReadWorkbookFromSheetXml(
            "<sheetData><row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c></row></sheetData>",
            strings.ToString()));
    }

    public static void WorksheetCellsBeyondBudgetThrow()
    {
        TestAssert.Throws<InvalidDataException>(() => ReadWorkbookFromSheetXml(BuildManyCellSheet(100001)));
    }
    public static void WorkbookDefinedNamesBeyondBudgetThrow()
    {
        var names = new System.Text.StringBuilder();
        for (int i = 0; i < 100001; i++)
        {
            names.Append("<definedName name=\"N" + i + "\">Sheet1!$A$1</definedName>");
        }

        TestAssert.Throws<InvalidDataException>(() => ReadWorkbookPackage(
            "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets><definedNames>" + names.ToString() + "</definedNames>",
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>",
            new Dictionary<string, string>
            {
                ["xl/worksheets/sheet1.xml"] = "<sheetData><row r=\"1\"><c r=\"A1\"><v>5</v></c></row></sheetData>",
            },
            null));
    }


    public static void WorksheetCellsAtBudgetBoundaryPass()
    {
        object workbook = ReadWorkbookFromSheetXml(BuildManyCellSheet(100000));
        TestAssert.Equal(1, ReadCellCount(workbook, "Sheet1!$A$100000"));
    }

    public static void WorkbookSheetsBeyondBudgetThrow()
    {
        var sheets = new System.Text.StringBuilder();
        var rels = new System.Text.StringBuilder();
        for (int i = 1; i <= 1025; i++)
        {
            sheets.Append("<sheet name=\"S" + i + "\" sheetId=\"" + i + "\" r:id=\"rId" + i + "\"/>");
            rels.Append("<Relationship Id=\"rId" + i + "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet" + i + ".xml\"/>");
        }

        TestAssert.Throws<InvalidDataException>(() => ReadWorkbookPackage(
            "<sheets>" + sheets.ToString() + "</sheets>",
            rels.ToString(),
            new Dictionary<string, string>(),
            null));
    }
    public static void WorkbookCellFormatsBeyondBudgetThrow()
    {
        var formats = new System.Text.StringBuilder();
        formats.Append("<cellXfs>");
        for (int i = 0; i < 100001; i++)
        {
            formats.Append("<xf numFmtId=\"0\"/>");
        }

        formats.Append("</cellXfs>");
        TestAssert.Throws<InvalidDataException>(() => ReadWorkbookPackage(
            "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>",
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>",
            new Dictionary<string, string>
            {
                ["xl/worksheets/sheet1.xml"] = "<sheetData><row r=\"1\"><c r=\"A1\"><v>5</v></c></row></sheetData>",
            },
            null,
            "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" + formats.ToString() + "</styleSheet>"));
    }


    private static string BuildManyCellSheet(int cellCount)
    {
        var rows = new System.Text.StringBuilder();
        rows.Append("<sheetData>");
        for (int i = 1; i <= cellCount; i++)
        {
            rows.Append("<row r=\"" + i + "\"><c r=\"A" + i + "\"><v>1</v></c></row>");
        }

        rows.Append("</sheetData>");
        return rows.ToString();
    }


    private static object ReadWorkbookFromSheetXml(string sheetBody)
    {
        return ReadWorkbookFromSheetXml(sheetBody, null);
    }

    private static object ReadWorkbookFromSheetXml(string sheetBody, string? sharedStringsBody)
    {
        return ReadWorkbookPackage(
            "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>",
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>",
            new Dictionary<string, string>
            {
                ["xl/worksheets/sheet1.xml"] = sheetBody,
            },
            sharedStringsBody, null);
    }

    private static object ReadWorkbookPackage(string sheetsFragment, string workbookRelsFragment, IReadOnlyDictionary<string, string> sheetBodies, string? sharedStringsBody, string? stylesBody = null)
    {
        const string spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        const string officeRels = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        const string packageRels = "http://schemas.openxmlformats.org/package/2006/relationships";
        string contentTypes =
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
            "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
            "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
            (sharedStringsBody is null ? string.Empty : "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>") +
            (stylesBody is null ? string.Empty : "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
        foreach (string partName in sheetBodies.Keys)
        {
            contentTypes += "<Override PartName=\"/" + partName + "\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>";
        }

        contentTypes += "</Types>";
        var entries = new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = contentTypes,
            ["xl/workbook.xml"] =
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<workbook xmlns=\"" + spreadsheet + "\" xmlns:r=\"" + officeRels + "\">" +
                sheetsFragment + "</workbook>",
            ["xl/_rels/workbook.xml.rels"] =
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<Relationships xmlns=\"" + packageRels + "\">" +
                workbookRelsFragment +
                (sharedStringsBody is null ? string.Empty : "<Relationship Id=\"rIdShared\" Type=\"" + officeRels + "/sharedStrings\" Target=\"sharedStrings.xml\"/>") +
                (stylesBody is null ? string.Empty : "<Relationship Id=\"rIdStyles\" Type=\"" + officeRels + "/styles\" Target=\"styles.xml\"/>") +
                "</Relationships>",
        };
        foreach ((string name, string body) in sheetBodies)
        {
            entries[name] =
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<worksheet xmlns=\"" + spreadsheet + "\">" + body + "</worksheet>";
        }

        if (sharedStringsBody is not null)
        {
            entries["xl/sharedStrings.xml"] =
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<sst xmlns=\"" + spreadsheet + "\">" + sharedStringsBody + "</sst>";
        }

        if (stylesBody is not null)
        {
            entries["xl/styles.xml"] = stylesBody;
        }

        using MemoryStream packageStream = TestFixtures.CreateZipPackage(entries);
        OoxPackage package = OoxPackage.Open(packageStream, CancellationToken.None);
        MethodInfo read = typeof(PptxRenderer).GetMethod(
            "ReadWorkbookDataCore",
            BindingFlags.NonPublic | BindingFlags.Static) ?? throw new InvalidOperationException("Expected workbook reader.");
        try
        {
            return read.Invoke(null, [package, CancellationToken.None])
                ?? throw new InvalidOperationException("Expected workbook data.");
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }
    }

    private static Array ReadRangeCells(object workbook, string formula)
    {
        MethodInfo read = workbook.GetType().GetMethod("ReadRangeCells") ?? throw new InvalidOperationException("Expected range reader.");
        try
        {
            return (Array)read.Invoke(workbook, [formula])!;
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }
    }

    private static object FindRangeCell(Array cells, int row, int column)
    {
        foreach (object? item in cells)
        {
            object cell = item ?? throw new InvalidOperationException("Expected range cell.");
            int cellRow = (int)(cell.GetType().GetProperty("SheetRow")?.GetValue(cell) ?? throw new InvalidOperationException("Expected sheet row."));
            int cellColumn = (int)(cell.GetType().GetProperty("SheetColumn")?.GetValue(cell) ?? throw new InvalidOperationException("Expected sheet column."));
            if (cellRow == row && cellColumn == column)
            {
                return cell;
            }
        }

        throw new InvalidOperationException("Range cell was not found.");
    }

    private static bool ReadRowHidden(Array cells, int row, int column)
    {
        object cell = FindRangeCell(cells, row, column);
        return (bool)(cell.GetType().GetProperty("RowHidden")?.GetValue(cell) ?? throw new InvalidOperationException("Expected row-hidden flag."));
    }

    private static bool ReadColumnHidden(Array cells, int row, int column)
    {
        object cell = FindRangeCell(cells, row, column);
        return (bool)(cell.GetType().GetProperty("ColumnHidden")?.GetValue(cell) ?? throw new InvalidOperationException("Expected column-hidden flag."));
    }
}

