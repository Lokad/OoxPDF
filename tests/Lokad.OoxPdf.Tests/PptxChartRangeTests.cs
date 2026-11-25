using System.Reflection;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxChartRangeTests
{

    public static void FullGridAndColumnFormulasExceedRangeBudget()
    {
        var workbook = CreateWorkbook();
        TestAssert.Throws<OoxPdfLimitExceededException>(() => workbook.ReadRangeCells("Sheet1!$A$1:$XFD$1048576"));
        TestAssert.Throws<OoxPdfLimitExceededException>(() => workbook.ReadRangeCells("Sheet1!$A$1:$A$1048576"));
    }

    public static void FullRowFormulaReturnsEveryColumn()
    {
        var workbook = CreateWorkbook();
        TestAssert.Equal(16384, workbook.ReadRangeCells("Sheet1!$A$1048576:$XFD$1048576").Length);
    }

    public static void OutOfGridReferencesResolveEmpty()
    {
        var workbook = CreateWorkbook();
        TestAssert.Equal(0, workbook.ReadRangeCells("Sheet1!$A$1:$XFE$1").Length);
        TestAssert.Equal(0, workbook.ReadRangeCells("Sheet1!$A$0:$A$5").Length);
        TestAssert.Equal(0, workbook.ReadRangeCells("Sheet1!$AAAA$1:$AAAA$2").Length);
    }

    public static void TableReferencesKeepUnboundedValidation()
    {
        MethodInfo parse = typeof(PptxRenderer).GetMethod(
            "TryParseTableReference",
            BindingFlags.NonPublic | BindingFlags.Static) ?? throw new InvalidOperationException("Expected table reference parser.");
        object?[] args = ["A1:XFE1048577", 0, 0, 0, 0];
        bool parsed = (bool)parse.Invoke(null, args)!;
        TestAssert.True(parsed, "Worksheet table refs keep unbounded validation (column 16385, row 1048577 must parse).");
        TestAssert.Equal(1, (int)args[1]!);
        TestAssert.Equal(1, (int)args[2]!);
        TestAssert.Equal(16385, (int)args[3]!);
        TestAssert.Equal(1048577, (int)args[4]!);
    }

    public static void SmallRangeReadsCells()
    {
        var workbook = CreateWorkbook();
        TestAssert.Equal(2, workbook.ReadRangeCells("Sheet1!$A$1:$A$2").Length);
        TestAssert.Equal(1, workbook.ReadRangeCells("Sheet1!$A$1").Length);
    }

    public static void WorkbookScopedNamesUseDeclaredSheetOrder()
    {
        var workbook = ReadWorkbookPackage(
            """
            <sheets><sheet name="First" sheetId="91" r:id="first"/><sheet name="Second" sheetId="7" r:id="second"/></sheets>
            <definedNames>
              <definedName name="Global">Second!$A$1</definedName>
              <definedName name="global">First!$A$1</definedName>
              <definedName name="Scoped" localSheetId="17">Second!$A$1</definedName>
              <definedName name="Scoped" localSheetId="0">First!$A$1</definedName>
              <definedName name="scoped" localSheetId="0">Second!$A$1</definedName>
              <definedName name="Scoped" localSheetId="1">Second!$A$1</definedName>
              <definedName name="Unresolved" localSheetId="19">First!$A$1</definedName>
              <definedName name="MalformedScope" localSheetId="-1">Second!$A$1</definedName>
            </definedNames>
            """,
            """
            <Relationship Id="first" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/first.xml"/>
            <Relationship Id="second" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/second.xml"/>
            """,
            new Dictionary<string, string>
            {
                ["xl/worksheets/first.xml"] = "<sheetData><row r=\"1\"><c r=\"A1\"><v>5</v></c></row></sheetData>",
                ["xl/worksheets/second.xml"] = "<sheetData><row r=\"1\"><c r=\"A1\"><v>9</v></c></row></sheetData>"
            }, null);
        TestAssert.Equal("5", workbook.ReadRangeCells("Global").Single().Text);
        TestAssert.Equal("5", workbook.ReadRangeCells("First!Scoped").Single().Text);
        TestAssert.Equal("9", workbook.ReadRangeCells("Second!Scoped").Single().Text);
        TestAssert.Equal(0, workbook.ReadRangeCells("Unresolved").Length);
        TestAssert.Equal(0, workbook.ReadRangeCells("First!Unresolved").Length);
        TestAssert.Equal("9", workbook.ReadRangeCells("MalformedScope").Single().Text);
    }

    private static PptxRenderer.ChartWorkbookData CreateWorkbook()
    {
        var sheets = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase)
        {
            ["Sheet1"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["A1"] = "5",
                ["A2"] = "6",
            },
        };
        return PptxChartRangeTests.ReadWorkbookCells(sheets);
    }

    public static void HiddenColumnRunBeyondGridIsClamped()
    {
        var workbook = ReadWorkbookFromSheetXml(
            "<cols><col min=\"1\" max=\"2147483647\" hidden=\"1\"/></cols>" +
            "<sheetData><row r=\"1\"><c r=\"A1\"><v>5</v></c><c r=\"B1\"><v>6</v></c></row></sheetData>");
        var cells = workbook.ReadRangeCells("Sheet1!$A$1:$B$1");
        TestAssert.Equal(2, cells.Length);
        TestAssert.True(cells[0].ColumnHidden, "Column A must report hidden.");
        TestAssert.True(cells[1].ColumnHidden, "Column B must report hidden.");
    }

    public static void NarrowHiddenColumnRunPreservesSemantics()
    {
        var workbook = ReadWorkbookFromSheetXml(
            "<cols><col min=\"1\" max=\"1\" hidden=\"1\"/></cols>" +
            "<sheetData><row r=\"1\"><c r=\"A1\"><v>5</v></c><c r=\"B1\"><v>6</v></c><c r=\"C1\"><v>7</v></c></row></sheetData>");
        var cells = workbook.ReadRangeCells("Sheet1!$A$1:$C$1");
        TestAssert.Equal(3, cells.Length);
        TestAssert.True(cells[0].ColumnHidden, "Column A must report hidden.");
        TestAssert.True(!cells[1].ColumnHidden, "Column B must report visible.");
        TestAssert.True(!cells[2].ColumnHidden, "Column C must report visible.");
    }

    public static void OutOfGridHiddenRowIsIgnored()
    {
        var workbook = ReadWorkbookFromSheetXml(
            "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row>" +
            "<row r=\"2\" hidden=\"1\"><c r=\"A2\"><v>2</v></c></row>" +
            "<row r=\"99999999\" hidden=\"1\"/></sheetData>");
        var cells = workbook.ReadRangeCells("Sheet1!$A$1:$A$2");
        TestAssert.Equal(2, cells.Length);
        TestAssert.True(!cells[0].RowHidden, "Row 1 must report visible.");
        TestAssert.True(cells[1].RowHidden, "Row 2 must report hidden.");
    }

    public static void HiddenColumnStorageStaysBoundedAcrossWideRuns()
    {
        static PptxRenderer.ChartWorkbookData Read(int last) => ReadWorkbookFromSheetXml(
            $"<cols><col min=\"4\" max=\"{last}\" hidden=\"1\"/></cols>" +
            "<sheetData><row r=\"1\"><c r=\"A1\"><v>5</v></c></row></sheetData>");
        Read(4);
        static long Allocated(int last)
        {
            long before = GC.GetAllocatedBytesForCurrentThread();
            var workbook = Read(last);
            long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
            GC.KeepAlive(workbook);
            return allocated;
        }
        long narrow = Allocated(4), wide = Allocated(16_383);
        TestAssert.True(wide < narrow + 128 * 1024,
            "Changing a range endpoint must not cause significant per-column construction growth.");
    }

    public static void HiddenColumnRunsIgnoreInvalidAndVisibleRanges()
    {
        foreach (string declaration in new[] { "", "<col min=\"1\" max=\"16384\" hidden=\"0\"/>",
            "<col min=\"0\" max=\"0\" hidden=\"1\"/>", "<col min=\"16385\" max=\"2147483647\" hidden=\"1\"/>",
            "<col min=\"8\" max=\"4\" hidden=\"1\"/>" })
        {
            var workbook = ReadWorkbookFromSheetXml("<cols>" + declaration +
                "</cols><sheetData><row r=\"1\"><c r=\"A1\"><v>5</v></c></row></sheetData>");
            TestAssert.True(!workbook.ReadRangeCells("Sheet1!A1").Single().ColumnHidden,
                "Absent, visible and invalid runs preserve a visible first column.");
            TestAssert.True(!workbook.ReadRangeCells("Sheet1!XFD1").Single().ColumnHidden,
                "Absent, visible and invalid runs preserve a visible last column.");
        }
    }

    public static void WorkbookSharedStringsBeyondBudgetThrow()
    {
        var strings = new System.Text.StringBuilder();
        for (int i = 0; i < 100001; i++)
        {
            strings.Append("<si><t>x</t></si>");
        }

        TestAssert.Throws<IOException>(() => ReadWorkbookFromSheetXml(
            "<sheetData><row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c></row></sheetData>",
            strings.ToString()));
    }

    public static void WorksheetCellsBeyondBudgetThrow()
    {
        TestAssert.Throws<IOException>(() => ReadWorkbookFromSheetXml(BuildManyCellSheet(100001)));
    }
    public static void WorkbookDefinedNamesBeyondBudgetThrow()
    {
        var names = new System.Text.StringBuilder();
        for (int i = 0; i < 100001; i++)
        {
            names.Append("<definedName name=\"N" + i + "\">Sheet1!$A$1</definedName>");
        }

        TestAssert.Throws<IOException>(() => ReadWorkbookPackage(
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
        var workbook = ReadWorkbookFromSheetXml(BuildManyCellSheet(100000));
        TestAssert.Equal(1, workbook.ReadRangeCells("Sheet1!$A$100000").Length);
    }

    public static void PptxWorkbookReplacementRetainsCombinedCellLimit()
    {
        string original = BuildManyCellSheet(60_000);
        var workbook = Read(40_000);
        TestAssert.Equal("1", workbook.ReadRangeCells("First!A40000").Single().Text);
        TestAssert.Equal(false, workbook.ReadRangeCells("First!A40001").Single().HasValue);
        TestAssert.Equal("1", workbook.ReadRangeCells("Tail!A60000").Single().Text);

        var failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => Read(40_001));
        TestAssert.Contains("maximum supported total cell count", failure.Message);

        PptxRenderer.ChartWorkbookData Read(int replacementCells) => ReadWorkbookPackage(
            """
            <sheets><sheet name="First" sheetId="1" r:id="first"/>
              <sheet name="fIRSt" sheetId="2" r:id="replacement"/>
              <sheet name="Tail" sheetId="3" r:id="tail"/></sheets>
            """,
            """
            <Relationship Id="first" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/first.xml"/>
            <Relationship Id="replacement" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/replacement.xml"/>
            <Relationship Id="tail" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/tail.xml"/>
            """,
            new Dictionary<string, string>
            {
                ["xl/worksheets/first.xml"] = original,
                ["xl/worksheets/replacement.xml"] = BuildManyCellSheet(replacementCells),
                ["xl/worksheets/tail.xml"] = original
            }, null);
    }

    public static void WorkbookSheetsBeyondBudgetThrow()
    {
        const int maxSheets = 1024;
        string preceding = string.Concat(Enumerable.Range(1, maxSheets - 1).Select(i =>
            $"<sheet name='Missing{i}' sheetId='{i}' r:id='unresolved'/>"));
        string tail = $"<sheet name=' Tail ' sheetId='{maxSheets}' r:id=' valid '/>";
        string localName = $"<definedNames><definedName name='AtLimit' localSheetId='{maxSheets - 1}'>Tail!$A$1</definedName></definedNames>";
        const string relationships = """
            <Relationship Id="valid" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/tail.xml"/>
            """;
        var sheetBodies = new Dictionary<string, string>
        {
            ["xl/worksheets/tail.xml"] = "<sheetData><row r='1'><c r='A1'><v>42</v></c></row></sheetData>"
        };

        var exact = ReadWorkbookPackage($"<sheets>{preceding}{tail}</sheets>{localName}", relationships, sheetBodies, null);
        TestAssert.Equal("42", exact.ReadRangeCells("Tail!AtLimit").Single().Text);
        var refusal = TestAssert.Throws<OoxPdfLimitExceededException>(() => ReadWorkbookPackage(
            $"<sheets>{preceding}{tail}<sheet name='TooMany' sheetId='1025' r:id='unresolved'/></sheets>{localName}",
            relationships, sheetBodies, null));
        TestAssert.Contains("more sheets than the maximum supported count", refusal.Message);
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
        TestAssert.Throws<IOException>(() => ReadWorkbookPackage(
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


    private static PptxRenderer.ChartWorkbookData ReadWorkbookFromSheetXml(string sheetBody, string? sharedStringsBody = null)
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

    internal static PptxRenderer.ChartWorkbookData ReadWorkbookCells(IReadOnlyDictionary<string, Dictionary<string, string>> sheets,
        CancellationToken cancellationToken = default)
    {
        XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XNamespace officeRels = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        XNamespace packageRels = "http://schemas.openxmlformats.org/package/2006/relationships";
        var sheetList = new XElement(spreadsheet+"sheets");
        var relationships = new List<XElement>();
        var bodies = new Dictionary<string, string>();
        foreach (var sheet in sheets)
        {
            int index = bodies.Count+1;
            string id = "rId"+index;
            sheetList.Add(new XElement(spreadsheet+"sheet", new XAttribute("name", sheet.Key),
                new XAttribute("sheetId", index), new XAttribute(officeRels+"id", id)));
            relationships.Add(new XElement(packageRels+"Relationship", new XAttribute("Id", id),
                new XAttribute("Type", officeRels.NamespaceName+"/worksheet"), new XAttribute("Target", "worksheets/sheet"+index+".xml")));
            bodies["xl/worksheets/sheet"+index+".xml"] = new XElement(spreadsheet+"sheetData", sheet.Value.Select(cell =>
                new XElement(spreadsheet+"row", new XAttribute("r", new string(cell.Key.Where(char.IsDigit).ToArray())),
                    new XElement(spreadsheet+"c", new XAttribute("r", cell.Key), new XElement(spreadsheet+"v", cell.Value))))).ToString();
        }
        return ReadWorkbookPackage(sheetList.ToString(), string.Concat(relationships), bodies, null,
            cancellationToken: cancellationToken);
    }

    private static PptxRenderer.ChartWorkbookData ReadWorkbookPackage(string sheetsFragment, string workbookRelsFragment, IReadOnlyDictionary<string, string> sheetBodies, string? sharedStringsBody, string? stylesBody = null,
        CancellationToken cancellationToken = default)
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
        using OoxPackage package = OoxPackage.Open(packageStream, CancellationToken.None);
        return ReadWorkbook(package, cancellationToken);
    }

    internal static PptxRenderer.ChartWorkbookData ReadWorkbook(OoxPackage package, CancellationToken cancellationToken = default) =>
        PptxRenderer.ReadWorkbookDataCore(package, cancellationToken)
            ?? throw new InvalidOperationException("Expected workbook data.");
}
