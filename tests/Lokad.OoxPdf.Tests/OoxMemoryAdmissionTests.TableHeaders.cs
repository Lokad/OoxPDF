using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxMemoryAdmissionTests
{
    public static void LayoutAdmissionRepeatedTableHeadersReuseCapacityAcrossPasses()
    {
        long narrow = LayoutTableFirstRequest(LayoutTableHeaderInput(1, cells: 1));
        var pool = new TestMemoryBudget(256L * 1024 * 1024);
        var measurer = LayoutAdmissionMeasurer();
        using var input = new MemoryStream(LayoutTableHeaderInput(100));
        using (var operation = OoxConversionBudget.BeginScope(null, pool))
        {
            var package = OoxPackage.Open(input, default);
            var document = new DocxReader().Read(package, null, default, OoxPdfDocxMarkupMode.Final);
            long initial = 0;
            pool.RefuseChange = (before, after) => { if (initial == 0 && after > before) initial = after; return false; };
            var first = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, measurer, default);
            TestAssert.True(initial > narrow, "Repeated-header content must affect capacity beyond the ordinary page frame.");
            TestAssert.Equal(100, first.Pages.Count);
            TestAssert.Equal(100, first.Pages.Sum(page => page.Items.OfType<DocxTableRowLayout>().Count(row => row.IsHeader)));
            long held = pool.Reserved;
            pool.Capacity = held;
            var second = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, measurer, default);
            TestAssert.Equal(first.Pages.Count, second.Pages.Count);
            TestAssert.Equal(held, pool.Reserved);
            package.Dispose();
            TestAssert.True(pool.Reserved > 0, "Keep both candidates funded after package readers finish.");
            TestAssert.Equal(64, first.Pages[^1].Items.OfType<DocxTableRowLayout>().First(row => row.IsHeader).Cells.Count);
            TestAssert.Equal(64, second.Pages[^1].Items.OfType<DocxTableRowLayout>().First(row => row.IsHeader).Cells.Count);
        }
        TestAssert.Equal(0L, pool.Reserved);
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        TestAssert.Equal(0, pool.CallerDisposals);
    }

    public static void LayoutAdmissionTableHeaderSummariesFollowSourceStructure()
    {
        long single = LayoutTableFirstRequest(LayoutTableHeaderInput(1, cells: 1));
        long wide = LayoutTableFirstRequest(LayoutTableHeaderInput(1));
        long nested = LayoutTableFirstRequest(LayoutTableHeaderInput(1, cells: 1, nestedCells: 16));
        long twoColumns = LayoutTableFirstRequest(LayoutTableHeaderInput(1, cells: 1, columns: 2));
        long customColumns = LayoutTableFirstRequest(LayoutTableHeaderInput(1, cells: 1, columns: 2, customColumns: 3));
        TestAssert.True(wide > single && nested > single, "Header cells and nested source content contribute to the shared summary.");
        TestAssert.True(twoColumns > single && customColumns > twoColumns, "Reserve for headers repeated in equal or custom column frames.");
        long noHeader = LayoutTableFirstRequest(LayoutTableHeaderInput(1, cells: 1, header: false));
        long lateFlag = LayoutTableFirstRequest(LayoutTableHeaderInput(1, cells: 1, header: false, lateHeader: true));
        TestAssert.Equal(noHeader, lateFlag); // Layout repeats only the leading header rows.
        TestAssert.True(single > noHeader, "Ordinary first rows are not repeated-header storage.");
    }

    public static void LayoutAdmissionRepeatedTableHeaderRefusalAndCancellationAreTerminal()
    {
        foreach (bool cancel in new[] { false, true })
        foreach (int stop in new[] { 1, 3 })
        {
            var pool = new TestMemoryBudget();
            using var cancellation = new CancellationTokenSource();
            using var input = new MemoryStream(LayoutTableHeaderInput(5, cells: 16));
            using (var operation = OoxConversionBudget.BeginScope(null, pool))
            using (var package = OoxPackage.Open(input, default))
            {
                var document = new DocxReader().Read(package, null, default, OoxPdfDocxMarkupMode.Final);
                int growths = 0;
                pool.RefuseChange = (before, after) => {
                    if (after <= before || ++growths != stop) return false;
                    if (cancel) cancellation.Cancel();
                    return !cancel;
                };
                var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
                if (cancel)
                    TestAssert.Equal(cancellation.Token, TestAssert.Throws<OperationCanceledException>(() => engine.Create(document, LayoutAdmissionMeasurer(), cancellation.Token)).CancellationToken);
                else
                    TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, TestAssert.Throws<OoxPdfLimitExceededException>(() => engine.Create(document, LayoutAdmissionMeasurer(), default)).Kind);
                TestAssert.Equal(stop, growths);
                pool.RefuseChange = null;
                TestAssert.Equal(5, engine.Create(document, LayoutAdmissionMeasurer(), default).Pages.Count);
            }
            TestAssert.Equal(0L, pool.Reserved);
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        }
    }

    public static void LayoutAdmissionRepeatedTableHeaderRefusalPreservesDestinations()
    {
        byte[] bytes = LayoutTableHeaderInput(4, cells: 12);
        long request = LayoutTableFirstRequest(bytes);
        string directory = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "ooxpdf-table-admission-" + Guid.NewGuid().ToString("N"))).FullName;
        try
        {
            string sourcePath = Path.Combine(directory, "source.docx"), target = Path.Combine(directory, "result.pdf");
            File.WriteAllBytes(sourcePath, bytes);
            foreach (bool file in new[] { false, true })
            {
                bool reached = false;
                var pool = new TestMemoryBudget { RefuseChange = (before, after) => {
                    if (before != 0 || after != request) return false;
                    reached = true;
                    return true;
                }};
                var options = new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx, MemoryBudget = pool,
                    FontResolver = new ProgramResolver(new ObservedFontProgram(() => { })) };
                File.WriteAllBytes(target, [1, 2, 3]);
                using var input = new MemoryStream(bytes);
                using var output = new MemoryStream();
                Task conversion = file ? OoxPdfConverter.ConvertAsync(sourcePath, target, options) : OoxPdfConverter.ConvertAsync(input, output, options);
                TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, TestAssert.Throws<OoxPdfLimitExceededException>(() => conversion.GetAwaiter().GetResult()).Kind);
                TestAssert.True(reached, "Refusal occurs at live repeated-header admission.");
                TestAssert.Equal(0L, output.Length);
                TestAssert.True(File.ReadAllBytes(target).SequenceEqual(new byte[] { 1, 2, 3 }), "Preserve the destination.");
                TestAssert.True(input.CanRead && output.CanWrite, "Caller streams stay open.");
                TestAssert.Equal(0, Directory.GetFiles(directory, "*.tmp-*").Length);
                TestAssert.Equal(0L, pool.Reserved);
                TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
            }
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    private static long LayoutTableFirstRequest(byte[] bytes)
    {
        var pool = new TestMemoryBudget();
        using var cancellation = new CancellationTokenSource();
        long request = 0;
        using (var operation = OoxConversionBudget.BeginScope(null, pool))
        using (var input = new MemoryStream(bytes))
        using (var package = OoxPackage.Open(input, default))
        {
            var document = new DocxReader().Read(package, null, default, OoxPdfDocxMarkupMode.Final);
            pool.RefuseChange = (before, after) => { if (after > before) { request = after; cancellation.Cancel(); } return false; };
            TestAssert.Equal(cancellation.Token, TestAssert.Throws<OperationCanceledException>(() => new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(document, (IDocxTextMeasurer?)null, cancellation.Token)).CancellationToken);
        }
        TestAssert.True(request > 0, "Observe the first live placement request before measurement/copying.");
        TestAssert.Equal(0L, pool.Reserved);
        return request;
    }

    private static byte[] LayoutTableHeaderInput(int pages, int cells = 64, int nestedCells = 0, int columns = 1, int customColumns = 0, bool header = true, bool lateHeader = false)
    {
        const string w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        string text = "<w:p><w:r><w:rPr><w:sz w:val=\"12\"/></w:rPr><w:t>A</w:t></w:r></w:p>";
        string nested = nestedCells == 0 ? "" : "<w:tbl><w:tblGrid>" + string.Concat(Enumerable.Repeat("<w:gridCol w:w=\"100\"/>", nestedCells)) + "</w:tblGrid><w:tr>" + string.Concat(Enumerable.Repeat("<w:tc>" + text + "</w:tc>", nestedCells)) + "</w:tr></w:tbl>";
        string head = string.Concat(Enumerable.Range(0, cells).Select(index => "<w:tc><w:tcPr><w:tcW w:w=\"146\" w:type=\"dxa\"/></w:tcPr>" + text + (index == 0 ? nested : "") + "</w:tc>"));
        string body = string.Concat(Enumerable.Range(0, pages).Select(index => "<w:tr><w:trPr><w:trHeight w:val=\"12000\" w:hRule=\"atLeast\"/><w:cantSplit/>" + (lateHeader && index == 0 ? "<w:tblHeader/>" : "") + "</w:trPr><w:tc><w:tcPr><w:gridSpan w:val=\"" + cells + "\"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr>"));
        string table = "<w:tbl><w:tblPr><w:tblW w:w=\"9360\" w:type=\"dxa\"/><w:tblCellMar><w:left w:w=\"0\" w:type=\"dxa\"/><w:right w:w=\"0\" w:type=\"dxa\"/></w:tblCellMar></w:tblPr><w:tblGrid>" + string.Concat(Enumerable.Repeat("<w:gridCol w:w=\"146\"/>", cells)) + "</w:tblGrid><w:tr><w:trPr>" + (header ? "<w:tblHeader/>" : "") + "</w:trPr>" + head + "</w:tr>" + body + "</w:tbl>";
        string cols = columns == 1 ? "" : "<w:cols w:num=\"" + columns + "\" w:equalWidth=\"" + (customColumns > 0 ? "0" : "1") + "\" w:space=\"0\">" + string.Concat(Enumerable.Repeat("<w:col w:w=\"2000\"/>", customColumns)) + "</w:cols>";
        using var input = TestFixtures.CreateZipPackage(new Dictionary<string, string> {
            ["[Content_Types].xml"] = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>",
            ["word/document.xml"] = $"<w:document xmlns:w=\"{w}\"><w:body>{table}<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:right=\"1440\"/>{cols}</w:sectPr></w:body></w:document>"
        });
        return input.ToArray();
    }
}
