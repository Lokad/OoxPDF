using System.Globalization;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxRepeatedTableHeaderTests
{
    public static void DocxTableHeaderGroupsAreAdmittedWithTheirKeptHeading()
    {
        foreach (var mode in Modes)
        foreach (int headers in new[] { 2, 3 })
        foreach (double height in new[] { 20d, 30d })
        foreach (bool keep in new[] { false, true })
        foreach (bool boundary in new[] { false, true })
        {
            var source = HeaderGroupDocument(headers, height, boundary, keep, heading: true);
            var layout = Layout(source, mode);
            int sourcePage = boundary ? 2 : 1;
            var tablePages = layout.Pages.Select((page, index) => (page, number: index + 1))
                .Where(pair => pair.page.Items.OfType<DocxTableRowLayout>().Any(row => row.IsHeader)).ToArray();
            TestAssert.Equal(sourcePage, tablePages[0].number);
            foreach (var (page, number) in tablePages)
            {
                var rows = page.Items.OfType<DocxTableRowLayout>().ToArray();
                TestAssert.Equal(headers, rows.Count(row => row.IsHeader));
                TestAssert.True(rows.Take(headers).Select(row => row.RowIndex).SequenceEqual(Enumerable.Range(0, headers)),
                    "Every admitted table page starts with the complete ordered header group.");
                TestAssert.True(rows.Where(row => row.IsHeader).All(row => Text(row) == "H" + row.RowIndex + "=" + sourcePage),
                    "Repeated fields retain the page where the complete source group was admitted.");
            }
            int headingPage = layout.Pages.Select((page, index) => (page, number: index + 1))
                .Single(pair => pair.page.Items.OfType<DocxTextLineLayout>().Any(line => line.Text == "Kept heading")).number;
            TestAssert.Equal(boundary && keep ? 2 : 1, headingPage);
        }
    }

    public static void DocxTableHeaderGroupCountersBindAfterCompleteAdmission()
    {
        foreach (var mode in Modes)
        foreach (int headers in new[] { 2, 3 })
        foreach (int precedingPages in new[] { 0, 8 })
        {
            var source = HeaderGroupDocument(headers, 20d, boundary: true, keep: false, heading: false, precedingPages);
            var layout = Layout(source, mode);
            int sourcePage = precedingPages + 2;
            foreach (var (page, index) in layout.Pages.Select((page, index) => (page, index)))
            foreach (var row in page.Items.OfType<DocxTableRowLayout>().Where(row => row.Cells[0].Cell.Text != "Filler"))
                TestAssert.Equal(row.IsHeader ? "H" + row.RowIndex + "=" + sourcePage : "B=" + (index + 1), Text(row));
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => new DocxLayoutEngine(mode).Create(source, new Measurer(), canceled.Token));
        }
    }

    public static void DocxTableHeaderGroupsRetainExplicitCellBreakAdmission()
    {
        foreach (var mode in Modes)
        {
            var source = HeaderGroupDocument(2, 30d, boundary: true, keep: false, heading: false);
            var table = source.Tables[^1];
            var before = DocxTests.CreateDocxLayoutParagraph("Explicit before", 10d, 10d);
            var after = DocxTests.CreateDocxLayoutParagraph("Explicit after", 10d, 10d);
            var cell = table.Rows[0].Cells[0] with { Paragraphs = [], BodyElements = [new DocxParagraphElement(before),
                new DocxPageBreakElement(DocxBreakSourceKind.PageBreakBefore, "1", null), new DocxParagraphElement(after)] };
            var changed = table with { Rows = [table.Rows[0] with { Cells = [cell], CantSplit = false }, .. table.Rows.Skip(1)] };
            source = source with
            {
                BodyElements = source.BodyElements.Select(block => block is DocxTableElement entry && ReferenceEquals(entry.Table, table)
                    ? new DocxTableElement(changed) : block).ToArray()
            };
            var layout = Layout(source, mode);
            int PageOf(string value) => layout.Pages.Select((page, index) => (page, number: index + 1))
                .Single(pair => pair.page.Items.OfType<DocxTableRowLayout>().Any(row => Text(row).Contains(value, StringComparison.Ordinal))).number;
            TestAssert.Equal(1, PageOf("Explicit before"));
            TestAssert.Equal(2, PageOf("Explicit after"));
        }
    }

    private static DocxDocument HeaderGroupDocument(int headers, double headerHeight, bool boundary,
        bool keep, bool heading, int precedingPages = 0)
    {
        var rows = new List<DocxTableRow>();
        for (int index = 0; index < headers + 8; index++)
        {
            var paragraph = DocxTests.CreateDocxLayoutParagraph(index < headers ? "H" + index + "={PAGE}" : "B={PAGE}", 10d, 10d);
            rows.Add(new([Cell(paragraph)], index < headers ? headerHeight : 40d)
                { IsHeader = index < headers, CantSplit = true, HeightRuleValue = "exact" });
        }
        var table = new DocxTable("fixed", [80d], rows);
        var blocks = new List<DocxBodyElement>();
        for (int index = 0; index < precedingPages; index++)
        {
            blocks.Add(new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Before", 10d, 10d)));
            blocks.Add(new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null));
        }
        if (boundary)
        {
            var filler = new DocxTable("fixed", [80d], [new DocxTableRow(
                [Cell(DocxTests.CreateDocxLayoutParagraph("Filler", 10d, 10d))], 190d - headers * headerHeight)
                { HeightRuleValue = "exact", CantSplit = true }]);
            blocks.Add(new DocxTableElement(filler));
        }
        if (heading) blocks.Add(new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Kept heading", 10d, 10d) with
            { KeepRules = DocxParagraphKeepRules.Empty with { KeepNext = keep } }));
        blocks.Add(new DocxTableElement(table));
        return DocxTests.CreateLayoutTestDocument(blocks, blocks.OfType<DocxTableElement>().Select(block => block.Table).ToArray());
    }

    public static void DocxRepeatedTableHeaderFieldsRetainSourcePagesAndBodyPlacement()
    {
        foreach (var mode in Modes)
        foreach (int precedingPages in new[] { 0, 8 })
        foreach (int headerCount in new[] { 1, 2 })
        {
            var table = Table(headerCount);
            var layout = Layout(Document(table, precedingPages), mode);
            var pages = layout.Pages.Select((page, index) => (page, number: index + 1))
                .Where(pair => pair.page.Items.OfType<DocxTableRowLayout>().Any()).ToArray();
            TestAssert.True(pages.Length > 1, "The source table must cross a page boundary.");
            foreach (var (page, number) in pages)
            foreach (var row in page.Items.OfType<DocxTableRowLayout>())
            {
                string text = Text(row);
                if (row.IsHeader)
                {
                    string expected = "H" + row.RowIndex + "=" + pages[0].number.ToString(CultureInfo.InvariantCulture);
                    TestAssert.Equal(expected, text);
                    TestAssert.True(row.Cells[0].TextLines.All(line =>
                        ReferenceEquals(line.SourceParagraph, table.Rows[row.RowIndex].Cells[0].Paragraphs[0])),
                        "Repeated content retains the original paragraph owner.");
                }
                else TestAssert.Equal("B=" + number.ToString(CultureInfo.InvariantCulture), text);
            }
        }
    }

    public static void DocxRepeatedTableHeaderFieldsBindAfterPageAdmission()
    {
        foreach (var mode in Modes)
        {
            var table = Table(1);
            // A single paragraph's large line pitch does not establish its
            // occupied trailing frame. Use an explicit fixed row as the filler.
            var before = new DocxTable("fixed", [80d], [new DocxTableRow(
                [Cell(DocxTests.CreateDocxLayoutParagraph("Before", 10d, 10d))], 175d)
                { HeightRuleValue = "exact", HeightValue = "3500", CantSplit = true }]);
            var source = DocxTests.CreateLayoutTestDocument(
                [new DocxTableElement(before), new DocxTableElement(table)], [before, table]);
            var layout = Layout(source, mode);
            TestAssert.Equal(175d, layout.Pages[0].Items.OfType<DocxTableRowLayout>().Single().Height);
            var headers = layout.Pages.SelectMany(p => p.Items.OfType<DocxTableRowLayout>().Where(r => r.IsHeader)).ToArray();
            TestAssert.True(headers.Length > 1, "The deferred table still repeats its source header.");
            TestAssert.True(headers.All(r => Text(r) == "H0=2"),
                "Bind the admitted source page, not the rejected first-page proposal.");
        }
    }

    public static void DocxRepeatedTableHeaderFieldsKeepMeasuredWidthAcrossDigitBoundary()
    {
        foreach (var mode in Modes)
        foreach (string layoutKind in new[] { "fixed", "autofit" })
        {
            var table = Table(1, narrow: true) with { LayoutValue = layoutKind };
            var layout = Layout(Document(table, 8), mode);
            var headers = layout.Pages.SelectMany(p => p.Items.OfType<DocxTableRowLayout>().Where(r => r.IsHeader)).ToArray();
            TestAssert.True(headers.Length > 1, "Exercise a repeated source row after page9.");
            TestAssert.True(headers.All(r => Text(r) == "H9"), "The source value remains a single digit.");
            foreach (var header in headers.Skip(1))
            {
                TestAssert.Equal(headers[0].Height, header.Height);
                TestAssert.Equal(headers[0].Cells[0].Width, header.Cells[0].Width);
                TestAssert.Equal(headers[0].Cells[0].TextLines.Count, header.Cells[0].TextLines.Count);
            }
        }
    }

    public static void DocxRepeatedTableHeaderFieldsSurviveSplitRowsAndKeepPhysicalImages()
    {
        foreach (var mode in Modes)
        {
            var table = Table(1);
            var header = table.Rows[0];
            var paragraph = header.Cells[0].Paragraphs[0];
            var image = new DocxInlineImage(8d, 8d, "image/png", null, "header.png") { SourceRunIndex = 0 };
            paragraph = paragraph with { Images = [image] };
            header = header with { Cells = [header.Cells[0] with { Paragraphs = [paragraph] }] };
            var body = DocxTests.CreateDocxLayoutParagraph(string.Join(" ", Enumerable.Repeat("word", 120)), 10d, 10d);
            var row = new DocxTableRow([Cell(body)], null);
            table = table with { Rows = [header, row] };
            var layout = Layout(Document(table, 0), mode);
            int repeated = 0;
            foreach (var (page, index) in layout.Pages.Select((p, i) => (p, i)))
            foreach (var placed in page.Items.OfType<DocxTableRowLayout>().Where(r => r.IsHeader))
            {
                repeated++;
                TestAssert.Equal("H0=1", Text(placed));
                TestAssert.Equal(index, placed.Cells[0].InlineImages.Single().PageIndex);
            }
            TestAssert.True(repeated > 1 && layout.Pages.SelectMany(p => p.Items.OfType<DocxTableRowLayout>())
                .Any(r => !r.IsHeader && r.FragmentCount > 1), "Exercise actual body-row continuations.");
        }
    }

    public static void DocxRepeatedTableHeaderNestedFieldsKeepPhysicalImagePages()
    {
        foreach (var mode in Modes)
        {
            var paragraph = DocxTests.CreateDocxLayoutParagraph("Nested={PAGE}", 10d, 10d) with
            { Images = [new DocxInlineImage(8d, 8d, "image/png", null, "nested.png") { SourceRunIndex = 0 }] };
            var nested = new DocxTable("fixed", [80d], [new DocxTableRow([Cell(paragraph)], 20d)]);
            var table = Table(1);
            var headerCell = table.Rows[0].Cells[0] with { Paragraphs = [], BodyElements = [new DocxTableElement(nested)] };
            table = table with { Rows = [table.Rows[0] with { Cells = [headerCell] }, .. table.Rows.Skip(1)] };
            var layout = Layout(Document(table, 0), mode);
            int count = 0;
            foreach (var (page, index) in layout.Pages.Select((p, i) => (p, i)))
            foreach (var header in page.Items.OfType<DocxTableRowLayout>().Where(r => r.IsHeader))
            {
                var row = header.Cells[0].NestedRows.Single();
                TestAssert.Equal("Nested=1", Text(row));
                TestAssert.Equal(index, row.Cells[0].InlineImages.Single().PageIndex);
                count++;
            }
            TestAssert.True(count > 1, "Nested source content must repeat on a later physical page.");
        }
    }

    private static readonly OoxPdfDocxMarkupGeometryMode[] Modes =
    [OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup];

    private sealed class Measurer : IDocxTextMeasurer, IDocxStaticTextMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double size) => text.Length * 5d;
        public double MeasureWindowsAscender(DocxTextRun? run, double size) => 8d;
        public double MeasureWindowsDescender(DocxTextRun? run, double size) => 2d;
    }

    private static DocxLayout Layout(DocxDocument document, OoxPdfDocxMarkupGeometryMode mode) =>
        // Keep admission geometry fixed across modes in these semantic tests.
        // Independent source/PDF controls exercise the actual public print scale.
        new DocxLayoutEngine(mode, wordCompatiblePrintScale: 1d, knownPageCount: 20).Create(document, new Measurer(), default);

    private static DocxTableCell Cell(DocxParagraph paragraph) =>
        new(paragraph.Runs[0].Text, [paragraph], null, null, null, null, [],
            DocxTableCellMargins.Empty with { LeftPoints = 0d, RightPoints = 0d, TopPoints = 0d, BottomPoints = 0d });

    private static DocxTable Table(int headers, bool narrow = false)
    {
        var rows = new List<DocxTableRow>();
        for (int i = 0; i < headers + 8; i++)
        {
            var paragraph = DocxTests.CreateDocxLayoutParagraph(i < headers
                ? narrow ? "H{PAGE}" : "H" + i + "={PAGE}" : narrow ? "B" : "B={PAGE}", 10d, 10d);
            rows.Add(new([Cell(paragraph)], i < headers ? null : 40d) { IsHeader = i < headers, CantSplit = true });
        }
        return new("fixed", [narrow ? 11d : 80d], rows);
    }

    private static DocxDocument Document(DocxTable table, int precedingPages)
    {
        var blocks = new List<DocxBodyElement>();
        for (int i = 0; i < precedingPages; i++)
        {
            blocks.Add(new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Before", 10d, 10d)));
            blocks.Add(new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null));
        }
        blocks.Add(new DocxTableElement(table));
        return DocxTests.CreateLayoutTestDocument(blocks, [table]);
    }

    private static string Text(DocxTableRowLayout row) => string.Concat(row.Cells[0].TextLines.Select(line => line.Text));
}
