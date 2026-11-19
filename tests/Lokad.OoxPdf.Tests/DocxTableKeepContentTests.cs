using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Tests;

internal static class DocxTableKeepContentTests
{
    public static void DocxContinuingTableUsesTheActivePageNoteBoundary()
    {
        const string word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const string relations = "http://schemas.openxmlformats.org/package/2006/relationships";
        const string relationType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var resolver = new PortableRenderingTests.PinnedResolver();
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (int noteLines in new[] { 0, 5, 8 })
        {
            string reference = noteLines == 0 ? "" : """<w:r><w:footnoteReference w:id="1"/></w:r>""";
            string note = string.Concat(Enumerable.Range(1, noteLines).Select(index => ParagraphXml($"NOTE LINE {index}")));
            using var input = TestFixtures.CreateZipPackage(new Dictionary<string, string>
            {
                ["[Content_Types].xml"] = """
                    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                      <Default Extension="xml" ContentType="application/xml"/>
                      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                      <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
                    </Types>
                    """,
                ["_rels/.rels"] = $"""<Relationships xmlns="{relations}"><Relationship Id="doc" Type="{relationType}/officeDocument" Target="word/document.xml"/></Relationships>""",
                ["word/_rels/document.xml.rels"] = $"""<Relationships xmlns="{relations}"><Relationship Id="notes" Type="{relationType}/footnotes" Target="footnotes.xml"/></Relationships>""",
                ["word/footnotes.xml"] = $"""<w:footnotes xmlns:w="{word}"><w:footnote w:id="1">{note}</w:footnote></w:footnotes>""",
                ["word/document.xml"] = $$"""
                    <w:document xmlns:w="{{word}}"><w:body>
                      {{ParagraphXml("PRECEDING OWNER", extraRun: reference)}}
                      {{ParagraphXml("PRECEDING SPACE", lineHeight: 2400)}}
                      <w:tbl>
                        <w:tblPr><w:tblW w:w="7200" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>
                        <w:tblGrid><w:gridCol w:w="7200"/></w:tblGrid>
                        {{RowXml("ROW ALPHA")}}{{RowXml("ROW BRAVO")}}
                      </w:tbl>
                      {{ParagraphXml("FOLLOWING OWNER")}}
                      <w:sectPr><w:pgSz w:w="12240" w:h="6000"/>
                        <w:pgMar w:top="480" w:bottom="480" w:left="720" w:right="720" w:header="240" w:footer="240" w:gutter="0"/>
                      </w:sectPr>
                    </w:body></w:document>
                    """
            });
            using var package = OoxPackage.Open(input, default);
            var document = new DocxReader().Read(package, null, default, markupMode: mode);
            var geometry = mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
            var pages = new DocxRenderer(resolver, mode, geometry)
                .RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();

            TestAssert.Equal(2, pages.Length);
            string[] pageText = pages.Select(DocxPageFootnoteNumberingTests.DecodeText).ToArray();
            int alphaPage = noteLines == 0 ? 0 : 1;
            foreach (var (label, page) in new[] { ("ROW ALPHA", alphaPage), ("ROW BRAVO", 1), ("FOLLOWING OWNER", 1) })
            {
                PortableRenderingTests.ReadPaintedSourceLabel(pages[page], label);
                TestAssert.DoesNotContain(label, pageText[1 - page]);
            }
            if (noteLines != 0)
            {
                string noteText = string.Join(" ", Enumerable.Range(1, noteLines).Select(index => $"NOTE LINE {index}"));
                PortableRenderingTests.ReadPaintedSourceLabel(pages[0], noteText);
                TestAssert.DoesNotContain("NOTE LINE", pageText[1]);
            }
        }

        static string ParagraphXml(string text, int lineHeight = 240, string extraRun = "") => $"""
            <w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="{lineHeight}" w:lineRule="exact"/></w:pPr>
              <w:r><w:t>{text}</w:t></w:r>{extraRun}
            </w:p>
            """;
        static string RowXml(string text) => $"""
            <w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="2100" w:hRule="exact"/></w:trPr>
              <w:tc><w:tcPr><w:tcW w:w="7200" w:type="dxa"/></w:tcPr>{ParagraphXml(text)}</w:tc>
            </w:tr>
            """;
    }

    public static void DocxOversizedCellKeepLinesPreservesSourceFrames()
    {
        foreach (var mode in Modes)
        foreach (int count in new[] { 12, 52 })
        foreach (bool keep in new[] { false, true })
        foreach (bool prefix in new[] { false, true })
        {
            var paragraph = Paragraph(count, keep);
            var source = Document([paragraph], prefix);
            var layout = Layout(source, mode);
            AssertLines(layout, paragraph, count);
            if (count == 52) TestAssert.True(layout.Pages.Count >= 3, "Oversized prose flows across enough pages to retain every source line.");
            if (count == 12 && keep)
            {
                var rows = layout.Pages.SelectMany(page => page.Items.OfType<DocxTableRowLayout>())
                    .Where(row => row.RowIndex == (prefix ? 1 : 0)).ToArray();
                TestAssert.Equal(1, rows.Length);
                TestAssert.Equal(1, rows[0].FragmentCount);
            }
        }
    }

    public static void DocxOversizedCellKeepLinesRetainsNeighborsAndCancellation()
    {
        foreach (var mode in Modes)
        {
            var heading = Paragraph(4, keep: true, label: "H");
            var content = Paragraph(52, keep: true);
            var source = Document([heading, content], prefix: true);
            var layout = Layout(source, mode);
            AssertLines(layout, content, 52);
            var headingPages = layout.Pages.Select((page, index) => (page, index)).Where(pair => pair.page.Items
                .OfType<DocxTableRowLayout>().SelectMany(row => row.Cells).SelectMany(cell => cell.TextLines)
                .Any(line => ReferenceEquals(line.SourceParagraph, heading))).ToArray();
            TestAssert.Equal(1, headingPages.Length);
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => new DocxLayoutEngine(mode, wordCompatiblePrintScale: 1d)
                .Create(source, new Measurer(), canceled.Token));
        }
    }

    public static void DocxDeclaredKeptRowsRetainTheirAtomicFallback()
    {
        foreach (var mode in Modes)
        foreach (string rule in new[] { "exact", "atLeast", "header" })
        foreach (bool nextOnly in new[] { false, true })
        {
            var paragraph = Paragraph(52, keep: !nextOnly);
            if (nextOnly) paragraph = paragraph with { KeepRules = paragraph.KeepRules with { KeepNext = true } };
            var source = Document([paragraph], prefix: false, declaredRule: rule);
            var layout = Layout(source, mode);
            var row = layout.Pages.SelectMany(page => page.Items.OfType<DocxTableRowLayout>()).Single();
            if (rule == "header") TestAssert.True(row.IsHeader && row.Height > 180d, "An oversized repeating header retains its existing atomic fallback.");
            else TestAssert.Equal(600d, row.Height);
            TestAssert.Equal(1, row.FragmentCount);
        }
    }

    public static void DocxOversizedCellKeepNextPreservesSourceFrames()
    {
        foreach (var mode in Modes)
        foreach (bool prefix in new[] { false, true })
        foreach (bool following in new[] { false, true })
        foreach (bool keepLines in new[] { false, true })
        {
            var paragraph = Paragraph(52, keepLines);
            paragraph = paragraph with { KeepRules = paragraph.KeepRules with { KeepNext = true } };
            var target = Paragraph(4, keep: false, label: "F");
            var source = Document(following ? [paragraph, target] : [paragraph], prefix);
            var layout = Layout(source, mode);
            AssertLines(layout, paragraph, 52);
            if (following) AssertLines(layout, target, 4, "F");
        }
    }

    public static void DocxFittingCellKeepNextPreservesTargetAndPageAdmission()
    {
        foreach (var mode in Modes)
        {
            var paragraph = Paragraph(12, keep: false);
            paragraph = paragraph with { KeepRules = paragraph.KeepRules with { KeepNext = true } };
            var target = Paragraph(4, keep: true, label: "F");
            var layout = Layout(Document([paragraph, target], prefix: true), mode);
            var rows = layout.Pages.SelectMany(page => page.Items.OfType<DocxTableRowLayout>()).Where(row => row.RowIndex == 1).ToArray();
            TestAssert.Equal(1, rows.Length);
            TestAssert.True(layout.Pages[1].Items.Contains(rows[0]), "A fitting keep group is admitted together after the page boundary.");
            AssertLines(layout, paragraph, 12);
            AssertLines(layout, target, 4, "F");
        }
    }

    public static void DocxFittingKeptCellParagraphsUseCompleteSourceBoundaries()
    {
        foreach (var mode in Modes)
        foreach (bool prefix in new[] { false, true })
        {
            var paragraphs = new[] { Paragraph(16, true, "A"), Paragraph(16, true, "B"), Paragraph(16, true, "C") };
            var layout = Layout(Document(paragraphs, prefix), mode);
            TestAssert.Equal(prefix ? 4 : 3, layout.Pages.Count);
            for (int index = 0; index < paragraphs.Length; index++)
            {
                var paragraph = paragraphs[index];
                AssertLines(layout, paragraph, 16, "ABC"[index].ToString());
                TestAssert.Equal(1, layout.Pages.Count(page => page.Items.OfType<DocxTableRowLayout>()
                    .SelectMany(row => row.Cells).SelectMany(cell => cell.TextLines)
                    .Any(line => ReferenceEquals(line.SourceParagraph, paragraph))));
            }
        }
    }

    public static void DocxFittingKeptCellParagraphsPreserveFollowingGroups()
    {
        foreach (var mode in Modes)
        foreach (bool prefix in new[] { false, true })
        {
            var heading = Paragraph(8, true, "H");
            heading = heading with { KeepRules = heading.KeepRules with { KeepNext = true } };
            var target = Paragraph(8, true, "T");
            var following = Paragraph(16, true, "F");
            var layout = Layout(Document([heading, target, following], prefix), mode);
            TestAssert.Equal(prefix ? 3 : 2, layout.Pages.Count);
            AssertLines(layout, heading, 8, "H");
            AssertLines(layout, target, 8, "T");
            AssertLines(layout, following, 16, "F");
            var first = layout.Pages[prefix ? 1 : 0].Items.OfType<DocxTableRowLayout>()
                .SelectMany(row => row.Cells).SelectMany(cell => cell.TextLines).ToArray();
            TestAssert.True(first.Any(line => ReferenceEquals(line.SourceParagraph, heading)) &&
                first.Any(line => ReferenceEquals(line.SourceParagraph, target)) &&
                !first.Any(line => ReferenceEquals(line.SourceParagraph, following)),
                "A fitting next-target group remains together while its following paragraph starts on another page.");
        }
    }

    public static void DocxFittingKeptCellParagraphsRetainUnqualifiedAtomicFallbacks()
    {
        foreach (var mode in Modes)
        foreach (string rule in new[] { "exact", "atLeast", "header" })
        foreach (bool next in new[] { false, true })
        {
            var paragraphs = Enumerable.Range(0, 3).Select(index => Paragraph(16, true, index.ToString())).ToArray();
            if (next)
                for (int index = 0; index < 2; index++)
                    paragraphs[index] = paragraphs[index] with { KeepRules = paragraphs[index].KeepRules with { KeepNext = true } };
            var layout = Layout(Document(paragraphs, false, rule), mode);
            var row = layout.Pages.SelectMany(page => page.Items.OfType<DocxTableRowLayout>()).Single();
            TestAssert.Equal(1, row.FragmentCount);
            TestAssert.True(row.Height > 180d, "Unqualified header and declared-height fallbacks remain explicit; this does not claim visible completeness.");
        }
    }

    public static void DocxUnfitCellKeepNextGroupsPreserveCompleteParagraphs()
    {
        foreach (var mode in Modes)
        foreach (bool prefix in new[] { false, true })
        foreach (int[] counts in new int[][] { [16, 16, 16], [4, 14, 4], [5, 5, 9], [2, 8, 10], [4, 52, 4] })
        {
            var paragraphs = counts.Select((count, index) => Paragraph(count, true, "ABC"[index].ToString())).ToArray();
            for (int index = 0; index < 2; index++)
                paragraphs[index] = paragraphs[index] with { KeepRules = paragraphs[index].KeepRules with { KeepNext = true } };
            var layout = Layout(Document(paragraphs, prefix), mode);
            for (int index = 0; index < paragraphs.Length; index++)
            {
                var paragraph = paragraphs[index];
                AssertLines(layout, paragraph, counts[index], "ABC"[index].ToString());
                if (counts[index] <= 18)
                    TestAssert.Equal(1, layout.Pages.Count(page => page.Items.OfType<DocxTableRowLayout>()
                        .SelectMany(row => row.Cells).SelectMany(cell => cell.TextLines)
                        .Any(line => ReferenceEquals(line.SourceParagraph, paragraph))));
            }
        }
    }

    public static void DocxCellKeepGroupsRespectEmptyContentAfterHeadersAndPadding()
    {
        foreach (var mode in Modes)
        foreach (int headerLines in new[] { 0, 3 })
        foreach (double padding in new[] { 0d, 15d })
        {
            var paragraphs = Enumerable.Range(0, 3).Select(index => Paragraph(6, true, "ABC"[index].ToString())).ToArray();
            for (int index = 0; index < 2; index++)
                paragraphs[index] = paragraphs[index] with { KeepRules = paragraphs[index].KeepRules with { KeepNext = true } };
            var layout = Layout(Document(paragraphs, false, headerLines: headerLines, verticalPadding: padding), mode);
            TestAssert.Equal(headerLines == 0 && padding == 0d ? 1 : 2, layout.Pages.Count);
            for (int index = 0; index < paragraphs.Length; index++)
                AssertLines(layout, paragraphs[index], 6, "ABC"[index].ToString());
            if (headerLines != 0)
                TestAssert.True(layout.Pages.All(page => page.Items.OfType<DocxTableRowLayout>().Any(row => row.IsHeader)),
                    "Each admitted continuation keeps its persistent header without hiding body source frames.");
        }
    }

    private static void AssertLines(DocxLayout layout, DocxParagraph paragraph, int count, string label = "L")
    {
        var lines = layout.Pages.SelectMany(page => page.Items.OfType<DocxTableRowLayout>())
            .SelectMany(row => row.Cells).SelectMany(cell => cell.TextLines)
            .Where(line => ReferenceEquals(line.SourceParagraph, paragraph)).ToArray();
        TestAssert.True(lines.Select(line => line.Text.Trim()).SequenceEqual(Enumerable.Range(0, count).Select(index => label + index.ToString("000"))),
            "Each original line retains its owner and paints exactly once in source order.");
        foreach (var line in lines)
        {
            TestAssert.True(DocxInlineLineGeometry.TryGetOccupiedBounds(line, out double lower, out double upper), "The source line has a complete measured frame.");
            TestAssert.True(lower >= 10d - .001d && upper <= 190d + .001d, "Every admitted source frame stays inside the physical page body.");
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

    private static DocxLayout Layout(DocxDocument source, OoxPdfDocxMarkupGeometryMode mode) =>
        new DocxLayoutEngine(mode, wordCompatiblePrintScale: 1d).Create(source, new Measurer(), default);

    private static DocxParagraph Paragraph(int count, bool keep, string label = "L") =>
        DocxTests.CreateDocxLayoutParagraph(string.Join('\n', Enumerable.Range(0, count).Select(index => label + index.ToString("000"))),
            10d, 10d, keepRules: DocxParagraphKeepRules.Empty with { KeepLines = keep, WidowControl = false });

    private static DocxDocument Document(IReadOnlyList<DocxParagraph> paragraphs, bool prefix, string? declaredRule = null,
        int headerLines = 0, double verticalPadding = 0d)
    {
        var cells = new DocxTableCell(string.Empty, paragraphs, null, null, null, null, [],
            DocxTableCellMargins.Empty with { LeftPoints = 0d, RightPoints = 0d, TopPoints = verticalPadding, BottomPoints = verticalPadding });
        var rows = new List<DocxTableRow>();
        if (headerLines != 0)
            rows.Add(new([new DocxTableCell(string.Empty, [Paragraph(headerLines, false, "R")], null, null, null, null, [],
                DocxTableCellMargins.Empty with { LeftPoints = 0d, RightPoints = 0d, TopPoints = 0d, BottomPoints = 0d })], null) { IsHeader = true });
        if (prefix) rows.Add(new([new DocxTableCell("Prefix", [], null, null, null, null, [], DocxTableCellMargins.Empty)], 90d)
            { HeightRuleValue = "exact", CantSplit = true });
        rows.Add(new([cells], declaredRule is null or "header" ? null : 600d)
            { HeightRuleValue = declaredRule == "header" ? null : declaredRule, IsHeader = declaredRule == "header" });
        var table = new DocxTable("fixed", [160d], rows);
        return DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);
    }
}
