using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxNoteMathFieldsStaticTablesPreserveCompleteMixedContent()
    {
        foreach (string position in new[] { "docEnd", "sectEnd" })
        foreach (bool direct in new[] { false, true })
        foreach (bool footer in new[] { false, true })
        foreach (string shape in new[] { "row", "two-rows", "nested", "math-picture", "standalone" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            int count = footer ? 3 : 2, reference = footer ? 2 : 1;
            string value = reference.ToString(System.Globalization.CultureInfo.InvariantCulture);
            DocxDocument Source(string? field, string cached, bool table)
            {
                var source = ReadNoteMathField("endnote", mode, field, cached, direct);
                source = source with
                {
                    Settings = source.Settings with { EndnoteReferenceSettings =
                        source.Settings.EndnoteReferenceSettings with { PositionValue = position } },
                    BodyElements = footer
                        ? [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Public earlier page", 12d, 14d)),
                            new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null), .. source.BodyElements]
                        : source.BodyElements
                };
                if (!table) return source;
                var paragraph = shape is "math-picture" or "standalone"
                    ? StaticNoteEquationParagraph(source, mode, shape == "standalone")
                    : PlainNoteMathStaticParagraph();
                if (shape == "math-picture")
                {
                    paragraph = NoteMathStaticPictureParagraph(source, paragraph, false);
                    var equation = paragraph.Equations[0] with { SourceRunIndex = paragraph.Runs.Max(run => run.SourceRunIndex) + 1 };
                    paragraph = paragraph with { Equations = [equation], InlineContent =
                        [.. paragraph.InlineContent!, new DocxInlineEquation(equation)] };
                }
                var block = new DocxTableElement(StaticNoteTable(paragraph, shape));
                source = WithNoteMathStaticBlocks(source, footer, [block]);
                return source with { EquationParts = [.. source.EquationParts,
                    new DocxEquationPartSource(footer ? "/word/public-table-footer.xml" : "/word/public-table-header.xml",
                        paragraph.Equations.Count, [block])] };
            }
            var renderer = PictureRangeRenderer(mode);
            var baseline = renderer.RenderPagesAsync(Source(null, value, false), null, default).ToBlockingEnumerable().ToArray();
            var literal = renderer.RenderPagesAsync(Source(null, value, true), null, default).ToBlockingEnumerable().ToArray();
            var diagnostics = new List<Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic>();
            var counter = renderer.RenderPagesAsync(Source("PAGE", "9", true), diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            string context = $"{position}/{direct}/{footer}/{shape}/{mode}";
            TestAssert.Equal(count, counter.Length); TestAssert.Equal(count, literal.Length); TestAssert.Equal(count, baseline.Length);
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "A complete static table cannot discard supported note math: " + context);
            var noteText = new System.Text.StringBuilder();
            for (int page = 0; page < count; page++)
            {
                TestAssert.Equal(NormalizeNoteMathFieldText(DocxPageFootnoteNumberingTests.DecodeText(literal[page])),
                    NormalizeNoteMathFieldText(DocxPageFootnoteNumberingTests.DecodeText(counter[page])));
                var a = StaticNoteMathShows(baseline[page], false); var b = StaticNoteMathShows(literal[page], false);
                var c = StaticNoteMathShows(counter[page], false);
                TestAssert.True(a.SequenceEqual(b) && a.SequenceEqual(c), "Unrelated static containers preserve complete note glyph frames: " + context);
                noteText.Append(string.Concat(c.Select(g => g.Text)));
                TestAssert.True(literal[page].Annotations.SequenceEqual(counter[page].Annotations), "All navigation stays source-equivalent: " + context);
                TestAssert.Equal(literal[page].Images.Count, counter[page].Images.Count);
                if (shape is "math-picture" or "standalone") TestAssert.Equal("ratetime",
                    NormalizeNoteMathFieldText(string.Concat(StaticNoteMathShows(counter[page], true).Select(g => g.Text))));
                PdfContentValidator.ValidatePage(counter[page], page, default);
            }
            TestAssert.Equal("ratetimeratetime", NormalizeNoteMathFieldText(noteText.ToString()));
        }
    }

    public static void DocxNoteMathFieldsStaticTablesRetainExactClippedAndUnsupportedFallbacks()
    {
        foreach (string position in new[] { "docEnd", "sectEnd" })
        foreach (bool footer in new[] { false, true })
        foreach (string variant in new[] { "exact", "clipped", "intrusive", "direction", "missing-math" })
        {
            var source = ReadNoteMathField("endnote", OoxPdfDocxMarkupMode.Final, "PAGE", "9", false);
            source = source with { Settings = source.Settings with { EndnoteReferenceSettings =
                source.Settings.EndnoteReferenceSettings with { PositionValue = position } } };
            var paragraph = variant == "missing-math" ? StaticNoteEquationParagraph(source, OoxPdfDocxMarkupMode.Final, true) : PlainNoteMathStaticParagraph();
            var table = StaticNoteTable(paragraph, "row");
            if (variant is "exact" or "clipped") table = table with { Rows = [table.Rows[0] with
                { HeightPoints = variant == "exact" ? 20d : 1d, HeightValue = variant == "exact" ? "400" : "20", HeightRuleValue = "exact" }] };
            if (variant == "intrusive") table = table with { Rows = Enumerable.Repeat(table.Rows[0], 6).ToArray() };
            if (variant == "direction") table = table with { Rows = [table.Rows[0] with
                { Cells = [table.Rows[0].Cells[0] with { TextDirectionValue = "tbRl" }] }] };
            source = WithNoteMathStaticBlocks(source, footer, [new DocxTableElement(table)]);
            var notes = CreatePageCounterLayout(source).Pages.SelectMany(page => page.PlacedRelatedStories)
                .Where(note => note.StoryLayout.Story.Id == "37").ToArray();
            TestAssert.True(notes.Length > 0 && notes.All(note => !note.StoryLayout.HasResolvedCounterInputs && !note.UsesCompleteDocumentEndnoteFlow),
                $"An unqualified static table retains its explicit fallback: {position}/{footer}/{variant}.");
        }
    }

    private static DocxTable StaticNoteTable(DocxParagraph paragraph, string shape)
    {
        var margins = new DocxTableCellMargins(0d, 0d, 0d, 0d, "0", "0", "0", "0");
        DocxTableCell Cell(DocxParagraph p) => new(string.Empty, [p], null, null, null, "top", [], margins);
        var row = new DocxTableRow([Cell(paragraph)], null);
        if (shape == "two-rows") return new("fixed", [468d], [row, new([Cell(paragraph with { })], null)]);
        if (shape == "nested")
        {
            var inner = new DocxTable("fixed", [468d], [row]);
            var tail = DocxTests.CreateDocxLayoutParagraph(string.Empty, 1d, 1d) with { LineSpacingPoints = null };
            row = new([Cell(tail) with { BodyElements = [new DocxTableElement(inner), new DocxParagraphElement(tail)] }], null);
        }
        return new("fixed", [468d], [row]);
    }
}
