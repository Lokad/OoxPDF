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
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            int count = footer ? 3 : 2, reference = footer ? 2 : 1;
            string value = reference.ToString(System.Globalization.CultureInfo.InvariantCulture);
            var renderer = PictureRangeRenderer(mode);
            var baseline = renderer.RenderPagesAsync(ReadEndnoteFieldSource(position, direct, footer, mode, null, value),
                null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(count, baseline.Length);
            foreach (string shape in new[] { "row", "two-rows", "nested", "math-picture", "standalone" })
            {
                DocxDocument Source(string? field, string cached)
                {
                    var source = ReadEndnoteFieldSource(position, direct, footer, mode, field, cached);
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
                var literal = renderer.RenderPagesAsync(Source(null, value), null, default).ToBlockingEnumerable().ToArray();
                var diagnostics = new List<Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic>();
                var counter = renderer.RenderPagesAsync(Source("PAGE", "9"), diagnostics.Add, default).ToBlockingEnumerable().ToArray();
                string context = $"{position}/{direct}/{footer}/{shape}/{mode}";
                TestAssert.Equal(count, counter.Length); TestAssert.Equal(count, literal.Length);
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
                    AssertStaticNoteTablePaint(literal[page], footer, shape, context);
                    AssertStaticNoteTablePaint(counter[page], footer, shape, context);
                    PdfContentValidator.ValidatePage(counter[page], page, default);
                }
                TestAssert.Equal("ratetimeratetime", NormalizeNoteMathFieldText(noteText.ToString()));
                if (position == "docEnd" && !direct && !footer && mode == OoxPdfDocxMarkupMode.Final && shape == "math-picture")
                {
                    var sharedPage = counter.First(page =>
                        ReadPictureBounds(page).Any(picture => !InStaticNoteMargin(page, footer, picture.Bounds)));
                    AssertStaticNoteTablePaintRejectsOmissions(sharedPage);
                }
            }
        }
    }

    private static void AssertStaticNoteTablePaint(PdfPage page, bool footer, string shape, string context)
    {
        if (shape == "standalone") return;
        AssertLabel("Public margin story");
        if (shape == "two-rows") AssertLabel("Public margin row two");
        if (shape != "math-picture") return;

        var pictures = ReadPictureBounds(page).Where(picture => InStaticNoteMargin(page, footer, picture.Bounds)).ToArray();
        TestAssert.True(pictures.Length == 1, "The static table paints exactly one margin picture: " + context);
        var (bounds, clip) = pictures[0];
        TestAssert.True(bounds.Width > 0d && bounds.Height > 0d && Math.Abs(bounds.Width * 2d - bounds.Height * 3d) < .01d &&
            bounds.X >= clip.X - .01d && bounds.X + bounds.Width <= clip.X + clip.Width + .01d &&
            bounds.Y >= clip.Y - .01d && bounds.Y + bounds.Height <= clip.Y + clip.Height + .01d,
            "The complete margin picture fits within its paint clip: " + context);

        void AssertLabel(string label)
        {
            var glyphs = PortableRenderingTests.ReadPaintedSourceLabel(page, label);
            TestAssert.True(glyphs.All(glyph => InStaticNoteMargin(page, footer, glyph.Ink)),
                "Every source label glyph paints in the selected margin: " + context + "/" + label);
        }
    }

    private static bool InStaticNoteMargin(PdfPage page, bool footer, PdfRectangle ink) => footer
        ? ink.Y >= 0d && ink.Y + ink.Height <= 72d
        : ink.Y >= page.Height - 72d && ink.Y + ink.Height <= page.Height;

    private static void AssertStaticNoteTablePaintRejectsOmissions(PdfPage page)
    {
        var frames = DocxNoteClippingTests.ReadPhysicalPaintFrames(page).ToArray();
        var text = frames.Single(frame => DocxCommentsTests.ReadEmbeddedGlyphTextShows(
            new PdfPage(page.Width, page.Height, frame.Block, page.Fonts))
            .Any(show => show.Text.Contains("Public margin story", StringComparison.Ordinal)));
        var picture = frames.Single(frame => frame.Block.EndsWith(" Do", StringComparison.Ordinal) &&
            frame.Transform.M32 >= page.Height - 72d);

        PdfPage noText = WithoutPaint(text);
        TestAssert.Contains("Each complete source label paints once", TestAssert.Throws<InvalidOperationException>(
            () => AssertStaticNoteTablePaint(noText, false, "math-picture", "removed prose")).Message);
        PdfPage noPicture = WithoutPaint(picture);
        TestAssert.Contains("exactly one margin picture", TestAssert.Throws<InvalidOperationException>(
            () => AssertStaticNoteTablePaint(noPicture, false, "math-picture", "removed picture")).Message);
        TestAssert.True(ReadPictureBounds(noPicture).Any(draw => !InStaticNoteMargin(noPicture, false, draw.Bounds)),
            "The note's picture remains independently painted.");

        PdfPage WithoutPaint(DocxNoteClippingTests.PhysicalPaintFrame frame)
        {
            var result = page with { Content = page.Content.Remove(frame.ContentOffset, frame.Block.Length) };
            TestAssert.True(ReferenceEquals(page.Images, result.Images) &&
                page.Annotations.SequenceEqual(result.Annotations) &&
                StaticNoteMathShows(page, false).SequenceEqual(StaticNoteMathShows(result, false)) &&
                StaticNoteMathShows(page, true).SequenceEqual(StaticNoteMathShows(result, true)),
                "The mutation preserves image resources, navigation and both formula stories.");
            PdfContentValidator.ValidatePage(result, 0, default);
            return result;
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
            TestAssert.True(notes.Length > 0 && notes.All(note => note.StoryLayout.HasResolvedCounterInputs &&
                note.UsesCompleteDocumentEndnoteFlow == (variant == "intrusive" && position == "docEnd")),
                $"Complete rows may intrude; exact/clipped/unsupported rows retain fallback: {position}/{footer}/{variant}.");
        }
    }

    private static DocxTable StaticNoteTable(DocxParagraph paragraph, string shape)
    {
        var margins = new DocxTableCellMargins(0d, 0d, 0d, 0d, "0", "0", "0", "0");
        DocxTableCell Cell(DocxParagraph p) => new([p], null, null, null, "top", [], margins);
        var row = new DocxTableRow([Cell(paragraph)], null);
        if (shape == "two-rows") return new("fixed", [468d], [row, new([Cell(PlainNoteMathStaticParagraph("Public margin row two"))], null)]);
        if (shape == "nested")
        {
            var inner = new DocxTable("fixed", [468d], [row]);
            var tail = DocxTests.CreateDocxLayoutParagraph(string.Empty, 1d, 1d) with { LineSpacingPoints = null };
            row = new([Cell(tail) with { BodyElements = [new DocxTableElement(inner), new DocxParagraphElement(tail)] }], null);
        }
        return new("fixed", [468d], [row]);
    }
}
