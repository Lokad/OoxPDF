using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxNoteMathFieldsPreparedStaticEquationsPreserveBothStories()
    {
        foreach (string position in new[] { "docEnd", "sectEnd" })
        foreach (bool direct in new[] { false, true })
        foreach (bool footer in new[] { false, true })
        foreach (bool standalone in new[] { false, true })
        foreach (string instruction in new[] { "PAGE", "NUMPAGES" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            int count = footer ? 3 : 2, reference = footer ? 2 : 1;
            string value = (instruction == "PAGE" ? reference : count).ToString(System.Globalization.CultureInfo.InvariantCulture);
            DocxDocument Source(string? field, string cached, bool marginMath)
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
                if (!marginMath) return source;
                var paragraph = StaticNoteEquationParagraph(source, mode, standalone);
                if (footer && !standalone)
                {
                    paragraph = NoteMathStaticPictureParagraph(source, paragraph, false);
                    var equation = paragraph.Equations[0] with { SourceRunIndex = paragraph.Runs.Max(run => run.SourceRunIndex) + 1 };
                    paragraph = paragraph with { Equations = [equation], InlineContent =
                        [.. paragraph.InlineContent!, new DocxInlineEquation(equation)] };
                }
                return WithStaticNoteEquation(source, footer, paragraph);
            }
            var renderer = PictureRangeRenderer(mode);
            var baseline = renderer.RenderPagesAsync(Source(null, value, false), null, default).ToBlockingEnumerable().ToArray();
            var literal = renderer.RenderPagesAsync(Source(null, value, true), null, default).ToBlockingEnumerable().ToArray();
            var source = Source(instruction, "9", true);
            var runs = NoteMathFieldParagraph(source).Runs.Select(run => (run.Text, run.FieldKind)).ToArray();
            var diagnostics = new List<Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic>();
            var counter = renderer.RenderPagesAsync(source, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            string context = $"{position}/direct={direct}/footer={footer}/standalone={standalone}/{instruction}/{mode}";
            TestAssert.Equal(count, counter.Length); TestAssert.Equal(count, literal.Length); TestAssert.Equal(count, baseline.Length);
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Supported static and note formulas coexist: " + context);
            TestAssert.True(runs.SequenceEqual(NoteMathFieldParagraph(source).Runs.Select(run => (run.Text, run.FieldKind))),
                "Binding preserves the cached note source.");
            var noteText = new System.Text.StringBuilder();
            for (int page = 0; page < count; page++)
            {
                TestAssert.Equal(NormalizeNoteMathFieldText(DocxPageFootnoteNumberingTests.DecodeText(literal[page])),
                    NormalizeNoteMathFieldText(DocxPageFootnoteNumberingTests.DecodeText(counter[page])));
                var a = StaticNoteMathShows(baseline[page], false); var b = StaticNoteMathShows(literal[page], false);
                var c = StaticNoteMathShows(counter[page], false);
                TestAssert.True(a.SequenceEqual(b) && a.SequenceEqual(c), "Static math preserves each note glyph frame: " + context);
                noteText.Append(string.Concat(c.Select(glyph => glyph.Text)));
                var margin = StaticNoteMathShows(counter[page], true);
                TestAssert.Equal("ratetime", NormalizeNoteMathFieldText(string.Concat(margin.Select(glyph => glyph.Text))));
                TestAssert.True(margin.SequenceEqual(StaticNoteMathShows(literal[page], true)), "Literal and field sources keep every margin math frame.");
                TestAssert.True(literal[page].Annotations.SequenceEqual(counter[page].Annotations), "Source-equivalent counters preserve all navigation.");
                TestAssert.Equal(literal[page].Images.Count, counter[page].Images.Count);
                PdfContentValidator.ValidatePage(counter[page], page, default);
            }
            TestAssert.Equal("ratetimeratetime", NormalizeNoteMathFieldText(noteText.ToString()));
        }
    }

    public static void DocxNoteMathFieldsStaticEquationFallbackRequiresAvailableFittingFrames()
    {
        foreach (string position in new[] { "docEnd", "sectEnd" })
        foreach (bool footer in new[] { false, true })
        foreach (bool standalone in new[] { false, true })
        foreach (string variant in new[] { "missing", "fixed", "tall", "wide" })
        {
            var source = ReadNoteMathField("endnote", OoxPdfDocxMarkupMode.Final, "PAGE", "9", false);
            source = source with { Settings = source.Settings with { EndnoteReferenceSettings =
                source.Settings.EndnoteReferenceSettings with { PositionValue = position } } };
            var paragraph = StaticNoteEquationParagraph(source, OoxPdfDocxMarkupMode.Final, standalone);
            if (variant == "fixed") paragraph = paragraph with { LineSpacingPoints = 14d };
            source = WithStaticNoteEquation(source, footer, paragraph);
            var math = source.RelatedStories.SelectMany(DocxBlockTraversal.EnumerateBodyParagraphs).SelectMany(p => p.Equations)
                .ToDictionary(e => e, e => new DocxPreparedMathParagraph(e, new(18d, 9d, 4d, []), 12d, 16d));
            if (variant != "missing")
            {
                var equation = paragraph.Equations[0];
                math[equation] = variant == "tall" ? new(equation, new(18d, 50d, 40d, []), 50d, 90d)
                    : new(equation, new(variant == "wide" ? 700d : 18d, 9d, 4d, []), 12d, 16d);
            }
            var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: math);
            var notes = layout.Pages.SelectMany(page => page.PlacedRelatedStories).Where(note => note.StoryLayout.Story.Id == "37").ToArray();
            TestAssert.True(notes.Length > 0 && notes.All(note => note.StoryLayout.HasResolvedCounterInputs && !note.UsesCompleteDocumentEndnoteFlow),
                $"Static-equation typography keeps fallback while note counters use their rendered reference: {position}/{footer}/{standalone}/{variant}.");
        }
    }

    private static (string Text, double X, double Y, double Size)[] StaticNoteMathShows(PdfPage page, bool margin) =>
        DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true,
            fontResourceFilter: name => name.StartsWith("Eq", StringComparison.Ordinal))
            .Where(glyph => (glyph.Y < 72d || glyph.Y > page.Height - 72d) == margin)
            .Select(glyph => (glyph.Text, glyph.X, glyph.Y, glyph.Size)).ToArray();

    private static DocxParagraph StaticNoteEquationParagraph(DocxDocument source, OoxPdfDocxMarkupMode mode, bool standalone)
    {
        var paragraph = standalone ? PlainNoteMathStaticParagraph() with { Runs = [] } : ReadNoteMathStaticCounters(mode, false);
        var equation = NoteMathFieldParagraph(source).Equations[0] with { TextRunIndex = paragraph.Runs.Count,
            SourceRunIndex = paragraph.Runs.Select(run => run.SourceRunIndex).DefaultIfEmpty(-1).Max() + 1,
            Display = standalone, ParentFontSize = 12d };
        var items = DocxInlineContent.FromPrefix(paragraph.Runs, [], default);
        items.Add(new DocxInlineEquation(equation));
        return paragraph with { Equations = [equation], InlineContent = items };
    }

    private static DocxDocument WithStaticNoteEquation(DocxDocument source, bool footer, DocxParagraph paragraph) =>
        WithNoteMathStaticStory(source, footer, paragraph) with { EquationParts = [.. source.EquationParts,
            new DocxEquationPartSource(footer ? "/word/public-math-footer.xml" : "/word/public-math-header.xml", 1,
                [new DocxParagraphElement(paragraph)])] };
}
