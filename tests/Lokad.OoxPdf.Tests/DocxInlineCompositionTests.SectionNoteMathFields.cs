using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxNoteMathFieldsSingleSectionSharesPlacementAcrossEndnotePositions()
    {
        foreach (bool direct in new[] { false, true })
        foreach (string? instruction in new string?[] { null, "PAGE" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = ReadNoteMathField("endnote", mode, instruction, instruction is null ? "1" : "9", direct);
            source = source with { BodyElements = source.BodyElements.Where(element => element is not DocxSectionBreakElement).ToArray() };
            DocxDocument Position(string position) => source with { Settings = source.Settings with { EndnoteReferenceSettings =
                source.Settings.EndnoteReferenceSettings with { PositionValue = position } } };
            var renderer = PictureRangeRenderer(mode);
            var document = renderer.RenderPagesAsync(Position("docEnd"), null, default).ToBlockingEnumerable().Single();
            var section = renderer.RenderPagesAsync(Position("sectEnd"), null, default).ToBlockingEnumerable().Single();
            var a = DocxCommentsTests.ReadEmbeddedGlyphTextShows(document, true).Select(g => (g.Text, g.X, g.Y, g.Size));
            var b = DocxCommentsTests.ReadEmbeddedGlyphTextShows(section, true).Select(g => (g.Text, g.X, g.Y, g.Size));
            TestAssert.True(a.SequenceEqual(b), "A complete group has the same paragraph/separator geometry when the only section ends the document.");
            TestAssert.True(document.Annotations.SequenceEqual(section.Annotations), "Navigation uses the same completed source boundaries.");
            TestAssert.Equal(document.Images.Count, section.Images.Count);
        }
    }

    public static void DocxNoteMathFieldsSectionPlacementKeepsMixedFramesAndFollowingBody()
    {
        foreach (bool direct in new[] { false, true })
        foreach (bool beforeMath in new[] { false, true })
        foreach (bool shiftReference in new[] { false, true })
        foreach (string instruction in new[] { "PAGE", "NUMPAGES" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            int pageCount = shiftReference ? 3 : 2, referencePage = shiftReference ? 2 : 1;
            string value = (instruction == "PAGE" ? referencePage : pageCount).ToString(System.Globalization.CultureInfo.InvariantCulture);
            DocxDocument Source(string? field, string cached)
            {
                var source = ReadNoteMathField("endnote", mode, field, cached, direct, beforeMath: beforeMath);
                return source with
                {
                    Settings = source.Settings with { EndnoteReferenceSettings =
                        source.Settings.EndnoteReferenceSettings with { PositionValue = "sectEnd" } },
                    BodyElements = shiftReference
                        ? [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Public earlier page", 12d, 14d)),
                            new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null), .. source.BodyElements]
                        : source.BodyElements
                };
            }
            var renderer = PictureRangeRenderer(mode);
            var literal = renderer.RenderPagesAsync(Source(null, value), null, default).ToBlockingEnumerable().ToArray();
            var source = Source(instruction, "9");
            var paragraph = NoteMathFieldParagraph(source);
            var sourceRuns = paragraph.Runs.Select(run => (run.Text, run.FieldKind)).ToArray();
            var diagnostics = new List<Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic>();
            var counter = renderer.RenderPagesAsync(source, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            string context = $"direct={direct}/before={beforeMath}/shift={shiftReference}/{instruction}/{mode}";
            TestAssert.Equal(pageCount, literal.Length); TestAssert.Equal(literal.Length, counter.Length);
            TestAssert.True(sourceRuns.SequenceEqual(paragraph.Runs.Select(run => (run.Text, run.FieldKind))),
                "Page binding keeps the source counter and cached result intact: " + context);
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Section placement cannot discard a supported mixed note: " + context);
            for (int page = 0; page < counter.Length; page++)
            {
                TestAssert.Equal(NormalizeNoteMathFieldText(DocxPageFootnoteNumberingTests.DecodeText(literal[page])),
                    NormalizeNoteMathFieldText(DocxPageFootnoteNumberingTests.DecodeText(counter[page])));
                var a = DocxCommentsTests.ReadEmbeddedGlyphTextShows(literal[page], true,
                    fontResourceFilter: name => name.StartsWith("Eq", StringComparison.Ordinal)).ToArray();
                var b = DocxCommentsTests.ReadEmbeddedGlyphTextShows(counter[page], true,
                    fontResourceFilter: name => name.StartsWith("Eq", StringComparison.Ordinal)).ToArray();
                TestAssert.True(a.Select(g => (g.Text, g.X, g.Y, g.Size)).SequenceEqual(b.Select(g => (g.Text, g.X, g.Y, g.Size))),
                    "Equivalent counters preserve each placed formula frame: " + context);
                TestAssert.Equal(literal[page].Images.Count, counter[page].Images.Count);
                TestAssert.True(literal[page].Annotations.SequenceEqual(counter[page].Annotations),
                    "Picture and navigation ownership survives section placement: " + context);
                PdfContentValidator.ValidatePage(counter[page], page, default);
            }
            string noteText = NormalizeNoteMathFieldText(DocxPageFootnoteNumberingTests.DecodeText(counter[referencePage - 1]));
            TestAssert.Contains("Counter" + value + ".", noteText);
            var math = DocxCommentsTests.ReadEmbeddedGlyphTextShows(counter[referencePage - 1], true,
                fontResourceFilter: name => name.StartsWith("Eq", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal("ratetimeratetime", NormalizeNoteMathFieldText(string.Concat(math.Select(g => g.Text))));
            TestAssert.Contains("Delta", DocxPageFootnoteNumberingTests.DecodeText(counter[^1]));
            TestAssert.True(!DocxPageFootnoteNumberingTests.DecodeText(counter[^1]).Contains("Counter", StringComparison.Ordinal),
                "The complete mixed note precedes the following source section: " + context);
        }
    }
}
