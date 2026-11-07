using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxNoteMathFieldsCompleteStaticFramesPreserveNotePlacement()
    {
        var cases = new[]
        {
            ("docEnd", false, "plain", OoxPdfDocxMarkupMode.Final),
            ("sectEnd", true, "plain", OoxPdfDocxMarkupMode.AllMarkup),
            ("docEnd", false, "table", OoxPdfDocxMarkupMode.AllMarkup),
            ("sectEnd", true, "table", OoxPdfDocxMarkupMode.Final),
            ("docEnd", true, "math", OoxPdfDocxMarkupMode.Final),
            ("sectEnd", false, "math", OoxPdfDocxMarkupMode.AllMarkup),
            ("docEnd", false, "math-table", OoxPdfDocxMarkupMode.Final),
            ("sectEnd", true, "math-table", OoxPdfDocxMarkupMode.AllMarkup)
        };
        foreach (var (position, direct, shape, mode) in cases)
        {
            DocxDocument Source(double distance)
            {
                var source = ReadNoteMathField("endnote", mode, "PAGE", "9", direct);
                source = source with { Settings = source.Settings with { EndnoteReferenceSettings =
                    source.Settings.EndnoteReferenceSettings with { PositionValue = position } },
                    BodyElements = [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Public earlier page", 12d, 14d)),
                        new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null), .. source.BodyElements] };
                bool math = shape is "math" or "math-table";
                var paragraph = math ? StaticNoteEquationParagraph(source, mode, true) : PlainNoteMathStaticParagraph();
                if (math) source = WithStaticNoteEquation(source, true, paragraph);
                DocxBodyElement block = shape is "table" or "math-table"
                    ? new DocxTableElement(StaticNoteTable(paragraph, "row")) : new DocxParagraphElement(paragraph);
                return WithNoteMathStaticBlocks(source, true, [block], distance);
            }
            var renderer = PictureRangeRenderer(mode);
            var baseline = renderer.RenderPagesAsync(Source(30d), null, default).ToBlockingEnumerable().ToArray();
            var intrusive = renderer.RenderPagesAsync(Source(120d), null, default).ToBlockingEnumerable().ToArray();
            string context = $"{position}/{direct}/{shape}/{mode}";
            TestAssert.Equal(3, baseline.Length); TestAssert.Equal(baseline.Length, intrusive.Length);
            var noteText = new System.Text.StringBuilder();
            for (int page = 0; page < baseline.Length; page++)
            {
                // This fixture's notes occupy the upper half; the independently
                // moved footer stays below it. Do not classify intrusion as body math.
                var a = NoteGlyphs(baseline[page]); var e = NoteGlyphs(intrusive[page]);
                TestAssert.True(a.SequenceEqual(e), "Complete footer intrusion cannot move fitting note glyphs: " + context);
                noteText.Append(string.Concat(e.Select(glyph => glyph.Text)));
                TestAssert.True(baseline[page].Annotations.Where(link => link.Y > baseline[page].Height / 2d)
                    .SequenceEqual(intrusive[page].Annotations.Where(link => link.Y > intrusive[page].Height / 2d)),
                    "Note navigation follows the unchanged paragraph frames: " + context);
                TestAssert.Equal(baseline[page].Images.Count, intrusive[page].Images.Count);
                PdfContentValidator.ValidatePage(intrusive[page], page, default);
            }
            TestAssert.Equal("ratetimeratetime", NormalizeNoteMathFieldText(noteText.ToString()));
            TestAssert.True(NormalizeNoteMathFieldText(string.Concat(intrusive.Select(DocxPageFootnoteNumberingTests.DecodeText)))
                .Contains("Counter2.", StringComparison.Ordinal),
                "The actual rendered reference still owns PAGE: " + context);
        }

        static (string Text, double X, double Y, double Size)[] NoteGlyphs(PdfPage page) =>
            DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true,
                fontResourceFilter: name => name.StartsWith("Eq", StringComparison.Ordinal))
                .Where(glyph => glyph.Y > page.Height / 2d)
                .Select(glyph => (glyph.Text, glyph.X, glyph.Y, glyph.Size)).ToArray();
    }

    public static void DocxNoteMathFieldsCompleteStaticFramesRequirePageAndNoteSpace()
    {
        foreach (double distance in new[] { 36d, 480d, -2d })
        {
            var source = ReadNoteCounter("endnote", OoxPdfDocxMarkupMode.Final, "PAGE", "9", false, (xml, w) =>
            {
                var note = xml.Root!.Elements(w + "endnote").Single(item => (string?)item.Attribute(w + "id") == "37");
                note.Elements(w + "p").First().Descendants(w + "t").First().Value +=
                    string.Concat(Enumerable.Repeat(" Additional note words.", 40));
            });
            source = WithNoteMathStaticStory(source, true, PlainNoteMathStaticParagraph(), distance);
            var notes = CreatePageCounterLayout(source).Pages.SelectMany(page => page.PlacedRelatedStories)
                .Where(note => note.StoryLayout.Story.Id == "37").ToArray();
            TestAssert.True(notes.Length > 0, "The source note is actually placed.");
            TestAssert.True(notes.All(note => note.UsesCompleteDocumentEndnoteFlow == (distance == 36d)),
                "Complete boundary placement requires a page-contained frame and actual free note space: " + distance);
        }
    }
}
