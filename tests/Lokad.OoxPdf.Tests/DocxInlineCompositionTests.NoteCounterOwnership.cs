using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxNoteMathFieldsReferenceBindingIsIndependentOfStaticTypography()
    {
        foreach (string position in new[] { "docEnd", "sectEnd" })
        foreach (bool direct in new[] { false, true })
        foreach (int reference in new[] { 2, 12 })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            DocxDocument Source(string? field, string cached)
            {
                var source = ReadNoteMathField("endnote", mode, field, cached, direct);
                var earlier = new List<DocxBodyElement>();
                for (int page = 1; page < reference; page++)
                {
                    earlier.Add(new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Public earlier page", 12d, 14d)));
                    earlier.Add(new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null));
                }
                source = source with { Settings = source.Settings with { EndnoteReferenceSettings =
                    source.Settings.EndnoteReferenceSettings with { PositionValue = position } },
                    BodyElements = [.. earlier, .. source.BodyElements] };
                return WithNoteMathStaticStory(source, true, PlainNoteMathStaticParagraph() with { LineSpacingPoints = 14d }, 120d);
            }
            var source = Source("PAGE", "9");
            var cached = NoteMathFieldParagraph(source).Runs.Select(run => (run.Text, run.FieldKind)).ToArray();
            var renderer = PictureRangeRenderer(mode);
            var literal = renderer.RenderPagesAsync(Source(null, reference.ToString(CultureInfo.InvariantCulture)), null, default)
                .ToBlockingEnumerable().ToArray();
            var counter = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            string context = $"{position}/{direct}/{reference}/{mode}";
            TestAssert.Equal(literal.Length, counter.Length);
            var noteText = new System.Text.StringBuilder();
            for (int page = 0; page < counter.Length; page++)
            {
                string text = NormalizeNoteMathFieldText(DocxPageFootnoteNumberingTests.DecodeText(counter[page]));
                TestAssert.Equal(NormalizeNoteMathFieldText(DocxPageFootnoteNumberingTests.DecodeText(literal[page])), text);
                noteText.Append(text);
                TestAssert.True(StaticNoteMathShows(literal[page], false).SequenceEqual(StaticNoteMathShows(counter[page], false)),
                    "Reference-owned counters preserve literal fallback geometry: " + context);
                TestAssert.True(literal[page].Annotations.SequenceEqual(counter[page].Annotations),
                    "Navigation follows the unchanged measured geometry: " + context);
                TestAssert.Equal(literal[page].Images.Count, counter[page].Images.Count);
                PdfContentValidator.ValidatePage(counter[page], page, default);
            }
            TestAssert.True(noteText.ToString().Contains("Counter" + reference + ".", StringComparison.Ordinal),
                "PAGE follows the rendered reference despite unqualified static typography: " + context);
            TestAssert.True(cached.SequenceEqual(NoteMathFieldParagraph(source).Runs.Select(run => (run.Text, run.FieldKind))),
                "Placement-local binding preserves cached source data.");
        }
    }

    public static void DocxNoteMathFieldsCounterAdmissionUsesTheSelectedFooterEdge()
    {
        foreach (string position in new[] { "docEnd", "sectEnd" })
        foreach (double distance in new[] { 36d, 480d })
        {
            var source = ReadNoteCounter("endnote", OoxPdfDocxMarkupMode.Final, "PAGE", "9", false, (xml, w) =>
            {
                var note = xml.Root!.Elements(w + "endnote").Single(item => (string?)item.Attribute(w + "id") == "37");
                note.Elements(w + "p").First().Descendants(w + "t").First().Value +=
                    string.Concat(Enumerable.Repeat(" Additional note words.", 40));
            });
            source = source with { Settings = source.Settings with { EndnoteReferenceSettings =
                source.Settings.EndnoteReferenceSettings with { PositionValue = position } } };
            source = WithNoteMathStaticStory(source, true,
                DocxTests.CreateDocxLayoutParagraph("Public fixed footer", 12d, 14d), distance);
            var notes = CreatePageCounterLayout(source).Pages.SelectMany(page => page.PlacedRelatedStories)
                .Where(note => note.StoryLayout.Story.Id == "37").ToArray();
            TestAssert.True(notes.Length > 0, "The source note is actually placed.");
            TestAssert.True(notes.All(note => note.StoryLayout.HasResolvedCounterInputs == (distance == 36d)),
                "Complete counter admission uses actual footer space: " + position + "/" + distance);
            TestAssert.True(notes.All(note => !note.UsesCompleteDocumentEndnoteFlow),
                "Counter ownership does not enable the separate fixed-footer typography correction.");
        }
    }
}
