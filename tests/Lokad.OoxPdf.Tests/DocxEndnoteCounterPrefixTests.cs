using System.Globalization;
using System.Reflection;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxEndnoteCounterPrefixTests
{
    public static void DocxEndnoteCounterPrefixPreservesReferenceValueAndLiteralGeometry()
    {
        foreach (string position in new[] { "docEnd", "sectEnd" })
        foreach (int reference in new[] { 1, 12 })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = Read(mode, position, reference, "PAGE", "9");
            var paragraph = source.RelatedStories.Single(story => story.Id == "37").BodyElements.OfType<DocxParagraphElement>().First().Paragraph;
            var cached = paragraph.Runs.Select(run => (run.Text, run.FieldKind)).ToArray();
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                    : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var literal = renderer.RenderPagesAsync(Read(mode, position, reference, null,
                reference.ToString(CultureInfo.InvariantCulture)), null, default).ToBlockingEnumerable().ToArray();
            var counter = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(literal.Length, counter.Length);
            string context = $"{position}/{reference}/{mode}";
            string Normalize(string value) => string.Concat(value.Normalize(NormalizationForm.FormKC).Where(c => !char.IsWhiteSpace(c)));
            var text = new StringBuilder();
            for (int page = 0; page < counter.Length; page++)
            {
                var a = DocxCommentsTests.ReadEmbeddedGlyphTextShows(literal[page], true).ToArray();
                var b = DocxCommentsTests.ReadEmbeddedGlyphTextShows(counter[page], true).ToArray();
                TestAssert.Equal(Normalize(string.Concat(a.Select(show => show.Text))), Normalize(string.Concat(b.Select(show => show.Text))));
                TestAssert.True(a.Select(show => (show.Text, show.X, show.Y, show.Size)).SequenceEqual(
                    b.Select(show => (show.Text, show.X, show.Y, show.Size))), "Reference fields preserve literal glyph frames: " + context);
                TestAssert.True(literal[page].Annotations.SequenceEqual(counter[page].Annotations), "Keep actual destination pages: " + context);
                TestAssert.Equal(literal[page].Images.Count, counter[page].Images.Count);
                PdfContentValidator.ValidatePage(counter[page], page, default);
                text.Append(string.Concat(b.Select(show => show.Text)));
            }
            TestAssert.True(Normalize(text.ToString()).Contains("Counter" + reference + ".", StringComparison.Ordinal),
                "Later overflow cannot replace the reference PAGE with the note page: " + context);
            TestAssert.True(cached.SequenceEqual(paragraph.Runs.Select(run => (run.Text, run.FieldKind))), "Binding preserves cached source text and field kinds.");
        }
    }

    public static void DocxEndnoteCounterPrefixStopsAtUnfitOrUnownedSource()
    {
        var source = Read(OoxPdfDocxMarkupMode.Final, "docEnd", 1, null, "1");
        var math = source.RelatedStories.SelectMany(story => DocxBlockTraversal.EnumerateBodyParagraphs(story))
            .SelectMany(p => p.Equations).ToDictionary(e => e, e => new DocxPreparedMathParagraph(e, new(18d, 9d, 4d, []), 12d, 16d));
        var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: math);
        var page = layout.Pages.First(p => p.PlacedRelatedStories.Any(story => story.StoryLayout.Story.Id == "37")) with { PlacedRelatedStories = [] };
        var stories = layout.RelatedStories.Where(story => story.Story.IsNormalStoryType).ToArray();
        var separator = layout.RelatedStories.Single(story => story.Story.Type == DocxRelatedStoryType.Separator);
        var count = typeof(DocxLayoutEngine).GetMethod("CountReferenceBoundEndnotes", BindingFlags.Static | BindingFlags.NonPublic)!
            .CreateDelegate<Func<DocxLayoutPage, IReadOnlyList<DocxPlacedRelatedStoryLayout>, DocxRelatedStoryLayout?, double,
                IReadOnlyList<DocxRelatedStoryLayout>, CancellationToken, double, int>>();
        TestAssert.Equal(3, count(page, [], separator, 1d, stories, default, page.MarginBottom));
        var unowned = stories.ToArray(); unowned[1] = unowned[1] with { HasResolvedCounterInputs = false };
        TestAssert.Equal(1, count(page, [], separator, 1d, unowned, default, page.MarginBottom));
        var unfit = stories.ToArray(); unfit[0] = unfit[0] with { ContentHeight = page.Height * 2d };
        TestAssert.Equal(0, count(page, [], separator, 1d, unfit, default, page.MarginBottom));
        var clipped = stories.ToArray(); var row = clipped[0].TableRows.Single();
        clipped[0] = clipped[0] with { TableRows = [row with { Cells = row.Cells.Select(cell => cell with { Height = .001d }).ToArray() }] };
        TestAssert.Equal(0, count(page, [], separator, 1d, clipped, default, page.MarginBottom));
        TestAssert.Equal(0, count(page, [], separator, 1d, stories, default, page.Height - page.MarginTop - 10d));
        TestAssert.Equal(0, count(page, [], separator, .8d, stories, default, page.MarginBottom));
        TestAssert.Equal(0, count(page, [], separator, 1d, stories, default, double.NaN));
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        TestAssert.Equal(cancellation.Token, TestAssert.Throws<OperationCanceledException>(() =>
            count(page, [], separator, 1d, stories, cancellation.Token, page.MarginBottom)).CancellationToken);
    }

    private static DocxDocument Read(OoxPdfDocxMarkupMode mode, string position, int reference, string? instruction, string value,
        int laterWords = 400, Action<XDocument, XNamespace>? change = null)
    {
        var source = DocxInlineCompositionTests.ReadMutation("CompositionProbes/mixed-note-endnote-table", mode, false, false,
            (xml, w, _) =>
            {
                var first = xml.Root!.Elements(w + "endnote").Single(note => (string?)note.Attribute(w + "id") == "37").Element(w + "p")!;
                first.Add(new XElement(w + "r", new XElement(w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), " Counter ")));
                var run = new XElement(w + "r", new XElement(w + "t", value));
                first.Add(instruction is null ? run : new XElement(w + "fldSimple", new XAttribute(w + "instr", instruction), run));
                first.Add(new XElement(w + "r", new XElement(w + "t", ".")));
                string laterId = position == "sectEnd" ? "4" : "8";
                xml.Root.Elements(w + "endnote").Single(note => (string?)note.Attribute(w + "id") == laterId).Element(w + "p")!
                    .Add(new XElement(w + "r", new XElement(w + "t", " " + string.Join(" ",
                        Enumerable.Range(0, laterWords).Select(index => $"LaterNote{index:000}")))));
                change?.Invoke(xml, w);
            }, partOverride: "word/endnotes.xml");
        var earlier = new List<DocxBodyElement>();
        for (int page = 1; page < reference; page++)
        {
            earlier.Add(new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Public earlier page", 12d, 14d)));
            earlier.Add(new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null));
        }
        foreach (var element in source.BodyElements)
        {
            if (position == "sectEnd" && element is DocxParagraphElement paragraph &&
                paragraph.Paragraph.InlineReferences.Any(item => item.Kind == DocxRelatedStoryKind.Endnote && item.Id == "4"))
                earlier.Add(new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null));
            earlier.Add(element);
        }
        return source with { Settings = source.Settings with { EndnoteReferenceSettings =
            source.Settings.EndnoteReferenceSettings with { PositionValue = position } }, BodyElements = earlier };
    }
}
