using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxContinuedPlainEndnoteCountersPreserveLiteralPagesAndGeometry()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = PictureRangeRenderer(mode);
            var literal = renderer.RenderPagesAsync(ReadContinuedCounter(null), null, default).ToBlockingEnumerable().ToArray();
            var counter = renderer.RenderPagesAsync(ReadContinuedCounter("PAGE"), null, default).ToBlockingEnumerable().ToArray();
            TestAssert.True(literal.Length >= 3, "The note actually continues across pages.");
            TestAssert.Equal(literal.Length, counter.Length);
            TestAssert.Contains("Count 1", DocxPageFootnoteNumberingTests.DecodeText(counter[^1]));
            for (int index = 0; index < literal.Length; index++)
            {
                var a = DocxCommentsTests.ReadEmbeddedGlyphTextShows(literal[index], usePageCoordinates: true)
                    .Select(g => (g.Text, g.X, g.Y, g.Size)).ToArray();
                var b = DocxCommentsTests.ReadEmbeddedGlyphTextShows(counter[index], usePageCoordinates: true)
                    .Select(g => (g.Text, g.X, g.Y, g.Size)).ToArray();
                TestAssert.True(a.SequenceEqual(b), "Reference-owned PAGE preserves literal continuation glyphs and page allocation: " + mode);
                TestAssert.True(literal[index].Annotations.SequenceEqual(counter[index].Annotations),
                    "Counter binding preserves the existing note destinations and hit areas.");
            }

            DocxDocument ReadContinuedCounter(string? instruction) => ReadNoteCounter("endnote", mode,
                instruction, instruction is null ? "1" : "7", false, (xml, w) =>
                {
                    var note = xml.Root!.Elements(w + "endnote").Single(n => (string?)n.Attribute(w + "id") == "37");
                    var first = note.Elements(w + "p").First();
                    foreach (var element in note.Elements().Where(e => !ReferenceEquals(e, first)).ToArray()) element.Remove();
                    first.Descendants(w + "t").First().Value = " Public continued note " +
                        string.Join(" ", Enumerable.Range(0, 900).Select(n => "item" + n.ToString("D3", System.Globalization.CultureInfo.InvariantCulture)));
                });
        }
    }

    public static void DocxFittingEndnotePageFieldsPreserveLiteralFramesAndWrapping()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (bool narrow in new[] { false, true })
        {
            var renderer = PictureRangeRenderer(mode);
            var literal = renderer.RenderPagesAsync(ReadNoteCounter("endnote", mode, null, "1", narrow), null, default)
                .ToBlockingEnumerable().ToArray();
            var counter = renderer.RenderPagesAsync(ReadNoteCounter("endnote", mode, "PAGE", "7", narrow), null, default)
                .ToBlockingEnumerable().ToArray();
            TestAssert.Equal(2, counter.Length);
            TestAssert.Equal(literal.Length, counter.Length);
            for (int index = 0; index < literal.Length; index++)
            {
                var first = literal[index]; var second = counter[index];
                var a = DocxCommentsTests.ReadEmbeddedGlyphTextShows(first, usePageCoordinates: true)
                    .Select(g => (g.Text, g.X, g.Y, g.Size)).ToArray();
                var b = DocxCommentsTests.ReadEmbeddedGlyphTextShows(second, usePageCoordinates: true)
                    .Select(g => (g.Text, g.X, g.Y, g.Size)).ToArray();
                TestAssert.True(a.SequenceEqual(b), "Fitting PAGE results preserve measured glyphs and wrapping: " + mode + "/" + narrow);
                TestAssert.Equal(first.Images.Count, second.Images.Count);
                if (first.Images.Count != 0) TestAssert.Equal(ReadPictureRangeFrame(first), ReadPictureRangeFrame(second));
                TestAssert.True(first.Annotations.SequenceEqual(second.Annotations), "Placement-bound fields preserve note targets and hit areas.");
            }
            TestAssert.Contains("Count 1", string.Concat(counter.Select(DocxPageFootnoteNumberingTests.DecodeText)));
        }
    }

    public static void DocxFittingEndnotePageMeasurementUsesTwoDigitReferencePage()
    {
        var source = ReadNoteCounter("endnote", OoxPdfDocxMarkupMode.Final, "PAGE", "1", narrow: true);
        source = source with { BodyElements = Enumerable.Range(0, 11).SelectMany(_ =>
            new DocxBodyElement[] { new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Public earlier page", 12d, 14d)),
                new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null) }).Concat(source.BodyElements).ToArray() };
        var layout = CreatePageCounterLayout(source);
        TestAssert.Equal(13, layout.Pages.Count);
        var sourceLayout = layout.RelatedStories.Single(s => s.Story.Id == "37");
        TestAssert.True(!sourceLayout.HasResolvedCounterInputs, "Placement does not publish a page-bound result into the width-only memo.");
        var placed = layout.Pages.SelectMany(p => p.PlacedRelatedStories).Single(s => s.StoryLayout.Story.Id == "37");
        TestAssert.True(layout.Pages[12].PlacedRelatedStories.Contains(placed), "The complete note is on physical page thirteen, after reference page twelve.");
        TestAssert.True(placed.UsesCompleteDocumentEndnoteFlow && placed.StoryLayout.HasResolvedCounterInputs,
            "Only the complete fitting placement owns its rendered reference binding.");
        TestAssert.True(placed.StoryLayout.TextLines.Count >= 3, "The narrow note exercises measured wrapping.");
        TestAssert.Contains("12", string.Concat(placed.StoryLayout.TextLines.Select(line => line.Text)));
        TestAssert.True(placed.StoryLayout.TextLines.All(line => !line.Text.Contains("{PAGE}", StringComparison.Ordinal)),
            "No placeholder-sized line survives into admitted geometry.");
    }

    public static void DocxEndnotePageBindingRetainsUnfitAndUnsupportedOwners()
    {
        foreach (string variant in new[] { "keep", "overflow", "static", "scaled", "clipped" })
        {
            var source = ReadNoteCounter("endnote", OoxPdfDocxMarkupMode.Final, "PAGE", "2", false, (xml, w) =>
            {
                var note = xml.Root!.Elements(w + "endnote").Single(n => (string?)n.Attribute(w + "id") == "37");
                var paragraph = note.Elements(w + "p").First();
                if (variant == "keep") paragraph.Element(w + "pPr")!.Add(new XElement(w + "keepLines"));
                if (variant == "overflow") paragraph.Descendants(w + "t").First().Value += string.Concat(Enumerable.Repeat(" Additional note words.", 200));
                if (variant == "clipped") note.Element(w + "tbl")!.Element(w + "tr")!.AddFirst(new XElement(w + "trPr",
                    new XElement(w + "trHeight", new XAttribute(w + "val", "20"), new XAttribute(w + "hRule", "exact"))));
            });
            if (variant == "static") source = source with { HeaderParagraphs = [DocxTests.CreateDocxLayoutParagraph("Public header", 12d, 14d)] };
            var layout = CreatePageCounterLayout(source, scaled: variant == "scaled");
            var notes = layout.Pages.SelectMany(p => p.PlacedRelatedStories).Where(s => s.StoryLayout.Story.Id == "37").ToArray();
            TestAssert.True(notes.Length != 0 && notes.All(s => !s.StoryLayout.HasResolvedCounterInputs && !s.UsesCompleteDocumentEndnoteFlow),
                "Unfit/unsupported ownership retains its unresolved placement fallback: " + variant);
        }
    }

    private static DocxLayout CreatePageCounterLayout(DocxDocument source, bool scaled = false)
    {
        var math = source.RelatedStories.SelectMany(story => DocxBlockTraversal.EnumerateBodyParagraphs(story))
            .SelectMany(p => p.Equations).ToDictionary(e => e, e => new DocxPreparedMathParagraph(e, new(18d, 9d, 4d, []), 12d, 16d));
        return new DocxLayoutEngine(scaled ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
            : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, wordCompatiblePrintScale: .8d)
            .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: math);
    }
}
