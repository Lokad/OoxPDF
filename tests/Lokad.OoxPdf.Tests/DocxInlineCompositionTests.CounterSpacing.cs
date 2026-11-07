using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxNotePageCountResultsPreserveLiteralFramesAndWrapping()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (bool narrow in new[] { false, true })
        {
            var renderer = PictureRangeRenderer(mode);
            var literal = renderer.RenderPagesAsync(ReadNoteCounter(kind, mode, null, "2", narrow), null, default)
                .ToBlockingEnumerable().ToArray();
            var counter = renderer.RenderPagesAsync(ReadNoteCounter(kind, mode, "NUMPAGES", "2", narrow), null, default)
                .ToBlockingEnumerable().ToArray();
            TestAssert.Equal(2, literal.Length);
            TestAssert.Equal(literal.Length, counter.Length);
            for (int index = 0; index < literal.Length; index++)
            {
                var first = literal[index]; var second = counter[index];
                var a = DocxCommentsTests.ReadEmbeddedGlyphTextShows(first, usePageCoordinates: true)
                    .Select(g => (g.Text, g.X, g.Y, g.Size)).ToArray();
                var b = DocxCommentsTests.ReadEmbeddedGlyphTextShows(second, usePageCoordinates: true)
                    .Select(g => (g.Text, g.X, g.Y, g.Size)).ToArray();
                TestAssert.True(a.SequenceEqual(b), "A resolved page count preserves glyph positions and wrapping: " + kind + "/" + mode + "/" + narrow);
                TestAssert.Equal(first.Images.Count, second.Images.Count);
                if (first.Images.Count != 0) TestAssert.Equal(ReadPictureRangeFrame(first), ReadPictureRangeFrame(second));
                TestAssert.True(first.Annotations.SequenceEqual(second.Annotations), "Resolved counters preserve existing note destinations and hit areas.");
            }
            TestAssert.Contains("Count 2", string.Concat(counter.Select(DocxPageFootnoteNumberingTests.DecodeText)));
        }
    }

    public static void DocxNoteCounterMeasurementUsesBindingsAndRetainsPageFallback()
    {
        foreach (int? count in new int?[] { null, 2, 12 })
        foreach (string instruction in new[] { "NUMPAGES", "PAGE" })
        {
            var source = ReadNoteCounter("endnote", OoxPdfDocxMarkupMode.Final, instruction, "2", narrow: true);
            var math = source.RelatedStories.SelectMany(story => DocxBlockTraversal.EnumerateBodyParagraphs(story))
                .SelectMany(p => p.Equations).ToDictionary(e => e, e => new DocxPreparedMathParagraph(e, new(18d, 9d, 4d, []), 12d, 16d));
            var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, knownPageCount: count)
                .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: math);
            var story = layout.RelatedStories.Single(s => s.Story.Id == "37");
            bool resolved = count is not null && instruction == "NUMPAGES";
            TestAssert.Equal(resolved, story.HasResolvedCounterInputs);
            if (resolved)
            {
                TestAssert.True(story.TextLines.All(line => !line.Text.Contains("{NUMPAGES}", StringComparison.Ordinal)),
                    "Note measurement uses the bound count rather than placeholder width.");
                TestAssert.True(story.TextLines.Count >= 3, "The narrow note actually exercises wrapping.");
                TestAssert.Contains(count!.Value.ToString(System.Globalization.CultureInfo.InvariantCulture), string.Concat(story.TextLines.Select(line => line.Text)));
            }
            else
                TestAssert.True(!layout.Pages.SelectMany(p => p.PlacedRelatedStories).Any(s => s.UsesCompleteDocumentEndnoteFlow),
                    "An unresolved page/count dependency cannot promote placeholder-sized content into complete flow.");
        }
    }

    public static void DocxCounterSpacingKeepsUnknownAndInconsistentResultFallbacks()
    {
        var source = ReadNoteCounter("endnote", OoxPdfDocxMarkupMode.Final, "NUMPAGES", "2", false);
        var paragraph = DocxBlockTraversal.EnumerateBodyParagraphs(source.RelatedStories.Single(s => s.Id == "37"))
            .Single(p => p.FieldReferences.Count != 0);
        TestAssert.True(DocxLayoutEngine.IsPlainContextualSpacingParagraph(paragraph), "The supported source/result contract admits numeric counters.");
        var field = paragraph.FieldReferences.Single();
        foreach (var broken in new[] { field with { Kind = DocxFieldKind.Other }, field with { Placeholder = "bad" },
            field with { UsesPlaceholder = false, RendersCachedResult = false } })
            TestAssert.True(!DocxLayoutEngine.IsPlainContextualSpacingParagraph(paragraph with { FieldReferences = [broken] }),
                "Unknown kinds and inconsistent result states retain their spacing fallback.");
        var unknown = ReadNoteCounter("endnote", OoxPdfDocxMarkupMode.Final, "REF MissingTarget", "2", false);
        var cached = DocxBlockTraversal.EnumerateBodyParagraphs(unknown.RelatedStories.Single(s => s.Id == "37"))
            .Single(p => p.FieldReferences.Count != 0);
        TestAssert.True(!DocxLayoutEngine.IsPlainContextualSpacingParagraph(cached), "An arbitrary cached field is not promoted to a known counter.");
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxFieldText.GetDependencies([], canceled.Token));
        TestAssert.Throws<OperationCanceledException>(() => PictureRangeRenderer(OoxPdfDocxMarkupMode.Final)
            .RenderPagesAsync(source, null, canceled.Token).ToBlockingEnumerable().ToArray());
    }

    private static DocxDocument ReadNoteCounter(string kind, OoxPdfDocxMarkupMode mode, string? instruction,
        string value, bool narrow) => ReadMutation("CompositionProbes/mixed-note-" + kind + "-table", mode, false, false,
        (xml, w, _) =>
        {
            var note = xml.Root!.Elements(w + kind).Single(n => (string?)n.Attribute(w + "id") == "37");
            var paragraph = note.Elements(w + "p").First();
            var run = new XElement(w + "r", new XElement(w + "t", value));
            paragraph.Add(new XElement(w + "r", new XElement(w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), " Count ")));
            paragraph.Add(instruction is null ? run : new XElement(w + "fldSimple", new XAttribute(w + "instr", instruction), run));
            if (narrow)
            {
                paragraph.Element(w + "pPr")!.Add(new XElement(w + "ind", new XAttribute(w + "right", "5040")));
                paragraph.Add(new XElement(w + "r", new XElement(w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"),
                    " Public trailing note words before the table.")));
            }
        }, partOverride: "word/" + kind + "s.xml");
}
