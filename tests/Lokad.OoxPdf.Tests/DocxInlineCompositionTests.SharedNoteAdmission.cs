using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxCompleteFootnoteAdmissionOwnsSharedRunMarkersOnTheirMeasuredPages()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (bool merge in new[] { false, true })
        {
            var source = ReadSharedAdmissionFixture(mode, merge);
            var layout = DocxPageFootnoteNumberingTests.CreateLayout(source);
            TestAssert.True(layout.UsesSourceOwnedFootnoteReserve, "Identically formatted run fragmentation must not change complete-note admission.");
            TestAssert.Equal(4, source.Paragraphs.Single().InlineReferences.Count);
            var notes = layout.Pages.SelectMany((page, index) => page.PlacedRelatedStories
                .Where(note => note.StoryLayout.Story.IsNormalStoryType).Select(note => (Page: index, Note: note))).ToArray();
            TestAssert.Equal(4, notes.Length);
            TestAssert.Equal(4, notes.Select(n => n.Note.StoryLayout.Story.Id).Distinct().Count());
            foreach (var (page, note) in notes)
            {
                var reference = source.Paragraphs.Single().InlineReferences.Single(r => r.Id == note.StoryLayout.Story.Id);
                TestAssert.True(layout.Pages[page].Items.OfType<DocxTextLineLayout>().SelectMany(l => l.Segments)
                    .Any(segment => ReferenceEquals(segment.StyleRun, reference.DisplayRun) && segment.Text == reference.DisplayText),
                    "A complete note shares its exact generated marker's physical page.");
            }
            TestAssert.True(layout.Pages.All(page => page.PlacedRelatedStories.Count(note =>
                note.StoryLayout.Story.Type == DocxRelatedStoryType.Separator) <= 1), "A physical page owns one shared separator.");
        }
    }

    public static void DocxCompleteFootnoteAdmissionSharedRunsPreserveRenderedPagesAndTargets()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var geometry = mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
            var variants = new[] { false, true }.Select(merge => new DocxRenderer(new TestFaceFontResolver(), mode, geometry)
                .RenderPagesAsync(ReadSharedAdmissionFixture(mode, merge), null, default).ToBlockingEnumerable().ToArray()).ToArray();
            TestAssert.Equal(variants[0].Length, variants[1].Length);
            TestAssert.Equal(4, variants[1].Sum(page => page.Annotations.Count));
            for (int page = 0; page < variants[0].Length; page++)
            {
                TestAssert.Equal(DocxPageFootnoteNumberingTests.DecodeText(variants[0][page]), DocxPageFootnoteNumberingTests.DecodeText(variants[1][page]));
                TestAssert.True(variants[0][page].Annotations.SequenceEqual(variants[1][page].Annotations),
                    "Equivalent run fragmentation preserves both source-marker pages and target geometry.");
            }
        }
    }

    public static void DocxCompleteFootnoteAdmissionSharedRunsRejectUnownedAutomaticMarkers()
    {
        foreach (string variant in new[] { "missing", "foreign", "coordinate" })
        {
            var source = ReadSharedAdmissionFixture(OoxPdfDocxMarkupMode.Final, merge: true);
            var paragraph = source.Paragraphs.Single();
            var references = paragraph.InlineReferences.ToArray();
            var reference = references[0];
            references[0] = variant switch {
                "missing" => reference with { DisplayRun = null },
                "foreign" => reference with { DisplayRun = reference.DisplayRun! with { } },
                _ => reference with { TextOffsetInRun = reference.TextOffsetInRun + 1 }
            };
            var changed = paragraph with { InlineReferences = references };
            source = source with { BodyElements = [new DocxParagraphElement(changed)], FallbackParagraphs = [changed] };
            TestAssert.True(!DocxPageFootnoteNumberingTests.CreateLayout(source).UsesSourceOwnedFootnoteReserve,
                "Missing, foreign or mismatched automatic source owners retain fallback: " + variant);
        }
    }

    private static DocxDocument ReadSharedAdmissionFixture(OoxPdfDocxMarkupMode mode, bool merge) =>
        ReadMutation("CompositionProbes/spanning-page-notes-calibri-12-zero", mode, false, false, (xml, w, _) =>
        {
            var paragraph = xml.Descendants(w + "p").Single();
            var runs = paragraph.Elements(w + "r").ToArray();
            var properties = runs[0].Element(w + "rPr")!;
            foreach (var run in runs.Where(r => r.Element(w + "footnoteReference") is not null))
            {
                run.Element(w + "rPr")?.Remove();
                run.AddFirst(new XElement(properties));
            }
            TestAssert.True(runs.All(r => XNode.DeepEquals(properties, r.Element(w + "rPr"))),
                "Every merged run must have identical authored formatting.");
            if (!merge) return;
            foreach (var run in runs.Skip(1))
            {
                runs[0].Add(run.Elements().Where(e => e.Name != w + "rPr").Select(e => new XElement(e)));
                run.Remove();
            }
        });
}
