using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxFootnotesTests
{
    public static void DocxDisabledSectionHeadersRetainSourcePagesAndVisibleOrigins() =>
        CheckDisabledSectionTemplates("header");

    public static void DocxDisabledSectionFootersRetainSourcePagesAndVisibleOrigins() =>
        CheckDisabledSectionTemplates("footer");

    private static void CheckDisabledSectionTemplates(string story)
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach ((string? preceding, string? following, int expected) in new[]
        {
            ((string?)"nextPage", (string?)"continuous", 1),
            ((string?)"continuous", (string?)"nextPage", 2),
            ((string?)null, (string?)"continuous", 1),
            ((string?)"continuous", (string?)null, 2)
        })
        {
            var renderer = new DocxRenderer(new TestFaceFontResolver(), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                    : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var control = renderer.RenderPagesAsync(ReadSectionTemplate(preceding, following, mode), null, default)
                .ToBlockingEnumerable().ToArray();
            TestAssert.Equal(expected, control.Length);
            foreach (string variant in new[] { "first", "even" })
            foreach (string owner in new[] { "earlier", "final", "both" })
            {
                var source = ReadSectionTemplate(preceding, following, mode, story, variant, owner);
                TestAssert.True(DocxPlainBodyFrame.HasNoSelectableStaticStories(source),
                    "Only known disabled templates participate in this invariant.");
                var pages = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
                TestAssert.True(expected == pages.Length,
                    $"Disabled {story}/{variant}/{owner} must preserve the following section's page assignment.");
                for (int page = 0; page < pages.Length; page++)
                {
                    var baseline = PhysicalSectionGlyphs(control[page]);
                    var actual = PhysicalSectionGlyphs(pages[page]);
                    TestAssert.Equal(baseline.Length, actual.Length);
                    for (int glyph = 0; glyph < actual.Length; glyph++) TestAssert.Equal(baseline[glyph], actual[glyph]);
                    TestAssert.Equal(control[page].Annotations.Count, pages[page].Annotations.Count);
                    TestAssert.True(!string.Concat(actual.Select(x => x.Text)).Contains("STATIC TEMPLATE", StringComparison.Ordinal),
                        "Disabled story text must remain absent from physical glyph emission.");
                }
            }
        }
    }

    public static void DocxSectionTemplatesKeepActiveUnknownInheritedAndCancelledFallbacks()
    {
        foreach (string story in new[] { "header", "footer" })
        foreach (string variant in new[] { "default", "first", "even", "unresolved" })
        {
            var source = ReadSectionTemplate("nextPage", "continuous", OoxPdfDocxMarkupMode.Final,
                story, variant, "earlier", enabled: true);
            TestAssert.True(!DocxPlainBodyFrame.HasNoSelectableStaticStories(source),
                "Active, unknown and later-inherited variants retain the existing static frame fallback.");
        }
        var disabled = ReadSectionTemplate("nextPage", "continuous", OoxPdfDocxMarkupMode.Final,
            "header", "first", "earlier");
        TestAssert.True(DocxPlainBodyFrame.HasNoSelectableStaticStories(disabled), "The inactive source is understood.");
        var unowned = disabled with { HeaderParagraphs = [disabled.Paragraphs[0]] };
        TestAssert.True(!DocxPlainBodyFrame.HasNoSelectableStaticStories(unowned), "Unbound legacy roots remain conservative.");
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxPlainBodyFrame.HasNoSelectableStaticStories(disabled, cancelled.Token));
        TestAssert.Throws<OperationCanceledException>(() => new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(disabled, new DocxTests.FamilyWidthTextMeasurer(), cancelled.Token));
    }

    private static (string Text, double Size, double X, double Y)[] PhysicalSectionGlyphs(PdfPage page) =>
        DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true, includeTextTransforms: true, honorActualText: false)
            .Select(show => (show.Text, show.Size, show.X, show.Y)).ToArray();

    private static DocxDocument ReadSectionTemplate(string? preceding, string? following, OoxPdfDocxMarkupMode mode,
        string? story = null, string variant = "first", string owner = "both", bool enabled = false) =>
        ReadSectionNumberingFixture("footnote", (xml, settings, parts) =>
        {
            SetSectionTypes(xml, preceding, following);
            SetNotePresence(xml, parts, "footnote", false, false);
            settings.Root!.Elements(NoteWord + "evenAndOddHeaders").Remove();
            var sections = xml.Descendants(NoteWord + "sectPr").ToArray();
            foreach (var section in sections) section.Elements(NoteWord + "titlePg").Remove();
            if (story is null) return;
            XNamespace relationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            XNamespace packageRelationship = "http://schemas.openxmlformats.org/package/2006/relationships";
            string part = "section-template-" + story + ".xml";
            parts["word/" + part] = new XElement(NoteWord + (story == "header" ? "hdr" : "ftr"),
                new XElement(NoteWord + "p", new XElement(NoteWord + "r", new XElement(NoteWord + "t", "STATIC TEMPLATE")))).ToString();
            var relationships = XDocument.Parse(parts["word/_rels/document.xml.rels"]);
            relationships.Root!.Add(new XElement(packageRelationship + "Relationship", new XAttribute("Id", "rSectionTemplate"),
                new XAttribute("Type", relationship.NamespaceName + "/" + story), new XAttribute("Target", part)));
            parts["word/_rels/document.xml.rels"] = relationships.ToString();
            var types = XDocument.Parse(parts["[Content_Types].xml"]);
            types.Root!.Add(new XElement(types.Root.Name.Namespace + "Override", new XAttribute("PartName", "/word/" + part),
                new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml." + story + "+xml")));
            parts["[Content_Types].xml"] = types.ToString();
            for (int index = 0; index < sections.Length; index++)
            {
                if (owner == "both" || owner == "earlier" && index == 0 || owner == "final" && index == sections.Length - 1)
                    sections[index].AddFirst(new XElement(NoteWord + story + "Reference", new XAttribute(NoteWord + "type", variant),
                        new XAttribute(relationship + "id", "rSectionTemplate")));
            }
            // A first-page story defined only in the earlier section becomes
            // selectable through inheritance when the final section enables it.
            if (enabled && variant == "first") sections[^1].Add(new XElement(NoteWord + "titlePg"));
            if (enabled && variant == "even") settings.Root.Add(new XElement(NoteWord + "evenAndOddHeaders"));
        }, mode);
}
