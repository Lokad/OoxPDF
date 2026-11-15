using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxCustomMarkEmptyTextPreservesFractionsPaintAndNoteLinks()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (int count in new[] { 1, 2 })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode, CachedMathGeometry(mode));
            var control = renderer.RenderPagesAsync(ReadCustomMarkText(kind, count, mode, "original"), null, default)
                .ToBlockingEnumerable().ToArray();
            foreach (string variant in new[] { "empty", "several-empty", "trimmed-empty", "trimmed-mark", "page-hint", "page-hint-namespace", "page-hints-and-empty" })
            {
                var source = ReadCustomMarkText(kind, count, mode, variant);
                var paragraph = source.Paragraphs.Single(p => p.Equations.Count != 0);
                var reference = paragraph.InlineReferences.Single();
                TestAssert.Equal("*", reference.CustomMarkText!);
                TestAssert.True(DocxCustomNoteMarkerOwners.TryGetSourceRun(reference,
                    DocxCustomNoteMarkerOwners.IndexSourceRuns(paragraph.Runs, default), out var owner) && owner.Text.StartsWith('*'),
                    "Nonprinting XML text or saved pagination metadata cannot replace the mark's exact source owner.");
                TestAssert.Equal(count, DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
                var diagnostics = new List<OoxPdfDiagnostic>();
                var actual = renderer.RenderPagesAsync(source, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
                TestAssert.Equal(control.Length, actual.Length);
                for (int page = 0; page < actual.Length; page++)
                {
                    TestAssert.Equal(control[page].Content, actual[page].Content);
                    var a = CustomMarkPhysicalGlyphs(actual[page]); var b = CustomMarkPhysicalGlyphs(control[page]);
                    TestAssert.Equal(b.Length, a.Length);
                    for (int glyph = 0; glyph < a.Length; glyph++) TestAssert.Equal(b[glyph], a[glyph]);
                    var links = actual[page].Annotations; var expected = control[page].Annotations;
                    TestAssert.Equal(expected.Count, links.Count);
                    for (int link = 0; link < links.Count; link++)
                        TestAssert.Equal((expected[link].X, expected[link].Y, expected[link].Width, expected[link].Height, expected[link].Destination),
                            (links[link].X, links[link].Y, links[link].Width, links[link].Height, links[link].Destination));
                }
                string text = string.Concat(actual.SelectMany(CustomMarkPhysicalGlyphs).Select(g => g.Text));
                TestAssert.Equal(count, text.Split("ratetime", StringSplitOptions.None).Length - 1);
                TestAssert.Equal(4, actual.Sum(page => page.Annotations.Count(link => link.Destination is not null)));
                TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                    "Nonprinting source metadata must not remove an independently supported fraction.");
            }
        }
    }

    public static void DocxCustomMarkTextRetainsVisibleWhitespaceAndOwnerBoundaries()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (string variant in new[] { "preserved-space", "tab", "break", "page-break", "reference", "unknown", "nested-empty", "surrogate", "end-of-run", "other-run", "page-hint-attribute", "page-hint-text", "page-hint-child" })
        {
            var source = ReadCustomMarkText(kind, 1, OoxPdfDocxMarkupMode.Final, variant);
            var paragraph = source.Paragraphs.Single(p => p.Equations.Count != 0);
            var reference = paragraph.InlineReferences.First(r => r.CustomMarkFollowsValue == "1");
            TestAssert.True(reference.CustomMarkText is null,
                "A mark cannot cross visible whitespace, controls, another owner or an unrecognized payload: " + variant);
            TestAssert.True(!DocxCustomNoteMarkerOwners.TryGetSourceRun(reference,
                DocxCustomNoteMarkerOwners.IndexSourceRuns(paragraph.Runs, default), out _), "Unknown ownership retains its fallback.");
        }
    }

    public static void DocxCustomMarkCasingKeepsResolvedFractionsAndFirstGlyphLinks()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (int count in new[] { 1, 2 })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode, CachedMathGeometry(mode));
            var control = renderer.RenderPagesAsync(ReadCasedCustomMark(kind, count, mode, "A", empty: false), null, default)
                .ToBlockingEnumerable().ToArray();
            foreach (bool empty in new[] { false, true })
            {
                var source = ReadCasedCustomMark(kind, count, mode, "a", empty);
                var reference = source.Paragraphs.SelectMany(p => p.InlineReferences).Single(r => r.CustomMarkFollowsValue == "1");
                TestAssert.Equal("A", reference.CustomMarkText!);
                var diagnostics = new List<OoxPdfDiagnostic>();
                var actual = renderer.RenderPagesAsync(source, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
                TestAssert.Equal(control.Length, actual.Length);
                for (int page = 0; page < actual.Length; page++)
                {
                    TestAssert.Equal(control[page].Content, actual[page].Content);
                    var a = CustomMarkPhysicalGlyphs(actual[page]); var b = CustomMarkPhysicalGlyphs(control[page]);
                    TestAssert.Equal(b.Length, a.Length);
                    for (int glyph = 0; glyph < a.Length; glyph++) TestAssert.Equal(b[glyph], a[glyph]);
                    TestAssert.Equal(control[page].Annotations.Count, actual[page].Annotations.Count);
                    for (int link = 0; link < actual[page].Annotations.Count; link++)
                        TestAssert.Equal(control[page].Annotations[link], actual[page].Annotations[link]);
                }
                string text = string.Concat(actual.SelectMany(CustomMarkPhysicalGlyphs).Select(g => g.Text));
                TestAssert.Equal(count, text.Split("ratetime", StringSplitOptions.None).Length - 1);
                TestAssert.Equal(4, actual.Sum(page => page.Annotations.Count(link => link.Destination is not null)));
                TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Resolved casing cannot discard supported mathematics.");
            }
        }
    }

    private static (string Text, double Size, double X, double Y)[] CustomMarkPhysicalGlyphs(PdfPage page) =>
        DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true, includeTextTransforms: true, honorActualText: false)
            .Select(show => (show.Text, show.Size, show.X, show.Y)).ToArray();

    private static DocxDocument ReadCustomMarkText(string kind, int count, OoxPdfDocxMarkupMode mode, string variant) =>
        ReadMutation("CompositionProbes/math-custom-note-" + kind + "-" + count, mode, false, false, (xml, w, _) =>
        {
            var reference = xml.Descendants(w + kind + "Reference").Single(r => r.Attribute(w + "customMarkFollows")?.Value == "1");
            var mark = reference.ElementsAfterSelf().First(e => e.Name == w + "t");
            switch (variant)
            {
                case "original": break;
                case "empty": mark.AddBeforeSelf(new XElement(w + "t")); break;
                case "several-empty": mark.AddBeforeSelf(Enumerable.Range(0, 4).Select(_ =>
                    new XElement(w + "t", new XAttribute(XNamespace.Xml + "space", "preserve")))); break;
                case "trimmed-empty": mark.AddBeforeSelf(new XElement(w + "t", " \t\n ")); break;
                case "trimmed-mark": mark.Attribute(XNamespace.Xml + "space")?.Remove(); mark.Value = " \t* \n"; break;
                case "page-hint": mark.AddBeforeSelf(new XElement(w + "lastRenderedPageBreak")); break;
                case "page-hint-namespace": mark.AddBeforeSelf(new XElement(w + "lastRenderedPageBreak",
                    new XAttribute(XNamespace.Xmlns + "w", w.NamespaceName))); break;
                case "page-hints-and-empty": mark.AddBeforeSelf(new XElement(w + "lastRenderedPageBreak"),
                    new XElement(w + "t"), new XElement(w + "lastRenderedPageBreak", string.Empty), new XElement(w + "t", " \t ")); break;
                case "preserved-space": mark.AddBeforeSelf(new XElement(w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), " ")); break;
                case "tab": mark.AddBeforeSelf(new XElement(w + "tab")); break;
                case "break": mark.AddBeforeSelf(new XElement(w + "br")); break;
                case "page-break": mark.AddBeforeSelf(new XElement(w + "br", new XAttribute(w + "type", "page"))); break;
                case "page-hint-attribute": mark.AddBeforeSelf(new XElement(w + "lastRenderedPageBreak", new XAttribute(w + "future", "1"))); break;
                case "page-hint-text": mark.AddBeforeSelf(new XElement(w + "lastRenderedPageBreak", "payload")); break;
                case "page-hint-child": mark.AddBeforeSelf(new XElement(w + "lastRenderedPageBreak", new XElement(w + "tab"))); break;
                case "reference": mark.AddBeforeSelf(new XElement(w + kind + "Reference", new XAttribute(w + "id", "4"))); break;
                case "unknown": mark.AddBeforeSelf(new XElement(XName.Get("future", "urn:future"))); break;
                case "nested-empty": mark.AddBeforeSelf(new XElement(w + "t", new XElement(XName.Get("future", "urn:future")))); break;
                case "surrogate": mark.AddBeforeSelf(new XElement(w + "t", "\ud83d\ude00")); break;
                case "end-of-run": mark.Remove(); reference.AddAfterSelf(new XElement(w + "t")); break;
                case "other-run":
                    mark.Remove(); reference.AddAfterSelf(new XElement(w + "t"));
                    reference.Parent!.AddAfterSelf(new XElement(w + "r", new XElement(w + "t", "*"))); break;
                default: throw new InvalidOperationException("Unknown custom-mark text mutation.");
            }
        });

    private static DocxDocument ReadCasedCustomMark(string kind, int count, OoxPdfDocxMarkupMode mode, string letter, bool empty)
    {
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        void SetMark(XElement text)
        {
            text.Value = letter + text.Value[1..];
            var run = text.Parent!; var properties = run.Element(w + "rPr");
            if (properties is null) { properties = new XElement(w + "rPr"); run.AddFirst(properties); }
            properties.Elements(w + "caps").Remove(); properties.Add(new XElement(w + "caps"));
        }
        return ReadMutation("CompositionProbes/math-custom-note-" + kind + "-" + count, mode, false, false,
            (xml, _, _) =>
            {
                var reference = xml.Descendants(w + kind + "Reference").Single(r => r.Attribute(w + "customMarkFollows")?.Value == "1");
                var mark = reference.ElementsAfterSelf().First(e => e.Name == w + "t"); SetMark(mark);
                if (empty) mark.AddBeforeSelf(new XElement(w + "t"));
            }, packageChange: archive =>
            {
                var entry = archive.GetEntry("word/" + kind + "s.xml")!; XDocument notes;
                using (var input = entry.Open()) notes = XDocument.Load(input);
                var note = notes.Root!.Elements(w + kind).Single(n => n.Attribute(w + "id")?.Value == "37");
                SetMark(note.Descendants(w + "t").First(t => t.Value.StartsWith('*')));
                entry.Delete(); using var output = archive.CreateEntry("word/" + kind + "s.xml").Open(); notes.Save(output);
            });
    }
}
