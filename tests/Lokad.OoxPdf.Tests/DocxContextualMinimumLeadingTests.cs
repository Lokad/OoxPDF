using System.IO.Compression;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static class DocxContextualMinimumLeadingTests
{
    public static void DocxContextualMinimumKeepsResolvedStylesAndAutomaticNeighbors()
    {
        foreach (string styleMode in new[] { "implicit", "named", "mixed-default" })
        {
            var source = Read(false, (xml, styles) =>
            {
                if (styleMode == "implicit") return;
                AddStyle(styles, "PublicContextualA", styleMode == "mixed-default");
                var paragraphs = Body(xml);
                for (int i = 0; i < paragraphs.Length; i++)
                    if (styleMode == "named" || i % 2 != 0)
                        paragraphs[i].Element(Word + "pPr")!.AddFirst(new XElement(Word + "pStyle", new XAttribute(Word + "val", "PublicContextualA")));
            });
            TestAssert.True(source.Paragraphs.All(p => p.Spacing.ContextualSpacing == true && p.Spacing.BeforeValue == "80" && p.Spacing.AfterValue == "80"),
                "Raw contextual flags and numeric tokens remain available.");
            var original = Layout(source);
            TestAssert.True(DocxMinimumBodyLineLeading.TryCreateLayout(source, original, out var adjusted), "Resolved contextual identities admit plain minimum leading.");
            var before = Lines(original); var after = Lines(adjusted);
            TestAssert.Equal(before.Length, after.Length);
            for (int i = 0; i < before.Length; i++)
            {
                if (before[i].LineHeightSource != DocxLineHeightSource.AtLeastLineSpacing)
                {
                    TestAssert.True(ReferenceEquals(before[i], after[i]), "Automatic neighbors retain their entire original lines.");
                    continue;
                }
                Near(before[i].LineHeight!.Value - before[i].SingleLineHeight!.Value, before[i].BaselineY - after[i].BaselineY, "Minimum excess moves its visible baseline.");
                Near(before[i].BaselineY + before[i].FirstLineInsetPoints!.Value, after[i].BaselineY + after[i].FirstLineInsetPoints!.Value,
                    "Contextual gap resolution and the authored paragraph top remain fixed.");
                TestAssert.Equal(before[i].AppliedBeforeSpacing, after[i].AppliedBeforeSpacing);
                TestAssert.Equal(before[i].PendingAfterSpacing, after[i].PendingAfterSpacing);
                TestAssert.True(ReferenceEquals(before[i].Segments, after[i].Segments), "Horizontal glyph plans retain identity.");
            }
            foreach (var mode in new[] { OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup })
                TestAssert.Equal(0, Renderer(mode).RenderBlankPages(source, null, CancellationToken.None).Sum(p => p.Annotations.Count));
        }
    }

    public static void DocxContextualMinimumActualNoteLinksRetainContextualGapOwnership()
    {
        foreach (var spec in new[] { (Mode: "same", Height: 20d), (Mode: "different", Height: 24d),
            (Mode: "predecessor-only", Height: 28d), (Mode: "current-only", Height: 20d) })
        {
            var source = Read(true, (xml, styles) =>
            {
                var paragraphs = Body(xml);
                if (spec.Mode == "different") { AddStyle(styles, "PublicContextualA"); AddStyle(styles, "PublicContextualB"); }
                for (int i = 0; i < paragraphs.Length; i++)
                {
                    var pr = paragraphs[i].Element(Word + "pPr")!;
                    if (spec.Mode == "different") pr.AddFirst(new XElement(Word + "pStyle", new XAttribute(Word + "val", i % 2 == 0 ? "PublicContextualA" : "PublicContextualB")));
                    if (spec.Mode is "predecessor-only" or "current-only")
                        pr.Element(Word + "contextualSpacing")!.SetAttributeValue(Word + "val", i % 2 == (spec.Mode == "current-only" ? 1 : 0) ? "1" : "0");
                }
                if (spec.Mode is "predecessor-only" or "current-only") Spacing(paragraphs[1]).SetAttributeValue(Word + "before", "160");
            });
            var original = Layout(source);
            TestAssert.True(DocxMinimumBodyLineLeading.TryCreateLayout(source, original, out var adjusted), "Complete contextual note owners admit the baseline correction.");
            TestAssert.True(DocxPageFootnoteNumbering.HasStableCompleteNoteOwners(source, original, adjusted), "Every actual source marker and complete note retains its page.");
            var pages = Renderer(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(source, null, CancellationToken.None).ToArray();
            var links = pages.SelectMany(p => p.Annotations).ToArray();
            TestAssert.Equal(4, links.Length);
            Near(spec.Height, links[0].Height, "The actual link owns the minimum slot and only the remaining contextual gaps.");
            TestAssert.Equal(4, links.Select(a => a.Destination).Distinct().Count());
            var preserve = Renderer(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(source, null, CancellationToken.None).SelectMany(p => p.Annotations).ToArray();
            Near(20, preserve[0].Height, "Preserve keeps its original slot hit area.");
        }
    }

    public static void DocxContextualMinimumRejectsUnknownStyleIdentitiesAndRichFlow()
    {
        foreach (bool notes in new[] { false, true })
        foreach (string guard in new[] { "missing-styles", "ambiguous-defaults", "unknown-named", "bold", "auto-spacing", "keep", "wrapped", "columns" })
        {
            var source = Read(notes, (xml, styles) =>
            {
                var paragraphs = Body(xml);
                if (guard == "ambiguous-defaults") { AddStyle(styles, "PublicContextualA", true); AddStyle(styles, "PublicContextualB", true); }
                if (guard == "unknown-named") paragraphs[0].Element(Word + "pPr")!.AddFirst(new XElement(Word + "pStyle", new XAttribute(Word + "val", "PublicMissing")));
                if (guard == "bold") paragraphs[0].Descendants(Word + "rPr").First().Add(new XElement(Word + "b"));
                if (guard == "auto-spacing") Spacing(paragraphs[0]).SetAttributeValue(Word + "beforeAutospacing", "1");
                if (guard == "keep") paragraphs[0].Element(Word + "pPr")!.Add(new XElement(Word + "keepLines"));
                if (guard == "wrapped") paragraphs[0].Descendants(Word + "t").First().Value += string.Concat(Enumerable.Repeat(" Public contextual wrapped paragraph.", 30));
                if (guard == "columns") xml.Descendants(Word + "sectPr").Single().Add(new XElement(Word + "cols", new XAttribute(Word + "num", "2")));
            }, removeStyles: guard == "missing-styles");
            var original = Layout(source);
            TestAssert.True(!DocxMinimumBodyLineLeading.TryCreateLayout(source, original, out var fallback), "Unknown contextual identities and unsupported complete frames retain fallback.");
            TestAssert.True(ReferenceEquals(original, fallback), "An excluded frame retains its entire original layout.");
        }
    }

    public static void DocxContextualMinimumInheritedFlagsKeepPreserveCapacityAndOwners()
    {
        foreach (bool notes in new[] { false, true })
        {
            var source = Read(notes, (xml, styles) =>
            {
                var defaults = styles.Root!.Element(Word + "docDefaults")!;
                defaults.Add(new XElement(Word + "pPrDefault", new XElement(Word + "pPr", new XElement(Word + "contextualSpacing"))));
                var paragraphs = Body(xml);
                foreach (var p in paragraphs) p.Element(Word + "pPr")!.Element(Word + "contextualSpacing")!.Remove();
                for (int i = 0; i < 40; i++)
                {
                    var spacer = new XElement(paragraphs[0]); spacer.Descendants(Word + "t").First().Value = $"Public contextual capacity spacer {i:00}.";
                    paragraphs[4].AddBeforeSelf(spacer);
                }
            });
            TestAssert.True(source.Paragraphs.All(p => p.Spacing.ContextualSpacing == true && p.ContextualSpacingDefaultStyleId == "Normal"), "Document-default flags retain a loaded implicit Normal identity.");
            var original = Layout(source);
            TestAssert.True(original.Pages.Count > 1, "The synthetic contextual frame exercises page boundaries.");
            TestAssert.True(DocxMinimumBodyLineLeading.TryCreateLayout(source, original, out var adjusted), "The complete inherited frame qualifies without repagination.");
            TestAssert.Equal(original.Pages.Count, adjusted.Pages.Count);
            foreach (var pair in original.Pages.Zip(adjusted.Pages))
                TestAssert.True(pair.First.Items.OfType<DocxTextLineLayout>().Select(l => l.SourceBlockIndex).SequenceEqual(pair.Second.Items.OfType<DocxTextLineLayout>().Select(l => l.SourceBlockIndex)), "Each actual body owner retains its page.");
            var first = Lines(original)[0];
            var preserve = Renderer(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectTextEmission(source).Lines.First(l => l.StoryKind == "Body" && l.SourceBlockIndex == first.SourceBlockIndex);
            var word = Renderer(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).InspectTextEmission(source).Lines.First(l => l.StoryKind == "Body" && l.SourceBlockIndex == first.SourceBlockIndex);
            Near(first.LineHeight!.Value - first.SingleLineHeight!.Value, preserve.Segments.First().BaselineY - word.Segments.First().BaselineY, "Actual requested Word emission changes while Preserve keeps the original baseline.");
            var pages = Renderer(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(source, null, CancellationToken.None).ToArray();
            TestAssert.Equal(notes ? 4 : 0, pages.Sum(p => p.Annotations.Count));
            if (notes)
            {
                TestAssert.True(DocxPageFootnoteNumbering.HasStableCompleteNoteOwners(source, original, adjusted), "Complete inherited note ownership survives capacity boundaries.");
                TestAssert.Equal("105", source.Paragraphs.First(p => p.InlineReferences.Count != 0).InlineReferences[0].DisplayText);
                for (int i = 0; i < pages.Length; i++) TestAssert.True(pages[i].Annotations.All(a => a.Destination!.Value.PageIndex == i), "Each actual note link targets its owner page.");
            }
        }
    }

    private static readonly XNamespace Word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static XElement[] Body(XDocument xml) => xml.Root!.Element(Word + "body")!.Elements(Word + "p").ToArray();
    private static XElement Spacing(XElement p) => p.Element(Word + "pPr")!.Element(Word + "spacing")!;
    private static DocxTextLineLayout[] Lines(DocxLayout layout) => layout.Pages.SelectMany(p => p.Items).OfType<DocxTextLineLayout>().ToArray();
    private static DocxRenderer Renderer(OoxPdfDocxMarkupGeometryMode mode) => new(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup, mode);
    private static void Near(double expected, double actual, string message) => TestAssert.True(Math.Abs(expected - actual) < .00001d, message);
    private static void AddStyle(XDocument styles, string id, bool isDefault = false)
    {
        var style = new XElement(Word + "style", new XAttribute(Word + "type", "paragraph"), new XAttribute(Word + "styleId", id), new XElement(Word + "name", new XAttribute(Word + "val", id)));
        if (isDefault) style.SetAttributeValue(Word + "default", "1");
        styles.Root!.Add(style);
    }
    private static DocxLayout Layout(DocxDocument document)
    {
        var resolver = new TestFaceFontResolver();
        var measurer = new DocxFontPlanTextMeasurer(DocxFontPlan.Create(document, resolver, CancellationToken.None), resolver.Resolve(new FontRequest("Calibri", false, false)), CancellationToken.None, resolver);
        return new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, measurer, CancellationToken.None);
    }
    private static DocxDocument Read(bool notes, Action<XDocument, XDocument>? modify = null, bool removeStyles = false)
    {
        string input = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", notes ? "note-marked-minimum-line-20.docx" : "ordinary-minimum-line-20.docx"));
        var parts = new Dictionary<string, string>();
        using (var zip = ZipFile.OpenRead(input))
            foreach (var entry in zip.Entries) { using var reader = new StreamReader(entry.Open()); parts[entry.FullName] = reader.ReadToEnd(); }
        var xml = XDocument.Parse(parts["word/document.xml"]); var styles = XDocument.Parse(parts["word/styles.xml"]);
        foreach (var p in Body(xml))
        {
            Spacing(p).SetAttributeValue(Word + "before", "80"); Spacing(p).SetAttributeValue(Word + "after", "80");
            p.Element(Word + "pPr")!.Add(new XElement(Word + "contextualSpacing"));
        }
        modify?.Invoke(xml, styles);
        parts["word/document.xml"] = xml.ToString(); parts["word/styles.xml"] = styles.ToString();
        if (notes)
        {
            var footnotes = XDocument.Parse(parts["word/footnotes.xml"]);
            foreach (var p in footnotes.Descendants(Word + "p"))
            {
                var pr = p.Element(Word + "pPr"); if (pr is null) { pr = new XElement(Word + "pPr"); p.AddFirst(pr); }
                var sp = pr.Element(Word + "spacing"); if (sp is null) { sp = new XElement(Word + "spacing"); pr.Add(sp); }
                sp.SetAttributeValue(Word + "before", "0");
                pr.Add(new XElement(Word + "contextualSpacing", new XAttribute(Word + "val", "0")));
            }
            parts["word/footnotes.xml"] = footnotes.ToString();
        }
        if (removeStyles) parts.Remove("word/styles.xml");
        string temporary = TestFixtures.WriteTempPackage(".docx", parts);
        try { return DocxTests.ReadDocx(temporary, OoxPdfDocxMarkupMode.AllMarkup); } finally { File.Delete(temporary); }
    }
}
