using System.IO.Compression;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static class DocxMinimumBodyFlowTests
{
    public static void DocxMinimumBodyFlowDefaultNotesKeepSlotsRegionsAndActualTargets()
    {
        var document = Read(ordinary: false);
        TestAssert.True(document.Paragraphs.All(p => p.StyleResolution.HasDocumentDefaultParagraphProperties), "The fixture exercises inherited minimum defaults.");
        var before = Layout(document);
        TestAssert.True(DocxMinimumBodyLineLeading.TryCreateLayout(document, before, out var after), "Coherent inherited minimum defaults qualify with complete note owners.");
        AssertPlacement(before, after);
        TestAssert.True(ReferenceEquals(before.RelatedStories, after.RelatedStories), "Global note regions retain their layout identity.");
        for (int i = 0; i < before.Pages.Count; i++)
            TestAssert.True(ReferenceEquals(before.Pages[i].PlacedRelatedStories, after.Pages[i].PlacedRelatedStories), "Placed note regions retain page geometry.");
        var pages = Renderer(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(document, null, CancellationToken.None).ToArray();
        TestAssert.Equal(4, pages.Sum(p => p.Annotations.Count));
        TestAssert.Equal(4, pages.SelectMany(p => p.Annotations).Select(a => a.Destination).Distinct().Count());
        for (int i = 0; i < pages.Length; i++)
            TestAssert.True(pages[i].Annotations.All(a => a.Destination!.Value.PageIndex == i), "Actual note targets retain same-page owners.");
        TestAssert.Equal("105", document.Paragraphs.First(p => p.InlineReferences.Count != 0).InlineReferences[0].DisplayText);
    }

    public static void DocxMinimumBodyFlowOrdinaryPlacementKeepsCapacityAndPreserveEmission()
    {
        var document = Read(ordinary: true);
        TestAssert.Equal(0, document.RelatedStories.Count);
        TestAssert.True(document.Paragraphs.All(p => p.InlineReferences.Count == 0), "Ordinary body flow has no note references.");
        TestAssert.True(document.FinalSectionBreak!.PageSettings.FootnoteReferenceSettings.NumberRestartValue is null, "Ordinary admission must not require an unused each-page setting.");
        var before = Layout(document);
        TestAssert.True(DocxMinimumBodyLineLeading.TryCreateLayout(document, before, out var after), "Ordinary minimum lines qualify independently of note numbering.");
        AssertPlacement(before, after);
        double Baseline(DocxTextEmissionSnapshot snapshot) => snapshot.Lines.Single(l => l.StoryKind == "Body" && l.SourceBlockIndex == 0).Segments.First().BaselineY;
        TestAssert.Equal(Minimum(before)[0].BaselineY, Baseline(Renderer(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectTextEmission(document)));
        TestAssert.Equal(Minimum(after)[0].BaselineY, Baseline(Renderer(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).InspectTextEmission(document)));
        TestAssert.Equal(0, Renderer(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(document, null, CancellationToken.None).Sum(p => p.Annotations.Count));
    }

    public static void DocxMinimumBodyFlowOrdinaryRichWrappedAndColumnPathsKeepFallback()
    {
        Action<XDocument>[] guards =
        [
            xml => MinimumParagraph(xml).Element(Word + "r")!.Element(Word + "rPr")!.Add(new XElement(Word + "highlight", new XAttribute(Word + "val", "yellow"))),
            xml => MinimumParagraph(xml).Element(Word + "r")!.Element(Word + "rPr")!.Add(new XElement(Word + "shd", new XAttribute(Word + "fill", "FFEEAA"))),
            xml => MinimumParagraph(xml).Element(Word + "pPr")!.Add(new XElement(Word + "keepNext")),
            xml => MinimumParagraph(xml).Descendants(Word + "t").First().Value += string.Concat(Enumerable.Repeat(" Public wrapped minimum.", 40)),
            xml => xml.Descendants(Word + "sectPr").Single().Add(new XElement(Word + "cols", new XAttribute(Word + "num", "2"))),
            xml => MinimumParagraph(xml).Descendants(Word + "t").First().Value += " e\u0301"
        ];
        foreach (var modify in guards)
        {
            var document = Read(ordinary: true, modify); var before = Layout(document);
            TestAssert.True(!DocxMinimumBodyLineLeading.TryCreateLayout(document, before, out var after), "Painted, kept, wrapped, column and combining-text paths retain the original layout.");
            TestAssert.True(ReferenceEquals(before, after), "Exclusions preserve the complete layout object.");
        }
    }

    private static void AssertPlacement(DocxLayout before, DocxLayout after)
    {
        var oldLines = Minimum(before); var newLines = Minimum(after);
        TestAssert.Equal(4, oldLines.Length); TestAssert.Equal(4, newLines.Length);
        TestAssert.Equal(before.Pages.Count, after.Pages.Count);
        for (int i = 0; i < oldLines.Length; i++)
        {
            var a = oldLines[i]; var z = newLines[i]; double excess = a.LineHeight!.Value - a.SingleLineHeight!.Value;
            TestAssert.True(excess > 0 && Math.Abs(a.BaselineY - z.BaselineY - excess) < .00001d, "Positive metric excess moves above the baseline.");
            TestAssert.True(Math.Abs(a.BaselineY + a.FirstLineInsetPoints!.Value - z.BaselineY - z.FirstLineInsetPoints!.Value) < .00001d, "The paragraph top remains fixed.");
            TestAssert.Equal(a.LineHeight, z.LineHeight);
            TestAssert.True(ReferenceEquals(a.Segments, z.Segments), "Glyph segments and advances retain their identity.");
        }
    }

    private static DocxTextLineLayout[] Minimum(DocxLayout layout) => layout.Pages.SelectMany(p => p.Items).OfType<DocxTextLineLayout>()
        .Where(l => l.Text.StartsWith("Public minimum ")).ToArray();
    private static DocxRenderer Renderer(OoxPdfDocxMarkupGeometryMode geometry) => new(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup, geometry);
    private static DocxLayout Layout(DocxDocument document)
    {
        var resolver = new TestFaceFontResolver();
        var measurer = new DocxFontPlanTextMeasurer(DocxFontPlan.Create(document, resolver, CancellationToken.None),
            resolver.Resolve(new FontRequest("Calibri", false, false)), CancellationToken.None, resolver);
        return new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, measurer, CancellationToken.None);
    }
    private static readonly XNamespace Word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static XElement MinimumParagraph(XDocument xml) => xml.Descendants(Word + "p").First(p => p.Descendants(Word + "t").Any(t => t.Value.StartsWith("Public minimum ")));
    private static DocxDocument Read(bool ordinary, Action<XDocument>? modify = null)
    {
        string fixture = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "note-minimum-body-line-16.docx"));
        var parts = new Dictionary<string, string>();
        using (var archive = ZipFile.OpenRead(fixture))
            foreach (var entry in archive.Entries) { using var reader = new StreamReader(entry.Open()); parts[entry.FullName] = reader.ReadToEnd(); }
        var xml = XDocument.Parse(parts["word/document.xml"]);
        var styles = XDocument.Parse(parts["word/styles.xml"]);
        var defaults = styles.Root!.Element(Word + "docDefaults")!;
        defaults.Element(Word + "pPrDefault")?.Remove();
        defaults.Add(new XElement(Word + "pPrDefault", new XElement(Word + "pPr", new XElement(Word + "spacing", new XAttribute(Word + "line", "400"), new XAttribute(Word + "lineRule", "atLeast")))));
        foreach (var p in xml.Descendants(Word + "p").Where(p => p.Descendants(Word + "t").Any(t => t.Value.StartsWith("Public minimum "))))
        {
            var spacing = p.Element(Word + "pPr")!.Element(Word + "spacing")!;
            spacing.Attribute(Word + "line")?.Remove(); spacing.Attribute(Word + "lineRule")?.Remove();
        }
        if (ordinary)
        {
            xml.Descendants(Word + "footnoteReference").Remove(); xml.Descendants(Word + "footnotePr").Remove();
            var settings = XDocument.Parse(parts["word/settings.xml"]); settings.Descendants(Word + "footnotePr").Remove(); parts["word/settings.xml"] = settings.ToString();
            parts.Remove("word/footnotes.xml");
            foreach (string name in new[] { "word/_rels/document.xml.rels", "[Content_Types].xml" })
            {
                var package = XDocument.Parse(parts[name]);
                package.Root!.Elements().Where(e => (e.Attribute("Type")?.Value.Contains("footnotes") ?? false) || e.Attribute("PartName")?.Value == "/word/footnotes.xml").Remove();
                parts[name] = package.ToString();
            }
        }
        modify?.Invoke(xml); parts["word/document.xml"] = xml.ToString(); parts["word/styles.xml"] = styles.ToString();
        string path = TestFixtures.WriteTempPackage(".docx", parts);
        try { return DocxTests.ReadDocx(path, OoxPdfDocxMarkupMode.AllMarkup); }
        finally { File.Delete(path); }
    }
}
