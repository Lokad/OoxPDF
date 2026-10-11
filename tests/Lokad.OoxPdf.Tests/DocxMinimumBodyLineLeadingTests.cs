using System.IO.Compression;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static class DocxMinimumBodyLineLeadingTests
{
    public static void DocxMinimumBodyLineLeadingKeepsCapacityAndParagraphTop()
    {
        var document = ReadFixture();
        var before = Layout(document);
        TestAssert.True(DocxMinimumBodyLineLeading.TryCreateLayout(document, before, out var after), "A regular minimum-height line must qualify with complete actual note owners.");
        var a = Minimum(before); var z = Minimum(after);
        double excess = a.LineHeight!.Value - a.SingleLineHeight!.Value;
        TestAssert.True(excess > 0 && Math.Abs(a.BaselineY - z.BaselineY - excess) < .00001d, "The baseline uses measured positive minimum excess.");
        TestAssert.True(Math.Abs(a.BaselineY + a.FirstLineInsetPoints!.Value - z.BaselineY - z.FirstLineInsetPoints!.Value) < .00001d, "The authored paragraph top stays fixed.");
        TestAssert.Equal(a.LineHeight, z.LineHeight);
        TestAssert.Equal(before.Pages.Count, after.Pages.Count);
        TestAssert.True(ReferenceEquals(before.RelatedStories, after.RelatedStories), "Note line boxes and destinations retain their original layout.");
        TestAssert.True(before.Pages.Zip(after.Pages).All(pair => ReferenceEquals(pair.First.PlacedRelatedStories, pair.Second.PlacedRelatedStories)), "Placed notes retain their actual page and geometry.");
        TestAssert.True(DocxPageFootnoteNumbering.HasStableCompleteNoteOwners(document, before, after), "Minimum leading does not repaginate complete owners.");
        TestAssert.Equal(a.Text, z.Text);
        TestAssert.True(ReferenceEquals(a.Segments, z.Segments), "Horizontal positioning and glyph plans remain the same.");
    }

    public static void DocxMinimumBodyLineLeadingKeepsRuleMarkerAndInheritedGuards()
    {
        Action<XDocument, XDocument>[] guards = [
            (xml, _) => Target(xml).Element(Word + "pPr")!.Element(Word + "spacing")!.SetAttributeValue(Word + "lineRule", "exact"),
            (xml, _) => Target(xml).Element(Word + "pPr")!.Element(Word + "spacing")!.SetAttributeValue(Word + "lineRule", "auto"),
            (xml, _) => Target(xml).Element(Word + "pPr")!.Element(Word + "spacing")!.SetAttributeValue(Word + "line", "560"),
            (xml, _) => Target(xml).Element(Word + "pPr")!.Add(new XElement(Word + "keepNext")),
            (xml, _) => Target(xml).Element(Word + "pPr")!.Add(new XElement(Word + "keepLines")),
            (xml, _) => Target(xml).Descendants(Word + "rPr").First().Add(new XElement(Word + "b")),
            (xml, _) => Target(xml).Descendants(Word + "sz").First().SetAttributeValue(Word + "val", "28"),
            (xml, _) => { var p = Target(xml); p.Element(Word + "pPr")!.Element(Word + "spacing")!.SetAttributeValue(Word + "line", "280");
                var fonts = p.Descendants(Word + "rFonts").First(); fonts.SetAttributeValue(Word + "ascii", "Times New Roman"); fonts.SetAttributeValue(Word + "hAnsi", "Times New Roman"); },
            (xml, _) => { var p = Target(xml); var run = new XElement(p.Element(Word + "r")!); var fonts = run.Descendants(Word + "rFonts").First();
                fonts.SetAttributeValue(Word + "ascii", "Arial"); fonts.SetAttributeValue(Word + "hAnsi", "Arial"); run.Element(Word + "t")!.Value = " Mixed face."; p.Add(run); },
            (xml, _) => Target(xml).Descendants(Word + "t").First().Value += string.Concat(Enumerable.Repeat(" Public wrapped minimum.", 40)),
            (xml, _) => { var p = xml.Descendants(Word + "p").First(p => p.Descendants(Word + "footnoteReference").Any());
                var spacing = p.Element(Word + "pPr")!.Element(Word + "spacing")!;
                spacing.SetAttributeValue(Word + "line", "400"); spacing.SetAttributeValue(Word + "lineRule", "atLeast"); },
        ];
        foreach (var modify in guards)
        {
            var document = ReadFixture(modify); var before = Layout(document);
            TestAssert.True(!DocxMinimumBodyLineLeading.TryCreateLayout(document, before, out var fallback), "Exact/auto, unqualified slots/styles, wrapped, marked, keep and inherited-default paths retain fallback.");
            TestAssert.True(ReferenceEquals(before, fallback), "An exclusion keeps the entire original layout.");
        }
    }

    public static void DocxMinimumBodyLineLeadingKeepsEmissionAndPreserveGeometryCoherent()
    {
        var document = ReadFixture(); var original = Minimum(Layout(document));
        DocxRenderer Renderer(OoxPdfDocxMarkupGeometryMode geometry) => new(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup, geometry);
        var preserve = Renderer(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectTextEmission(document);
        var word = Renderer(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).InspectTextEmission(document);
        double Baseline(DocxTextEmissionSnapshot snapshot) => snapshot.Lines.Single(l => l.StoryKind == "Body" && l.SourceBlockIndex == original.SourceBlockIndex).Segments.First().BaselineY;
        double excess = original.LineHeight!.Value - original.SingleLineHeight!.Value;
        TestAssert.True(Math.Abs(Baseline(preserve) - original.BaselineY) < .00001d, "Preserve output retains the original body baseline.");
        TestAssert.True(Math.Abs(Baseline(preserve) - Baseline(word) - excess) < .00001d, "The emitted baseline follows the same minimum-height layout rule.");
        var pages = Renderer(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(document, null, CancellationToken.None).ToArray();
        TestAssert.Equal(4, pages.Sum(p => p.Annotations.Count));
        TestAssert.Equal(4, pages.SelectMany(p => p.Annotations).Select(a => a.Destination).Distinct().Count());
        for (int i = 0; i < pages.Length; i++)
            TestAssert.True(pages[i].Annotations.All(a => a.Destination!.Value.PageIndex == i), "Text correction retains actual same-page note targets.");
        TestAssert.Equal("105", document.Paragraphs.First(p => p.InlineReferences.Count != 0).InlineReferences[0].DisplayText);
    }

    private static DocxTextLineLayout Minimum(DocxLayout layout) => layout.Pages.SelectMany(p => p.Items)
        .OfType<DocxTextLineLayout>().Single(l => l.Text.StartsWith("Public minimum body"));
    private static DocxLayout Layout(DocxDocument document)
    {
        var resolver = new TestFaceFontResolver();
        var measurer = new DocxFontPlanTextMeasurer(DocxFontPlan.Create(document, resolver, CancellationToken.None),
            resolver.Resolve(new FontRequest("Calibri", false, false)), CancellationToken.None, resolver);
        return new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, measurer, CancellationToken.None);
    }
    private static readonly XNamespace Word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static XElement Target(XDocument xml) => xml.Descendants(Word + "p").Single(p => p.Descendants(Word + "t").Any(t => t.Value.StartsWith("Public minimum body")));
    private static DocxDocument ReadFixture(Action<XDocument, XDocument>? modify = null)
    {
        string input = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "note-complete-reserve-3.docx"));
        var parts = new Dictionary<string, string>();
        using (ZipArchive archive = ZipFile.OpenRead(input))
            foreach (var entry in archive.Entries) { using var reader = new StreamReader(entry.Open()); parts[entry.FullName] = reader.ReadToEnd(); }
        var xml = XDocument.Parse(parts["word/document.xml"]); var styles = XDocument.Parse(parts["word/styles.xml"]);
        var p = xml.Descendants(Word + "p").First(p => !p.Descendants(Word + "footnoteReference").Any() && p.Element(Word + "pPr")?.Element(Word + "spacing")?.Attribute(Word + "lineRule")?.Value == "atLeast");
        p.Descendants(Word + "sz").First().SetAttributeValue(Word + "val", "24");
        p.Element(Word + "pPr")!.Element(Word + "spacing")!.SetAttributeValue(Word + "line", "400");
        p.Descendants(Word + "t").First().Value = "Public minimum body AB.";
        modify?.Invoke(xml, styles);
        parts["word/document.xml"] = xml.ToString(); parts["word/styles.xml"] = styles.ToString();
        string temporary = TestFixtures.WriteTempPackage(".docx", parts);
        try { return DocxTests.ReadDocx(temporary, OoxPdfDocxMarkupMode.AllMarkup); }
        finally { File.Delete(temporary); }
    }
}
