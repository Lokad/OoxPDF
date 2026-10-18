using System.IO.Compression;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static class DocxOrdinaryKerningTests
{
    public static void DocxOrdinaryKerningReadsBasedOnCascadeAndExplicitZero()
    {
        var document = Read((xml, styles) =>
        {
            styles.Root!.Add(Style("OrdinaryBase", "paragraph", "16"));
            styles.Root.Add(Style("OrdinaryChild", "paragraph", "40", "OrdinaryBase"));
            styles.Root.Add(Style("OrdinaryCharacter", "character", "16"));
            var paragraphs = xml.Descendants(Word + "p").Take(4).ToArray();
            foreach (int i in new[] { 1, 2, 3 })
            {
                paragraphs[i].Element(Word + "pPr")!.Elements(Word + "pStyle").Remove();
                paragraphs[i].Element(Word + "pPr")!.AddFirst(new XElement(Word + "pStyle", new XAttribute(Word + "val", i == 1 ? "OrdinaryBase" : "OrdinaryChild")));
            }
            var rp = paragraphs[3].Element(Word + "r")!.Element(Word + "rPr")!;
            rp.AddFirst(new XElement(Word + "rStyle", new XAttribute(Word + "val", "OrdinaryCharacter")));
            rp.Add(Kern("0"));
        });
        TestAssert.Equal(0, document.RelatedStories.Count);
        TestAssert.True(document.Paragraphs.All(p => p.InlineReferences.Count == 0), "Ordinary cascade does not depend on unused note settings.");
        var runs = document.Paragraphs.Take(4).Select(p => p.Runs[0]).ToArray();
        TestAssert.Equal("28,16,40,0", string.Join(',', runs.Select(r => r.KerningMinimumHalfPointsValue)));
        TestAssert.Equal("28,16,40,0", string.Join(',', runs.Select(r => r.EffectiveProperties.KerningMinimumHalfPointsValue)));
        TestAssert.True(runs.All(r => r.PairKerningEnabled), "Reading retains source metadata and fallback pair policy.");
    }

    public static void DocxOrdinaryKerningKeepsMinimumSlotsAndActualNominalEmission()
    {
        var source = Read(); var before = Layout(source);
        TestAssert.True(DocxPlainFootnoteKerning.TryCreateDocument(source, before, out var adjusted, out var pages), "Plain ordinary 12pt text admits the measured source threshold policy.");
        TestAssert.Equal(0, pages.Count);
        var after = Layout(adjusted);
        TestAssert.Equal(before.Pages.Count, after.Pages.Count);
        var oldLines = before.Pages.SelectMany(p => p.Items).OfType<DocxTextLineLayout>().ToArray();
        var newLines = after.Pages.SelectMany(p => p.Items).OfType<DocxTextLineLayout>().ToArray();
        TestAssert.Equal(oldLines.Length, newLines.Length);
        for (int i = 0; i < oldLines.Length; i++)
        {
            TestAssert.Equal(oldLines[i].BaselineY, newLines[i].BaselineY);
            TestAssert.Equal(oldLines[i].LineHeight, newLines[i].LineHeight);
        }
        TestAssert.True(DocxMinimumBodyLineLeading.TryCreateLayout(source, before, out var oldPlaced), "The parent minimum-leading model admits ordinary source.");
        TestAssert.True(DocxMinimumBodyLineLeading.TryCreateLayout(adjusted, after, out var newPlaced), "Nominal advances retain minimum-leading admission.");
        TestAssert.Equal(oldPlaced.Pages[0].Items.OfType<DocxTextLineLayout>().First().BaselineY,
            newPlaced.Pages[0].Items.OfType<DocxTextLineLayout>().First().BaselineY);
        var resolver = new TestFaceFontResolver();
        double Measure(DocxDocument doc) => Measurer(doc, resolver).MeasureText(doc.Paragraphs[0].Runs[0], "AB", 12d);
        TestAssert.True(Math.Abs(Measure(adjusted) - Measure(source) - .6d) < .00001d, "The independent synthetic AB pair expands by its real -50-unit adjustment.");
        var word = Renderer(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        var emission = word.InspectTextEmission(source);
        TestAssert.True(emission.Lines.SelectMany(l => l.Segments).All(s => s.GlyphAdvanceSignature.KerningUnits == 0), "Actual Word-compatible emission uses nominal advances.");
        TestAssert.True(emission.Lines.SelectMany(l => l.Segments).All(s => Math.Abs(s.AdvanceProfile.KerningAdjustmentTotal) < .00001d), "Measured and emitted pair positioning agree.");
        TestAssert.Equal(0, word.RenderBlankPages(source, null, CancellationToken.None).Sum(p => p.Annotations.Count));
        var preserve = Renderer(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectTextEmission(source);
        TestAssert.True(preserve.Lines.SelectMany(l => l.Segments).All(s => s.GlyphAdvanceSignature.KerningUnits == 0),
            "Preserve and Word-compatible painting share the authored minimum, independent of layout admission.");
        TestAssert.True(source.Paragraphs[0].Runs[0].PairKerningEnabled, "Rendering leaves source pair policy unchanged.");
    }

    public static void DocxOrdinaryKerningRetainsPaintedWrappedKeptAndMalformedFallbacks()
    {
        Action<XDocument, XDocument>[] guards =
        [
            (xml, _) => FirstRunProperties(xml).Add(new XElement(Word + "b")),
            (xml, _) => FirstRunProperties(xml).Add(new XElement(Word + "highlight", new XAttribute(Word + "val", "yellow"))),
            (xml, _) => FirstRunProperties(xml).Add(new XElement(Word + "shd", new XAttribute(Word + "fill", "FFEEAA"))),
            (xml, _) => FirstRunProperties(xml).Add(Kern("-1")),
            (xml, _) => FirstRunProperties(xml).Add(new XElement(Word + "kern")),
            (xml, _) => FirstRunProperties(xml).Element(Word + "sz")!.SetAttributeValue(Word + "val", "28"),
            (xml, _) => xml.Descendants(Word + "p").First().Element(Word + "pPr")!.Add(new XElement(Word + "keepNext")),
            (xml, _) => xml.Descendants(Word + "t").First().Value += string.Concat(Enumerable.Repeat(" Public wrapped AB.", 40)),
            (xml, _) => xml.Descendants(Word + "t").First().Value += " e\u0301",
            (xml, _) => xml.Descendants(Word + "sectPr").Single().Add(new XElement(Word + "cols", new XAttribute(Word + "num", "2"))),
            (xml, styles) => { foreach (var rp in xml.Descendants(Word + "rPr")) rp.Add(Kern("24")); }
        ];
        foreach (var modify in guards)
        {
            var source = Read(modify);
            TestAssert.True(!DocxPlainFootnoteKerning.TryCreateDocument(source, Layout(source), out var fallback, out _), "Rich/painted, malformed, other-size, kept, wrapped, combining-text, column and already-enabled inputs retain fallback.");
            TestAssert.True(ReferenceEquals(source, fallback), "Excluded ordinary flow retains the original document object.");
        }
    }

    private static DocxRenderer Renderer(OoxPdfDocxMarkupGeometryMode geometry) => new(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup, geometry);
    private static DocxFontPlanTextMeasurer Measurer(DocxDocument document, TestFaceFontResolver resolver) =>
        new(DocxFontPlan.Create(document, resolver, CancellationToken.None), resolver.Resolve(new FontRequest("Calibri", false, false)), CancellationToken.None, resolver);
    private static DocxLayout Layout(DocxDocument document) => new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
        .Create(document, Measurer(document, new TestFaceFontResolver()), CancellationToken.None);
    private static readonly XNamespace Word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static XElement Kern(string value) => new(Word + "kern", new XAttribute(Word + "val", value));
    private static XElement Style(string id, string type, string threshold, string? parent = null) => new(Word + "style",
        new XAttribute(Word + "type", type), new XAttribute(Word + "styleId", id),
        parent is null ? null : new XElement(Word + "basedOn", new XAttribute(Word + "val", parent)), new XElement(Word + "rPr", Kern(threshold)));
    private static XElement FirstRunProperties(XDocument xml) => xml.Descendants(Word + "rPr").First();
    private static DocxDocument Read(Action<XDocument, XDocument>? modify = null)
    {
        string input = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "ordinary-minimum-line-20.docx"));
        var parts = new Dictionary<string, string>();
        using (var archive = ZipFile.OpenRead(input))
            foreach (var entry in archive.Entries) { using var reader = new StreamReader(entry.Open()); parts[entry.FullName] = reader.ReadToEnd(); }
        var xml = XDocument.Parse(parts["word/document.xml"]); var styles = XDocument.Parse(parts["word/styles.xml"]);
        xml.Descendants(Word + "t").First().Value = "Public AB ordinary minimum.";
        var defaults = styles.Root!.Element(Word + "docDefaults")!;
        var runDefault = defaults.Element(Word + "rPrDefault");
        if (runDefault is null) { runDefault = new XElement(Word + "rPrDefault"); defaults.AddFirst(runDefault); }
        var rp = runDefault.Element(Word + "rPr");
        if (rp is null) { rp = new XElement(Word + "rPr"); runDefault.Add(rp); }
        rp.Elements(Word + "kern").Remove(); rp.Add(Kern("28"));
        modify?.Invoke(xml, styles); parts["word/document.xml"] = xml.ToString(); parts["word/styles.xml"] = styles.ToString();
        string temporary = TestFixtures.WriteTempPackage(".docx", parts);
        try { return DocxTests.ReadDocx(temporary, OoxPdfDocxMarkupMode.AllMarkup); }
        finally { File.Delete(temporary); }
    }
}
