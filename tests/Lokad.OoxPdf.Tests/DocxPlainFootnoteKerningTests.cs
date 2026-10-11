using System.IO.Compression;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static class DocxPlainFootnoteKerningTests
{
    public static void DocxPlainFootnoteKerningRetainsStyleCascadeAndExplicitZero()
    {
        var document = ReadFixture((xml, _, styles) =>
        {
            styles.Root!.Element(Word + "docDefaults")?.Remove();
            styles.Root.AddFirst(new XElement(Word + "docDefaults", new XElement(Word + "rPrDefault", RPr("28"))));
            styles.Root.Add(new XElement(Word + "style", new XAttribute(Word + "type", "paragraph"),
                new XAttribute(Word + "styleId", "KernParagraph"), RPr("16")));
            styles.Root.Add(new XElement(Word + "style", new XAttribute(Word + "type", "character"),
                new XAttribute(Word + "styleId", "KernCharacter"), RPr("40")));
            XElement[] paragraphs = xml.Descendants(Word + "p").Where(p => p.Descendants(Word + "footnoteReference").Any()).ToArray();
            foreach (int i in new[] { 1, 3 }) paragraphs[i].Element(Word + "pPr")!.AddFirst(
                new XElement(Word + "pStyle", new XAttribute(Word + "val", "KernParagraph")));
            foreach (int i in new[] { 2, 3 }) paragraphs[i].Elements(Word + "r").First().Element(Word + "rPr")!.AddFirst(
                new XElement(Word + "rStyle", new XAttribute(Word + "val", "KernCharacter")));
            paragraphs[3].Elements(Word + "r").First().Element(Word + "rPr")!.Add(Kern("0"));
        });
        var runs = document.Paragraphs.Where(p => p.InlineReferences.Count != 0).Select(p => p.Runs[0]).ToArray();
        TestAssert.Equal("28,16,40,0", string.Join(',', runs.Select(r => r.KerningMinimumHalfPointsValue)));
        TestAssert.Equal("28,16,40,0", string.Join(',', runs.Select(r => r.EffectiveProperties.KerningMinimumHalfPointsValue)));
        TestAssert.True(runs.All(r => r.PairKerningEnabled), "Raw reading retains the renderer fallback while recording inherited thresholds and an explicit zero.");
    }

    public static void DocxPlainFootnoteKerningKeepsMeasurementAndEmissionCoherent()
    {
        var source = ReadFixture();
        var layout = Layout(source);
        TestAssert.True(DocxPlainFootnoteKerning.TryCreateDocument(source, layout, out var adjusted, out var pages), "The complete plain fixture must admit nominal advances.");
        var before = source.Paragraphs[0].Runs[0]; var after = adjusted.Paragraphs[0].Runs[0];
        TestAssert.Equal(before.Text, after.Text);
        TestAssert.Equal(before.SourceRunIndex, after.SourceRunIndex);
        TestAssert.Equal(before.SourceTextOffsetInRun, after.SourceTextOffsetInRun);
        TestAssert.True(before.PairKerningEnabled && !after.PairKerningEnabled, "Only the derived rendering document changes its pair policy.");
        var resolver = new TestFaceFontResolver();
        double Measure(DocxDocument doc, DocxTextRun run) => new DocxFontPlanTextMeasurer(
            DocxFontPlan.Create(doc, resolver, CancellationToken.None), resolver.Resolve(new FontRequest("Calibri", false, false)),
            CancellationToken.None, resolver).MeasureText(run, "AB", 12d);
        double kerned = Measure(source, before), nominal = Measure(adjusted, after);
        TestAssert.True(Math.Abs(nominal - kerned - 0.6d) < 0.00001d, "The synthetic AB pair has a genuine -50-unit adjustment at 12pt.");
        var renderer = Renderer(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        var emission = renderer.InspectTextEmission(source);
        var segments = emission.Lines.Where(l => l.StoryKind == "Body" && l.SourceBlockIndex == 0).SelectMany(l => l.Segments).ToArray();
        TestAssert.True(segments.Length > 0 && segments.All(s => s.GlyphAdvanceSignature.KerningUnits == 0), "Emission signatures must describe the actual nominal glyph advances.");
        TestAssert.True(segments.All(s => Math.Abs(s.AdvanceProfile.KerningAdjustmentTotal) < 0.00001d), "Layout measurement and emitted positioning use the same pair policy.");
        TestAssert.True(DocxPageFootnoteNumbering.HasStablePages(adjusted, Layout(adjusted), pages), "The derived document retains complete marker/note owners.");
        var rendered = renderer.RenderBlankPages(source, null, CancellationToken.None).ToArray();
        TestAssert.Equal(4, rendered.Sum(p => p.Annotations.Count));
        TestAssert.Equal(4, rendered.SelectMany(p => p.Annotations).Select(a => a.Destination).Distinct().Count());
        for (int i = 0; i < rendered.Length; i++)
            TestAssert.True(rendered[i].Annotations.All(a => a.Destination!.Value.PageIndex == i), "Nominal advances retain actual note destinations.");
    }

    public static void DocxPlainFootnoteKerningRetainsMalformedAndUnqualifiedFallbacks()
    {
        Action<XDocument, XDocument, XDocument>[] guards = [
            (xml, _, _) => xml.Descendants(Word + "rPr").First().Element(Word + "sz")!.SetAttributeValue(Word + "val", "28"),
            (xml, _, _) => xml.Descendants(Word + "rPr").First().Add(Kern("bad")),
            (xml, _, _) => xml.Descendants(Word + "rPr").First().Add(new XElement(Word + "kern")),
            (xml, _, _) => xml.Descendants(Word + "rPr").First().Add(new XElement(Word + "i")),
            (xml, _, _) => xml.Descendants(Word + "rPr").First().Add(new XElement(Word + "spacing", new XAttribute(Word + "val", "20"))),
            (xml, _, _) => xml.Descendants(Word + "t").First().Value += string.Concat(Enumerable.Repeat(" Public wrapped body AB.", 30)),
            (xml, _, _) => xml.Descendants(Word + "footnotePr").Single().Element(Word + "pos")!.SetAttributeValue(Word + "val", "beneathText"),
            (xml, notes, _) => { foreach (var rp in xml.Descendants(Word + "rPr").Concat(notes.Descendants(Word + "rPr"))) rp.Add(Kern("16")); }
        ];
        foreach (var modify in guards)
        {
            var source = ReadFixture(modify);
            TestAssert.True(!DocxPlainFootnoteKerning.TryCreateDocument(source, Layout(source), out var fallback, out _), "Malformed, wrapped, styled, positioned and enabled-threshold controls retain fallback.");
            TestAssert.True(ReferenceEquals(source, fallback), "Rejected nominal advances retain the original document.");
        }
        var malformed = ReadFixture((xml, _, _) => xml.Descendants(Word + "rPr").First().Add(new XElement(Word + "kern")));
        TestAssert.Equal(string.Empty, malformed.Paragraphs[0].Runs[0].KerningMinimumHalfPointsValue);
    }

    public static void DocxPlainFootnoteKerningKeepsPreserveAndReviewPairPositioning()
    {
        var source = ReadFixture();
        var preserve = Renderer(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectTextEmission(source);
        TestAssert.True(preserve.Lines.SelectMany(l => l.Segments).Any(s => s.GlyphAdvanceSignature.KerningUnits < 0), "Preserve output retains its measured pair positioning.");
        var review = ReadFixture((xml, _, _) => xml.Descendants(Word + "p").First().Add(new XElement(Word + "ins",
            new XAttribute(Word + "id", "19"), new XAttribute(Word + "author", "Public Author"),
            new XElement(Word + "r", new XElement(Word + "t", " Public revision AB.")))));
        var emission = Renderer(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).InspectTextEmission(review);
        TestAssert.True(emission.Lines.SelectMany(l => l.Segments).Any(s => s.GlyphAdvanceSignature.KerningUnits < 0), "The established review-text branch retains its pair positioning.");
        TestAssert.True(source.Paragraphs[0].Runs[0].PairKerningEnabled, "Rendering never mutates the source pair policy.");
    }

    private static DocxRenderer Renderer(OoxPdfDocxMarkupGeometryMode geometry) =>
        new(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup, geometry);
    private static DocxLayout Layout(DocxDocument source) =>
        new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(source, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
    private static readonly XNamespace Word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static XElement Kern(string value) => new(Word + "kern", new XAttribute(Word + "val", value));
    private static XElement RPr(string value) => new(Word + "rPr", Kern(value));
    private static DocxDocument ReadFixture(Action<XDocument, XDocument, XDocument>? modify = null)
    {
        string input = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "note-complete-reserve-3.docx"));
        var parts = new Dictionary<string, string>();
        using (ZipArchive archive = ZipFile.OpenRead(input))
            foreach (var entry in archive.Entries) { using var reader = new StreamReader(entry.Open()); parts[entry.FullName] = reader.ReadToEnd(); }
        XDocument xml = XDocument.Parse(parts["word/document.xml"]), notes = XDocument.Parse(parts["word/footnotes.xml"]), styles = XDocument.Parse(parts["word/styles.xml"]);
        xml.Descendants(Word + "t").First().Value = "Alpha AB public body before note ";
        modify?.Invoke(xml, notes, styles);
        parts["word/document.xml"] = xml.ToString(); parts["word/footnotes.xml"] = notes.ToString(); parts["word/styles.xml"] = styles.ToString();
        string temporary = TestFixtures.WriteTempPackage(".docx", parts);
        try { return DocxTests.ReadDocx(temporary, OoxPdfDocxMarkupMode.AllMarkup); }
        finally { File.Delete(temporary); }
    }
}
