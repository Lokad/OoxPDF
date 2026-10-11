using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxNominalGlyphAdvanceTests
{
    public static void DocxNominalGlyphAdvanceKeepsPhysicalGlyphSizeAndLastNaturalWidth()
    {
        var font = OpenTypeFont.Load(TestFontBuilder.CreateTestFont());
        var embedded = PdfEmbeddedFont.Create(font, "AB".EnumerateRunes().Select(r => r.Value), CancellationToken.None);
        string original = embedded.EncodeGlyphPositioningArray("AB", 0, 11.04, true, false)!;
        string explicitDefault = embedded.EncodeGlyphPositioningArray("AB", 0, 11.04, true, false, null)!;
        TestAssert.Equal(original, explicitDefault);
        foreach (double spacing in new[] { 0d, .5d })
        {
            string payload = embedded.EncodeGlyphPositioningArray("AB", spacing, 11.04, true, false, 11)!;
            double adjustment = Regex.Matches(payload, @"(?<= )-?(?:\d+(?:\.\d*)?|\.\d+)(?= )")
                .Sum(m => double.Parse(m.Value, CultureInfo.InvariantCulture));
            // The independently authored face has A=.600em and B=.620em.
            // A advances at 11pt; B's physical final width remains at 11.04pt.
            double emitted = 1.220d * 11.04d - adjustment * 11.04d / 1000d;
            Near(.600d * 11d + .620d * 11.04d + spacing, emitted, "Actual TJ preserves physical glyph size and final width while correcting the preceding advance.");
        }
        Near(-.024d, embedded.MeasureNominalAdvanceAdjustmentPoints("AB", 11.04, 11), "Inspection accounts for the actual inter-glyph correction.");
        Near(0, embedded.MeasureNominalAdvanceAdjustmentPoints("A", 11.04, 11), "A single glyph keeps its complete natural width.");
        TestAssert.Equal(embedded.EncodeGlyphPositioningArray("A", 0, 11.04, true, false), embedded.EncodeGlyphPositioningArray("A", 0, 11.04, true, false, 11));
    }

    public static void DocxNominalGlyphAdvanceUsesQualifiedBodyPolicyAndPreservesSource()
    {
        foreach (bool notes in new[] { false, true })
        foreach (string? threshold in new string?[] { null, "0", "28", "20" })
        {
            var source = Read(notes, (xml, _) =>
            {
                if (threshold is null) return;
                foreach (var rp in Body(xml).SelectMany(p => p.Descendants(Word + "rPr"))) rp.Add(new XElement(Word + "kern", new XAttribute(Word + "val", threshold)));
            });
            var word = Renderer(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).InspectTextEmission(source);
            var body = word.Lines.Where(l => l.StoryKind == "Body").SelectMany(l => l.Segments).Where(s => !s.IsTerminalLineSpace && s.PdfFontSize == 11.04).ToArray();
            TestAssert.True(body.Length > 0, "The actual body emits its physical 11.04pt font.");
            if (threshold == "20" && !notes) TestAssert.True(body.All(s => s.AdvanceProfile.NominalAdvanceAdjustmentTotal == 0), "Excluded ordinary Calibri met frames retain accepted advances.");
            else if (threshold == "20") TestAssert.True(body.Any(s => s.AdvanceProfile.NominalAdvanceAdjustmentTotal < 0 && s.AdvanceProfile.PairAdvanceAdjustmentTotal > 0), "Qualified met note frames combine logical glyph advances with the authored pair policy.");
            else TestAssert.True(body.Any(s => s.AdvanceProfile.NominalAdvanceAdjustmentTotal < 0), "Only qualified nominal body emission applies logical advances.");
            TestAssert.True(word.Lines.Where(l => l.StoryKind != "Body").SelectMany(l => l.Segments).All(s => s.AdvanceProfile.NominalAdvanceAdjustmentTotal == 0), "Related note stories retain their accepted advance policy.");
            TestAssert.True(word.Lines.SelectMany(l => l.Segments).Where(s => s.PdfFontSize == 6.96).All(s => s.AdvanceProfile.NominalAdvanceAdjustmentTotal == 0), "Source-owned body marks keep their superscript policy.");
            var preserve = Renderer(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectTextEmission(source);
            TestAssert.True(preserve.Lines.SelectMany(l => l.Segments).All(s => s.AdvanceProfile.NominalAdvanceAdjustmentTotal == 0), "Preserve applies no logical-advance correction.");
            TestAssert.True(source.Paragraphs.SelectMany(p => p.Runs).All(r => !r.UseLogicalSizeForPdfAdvances && r.PairKerningEnabled), "Rendering retains original metadata and fallback pair policy.");
        }
    }

    public static void DocxNominalGlyphAdvanceCloneKeepsStyleAndPlannedEmissionProvenance()
    {
        var source = Read(false);var layout = Layout(source);
        TestAssert.True(DocxPlainFootnoteKerning.TryCreateDocument(source, layout, out var adjusted, out _), "Complete nominal 11pt source qualifies.");
        var run = adjusted.Paragraphs[0].Runs.First(r => r.Text.Any(char.IsLetter));
        TestAssert.True(run.UseLogicalSizeForPdfAdvances && !run.PairKerningEnabled, "The emission-only clone records nominal advance ownership.");
        var copied = DocxTextRunStyle.Empty.ApplyTo(run, "AB", 12);
        TestAssert.True(copied.UseLogicalSizeForPdfAdvances && copied.FontSize == 11, "Style application preserves the qualified clone provenance.");
        var plan = DocxTextEmissionPlanner.Create(copied, 11, 0, false, DocxTextStateCharacterSpacingSource.None);
        Near(11.04, plan.PdfFontSize, "Physical export size remains fixed.");
        Near(11, plan.NominalAdvanceFontSize!.Value, "Only advance metrics use logical size.");
        var font = OpenTypeFont.Load(TestFontBuilder.CreateTestFont());
        var embedded = PdfEmbeddedFont.Create(font, "AB".EnumerateRunes().Select(r => r.Value), CancellationToken.None);
        var profile = DocxTextEmissionPlanner.MeasureAdvanceProfile("AB", embedded, 1.220d * 11d, plan);
        Near(.600d * 11d + .620d * 11.04d, profile.PlannedEmittedAdvance, "Inspection predicts the actual corrected TJ endpoint.");
        Near(-.024d, profile.NominalAdvanceAdjustmentTotal, "The separate correction remains inspectable.");
    }

    public static void DocxNominalGlyphAdvanceDisablesOnlyOwnedElevenPointMarkerPairs()
    {
        var metSource = Read(true, (xml, _) =>
        {
            foreach (var r in Body(xml).SelectMany(p => p.Elements(Word + "r")).Where(r => r.Element(Word + "t") is not null))
                r.Element(Word + "rPr")!.Add(new XElement(Word + "kern", new XAttribute(Word + "val", "20")));
        });
        TestAssert.True(DocxPlainFootnoteKerning.TryCreateDocument(metSource, Layout(metSource), out var metAdjusted, out _), "The accepted met-body-threshold frame still qualifies its legacy policy.");
        TestAssert.True(metAdjusted.Paragraphs.SelectMany(p => p.Runs).Where(r => r.VerticalAlignment == DocxRunVerticalAlignment.Superscript).All(r => !r.PairKerningEnabled), "The diagnostic coupled met body model uses measured automatic-marker pairs.");
        foreach (string? threshold in new string?[] { null, "0", "12", "14", "16", "20", "28" })
        {
            var source = Read(true, (xml, _) =>
            {
                if (threshold is null) return;
                foreach (var r in Body(xml).SelectMany(p => p.Elements(Word + "r")).Where(r => r.Element(Word + "footnoteReference") is not null))
                    r.Element(Word + "rPr")!.Add(new XElement(Word + "kern", new XAttribute(Word + "val", threshold)));
            });
            TestAssert.True(DocxPlainFootnoteKerning.TryCreateDocument(source, Layout(source), out var adjusted, out _), "Complete source-owned 11pt note frame qualifies.");
            var marks = adjusted.Paragraphs.SelectMany(p => p.Runs.Where(r => p.InlineReferences.Any(x => x.SourceRunIndex == r.SourceRunIndex))).ToArray();
            TestAssert.Equal(4, marks.Length);
            TestAssert.True(marks.All(r => !r.PairKerningEnabled && !r.UseLogicalSizeForPdfAdvances), "Automatic superscript marks keep physical widths without pair or logical body compensation.");
            TestAssert.True(source.Paragraphs.SelectMany(p => p.Runs).All(r => r.PairKerningEnabled), "Original source metadata remains unchanged.");
            var word = Renderer(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).InspectTextEmission(source);
            TestAssert.True(word.Lines.SelectMany(l => l.Segments).Where(s => s.PdfFontSize == 6.96).All(s => s.AdvanceProfile.KerningAdjustmentTotal == 0), "Actual 6.96pt marker emission contains no pair correction.");
        }
    }

    public static void DocxNominalGlyphAdvanceRetainsPitchAndCompleteNoteNavigation()
    {
        foreach (string rule in new[] { "atLeast", "exact", "auto" })
        {
            var source = Read(true, (xml, _) =>
            {
                foreach (var p in Body(xml))
                {
                    var sp = Spacing(p);
                    if ((string?)sp.Attribute(Word + "lineRule") != "atLeast") continue;
                    sp.SetAttributeValue(Word + "lineRule", rule);sp.SetAttributeValue(Word + "line", rule == "auto" ? "240" : "400");
                }
            });
            var before = Layout(source);
            TestAssert.True(DocxPlainFootnoteKerning.TryCreateDocument(source, before, out var adjusted, out var owners), "Qualified note frame normalizes the emission clone.");
            var after = Layout(adjusted);
            TestAssert.Equal(before.Pages.Count, after.Pages.Count);
            TestAssert.True(DocxPageFootnoteNumbering.HasStablePages(adjusted, after, owners), "Every complete note retains its owner page.");
            var a = Lines(before);var z = Lines(after);TestAssert.Equal(a.Length, z.Length);
            for (int i = 0; i < a.Length; i++) { Near(a[i].BaselineY, z[i].BaselineY, "Logical glyph correction retains layout baselines.");TestAssert.Equal(a[i].LineHeight,z[i].LineHeight); }
            var pages = Renderer(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(source, null, CancellationToken.None).ToArray();
            TestAssert.Equal(4, pages.Sum(p => p.Annotations.Count));
            for (int i = 0; i < pages.Length; i++) TestAssert.True(pages[i].Annotations.All(a => a.Destination!.Value.PageIndex == i), "Actual link destinations remain on the owned page.");
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
    internal static DocxLayout Layout(DocxDocument document)
    {
        var resolver = new TestFaceFontResolver();
        var measurer = new DocxFontPlanTextMeasurer(DocxFontPlan.Create(document, resolver, CancellationToken.None), resolver.Resolve(new FontRequest("Calibri", false, false)), CancellationToken.None, resolver);
        return new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, measurer, CancellationToken.None);
    }
    internal static DocxDocument Read(bool notes, Action<XDocument, XDocument>? modify = null, bool removeStyles = false)
    {
        string input = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", notes ? "note-marked-minimum-line-20.docx" : "ordinary-minimum-line-20.docx"));
        var parts = new Dictionary<string, string>();
        using (var zip = ZipFile.OpenRead(input))
            foreach (var entry in zip.Entries) { using var reader = new StreamReader(entry.Open()); parts[entry.FullName] = reader.ReadToEnd(); }
        var xml = XDocument.Parse(parts["word/document.xml"]); var styles = XDocument.Parse(parts["word/styles.xml"]);
        foreach (var p in Body(xml))
        {
            foreach (var run in p.Elements(Word + "r"))
            {
                var rp = run.Element(Word + "rPr");
                if (rp is null) { rp = new XElement(Word + "rPr"); run.AddFirst(rp); }
                var size = rp.Element(Word + "sz");
                if (size is null) { size = new XElement(Word + "sz"); rp.Add(size); }
                size.SetAttributeValue(Word + "val", "22");
                if (run.Element(Word + "t") is { } text) text.Value += " AB.";
            }
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
