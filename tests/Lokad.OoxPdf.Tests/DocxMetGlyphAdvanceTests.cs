using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxMetGlyphAdvanceTests
{
    public static void DocxMetGlyphAdvanceCombinesLogicalPairAndPhysicalFinalGlyph()
    {
        var font = OpenTypeFont.Load(TestFontBuilder.CreateTestFont());
        var embedded = PdfEmbeddedFont.Create(font, "AB".EnumerateRunes().Select(r => r.Value), CancellationToken.None);
        foreach (double spacing in new[] { 0d, .5d })
        {
            string payload = embedded.EncodeGlyphPositioningArray("AB", spacing, 11.04, true, true, 11)!;
            double adjustment = Regex.Matches(payload, @"(?<= )-?(?:\d+(?:\.\d*)?|\.\d+)(?= )").Sum(m => double.Parse(m.Value, CultureInfo.InvariantCulture));
            double actual = 1.220d * 11.04d - adjustment * 11.04d / 1000d;
            Near(.600d * 11d - .050d * 11d + .620d * 11.04d + spacing, actual, "Actual TJ uses logical A/pair advances and keeps the final B physical width.");
        }
        var style = DocxNominalGlyphAdvanceTests.Read(false).Paragraphs[0].Runs.First(r => r.Text.Any(char.IsLetter))
            with { Text = "AB", FontFamily = "Calibri", FontSize = 11, PairKerningEnabled = true, UseLogicalSizeForPdfAdvances = true };
        var plan = DocxTextEmissionPlanner.Create(style, 11, 0, false, DocxTextStateCharacterSpacingSource.None);
        var profile = DocxTextEmissionPlanner.MeasureAdvanceProfile("AB", embedded, 1.170d * 11d, plan);
        Near(.002d, profile.PairAdvanceAdjustmentTotal, "The logical pair adjustment is separately accounted for.");
        Near(12.8948d, profile.PlannedEmittedAdvance, "Inspection predicts actual coupled emission.");
        TestAssert.Equal(embedded.EncodeGlyphPositioningArray("A", 0, 11.04, true, true), embedded.EncodeGlyphPositioningArray("A", 0, 11.04, true, true, 11));
    }

    public static void DocxMetGlyphAdvanceKeepsRawThresholdPolicyAndRelatedStories()
    {
        foreach (string threshold in new[] { "2", "20", "22", "24" })
        {
            var source = DocxNominalGlyphAdvanceTests.Read(true, (xml, _) =>
            {
                foreach (var rp in xml.Root!.Element(Word + "body")!.Elements(Word + "p").SelectMany(p => p.Descendants(Word + "rPr")))
                    rp.Add(new XElement(Word + "kern", new XAttribute(Word + "val", threshold)));
            });
            TestAssert.True(DocxPlainFootnoteKerning.TryCreateDocument(source, DocxNominalGlyphAdvanceTests.Layout(source), out var adjusted, out _), "Complete source-owned frame qualifies the diagnostic clone.");
            var body = adjusted.Paragraphs.SelectMany(p => p.Runs).Where(r => r.Text.Length != 0 && r.VerticalAlignment == DocxRunVerticalAlignment.Baseline).ToArray();
            bool met = threshold != "24";
            TestAssert.True(body.All(r => r.UseLogicalSizeForPdfAdvances && r.PairKerningEnabled == met), "Logical advances retain the authored positive-threshold pair policy.");
            TestAssert.True(adjusted.Paragraphs.SelectMany(p => p.Runs).Where(r => r.VerticalAlignment == DocxRunVerticalAlignment.Superscript).All(r => !r.PairKerningEnabled), "Exact owned markers use the independently measured unkerned policy.");
            TestAssert.True(adjusted.RelatedStories.SelectMany(s => s.Paragraphs).SelectMany(p => p.Runs).All(r => !r.UseLogicalSizeForPdfAdvances), "Related 12pt stories receive no new glyph/pair compensation.");
            TestAssert.True(source.Paragraphs.SelectMany(p => p.Runs).All(r => !r.UseLogicalSizeForPdfAdvances && r.PairKerningEnabled && r.KerningMinimumHalfPointsValue == threshold), "Source flags and raw threshold tokens remain untouched.");
        }
    }

    public static void DocxMetGlyphAdvanceRetainsCompleteOwnerPagesAndActualTargets()
    {
        var source = DocxNominalGlyphAdvanceTests.Read(true, (xml, _) =>
        {
            foreach (var rp in xml.Root!.Element(Word + "body")!.Elements(Word + "p").SelectMany(p => p.Descendants(Word + "rPr")))
                rp.Add(new XElement(Word + "kern", new XAttribute(Word + "val", "22")));
        });
        var old = DocxNominalGlyphAdvanceTests.Layout(source);
        TestAssert.True(DocxPlainFootnoteKerning.TryCreateDocument(source, old, out var adjusted, out var owners), "Every complete note has one exact source owner.");
        var current = DocxNominalGlyphAdvanceTests.Layout(adjusted);
        TestAssert.Equal(old.Pages.Count, current.Pages.Count);
        TestAssert.True(DocxPageFootnoteNumbering.HasStablePages(adjusted, current, owners), "Each complete note retains its source owner's page.");
        var renderer = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        var pages = renderer.RenderBlankPages(source, null, CancellationToken.None).ToArray();
        TestAssert.Equal(4, pages.Sum(p => p.Annotations.Count));
        for (int i = 0; i < pages.Length; i++) TestAssert.True(pages[i].Annotations.All(a => a.Destination!.Value.PageIndex == i), "Actual destinations retain the complete owned page.");
    }

    public static void DocxMetGlyphAdvanceRetainsExcludedOrdinaryFacePolicy()
    {
        foreach (string family in new[] { "Calibri", "Arial" })
        {
            var source = DocxNominalGlyphAdvanceTests.Read(false, (xml, _) =>
            {
                foreach (var rp in xml.Root!.Element(Word + "body")!.Elements(Word + "p").SelectMany(p => p.Descendants(Word + "rPr")))
                {
                    var fonts = rp.Element(Word + "rFonts");
                    if (fonts is null) { fonts = new XElement(Word + "rFonts"); rp.AddFirst(fonts); }
                    foreach (string attribute in new[] { "ascii", "hAnsi", "eastAsia", "cs" }) fonts.SetAttributeValue(Word + attribute, family);
                    rp.Add(new XElement(Word + "kern", new XAttribute(Word + "val", "22")));
                }
            });
            TestAssert.True(DocxPlainElevenPointBody.IsEligible(source, DocxNominalGlyphAdvanceTests.Layout(source)), "The excluded met face remains inside the accepted 11pt source frame.");
            var renderer = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var word = renderer.InspectTextEmission(source);
            TestAssert.True(word.Lines.SelectMany(l => l.Segments).All(s => s.AdvanceProfile.NominalAdvanceAdjustmentTotal == 0 && s.AdvanceProfile.PairAdvanceAdjustmentTotal == 0), "Excluded ordinary faces keep physical pair/advance emission.");
            TestAssert.True(source.Paragraphs.SelectMany(p => p.Runs).All(r => !r.UseLogicalSizeForPdfAdvances && r.PairKerningEnabled), "The excluded source retains its metadata and pair policy.");
        }

        var mixed = DocxNominalGlyphAdvanceTests.Read(true, (xml, _) =>
        {
            var properties = xml.Root!.Element(Word + "body")!.Elements(Word + "p").SelectMany(p => p.Descendants(Word + "rPr")).ToArray();
            foreach (var rp in properties) rp.Add(new XElement(Word + "kern", new XAttribute(Word + "val", "22")));
            properties[0].Element(Word + "kern")!.SetAttributeValue(Word + "val", "20");
        });
        TestAssert.True(DocxPlainElevenPointBody.IsEligible(mixed, DocxNominalGlyphAdvanceTests.Layout(mixed)), "The mixed positive thresholds retain the accepted complete source frame.");
        var mixedRenderer = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        TestAssert.True(mixedRenderer.InspectTextEmission(mixed).Lines.SelectMany(l => l.Segments).All(s => s.AdvanceProfile.NominalAdvanceAdjustmentTotal == 0 && s.AdvanceProfile.PairAdvanceAdjustmentTotal == 0), "Mixed met thresholds retain accepted physical advance emission.");
        TestAssert.True(DocxPlainFootnoteKerning.TryCreateDocument(mixed, DocxNominalGlyphAdvanceTests.Layout(mixed), out var mixedAdjusted, out _), "Existing complete-note policy still applies.");
        TestAssert.True(mixedAdjusted.Paragraphs.SelectMany(p => p.Runs).Where(r => r.VerticalAlignment == DocxRunVerticalAlignment.Superscript).All(r => r.PairKerningEnabled), "Mixed thresholds retain the accepted source-owned marker pair policy.");
    }

    private static readonly XNamespace Word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static void Near(double expected, double actual, string message) => TestAssert.True(Math.Abs(expected - actual) < .00001d, message);
}
