using System.IO.Compression;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxMinimumBeforeLeadingTests
{
    public static void DocxMinimumBeforeLeadingKeepsMetricBaselineAndParagraphTop()
    {
        var source = Read(4, 4);
        var original = DocxMinimumBodyLineLeadingTests.CreatePreserveLayout(source);
        TestAssert.True(DocxMinimumBodyLineLeading.TryCreateLayout(source, original, out var adjusted),
            "Bounded before/after spacing qualifies with complete actual note owners.");
        var a = original.Pages.SelectMany(p => p.Items).OfType<DocxTextLineLayout>().First(l => l.SourceParagraph!.InlineReferences.Count != 0);
        var z = adjusted.Pages.SelectMany(p => p.Items).OfType<DocxTextLineLayout>().Single(l => ReferenceEquals(l.SourceParagraph, a.SourceParagraph));
        double excess = a.LineHeight!.Value - a.SingleLineHeight!.Value;
        Near(excess, a.BaselineY - z.BaselineY, "Before-spacing does not change the metric-derived minimum correction.");
        Near(a.BaselineY + a.FirstLineInsetPoints!.Value, z.BaselineY + z.FirstLineInsetPoints!.Value,
            "The authored paragraph top, including its spacing, remains fixed.");
        TestAssert.True(DocxPageFootnoteNumbering.HasStableCompleteNoteOwners(source, original, adjusted),
            "The correction retains complete actual marker/note pages.");
        TestAssert.Equal(a.AppliedBeforeSpacing, z.AppliedBeforeSpacing);
        TestAssert.Equal(a.PendingAfterSpacing, z.PendingAfterSpacing);
        var p = Renderer(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectTextEmission(source)
            .Lines.Single(l => l.StoryKind == "Body" && l.SourceBlockIndex == a.SourceBlockIndex);
        var w = Renderer(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).InspectTextEmission(source)
            .Lines.Single(l => l.StoryKind == "Body" && l.SourceBlockIndex == a.SourceBlockIndex);
        Near(excess, p.Segments.First().BaselineY - w.Segments.First().BaselineY, "Actual body emission uses the same corrected baseline.");
        var pm = p.Segments.Single(s => s.TextLength > 0 && s.CharacterProfile.DigitCount == s.TextLength);
        var wm = w.Segments.Single(s => s.TextLength > 0 && s.CharacterProfile.DigitCount == s.TextLength);
        Near(excess, pm.BaselineY - wm.BaselineY, "The source-owned marker moves with its body.");
        TestAssert.Equal("105", source.Paragraphs.First(q => q.InlineReferences.Count != 0).InlineReferences[0].DisplayText);
    }

    public static void DocxMinimumBeforeLeadingActualLinksOwnOnlyRemainingBeforeGap()
    {
        foreach (int predecessorAfter in new[] { 0, 4, 8 })
        foreach (int currentAfter in new[] { 0, 4, 8 })
        {
            var source = Read(0, 0, (xml, _) =>
            {
                var paragraphs = Body(xml);
                Spacing(paragraphs[0]).SetAttributeValue(Word + "after", predecessorAfter * 20);
                Spacing(paragraphs[1]).SetAttributeValue(Word + "before", "160");
                Spacing(paragraphs[1]).SetAttributeValue(Word + "after", currentAfter * 20);
            });
            var links = Renderer(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderPagesAsync(source, null, CancellationToken.None).ToBlockingEnumerable().SelectMany(p => p.Annotations).ToArray();
            TestAssert.Equal(4, links.Length);
            Near(20 + 8 - predecessorAfter + currentAfter, links[0].Height,
                "The predecessor owns its after-spacing; the real note link covers only remaining before-gap and current after-spacing.");
            var preserve = Renderer(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .RenderPagesAsync(source, null, CancellationToken.None).ToBlockingEnumerable().SelectMany(p => p.Annotations).ToArray();
            Near(20, preserve[0].Height, "Preserve retains its original minimum-slot hit area.");
            TestAssert.Equal(4, links.Select(a => a.Destination).Distinct().Count());
        }
    }

    public static void DocxMinimumBeforeLeadingInheritedFirstNoteKeepsPageTopAndRawSpacing()
    {
        var source = Read(8, 4, (xml, styles) =>
        {
            Body(xml)[0].Remove();
            var defaults = styles.Root!.Element(Word + "docDefaults")!;
            defaults.Element(Word + "pPrDefault")?.Remove();
            defaults.Add(new XElement(Word + "pPrDefault", new XElement(Word + "pPr",
                new XElement(Word + "spacing", new XAttribute(Word + "before", "160")))));
            foreach (var paragraph in Body(xml)) Spacing(paragraph).Attribute(Word + "before")!.Remove();
        });
        TestAssert.True(source.Paragraphs.All(p => p.SpacingBeforePoints == 8 && p.Spacing.BeforeValue == "160"),
            "The inherited raw before token and effective points remain coherent.");
        var pages = Renderer(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderPagesAsync(source, null, CancellationToken.None).ToBlockingEnumerable().ToArray();
        var first = pages[0].Annotations.First();
        Near(32, first.Height, "The first marked minimum slot owns its before-gap and after-spacing.");
        Near(720, first.Y + first.Height, "Its actual hit area reaches the authored page body top.");
        TestAssert.Equal(4, pages.Sum(p => p.Annotations.Count));
        for (int i = 0; i < pages.Length; i++)
            TestAssert.True(pages[i].Annotations.All(a => a.Destination!.Value.PageIndex == i), "Every complete note still targets its actual page.");
    }

    public static void DocxMinimumBeforeLeadingRetainsUnsupportedSpacingAndComposition()
    {
        DocxMinimumBodyLineLeadingTests.AssertLocalFallbacks(Read, paragraphIndex: 1);
    }

    private static readonly XNamespace Word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static XElement[] Body(XDocument xml) => xml.Root!.Element(Word + "body")!.Elements(Word + "p").ToArray();
    private static XElement Spacing(XElement paragraph) => paragraph.Element(Word + "pPr")!.Element(Word + "spacing")!;
    private static DocxRenderer Renderer(OoxPdfDocxMarkupGeometryMode mode) => new(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup, mode);
    private static void Near(double expected, double actual, string message) => TestAssert.True(Math.Abs(expected - actual) < .00001d, message);
    private static DocxDocument Read(int before, int after, Action<XDocument, XDocument>? modify = null)
    {
        string input = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "note-marked-minimum-line-20.docx"));
        var parts = new Dictionary<string, string>();
        using (var zip = ZipFile.OpenRead(input))
            foreach (var entry in zip.Entries) { using var reader = new StreamReader(entry.Open()); parts[entry.FullName] = reader.ReadToEnd(); }
        var xml = XDocument.Parse(parts["word/document.xml"]); var styles = XDocument.Parse(parts["word/styles.xml"]);
        foreach (var paragraph in Body(xml))
        {
            Spacing(paragraph).SetAttributeValue(Word + "before", before * 20);
            Spacing(paragraph).SetAttributeValue(Word + "after", after * 20);
        }
        modify?.Invoke(xml, styles);
        parts["word/document.xml"] = xml.ToString(); parts["word/styles.xml"] = styles.ToString();
        var notes = XDocument.Parse(parts["word/footnotes.xml"]);
        foreach (var paragraph in notes.Descendants(Word + "p"))
        {
            var properties = paragraph.Element(Word + "pPr");
            if (properties is null) { properties = new XElement(Word + "pPr"); paragraph.AddFirst(properties); }
            var spacing = properties.Element(Word + "spacing");
            if (spacing is null) { spacing = new XElement(Word + "spacing"); properties.Add(spacing); }
            spacing.SetAttributeValue(Word + "before", "0");
        }
        parts["word/footnotes.xml"] = notes.ToString();
        string temporary = TestFixtures.WriteTempPackage(".docx", parts);
        try { return DocxTests.ReadDocx(temporary, OoxPdfDocxMarkupMode.AllMarkup); }
        finally { File.Delete(temporary); }
    }
}
