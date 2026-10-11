using System.IO.Compression;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static class DocxElevenPointBodyTests
{
    public static void DocxElevenPointMinimumKeepsParagraphTopsAndAutomaticNeighbors()
    {
        foreach (bool notes in new[] { false, true })
        {
            var source = Read(notes); var original = Layout(source);
            TestAssert.True(DocxMinimumBodyLineLeading.TryCreateLayout(source, original, out var adjusted), "The complete regular 11pt frame admits minimum leading.");
            var before = Lines(original); var after = Lines(adjusted);
            TestAssert.Equal(before.Length, after.Length);
            for (int i = 0; i < before.Length; i++)
            {
                if (before[i].LineHeightSource != DocxLineHeightSource.AtLeastLineSpacing)
                {
                    TestAssert.True(ReferenceEquals(before[i], after[i]), "Automatic neighbors retain their original lines.");
                    continue;
                }
                Near(before[i].LineHeight!.Value - before[i].SingleLineHeight!.Value, before[i].BaselineY - after[i].BaselineY, "Only positive minimum excess moves the baseline.");
                Near(before[i].BaselineY + before[i].FirstLineInsetPoints!.Value, after[i].BaselineY + after[i].FirstLineInsetPoints!.Value, "Authored paragraph tops remain fixed.");
                TestAssert.Equal(before[i].AppliedBeforeSpacing, after[i].AppliedBeforeSpacing);
                TestAssert.Equal(before[i].PendingAfterSpacing, after[i].PendingAfterSpacing);
            }
            var preserve = Renderer(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectTextEmission(source);
            var word = Renderer(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).InspectTextEmission(source);
            var p = preserve.Lines.First(l => l.StoryKind == "Body" && l.SourceBlockIndex == before[0].SourceBlockIndex);
            var w = word.Lines.First(l => l.StoryKind == "Body" && l.SourceBlockIndex == before[0].SourceBlockIndex);
            Near(before[0].LineHeight!.Value - before[0].SingleLineHeight!.Value, p.Segments.First().BaselineY - w.Segments.First().BaselineY, "Requested Word emission moves the minimum baseline while Preserve retains it.");
            TestAssert.True(source.Paragraphs.SelectMany(x => x.Runs).Where(r => r.Text.Length != 0).All(r => r.FontSize == 11 && r.PairKerningEnabled), "Rendering preserves source sizes and pair policy.");
        }
    }

    public static void DocxElevenPointKerningUsesLogicalThresholdAndKeepsNoteStorySizes()
    {
        foreach (bool notes in new[] { false, true })
        {
            var source = Read(notes, (xml, _) =>
            {
                var paragraphs = Body(xml);
                string?[] thresholds = [null, "0", "28", "20"];
                for (int i = 0; i < paragraphs.Length; i++)
                foreach (var rp in paragraphs[i].Descendants(Word + "rPr"))
                    if (thresholds[i % 4] is { } threshold) rp.Add(new XElement(Word + "kern", new XAttribute(Word + "val", threshold)));
            });
            var original = Layout(source);
            TestAssert.True(DocxPlainFootnoteKerning.TryCreateDocument(source, original, out var adjusted, out var pages), "Absent, zero, above and met thresholds use the measured logical 11pt policy.");
            foreach (var run in adjusted.Paragraphs.SelectMany(p => p.Runs).Where(r => r.Text.Any(char.IsLetter)))
                TestAssert.Equal(run.KerningMinimumHalfPointsValue == "20", run.PairKerningEnabled);
            TestAssert.True(adjusted.RelatedStories.SelectMany(s => s.Paragraphs).SelectMany(p => p.Runs).Where(r => r.Text.Any(char.IsLetter)).All(r => r.FontSize == 12 && !r.PairKerningEnabled), "Note letters retain their separate 12pt size and qualified nominal advances.");
            TestAssert.Equal(notes ? 4 : 0, pages.Count);
            TestAssert.True(source.Paragraphs.SelectMany(p => p.Runs).All(r => r.PairKerningEnabled), "Source threshold metadata does not mutate fallback policy.");
            var preserve = Renderer(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectTextEmission(source);
            var word = Renderer(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).InspectTextEmission(source);
            TestAssert.True(preserve.Lines.SelectMany(l => l.Segments).Any(s => s.GlyphAdvanceSignature.KerningUnits < 0), "Preserve retains actual pair positioning.");
            TestAssert.True(word.Lines.Where(l => l.StoryKind == "Body" && l.SourceBlockIndex == 0).SelectMany(l => l.Segments).All(s => s.GlyphAdvanceSignature.KerningUnits == 0), "Actual absent-threshold Word body emission uses nominal advances.");
        }
    }

    public static void DocxElevenPointExactAndAutomaticRulesRetainBaselines()
    {
        foreach (bool notes in new[] { false, true })
        foreach (string rule in new[] { "exact", "auto" })
        {
            var source = Read(notes, (xml, _) =>
            {
                foreach (var p in Body(xml))
                {
                    var sp = Spacing(p);
                    if ((string?)sp.Attribute(Word + "lineRule") != "atLeast") continue;
                    sp.SetAttributeValue(Word + "lineRule", rule);
                    sp.SetAttributeValue(Word + "line", rule == "auto" ? "240" : "400");
                }
            });
            var original = Layout(source);
            TestAssert.True(!DocxMinimumBodyLineLeading.TryCreateLayout(source, original, out var unchanged) && ReferenceEquals(original, unchanged), "Exact and automatic counterparts do not acquire minimum excess.");
            TestAssert.True(DocxPlainFootnoteKerning.TryCreateDocument(source, original, out var adjusted, out _), "The independently qualified pitch counterpart still normalizes kerning.");
            var after = Lines(Layout(adjusted)); var before = Lines(original);
            TestAssert.Equal(before.Length, after.Length);
            for (int i = 0; i < before.Length; i++) Near(before[i].BaselineY, after[i].BaselineY, "Horizontal correction retains each exact or automatic baseline.");
        }
    }

    public static void DocxElevenPointRejectsUnknownMixedAndUnsupportedCompleteFrames()
    {
        foreach (bool notes in new[] { false, true })
        foreach (string guard in new[] { "missing-styles", "unknown-style", "ambiguous-defaults", "bold", "mixed-size", "mixed-family", "malformed-kern", "large-spacing", "auto-spacing", "keep", "wrapped", "columns", "center", "right" })
        {
            var source = Read(notes, (xml, styles) =>
            {
                var paragraphs = Body(xml); var first = paragraphs[0]; var rp = first.Descendants(Word + "rPr").First();
                if (guard == "unknown-style") first.Element(Word + "pPr")!.AddFirst(new XElement(Word + "pStyle", new XAttribute(Word + "val", "PublicMissing")));
                if (guard == "ambiguous-defaults") { AddStyle(styles, "PublicDefaultA", true); AddStyle(styles, "PublicDefaultB", true); }
                if (guard == "bold") rp.Add(new XElement(Word + "b"));
                if (guard == "mixed-size") rp.Element(Word + "sz")!.SetAttributeValue(Word + "val", "24");
                if (guard == "mixed-family") rp.Element(Word + "rFonts")!.SetAttributeValue(Word + "ascii", "Arial");
                if (guard == "malformed-kern") rp.Add(new XElement(Word + "kern", new XAttribute(Word + "val", "-1")));
                if (guard == "large-spacing") Spacing(first).SetAttributeValue(Word + "before", "240");
                if (guard == "auto-spacing") Spacing(first).SetAttributeValue(Word + "beforeAutospacing", "1");
                if (guard == "keep") first.Element(Word + "pPr")!.Add(new XElement(Word + "keepLines"));
                if (guard == "wrapped") first.Descendants(Word + "t").First().Value += string.Concat(Enumerable.Repeat(" Public wrapped AB.", 80));
                if (guard == "columns") xml.Descendants(Word + "sectPr").Single().Add(new XElement(Word + "cols", new XAttribute(Word + "num", "2")));
                if (guard is "center" or "right") first.Element(Word + "pPr")!.Add(new XElement(Word + "jc", new XAttribute(Word + "val", guard)));
            }, removeStyles: guard == "missing-styles");
            var layout = Layout(source);
            TestAssert.True(!DocxMinimumBodyLineLeading.TryCreateLayout(source, layout, out var unchanged) && ReferenceEquals(layout, unchanged), "Unsupported complete 11pt frames retain minimum fallback: " + guard);
            TestAssert.True(!DocxPlainFootnoteKerning.TryCreateDocument(source, layout, out var original, out _) && ReferenceEquals(source, original), "Unsupported complete 11pt frames retain kerning fallback: " + guard);
        }
    }

    public static void DocxElevenPointInheritedFontKeepsCompleteNoteAndLinkOwners()
    {
        foreach (bool notes in new[] { false, true })
        {
            var source = Read(notes, (xml, styles) =>
            {
                AddStyle(styles, "PublicElevenPoint");
                styles.Root!.Elements(Word + "style").Last().Add(new XElement(Word + "rPr", new XElement(Word + "sz", new XAttribute(Word + "val", "22"))));
                foreach (var p in Body(xml))
                {
                    p.Element(Word + "pPr")!.AddFirst(new XElement(Word + "pStyle", new XAttribute(Word + "val", "PublicElevenPoint")));
                    p.Descendants(Word + "sz").Remove();
                }
            });
            TestAssert.True(source.Paragraphs.SelectMany(p => p.Runs).Where(r => r.Text.Length != 0).All(r => r.FontSize == 11), "Known paragraph font inheritance resolves logical 11pt.");
            var layout = Layout(source);
            TestAssert.True(DocxMinimumBodyLineLeading.TryCreateLayout(source, layout, out var adjusted), "Inherited source size qualifies the complete frame.");
            var pages = Renderer(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(source, null, CancellationToken.None).ToArray();
            TestAssert.Equal(notes ? 4 : 0, pages.Sum(p => p.Annotations.Count));
            if (notes)
            {
                TestAssert.True(DocxPageFootnoteNumbering.HasStableCompleteNoteOwners(source, layout, adjusted), "Every complete note retains its original page.");
                for (int i = 0; i < pages.Length; i++) TestAssert.True(pages[i].Annotations.All(a => a.Destination!.Value.PageIndex == i), "Actual destinations remain on their source owner pages.");
                foreach (var annotation in pages.SelectMany(p => p.Annotations)) Near(20, annotation.Height, "The marked minimum link retains its authored slot height.");
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
