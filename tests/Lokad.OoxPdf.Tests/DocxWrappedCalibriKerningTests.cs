using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxWrappedCalibriKerningTests
{
    public static void DocxWrappedCalibriKerningQualifiesNominalRawThresholdsBeforeNumbering()
    {
        foreach (string? threshold in new string?[] { null, "0", "24", "26", "bad", "32768" })
        {
            var source = Read(threshold);
            var first = DocxPageFootnoteNumberingTests.CreateLayout(source);
            TestAssert.True(!DocxPageFootnoteNumbering.TryCreateDocument(source, first, out _, out _, allowWrappedBody: true), "Raw Calibri wrapped numbering retains fallback before the policy retry.");
            bool admitted = threshold is null or "0" or "26";
            bool actual = DocxWrappedCalibriKerning.TryCreateDocument(source, first, out var adjusted, out var owners);
            TestAssert.Equal(admitted, actual);
            if (!actual) { TestAssert.True(ReferenceEquals(source, adjusted), "Met and malformed thresholds retain exact source.");continue; }
            var body = adjusted.Paragraphs.SelectMany(p => p.Runs).Where(r => r.Text.Length != 0 && r.VerticalAlignment == DocxRunVerticalAlignment.Baseline);
            TestAssert.True(body.All(r => !r.PairKerningEnabled && !r.UseLogicalSizeForPdfAdvances && r.KerningMinimumHalfPointsValue == threshold), "The derived 12pt frame normalizes pairs and preserves raw metadata without 11pt grid compensation.");
            TestAssert.True(source.Paragraphs.SelectMany(p => p.Runs).All(r => r.PairKerningEnabled), "The reader's source flags remain intact.");
            var current = DocxPageFootnoteNumberingTests.CreateLayout(adjusted);
            TestAssert.True(DocxPageFootnoteNumbering.HasStablePages(adjusted, current, owners, allowWrappedBody: true), "All exact source owners retain complete note pages.");
            TestAssert.True(DocxPageFootnoteNumbering.TryCreateDocument(adjusted, current, out var numbered, out _, allowWrappedBody: true), "Only the qualified nominal clone admits Calibri numbering.");
            TestAssert.Equal("09,10,09,10", string.Join(',', numbered.Paragraphs.SelectMany(p => p.InlineReferences).Select(r => r.DisplayText)));
        }
    }

    public static void DocxWrappedCalibriKerningRetainsOtherFontsSizesAndLayoutGates()
    {
        Action<XDocument, XDocument, XDocument>[] exclusions = [
            (xml, _, notes) => { foreach (var fonts in xml.Descendants(W + "rFonts").Concat(notes.Descendants(W + "rFonts"))) foreach (string name in new[] { "ascii", "hAnsi", "eastAsia", "cs" }) fonts.SetAttributeValue(W + name, "Arial"); },
            (xml, _, _) => { foreach (var size in xml.Descendants(W + "sz")) size.SetAttributeValue(W + "val", "28"); },
            (xml, _, _) => xml.Descendants(W + "rPr").First().Add(new XElement(W + "b")),
            (xml, _, _) => xml.Descendants(W + "spacing").First().SetAttributeValue(W + "after", "20"),
            (xml, _, _) => xml.Descendants(W + "pPr").First().Add(new XElement(W + "keepNext")),
            (xml, _, _) => xml.Descendants(W + "footnoteReference").First().SetAttributeValue(W + "customMarkFollows", "1"),
            (_, _, notes) => notes.Root!.Elements(W + "footnote").First(n => n.Attribute(W + "type") is null).Add(new XElement(W + "p", new XElement(W + "r", new XElement(W + "t", "separate note paragraph"))))
        ];
        foreach (var exclusion in exclusions)
        {
            var source = Read("0", exclusion);
            TestAssert.True(!DocxWrappedCalibriKerning.TryCreateDocument(source, DocxPageFootnoteNumberingTests.CreateLayout(source), out var fallback, out _), "Unsupported wrapped text frames keep fallback.");
            TestAssert.True(ReferenceEquals(source, fallback), "Excluded source retains its flags and metadata.");
        }
        var plain = DocxPageFootnoteNumberingTests.ReadFixture();
        TestAssert.True(!DocxWrappedCalibriKerning.TryCreateDocument(plain, DocxPageFootnoteNumberingTests.CreateLayout(plain), out _, out _), "Existing one-line admission stays separate.");
        var wrapped = Read("0");var layout = DocxPageFootnoteNumberingTests.CreateLayout(wrapped);
        TestAssert.True(DocxWrappedCalibriKerning.TryCreateDocument(wrapped, layout, out var adjusted, out _), "Nominal frame must qualify.");
        var current = DocxPageFootnoteNumberingTests.CreateLayout(adjusted);
        TestAssert.True(!DocxPageFootnoteNumbering.HasCompleteNoteOwners(adjusted, current), "Complete-reserve owner mapping keeps its original one-line gate.");
        TestAssert.True(!DocxMinimumBodyLineLeading.TryCreateLayout(adjusted, current, out _), "Minimum leading remains outside wrapped flow.");
    }

    public static void DocxWrappedCalibriKerningRejectsShiftedSourceOwners()
    {
        var source = Read("0");var first = DocxPageFootnoteNumberingTests.CreateLayout(source);
        var p = source.Paragraphs.First(q => q.InlineReferences.Count != 0);var r = p.InlineReferences[0];
        var markerLine = first.Pages.SelectMany(page => page.Items).OfType<DocxTextLineLayout>().Single(l => ReferenceEquals(l.SourceParagraph, p) && l.Segments.Any(s => s.SourceTextRunIndex == r.SourceRunIndex));
        var changed = markerLine with { Segments = markerLine.Segments.Select(s => s.SourceTextRunIndex == r.SourceRunIndex ? s with { SourceTextOffsetInRun = 1 } : s).ToArray() };
        var layout = first with { Pages = first.Pages.Select(page => page with { Items = page.Items.Select(i => ReferenceEquals(i, markerLine) ? changed : i).ToArray() }).ToArray() };
        TestAssert.True(!DocxWrappedCalibriKerning.TryCreateDocument(source, layout, out var fallback, out _), "The pre-policy retry needs exact source offsets, not paragraph-only ownership.");
        TestAssert.True(ReferenceEquals(source, fallback), "Rejected ownership retains source.");
    }

    public static void DocxWrappedCalibriKerningRendersNumberingAndPreserveSeparately()
    {
        var source = Read("0");
        foreach (var geometry in new[] { OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout })
        {
            var renderer = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup, geometry);
            var pages = renderer.RenderBlankPages(source, null, CancellationToken.None).ToArray();TestAssert.Equal(2, pages.Length);
            string text = string.Join('|', pages.Select(DocxPageFootnoteNumberingTests.DecodeText));
            TestAssert.Contains(geometry == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup ? "09 Public destination" : "105 Public destination", text);
            TestAssert.Equal(4, pages.Sum(p => p.Annotations.Count));
            for (int i = 0; i < pages.Length; i++) TestAssert.True(pages[i].Annotations.All(a => a.Destination!.Value.PageIndex == i), "Every actual destination retains its complete source owner page.");
        }
        TestAssert.Equal("105,106,107,108", string.Join(',', source.Paragraphs.SelectMany(p => p.InlineReferences).Select(r => r.DisplayText)));
    }

    public static void DocxMultipleWrappedCalibriNormalizesOwnedMarkersWithExplicitScope()
    {
        foreach (string? threshold in new string?[] { null, "0", "26", "mixed" })
        {
            var source = ReadMultiple(threshold);
            var first = DocxPageFootnoteNumberingTests.CreateLayout(source);
            TestAssert.True(!DocxWrappedCalibriKerning.TryCreateDocument(source, first, out _, out _, DocxWrappedCalibriPolicy.SourceThresholds),
                "Legacy callers retain their single-reference frame.");
            TestAssert.True(DocxWrappedCalibriKerning.TryCreateDocument(source, first, out var adjusted, out var owners,
                DocxWrappedCalibriPolicy.SourceThresholds, DocxFootnoteOwnerScope.AllSourceReferences), "The shared complete-owner scope admits source pair normalization.");
            foreach (var run in adjusted.Paragraphs.SelectMany(p => p.Runs).Where(r => r.Text.Length != 0 && r.VerticalAlignment == DocxRunVerticalAlignment.Baseline))
            {
                TestAssert.True(DocxKerningThreshold.TryGetMinimum(run, out ushort minimum), "Raw thresholds remain parseable.");
                TestAssert.Equal(minimum is > 0 and <= 24, run.PairKerningEnabled);
            }
            var current = DocxPageFootnoteNumberingTests.CreateLayout(adjusted);
            TestAssert.True(DocxPageFootnoteNumbering.HasStablePages(adjusted, current, owners, allowWrappedBody: true,
                calibriPolicy: DocxWrappedCalibriPolicy.NormalizedThresholds, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences), "Every complete source owner retains its page.");
            TestAssert.True(DocxPageFootnoteNumbering.TryCreateDocument(adjusted, current, out var numbered, out _, allowWrappedBody: true,
                calibriPolicy: DocxWrappedCalibriPolicy.NormalizedThresholds, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences), "Only the normalized clone admits numbering.");
            TestAssert.Equal("09,10,09,10", string.Join(',', numbered.Paragraphs.SelectMany(p => p.InlineReferences).Select(r => r.DisplayText)));
            TestAssert.True(!DocxPageFootnoteNumbering.HasCompleteNoteOwners(adjusted, current), "Separator-reserve admission remains separate.");
            TestAssert.Equal("105,106,107,108", string.Join(',', source.Paragraphs.SelectMany(p => p.InlineReferences).Select(r => r.DisplayText)));
        }
    }

    public static void DocxMultipleWrappedCalibriKeepsMalformedAndIncompleteOwnerFallbacks()
    {
        foreach (string threshold in new[] { "bad", "32768", "24" })
        {
            var source = ReadMultiple(threshold);
            TestAssert.True(!DocxWrappedCalibriKerning.TryCreateDocument(source, DocxPageFootnoteNumberingTests.CreateLayout(source), out var fallback, out _,
                DocxWrappedCalibriPolicy.SourceThresholds, DocxFootnoteOwnerScope.AllSourceReferences), "Malformed and uniform-met source policies retain their prior fallback.");
            TestAssert.True(ReferenceEquals(source, fallback), "Rejected normalization preserves the reader's original model.");
        }
        var complete = ReadMultiple(null);var layout = DocxPageFootnoteNumberingTests.CreateLayout(complete);
        var pages = layout.Pages.ToArray();var notes = pages[0].PlacedRelatedStories.ToArray();
        int index = Array.FindIndex(notes, n => n.StoryLayout.Story.IsNormalStoryType);
        notes[index] = notes[index] with { ContentTopOffset = 1d };
        pages[0] = pages[0] with { PlacedRelatedStories = notes };
        TestAssert.True(!DocxWrappedCalibriKerning.TryCreateDocument(complete, layout with { Pages = pages }, out _, out _,
            DocxWrappedCalibriPolicy.SourceThresholds, DocxFootnoteOwnerScope.AllSourceReferences), "A partial note cannot acquire normalized owners.");
    }

    private static DocxDocument ReadMultiple(string? threshold) => Read(threshold == "mixed" ? null : threshold, (xml, _, _) =>
    {
        var paragraphs = xml.Root!.Element(W + "body")!.Elements(W + "p").Where(p => p.Descendants(W + "footnoteReference").Any()).ToArray();
        for (int i = 0; i < paragraphs.Length; i += 2)
        {
            paragraphs[i].Add(paragraphs[i + 1].Elements(W + "r").Select(r => new XElement(r)));paragraphs[i + 1].Remove();
        }
        if (threshold == "mixed")
        {
            int index = 0;
            foreach (var rp in xml.Root.Element(W + "body")!.Descendants(W + "rPr"))
                rp.Add(new XElement(W + "kern", new XAttribute(W + "val", index++ % 2 == 0 ? "24" : "32")));
        }
    });

    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static DocxDocument Read(string? threshold, Action<XDocument, XDocument, XDocument>? modify = null) =>
        DocxWrappedFootnoteNumberingTests.Read(later: true, format: "decimalZero", family: "Calibri", modify: (xml, settings, notes) =>
        {
            if (threshold is not null) foreach (var rp in xml.Root!.Element(W + "body")!.Descendants(W + "rPr")) rp.Add(new XElement(W + "kern", new XAttribute(W + "val", threshold)));
            modify?.Invoke(xml, settings, notes);
        });
}
