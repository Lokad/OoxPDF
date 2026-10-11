using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxSpacedWrappedFootnoteTests
{
    public static void DocxSpacedWrappedAfterUsesSharedThresholdRules()
    {
        foreach (var raw in new[] { ("0", "0"), ("24", "24"), ("0", "24") })
        {
            var source = Read(false, 4, raw.Item1, raw.Item2);
            var first = DocxPageFootnoteNumberingTests.CreateLayout(source);
            var adjusted = source;
            if (raw.Item1 != "24")
                TestAssert.True(DocxWrappedCalibriKerning.TryCreateDocument(source, first, out adjusted, out _,
                    DocxWrappedCalibriPolicy.SourceThresholds), "Source normalization shares numeric after-spacing admission.");
            var layout = DocxPageFootnoteNumberingTests.CreateLayout(adjusted);
            TestAssert.True(DocxPageFootnoteNumbering.TryCreateDocument(adjusted, layout, out var numbered, out _, true,
                DocxWrappedCalibriPolicy.NormalizedThresholds), "Complete owners admit normalized threshold numbering.");
            TestAssert.Equal("09,10,09,10", string.Join(',', numbered.Paragraphs.SelectMany(p => p.InlineReferences).Select(r => r.DisplayText)));
            TestAssert.True(!DocxPageFootnoteNumbering.TryCreateDocument(adjusted, layout, out _, out _, true),
                "Earlier nominal-only callers retain zero-spacing qualification.");
            TestAssert.Equal("105,106,107,108", string.Join(',', source.Paragraphs.SelectMany(p => p.InlineReferences).Select(r => r.DisplayText)));
            TestAssert.True(source.Paragraphs.All(p => p.SpacingAfterPoints == 4), "Source numeric spacing remains immutable.");
        }
    }

    public static void DocxSpacedWrappedAfterOwnsOnlyFinalVisibleLine()
    {
        foreach (bool later in new[] { false, true })
        foreach (bool runBreak in new[] { false, true })
        foreach (var geometry in new[] { OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout })
        {
            var plain = Render(Read(later, 0, runBreak: runBreak), geometry);
            var spaced = Render(Read(later, 4, runBreak: runBreak), geometry);
            TestAssert.Equal(4, plain.Length); TestAssert.Equal(4, spaced.Length);
            for (int i = 0; i < spaced.Length; i++)
            {
                double expected = later && geometry == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup ? 4d : 0d;
                TestAssert.True(Math.Abs(spaced[i].Height - plain[i].Height - expected) < .000001d,
                    $"After-gap owns final line: later={later}, runBreak={runBreak}, mode={geometry}, index={i}, plain={plain[i].Height}, spaced={spaced[i].Height}, expected={expected}.");
                TestAssert.Equal(plain[i].Destination!.Value.PageIndex, spaced[i].Destination!.Value.PageIndex);
            }
        }
    }

    public static void DocxSpacedWrappedAfterRetainsUnqualifiedSpacingFallbacks()
    {
        Action<XDocument>[] exclusions = [
            xml => xml.Descendants(W + "spacing").First().SetAttributeValue(W + "before", "180"),
            xml => xml.Descendants(W + "spacing").First().SetAttributeValue(W + "after", "180"),
            xml => xml.Descendants(W + "spacing").First().SetAttributeValue(W + "afterLines", "100"),
            xml => xml.Descendants(W + "spacing").First().SetAttributeValue(W + "afterAutospacing", "1"),
            xml => xml.Descendants(W + "pPr").First().Add(new XElement(W + "contextualSpacing")),
            xml => xml.Descendants(W + "pPr").First().Add(new XElement(W + "keepLines"))
        ];
        foreach (var exclude in exclusions)
        {
            var source = Read(true, 4, modify: exclude);
            TestAssert.True(!DocxWrappedFootnoteFrame.IsEligible(source, DocxWrappedCalibriPolicy.SourceThresholds),
                "Existing raw numeric-spacing and kept-paragraph fallbacks remain explicit.");
            TestAssert.True(!DocxWrappedCalibriKerning.TryCreateDocument(source, DocxPageFootnoteNumberingTests.CreateLayout(source),
                out var fallback, out _, DocxWrappedCalibriPolicy.SourceThresholds), "Unqualified frames retain source.");
            TestAssert.True(ReferenceEquals(source, fallback), "Rejected normalization returns unchanged source.");
        }
    }

    public static void DocxSpacedWrappedAfterRejectsIncompleteMarkerOwners()
    {
        var source = Read(true, 4, "24", "24");
        var layout = DocxPageFootnoteNumberingTests.CreateLayout(source);
        TestAssert.True(DocxPageFootnoteNumbering.HasCompleteWrappedNoteOwners(source, layout), "Complete owners qualify.");
        var p = source.Paragraphs.First(); var marker = p.InlineReferences.Single();
        var line = layout.Pages.SelectMany(p => p.Items).OfType<DocxTextLineLayout>().Single(l => ReferenceEquals(l.SourceParagraph, p) &&
            l.Segments.Any(s => s.SourceTextRunIndex == marker.SourceRunIndex));
        var segment = line.Segments.Single(s => s.SourceTextRunIndex == marker.SourceRunIndex);
        var duplicate = line with { Segments = [.. line.Segments, segment] };
        var changed = layout with { Pages = layout.Pages.Select(p => p with {
            Items = p.Items.Select(i => ReferenceEquals(i, line) ? duplicate : i).ToArray() }).ToArray() };
        TestAssert.True(!DocxPageFootnoteNumbering.HasCompleteWrappedNoteOwners(source, changed),
            "The same exact owner mapper rejects duplicate markers for numbering and annotation spacing.");
    }

    private static PdfLinkAnnotation[] Render(DocxDocument document, OoxPdfDocxMarkupGeometryMode geometry) =>
        new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup, geometry)
            .RenderBlankPages(document, null, default).SelectMany(p => p.Annotations).ToArray();

    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static DocxDocument Read(bool later, int after, string first = "0", string second = "24", bool runBreak = true,
        Action<XDocument>? modify = null) => DocxWrappedFootnoteNumberingTests.Read(later, runBreak, "decimalZero", family: "Calibri",
        modify: (xml, _, _) => {
            int index = 0;
            foreach (var paragraph in xml.Root!.Element(W + "body")!.Elements(W + "p").Where(p => p.Descendants(W + "footnoteReference").Any()))
            {
                if (later) paragraph.Descendants(W + "t").Last().Value = "";
                paragraph.Element(W + "pPr")!.Element(W + "spacing")!.SetAttributeValue(W + "after", after * 20);
                string threshold = index++ % 2 == 0 ? first : second;
                foreach (var rp in paragraph.Descendants(W + "rPr")) rp.Add(new XElement(W + "kern", new XAttribute(W + "val", threshold)));
            }
            modify?.Invoke(xml);
        });
}
