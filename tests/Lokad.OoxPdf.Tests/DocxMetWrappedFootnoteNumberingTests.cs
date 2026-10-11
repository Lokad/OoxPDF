using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxMetWrappedFootnoteNumberingTests
{
    public static void DocxMetWrappedFootnoteNumbersRequireExplicitUniformMetPolicy()
    {
        foreach (string threshold in new[] { "1", "22", "24" })
        {
            var source = Read(threshold);
            var first = DocxPageFootnoteNumberingTests.CreateLayout(source);
            TestAssert.True(!DocxPageFootnoteNumbering.TryCreateDocument(source, first, out _, out _, allowWrappedBody: true), "Legacy and Preserve callers keep the met-Calibri exclusion.");
            TestAssert.True(Create(source, first, out var numbered, out var owners), "Uniform met thresholds qualify complete wrapped owners.");
            TestAssert.Equal("09,10,09,10", string.Join(',', numbered.Paragraphs.SelectMany(p => p.InlineReferences).Select(r => r.DisplayText)));
            TestAssert.Equal("105,106,107,108", string.Join(',', source.Paragraphs.SelectMany(p => p.InlineReferences).Select(r => r.DisplayText)));
            TestAssert.True(numbered.Paragraphs.SelectMany(p => p.Runs).All(r => r.PairKerningEnabled && r.KerningMinimumHalfPointsValue == threshold && !r.UseLogicalSizeForPdfAdvances), "Numbering retains raw source thresholds, pairs and physical advance policy.");
            var current = DocxPageFootnoteNumberingTests.CreateLayout(numbered);
            TestAssert.True(DocxPageFootnoteNumbering.HasStablePages(numbered, current, owners, allowWrappedBody: true, allowMetWrappedCalibri: true), "Every complete owner retains its original page.");
            TestAssert.True(!DocxPageFootnoteNumbering.HasCompleteNoteOwners(numbered, current), "Complete reserve callers retain the one-line gate.");
        }
    }

    public static void DocxMetWrappedFootnoteNumbersRejectUnqualifiedThresholdsAndFrames()
    {
        foreach (string? threshold in new string?[] { null, "0", "26", "bad", "32768" })
        {
            var source = Read(threshold);
            TestAssert.True(!Create(source, DocxPageFootnoteNumberingTests.CreateLayout(source), out var fallback, out _), "Nominal and malformed raw frames retain their separate policy.");
            TestAssert.True(ReferenceEquals(source, fallback), "Exclusions do not mutate source.");
        }
        Action<XDocument, XDocument, XDocument>[] exclusions = [
            (xml, _, _) => xml.Descendants(W + "kern").First().SetAttributeValue(W + "val", "22"),
            (xml, _, _) => { foreach (var size in xml.Descendants(W + "sz")) size.SetAttributeValue(W + "val", "28"); },
            (xml, _, _) => xml.Descendants(W + "rPr").First().Add(new XElement(W + "b")),
            (xml, _, _) => xml.Descendants(W + "spacing").First().SetAttributeValue(W + "after", "20"),
            (xml, _, _) => xml.Descendants(W + "pPr").First().Add(new XElement(W + "keepNext")),
            (xml, _, _) => xml.Descendants(W + "footnoteReference").First().SetAttributeValue(W + "customMarkFollows", "1"),
            (_, _, notes) => notes.Root!.Elements(W + "footnote").First(n => n.Attribute(W + "type") is null).Add(new XElement(W + "p", new XElement(W + "r", new XElement(W + "t", "separate note paragraph"))))
        ];
        foreach (var exclusion in exclusions)
        {
            var source = Read("24", exclusion);
            TestAssert.True(!Create(source, DocxPageFootnoteNumberingTests.CreateLayout(source), out var fallback, out _), "Mixed thresholds and unsupported formatting/notes retain fallback.");
            TestAssert.True(ReferenceEquals(source, fallback), "Rejected frame stays unchanged.");
        }
    }

    public static void DocxMetWrappedFootnoteNumbersRejectShiftedAndDuplicateMarkerSegments()
    {
        var source = Read("24");var first = DocxPageFootnoteNumberingTests.CreateLayout(source);
        var p = source.Paragraphs.First(q => q.InlineReferences.Count != 0);var r = p.InlineReferences[0];
        var line = first.Pages.SelectMany(page => page.Items).OfType<DocxTextLineLayout>().Single(l => ReferenceEquals(l.SourceParagraph, p) && l.Segments.Any(s => s.SourceTextRunIndex == r.SourceRunIndex));
        var marker = line.Segments.Single(s => s.SourceTextRunIndex == r.SourceRunIndex);
        foreach (var segments in new[] { line.Segments.Select(s => ReferenceEquals(s, marker) ? s with { SourceTextOffsetInRun = 1 } : s).ToArray(), line.Segments.Append(marker).ToArray() })
        {
            var changed = line with { Segments = segments };
            var layout = first with { Pages = first.Pages.Select(page => page with { Items = page.Items.Select(i => ReferenceEquals(i, line) ? changed : i).ToArray() }).ToArray() };
            TestAssert.True(!Create(source, layout, out var fallback, out _), "A wrapped owner needs one exact source run/offset/display segment.");
            TestAssert.True(ReferenceEquals(source, fallback), "Rejected ownership retains source.");
        }
    }

    public static void DocxMetWrappedFootnoteNumbersRenderWordAndPreserveSeparately()
    {
        var source = Read("24");
        foreach (var geometry in new[] { OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout })
        {
            var renderer = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup, geometry);
            var pages = renderer.RenderBlankPages(source, null, CancellationToken.None).ToArray();TestAssert.Equal(2, pages.Length);
            string text = string.Join('|', pages.Select(DocxPageFootnoteNumberingTests.DecodeText));
            TestAssert.Contains(geometry == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup ? "09 Public destination" : "105 Public destination", text);
            TestAssert.Equal(4, pages.Sum(p => p.Annotations.Count));
            for (int i = 0; i < pages.Length; i++) TestAssert.True(pages[i].Annotations.All(a => a.Destination!.Value.PageIndex == i), "Actual destinations retain source-owned pages.");
        }
        TestAssert.Equal("105,106,107,108", string.Join(',', source.Paragraphs.SelectMany(p => p.InlineReferences).Select(r => r.DisplayText)));
    }

    private static bool Create(DocxDocument source, DocxLayout layout, out DocxDocument numbered, out IReadOnlyDictionary<string, int> owners) =>
        DocxPageFootnoteNumbering.TryCreateDocument(source, layout, out numbered, out owners, allowWrappedBody: true, allowMetWrappedCalibri: true);
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static DocxDocument Read(string? threshold, Action<XDocument, XDocument, XDocument>? modify = null) =>
        DocxWrappedFootnoteNumberingTests.Read(later: true, format: "decimalZero", family: "Calibri", modify: (xml, settings, notes) =>
        {
            if (threshold is not null) foreach (var rp in xml.Root!.Element(W + "body")!.Descendants(W + "rPr")) rp.Add(new XElement(W + "kern", new XAttribute(W + "val", threshold)));
            modify?.Invoke(xml, settings, notes);
        });
}
