using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxMixedWrappedFootnoteTests
{
    public static void DocxMixedWrappedThresholdsRequireDerivedPolicyBeforeNumbering()
    {
        foreach (var thresholds in new (string?, string?)[] { ("0", "24"), (null, "22"), ("26", "1"), ("1", "22") })
        {
            var source = Read(thresholds.Item1, thresholds.Item2);
            var first = DocxPageFootnoteNumberingTests.CreateLayout(source);
            TestAssert.Equal(DocxBodyKerningKind.Mixed, DocxWrappedFootnoteFrame.ClassifyKerning(source).Kind);
            TestAssert.True(!DocxPageFootnoteNumbering.TryCreateDocument(source, first, out _, out _, allowWrappedBody: true,
                calibriPolicy: DocxWrappedCalibriPolicy.NormalizedUniformMet), "The existing uniform-met qualification excludes mixed source thresholds.");
            TestAssert.True(!DocxWrappedCalibriKerning.TryCreateDocument(source, first, out _, out _), "The existing nominal-only retry excludes mixed thresholds.");
            TestAssert.True(DocxWrappedCalibriKerning.TryCreateDocument(source, first, out var adjusted, out var owners,
                DocxWrappedCalibriPolicy.SourceThresholds), "Shared source-threshold rules admit the complete regular wrapped frame.");
            var current = DocxPageFootnoteNumberingTests.CreateLayout(adjusted);
            TestAssert.True(DocxWrappedFootnoteFrame.ClassifyKerning(adjusted).PairsMatchThresholds, "Every derived baseline run follows its retained raw threshold.");
            TestAssert.True(DocxPageFootnoteNumbering.HasStablePages(adjusted, current, owners, allowWrappedBody: true,
                calibriPolicy: DocxWrappedCalibriPolicy.NormalizedThresholds), "Normalization retains exact complete owner pages.");
            TestAssert.True(DocxPageFootnoteNumbering.TryCreateDocument(adjusted, current, out var numbered, out _, allowWrappedBody: true,
                calibriPolicy: DocxWrappedCalibriPolicy.NormalizedThresholds), "Only the normalized threshold policy admits mixed numbering.");
            TestAssert.Equal("09,10,09,10", string.Join(',', numbered.Paragraphs.SelectMany(p => p.InlineReferences).Select(r => r.DisplayText)));
            TestAssert.Equal("105,106,107,108", string.Join(',', source.Paragraphs.SelectMany(p => p.InlineReferences).Select(r => r.DisplayText)));
            var original = source.Paragraphs.SelectMany(p => p.Runs).ToArray();
            var derived = adjusted.Paragraphs.SelectMany(p => p.Runs).ToArray();
            for (int i = 0; i < original.Length; i++)
            {
                TestAssert.Equal(original[i].KerningMinimumHalfPointsValue, derived[i].KerningMinimumHalfPointsValue);
                TestAssert.Equal(original[i].FontSize, derived[i].FontSize);
                TestAssert.Equal(original[i].UseLogicalSizeForPdfAdvances, derived[i].UseLogicalSizeForPdfAdvances);
                TestAssert.Equal(original[i].Text, derived[i].Text);
                TestAssert.True(original[i].PairKerningEnabled, "Source pair flags remain immutable.");
            }
            TestAssert.True(!DocxPageFootnoteNumbering.HasCompleteNoteOwners(adjusted, current), "Reserve callers retain their separate one-line fallback.");
        }
    }

    public static void DocxMixedWrappedThresholdsRetainMalformedAndFrameFallbacks()
    {
        foreach (string token in new[] { "bad", "32768", "-1" })
        {
            var source = Read("24", token);
            TestAssert.Equal(DocxBodyKerningKind.Invalid, DocxWrappedFootnoteFrame.ClassifyKerning(source).Kind);
            TestAssert.True(!Normalize(source, out var fallback), "Malformed thresholds retain fallback.");
            TestAssert.True(ReferenceEquals(source, fallback), "Rejected source remains unchanged.");
        }
        Action<XDocument, XDocument, XDocument>[] exclusions = [
            (xml, _, _) => { foreach (var size in xml.Descendants(W + "sz")) size.SetAttributeValue(W + "val", "28"); },
            (xml, _, _) => xml.Descendants(W + "rPr").First().Add(new XElement(W + "b")),
            (xml, _, _) => xml.Descendants(W + "spacing").First().SetAttributeValue(W + "after", "180"),
            (xml, _, _) => xml.Descendants(W + "pPr").First().Add(new XElement(W + "keepNext")),
            (xml, _, _) => xml.Descendants(W + "footnoteReference").First().SetAttributeValue(W + "customMarkFollows", "1"),
            (_, _, notes) => notes.Root!.Elements(W + "footnote").First(n => n.Attribute(W + "type") is null).Add(new XElement(W + "p", new XElement(W + "r", new XElement(W + "t", "separate note paragraph"))))
        ];
        foreach (var exclusion in exclusions)
        {
            var source = Read("0", "24", exclusion);
            TestAssert.True(!Normalize(source, out var fallback), "Unqualified formatting and note shapes retain fallback.");
            TestAssert.True(ReferenceEquals(source, fallback), "An excluded frame does not change source policy.");
        }
    }

    public static void DocxMixedWrappedThresholdsRejectShiftedAndDuplicateOwners()
    {
        var source = Read("0", "24");
        var first = DocxPageFootnoteNumberingTests.CreateLayout(source);
        var p = source.Paragraphs.First(q => q.InlineReferences.Count != 0);
        var r = p.InlineReferences[0];
        var line = first.Pages.SelectMany(page => page.Items).OfType<DocxTextLineLayout>().Single(l => ReferenceEquals(l.SourceParagraph, p) && l.Segments.Any(s => s.SourceTextRunIndex == r.SourceRunIndex));
        var marker = line.Segments.Single(s => s.SourceTextRunIndex == r.SourceRunIndex);
        foreach (var segments in new[] { line.Segments.Select(s => ReferenceEquals(s, marker) ? s with { SourceTextOffsetInRun = 1 } : s).ToArray(), line.Segments.Append(marker).ToArray() })
        {
            var changed = line with { Segments = segments };
            var layout = first with { Pages = first.Pages.Select(page => page with { Items = page.Items.Select(i => ReferenceEquals(i, line) ? changed : i).ToArray() }).ToArray() };
            TestAssert.True(!DocxWrappedCalibriKerning.TryCreateDocument(source, layout, out var fallback, out _, DocxWrappedCalibriPolicy.SourceThresholds), "Every mixed owner needs one exact source marker segment.");
            TestAssert.True(ReferenceEquals(source, fallback), "Ownership failures retain source.");
        }
    }

    public static void DocxMixedWrappedThresholdsRenderWordAndPreserveSeparately()
    {
        var source = Read("0", "24");
        foreach (var geometry in new[] { OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout })
        {
            var renderer = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup, geometry);
            var pages = renderer.RenderBlankPages(source, null, default).ToArray();
            TestAssert.Equal(2, pages.Length);
            string text = string.Join('|', pages.Select(DocxPageFootnoteNumberingTests.DecodeText));
            TestAssert.Contains(geometry == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup ? "09 Public destination" : "105 Public destination", text);
            TestAssert.Equal(4, pages.Sum(p => p.Annotations.Count));
            for (int i = 0; i < pages.Length; i++) TestAssert.True(pages[i].Annotations.All(a => a.Destination!.Value.PageIndex == i), "Actual destinations retain complete source owner pages.");
        }
        TestAssert.Equal("105,106,107,108", string.Join(',', source.Paragraphs.SelectMany(p => p.InlineReferences).Select(r => r.DisplayText)));
    }

    private static bool Normalize(DocxDocument source, out DocxDocument adjusted) =>
        DocxWrappedCalibriKerning.TryCreateDocument(source, DocxPageFootnoteNumberingTests.CreateLayout(source), out adjusted, out _, DocxWrappedCalibriPolicy.SourceThresholds);

    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static DocxDocument Read(string? first, string? second, Action<XDocument, XDocument, XDocument>? modify = null) =>
        DocxWrappedFootnoteNumberingTests.Read(later: true, format: "decimalZero", family: "Calibri", modify: (xml, settings, notes) =>
        {
            int index = 0;
            foreach (var p in xml.Root!.Element(W + "body")!.Elements(W + "p"))
            {
                string? threshold = index++ % 2 == 0 ? first : second;
                if (threshold is not null) foreach (var rp in p.Descendants(W + "rPr")) rp.Add(new XElement(W + "kern", new XAttribute(W + "val", threshold)));
            }
            modify?.Invoke(xml, settings, notes);
        });
}
