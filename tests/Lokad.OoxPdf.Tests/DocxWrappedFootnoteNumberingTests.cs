using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxWrappedFootnoteNumberingTests
{
    public static void DocxWrappedFootnoteNumbersRestartOnExactEarlyAndLaterSegments()
    {
        foreach (string family in new[] { "Arial", "Times New Roman" })
        foreach (bool later in new[] { false, true })
        foreach (bool runBreak in new[] { false, true })
        foreach (string format in new[] { "decimal", "decimalZero" })
        {
            var source = Read(later, runBreak, format, family: family);
            var first = DocxPageFootnoteNumberingTests.CreateLayout(source);
            TestAssert.Equal(2, first.Pages.Count);
            TestAssert.True(!DocxPageFootnoteNumbering.TryCreateDocument(source, first, out _, out _), "Legacy layout-policy callers retain their one-line gate.");
            TestAssert.True(!DocxPageFootnoteNumbering.HasCompleteNoteOwners(source, first), "Complete-reserve admission remains unchanged.");
            TestAssert.True(DocxPageFootnoteNumbering.TryCreateDocument(source, first, out var numbered, out var owners, allowWrappedBody: true), "The explicit numbering retry resolves all complete wrapped owners.");
            string[] expected = format == "decimalZero" ? ["09", "10", "09", "10"] : ["9", "10", "9", "10"];
            TestAssert.Equal(string.Join(',', expected), string.Join(',', numbered.Paragraphs.SelectMany(p => p.InlineReferences).Select(r => r.DisplayText)));
            TestAssert.Equal("1,1,2,2", string.Join(',', owners.Values));
            foreach (var p in source.Paragraphs.Where(p => p.InlineReferences.Count != 0))
            {
                var lines = first.Pages.SelectMany(page => page.Items).OfType<DocxTextLineLayout>().Where(l => ReferenceEquals(l.SourceParagraph, p)).ToArray();
                var marker = lines.Single(l => l.Segments.Any(s => s.SourceTextRunIndex == p.InlineReferences[0].SourceRunIndex && s.Text == p.InlineReferences[0].DisplayText));
                TestAssert.True(lines.Length > 1 && (later ? !ReferenceEquals(lines[0], marker) : ReferenceEquals(lines[0], marker)), "The fixture exercises the requested actual marker line.");
            }
            var current = DocxPageFootnoteNumberingTests.CreateLayout(numbered);
            TestAssert.True(DocxPageFootnoteNumbering.HasStablePages(numbered, current, owners, allowWrappedBody: true), "Renumbering retains each complete note's actual owner page.");
            TestAssert.Equal("105,106,107,108", string.Join(',', source.Paragraphs.SelectMany(p => p.InlineReferences).Select(r => r.DisplayText)));
            var originalAuthored = source.Paragraphs.SelectMany(p => p.Runs.Where(r => !p.InlineReferences.Any(i => i.SourceRunIndex == r.SourceRunIndex))).Select(r => r.Text);
            var numberedAuthored = numbered.Paragraphs.SelectMany(p => p.Runs.Where(r => !p.InlineReferences.Any(i => i.SourceRunIndex == r.SourceRunIndex))).Select(r => r.Text);
            TestAssert.Equal(string.Join('|', originalAuthored), string.Join('|', numberedAuthored));
        }
    }

    public static void DocxWrappedFootnoteNumbersRejectMissingAmbiguousOrShiftedSegments()
    {
        var source = Read(later: true);
        var first = DocxPageFootnoteNumberingTests.CreateLayout(source);
        var p = source.Paragraphs.First(q => q.InlineReferences.Count != 0);var r = p.InlineReferences[0];
        var line = first.Pages.SelectMany(page => page.Items).OfType<DocxTextLineLayout>().Single(l => ReferenceEquals(l.SourceParagraph, p) && l.Segments.Any(s => s.SourceTextRunIndex == r.SourceRunIndex));
        var segment = line.Segments.Single(s => s.SourceTextRunIndex == r.SourceRunIndex);
        foreach (var replacement in new[] {
            line with { Segments = line.Segments.Select(s => ReferenceEquals(s, segment) ? s with { SourceTextRunIndex = -1 } : s).ToArray() },
            line with { Segments = line.Segments.Select(s => ReferenceEquals(s, segment) ? s with { SourceTextOffsetInRun = 1 } : s).ToArray() },
            line with { Segments = line.Segments.Select(s => ReferenceEquals(s, segment) ? s with { Text = "999" } : s).ToArray() },
            line with { Segments = [.. line.Segments, segment] }
        })
        {
            var changed = first with { Pages = first.Pages.Select(page => page with { Items = page.Items.Select(i => ReferenceEquals(i, line) ? replacement : i).ToArray() }).ToArray() };
            TestAssert.True(!DocxPageFootnoteNumbering.TryCreateDocument(source, changed, out var fallback, out _, allowWrappedBody: true), "Only one exact source run, offset and display segment can own the note.");
            TestAssert.True(ReferenceEquals(source, fallback), "Rejected ownership preserves the source.");
        }
        var otherLine = first.Pages[0].Items.OfType<DocxTextLineLayout>().First(l => ReferenceEquals(l.SourceParagraph, p) && !ReferenceEquals(l, line));
        var splitPages = first.Pages.ToArray();
        splitPages[0] = splitPages[0] with { Items = splitPages[0].Items.Where(i => !ReferenceEquals(i, otherLine)).ToArray() };
        splitPages[1] = splitPages[1] with { Items = [otherLine, .. splitPages[1].Items] };
        var split = first with { Pages = splitPages };
        TestAssert.True(DocxPageFootnoteNumbering.TryCreateDocument(source, split, out _, out var owners, allowWrappedBody: true),
            "Unmarked paragraph lines on another page do not invalidate complete marker/note ownership for numbering.");
        TestAssert.True(!DocxPageFootnoteNumbering.HasStablePages(source, split, owners, allowWrappedBody: true),
            "Typography purposes retain their existing same-paragraph-page complete-frame admission.");
    }

    public static void DocxWrappedFootnoteNumbersKeepUnqualifiedSourceFrames()
    {
        Action<XDocument, XDocument, XDocument>[] exclusions = [
            (xml, _, notes) => { foreach (var fonts in xml.Descendants(W + "rFonts").Concat(notes.Descendants(W + "rFonts"))) foreach (string attribute in new[] { "ascii", "hAnsi", "eastAsia", "cs" }) fonts.SetAttributeValue(W + attribute, "Calibri"); },
            (xml, _, _) => xml.Descendants(W + "spacing").First().SetAttributeValue(W + "after", "20"),
            (xml, _, _) => xml.Descendants(W + "rPr").First().Add(new XElement(W + "i")),
            (xml, _, _) => xml.Descendants(W + "sz").First().SetAttributeValue(W + "val", "22"),
            (xml, _, _) => xml.Descendants(W + "pPr").First().Add(new XElement(W + "keepNext")),
            (xml, _, _) => xml.Descendants(W + "sectPr").Single().Add(new XElement(W + "cols", new XAttribute(W + "num", "2"))),
            (xml, _, _) => xml.Descendants(W + "footnoteReference").First().SetAttributeValue(W + "customMarkFollows", "1"),
            (_, _, notes) => notes.Root!.Elements(W + "footnote").First(n => n.Attribute(W + "type") is null).Add(new XElement(W + "p", new XElement(W + "r", new XElement(W + "t", "another complete note paragraph"))))
        ];
        foreach (var exclusion in exclusions)
        {
            var source = Read(later: true, modify: exclusion);
            TestAssert.True(!DocxPageFootnoteNumbering.TryCreateDocument(source, DocxPageFootnoteNumberingTests.CreateLayout(source), out var fallback, out _, allowWrappedBody: true), "The new wrapped frame keeps its independent exclusions.");
            TestAssert.True(ReferenceEquals(source, fallback), "Excluded source metadata stays unchanged.");
        }
    }

    public static void DocxWrappedFootnoteNumbersRetainActualTargetsAndLayoutPolicies()
    {
        var source = Read(later: true, format: "decimalZero");
        var layout = DocxPageFootnoteNumberingTests.CreateLayout(source);
        TestAssert.True(!DocxPlainFootnoteKerning.TryCreateDocument(source, layout, out _, out _), "Wrapped numbering does not extend kerning admission.");
        TestAssert.True(!DocxMinimumBodyLineLeading.TryCreateLayout(source, layout, out _), "Wrapped numbering does not extend minimum leading.");
        foreach (var geometry in new[] { OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup })
        {
            var renderer = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup, geometry);
            var pages = renderer.RenderBlankPages(source, null, CancellationToken.None).ToArray();
            TestAssert.Equal(4, pages.Sum(p => p.Annotations.Count));
            for (int i = 0; i < pages.Length; i++)
                TestAssert.True(pages[i].Annotations.All(a => a.Destination!.Value.PageIndex == i), "Actual links retain the complete physical owner page.");
            TestAssert.Equal(4, pages.SelectMany(p => p.Annotations).Select(a => a.Destination).Distinct().Count());
        }
    }

    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    internal static DocxDocument Read(bool later = false, bool runBreak = true, string format = "decimal",
        Action<XDocument, XDocument, XDocument>? modify = null, string family = "Arial") => DocxPageFootnoteNumberingTests.ReadFixture(runBreak, format, 9, (xml, settings, notes) =>
        {
            foreach (var p in xml.Root!.Element(W + "body")!.Elements(W + "p").Where(p => p.Descendants(W + "footnoteReference").Any()))
            {
                var texts = p.Descendants(W + "t").ToArray();
                string filler = string.Join(' ', Enumerable.Repeat("Public wrapped words retain authored digits 108", 9));
                if (later) texts[0].Value += filler + " "; else texts[^1].Value += " " + filler;
            }
            foreach (var rp in xml.Descendants(W + "rPr").Concat(notes.Descendants(W + "rPr")))
            {
                var fonts = rp.Element(W + "rFonts");if (fonts is null) { fonts = new XElement(W + "rFonts");rp.AddFirst(fonts); }
                foreach (string attribute in new[] { "ascii", "hAnsi", "eastAsia", "cs" }) fonts.SetAttributeValue(W + attribute, family);
            }
            modify?.Invoke(xml, settings, notes);
        });
}
