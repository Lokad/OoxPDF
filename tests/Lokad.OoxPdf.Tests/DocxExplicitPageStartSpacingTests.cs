using System.Xml.Linq;
using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxExplicitPageStartSpacingTests
{
    public static void DocxExplicitPageStartConsumesOnlyTheAuthoredGap()
    {
        foreach (bool runBreak in new[] { false, true })
        foreach (var gap in new[] { (4,0), (8,8), (8,4), (4,8), (8,0), (9,4), (24,18), (24,0) })
        {
            var source = Read(runBreak, gap.Item1, gap.Item2);
            var layout = Layout(source);
            var first = First(layout, source.Paragraphs[0]);
            var nextPage = First(layout, source.Paragraphs[2]);
            TestAssert.Equal((double)gap.Item1, first.AppliedBeforeSpacing!.Value);
            TestAssert.Equal(runBreak ? 0d : Math.Max(0d, gap.Item1 - gap.Item2), nextPage.AppliedBeforeSpacing!.Value);
            TestAssert.Equal(0d, nextPage.PendingAfterSpacing!.Value);
            TestAssert.Equal((double)gap.Item1, nextPage.ParagraphBeforeSpacing!.Value);
            TestAssert.Equal((double)gap.Item2, nextPage.ParagraphAfterSpacing!.Value);
            TestAssert.True(source.Paragraphs.All(p => p.SpacingBeforePoints == gap.Item1 && p.SpacingAfterPoints == gap.Item2),
                "Page-start consumption does not rewrite authored numeric spacing.");
        }
    }

    public static void DocxExplicitPageStartSharesAuthoredGapsAcrossUnscaledModes()
    {
        foreach (bool runBreak in new[] { false, true })
        foreach (var geometry in new[] { OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout,
            OoxPdfDocxMarkupGeometryMode.ReserveMarkupMargin, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup })
        {
            var source = Read(runBreak, 8, 4);
            var line = First(Layout(source, geometry), source.Paragraphs[2]);
            TestAssert.Equal(runBreak ? 0d : 4d, line.AppliedBeforeSpacing!.Value);
        }
    }

    public static void DocxExplicitPageStartKeepsScaledProfiles()
    {
        foreach (bool runBreak in new[] { false, true })
        {
            var source = Read(runBreak, 8, 4);
            var line = First(Layout(source, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, .8d), source.Paragraphs[2]);
            TestAssert.True(Math.Abs(line.AppliedBeforeSpacing!.Value - 6.4d) < .000001d,
                "Scaled print keeps its existing spacing fallback until independently qualified.");
        }
    }

    public static void DocxExplicitPageStartRetainsNonDistanceSpacingFallbacks()
    {
        Action<XDocument>[] exclusions = [
            xml => xml.Descendants(W + "spacing").Last().SetAttributeValue(W + "beforeLines", "100"),
            xml => xml.Descendants(W + "spacing").Last().SetAttributeValue(W + "beforeAutospacing", "1"),
            xml => xml.Descendants(W + "pPr").Last().Add(new XElement(W + "contextualSpacing"))
        ];
        foreach (var exclude in exclusions)
        {
            var source = Read(true, 8, 4, xml => {
                // Modify the paragraph immediately after the break, preserving
                // the control's other paragraphs and source marker identities.
                var p = xml.Root!.Element(W + "body")!.Elements(W + "p").Where(p => p.Descendants(W + "footnoteReference").Any()).ElementAt(2);
                var isolated = new XDocument(new XElement("test",new XElement(p)));
                exclude(isolated);p.ReplaceWith(isolated.Root!.Element(W + "p"));
            });
            var line = First(Layout(source), source.Paragraphs[2]);
            TestAssert.Equal(source.Paragraphs[2].SpacingBeforePoints, line.AppliedBeforeSpacing!.Value);
        }
    }

    public static void DocxExplicitPageStartRetainsNaturalAndDocumentStartSpacing()
    {
        var source = Read(false, 8, 4);
        source = source with { BodyElements = source.BodyElements.Where(e => e is not DocxPageBreakElement).ToArray() };
        var layout = Layout(source);
        foreach (var paragraph in source.Paragraphs)
            TestAssert.Equal(8d, First(layout, paragraph).AppliedBeforeSpacing!.Value);
        TestAssert.True(layout.Pages.SelectMany(p => p.Items).OfType<DocxTextLineLayout>().All(l => l.SourceParagraph is null ||
            source.Paragraphs.Any(p => ReferenceEquals(p, l.SourceParagraph))), "Natural flow retains source paragraph ownership.");
    }

    public static void DocxExplicitPageStartSharesNumericBeforeThresholdPolicy()
    {
        var source = ReadWrapped(false, true, 8, 8);
        var initial = Layout(source);
        TestAssert.True(DocxWrappedCalibriKerning.TryCreateDocument(source, initial, out var adjusted, out _,
            DocxWrappedCalibriPolicy.SourceThresholds), "The existing shared numeric qualifier admits both spacing contributions.");
        var layout = Layout(adjusted);
        TestAssert.True(DocxPageFootnoteNumbering.TryCreateDocument(adjusted, layout, out var numbered, out _, true,
            DocxWrappedCalibriPolicy.NormalizedThresholds), "Complete source owners govern numbering after page-start consumption.");
        TestAssert.Equal("09,10,09,10",string.Join(',',numbered.Paragraphs.SelectMany(p => p.InlineReferences).Select(r => r.DisplayText)));
        TestAssert.True(!DocxPageFootnoteNumbering.TryCreateDocument(adjusted, layout, out _, out _, true),
            "Original nominal-only purpose still requires zero numeric spacing.");
        TestAssert.Equal("105,106,107,108",string.Join(',',source.Paragraphs.SelectMany(p => p.InlineReferences).Select(r => r.DisplayText)));
    }

    public static void DocxExplicitPageStartBeforeHitAreaBelongsToFirstVisibleLine()
    {
        foreach (bool later in new[] { false, true })
        foreach (bool runBreak in new[] { false, true })
        foreach (var geometry in new[] { OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout })
        {
            Lokad.OoxPdf.Pdf.PdfLinkAnnotation[] Links(int before) =>
                new DocxRenderer(new TestFaceFontResolver(),OoxPdfDocxMarkupMode.AllMarkup,geometry)
                    .RenderPagesAsync(ReadWrapped(later,runBreak,before,0),null,default).ToBlockingEnumerable().SelectMany(p => p.Annotations).ToArray();
            var plain=Links(0);var spaced=Links(4);TestAssert.Equal(4,plain.Length);TestAssert.Equal(4,spaced.Length);
            for(int i=0;i<4;i++)
            {
                double gap=geometry==OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup && !later && !(runBreak && i==2) ? 4d : 0d;
                TestAssert.True(Math.Abs(spaced[i].Height-plain[i].Height-gap)<.000001d,
                    $"First-line before ownership: later={later}, run={runBreak}, mode={geometry}, owner={i}, expected={gap}.");
                TestAssert.Equal(plain[i].Destination!.Value.PageIndex,spaced[i].Destination!.Value.PageIndex);
            }
        }
    }

    public static void DocxExplicitPageStartMatchesWordAcrossModesSizesAndSurroundingContent()
    {
        string directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "ExplicitPageStart"));
        using var golden = JsonDocument.Parse(File.ReadAllText(Path.Combine(directory, "office.json")));
        var cases = golden.RootElement.GetProperty("Cases").EnumerateArray()
            .ToDictionary(item => item.GetProperty("Id").GetString()!);
        foreach (var reference in golden.RootElement.GetProperty("References").EnumerateArray())
        {
            string input = Path.Combine(directory, reference.GetProperty("Input").GetString()!);
            TestAssert.Equal(reference.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            bool review = reference.GetProperty("Mode").GetString() == "all-word-compatible";
            var source = DocxTests.ReadDocx(input, review ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final);
            var layout = Layout(source, review ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            TestAssert.Equal(reference.GetProperty("Pages").GetInt32(), layout.Pages.Count);
            var lines = layout.Pages.SelectMany((page, index) => page.Items.OfType<DocxTextLineLayout>()
                .Select(line => (Line: line, Page: index + 1))).ToArray();
            var paragraphs = source.Paragraphs.ToDictionary(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text)).Trim());
            foreach (var sample in reference.GetProperty("Samples").EnumerateArray())
            {
                string id = sample.GetProperty("Id").GetString()!;
                var spec = cases[id];
                var first = lines.Single(item => ReferenceEquals(item.Line.SourceParagraph, paragraphs[id + " FIRST"]) && item.Line.SourceLineIndex == 0);
                var next = lines.Single(item => ReferenceEquals(item.Line.SourceParagraph, paragraphs[id + " NEXT"]) && item.Line.SourceLineIndex == 0);
                TestAssert.Equal(sample.GetProperty("First").GetProperty("Page").GetInt32(), first.Page);
                TestAssert.Equal(sample.GetProperty("Next").GetProperty("Page").GetInt32(), next.Page);
                TestAssert.Equal(spec.GetProperty("ExpectedConsumedBeforeSpacing").GetDouble(), next.Line.AppliedBeforeSpacing!.Value);
                TestAssert.Equal(spec.GetProperty("Before").GetDouble(), next.Line.ParagraphBeforeSpacing!.Value);
                TestAssert.Equal(spec.GetProperty("After").GetDouble(), first.Line.SourceParagraph!.SpacingAfterPoints);
                TestAssert.Equal(0d, next.Line.PendingAfterSpacing!.Value);
            }
        }
    }

    public static void DocxExplicitPageStartPreservesComposedContent()
    {
        string input = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "ExplicitPageStart", "composition.docx"));
        var document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
        TestAssert.Equal(2, document.Paragraphs.Sum(paragraph => paragraph.Equations.Count));
        var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true),
            OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var pages = renderer.RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();
        TestAssert.Equal(8, pages.Length);
        int pictures = pages.Sum(page => System.Text.RegularExpressions.Regex.Matches(page.Content, @"/(?<name>[^\s/]+) Do")
            .Count(match => page.Images.Any(image => image.ResourceName == match.Groups["name"].Value)));
        TestAssert.Equal(2, pictures);
        TestAssert.True(pages.Take(2).All(page => page.Content.Contains(" re f", StringComparison.Ordinal)),
            "Each composed fraction retains its visible bar on its source page.");
        string text = string.Concat(pages.SelectMany(DocxCommentsTests.ReadEmbeddedGlyphTextShows).Select(show => show.Text));
        foreach (string label in new[] { "MATH FIRST", "MATH NEXT", "PICTURE FIRST", "PICTURE NEXT", "TABLECELL", "STATIC NEXT", "Public static header" })
            TestAssert.True(text.Contains(label, StringComparison.Ordinal), "Supported content survives explicit page starts: " + label);
    }

    private static DocxDocument ReadWrapped(bool later,bool runBreak,int before,int after) =>
        DocxWrappedFootnoteNumberingTests.Read(later,runBreak,"decimalZero",family:"Calibri",modify:(xml,_,_) => {
            foreach(var p in xml.Root!.Element(W+"body")!.Elements(W+"p").Where(p => p.Descendants(W+"footnoteReference").Any()))
            {
                if(later)p.Descendants(W+"t").Last().Value="";
                var spacing=p.Element(W+"pPr")!.Element(W+"spacing")!;
                spacing.SetAttributeValue(W+"before",before*20);spacing.SetAttributeValue(W+"after",after*20);
            }
        });

    private static DocxTextLineLayout First(DocxLayout layout, DocxParagraph paragraph) =>
        layout.Pages.SelectMany(p => p.Items).OfType<DocxTextLineLayout>().First(l => ReferenceEquals(l.SourceParagraph, paragraph) &&
            l.Segments.Any(s => s.Role == DocxTextSegmentRole.Text));
    private static DocxLayout Layout(DocxDocument source,
        OoxPdfDocxMarkupGeometryMode geometry = OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, double scale = 1d) =>
        new DocxLayoutEngine(geometry,scale).Create(source,new DocxTests.FamilyWidthTextMeasurer(),default);
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static DocxDocument Read(bool runBreak, int before, int after, Action<XDocument>? modify = null) =>
        DocxPageFootnoteNumberingTests.ReadFixture(runBreak, modify:(xml,_,_) => {
            foreach (var spacing in xml.Descendants(W + "spacing"))
            { spacing.SetAttributeValue(W + "before",before * 20); spacing.SetAttributeValue(W + "after",after * 20); }
            modify?.Invoke(xml);
        });
}
