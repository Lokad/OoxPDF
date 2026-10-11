using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxExplicitPageStartSpacingTests
{
    public static void DocxExplicitPageStartConsumesOnlyTheQualifiedGap()
    {
        foreach (bool runBreak in new[] { false, true })
        foreach (var gap in new[] { (4,0), (8,8), (8,4), (4,8), (8,0) })
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

    public static void DocxExplicitPageStartKeepsPreserveAndScaledProfiles()
    {
        foreach (bool runBreak in new[] { false, true })
        {
            var source = Read(runBreak, 8, 4);
            foreach (var profile in new[] { (OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout,1d),
                (OoxPdfDocxMarkupGeometryMode.ReserveMarkupMargin,1d), (OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup,.8d) })
            {
                var line = First(Layout(source, profile.Item1, profile.Item2), source.Paragraphs[2]);
                TestAssert.True(Math.Abs(line.AppliedBeforeSpacing!.Value - 8d * profile.Item2) < .000001d,
                    "Preserve, reserve-only and scaled output keep prior spacing fallback.");
            }
        }
    }

    public static void DocxExplicitPageStartRetainsRawSpacingAndKeptFallbacks()
    {
        Action<XDocument>[] exclusions = [
            xml => xml.Descendants(W + "spacing").Last().SetAttributeValue(W + "before", "180"),
            xml => xml.Descendants(W + "spacing").Last().SetAttributeValue(W + "beforeLines", "100"),
            xml => xml.Descendants(W + "spacing").Last().SetAttributeValue(W + "beforeAutospacing", "1"),
            xml => xml.Descendants(W + "pPr").Last().Add(new XElement(W + "contextualSpacing")),
            xml => xml.Descendants(W + "pPr").Last().Add(new XElement(W + "keepLines"))
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
                    .RenderBlankPages(ReadWrapped(later,runBreak,before,0),null,default).SelectMany(p => p.Annotations).ToArray();
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
