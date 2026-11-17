using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxChartLabelDefaultsTests
{
    public static void PptxDataLabelInkSurvivesFontsRichRunsAndRangeMinima()
    {
        foreach (bool mathFont in new[] { false, true })
        foreach (bool rich in new[] { false, true })
        foreach (int minimum in new[] { 0, 4 })
        {
            PdfPage page = RenderInkLabels(minimum, mathFont, rich);
            foreach (int value in new[] { minimum, minimum + 4 })
            {
                string caption = "gjpqy | " + value.ToString(CultureInfo.InvariantCulture);
                PortableRenderingTests.ReadPaintedSourceLabel(page, caption, fontPrefix: "CLD");
                if (!rich) continue;
                var glyphs = PortableRenderingTests.ReadPaintedSourceLabel(page, "qy | " + value, fontPrefix: "CLD");
                var link = page.Annotations.Single(annotation => annotation.Uri == "https://example.org/labels/" + value);
                TestAssert.True(glyphs.All(glyph => glyph.Ink.X >= link.X - .002d &&
                    glyph.Ink.Y >= link.Y - .002d && glyph.Ink.X + glyph.Ink.Width <= link.X + link.Width + .002d &&
                    glyph.Ink.Y + glyph.Ink.Height <= link.Y + link.Height + .002d),
                    "Linked rich-run ink and navigation share the admitted clip.");
            }
            TestAssert.Equal(rich ? 2 : 0, page.Annotations.Count);
        }
    }

    public static void PptxDataLabelInkPreservesManualTitleAndPageClips()
    {
        PdfPage manual = RenderInkLabels(0, mathFont: false, rich: false, manualTitle: true);
        PortableRenderingTests.ReadPaintedSourceLabel(manual, "gjpqy | 0", fontPrefix: "CLD");
        TestAssert.Throws<InvalidOperationException>(() =>
            PortableRenderingTests.ReadPaintedSourceLabel(manual, "Croppedgjpq", fontPrefix: "CT"));
        PdfPage offPage = RenderInkLabels(0, mathFont: false, rich: false, offPage: true);
        TestAssert.Throws<InvalidOperationException>(() =>
            PortableRenderingTests.ReadPaintedSourceLabel(offPage, "gjpqy | 0", fontPrefix: "CLD"));
    }

    private static PdfPage RenderInkLabels(int minimum, bool mathFont, bool rich,
        bool manualTitle = false, bool offPage = false)
    {
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var (_, plot) = Source("");
        plot.Name = C + "lineChart";
        plot.Element(C + "ser")!.Element(C + "tx")!.Element(C + "v")!.Value = "gjpqy";
        var values = plot.Descendants(C + "val").Single().Descendants(C + "pt").ToArray();
        values[0].Element(C + "v")!.Value = minimum.ToString(CultureInfo.InvariantCulture);
        values[1].Element(C + "v")!.Value = (minimum + 4).ToString(CultureInfo.InvariantCulture);
        values[2].Remove();
        plot.Descendants(C + "cat").Single().Descendants(C + "pt").Last().Remove();
        plot.Add(new XElement(C + "axId", new XAttribute("val", 1)), new XElement(C + "axId", new XAttribute("val", 2)));
        plot.Parent!.Add(XElement.Parse($$"""
            <c:catAx xmlns:c="{{C}}"><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/><c:tickLblPos val="none"/></c:catAx>
            """), XElement.Parse($$"""
            <c:valAx xmlns:c="{{C}}"><c:axId val="2"/><c:scaling><c:min val="{{minimum}}"/><c:max val="{{minimum + 4}}"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/><c:tickLblPos val="none"/></c:valAx>
            """));
        var labels = new XElement(C + "dLbls");
        if (rich)
            for (int i = 0; i < 2; i++)
                labels.Add(new XElement(C + "dLbl", new XElement(C + "idx", new XAttribute("val", i)),
                    new XElement(C + "tx", new XElement(C + "rich", new XElement(a + "bodyPr"), new XElement(a + "lstStyle"),
                        new XElement(a + "p", new XElement(a + "r", new XElement(a + "rPr", new XAttribute("sz", 2800)), new XElement(a + "t", "gjp")),
                            new XElement(a + "r", new XElement(a + "rPr", new XAttribute("sz", 800),
                                new XElement(a + "hlinkClick", new XAttribute(r + "id", "label" + i))), new XElement(a + "t", "qy | " + (minimum + i * 4))))))));
        labels.Add(new XElement(C + "dLblPos", new XAttribute("val", "b")),
            Components.Select(component => new XElement(C + component.Element,
                new XAttribute("val", component.Property is "ShowValue" or "ShowSeriesName" ? 1 : 0))),
            new XElement(C + "separator", " | "), new XElement(C + "txPr", new XElement(a + "bodyPr"), new XElement(a + "lstStyle"),
                new XElement(a + "p", new XElement(a + "pPr", new XElement(a + "defRPr", new XAttribute("sz", 1200))))));
        plot.Add(labels);
        if (manualTitle)
            plot.Parent.Parent!.AddFirst(XElement.Parse($$"""
                <c:title xmlns:c="{{C}}" xmlns:a="{{a}}"><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1600"/><a:t>Croppedgjpq</a:t></a:r></a:p></c:rich></c:tx>
                  <c:layout><c:manualLayout><c:xMode val="factor"/><c:yMode val="factor"/><c:wMode val="factor"/><c:hMode val="factor"/>
                    <c:x val="0.2"/><c:y val="0.1"/><c:w val="0.4"/><c:h val="0.001"/></c:manualLayout></c:layout></c:title>
                """));
        var parts = new Dictionary<string, byte[]>
        {
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8($$"""
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="{{a}}" xmlns:r="{{r}}" xmlns:c="{{C}}"><p:cSld><p:spTree><p:graphicFrame>
                  <p:xfrm><a:off x="914400" y="{{(offPage ? 500 : 72) * 12700}}"/><a:ext cx="7315200" cy="4572000"/></p:xfrm>
                  <a:graphic><a:graphicData uri="{{C}}"><c:chart r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>
                """),
            ["ppt/charts/_rels/chart1.xml.rels"] = TestFixtures.Utf8($$"""
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="label0" Type="{{r}}/hyperlink" Target="https://example.org/labels/{{minimum}}" TargetMode="External"/>
                  <Relationship Id="label1" Type="{{r}}/hyperlink" Target="https://example.org/labels/{{minimum + 4}}" TargetMode="External"/>
                </Relationships>
                """ )
        };
        return RenderChartPage(new XDocument(plot.Document!), parts, new PortableRenderingTests.PinnedResolver(mathFont));
    }
}
