using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxScatterIntervalTests
{
    public static void PptxSparsePointSeriesKeepAuthoredLabelsLinksAndPaint()
    {
        foreach (var (input, marks) in new[]
        {
            ("pptx-ladder-11-chart-scatter-clusters-port.pptx", 5),
            ("bubble-two-series-source.pptx", 2)
        })
        {
            const string chartPart = "ppt/charts/chart1.xml";
            var parts = ReadChartParts(input);
            var xml = XDocument.Parse(Encoding.UTF8.GetString(parts[chartPart]));
            var chart = xml.Descendants(CurveC + "chart").Single();
            chart.Element(CurveC + "legend")?.Remove();
            var plot = chart.Element(CurveC + "plotArea")!.Elements().Single(node => node.Elements(CurveC + "ser").Any());
            plot.Element(CurveC + "dLbls")?.Remove();
            XElement[] series = plot.Elements(CurveC + "ser").ToArray();
            const string firstLink = "https://example.invalid/empty-owner";
            const string nextLink = "https://example.invalid/live-owner";
            string relationshipsPart = OoxPath.GetRelationshipPartName("/" + chartPart).TrimStart('/');
            XNamespace relationships = "http://schemas.openxmlformats.org/package/2006/relationships";
            var rels = parts.TryGetValue(relationshipsPart, out var existing)
                ? XDocument.Parse(Encoding.UTF8.GetString(existing))
                : new XDocument(new XElement(relationships + "Relationships"));
            SetSeries(series[0], "EmptyOwner", "A51515", "rIdEmptyOwner", firstLink);
            SetSeries(series[1], "LiveOwner", "067649", "rIdLiveOwner", nextLink);
            parts[relationshipsPart] = TestFixtures.Utf8(rels.ToString(SaveOptions.DisableFormatting));

            PdfPage full = Render();
            AssertSeries(full, ["EmptyOwner", "LiveOwner"], [firstLink, nextLink], marks, marks);
            DropPoints(series[0]);
            PdfPage sparse = Render();
            AssertSeries(sparse, ["LiveOwner"], [nextLink], 0, marks);
            series[0].Remove();
            PdfPage literal = Render();
            TestAssert.Equal(literal.Content, sparse.Content);
            TestAssert.Equal(literal.Annotations.Single(), sparse.Annotations.Single());
            DropPoints(series[1]);
            var diagnostics = new List<OoxPdfDiagnostic>();
            PdfPage empty = Render(diagnostics.Add);
            AssertSeries(empty, [], [], 0, 0);
            TestAssert.Equal(1, diagnostics.Count(d => d.Id == "PPTX_CHART_MISSING_CACHED_DATA"));

            PdfPage Render(Action<OoxPdfDiagnostic>? sink = null)
            {
                parts[chartPart] = TestFixtures.Utf8(xml.ToString(SaveOptions.DisableFormatting));
                using var stream = TestFixtures.CreateZipPackage(parts);
                using var package = OoxPackage.Open(stream, default);
                return new PptxRenderer(new PortableRenderingTests.PinnedResolver())
                    .RenderPagesAsync(new PptxReader().Read(package, default), package, sink, default)
                    .ToBlockingEnumerable().Single();
            }

            void SetSeries(XElement owner, string text, string color, string linkId, string link)
            {
                const string relationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                const string disabled = """
                    <c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/>
                    <c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/>
                    """;
                XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
                XNamespace r = relationshipType;
                owner.Element(CurveC + "dLbls")?.Remove();
                owner.Add(XElement.Parse($"""
                    <c:dLbls xmlns:c="{CurveC}" xmlns:a="{a}" xmlns:r="{r}">
                      <c:dLbl><c:idx val="0"/><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p>
                        <a:r><a:rPr sz="1000"><a:hlinkClick r:id="{linkId}"/></a:rPr><a:t>{text}</a:t></a:r>
                      </a:p></c:rich></c:tx>{disabled}</c:dLbl>
                      {disabled}
                    </c:dLbls>
                    """));
                rels.Root!.Add(new XElement(relationships + "Relationship", new XAttribute("Id", linkId),
                    new XAttribute("Type", relationshipType + "/hyperlink"), new XAttribute("Target", link),
                    new XAttribute("TargetMode", "External")));
                var fill = new XElement(a + "solidFill", new XElement(a + "srgbClr", new XAttribute("val", color)));
                var style = owner.Element(CurveC + "spPr");
                if (style is null) owner.Add(style = new XElement(CurveC + "spPr"));
                style.Element(a + "solidFill")?.Remove();
                style.AddFirst(fill);
                if (plot.Name != CurveC + "scatterChart") return;
                owner.Element(CurveC + "marker")?.Remove();
                owner.Add(new XElement(CurveC + "marker",
                    new XElement(CurveC + "symbol", new XAttribute("val", "circle")),
                    new XElement(CurveC + "size", new XAttribute("val", "9")),
                    new XElement(CurveC + "spPr", new XElement(fill), new XElement(a + "ln", new XElement(a + "noFill")))));
            }
        }

        static void DropPoints(XElement series)
        {
            foreach (var source in series.Elements().Where(node => node.Name.LocalName is "xVal" or "yVal" or "bubbleSize"))
            {
                source.Descendants(CurveC + "pt").Remove();
                foreach (var count in source.Descendants(CurveC + "ptCount"))
                    count.SetAttributeValue("val", "0");
            }
        }

        static void AssertSeries(PdfPage page, string[] labels, string[] links, int red, int green)
        {
            TestAssert.True(BubbleText(page, "CSD").SequenceEqual(labels), "Data labels retain their authored point series.");
            TestAssert.True(page.Annotations.Select(link => link.Uri).SequenceEqual(links), "Only the painted label owners publish links.");
            RgbColor[] fills = PptxChartPaletteTests.ReadColoredPaints(page.Content).ToArray();
            TestAssert.Equal(red, fills.Count(fill => fill == new RgbColor(0xA5, 0x15, 0x15)));
            TestAssert.Equal(green, fills.Count(fill => fill == new RgbColor(0x06, 0x76, 0x49)));
        }
    }
}
