using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxChartLabelDefaultsTests
{
    public static void PptxActivePlotPointsOwnDefaults()
    {
        string fourOff = string.Concat(Components.Where(c => c.Property is not ("ShowPercent" or "ShowBubbleSize"))
            .Select(c => $"<c:{c.Element} val=\"0\"/>"));
        foreach (string parent in new[] { "<c:showVal val=\"0\"/>", fourOff,
            fourOff + "<c:showPercent val=\"0\"/>", fourOff + "<c:showBubbleSize val=\"0\"/>" })
        {
            var (chart, plot) = PointDefaultsSource("lineChart", parent, "");
            TestAssert.Equal(null, chart.Plots[0].DataLabels.Overrides[0].ShowValue);
            object options = Read(chart);
            object point = Invoke("ResolveChartDataLabelOptions", options, 0);
            foreach (var component in Components)
            {
                TestAssert.True(Flag(point, component.Property), "An admitted point owns its absent-component defaults.");
                TestAssert.Equal(null, SourceFlag(chart.Plots[0].DataLabels.Overrides[0], component.Property));
            }
        }
        foreach (var component in Components)
        foreach (string value in new[] { "0", "false", "1", "", "unknown" })
        {
            string attribute = value.Length == 0 ? "" : $" val=\"{value}\"";
            var (chart, plot) = PointDefaultsSource("lineChart", "<c:showVal val=\"0\"/>", $"<c:{component.Element}{attribute}/>");
            object options = Read(chart);
            object point = Invoke("ResolveChartDataLabelOptions", options, 0);
            foreach (var other in Components)
                TestAssert.Equal(other == component ? value is "1" or "" : true, Flag(point, other.Property));
            TestAssert.Equal(value is "1" or "", SourceFlag(chart.Plots[0].DataLabels.Overrides[0], component.Property));
        }
    }

    public static void PptxDisabledCartesianPlotsSuppressPointRequestsButKeepSource()
    {
        foreach (string kind in new[] { "barChart", "lineChart", "areaChart" })
        foreach (string point in new[] { "", "<c:showVal val=\"1\"/>", PointDefaultsCustomText })
        {
            var (chart, plot) = PointDefaultsSource(kind, PointDefaultsDisabledFlags, point);
            TestAssert.Equal(1, chart.Plots[0].DataLabels.Overrides.Count);
            object resolved = Invoke("ResolveChartDataLabelOptions", Read(chart), 0);
            TestAssert.True(!Flag(resolved, "HasVisibleText") && !Flag(resolved, "ShowLegendKey"),
                "A fully disabled Cartesian plot leaves the point's text and key disabled.");
        }
    }

    public static void PptxSectorPointsRemainIndependentOfPlotDisables()
    {
        foreach (string kind in new[] { "pieChart", "doughnutChart" })
        {
            var (chart, plot) = PointDefaultsSource(kind, PointDefaultsDisabledFlags, "<c:showVal val=\"0\"/>");
            object point = Invoke("ResolveChartDataLabelOptions", Read(chart), 0);
            foreach (var component in Components)
                TestAssert.Equal(component.Property != "ShowValue", Flag(point, component.Property));
        }
    }

    public static void PptxPieUnmatchedLabelKeepsUnlabeledPlot()
    {
        var resolver = new PortableRenderingTests.PinnedResolver();
        PdfPage Render(int? pointIndex)
        {
            string point = pointIndex is int index
                ? $"<c:dLbl><c:idx val=\"{index}\"/>{PointDefaultsCustomText}{PointDefaultsDisabledFlags}</c:dLbl>"
                : "";
            var source = XDocument.Parse($$"""
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:pieChart>
                    <c:ser><c:idx val="0"/><c:order val="0"/><c:dLbls>{{point}}{{PointDefaultsDisabledFlags}}</c:dLbls>
                    <c:val><c:numLit><c:ptCount val="4"/>
                      <c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt><c:pt idx="2"><c:v>0</c:v></c:pt>
                    </c:numLit></c:val></c:ser>
                  </c:pieChart></c:plotArea></c:chart>
                </c:chartSpace>
                """);
            return RenderChartPage(source, fontResolver: resolver);
        }

        static string LabelText(PdfPage page) => string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, false,
            fontResourceFilter: resource => resource.StartsWith("CP", StringComparison.Ordinal)).Select(show => show.Text));

        PdfPage unlabeled = Render(null);
        PdfPage unmatched = Render(8);
        TestAssert.True(unlabeled.Content == unmatched.Content, "An absent point cannot change the pie's painted geometry.");
        TestAssert.Equal("", LabelText(unmatched));
        TestAssert.Equal("Pinned", LabelText(Render(0)));

        PdfPage zero = Render(2);
        PdfPage blank = Render(3);
        TestAssert.Equal("", LabelText(zero));
        TestAssert.Equal("", LabelText(blank));
        TestAssert.True(zero.Content == blank.Content && zero.Content != unlabeled.Content,
            "The declared zero and blank slots retain the existing label reservation.");
    }

    public static void PptxPolarLegendsUseSourceIndicesForUnlabeledKeys()
    {
        var renderer = new PptxRenderer(new PortableRenderingTests.PinnedResolver());
        RgbColor red = new(197, 32, 33);
        RgbColor green = new(24, 148, 72);
        RgbColor blue = new(36, 73, 213);
        foreach (string kind in new[] { "pieChart", "doughnutChart" })
        foreach (int trailingIndex in new[] { 1, 3 })
        {
            string hole = kind == "doughnutChart" ? "<c:holeSize val=\"50\"/>" : "";
            var chart = XDocument.Parse($$"""
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:{{kind}}>
                    <c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/>
                    {{Point(0, "C52021")}}{{Point(1, "189448")}}{{Point(3, "2449D5")}}
                    <c:cat><c:strLit><c:ptCount val="1"/><c:pt idx="0"><c:v>Source Anchor</c:v></c:pt></c:strLit></c:cat>
                    <c:val><c:numLit><c:ptCount val="4"/>
                      <c:pt idx="0"><c:v>16</c:v></c:pt><c:pt idx="{{trailingIndex}}"><c:v>27</c:v></c:pt>
                    </c:numLit></c:val></c:ser>{{hole}}
                  </c:{{kind}}></c:plotArea><c:legend><c:legendPos val="r"/>
                    <c:layout><c:manualLayout><c:xMode val="edge"/><c:yMode val="edge"/><c:wMode val="factor"/><c:hMode val="factor"/>
                      <c:x val=".5"/><c:y val=".25"/><c:w val=".45"/><c:h val=".35"/>
                    </c:manualLayout></c:layout><c:overlay val="1"/>
                    <c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="1000"/></a:pPr></a:p></c:txPr>
                  </c:legend></c:chart>
                </c:chartSpace>
                """);
            using var input = PptxTests.CreateSingleChartPackageStream(chart.ToString());
            using var package = OoxPackage.Open(input, default);
            var document = new PptxReader().Read(package, default);
            PdfPage page = renderer.RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
            var operations = new List<string>();
            RgbColor[] colors = PptxChartPaletteTests.ReadColoredPaints(page.Content, operations.Add).ToArray();
            var paints = colors.Zip(operations, (color, op) => (Color: color, Operation: op))
                .Where(paint => paint.Color == red || paint.Color == green || paint.Color == blue).ToArray();
            RgbColor[] expected = [red, trailingIndex == 1 ? green : blue];
            TestAssert.True(paints.Where(p => p.Operation == "f*").Select(p => p.Color).SequenceEqual(expected),
                $"{kind} must paint both authored slice colors for point {trailingIndex}.");
            TestAssert.True(paints.Where(p => p.Operation.EndsWith(" re f", StringComparison.Ordinal)).Select(p => p.Color).SequenceEqual(expected),
                $"{kind} legend keys must use the actual source points for index {trailingIndex}.");
            PortableRenderingTests.ReadPaintedSourceLabel(page, "Source Anchor");
        }

        static string Point(int index, string color) => $"""
            <c:dPt><c:idx val="{index}"/><c:spPr><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></c:spPr></c:dPt>
            """;
    }

    public static void PptxDefinedSeriesOwnPointDeclarationsWithoutLeakingPlotPoints()
    {
        string parent = "<c:dLbls><c:dLbl><c:idx val=\"0\"/><c:delete val=\"1\"/></c:dLbl>" +
            "<c:dLbl><c:idx val=\"2\"/><c:showVal val=\"0\"/></c:dLbl><c:showVal val=\"0\"/></c:dLbls>";
        foreach (string series in new[] { "<c:dLbls/>", "<c:dLbls>" + PointDefaultsDisabledFlags + "</c:dLbls>",
            "<c:dLbls><c:dLbl><c:idx val=\"1\"/><c:showVal val=\"0\"/></c:dLbl>" + PointDefaultsDisabledFlags + "</c:dLbls>" })
        {
            var (chart, plot) = Source(parent, series);
            object options = Read(chart);
            object perSeries = Invoke("ReadSceneSeriesDataLabelOptions", chart,
                chart.Plots[0]);
            object merged = Invoke("ResolveChartDataLabelOptionsForSeries", options, perSeries, 0);
            if (series == "<c:dLbls/>")
            {
                TestAssert.True(Flag(Invoke("ResolveChartDataLabelOptions", merged, 0), "ShowValue"),
                    "A defined empty series does not inherit the plot's point deletion.");
                TestAssert.True(Flag(Invoke("ResolveChartDataLabelOptions", merged, 2), "ShowValue"),
                    "The plot point's explicit false does not leak into the new series scope.");
            }
            else
            {
                TestAssert.True(!Flag(Invoke("ResolveChartDataLabelOptions", merged, 2), "HasVisibleText"),
                    "A disabled series does not inherit another plot point's default components.");
                if (series.Contains("<c:dLbl>", StringComparison.Ordinal))
                {
                    object point = Invoke("ResolveChartDataLabelOptions", merged, 1);
                    TestAssert.True(!Flag(point, "ShowValue") && Flag(point, "ShowCategoryName"),
                        "The series point keeps explicit false and its own other defaults.");
                }
            }
        }
        foreach (string point in new[] { PointDefaultsCustomText,
            "<c:spPr><a:solidFill><a:srgbClr val=\"00FF00\"/></a:solidFill></c:spPr>" })
        {
            XDocument xml = PointDefaultsRenderSource(false, point);
            var original = RenderChartPage(xml);
            if (point == PointDefaultsCustomText)
                TestAssert.Equal("Pinned", string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(original, false,
                    fontResourceFilter: resource => resource.StartsWith("CLD", StringComparison.Ordinal)).Select(s => s.Text)));
            else
                TestAssert.Equal(1, MarkerRectangles(original, new RgbColor(0, 255, 0)).Length);
            xml.Descendants(C + "ser").Single().Element(C + "cat")!.AddBeforeSelf(new XElement(C + "dLbls"));
            var page = RenderChartPage(xml);
            string text = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, false,
                fontResourceFilter: resource => resource.StartsWith("CLD", StringComparison.Ordinal)).Select(s => s.Text));
            TestAssert.Equal("Actual, W1, 2Actual, W2, 4Actual, W3, 3", text);
            TestAssert.Equal(0, MarkerRectangles(page, new RgbColor(0, 255, 0)).Length);
            TestAssert.Equal(2, MarkerRectangles(page, new RgbColor(255, 0, 0)).Length);
        }
    }

    public static void PptxPlotPointDefaultRenderingRestoresCompleteContentAndKey()
    {
        var page = RenderChartPage(PointDefaultsRenderSource(false, ""));
        string text = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, false,
            fontResourceFilter: resource => resource.StartsWith("CLD", StringComparison.Ordinal)).Select(s => s.Text));
        TestAssert.Equal("Actual, W3, 3", text);
        TestAssert.Equal(2, MarkerRectangles(page, new RgbColor(255, 0, 0)).Length);
    }

    public static void PptxDisabledPlotDefaultRenderingOmitsPointContentAndKeys()
    {
        foreach (string point in new[] { "<c:showVal val=\"1\"/>", PointDefaultsCustomText })
        {
            var page = RenderChartPage(PointDefaultsRenderSource(true, point));
            string text = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, false,
                fontResourceFilter: resource => resource.StartsWith("CLD", StringComparison.Ordinal)).Select(s => s.Text));
            TestAssert.Equal("", text);
            TestAssert.Equal(1, MarkerRectangles(page, new RgbColor(255, 0, 0)).Length);
        }
    }

    private static string PointDefaultsDisabledFlags => string.Concat(Components.Select(c => $"<c:{c.Element} val=\"0\"/>"));
    private const string PointDefaultsCustomText = "<c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Pinned</a:t></a:r></a:p></c:rich></c:tx>";

    private static (PptxSceneChart Chart, XElement Plot) PointDefaultsSource(string kind, string parentFlags, string point)
    {
        var (_, plot) = Source("<c:dLbls><c:dLbl><c:idx val=\"0\"/>" + point + "</c:dLbl>" + parentFlags + "</c:dLbls>");
        plot.Name = C + kind;
        return (PptxTests.BuildSingleChartScene(plot.Document!.ToString())!, plot);
    }

    private static XDocument PointDefaultsRenderSource(bool disabled, string point)
    {
        XDocument xml = MarkerSource([PointMarker(2, RedMarker)]);
        string flags = disabled ? PointDefaultsDisabledFlags : string.Concat(Components
            .Where(c => c.Property is not ("ShowPercent" or "ShowBubbleSize"))
            .Select(c => $"<c:{c.Element} val=\"0\"/>"));
        XElement labels = XElement.Parse("<c:dLbls xmlns:c=\"" + C + "\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
            "<c:dLbl><c:idx val=\"0\"/><c:delete val=\"1\"/></c:dLbl>" +
            "<c:dLbl><c:idx val=\"1\"/><c:delete val=\"1\"/></c:dLbl>" +
            "<c:dLbl><c:idx val=\"2\"/>" + point + "</c:dLbl>" + flags + "</c:dLbls>");
        xml.Descendants(C + "dLbls").Single().ReplaceWith(labels);
        return xml;
    }
}
