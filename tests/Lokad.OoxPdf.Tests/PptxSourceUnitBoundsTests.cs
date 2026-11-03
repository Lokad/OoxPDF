using System.Globalization;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxSourceUnitBoundsTests
{
    private sealed record Sample(double[] Values, double Unit, double? Minimum, double Min, double Max);

    public static void PptxAuthoredUnitsOwnAutomaticCartesianBounds()
    {
        // Independent Office exports, including translated and sub-unit controls.
        Sample[] samples =
        [
            new([4, 4], 1, null, 0, 5), new([4, 4], 2, null, 0, 6),
            new([4, 4], 5, null, 0, 5), new([4, 4], 1.25, null, 0, 5),
            new([-4, -4], 2, null, -6, 0), new([0, 0], 2, null, 0, 2),
            new([-4, 4], 2, null, -6, 6), new([67, 67], 20, 7, 7, 87),
            new([4, 4], 1.25, .3, .3, 5.3), new([64, 64], 20, 7, 7, 67),
            new([164, 164], 20, 107, 107, 167), new([.04, .04], .02, null, 0, .06)
        ];
        foreach (Sample sample in samples)
        foreach (string kind in new[] { "bar", "line", "area" })
        {
            double[] labels = Labels(Render(sample, kind));
            TestAssert.True(labels.Length >= 2 && Math.Abs(labels[0] - sample.Min) < 1e-8 && Math.Abs(labels[^1] - sample.Max) < 1e-8,
                $"{kind}, unit {sample.Unit}, minimum {sample.Minimum}: {string.Join(",", labels)}");
            TestAssert.True(labels.Zip(labels.Skip(1)).All(pair => Math.Abs(pair.Second - pair.First - sample.Unit) < 1e-8),
                "The automatic limits and all painted labels share the authored interval.");
        }
    }

    public static void PptxAuthoredUnitBoundsShareTicksGridsAndBarGeometry()
    {
        foreach (bool reversed in new[] { false, true })
        foreach (bool manual in new[] { false, true })
        {
            PdfPage page = Render(new([4, 4, 4], 2, null, 0, 6), "bar", reversed, manual);
            var lines = PptxColumnIntervalTests.Lines(page);
            double[] axis = lines.Single(p => p.Color == new RgbColor(255, 0, 0) &&
                Math.Abs(p.V[0] - p.V[2]) < .001 && Math.Abs(p.V[1] - p.V[3]) > 50).V;
            double bottom = Math.Min(axis[1], axis[3]), height = Math.Abs(axis[3] - axis[1]);
            double Value(double y) => 6 * (reversed ? 1 - (y - bottom) / height : (y - bottom) / height);
            double[] Levels(RgbColor color, bool ticks) => lines
                .Where(p => p.Color == color && Math.Abs(p.V[1] - p.V[3]) < .001 &&
                    (ticks ? Math.Abs(p.V[0] - p.V[2]) is > 0 and < 20 : Math.Abs(p.V[0] - p.V[2]) > 100))
                .Select(p => Value(p.V[1])).Order().ToArray();
            void Check(double[] actual, double[] expected) => TestAssert.True(actual.Length == expected.Length &&
                actual.Zip(expected).All(pair => Math.Abs(pair.First - pair.Second) < .002),
                "Expected source-unit geometry, got " + string.Join(",", actual));
            Check(Labels(page), [0, 2, 4, 6]);
            Check(Levels(new(255, 0, 0), true), [0, 2, 4, 6]);
            Check(Levels(new(0, 0, 255), false), [2, 4, 6]);
            Check(Levels(new(255, 0, 255), false), Enumerable.Range(1, 14).Where(i => i % 5 != 0).Select(i => i * .4).ToArray());
            var rectangles = Regex.Matches(page.Content, @"([-\d.]+) ([-\d.]+) ([\d.]+) ([\d.]+) re f")
                .Select(m => m.Groups.Cast<Group>().Skip(1).Select(g => double.Parse(g.Value, CultureInfo.InvariantCulture)).ToArray());
            TestAssert.True(3 == rectangles.Count(r => Math.Abs(r[3] - height * 4 / 6) < .003),
                "All source columns use the same resolved range as the numeric labels and grids.");
        }
    }

    public static void PptxSourceUnitBoundsPreserveSceneAuthorityAndFallbacks()
    {
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        string xml = Chart(new([4, 4], 2, null, 0, 6), "bar", false, false);
        var scene = PptxTests.BuildSingleChartScene(xml)!;
        var axis = scene.Axes.Single(a => a.AxisKind == PptxSceneChartAxisKind.Value);
        XElement mismatch = XElement.Parse("<c:valAx xmlns:c='" + c + "'><c:scaling><c:max val='100'/></c:scaling><c:majorUnit val='20'/></c:valAx>");
        var data = new PptxRenderer.ChartValueExtents(0, 4);
        object Invoke(string method, params object?[] args) => typeof(PptxRenderer)
            .GetMethod(method, BindingFlags.Static | BindingFlags.NonPublic)!.Invoke(null, args)!;
        TestAssert.Equal(new PptxRenderer.ChartValueExtents(0, 6),
            (PptxRenderer.ChartValueExtents)Invoke("ReadCartesianChartValueAxisExtents", scene, axis, mismatch, data, false, true, .96d));

        foreach (string fallback in new[] { "log", "secondary", "composed", "horizontal", "percent" })
        {
            XElement source = XElement.Parse(xml);
            XElement plot = source.Descendants(c + "plotArea").Single();
            XElement valueAxis = plot.Element(c + "valAx")!;
            if (fallback == "log") valueAxis.Element(c + "scaling")!.Add(new XElement(c + "logBase", new XAttribute("val", 10)));
            if (fallback == "secondary") plot.Add(new XElement(valueAxis));
            if (fallback == "composed") plot.Add(new XElement(plot.Element(c + "barChart")!));
            if (fallback == "horizontal") plot.Element(c + "barChart")!.Element(c + "barDir")!.SetAttributeValue("val", "bar");
            var model = PptxTests.BuildSingleChartScene(source.ToString())!;
            var modelAxis = model.Axes.First(a => a.AxisKind == PptxSceneChartAxisKind.Value);
            var expected = (PptxRenderer.ChartValueExtents)Invoke("ReadPercentStackedAwareValueAxisExtents", modelAxis, valueAxis, data, fallback == "percent", true, .96d);
            var actual = (PptxRenderer.ChartValueExtents)Invoke("ReadCartesianChartValueAxisExtents", model, modelAxis, valueAxis, data, fallback == "percent", true, .96d);
            TestAssert.True(expected == actual, "Unqualified " + fallback + " scale retains its existing owner.");
        }
        foreach (bool reversed in new[] { false, true })
        foreach (bool manual in new[] { false, true })
            TestAssert.True(Labels(Render(new([4, 4], 1.25, .3, .3, 5.3), "bar", reversed, manual))
                .SequenceEqual(new[] { .3, 1.55, 2.8, 4.05, 5.3 }), "Source intervals survive reversal and authored placement.");
    }

    private static double[] Labels(PdfPage page) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
        fontResourceFilter: name => name.StartsWith("CVA", StringComparison.Ordinal))
        .Select(show => double.Parse(show.Text, CultureInfo.InvariantCulture)).ToArray();

    private static PdfPage Render(Sample sample, string kind, bool reversed = false, bool manual = false)
    {
        using var stream = PptxTests.CreateSingleChartPackageStream(Chart(sample, kind, reversed, manual));
        using var package = OoxPackage.Open(stream, default);
        return new PptxRenderer(new PortableRenderingTests.PinnedResolver())
            .RenderPagesAsync(new PptxReader().Read(package, default), package, null, default).ToBlockingEnumerable().Single();
    }

    private static string Chart(Sample sample, string kind, bool reversed, bool manual)
    {
        string Number(double value) => value.ToString("R", CultureInfo.InvariantCulture);
        string points = string.Concat(sample.Values.Select((value, i) => $"<c:pt idx=\"{i}\"><c:v>{Number(value)}</c:v></c:pt>"));
        string categories = string.Concat(sample.Values.Select((_, i) => $"<c:pt idx=\"{i}\"><c:v>Category{i}</c:v></c:pt>"));
        return $$"""
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:autoTitleDeleted val="1"/><c:plotArea>
              {{(manual ? "<c:layout><c:manualLayout><c:layoutTarget val=\"inner\"/><c:x val=\".12\"/><c:y val=\".1\"/><c:w val=\".78\"/><c:h val=\".72\"/></c:manualLayout></c:layout>" : "")}}
              <c:{{kind}}Chart>{{(kind == "bar" ? "<c:barDir val=\"col\"/>" : "")}}<c:grouping val="{{(kind == "bar" ? "clustered" : "standard")}}"/>
                <c:ser><c:idx val="0"/><c:order val="0"/><c:spPr><a:solidFill><a:srgbClr val="00AA00"/></a:solidFill></c:spPr>
                  <c:cat><c:strLit>{{categories}}</c:strLit></c:cat><c:val><c:numLit>{{points}}</c:numLit></c:val>
                </c:ser><c:axId val="1"/><c:axId val="2"/></c:{{kind}}Chart>
              <c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
              <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="{{(reversed ? "maxMin" : "minMax")}}"/>
                {{(sample.Minimum is { } min ? "<c:min val=\"" + Number(min) + "\"/>" : "")}}</c:scaling><c:axPos val="l"/>
                <c:numFmt formatCode="0.00" sourceLinked="0"/><c:crossAx val="1"/><c:majorUnit val="{{Number(sample.Unit)}}"/><c:majorTickMark val="out"/>
                <c:majorGridlines><c:spPr><a:ln><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>
                <c:minorGridlines><c:spPr><a:ln><a:solidFill><a:srgbClr val="FF00FF"/></a:solidFill></a:ln></c:spPr></c:minorGridlines>
                <c:spPr><a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></c:spPr>
                <c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="1000"/></a:pPr></a:p></c:txPr>
              </c:valAx>
            </c:plotArea></c:chart></c:chartSpace>
            """;
    }
}
