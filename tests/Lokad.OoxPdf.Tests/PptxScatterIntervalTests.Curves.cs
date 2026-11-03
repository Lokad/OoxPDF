using System.Globalization;
using System.Reflection;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxScatterIntervalTests
{
    public static void PptxScatterFixedUnitRangesMatchIndependentOfficeBounds()
    {
        foreach (RangeCase sample in RangeCases())
        foreach (bool sceneAxis in new[] { false, true })
        {
            var (actual, _) = ResolveRange(sample, sceneAxis);
            TestAssert.True(Math.Abs(actual.Min - sample.ExpectedMinimum) < 1e-9d &&
                Math.Abs(actual.Max - sample.ExpectedMaximum) < 1e-9d,
                sample.Id + ": expected " + sample.ExpectedMinimum + ".." + sample.ExpectedMaximum + ", got " + actual);
        }
    }

    public static void PptxScatterFixedUnitLabelsMarksAndGridsUseOneScale()
    {
        PdfPage page = RenderRange(RangeCases().Single(c => c.Id == "range-scatter-103-25"));
        string labels = string.Join(",", DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
            fontResourceFilter: name => name.StartsWith("CVA", StringComparison.Ordinal)).Select(label => label.Text));
        TestAssert.Equal("0,1,2,3,4,5,6,7,8,9,-22,3,28,53,78,103", labels);
        var lines = PptxColumnIntervalTests.Lines(page).Where(line => line.Color == new RgbColor(0, 0, 255) &&
            Math.Abs(line.V[1] - line.V[3]) < .001d).ToArray();
        double[] ticks = lines.Where(line => Math.Abs(line.V[0] - line.V[2]) < 20d).Select(line => line.V[1]).ToArray();
        double[] grids = lines.Where(line => Math.Abs(line.V[0] - line.V[2]) > 100d).Select(line => line.V[1]).ToArray();
        TestAssert.Equal(6, ticks.Length);
        TestAssert.Equal(6, grids.Length);
        TestAssert.True(grids.All(y => ticks.Any(t => Math.Abs(t - y) < .002d)), "Grid lines and major marks share the source interval origin.");
    }

    public static void PptxScatterCurveOrderOwnsBoundsAndHiddenLinesKeepPointBounds()
    {
        RangeCase original = RangeCases().Single(c => c.Id == "range-fine-positive103");
        RangeCase ascending = RangeCases().Single(c => c.Id == "curve-ascending");
        string[] Labels(RangeCase sample) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(RenderRange(sample), usePageCoordinates: false,
            fontResourceFilter: name => name.StartsWith("CVA", StringComparison.Ordinal)).Select(label => label.Text).ToArray();
        TestAssert.True(Labels(original).Contains("-10"), "Connected overshoot participates in automatic bounds.");
        TestAssert.True(Labels(ascending).All(text => !text.StartsWith("-", StringComparison.Ordinal)), "Reordering unchanged extrema can remove curve undershoot.");
        TestAssert.True(Labels(original with { LineHidden = true }).All(text => !text.StartsWith("-", StringComparison.Ordinal)), "Hidden connections retain point-only bounds.");
    }

    public static void PptxScatterFixedUnitFallbacksKeepSceneAuthorityAndCancellation()
    {
        RangeCase sample = RangeCases().Single(c => c.Id == "range-fine-positive103");
        XElement mismatch = XElement.Parse(Axis("2", "l", "1", "400", "500", "100", "0000FF"));
        var (owned, _) = ResolveRange(sample, true, mismatch);
        TestAssert.Equal(new PptxRenderer.ChartValueExtents(-10d, 103d), owned);
        foreach (bool logarithmic in new[] { false, true })
        {
            RangeCase source = logarithmic ? sample : sample with { Minimum = double.NaN };
            string xml = RangeChart(source, logarithmic);
            var chart = PptxTests.BuildSingleChartScene(xml)!;
            var axis = chart.Axes.Single(a => a.PositionKind == PptxSceneChartAxisPosition.Left);
            XElement element = XDocument.Parse(xml).Descendants(CurveC + "valAx").Single(a => (string?)a.Element(CurveC + "axPos")?.Attribute("val") == "l");
            var (actual, _) = ResolveRange(source, true, logarithmic: logarithmic);
            var legacy = (PptxRenderer.ChartValueExtents)CurveInvoke("ReadSceneOrXmlBubbleChartValueAxisExtents", axis, element,
                new PptxRenderer.ChartValueExtents(1d, 79d), 5d, true);
            TestAssert.Equal(legacy, actual);
        }
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => ResolveRange(sample, true, cancellationToken: cancellation.Token));
    }

    private static readonly XNamespace CurveC = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private sealed record RangeCase(string Id, string Axis, double[] X, double[] Y, double? Minimum, double? Maximum,
        double? MajorUnit, string? Smooth, bool LineHidden, double ExpectedMinimum, double ExpectedMaximum, double? ExpectedUnit = null);

    private static RangeCase[] RangeCases()
    {
        string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "scatter-fixed-intervals-office.json"));
        using var document = JsonDocument.Parse(File.ReadAllText(path));
        return JsonSerializer.Deserialize<RangeCase[]>(document.RootElement.GetProperty("Cases").GetRawText())!;
    }

    private static (PptxRenderer.ChartValueExtents Extents, object Options) ResolveRange(RangeCase sample, bool useScene,
        XElement? mismatchedXml = null, bool logarithmic = false, CancellationToken cancellationToken = default, bool bubble = false)
    {
        string xml = RangeChart(sample, logarithmic);
        var chart = PptxTests.BuildSingleChartScene(xml)!;
        XElement plot = XDocument.Parse(xml).Descendants(CurveC + "scatterChart").Single();
        object series = CurveInvoke("ReadSceneOrXmlScatterSeries", chart.Plots.Single(), plot, false, null, true);
        object options = CurveInvoke("ReadSceneOrXmlChartScatterOptions", chart.Plots.Single(), plot);
        object hidden = CurveInvoke("ReadSceneOrXmlSeriesLineHidden", chart.Plots.Single(), plot);
        object measured = CurveInvoke("MeasureScatterChartExtents", series, options, hidden, cancellationToken);
        int coordinate = sample.Axis == "X" ? 1 : 2;
        object data = measured.GetType().GetField("Item" + coordinate)!.GetValue(measured)!;
        object occupied = measured.GetType().GetField("Item" + (coordinate + 2))!.GetValue(measured)!;
        string side = sample.Axis == "X" ? "b" : "l";
        var axis = chart.Axes.Single(a => a.PositionKind == (side == "b" ? PptxSceneChartAxisPosition.Bottom : PptxSceneChartAxisPosition.Left));
        XElement element = XDocument.Parse(xml).Descendants(CurveC + "valAx").Single(a => (string?)a.Element(CurveC + "axPos")?.Attribute("val") == side);
        Type sourceType = typeof(PptxRenderer).GetNestedType("ChartAxisSource", BindingFlags.NonPublic)!;
        object source = Activator.CreateInstance(sourceType, [useScene ? axis : null, mismatchedXml ?? element])!;
        object scale = CurveInvoke("ReadChartPointValueAxisScale", PptxTheme.Empty, source, data, occupied, bubble, coordinate == 1, null);
        return ((PptxRenderer.ChartValueExtents)scale.GetType().GetField("Item1")!.GetValue(scale)!, scale.GetType().GetField("Item2")!.GetValue(scale)!);
    }

    private static object CurveInvoke(string name, params object?[] arguments)
    {
        try { return typeof(PptxRenderer).GetMethod(name, BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null, arguments)!; }
        catch (TargetInvocationException error) { throw error.InnerException ?? error; }
    }

    private static PdfPage RenderRange(RangeCase sample)
    {
        using var stream = PptxTests.CreateSingleChartPackageStream(RangeChart(sample));
        using var package = OoxPackage.Open(stream, default);
        return new PptxRenderer(new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(new PptxReader().Read(package, default), package, null, default).ToBlockingEnumerable().Single();
    }

    private static string RangeChart(RangeCase sample, bool logarithmic = false)
    {
        string Number(double value) => value.ToString("R", CultureInfo.InvariantCulture);
        string Values(string name, double[] values) => "<c:" + name + "><c:numLit>" + string.Concat(values.Select((value, index) =>
            $"<c:pt idx=\"{index}\"><c:v>{Number(value)}</c:v></c:pt>")) + "</c:numLit></c:" + name + ">";
        string AxisSource(bool horizontal)
        {
            bool selected = horizontal == (sample.Axis == "X");
            double? lower = selected ? sample.Minimum : 0d, upper = selected ? sample.Maximum : horizontal ? 9d : 103d;
            double? major = selected ? sample.MajorUnit : horizontal ? 1d : 25d;
            XElement axis = XElement.Parse(Axis(horizontal ? "1" : "2", horizontal ? "b" : "l", horizontal ? "2" : "1",
                lower is { } min ? Number(min) : "", upper is { } max ? Number(max) : "", Number(major ?? 1d), horizontal ? "00FF00" : "0000FF"));
            XElement scaling = axis.Element(CurveC + "scaling")!;
            if (lower is null) scaling.Element(CurveC + "min")!.Remove();
            if (upper is null) scaling.Element(CurveC + "max")!.Remove();
            if (major is null) axis.Element(CurveC + "majorUnit")!.Remove();
            if (selected && logarithmic) scaling.Add(new XElement(CurveC + "logBase", new XAttribute("val", "10")));
            return axis.ToString(SaveOptions.DisableFormatting);
        }
        return $$"""
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:autoTitleDeleted val="1"/><c:plotArea>
              <c:scatterChart><c:scatterStyle val="lineMarker"/><c:ser><c:idx val="0"/><c:order val="0"/>
                {{(sample.LineHidden ? "<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>" : "")}}<c:marker><c:symbol val="none"/></c:marker>
                {{Values("xVal", sample.X)}}{{Values("yVal", sample.Y)}}{{(sample.Smooth is null ? "" : $"<c:smooth val=\"{sample.Smooth}\"/>")}}
              </c:ser><c:axId val="1"/><c:axId val="2"/></c:scatterChart>{{AxisSource(true)}}{{AxisSource(false)}}
            </c:plotArea></c:chart></c:chartSpace>
            """;
    }
}
