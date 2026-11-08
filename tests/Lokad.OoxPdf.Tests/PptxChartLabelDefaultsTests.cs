using System.Reflection;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxChartLabelDefaultsTests
{
    private static readonly (string Element, string Property)[] Components =
    [
        ("showVal", "ShowValue"), ("showPercent", "ShowPercent"),
        ("showCatName", "ShowCategoryName"), ("showSerName", "ShowSeriesName"),
        ("showLegendKey", "ShowLegendKey"), ("showBubbleSize", "ShowBubbleSize")
    ];

    public static void PptxDefinedLabelsUseIndependentOfficeComponentDefaults()
    {
        // Native controls cover line, column, area, scatter, pie/doughnut and
        // bubble. Missing containers stay distinct from missing component flags.
        foreach (bool defined in new[] { false, true })
        {
            var (chart, plot) = Source(defined ? "<c:dLbls/>" : "");
            object options = Read(chart);
            foreach (var component in Components)
                TestAssert.Equal(defined, Flag(options, component.Property));
            TestAssert.True(!Flag(options, "ShowLeaderLines"), "Unproven leader-line fallback remains unchanged.");
            TestAssert.Equal(null, chart.Plots[0].DataLabels.ShowSeriesName);
            TestAssert.Equal(defined, chart.Plots[0].DataLabels.IsDefined);
        }
        foreach (var component in Components)
        foreach (string value in new[] { "0", "false", "1", "" })
        {
            string attribute = value.Length == 0 ? "" : $" val=\"{value}\"";
            var (chart, plot) = Source($"<c:dLbls><c:{component.Element}{attribute}/></c:dLbls>");
            object options = Read(chart);
            foreach (var other in Components)
                TestAssert.Equal(other == component ? value is not ("0" or "false") : true, Flag(options, other.Property));
            object raw = PptxTests.ChartDataLabelFlagOption(options, component.Element);
            TestAssert.True(PptxTests.ChartBooleanOptionIsDefined(raw), "Explicit flags retain source presence.");
            TestAssert.Equal(value, PptxTests.ChartBooleanOptionRawValue(raw));
        }
    }

    public static void PptxLabelContainersOwnComponentDefaultsAcrossSeriesAndPoints()
    {
        string disabled = "<c:dLbls>" + string.Concat(Components.Select(c => $"<c:{c.Element} val=\"0\"/>")) + "</c:dLbls>";
        const string series = """
            <c:dLbls><c:dLbl><c:idx val="0"/><c:showSerName val="1"/></c:dLbl>
              <c:dLbl><c:idx val="1"/><c:delete val="1"/></c:dLbl><c:showVal val="0"/><c:showSerName val="0"/></c:dLbls>
            """;
        foreach (string plotLabels in new[] { "", disabled })
        foreach (string seriesLabels in new[] { "<c:dLbls/>", series })
        {
            var (chart, plot) = Source(plotLabels, seriesLabels);
            object options = Read(chart);
            object perSeries = Invoke("ReadSceneSeriesDataLabelOptions", chart,
                chart.Plots[0]);
            object merged = Invoke("ResolveChartDataLabelOptionsForSeries", options, perSeries, 0);
            foreach (var component in Components)
                TestAssert.Equal(!(seriesLabels == series && component.Property is "ShowValue" or "ShowSeriesName"), Flag(merged, component.Property));
            if (seriesLabels == series)
            {
                object point = Invoke("ResolveChartDataLabelOptions", merged, 0);
                TestAssert.True(Flag(point, "ShowSeriesName"), "The point's explicit true overrides an inherited false.");
                TestAssert.True(Flag(point, "ShowValue"), "A defined point takes its own missing-component defaults.");
                object deleted = Invoke("ResolveChartDataLabelOptions", merged, 1);
                TestAssert.True(!Flag(deleted, "HasVisibleContent"), "Deleted labels remain absent.");
            }
        }
        foreach (bool deleted in new[] { false, true })
        {
            string point = "<c:dLbl><c:idx val=\"0\"/>" + (deleted ? "<c:delete val=\"1\"/>" : "") + "</c:dLbl>";
            var (chart, plot) = Source("", disabled.Replace("<c:dLbls>", "<c:dLbls>" + point, StringComparison.Ordinal));
            object options = Read(chart);
            object perSeries = Invoke("ReadSceneSeriesDataLabelOptions", chart, chart.Plots[0]);
            object merged = Invoke("ResolveChartDataLabelOptionsForSeries", options, perSeries, 0);
            TestAssert.Equal(!deleted, Flag(merged, "HasVisibleContent"));
        }
    }

    public static void PptxLineLabelDefaultsPreserveCompleteSourceTextWithoutKeys()
    {
        foreach (bool keys in new[] { false, true })
        {
            const string name = "A longer source series";
            var page = RenderLabelPage($"<c:dLbls><c:showLegendKey val=\"{(keys ? 1 : 0)}\"/></c:dLbls>", name);
            string text = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, false,
                fontResourceFilter: resource => resource.StartsWith("CLD", StringComparison.Ordinal)).Select(show => show.Text));
            TestAssert.Equal($"{name}, W1, 2{name}, W2, 4{name}, W3, 3", text);
        }
    }

    public static void PptxAutomaticLabelTextOwnsItsClippingRegion()
    {
        const string labels = "<c:dLbls><c:dLblPos val=\"t\"/><c:showLegendKey val=\"0\"/><c:showVal val=\"1\"/><c:showCatName val=\"1\"/><c:showSerName val=\"1\"/></c:dLbls>";
        var page = RenderLabelPage(labels, "Actual", atMaximum: true);
        var clips = Regex.Matches(page.Content, @"([-\d.]+) ([-\d.]+) ([\d.]+) ([\d.]+) re W\*? n").Cast<Match>().ToArray();
        var shows = Regex.Matches(page.Content, @"/CLD\S+ [\d.]+ Tf\s+\S+ Tc\s+1 0 0 1 (?<x>[-\d.]+) (?<y>[-\d.]+) Tm").Cast<Match>().ToArray();
        TestAssert.Equal(3, shows.Length);
        foreach (Match show in shows)
        {
            // Inspect the generated text's active clip, not merely its encoded
            // characters: a clipped label can still be present in PDF text data.
            var clip = clips.Last(match => match.Index < show.Index);
            double[] bounds = clip.Groups.Cast<Group>().Skip(1).Select(group => double.Parse(group.Value, CultureInfo.InvariantCulture)).ToArray();
            double x = double.Parse(show.Groups["x"].Value, CultureInfo.InvariantCulture);
            double y = double.Parse(show.Groups["y"].Value, CultureInfo.InvariantCulture);
            TestAssert.True(x >= bounds[0] - .001d && x <= bounds[0] + bounds[2] + .001d &&
                y >= bounds[1] - .001d && y <= bounds[1] + bounds[3] + .001d,
                "An automatic label baseline beyond the data plot remains inside its own text clip.");
        }
    }

    public static void PptxScatterLabelsPreserveFormattedXValues()
    {
        foreach (bool bubble in new[] { false, true })
        {
            var page = RenderScatterPage(bubble, null, xOnly: true);
            string text = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, false,
                fontResourceFilter: resource => resource.StartsWith("CSD", StringComparison.Ordinal)).Select(show => show.Text));
            TestAssert.Equal("1.252.503.75", text);
        }
    }

    public static void PptxScatterAndBubbleLabelsDefaultToRight()
    {
        // Independent Office exports have identical page streams for an absent
        // position and explicit r, for both scatter and bubble controls.
        foreach (bool bubble in new[] { false, true })
            TestAssert.Equal(RenderScatterPage(bubble, "r").Content, RenderScatterPage(bubble, null).Content);
    }

    public static void PptxBubbleLabelPlacementFollowsCircleEdgeAndKeepsCenter()
    {
        foreach (string position in new[] { "ctr", "r", "l", "t", "b" })
        {
            (double X, double Y) Origin(double size)
            {
                string content = RenderScatterPage(true, position, firstSize: size).Content;
                Match show = Regex.Match(content, @"/CSD\S+ [\d.]+ Tf\s+\S+ Tc\s+1 0 0 1 (?<x>[-\d.]+) (?<y>[-\d.]+) Tm");
                TestAssert.True(show.Success, "Expected a bubble label text origin.");
                return (double.Parse(show.Groups["x"].Value, CultureInfo.InvariantCulture),
                    double.Parse(show.Groups["y"].Value, CultureInfo.InvariantCulture));
            }
            var small = Origin(.25d);
            var large = Origin(1d);
            if (position == "ctr") TestAssert.Equal(small, large);
            else if (position == "r") TestAssert.True(large.X > small.X && large.Y == small.Y, "Right labels follow the right circle edge.");
            else if (position == "l") TestAssert.True(large.X < small.X && large.Y == small.Y, "Left labels follow the left circle edge.");
            else if (position == "t") TestAssert.True(large.Y > small.Y && large.X == small.X, "Top labels follow the top circle edge.");
            else TestAssert.True(large.Y < small.Y && large.X == small.X, "Bottom labels follow the bottom circle edge.");
        }
    }

    private static PdfPage RenderScatterPage(bool bubble, string? position, bool xOnly = false, double firstSize = 1d)
    {
        string labels = "<c:dLbls>" + string.Concat(Components.Select(c =>
            $"<c:{c.Element} val=\"{(c.Property == (xOnly ? "ShowCategoryName" : "ShowValue") ? 1 : 0)}\"/>")) +
            (position is null ? "" : $"<c:dLblPos val=\"{position}\"/>") + "</c:dLbls>";
        var (_, plot) = Source(labels);
        plot.Name = C + (bubble ? "bubbleChart" : "scatterChart");
        XElement series = plot.Element(C + "ser")!;
        series.Element(C + "cat")!.ReplaceWith(XElement.Parse($$"""
            <c:xVal xmlns:c="{{C}}"><c:numLit><c:formatCode>0.00</c:formatCode><c:pt idx="0"><c:v>1.25</c:v></c:pt><c:pt idx="1"><c:v>2.5</c:v></c:pt><c:pt idx="2"><c:v>3.75</c:v></c:pt></c:numLit></c:xVal>
            """));
        series.Element(C + "val")!.Name = C + "yVal";
        if (bubble)
            series.Add(XElement.Parse($$"""
                <c:bubbleSize xmlns:c="{{C}}"><c:numLit><c:pt idx="0"><c:v>{{firstSize.ToString(CultureInfo.InvariantCulture)}}</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numLit></c:bubbleSize>
                """));
        return RenderChartPage(new XDocument(plot.Document!));
    }

    private static PdfPage RenderLabelPage(string labels, string name, bool atMaximum = false)
    {
        var (_, plot) = Source(labels);
        plot.Name = C + "lineChart";
        plot.Descendants(C + "tx").First().Element(C + "v")!.Value = name;
        if (atMaximum)
        {
            plot.Add(new XElement(C + "axId", new XAttribute("val", 1)), new XElement(C + "axId", new XAttribute("val", 2)));
            plot.Parent!.Add(XElement.Parse("<c:catAx xmlns:c=\"" + C + "\"><c:axId val=\"1\"/><c:axPos val=\"b\"/><c:crossAx val=\"2\"/><c:tickLblPos val=\"none\"/></c:catAx>"));
            plot.Parent.Add(XElement.Parse("<c:valAx xmlns:c=\"" + C + "\"><c:axId val=\"2\"/><c:scaling><c:min val=\"0\"/><c:max val=\"4\"/></c:scaling><c:axPos val=\"l\"/><c:crossAx val=\"1\"/><c:tickLblPos val=\"none\"/></c:valAx>"));
        }
        return RenderChartPage(new XDocument(plot.Document!));
    }

    private static PdfPage RenderChartPage(XDocument xml, IReadOnlyDictionary<string, byte[]>? additionalParts = null, IFontResolver? fontResolver = null)
    {
        const string slide = """
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><p:cSld><p:spTree><p:graphicFrame><p:xfrm><a:off x="914400" y="457200"/><a:ext cx="9144000" cy="5486400"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>
            """;
        var parts = new Dictionary<string, byte[]> { ["ppt/slides/slide1.xml"] = TestFixtures.Utf8(slide) };
        if (additionalParts is not null)
            foreach (var part in additionalParts) parts[part.Key] = part.Value;
        using var stream = PptxTests.CreateSingleChartPackageStream(xml.ToString(), parts);
        using var package = OoxPackage.Open(stream, default);
        var document = new PptxReader().Read(package, default);
        return new PptxRenderer(fontResolver ?? new TestFaceFontResolver()).RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
    }

    private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    private static (PptxSceneChart Chart, XElement Plot) Source(string plotLabels, string seriesLabels = "")
    {
        var xml = XDocument.Parse($$"""
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="800"/></a:pPr></a:p></c:txPr><c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Actual</c:v></c:tx><c:cat><c:strLit><c:pt idx="0"><c:v>W1</c:v></c:pt><c:pt idx="1"><c:v>W2</c:v></c:pt><c:pt idx="2"><c:v>W3</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numLit></c:val>{{seriesLabels}}</c:ser>{{plotLabels}}</c:barChart></c:plotArea></c:chart></c:chartSpace>
            """);
        return (PptxTests.BuildSingleChartScene(xml.ToString())!, xml.Descendants(C + "barChart").Single());
    }

    private static object Read(PptxSceneChart chart)
        => Invoke("ReadSceneDataLabelOptions", chart, chart.Plots[0]);

    private static object Invoke(string name, params object?[] arguments)
        => typeof(PptxRenderer).GetMethod(name, BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null, arguments)!;

    private static bool Flag(object value, string name) => (bool)value.GetType().GetProperty(name)!.GetValue(value)!;
}
