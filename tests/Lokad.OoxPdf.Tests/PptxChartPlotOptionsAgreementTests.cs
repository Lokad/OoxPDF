using System.IO.Compression;
using System.Reflection;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

// Parsed plot/axis options use explicit source-backed expectations.
internal static class PptxChartPlotOptionsAgreementTests
{
    private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    public static void BarScalarOptionsFollowSourceDefaultsAndClamps()
    {
        (string Xml, PptxSceneChartGrouping Expected)[] groups = [
            ("<c:grouping val=\"clustered\"/>", PptxSceneChartGrouping.Clustered),
            ("<c:grouping val=\"stacked\"/>", PptxSceneChartGrouping.Stacked),
            ("<c:grouping val=\"bogus\"/>", PptxSceneChartGrouping.Standard), ("", PptxSceneChartGrouping.Standard)];
        (string Xml, PptxSceneChartBarDirection Expected)[] directions = [
            ("<c:barDir val=\"bar\"/>", PptxSceneChartBarDirection.Bar),
            ("<c:barDir val=\"col\"/>", PptxSceneChartBarDirection.Column),
            ("<c:barDir val=\"bogus\"/>", PptxSceneChartBarDirection.Column), ("", PptxSceneChartBarDirection.Column)];
        (string Xml, bool Value, bool Defined)[] varies = [
            ("<c:varyColors/>", true, true), ("<c:varyColors val=\"0\"/>", false, true),
            ("<c:varyColors val=\"1\"/>", true, true), ("", true, false)];
        (string Xml, double Expected)[] gaps = [("<c:gapWidth val=\"150\"/>", 150),
            ("<c:gapWidth val=\"999\"/>", 500), ("<c:gapWidth val=\"-5\"/>", 0), ("", 150)];
        (string Xml, double Expected)[] overlaps = [("<c:overlap val=\"0\"/>", 0),
            ("<c:overlap val=\"250\"/>", 100), ("<c:overlap val=\"-150\"/>", -100), ("", 0)];
        foreach (var group in groups)
        foreach (var direction in directions)
        foreach (var vary in varies)
        foreach (var gap in gaps)
        foreach (var overlap in overlaps)
        {
            var (plot, _) = LoadPlot("barChart", group.Xml + direction.Xml + vary.Xml + gap.Xml + overlap.Xml);
            object options = Invoke("ReadSceneChartBarOptions", [typeof(PptxSceneChartPlot), typeof(PptxSceneChartGrouping)],
                [plot, PptxSceneChartGrouping.Standard])!;
            object Value(string name) => options.GetType().GetProperty(name)!.GetValue(options)!;
            TestAssert.Equal(group.Expected, (PptxSceneChartGrouping)Value("Grouping"));
            TestAssert.Equal(direction.Expected, (PptxSceneChartBarDirection)Value("BarDirection"));
            TestAssert.Equal(vary.Value, PptxTests.ChartBooleanOptionValue(Value("VaryColors")));
            TestAssert.Equal(vary.Defined, PptxTests.ChartBooleanOptionIsDefined(Value("VaryColors")));
            TestAssert.Equal(gap.Expected, (double)Value("GapWidth"));
            TestAssert.Equal(overlap.Expected, (double)Value("Overlap"));
        }
    }

    public static void ScatterStyleFollowsSourceDefaults()
    {
        (string Xml, PptxSceneChartScatterStyle Expected)[] styles = [
            ("<c:scatterStyle val=\"marker\"/>", PptxSceneChartScatterStyle.Marker),
            ("<c:scatterStyle val=\"smoothMarker\"/>", PptxSceneChartScatterStyle.SmoothMarker),
            ("<c:scatterStyle val=\"bogus\"/>", PptxSceneChartScatterStyle.Unknown),
            ("", PptxSceneChartScatterStyle.LineMarker)];
        foreach (var style in styles)
        {
            var (plot, _) = LoadPlot("scatterChart", style.Xml);
            object options = Invoke("ReadSceneChartScatterOptions", [typeof(PptxSceneChartPlot)], [plot])!;
            TestAssert.Equal(style.Expected, (PptxSceneChartScatterStyle)options.GetType().GetProperty("ScatterStyle")!.GetValue(options)!);
        }
    }

    public static void RadarStyleFollowsSourceDefaults()
    {
        (string Xml, PptxSceneChartRadarStyle Expected)[] styles = [
            ("<c:radarStyle val=\"standard\"/>", PptxSceneChartRadarStyle.Standard),
            ("<c:radarStyle val=\"marker\"/>", PptxSceneChartRadarStyle.Marker),
            ("<c:radarStyle val=\"filled\"/>", PptxSceneChartRadarStyle.Filled),
            ("<c:radarStyle val=\"bogus\"/>", PptxSceneChartRadarStyle.Standard), ("", PptxSceneChartRadarStyle.Standard)];
        foreach (var style in styles)
        {
            var (plot, _) = LoadPlot("radarChart", style.Xml);
            object result = Invoke("ReadSceneChartRadarStyle", [typeof(PptxSceneChartPlot)], [plot])!;
            TestAssert.Equal(style.Expected, (PptxSceneChartRadarStyle)result);
        }
    }

    public static void SeriesSmoothUsesSourceAndExplicitFallback()
    {
        (string Xml, bool? Value)[] forms = [("<c:smooth/>", true), ("<c:smooth val=\"0\"/>", false),
            ("<c:smooth val=\"1\"/>", true), ("<c:smooth val=\"bogus\"/>", false), ("", null)];
        foreach (bool defaultValue in new[] { false, true })
        foreach (bool plotFallback in new[] { false, true })
        foreach (var plotForm in forms)
        foreach (var seriesForm in forms)
        {
            var (plot, _) = LoadPlot("barChart", plotForm.Xml, seriesForm.Xml);
            var options = (System.Collections.IEnumerable)Invoke("ReadSceneSmoothSeries",
                [typeof(PptxSceneChartPlot), typeof(bool), typeof(bool)], [plot, plotFallback, defaultValue])!;
            object option = options.Cast<object>().Single();
            TestAssert.Equal(seriesForm.Value ?? (plotFallback ? plotForm.Value : null) ?? defaultValue,
                PptxTests.ChartBooleanOptionValue(option));
            TestAssert.Equal(seriesForm.Value.HasValue, PptxTests.ChartBooleanOptionIsDefined(option));
        }
    }

    public static void SeriesLineVisibilityUsesSource()
    {
        (string Xml, bool Hidden)[] forms = [("", false), ("<c:spPr><a:ln w=\"12700\"/></c:spPr>", false),
            ("<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>", true)];
        foreach (var form in forms)
        {
            var (plot, _) = LoadPlot("barChart", "", form.Xml);
            var values = (IReadOnlyList<bool>)Invoke("ReadSceneSeriesLineHidden", [typeof(PptxSceneChartPlot)], [plot])!;
            TestAssert.Equal(form.Hidden, values.Single());
        }
    }

    public static void SeriesNamesUseParsedSourceValues()
    {
        (string Xml, string Expected, string Source)[] cases =
        [
            ("", "Series 1", "Default"),
            ("<c:tx><c:v>Solo</c:v></c:tx>", "Solo", "Cache"),
            ("<c:tx><c:v>  Padded  </c:v></c:tx>", "Padded", "Cache"),
            ("<c:tx><c:strRef><c:f>Sheet1!$A$1</c:f><c:strCache><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>Cached</c:v></c:pt></c:strCache></c:strRef></c:tx>", "Cached", "Cache"),
            ("<c:tx><c:strRef><c:f>Sheet1!$A$1</c:f></c:strRef></c:tx>", "Series 1", "Default")
        ];
        Type workbookType = TestAssert.NotNull(typeof(PptxRenderer).Assembly.GetType("Lokad.OoxPdf.Pptx.PptxRenderer+ChartWorkbookData"));
        foreach (var test in cases)
        {
            var (plot, _) = LoadPlot("barChart", "", test.Xml);
            object record = ((System.Collections.IEnumerable)TestAssert.NotNull(Invoke("ReadSceneChartSeriesNameRecords",
                [typeof(PptxSceneChartPlot), workbookType], [plot, null]))).Cast<object>().Single();
            TestAssert.Equal(test.Expected, (string?)record.GetType().GetProperty("ActiveName")?.GetValue(record) ?? string.Empty);
            TestAssert.Equal(test.Source, record.GetType().GetProperty("ActiveNameSource")?.GetValue(record)?.ToString() ?? string.Empty);
        }
    }

    public static void SeriesExplosionsRespectParsedValuesAndPointOverrides()
    {
        (string Xml, double? Expected)[] cases =
        [
            ("", null),
            ("<c:explosion val=\"25\"/>", 0.25d),
            ("<c:explosion val=\"150\"/>", 1d),
            ("<c:explosion val=\"-25\"/>", 0d),
            ("<c:explosion val=\"bogus\"/>", null),
            ("<c:dPt><c:idx val=\"0\"/><c:explosion val=\"50\"/></c:dPt>", 0.5d),
            ("<c:dPt><c:idx val=\"bogus\"/><c:explosion val=\"50\"/></c:dPt>", null),
            ("<c:explosion val=\"25\"/><c:dPt><c:idx val=\"0\"/><c:explosion val=\"50\"/></c:dPt>", 0.5d),
            ("<c:explosion val=\"150\"/><c:dPt><c:idx val=\"0\"/><c:explosion val=\"-50\"/></c:dPt>", 0d)
        ];
        Type workbookType = TestAssert.NotNull(typeof(PptxRenderer).Assembly.GetType("Lokad.OoxPdf.Pptx.PptxRenderer+ChartWorkbookData"));

        foreach (var test in cases)
        {
            var (plot, _) = LoadPlot("pieChart", "", test.Xml);
            var actual = (IReadOnlyDictionary<int, double>)TestAssert.NotNull(Invoke("ReadSceneChartPointExplosions",
                [typeof(PptxSceneChartPlot), workbookType], [plot, null]));
            TestAssert.Equal(test.Expected.HasValue ? 1 : 0, actual.Count);
            if (test.Expected is { } expected) TestAssert.Equal(expected, actual[0]);
        }
    }

    public static void ParsedAxisNumericOptionsKeepUnitsAndOrientation()
    {
        (string Xml, double? Major, double? Minor, bool Reversed)[] cases =
        [
            ("", null, null, false),
            ("<c:majorUnit val=\"10\"/>", 10, null, false),
            ("<c:minorUnit val=\"5\"/>", null, 5, false),
            ("<c:majorUnit val=\"bogus\"/><c:minorUnit val=\"0\"/>", null, null, false),
            ("<c:majorUnit val=\"-2\"/>", null, null, false),
            ("<c:scaling><c:orientation val=\"maxMin\"/></c:scaling>", null, null, true),
            ("<c:scaling><c:orientation val=\"minMax\"/></c:scaling>", null, null, false),
            ("<c:scaling><c:orientation val=\"bogus\"/></c:scaling>", null, null, false)
        ];
        foreach (var test in cases)
        {
            var axis = LoadAxis(test.Xml);
            object units = TestAssert.NotNull(Invoke("ReadSceneChartValueAxisUnits", [typeof(PptxSceneChartAxis)], [axis]));
            TestAssert.Equal(test.Major, (double?)units.GetType().GetProperty("MajorUnit")!.GetValue(units));
            TestAssert.Equal(test.Minor, (double?)units.GetType().GetProperty("MinorUnit")!.GetValue(units));
            TestAssert.Equal(test.Reversed, (bool)TestAssert.NotNull(Invoke("ReadSceneValueAxisReversed", [typeof(PptxSceneChartAxis)], [axis])));
        }
    }

    public static void MissingAxisNumericOptionsRetainAutomaticFallbacks()
    {
        object units = TestAssert.NotNull(Invoke("ReadSceneChartValueAxisUnits", [typeof(PptxSceneChartAxis)], [null]));
        TestAssert.Equal(null, units.GetType().GetProperty("MajorUnit")!.GetValue(units));
        TestAssert.Equal(null, units.GetType().GetProperty("MinorUnit")!.GetValue(units));
        TestAssert.Equal(false, (bool)TestAssert.NotNull(Invoke("ReadSceneValueAxisReversed", [typeof(PptxSceneChartAxis)], [null])));
        Type extentsType = TestAssert.NotNull(typeof(PptxRenderer).GetNestedType("ChartValueExtents", BindingFlags.NonPublic));
        foreach (var test in new[] { (-5d, 5d, 0d), (2d, 5d, 2d), (-5d, -2d, -2d), (0d, 5d, 0d), (-5d, 0d, 0d), (0d, 0d, 0d) })
        {
            object extents = TestAssert.NotNull(Activator.CreateInstance(extentsType, [test.Item1, test.Item2]));
            TestAssert.Equal(test.Item3, (double)TestAssert.NotNull(Invoke("ReadSceneValueAxisCrossingValue",
                [typeof(PptxSceneChartAxis), extentsType], [null, extents])));
        }
    }

    public static void ParsedAxisGridlineOptionsKeepVisibilityAndStroke()
    {
        (string Xml, bool Major, bool Minor, bool MajorElement)[] cases =
        [
            ("", false, false, false),
            ("<c:majorGridlines/>", true, false, true),
            ("<c:minorGridlines/>", false, true, false),
            ("<c:majorGridlines/><c:minorGridlines/>", true, true, true),
            ("<c:majorGridlines><c:spPr><a:ln><a:noFill/></a:ln></c:spPr></c:majorGridlines>", false, false, true),
            ("<c:minorGridlines><c:spPr><a:ln><a:noFill/></a:ln></c:spPr></c:minorGridlines>", false, false, false),
            ("<c:scaling><c:orientation val=\"bogus\"/></c:scaling>", false, false, false),
            ("<c:majorUnit val=\"bogus\"/><c:numFmt formatCode=\"0.0\"/>", false, false, false)
        ];
        var extents = new PptxRenderer.ChartValueExtents(-5d, 5d);
        object Options(PptxSceneChartAxis? axis) => TestAssert.NotNull(Invoke("ReadSceneChartValueAxisRenderOptions",
            [typeof(PptxSceneChartAxis), typeof(PptxRenderer.ChartValueExtents), typeof(bool)], [axis, extents, false]));
        object? Property(object value, string name) => value.GetType().GetProperty(name)!.GetValue(value);
        foreach (var test in cases)
        {
            object options = Options(LoadAxis(test.Xml));
            TestAssert.Equal(test.Major, (bool)Property(options, "MajorGridlines")!);
            TestAssert.Equal(test.Minor, (bool)Property(options, "MinorGridlines")!);
            TestAssert.Equal(test.MajorElement, (bool)Property(options, "HasMajorGridlineElement")!);
        }
        foreach (string kind in new[] { "major", "minor" })
        {
            object options = Options(LoadAxis($"<c:{kind}Gridlines><c:spPr><a:ln w=\"25400\"><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill></a:ln></c:spPr></c:{kind}Gridlines>"));
            object style = TestAssert.NotNull(Property(options, "GridlineStyle"));
            object stroke = TestAssert.NotNull(Property(style, kind == "major" ? "Major" : "Minor"));
            TestAssert.Equal(new RgbColor(255, 0, 0), (RgbColor)Property(stroke, "Color")!);
            TestAssert.Equal(2d, (double)Property(stroke, "Width")!);
        }
        object missing = Options(null);
        TestAssert.Equal(false, (bool)Property(missing, "MajorGridlines")!);
        TestAssert.Equal(false, (bool)Property(missing, "MinorGridlines")!);
        TestAssert.Equal(false, (bool)Property(missing, "HasMajorGridlineElement")!);
        object emptyStyle = TestAssert.NotNull(Property(missing, "GridlineStyle"));
        TestAssert.Equal(null, Property(emptyStyle, "Major"));
        TestAssert.Equal(null, Property(emptyStyle, "Minor"));
    }

    public static void ParsedAxisVisibilityKeepsAbsentDeletedAndUnknownDefaults()
    {
        bool Visible(PptxSceneChartAxis? axis) => (bool)TestAssert.NotNull(Invoke("IsSceneChartAxisLabelVisible",
            [typeof(PptxSceneChartAxis)], [axis]));
        TestAssert.Equal(true, Visible(null));
        bool ExplicitValue(PptxSceneChartAxis? axis) => (bool)TestAssert.NotNull(Invoke("IsSceneVisibleValueAxis",
            [typeof(PptxSceneChartAxis)], [axis]));
        TestAssert.Equal(false, ExplicitValue(null));
        TestAssert.Equal(true, ExplicitValue(LoadAxis("")));
        TestAssert.Equal(false, ExplicitValue(LoadAxis("<c:delete val=\"1\"/>")));
        foreach (var test in new[]
        {
            (Xml: "", Expected: true),
            (Xml: "<c:delete val=\"1\"/>", Expected: false),
            (Xml: "<c:delete val=\"0\"/>", Expected: true),
            (Xml: "<c:tickLblPos val=\"none\"/>", Expected: false),
            (Xml: "<c:tickLblPos val=\"bogus\"/>", Expected: true),
            (Xml: "<c:delete val=\"1\"/><c:tickLblPos val=\"high\"/>", Expected: false)
        })
            TestAssert.Equal(test.Expected, Visible(LoadAxis(test.Xml)));
    }

    public static void ParsedAxisPairsKeepCrossingAndLabelSideOwners()
    {
        (string Kind, string Orientation, string Cross, bool Reciprocal, string Label,
            string CrossAt, bool ExpectedSide, bool ExpectedLabelSide)[] cases =
        [
            ("catAx", "maxMin", "max", true, "low", "", false, true),
            ("catAx", "minMax", "max", true, "high", "", true, true),
            ("catAx", "maxMin", "min", true, "high", "", true, false),
            ("catAx", "maxMin", "max", false, "low", "", true, false),
            ("dateAx", "maxMin", "max", true, "low", "", true, false),
            ("catAx", "maxMin", "", true, "high", "<c:crossesAt val=\"2\"/>", true, false),
            ("catAx", "minMax", "bogus", true, "nextTo", "", true, true),
            ("catAx", "minMax", "", true, "nextTo", "", false, true)
        ];
        bool Resolve(string method, PptxSceneChart? chart, PptxSceneChartAxis? axis, bool fallback) =>
            (bool)TestAssert.NotNull(Invoke(method,
                [typeof(PptxSceneChart), typeof(PptxSceneChartAxis), typeof(bool)], [chart, axis, fallback]));
        foreach (var test in cases)
        {
            string cross = test.Cross.Length == 0 ? "" : $"<c:crosses val=\"{test.Cross}\"/>";
            PptxSceneChart chart = TestAssert.NotNull(PptxTests.BuildSingleChartScene(ChartSpace($$"""
                <c:barChart><c:axId val="10"/><c:axId val="20"/></c:barChart>
                <c:{{test.Kind}}><c:axId val="10"/><c:crossAx val="{{(test.Reciprocal ? "20" : "99")}}"/>
                  <c:scaling><c:orientation val="{{test.Orientation}}"/></c:scaling></c:{{test.Kind}}>
                <c:valAx><c:axId val="20"/><c:crossAx val="10"/><c:axPos val="r"/>
                  {{cross}}{{test.CrossAt}}<c:tickLblPos val="{{test.Label}}"/></c:valAx>
                """, "")));
            PptxSceneChartAxis axis = chart.Axes.Single(axis => axis.AxisKind == PptxSceneChartAxisKind.Value);
            TestAssert.Equal(test.ExpectedSide, Resolve("ResolveSceneValueAxisRightSide", chart, axis, true));
            TestAssert.Equal(test.ExpectedLabelSide, Resolve("ResolveSceneValueAxisLabelsRightSide", chart, axis, true));
        }
        foreach (bool fallback in new[] { false, true })
        {
            TestAssert.Equal(fallback, Resolve("ResolveSceneValueAxisRightSide", null, null, fallback));
            TestAssert.Equal(fallback, Resolve("ResolveSceneValueAxisLabelsRightSide", null, null, fallback));
        }
    }

    private static PptxSceneChartAxis LoadAxis(string axisInner)
    {
        string xml = ChartSpace("<c:barChart><c:ser><c:cat><c:strLit><c:pt idx=\"0\"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx=\"0\"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val=\"10\"/><c:axId val=\"20\"/></c:barChart><c:valAx><c:axId val=\"10\"/>" + axisInner + "</c:valAx>", "");
        PptxSceneChart sceneChart = TestAssert.NotNull(PptxTests.BuildSingleChartScene(xml));
        TestAssert.Equal(1, sceneChart.Axes.Count);
        return sceneChart.Axes[0];
    }

    public static void ChartTitleSceneRetainsAuthoredFormatting()
    {
        string[] titles =
        [
            "",
            """<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx><c:layout/><c:overlay val="0"/></c:title>""",
            """<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:rPr b="1" sz="1400"/><a:t>Q1 </a:t></a:r><a:r><a:rPr i="1"/><a:t>rev</a:t></a:r></a:p></c:rich></c:tx><c:layout/><c:overlay val="1"/></c:title>""",
            """<c:title><c:txPr><a:bodyPr rot="5400000"/></c:txPr><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Spun</a:t></a:r></a:p></c:rich></c:tx><c:layout/></c:title>"""
        ];
        string?[] expectedText = [null, "Sales", "Q1 rev", "Spun"];
        bool?[] expectedOverlay = [null, false, true, null];
        for (int index = 0; index < titles.Length; index++)
        {
            string xml = ChartSpace("""<c:barChart><c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val="10"/><c:axId val="20"/></c:barChart>""", titles[index]);
            PptxSceneChartTitle title = TestAssert.NotNull(PptxTests.BuildSingleChartScene(xml)).Title;
            TestAssert.True(title.Text == expectedText[index], "Authored title text is retained.");
            TestAssert.True(title.Overlay == expectedOverlay[index], "Title overlay retains its declared or absent value.");
            TestAssert.True(!title.IsAutoGenerated, "Absent automatic-title eligibility cannot invent a title.");
            if (index == 0)
            {
                TestAssert.Equal(0, title.TextRuns.Count);
            }
            else if (index == 2)
            {
                TestAssert.Equal(2, title.TextRuns.Count);
                TestAssert.Equal("Q1 ", title.TextRuns[0].Text);
                TestAssert.Equal(14d, title.TextRuns[0].TextStyle.FontSize ?? 0d);
                TestAssert.True(title.TextRuns[0].TextStyle.Bold == true, "First title run stays bold.");
                TestAssert.Equal("rev", title.TextRuns[1].Text);
                TestAssert.True(title.TextRuns[1].TextStyle.Italic == true, "Second title run stays italic.");
            }
            else if (index == 3)
            {
                TestAssert.Equal(90d, title.TextBodyProperties.RotationDegrees ?? 0d);
                TestAssert.Equal("5400000", title.TextBodyProperties.RotationValue);
            }
        }
    }

    public static void LegendOptionsUseParsedSceneDefaults()
    {
        (string Xml, bool Visible, PptxSceneChartLegendPosition Position, bool Overlay, double Size)[] legends =
        [
            ("", false, PptxSceneChartLegendPosition.Right, true, 9d),
            ("<c:legend><c:legendPos val=\"r\"/><c:layout/><c:overlay val=\"0\"/></c:legend>", true, PptxSceneChartLegendPosition.Right, false, 9d),
            ("<c:legend><c:legendPos val=\"b\"/><c:layout/><c:overlay val=\"1\"/></c:legend>", true, PptxSceneChartLegendPosition.Bottom, true, 9d),
            ("<c:legend><c:legendPos val=\"bogus\"/><c:layout/></c:legend>", true, PptxSceneChartLegendPosition.Right, true, 9d),
            ("<c:legend><c:legendPos val=\"r\"/><c:overlay val=\"0\"/><c:delete val=\"1\"/></c:legend>", false, PptxSceneChartLegendPosition.Right, false, 9d),
            ("<c:legend><c:txPr><a:p><a:pPr><a:defRPr sz=\"1800\" b=\"1\"/></a:pPr></a:p></c:txPr></c:legend>", true, PptxSceneChartLegendPosition.Right, true, 18d)
        ];
        foreach (var legend in legends)
        {
            string xml = ChartSpace("<c:barChart><c:ser><c:cat><c:strLit><c:pt idx=\"0\"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx=\"0\"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val=\"10\"/><c:axId val=\"20\"/></c:barChart>", legend.Xml);
            PptxSceneChart sceneChart = TestAssert.NotNull(PptxTests.BuildSingleChartScene(xml));
            object layout = TestAssert.NotNull(Invoke("ReadSceneChartLegendLayout", new[] { typeof(PptxSceneChart) }, new object?[] { sceneChart }));
            Type type = layout.GetType();
            TestAssert.Equal(legend.Visible, (bool)TestAssert.NotNull(type.GetProperty("Visible")!.GetValue(layout)));
            if (legend.Visible)
            {
                TestAssert.Equal(legend.Position, (PptxSceneChartLegendPosition)TestAssert.NotNull(type.GetProperty("PositionKind")!.GetValue(layout)));
                TestAssert.Equal(legend.Overlay, (bool)TestAssert.NotNull(type.GetProperty("Overlay")!.GetValue(layout)));
            }
            object style = TestAssert.NotNull(Invoke("ReadSceneChartLegendTextStyle", new[] { typeof(PptxTheme), typeof(PptxSceneChart) }, new object?[] { PptxTheme.Empty, sceneChart }));
            TestAssert.Equal(legend.Size, (double)TestAssert.NotNull(style.GetType().GetProperty("FontSize")!.GetValue(style)));
            TestAssert.Equal(legend.Size == 18d, (bool)TestAssert.NotNull(style.GetType().GetProperty("Bold")!.GetValue(style)));
        }
    }

    public static void ChartLevelOptionsKeepParsedSceneDefaults()
    {
        (string Xml, PptxSceneChartDisplayBlanksAs Expected)[] blanks =
        [
            ("<c:dispBlanksAs val=\"span\"/>", PptxSceneChartDisplayBlanksAs.Span),
            ("<c:dispBlanksAs val=\"zero\"/>", PptxSceneChartDisplayBlanksAs.Zero),
            ("<c:dispBlanksAs val=\"gap\"/>", PptxSceneChartDisplayBlanksAs.Gap),
            ("<c:dispBlanksAs val=\"bogus\"/>", PptxSceneChartDisplayBlanksAs.Gap),
            ("", PptxSceneChartDisplayBlanksAs.Gap)
        ];
        (string Xml, bool? Expected)[] visible = [("<c:plotVisOnly val=\"0\"/>", false),
            ("<c:plotVisOnly val=\"1\"/>", true), ("", null)];
        foreach (var blank in blanks)
        {
            foreach (var vis in visible)
            {
                string xml = ChartSpace("<c:barChart><c:ser><c:cat><c:strLit><c:pt idx=\"0\"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx=\"0\"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val=\"10\"/><c:axId val=\"20\"/></c:barChart>", blank.Xml + vis.Xml);
                PptxSceneChart sceneChart = TestAssert.NotNull(PptxTests.BuildSingleChartScene(xml));
                object? effective = Invoke("ResolveChartDisplayBlanksAs", new[] { typeof(PptxSceneChart) }, new object?[] { sceneChart });
                TestAssert.Equal(blank.Expected, (PptxSceneChartDisplayBlanksAs)TestAssert.NotNull(effective));
                TestAssert.Equal(vis.Expected, sceneChart.Options.PlotVisibleOnly);
            }
        }
    }

    public static void ScenePlotsCorrespondToXmlPlotElementsAcrossFixtures()
    {
        // Every plotArea chart child keeps one scene owner, its source element,
        // order and per-kind index. Read the qualified fixture snapshot.
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases"));
        int charts = 0;
        int plots = 0;
        foreach (string fixture in Directory.GetFiles(cases, "*.pptx"))
        {
            if (!ZipHasChartPart(fixture))
            {
                continue;
            }

            using FileStream stream = File.OpenRead(fixture);
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
            PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
            foreach (PptxSceneSlide slide in scene.Slides)
            {
                foreach (PptxSceneNode node in AllNodes(slide))
                {
                    if (node.Chart is not { ChartXml: { } xml } chart)
                    {
                        continue;
                    }

                    charts++;
                    List<XElement> elements = xml.Descendants(C + "plotArea").FirstOrDefault()?
                        .Elements()
                        .Where(e => e.Name.Namespace == C && e.Name.LocalName.EndsWith("Chart", StringComparison.Ordinal))
                        .ToList() ?? new List<XElement>();
                    TestAssert.Equal(elements.Count, chart.Plots.Count);
                    for (int i = 0; i < elements.Count; i++)
                    {
                        TestAssert.Equal(elements[i].Name.LocalName, chart.Plots[i].Kind);
                        TestAssert.True(ReferenceEquals(elements[i], chart.Plots[i].Source), "Scene plot must wrap the plotArea child element itself.");
                    }

                    var seen = new Dictionary<string, int>(StringComparer.Ordinal);
                    foreach (PptxSceneChartPlot plot in chart.Plots)
                    {
                        int expected = seen.TryGetValue(plot.Kind, out int n) ? n : 0;
                        seen[plot.Kind] = expected + 1;
                        TestAssert.Equal(expected, plot.KindIndex);
                    }

                    plots += elements.Count;
                }
            }
        }

        TestAssert.True(charts > 0 && plots > 0, "Chart corpus must not be empty.");
    }

    private static bool ZipHasChartPart(string path)
    {
        using ZipArchive archive = ZipFile.OpenRead(path);
        foreach (ZipArchiveEntry entry in archive.Entries)
        {
            if (entry.FullName.Contains("/charts/", StringComparison.Ordinal))
            {
                return true;
            }
        }

        return false;
    }

    private static IEnumerable<PptxSceneNode> AllNodes(PptxSceneSlide slide)
    {
        foreach (PptxSceneNode node in slide.MasterNodes.Concat(slide.LayoutNodes).Concat(slide.SlideNodes))
        {
            foreach (PptxSceneNode nested in Walk(node))
            {
                yield return nested;
            }
        }
    }

    private static IEnumerable<PptxSceneNode> Walk(PptxSceneNode node)
    {
        yield return node;
        foreach (PptxSceneNode child in node.Children)
        {
            foreach (PptxSceneNode nested in Walk(child))
            {
                yield return nested;
            }
        }
    }

    private static bool SequenceEqual(object? xml, object? scene)
    {
        if (xml is System.Collections.IEnumerable xmlItems && scene is System.Collections.IEnumerable sceneItems)
        {
            var xmlList = xmlItems.Cast<object?>().ToList();
            var sceneList = sceneItems.Cast<object?>().ToList();
            return xmlList.Count == sceneList.Count && xmlList.SequenceEqual(sceneList);
        }

        return Equals(xml, scene);
    }

    private static bool SeriesNameEqual(object? xml, object? scene)
    {
        if (xml is null || scene is null)
        {
            return xml is null && scene is null;
        }

        // WorkbookPoints is a fresh list per arm, so compare its contents rather
        // than the record (list identity would never agree across arms).
        return Prop(xml, "ActiveName") == Prop(scene, "ActiveName") &&
            Prop(xml, "CacheName") == Prop(scene, "CacheName") &&
            Prop(xml, "ActiveNameSource") == Prop(scene, "ActiveNameSource") &&
            PointsText(xml, "WorkbookPoints").SequenceEqual(PointsText(scene, "WorkbookPoints"));
    }

    private static string? Prop(object record, string name)
    {
        return record.GetType().GetProperty(name)?.GetValue(record)?.ToString();
    }

    private static List<string?> PointsText(object record, string name)
    {
        if (record.GetType().GetProperty(name)?.GetValue(record) is System.Collections.IEnumerable points)
        {
            return points.Cast<object?>().Select(point => point?.ToString()).ToList();
        }

        return new List<string?>();
    }

    private static (PptxSceneChartPlot Plot, XElement Element) LoadPlot(string plotKind, string plotInner, string seriesInner = "")
    {
        string xml = ChartSpace("<c:" + plotKind + ">" + plotInner + "<c:ser>" + seriesInner + "<c:cat><c:strLit><c:pt idx=\"0\"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx=\"0\"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val=\"10\"/><c:axId val=\"20\"/></c:" + plotKind + ">", "");
        PptxSceneChart sceneChart = TestAssert.NotNull(PptxTests.BuildSingleChartScene(xml));
        PptxSceneChart chart = sceneChart;
        TestAssert.Equal(1, chart.Plots.Count);
        XElement element = XDocument.Parse(xml).Descendants(C + plotKind).First();
        return (chart.Plots[0], element);
    }

    private static string ChartSpace(string plotAreaInner, string options)
    {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
            + "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">"
            + "<c:chart><c:plotArea>" + plotAreaInner + "</c:plotArea>"
            + options
            + "</c:chart></c:chartSpace>";
    }

    private static object? Invoke(string name, Type[] types, object?[] args)
    {
        MethodInfo method = typeof(PptxRenderer).GetMethod(
            name,
            BindingFlags.NonPublic | BindingFlags.Static,
            null,
            types,
            null) ?? throw new InvalidOperationException("Expected renderer bridge: " + name);
        try
        {
            return method.Invoke(null, args);
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }
    }
}
