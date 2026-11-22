using System.IO.Compression;
using System.Reflection;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

// Parsed plot/axis options use explicit source-backed expectations.
internal static class PptxChartPlotOptionsAgreementTests
{
    private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";


    public static void BarScalarOptionsFollowSourceDefaultsAndClamps()
    {
        Check("");
        Check("""<c:grouping val="clustered"/>""", grouping: PptxSceneChartGrouping.Clustered);
        Check("""<c:grouping val="stacked"/>""", grouping: PptxSceneChartGrouping.Stacked);
        Check("""<c:grouping val="bogus"/>""");
        Check("""<c:barDir val="bar"/>""", direction: PptxSceneChartBarDirection.Bar);
        Check("""<c:barDir val="col"/>""");
        Check("""<c:barDir val="bogus"/>""");
        Check("<c:varyColors/>", varyColors: true);
        Check("""<c:varyColors val="bogus"/>""", varyColors: false);
        Check("""<c:varyColors val="0"/>""", varyColors: false);
        Check("""<c:varyColors val="1"/>""", varyColors: true);
        Check("""<c:gapWidth val="150"/>""");
        Check("""<c:gapWidth val="999"/>""", gapWidth: 500);
        Check("""<c:gapWidth val="-5"/>""", gapWidth: 0);
        Check("""<c:overlap val="0"/>""");
        Check("""<c:overlap val="250"/>""", overlap: 100);
        Check("""<c:overlap val="-150"/>""", overlap: -100);
        Check("""
            <c:barDir val="bar"/><c:grouping val="stacked"/><c:varyColors val="0"/>
            <c:gapWidth val="999"/><c:overlap val="-150"/>
            """, PptxSceneChartGrouping.Stacked, PptxSceneChartBarDirection.Bar,
            varyColors: false, gapWidth: 500, overlap: -100);

        static void Check(string xml, PptxSceneChartGrouping grouping = PptxSceneChartGrouping.Standard,
            PptxSceneChartBarDirection direction = PptxSceneChartBarDirection.Column,
            bool? varyColors = null, double gapWidth = 150, double overlap = 0)
        {
            PptxSceneChartPlot plot = LoadPlot("barChart", xml);
            object options = Invoke("ReadSceneChartBarOptions", [typeof(PptxSceneChartPlot), typeof(PptxSceneChartGrouping)],
                [plot, PptxSceneChartGrouping.Standard])!;
            object Value(string name) => options.GetType().GetProperty(name)!.GetValue(options)!;
            TestAssert.Equal(grouping, (PptxSceneChartGrouping)Value("Grouping"));
            TestAssert.Equal(direction, (PptxSceneChartBarDirection)Value("BarDirection"));
            TestAssert.Equal(varyColors, plot.VaryColors);
            TestAssert.Equal(varyColors ?? true, (bool)Value("VaryColors"));
            TestAssert.Equal(gapWidth, (double)Value("GapWidth"));
            TestAssert.Equal(overlap, (double)Value("Overlap"));
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
            PptxSceneChartPlot plot = LoadPlot("scatterChart", style.Xml);
            object options = Invoke("ReadSceneChartScatterOptions", [typeof(PptxSceneChartPlot)], [plot])!;
            TestAssert.Equal(style.Expected, (PptxSceneChartScatterStyle)options.GetType().GetProperty("ScatterStyle")!.GetValue(options)!);
        }
    }

    public static void RadarStyleFollowsSourceDefaults()
    {
        (string Xml, PptxSceneChartRadarStyle Parsed, PptxSceneChartRadarStyle Effective)[] styles = [
            ("""<c:radarStyle val="standard"/>""", PptxSceneChartRadarStyle.Standard, PptxSceneChartRadarStyle.Standard),
            ("""<c:radarStyle val="marker"/>""", PptxSceneChartRadarStyle.Marker, PptxSceneChartRadarStyle.Marker),
            ("""<c:radarStyle val="filled"/>""", PptxSceneChartRadarStyle.Filled, PptxSceneChartRadarStyle.Filled),
            ("""<c:radarStyle val="bogus"/>""", PptxSceneChartRadarStyle.Unknown, PptxSceneChartRadarStyle.Standard),
            ("", PptxSceneChartRadarStyle.Unknown, PptxSceneChartRadarStyle.Standard)];
        foreach (var style in styles)
        {
            PptxSceneChartPlot plot = LoadPlot("radarChart", style.Xml);
            TestAssert.Equal(style.Parsed, plot.RadarStyleKind);
            object result = Invoke("ReadSceneChartRadarStyle", [typeof(PptxSceneChartPlot)], [plot])!;
            TestAssert.Equal(style.Effective, (PptxSceneChartRadarStyle)result);
        }
    }

    public static void PptxChartRadarUnknownAndMissingStylesPaintStandardSeries()
    {
        byte[] standard = Render("""<c:radarStyle val="standard"/>""");
        TestAssert.Equal(3, PptxTests.CountLineChartSeriesSegments(Encoding.ASCII.GetString(standard), "1 0 0 RG"));
        TestAssert.True(standard.AsSpan().SequenceEqual(Render("""<c:radarStyle val="bogus"/>""")),
            "An unrecognized radar style retains the standard series paint.");
        TestAssert.True(standard.AsSpan().SequenceEqual(Render("")),
            "An omitted radar style retains the standard series paint.");

        static byte[] Render(string style)
        {
            string xml = ChartSpace($"""
                <c:radarChart>{style}
                  <c:ser>
                    <c:spPr><a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></c:spPr>
                    <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt></c:strLit></c:cat>
                    <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt><c:pt idx="2"><c:v>4</c:v></c:pt></c:numLit></c:val>
                  </c:ser>
                </c:radarChart>
                """, "");
            using var input = PptxTests.CreateSingleChartPackageStream(xml);
            using var output = new MemoryStream();
            var diagnostics = new List<OoxPdfDiagnostic>();
            OoxPdfConverter.Convert(input, output, new OoxPdfOptions
            {
                InputKind = OoxPdfInputKind.Pptx,
                FontResolver = new PortableRenderingTests.PinnedResolver(),
                DiagnosticSink = diagnostics.Add
            });
            TestAssert.Equal(0, diagnostics.Count);
            return output.ToArray();
        }
    }

    public static void PptxSparseRadarSeriesKeepAuthoredStrokeAndFill()
    {
        var red = new RgbColor(0xA5, 0x15, 0x15);
        var green = new RgbColor(0x06, 0x76, 0x49);
        foreach (string style in new[] { "marker", "filled" })
        {
            string first = RadarSeries(0, "A51515", true);
            string emptyFirst = RadarSeries(0, "A51515", false);
            string second = RadarSeries(1, "067649", true);
            PdfPage full = RenderRadar(style, first + second);
            PdfPage sparse = RenderRadar(style, emptyFirst + second);
            PdfPage literal = RenderRadar(style, second);
            TestAssert.True(Paints(full, strokes: true).SequenceEqual([red, green]),
                "Both outlines keep their authored stroke in the populated plot.");
            TestAssert.True(Paints(sparse, strokes: true).SequenceEqual([green]),
                "Only the surviving series may paint its own outline.");
            TestAssert.True(Paints(full, strokes: false).SequenceEqual(style == "filled" ? [red, green] : []),
                "Only filled radars paint both authored series fills.");
            TestAssert.True(Paints(sparse, strokes: false).SequenceEqual(style == "filled" ? [green] : []),
                "Only the surviving filled series paints its own fill.");
            TestAssert.True(literal.Content == sparse.Content,
                "The sparse chart keeps the surviving series' polygon and grid geometry.");

            var diagnostics = new List<OoxPdfDiagnostic>();
            PdfPage empty = RenderRadar(style, emptyFirst + RadarSeries(1, "067649", false), sink: diagnostics.Add);
            TestAssert.Equal(0, Paints(empty, strokes: true).Count());
            TestAssert.Equal(0, Paints(empty, strokes: false).Count());
            TestAssert.Equal(1, diagnostics.Count(d => d.Id == "PPTX_CHART_MISSING_CACHED_DATA"));
        }

        IEnumerable<RgbColor> Paints(PdfPage page, bool strokes) =>
            PptxChartPaletteTests.ReadColoredPaints(page.Content, strokes: strokes)
                .Where(color => color == red || color == green);

    }

    public static void PptxRadarRightLegendPaintsCompleteNamesAndOwnedKeys()
    {
        const string right = """<c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>""";
        const string largeRight = """
            <c:legend><c:legendPos val="r"/><c:overlay val="0"/>
              <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1800"/></a:pPr></a:p></c:txPr>
            </c:legend>
            """;
        const string narrow = """
            <c:legend><c:legendPos val="r"/><c:layout><c:manualLayout>
              <c:wMode val="factor"/><c:w val="0.05"/>
            </c:manualLayout></c:layout><c:overlay val="0"/></c:legend>
            """;
        const string authoredPlot = """
            <c:layout><c:manualLayout><c:layoutTarget val="inner"/>
              <c:xMode val="edge"/><c:yMode val="edge"/><c:wMode val="factor"/><c:hMode val="factor"/>
              <c:x val="0.08"/><c:y val="0.1"/><c:w val="0.48"/><c:h val="0.7"/>
            </c:manualLayout></c:layout>
            """;
        string series = RadarLegendSeries();
        foreach (string style in new[] { "marker", "filled" })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            PdfPage omitted = RenderRadar(style, series, plotLayout: authoredPlot);
            PdfPage visible = RenderRadar(style, series, right, authoredPlot, diagnostics.Add);
            TestAssert.Equal(0, diagnostics.Count);
            AssertRadarLegendNames(visible);
            TestAssert.True(ReadRequiredRadarAxisCaptions(omitted).SequenceEqual(ReadRequiredRadarAxisCaptions(visible)),
                "A fitting legend preserves the authored plot and both sets of axis captions.");
            AssertRadarLegendKeys(visible, filled: style == "filled");
            TestAssert.DoesNotContain("/CL", omitted.Content);

            PdfPage automatic = RenderRadar(style, series, largeRight, sink: diagnostics.Add);
            TestAssert.Equal(0, diagnostics.Count);
            AssertRadarLegendNames(automatic);
            ReadRequiredRadarAxisCaptions(automatic);

            var refused = RenderRadar(style, series, narrow, authoredPlot, diagnostics.Add);
            TestAssert.True(refused.Content == omitted.Content, "A rejected legend preserves the authored chart paint.");
            TestAssert.True(diagnostics.Any(d => d.Id == "PPTX_UNSUPPORTED_CHART" && d.Fallback == "Legend ignored"),
                "A declared legend box that cannot contain its names has a diagnostic.");
        }
    }

    public static void PptxRadarLeftLegendUsesFittingSideAndAuthoredOverlayBoxes()
    {
        const string left = """
            <c:legend><c:legendPos val="l"/><c:layout><c:manualLayout>
              <c:xMode val="edge"/><c:yMode val="edge"/><c:wMode val="factor"/><c:hMode val="factor"/>
              <c:x val="0.04"/><c:y val="0.28"/><c:w val="0.35"/><c:h val="0.36"/>
            </c:manualLayout></c:layout><c:overlay val="0"/></c:legend>
            """;
        const string automatic = """<c:legend><c:legendPos val="l"/><c:overlay val="0"/></c:legend>""";
        const string authoredPlot = """
            <c:layout><c:manualLayout><c:layoutTarget val="inner"/>
              <c:xMode val="edge"/><c:yMode val="edge"/><c:wMode val="factor"/><c:hMode val="factor"/>
              <c:x val="0.50"/><c:y val="0.15"/><c:w val="0.42"/><c:h val="0.70"/>
            </c:manualLayout></c:layout>
            """;
        string overlappingPlot = authoredPlot.Replace("0.50", "0.08", StringComparison.Ordinal);
        string overlay = left.Replace("""<c:overlay val="0"/>""", """<c:overlay val="1"/>""", StringComparison.Ordinal);
        string rightOverlay = overlay.Replace("""<c:legendPos val="l"/>""", """<c:legendPos val="r"/>""", StringComparison.Ordinal);
        string series = RadarLegendSeries();

        foreach (string style in new[] { "marker", "filled" })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            PdfPage baseline = RenderRadar(style, series, plotLayout: authoredPlot);
            PdfPage fitted = RenderRadar(style, series, left, authoredPlot, diagnostics.Add);
            AssertRadarLegendNames(fitted);
            TestAssert.True(ReadRequiredRadarAxisCaptions(baseline).SequenceEqual(ReadRequiredRadarAxisCaptions(fitted)),
                "A left legend in a free band preserves the manually placed axis captions.");
            AssertRadarLegendKeys(fitted, filled: style == "filled");

            PdfPage overlaid = RenderRadar(style, series, overlay, overlappingPlot, diagnostics.Add);
            AssertRadarLegendNames(overlaid);
            PdfPage equivalent = RenderRadar(style, series, rightOverlay, overlappingPlot, diagnostics.Add);
            TestAssert.Equal(equivalent.Content, overlaid.Content);
            PdfPage automaticPage = RenderRadar(style, series, automatic, sink: diagnostics.Add);
            AssertRadarLegendNames(automaticPage);
            ReadRequiredRadarAxisCaptions(automaticPage);
            TestAssert.Equal(0, diagnostics.Count);

            PdfPage refused = RenderRadar(style, series, left, overlappingPlot, diagnostics.Add);
            PdfPage omitted = RenderRadar(style, series, plotLayout: overlappingPlot);
            TestAssert.Equal(omitted.Content, refused.Content);
            TestAssert.True(diagnostics.Any(d => d.Id == "PPTX_UNSUPPORTED_CHART" && d.Fallback == "Legend ignored"),
                "A non-overlay legend cannot displace a conflicting manual plot.");
        }
    }

    private static string RadarLegendSeries() =>
        RadarSeries(0, "A51515", true, "MeridianSales") + RadarSeries(1, "067649", true, "StationCredits");

    private static void AssertRadarLegendNames(PdfPage page)
    {
        var names = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, false,
            fontResourceFilter: name => name.StartsWith("CL", StringComparison.Ordinal)).ToArray();
        TestAssert.True(names.Select(show => show.Text).SequenceEqual(["MeridianSales", "StationCredits"]),
            "A fitting legend paints each authored name once in series order.");
        TestAssert.True(names[0].Y > names[1].Y, "Each legend name has its own row.");
    }

    private static void AssertRadarLegendKeys(PdfPage page, bool filled)
    {
        var red = new RgbColor(0xA5, 0x15, 0x15);
        var green = new RgbColor(0x06, 0x76, 0x49);
        RgbColor[] expected = [red, green, red, green];
        TestAssert.True(Paints(strokes: true).SequenceEqual(expected),
            "Radar legend keys match the authored series strokes.");
        TestAssert.True(Paints(strokes: false).SequenceEqual(filled ? expected : []),
            "Radar legend fill follows the authored series fill.");

        IEnumerable<RgbColor> Paints(bool strokes) => PptxChartPaletteTests.ReadColoredPaints(page.Content, strokes: strokes)
            .Where(color => color == red || color == green);
    }

    private static IReadOnlyList<(string Text, double X, double Y)> ReadRequiredRadarAxisCaptions(PdfPage page)
    {
        var captions = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true,
            fontResourceFilter: name => name.StartsWith("RCA", StringComparison.Ordinal) || name.StartsWith("RVA", StringComparison.Ordinal))
            .Select(show => (show.Text, show.X, show.Y)).ToArray();
        TestAssert.True(captions.Select(show => show.Text).Where(text => text is "A" or "B" or "C")
            .SequenceEqual(["A", "B", "C"]),
            "Each authored radar category caption appears once in source order.");
        TestAssert.True(captions.Any(show => show.Text == "0"),
            "The radar value axis retains its zero caption.");
        return captions;
    }

    public static void PptxRadarInnerPlotRetainsCompleteCategoryInk()
    {
        const string plot = """
            <c:layout><c:manualLayout><c:layoutTarget val="inner"/>
              <c:xMode val="edge"/><c:yMode val="edge"/><c:wMode val="factor"/><c:hMode val="factor"/>
              <c:x val="0.20"/><c:y val="0.19"/><c:w val="0.60"/><c:h val="0.62"/>
            </c:manualLayout></c:layout>
            """;
        const string series = """
            <c:ser><c:idx val="0"/><c:order val="0"/>
              <c:cat><c:strLit>
                <c:pt idx="0"><c:v>Speed</c:v></c:pt><c:pt idx="1"><c:v>Power</c:v></c:pt>
                <c:pt idx="2"><c:v>Agility</c:v></c:pt><c:pt idx="3"><c:v>Defense</c:v></c:pt>
                <c:pt idx="4"><c:v>Stamina</c:v></c:pt>
              </c:strLit></c:cat>
              <c:val><c:numLit><c:ptCount val="5"/>
                <c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>5</c:v></c:pt>
                <c:pt idx="2"><c:v>4</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt>
                <c:pt idx="4"><c:v>2</c:v></c:pt>
              </c:numLit></c:val>
            </c:ser>
            """;
        foreach (string style in new[] { "marker", "filled" })
        {
            var page = RenderRadar(style, series, plotLayout: plot);
            foreach (string label in new[] { "Speed", "Power", "Agility", "Defense", "Stamina" })
                PortableRenderingTests.ReadPaintedSourceLabel(page, label);
        }
    }

    private static string RadarSeries(int index, string color, bool hasCache, string? name = null) => $"""
            <c:ser><c:idx val="{index}"/><c:order val="{index}"/>
              {(name is null ? "" : $"<c:tx><c:v>{name}</c:v></c:tx>")}
              <c:spPr><a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:ln w="25400">
                <a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:ln></c:spPr>
              <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt></c:strLit></c:cat>
              <c:val><c:numLit><c:ptCount val="{(hasCache ? 3 : 0)}"/>
                {(hasCache ? """<c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>5</c:v></c:pt><c:pt idx="2"><c:v>4</c:v></c:pt>""" : "")}
              </c:numLit></c:val>
            </c:ser>
            """;

    private static PdfPage RenderRadar(string style, string series, string legend = "", string plotLayout = "", Action<OoxPdfDiagnostic>? sink = null)
    {
        string xml = ChartSpace($"{plotLayout}<c:radarChart><c:radarStyle val=\"{style}\"/>{series}</c:radarChart>", legend);
        using var stream = PptxTests.CreateSingleChartPackageStream(xml);
        using var package = OoxPackage.Open(stream, default);
        return new PptxRenderer(new PortableRenderingTests.PinnedResolver())
            .RenderPagesAsync(new PptxReader().Read(package, default), package, sink, default)
            .ToBlockingEnumerable().Single();
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
            PptxSceneChartPlot plot = LoadPlot("barChart", plotForm.Xml, seriesForm.Xml);
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
            PptxSceneChartPlot plot = LoadPlot("barChart", "", form.Xml);
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
            PptxSceneChartPlot plot = LoadPlot("barChart", "", test.Xml);
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
            PptxSceneChartPlot plot = LoadPlot("pieChart", "", test.Xml);
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
        const string plot = """
            <c:barChart>
              <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
              <c:axId val="10"/><c:axId val="20"/>
            </c:barChart>
            """;
        (string Xml, PptxSceneChartDisplayBlanksAs Blanks, bool? VisibleOnly)[] samples =
        [
            ("<c:dispBlanksAs val=\"span\"/><c:plotVisOnly val=\"0\"/>", PptxSceneChartDisplayBlanksAs.Span, false),
            ("<c:dispBlanksAs val=\"zero\"/><c:plotVisOnly val=\"1\"/>", PptxSceneChartDisplayBlanksAs.Zero, true),
            ("<c:dispBlanksAs val=\"gap\"/><c:plotVisOnly/>", PptxSceneChartDisplayBlanksAs.Gap, true),
            ("<c:dispBlanksAs val=\"bogus\"/><c:plotVisOnly val=\"bogus\"/>", PptxSceneChartDisplayBlanksAs.Gap, false),
            ("", PptxSceneChartDisplayBlanksAs.Gap, null)
        ];
        foreach (var sample in samples)
        {
            PptxSceneChart sceneChart = TestAssert.NotNull(PptxTests.BuildSingleChartScene(ChartSpace(plot, sample.Xml)));
            object? effective = Invoke("ResolveChartDisplayBlanksAs", new[] { typeof(PptxSceneChart) }, new object?[] { sceneChart });
            TestAssert.Equal(sample.Blanks, (PptxSceneChartDisplayBlanksAs)TestAssert.NotNull(effective));
            TestAssert.Equal(sample.VisibleOnly, sceneChart.Options.PlotVisibleOnly);
        }
    }

    public static void ScenePlotsCorrespondToXmlPlotElementsAcrossFixtures()
    {
        // Every plotArea chart child keeps one scene owner in source order.
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

    private static PptxSceneChartPlot LoadPlot(string plotKind, string plotInner, string seriesInner = "")
    {
        string xml = ChartSpace("<c:" + plotKind + ">" + plotInner + "<c:ser>" + seriesInner + "<c:cat><c:strLit><c:pt idx=\"0\"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx=\"0\"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val=\"10\"/><c:axId val=\"20\"/></c:" + plotKind + ">", "");
        PptxSceneChart chart = TestAssert.NotNull(PptxTests.BuildSingleChartScene(xml));
        TestAssert.Equal(1, chart.Plots.Count);
        return chart.Plots[0];
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
