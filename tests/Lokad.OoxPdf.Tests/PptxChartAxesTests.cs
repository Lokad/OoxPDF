using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxChartAxesTests
{
    public static void PptxVerticalAxisMeasurementMatchesEquivalentAuthoredIntervals()
    {
        foreach (string kind in new[] { "column", "stacked", "dual-stacked", "line", "right-line", "right-area" })
        foreach (double maximum in new[] { .5d, 5d, 50d })
        {
            PdfPage automatic = Render(authoredUnit: false);
            PdfPage authored = Render(authoredUnit: true);
            string[] autoLabels = Labels(automatic), authoredLabels = Labels(authored);
            TestAssert.True(autoLabels.Length >= 11 && autoLabels.Contains((maximum / 10d).ToString("G", CultureInfo.InvariantCulture)),
                "The tall fixture must actually emit the selected automatic interval.");
            TestAssert.True(autoLabels.SequenceEqual(authoredLabels), "Equivalent source intervals must emit the same numeric labels.");
            string[] autoClips = Clips(automatic), authoredClips = Clips(authored);
            TestAssert.True(autoClips.Length >= 2, "Expected frame and data plot clips.");
            TestAssert.True(autoClips.SequenceEqual(authoredClips),
                $"Measuring emitted {kind} labels must preserve plot geometry for equivalent intervals on0..{maximum}.");

            PdfPage Render(bool authoredUnit)
            {
                string N(double value) => value.ToString("R", CultureInfo.InvariantCulture);
                string Series(int index, double value) => $$"""
                    <c:ser><c:idx val="{{index}}"/><c:order val="{{index}}"/><c:tx><c:v>Series{{index}}</c:v></c:tx>
                      <c:cat><c:strLit><c:pt idx="0"><c:v>Alfa</c:v></c:pt><c:pt idx="1"><c:v>Beta</c:v></c:pt></c:strLit></c:cat>
                      <c:val><c:numLit><c:pt idx="0"><c:v>{{N(value)}}</c:v></c:pt><c:pt idx="1"><c:v>{{N(value)}}</c:v></c:pt></c:numLit></c:val>
                    </c:ser>
                    """;
                string Axis(int id, string side) => $$"""
                    <c:valAx><c:axId val="{{id}}"/><c:scaling><c:orientation val="minMax"/><c:min val="0"/><c:max val="{{N(maximum)}}"/></c:scaling>
                      <c:axPos val="{{side}}"/><c:crossAx val="1"/>{{(authoredUnit ? "<c:majorUnit val=\"" + N(maximum / 10d) + "\"/>" : "")}}
                      <c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="1000"/></a:pPr></a:p></c:txPr>
                    </c:valAx>
                    """;
                bool stacked = kind is "stacked" or "dual-stacked";
                string plotKind = kind is "line" or "right-line" ? "line" : kind == "right-area" ? "area" : "bar";
                string options = (plotKind == "bar" ? "<c:barDir val=\"col\"/>" : "") +
                    "<c:grouping val=\"" + (stacked ? "stacked" : plotKind == "bar" ? "clustered" : "standard") + "\"/>";
                string series = stacked ? Series(0, maximum * .4d) + Series(1, maximum * .4d) : Series(0, maximum * .8d);
                string plot = $"<c:{plotKind}Chart>{options}{series}<c:axId val=\"1\"/><c:axId val=\"2\"/></c:{plotKind}Chart>";
                if (kind == "dual-stacked") plot += $"<c:barChart>{options}{Series(2, maximum * .3d)}{Series(3, maximum * .3d)}<c:axId val=\"1\"/><c:axId val=\"3\"/></c:barChart>";
                string chart = $$"""
                    <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:autoTitleDeleted val="1"/><c:plotArea>
                      {{plot}}<c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/><c:crosses val="autoZero"/></c:catAx>{{Axis(2, "l")}}{{(kind == "dual-stacked" ? Axis(3, "r") : "")}}
                    </c:plotArea>{{(kind.StartsWith("right-", StringComparison.Ordinal) ? "<c:legend><c:legendPos val=\"r\"/></c:legend>" : "")}}</c:chart></c:chartSpace>
                    """;
                const string slide = """
                    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><p:cSld><p:spTree><p:graphicFrame><p:xfrm><a:off x="914400" y="1371600"/><a:ext cx="5486400" cy="4572000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>
                    """;
                using var stream = PptxTests.CreateSingleChartPackageStream(chart, new Dictionary<string, byte[]> { ["ppt/slides/slide1.xml"] = TestFixtures.Utf8(slide) });
                using var package = OoxPackage.Open(stream, default);
                var document = new PptxReader().Read(package, default);
                return new PptxRenderer(new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
            }
        }

        static string[] Labels(PdfPage page)
        {
            var fonts = page.Fonts.Where(font => font.ResourceName.StartsWith("CVA", StringComparison.Ordinal)).Select(font => font.Font).ToHashSet();
            return DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).Where(show => fonts.Contains(show.Font)).Select(show => show.Text).Order(StringComparer.Ordinal).ToArray();
        }

        static string[] Clips(PdfPage page) => Regex.Matches(page.Content, @"([-\d.]+) ([-\d.]+) ([\d.]+) ([\d.]+) re W\* n")
            .Where(m => double.Parse(m.Groups[3].Value, CultureInfo.InvariantCulture) > 150d && double.Parse(m.Groups[4].Value, CultureInfo.InvariantCulture) > 50d)
            .Select(m => m.Value).Distinct(StringComparer.Ordinal).Order(StringComparer.Ordinal).ToArray();
    }

    public static void PptxColumnAxisPositionPreservesTheDataZeroOrigin()
    {
        foreach (bool negative in new[] { false, true })
        foreach (bool reversed in new[] { false, true })
        {
            double[][] bottom = Bars(RenderColumnDirectionProbe(negative, reversed, "none", "b"));
            double[][] top = Bars(RenderColumnDirectionProbe(negative, reversed, "none", "t"));
            TestAssert.Equal(2, bottom.Length);
            TestAssert.Equal(2, top.Length);
            for (int i = 0; i < 2; i++)
            {
                // Authored216pt frame, half-height inner plot, value4/range5.
                TestAssert.True(Math.Abs(top[i][3] - 86.4d) <= .001d,
                    "Axis position metadata must not replace zero with the maximum data origin.");
                for (int coordinate = 0; coordinate < 4; coordinate++)
                    TestAssert.True(Math.Abs(top[i][coordinate] - bottom[i][coordinate]) <= .001d,
                        "Changing category axis position must preserve every data rectangle.");
            }
        }

        static double[][] Bars(PdfPage page) => Regex.Matches(page.Content, @"([-\d.]+) ([-\d.]+) ([\d.]+) ([\d.]+) re f")
            .Select(m => m.Groups.Cast<Group>().Skip(1).Select(g => double.Parse(g.Value, CultureInfo.InvariantCulture)).ToArray())
            .Where(rect => rect[2] > 1d && rect[2] < 100d && rect[3] > 1d).ToArray();
    }

    public static void PptxColumnCaptionsFollowNumericPositionAndValueDirection()
    {
        foreach (bool negative in new[] { false, true })
        foreach (bool reversed in new[] { false, true })
        foreach (string position in new[] { "nextTo", "low", "high" })
        {
            PdfPage page = RenderColumnDirectionProbe(negative, reversed, position, "b");
            var bars = Regex.Matches(page.Content, @"([-\d.]+) ([-\d.]+) ([\d.]+) ([\d.]+) re f")
                .Select(m => m.Groups.Cast<Group>().Skip(1).Select(g => double.Parse(g.Value, CultureInfo.InvariantCulture)).ToArray())
                .Where(rect => rect[2] > 1d && rect[2] < 100d && rect[3] > 1d).ToArray();
            TestAssert.Equal(2, bars.Length);
            double plotHeight = bars[0][3] * 5d / 4d;
            bool zeroAtTop = negative != reversed;
            double plotTop = bars[0][1] + (zeroAtTop ? bars[0][3] : plotHeight);
            double plotBottom = plotTop - plotHeight;
            double anchor = position == "low" ? (reversed ? plotTop : plotBottom) :
                position == "high" ? (reversed ? plotBottom : plotTop) : zeroAtTop ? plotTop : plotBottom;
            bool above = (position == "high") != reversed;
            var captions = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
                fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(2, captions.Length);
            foreach (var show in captions)
            {
                double bottom = double.PositiveInfinity, top = double.NegativeInfinity;
                foreach (char character in show.Text)
                {
                    TestAssert.True(show.Font.Font.TryGetGlyphBounds(show.Font.Font.MapCodePoint(character), out var box), "Expected pinned glyph ink.");
                    bottom = Math.Min(bottom, show.Y + box.YMin * show.Size / show.Font.Font.UnitsPerEm);
                    top = Math.Max(top, show.Y + box.YMax * show.Size / show.Font.Font.UnitsPerEm);
                }
                TestAssert.True(above ? bottom > anchor : top < anchor,
                    $"Caption ink must follow numeric {position} placement (negative={negative}, reversed={reversed}).");
            }
        }
    }

    public static void PptxLineCaptionPlacementSurvivesTitlesAndBottomLegends()
    {
        foreach (bool negative in new[] { false, true })
        foreach (bool reversed in new[] { false, true })
        foreach (string position in new[] { "nextTo", "low", "high" })
        foreach (string furniture in new[] { "none", "title", "legend", "both" })
        {
            PdfPage page = RenderColumnDirectionProbe(negative, reversed, position, "b", line: true,
                title: furniture is "title" or "both", bottomLegend: furniture is "legend" or "both");
            double[] plot = Regex.Matches(page.Content, @"([-\d.]+) ([-\d.]+) ([\d.]+) ([\d.]+) re W\* n")
                .Select(m => m.Groups.Cast<Group>().Skip(1).Select(g => double.Parse(g.Value, CultureInfo.InvariantCulture)).ToArray())
                .Where(rect => rect[2] > 150d && rect[2] < 288d && rect[3] > 50d && rect[3] < 216d)
                .OrderByDescending(rect => rect[3]).First();
            bool zeroAtTop = negative != reversed;
            double ratio = position == "low" ? (reversed ? 1d : 0d) :
                position == "high" ? (reversed ? 0d : 1d) : zeroAtTop ? 1d : 0d;
            double anchor = plot[1] + plot[3] * ratio;
            bool above = (position == "high") != reversed;
            var captions = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
                fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(2, captions.Length);
            TestAssert.True(captions.Select(show => show.Text).Order(StringComparer.Ordinal).SequenceEqual(new[] { "Alfa", "Beta" }), "Keep both source captions.");
            foreach (var show in captions)
            {
                double bottom = double.PositiveInfinity, top = double.NegativeInfinity;
                foreach (char character in show.Text)
                {
                    TestAssert.True(show.Font.Font.TryGetGlyphBounds(show.Font.Font.MapCodePoint(character), out var box), "Expected pinned glyph ink.");
                    bottom = Math.Min(bottom, show.Y + box.YMin * show.Size / show.Font.Font.UnitsPerEm);
                    top = Math.Max(top, show.Y + box.YMax * show.Size / show.Font.Font.UnitsPerEm);
                }
                TestAssert.True(above ? bottom > anchor : top < anchor,
                    $"Line captions must retain numeric {position} placement with {furniture} (negative={negative}, reversed={reversed}).");
            }
        }
    }

    public static void PptxLineTitleKeepsNumericFootClearanceSeparateFromHeader()
    {
        foreach (bool legend in new[] { false, true })
        {
            PdfPage bare = RenderColumnDirectionProbe(false, true, "low", "b", line: true, bottomLegend: legend);
            PdfPage titled = RenderColumnDirectionProbe(false, true, "low", "b", line: true, title: true, bottomLegend: legend);
            double[] before = Plot(bare), after = Plot(titled);
            TestAssert.True(Math.Abs(before[1] - after[1]) <= .001d,
                "A title above top-side captions must not also consume numeric clearance at the foot.");
            TestAssert.True(after[3] < before[3], "The title still owns space at the top of the frame.");
        }

        static double[] Plot(PdfPage page) => Regex.Matches(page.Content, @"([-\d.]+) ([-\d.]+) ([\d.]+) ([\d.]+) re W\* n")
            .Select(m => m.Groups.Cast<Group>().Skip(1).Select(g => double.Parse(g.Value, CultureInfo.InvariantCulture)).ToArray())
            .Where(rect => rect[2] > 150d && rect[2] < 288d && rect[3] > 50d && rect[3] < 216d)
            .OrderByDescending(rect => rect[3]).First();
    }

    private static PdfPage RenderColumnDirectionProbe(bool negative, bool reversed, string position, string axisPosition,
        bool line = false, bool title = false, bool bottomLegend = false)
    {
        string chart = $$"""
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea>
              <c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/><c:wMode val="factor"/><c:hMode val="factor"/><c:x val="0.1"/><c:y val="0.1"/><c:w val="0.8"/><c:h val="0.5"/></c:manualLayout></c:layout>
              <c:barChart><c:barDir val="col"/><c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>Alfa</c:v></c:pt><c:pt idx="1"><c:v>Beta</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>{{(negative ? -4 : 4)}}</c:v></c:pt><c:pt idx="1"><c:v>{{(negative ? -4 : 4)}}</c:v></c:pt></c:numLit></c:val>
              </c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart>
              <c:catAx><c:axId val="1"/><c:axPos val="{{axisPosition}}"/><c:tickLblPos val="{{position}}"/>
                <c:txPr><a:bodyPr rot="0"/><a:p><a:pPr><a:defRPr sz="1000"/></a:pPr></a:p></c:txPr>
              </c:catAx>
              <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="{{(reversed ? "maxMin" : "minMax")}}"/><c:min val="{{(negative ? -5 : 0)}}"/><c:max val="{{(negative ? 0 : 5)}}"/></c:scaling><c:tickLblPos val="none"/></c:valAx>
            </c:plotArea></c:chart></c:chartSpace>
            """;
        if (line)
        {
            XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            XDocument parsed = XDocument.Parse(chart);
            XElement area = parsed.Descendants(c + "plotArea").Single();
            area.Element(c + "layout")!.Remove();
            XElement plot = area.Element(c + "barChart")!;
            plot.Name = c + "lineChart";
            plot.Element(c + "barDir")!.Remove();
            area.Element(c + "catAx")!.Add(new XElement(c + "crossAx", new XAttribute("val", "2")), new XElement(c + "crosses", new XAttribute("val", "autoZero")));
            area.Element(c + "valAx")!.Add(new XElement(c + "axPos", new XAttribute("val", "l")), new XElement(c + "crossAx", new XAttribute("val", "1")));
            if (title) parsed.Descendants(c + "chart").Single().AddFirst(new XElement(c + "title", new XElement(c + "tx", new XElement(c + "rich",
                new XElement(a + "bodyPr"), new XElement(a + "p", new XElement(a + "r", new XElement(a + "t", "Direction title")))))));
            if (bottomLegend) parsed.Descendants(c + "chart").Single().Add(new XElement(c + "legend", new XElement(c + "legendPos", new XAttribute("val", "b"))));
            chart = parsed.ToString(SaveOptions.DisableFormatting);
        }
        const string slide = """
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><p:cSld><p:spTree><p:graphicFrame><p:xfrm><a:off x="914400" y="1371600"/><a:ext cx="3657600" cy="2743200"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>
            """;
        using var stream = PptxTests.CreateSingleChartPackageStream(chart, new Dictionary<string, byte[]> { ["ppt/slides/slide1.xml"] = TestFixtures.Utf8(slide) });
        using var package = OoxPackage.Open(stream, default);
        var document = new PptxReader().Read(package, default);
        return new PptxRenderer(new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
    }

    public static void PptxNegativeOnlyAxesKeepOfficeRangeAndAuthoredBounds()
    {
        // Office16 columns, horizontal bars and lines agree on these bounds;
        // the retained scaled/partial-bound probes exercise the same reflection.
        foreach (string kind in new[] { "column", "bar", "line" })
        foreach (var bounds in new[]
        {
            (Minimum: (double?)null, Maximum: (double?)null, ExpectedMin: -4.5d, ExpectedMax: 0d),
            (Minimum: (double?)-5d, Maximum: (double?)null, ExpectedMin: -5d, ExpectedMax: 0d),
            (Minimum: (double?)null, Maximum: (double?)-1d, ExpectedMin: -4.5d, ExpectedMax: -1d),
            (Minimum: (double?)null, Maximum: (double?)1d, ExpectedMin: -5d, ExpectedMax: 1d),
            (Minimum: (double?)-5d, Maximum: (double?)1d, ExpectedMin: -5d, ExpectedMax: 1d)
        })
        foreach (double scale in new[] { 1d, .1d, .01d })
        {
            string N(double value) => (value * scale).ToString("R", CultureInfo.InvariantCulture);
            string axisBounds = (bounds.Minimum is { } min ? $"<c:min val=\"{N(min)}\"/>" : "") +
                (bounds.Maximum is { } max ? $"<c:max val=\"{N(max)}\"/>" : "");
            string plot = kind == "line" ? "line" : "bar";
            string direction = kind == "line" ? "" : $"<c:barDir val=\"{(kind == "bar" ? "bar" : "col")}\"/>";
            string manual = kind == "line" ? "<c:layout><c:manualLayout><c:layoutTarget val=\"inner\"/><c:x val=\"0.1\"/><c:y val=\"0.1\"/><c:w val=\"0.8\"/><c:h val=\"0.7\"/></c:manualLayout></c:layout>" : "";
            string chart = $$"""
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
                  <c:chart><c:plotArea>{{manual}}<c:{{plot}}Chart>{{direction}}<c:ser>
                    <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                    <c:val><c:numLit><c:pt idx="0"><c:v>{{N(-4d)}}</c:v></c:pt><c:pt idx="1"><c:v>{{N(-4d)}}</c:v></c:pt></c:numLit></c:val>
                  </c:ser><c:axId val="1"/><c:axId val="2"/></c:{{plot}}Chart>
                  <c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
                  <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="{{(kind == "bar" ? "maxMin" : "minMax")}}"/>{{axisBounds}}</c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
                  </c:plotArea></c:chart>
                </c:chartSpace>
                """;
            using var stream = PptxTests.CreateSingleChartPackageStream(chart);
            using var package = OoxPackage.Open(stream, default);
            var document = new PptxReader().Read(package, default);
            PdfPage page = new PptxRenderer(new PortableRenderingTests.PinnedResolver())
                .RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
            var fonts = page.Fonts.Where(font => font.ResourceName.StartsWith("CVA", StringComparison.Ordinal)).Select(font => font.Font).ToHashSet();
            double[] labels = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).Where(show => fonts.Contains(show.Font))
                .Select(show => double.Parse(show.Text, CultureInfo.InvariantCulture)).ToArray();
            TestAssert.True(labels.Length >= 2, "Expected both value-axis endpoints.");
            TestAssert.True(Math.Abs(labels.Min() - bounds.ExpectedMin * scale) < 1e-9 && Math.Abs(labels.Max() - bounds.ExpectedMax * scale) < 1e-9,
                $"Negative-only {kind} must retain Office bounds with source {bounds}, scale {scale}; got {labels.Min()}..{labels.Max()}.");
        }
    }

    public static void PptxChartRangeConsolidationKeepsSceneAuthorityAndXmlFallback()
    {
        var resolve = typeof(PptxRenderer).GetMethod("ReadSceneOrXmlChartValueAxisExtents",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;
        var fallback = new PptxRenderer.ChartValueExtents(-4d, 0d);
        XElement Xml(string scaling) => XElement.Parse($"<c:valAx xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">{scaling}</c:valAx>");
        XElement contradictory = Xml("<c:scaling><c:min val=\"-100\"/><c:max val=\"100\"/></c:scaling>");
        foreach (var sample in new[]
        {
            (Scaling: "<c:scaling/>", Minimum: -4.5d, Maximum: 0d),
            (Scaling: "<c:scaling><c:min val=\"-6\"/></c:scaling>", Minimum: -6d, Maximum: 0d),
            (Scaling: "", Minimum: -4d, Maximum: 0d),
            (Scaling: "<c:scaling><c:min val=\"-1\"/><c:max val=\"-5\"/></c:scaling>", Minimum: -4d, Maximum: 0d)
        })
        {
            PptxSceneChart chart = PptxTests.BuildSingleChartScene($$"""
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea>
                  <c:lineChart><c:axId val="2"/></c:lineChart><c:valAx><c:axId val="2"/>{{sample.Scaling}}</c:valAx>
                </c:plotArea></c:chart></c:chartSpace>
                """)!;
            var sceneResult = (PptxRenderer.ChartValueExtents)resolve.Invoke(null, [chart.Axes.Single(), contradictory, fallback, true, .96d, 9d, false])!;
            var xmlResult = (PptxRenderer.ChartValueExtents)resolve.Invoke(null, [null, Xml(sample.Scaling), fallback, true, .96d, 9d, false])!;
            TestAssert.Equal(new PptxRenderer.ChartValueExtents(sample.Minimum, sample.Maximum), sceneResult);
            TestAssert.Equal(sceneResult, xmlResult);
        }
        var explicitXml = (PptxRenderer.ChartValueExtents)resolve.Invoke(null, [null, contradictory, fallback, true, .96d, 9d, false])!;
        TestAssert.Equal(new PptxRenderer.ChartValueExtents(-100d, 100d), explicitXml);
    }

    public static void PptxAuthoredAxisGeometrySurvivesChangesOfUnits()
    {
        foreach (var sample in new[]
        {
            (Kind: "bar", Direction: "col", Min: 0d, Max: 5d, Reversed: false, Units: false),
            (Kind: "bar", Direction: "col", Min: -5d, Max: 0d, Reversed: false, Units: true),
            (Kind: "bar", Direction: "col", Min: -2d, Max: 3d, Reversed: true, Units: false),
            (Kind: "bar", Direction: "bar", Min: 0d, Max: 5d, Reversed: true, Units: true),
            (Kind: "line", Direction: "", Min: 0d, Max: 5d, Reversed: false, Units: true),
            (Kind: "area", Direction: "", Min: -2d, Max: 3d, Reversed: false, Units: false),
            (Kind: "scatter", Direction: "", Min: 0d, Max: 5d, Reversed: false, Units: true),
            (Kind: "scatter", Direction: "", Min: -2d, Max: 3d, Reversed: false, Units: false),
            (Kind: "bubble", Direction: "", Min: 0d, Max: 5d, Reversed: false, Units: false)
        })
        {
            double[] original = Render(1d);
            TestAssert.True(original.Length > 40, "Expected painted chart geometry.");
            foreach (double scale in new[] { .1d, .01d })
            {
                double[] scaled = Render(scale);
                TestAssert.True(original.Length == scaled.Length,
                    $"Changing units must preserve {sample} operation count at scale {scale}: {original.Length} versus {scaled.Length}.");
                for (int i = 0; i < original.Length; i++)
                    TestAssert.True(Math.Abs(original[i] - scaled[i]) <= .002d,
                        $"Changing units must preserve {sample.Kind}/{sample.Direction} geometry (scale {scale}, coordinate {i}).");
            }

            double[] Render(double scale)
            {
                string N(double value) => (value * scale).ToString("R", CultureInfo.InvariantCulture);
                string values = $"<c:pt idx=\"0\"><c:v>{N(sample.Min + 1d)}</c:v></c:pt><c:pt idx=\"1\"><c:v>{N(sample.Max - 1d)}</c:v></c:pt>";
                bool numericX = sample.Kind is "scatter" or "bubble";
                string source = numericX
                    ? $"<c:xVal><c:numLit>{values}</c:numLit></c:xVal><c:yVal><c:numLit>{values}</c:numLit></c:yVal>"
                    : $"<c:cat><c:strLit><c:pt idx=\"0\"><c:v>A</c:v></c:pt><c:pt idx=\"1\"><c:v>B</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit>{values}</c:numLit></c:val>";
                if (sample.Kind == "bubble") source += "<c:bubbleSize><c:numLit><c:pt idx=\"0\"><c:v>1</c:v></c:pt><c:pt idx=\"1\"><c:v>4</c:v></c:pt></c:numLit></c:bubbleSize>";
                string units = sample.Units ? $"<c:majorUnit val=\"{N(2d)}\"/><c:minorUnit val=\"{N(.25d)}\"/>" : "";
                string Axis(string id, string position) => $$"""
                    <c:valAx><c:axId val="{{id}}"/><c:scaling><c:orientation val="{{(sample.Reversed ? "maxMin" : "minMax")}}"/><c:min val="{{N(sample.Min)}}"/><c:max val="{{N(sample.Max)}}"/></c:scaling>
                      <c:axPos val="{{position}}"/><c:majorGridlines/><c:minorGridlines/><c:tickLblPos val="none"/>{{units}}
                    </c:valAx>
                    """;
                string firstAxis = numericX ? Axis("1", "b") : "<c:catAx><c:axId val=\"1\"/><c:axPos val=\"b\"/><c:tickLblPos val=\"none\"/></c:catAx>";
                string options = sample.Kind == "bar" ? $"<c:barDir val=\"{sample.Direction}\"/>" : sample.Kind == "scatter" ? "<c:scatterStyle val=\"lineMarker\"/>" : "";
                string chart = $$"""
                    <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                      <c:chart><c:plotArea><c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/><c:wMode val="factor"/><c:hMode val="factor"/><c:x val="0.1"/><c:y val="0.1"/><c:w val="0.8"/><c:h val="0.7"/></c:manualLayout></c:layout>
                        <c:{{sample.Kind}}Chart>{{options}}<c:ser><c:idx val="0"/><c:order val="0"/>
                          <c:spPr><a:solidFill><a:srgbClr val="CC2200"/></a:solidFill><a:ln w="25400"><a:solidFill><a:srgbClr val="CC2200"/></a:solidFill></a:ln></c:spPr>{{source}}
                        </c:ser><c:axId val="1"/><c:axId val="2"/></c:{{sample.Kind}}Chart>{{firstAxis}}{{Axis("2", "l")}}
                      </c:plotArea></c:chart>
                    </c:chartSpace>
                    """;
                using var stream = PptxTests.CreateSingleChartPackageStream(chart);
                using var package = OoxPackage.Open(stream, default);
                var document = new PptxReader().Read(package, default);
                PdfPage page = new PptxRenderer(new PortableRenderingTests.PinnedResolver())
                    .RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
                // Compare painted paths. Numeric label text changes with units,
                // so its glyph operands and text clipping boxes are not geometry
                // invariants (scatter/bubble currently emit those labels here).
                string paths = string.Join("\n", page.Content.Split('\n').Where(line =>
                    Regex.IsMatch(line, @"(?:^| )(?:m|l|c|re)(?: [fS]\*?)?\r?$")));
                return Regex.Matches(paths, @"-?\d+(?:\.\d+)?")
                    .Select(match => double.Parse(match.Value, CultureInfo.InvariantCulture)).ToArray();
            }
        }
    }

    public static void PptxAutomaticIntervalsScaleWithSubUnitBounds()
    {
        double[] ticks = ReadBoundedAxisTicks(-2d, 3d, null).ToArray();
        double[] minors = ReadMinorGridlineValues(5d, null, null).ToArray();
        foreach (double scale in new[] { .1d, .01d })
        {
            double[] scaledTicks = ReadBoundedAxisTicks(-2d * scale, 3d * scale, null).ToArray();
            double[] scaledMinors = ReadMinorGridlineValues(5d * scale, null, null).ToArray();
            TestAssert.Equal(ticks.Length, scaledTicks.Length);
            TestAssert.Equal(minors.Length, scaledMinors.Length);
            for (int i = 0; i < ticks.Length; i++) TestAssert.True(Math.Abs(ticks[i] * scale - scaledTicks[i]) < 1e-10, "Automatic major intervals must scale with their resolved bounds.");
            for (int i = 0; i < minors.Length; i++) TestAssert.True(Math.Abs(minors[i] * scale - scaledMinors[i]) < 1e-10, "Automatic minor intervals must subdivide the same scaled major interval.");
        }
        TestAssert.Equal(0, ReadBoundedAxisTicks(0d, 0d, null, includeEndpoints: false).Count);
    }

    public static void PptxCrowdedCategoryCaptionsPreservePlotSpaceAcrossFrameSizes()
    {
        foreach (int height in new[] { 120, 240, 420 })
        foreach (int rotation in new[] { -90, -45, 45 })
        {
            var automatic = Render(height, rotation, manual: false, longNames: true);
            double plotHeight = automatic.Bars[0][3] * 5d / 4d; // Authored0..5, all values4.
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(automatic.Page).ToArray();
            TestAssert.Equal(12, shows.Length);
            foreach (var show in shows)
            {
                double sin = Math.Sin(-rotation * Math.PI / 180d), cos = Math.Cos(rotation * Math.PI / 180d);
                double cursor = 0d, bottom = double.PositiveInfinity, top = double.NegativeInfinity;
                foreach (Rune rune in show.Text.EnumerateRunes())
                {
                    var font = show.Font.Font;
                    ushort glyph = font.MapCodePoint(rune.Value);
                    TestAssert.True(font.TryGetGlyphBounds(glyph, out var box), "Expected pinned caption glyph bounds.");
                    double scale = show.Size / font.UnitsPerEm;
                    foreach (double x in new[] { box.XMin * scale, box.XMax * scale })
                    foreach (double y in new[] { box.YMin * scale, box.YMax * scale })
                    {
                        double projected = (cursor + x) * sin + y * cos;
                        bottom = Math.Min(bottom, projected);
                        top = Math.Max(top, projected);
                    }
                    // Positions use the authored10pt advance; glyph ink uses
                    // the PDF's emitted size, which can be rounded separately.
                    cursor += font.GetAdvanceWidth(glyph) * 10d / font.UnitsPerEm;
                }
                TestAssert.True(top - bottom <= plotHeight + 0.1d,
                    "Caption ink must not consume more height than its plot in an automatic frame.");
            }
            var manualLong = Render(height, rotation, manual: true, longNames: true);
            var manualShort = Render(height, rotation, manual: true, longNames: false);
            TestAssert.True(manualLong.Bars.SelectMany(bar => bar).SequenceEqual(manualShort.Bars.SelectMany(bar => bar)),
                "Changing captions must preserve the authored plot rectangle and every source bar.");
        }

        static (PdfPage Page, double[][] Bars) Render(int height, int rotation, bool manual, bool longNames)
        {
            string names = string.Concat(Enumerable.Range(0, 12).Select(i =>
                $"<c:pt idx=\"{i}\"><c:v>{(longNames ? "NorthWarehouseDistrictWithLongCaption" : "Q")}{i:D2}</c:v></c:pt>"));
            string values = string.Concat(Enumerable.Range(0, 12).Select(i => $"<c:pt idx=\"{i}\"><c:v>4</c:v></c:pt>"));
            string layout = manual ? "<c:layout><c:manualLayout><c:layoutTarget val=\"inner\"/><c:xMode val=\"edge\"/><c:yMode val=\"edge\"/><c:wMode val=\"factor\"/><c:hMode val=\"factor\"/><c:x val=\"0.1\"/><c:y val=\"0.1\"/><c:w val=\"0.8\"/><c:h val=\"0.4\"/></c:manualLayout></c:layout>" : "";
            string chart = $$"""
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <c:chart><c:plotArea>{{layout}}<c:barChart><c:barDir val="col"/>
                    <c:ser><c:cat><c:strLit>{{names}}</c:strLit></c:cat><c:val><c:numLit>{{values}}</c:numLit></c:val></c:ser>
                    <c:axId val="1"/><c:axId val="2"/>
                  </c:barChart>
                  <c:catAx><c:axId val="1"/><c:axPos val="b"/><c:tickLblSkip val="1"/>
                    <c:txPr><a:bodyPr rot="{{rotation * 60000}}" vertOverflow="ellipsis"/><a:p><a:pPr><a:defRPr sz="1000"/></a:pPr></a:p></c:txPr>
                  </c:catAx>
                  <c:valAx><c:axId val="2"/><c:axPos val="l"/><c:tickLblPos val="none"/>
                    <c:scaling><c:min val="0"/><c:max val="5"/></c:scaling>
                  </c:valAx></c:plotArea></c:chart>
                </c:chartSpace>
                """;
            string slide = $$"""
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
                  <p:cSld><p:spTree><p:graphicFrame><p:xfrm><a:off x="914400" y="1371600"/><a:ext cx="3657600" cy="{{height * 12700}}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld>
                </p:sld>
                """;
            using var stream = PptxTests.CreateSingleChartPackageStream(chart, new Dictionary<string, byte[]> { ["ppt/slides/slide1.xml"] = TestFixtures.Utf8(slide) });
            using var package = OoxPackage.Open(stream, default);
            var document = new PptxReader().Read(package, default);
            PdfPage page = new PptxRenderer(new PortableRenderingTests.PinnedResolver())
                .RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
            double[][] bars = Regex.Matches(page.Content, @"([-\d.]+) ([-\d.]+) ([\d.]+) ([\d.]+) re f")
                .Select(m => m.Groups.Cast<Group>().Skip(1).Select(g => double.Parse(g.Value, CultureInfo.InvariantCulture)).ToArray())
                .Where(rect => rect[2] > 0d && rect[2] < 20d).ToArray();
            TestAssert.Equal(12, bars.Length);
            return (page, bars);
        }
    }

    public static void PptxAutomaticValueLabelsAvoidInkOverlapOnCrowdedFrames()
    {
        foreach (int height in new[] { 72, 216 })
        foreach (bool authoredInterval in new[] { false, true })
        {
            string categories = string.Concat(Enumerable.Range(0, 12).Select(i => $"<c:pt idx=\"{i}\"><c:v>NorthWarehouseDistrictWithLongCaption{i:D2}</c:v></c:pt>"));
            string values = string.Concat(Enumerable.Range(0, 12).Select(i => $"<c:pt idx=\"{i}\"><c:v>4</c:v></c:pt>"));
            string slide = $$"""
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
                  <p:cSld><p:spTree><p:graphicFrame><p:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="{{height * 12700}}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld>
                </p:sld>
                """;
            using var stream = PptxTests.CreateSingleChartPackageStream($$"""
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <c:chart><c:plotArea><c:barChart><c:barDir val="col"/>
                    <c:ser><c:cat><c:strLit>{{categories}}</c:strLit></c:cat><c:val><c:numLit>{{values}}</c:numLit></c:val></c:ser>
                    <c:axId val="1"/><c:axId val="2"/>
                  </c:barChart>
                  <c:catAx><c:axId val="1"/><c:axPos val="b"/><c:tickLblSkip val="1"/>
                    <c:txPr><a:bodyPr rot="-5400000" vertOverflow="ellipsis"/><a:p><a:pPr><a:defRPr sz="1000"/></a:pPr></a:p></c:txPr>
                  </c:catAx>
                  <c:valAx><c:axId val="2"/><c:axPos val="l"/>
                    <c:scaling><c:min val="0"/><c:max val="5"/></c:scaling>{{(authoredInterval ? "<c:majorUnit val=\"0.5\"/>" : "")}}
                    <c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="1000"/></a:pPr></a:p></c:txPr>
                  </c:valAx></c:plotArea></c:chart>
                </c:chartSpace>
                """, new Dictionary<string, byte[]> { ["ppt/slides/slide1.xml"] = TestFixtures.Utf8(slide) });
            using var package = OoxPackage.Open(stream, default);
            var document = new PptxReader().Read(package, default);
            PdfPage page = new PptxRenderer(new PortableRenderingTests.PinnedResolver())
                .RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
            var fonts = page.Fonts.Where(font => font.ResourceName.StartsWith("CVA", StringComparison.Ordinal)).Select(font => font.Font).ToHashSet();
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).Where(show => fonts.Contains(show.Font)).OrderBy(show => show.Y).ToArray();
            TestAssert.True(shows.Any(show => show.Text == "0") && shows.Any(show => show.Text == "5"),
                "Both range endpoints remain visible after automatic text admission.");
            if (authoredInterval)
            {
                TestAssert.Equal(11, shows.Length);
                continue;
            }
            TestAssert.True(shows.Length >= 2 && (height != 72 || shows.Length < 6), "A crowded automatic axis admits fewer labels, retaining its scale.");
            double previousTop = double.NegativeInfinity;
            foreach (var show in shows)
            {
                double bottom = double.PositiveInfinity, top = double.NegativeInfinity;
                foreach (char character in show.Text)
                {
                    TestAssert.True(show.Font.Font.TryGetGlyphBounds(show.Font.Font.MapCodePoint(character), out var bounds), "Expected pinned numeric glyph bounds.");
                    bottom = Math.Min(bottom, show.Y + bounds.YMin * show.Size / show.Font.Font.UnitsPerEm);
                    top = Math.Max(top, show.Y + bounds.YMax * show.Size / show.Font.Font.UnitsPerEm);
                }
                TestAssert.True(bottom >= previousTop - 0.002d, "Admitted numeric glyph ink must not overlap.");
                previousTop = top;
            }
        }
    }

    public static void PptxCategoryLabelRotationPreservesContentAcrossPlotKinds()
    {
        foreach (string kind in new[] { "bar", "line", "area" })
        foreach (string rotation in new[] { "-60000000", "-5400000", "-2700000", "2700000", "0", "missing" })
        {
            string direction = kind == "bar" ? "<c:barDir val=\"col\"/>" : string.Empty;
            string body = rotation == "missing" ? "" : $"rot=\"{rotation}\" vert=\"horz\"";
            string names = string.Concat(Enumerable.Range(0, 12).Select(i => $"<c:pt idx=\"{i}\"><c:v>Depot{i:D2}</c:v></c:pt>"));
            string values = string.Concat(Enumerable.Range(0, 12).Select(i => $"<c:pt idx=\"{i}\"><c:v>4</c:v></c:pt>"));
            string chart = $$"""
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <c:chart><c:plotArea><c:{{kind}}Chart>{{direction}}
                    <c:ser><c:cat><c:strLit>{{names}}</c:strLit></c:cat><c:val><c:numLit>{{values}}</c:numLit></c:val></c:ser>
                    <c:axId val="1"/><c:axId val="2"/>
                  </c:{{kind}}Chart>
                  <c:catAx><c:axId val="1"/><c:axPos val="b"/><c:tickLblSkip val="1"/>
                    <c:txPr><a:bodyPr {{body}} vertOverflow="ellipsis"/><a:p><a:pPr><a:defRPr sz="1000"/></a:pPr></a:p></c:txPr>
                  </c:catAx>
                  <c:valAx><c:axId val="2"/><c:axPos val="l"/><c:tickLblPos val="none"/>
                    <c:scaling><c:min val="0"/><c:max val="5"/></c:scaling>
                  </c:valAx></c:plotArea></c:chart>
                </c:chartSpace>
                """;
            using var stream = PptxTests.CreateSingleChartPackageStream(chart);
            using var package = OoxPackage.Open(stream, default);
            var document = new PptxReader().Read(package, default);
            PdfPage page = new PptxRenderer(new PortableRenderingTests.PinnedResolver())
                .RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).ToArray();
            TestAssert.True(shows.Select(show => show.Text).SequenceEqual(Enumerable.Range(0, 12).Select(i => $"Depot{i:D2}")),
                $"Authored spacing retains every complete category in {kind}, rotation {rotation}.");
            var matrices = Regex.Matches(page.Content, @"(?m)^(?<a>[-\d.]+) (?<b>[-\d.]+) (?<c>[-\d.]+) (?<d>[-\d.]+) [-\d.]+ [-\d.]+ cm\r?$")
                .Select(m => (A: double.Parse(m.Groups["a"].Value, CultureInfo.InvariantCulture), B: double.Parse(m.Groups["b"].Value, CultureInfo.InvariantCulture))).ToArray();
            double expected = rotation is "-60000000" or "missing" ? -45d : double.Parse(rotation, CultureInfo.InvariantCulture) / 60000d;
            if (expected != 0d)
                // PDF numbers are emitted to three decimal places.
                TestAssert.True(matrices.Count(m => Math.Abs(m.A - Math.Cos(-expected * Math.PI / 180d)) <= 0.0005d &&
                    Math.Abs(m.B - Math.Sin(-expected * Math.PI / 180d)) <= 0.0005d) == 12,
                    "Each label uses the authored rotation or the automatic measured angle, through the shared PDF transform.");
            PptxSceneChartAxis axis = PptxTests.BuildSingleChartScene(chart)!.Axes.Single(a => a.AxisKind == PptxSceneChartAxisKind.Category);
            TestAssert.Equal(rotation == "missing" ? "" : rotation, axis.TextBodyProperties.RotationValue);
            TestAssert.Equal(rotation == "missing" ? "" : "horz", axis.TextBodyProperties.OrientationValue);
            TestAssert.Equal("ellipsis", axis.TextBodyProperties.VerticalOverflowValue);
        }
    }

    public static void PptxSkippedCategoryLabelsUseSpaceBetweenVisibleTicks()
    {
        foreach (string kind in new[] { "bar", "line" })
        foreach (int skip in new[] { 4, int.MaxValue })
        {
            string barDirection = kind == "bar" ? "<c:barDir val=\"col\"/>" : string.Empty;
            string categories = string.Concat(Enumerable.Range(0, 12).Select(i =>
                $"<c:pt idx=\"{i}\"><c:v>NorthDepot{i:D2}</c:v></c:pt>"));
            string values = string.Concat(Enumerable.Range(0, 12).Select(i =>
                $"<c:pt idx=\"{i}\"><c:v>4</c:v></c:pt>"));
            using var stream = PptxTests.CreateSingleChartPackageStream($$"""
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <c:chart><c:plotArea><c:{{kind}}Chart>{{barDirection}}
                    <c:ser><c:cat><c:strLit>{{categories}}</c:strLit></c:cat>
                      <c:val><c:numLit>{{values}}</c:numLit></c:val></c:ser>
                    <c:axId val="1"/><c:axId val="2"/>
                  </c:{{kind}}Chart>
                  <c:catAx><c:axId val="1"/><c:axPos val="b"/><c:tickLblSkip val="{{skip}}"/>
                    <c:txPr><a:bodyPr rot="0"/><a:p><a:pPr><a:defRPr sz="1000"/></a:pPr></a:p></c:txPr>
                  </c:catAx>
                  <c:valAx><c:axId val="2"/><c:axPos val="l"/><c:tickLblPos val="none"/>
                    <c:scaling><c:min val="0"/><c:max val="5"/></c:scaling>
                  </c:valAx></c:plotArea></c:chart>
                </c:chartSpace>
                """);
            using var package = OoxPackage.Open(stream, default);
            var document = new PptxReader().Read(package, default);
            PdfPage page = new PptxRenderer(new PortableRenderingTests.PinnedResolver())
                .RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).ToArray();
            string[] expected = Enumerable.Range(0, 12).Where(i => i % skip == 0)
                .Select(i => $"NorthDepot{i:D2}").ToArray();
            TestAssert.True(shows.Select(show => show.Text).SequenceEqual(expected),
                "Explicit tick spacing preserves complete selected labels and suppresses the intervening labels.");
            var clips = Regex.Matches(page.Content,
                @"(?<x>[-\d.]+) [-\d.]+ (?<width>[\d.]+) [\d.]+ re W\* n\s+BT\s+(?<body>.*?)ET",
                RegexOptions.Singleline).Where(match => match.Groups["body"].Value.Contains("/CCA", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(shows.Length, clips.Length);
            for (int i = 0; i < shows.Length; i++)
            {
                double left = double.Parse(clips[i].Groups["x"].Value, CultureInfo.InvariantCulture);
                double width = double.Parse(clips[i].Groups["width"].Value, CultureInfo.InvariantCulture);
                double advance = shows[i].Font.MeasureTextPoints(shows[i].Text, shows[i].Size, kerningEnabled: false);
                TestAssert.True(shows[i].X >= left - 0.002d && shows[i].X + advance <= left + width + 0.002d,
                    "The emitted complete label fits its actual PDF clip, rather than being present only in extracted text.");
            }
        }
    }

    public static void PptxChartSceneAxesUseSceneOwnedXmlEvidence()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:lineChart>
                <c:catAx><c:axId val="10"/><c:tickLblPos val="low"/></c:catAx>
                <c:valAx><c:axId val="20"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        TestAssert.Equal(2, chart.Axes.Count);

        XDocument mismatchedChartXml = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart><c:axId val="10"/><c:axId val="20"/></c:lineChart>
                <c:catAx><c:axId val="10"/><c:tickLblPos val="high"/></c:catAx>
                <c:valAx><c:axId val="20"/><c:majorUnit val="99"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """);
        XNamespace chartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedPlot = mismatchedChartXml.Descendants(chartNamespace + "lineChart").First();
        System.Reflection.MethodInfo readValueAxes = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartValueAxesForPlot",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected renderer value-axis bridge.");
        System.Reflection.MethodInfo readCategoryAxis = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartCategoryAxisForPlot",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected renderer category-axis bridge.");

        object sceneValueAxis = ((System.Collections.IEnumerable)(readValueAxes.Invoke(null, [chart, chart.Plots[0], mismatchedChartXml, mismatchedPlot]) ?? throw new InvalidOperationException("Expected scene value-axis collection."))).Cast<object>().Single();
        object xmlOnlyValueAxis = ((System.Collections.IEnumerable)(readValueAxes.Invoke(null, [null, null, mismatchedChartXml, mismatchedPlot]) ?? throw new InvalidOperationException("Expected XML-only value-axis collection."))).Cast<object>().Single();
        object sceneCategoryAxis = readCategoryAxis.Invoke(null, [chart, chart.Plots[0], mismatchedChartXml, mismatchedPlot]) ?? throw new InvalidOperationException("Expected scene category-axis source.");
        object xmlOnlyCategoryAxis = readCategoryAxis.Invoke(null, [null, null, mismatchedChartXml, mismatchedPlot]) ?? throw new InvalidOperationException("Expected XML-only category-axis source.");

        XElement sceneValueXmlAxis = (XElement?)sceneValueAxis.GetType().GetProperty("XmlAxis")?.GetValue(sceneValueAxis) ?? throw new InvalidOperationException("Expected scene value-axis XML evidence.");
        XElement xmlOnlyValueXmlAxis = (XElement?)xmlOnlyValueAxis.GetType().GetProperty("XmlAxis")?.GetValue(xmlOnlyValueAxis) ?? throw new InvalidOperationException("Expected XML-only value-axis XML evidence.");
        XElement sceneCategoryXmlAxis = (XElement?)sceneCategoryAxis.GetType().GetProperty("XmlAxis")?.GetValue(sceneCategoryAxis) ?? throw new InvalidOperationException("Expected scene category-axis XML evidence.");
        XElement xmlOnlyCategoryXmlAxis = (XElement?)xmlOnlyCategoryAxis.GetType().GetProperty("XmlAxis")?.GetValue(xmlOnlyCategoryAxis) ?? throw new InvalidOperationException("Expected XML-only category-axis XML evidence.");

        TestAssert.True(sceneValueXmlAxis.Element(chartNamespace + "majorUnit") is null, "Expected scene value axis to keep scene-owned XML evidence.");
        TestAssert.Equal("99", xmlOnlyValueXmlAxis.Element(chartNamespace + "majorUnit")?.Attribute("val")?.Value ?? string.Empty);
        TestAssert.Equal("low", sceneCategoryXmlAxis.Element(chartNamespace + "tickLblPos")?.Attribute("val")?.Value ?? string.Empty);
        TestAssert.Equal("high", xmlOnlyCategoryXmlAxis.Element(chartNamespace + "tickLblPos")?.Attribute("val")?.Value ?? string.Empty);
    }

    public static void PptxChartValueAxisScaleFallbackUsesSceneAuthoritativeAbsence()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:lineChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        TestAssert.Equal(0, chart.Axes.Count);

        XDocument mismatchedChartXml = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:valAx><c:axId val="20"/><c:scaling><c:min val="-100"/><c:max val="100"/></c:scaling></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """);

        Type axisSourceType = typeof(PptxRenderer).GetNestedType(
            "ChartAxisSource",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart axis-source type.");
        object emptyAxisSource = Activator.CreateInstance(axisSourceType, [null, null]) ?? throw new InvalidOperationException("Expected empty axis source.");
        System.Reflection.MethodInfo resolveValueAxis = typeof(PptxRenderer).GetMethod(
            "ResolveXmlValueAxisForSource",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected value-axis fallback resolver.");

        object? sceneAxis = resolveValueAxis.Invoke(null, [chart, emptyAxisSource, mismatchedChartXml]);
        object? xmlOnlyAxis = resolveValueAxis.Invoke(null, [null, emptyAxisSource, mismatchedChartXml]);

        TestAssert.True(sceneAxis is null, "Expected scene-backed missing value axis not to be repaired from fallback XML for scale/layout estimation.");
        TestAssert.True(xmlOnlyAxis is XElement, "Expected XML-only value-axis scale lookup to keep reading XML.");
    }

    public static void PptxChartUnknownDataLabelPositionResolvesThroughExplicitDefault()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:dLbls><c:showVal val="1"/><c:dLblPos val="bogus"/></c:dLbls>
                  <c:ser><c:tx><c:v>Line Series</c:v></c:tx><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:lineChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        TestAssert.Equal(PptxSceneChartDataLabelPosition.Unknown, chart.Plots[0].DataLabels.PositionKind);
        TestAssert.Equal("bogus", chart.Plots[0].DataLabels.Position);

        System.Reflection.MethodInfo resolvePosition = typeof(PptxRenderer).GetMethod(
            "ResolveChartDataLabelPosition",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected data-label position resolver.");
        object position = resolvePosition.Invoke(null, [PptxSceneChartDataLabelPosition.Unknown]) ?? throw new InvalidOperationException("Expected resolved data-label position.");

        TestAssert.Equal(PptxSceneChartDataLabelPosition.OutsideEnd, (PptxSceneChartDataLabelPosition)position);
    }

    public static void PptxChartUnknownGroupingUsesSceneAuthoritativeDefault()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:grouping val="bogus"/>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:lineChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        TestAssert.Equal(PptxSceneChartGrouping.Unknown, chart.Plots[0].GroupingKind);
        TestAssert.Equal("bogus", chart.Plots[0].Grouping);

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:lineChart><c:grouping val="stacked"/></c:lineChart></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "lineChart").Single();
        System.Reflection.MethodInfo readGrouping = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartGrouping",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart grouping resolver.");
        object grouping = readGrouping.Invoke(null, [chart.Plots[0], mismatchedXmlFallback, PptxSceneChartGrouping.Standard]) ?? throw new InvalidOperationException("Expected resolved grouping.");

        TestAssert.Equal(PptxSceneChartGrouping.Standard, (PptxSceneChartGrouping)grouping);
    }

    public static void PptxChartUnknownBarDirectionUsesSceneAuthoritativeDefault()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:barChart>
                  <c:barDir val="bogus"/>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:barChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        TestAssert.Equal(PptxSceneChartBarDirection.Unknown, chart.Plots[0].BarDirectionKind);
        TestAssert.Equal("bogus", chart.Plots[0].BarDirection);

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:barChart><c:barDir val="bar"/></c:barChart></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "barChart").Single();
        System.Reflection.MethodInfo readBarDirection = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartBarDirection",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart bar-direction resolver.");
        object barDirection = readBarDirection.Invoke(null, [chart.Plots[0], mismatchedXmlFallback]) ?? throw new InvalidOperationException("Expected resolved bar direction.");

        TestAssert.Equal(PptxSceneChartBarDirection.Column, (PptxSceneChartBarDirection)barDirection);
    }

    public static void PptxChartBarOptionsUseSceneAuthoritativeDefaults()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:barChart>
                  <c:barDir val="bogus"/>
                  <c:grouping val="bogus"/>
                  <c:gapWidth val="futureGap"/>
                  <c:overlap val="futureOverlap"/>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:barChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartPlot plot = chart.Plots[0];
        TestAssert.Equal(PptxSceneChartGrouping.Unknown, plot.GroupingKind);
        TestAssert.Equal(PptxSceneChartBarDirection.Unknown, plot.BarDirectionKind);
        TestAssert.True(plot.VaryColors is null, "Expected missing varyColors metadata to remain distinct from the effective renderer default.");
        TestAssert.Equal(string.Empty, plot.VaryColorsValue);
        TestAssert.Equal(null, plot.GapWidth);
        TestAssert.Equal(null, plot.Overlap);

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:barChart>
                <c:barDir val="bar"/>
                <c:grouping val="stacked"/>
                <c:varyColors val="0"/>
                <c:gapWidth val="300"/>
                <c:overlap val="75"/>
              </c:barChart></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "barChart").Single();
        System.Reflection.MethodInfo readOptions = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartBarOptions",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart bar-option resolver.");
        object options = readOptions.Invoke(null, [plot, mismatchedXmlFallback, PptxSceneChartGrouping.Clustered]) ?? throw new InvalidOperationException("Expected resolved bar options.");
        Type optionsType = options.GetType();

        TestAssert.Equal(PptxSceneChartGrouping.Clustered, (PptxSceneChartGrouping)(optionsType.GetProperty("Grouping")?.GetValue(options) ?? default(PptxSceneChartGrouping)));
        TestAssert.Equal(PptxSceneChartBarDirection.Column, (PptxSceneChartBarDirection)(optionsType.GetProperty("BarDirection")?.GetValue(options) ?? default(PptxSceneChartBarDirection)));
        object varyColors = optionsType.GetProperty("VaryColors")?.GetValue(options) ?? throw new InvalidOperationException("Expected varyColors option.");
        TestAssert.True(PptxTests.ChartBooleanOptionValue(varyColors), "Expected missing effective scene varyColors to use the scene default, not the XML fallback.");
        TestAssert.True(!PptxTests.ChartBooleanOptionIsDefined(varyColors), "Expected missing scene varyColors state to remain distinguishable from an explicit true token.");
        TestAssert.Equal(string.Empty, PptxTests.ChartBooleanOptionRawValue(varyColors));
        TestAssert.Equal(150d, (double)(optionsType.GetProperty("GapWidth")?.GetValue(options) ?? double.NaN));
        TestAssert.Equal(0d, (double)(optionsType.GetProperty("Overlap")?.GetValue(options) ?? double.NaN));
    }

    public static void PptxChartMultiValueAxisStripFactorDistinguishesOppositeSides()
    {
        System.Reflection.MethodInfo readStripFactor = typeof(PptxRenderer).GetMethod(
            "GetMultiValueAxisStripFactor",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart axis strip-factor resolver.");

        double singleLeft = (double)(readStripFactor.Invoke(null, [1, false]) ?? double.NaN);
        double singleRight = (double)(readStripFactor.Invoke(null, [1, true]) ?? double.NaN);
        double multipleLeft = (double)(readStripFactor.Invoke(null, [2, false]) ?? double.NaN);
        double multipleRight = (double)(readStripFactor.Invoke(null, [2, true]) ?? double.NaN);
        Type metricRules = typeof(PptxRenderer).GetNestedType(
            "PptxChartMetricRules",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart metric rules.");
        double primaryFactor = (double)(metricRules.GetField("BarMultiValueAxisPrimaryStripFactor")?.GetRawConstantValue() ?? double.NaN);
        double secondaryFactor = (double)(metricRules.GetField("BarMultiValueAxisSecondaryStripFactor")?.GetRawConstantValue() ?? double.NaN);

        TestAssert.Equal(1d, singleLeft);
        TestAssert.Equal(1d, singleRight);
        TestAssert.Equal(primaryFactor, multipleLeft);
        TestAssert.Equal(secondaryFactor, multipleRight);
    }

    public static void PptxChartMinorGridlinesUseSelectedMajorSubdivisions()
    {
        foreach ((double max, double major, double expectedMinor, int count) in new[] { (100d, 20d, 4d, 20), (100d, 10d, 2d, 40), (1d, .2d, .04d, 20) })
        {
            IReadOnlyList<double> values = ReadMinorGridlineValues(max, major, null);
            TestAssert.Equal(count, values.Count);
            TestAssert.True(Math.Abs(values[0] - expectedMinor) < 1e-8, "First automatic minor must subdivide the selected major interval.");
            TestAssert.True(Math.Abs(values[^1] - (max - expectedMinor)) < 1e-8, "Last minor must stop before the major endpoint.");
        }

        TestAssert.Equal(20, ReadMinorGridlineValues(100d, null, null, targetCount: 9d).Count);
        TestAssert.Equal(40, ReadMinorGridlineValues(100d, null, null, targetCount: 10d).Count);
    }

    public static void PptxChartMinorGridlinesKeepAuthoredUnitsAndExcludeMajorLevels()
    {
        IReadOnlyList<double> dense = ReadMinorGridlineValues(100d, 20d, 1d);
        TestAssert.Equal(95, dense.Count);
        foreach (double value in new[] { 0d, 20d, 40d, 60d, 80d, 100d })
        {
            TestAssert.True(!dense.Contains(value), "Major levels and endpoints must not emit minor strokes.");
        }

        IReadOnlyList<double> independent = ReadMinorGridlineValues(100d, 10d, 4d);
        TestAssert.Equal(20, independent.Count);
        TestAssert.True(independent.Contains(8d) && independent.Contains(12d), "An authored minor unit must keep its independent spacing.");
        TestAssert.True(!independent.Contains(20d), "Coincident major levels must still be excluded.");
        IReadOnlyList<double> absent = ReadMinorGridlineValues(100d, 20d, 4d, hasMajorElement: false);
        TestAssert.Equal(25, absent.Count);
        TestAssert.True(absent.Contains(20d) && absent.Contains(100d), "Absent major-gridline source leaves coincident minor levels visible.");
    }

    public static void PptxChartMinorGridlinesRetainCrossingExclusion()
    {
        IReadOnlyList<double> ordinary = ReadMinorGridlineValues(100d, 20d, 1d);
        IReadOnlyList<double> crossing = ReadMinorGridlineValues(100d, 20d, 1d, crossing: 48d);
        TestAssert.Equal(94, crossing.Count);
        TestAssert.True(ordinary.Where(value => value != 48d).SequenceEqual(crossing), "Only the existing crossing exclusion may remove another minor level.");
        TestAssert.True(ordinary.SequenceEqual(ReadMinorGridlineValues(100d, 20d, 1d, crossing: 40d)), "A crossing already owned by a major level must not remove another minor.");
    }

    public static void PptxChartMinorGridlinesRetainTickBounds()
    {
        TestAssert.True(ReadMinorGridlineValues(100d, 20d, 4d).Count > 0, "Ordinary bounded minor levels remain available.");
        TestAssert.Throws<OoxPdfLimitExceededException>(() => ReadMinorGridlineValues(100d, 20d, 1e-12));
    }

    private static IReadOnlyList<double> ReadMinorGridlineValues(double max, double? major, double? minor, double? crossing = 0d, double targetCount = 9d, bool hasMajorElement = true)
    {
        const System.Reflection.BindingFlags flags = System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static;
        Type renderer = typeof(PptxRenderer);
        Type extentsType = renderer.GetNestedType("ChartValueExtents", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart extents.");
        Type unitsType = renderer.GetNestedType("ChartAxisUnits", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart units.");
        object extents = Activator.CreateInstance(extentsType, [0d, max]) ?? throw new InvalidOperationException("Expected chart extents instance.");
        object units = Activator.CreateInstance(unitsType, [major, minor]) ?? throw new InvalidOperationException("Expected chart units instance.");
        Type policyType = renderer.GetNestedType("ChartAxisTickPolicy", System.Reflection.BindingFlags.NonPublic)!;
        object policy = Enum.Parse(policyType, "ZeroAlignedWithEndpoints");
        System.Reflection.MethodInfo method = renderer.GetMethod("GetChartMinorGridlineValuesCore", flags) ?? throw new InvalidOperationException("Expected shared minor interval resolver.");
        try
        {
            return (IReadOnlyList<double>)(method.Invoke(null, [extents, units, crossing, targetCount, hasMajorElement, policy]) ?? throw new InvalidOperationException("Expected minor levels."));
        }
        catch (System.Reflection.TargetInvocationException exception) when (exception.InnerException is { } inner)
        {
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(inner).Throw();
            throw;
        }
    }

    public static void PptxChartAxisTickExpansionRejectsDenseUnits()
    {
        TestAssert.Throws<OoxPdfLimitExceededException>(() => ReadBoundedAxisTicks(0d, 1d, 1e-12));
        TestAssert.Throws<OoxPdfLimitExceededException>(() => ReadBoundedAxisTicks(0d, 99999d, 1d));
    }

    public static void PptxChartAxisTickExpansionKeepsBoundaryAndOrdinaryTicks()
    {
        IReadOnlyList<double> boundary = ReadBoundedAxisTicks(0d, 99998d, 1d);
        TestAssert.Equal(99999, boundary.Count);
        TestAssert.Equal(0d, boundary[0]);
        TestAssert.Equal(99998d, boundary[^1]);
        TestAssert.Equal("-7,-5,0,5,7", string.Join(",", ReadBoundedAxisTicks(-7d, 7d, 5d)));
        TestAssert.Equal("0,5,10,15,20,25,30,35,40,45,50", string.Join(",", ReadBoundedAxisTicks(0d, 50d, null)));
        TestAssert.Equal("1,2,3", string.Join(",", ReadBoundedAxisTicks(0d, 4d, 1d, includeEndpoints: false)));
    }

    public static void PptxChartAxisTickExpansionRejectsNonprogressAndInvalidBounds()
    {
        TestAssert.Throws<OoxPdfLimitExceededException>(() => ReadBoundedAxisTicks(1e16, 1e16 + 4d, 1d));
        TestAssert.Throws<OoxPdfLimitExceededException>(() => ReadBoundedAxisTicks(9007199254740990d, 9007199254740996d, 1d));
        foreach ((double min, double max) in new[] { (double.NaN, 4d), (0d, double.PositiveInfinity), (-double.MaxValue, double.MaxValue), (4d, 0d) })
        {
            TestAssert.Throws<OoxPdfLimitExceededException>(() => ReadBoundedAxisTicks(min, max, 1d));
        }
    }

    public static void PptxChartAxisTickExpansionConversionPropagatesLimitFailure()
    {
        foreach ((string min, string max, string unit) in new[] { ("0", "1", "1e-12"), ("10000000000000000", "10000000000000004", "1") })
        {
            string chart = $"""
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
                  <c:chart><c:plotArea>
                    <c:lineChart>
                      <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser>
                      <c:axId val="10"/><c:axId val="20"/>
                    </c:lineChart>
                    <c:catAx><c:axId val="10"/><c:crossAx val="20"/></c:catAx>
                    <c:valAx><c:axId val="20"/><c:scaling><c:min val="{min}"/><c:max val="{max}"/></c:scaling><c:majorUnit val="{unit}"/><c:crossAx val="10"/></c:valAx>
                  </c:plotArea></c:chart>
                </c:chartSpace>
                """;
            TestAssert.Throws<OoxPdfLimitExceededException>(() => PptxTests.ConvertSingleChartAndCollectDiagnostics(chart));
        }
    }

    public static void PptxChartSourceMinimumIntervalsUseAuthoredLatticeAndAlignedEndpoints()
    {
        foreach ((double min, double max, string expected) in new[]
        {
            (7d, 107d, "7,27,47,67,87,107"),
            (-7d, 93d, "-7,13,33,53,73,93"),
            (0.5d, 100.5d, "0.5,20.5,40.5,60.5,80.5,100.5"),
            (7d, 103d, "7,27,47,67,87"),
            (0d, 97d, "0,20,40,60,80")
        })
        {
            IReadOnlyList<double> ticks = ReadBoundedAxisTicks(min, max, 20d, sourceMinimumIntervals: true);
            TestAssert.Equal(expected, string.Join(",", ticks.Select(v => v.ToString(CultureInfo.InvariantCulture))));
        }
        TestAssert.Equal("7,20,40,60,80,100,107", string.Join(",", ReadBoundedAxisTicks(7d, 107d, 20d).Select(v => v.ToString(CultureInfo.InvariantCulture))));
    }

    public static void PptxChartSourceMinimumIntervalsRetainWorkAndProgressBounds()
    {
        TestAssert.Throws<OoxPdfLimitExceededException>(() => ReadBoundedAxisTicks(7d, 107d, 1e-12, sourceMinimumIntervals: true));
        TestAssert.Equal(6, ReadBoundedAxisTicks(7d, 107d, 20d, sourceMinimumIntervals: true).Count);
        TestAssert.Equal(6, ReadBoundedAxisTicks(-7d, 93d, 20d, sourceMinimumIntervals: true).Count);
        TestAssert.Throws<OoxPdfLimitExceededException>(() => ReadBoundedAxisTicks(1e16, 1e16 + 4d, 1d, sourceMinimumIntervals: true));
    }

    public static void PptxChartSourceMinimumPolicyKeepsSceneAuthorityAndUnsupportedFallbacks()
    {
        foreach (string log in new[] { "", "<c:logBase val=\"10\"/>", "<c:logBase val=\"invalid\"/>" })
        {
            PptxSceneChart chart = PptxTests.BuildSingleChartScene($$"""
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea>
                  <c:lineChart><c:axId val="10"/><c:axId val="20"/></c:lineChart>
                  <c:catAx><c:axId val="10"/><c:axPos val="b"/><c:crossAx val="20"/></c:catAx>
                  <c:valAx><c:axId val="20"/><c:scaling>{{log}}<c:min val="7"/><c:max val="107"/></c:scaling><c:axPos val="l"/><c:crossAx val="10"/></c:valAx>
                </c:plotArea></c:chart></c:chartSpace>
                """) ?? throw new InvalidOperationException("Expected source axis scene.");
            PptxSceneChartAxis axis = chart.Axes.Single(a => a.AxisKind == PptxSceneChartAxisKind.Value);
            TestAssert.Equal(log.Length != 0, axis.HasLogarithmicScale);
            const System.Reflection.BindingFlags flags = System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static;
            var method = typeof(PptxRenderer).GetMethod("ReadCartesianChartAxisTickPolicy", flags) ?? throw new InvalidOperationException("Expected source policy.");
            Type extentsType = typeof(PptxRenderer).GetNestedType("ChartValueExtents", System.Reflection.BindingFlags.NonPublic)!;
            object extents = Activator.CreateInstance(extentsType, [7d, 107d])!;
            TestAssert.Equal(log.Length == 0 ? "SourceMinimumIntervals" : "ZeroAlignedWithEndpoints", method.Invoke(null, [chart, axis, extents, false])!.ToString());
            object fallbackExtents = Activator.CreateInstance(extentsType, [0d, 100d])!;
            TestAssert.Equal("ZeroAlignedWithEndpoints", method.Invoke(null, [chart, axis, fallbackExtents, false])!.ToString());
        }
    }

    private static IReadOnlyList<double> ReadBoundedAxisTicks(double min, double max, double? unit, bool includeEndpoints = true, bool sourceMinimumIntervals = false)
    {
        const System.Reflection.BindingFlags flags = System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static;
        Type extentsType = typeof(PptxRenderer).GetNestedType("ChartValueExtents", System.Reflection.BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("Expected chart value extents.");
        object extents = Activator.CreateInstance(extentsType, [min, max]) ?? throw new InvalidOperationException("Expected chart value extents instance.");
        System.Reflection.MethodInfo method = typeof(PptxRenderer).GetMethod(sourceMinimumIntervals ? "GetChartAxisTickValuesCore" : "GetChartAxisTickValues", flags)
            ?? throw new InvalidOperationException("Expected chart axis tick expansion.");
        try
        {
            object[] arguments = sourceMinimumIntervals
                ? [extents, unit!, includeEndpoints, 10d, Enum.Parse(typeof(PptxRenderer).GetNestedType("ChartAxisTickPolicy", System.Reflection.BindingFlags.NonPublic)!, "SourceMinimumIntervals")]
                : [extents, unit!, includeEndpoints, 10d];
            return (IReadOnlyList<double>)(method.Invoke(null, arguments) ?? throw new InvalidOperationException("Expected chart ticks."));
        }
        catch (System.Reflection.TargetInvocationException exception) when (exception.InnerException is { } inner)
        {
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(inner).Throw();
            throw;
        }
    }

    public static void PptxChartVerticalValueAxisAutoTicksUseOfficeDenseDefault()
    {
        Type extentsType = typeof(PptxRenderer).GetNestedType(
            "ChartValueExtents",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart value extents.");
        object extents = Activator.CreateInstance(
            extentsType,
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic,
            binder: null,
            args: [0d, 50d],
            culture: CultureInfo.InvariantCulture) ?? throw new InvalidOperationException("Expected chart value extents instance.");
        System.Reflection.MethodInfo readTargetCount = typeof(PptxRenderer).GetMethod(
            "GetValueAxisAutoTickTargetCount",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart tick target resolver.");
        System.Reflection.MethodInfo readTickValues = typeof(PptxRenderer).GetMethods(
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)
            .Single(method => method.Name == "GetChartAxisTickValues" && method.GetParameters().Length == 4);

        double targetCount = (double)(readTargetCount.Invoke(null, [false, true, false]) ?? double.NaN);
        object values = readTickValues.Invoke(null, [extents, null, true, targetCount]) ?? throw new InvalidOperationException("Expected chart tick values.");
        string tickList = string.Join(
            ",",
            ((System.Collections.IEnumerable)values).Cast<object>().Select(value => Convert.ToDouble(value, CultureInfo.InvariantCulture).ToString("0.########", CultureInfo.InvariantCulture)));

        TestAssert.Equal("0,5,10,15,20,25,30,35,40,45,50", tickList);
    }

    public static void PptxChartLineOptionsUseSceneAuthoritativeDefaults()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart>
                <c:dispBlanksAs val="futureBlankPolicy"/>
                <c:plotArea>
                  <c:lineChart>
                    <c:grouping val="bogus"/>
                    <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                  </c:lineChart>
                </c:plotArea>
              </c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartPlot plot = chart.Plots[0];
        TestAssert.Equal(PptxSceneChartGrouping.Unknown, plot.GroupingKind);
        TestAssert.Equal(PptxSceneChartDisplayBlanksAs.Unknown, chart.Options.DisplayBlanksAsKind);
        TestAssert.True(plot.Series[0].Smooth is null, "Expected missing smooth metadata to remain distinct from the effective renderer default.");

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XDocument mismatchedChartXml = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart>
                <c:dispBlanksAs val="span"/>
                <c:plotArea><c:lineChart>
                  <c:grouping val="stacked"/>
                  <c:ser><c:smooth val="1"/></c:ser>
                </c:lineChart></c:plotArea>
              </c:chart>
            </c:chartSpace>
            """);
        XElement mismatchedXmlFallback = mismatchedChartXml.Descendants(c + "lineChart").Single();
        System.Reflection.MethodInfo readOptions = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartLineOptions",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart line-option resolver.");
        object options = readOptions.Invoke(null, [chart, plot, mismatchedChartXml, mismatchedXmlFallback, PptxSceneChartGrouping.Standard, false]) ?? throw new InvalidOperationException("Expected resolved line options.");
        Type optionsType = options.GetType();
        object[] smoothSeries = (((System.Collections.IEnumerable?)optionsType.GetProperty("SmoothSeries")?.GetValue(options)) ?? throw new InvalidOperationException("Expected smooth series options.")).Cast<object>().ToArray();

        TestAssert.Equal(PptxSceneChartGrouping.Standard, (PptxSceneChartGrouping)(optionsType.GetProperty("Grouping")?.GetValue(options) ?? default(PptxSceneChartGrouping)));
        TestAssert.True((bool)(optionsType.GetProperty("Stacked")?.GetValue(options) ?? true) == false, "Expected unknown scene grouping to use the supplied scene default, not fallback XML stacked.");
        TestAssert.True((bool)(optionsType.GetProperty("PercentStacked")?.GetValue(options) ?? true) == false, "Expected unknown scene grouping to stay non-percent-stacked.");
        TestAssert.True(smoothSeries.Length == 1 && !PptxTests.ChartBooleanOptionValue(smoothSeries[0]), "Expected missing scene smooth flag to use the scene default, not fallback XML smooth=1.");
        TestAssert.True(!PptxTests.ChartBooleanOptionIsDefined(smoothSeries[0]), "Expected missing scene smooth state to remain distinguishable from an explicit false token.");
        TestAssert.Equal(PptxSceneChartDisplayBlanksAs.Gap, (PptxSceneChartDisplayBlanksAs)(optionsType.GetProperty("DisplayBlanksAs")?.GetValue(options) ?? default(PptxSceneChartDisplayBlanksAs)));
    }

    public static void PptxChartAreaOptionsUseSceneAuthoritativeDefaults()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart>
                <c:dispBlanksAs val="futureBlankPolicy"/>
                <c:plotArea>
                  <c:areaChart>
                    <c:grouping val="bogus"/>
                    <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                  </c:areaChart>
                </c:plotArea>
              </c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartPlot plot = chart.Plots[0];
        TestAssert.Equal(PptxSceneChartGrouping.Unknown, plot.GroupingKind);
        TestAssert.Equal(PptxSceneChartDisplayBlanksAs.Unknown, chart.Options.DisplayBlanksAsKind);

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XDocument mismatchedChartXml = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart>
                <c:dispBlanksAs val="span"/>
                <c:plotArea><c:areaChart><c:grouping val="percentStacked"/></c:areaChart></c:plotArea>
              </c:chart>
            </c:chartSpace>
            """);
        XElement mismatchedXmlFallback = mismatchedChartXml.Descendants(c + "areaChart").Single();
        System.Reflection.MethodInfo readOptions = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartAreaOptions",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart area-option resolver.");
        object options = readOptions.Invoke(null, [chart, plot, mismatchedChartXml, mismatchedXmlFallback, PptxSceneChartGrouping.Standard]) ?? throw new InvalidOperationException("Expected resolved area options.");
        Type optionsType = options.GetType();

        TestAssert.Equal(PptxSceneChartGrouping.Standard, (PptxSceneChartGrouping)(optionsType.GetProperty("Grouping")?.GetValue(options) ?? default(PptxSceneChartGrouping)));
        TestAssert.True((bool)(optionsType.GetProperty("Stacked")?.GetValue(options) ?? true) == false, "Expected unknown scene grouping to use the supplied scene default, not fallback XML percentStacked.");
        TestAssert.True((bool)(optionsType.GetProperty("PercentStacked")?.GetValue(options) ?? true) == false, "Expected unknown scene grouping to stay non-percent-stacked.");
        TestAssert.Equal(PptxSceneChartDisplayBlanksAs.Gap, (PptxSceneChartDisplayBlanksAs)(optionsType.GetProperty("DisplayBlanksAs")?.GetValue(options) ?? default(PptxSceneChartDisplayBlanksAs)));
    }

    public static void PptxChartCategoryAxisLabelPolicyUsesSceneAuthoritativeDefaults()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:barChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:barChart>
                <c:catAx>
                  <c:axId val="10"/><c:axPos val="b"/><c:lblOffset val="futureOffset"/><c:tickLblSkip val="futureSkip"/><c:crossAx val="20"/>
                </c:catAx>
                <c:valAx><c:axId val="20"/><c:axPos val="l"/><c:crossAx val="10"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartAxis sceneAxis = chart.Axes.Single(axis => axis.AxisKind == PptxSceneChartAxisKind.Category);
        TestAssert.True(sceneAxis.LabelOffset is null, "Expected unparseable scene label offset to stay distinct from an authored value.");
        TestAssert.True(sceneAxis.TickLabelSkip is null, "Expected unparseable scene tick-label skip to stay distinct from an authored value.");

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedAxis = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:catAx><c:axId val="10"/><c:lblOffset val="200"/><c:tickLblSkip val="3"/></c:catAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "catAx").Single();
        System.Reflection.MethodInfo resolveLabelOffset = typeof(PptxRenderer).GetMethod(
            "ResolveSceneOrXmlCategoryAxisLabelOffsetScale",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected category-axis label-offset resolver.");
        System.Reflection.MethodInfo resolveTickLabelSkip = typeof(PptxRenderer).GetMethod(
            "ResolveSceneOrXmlCategoryAxisTickLabelSkip",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected category-axis tick-label-skip resolver.");

        double sceneOffsetScale = (double)resolveLabelOffset.Invoke(null, [sceneAxis, mismatchedAxis])!;
        int sceneSkip = (int)resolveTickLabelSkip.Invoke(null, [sceneAxis, mismatchedAxis])!;
        double xmlOnlyOffsetScale = (double)resolveLabelOffset.Invoke(null, [null, mismatchedAxis])!;
        int xmlOnlySkip = (int)resolveTickLabelSkip.Invoke(null, [null, mismatchedAxis])!;

        TestAssert.Equal(1d, sceneOffsetScale);
        TestAssert.Equal(1, sceneSkip);
        TestAssert.Equal(2d, xmlOnlyOffsetScale);
        TestAssert.Equal(3, xmlOnlySkip);
    }

    public static void PptxChartScatterOptionsUseSceneAuthoritativeDefaults()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:scatterChart>
                  <c:scatterStyle val="bogus"/>
                  <c:ser><c:xVal><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:xVal><c:yVal><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:yVal></c:ser>
                </c:scatterChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartPlot plot = chart.Plots[0];
        TestAssert.Equal(PptxSceneChartScatterStyle.Unknown, plot.ScatterStyleKind);
        TestAssert.True(plot.Series[0].Smooth is null, "Expected missing scatter smooth metadata to remain distinct from the effective renderer default.");

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:scatterChart>
                <c:scatterStyle val="lineMarker"/>
                <c:ser><c:smooth val="0"/></c:ser>
              </c:scatterChart></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "scatterChart").Single();
        System.Reflection.MethodInfo readOptions = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartScatterOptions",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart scatter-option resolver.");
        object options = readOptions.Invoke(null, [plot, mismatchedXmlFallback]) ?? throw new InvalidOperationException("Expected resolved scatter options.");
        Type optionsType = options.GetType();
        object[] smoothSeries = (((System.Collections.IEnumerable?)optionsType.GetProperty("SmoothSeries")?.GetValue(options)) ?? throw new InvalidOperationException("Expected smooth series options.")).Cast<object>().ToArray();

        TestAssert.Equal(PptxSceneChartScatterStyle.Unknown, (PptxSceneChartScatterStyle)(optionsType.GetProperty("ScatterStyle")?.GetValue(options) ?? default(PptxSceneChartScatterStyle)));
        TestAssert.True((bool)(optionsType.GetProperty("ConnectLines")?.GetValue(options) ?? true) == false, "Expected unknown scene scatterStyle to keep the no-line default, not fallback XML lineMarker.");
        TestAssert.True(smoothSeries.Length == 1 && PptxTests.ChartBooleanOptionValue(smoothSeries[0]), "Expected missing scene scatter smooth flag to use Office's true default, not fallback XML smooth=0.");
        TestAssert.True(!PptxTests.ChartBooleanOptionIsDefined(smoothSeries[0]), "Expected missing scene scatter smooth state to remain distinguishable from an explicit false token.");
    }

    public static void PptxChartSmoothOptionsPreserveRawBooleanState()
    {
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement chartElement = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:scatterChart>
                <c:scatterStyle val="lineMarker"/>
                <c:ser><c:smooth/></c:ser>
                <c:ser><c:smooth val="0"/></c:ser>
                <c:ser/>
              </c:scatterChart></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "scatterChart").Single();
        System.Reflection.MethodInfo readOptions = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartScatterOptions",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart scatter-option resolver.");
        object options = readOptions.Invoke(null, [null, chartElement]) ?? throw new InvalidOperationException("Expected resolved scatter options.");
        Type optionsType = options.GetType();
        object[] smoothSeries = (((System.Collections.IEnumerable?)optionsType.GetProperty("SmoothSeries")?.GetValue(options)) ?? throw new InvalidOperationException("Expected smooth series options.")).Cast<object>().ToArray();

        TestAssert.Equal(3, smoothSeries.Length);
        TestAssert.True(PptxTests.ChartBooleanOptionValue(smoothSeries[0]), "Expected shorthand smooth element to resolve through the shared OOXML boolean parser.");
        TestAssert.True(PptxTests.ChartBooleanOptionIsDefined(smoothSeries[0]), "Expected shorthand smooth element presence to remain explicit.");
        TestAssert.Equal(string.Empty, PptxTests.ChartBooleanOptionRawValue(smoothSeries[0]));
        TestAssert.True(!PptxTests.ChartBooleanOptionValue(smoothSeries[1]), "Expected explicit smooth=0 to resolve false.");
        TestAssert.True(PptxTests.ChartBooleanOptionIsDefined(smoothSeries[1]), "Expected explicit smooth=0 presence to remain explicit.");
        TestAssert.Equal("0", PptxTests.ChartBooleanOptionRawValue(smoothSeries[1]));
        TestAssert.True(PptxTests.ChartBooleanOptionValue(smoothSeries[2]), "Expected missing scatter smooth element to use Office's true default.");
        TestAssert.True(!PptxTests.ChartBooleanOptionIsDefined(smoothSeries[2]), "Expected missing smooth element to remain distinct from explicit false.");
        TestAssert.Equal(string.Empty, PptxTests.ChartBooleanOptionRawValue(smoothSeries[2]));
    }

    public static void PptxChartVaryColorsOptionsPreserveRawBooleanState()
    {
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        System.Reflection.MethodInfo readOptions = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartBarOptions",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart bar-option resolver.");

        object shorthand = ReadBarOptions("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:barChart><c:varyColors/></c:barChart></c:plotArea></c:chart>
            </c:chartSpace>
            """);
        object explicitFalse = ReadBarOptions("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:barChart><c:varyColors val="0"/></c:barChart></c:plotArea></c:chart>
            </c:chartSpace>
            """);
        object missing = ReadBarOptions("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:barChart/></c:plotArea></c:chart>
            </c:chartSpace>
            """);

        object shorthandVaryColors = shorthand.GetType().GetProperty("VaryColors")?.GetValue(shorthand) ?? throw new InvalidOperationException("Expected shorthand varyColors option.");
        TestAssert.True(PptxTests.ChartBooleanOptionValue(shorthandVaryColors), "Expected shorthand varyColors element to resolve through the shared OOXML boolean parser.");
        TestAssert.True(PptxTests.ChartBooleanOptionIsDefined(shorthandVaryColors), "Expected shorthand varyColors element presence to remain explicit.");
        TestAssert.Equal(string.Empty, PptxTests.ChartBooleanOptionRawValue(shorthandVaryColors));

        object explicitFalseVaryColors = explicitFalse.GetType().GetProperty("VaryColors")?.GetValue(explicitFalse) ?? throw new InvalidOperationException("Expected explicit false varyColors option.");
        TestAssert.True(!PptxTests.ChartBooleanOptionValue(explicitFalseVaryColors), "Expected explicit varyColors=0 to resolve false.");
        TestAssert.True(PptxTests.ChartBooleanOptionIsDefined(explicitFalseVaryColors), "Expected explicit varyColors=0 presence to remain explicit.");
        TestAssert.Equal("0", PptxTests.ChartBooleanOptionRawValue(explicitFalseVaryColors));

        object missingVaryColors = missing.GetType().GetProperty("VaryColors")?.GetValue(missing) ?? throw new InvalidOperationException("Expected missing varyColors option.");
        TestAssert.True(PptxTests.ChartBooleanOptionValue(missingVaryColors), "Expected missing varyColors element to use the renderer default true.");
        TestAssert.True(!PptxTests.ChartBooleanOptionIsDefined(missingVaryColors), "Expected missing varyColors element to remain distinct from explicit true.");
        TestAssert.Equal(string.Empty, PptxTests.ChartBooleanOptionRawValue(missingVaryColors));

        object ReadBarOptions(string xml)
        {
            XElement chartElement = XDocument.Parse(xml).Descendants(c + "barChart").Single();
            return readOptions.Invoke(null, [null, chartElement, PptxSceneChartGrouping.Clustered]) ?? throw new InvalidOperationException("Expected resolved bar options.");
        }
    }

    public static void PptxChartDataLabelOptionsPreserveRawBooleanState()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:plotArea><c:barChart>
                <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                <c:dLbls>
                  <c:txPr><a:bodyPr rot="3600000"/><a:lstStyle/><a:p/></c:txPr>
                  <c:leaderLines><c:spPr><a:ln w="38100"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:ln></c:spPr></c:leaderLines>
                  <c:showVal/>
                  <c:showPercent val="0"/>
                  <c:dLbl><c:idx val="0"/><c:txPr><a:bodyPr rot="-1200000"/><a:lstStyle/><a:p/></c:txPr><c:leaderLines><c:spPr><a:ln w="25400"><a:solidFill><a:srgbClr val="445566"/></a:solidFill></a:ln></c:spPr></c:leaderLines><c:showVal val="0"/><c:showLegendKey/></c:dLbl>
                </c:dLbls>
              </c:barChart></c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartPlot plot = chart.Plots[0];

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:plotArea><c:barChart>
                <c:dLbls><c:txPr><a:bodyPr rot="600000"/><a:lstStyle/><a:p/></c:txPr><c:leaderLines><c:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill></a:ln></c:spPr></c:leaderLines><c:showVal val="0"/><c:showLegendKey val="1"/><c:dLbl><c:idx val="0"/><c:txPr><a:bodyPr rot="2400000"/><a:lstStyle/><a:p/></c:txPr><c:leaderLines><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="FEDCBA"/></a:solidFill></a:ln></c:spPr></c:leaderLines><c:showVal/><c:showLegendKey val="0"/></c:dLbl></c:dLbls>
              </c:barChart></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "barChart").Single();
        System.Reflection.MethodInfo readOptions = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlDataLabelOptions",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected data-label option resolver.");
        System.Reflection.MethodInfo resolveOptions = typeof(PptxRenderer).GetMethod(
            "ResolveChartDataLabelOptions",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected data-label override resolver.");
        System.Reflection.MethodInfo readSeriesOptions = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlSeriesDataLabelOptions",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected series data-label option resolver.");
        System.Reflection.MethodInfo resolveSeriesOptions = typeof(PptxRenderer).GetMethod(
            "ResolveChartDataLabelOptionsForSeries",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected series data-label merge resolver.");
        object sceneOptions = readOptions.Invoke(null, [chart, plot, mismatchedXmlFallback, PptxTheme.Empty, PptxColorMap.Default]) ?? throw new InvalidOperationException("Expected scene-backed label options.");

        object sceneShowValue = PptxTests.ChartDataLabelFlagOption(sceneOptions, "showVal");
        TestAssert.True(PptxTests.ChartBooleanOptionValue(sceneShowValue), "Expected scene-backed shorthand showVal to resolve true.");
        TestAssert.True(PptxTests.ChartBooleanOptionIsDefined(sceneShowValue), "Expected scene-backed shorthand showVal presence to remain explicit.");
        TestAssert.Equal(string.Empty, PptxTests.ChartBooleanOptionRawValue(sceneShowValue));
        object sceneShowPercent = PptxTests.ChartDataLabelFlagOption(sceneOptions, "showPercent");
        TestAssert.True(!PptxTests.ChartBooleanOptionValue(sceneShowPercent), "Expected scene-backed showPercent=0 to resolve false.");
        TestAssert.True(PptxTests.ChartBooleanOptionIsDefined(sceneShowPercent), "Expected scene-backed showPercent=0 presence to remain explicit.");
        TestAssert.Equal("0", PptxTests.ChartBooleanOptionRawValue(sceneShowPercent));
        object sceneShowLegendKey = PptxTests.ChartDataLabelFlagOption(sceneOptions, "showLegendKey");
        TestAssert.True(!PptxTests.ChartBooleanOptionValue(sceneShowLegendKey), "Expected absent scene showLegendKey to keep its raw source value, not consume fallback XML.");
        TestAssert.True(!PptxTests.ChartBooleanOptionIsDefined(sceneShowLegendKey), "Expected missing scene showLegendKey to remain distinct from explicit false.");
        object sceneLeaderLines = PptxTests.ChartDataLabelLeaderLines(sceneOptions);
        TestAssert.True(PptxTests.ChartDataLabelLeaderLinesIsDefined(sceneLeaderLines), "Expected scene-backed data-label leader-line source to remain defined.");
        TestAssert.Equal(new RgbColor(17, 34, 51), PptxTests.ChartSeriesStrokeColor(PptxTests.ChartDataLabelLeaderLinesStroke(sceneLeaderLines)));
        TestAssert.Equal(3d, PptxTests.ChartSeriesStrokeWidth(PptxTests.ChartDataLabelLeaderLinesStroke(sceneLeaderLines)));
        object sceneBody = PptxTests.ChartDataLabelTextBodyProperties(sceneOptions);
        TestAssert.Equal(60d, PptxTests.ChartTextBodyRotationDegrees(sceneBody) ?? 0d);
        TestAssert.Equal("3600000", PptxTests.ChartTextBodyRotationValue(sceneBody));
        object sceneOverride = PptxTests.ChartDataLabelOverride(sceneOptions, 0);
        object sceneOverrideShowValue = PptxTests.ChartDataLabelFlagOption(sceneOverride, "showVal");
        TestAssert.True(!PptxTests.ChartBooleanOptionValue(sceneOverrideShowValue), "Expected scene-backed override showVal=0 to resolve false.");
        TestAssert.True(PptxTests.ChartBooleanOptionIsDefined(sceneOverrideShowValue), "Expected scene-backed override showVal=0 presence to remain explicit.");
        TestAssert.Equal("0", PptxTests.ChartBooleanOptionRawValue(sceneOverrideShowValue));
        object sceneOverrideShowLegendKey = PptxTests.ChartDataLabelFlagOption(sceneOverride, "showLegendKey");
        TestAssert.True(PptxTests.ChartBooleanOptionValue(sceneOverrideShowLegendKey), "Expected scene-backed override shorthand showLegendKey to resolve true.");
        TestAssert.True(PptxTests.ChartBooleanOptionIsDefined(sceneOverrideShowLegendKey), "Expected scene-backed override shorthand showLegendKey presence to remain explicit.");
        object sceneOverrideLeaderLines = PptxTests.ChartDataLabelLeaderLines(sceneOverride);
        TestAssert.True(PptxTests.ChartDataLabelLeaderLinesIsDefined(sceneOverrideLeaderLines), "Expected scene-backed data-label override leader-line source to remain defined.");
        TestAssert.Equal(new RgbColor(68, 85, 102), PptxTests.ChartSeriesStrokeColor(PptxTests.ChartDataLabelLeaderLinesStroke(sceneOverrideLeaderLines)));
        TestAssert.Equal(2d, PptxTests.ChartSeriesStrokeWidth(PptxTests.ChartDataLabelLeaderLinesStroke(sceneOverrideLeaderLines)));
        object sceneOverrideBody = PptxTests.ChartDataLabelTextBodyProperties(sceneOverride);
        TestAssert.Equal(-20d, PptxTests.ChartTextBodyRotationDegrees(sceneOverrideBody) ?? 0d);
        TestAssert.Equal("-1200000", PptxTests.ChartTextBodyRotationValue(sceneOverrideBody));
        object resolvedSceneOptions = resolveOptions.Invoke(null, [sceneOptions, 0]) ?? throw new InvalidOperationException("Expected resolved scene-backed label options.");
        object resolvedSceneShowValue = PptxTests.ChartDataLabelFlagOption(resolvedSceneOptions, "showVal");
        TestAssert.True(!PptxTests.ChartBooleanOptionValue(resolvedSceneShowValue), "Expected resolved scene-backed override showVal=0 to replace base shorthand showVal.");
        TestAssert.True(PptxTests.ChartBooleanOptionIsDefined(resolvedSceneShowValue), "Expected resolved scene-backed override showVal=0 presence to remain explicit.");
        TestAssert.Equal("0", PptxTests.ChartBooleanOptionRawValue(resolvedSceneShowValue));
        object resolvedSceneShowPercent = PptxTests.ChartDataLabelFlagOption(resolvedSceneOptions, "showPercent");
        TestAssert.True(!PptxTests.ChartBooleanOptionValue(resolvedSceneShowPercent), "Missing point showPercent retains its raw source value.");
        TestAssert.True(!PptxTests.ChartBooleanOptionIsDefined(resolvedSceneShowPercent), "A parent explicit false does not become a present point flag.");
        TestAssert.True((bool)resolvedSceneOptions.GetType().GetProperty("ShowPercent")!.GetValue(resolvedSceneOptions)!,
            "The defined point supplies its own enabled percentage default.");
        object resolvedSceneLeaderLines = PptxTests.ChartDataLabelLeaderLines(resolvedSceneOptions);
        TestAssert.True(PptxTests.ChartDataLabelLeaderLinesIsDefined(resolvedSceneLeaderLines), "Expected resolved scene-backed leader-line source to remain defined.");
        TestAssert.Equal(new RgbColor(68, 85, 102), PptxTests.ChartSeriesStrokeColor(PptxTests.ChartDataLabelLeaderLinesStroke(resolvedSceneLeaderLines)));
        TestAssert.Equal(2d, PptxTests.ChartSeriesStrokeWidth(PptxTests.ChartDataLabelLeaderLinesStroke(resolvedSceneLeaderLines)));
        object resolvedSceneBody = PptxTests.ChartDataLabelTextBodyProperties(resolvedSceneOptions);
        TestAssert.Equal(-20d, PptxTests.ChartTextBodyRotationDegrees(resolvedSceneBody) ?? 0d);
        TestAssert.Equal("-1200000", PptxTests.ChartTextBodyRotationValue(resolvedSceneBody));

        object xmlOptions = readOptions.Invoke(null, [null, null, mismatchedXmlFallback, PptxTheme.Empty, PptxColorMap.Default]) ?? throw new InvalidOperationException("Expected XML-backed label options.");
        object xmlShowValue = PptxTests.ChartDataLabelFlagOption(xmlOptions, "showVal");
        TestAssert.True(!PptxTests.ChartBooleanOptionValue(xmlShowValue), "Expected XML showVal=0 to resolve false.");
        TestAssert.True(PptxTests.ChartBooleanOptionIsDefined(xmlShowValue), "Expected XML showVal=0 presence to remain explicit.");
        TestAssert.Equal("0", PptxTests.ChartBooleanOptionRawValue(xmlShowValue));
        object xmlShowLegendKey = PptxTests.ChartDataLabelFlagOption(xmlOptions, "showLegendKey");
        TestAssert.True(PptxTests.ChartBooleanOptionValue(xmlShowLegendKey), "Expected XML showLegendKey=1 to resolve true.");
        TestAssert.True(PptxTests.ChartBooleanOptionIsDefined(xmlShowLegendKey), "Expected XML showLegendKey=1 presence to remain explicit.");
        TestAssert.Equal("1", PptxTests.ChartBooleanOptionRawValue(xmlShowLegendKey));
        object xmlLeaderLines = PptxTests.ChartDataLabelLeaderLines(xmlOptions);
        TestAssert.True(PptxTests.ChartDataLabelLeaderLinesIsDefined(xmlLeaderLines), "Expected XML-backed data-label leader-line source to remain defined.");
        TestAssert.Equal(new RgbColor(171, 205, 239), PptxTests.ChartSeriesStrokeColor(PptxTests.ChartDataLabelLeaderLinesStroke(xmlLeaderLines)));
        TestAssert.Equal(1d, PptxTests.ChartSeriesStrokeWidth(PptxTests.ChartDataLabelLeaderLinesStroke(xmlLeaderLines)));
        object xmlBody = PptxTests.ChartDataLabelTextBodyProperties(xmlOptions);
        TestAssert.Equal(10d, PptxTests.ChartTextBodyRotationDegrees(xmlBody) ?? 0d);
        TestAssert.Equal("600000", PptxTests.ChartTextBodyRotationValue(xmlBody));
        object xmlOverride = PptxTests.ChartDataLabelOverride(xmlOptions, 0);
        object xmlOverrideShowValue = PptxTests.ChartDataLabelFlagOption(xmlOverride, "showVal");
        TestAssert.True(PptxTests.ChartBooleanOptionValue(xmlOverrideShowValue), "Expected XML override shorthand showVal to resolve true.");
        TestAssert.True(PptxTests.ChartBooleanOptionIsDefined(xmlOverrideShowValue), "Expected XML override shorthand showVal presence to remain explicit.");
        object xmlOverrideShowLegendKey = PptxTests.ChartDataLabelFlagOption(xmlOverride, "showLegendKey");
        TestAssert.True(!PptxTests.ChartBooleanOptionValue(xmlOverrideShowLegendKey), "Expected XML override showLegendKey=0 to resolve false.");
        TestAssert.True(PptxTests.ChartBooleanOptionIsDefined(xmlOverrideShowLegendKey), "Expected XML override showLegendKey=0 presence to remain explicit.");
        TestAssert.Equal("0", PptxTests.ChartBooleanOptionRawValue(xmlOverrideShowLegendKey));
        object xmlOverrideLeaderLines = PptxTests.ChartDataLabelLeaderLines(xmlOverride);
        TestAssert.True(PptxTests.ChartDataLabelLeaderLinesIsDefined(xmlOverrideLeaderLines), "Expected XML-backed data-label override leader-line source to remain defined.");
        TestAssert.Equal(new RgbColor(254, 220, 186), PptxTests.ChartSeriesStrokeColor(PptxTests.ChartDataLabelLeaderLinesStroke(xmlOverrideLeaderLines)));
        TestAssert.Equal(0.5d, PptxTests.ChartSeriesStrokeWidth(PptxTests.ChartDataLabelLeaderLinesStroke(xmlOverrideLeaderLines)));
        object xmlOverrideBody = PptxTests.ChartDataLabelTextBodyProperties(xmlOverride);
        TestAssert.Equal(40d, PptxTests.ChartTextBodyRotationDegrees(xmlOverrideBody) ?? 0d);
        TestAssert.Equal("2400000", PptxTests.ChartTextBodyRotationValue(xmlOverrideBody));
        object resolvedXmlOptions = resolveOptions.Invoke(null, [xmlOptions, 0]) ?? throw new InvalidOperationException("Expected resolved XML-backed label options.");
        object resolvedXmlShowLegendKey = PptxTests.ChartDataLabelFlagOption(resolvedXmlOptions, "showLegendKey");
        TestAssert.True(!PptxTests.ChartBooleanOptionValue(resolvedXmlShowLegendKey), "Expected resolved XML override showLegendKey=0 to replace base showLegendKey=1.");
        TestAssert.True(PptxTests.ChartBooleanOptionIsDefined(resolvedXmlShowLegendKey), "Expected resolved XML override showLegendKey=0 presence to remain explicit.");
        TestAssert.Equal("0", PptxTests.ChartBooleanOptionRawValue(resolvedXmlShowLegendKey));
        object resolvedXmlLeaderLines = PptxTests.ChartDataLabelLeaderLines(resolvedXmlOptions);
        TestAssert.True(PptxTests.ChartDataLabelLeaderLinesIsDefined(resolvedXmlLeaderLines), "Expected resolved XML-backed leader-line source to remain defined.");
        TestAssert.Equal(new RgbColor(254, 220, 186), PptxTests.ChartSeriesStrokeColor(PptxTests.ChartDataLabelLeaderLinesStroke(resolvedXmlLeaderLines)));
        TestAssert.Equal(0.5d, PptxTests.ChartSeriesStrokeWidth(PptxTests.ChartDataLabelLeaderLinesStroke(resolvedXmlLeaderLines)));
        object resolvedXmlBody = PptxTests.ChartDataLabelTextBodyProperties(resolvedXmlOptions);
        TestAssert.Equal(40d, PptxTests.ChartTextBodyRotationDegrees(resolvedXmlBody) ?? 0d);
        TestAssert.Equal("2400000", PptxTests.ChartTextBodyRotationValue(resolvedXmlBody));

        XElement xmlWithSeriesLabels = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:barChart>
                <c:dLbls><c:showVal/><c:showPercent val="0"/><c:separator>, </c:separator></c:dLbls>
                <c:ser>
                  <c:dLbls><c:showLegendKey/><c:separator>; </c:separator></c:dLbls>
                  <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                  <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val>
                </c:ser>
              </c:barChart></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "barChart").Single();
        object plotOptions = readOptions.Invoke(null, [null, null, xmlWithSeriesLabels, PptxTheme.Empty, PptxColorMap.Default]) ?? throw new InvalidOperationException("Expected plot label options.");
        object seriesOptions = readSeriesOptions.Invoke(null, [null, null, xmlWithSeriesLabels, PptxTheme.Empty, PptxColorMap.Default]) ?? throw new InvalidOperationException("Expected series label options.");
        object resolvedSeriesOptions = resolveSeriesOptions.Invoke(null, [plotOptions, seriesOptions, 0]) ?? throw new InvalidOperationException("Expected resolved series label options.");
        object seriesShowValue = PptxTests.ChartDataLabelFlagOption(resolvedSeriesOptions, "showVal");
        TestAssert.True(!PptxTests.ChartBooleanOptionValue(seriesShowValue), "Missing series showVal retains its raw source value.");
        TestAssert.True(!PptxTests.ChartBooleanOptionIsDefined(seriesShowValue), "Missing series showVal remains absent in the active source container.");
        TestAssert.True((bool)resolvedSeriesOptions.GetType().GetProperty("ShowValue")!.GetValue(resolvedSeriesOptions)!, "The defined series supplies its own enabled value default.");
        object seriesShowPercent = PptxTests.ChartDataLabelFlagOption(resolvedSeriesOptions, "showPercent");
        TestAssert.True(!PptxTests.ChartBooleanOptionValue(seriesShowPercent), "Missing series showPercent retains its raw source value.");
        TestAssert.True(!PptxTests.ChartBooleanOptionIsDefined(seriesShowPercent), "A parent explicit false does not become a present series flag.");
        TestAssert.True((bool)resolvedSeriesOptions.GetType().GetProperty("ShowPercent")!.GetValue(resolvedSeriesOptions)!, "The defined series supplies its own enabled percentage default.");
        object seriesShowLegendKey = PptxTests.ChartDataLabelFlagOption(resolvedSeriesOptions, "showLegendKey");
        TestAssert.True(PptxTests.ChartBooleanOptionValue(seriesShowLegendKey), "Expected series-level showLegendKey to override the plot-level default.");
        TestAssert.True(PptxTests.ChartBooleanOptionIsDefined(seriesShowLegendKey), "Expected series-level showLegendKey presence to remain explicit.");
        TestAssert.Equal("; ", (string?)resolvedSeriesOptions.GetType().GetProperty("Separator")?.GetValue(resolvedSeriesOptions) ?? string.Empty);
    }

    public static void PptxChartRadarOptionsUseSceneAuthoritativeDefaults()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:radarChart>
                  <c:radarStyle val="bogus"/>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt><c:pt idx="2"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:radarChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartPlot plot = chart.Plots[0];
        TestAssert.Equal(PptxSceneChartRadarStyle.Unknown, plot.RadarStyleKind);

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:radarChart><c:radarStyle val="filled"/></c:radarChart></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "radarChart").Single();
        System.Reflection.MethodInfo readOptions = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartRadarOptions",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart radar-option resolver.");
        object options = readOptions.Invoke(null, [plot, mismatchedXmlFallback]) ?? throw new InvalidOperationException("Expected resolved radar options.");
        Type optionsType = options.GetType();

        TestAssert.Equal(PptxSceneChartRadarStyle.Standard, (PptxSceneChartRadarStyle)(optionsType.GetProperty("RadarStyle")?.GetValue(options) ?? default(PptxSceneChartRadarStyle)));
    }

    public static void PptxChartPolarPointOptionsUseSceneAuthoritativeDefaults()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:pieChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:pieChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartPlot plot = chart.Plots[0];
        TestAssert.True(plot.FirstSliceAngle is null, "Expected missing firstSliceAng metadata to remain distinct from the effective renderer default.");
        TestAssert.True(plot.Series[0].PointStyles.Count == 0, "Expected no scene point styles.");

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:pieChart>
                <c:ser><c:dPt><c:idx val="0"/><c:explosion val="25"/></c:dPt></c:ser>
                <c:firstSliceAng val="270"/>
              </c:pieChart></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "pieChart").Single();
        System.Reflection.MethodInfo readOptions = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartPolarPointOptions",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart polar-point option resolver.");
        object options = readOptions.Invoke(null, [plot, mismatchedXmlFallback, PptxTheme.Empty, PptxColorMap.Default, null]) ?? throw new InvalidOperationException("Expected resolved polar point options.");
        Type optionsType = options.GetType();
        object pointExplosions = optionsType.GetProperty("PointExplosions")?.GetValue(options) ?? throw new InvalidOperationException("Expected point explosion options.");

        TestAssert.True(((System.Collections.IEnumerable)pointExplosions).Cast<object>().Count() == 0, "Expected missing scene point explosions to stay empty, not import fallback XML dPt explosion.");
        TestAssert.Equal(0d, (double)(optionsType.GetProperty("FirstSliceAngle")?.GetValue(options) ?? double.NaN));
    }

    public static void PptxChartDoughnutOptionsUseSceneAuthoritativeDefaults()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:doughnutChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:doughnutChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartPlot plot = chart.Plots[0];
        TestAssert.True(plot.HoleSize is null, "Expected missing holeSize metadata to remain distinct from the effective renderer default.");

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:doughnutChart>
                <c:holeSize val="10"/>
                <c:firstSliceAng val="270"/>
              </c:doughnutChart></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "doughnutChart").Single();
        System.Reflection.MethodInfo readOptions = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartDoughnutOptions",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart doughnut option resolver.");
        object options = readOptions.Invoke(null, [plot, mismatchedXmlFallback, PptxTheme.Empty, PptxColorMap.Default, null]) ?? throw new InvalidOperationException("Expected resolved doughnut options.");
        Type optionsType = options.GetType();
        object polarPoints = optionsType.GetProperty("PolarPoints")?.GetValue(options) ?? throw new InvalidOperationException("Expected polar point options.");

        TestAssert.Equal(0.56d, (double)(optionsType.GetProperty("HoleSize")?.GetValue(options) ?? double.NaN));
        TestAssert.Equal(0d, (double)(polarPoints.GetType().GetProperty("FirstSliceAngle")?.GetValue(polarPoints) ?? double.NaN));
    }

    public static void PptxChartUnknownScatterStyleKeepsSceneLineConnectionDefault()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:scatterChart>
                  <c:scatterStyle val="bogus"/>
                  <c:ser><c:xVal><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:xVal><c:yVal><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:yVal></c:ser>
                </c:scatterChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        TestAssert.Equal(PptxSceneChartScatterStyle.Unknown, chart.Plots[0].ScatterStyleKind);
        TestAssert.Equal("bogus", chart.Plots[0].ScatterStyle);

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:scatterChart><c:scatterStyle val="lineMarker"/></c:scatterChart></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "scatterChart").Single();
        System.Reflection.MethodInfo readOptions = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartScatterOptions",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart scatter-option resolver.");
        object options = readOptions.Invoke(null, [chart.Plots[0], mismatchedXmlFallback])!;
        object scatterStyle = options.GetType().GetProperty("ScatterStyle")!.GetValue(options)!;
        object connectLines = options.GetType().GetProperty("ConnectLines")!.GetValue(options)!;

        TestAssert.Equal(PptxSceneChartScatterStyle.Unknown, (PptxSceneChartScatterStyle)scatterStyle);
        TestAssert.True((bool)connectLines == false, "Expected unknown scatterStyle to keep the scene-owned no-line default instead of borrowing XML fallback lineMarker.");
    }

    public static void PptxChartUnknownRadarStyleUsesSceneAuthoritativeDefault()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:radarChart>
                  <c:radarStyle val="bogus"/>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt><c:pt idx="2"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:radarChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        TestAssert.Equal(PptxSceneChartRadarStyle.Unknown, chart.Plots[0].RadarStyleKind);
        TestAssert.Equal("bogus", chart.Plots[0].RadarStyle);

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:radarChart><c:radarStyle val="filled"/></c:radarChart></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "radarChart").Single();
        System.Reflection.MethodInfo readRadarStyle = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartRadarStyle",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart radar-style resolver.");
        object radarStyle = readRadarStyle.Invoke(null, [chart.Plots[0], mismatchedXmlFallback]) ?? throw new InvalidOperationException("Expected resolved radar style.");

        TestAssert.Equal(PptxSceneChartRadarStyle.Standard, (PptxSceneChartRadarStyle)radarStyle);
    }

    public static void PptxChartUnknownAxisOrientationUsesSceneAuthoritativeDefault()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:lineChart>
                <c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="20"/></c:catAx>
                <c:valAx><c:axId val="20"/><c:scaling><c:orientation val="bogus"/></c:scaling><c:axPos val="l"/><c:crossAx val="10"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartAxis valueAxis = chart.Axes.First(axis => axis.AxisKind == PptxSceneChartAxisKind.Value);
        TestAssert.Equal(PptxSceneChartAxisOrientation.Unknown, valueAxis.OrientationKind);
        TestAssert.Equal("bogus", valueAxis.Orientation);

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:valAx><c:scaling><c:orientation val="maxMin"/></c:scaling></c:valAx></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "valAx").Single();
        System.Reflection.MethodInfo readAxisReversed = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlValueAxisReversed",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected value-axis orientation resolver.");
        object reversed = readAxisReversed.Invoke(null, [valueAxis, mismatchedXmlFallback]) ?? throw new InvalidOperationException("Expected resolved axis orientation.");

        TestAssert.True(!(bool)reversed, "Unknown scene-owned value-axis orientation should resolve to the Office default minMax, not fall back to mismatched XML.");
    }

    public static void PptxChartUnknownAxisCrossesUsesSceneAuthoritativeDefault()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:lineChart>
                <c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="20"/></c:catAx>
                <c:valAx><c:axId val="20"/><c:scaling><c:min val="-5"/><c:max val="5"/></c:scaling><c:axPos val="l"/><c:crossAx val="10"/><c:crosses val="bogus"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartAxis valueAxis = chart.Axes.First(axis => axis.AxisKind == PptxSceneChartAxisKind.Value);
        TestAssert.Equal(PptxSceneChartAxisCrosses.Unknown, valueAxis.CrossesKind);
        TestAssert.Equal("bogus", valueAxis.Crosses);

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:valAx><c:crosses val="max"/></c:valAx></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "valAx").Single();
        Type extentsType = typeof(PptxRenderer).GetNestedType(
            "ChartValueExtents",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart value extents type.");
        object extents = Activator.CreateInstance(extentsType, [-5d, 5d]) ?? throw new InvalidOperationException("Expected chart value extents.");
        System.Reflection.MethodInfo readCrossingValue = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlValueAxisCrossingValue",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected value-axis crossing resolver.");
        object crossingValue = readCrossingValue.Invoke(null, [valueAxis, mismatchedXmlFallback, extents]) ?? throw new InvalidOperationException("Expected resolved axis crossing.");

        TestAssert.Equal(0d, (double)crossingValue);
    }

    public static void PptxChartInvalidAxisCrossesAtUsesSceneAuthoritativeCrosses()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:lineChart>
                <c:catAx><c:axId val="10"/><c:axPos val="b"/><c:crossAx val="20"/></c:catAx>
                <c:valAx><c:axId val="20"/><c:scaling><c:min val="-5"/><c:max val="5"/></c:scaling><c:axPos val="l"/><c:crossAx val="10"/><c:crosses val="max"/><c:crossesAt val="futureCross"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartAxis valueAxis = chart.Axes.First(axis => axis.AxisKind == PptxSceneChartAxisKind.Value);
        TestAssert.Equal(PptxSceneChartAxisCrosses.Maximum, valueAxis.CrossesKind);
        TestAssert.True(valueAxis.CrossesAt is null, "Expected unparseable scene crossesAt to stay distinct from an authored value.");
        TestAssert.Equal("futureCross", valueAxis.CrossesAtValue);

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:valAx><c:crossesAt val="2"/></c:valAx></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "valAx").Single();
        Type extentsType = typeof(PptxRenderer).GetNestedType(
            "ChartValueExtents",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart value extents type.");
        object extents = Activator.CreateInstance(extentsType, [-5d, 5d]) ?? throw new InvalidOperationException("Expected chart value extents.");
        System.Reflection.MethodInfo readCrossingValue = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlValueAxisCrossingValue",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected value-axis crossing resolver.");
        object sceneCrossingValue = readCrossingValue.Invoke(null, [valueAxis, mismatchedXmlFallback, extents]) ?? throw new InvalidOperationException("Expected resolved scene axis crossing.");
        object xmlOnlyCrossingValue = readCrossingValue.Invoke(null, [null, mismatchedXmlFallback, extents]) ?? throw new InvalidOperationException("Expected resolved XML-only axis crossing.");

        TestAssert.Equal(5d, (double)sceneCrossingValue);
        TestAssert.Equal(2d, (double)xmlOnlyCrossingValue);
    }

    public static void PptxChartUnknownAxisCrossBetweenUsesSceneAuthoritativeDefault()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:lineChart>
                <c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="20"/></c:catAx>
                <c:valAx><c:axId val="20"/><c:scaling><c:min val="0"/><c:max val="5"/></c:scaling><c:axPos val="l"/><c:crossAx val="10"/><c:crossBetween val="bogus"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartAxis valueAxis = chart.Axes.First(axis => axis.AxisKind == PptxSceneChartAxisKind.Value);
        TestAssert.Equal(PptxSceneChartAxisCrossBetween.Unknown, valueAxis.CrossBetweenKind);
        TestAssert.Equal("bogus", valueAxis.CrossBetween);

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:valAx><c:crossBetween val="midCat"/></c:valAx></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "valAx").Single();
        System.Reflection.MethodInfo labelsOnTickMarks = typeof(PptxRenderer).GetMethod(
            "ResolveSceneOrXmlCategoryAxisLabelsOnTickMarks",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected category-axis label placement resolver.");
        object onTickMarks = labelsOnTickMarks.Invoke(null, [valueAxis, mismatchedXmlFallback]) ?? throw new InvalidOperationException("Expected resolved category-axis label placement.");

        TestAssert.True(!(bool)onTickMarks, "Unknown scene-owned crossBetween should resolve to the Office default between, not fall back to mismatched XML.");
    }

    public static void PptxChartUnknownTickLabelPositionUsesSceneAuthoritativeDefault()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:lineChart>
                <c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="20"/><c:tickLblPos val="bogus"/></c:catAx>
                <c:valAx><c:axId val="20"/><c:scaling><c:min val="0"/><c:max val="5"/></c:scaling><c:axPos val="l"/><c:crossAx val="10"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartAxis categoryAxis = chart.Axes.First(axis => axis.AxisKind == PptxSceneChartAxisKind.Category);
        TestAssert.Equal(PptxSceneChartTickLabelPosition.Unknown, categoryAxis.TickLabelPositionKind);
        TestAssert.Equal("bogus", categoryAxis.TickLabelPosition);

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:catAx><c:tickLblPos val="none"/></c:catAx></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "catAx").Single();
        System.Reflection.MethodInfo labelVisible = typeof(PptxRenderer).GetMethod(
            "IsSceneOrXmlChartAxisLabelVisible",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected axis label visibility resolver.");
        object visible = labelVisible.Invoke(null, [categoryAxis, mismatchedXmlFallback]) ?? throw new InvalidOperationException("Expected resolved axis label visibility.");

        TestAssert.True((bool)visible, "Unknown scene-owned tickLblPos should resolve to the Office default nextTo/visible, not fall back to mismatched XML.");
    }

    public static void PptxChartMissingMajorTickMarkResolvesThroughExplicitDefault()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:lineChart>
                <c:catAx><c:axId val="10"/><c:axPos val="b"/><c:crossAx val="20"/></c:catAx>
                <c:valAx><c:axId val="20"/><c:axPos val="l"/><c:crossAx val="10"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartAxis categoryAxis = chart.Axes.First(axis => axis.AxisKind == PptxSceneChartAxisKind.Category);
        TestAssert.Equal(PptxSceneChartAxisTickMark.Unknown, categoryAxis.MajorTickMarkKind);
        TestAssert.Equal(string.Empty, categoryAxis.MajorTickMark);

        System.Reflection.MethodInfo readMajorTickMark = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartAxisMajorTickMark",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart major tick-mark resolver.");
        object tickMark = readMajorTickMark.Invoke(null, [categoryAxis, null]) ?? throw new InvalidOperationException("Expected resolved major tick mark.");

        TestAssert.Equal(PptxSceneChartAxisTickMark.None, (PptxSceneChartAxisTickMark)tickMark);
    }

    public static void PptxChartValueAxisRenderOptionsUseSceneAuthoritativeDefaults()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:lineChart>
                <c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="20"/></c:catAx>
                <c:valAx><c:axId val="20"/><c:scaling><c:min val="-5"/><c:max val="5"/><c:orientation val="bogus"/></c:scaling><c:axPos val="l"/><c:crossAx val="10"/><c:crosses val="bogus"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartAxis valueAxis = chart.Axes.First(axis => axis.AxisKind == PptxSceneChartAxisKind.Value);

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:plotArea><c:valAx>
                <c:scaling><c:orientation val="maxMin"/></c:scaling>
                <c:crosses val="max"/>
                <c:majorUnit val="0.25"/>
                <c:majorGridlines><c:spPr><a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>
              </c:valAx></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "valAx").Single();
        Type extentsType = typeof(PptxRenderer).GetNestedType(
            "ChartValueExtents",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart value extents type.");
        object extents = Activator.CreateInstance(extentsType, [-5d, 5d]) ?? throw new InvalidOperationException("Expected chart value extents.");
        System.Reflection.MethodInfo readOptions = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartValueAxisRenderOptions",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart value-axis option resolver.");
        object options = readOptions.Invoke(null, [valueAxis, mismatchedXmlFallback, PptxTheme.Empty, extents, true]) ?? throw new InvalidOperationException("Expected value-axis render options.");
        Type optionsType = options.GetType();

        object units = optionsType.GetProperty("Units")?.GetValue(options) ?? throw new InvalidOperationException("Expected axis units.");
        TestAssert.Equal(0.1d, (double)(units.GetType().GetProperty("MajorUnit")?.GetValue(units) ?? double.NaN));
        TestAssert.True((bool)(optionsType.GetProperty("Reversed")?.GetValue(options) ?? true) == false, "Expected unknown scene orientation to use the default minMax, not fallback XML maxMin.");
        TestAssert.Equal(0d, (double)(optionsType.GetProperty("CrossingValue")?.GetValue(options) ?? double.NaN));
        TestAssert.True((bool)(optionsType.GetProperty("MajorGridlines")?.GetValue(options) ?? true) == false, "Expected missing scene gridlines to stay missing, not import fallback XML gridlines.");
    }

    public static void PptxChartAxisNumberFormatUsesSceneAuthoritativeAbsence()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>1.25</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:lineChart>
                <c:catAx><c:axId val="10"/><c:axPos val="b"/><c:crossAx val="20"/></c:catAx>
                <c:valAx><c:axId val="20"/><c:axPos val="l"/><c:crossAx val="10"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartAxis valueAxis = chart.Axes.First(axis => axis.AxisKind == PptxSceneChartAxisKind.Value);
        TestAssert.True(valueAxis.NumberFormatInfo.IsDefined == false, "Expected scene value axis to carry an authoritative missing number format.");

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:valAx><c:axId val="20"/><c:numFmt formatCode="0%" sourceLinked="0"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "valAx").Single();
        System.Reflection.MethodInfo formatAxisLabel = typeof(PptxRenderer).GetMethod(
            "FormatSceneOrXmlChartAxisLabel",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart axis label formatter.");

        string sceneLabel = (string?)formatAxisLabel.Invoke(null, [1.25d, valueAxis, mismatchedXmlFallback, null]) ?? string.Empty;
        string xmlOnlyLabel = (string?)formatAxisLabel.Invoke(null, [1.25d, null, mismatchedXmlFallback, null]) ?? string.Empty;

        TestAssert.Equal("1.25", sceneLabel);
        TestAssert.Equal("125%", xmlOnlyLabel);
    }

    public static void PptxChartAxisTextStyleUsesSceneAuthoritativeMissingAxis()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:lineChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XDocument mismatchedXml = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:plotArea>
                <c:valAx>
                  <c:txPr><a:p><a:pPr><a:defRPr sz="4000" b="1"/></a:pPr></a:p></c:txPr>
                </c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """);
        XElement mismatchedAxis = mismatchedXml.Descendants(c + "valAx").Single();

        System.Reflection.MethodInfo readStyle = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartTextStyle",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart text-style resolver.");

        object sceneStyle = readStyle.Invoke(null, [PptxTheme.Empty, chart, null, mismatchedXml, mismatchedAxis, 10d, null]) ?? throw new InvalidOperationException("Expected scene text style.");
        object xmlStyle = readStyle.Invoke(null, [PptxTheme.Empty, null, null, mismatchedXml, mismatchedAxis, 10d, null]) ?? throw new InvalidOperationException("Expected XML text style.");

        TestAssert.Equal(10d, (double)(sceneStyle.GetType().GetProperty("FontSize")?.GetValue(sceneStyle) ?? double.NaN));
        TestAssert.True((bool)(sceneStyle.GetType().GetProperty("Bold")?.GetValue(sceneStyle) ?? true) == false, "Expected missing scene axis to keep scene defaults, not import fallback XML text properties.");
        TestAssert.Equal(40d, (double)(xmlStyle.GetType().GetProperty("FontSize")?.GetValue(xmlStyle) ?? double.NaN));
        TestAssert.True((bool)(xmlStyle.GetType().GetProperty("Bold")?.GetValue(xmlStyle) ?? false), "Expected XML-only compatibility to keep reading axis text properties.");
    }

    public static void PptxChartTextStyleCarriesTypefaceSourceThroughRendererStyle()
    {
        const string chartXml = """
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:txPr><a:p><a:pPr><a:defRPr sz="1100" spc="125"><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr>
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:lineChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """;
        PptxSceneChart chart = PptxTests.BuildSingleChartScene(chartXml) ?? throw new InvalidOperationException("Expected chart scene.");
        XDocument document = XDocument.Parse(chartXml);
        System.Reflection.MethodInfo readStyle = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartTextStyle",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart text-style resolver.");

        object style = readStyle.Invoke(null, [PptxTheme.Empty, chart, null, document, document.Root, 10d, null]) ?? throw new InvalidOperationException("Expected chart text style.");

        TestAssert.Equal("Arial", (string?)style.GetType().GetProperty("FontFamily")?.GetValue(style) ?? string.Empty);
        TestAssert.Equal(1.25d, (double)(style.GetType().GetProperty("CharacterSpacing")?.GetValue(style) ?? double.NaN));
        TestAssert.Equal("Arial", (string?)style.GetType().GetProperty("RequestedTypeface")?.GetValue(style) ?? string.Empty);
        TestAssert.Equal(PptxThemeTypefaceSource.Direct, (PptxThemeTypefaceSource?)style.GetType().GetProperty("TypefaceSource")?.GetValue(style) ?? default);
    }

    public static void PptxChartUnknownDisplayBlanksAsUsesSceneAuthoritativeDefault()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:lineChart>
              </c:plotArea><c:dispBlanksAs val="bogus"/></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        TestAssert.Equal(PptxSceneChartDisplayBlanksAs.Unknown, chart.Options.DisplayBlanksAsKind);
        TestAssert.Equal("bogus", chart.Options.DisplayBlanksAs);

        XDocument mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:dispBlanksAs val="span"/></c:chart>
            </c:chartSpace>
            """);
        System.Reflection.MethodInfo readDisplayBlanksAs = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartDisplayBlanksAs",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected display-blanks resolver.");
        object displayBlanksAs = readDisplayBlanksAs.Invoke(null, [chart, mismatchedXmlFallback]) ?? throw new InvalidOperationException("Expected resolved display-blanks value.");

        TestAssert.Equal(PptxSceneChartDisplayBlanksAs.Gap, (PptxSceneChartDisplayBlanksAs)displayBlanksAs);
    }

    public static void PptxChartPlotElementSelectionUsesSceneSource()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:barChart>
                  <c:barDir val="bar"/>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:barChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");

        XDocument mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser></c:lineChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """);
        System.Reflection.MethodInfo readPlotElement = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlFirstChartPlotElement",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected plot element resolver.");
        object sceneBacked = readPlotElement.Invoke(null, [chart, mismatchedXmlFallback, PptxSceneChartPlotKind.Bar]) ?? throw new InvalidOperationException("Expected scene-backed bar plot.");
        object? xmlBacked = readPlotElement.Invoke(null, [null, mismatchedXmlFallback, PptxSceneChartPlotKind.Bar]);

        TestAssert.Equal("barChart", ((XElement)sceneBacked).Name.LocalName);
        TestAssert.True(xmlBacked is null, "Expected XML-only fallback to report the missing bar plot instead of using scene data.");
    }

    public static void PptxChartPlotElementsSelectionUsesSceneSources()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:barChart>
                  <c:barDir val="col"/>
                  <c:axId val="10"/><c:axId val="20"/>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:barChart>
                <c:barChart>
                  <c:barDir val="bar"/>
                  <c:axId val="10"/><c:axId val="30"/>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>B</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>3</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:barChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");

        XDocument mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser></c:lineChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """);
        System.Reflection.MethodInfo readPlotElements = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartPlotElements",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected plot elements resolver.");
        object sceneBacked = readPlotElements.Invoke(null, [chart, mismatchedXmlFallback, PptxSceneChartPlotKind.Bar]) ?? throw new InvalidOperationException("Expected scene-backed bar plots.");
        object xmlBacked = readPlotElements.Invoke(null, [null, mismatchedXmlFallback, PptxSceneChartPlotKind.Bar]) ?? throw new InvalidOperationException("Expected XML-backed plot list.");

        IReadOnlyList<XElement> scenePlots = (IReadOnlyList<XElement>)sceneBacked;
        IReadOnlyList<XElement> xmlPlots = (IReadOnlyList<XElement>)xmlBacked;
        XNamespace chartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        TestAssert.Equal(2, scenePlots.Count);
        TestAssert.Equal("col", scenePlots[0].Element(chartNamespace + "barDir")?.Attribute("val")?.Value ?? string.Empty);
        TestAssert.Equal("bar", scenePlots[1].Element(chartNamespace + "barDir")?.Attribute("val")?.Value ?? string.Empty);
        TestAssert.Equal(0, xmlPlots.Count);
    }

    public static void PptxChartSecondaryRightValueAxisSelectionUsesSceneSource()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:barChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:barChart>
                <c:barChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>B</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>30</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:axId val="10"/><c:axId val="30"/>
                </c:barChart>
                <c:catAx><c:axId val="10"/><c:axPos val="b"/><c:crossAx val="20"/></c:catAx>
                <c:valAx><c:axId val="20"/><c:axPos val="l"/><c:crossAx val="10"/></c:valAx>
                <c:valAx><c:axId val="30"/><c:axPos val="r"/><c:crossAx val="10"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");

        XDocument mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:valAx><c:axId val="20"/><c:axPos val="r"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """);
        System.Reflection.MethodInfo readSecondaryAxis = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlSecondaryRightValueAxis",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected secondary value-axis resolver.");
        object sceneBacked = readSecondaryAxis.Invoke(null, [chart, mismatchedXmlFallback]) ?? throw new InvalidOperationException("Expected scene-backed secondary value axis.");
        object xmlBacked = readSecondaryAxis.Invoke(null, [null, mismatchedXmlFallback]) ?? throw new InvalidOperationException("Expected XML-backed secondary value axis.");

        PptxSceneChartAxis? sceneAxis = (PptxSceneChartAxis?)sceneBacked.GetType().GetProperty("SceneAxis")?.GetValue(sceneBacked);
        XElement? sceneXmlAxis = (XElement?)sceneBacked.GetType().GetProperty("XmlAxis")?.GetValue(sceneBacked);
        PptxSceneChartAxis? xmlSceneAxis = (PptxSceneChartAxis?)xmlBacked.GetType().GetProperty("SceneAxis")?.GetValue(xmlBacked);
        XElement? xmlAxis = (XElement?)xmlBacked.GetType().GetProperty("XmlAxis")?.GetValue(xmlBacked);

        XNamespace chartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        TestAssert.Equal("30", sceneAxis?.Id ?? string.Empty);
        TestAssert.Equal("30", sceneXmlAxis?.Element(chartNamespace + "axId")?.Attribute("val")?.Value ?? string.Empty);
        TestAssert.True(xmlSceneAxis is null, "Expected XML-only secondary axis selection to remain XML-backed.");
        TestAssert.Equal("20", xmlAxis?.Element(chartNamespace + "axId")?.Attribute("val")?.Value ?? string.Empty);
    }

    public static void PptxChartRotatedTextClipMapsIntoTransformedSpace()
    {
        Type textRunType = typeof(PptxRenderer).GetNestedType(
            "TextRun",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart text run.");
        Type alignmentType = typeof(PptxRenderer).GetNestedType(
            "TextAlignment",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected text alignment.");
        System.Reflection.MethodInfo inverseClip = typeof(PptxRenderer).GetMethod(
            "InverseTransformClip",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected rotated-clip bridge.");
        object center = Enum.ToObject(alignmentType, 1);
        object? plainRun = Activator.CreateInstance(textRunType, ["", 0d, 0d, 0d, 0d, 0d, 0d, 0d, 0d, 12d, 0d, 0d, new RgbColor(0, 0, 0), 1d, null, false, false, false, false, true, center, null, 0d, 0d, 0d, false, false, false, null, false, 0]);
        object? rotatedRun = Activator.CreateInstance(textRunType, ["", 0d, 0d, 0d, 0d, 0d, 0d, 0d, 0d, 12d, 0d, 0d, new RgbColor(0, 0, 0), 1d, null, false, false, false, false, true, center, null, -90d, 100d, 200d, false, false, false, null, false, 0]);
        (double plainX, double plainY, double plainW, double plainH) = ((double, double, double, double))(inverseClip.Invoke(null, [plainRun, 10d, 20d, 30d, 40d]) ?? throw new InvalidOperationException("Expected clip."));
        (double rotX, double rotY, double rotW, double rotH) = ((double, double, double, double))(inverseClip.Invoke(null, [rotatedRun, 10d, 20d, 30d, 40d]) ?? throw new InvalidOperationException("Expected clip."));
        TestAssert.Equal((10d, 20d, 30d, 40d), (plainX, plainY, plainW, plainH));
        TestAssert.True(Math.Abs(rotX - -80d) < 1e-6 && Math.Abs(rotY - 260d) < 1e-6 && Math.Abs(rotW - 40d) < 1e-6 && Math.Abs(rotH - 30d) < 1e-6, "Expected a -90-degree clip to map frame bounds into rotated space, covering text past the frame edge.");
    }

    public static void PptxChartSecondaryValueAxisStyleSelectionUsesSceneSource()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:barChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:barChart>
                <c:barChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>B</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>30</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:axId val="10"/><c:axId val="30"/>
                </c:barChart>
                <c:catAx><c:axId val="10"/><c:axPos val="b"/><c:crossAx val="20"/></c:catAx>
                <c:valAx><c:axId val="20"/><c:axPos val="l"/><c:crossAx val="10"/></c:valAx>
                <c:valAx><c:axId val="30"/><c:axPos val="r"/><c:crossAx val="10"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");

        XDocument mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:barChart><c:axId val="10"/><c:axId val="20"/></c:barChart>
                <c:catAx><c:axId val="10"/></c:catAx>
                <c:valAx><c:axId val="20"/><c:axPos val="l"/></c:valAx>
                <c:valAx><c:axId val="99"/><c:axPos val="r"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """);
        XNamespace chartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedPlot = mismatchedXmlFallback.Descendants(chartNamespace + "barChart").First();
        System.Reflection.MethodInfo readValueAxes = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartValueAxesForPlot",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected renderer value-axis bridge.");
        System.Reflection.MethodInfo readSecondaryAxis = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlSecondaryValueAxisForChart",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected secondary style value-axis resolver.");

        object scenePrimaryAxis = ((System.Collections.IEnumerable)(readValueAxes.Invoke(null, [chart, chart.Plots[0], mismatchedXmlFallback, mismatchedPlot]) ?? throw new InvalidOperationException("Expected scene value axes."))).Cast<object>().First();
        object xmlPrimaryAxis = ((System.Collections.IEnumerable)(readValueAxes.Invoke(null, [null, null, mismatchedXmlFallback, mismatchedPlot]) ?? throw new InvalidOperationException("Expected XML value axes."))).Cast<object>().First();
        object sceneBacked = readSecondaryAxis.Invoke(null, [chart, mismatchedXmlFallback, scenePrimaryAxis]) ?? throw new InvalidOperationException("Expected scene-backed secondary style axis.");
        object xmlBacked = readSecondaryAxis.Invoke(null, [null, mismatchedXmlFallback, xmlPrimaryAxis]) ?? throw new InvalidOperationException("Expected XML-backed secondary style axis.");

        PptxSceneChartAxis? sceneAxis = (PptxSceneChartAxis?)sceneBacked.GetType().GetProperty("SceneAxis")?.GetValue(sceneBacked);
        XElement? sceneXmlAxis = (XElement?)sceneBacked.GetType().GetProperty("XmlAxis")?.GetValue(sceneBacked);
        PptxSceneChartAxis? xmlSceneAxis = (PptxSceneChartAxis?)xmlBacked.GetType().GetProperty("SceneAxis")?.GetValue(xmlBacked);
        XElement? xmlAxis = (XElement?)xmlBacked.GetType().GetProperty("XmlAxis")?.GetValue(xmlBacked);

        TestAssert.Equal("30", sceneAxis?.Id ?? string.Empty);
        TestAssert.Equal("30", sceneXmlAxis?.Element(chartNamespace + "axId")?.Attribute("val")?.Value ?? string.Empty);
        TestAssert.True(xmlSceneAxis is null, "Expected XML-only secondary style axis selection to remain XML-backed.");
        TestAssert.Equal("99", xmlAxis?.Element(chartNamespace + "axId")?.Attribute("val")?.Value ?? string.Empty);
    }

    public static void PptxChartSeriesStylesUseSceneSource()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:plotArea><c:barChart>
                <c:ser>
                  <c:spPr>
                    <a:solidFill><a:srgbClr val="112233"/></a:solidFill>
                    <a:ln w="25400"><a:solidFill><a:srgbClr val="445566"/></a:solidFill></a:ln>
                  </c:spPr>
                  <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                  <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val>
                </c:ser>
              </c:barChart></c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartPlot plot = chart.Plots[0];

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:plotArea><c:barChart>
                <c:ser>
                  <c:spPr>
                    <a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill>
                    <a:ln w="12700"><a:solidFill><a:srgbClr val="FEDCBA"/></a:solidFill></a:ln>
                  </c:spPr>
                </c:ser>
              </c:barChart></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "barChart").Single();
        System.Reflection.MethodInfo readFills = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlSeriesFills",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected series fill resolver.");
        System.Reflection.MethodInfo readStrokes = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlSeriesStrokes",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected series stroke resolver.");

        object sceneFill = (((System.Collections.IEnumerable?)readFills.Invoke(null, [plot, mismatchedXmlFallback, PptxTheme.Empty, PptxColorMap.Default])) ?? throw new InvalidOperationException("Expected scene fills.")).Cast<object>().Single();
        object sceneStroke = (((System.Collections.IEnumerable?)readStrokes.Invoke(null, [plot, mismatchedXmlFallback, PptxTheme.Empty, PptxColorMap.Default, null])) ?? throw new InvalidOperationException("Expected scene strokes.")).Cast<object>().Single();
        object xmlFill = (((System.Collections.IEnumerable?)readFills.Invoke(null, [null, mismatchedXmlFallback, PptxTheme.Empty, PptxColorMap.Default])) ?? throw new InvalidOperationException("Expected XML fills.")).Cast<object>().Single();
        object xmlStroke = (((System.Collections.IEnumerable?)readStrokes.Invoke(null, [null, mismatchedXmlFallback, PptxTheme.Empty, PptxColorMap.Default, null])) ?? throw new InvalidOperationException("Expected XML strokes.")).Cast<object>().Single();

        TestAssert.Equal(new RgbColor(17, 34, 51), PptxTests.ChartSeriesFillColor(sceneFill));
        TestAssert.Equal(new RgbColor(68, 85, 102), PptxTests.ChartSeriesStrokeColor(sceneStroke));
        TestAssert.Equal(2d, PptxTests.ChartSeriesStrokeWidth(sceneStroke));
        TestAssert.Equal(new RgbColor(171, 205, 239), PptxTests.ChartSeriesFillColor(xmlFill));
        TestAssert.Equal(new RgbColor(254, 220, 186), PptxTests.ChartSeriesStrokeColor(xmlStroke));
        TestAssert.Equal(1d, PptxTests.ChartSeriesStrokeWidth(xmlStroke));
    }

    public static void PptxChartSeriesPointStylesUseSceneSource()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:plotArea><c:barChart>
                <c:ser>
                  <c:dPt>
                    <c:idx val="1"/>
                    <c:spPr>
                      <a:solidFill><a:srgbClr val="123456"/></a:solidFill>
                      <a:ln w="38100"><a:solidFill><a:srgbClr val="654321"/></a:solidFill></a:ln>
                    </c:spPr>
                  </c:dPt>
                  <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                  <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt></c:numLit></c:val>
                </c:ser>
              </c:barChart></c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartPlot plot = chart.Plots[0];

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:plotArea><c:barChart>
                <c:ser>
                  <c:dPt>
                    <c:idx val="0"/>
                    <c:spPr>
                      <a:solidFill><a:srgbClr val="A0B0C0"/></a:solidFill>
                      <a:ln w="12700"><a:solidFill><a:srgbClr val="0C0B0A"/></a:solidFill></a:ln>
                    </c:spPr>
                  </c:dPt>
                </c:ser>
              </c:barChart></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "barChart").Single();
        System.Reflection.MethodInfo readPointFills = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlSeriesPointFills",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected series point fill resolver.");
        System.Reflection.MethodInfo readPointStrokes = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlSeriesPointStrokes",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected series point stroke resolver.");

        object scenePointFills = (((System.Collections.IEnumerable?)readPointFills.Invoke(null, [plot, mismatchedXmlFallback, PptxTheme.Empty, PptxColorMap.Default])) ?? throw new InvalidOperationException("Expected scene point fills.")).Cast<object>().Single();
        object scenePointStrokes = (((System.Collections.IEnumerable?)readPointStrokes.Invoke(null, [plot, mismatchedXmlFallback, PptxTheme.Empty, PptxColorMap.Default])) ?? throw new InvalidOperationException("Expected scene point strokes.")).Cast<object>().Single();
        object xmlPointFills = (((System.Collections.IEnumerable?)readPointFills.Invoke(null, [null, mismatchedXmlFallback, PptxTheme.Empty, PptxColorMap.Default])) ?? throw new InvalidOperationException("Expected XML point fills.")).Cast<object>().Single();
        object xmlPointStrokes = (((System.Collections.IEnumerable?)readPointStrokes.Invoke(null, [null, mismatchedXmlFallback, PptxTheme.Empty, PptxColorMap.Default])) ?? throw new InvalidOperationException("Expected XML point strokes.")).Cast<object>().Single();

        object sceneFill = PptxTests.ChartDictionaryValue(scenePointFills, 1);
        object sceneStroke = PptxTests.ChartDictionaryValue(scenePointStrokes, 1);
        object xmlFill = PptxTests.ChartDictionaryValue(xmlPointFills, 0);
        object xmlStroke = PptxTests.ChartDictionaryValue(xmlPointStrokes, 0);

        TestAssert.Equal(new RgbColor(18, 52, 86), PptxTests.ChartSeriesFillColor(sceneFill));
        TestAssert.Equal(new RgbColor(101, 67, 33), PptxTests.ChartSeriesStrokeColor(sceneStroke));
        TestAssert.Equal(3d, PptxTests.ChartSeriesStrokeWidth(sceneStroke));
        TestAssert.Equal(new RgbColor(160, 176, 192), PptxTests.ChartSeriesFillColor(xmlFill));
        TestAssert.Equal(new RgbColor(12, 11, 10), PptxTests.ChartSeriesStrokeColor(xmlStroke));
        TestAssert.Equal(1d, PptxTests.ChartSeriesStrokeWidth(xmlStroke));
    }
}
