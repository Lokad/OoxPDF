using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxLinePointStyleTests
{
    public static void PptxColorSaturationPreservesGrayAndTransformOrder()
    {
        RgbColor Transform(string source, string transforms)
        {
            XElement color = XElement.Parse($"<a:srgbClr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" val=\"{source}\">{transforms}</a:srgbClr>");
            TestAssert.True(RgbColor.TryParse(source, out var initial), "Expected RGB fixture.");
            return PptxColorResolver.ApplyColorTransforms(color, initial);
        }
        TestAssert.Equal(new RgbColor(128, 128, 128), Transform("808080", "<a:satMod val=\"500000\"/>"));
        TestAssert.Equal(new RgbColor(127, 127, 127), Transform("FE0000", "<a:satMod val=\"0\"/>"));
        TestAssert.Equal(new RgbColor(255, 65, 65), Transform("C08080", "<a:satMod val=\"500000\"/>"));
        TestAssert.Equal(Transform("335577", ""), Transform("335577", "<a:satMod val=\"100000\"/>"));
        RgbColor shadeFirst = Transform("FE0000", "<a:shade val=\"50000\"/><a:satMod val=\"0\"/>");
        RgbColor saturationFirst = Transform("FE0000", "<a:satMod val=\"0\"/><a:shade val=\"50000\"/>");
        TestAssert.True(shadeFirst.Red == shadeFirst.Green && shadeFirst.Green == shadeFirst.Blue &&
            saturationFirst.Red == saturationFirst.Green && saturationFirst.Green == saturationFirst.Blue && shadeFirst != saturationFirst,
            "Preserve ordered transformations and HSL lightness when saturation changes.");
    }

    public static void PptxLinePointPaintAndLegendNamesMatchIndependentOffice()
    {
        string directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "LinePointStyles"));
        using var reference = JsonDocument.Parse(File.ReadAllText(Path.Combine(directory, "office.json")));
        foreach (var row in reference.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            using var stream = File.OpenRead(Path.Combine(directory, row.GetProperty("Input").GetString()!));
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(stream)));
            stream.Position = 0;
            using var package = OoxPackage.Open(stream, default);
            PdfPage page = Render(package);
            string[] names = LegendNames(page);
            TestAssert.True(names.SequenceEqual(row.GetProperty("LegendNames").EnumerateArray().Select(v => v.GetString()!)), id + ": preserve Office legend-role names and source order.");
            var segments = NativeSegments(page.Content).ToArray();
            var expected = row.GetProperty("Segments").EnumerateArray().ToArray();
            TestAssert.Equal(expected.Length, segments.Length);
            for (int index = 0; index < segments.Length; index++)
            {
                byte[] color = Convert.FromHexString(expected[index].GetProperty("Color").GetString()!);
                int error = Math.Max(Math.Abs(color[0] - segments[index].Color.Red), Math.Max(Math.Abs(color[1] - segments[index].Color.Green), Math.Abs(color[2] - segments[index].Color.Blue)));
                TestAssert.True(error <= reference.RootElement.GetProperty("MaxRgbChannelError").GetInt32(), $"{id} segment{index}: RGB error{error}.");
                TestAssert.True(Math.Abs(segments[index].Width - expected[index].GetProperty("Width").GetDouble()) <= .001d, id + ": preserve segment stroke width.");
                TestAssert.Equal(expected[index].GetProperty("Operator").GetString(), segments[index].Operator);
            }
        }
    }

    public static void PptxLinePointLegendsSharePathAndDataLabelPaint()
    {
        foreach (int count in new[] { 2, 4, 8, 19 })
        foreach (bool keys in new[] { false, true })
        {
            PdfPage explicitOn = Probe(count, "1", keys), omitted = Probe(count, null, keys), off = Probe(count, "0", keys);
            TestAssert.Equal(explicitOn.Content, omitted.Content);
            TestAssert.Equal(count, LegendNames(explicitOn).Length);
            TestAssert.True(LegendNames(explicitOn).SequenceEqual(Enumerable.Range(0, count).Select(i => "P" + i)), "Keep every category key in source order.");
            TestAssert.True(LegendNames(off).SequenceEqual(new[] { "Series" }), "Explicit false keeps a series legend.");
            RgbColor[] paints = PptxChartPaletteTests.ReadColoredPaints(explicitOn.Content, strokes: true).ToArray();
            TestAssert.Equal((keys ? 3 : 2) * count - 1, paints.Length);
            var legend = paints.Skip(count - 1).Take(count).ToArray();
            TestAssert.True(paints.Take(count - 1).SequenceEqual(legend.Skip(1)), "Each incoming segment shares its destination point's legend paint.");
            if (keys) TestAssert.True(legend.SequenceEqual(paints.Skip(2 * count - 1)), "Data-label keys share point paint.");
        }
    }

    public static void PptxLineAutomaticMarkersRetainSeriesNoneAndRawPlotSwitch()
    {
        foreach (string? flag in new string?[] { null, "0", "1" })
        {
            var chart = PptxTests.BuildSingleChartScene($$"""
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:lineChart>
                  {{(flag is null ? "" : "<c:marker val=\"" + flag + "\"/>")}}
                  <c:ser/><c:ser><c:marker><c:symbol val="none"/></c:marker></c:ser>
                </c:lineChart></c:plotArea></c:chart></c:chartSpace>
                """)!;
            var plot = chart.Plots.Single();
            TestAssert.Equal(flag ?? "", plot.MarkersEnabledValue);
            TestAssert.Equal(PptxSceneChartMarkerSymbol.Diamond, plot.Series[0].Marker.SymbolKind);
            TestAssert.Equal(7d, plot.Series[0].Marker.Size);
            TestAssert.True(!plot.Series[0].Marker.IsDefined && plot.Series[1].Marker.IsDefined, "Keep source marker provenance.");
            TestAssert.Equal(PptxSceneChartMarkerSymbol.None, plot.Series[1].Marker.SymbolKind);
        }
    }

    private static PdfPage Probe(int count, string? vary, bool keys)
    {
        string labels = string.Concat(Enumerable.Range(0, count).Select(i => $"<c:pt idx=\"{i}\"><c:v>P{i}</c:v></c:pt>"));
        string values = string.Concat(Enumerable.Range(0, count).Select(i => $"<c:pt idx=\"{i}\"><c:v>4</c:v></c:pt>"));
        string chart = $$"""
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:lineChart>
              {{(vary is null ? "" : "<c:varyColors val=\"" + vary + "\"/>")}}
              <c:ser><c:tx><c:v>Series</c:v></c:tx><c:marker><c:symbol val="none"/></c:marker>
                <c:cat><c:strLit><c:ptCount val="{{count}}"/>{{labels}}</c:strLit></c:cat>
                <c:val><c:numLit><c:ptCount val="{{count}}"/>{{values}}</c:numLit></c:val></c:ser>
              {{(keys ? "<c:dLbls><c:showLegendKey val=\"1\"/><c:showVal val=\"0\"/></c:dLbls>" : "")}}
              <c:axId val="1"/><c:axId val="2"/>
            </c:lineChart><c:catAx><c:axId val="1"/><c:axPos val="b"/><c:tickLblPos val="none"/></c:catAx>
              <c:valAx><c:axId val="2"/><c:scaling><c:min val="0"/><c:max val="5"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/></c:valAx>
            </c:plotArea><c:legend><c:legendPos val="b"/></c:legend></c:chart></c:chartSpace>
            """;
        var parts = new Dictionary<string, byte[]>
        {
            ["ppt/charts/_rels/chart1.xml.rels"] = TestFixtures.Utf8("""
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="colors" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors.xml"/></Relationships>
                """),
            ["ppt/charts/colors.xml"] = TestFixtures.Utf8("""
                <cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" meth="cycle"><a:srgbClr val="156082"/><a:srgbClr val="E97132"/><a:srgbClr val="196B24"/><a:srgbClr val="0F9ED5"/><a:srgbClr val="A02B93"/><a:srgbClr val="4EA72E"/></cs:colorStyle>
                """)
        };
        using var stream = PptxTests.CreateSingleChartPackageStream(chart, parts);
        using var package = OoxPackage.Open(stream, default);
        return Render(package);
    }

    private static PdfPage Render(OoxPackage package) => new PptxRenderer(new PortableRenderingTests.PinnedResolver())
        .RenderPagesAsync(new PptxReader().Read(package, default), package, null, default).ToBlockingEnumerable().Single();

    private static string[] LegendNames(PdfPage page)
    {
        return DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
            fontResourceFilter: name => name.StartsWith("CL", StringComparison.Ordinal) && name.Length > 2 && char.IsAsciiDigit(name[2]))
            .Select(show => show.Text).ToArray();
    }

    // These source fixtures have eight slots in a720pt frame. Plot segments
    // span over40pt; marker and legend-key paths are smaller. Keep curves and
    // straight edges separate without using any production style helper.
    private static IEnumerable<(RgbColor Color, double Width, string Operator)> NativeSegments(string content)
    {
        RgbColor? color = null; double width = 1d, x = 0d;
        var states = new Stack<(RgbColor?, double)>();
        var edges = new List<string>();
        foreach (string raw in content.Split('\n'))
        {
            string line = raw.Trim(); string[] words = line.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            if (words.Length == 0) continue;
            string op = words[^1];
            double N(int index) => double.Parse(words[index], CultureInfo.InvariantCulture);
            if (line == "q") states.Push((color, width));
            else if (line == "Q") (color, width) = states.Pop();
            else if (op == "RG") color = new RgbColor((byte)Math.Round(N(0) * 255), (byte)Math.Round(N(1) * 255), (byte)Math.Round(N(2) * 255));
            else if (op == "G") color = null;
            else if (op == "w") width = N(0);
            else if (op == "m") x = N(0);
            else if (op is "l" or "c")
            {
                double next = N(op == "c" ? 4 : 0);
                if (Math.Abs(next - x) > 40d) edges.Add(op);
                x = next;
            }
            else if (words.Length == 7 && words[2] == "m" && words[5] == "l" && op == "S")
            {
                if (Math.Abs(N(3) - N(0)) > 40d && color is { } ink && (ink.Red != ink.Green || ink.Green != ink.Blue))
                    yield return (ink, width, "l");
                edges.Clear();
            }
            else if (op is "S" or "s")
            {
                if (color is { } ink && (ink.Red != ink.Green || ink.Green != ink.Blue))
                    foreach (string edge in edges) yield return (ink, width, edge);
                edges.Clear();
            }
            else if (op is "n" or "f" or "f*") edges.Clear();
        }
    }
}
