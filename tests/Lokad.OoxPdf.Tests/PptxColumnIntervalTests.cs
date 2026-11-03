using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxColumnIntervalTests
{
    public static void PptxColumnIntervalsShareOfficeSourceMinimums()
    {
        // Independently exported Office controls include negative minima, auto
        // units and a unit wider than the range; no extra endpoint is a tick.
        (int Min, int? Major, string Expected)[] cases =
        [
            (-7, 20, "-7,13,33,53,73,93"), (-7, null, "-7,13,33,53,73,93"),
            (7, 25, "7,32,57,82"), (7, null, "7,17,27,37,47,57,67,77,87,97"),
            (0, null, "0,20,40,60,80,100"), (7, 500, "7")
        ];
        foreach (var test in cases)
        foreach (bool reversed in new[] { false, true })
        {
            var labels = Labels(Render(test.Min, test.Major, reversed: reversed));
            TestAssert.True(string.Join(",", labels) == test.Expected,
                $"Source minimum {test.Min}, unit {test.Major}, reversed {reversed}: {string.Join(",", labels)}");
        }
    }

    public static void PptxColumnIntervalsSurviveStackingAndAuthoredPlotPlacement()
    {
        foreach (string grouping in new[] { "clustered", "stacked" })
        foreach (bool manual in new[] { false, true })
            TestAssert.True(Labels(Render(7, 25, grouping: grouping, manual: manual)).SequenceEqual(new[] { "7", "32", "57", "82" }),
                "Changing placement or stacking keeps the supported numeric source intervals.");
        TestAssert.True(Labels(Render(null, 25)).SequenceEqual(new[] { "-22", "3", "28", "53", "78", "103" }),
            "An authored upper bound anchors whole intervals down to the automatic lower bound, as in the native control.");
    }

    public static void PptxColumnTicksAndBothGridKindsUseTheLabelIntervals()
    {
        foreach (bool reversed in new[] { false, true })
        {
            PdfPage page = Render(7, 25, reversed: reversed, grids: true);
            var lines = Lines(page);
            double[] axis = lines.Single(p => p.Color == new RgbColor(255, 0, 0) && Math.Abs(p.V[0] - p.V[2]) < .001d && Math.Abs(p.V[1] - p.V[3]) > 100d).V;
            double lower = Math.Min(axis[1], axis[3]), span = Math.Abs(axis[3] - axis[1]);
            double Value(double y) => 7d + (reversed ? 1d - (y - lower) / span : (y - lower) / span) * 96d;
            double[] Levels(RgbColor color, bool ticks) => lines
                .Where(p => p.Color == color && Math.Abs(p.V[1] - p.V[3]) < .001d &&
                    (ticks ? Math.Abs(p.V[0] - p.V[2]) is > 0d and < 20d : Math.Abs(p.V[0] - p.V[2]) > 100d))
                .Select(p => Value(p.V[1])).Order().ToArray();
            void Check(double[] actual, double[] expected, string description) => TestAssert.True(actual.Length == expected.Length &&
                actual.Zip(expected).All(p => Math.Abs(p.First - p.Second) < .002d),
                description + ": " + string.Join(",", actual.Select(v => v.ToString("F3", CultureInfo.InvariantCulture))));
            Check(Levels(new RgbColor(255, 0, 0), true), [7d, 32d, 57d, 82d], "Numeric tick marks share the source lattice");
            Check(Levels(new RgbColor(0, 0, 255), false), [32d, 57d, 82d], "Major grids exclude the category crossing");
            Check(Levels(new RgbColor(255, 0, 255), false), Enumerable.Range(1, 19).Where(i => i % 5 != 0).Select(i => 7d + 5d * i).ToArray(),
                "Minor grids subdivide the same major lattice and exclude its painted levels");
        }
    }

    public static void PptxColumnIntervalMeasurementRetainsEndpointAdvance()
    {
        PdfPage shortBound = Render(7, 25, maximum: 99), longBound = Render(7, 25, maximum: 103);
        TestAssert.True(Labels(shortBound).SequenceEqual(Labels(longBound)), "Neither non-lattice maximum becomes a painted tick.");
        double AxisX(PdfPage page) => Lines(page).Single(p => p.Color == new RgbColor(255, 0, 0) && Math.Abs(p.V[0] - p.V[2]) < .001d && Math.Abs(p.V[1] - p.V[3]) > 100d).V[0];
        TestAssert.True(AxisX(longBound) > AxisX(shortBound) + 1d,
            "The three-digit source maximum still reserves its measured advance, independently of emitted ticks.");
    }

    private static string[] Labels(PdfPage page) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
        fontResourceFilter: name => name.StartsWith("CVA", StringComparison.Ordinal)).Select(l => l.Text).ToArray();

    internal static (RgbColor Color, double[] V)[] Lines(PdfPage page)
    {
        var result = new List<(RgbColor, double[])>();
        var saved = new Stack<RgbColor>();
        RgbColor color = default;
        string path = "";
        foreach (string line in page.Content.Split('\n').Select(l => l.Trim()))
        {
            if (line == "q") saved.Push(color);
            else if (line == "Q") color = saved.Pop();
            else if (line.EndsWith(" RG", StringComparison.Ordinal))
            {
                byte[] rgb = line[..^3].Split(' ').Select(s => (byte)Math.Round(double.Parse(s, CultureInfo.InvariantCulture) * 255d)).ToArray();
                color = new RgbColor(rgb[0], rgb[1], rgb[2]);
            }
            else if (line.EndsWith(" G", StringComparison.Ordinal)) color = default;
            else if (line.EndsWith(" m", StringComparison.Ordinal) || line.EndsWith(" l", StringComparison.Ordinal) || line.EndsWith(" S", StringComparison.Ordinal) || line == "S")
            {
                path += line + " ";
                if (line == "S" || line.EndsWith(" S", StringComparison.Ordinal))
                {
                    result.AddRange(Regex.Matches(path, @"([-\d.]+) ([-\d.]+) m\s+([-\d.]+) ([-\d.]+) l")
                        .Select(m => (color, m.Groups.Cast<Group>().Skip(1).Select(g => double.Parse(g.Value, CultureInfo.InvariantCulture)).ToArray())));
                    path = "";
                }
            }
        }
        return result.ToArray();
    }

    private static PdfPage Render(int? minimum, int? major, int maximum = 103, bool reversed = false, string grouping = "clustered", bool manual = false, bool grids = false, bool horizontal = false, Action<XElement>? editChart = null)
    {
        string Grid(string name, string color) => grids ? $"<c:{name}><c:spPr><a:ln><a:solidFill><a:srgbClr val=\"{color}\"/></a:solidFill></a:ln></c:spPr></c:{name}>" : "";
        string chart = $$"""
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:autoTitleDeleted val="1"/><c:plotArea>
              {{(manual ? "<c:layout><c:manualLayout><c:layoutTarget val=\"inner\"/><c:xMode val=\"factor\"/><c:yMode val=\"factor\"/><c:wMode val=\"factor\"/><c:hMode val=\"factor\"/><c:x val=\".12\"/><c:y val=\".1\"/><c:w val=\".78\"/><c:h val=\".72\"/></c:manualLayout></c:layout>" : "")}}
              <c:barChart><c:barDir val="{{(horizontal ? "bar" : "col")}}"/><c:grouping val="{{grouping}}"/><c:ser><c:idx val="0"/><c:order val="0"/>
                <c:cat><c:strLit><c:pt idx="0"><c:v>First</c:v></c:pt><c:pt idx="1"><c:v>Last</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>82</c:v></c:pt><c:pt idx="1"><c:v>96</c:v></c:pt></c:numLit></c:val>
              </c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart>
              <c:catAx><c:axId val="1"/><c:axPos val="{{(horizontal ? "l" : "b")}}"/><c:crossAx val="2"/><c:crosses val="autoZero"/><c:majorTickMark val="out"/>
                <c:spPr><a:ln><a:solidFill><a:srgbClr val="00AA00"/></a:solidFill></a:ln></c:spPr>
              </c:catAx>
              <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="{{(reversed ? "maxMin" : "minMax")}}"/>
                {{(minimum is { } min ? $"<c:min val=\"{min}\"/>" : "")}}<c:max val="{{maximum}}"/></c:scaling>
                <c:axPos val="{{(horizontal ? "b" : "l")}}"/><c:crossAx val="1"/><c:crosses val="autoZero"/><c:majorTickMark val="out"/>
                {{(major is { } unit ? $"<c:majorUnit val=\"{unit}\"/>" : "")}}{{(grids ? "<c:minorUnit val=\"5\"/>" : "")}}
                {{Grid("majorGridlines", "0000FF")}}{{Grid("minorGridlines", "FF00FF")}}
                <c:spPr><a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></c:spPr>
              </c:valAx>
            </c:plotArea></c:chart></c:chartSpace>
            """;
        const string slide = """
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><p:cSld><p:spTree><p:graphicFrame><p:xfrm><a:off x="914400" y="457200"/><a:ext cx="9144000" cy="5486400"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>
            """;
        if (editChart is not null)
        {
            XElement source = XElement.Parse(chart);
            editChart(source);
            chart = source.ToString(SaveOptions.DisableFormatting);
        }
        using var stream = PptxTests.CreateSingleChartPackageStream(chart, new Dictionary<string, byte[]> { ["ppt/slides/slide1.xml"] = TestFixtures.Utf8(slide) });
        using var package = OoxPackage.Open(stream, default);
        return new PptxRenderer(new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(new PptxReader().Read(package, default), package, null, default).ToBlockingEnumerable().Single();
    }
}
