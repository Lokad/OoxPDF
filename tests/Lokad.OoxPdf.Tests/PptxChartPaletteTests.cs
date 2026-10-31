using System.Globalization;
using System.Text.Json;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxChartPaletteTests
{
    public static void PptxAutomaticBarPalettesMatchIndependentOfficeAcrossCountsAndThemes()
    {
        string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "chart-auto-palette-office.json"));
        using var reference = JsonDocument.Parse(File.ReadAllText(path));
        foreach (JsonElement row in reference.RootElement.GetProperty("Rows").EnumerateArray())
        {
            int count = row.GetProperty("PointCount").GetInt32();
            string id = row.GetProperty("Id").GetString()!;
            string direction = row.GetProperty("Direction").GetString()!;
            string palette = string.Join("", row.GetProperty("Palette").EnumerateArray().Select(color =>
                "<a:srgbClr val=\"" + color.GetString() + "\"/>"));
            string values = string.Join("", Enumerable.Range(0, count).Select(i =>
                "<c:pt idx=\"" + i + "\"><c:v>4</c:v></c:pt>"));
            string chart = $$"""
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
                  <c:chart><c:plotArea><c:barChart>
                    <c:barDir val="{{direction}}"/><c:grouping val="clustered"/><c:varyColors val="1"/>
                    <c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numLit>
                      <c:ptCount val="{{count}}"/>{{values}}
                    </c:numLit></c:val></c:ser>
                  </c:barChart></c:plotArea></c:chart>
                </c:chartSpace>
                """;
            var page = Render(chart, palette);
            RgbColor[] paints = ReadColoredPaints(page.Content).ToArray();
            string[] expected = row.GetProperty("Colors").EnumerateArray().Select(color => color.GetString()!).ToArray();
            TestAssert.Equal(count, paints.Length);
            var residuals = row.GetProperty("Residuals").EnumerateArray().ToDictionary(
                residual => residual.GetProperty("PointIndex").GetInt32(),
                residual => residual.GetProperty("MaxChannelError").GetInt32());
            for (int i = 0; i < count; i++)
            {
                byte[] bytes = Convert.FromHexString(expected[i]);
                int error = Math.Max(Math.Abs(paints[i].Red - bytes[0]), Math.Max(
                    Math.Abs(paints[i].Green - bytes[1]), Math.Abs(paints[i].Blue - bytes[2])));
                int allowed = residuals.GetValueOrDefault(i);
                TestAssert.True(error <= allowed, $"{id} point {i}: RGB error {error}, allowed {allowed}.");
            }
        }
    }

    public static void PptxVaryingBarCategoryLegendsSharePointPaintAndPreserveSeriesLegends()
    {
        const int count = 6;
        foreach (var (vary, seriesFill, pointFill, negative) in new[]
        {
            (true, false, false, false), (false, false, false, false),
            (true, true, false, false), (true, false, true, false),
            (true, false, false, true), (false, false, true, false)
        })
        {
            PdfPage page = RenderLegend(count, vary, seriesFill, pointFill, negative);
            RgbColor[] paints = ReadColoredPaints(page.Content).ToArray();
            int bars = negative ? count - 1 : count; // White inverted paint is excluded.
            bool categoryLegend = vary || pointFill;
            int keys = categoryLegend ? bars : 1;
            TestAssert.Equal(bars + keys, paints.Length);
            if (categoryLegend) TestAssert.True(paints.Take(bars).SequenceEqual(paints.Skip(bars)), "Category legend keys must share the resolved point paint.");
        }
    }

    public static void PptxVaryingBarCategoryLegendsWrapAndReservePlotSpace()
    {
        const int count = 13;
        PdfPage page = RenderLegend(count);
        var rectangles = new List<(double X, double Y, double Width, double Height)>();
        _ = ReadColoredPaints(page.Content, line =>
        {
            if (!line.EndsWith(" re f", StringComparison.Ordinal)) return;
            double[] values = line[..^5].Split(' ', StringSplitOptions.RemoveEmptyEntries)
                .Select(value => double.Parse(value, CultureInfo.InvariantCulture)).ToArray();
            rectangles.Add((values[0], values[1], values[2], values[3]));
        }).ToArray();
        TestAssert.Equal(2 * count, rectangles.Count);
        var bars = rectangles.Take(count).ToArray();
        var keys = rectangles.Skip(count).ToArray();
        TestAssert.True(keys.Select(key => key.Y).Distinct().Count() > 1, "The category legend must wrap into measured rows.");
        TestAssert.True(keys.Max(key => key.Y + key.Height) < bars.Min(bar => bar.Y), "Legend space must be reserved before placing plot bars.");
        TestAssert.True(keys.All(key => key.X >= 72 && key.X + key.Width <= 360 && key.Y >= page.Height - 288 && key.Y + key.Height <= page.Height - 72), "Legend keys must stay inside the source chart frame.");
    }

    private static PdfPage RenderLegend(int count, bool vary = true, bool seriesFill = false, bool pointFill = false, bool negative = false)
    {
        string values = string.Join("", Enumerable.Range(0, count).Select(i =>
            $"<c:pt idx=\"{i}\"><c:v>{(negative && i == 0 ? -4 : 4)}</c:v></c:pt>"));
        string labels = string.Join("", Enumerable.Range(0, count).Select(i =>
            $"<c:pt idx=\"{i}\"><c:v>Category {i + 1}</c:v></c:pt>"));
        const string fill = "<c:spPr><a:solidFill><a:srgbClr val=\"E52C40\"/></a:solidFill></c:spPr>";
        string point = pointFill ? "<c:dPt><c:idx val=\"0\"/>" + fill + "</c:dPt>" : "";
        string chart = $$"""
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:plotArea><c:barChart><c:barDir val="col"/>
                <c:grouping val="clustered"/><c:varyColors val="{{(vary ? 1 : 0)}}"/>
                <c:ser><c:idx val="0"/><c:order val="0"/>
                  {{(seriesFill ? fill : "")}}<c:invertIfNegative val="1"/>{{point}}
                  <c:cat><c:strLit><c:ptCount val="{{count}}"/>{{labels}}</c:strLit></c:cat>
                  <c:val><c:numLit><c:ptCount val="{{count}}"/>{{values}}</c:numLit></c:val>
                </c:ser>
              </c:barChart></c:plotArea><c:legend><c:legendPos val="b"/></c:legend></c:chart>
            </c:chartSpace>
            """;
        string palette = string.Join("", new[] { "156082", "E97132", "196B24", "0F9ED5", "A02B93", "4EA72E" }
            .Select(color => "<a:srgbClr val=\"" + color + "\"/>"));
        return Render(chart, palette);
    }

    private static PdfPage Render(string chart, string palette)
    {
        var parts = new Dictionary<string, byte[]>
        {
            ["ppt/charts/_rels/chart1.xml.rels"] = TestFixtures.Utf8("""
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="color" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors.xml"/>
                </Relationships>
                """),
            ["ppt/charts/colors.xml"] = TestFixtures.Utf8($$"""
                <cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"
                    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" meth="cycle">{{palette}}</cs:colorStyle>
                """)
        };
        var (document, package) = PptxTests.BuildSingleChartPackage(chart, parts);
        var page = new PptxRenderer(new PortableRenderingTests.PinnedResolver())
            .RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
        return page;
    }


    // Inspect painted colors rather than private helpers, branch names or constants.
    internal static IEnumerable<RgbColor> ReadColoredPaints(string content, Action<string>? painted = null, bool strokes = false)
    {
        RgbColor? color = null;
        var saved = new Stack<RgbColor?>();
        foreach (string raw in content.Split('\n'))
        {
            string line = raw.Trim();
            if (line == "q") { saved.Push(color); }
            else if (line == "Q") { color = saved.Pop(); }
            else if (line.EndsWith(strokes ? " RG" : " rg", StringComparison.Ordinal))
            {
                double[] channels = line[..^3].Split(' ', StringSplitOptions.RemoveEmptyEntries)
                    .Select(value => double.Parse(value, CultureInfo.InvariantCulture)).ToArray();
                color = new RgbColor(Channel(channels[0]), Channel(channels[1]), Channel(channels[2]));
            }
            else if (line.EndsWith(strokes ? " G" : " g", StringComparison.Ordinal)) { color = null; }
            else if (color is { } paint && (paint.Red != paint.Green || paint.Green != paint.Blue) &&
                (strokes ? line is "S" or "s" || line.EndsWith(" S", StringComparison.Ordinal) || line.EndsWith(" s", StringComparison.Ordinal) :
                    line is "f" or "f*" || line.EndsWith(" f", StringComparison.Ordinal) || line.EndsWith(" f*", StringComparison.Ordinal)))
            { painted?.Invoke(line); yield return paint; }
        }
        static byte Channel(double value) => (byte)Math.Round(value * 255, MidpointRounding.AwayFromZero);
    }
}
