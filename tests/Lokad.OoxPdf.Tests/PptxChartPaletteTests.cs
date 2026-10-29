using System.Globalization;
using System.Text.Json;
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
            using var ownedPackage = package;
            var page = new PptxRenderer(new PortableRenderingTests.PinnedResolver())
                .RenderPages(document, package, null, default).Single();
            RgbColor[] paints = ReadColoredFills(page.Content).ToArray();
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

    // Inspect painted colors rather than private helpers, branch names or constants.
    private static IEnumerable<RgbColor> ReadColoredFills(string content)
    {
        RgbColor? color = null;
        var saved = new Stack<RgbColor?>();
        foreach (string raw in content.Split('\n'))
        {
            string line = raw.Trim();
            if (line == "q") { saved.Push(color); }
            else if (line == "Q") { color = saved.Pop(); }
            else if (line.EndsWith(" rg", StringComparison.Ordinal))
            {
                double[] channels = line[..^3].Split(' ', StringSplitOptions.RemoveEmptyEntries)
                    .Select(value => double.Parse(value, CultureInfo.InvariantCulture)).ToArray();
                color = new RgbColor(Channel(channels[0]), Channel(channels[1]), Channel(channels[2]));
            }
            else if (line.EndsWith(" g", StringComparison.Ordinal)) { color = null; }
            else if (color is { } paint && (paint.Red != paint.Green || paint.Green != paint.Blue) &&
                (line is "f" or "f*" || line.EndsWith(" f", StringComparison.Ordinal) || line.EndsWith(" f*", StringComparison.Ordinal)))
            { yield return paint; }
        }
        static byte Channel(double value) => (byte)Math.Round(value * 255, MidpointRounding.AwayFromZero);
    }
}
