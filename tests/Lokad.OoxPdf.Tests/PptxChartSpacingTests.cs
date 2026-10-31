using System.Globalization;
using System.Text.Json;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxChartSpacingTests
{
    public static void PptxClusteredSpacingMatchesIndependentOfficeAcrossSeriesAndOverlap()
    {
        string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "ClusteredBarSpacing", "office.json"));
        using var reference = JsonDocument.Parse(File.ReadAllText(path));
        foreach (var row in reference.RootElement.GetProperty("Rows").EnumerateArray())
        {
            int seriesCount = row.GetProperty("SeriesCount").GetInt32();
            string direction = row.GetProperty("Direction").GetString()!;
            int gap = row.GetProperty("GapWidth").GetInt32();
            int overlap = row.GetProperty("Overlap").GetInt32();
            string series = string.Concat(Enumerable.Range(0, seriesCount).Select(i => $$"""
                <c:ser><c:idx val="{{i}}"/><c:order val="{{i}}"/>
                  <c:val><c:numLit><c:ptCount val="6"/>
                    {{string.Concat(Enumerable.Range(0, 6).Select(j => $"<c:pt idx=\"{j}\"><c:v>4</c:v></c:pt>"))}}
                  </c:numLit></c:val>
                </c:ser>
                """));
            string chart = $$"""
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
                  <c:chart><c:plotArea><c:barChart><c:barDir val="{{direction}}"/>
                    <c:grouping val="clustered"/><c:varyColors val="0"/>
                    {{series}}<c:gapWidth val="{{gap}}"/><c:overlap val="{{overlap}}"/>
                  </c:barChart></c:plotArea></c:chart>
                </c:chartSpace>
                """;
            using var stream = PptxTests.CreateSingleChartPackageStream(chart);
            using var package = OoxPackage.Open(stream, default);
            var document = new PptxReader().Read(package, default);
            PdfPage page = new PptxRenderer(new PortableRenderingTests.PinnedResolver())
                .RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
            bool horizontal = direction == "bar";
            var bars = ReadBarRectangles(page.Content).OrderBy(rect => horizontal ? rect.Y : rect.X).ToArray();
            TestAssert.Equal(6 * seriesCount, bars.Length);
            double band = ((horizontal ? bars[5 * seriesCount].Y : bars[5 * seriesCount].X) -
                (horizontal ? bars[0].Y : bars[0].X)) / 5;
            double thickness = bars.Average(rect => horizontal ? rect.Height : rect.Width);
            double officeBand = row.GetProperty("CategoryBand").GetDouble();
            double expected = band * row.GetProperty("BarThickness").GetDouble() / officeBand;
            // Isolate interseries spacing from the still-approximate automatic
            // plot frame. Expected proportions come from independent Office PDFs.
            TestAssert.True(Math.Abs(thickness - expected) < 0.01,
                $"{row.GetProperty("Id").GetString()}: bar thickness {thickness}, Office proportion {expected}.");
            if (seriesCount > 1 && row.TryGetProperty("SeriesStep", out var officeStep))
            {
                double step = Enumerable.Range(0, 6).Average(category =>
                    ((horizontal ? bars[category * seriesCount + seriesCount - 1].Y : bars[category * seriesCount + seriesCount - 1].X) -
                     (horizontal ? bars[category * seriesCount].Y : bars[category * seriesCount].X)) / (seriesCount - 1));
                TestAssert.True(Math.Abs(step - band * officeStep.GetDouble() / officeBand) < 0.01,
                    $"{row.GetProperty("Id").GetString()}: series step must retain Office spacing.");
            }
        }
    }

    internal static IEnumerable<(double X, double Y, double Width, double Height)> ReadBarRectangles(string content)
    {
        bool nonWhite = false;
        var saved = new Stack<bool>();
        foreach (string raw in content.Split('\n'))
        {
            string line = raw.Trim();
            if (line == "q") saved.Push(nonWhite);
            else if (line == "Q") nonWhite = saved.Pop();
            else if (line.EndsWith(" rg", StringComparison.Ordinal))
                nonWhite = line[..^3].Split(' ', StringSplitOptions.RemoveEmptyEntries).Select(Number).Any(channel => channel < 1d);
            else if (line.EndsWith(" g", StringComparison.Ordinal)) nonWhite = Number(line[..^2]) < 1d;
            else if (nonWhite && line.EndsWith(" re f", StringComparison.Ordinal))
            {
                double[] values = line[..^5].Split(' ', StringSplitOptions.RemoveEmptyEntries).Select(Number).ToArray();
                yield return (values[0], values[1], values[2], values[3]);
            }
        }
        static double Number(string value) => double.Parse(value, CultureInfo.InvariantCulture);
    }
}
