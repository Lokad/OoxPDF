using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxChartNegativeInversionTests
{
    private static readonly RgbColor SeriesFill = new(51, 102, 204);
    private static readonly RgbColor InvertedFill = new(255, 255, 255);

    public static void PptxNegativeBarScenePreservesSeriesAndPointInversion()
    {
        foreach ((string declaration, bool expected) in new[]
        {
            ("", true), ("<c:invertIfNegative/>", true),
            ("<c:invertIfNegative val=\"1\"/>", true),
            ("<c:invertIfNegative val=\"true\"/>", true),
            ("<c:invertIfNegative val=\"0\"/>", false),
            ("<c:invertIfNegative val=\"false\"/>", false)
        })
        {
            PptxSceneChart chart = TestAssert.NotNull(PptxTests.BuildSingleChartScene(Chart(declaration)));
            var policy = chart.Plots[0].Series[0].NegativeInversion;
            TestAssert.Equal(expected, policy.ForPoint(0));
            TestAssert.Equal(expected, policy.ForPoint(7));
        }
        string overrides = """
            <c:dPt><c:idx val="0"/><c:invertIfNegative val="0"/></c:dPt>
            <c:dPt><c:idx val="1"/><c:invertIfNegative val="0"/></c:dPt>
            <c:dPt><c:idx val="1"/><c:invertIfNegative/></c:dPt>
            <c:dPt><c:idx val="-1"/><c:invertIfNegative val="0"/></c:dPt>
            """;
        var series = TestAssert.NotNull(PptxTests.BuildSingleChartScene(Chart("<c:invertIfNegative val=\"1\"/>", overrides))).Plots[0].Series[0];
        TestAssert.Equal(false, series.NegativeInversion.ForPoint(0));
        TestAssert.Equal(true, series.NegativeInversion.ForPoint(1));
        TestAssert.Equal(true, series.NegativeInversion.ForPoint(2));
        TestAssert.Equal(2, TestAssert.NotNull(series.NegativeInversion.PointValues).Count);
        var pointDefault = TestAssert.NotNull(PptxTests.BuildSingleChartScene(Chart(
            "<c:invertIfNegative val=\"0\"/>", "<c:dPt><c:idx val=\"0\"/></c:dPt>"))).Plots[0].Series[0];
        TestAssert.Equal(true, pointDefault.NegativeInversion.ForPoint(0));
        TestAssert.Equal(false, pointDefault.NegativeInversion.ForPoint(1));
    }

    public static void PptxNegativeBarPaintHonorsInversionAcrossLayouts()
    {
        foreach (string direction in new[] { "col", "bar" })
        foreach (string grouping in new[] { "clustered", "stacked" })
        foreach ((string declaration, bool blue) in new[]
        {
            ("<c:invertIfNegative val=\"0\"/>", true),
            ("<c:invertIfNegative val=\"1\"/>", false),
            ("", false)
        })
        {
            PdfPage page = Render(Chart(declaration, direction: direction, grouping: grouping));
            // Stacked bars can use closed paths. This chart paints only the slide background and one bar.
            var fills = PptxChartPaletteTests.ReadColoredPaints(page.Content, includeNeutral: true).ToArray();
            TestAssert.Equal(2, fills.Length);
            TestAssert.Equal(InvertedFill, fills[0]);
            TestAssert.Equal(blue ? SeriesFill : InvertedFill, fills[1]);
        }
    }

    public static void PptxNegativeBarPointOverridesAlsoControlLabelSwatches()
    {
        foreach ((bool series, bool point) in new[] { (true, false), (false, true), (false, false), (true, true) })
        {
            string declaration = "<c:invertIfNegative val=\"" + (series ? "1" : "0") + "\"/>";
            string pointStyle = "<c:dPt><c:idx val=\"0\"/><c:invertIfNegative val=\"" + (point ? "1" : "0") + "\"/></c:dPt>";
            PdfPage page = Render(Chart(declaration, pointStyle, labels: true));
            var expected = point ? InvertedFill : SeriesFill;
            var (barFill, swatchFill) = CheckBarAndSwatch(page, expected);

            if (point && !series)
            {
                var wrongBar = page with
                {
                    Content = page.Content.Replace(barFill, "1 0 0 rg\n" + barFill, StringComparison.Ordinal)
                };
                var missingSwatch = page with
                {
                    Content = page.Content.Replace(swatchFill, "", StringComparison.Ordinal)
                };
                TestAssert.Equal("Unexpected bar fill.", TestAssert.Throws<InvalidOperationException>(
                    () => CheckBarAndSwatch(wrongBar, expected)).Message);
                TestAssert.Equal("Expected both the bar and the data-label swatch.", TestAssert.Throws<InvalidOperationException>(
                    () => CheckBarAndSwatch(missingSwatch, expected)).Message);
            }
        }
    }

    private static (string Bar, string Swatch) CheckBarAndSwatch(PdfPage page, RgbColor expected)
    {
        var fills = ChartFills(page).ToArray();
        TestAssert.True(fills.Length == 2, "Expected both the bar and the data-label swatch.");
        var bar = fills[0];
        var swatch = fills[1];
        TestAssert.True(bar.Color == expected, "Unexpected bar fill.");
        TestAssert.True(swatch.Color == expected, "Unexpected data-label swatch fill.");
        TestAssert.True(bar.Width > swatch.Width && bar.Height > swatch.Height,
            "The data-label swatch must be smaller than its bar.");

        var label = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
            fontResourceFilter: name => name.StartsWith("CBD", StringComparison.Ordinal)).ToArray();
        TestAssert.Equal("Series 1, Negative, -2", string.Concat(label.Select(show => show.Text)));
        double baseline = label[0].Y;
        TestAssert.True(swatch.X + swatch.Width < label[0].X &&
            Math.Abs(swatch.Y + swatch.Height / 2d - baseline) < label[0].Size,
            "The swatch must sit to the left of its complete data label at the same text line.");
        return (bar.Operation, swatch.Operation);
    }

    private static IEnumerable<(RgbColor Color, double X, double Y, double Width, double Height, string Operation)> ChartFills(PdfPage page) =>
        // BuildSingleChartPackage places the 288 by 216 point chart at (72, 72) from the slide top.
        PptxComposedPaletteTests.ReadFilledRectangles(page).Where(rect =>
            rect.X >= 72d && rect.X + rect.Width <= 360d &&
            rect.Y >= page.Height - 288d && rect.Y + rect.Height <= page.Height - 72d);

    public static void PptxNegativeBarPointDefaultsAndExplicitPaintMatchOffice()
    {
        const string fill = "<c:spPr><a:solidFill><a:srgbClr val=\"3366CC\"/></a:solidFill></c:spPr>";
        foreach ((string flag, string paint, bool negative, bool blue) in new[]
        {
            ("", "", true, false),
            ("", fill, true, false),
            ("<c:invertIfNegative val=\"0\"/>", fill, true, true),
            ("<c:invertIfNegative val=\"1\"/>", fill, true, false),
            ("", fill, false, true)
        })
        {
            string point = "<c:dPt><c:idx val=\"0\"/>" + flag + paint + "</c:dPt>";
            PdfPage page = Render(Chart("<c:invertIfNegative val=\"0\"/>", point, negative: negative));
            TestAssert.Equal(blue ? SeriesFill : InvertedFill, ChartFills(page).Single().Color);
        }
    }

    private static string Chart(string inversion, string points = "", string direction = "col", string grouping = "clustered", bool labels = false, bool negative = true) => $$"""
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <c:chart><c:plotArea><c:barChart>
            <c:barDir val="{{direction}}"/><c:grouping val="{{grouping}}"/><c:varyColors val="0"/>
            <c:ser><c:idx val="0"/><c:order val="0"/>
              <c:spPr><a:solidFill><a:srgbClr val="3366CC"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>
              {{inversion}}{{points}}
              <c:cat><c:strLit><c:ptCount val="1"/><c:pt idx="0"><c:v>Negative</c:v></c:pt></c:strLit></c:cat>
              <c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>{{(negative ? "-2" : "2")}}</c:v></c:pt></c:numLit></c:val>
              {{(labels ? "<c:dLbls><c:dLblPos val=\"outEnd\"/><c:showLegendKey val=\"1\"/><c:showVal val=\"1\"/></c:dLbls>" : "")}}
            </c:ser>
          </c:barChart></c:plotArea><c:legend><c:delete val="1"/></c:legend></c:chart>
        </c:chartSpace>
        """;

    private static PdfPage Render(string chart)
    {
        var (document, package) = PptxTests.BuildSingleChartPackage(chart);
        return new PptxRenderer(new PortableRenderingTests.PinnedResolver())
            .RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
    }
}
