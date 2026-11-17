using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxAuthoredColorTests
{
    public static void PptxScRgbChannelsMatchStandardBytesAndClampEndpoints()
    {
        // DrawingML's standard example: 50% linear gray is sRGB BCBCBC.
        // Literal byte controls also bracket the linear/encoded transfer boundary.
        foreach (var (channel, expected) in new[] {
            (-1000L, (byte)0), (0L, (byte)0), (313L, (byte)10), (400L, (byte)13),
            (10000L, (byte)89), (25000L, (byte)137), (50000L, (byte)188),
            (75000L, (byte)225), (100000L, (byte)255), (120000L, (byte)255),
            (long.MinValue, (byte)0), (long.MaxValue, (byte)255) })
        {
            string value = channel.ToString(System.Globalization.CultureInfo.InvariantCulture);
            var actual = ReadScRgbColor($"<a:scrgbClr r=\"{value}\" g=\"{value}\" b=\"{value}\"/>");
            TestAssert.Equal(new RgbColor(expected, expected, expected), actual.Color);
            TestAssert.Equal(1d, actual.Alpha);
        }
        TestAssert.Equal(new RgbColor(188, 137, 0), ReadScRgbColor("<a:scrgbClr r=\"50000\" g=\"25000\" b=\"0\"/>").Color);
        TestAssert.Equal(new RgbColor(0, 0, 255), ReadScRgbColor("<a:scrgbClr r=\"bad\" b=\"100000\"/>").Color);
    }

    public static void PptxScRgbPreservesAlphaAndExistingTransformOrder()
    {
        var actual = ReadScRgbColor("<a:scrgbClr r=\"50000\" g=\"25000\" b=\"0\"><a:alpha val=\"40000\"/><a:lumMod val=\"50000\"/></a:scrgbClr>");
        TestAssert.Equal(new RgbColor(94, 68, 0), actual.Color);
        TestAssert.Equal(.4d, actual.Alpha);
        foreach (string transforms in new[] {
            "<a:tint val=\"35000\"/><a:shade val=\"60000\"/>",
            "<a:shade val=\"60000\"/><a:tint val=\"35000\"/>" })
        {
            var linear = ReadScRgbColor("<a:scrgbClr r=\"50000\" g=\"25000\" b=\"0\">" + transforms + "</a:scrgbClr>");
            var encoded = ReadScRgbColor("<a:srgbClr val=\"BC8900\">" + transforms + "</a:srgbClr>");
            TestAssert.Equal(encoded, linear);
        }
    }

    public static void PptxScRgbCachedChartPaintMatchesEquivalentSrgb()
    {
        PdfPage Render(string color)
        {
            string chart = $$"""
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea>
                  <c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
                    <c:ser><c:idx val="0"/><c:order val="0"/><c:spPr><a:solidFill>{{color}}</a:solidFill></c:spPr>
                      <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                    <c:axId val="1"/><c:axId val="2"/></c:barChart>
                  <c:catAx><c:axId val="1"/><c:scaling/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
                  <c:valAx><c:axId val="2"/><c:scaling><c:min val="0"/><c:max val="4"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
                </c:plotArea></c:chart></c:chartSpace>
                """;
            using var input = PptxTests.CreateSingleChartPackageStream(chart);
            using var package = OoxPackage.Open(input, default);
            return new PptxRenderer(new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(
                new PptxReader().Read(package, default), package, null, default).ToBlockingEnumerable().Single();
        }
        var linear = Render("<a:scrgbClr r=\"50000\" g=\"25000\" b=\"0\"/>");
        var encoded = Render("<a:srgbClr val=\"BC8900\"/>");
        TestAssert.Contains("0.737 0.537 0 rg", linear.Content);
        TestAssert.Equal(encoded.Content, linear.Content);
    }

    private static (RgbColor Color, double Alpha) ReadScRgbColor(string color)
    {
        var source = XElement.Parse("<a:solidFill xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" + color + "</a:solidFill>");
        TestAssert.True(PptxColorResolver.TryReadSolidColorWithAlpha(source, PptxTheme.Empty, out var rgb, out double alpha),
            "The supplied color form must resolve through the live shared reader.");
        return (rgb, alpha);
    }
}
