using System.Globalization;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PdfWriterTests
{
    public static void ShadingStopPartitionsPreserveNearbyAndEndpointOffsets()
    {
        foreach (double[] positions in new[] {
            new[] { 0d, .25d, .75d, 1d },
            new[] { 0d, .5001d, .5002d, 1d },
            new[] { 0d, .0001d, .0002d, 1d },
            new[] { 0d, .000001d, .000002d, 1d },
            new[] { 0d, .9999998d, 1d },
            new[] { 0d, double.BitDecrement(1d), 1d } })
        {
            var graphics = new PdfGraphicsBuilder();
            var stops = positions.Select((offset, i) => new PdfShadingStop(offset, (byte)(i * 60), 0, 255)).ToArray();
            // The final two source stops can coalesce by the existing epsilon.
            // Exercise a synthesized terminal interval as a separate endpoint case.
            if (positions.Length == 3) stops = stops[..^1];
            graphics.PaintAxialShading(0, 0, 100, 0, stops);
            var page = new PdfPage(100, 100, graphics.ToString(), [], [], graphics.ExtGStates, graphics.Shadings);
            string pdf = WritePdfText([page]);
            double[] bounds = ReadShadingBounds(pdf);
            TestAssert.True(bounds.SequenceEqual(positions[1..^1]), "Printed bounds preserve each intended unitless position.");
            double previous = 0;
            foreach (double bound in bounds)
            {
                TestAssert.True(bound > previous && bound < 1, "Interior stitching bounds are strictly ordered within the domain.");
                previous = bound;
            }
        }
    }

    public static void ShadingResourceIdentityPreservesDistinctStopPositions()
    {
        PdfShadingStop[] Palette(double before, double after) =>
            [new(0, 255, 0, 0), new(before, 0, 255, 0), new(after, 0, 255, 255), new(1, 0, 0, 255)];
        var graphics = new PdfGraphicsBuilder();
        graphics.PaintShading(new PdfRadialShading(Palette(.500002d, .5000021d)));
        graphics.PaintShading(new PdfRadialShading(Palette(.500002d, .5000022d)));
        graphics.PaintShading(new PdfRadialShading(Palette(.500002d, .5000021d)));
        TestAssert.Equal(2, graphics.Shadings.Count);
        var page = new PdfPage(100, 100, graphics.ToString(), [], [], graphics.ExtGStates, graphics.Shadings);
        string pdf = WritePdfText([page]);
        TestAssert.Equal(2, CountOccurrences(pdf, "/ShadingType 3"));
        TestAssert.Contains("/Bounds [0.500002 0.5000021]", pdf);
        TestAssert.Contains("/Bounds [0.500002 0.5000022]", pdf);
        TestAssert.Equal(2, CountOccurrences(pdf, "/Sh1 sh"));
    }

    public static void ShadingCoalescingKeepsDomainEndpointsAndLatestColor()
    {
        var shading = new PdfAxialShading(0, 0, 1, 0, [
            new(0, 255, 0, 0), new(1e-20, 0, 255, 0),
            new(.5d, 255, 0, 0), new(.5000004d, 0, 0, 255), new(1, 255, 255, 255)]);
        TestAssert.True(shading.Stops.Select(stop => stop.Offset).SequenceEqual(new[] { 0d, .5000004d, 1d }),
            "Coalescing retains the domain endpoints and the latest nearby source stop.");
        TestAssert.Equal((byte)255, shading.Stops[0].Green);
        TestAssert.Equal((byte)255, shading.Stops[1].Blue);
    }

    public static void ShadingHardStopsKeepSeparateIntervalColors()
    {
        var graphics = new PdfGraphicsBuilder();
        graphics.PaintShading(new PdfAxialShading(0, 0, 100, 0, [
            new(0, 255, 0, 0), new(0, 0, 0, 255),
            new(.5d, 0, 255, 0), new(.5d, 255, 255, 0), new(.5d, 1, 2, 3),
            new(1, 0, 0, 255), new(1, 0, 0, 0)], preserveHardStops: true));
        var page = new PdfPage(100, 100, graphics.ToString(), [], [], graphics.ExtGStates, graphics.Shadings);
        string pdf = WritePdfText([page]);

        TestAssert.True(ReadShadingBounds(pdf).SequenceEqual(new[] { .5d }), "A coincident group contributes one interior bound.");
        TestAssert.Equal(2, CountOccurrences(pdf, "/FunctionType 2"));
        TestAssert.Contains("/C0 [0 0 1] /C1 [0 1 0]", pdf);
        TestAssert.Contains("/C0 [0.004 0.008 0.012] /C1 [0 0 1]", pdf);
        TestAssert.Contains("/Encode [0 1 0 1]", pdf);
    }

    private static double[] ReadShadingBounds(string pdf)
    {
        Match match = Regex.Match(pdf, @"/Bounds \[([^\]]*)\]");
        TestAssert.True(match.Success, "A multi-stop shading emits a stitching function.");
        string bounds = match.Groups[1].Value;
        TestAssert.True(bounds.IndexOfAny(['e', 'E']) < 0, "PDF numbers use decimal notation.");
        return bounds.Split(' ', StringSplitOptions.RemoveEmptyEntries).Select(value => double.Parse(value, CultureInfo.InvariantCulture)).ToArray();
    }
}
