using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxScatterIntervalTests
{
    public static void PptxScatterDirectionsMoveDataLabelsWithTheirPoints()
    {
        foreach (bool bubble in new[] { false, true })
        {
            PdfPage normal = Render(bubble, .5d, 7d, 103d, labels: true);
            var axes = Axes(normal);
            var origin = Label(normal);
            foreach (bool reverseX in new[] { false, true })
            foreach (bool reverseY in new[] { false, true })
            {
                PdfPage reversed = Render(bubble, .5d, 7d, 103d, reverseX, reverseY, labels: true);
                var actual = Label(reversed);
                var placed = Axes(reversed);
                // X=8 in .5..8.5 and Y=79 in7..103; default right labels
                // retain the same occupied offset from their reflected point.
                // Caption admission can move the plot when the side changes.
                Near(placed.Left + (reverseX ? .0625d : .9375d) * (placed.Right-placed.Left) -
                    (axes.Left + .9375d * (axes.Right-axes.Left)), actual.X-origin.X);
                Near(placed.Bottom + (reverseY ? .25d : .75d) * (placed.Top-placed.Bottom) -
                    (axes.Bottom + .75d * (axes.Top-axes.Bottom)), actual.Y-origin.Y);
            }
        }

        static (double X, double Y) Label(PdfPage page)
        {
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, false,
                fontResourceFilter: name => name.StartsWith("CSD", StringComparison.Ordinal)).ToArray();
            TestAssert.True(shows.Any(s => s.Text == "79"), "Expected the authored value-only label; visible labels: " + string.Join("|", shows.Select(s => s.Text)));
            var show = shows.Single(s => s.Text == "79");
            return (show.X, show.Y);
        }
    }

    public static void PptxScatterDirectionsShareCrossingsLabelsAndTickSides()
    {
        foreach (bool bubble in new[] { false, true })
        foreach (bool reversed in new[] { false, true })
        foreach (bool numericCrossing in new[] { false, true })
        foreach (string? labelPosition in new[] { null, "low", "high" })
        {
            PdfPage page = Render(bubble, .5d, 7d, 103d, reversed, reversed, labelPosition,
                horizontalCrossing: numericCrossing ? 50d : null, verticalCrossing: numericCrossing ? 4d : null);
            var axes = Axes(page);
            double X(double value) => axes.Left + (reversed ? 1d-(value-.5d)/8d : (value-.5d)/8d) * (axes.Right-axes.Left);
            double Y(double value) => axes.Bottom + (reversed ? 1d-(value-7d)/96d : (value-7d)/96d) * (axes.Top-axes.Bottom);
            Near(X(numericCrossing ? 4d : .5d), axes.VerticalX);
            Near(Y(numericCrossing ? 50d : 7d), axes.HorizontalY);
            var lines = PptxColumnIntervalTests.Lines(page);
            var xTicks = lines.Where(p => p.Color == new RgbColor(0, 255, 0) &&
                Math.Abs(p.V[0]-p.V[2]) < .001d && Math.Abs(p.V[1]-p.V[3]) is > 0d and < 20d).ToArray();
            var yTicks = lines.Where(p => p.Color == new RgbColor(0, 0, 255) &&
                Math.Abs(p.V[1]-p.V[3]) < .001d && Math.Abs(p.V[0]-p.V[2]) is > 0d and < 20d).ToArray();
            TestAssert.Equal(7, xTicks.Length); TestAssert.Equal(4, yTicks.Length);
            for (int i=0; i<xTicks.Length; i++) Near(X(.5d+1.25d*i), xTicks[i].V[0]);
            for (int i=0; i<yTicks.Length; i++) Near(Y(7d+25d*i), yTicks[i].V[1]);
            double[] gridY = lines.Where(p => p.Color == new RgbColor(0, 0, 255) &&
                Math.Abs(p.V[1]-p.V[3]) < .001d && Math.Abs(p.V[0]-p.V[2]) > 100d).Select(p => p.V[1]).Order().ToArray();
            double[] expectedGridY = new[] { 7d, 32d, 57d, 82d }.Where(v => numericCrossing || v != 7d).Select(Y).Order().ToArray();
            TestAssert.Equal(expectedGridY.Length, gridY.Length);
            for (int i=0; i<gridY.Length; i++) Near(expectedGridY[i], gridY[i]);
            // Numeric crossesAt keeps the authored physical bottom/left out
            // direction. Automatic crossings follow the perpendicular scale.
            bool marksHigh = !numericCrossing && reversed;
            TestAssert.True(xTicks.All(p => ((p.V[1]+p.V[3])/2d > axes.HorizontalY) == marksHigh), "Horizontal out marks follow axis ownership.");
            TestAssert.True(yTicks.All(p => ((p.V[0]+p.V[2])/2d > axes.VerticalX) == marksHigh), "Vertical out marks follow axis ownership.");
            bool labelsHigh = (labelPosition == "high") != reversed;
            double labelAxisX = labelPosition is null ? axes.VerticalX : X(labelPosition == "high" ? 8.5d : .5d);
            var yLabels = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, false,
                fontResourceFilter: name => name.StartsWith("CVA", StringComparison.Ordinal)).Where(s => s.Text is "7" or "32" or "57" or "82").ToArray();
            TestAssert.Equal(4, yLabels.Length);
            TestAssert.True(yLabels.All(s => (s.X > labelAxisX) == labelsHigh), "Explicit low/high labels follow value direction independently of the stroke and ticks.");
        }
    }

    public static void PptxScatterBetweenCrossingsKeepValueUnitsAndMaximumOwnership()
    {
        foreach (bool bubble in new[] { false, true })
        foreach (bool reversed in new[] { false, true })
        foreach (string? between in new[] { null, "midCat", "between" })
        {
            var axes = Axes(Render(bubble, .5d, 7d, 103d, reversed, reversed,
                horizontalCrossing: 50d, verticalCrossing: 4d, crossBetween: between));
            double xValue = between == "between" ? 3.5d : 4d;
            double yValue = between == "between" ? 49.5d : 50d;
            Near(axes.Left + (reversed ? 1d-(xValue-.5d)/8d : (xValue-.5d)/8d) * (axes.Right-axes.Left), axes.VerticalX);
            Near(axes.Bottom + (reversed ? 1d-(yValue-7d)/96d : (yValue-7d)/96d) * (axes.Top-axes.Bottom), axes.HorizontalY);
            var maximum = Axes(Render(bubble, .5d, 7d, 103d, reversed, reversed, crossBetween: between, verticalMaximum: true));
            Near(reversed ? maximum.Left : maximum.Right, maximum.VerticalX);
        }
    }

    private static (double Left, double Right, double Bottom, double Top, double HorizontalY, double VerticalX) Axes(PdfPage page)
    {
        var lines = PptxColumnIntervalTests.Lines(page);
        var x = lines.Single(p => p.Color == default && Math.Abs(p.V[1]-p.V[3]) < .001d && Math.Abs(p.V[0]-p.V[2]) > 100d).V;
        var y = lines.Single(p => p.Color == default && Math.Abs(p.V[0]-p.V[2]) < .001d && Math.Abs(p.V[1]-p.V[3]) > 100d).V;
        return (x[0], x[2], y[1], y[3], x[1], y[0]);
    }

    private static void Near(double expected, double actual) => TestAssert.True(Math.Abs(expected-actual) < .002d,
        $"Expected coordinate {expected}, got {actual}.");
}
