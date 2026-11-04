using System.Text.Json;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxColumnIntervalTests
{
    private static string OuterCases => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", "OuterPlotFrames"));

    public static void PptxManualOuterFramesUseIndependentCaptionSides()
    {
        using var oracle = OuterOracle();
        foreach (string id in new[] { "outer-factor-y16", "outer-x10-y0", "outer-numeric-right", "outer-line-x10-y0", "outer-area-x10-y0" })
        {
            var row = OuterRow(oracle, id);
            PdfPage page = RenderAdmission(row, caseDirectory: OuterCases);
            var box = AdmissionBounds(page);
            double[] native = AdmissionNumbers(row, "NativeFrame");
            double advanceDelta = AdmissionNumericShows(page).Max(show => show.Font.MeasureTextPoints(show.Text, show.Size)) -
                row.GetProperty("NativeNumericAdvance").GetDouble();
            bool right = row.GetProperty("NumericRight").GetBoolean();
            TestAssert.True(Math.Abs(box.Left - native[0] - (right ? 0d : advanceDelta)) <= 1d &&
                Math.Abs(box.Right - native[2] + (right ? advanceDelta : 0d)) <= 1d &&
                Math.Abs(box.Top - native[1]) <= 1d && Math.Abs(box.Bottom - native[3]) <= 1d,
                id + ": measured font substitution retains the independent native outer-frame owners.");
        }
    }

    public static void PptxManualOuterHiddenAxesRetainTheirIndependentMinimumBand()
    {
        using var oracle = OuterOracle();
        foreach (string id in new[] { "outer-hide-both-y16", "outer-hidden-no-ticks", "outer-hidden-inside-category",
            "outer-hidden-category-out", "outer-hidden-out12", "outer-hidden-out24", "outer-hidden-offset0", "outer-hidden-offset200", "outer-hidden-arial18" })
        {
            var row = OuterRow(oracle, id);
            PdfPage page = RenderAdmission(row, caseDirectory: OuterCases);
            var box = AdmissionBounds(page);
            double[] native = AdmissionNumbers(row, "NativeFrame");
            TestAssert.True(new[] { box.Left, box.Top, box.Right, box.Bottom }.Zip(native).All(p => Math.Abs(p.First - p.Second) <= 1d),
                id + ": hidden ink and label offsets do not release the present axis's minimum external band.");
            TestAssert.Equal(0, AdmissionNumericShows(page).Length);
        }
        PdfPage deleted = RenderAdmission(OuterRow(oracle, "outer-deleted-axes"), caseDirectory: OuterCases);
        TestAssert.True(!Lines(deleted).Any(line => line.Color == new RgbColor(255, 0, 0) || line.Color == new RgbColor(0, 170, 0)),
            "Deleted axes remain absent; minimum layout reserves do not create painted axes.");
    }

    public static void PptxManualOuterEndpointsOwnTheirMeasuredCaptionFrames()
    {
        using var oracle = OuterOracle();
        PdfPage wide = RenderAdmission(OuterRow(oracle, "outer-width35-left12-y03"), caseDirectory: OuterCases);
        PdfPage shortened = RenderAdmission(OuterRow(oracle, "outer-width35-last-short-fixed"), caseDirectory: OuterCases);
        var before = AdmissionBounds(wide);
        var after = AdmissionBounds(shortened);
        TestAssert.True(after.Right > before.Right + 1d && Math.Abs(after.Left - before.Left) < .002d &&
            Math.Abs(after.Top - before.Top) < .002d && Math.Abs(after.Bottom - before.Bottom) < .002d,
            "The independently observed endpoint mutation releases only its horizontal caption reserve.");
        var captions = DocxCommentsTests.ReadEmbeddedGlyphTextShows(shortened, usePageCoordinates: false,
            fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal)).Select(show => show.Text).ToArray();
        TestAssert.True(captions.SequenceEqual(Enumerable.Range(1, 7).Select(i => "W" + i).Append("x")),
            "Frame admission retains every source category, including the changed endpoint.");
    }

    public static void PptxManualOuterPlotFillUsesTheNativeDataRectangle()
    {
        using var oracle = OuterOracle();
        PdfPage page = RenderAdmission(OuterRow(oracle, "outer-painted-plot"), caseDirectory: OuterCases);
        var box = AdmissionBounds(page);
        var fill = PptxChartSpacingTests.ReadBarRectangles(page.Content).Single(rect => rect.Width > 200d && rect.Height > 200d);
        TestAssert.True(Math.Abs(fill.X - box.Left) < .002d && Math.Abs(fill.Y - page.Height + box.Bottom) < .002d &&
            Math.Abs(fill.Width - box.Right + box.Left) < .002d && Math.Abs(fill.Height - box.Bottom + box.Top) < .002d,
            "The independently observed native plot fill follows data axes instead of covering the caption bands.");
    }

    public static void PptxManualOuterLinkedLabelsFollowTheirAdmittedPlot()
    {
        using var oracle = OuterOracle();
        PdfPage page = RenderAdmission(OuterRow(oracle, "outer-x10-y0"), link: true, caseDirectory: OuterCases);
        var box = AdmissionBounds(page);
        var show = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
            fontResourceFilter: name => name.StartsWith("CBD", StringComparison.Ordinal)).Single(show => show.Text == "L");
        var link = page.Annotations.Single();
        TestAssert.Equal("https://example.invalid/manual-caption", link.Uri);
        TestAssert.True(link.X <= show.X + .002d && link.X + link.Width >= show.X + show.Font.MeasureTextPoints(show.Text, show.Size) - .002d &&
            link.Y < show.Y && link.Y + link.Height > show.Y && link.X + link.Width / 2d > box.Left && link.X + link.Width / 2d < box.Left + 45d,
            "The rich source label and its annotation share the admitted source-zero slot.");
    }

    private static JsonDocument OuterOracle() => JsonDocument.Parse(File.ReadAllText(Path.Combine(OuterCases, "office.json")));
    private static JsonElement OuterRow(JsonDocument oracle, string id) =>
        oracle.RootElement.GetProperty("Rows").EnumerateArray().Single(row => row.GetProperty("Id").GetString() == id);
}
