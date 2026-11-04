using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxColumnIntervalTests
{
    private static string CategoryOrderCases => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", "OuterCategoryFrames"));

    public static void PptxOuterCategoryOrientationUsesWidthAfterNumericAdmission()
    {
        using var oracle = CategoryOrderOracle();
        foreach (string id in new[] { "outer-arial24-heldout", "outer-arial24-width60", "outer-arial18-width44", "outer-arial24-no-numeric", "outer-arial24-width50" })
            CheckCategoryOrderPage(RenderCategoryOrder(oracle, id), CategoryOrderRow(oracle, id));
    }

    public static void PptxOuterCategoryOrientationRespectsAxisFontAndContainerOwners()
    {
        using var oracle = CategoryOrderOracle();
        foreach (string id in new[] { "outer-line90", "outer-line-axis24-90", "outer-area-axis24-90", "outer-high90" })
            CheckCategoryOrderPage(RenderCategoryOrder(oracle, id), CategoryOrderRow(oracle, id));
    }

    public static void PptxOuterProjectedCategoryInkStaysInsideItsChart()
    {
        using var oracle = CategoryOrderOracle();
        foreach (string id in new[] { "outer-quarter45", "outer-area45", "outer-authored45-low", "outer-authored45-high" })
        {
            PdfPage page = RenderCategoryOrder(oracle, id);
            var row = CategoryOrderRow(oracle, id);
            CheckCategoryOrderPage(page, row);
            PptxSourceUnitBoundsTests.CheckCaptionInk(page, row.GetProperty("ExpectedRendererAngle").GetInt32(),
                expectedCount: 8, frameRight: 792d, frameBottom: 72d, frameTop: 504d);
        }
    }

    public static void PptxOuterCategoryAdmissionKeepsFillLegendAndLinkedLabelsTogether()
    {
        using var oracle = CategoryOrderOracle();
        var row = CategoryOrderRow(oracle, "outer-painted90");
        PdfPage page = RenderAdmission(row, link: true, caseDirectory: CategoryOrderCases);
        CheckCategoryOrderPage(page, row);
        var box = AdmissionBounds(page);
        var fill = PptxChartSpacingTests.ReadBarRectangles(page.Content).Single(rect => rect.Width > 200d && rect.Height > 100d);
        TestAssert.True(Math.Abs(fill.X - box.Left) < .002d && Math.Abs(fill.Y - page.Height + box.Bottom) < .002d &&
            Math.Abs(fill.Width - box.Right + box.Left) < .002d && Math.Abs(fill.Height - box.Bottom + box.Top) < .002d,
            "Plot fill and source axes use the same admitted data frame after category fitting.");
        TestAssert.Equal("Regional sales performance", string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page,
            usePageCoordinates: false, fontResourceFilter: name => name.StartsWith("CL", StringComparison.Ordinal)).Select(show => show.Text)));
        var label = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
            fontResourceFilter: name => name.StartsWith("CBD", StringComparison.Ordinal)).Single(show => show.Text == "L");
        var link = page.Annotations.Single();
        TestAssert.Equal("https://example.invalid/manual-caption", link.Uri);
        TestAssert.True(link.X <= label.X + .002d && link.X + link.Width >= label.X + label.Font.MeasureTextPoints(label.Text, label.Size) - .002d &&
            link.Y < label.Y && link.Y + link.Height > label.Y && link.X + link.Width / 2d > box.Left && link.X + link.Width / 2d < box.Left + 45d,
            "The rich data label and link follow the final source-zero slot.");
    }

    public static void PptxOuterExplicitCategoryOrientationPreservesSourceContent()
    {
        using var oracle = CategoryOrderOracle();
        CheckCategoryOrderPage(RenderCategoryOrder(oracle, "outer-arial24-explicit-horizontal"),
            CategoryOrderRow(oracle, "outer-arial24-explicit-horizontal"));
    }

    private static void CheckCategoryOrderPage(PdfPage page, JsonElement row)
    {
        string id = row.GetProperty("Id").GetString()!;
        bool quarter = id.Contains("quarter", StringComparison.Ordinal) || id == "outer-area45";
        var text = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
            fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal)).Select(show => show.Text).ToArray();
        TestAssert.True(text.SequenceEqual(Enumerable.Range(1, 8).Select(i => (quarter ? "Quarter" : "W") + i)),
            id + ": fitting and projected placement preserve every complete source category.");
        int angle = row.GetProperty("ExpectedRendererAngle").GetInt32();
        var matrices = Regex.Matches(page.Content, @"(?m)^(?<a>[-\d.]+) (?<b>[-\d.]+) [-\d.]+ [-\d.]+ [-\d.]+ [-\d.]+ cm\r?$")
            .Select(m => (A: double.Parse(m.Groups["a"].Value, CultureInfo.InvariantCulture), B: double.Parse(m.Groups["b"].Value, CultureInfo.InvariantCulture))).ToArray();
        if (angle == 0)
            TestAssert.True(matrices.All(m => Math.Abs(m.B) < .0005d), id + ": the source/native horizontal control retains its angle.");
        else
            TestAssert.Equal(8, matrices.Count(m => Math.Abs(m.A - Math.Cos(-angle * Math.PI / 180d)) <= .0005d &&
                Math.Abs(m.B - Math.Sin(-angle * Math.PI / 180d)) <= .0005d));
    }

    private static JsonDocument CategoryOrderOracle() => JsonDocument.Parse(File.ReadAllText(Path.Combine(CategoryOrderCases, "office.json")));
    private static JsonElement CategoryOrderRow(JsonDocument oracle, string id) =>
        oracle.RootElement.GetProperty("Rows").EnumerateArray().Single(row => row.GetProperty("Id").GetString() == id);
    private static PdfPage RenderCategoryOrder(JsonDocument oracle, string id) =>
        RenderAdmission(CategoryOrderRow(oracle, id), caseDirectory: CategoryOrderCases);
}
