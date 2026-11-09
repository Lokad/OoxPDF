using System.Text.Json;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxNaturalBodyLinesMatchIndependentWordOriginsInFinalAndReview()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent natural body line origins use installed Windows fonts.");
        using var cases = ReadNaturalBodyLinesJson("natural-lines.json");
        using var gold = ReadNaturalBodyLinesJson("natural-lines-office.json");
        foreach (var expected in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            var source = cases.RootElement.EnumerateArray().Single(row => row.GetProperty("Id").GetString() == expected.GetProperty("Id").GetString());
            TestAssert.Equal(source.GetProperty("Sha256").GetString(), expected.GetProperty("InputSha256").GetString());
            var mode = expected.GetProperty("Mode").GetString() == "final-preserve" ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var page = RenderContainerGrid(source, mode, null);
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true).ToArray();
            foreach (var origin in expected.GetProperty("TextOrigins").EnumerateArray())
            {
                var actual = shows.Single(show => show.Text.Trim() == origin.GetProperty("Text").GetString());
                TestAssert.True(Math.Abs(actual.X - origin.GetProperty("X").GetDouble()) < 1d &&
                    Math.Abs(page.Height - actual.Y - origin.GetProperty("Y").GetDouble()) < 1d,
                    "Actual line metrics retain independent absolute origins: " + source.GetProperty("Id").GetString() + "/" + mode + "/" + origin.GetProperty("Text").GetString());
            }
        }
    }

    private static JsonDocument ReadNaturalBodyLinesJson(string name) => JsonDocument.Parse(File.ReadAllText(Path.GetFullPath(
        Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "NaturalBodyLines", name))));
}
