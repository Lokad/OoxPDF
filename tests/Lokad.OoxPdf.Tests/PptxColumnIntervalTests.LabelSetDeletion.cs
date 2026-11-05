using System.Text.Json;
using System.Text.RegularExpressions;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxColumnIntervalTests
{
    public static void PptxLabelSetDeletionMatchesIndependentOfficeScopes()
    {
        string directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "Cases", "LabelSetDeletion"));
        using var oracle = JsonDocument.Parse(File.ReadAllText(Path.Combine(directory, "office.json")));
        foreach (var row in oracle.RootElement.GetProperty("Rows").EnumerateArray())
        {
            var page = RenderAdmission(row, caseDirectory: directory);
            string actual = string.Join(" ", DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, false,
                fontResourceFilter: name => name.StartsWith("CBD", StringComparison.Ordinal)).Select(show => show.Text));
            string expected = string.Join(" ", row.GetProperty("ExpectedDataLabels").EnumerateArray().Select(value => value.GetString()));
            // Source order differs between Office and the renderer. Compare
            // complete visible tokens, including point names and punctuation,
            // across font-dependent line wrapping without assuming one show.
            static string Tokens(string text) => string.Join("|", Regex.Matches(text, @"[A-Za-z]+\d*|\d+|[,;%]")
                .Select(match => match.Value).Order(StringComparer.Ordinal));
            TestAssert.Equal(Tokens(expected), Tokens(actual));
        }
    }

    public static void PptxDeletedSeriesKeepsAxisCaptionsAndLegend()
    {
        using var oracle = LogicalCaptionOracle();
        foreach (string id in new[] { "outer-high45-no-data-labels", "reference-numeric-format", "logical-reversed-hidden12" })
        {
            var page = RenderLogicalCaption(oracle, id);
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, false).ToArray();
            TestAssert.True(!page.Fonts.Any(font => font.ResourceName.StartsWith("CBD", StringComparison.Ordinal)),
                "A deleted series with no point requests emits no data-label resources or paint.");
            TestAssert.Equal(1, shows.Count(show => show.Text.Contains("Regional", StringComparison.Ordinal)));
            string categories = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, false,
                fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal)).Select(show => show.Text));
            TestAssert.Equal(string.Concat(LogicalCaptionRow(oracle, id).GetProperty("ExpectedCaptions")
                .EnumerateArray().Select(value => value.GetString())), categories);
            if (id == "reference-numeric-format")
                TestAssert.True(AdmissionNumericShows(page).Select(show => show.Text)
                    .SequenceEqual(new[] { "7.00", "32.00", "57.00", "82.00", "107.00" }),
                    "Suppressing series labels preserves complete formatted numeric-axis captions.");
        }
    }
}
