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
}
