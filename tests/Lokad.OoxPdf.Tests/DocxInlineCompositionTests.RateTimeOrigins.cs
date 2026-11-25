using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    private static void CheckIndependentRateTimeOrigins(string manifest, string[] inputs, bool checkBodyOrigin = false)
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for these independent Word origins.");

        const string finalName = "final-preserve";
        const string reviewName = "all-word-compatible";
        const int formulaCount = 4;
        const int glyphsPerFormula = 4;
        using var gold = JsonDocument.Parse(File.ReadAllText(TestFixtures.PublicCasePath("CompositionProbes/" + manifest)));
        var rows = gold.RootElement.GetProperty("Rows").EnumerateArray().ToArray();
        var expected = new HashSet<(string Input, string Mode)>();
        foreach (string input in inputs)
        {
            expected.Add((input, finalName));
            expected.Add((input, reviewName));
        }
        foreach (var row in rows)
        {
            string input = row.GetProperty("Input").GetString()!;
            string mode = row.GetProperty("Mode").GetString()!;
            TestAssert.True(expected.Remove((input, mode)), $"Unexpected or duplicate Word reference row {input}/{mode} in {manifest}.");
            TestAssert.True(row.GetProperty("MathOrigins").GetArrayLength() == formulaCount,
                $"Expected all four formula origins for {input}/{mode} in {manifest}.");
            if (checkBodyOrigin) TestAssert.True(row.TryGetProperty("BodyOrigin", out _), $"Missing body origin for {input}/{mode}.");
        }
        TestAssert.True(expected.Count == 0, $"Missing Word reference rows in {manifest}: {string.Join(", ", expected)}.");

        foreach (var source in rows.GroupBy(row => row.GetProperty("Input").GetString()!, StringComparer.Ordinal))
        {
            byte[] bytes = File.ReadAllBytes(TestFixtures.PublicCasePath("CompositionProbes/" + source.Key));
            string inputSha256 = Convert.ToHexString(SHA256.HashData(bytes));
            foreach (var row in source)
            {
                TestAssert.Equal(row.GetProperty("InputSha256").GetString(), inputSha256);
                var mode = row.GetProperty("Mode").GetString() == finalName ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
                using var input = new MemoryStream(bytes, writable: false);
                var page = new DocxRenderer(null, mode, CachedMathGeometry(mode))
                    .RenderPagesAsync(DocxTests.ReadDocx(input, mode), null, default).ToBlockingEnumerable().Single();
                var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).ToArray();
                var math = shows.Where(show => page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) &&
                    ReferenceEquals(font.Font, show.Font))).ToArray();
                TestAssert.Equal("ratetimeratetime", string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
                TestAssert.Equal(formulaCount * glyphsPerFormula, math.Length);
                var origins = row.GetProperty("MathOrigins");
                for (int index = 0; index < formulaCount; index++)
                {
                    var actual = math[glyphsPerFormula * index];
                    var anchor = origins[index];
                    TestAssert.Equal(anchor.GetProperty("Text").GetString(), actual.Text.Normalize(NormalizationForm.FormKC));
                    CheckOrigin(actual.X, actual.Y, anchor, $"{source.Key}/{mode} formula {index + 1}");
                }

                if (!checkBodyOrigin) continue;
                // Find the margin-adjacent body line by its visible text,
                // independent of header field fragments and PDF show grouping.
                var body = shows.GroupBy(show => Math.Round(show.Y, 3)).Single(line =>
                    string.Concat(string.Concat(line.OrderBy(show => show.X).Select(show => show.Text))
                        .Normalize(NormalizationForm.FormKC).Where(c => !char.IsWhiteSpace(c))) == "NorthSouth");
                CheckOrigin(body.Min(show => show.X), body.First().Y, row.GetProperty("BodyOrigin"), $"{source.Key}/{mode} body");
            }
        }

        static void CheckOrigin(double x, double y, JsonElement anchor, string label) =>
            TestAssert.True(Math.Abs(x - anchor.GetProperty("X").GetDouble()) < 1d &&
                Math.Abs(y - anchor.GetProperty("Y").GetDouble()) < 1d, $"{label} must retain the one-point Word origin gate.");
    }
}
