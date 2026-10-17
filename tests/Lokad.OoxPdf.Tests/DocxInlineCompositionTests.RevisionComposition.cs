using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionCompleteRevisionsKeepDistinctWordStoryBands()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent complete-report frames use installed Windows fonts.");
        CheckCompleteRevisionComposition(null, true);
    }

    internal static void CheckCompleteRevisionComposition(IFontResolver? resolver, bool compareWordFrame)
    {
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "Composition"));
        string input = Path.Combine(folder, "revision-composition-report.docx");
        using var golden = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "revision-composition-office.json")));
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray())
        {
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            bool review = row.GetProperty("Mode").GetString() == "all-word-compatible";
            var mode = review ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final;
            var pages = new DocxRenderer(resolver, mode, review ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup :
                OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(DocxTests.ReadDocx(input, mode), null, default).ToArray();
            if (compareWordFrame) TestAssert.Equal(row.GetProperty("Pages").GetInt32(), pages.Length);
            var actual = new StringBuilder();
            int redMath = 0, totalBars = 0;
            for (int pageNumber = 1; pageNumber <= pages.Length; pageNumber++)
            {
                var page = pages[pageNumber - 1];
                var fonts = page.Fonts.Where(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal)).ToArray();
                var embedded = fonts.Select(font => font.Font).ToHashSet();
                foreach (var glyph in DocxCommentsTests.ReadEmbeddedGlyphTextShows(page))
                    if (embedded.Contains(glyph.Font)) actual.Append(glyph.Text);
                var names = fonts.Select(font => font.ResourceName).ToHashSet();
                foreach (Match block in Regex.Matches(page.Content, @"(?s)BT\s+(?<body>.*?)\s+ET"))
                {
                    if (!names.Contains(Regex.Match(block.Value, @"/(?<name>\S+) [\d.]+ Tf").Groups["name"].Value)) continue;
                    var color = Regex.Match(block.Value, @"(?<r>[\d.]+) (?<g>[\d.]+) (?<b>[\d.]+) rg");
                    if (color.Success && Math.Abs(Number(color, "r") - .82d) < .002d &&
                        Math.Abs(Number(color, "g") - .204d) < .002d && Math.Abs(Number(color, "b") - .22d) < .002d) redMath++;
                }
                var bars = Regex.Matches(page.Content, @"(?m)^(?<x>-?[\d.]+) (?<y>-?[\d.]+) (?<width>[\d.]+) (?<height>[\d.]+) re f\r?$")
                    .Cast<Match>().Where(match => Number(match, "width") < 2d && Number(match, "height") > 5d &&
                        Number(match, "x") is >= 20d and <= 40d).OrderBy(match => Number(match, "y")).ToArray();
                totalBars += bars.Length;
                if (compareWordFrame)
                {
                    var expected = row.GetProperty("Bars").EnumerateArray().Where(bar => bar.GetProperty("PageNumber").GetInt32() == pageNumber)
                        .OrderBy(bar => bar.GetProperty("Y").GetDouble()).ToArray();
                    TestAssert.Equal(expected.Length, bars.Length);
                    for (int i = 0; i < bars.Length; i++)
                    {
                        TestAssert.True(Math.Abs(Number(bars[i], "x") - expected[i].GetProperty("X").GetDouble()) < .03d &&
                            Math.Abs(Number(bars[i], "width") - expected[i].GetProperty("Width").GetDouble()) < .03d,
                            "Independent Word story bands retain horizontal print frames.");
                        if (i != 0) TestAssert.True(Number(bars[i - 1], "y") + Number(bars[i - 1], "height") < Number(bars[i], "y"),
                            "Distinct stories cannot join through the unchanged page area.");
                    }
                }
                TestAssert.True(page.Images.Count != 0, "The repeated footer image remains present.");
            }
            string expectedMath = row.GetProperty("FixedMath").GetString() +
                string.Concat(Enumerable.Repeat(row.GetProperty("RepeatedMathPerPage").GetString(), pages.Length));
            TestAssert.True(actual.ToString().Normalize(NormalizationForm.FormKC).EnumerateRunes().Select(rune => rune.Value).Order()
                .SequenceEqual(expectedMath.Normalize(NormalizationForm.FormKC).EnumerateRunes().Select(rune => rune.Value).Order()),
                "Source-backed independent Word math content survives body/static/note composition, including normal-font formulas.");
            TestAssert.Equal(row.GetProperty("FixedRedMathGlyphs").GetInt32() + pages.Length * row.GetProperty("RepeatedRedMathGlyphsPerPage").GetInt32(), redMath);
            TestAssert.Equal(review ? 2 * pages.Length + 2 : 0, totalBars);
        }

        static double Number(Match match, string field) => double.Parse(match.Groups[field].Value, CultureInfo.InvariantCulture);
    }
}
