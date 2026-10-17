using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionMappedNoteBandBottomMatchesIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent mapped note-band frames use installed Windows fonts.");
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "Composition"));
        string input = Path.Combine(folder, "revision-composition-report.docx");
        using var golden = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "revision-composition-office.json")));
        var row = golden.RootElement.GetProperty("Rows").EnumerateArray().Single(r => r.GetProperty("Mode").GetString() == "all-word-compatible");
        TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
        // Page two's middle band belongs to the endnote; separate header/footer
        // bands and their remaining vertical failures stay independently recorded.
        var expected = row.GetProperty("Bars").EnumerateArray().Where(bar => bar.GetProperty("PageNumber").GetInt32() == 2)
            .OrderBy(bar => bar.GetProperty("Y").GetDouble()).ElementAt(1);
        var pages = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .RenderBlankPages(DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup), null, default).ToArray();
        var bands = Regex.Matches(pages[1].Content, @"(?m)^(?<x>-?[\d.]+) (?<y>-?[\d.]+) (?<width>[\d.]+) (?<height>[\d.]+) re f\r?$")
            .Cast<Match>().Where(match => Number(match, "width") < 2d && Number(match, "height") > 5d &&
                Number(match, "x") is >= 20d and <= 40d).OrderBy(match => Number(match, "y")).ToArray();
        TestAssert.Equal(3, bands.Length);
        TestAssert.True(Math.Abs(Number(bands[1], "y") - expected.GetProperty("Y").GetDouble()) < 1d,
            "A note's mapped physical height retains the independently measured Word bottom edge.");
        static double Number(Match match, string field) => double.Parse(match.Groups[field].Value, CultureInfo.InvariantCulture);
    }
}
