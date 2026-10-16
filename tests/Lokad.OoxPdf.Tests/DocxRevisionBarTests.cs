using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static class DocxRevisionBarTests
{
    public static void DocxRevisionBarHorizontalFramesMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native Word controls use installed Windows fonts.");
        CheckHorizontalFrames(null);
    }

    internal static void CheckHorizontalFrames(IFontResolver? resolver)
    {
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "RevisionBars"));
        using var reference = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "revision-bars-office-geometry.json")));
        foreach (var row in reference.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            string input = Path.Combine(folder, id + ".docx");
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            TestAssert.Equal(1, row.GetProperty("OfficeView").GetProperty("RevisionsMode").GetInt32());
            var pages = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup), null, default).ToArray();
            var page = pages[row.GetProperty("PageNumber").GetInt32() - 1];
            double x = row.GetProperty("X").GetDouble();
            double width = row.GetProperty("Width").GetDouble();
            var rectangles = Regex.Matches(page.Content,
                @"(?m)^(?<x>-?[\d.]+) (?<y>-?[\d.]+) (?<width>[\d.]+) (?<height>[\d.]+) re f\r?$");
            int matched = rectangles.Cast<Match>().Count(match =>
                double.Parse(match.Groups["height"].Value, CultureInfo.InvariantCulture) > 8d &&
                Math.Abs(double.Parse(match.Groups["x"].Value, CultureInfo.InvariantCulture) - x) < .03d &&
                Math.Abs(double.Parse(match.Groups["width"].Value, CultureInfo.InvariantCulture) - width) < .03d);
            TestAssert.True(matched == 1, "Independent Word X/thickness across margins, orientation and fitted review frame: " + id);
        }
    }
}
