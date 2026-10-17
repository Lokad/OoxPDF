using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionCellConnectorsKeepFirstIndependentWordAnchorAndTypeColor()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent connector geometry uses installed Windows fonts.");
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var golden = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "cell-connectors-first-office.json")));
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            var document = ReadMutation("CompositionProbes/" + id, OoxPdfDocxMarkupMode.AllMarkup, false, false);
            var renderer = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var anchor = renderer.InspectMarkupBalloons(document).MaxBy(balloon => balloon.Y)!;
            TestAssert.True(Math.Abs(anchor.AnchorConnectorX - row.GetProperty("FirstAnchorX").GetDouble()) < 1d &&
                Math.Abs(anchor.AnchorY - row.GetProperty("FirstAnchorY").GetDouble()) < 1d,
                $"The first structural connector retains its independent source frame: {id}/{anchor.AnchorConnectorX}/{anchor.AnchorY}.");
            var page = renderer.RenderBlankPages(document, null, default).Single();
            string stroke = row.GetProperty("StrokeColor").GetString()!.Replace("RG:", "", StringComparison.Ordinal).Replace(',', ' ') + " RG";
            TestAssert.Contains(stroke, page.Content);
        }
    }

    public static void DocxInlineCompositionCellConnectorsKeepMathAndFirstSourceAttachment()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native connector content uses installed Windows fonts.");
        CheckCellConnectorContent(null);
    }

    internal static void CheckCellConnectorContent(IFontResolver? resolver)
    {
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var golden = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "cell-connectors-office.json")));
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(Path.Combine(folder, id + ".docx")))));
            var document = ReadMutation("CompositionProbes/" + id, OoxPdfDocxMarkupMode.AllMarkup, false, false);
            var renderer = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var page = renderer.RenderBlankPages(document, null, default).Single();
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).ToArray();
            var fonts = page.Fonts.Where(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal)).Select(font => font.Font).ToHashSet();
            var math = shows.Where(show => fonts.Contains(show.Font)).ToArray();
            TestAssert.Equal(row.GetProperty("MathGlyphCount").GetInt32(), math.Length);
            TestAssert.Equal(row.GetProperty("MathText").GetString(), string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
            TestAssert.Equal(row.GetProperty("ImageCount").GetInt32(), page.Images.Count);
            // A line paints its prose before its math, so math operations can
            // sit between two wrapped prose words in paint order. Screen prose
            // separately; the independent math text/count checks above remain.
            string prose = System.Text.RegularExpressions.Regex.Replace(
                string.Join(" ", shows.Where(show => !fonts.Contains(show.Font)).Select(show => show.Text)), @"\s+", " ").Trim();
            foreach (var text in row.GetProperty("ExpectedText").EnumerateArray())
                TestAssert.Contains(text.GetString()!, prose);
            var summary = shows.First(show => show.Text.Contains("Summary:", StringComparison.Ordinal));
            var anchor = renderer.InspectMarkupBalloons(document).MaxBy(balloon => balloon.Y)!;
            TestAssert.True(Math.Abs(anchor.AnchorConnectorX - summary.X) < .01d,
                "The first connector starts at its printed source paragraph on both font platforms.");
            // These checks cover content and first-source attachment. Raw Word
            // counts still expose the separate same-author range coalescing gap.
        }
    }
}
