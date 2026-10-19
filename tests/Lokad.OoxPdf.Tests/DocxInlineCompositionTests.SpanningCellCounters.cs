using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxSpanningCellTotalCountersRetainWordPagesAndCompleteContent()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent spanning-cell origins use installed Windows fonts.");
        CheckSpanningCellTotalCount(null, compareWord: true);
    }

    internal static void CheckSpanningCellTotalCount(IFontResolver? resolver, bool compareWord)
    {
        string directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var gold = JsonDocument.Parse(File.ReadAllText(Path.Combine(directory, "spanning-cell-total-count-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string input = Path.Combine(directory, row.GetProperty("Id").GetString() + ".docx");
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var source = DocxTests.ReadDocx(input, mode);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(source, diagnostics.Add, default).ToArray();
            var byPage = pages.Select(page => Regex.Replace(DocxPageFootnoteNumberingTests.DecodeText(page), @"\s+", " ").Trim()).ToArray();
            foreach (var text in row.GetProperty("ExpectedText").EnumerateArray())
                TestAssert.Contains(text.GetString()!, string.Join(" ", byPage));
            int first = Array.FindIndex(byPage, text => text.Contains("Planning row 1 stays visible.", StringComparison.Ordinal));
            int last = Array.FindIndex(byPage, text => text.Contains("Planning row 84 stays visible.", StringComparison.Ordinal));
            int following = Array.FindIndex(byPage, text => text.Contains("After public math table.", StringComparison.Ordinal));
            TestAssert.True(first >= 0 && last > first && following > last, "The complete cell spans pages and following prose retains its source break.");
            TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"), "A spanning total-count cell retains its child equation.");
            var page = pages[0];
            var mathFonts = page.Fonts.Where(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal)).Select(font => font.Font).ToHashSet();
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).ToArray();
            var math = shows.Where(show => mathFonts.Contains(show.Font) && show.Text.Trim().Length != 0).ToArray();
            TestAssert.Equal("x+2y", string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
            var before = shows.First(show => show.Text.Contains("Before inline", StringComparison.Ordinal));
            string count = pages.Length.ToString(System.Globalization.CultureInfo.InvariantCulture);
            var fields = shows.Where(show => !mathFonts.Contains(show.Font) && show.Text.Trim() == count &&
                Math.Abs(show.Y - before.Y) < .01d).ToArray();
            TestAssert.True(fields.Length != 0, "The field binds to the actual final count within its source box row.");
            if (!compareWord) continue;
            TestAssert.Equal(row.GetProperty("WordPages").GetInt32(), pages.Length);
            var anchors = row.GetProperty("WordMathAnchors").EnumerateArray().ToArray();
            TestAssert.Equal(anchors.Length, math.Length);
            for (int index = 0; index < math.Length; index++) Check(math[index].X, page.Height - math[index].Y, math[index].Size, anchors[index]);
            var field = fields[0];
            Check(field.X, page.Height - field.Y, field.Size, row.GetProperty("WordFieldGlyphs")[0]);
        }

        static void Check(double x, double y, double size, JsonElement anchor) =>
            TestAssert.True(Math.Abs(x - anchor.GetProperty("X").GetDouble()) < 1d &&
                Math.Abs(y - anchor.GetProperty("Y").GetDouble()) < 1d && Math.Abs(size - anchor.GetProperty("Size").GetDouble()) < .1d,
                "Math and count field retain their independent Word glyph gates; separate prose residuals remain in raw reports.");
    }
}
