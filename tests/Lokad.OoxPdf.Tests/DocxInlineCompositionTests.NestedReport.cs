using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionNestedReportKeepsCompleteContentNavigationAndNeutralTransformations()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent compound report controls use installed Windows fonts.");
        CheckNestedReport(null);
    }

    internal static void CheckNestedReport(IFontResolver? resolver)
    {
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "Composition"));
        using var sourceManifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "nested-report.json")));
        var sourceCase = sourceManifest.RootElement[0];
        TestAssert.True(!sourceCase.GetProperty("HeldOut").GetBoolean(),
            "Development composition does not derive from or fit a held-out document.");
        TestAssert.Equal(sourceCase.GetProperty("Sha256").GetString(), Convert.ToHexString(SHA256.HashData(
            File.ReadAllBytes(Path.Combine(cases, "composed-nested-navigation-report.docx")))));
        using var gold = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "nested-report-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            var mode = row.GetProperty("Mode").GetString() == "final-preserve" ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var renderer = new DocxRenderer(resolver, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            string? baselineText = null, baselineMath = null;
            foreach (bool transformed in new[] { false, true })
            {
                var source = ReadMutation("Composition/composed-nested-navigation-report", mode, transformed, transformed);
                TestAssert.Equal(8, source.EquationParts.Sum(part => part.SourceCount));
                var diagnostics = new List<OoxPdfDiagnostic>();
                var pages = renderer.RenderBlankPages(source, diagnostics.Add, default).ToArray();
                TestAssert.True(pages.Length >= 2, "The report exercises an actual page boundary and repeated static stories.");
                if (resolver is null) TestAssert.Equal(row.GetProperty("Pages").GetInt32(), pages.Length);
                TestAssert.True(diagnostics.All(diagnostic => diagnostic.Id != "DOCX_UNSUPPORTED_EQUATION"),
                    "Every independently supported equation survives the compound document.");
                string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
                foreach (var prose in sourceCase.GetProperty("ExpectedText").EnumerateArray()) TestAssert.Contains(prose.GetString()!, text);
                foreach (var math in row.GetProperty("BodyMath").EnumerateObject())
                {
                    int count = math.Value.GetInt32();
                    if (row.GetProperty("HeaderMathPerPage").TryGetProperty(math.Name, out var header)) count += header.GetInt32() * pages.Length;
                    TestAssert.Equal(count, text.Split(math.Name, StringSplitOptions.None).Length - 1);
                }
                TestAssert.Equal(row.GetProperty("Images").GetInt32(), pages.Sum(page => page.Images.Sum(image =>
                    Regex.Matches(page.Content, "/" + Regex.Escape(image.ResourceName) + @"\s+Do\b").Count)));
                var links = pages.SelectMany(page => page.Annotations).ToArray();
                TestAssert.Equal(1, links.Count(link => link.Uri == row.GetProperty("ExternalLink").GetString()));
                var destinations = links.Where(link => link.Destination is not null).ToArray();
                TestAssert.Equal(row.GetProperty("InternalLinksIncludingFootnote").GetInt32(), destinations.Length);
                TestAssert.Equal(2, destinations.Select(link => link.Destination).Distinct().Count());
                TestAssert.True(destinations.All(link => link.Width > 0d && link.Height > 0d &&
                    link.Destination!.Value.PageIndex >= 0 && link.Destination.Value.PageIndex < pages.Length),
                    "The fraction and footnote retain distinct concrete destinations.");
                var nested = destinations.Single(link => link.Destination!.Value.PageIndex == 0);
                var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(pages[0], usePageCoordinates: true);
                string linked = string.Concat(shows.Where(show => show.X >= nested.X - .01d && show.X <= nested.X + nested.Width + .01d &&
                    show.Y >= nested.Y - .01d && show.Y <= nested.Y + nested.Height + .01d).Select(show => show.Text))
                    .Normalize(NormalizationForm.FormKC);
                TestAssert.Contains("ratetime", linked);
                // Synthetic bold paints a title twice. Run splitting changes
                // the packet order of that overpaint, so compare character
                // multisets per physical page rather than PDF show order.
                string pageContent = string.Join("|", pages.Select(page => new string(
                    DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC)
                        .Where(character => !char.IsWhiteSpace(character)).OrderBy(character => character).ToArray())));
                if (baselineText is not null) TestAssert.Equal(baselineText, pageContent);
                baselineText = pageContent;
                if (resolver is not null) continue;
                string mathGeometry = string.Join("|", pages.SelectMany((page, pageIndex) =>
                    DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                        .Where(show => page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) &&
                            ReferenceEquals(font.Font, show.Font)))
                        .Select(show => pageIndex + ":" + show.Text + "@" + show.X.ToString("0.###", CultureInfo.InvariantCulture) +
                            "," + show.Y.ToString("0.###", CultureInfo.InvariantCulture))));
                if (baselineMath is not null) TestAssert.Equal(baselineMath, mathGeometry);
                baselineMath = mathGeometry;
            }
        }
    }
}
