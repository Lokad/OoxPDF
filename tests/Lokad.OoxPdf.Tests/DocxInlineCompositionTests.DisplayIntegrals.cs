using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionDisplayIntegralsKeepIndependentWordLimitsAndFlow()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent display-integral origins require installed Windows fonts.");
        CheckDisplayIntegrals(null);
    }

    internal static void CheckDisplayIntegrals(IFontResolver? resolver)
    {
        using var gold = JsonDocument.Parse(File.ReadAllText(NestedMathCasePath("display-integral-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(),
                Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(NestedMathCasePath(id + ".docx")))));
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            string? originalOrigins = null;
            foreach (bool transformed in new[] { false, true })
            {
                var diagnostics = new List<OoxPdfDiagnostic>();
                var document = ReadMutation("CompositionProbes/" + id, mode, transformed, transformed);
                TestAssert.True(document.Paragraphs.SelectMany(p => p.Equations).Single().Display,
                    "The independent source is standalone display math rather than a rewritten inline equation.");
                var page = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(document, diagnostics.Add, default).Single();
                TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                    "Supported integral source survives its display placement and limit location.");
                var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).ToArray();
                var math = shows.Where(show => page.Fonts.Any(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal)
                    && ReferenceEquals(resource.Font, show.Font))).ToArray();
                TestAssert.Equal(row.GetProperty("MathText").GetString(),
                    string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
                TestAssert.Equal(4, math.Length);
                TestAssert.True(math.All(show => show.X > 0d && show.X < page.Width && show.Y > 0d && show.Y < page.Height),
                    "Operator, body and both limits remain on the page.");
                string origins = string.Join("|", math.Select(show => show.X.ToString("R", CultureInfo.InvariantCulture) + "," +
                    show.Y.ToString("R", CultureInfo.InvariantCulture)));
                if (originalOrigins is not null) TestAssert.Equal(originalOrigins, origins);
                originalOrigins = origins;
                TestAssert.Contains("Following prose.", DocxPageFootnoteNumberingTests.DecodeText(page));
                if (resolver is not null) continue;
                foreach (var anchor in row.GetProperty("PhysicalAnchors").EnumerateArray())
                {
                    var show = math[anchor.GetProperty("Offset").GetInt32()];
                    TestAssert.True(Math.Abs(show.X - anchor.GetProperty("X").GetDouble()) < 1d &&
                        Math.Abs(show.Y - anchor.GetProperty("Y").GetDouble()) < 1d &&
                        Math.Abs(show.Size - anchor.GetProperty("Size").GetDouble()) < 0.1d,
                        "Both limit placements preserve the independent one-point Word origin gate and authored size.");
                }
                var following = shows.First(show => show.Text.StartsWith("Follow", StringComparison.Ordinal));
                TestAssert.True(Math.Abs(following.Y - row.GetProperty("FollowingY").GetDouble()) < 1d,
                    "The complete font-owned display group preserves independent following-paragraph flow.");
            }
        }
    }
}
