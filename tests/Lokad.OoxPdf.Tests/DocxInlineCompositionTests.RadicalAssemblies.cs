using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionRadicalAssemblyMatchesIndependentWordPartsAndFollowingProse()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent root assembly parts use installed Windows fonts.");
        CheckRadicalAssembly(null);
    }

    internal static void CheckRadicalAssembly(IFontResolver? resolver)
    {
        using var gold = JsonDocument.Parse(File.ReadAllText(NestedMathCasePath("radical-assembly-office.json")));
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
                var page = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(
                        ReadMutation("CompositionProbes/" + id, mode, transformed, transformed), diagnostics.Add, default).Single();
                TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                    "The complete supported root and child survive when ready-made variants are too small.");
                var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).ToArray();
                bool IsMath(Lokad.OoxPdf.Pdf.PdfEmbeddedFont font) => page.Fonts.Any(resource =>
                    resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(resource.Font, font));
                var logical = shows.Where(show => IsMath(show.Font)).ToArray();
                TestAssert.Equal(row.GetProperty("MathText").GetString(), string.Concat(logical.Select(show => show.Text)));
                TestAssert.Equal(12d, logical[0].Size);
                TestAssert.Equal(200.04d, logical[1].Size); // Existing authored-size export profile; Word emits200.06pt.
                var physical = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true, honorActualText: false)
                    .Where(show => IsMath(show.Font)).ToArray();
                TestAssert.True(physical.Length > logical.Length && physical.All(show => show.X > 0d && show.X < page.Width &&
                    show.Y > 0d && show.Y < page.Height), "All physical root pieces remain visible on the page.");
                string origins = string.Join("|", physical.Select(show => show.X.ToString("R", CultureInfo.InvariantCulture) + "," +
                    show.Y.ToString("R", CultureInfo.InvariantCulture)));
                if (originalOrigins is not null) TestAssert.Equal(originalOrigins, origins);
                originalOrigins = origins;
                TestAssert.Contains("Following prose.", DocxPageFootnoteNumberingTests.DecodeText(page));
                var following = shows.First(show => show.Text.StartsWith("Follow", StringComparison.Ordinal));
                if (resolver is not null) continue;
                TestAssert.Equal(row.GetProperty("PhysicalGlyphOperations").GetInt32(), physical.Length);
                foreach (var anchor in row.GetProperty("PhysicalAnchors").EnumerateArray())
                {
                    var show = physical[anchor.GetProperty("Offset").GetInt32()];
                    TestAssert.True(Math.Abs(show.X - anchor.GetProperty("X").GetDouble()) < 1d &&
                        Math.Abs(show.Y - anchor.GetProperty("Y").GetDouble()) < 1d,
                        "Every physical part retains the independent one-point Word origin gate.");
                }
                TestAssert.True(Math.Abs(following.Y - row.GetProperty("FollowingY").GetDouble()) < 1d,
                    "The control owns the descent floor; unused large-child font descent cannot displace following prose.");
            }
        }
    }
}
