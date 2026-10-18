using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionDelimiterAssembliesKeepContentAndNeutralSourceOrigins()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent delimiter controls use installed Windows fonts.");
        CheckDelimiterAssemblies(null);
    }

    internal static void CheckDelimiterAssemblies(IFontResolver? resolver)
    {
        using var gold = JsonDocument.Parse(File.ReadAllText(NestedMathCasePath("delimiter-assembly-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(),
                Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(NestedMathCasePath(id + ".docx")))));
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            bool grow = row.GetProperty("Grow").GetBoolean();
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            string? originalOrigins = null;
            foreach (bool transformed in new[] { false, true })
            {
                var diagnostics = new List<OoxPdfDiagnostic>();
                var page = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(
                        ReadMutation("CompositionProbes/" + id, mode, transformed, transformed), diagnostics.Add, default).Single();
                TestAssert.True(diagnostics.All(diagnostic => diagnostic.Id != "DOCX_UNSUPPORTED_EQUATION"),
                    "Supported tall arguments cannot make the complete growing formula disappear.");
                bool IsMath(Lokad.OoxPdf.Pdf.PdfEmbeddedFont font) => page.Fonts.Any(resource =>
                    resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(resource.Font, font));
                var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).ToArray();
                var logical = shows.Where(show => IsMath(show.Font) && show.Text.Length != 0).ToArray();
                TestAssert.Equal(row.GetProperty("MathText").GetString(), string.Concat(logical.Select(show => show.Text)));
                TestAssert.Equal(3, logical.Length);
                TestAssert.Equal(12d, logical[0].Size);
                TestAssert.Equal(12d, logical[^1].Size);
                TestAssert.True(Math.Abs(logical[1].Size - row.GetProperty("ChildSize").GetDouble()) < 0.1d,
                    "The larger argument keeps its independently authored size/export profile.");
                var physical = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true, honorActualText: false)
                    .Where(show => IsMath(show.Font)).ToArray();
                TestAssert.True((grow ? physical.Length > 3 : physical.Length == 3) &&
                    physical.All(show => show.X > 0d && show.X < page.Width && show.Y > 0d && show.Y < page.Height),
                    "All physical pieces stay on the page; fixed controls retain single glyphs.");
                string origins = string.Join("|", physical.Select(show => show.X.ToString("R", CultureInfo.InvariantCulture) + "," +
                    show.Y.ToString("R", CultureInfo.InvariantCulture)));
                if (originalOrigins is not null) TestAssert.Equal(originalOrigins, origins);
                originalOrigins = origins;
                TestAssert.Contains("Following prose.", DocxPageFootnoteNumberingTests.DecodeText(page));
                if (resolver is not null) continue;
                TestAssert.Equal(row.GetProperty("PhysicalGlyphOperations").GetInt32(), physical.Length);
                foreach (var anchor in row.GetProperty("PhysicalAnchors").EnumerateArray())
                {
                    var show = physical[anchor.GetProperty("Offset").GetInt32()];
                    TestAssert.True(Math.Abs(show.X - anchor.GetProperty("X").GetDouble()) < 1d,
                        "Every delimiter piece keeps the independent one-point horizontal Word gate.");
                    if (!grow || anchor.GetProperty("Text").GetString() == "A")
                        TestAssert.True(Math.Abs(show.Y - anchor.GetProperty("Y").GetDouble()) < 1d,
                            "Fixed pieces and the argument retain independent one-point baseline gates.");
                }
                // Growing-piece Y and following flow remain explicit Office-comparison
                // residuals. Content restoration does not relabel those failed gates.
                if (!grow)
                {
                    var following = shows.First(show => show.Text.StartsWith("Follow", StringComparison.Ordinal));
                    TestAssert.True(Math.Abs(following.Y - row.GetProperty("FollowingY").GetDouble()) < 1d,
                        "Fixed delimiters preserve the validated following paragraph geometry.");
                }
            }
        }
    }
}
