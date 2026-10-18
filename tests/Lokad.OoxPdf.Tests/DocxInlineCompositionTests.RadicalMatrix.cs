using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionRadicalMatrixKeepsIndependentWordContentAndOrigins()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent radical/matrix origins use installed Windows fonts.");
        CheckRadicalMatrix(null);
    }

    internal static void CheckRadicalMatrix(IFontResolver? resolver)
    {
        using var gold = JsonDocument.Parse(File.ReadAllText(NestedMathCasePath("radical-matrix-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        foreach (bool cell in new[] { false, true })
        {
            string id = row.GetProperty("Id").GetString()!;
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(),
                Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(NestedMathCasePath(id + ".docx")))));
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            string? originalOrigins = null;
            foreach (bool transformed in new[] { false, true })
            {
                var source = ReadMutation("CompositionProbes/" + id, mode, transformed, transformed, (xml, w, _) =>
                {
                    if (!cell) return;
                    var paragraph = xml.Root!.Element(w + "body")!.Elements(w + "p").First();
                    paragraph.ReplaceWith(new XElement(w + "tbl",
                        new XElement(w + "tblPr", new XElement(w + "tblW", new XAttribute(w + "w", "9360"), new XAttribute(w + "type", "dxa"))),
                        new XElement(w + "tblGrid", new XElement(w + "gridCol", new XAttribute(w + "w", "9360"))),
                        new XElement(w + "tr", new XElement(w + "tc",
                            new XElement(w + "tcPr", new XElement(w + "tcW", new XAttribute(w + "w", "9360"), new XAttribute(w + "type", "dxa"))),
                            new XElement(paragraph)))));
                });
                var diagnostics = new List<OoxPdfDiagnostic>();
                var renderer = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
                var page = renderer.RenderBlankPages(source, diagnostics.Add, default).Single();
                TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                    "Supported radical/matrix composition survives ordinary cell placement and neutral source metadata.");
                var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true);
                var math = shows.Where(show => page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) &&
                    ReferenceEquals(font.Font, show.Font))).ToArray();
                TestAssert.Equal(row.GetProperty("MathGlyphCount").GetInt32(), math.Length);
                TestAssert.Equal(row.GetProperty("MathText").GetString(),
                    string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
                TestAssert.Contains("Following prose.", DocxPageFootnoteNumberingTests.DecodeText(page));
                TestAssert.True(math.All(show => show.X >= 0d && show.X <= page.Width && show.Y >= 0d && show.Y <= page.Height),
                    "Every mathematical origin remains inside its physical page.");
                string origins = string.Join("|", math.Select(show => show.X.ToString("R", CultureInfo.InvariantCulture) + "," +
                    show.Y.ToString("R", CultureInfo.InvariantCulture)));
                if (originalOrigins is not null) TestAssert.Equal(originalOrigins, origins);
                originalOrigins = origins;
                if (resolver is not null || cell || transformed) continue;
                foreach (var anchor in row.GetProperty("Anchors").EnumerateArray())
                {
                    var show = math[anchor.GetProperty("Offset").GetInt32()];
                    TestAssert.True(Math.Abs(show.X - anchor.GetProperty("X").GetDouble()) < 1d &&
                        Math.Abs(show.Y - anchor.GetProperty("Y").GetDouble()) < 1d,
                        $"Bounded radical/matrix geometry meets independent Word's existing one-point gate: {id}/{mode}.");
                }
                var following = shows.Single(show => show.Text == "Following prose.");
                TestAssert.True(Math.Abs(following.Y - row.GetProperty("FollowingY").GetDouble()) < 1d,
                    "Following prose consumes the complete composed mathematical height.");
            }
        }
    }
}
