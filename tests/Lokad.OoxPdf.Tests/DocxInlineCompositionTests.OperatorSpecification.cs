using System.Globalization;
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
    public static void DocxInlineCompositionOperatorMatrixSpecificationRetainsSourceContentAndNeutralOrigins()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native compound specification uses installed Windows fonts.");
        CheckOperatorMatrixSpecification(null);
    }

    internal static void CheckOperatorMatrixSpecification(IFontResolver? resolver)
        => CheckComposedOperatorSpecification(resolver, "operator-matrix-specification", requireWordPageCount: false);

    public static void DocxInlineCompositionMixedArgumentSpecificationRetainsIndependentContentAndPages()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native mixed-argument specification uses installed Windows fonts.");
        CheckMixedArgumentSpecification(null);
    }

    internal static void CheckMixedArgumentSpecification(IFontResolver? resolver)
        => CheckComposedOperatorSpecification(resolver, "mixed-argument-specification", requireWordPageCount: true);

    private static void CheckComposedOperatorSpecification(IFontResolver? resolver, string name, bool requireWordPageCount)
    {
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "Composition"));
        using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, name + ".json")));
        var fixture = manifest.RootElement[0];
        TestAssert.True(!fixture.GetProperty("HeldOut").GetBoolean(), "Compound controls derive only from development inputs.");
        using var gold = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, name + "-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(Path.Combine(cases, id + ".docx")))));
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            string? originalPages = null, originalMath = null;
            foreach (bool transformed in new[] { false, true })
            {
                var source = ReadMutation("Composition/" + id, mode, transformed, transformed);
                TestAssert.Equal(6, source.EquationParts.Sum(part => part.SourceCount));
                var diagnostics = new List<OoxPdfDiagnostic>();
                var renderer = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
                var pages = renderer.RenderBlankPages(source, diagnostics.Add, default).ToArray();
                TestAssert.True(pages.Length >= 2, "The specification exercises page boundaries and repeated static content.");
                // Portable faces have their own metrics. Native composed controls
                // can independently qualify Word page membership when requested.
                if (requireWordPageCount && resolver is null)
                    TestAssert.Equal(row.GetProperty("WordPages").GetInt32(), pages.Length);
                TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Every supported source formula survives the compound document.");
                string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
                foreach (var prose in fixture.GetProperty("ExpectedText").EnumerateArray()) TestAssert.Contains(prose.GetString()!, text);
                var math = pages.SelectMany((page, index) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                    .Where(show => page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(font.Font, show.Font)))
                    .Select(show => (Page: index, Show: show))).ToArray();
                string mathematicalText = string.Concat(math.Select(item => item.Show.Text)).Normalize(NormalizationForm.FormKC);
                var expected = row.GetProperty("BodyMathCharacters").EnumerateObject().ToDictionary(item => item.Name[0], item => item.Value.GetInt32());
                foreach (var item in row.GetProperty("HeaderMathCharactersPerPage").EnumerateObject())
                    expected[item.Name[0]] = expected.GetValueOrDefault(item.Name[0]) + item.Value.GetInt32() * pages.Length;
                var actual = mathematicalText.GroupBy(character => character).ToDictionary(group => group.Key, group => group.Count());
                TestAssert.Equal(expected.Count, actual.Count);
                foreach (var item in expected) TestAssert.Equal(item.Value, actual.GetValueOrDefault(item.Key));
                TestAssert.Equal(row.GetProperty("Images").GetInt32(), pages.Sum(page => page.Images.Sum(image => Regex.Matches(page.Content, "/" + Regex.Escape(image.ResourceName) + @"\s+Do\b").Count)));
                var links = pages.SelectMany(page => page.Annotations).ToArray();
                TestAssert.Equal(1, links.Count(link => link.Uri == row.GetProperty("ExternalLink").GetString()));
                TestAssert.Equal(row.GetProperty("InternalLinksIncludingFootnote").GetInt32(), links.Count(link => link.Destination is not null));
                string pageContent = string.Join("|", pages.Select(page => new string(DocxPageFootnoteNumberingTests.DecodeText(page)
                    .Normalize(NormalizationForm.FormKC).Where(character => !char.IsWhiteSpace(character)).OrderBy(character => character).ToArray())));
                string mathOrigins = string.Join("|", math.Select(item => item.Page + ":" + item.Show.X.ToString("R", CultureInfo.InvariantCulture) + "," + item.Show.Y.ToString("R", CultureInfo.InvariantCulture)));
                if (originalPages is not null) TestAssert.Equal(originalPages, pageContent);
                if (originalMath is not null) TestAssert.Equal(originalMath, mathOrigins);
                originalPages = pageContent; originalMath = mathOrigins;
            }
        }
    }
}
