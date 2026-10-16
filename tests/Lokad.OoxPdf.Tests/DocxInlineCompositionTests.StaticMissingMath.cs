using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    internal static void CheckStaticMissingMath(IFontResolver resolver)
    {
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var golden = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "static-source-frames-office.json")));
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray())
        foreach (bool final in new[] { true, false })
        {
            string id = row.GetProperty("Id").GetString()!;
            string path = Path.Combine(folder, row.GetProperty("File").GetString()!);
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var document = DocxTests.ReadDocx(path, mode);
            var paragraphs = document.HeaderParagraphsByType.Values.Concat(document.FooterParagraphsByType.Values)
                .SelectMany(value => value).ToArray();
            var owners = paragraphs.SelectMany(paragraph => paragraph.Runs).ToArray();
            var diagnostics = new List<OoxPdfDiagnostic>();
            var page = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(document, diagnostics.Add, default).Single();
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).ToArray();
            TestAssert.Equal("1.,2.,3.", string.Join(',', shows.Where(show => show.Text.Trim() is "1." or "2." or "3.")
                .Select(show => show.Text.Trim()).Distinct()));
            foreach (string prefix in new[] { "Summary: ", "The two relationships ", "The trend marker " })
            {
                var prose = shows.Where(show => show.Text.StartsWith(prefix, StringComparison.Ordinal)).ToArray();
                TestAssert.True(prose.Length != 0 && prose.All(show => show.Y > 0d && show.Y < page.Height),
                    "Unavailable math cannot displace supported static prose off the page in either public preset: " + id);
            }
            if (row.GetProperty("SourceEquations").GetInt32() != 0)
            {
                TestAssert.True(diagnostics.Any(diagnostic => diagnostic.Id == "DOCX_UNSUPPORTED_EQUATION"),
                    "Missing mathematical italic glyphs retain explicit unsupported diagnostics: " + id);
                TestAssert.True(page.Images.Any(image => page.Content.Contains("/" + image.ResourceName + " Do", StringComparison.Ordinal)),
                    "Supported pictures remain painted beside unavailable math: " + id);
            }
            TestAssert.True(owners.Zip(paragraphs.SelectMany(paragraph => paragraph.Runs))
                .All(pair => ReferenceEquals(pair.First, pair.Second)), "Fallback placement retains original static source runs.");
        }
    }
}
