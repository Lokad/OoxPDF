using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionUndeclaredTypefacesMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent emitted faces use installed Word-compatible fonts.");
        CheckDefaultTypefaces(null);
    }

    internal static void CheckDefaultTypefaces(IFontResolver? resolver)
    {
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var golden = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "default-typefaces-office.json")));
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string path = Path.Combine(folder, row.GetProperty("File").GetString()!);
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var document = DocxTests.ReadDocx(path, mode);
            var paragraphs = DocxBlockTraversal.EnumerateSourceParagraphs(document).ToArray();
            var owners = paragraphs.SelectMany(p => p.Runs).ToArray();
            var page = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(document, null, default).Single();
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true, includeTextTransforms: true).ToArray();
            foreach (var expected in row.GetProperty("Anchors").EnumerateArray())
            {
                string text = expected.GetProperty("Text").GetString()!;
                var paints = shows.Where(show => show.Text.Trim() == text).ToArray();
                TestAssert.True(paints.Length != 0 && paints.All(show => show.Y > 0d && show.Y < page.Height),
                    "Every body/cell/static source anchor must survive fallback resolution: " + text);
                if (resolver is null)
                    foreach (var paint in paints)
                        TestAssert.Equal(expected.GetProperty("Family").GetString(), paint.Font.Font.FamilyName);
            }
            TestAssert.True(owners.Zip(paragraphs.SelectMany(p => p.Runs)).All(pair => ReferenceEquals(pair.First, pair.Second)),
                "Resolving undeclared faces cannot rewrite source formatting or owners.");
        }
    }
}
