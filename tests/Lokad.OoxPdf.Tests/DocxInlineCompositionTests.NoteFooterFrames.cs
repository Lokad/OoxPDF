using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionNoteFooterFramesKeepWordOriginsAcrossEmptyContinuations()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent note/footer origins require installed Windows fonts.");
        CheckNoteFooterFrames(null);
    }

    internal static void CheckNoteFooterFrames(IFontResolver? resolver)
    {
        using var gold = JsonDocument.Parse(File.ReadAllText(NestedMathCasePath("note-footer-frames-office.json")));
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
                var document = ReadMutation("CompositionProbes/" + id, mode, transformed, transformed, partOverride: "word/footnotes.xml");
                var diagnostics = new List<OoxPdfDiagnostic>();
                var pages = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(document, diagnostics.Add, default).ToArray();
                TestAssert.Equal(row.GetProperty("Pages").GetInt32(), pages.Length);
                var shows = pages.SelectMany((page, index) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                    .Select(show => (Page: index + 1, Show: show, IsMath: page.Fonts.Any(resource =>
                        resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(resource.Font, show.Font))))).ToArray();
                var math = shows.Where(item => item.IsMath).ToArray();
                TestAssert.Equal("ratetimeratetimeratetimeratetime", string.Concat(math.Select(item => item.Show.Text)).Normalize(NormalizationForm.FormKC));
                string prose = string.Concat(shows.Where(item => !item.IsMath).Select(item => item.Show.Text));
                TestAssert.Contains("Public calculation note.", prose);
                TestAssert.Contains("End of supported note.", prose);
                TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Continuation containers preserve all supported source equations.");
                TestAssert.Equal(1, pages.Sum(page => page.Annotations.Count(annotation => annotation.Destination is not null)));
                foreach (int pageNumber in Enumerable.Range(1, pages.Length))
                {
                    var footer = shows.Where(item => item.Page == pageNumber && item.Show.Text.StartsWith("Public review footer", StringComparison.Ordinal)).ToArray();
                    TestAssert.Equal(3, footer.Length);
                    TestAssert.True(footer.All(item => item.Show.Y > 0d && item.Show.Y < pages[pageNumber - 1].Height),
                        "Every footer row remains visible on body and dedicated note continuation pages.");
                }
                string origins = string.Join("|", math.Select(item => FormattableString.Invariant($"{item.Page}:{item.Show.X:R},{item.Show.Y:R}")));
                if (originalOrigins is not null) TestAssert.Equal(originalOrigins, origins);
                originalOrigins = origins;
                if (resolver is not null) continue;
                foreach (var anchor in row.GetProperty("Anchors").EnumerateArray())
                {
                    string prefix = anchor.GetProperty("Prefix").GetString()![..8];
                    var actual = shows.First(item => item.Show.Text.TrimStart().StartsWith(prefix, StringComparison.Ordinal));
                    TestAssert.Equal(anchor.GetProperty("Page").GetInt32(), actual.Page);
                    TestAssert.True(Math.Abs(actual.Show.Y - anchor.GetProperty("Y").GetDouble()) < 1d,
                        "Notes share Word's bottom-anchored continuation frame and preserve body admission.");
                }
                foreach (var anchor in row.GetProperty("FooterAnchors").EnumerateArray())
                {
                    var actual = shows.Single(item => item.Page == anchor.GetProperty("Page").GetInt32() && item.Show.Text == anchor.GetProperty("Text").GetString());
                    TestAssert.True(Math.Abs(actual.Show.Y - anchor.GetProperty("Y").GetDouble()) < 1d,
                        "Static footer and note placement retain their distinct independent Word origins.");
                }
                TestAssert.Equal(32, math.Length);
                foreach (var anchor in row.GetProperty("MathAnchors").EnumerateArray())
                {
                    var actual = math[anchor.GetProperty("Offset").GetInt32()];
                    TestAssert.Equal(anchor.GetProperty("Page").GetInt32(), actual.Page);
                    TestAssert.True(Math.Abs(actual.Show.X - anchor.GetProperty("X").GetDouble()) < 1d &&
                        Math.Abs(actual.Show.Y - anchor.GetProperty("Y").GetDouble()) < 1d,
                        "First glyphs of grouped Word fraction words retain their source page and one-point origins.");
                }
            }
        }
    }
}
