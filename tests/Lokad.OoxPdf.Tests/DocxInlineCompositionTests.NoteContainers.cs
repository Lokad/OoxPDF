using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionOwnedNoteContainersKeepWordHeadsAndEmptyPageTakes()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent note-container origins require installed Windows fonts.");
        CheckOwnedNoteContainers(null);
    }

    internal static void CheckOwnedNoteContainers(IFontResolver? resolver)
    {
        using var gold = JsonDocument.Parse(File.ReadAllText(NestedMathCasePath("owned-note-containers-office.json")));
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
                var shows = pages.SelectMany((page, index) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                    .Select(show => (Page: index + 1, Show: show, IsMath: page.Fonts.Any(resource =>
                        resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(resource.Font, show.Font))))).ToArray();
                var math = shows.Where(item => item.IsMath).ToArray();
                TestAssert.Equal("ratetimeratetimeratetimeratetime", string.Concat(math.Select(item => item.Show.Text)).Normalize(NormalizationForm.FormKC));
                string prose = string.Concat(shows.Where(item => !item.IsMath).Select(item => item.Show.Text));
                TestAssert.Contains("Public calculation note.", prose);
                TestAssert.Contains("Final tail.", prose);
                TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Owned note containers retain all source equations.");
                TestAssert.Equal(1, pages.Sum(page => page.Annotations.Count(annotation => annotation.Destination is not null)));
                string origins = string.Join("|", math.Select(item => FormattableString.Invariant($"{item.Page}:{item.Show.X:R},{item.Show.Y:R}")));
                if (originalOrigins is not null) TestAssert.Equal(originalOrigins, origins);
                originalOrigins = origins;
                foreach (int pageNumber in Enumerable.Range(1, pages.Length))
                {
                    var pageShows = shows.Where(item => item.Page == pageNumber).ToArray();
                    var footer = pageShows.Where(item => item.Show.Text.StartsWith("Public review footer", StringComparison.Ordinal)).ToArray();
                    TestAssert.Equal(3, footer.Length);
                    double footerTop = Ink(footer[0].Show).Max();
                    double proseBottom = pageShows.Where(item => !item.IsMath && !item.Show.Text.StartsWith("Public review footer", StringComparison.Ordinal))
                        .SelectMany(item => Ink(item.Show)).Min();
                    TestAssert.True(proseBottom > footerTop, "Continued note prose clears actual footer ink on intermediate and final pages.");
                }
                if (resolver is not null) continue;
                TestAssert.Equal(row.GetProperty("Pages").GetInt32(), pages.Length);
                // The retained exact-height Review body-admission failure is recorded
                // separately in UnresolvedAnchors and in the raw Office comparison.
                foreach (var anchor in row.GetProperty("Anchors").EnumerateArray())
                {
                    string prefix = anchor.GetProperty("Prefix").GetString()!;
                    var actual = anchor.GetProperty("EndsWith").GetBoolean()
                        ? shows.Last(item => item.Show.Text.TrimEnd().EndsWith(prefix, StringComparison.Ordinal))
                        : shows.First(item => item.Show.Text.TrimStart().StartsWith(prefix[..Math.Min(8, prefix.Length)], StringComparison.Ordinal));
                    TestAssert.Equal(anchor.GetProperty("Page").GetInt32(), actual.Page);
                    TestAssert.True(Math.Abs(actual.Show.Y - anchor.GetProperty("Y").GetDouble()) < 1d,
                        "Qualified note heads, tails and body origins retain independent one-point Word gates.");
                }
                if (!transformed && !row.GetProperty("Page2RowsUnresolved").GetBoolean())
                {
                    var rows = shows.Where(item => item.Page == 2 && !string.IsNullOrWhiteSpace(item.Show.Text) &&
                        !item.Show.Text.StartsWith("Public review footer", StringComparison.Ordinal))
                        .Select(item => item.Show.Y).Distinct().Count();
                    TestAssert.Equal(row.GetProperty("Page2NoteRows").GetInt32(), rows);
                }
                foreach (var anchor in row.GetProperty("FooterAnchors").EnumerateArray())
                {
                    var actual = shows.Single(item => item.Page == anchor.GetProperty("Page").GetInt32() && item.Show.Text == anchor.GetProperty("Text").GetString());
                    TestAssert.True(Math.Abs(actual.Show.Y - anchor.GetProperty("Y").GetDouble()) < 1d, "Static footer placement retains Word's origin on every note page.");
                }
                TestAssert.Equal(32, math.Length);
                foreach (var anchor in row.GetProperty("MathAnchors").EnumerateArray())
                {
                    var actual = math[anchor.GetProperty("Offset").GetInt32()];
                    TestAssert.Equal(anchor.GetProperty("Page").GetInt32(), actual.Page);
                    TestAssert.True(Math.Abs(actual.Show.X - anchor.GetProperty("X").GetDouble()) < 1d &&
                        Math.Abs(actual.Show.Y - anchor.GetProperty("Y").GetDouble()) < 1d, "First glyphs of Word's grouped fraction words retain independent origins.");
                }
            }
        }

        static IEnumerable<double> Ink((PdfEmbeddedFont Font, string Text, double Size, double X, double Y) show)
        {
            var font = show.Font.Font;
            foreach (var rune in show.Text.EnumerateRunes())
            {
                if (Rune.IsWhiteSpace(rune)) continue;
                TestAssert.True(font.TryGetGlyphBounds(font.MapCodePoint(rune.Value), out var bounds), "The resolved prose glyph has measurable ink bounds.");
                yield return show.Y + bounds.YMin * show.Size / font.UnitsPerEm;
                yield return show.Y + bounds.YMax * show.Size / font.UnitsPerEm;
            }
        }
    }
}
