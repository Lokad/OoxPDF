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
    public static void DocxInlineCompositionOwnedNoteContainersKeepWordHeadsAndEmptyPageTakes()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent note-container origins require installed Windows fonts.");
        CheckOwnedNoteContainers(null);
    }

    internal static void CheckOwnedNoteContainers(IFontResolver? resolver)
        => CheckNoteContinuationFrames("owned-note-containers-office.json", resolver, checkSeparatorRules: true);

    public static void DocxInlineCompositionNoteFooterFramesKeepWordOriginsAcrossEmptyContinuations()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent note/footer origins require installed Windows fonts.");
        CheckNoteFooterFrames(null);
    }

    internal static void CheckNoteFooterFrames(IFontResolver? resolver)
        => CheckNoteContinuationFrames("note-footer-frames-office.json", resolver, checkSeparatorRules: false);

    private static void CheckNoteContinuationFrames(string oracle, IFontResolver? resolver, bool checkSeparatorRules)
    {
        using var gold = JsonDocument.Parse(File.ReadAllText(NestedMathCasePath(oracle)));
        var inputHashes = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            if (!inputHashes.TryGetValue(id, out string? hash))
            {
                using var input = File.OpenRead(NestedMathCasePath(id + ".docx"));
                inputHashes.Add(id, hash = Convert.ToHexString(SHA256.HashData(input)));
            }
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), hash);
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var geometry = final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
            // Exercise each source mutation on one representative row from each oracle.
            bool separateTransforms = checkSeparatorRules ? id == "owned-note-container-intermediate" && final :
                id == "note-footer-frame-empty-tail" && !final;
            (bool Bookmark, bool Split)[] variants = separateTransforms
                ? [(false, false), (true, false), (false, true), (true, true)]
                : [(false, false), (true, true)];
            string? originalOrigins = null;
            foreach ((bool bookmark, bool split) in variants)
            {
                var document = ReadMutation("CompositionProbes/" + id, mode, bookmark, split, partOverride: "word/footnotes.xml");
                var diagnostics = new List<OoxPdfDiagnostic>();
                var pages = new DocxRenderer(resolver, mode, geometry)
                    .RenderPagesAsync(document, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
                if (!checkSeparatorRules || resolver is null) TestAssert.Equal(row.GetProperty("Pages").GetInt32(), pages.Length);
                var shows = pages.SelectMany((page, index) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                    .Select(show => (Page: index + 1, Show: show, IsMath: page.Fonts.Any(resource =>
                        resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(resource.Font, show.Font))))).ToArray();
                var math = shows.Where(item => item.IsMath).ToArray();
                TestAssert.Equal("ratetimeratetimeratetimeratetime", string.Concat(math.Select(item => item.Show.Text)).Normalize(NormalizationForm.FormKC));
                string prose = string.Concat(shows.Where(item => !item.IsMath).Select(item => item.Show.Text));
                TestAssert.Contains("Public calculation note.", prose);
                TestAssert.Contains(checkSeparatorRules ? "Final tail." : "End of supported note.", prose);
                TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Owned note containers retain all source equations.");
                TestAssert.Equal(1, pages.Sum(page => page.Annotations.Count(annotation => annotation.Destination is not null)));
                string origins = string.Join("|", math.Select(item => FormattableString.Invariant($"{item.Page}:{item.Show.X:R},{item.Show.Y:R}")));
                originalOrigins ??= origins;
                TestAssert.Equal(originalOrigins, origins);
                foreach (int pageNumber in Enumerable.Range(1, pages.Length))
                {
                    var pageShows = shows.Where(item => item.Page == pageNumber).ToArray();
                    var footer = pageShows.Where(item => item.Show.Text.StartsWith("Public review footer", StringComparison.Ordinal)).ToArray();
                    TestAssert.Equal(3, footer.Length);
                    if (checkSeparatorRules)
                    {
                        double footerTop = Ink(footer[0].Show).Max();
                        double proseBottom = pageShows.Where(item => !item.IsMath && !item.Show.Text.StartsWith("Public review footer", StringComparison.Ordinal))
                            .SelectMany(item => Ink(item.Show)).Min();
                        TestAssert.True(proseBottom > footerTop, "Continued note prose clears actual footer ink on intermediate and final pages.");
                    }
                    else
                    {
                        TestAssert.True(footer.All(item => item.Show.Y > 0d && item.Show.Y < pages[pageNumber - 1].Height),
                            "Every footer row remains visible on body and dedicated note continuation pages.");
                    }
                }
                if (checkSeparatorRules)
                    CheckNoteContinuationSeparators(row, document, pages, mode, geometry, resolver is null);
                if (resolver is not null) continue;
                // Body admission and continued notes use the same container boundary;
                // every pinned Word body/tail page anchor now participates.
                foreach (var anchor in row.GetProperty("Anchors").EnumerateArray())
                {
                    string prefix = anchor.GetProperty("Prefix").GetString()!;
                    var actual = anchor.TryGetProperty("EndsWith", out var endsWith) && endsWith.GetBoolean()
                        ? shows.Last(item => item.Show.Text.TrimEnd().EndsWith(prefix, StringComparison.Ordinal))
                        : shows.First(item => item.Show.Text.TrimStart().StartsWith(prefix[..Math.Min(8, prefix.Length)], StringComparison.Ordinal));
                    TestAssert.Equal(anchor.GetProperty("Page").GetInt32(), actual.Page);
                    TestAssert.True(Math.Abs(actual.Show.Y - anchor.GetProperty("Y").GetDouble()) < 1d,
                        "Qualified note heads, tails and body origins retain independent one-point Word gates.");
                }
                if (checkSeparatorRules && !bookmark && !split)
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

    private static void CheckNoteContinuationSeparators(JsonElement row, DocxDocument document, PdfPage[] pages,
        OoxPdfDocxMarkupMode mode, OoxPdfDocxMarkupGeometryMode geometry, bool checkWord)
    {
        var rules = pages.SelectMany((page, index) => Regex.Matches(page.Content,
            @"(?m)^0 g\r?\n(?<x>[-\d.]+) (?<y>[-\d.]+) (?<w>[-\d.]+) (?<h>[-\d.]+) re f")
            .Cast<Match>().Select(match => (Page: index + 1, X: Number(match, "x"), Y: Number(match, "y"),
                Width: Number(match, "w"), Height: Number(match, "h"))))
            .Where(rule => rule.Height < 2d && rule.Width > 50d).ToArray();
        TestAssert.True(rules.Length >= 2, "A continued source note retains its head and continuation rules.");
        double scale = mode == OoxPdfDocxMarkupMode.Final ? 1d : DocxRenderer.ResolveWordCompatiblePrintScale(document,
            DocxMarkupContext.FromMode(mode, geometry));
        double designWidth = document.PageWidthPoints - document.MarginLeftPoints - document.MarginRightPoints;
        TestAssert.True(Math.Abs(rules[0].Width - 144d * scale) < .002d, "The default head retains its qualified short rule.");
        TestAssert.True(rules.Skip(1).All(rule => Math.Abs(rule.Width - designWidth * scale) < .002d),
            "Omitted continuation stories use the owning section's full design width mapped once.");
        if (!checkWord) return;
        var expectedRules = row.GetProperty("SeparatorRules").EnumerateArray().ToArray();
        TestAssert.Equal(expectedRules.Length, rules.Length);
        foreach (var expectedRule in expectedRules)
        {
            var actual = rules.Single(rule => rule.Page == expectedRule.GetProperty("Page").GetInt32());
            TestAssert.True(Math.Abs(actual.X - expectedRule.GetProperty("X").GetDouble()) < 1d &&
                Math.Abs(actual.Y - expectedRule.GetProperty("Y").GetDouble()) < 1d &&
                Math.Abs(actual.Width - expectedRule.GetProperty("Width").GetDouble()) < 1d &&
                Math.Abs(actual.Height - expectedRule.GetProperty("Height").GetDouble()) < 1d,
                "Head and default continuation rule geometry retain the independent one-point Word gate.");
        }

        static double Number(Match match, string name) => double.Parse(match.Groups[name].Value, System.Globalization.CultureInfo.InvariantCulture);
    }
}
