using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionOccupiedFooterKeepsNotesAtIndependentWordOrigins()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent footer/note origins require installed Windows fonts.");
        CheckOccupiedFooterNotes(null);
    }

    internal static void CheckOccupiedFooterNotes(IFontResolver? resolver)
    {
        using var gold = JsonDocument.Parse(File.ReadAllText(NestedMathCasePath("linked-inline-stories-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(),
                Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(NestedMathCasePath(id + ".docx")))));
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            foreach (bool transformed in new[] { false, true })
            {
                var document = ReadMutation("CompositionProbes/" + id, mode, transformed, transformed,
                    partOverride: id.Contains("footer", StringComparison.Ordinal) ? "word/footer1.xml" : "word/footnotes.xml");
                var diagnostics = new List<Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic>();
                var pages = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(document, diagnostics.Add, default).ToArray();
                TestAssert.True(pages.Length > 0, "The complete source report remains renderable.");
                TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Footer availability must retain every supported formula and source story.");
                var showsByPage = pages.Select(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).ToArray()).ToArray();
                var math = showsByPage.SelectMany((shows, index) => shows
                    .Where(show => pages[index].Fonts.Any(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal) &&
                        ReferenceEquals(resource.Font, show.Font)))).ToArray();
                // Body22/header5 glyphs plus the added16-glyph formula; static
                // stories repeat on every page, including portable-font reflow.
                TestAssert.Equal(22 + 5 * pages.Length + (id.Contains("footer", StringComparison.Ordinal) ? 16 * pages.Length : 16), math.Length);
                int notePage = Array.FindIndex(showsByPage, shows => shows.Any(show => show.Text.StartsWith("All quantities", StringComparison.Ordinal)));
                TestAssert.True(notePage >= 0, "The original source note remains visible.");
                var shows = showsByPage[notePage];
                var lines = shows.GroupBy(show => show.Y).Select(group => group.OrderBy(show => show.X).ToArray()).ToArray();
                var note = lines.First(line => string.Concat(line.Select(show => show.Text)).StartsWith("All quantities", StringComparison.Ordinal));
                var footer = lines.First(line => string.Concat(line.Select(show => show.Text)).StartsWith("Public synthetic composition portfolio", StringComparison.Ordinal));
                double scale = final ? 1d : DocxRenderer.ResolveWordCompatiblePrintScale(document,
                    DocxMarkupContext.FromMode(mode, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup));
                TestAssert.True(note.SelectMany(show => Ink(show, scale)).Min() > footer.SelectMany(show => Ink(show, scale)).Max(),
                    "Actual note and footer prose ink must remain disjoint after physical placement, with native or portable fonts.");
                if (resolver is not null) continue;
                TestAssert.Equal(2, pages.Length);
                TestAssert.Equal(1, notePage);
                TestAssert.Equal(row.GetProperty("MathGlyphCount").GetInt32(), math.Length);
                var rule = Regex.Matches(pages[notePage].Content,
                    @"(?m)^0 g\r?\n(?<x>[-\d.]+) (?<y>[-\d.]+) (?<w>[-\d.]+) (?<h>[-\d.]+) re f")
                    .Cast<Match>().Select(match => (X: Number(match, "x"), Y: Number(match, "y"),
                        Width: Number(match, "w"), Height: Number(match, "h")))
                    .Single(rect => rect.Width > 50d && rect.Height < 2d);
                var expectedRule = row.GetProperty("SeparatorRule");
                TestAssert.True(Math.Abs(rule.X - expectedRule.GetProperty("X").GetDouble()) < 1d &&
                    Math.Abs(rule.Y - expectedRule.GetProperty("Y").GetDouble()) < 1d &&
                    Math.Abs(rule.Width - expectedRule.GetProperty("Width").GetDouble()) < 1d &&
                    Math.Abs(rule.Height - expectedRule.GetProperty("Height").GetDouble()) < 1d,
                    "The separator rule shares its measured placement frame without changing typography eligibility.");
                TestAssert.True(Math.Abs(note[0].Y - row.GetProperty("NoteY").GetDouble()) < 1d &&
                    Math.Abs(footer[0].Y - row.GetProperty("FooterY").GetDouble()) < 1d,
                    "Note seating and footer painting retain independent one-point Word origins in both presets.");
                var destination = pages[1].Annotations.Single(link => link.Destination?.PageIndex == 1).Destination!.Value;
                TestAssert.True(destination.Top is { } targetTop && Math.Abs(targetTop - row.GetProperty("NoteDestinationY").GetDouble()) < 1d,
                    "The automatic note target follows the same corrected placed note frame.");
            }
        }

        static double Number(Match match, string group) => double.Parse(match.Groups[group].Value, System.Globalization.CultureInfo.InvariantCulture);

        static IEnumerable<double> Ink((PdfEmbeddedFont Font, string Text, double Size, double X, double Y) show, double scale)
        {
            var font = show.Font.Font;
            foreach (var rune in show.Text.EnumerateRunes())
            {
                if (Rune.IsWhiteSpace(rune)) continue;
                TestAssert.True(font.TryGetGlyphBounds(font.MapCodePoint(rune.Value), out var bounds), "Resolved source prose has measurable ink bounds.");
                yield return show.Y + bounds.YMin * show.Size * scale / font.UnitsPerEm;
                yield return show.Y + bounds.YMax * show.Size * scale / font.UnitsPerEm;
            }
        }
    }
}
