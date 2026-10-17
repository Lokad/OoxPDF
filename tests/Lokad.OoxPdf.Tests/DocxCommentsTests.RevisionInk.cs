using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxCommentsTests
{
    public static void DocxReviewRevisionInkMatchesIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native controls use installed Windows fonts.");
        CheckReviewRevisionInk(null);
    }

    public static void DocxReviewRevisionInkFallbackKeepsSourcePalette() => CheckReviewRevisionInk(new MissingInkFontResolver());

    internal static void CheckReviewRevisionInk(IFontResolver? resolver)
    {
        using var golden = JsonDocument.Parse(File.ReadAllText(Path.Combine(NonTextAuthorCases(), "revision-text-office-ink.json")));
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray())
        {
            var source = ReadNonTextAuthorSource(row, OoxPdfDocxMarkupMode.AllMarkup);
            var page = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(source, null, default).Single();
            var ink = ReadNonTextAuthorRgb(row, "FillRgb");
            string expected = DocxTests.FormatPdfRgb(ink, "rg");
            string target = row.GetProperty("TargetWord").GetString()!;
            int found = 0;
            foreach (Match block in Regex.Matches(page.Content, @"(?s)BT\s+(?<body>.*?)\s+ET"))
            {
                string name = Regex.Match(block.Value, @"/(?<name>\S+) [\d.]+ Tf").Groups["name"].Value;
                if (page.Fonts.Any(resource => resource.ResourceName == name))
                {
                    var shows = ReadEmbeddedGlyphTextShows(new PdfPage(page.Width, page.Height, block.Value, page.Fonts))
                        .Where(show => show.X < 400d && show.Text.Contains(target, StringComparison.Ordinal)).ToArray();
                    foreach (var show in shows)
                    {
                        found++;
                        TestAssert.Contains(expected, block.Value);
                        string underline = Regex.Escape(DocxTests.FormatPdfRgb(ReadNonTextAuthorRgb(row, "DecorationFillRgb"), "rg")) +
                            @"\s+(?<x>[-\d.]+) (?<y>[-\d.]+) (?<w>[-\d.]+) (?<h>[-\d.]+) re f";
                        bool sameInk = Regex.Matches(page.Content, underline).Cast<Match>().Any(rect =>
                            Math.Abs(Number(rect, "x") - show.X) < .1d && Number(rect, "w") > 0d &&
                            Math.Abs(show.Y - Number(rect, "y")) < show.Size &&
                            (row.GetProperty("Decoration").GetString() == "Strike" ? Number(rect, "y") > show.Y : Number(rect, "y") < show.Y) &&
                            Number(rect, "h") > 0d && Number(rect, "h") < show.Size);
                        TestAssert.True(sameInk, "The source revision decoration must share its independently measured author ink: " + row.GetProperty("Id").GetString());
                    }
                }
                else if (page.FallbackFonts.Any(resource => resource.ResourceName == name))
                {
                    // The generated fallback positions one WinAnsi glyph per
                    // byte code. These source probes use ASCII words only.
                    string text = string.Concat(Regex.Matches(block.Value, @"<(?<byte>[0-9A-F]{2})> Tj").Cast<Match>()
                        .Select(glyph => (char)byte.Parse(glyph.Groups["byte"].Value, NumberStyles.HexNumber, CultureInfo.InvariantCulture)));
                    var matrix = Regex.Match(block.Value, @"1 0 0 1 (?<x>[-\d.]+) (?<y>[-\d.]+) Tm");
                    if (!text.Contains(target, StringComparison.Ordinal) || !matrix.Success || Number(matrix, "x") >= 400d) continue;
                    found++;
                    TestAssert.Contains(expected, block.Value);
                }
            }
            TestAssert.Equal(1, found);
        }
        static double Number(Match match, string group) => double.Parse(match.Groups[group].Value, CultureInfo.InvariantCulture);
    }

    private sealed class MissingInkFontResolver : IFontResolver
    {
        public FontFaceResolution Resolve(FontRequest request) => new(request.FamilyName, "Missing", new FontStyleKey(),
            new MemoryFontProgramSource("missing-ink", ReadOnlyMemory<byte>.Empty), IsFallback: false);
    }
}
