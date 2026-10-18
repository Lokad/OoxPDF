using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionContinuedFooterNotesKeepWordBottomsAndSourceOrder()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent continuation origins require installed Windows fonts.");
        CheckContinuedFooterNotes(null);
    }

    internal static void CheckContinuedFooterNotes(IFontResolver? resolver)
    {
        using var gold = JsonDocument.Parse(File.ReadAllText(NestedMathCasePath("occupied-footer-continuations-office.json")));
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
                    partOverride: "word/footer1.xml");
                var diagnostics = new List<Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic>();
                var pages = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(document, diagnostics.Add, default).ToArray();
                TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Source formulas survive ordinary note continuation.");
                double scale = final ? 1d : DocxRenderer.ResolveWordCompatiblePrintScale(document,
                    DocxMarkupContext.FromMode(mode, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup));
                var seen = new List<int>();
                int mathCount = 0;
                for (int index = 0; index < pages.Length; index++)
                {
                    var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(pages[index], usePageCoordinates: true).ToArray();
                    mathCount += shows.Count(show => pages[index].Fonts.Any(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal) &&
                        ReferenceEquals(resource.Font, show.Font)));
                    var lines = shows.GroupBy(show => show.Y).OrderByDescending(group => group.Key)
                        .Select(group => group.OrderBy(show => show.X).ToArray()).ToArray();
                    string Text((Lokad.OoxPdf.Pdf.PdfEmbeddedFont Font, string Text, double Size, double X, double Y)[] line) => string.Concat(line.Select(show => show.Text));
                    var note = lines.Where(line => Text(line).StartsWith("All quantities", StringComparison.Ordinal) ||
                        Text(line).StartsWith("Public continuation note line", StringComparison.Ordinal)).ToArray();
                    foreach (Match match in Regex.Matches(string.Join(" ", lines.Select(Text)), @"Public\s+continuation\s+note\s+line\s+(\d\d):"))
                        seen.Add(int.Parse(match.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture));
                    if (note.Length == 0) continue;
                    var footer = lines.First(line => Text(line).StartsWith("Public synthetic composition portfolio", StringComparison.Ordinal));
                    TestAssert.True(note.SelectMany(line => line).SelectMany(show => OccupiedFooterInk(show, scale)).Min() >
                        footer.SelectMany(show => OccupiedFooterInk(show, scale)).Max(),
                        "Every source note page, including dedicated continuations, clears measured footer prose ink.");
                    if (resolver is not null) continue;
                    var expected = row.GetProperty("NotePages").EnumerateArray().Single(p => p.GetProperty("Page").GetInt32() == index + 1);
                    TestAssert.True(Math.Abs(note[^1][0].Y - expected.GetProperty("NoteY").GetDouble()) < 1d &&
                        Math.Abs(footer[0].Y - expected.GetProperty("FooterY").GetDouble()) < 1d,
                        "Last-note and footer baselines retain independent one-point Word bounds on every page.");
                }
                TestAssert.Equal(22 + 21 * pages.Length, mathCount);
                TestAssert.True(seen.SequenceEqual(id.Contains("long-note", StringComparison.Ordinal) ? Enumerable.Range(1, 64) : []),
                    "All64 numbered source paragraphs survive exactly once, in source order, across page turns.");
                if (resolver is null) TestAssert.Equal(row.GetProperty("Pages").GetInt32(), pages.Length);
            }
        }
    }
}
