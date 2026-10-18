using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionTableNoteReferencesKeepIndependentWordPages()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent table/reference pages require installed Windows fonts.");
        CheckTableNoteReferences(null);
    }

    internal static void CheckTableNoteReferences(IFontResolver? resolver)
    {
        using var gold = JsonDocument.Parse(File.ReadAllText(NestedMathCasePath("table-note-reference-office.json")));
        TestAssert.Equal(gold.RootElement.GetProperty("InputSha256").GetString(),
            Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(NestedMathCasePath("table-note-reference-report.docx")))));
        var expectedNumbers = gold.RootElement.GetProperty("ExpectedNumberOrder").EnumerateArray().Select(n => n.GetInt32()).ToArray();
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            foreach (bool transformed in new[] { false, true })
            {
                var document = ReadMutation("CompositionProbes/table-note-reference-report", mode, transformed, transformed,
                    partOverride: "word/document.xml");
                var diagnostics = new List<Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic>();
                var pages = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(document, diagnostics.Add, default).ToArray();
                TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Moving a note reference into a cell retains supported math.");
                double scale = final ? 1d : DocxRenderer.ResolveWordCompatiblePrintScale(document,
                    DocxMarkupContext.FromMode(mode, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup));
                var seen = new List<int>();
                int mathCount = 0, heads = 0;
                for (int index = 0; index < pages.Length; index++)
                {
                    var page = pages[index];
                    var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).ToArray();
                    mathCount += shows.Count(show => page.Fonts.Any(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal) &&
                        ReferenceEquals(resource.Font, show.Font)));
                    var lines = shows.GroupBy(show => show.Y).OrderByDescending(group => group.Key)
                        .Select(group => group.OrderBy(show => show.X).ToArray()).ToArray();
                    string[] text = lines.Select(line => string.Concat(line.Select(show => show.Text))).ToArray();
                    int priorCount = seen.Count;
                    foreach (Match match in Regex.Matches(string.Join(" ", text), @"Public\s+continuation\s+note\s+line\s+(\d\d):"))
                        seen.Add(int.Parse(match.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture));
                    int[] noteIndices = Enumerable.Range(0, lines.Length).Where(i => text[i].StartsWith("All quantities", StringComparison.Ordinal) ||
                        text[i].StartsWith("Second public continuation note.", StringComparison.Ordinal) ||
                        text[i].StartsWith("Public continuation note line", StringComparison.Ordinal)).ToArray();
                    if (noteIndices.Length == 0) continue;
                    if (index == 0)
                    {
                        var tableProse = Enumerable.Range(0, lines.Length).Where(i => text[i].Contains("Observed series", StringComparison.Ordinal) ||
                            text[i].Contains("Additional note", StringComparison.Ordinal)).SelectMany(i => lines[i]);
                        var tableMath = shows.Where(show => page.Fonts.Any(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal) &&
                            ReferenceEquals(resource.Font, show.Font)));
                        TestAssert.True(noteIndices.SelectMany(i => lines[i]).SelectMany(show => OccupiedFooterInk(show, scale)).Max() <
                            tableProse.Concat(tableMath).SelectMany(show => OccupiedFooterInk(show, scale)).Min(),
                            "The first note clears actual table prose and mathematical ink, including denominator descent.");
                    }
                    var footer = lines[Array.FindIndex(text, t => t.StartsWith("Public synthetic composition portfolio", StringComparison.Ordinal))];
                    TestAssert.True(noteIndices.SelectMany(i => lines[i]).SelectMany(show => OccupiedFooterInk(show, scale)).Min() >
                        footer.SelectMany(show => OccupiedFooterInk(show, scale)).Max(), "Table-owned notes clear the occupied footer.");
                    for (int n = 1; n < noteIndices.Length; n++)
                        TestAssert.True(lines[noteIndices[n]].SelectMany(show => OccupiedFooterInk(show, scale)).Max() <
                            lines[noteIndices[n - 1]].SelectMany(show => OccupiedFooterInk(show, scale)).Min(),
                            "Successive source note paragraphs occupy disjoint ink bands.");
                    foreach (var head in row.GetProperty("Heads").EnumerateArray())
                    {
                        int lineIndex = Array.FindIndex(text, t => t.StartsWith(head.GetProperty("Prefix").GetString()!, StringComparison.Ordinal));
                        if (lineIndex < 0) continue;
                        heads++;
                        var target = page.Annotations.Single(link => link.Destination?.PageIndex == index).Destination!.Value;
                        TestAssert.True(target.Top > lines[lineIndex][0].Y, "The automatic target lands above its own first note baseline.");
                        if (resolver is null)
                        {
                            TestAssert.Equal(head.GetProperty("Page").GetInt32(), index + 1);
                            TestAssert.True(Math.Abs(lines[lineIndex][0].Y - head.GetProperty("BaselineY").GetDouble()) < 1d,
                                "The note head retains its independent one-point Word baseline.");
                            TestAssert.True(target.Top is { } destination && Math.Abs(destination - head.GetProperty("DestinationY").GetDouble()) < 1d,
                                "The automatic target retains its independent one-point Word origin.");
                        }
                    }
                    if (resolver is null)
                    {
                        var expected = row.GetProperty("NotePages").EnumerateArray().Single(p => p.GetProperty("Page").GetInt32() == index + 1);
                        TestAssert.True(seen.Skip(priorCount).SequenceEqual(expected.GetProperty("Numbers").EnumerateArray().Select(n => n.GetInt32())),
                            "Every page retains Word's independent source-paragraph membership.");
                    }
                }
                TestAssert.Equal(2, heads);
                TestAssert.True(seen.SequenceEqual(expectedNumbers), "Every source paragraph survives once, following reference order rather than note XML order.");
                TestAssert.Equal(22 + 5 * pages.Length, mathCount);
                TestAssert.Equal(2, pages.Sum(page => page.Annotations.Count(link => link.Destination is not null)));
                if (resolver is null) TestAssert.Equal(row.GetProperty("Pages").GetInt32(), pages.Length);
                {
                    var firstPage = DocxCommentsTests.ReadEmbeddedGlyphTextShows(pages[0], usePageCoordinates: true).ToArray();
                    TestAssert.True(firstPage.GroupBy(show => show.Y).Any(line => string.Concat(line.OrderBy(show => show.X).Select(show => show.Text)).Contains("Observed series", StringComparison.Ordinal)),
                        "The reference-bearing table row stays on Word's first page instead of following isolated header pages.");
                }
            }
        }
    }
}
