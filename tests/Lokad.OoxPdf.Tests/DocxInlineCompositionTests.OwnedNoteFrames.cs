using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionOwnedNoteFramesMatchWordContinuationAndKeepBodyOrigins()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent note origins require installed Windows fonts.");
        CheckOwnedNoteFrames(null);
    }

    internal static void CheckOwnedNoteFrames(IFontResolver? resolver)
    {
        using var gold = JsonDocument.Parse(File.ReadAllText(NestedMathCasePath("owned-note-frames-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(),
                Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(NestedMathCasePath(id + ".docx")))));
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            string? originalOrigins = null;
            foreach (bool transformed in new[] { false, true })
            foreach (bool separator in new[] { false, true })
            {
                var document = ReadMutation("CompositionProbes/" + id, mode, transformed, transformed,
                    (xml, w, _) =>
                    {
                        if (separator) xml.Root!.AddFirst(new XElement(w + "footnote", new XAttribute(w + "id", "-1"),
                            new XAttribute(w + "type", "separator"), new XElement(w + "p",
                                new XElement(w + "r", new XElement(w + "separator")))));
                    }, "word/footnotes.xml");
                var diagnostics = new List<OoxPdfDiagnostic>();
                var pages = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(document, diagnostics.Add, default).ToArray();
                var shows = pages.SelectMany((page, index) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                    .Select(show => (Page: index + 1, Show: show, IsMath: page.Fonts.Any(resource =>
                        resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(resource.Font, show.Font))))).ToArray();
                var math = shows.Where(item => item.IsMath).ToArray();
                TestAssert.Equal("ratetimeratetimeratetimeratetime", string.Concat(math.Select(item => item.Show.Text))
                    .Normalize(NormalizationForm.FormKC));
                string prose = string.Concat(shows.Where(item => !item.IsMath).Select(item => item.Show.Text));
                TestAssert.Contains("Public calculation note.", prose);
                TestAssert.Contains("End of supported note.", prose);
                TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Owned note continuations retain all supported equations.");
                TestAssert.Equal(1, pages.Sum(page => page.Annotations.Count(annotation => annotation.Destination is not null)));
                TestAssert.True(math.All(item => item.Show.Y > 0d && item.Show.Y < pages[item.Page - 1].Height),
                    "All note and body math remains within the page.");
                string origins = string.Join("|", math.Select(item => item.Page + ":" + item.Show.X.ToString("R", CultureInfo.InvariantCulture) +
                    "," + item.Show.Y.ToString("R", CultureInfo.InvariantCulture)));
                if (!separator)
                {
                    if (originalOrigins is not null) TestAssert.Equal(originalOrigins, origins);
                    originalOrigins = origins;
                }
                if (resolver is not null || separator) continue;
                TestAssert.Equal(row.GetProperty("Pages").GetInt32(), pages.Length);
                foreach (var anchor in row.GetProperty("ProseAnchors").EnumerateArray())
                {
                    // Neutral run splits can divide the prose anchor across PDF shows.
                    string prefix = anchor.GetProperty("Prefix").GetString()![..3];
                    var actual = shows.First(item => item.Show.Text.TrimStart().StartsWith(prefix, StringComparison.Ordinal));
                    TestAssert.Equal(anchor.GetProperty("Page").GetInt32(), actual.Page);
                    TestAssert.True(Math.Abs(actual.Show.Y - anchor.GetProperty("Y").GetDouble()) < 1d,
                        "The independent one-point Word gate covers both continued note and unchanged body baselines.");
                }
                TestAssert.Equal(32, math.Length);
                foreach (var anchor in row.GetProperty("MathAnchors").EnumerateArray())
                {
                    var actual = math[anchor.GetProperty("Offset").GetInt32()];
                    TestAssert.Equal(anchor.GetProperty("Page").GetInt32(), actual.Page);
                    TestAssert.True(Math.Abs(actual.Show.X - anchor.GetProperty("X").GetDouble()) < 1d &&
                        Math.Abs(actual.Show.Y - anchor.GetProperty("Y").GetDouble()) < 1d,
                        "First glyphs of Word's grouped fraction words retain their independent origins and pages.");
                }
            }
        }
    }
}
