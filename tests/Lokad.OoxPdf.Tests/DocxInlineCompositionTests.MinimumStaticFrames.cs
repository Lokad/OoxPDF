using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionMinimumStaticFramesMatchWordAndKeepBodyAdmission()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent static origins require installed Windows fonts.");
        CheckMinimumStaticFrames(null);
    }

    public static void DocxInlineCompositionTypographicMinimumStaticFramesKeepIndependentWordBaselines()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent typographic static origins require installed Windows fonts.");
        CheckMinimumStaticFrames(null, "typographic-static-frames-office.json");
    }

    internal static void CheckMinimumStaticFrames(IFontResolver? resolver, string referenceFile = "minimum-static-frames-office.json")
    {
        using var gold = JsonDocument.Parse(File.ReadAllText(NestedMathCasePath(referenceFile)));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(),
                Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(NestedMathCasePath(id + ".docx")))));
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            double[]? original = null;
            foreach (bool transformed in new[] { false, true })
            {
                var document = ReadMutation("CompositionProbes/" + id, mode, false, transformed,
                    (xml, w, _) =>
                    {
                        if (!transformed) return;
                        var paragraph = xml.Descendants(w + "p").First();
                        paragraph.AddFirst(new XElement(w + "bookmarkStart", new XAttribute(w + "id", "41"), new XAttribute(w + "name", "StaticSource")));
                        paragraph.Add(new XElement(w + "bookmarkEnd", new XAttribute(w + "id", "41")));
                    }, row.GetProperty("Part").GetString());
                var owners = document.HeaderParagraphsByType.Values.Concat(document.FooterParagraphsByType.Values)
                    .SelectMany(p => p).SelectMany(p => p.Runs).ToArray();
                var page = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(document, null, default).Single();
                var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).ToArray();
                string prose = string.Concat(shows.Select(s => s.Text));
                foreach (int line in new[] { 1, 2, 3 })
                    TestAssert.Contains("Public review " + (id.EndsWith("header", StringComparison.Ordinal) ? "header" : "footer") + " line " + line + ".", prose);
                var rows = shows.Where(show => show.Text.StartsWith("Public review", StringComparison.Ordinal)).OrderByDescending(show => show.Y).ToArray();
                TestAssert.Equal(3, rows.Length);
                TestAssert.True(rows.All(show => show.Y > 0d && show.Y < page.Height),
                    "Every supported minimum-height header/footer line must be visible on its page.");
                if (original is not null) TestAssert.True(original.SequenceEqual(rows.Select(s => s.Y)),
                    "Neutral source metadata and run splits preserve static baseline geometry.");
                original = rows.Select(s => s.Y).ToArray();
                TestAssert.True(owners.SequenceEqual(document.HeaderParagraphsByType.Values.Concat(document.FooterParagraphsByType.Values)
                    .SelectMany(p => p).SelectMany(p => p.Runs), ReferenceEqualityComparer.Instance), "Placement retains the original static source runs.");
                if (resolver is not null) continue;
                foreach (var pair in rows.Zip(row.GetProperty("StaticY").EnumerateArray()))
                    TestAssert.True(Math.Abs(pair.First.Y - pair.Second.GetDouble()) < 1d,
                        "Complete minimum-height static stories retain the independent one-point Word baseline gate.");
                foreach (var anchor in row.GetProperty("BodyAnchors").EnumerateArray())
                    TestAssert.True(Math.Abs(shows.First(s => s.Text.StartsWith(anchor.GetProperty("Prefix").GetString()!, StringComparison.Ordinal)).Y -
                        anchor.GetProperty("Y").GetDouble()) < 1d, "Static admission preserves independent body origins.");
            }
        }
    }
}
