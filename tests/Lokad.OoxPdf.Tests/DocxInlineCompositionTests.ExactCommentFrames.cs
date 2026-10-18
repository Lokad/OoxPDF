using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionExactCommentFramesMatchIndependentWordVerticalBounds()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent exact comment origins require installed Windows fonts.");
        CheckExactCommentFrames(null);
    }

    internal static void CheckExactCommentHeightChanges(IFontResolver resolver)
    {
        double? originalMarkerHeight = null;
        foreach (int twips in new[] { 1600, 2400 })
        {
            var document = ReadMutation("CompositionProbes/owned-note-container-intermediate", OoxPdfDocxMarkupMode.AllMarkup,
                false, false, (xml, w, _) =>
                {
                    var paragraph = xml.Descendants(w + "p").Single(p => p.Elements(w + "commentRangeStart").Any());
                    paragraph.Element(w + "pPr")!.Element(w + "spacing")!.SetAttributeValue(w + "line", twips);
                });
            var context = DocxMarkupContext.FromMode(OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            double scale = DocxRenderer.ResolveWordCompatiblePrintScale(document, context);
            var pages = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, null, default).ToArray();
            var selected = pages.SelectMany(page => Regex.Matches(page.Content,
                @"0\.973 0\.863 0\.867 rg\s+(?<x>[-\d.]+) (?<y>[-\d.]+) (?<w>[-\d.]+) (?<h>[-\d.]+) re f")
                .Cast<Match>().Select(match => (Page: page, X: Value(match, "x"), Width: Value(match, "w"), Height: Value(match, "h"))))
                .Single(fill => fill.X < 400d);
            TestAssert.True(Math.Abs(selected.Height - twips / 20d * scale) < .002d, "Selection consumes the changed authored box.");
            var verticals = Regex.Matches(selected.Page.Content,
                @"(?m)^(?<x>[-\d.]+) (?<y>[-\d.]+) m\s+(?<endX>[-\d.]+) (?<endY>[-\d.]+) l S")
                .Cast<Match>().Where(match => Math.Abs(Value(match, "x") - Value(match, "endX")) < .001d &&
                    Value(match, "x") >= selected.X && Value(match, "x") <= selected.X + selected.Width + 1d)
                .Select(match => Math.Abs(Value(match, "endY") - Value(match, "y"))).ToArray();
            TestAssert.True(verticals.Length >= 2, "The selected source retains both glyph brackets.");
            double markerHeight = verticals.Max();
            if (originalMarkerHeight is { } prior)
                TestAssert.True(Math.Abs(markerHeight - prior) < .002d, "Authored spacing changes cannot stretch glyph brackets.");
            originalMarkerHeight = markerHeight;
        }

        static double Value(Match match, string name) => double.Parse(match.Groups[name].Value, CultureInfo.InvariantCulture);
    }

    internal static void CheckExactCommentFrames(IFontResolver? resolver)
    {
        using var gold = JsonDocument.Parse(File.ReadAllText(NestedMathCasePath("exact-comment-frames-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(),
                Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(NestedMathCasePath(id + ".docx")))));
            string? originalOrigins = null;
            foreach (bool transformed in new[] { false, true })
            {
                var document = ReadMutation("CompositionProbes/" + id, OoxPdfDocxMarkupMode.AllMarkup, false, transformed,
                    (xml, w, _) =>
                    {
                        if (!transformed) return;
                        var xmlParagraph = xml.Descendants(w + "p").First(p => p.Elements(w + "commentRangeStart").Any());
                        var start = new XElement(w + "bookmarkStart", new XAttribute(w + "id", "7000"), new XAttribute(w + "name", "PublicCommentFrame"));
                        if (xmlParagraph.Element(w + "pPr") is { } properties) properties.AddAfterSelf(start);
                        else xmlParagraph.AddFirst(start);
                        xmlParagraph.Add(new XElement(w + "bookmarkEnd", new XAttribute(w + "id", "7000")));
                    });
                var context = DocxMarkupContext.FromMode(OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
                double scale = DocxRenderer.ResolveWordCompatiblePrintScale(document, context);
                var paragraph = document.Paragraphs.Single(p => p.CommentRanges.Count != 0);
                double sourceHeight = paragraph.LineSpacingPoints!.Value;
                var diagnostics = new List<OoxPdfDiagnostic>();
                var pages = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                    .RenderBlankPages(document, diagnostics.Add, default).ToArray();
                var fills = pages.SelectMany((page, index) => Regex.Matches(page.Content,
                    @"0\.973 0\.863 0\.867 rg\s+(?<x>[-\d.]+) (?<y>[-\d.]+) (?<w>[-\d.]+) (?<h>[-\d.]+) re f")
                    .Cast<Match>().Select(match => (Page: index + 1, X: Value(match, "x"), Bottom: Value(match, "y"), Height: Value(match, "h"))))
                    .Where(fill => fill.X < 400d).ToArray();
                var expected = row.GetProperty("SelectionBounds").EnumerateArray().ToArray();
                TestAssert.Equal(expected.Length, fills.Length);
                TestAssert.True(fills.All(fill => Math.Abs(fill.Height - sourceHeight * scale) < .002d),
                    "Authored exact boxes own selection expansion and compression independently of font ink.");
                var shows = pages.SelectMany((page, index) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                    .Select(show => (Page: index + 1, Show: show, IsMath: page.Fonts.Any(resource =>
                        resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(resource.Font, show.Font))))).ToArray();
                foreach (var text in row.GetProperty("ExpectedText").EnumerateArray())
                    TestAssert.Contains(text.GetString()!, string.Concat(shows.Select(item => item.Show.Text)));
                TestAssert.Equal(row.GetProperty("ExpectedMathGlyphs").GetInt32(), shows.Count(item => item.IsMath));
                TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Comment geometry preserves supported mathematical content.");
                var balloon = shows.Single(item => item.Show.Text.StartsWith("Commented [RV1]:", StringComparison.Ordinal));
                string origins = string.Join("|", fills.Select(fill => FormattableString.Invariant($"{fill.Page}:{fill.Bottom:R},{fill.Height:R}"))) +
                    FormattableString.Invariant($"|{balloon.Page}:{balloon.Show.Y:R}");
                if (originalOrigins is not null) TestAssert.Equal(originalOrigins, origins);
                originalOrigins = origins;
                if (resolver is not null) continue;
                for (int i = 0; i < fills.Length; i++)
                {
                    TestAssert.Equal(row.GetProperty("Page").GetInt32(), fills[i].Page);
                    TestAssert.True(Math.Abs(fills[i].Bottom - expected[i].GetProperty("Bottom").GetDouble()) < 1d &&
                        Math.Abs(fills[i].Bottom + fills[i].Height - expected[i].GetProperty("Top").GetDouble()) < 1d,
                        "Selection top and bottom retain independent one-point Word gates for compact and tall authored boxes.");
                }
                TestAssert.Equal(row.GetProperty("Page").GetInt32(), balloon.Page);
                TestAssert.True(Math.Abs(balloon.Show.Y - row.GetProperty("BalloonTitleY").GetDouble()) < 1d,
                    "The complete ending source run and its occupied top own the independent Word balloon origin.");
            }
        }

        static double Value(Match match, string name) => double.Parse(match.Groups[name].Value, CultureInfo.InvariantCulture);
    }
}
