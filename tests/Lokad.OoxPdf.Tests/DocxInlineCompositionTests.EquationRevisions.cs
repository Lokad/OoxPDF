using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionEquationRevisionsRequireActualSourceOwners()
    {
        var source = ReadMutation("CompositionProbes/equation-revision-body-ins", OoxPdfDocxMarkupMode.Final, false, false);
        var paragraph = source.Paragraphs.First(p => p.Equations.Any(e => e.Revision is not null));
        var equation = paragraph.Equations.Single();
        TestAssert.True(paragraph.InlineContent is not null && DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph, default),
            "The actual equation wrapper owns its revision independently of prose runs.");
        TestAssert.Equal(1, DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
        foreach (var invalid in new[] { equation with { Revision = equation.Revision! with { } },
            equation with { SourceRunIndex = -1 }, equation with { Revision = equation.Revision! with { Kind = DocxRevisionKind.MoveTo } } })
            TestAssert.True(!DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph with { Equations = [invalid] }, default),
                "Value-equal metadata, unknown kinds and wrong source positions cannot invent equation ownership.");
        TestAssert.True(!DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph with { Equations = [] }, default),
            "Aggregate paragraph metadata cannot replace the removed source equation.");
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph, canceled.Token));
        foreach (bool before in new[] { true, false })
        {
            var ordered = ReadMutation("CompositionProbes/equation-revision-body-ins", OoxPdfDocxMarkupMode.AllMarkup, false, false,
                (xml, w, m) =>
                {
                    var wrapper = xml.Descendants(w + "ins").Single();
                    var text = new XElement(w + "ins", new XAttribute(w + "id", "872"), new XAttribute(w + "author", "Other Reviewer"),
                        new XElement(w + "r", new XElement(w + "t", "Other inserted text.")));
                    if (before) wrapper.AddBeforeSelf(text); else wrapper.AddAfterSelf(text);
                });
            var slots = DocxRenderer.BuildWordReviewAuthorPalette(ordered, [], default).AuthorSlots;
            TestAssert.Equal(before ? 1 : 0, slots["Equation Reviewer"]);
            TestAssert.Equal(before ? 0 : 1, slots["Other Reviewer"]);
        }
    }

    public static void DocxInlineCompositionEquationRevisionsMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent equation frames use installed Windows math/text fonts.");
        CheckEquationRevisions(null, true);
    }

    internal static void CheckEquationRevisions(IFontResolver? resolver, bool compareGeometry,
        string referenceFile = "equation-revisions-office.json", bool splitMathRuns = false)
    {
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var golden = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, referenceFile)));
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            string fixture = row.TryGetProperty("Fixture", out var authored) ? authored.GetString()! : "CompositionProbes/" + id;
            string input = Path.GetFullPath(Path.Combine(folder, "..", fixture + ".docx"));
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            var mode = row.GetProperty("Mode").GetString() == "final-preserve" ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            string? part = id.Contains("header", StringComparison.Ordinal) ? "word/header1.xml" :
                id.Contains("footer", StringComparison.Ordinal) ? "word/footer1.xml" : null;
            var anchors = row.GetProperty("Anchors").EnumerateArray().ToArray();
            int expectedGlyphs = row.TryGetProperty("MathGlyphCount", out var count) ? count.GetInt32() : anchors.Length;
            var inkAnchors = new JsonElement[expectedGlyphs];
            for (int a = 0; a < anchors.Length; a++)
            {
                int start = anchors[a].TryGetProperty("Offset", out var offset) ? offset.GetInt32() : a;
                int end = a + 1 < anchors.Length ? (anchors[a + 1].TryGetProperty("Offset", out var next) ? next.GetInt32() : a + 1) : expectedGlyphs;
                for (int i = start; i < end; i++) inkAnchors[i] = anchors[a];
            }
            (double X, double Y, double Size)[]? original = null;
            foreach (bool transformed in new[] { false, true })
            {
                var source = ReadMutation(fixture, mode, transformed, transformed,
                    !splitMathRuns || !transformed ? null : (xml, w, m) =>
                    {
                        var run = xml.Descendants(m + "r").FirstOrDefault(node => node.Element(m + "t")?.Value.Length > 1);
                        if (run is null) return;
                        var text = run.Element(m + "t")!;
                        var second = new XElement(run);
                        second.Element(m + "t")!.Value = text.Value[1..];
                        text.Value = text.Value[..1];
                        run.AddAfterSelf(second);
                    }, partOverride: part);
                var page = new DocxRenderer(resolver, mode, mode == OoxPdfDocxMarkupMode.Final
                    ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                    .RenderBlankPages(source, null, default).Single();
                var fonts = page.Fonts.Where(f => f.ResourceName.StartsWith("Eq", StringComparison.Ordinal)).Select(f => f.Font).ToHashSet();
                var math = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).Where(show => fonts.Contains(show.Font)).ToArray();
                TestAssert.Equal(expectedGlyphs, math.Length);
                TestAssert.Equal(row.GetProperty("MathText").GetString(), string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
                TestAssert.Equal(row.TryGetProperty("ImageCount", out var images) ? images.GetInt32() : 1, page.Images.Count);
                foreach (string prose in row.TryGetProperty("ExpectedText", out var expectedText)
                    ? expectedText.EnumerateArray().Select(value => value.GetString()!) : ["share the same input series."])
                    TestAssert.Contains(prose, DocxPageFootnoteNumberingTests.DecodeText(page));
                if (compareGeometry)
                    for (int a = 0; a < anchors.Length; a++)
                    {
                        int index = anchors[a].TryGetProperty("Offset", out var offset) ? offset.GetInt32() : a;
                        TestAssert.True(Math.Abs(math[index].X - anchors[a].GetProperty("X").GetDouble()) < 1d &&
                            Math.Abs(math[index].Y - anchors[a].GetProperty("Y").GetDouble()) < 1d,
                            "Revised equations retain independent Word container geometry: " + id + "/" + mode + "/" + index);
                    }
                for (int i = 0; i < math.Length; i++)
                {
                    if (original is not null)
                        TestAssert.True(Math.Abs(math[i].X - original[i].X) < .01d && Math.Abs(math[i].Y - original[i].Y) < .01d && math[i].Size == original[i].Size,
                            "Bookmark/run splitting cannot remove or relocate a supported revised equation.");
                }
                original = math.Select(show => (show.X, show.Y, show.Size)).ToArray();
                var names = page.Fonts.Where(f => f.ResourceName.StartsWith("Eq", StringComparison.Ordinal)).Select(f => f.ResourceName).ToHashSet();
                var blocks = Regex.Matches(page.Content, @"(?s)BT\s+(?<body>.*?)\s+ET").Cast<Match>()
                    .Where(block => names.Contains(Regex.Match(block.Value, @"/(?<name>\S+) [\d.]+ Tf").Groups["name"].Value)).ToArray();
                TestAssert.Equal(math.Length, blocks.Length);
                for (int i = 0; i < blocks.Length; i++)
                {
                    var rgb = Regex.Match(blocks[i].Value, @"(?<r>[\d.]+) (?<g>[\d.]+) (?<b>[\d.]+) rg");
                    var gray = Regex.Match(blocks[i].Value, @"(?<v>[\d.]+) g");
                    TestAssert.True(rgb.Success || gray.Success, "Each painted equation glyph has explicit ink.");
                    string expected = inkAnchors[i].GetProperty("FillColor").GetString()!;
                    double[] target = expected.StartsWith("rg:", StringComparison.Ordinal)
                        ? expected[3..].Split(',').Select(v => double.Parse(v, CultureInfo.InvariantCulture)).ToArray() : [0d, 0d, 0d];
                    for (int channel = 0; channel < 3; channel++)
                        TestAssert.True(Math.Abs(double.Parse(rgb.Success ? rgb.Groups[new[] { "r", "g", "b" }[channel]].Value : gray.Groups["v"].Value, CultureInfo.InvariantCulture) - target[channel]) < .002d,
                            "Actual equation glyph ink matches the independent Word author owner: " + id + "/" + mode);
                }
            }
        }
    }
}
