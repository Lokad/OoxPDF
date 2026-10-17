using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionEmptyFractionKeepsOnlyStructuralOwnership()
    {
        var source = ReadMutation("CompositionProbes/equation-revision-body-del", OoxPdfDocxMarkupMode.Final, false, false);
        var paragraph = source.Paragraphs.First();
        var equation = paragraph.Equations.Single();
        TestAssert.True(equation.Revision is null && paragraph.Revisions.Count == 0 &&
            equation.Content.Children.Single() is DocxMathFraction { Numerator.Children.Count: 0, Denominator.Children.Count: 0 },
            "Final keeps fraction structure without deleted glyphs or review metadata.");
        TestAssert.True(!DocxRenderer.BuildWordReviewAuthorPalette(source, [], default).AuthorSlots.ContainsKey("Equation Reviewer"),
            "An excluded glyph owner cannot register author ink in Final.");
        foreach (string kind in new[] { "noBar", "unknown" })
        {
            var unsupported = ReadMutation("CompositionProbes/equation-revision-body-del", OoxPdfDocxMarkupMode.Final, false, false,
                (xml, w, m) => xml.Descendants(w + "del").Single().Descendants(m + "f").Single()
                    .AddFirst(new XElement(m + "fPr", new XElement(m + "type", new XAttribute(m + "val", kind)))));
            TestAssert.Equal(0, unsupported.Paragraphs.First().Equations.Count);
        }
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.RemoveDeletedFractionText(equation, canceled.Token));
    }

    public static void DocxInlineCompositionEmptyFractionsMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent fraction frames use installed Windows fonts.");
        CheckEmptyFractions(null, true);
    }

    internal static void CheckEmptyFractions(IFontResolver? resolver, bool compareGeometry)
    {
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var golden = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "empty-fractions-office.json")));
        var contents = new Dictionary<string, string>();
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            string input = Path.Combine(folder, id + ".docx");
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            var page = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final), null, default).Single();
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).ToArray();
            var fonts = page.Fonts.Where(f => f.ResourceName.StartsWith("Eq", StringComparison.Ordinal)).Select(f => f.Font).ToHashSet();
            TestAssert.Equal(14, shows.Count(show => fonts.Contains(show.Font)));
            string prose = Regex.Replace(string.Join(" ", shows.Where(show => !fonts.Contains(show.Font)).Select(show => show.Text)), @"\s+", " ");
            TestAssert.True(prose.Contains("share the same input series.", StringComparison.Ordinal),
                "Wrapped prose remains present independently of mathematical text-show order: " + id);
            TestAssert.Equal(1, page.Images.Count);
            if (compareGeometry)
            {
                double first = shows.First(show => show.Text.StartsWith("Summary:", StringComparison.Ordinal)).Y;
                double following = shows.First(show => show.Text.StartsWith("The two relationships", StringComparison.Ordinal)).Y;
                TestAssert.True(Math.Abs(first - row.GetProperty("SummaryBaseline").GetDouble()) < 1d &&
                    Math.Abs(following - row.GetProperty("FollowingBaseline").GetDouble()) < 1d,
                    "Empty fraction control owns independent paragraph flow without its deleted text: " + id + ", " + first + ", " + following);
            }
            contents.Add(id, page.Content);
        }
        foreach (string suffix in new[] { "", "-9", "-18" })
        {
            string deleted = suffix.Length == 0 ? "equation-revision-body-del" : "final-equation-flow-deleted" + suffix;
            TestAssert.Equal(contents["final-equation-flow-empty-fraction" + suffix], contents[deleted]);
        }
        // A structural-only equation must not require a dummy glyph or PDF font subset.
        var structural = ReadMutation("CompositionProbes/final-equation-flow-empty-fraction", OoxPdfDocxMarkupMode.Final, false, false,
            (xml, w, m) =>
            {
                var body = xml.Root!.Element(w + "body")!;
                body.ReplaceNodes(new XElement(body.Element(w + "p")!), new XElement(body.Element(w + "sectPr")!));
            });
        var only = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(structural, null, default).Single();
        TestAssert.True(only.Fonts.All(font => !font.ResourceName.StartsWith("Eq", StringComparison.Ordinal)),
            "Metric-only fraction placeholders create no mathematical PDF font resource.");
        TestAssert.True(Regex.IsMatch(only.Content, @"[-\d.]+ [-\d.]+ [-\d.]+ [-\d.]+ re f"),
            "The structural fraction rule remains painted without a mathematical text show.");
    }
}
