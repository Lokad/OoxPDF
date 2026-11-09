using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxNaturalCommentFillsMatchIndependentWordSourceAndLineBounds()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent natural comment fill bounds use installed Windows fonts.");
        using var cases = ReadNaturalCommentFillJson("fills.json");
        using var gold = ReadNaturalCommentFillJson("fills-office.json");
        foreach (var expected in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            var source = cases.RootElement.EnumerateArray().Single(row => row.GetProperty("Id").GetString() == expected.GetProperty("Id").GetString());
            TestAssert.Equal(source.GetProperty("Sha256").GetString(), expected.GetProperty("InputSha256").GetString());
            var page = RenderContainerGrid(source, OoxPdfDocxMarkupMode.AllMarkup, null);
            var fill = ReadNaturalCommentFillRectangles(page.Content).Single(rect => rect.X < 400d && rect.Width > 5d);
            double[] actual = [fill.X, page.Height - fill.Y - fill.Height, fill.X + fill.Width, page.Height - fill.Y];
            for (int index = 0; index < actual.Length; index++)
                TestAssert.True(Math.Abs(actual[index] - expected.GetProperty("Fill")[index].GetDouble()) < 1d,
                    "Source advances and the painting line retain the independent Word fill edge: " + source.GetProperty("Id").GetString() + "/" + index);
        }
    }

    public static void DocxNaturalCommentFillsPreservePortableContentAndCancellation()
    {
        using var cases = ReadNaturalCommentFillJson("fills.json");
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        foreach (var source in cases.RootElement.EnumerateArray())
        {
            var page = RenderContainerGrid(source, OoxPdfDocxMarkupMode.AllMarkup, resolver);
            TestAssert.Equal(1, ReadNaturalCommentFillRectangles(page.Content).Count(rect => rect.X < 400d && rect.Width > 5d));
            string text = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true).Select(show => show.Text));
            foreach (var expected in source.GetProperty("ExpectedText").EnumerateArray())
                TestAssert.True(Normalize(text).Contains(Normalize(expected.GetString()!), StringComparison.Ordinal),
                    "Fill changes cannot omit source prose: " + source.GetProperty("Id").GetString());
        }
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => RenderContainerGrid(cases.RootElement[0], OoxPdfDocxMarkupMode.AllMarkup, resolver, canceled.Token));

        static string Normalize(string text) => Regex.Replace(text, @"\s+", "");
    }

    public static void DocxNaturalCommentFillsKeepRangesAcrossBookmarksAndRunSplits()
    {
        PdfRectangle? original = null;
        string? originalText = null;
        foreach ((bool bookmark, bool split) in new[] { (false, false), (true, false), (false, true), (true, true) })
        {
            var document = ReadMutation("NaturalCommentFills/arial12-auto", OoxPdfDocxMarkupMode.AllMarkup, false, split,
                (xml, w, _) =>
                {
                    if (!bookmark) return;
                    var paragraph = xml.Descendants(w + "p").Single(p => p.Elements(w + "commentRangeStart").Any());
                    var start = new XElement(w + "bookmarkStart", new XAttribute(w + "id", "7010"), new XAttribute(w + "name", "PublicCommentFill"));
                    paragraph.Element(w + "pPr")!.AddAfterSelf(start);
                    paragraph.Add(new XElement(w + "bookmarkEnd", new XAttribute(w + "id", "7010")));
                });
            var page = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), OoxPdfDocxMarkupMode.AllMarkup,
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderPagesAsync(document, null, default).ToBlockingEnumerable().Single();
            var fill = ReadNaturalCommentFillRectangles(page.Content).Single(rect => rect.X < 400d && rect.Width > 5d);
            string text = Regex.Replace(string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true).Select(show => show.Text)), @"\s+", "");
            if (original is { } prior)
            {
                TestAssert.True(Math.Abs(prior.X - fill.X) < 1d && Math.Abs(prior.Y - fill.Y) < 1d &&
                    Math.Abs(prior.Width - fill.Width) < 1d && Math.Abs(prior.Height - fill.Height) < 1d,
                    "Non-rendering source transformations retain the selected range's geometry.");
                TestAssert.Equal(originalText!, text);
            }
            original = fill;
            originalText = text;
        }
    }

    private static IEnumerable<PdfRectangle> ReadNaturalCommentFillRectangles(string content) => Regex.Matches(content,
        @"0\.973 0\.863 0\.867 rg\s+(?<x>[-\d.]+) (?<y>[-\d.]+) (?<w>[\d.]+) (?<h>[\d.]+) re f")
        .Select(match => new PdfRectangle(Value(match, "x"), Value(match, "y"), Value(match, "w"), Value(match, "h")));

    private static double Value(Match match, string name) => double.Parse(match.Groups[name].Value, CultureInfo.InvariantCulture);

    private static JsonDocument ReadNaturalCommentFillJson(string name) => JsonDocument.Parse(File.ReadAllText(Path.GetFullPath(
        Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "NaturalCommentFills", name))));
}
