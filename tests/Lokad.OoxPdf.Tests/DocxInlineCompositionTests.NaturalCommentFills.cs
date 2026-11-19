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

    public static void DocxMixedCommentFillKeepsPictureAndFormulaGeometry()
    {
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        foreach (bool split in new[] { false, true })
        {
            const OoxPdfDocxMarkupMode mode = OoxPdfDocxMarkupMode.AllMarkup;
            var document = ReadMutation("Composition/composed-static-cell-report", mode, false, false, (xml, w, m) =>
            {
                var body = xml.Root!.Element(w + "body")!;
                var paragraph = body.Elements(w + "p").Single(p => p.Elements(w + "commentRangeStart").Any() &&
                    p.Descendants(w + "drawing").Any());
                body.Elements().Where(element => element != paragraph && element.Name != w + "sectPr").Remove();
                body.Descendants().Where(element => element.Name == w + "headerReference" ||
                    element.Name == w + "footerReference").Remove();
                var math = paragraph.Elements(m + "oMath").Single();
                // Keep a following word on the selected line with portable fonts.
                math.ElementsAfterSelf(w + "r").First().Element(w + "t")!.Value = " ends.";
                if (split)
                {
                    var run = paragraph.Element(w + "hyperlink")!.Element(w + "r")!;
                    var continuation = new XElement(run);
                    run.Element(w + "t")!.Value = "The trend ";
                    continuation.Element(w + "t")!.Value = "marker ";
                    run.AddAfterSelf(continuation);
                    math.AddBeforeSelf(new XElement(w + "bookmarkStart", new XAttribute(w + "id", "7010"),
                        new XAttribute(w + "name", "SelectedFormula")));
                    math.AddAfterSelf(new XElement(w + "bookmarkEnd", new XAttribute(w + "id", "7010")));
                }
            });
            var page = new DocxRenderer(resolver, mode, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderPagesAsync(document, null, default).ToBlockingEnumerable().Single();
            var picture = ReadPictureRangeFrame(page);
            var formula = PortableRenderingTests.ReadPaintedSourceLabel(page, "a+b=c", mathOnly: true);
            var left = PortableRenderingTests.ReadPaintedSourceLabel(page, "The trend marker");
            var right = PortableRenderingTests.ReadPaintedSourceLabel(page, "ends.");
            var fills = ReadNaturalCommentFillRectangles(page.Content).Where(rect => rect.X < 400d).ToArray();
            var leading = fills.Single(rect => Contains(rect, left[0].Ink) && Contains(rect, picture));
            var mathFill = fills.Single(rect => formula.All(glyph => Contains(rect, glyph.Ink)));
            var trailing = fills.Single(rect => Contains(rect, right[0].Ink));
            TestAssert.True(leading.Height > mathFill.Height + 1d && trailing.Height > mathFill.Height + 1d,
                "The formula uses its own selection geometry inside the selected prose.");
            TestAssert.True(Math.Abs(leading.X + leading.Width - mathFill.X) < .003d &&
                Math.Abs(mathFill.X + mathFill.Width - trailing.X) < .003d,
                "The source-owned bands meet at the equation without overpainting the entire line.");
        }

        static bool Contains(PdfRectangle box, PdfRectangle ink) => ink.X >= box.X - .02d &&
            ink.X + ink.Width <= box.X + box.Width + .02d && ink.Y >= box.Y - .02d &&
            ink.Y + ink.Height <= box.Y + box.Height + .02d;
    }

    private static IEnumerable<PdfRectangle> ReadNaturalCommentFillRectangles(string content)
    {
        foreach (Match block in Regex.Matches(content,
            @"0\.973 0\.863 0\.867 rg\s+(?:(?<x>[-\d.]+) (?<y>[-\d.]+) (?<w>[\d.]+) (?<h>[\d.]+) re f\s*)+"))
        {
            for (int index = 0; index < block.Groups["x"].Captures.Count; index++)
                yield return new PdfRectangle(Value("x", index), Value("y", index), Value("w", index), Value("h", index));

            double Value(string name, int index) => double.Parse(block.Groups[name].Captures[index].Value, CultureInfo.InvariantCulture);
        }
    }

    private static JsonDocument ReadNaturalCommentFillJson(string name) => JsonDocument.Parse(File.ReadAllText(Path.GetFullPath(
        Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "NaturalCommentFills", name))));
}
