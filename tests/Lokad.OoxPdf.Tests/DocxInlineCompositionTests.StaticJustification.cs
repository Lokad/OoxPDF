using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxStaticJustificationSharesBodyWordPositionsAndLinks()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode, CachedMathGeometry(mode));
            var body = renderer.RenderPagesAsync(ReadStaticJustification("body", mode), null, default).ToBlockingEnumerable().ToArray();
            var expected = JustifiedStaticWordPositions(body);
            TestAssert.Equal(18, expected.Length);
            TestAssert.True(expected.Select(word => word.Y).Distinct().Count() > 1, "The source exercises wrapped and final rows.");
            foreach (string container in new[] { "header", "footer" })
            {
                var source = ReadStaticJustification(container, mode, navigation: true);
                var pages = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
                var actual = JustifiedStaticWordPositions(pages);
                TestAssert.Equal(expected.Length, actual.Length);
                for (int index = 0; index < expected.Length; index++)
                {
                    TestAssert.Equal(expected[index].Word, actual[index].Word);
                    TestAssert.True(Math.Abs(expected[index].X - actual[index].X) <= .001d + 1e-9,
                        "Justified words share body placement without an equation, picture, tab or indent: " + container + "/" + mode);
                }
                TestAssert.Equal(1, pages.Sum(page => page.Annotations.Count(annotation => annotation.Destination is not null)));
                TestAssert.True(DocxBlockTraversal.EnumerateSourceParagraphs(source).All(p => p.Equations.Count == 0),
                    "Plain justification needs no synthetic math owner.");
                using var canceled = new CancellationTokenSource(); canceled.Cancel();
                TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(source, null, canceled.Token).ToBlockingEnumerable().ToArray());
            }
        }
    }

    public static void DocxStaticJustificationKeepsLastRowsAndHiddenOwnerPaint()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (string container in new[] { "header", "footer" })
        foreach (int fontSize in new[] { 11, 12, 24 })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode, CachedMathGeometry(mode));
            foreach (bool hidden in new[] { false, true })
            {
                var left = renderer.RenderPagesAsync(ReadStaticJustification(container, mode, justified: false,
                    longText: false, hidden: hidden, fontSize: fontSize), null, default).ToBlockingEnumerable().ToArray();
                var justified = renderer.RenderPagesAsync(ReadStaticJustification(container, mode,
                    longText: false, hidden: hidden, fontSize: fontSize), null, default).ToBlockingEnumerable().ToArray();
                AssertMetricPaint(left, justified);
                TestAssert.True(justified.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true,
                    honorActualText: false)).Any(show => show.Text.Contains("FollowingSpread.", StringComparison.Ordinal)),
                    "A non-stretched or hidden owner retains following visible text.");
            }
        }
    }

    private static (string Word, double X, double Y)[] JustifiedStaticWordPositions(PdfPage[] pages) =>
        pages.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true,
            includeTextTransforms: true, honorActualText: false))
            .Where(show => show.Text.Trim().StartsWith("Spread", StringComparison.Ordinal))
            .Select(show => (show.Text.Trim(), show.X, show.Y)).ToArray();

    private static DocxDocument ReadStaticJustification(string container, OoxPdfDocxMarkupMode mode,
        bool justified = true, bool longText = true, bool hidden = false, int fontSize = 12, bool navigation = false) =>
        ReadPlainTabFields(container, mode, "none", paragraphChange: (paragraph, w) =>
        {
            var properties = paragraph.Element(w + "pPr")!;
            properties.Elements(w + "jc").Remove();
            properties.Add(new XElement(w + "jc", new XAttribute(w + "val", justified ? "both" : "left")));
            var style = new XElement(paragraph.Elements(w + "r").First().Element(w + "rPr")!);
            style.Elements(w + "sz").Remove(); style.Add(new XElement(w + "sz", new XAttribute(w + "val", fontSize * 2)));
            if (hidden) style.Add(new XElement(w + "vanish"));
            paragraph.Elements().Where(element => element != properties).Remove();
            string[] words = longText ? ["SpreadAlpha", .. Enumerable.Range(1, 16).Select(index => $"Spread{index:00}"), "SpreadOmega."] : ["SpreadAlpha."];
            for (int index = 0; index < words.Length; index++)
            {
                var run = new XElement(w + "r", new XElement(style), new XElement(w + "t",
                    new XAttribute(XNamespace.Xml + "space", "preserve"), words[index] + (index + 1 < words.Length ? " " : "")));
                paragraph.Add(navigation && index == 0
                    ? new XElement(w + "hyperlink", new XAttribute(w + "anchor", "SpreadAnchor"), run) : run);
            }
            if (navigation)
            {
                properties.AddAfterSelf(new XElement(w + "bookmarkStart", new XAttribute(w + "id", "161"), new XAttribute(w + "name", "SpreadAnchor")));
                paragraph.Add(new XElement(w + "bookmarkEnd", new XAttribute(w + "id", "161")));
            }
            var followingStyle = new XElement(style); followingStyle.Elements(w + "vanish").Remove();
            paragraph.AddAfterSelf(new XElement(w + "p", new XElement(w + "r", followingStyle, new XElement(w + "t", "FollowingSpread."))));
        });
}
