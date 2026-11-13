using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxFallbackInlineBoxesAdmitCompleteFramesBesideUnsupportedMath()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (bool boundary in new[] { false, true })
        foreach (bool roomy in new[] { false, true })
        {
            var source = ReadFallbackGrowingBox(mode, boundary, roomy, counter: false);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode, CachedMathGeometry(mode))
                .RenderPagesAsync(source, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            int expected = boundary ? 1 : 0;
            int actual = Enumerable.Range(0, pages.Length).Single(index =>
                DocxPageFootnoteNumberingTests.DecodeText(pages[index]).Contains(GrowthLabels()[0], StringComparison.Ordinal));
            TestAssert.True(actual == expected, "Fallback admission uses the complete frame, including grown content and insets.");
            AssertSourceLabelsPaint(pages[expected], mode, GrowthLabels(), "fallback growing box");
            TestAssert.Equal(1, diagnostics.Count(item => item.Id == "DOCX_UNSUPPORTED_EQUATION"));
        }
    }

    public static void DocxFallbackInlineBoxPageCountersFollowTheAdmittedFrame()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (bool boundary in new[] { false, true })
        {
            var source = ReadFallbackGrowingBox(mode, boundary, roomy: false, counter: true);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode, CachedMathGeometry(mode));
            var pages = renderer.RenderPagesAsync(source, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            int expected = boundary ? 1 : 0;
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(pages[expected], usePageCoordinates: true).ToArray();
            TestAssert.Equal(1, shows.Count(show => show.Text.Trim() == (expected + 1).ToString()));
            TestAssert.True(!shows.Any(show => show.Text.Trim() == "9"), "The stale source PAGE result never paints.");
            AssertSourceLabelsPaint(pages[expected], mode, GrowthLabels(), "fallback counter box");
            TestAssert.Equal(1, diagnostics.Count(item => item.Id == "DOCX_UNSUPPORTED_EQUATION"));
            using var canceled = new CancellationTokenSource();
            canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(source, null, canceled.Token)
                .ToBlockingEnumerable().ToArray());
        }
    }

    public static void DocxFallbackCellBoxesReserveCompleteChildFrames()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (bool roomy in new[] { false, true })
        {
            var source = ReadFallbackGrowingBox(mode, boundary: false, roomy, counter: false, story: "cell");
            var diagnostics = new List<OoxPdfDiagnostic>();
            var page = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode, CachedMathGeometry(mode))
                .RenderPagesAsync(source, diagnostics.Add, default).ToBlockingEnumerable().Single();
            AssertSourceLabelsPaint(page, mode, GrowthLabels(), "fallback cell box");
            TestAssert.Equal(1, diagnostics.Count(item => item.Id == "DOCX_UNSUPPORTED_EQUATION"));
        }
    }

    private static DocxDocument ReadFallbackGrowingBox(OoxPdfDocxMarkupMode mode, bool boundary, bool roomy, bool counter,
        string story = "body") => ReadGrowingInlineBox(story, mode, 20d, 15d, (xml, w, m) =>
        {
            var content = xml.Descendants(w + "txbxContent").Single();
            var owner = content.Ancestors(w + "p").First();
            var equation = new XElement(m + "oMath", new XElement(m + "acc", new XElement(m + "e",
                new XElement(m + "r", new XElement(m + "t", "q")))));
            if (owner.Element(w + "pPr") is { } properties) properties.AddAfterSelf(equation);
            else owner.AddFirst(equation);
            if (boundary)
                owner.AddBeforeSelf(new XElement(w + "p", new XElement(w + "pPr", new XElement(w + "spacing",
                    new XAttribute(w + "before", "0"), new XAttribute(w + "after", "0"),
                    new XAttribute(w + "line", "12000"), new XAttribute(w + "lineRule", "exact")))));
            if (roomy)
                foreach (var extent in xml.Descendants().Where(element => (element.Name.LocalName is "extent" or "ext") && element.Attribute("cy") is not null))
                    extent.SetAttributeValue("cy", 200 * 12700);
            if (counter)
            {
                var paragraph = content.Elements(w + "p").First();
                paragraph.Add(new XElement(w + "fldSimple", new XAttribute(w + "instr", " PAGE "),
                    new XElement(w + "r", new XElement(paragraph.Element(w + "r")!.Element(w + "rPr")!), new XElement(w + "t", "9"))));
            }
        });
}
