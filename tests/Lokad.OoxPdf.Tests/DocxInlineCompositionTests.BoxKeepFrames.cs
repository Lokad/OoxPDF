using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxKeptFallbackBoxesKeepHeadingsWithCompleteFrames()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (bool boundary in new[] { false, true })
        foreach (var (keep, roomy) in new[] { (true, false), (true, true), (false, false) })
        {
            var source = ReadKeptFallbackBox(mode, keep, roomy, boundary, counter: false);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode, CachedMathGeometry(mode))
                .RenderPagesAsync(source, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(boundary ? 2 : 1, pages.Length);
            int labelPage = boundary ? 1 : 0;
            int headingPage = boundary && keep ? 1 : 0;
            int actualHeadingPage = Enumerable.Range(0, pages.Length).Single(index =>
                DocxPageFootnoteNumberingTests.DecodeText(pages[index]).Contains("KEPT BOX HEADING", StringComparison.Ordinal));
            TestAssert.True(actualHeadingPage == headingPage, "An authored keep-next heading follows the complete box frame; an explicitly separate heading stays put.");
            AssertSourceLabelsPaint(pages[headingPage], mode, ["KEPT BOX HEADING"], "kept box heading");
            AssertSourceLabelsPaint(pages[labelPage], mode, GrowthLabels(), "kept fallback box");
            TestAssert.Equal(1, diagnostics.Count(item => item.Id == "DOCX_UNSUPPORTED_EQUATION"));
        }
    }

    public static void DocxKeptFallbackBoxesBindPageAndTotalFieldsAfterAdmission()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (bool boundary in new[] { false, true })
        {
            var source = ReadKeptFallbackBox(mode, keep: true, roomy: false, boundary, counter: true);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode, CachedMathGeometry(mode));
            var pages = renderer.RenderPagesAsync(source, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            int expected = boundary ? 1 : 0;
            TestAssert.Equal(expected + 1, pages.Length);
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(pages[expected], usePageCoordinates: true).ToArray();
            string text = string.Concat(shows.SelectMany(show => show.Text).Where(character => !char.IsWhiteSpace(character)));
            TestAssert.True(text.Contains("KEPTBOXHEADING", StringComparison.Ordinal), "Counter inputs cannot detach the kept source heading.");
            TestAssert.True(text.Contains("GROWTHLINE1pyg" + (expected + 1) + pages.Length, StringComparison.Ordinal),
                "The child PAGE and NUMPAGES values follow the admitted frame and final document count.");
            TestAssert.True(!text.Contains('8') && !text.Contains('9'), "Stale PAGE and NUMPAGES cached results never paint.");
            AssertSourceLabelsPaint(pages[expected], mode, GrowthLabels(), "kept counter box");
            TestAssert.Equal(1, diagnostics.Count(item => item.Id == "DOCX_UNSUPPORTED_EQUATION"));
            using var canceled = new CancellationTokenSource();
            canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(source, null, canceled.Token)
                .ToBlockingEnumerable().ToArray());
        }
    }

    private static DocxDocument ReadKeptFallbackBox(OoxPdfDocxMarkupMode mode, bool keep, bool roomy, bool boundary, bool counter) =>
        ReadFallbackGrowingBox(mode, boundary, roomy, counter, mutate: (xml, w, _) =>
        {
            var content = xml.Descendants(w + "txbxContent").Single();
            var owner = content.Ancestors(w + "p").First();
            var style = content.Elements(w + "p").First().Element(w + "r")!.Element(w + "rPr")!;
            owner.AddBeforeSelf(new XElement(w + "p", new XElement(w + "pPr",
                new XElement(w + "keepNext", new XAttribute(w + "val", keep ? "1" : "0")),
                new XElement(w + "spacing", new XAttribute(w + "before", "0"), new XAttribute(w + "after", "0"),
                    new XAttribute(w + "line", "240"), new XAttribute(w + "lineRule", "exact"))),
                new XElement(w + "r", new XElement(style), new XElement(w + "t", "KEPT BOX HEADING"))));
            if (counter)
                content.Elements(w + "p").First().Add(new XElement(w + "fldSimple", new XAttribute(w + "instr", " NUMPAGES "),
                    new XElement(w + "r", new XElement(style), new XElement(w + "t", "8"))));
        });
}
