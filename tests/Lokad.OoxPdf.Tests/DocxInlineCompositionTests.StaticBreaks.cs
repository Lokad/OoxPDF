using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxStaticHardBreaksRetainAuthoredEmptyRows()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode, CachedMathGeometry(mode));
            for (int breaks = 1; breaks <= 3; breaks++)
            {
                var body = Render("body", breaks);
                double expected = BreakOrigin(body, "BreakOmega.") - BreakOrigin(body, "BreakAlpha.");
                TestAssert.True(Math.Abs(expected) > 0d, "Actual authored breaks separate the source words.");
                foreach (string container in new[] { "header", "footer" })
                {
                    var pages = Render(container, breaks);
                    double gap = BreakOrigin(pages, "BreakOmega.") - BreakOrigin(pages, "BreakAlpha.");
                    TestAssert.True(Math.Abs(gap - expected) < .001d,
                        "Every authored empty row has the same advance across containers: " + container + "/" + breaks + "/" + mode);
                    TestAssert.True(pages.All(page => page.Fonts.All(font => !font.ResourceName.StartsWith("Eq", StringComparison.Ordinal))),
                        "Plain hard breaks need no synthetic equation owner.");
                }
            }
            PdfPage[] Render(string container, int breaks) => renderer.RenderPagesAsync(
                ReadStaticBreakFrame(container, mode, breaks), null, default).ToBlockingEnumerable().ToArray();
        }
    }

    public static void DocxStaticHardBreaksKeepHiddenOwnersAndBoundaryRows()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode, CachedMathGeometry(mode));
            var body = Render("body", 1, leading: 2, trailing: 2);
            foreach (string container in new[] { "header", "footer" })
            {
                var control = Render(container, 0);
                AssertMetricPaint(control, Render(container, 3, hidden: true));
                var pages = Render(container, 1, leading: 2, trailing: 2);
                foreach (string word in new[] { "BreakOmega.", "BreakFollowing." })
                    TestAssert.True(Math.Abs((BreakOrigin(pages, word) - BreakOrigin(pages, "BreakAlpha.")) -
                        (BreakOrigin(body, word) - BreakOrigin(body, "BreakAlpha."))) < .001d,
                        "Leading/trailing rows retain following content and paragraph placement: " + container + "/" + mode);
                using var canceled = new CancellationTokenSource(); canceled.Cancel();
                TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(
                    ReadStaticBreakFrame(container, mode, 2), null, canceled.Token).ToBlockingEnumerable().ToArray());
            }
            PdfPage[] Render(string container, int breaks, bool hidden = false, int leading = 0, int trailing = 0) =>
                renderer.RenderPagesAsync(ReadStaticBreakFrame(container, mode, breaks, hidden, leading, trailing),
                    null, default).ToBlockingEnumerable().ToArray();
        }
    }

    private static double BreakOrigin(PdfPage[] pages, string word) => pages.SelectMany(page =>
        DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true, includeTextTransforms: true, honorActualText: false))
        .Single(show => show.Text.Contains(word, StringComparison.Ordinal)).Y;

    private static DocxDocument ReadStaticBreakFrame(string container, OoxPdfDocxMarkupMode mode,
        int breaks, bool hidden = false, int leading = 0, int trailing = 0) =>
        ReadPlainTabFields(container, mode, "none", paragraphChange: (paragraph, w) =>
        {
            var properties = paragraph.Element(w + "pPr")!;
            var style = paragraph.Elements(w + "r").First().Element(w + "rPr");
            paragraph.Elements().Where(element => element != properties).Remove();
            AddBreaks(leading, false);
            paragraph.Add(Run(new XElement(w + "t", "BreakAlpha.")));
            AddBreaks(breaks, hidden);
            paragraph.Add(Run(new XElement(w + "t", "BreakOmega.")));
            AddBreaks(trailing, false);
            paragraph.AddAfterSelf(new XElement(w + "p", new XElement(properties), Run(new XElement(w + "t", "BreakFollowing."))));

            XElement Run(params object[] content) => new(w + "r", style is null ? null : new XElement(style), content);
            void AddBreaks(int count, bool hide)
            {
                if (count == 0) return;
                var run = Run(Enumerable.Range(0, count).Select(_ => new XElement(w + "br")));
                if (hide)
                {
                    var runProperties = run.Element(w + "rPr");
                    if (runProperties is null) { runProperties = new XElement(w + "rPr"); run.AddFirst(runProperties); }
                    runProperties.Add(new XElement(w + "vanish"));
                }
                paragraph.Add(run);
            }
        });
}
