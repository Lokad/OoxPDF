using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxListTabsClearMeasuredPrefixesAcrossContainers()
    {
        foreach (string container in new[] { "body", "cell", "header", "footer" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode, CachedMathGeometry(mode));
            var control = renderer.RenderPagesAsync(ReadListTabs(container, mode, 72d, "tab"), null, default).ToBlockingEnumerable().ToArray();
            var source = ReadListTabs(container, mode, 18d, "tab");
            var paragraphs = NumberedCounterParagraphs(source);
            var labels = paragraphs.Select(p => p.ListLabel).ToArray();
            var diagnostics = new List<OoxPdfDiagnostic>();
            var actual = renderer.RenderPagesAsync(source, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            // Consuming an earlier tab cannot cover the next paragraph text.
            // The unconsumed 72pt control supplies the same explicit next stop.
            AssertMetricPaint(control, actual);
            var shows = actual.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true,
                includeTextTransforms: true, honorActualText: false)).ToArray();
            var prefixes = shows.Where(show => show.Text.StartsWith("STEP ", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(paragraphs.Sum(p => p.ListLabel!.NumberingChange is null ? 1 : 2), prefixes.Length);
            foreach (var page in actual)
            {
                var pageShows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true,
                    includeTextTransforms: true, honorActualText: false).ToArray();
                foreach (var row in pageShows.Where(show => show.Text.StartsWith("STEP ", StringComparison.Ordinal)).GroupBy(show => show.Y))
                {
                    // Static-story fixtures also retain unnumbered body prose.
                    // Bind each check to its actual page and generated-label row.
                    var prose = pageShows.Where(show => Math.Abs(show.Y - row.Key) < .001d &&
                        new[] { "Summary:", "The two relationships", "The trend marker" }.Any(text =>
                            show.Text.StartsWith(text, StringComparison.Ordinal))).ToArray();
                    TestAssert.Equal(1, prose.Length);
                    foreach (var label in row)
                    {
                        double advance = label.Text.Sum(c => (double)label.Font.Font.GetAdvanceWidth(label.Font.Font.MapCodePoint(c))) *
                            label.Size / label.Font.Font.UnitsPerEm;
                        TestAssert.True(label.X + advance <= prose[0].X + .001d, "Every source row clears its complete visible prefix.");
                    }
                }
            }
            TestAssert.True(actual.Any(page => page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal))),
                "Clearing list prefixes preserves supported equations in each container.");
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Tab placement cannot drop composed mathematics.");
            TestAssert.True(paragraphs.Zip(labels).All(pair => ReferenceEquals(pair.First.ListLabel, pair.Second)),
                "Placement retains current/original counter ownership.");
        }
    }

    public static void DocxListTabsPreserveShortLabelsAndNonTabSuffixes()
    {
        foreach (string container in new[] { "body", "footer" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (string suffix in new[] { "tab", "space", "nothing" })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode, CachedMathGeometry(mode));
            // Current and previous short counters fit before 36pt;
            // space/nothing ignore tab settings.
            string prefix = suffix == "tab" ? "" : "step ";
            double numberStop = suffix == "tab" ? 36d : 18d;
            var control = renderer.RenderPagesAsync(ReadListTabs(container, mode, numberStop, suffix, prefix, nextStop: false),
                null, default).ToBlockingEnumerable().ToArray();
            var actual = renderer.RenderPagesAsync(ReadListTabs(container, mode, numberStop, suffix, prefix),
                null, default).ToBlockingEnumerable().ToArray();
            AssertMetricPaint(control, actual);
        }
    }

    private static DocxDocument ReadListTabs(string container, OoxPdfDocxMarkupMode mode,
        double numberStop, string suffix, string prefix = "step ", bool nextStop = true) =>
        ReadListCasing(container, mode, true, false, false, suffix, labelPrefix: prefix, paragraphChange: xml =>
        {
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            foreach (var properties in xml.Descendants(w + "pPr").Where(p => p.Element(w + "numPr") is not null))
            {
                properties.Elements(w + "tabs").Remove();
                var tabs = new XElement(w + "tabs", new XElement(w + "tab", new XAttribute(w + "val", "num"),
                    new XAttribute(w + "pos", (int)(numberStop * 20d))));
                if (nextStop && numberStop < 72d) tabs.Add(new XElement(w + "tab", new XAttribute(w + "val", "left"),
                    new XAttribute(w + "pos", "1440")));
                properties.Add(tabs);
            }
        });
}
