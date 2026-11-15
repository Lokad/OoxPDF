using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxStaticTabsShareBodyFieldOffsetsWithoutEquations()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (string stops in new[] { "left", "default", "center", "right" })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode, CachedMathGeometry(mode));
            var control = renderer.RenderPagesAsync(ReadPlainTabFields("body", mode, stops), null, default).ToBlockingEnumerable().ToArray();
            var expected = PlainTabFieldOffsets(control);
            foreach (string container in new[] { "header", "footer" })
            {
                var source = ReadPlainTabFields(container, mode, stops);
                var actual = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
                var offsets = PlainTabFieldOffsets(actual);
                TestAssert.True(expected.Zip(offsets).All(pair => Math.Abs(pair.First - pair.Second) < .001d),
                    "Removing equations cannot change the authored tab contract across containers: " + container + "/" + stops + "/" + mode);
                TestAssert.True(actual.All(page => page.Fonts.All(font => !font.ResourceName.StartsWith("Eq", StringComparison.Ordinal))),
                    "The plain source has no formula that could select the tab-aware path.");
                TestAssert.True(DocxBlockTraversal.EnumerateSourceParagraphs(source).All(p => p.Equations.Count == 0),
                    "The same visible tabs retain source positions without synthetic equations.");
            }
        }
    }

    public static void DocxStaticTabsKeepFieldsLinksHiddenOwnersAndCancellation()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode, CachedMathGeometry(mode));
            var control = renderer.RenderPagesAsync(ReadPlainTabFields("body", mode, "left", fields: true), null, default).ToBlockingEnumerable().ToArray();
            var expected = PlainTabFieldOffsets(control, fields: true);
            foreach (string container in new[] { "header", "footer" })
            {
                var source = ReadPlainTabFields(container, mode, "left", fields: true);
                var actual = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
                TestAssert.True(expected.Zip(PlainTabFieldOffsets(actual, fields: true)).All(pair => Math.Abs(pair.First - pair.Second) < .001d),
                    "Static page fields and hyperlink text retain the same tab positions as body content.");
                TestAssert.Equal(1, actual.Sum(page => page.Annotations.Count(link => link.Destination is not null)));
                var plain = renderer.RenderPagesAsync(ReadPlainTabFields(container, mode, "none"), null, default).ToBlockingEnumerable().ToArray();
                var hidden = renderer.RenderPagesAsync(ReadPlainTabFields(container, mode, "none", hiddenTab: true), null, default).ToBlockingEnumerable().ToArray();
                AssertMetricPaint(plain, hidden);
                using var canceled = new CancellationTokenSource(); canceled.Cancel();
                TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(source, null, canceled.Token).ToBlockingEnumerable().ToArray());
            }
        }
    }

    private static double[] PlainTabFieldOffsets(PdfPage[] pages, bool fields = false)
    {
        var shows = pages.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true,
            includeTextTransforms: true, honorActualText: false)).ToArray();
        double start = shows.Single(show => show.Text.TrimStart().StartsWith("Start", StringComparison.Ordinal)).X;
        double middle = shows.Single(show => show.Text.TrimStart().StartsWith("Middle", StringComparison.Ordinal)).X;
        double end = shows.Single(show => fields ? show.Text.Trim() == "1" : show.Text.Trim() == "End.").X;
        return [middle - start, end - start];
    }

    private static DocxDocument ReadPlainTabFields(string container, OoxPdfDocxMarkupMode mode,
        string stops, bool fields = false, bool hiddenTab = false) =>
        ReadMutation("CompositionProbes/math-tabs-left72-" + container, mode, false, false, (xml, w, m) =>
        {
            var paragraph = xml.Descendants(w + "p").Single(p => p.Elements(m + "oMath").Any());
            paragraph.Elements(m + "oMath").Remove();
            var properties = paragraph.Element(w + "pPr")!;
            if (stops is "default" or "none") properties.Elements(w + "tabs").Remove();
            else if (stops is "center" or "right")
            {
                int position = 2160;
                foreach (var tab in properties.Element(w + "tabs")!.Elements(w + "tab"))
                {
                    tab.SetAttributeValue(w + "val", stops); tab.SetAttributeValue(w + "pos", position); position += 2160;
                }
            }
            if (stops == "none") paragraph.Descendants(w + "tab").Remove();
            if (fields)
            {
                var first = paragraph.Elements(w + "r").Single(r => r.Element(w + "t")?.Value == "Start ");
                first.AddBeforeSelf(new XElement(w + "bookmarkStart", new XAttribute(w + "id", "150"), new XAttribute(w + "name", "TabOrigin")));
                first.AddAfterSelf(new XElement(w + "bookmarkEnd", new XAttribute(w + "id", "150")));
                var middle = paragraph.Elements(w + "r").Single(r => r.Element(w + "t")?.Value == "Middle ");
                middle.ReplaceWith(new XElement(w + "hyperlink", new XAttribute(w + "anchor", "TabOrigin"), new XElement(middle)));
                var end = paragraph.Elements(w + "r").Single(r => r.Element(w + "t")?.Value == "End.");
                end.Element(w + "t")!.Value = "23";
                end.ReplaceWith(new XElement(w + "fldSimple", new XAttribute(w + "instr", "PAGE"), new XElement(end)));
            }
            if (hiddenTab)
                paragraph.Add(new XElement(w + "r", new XElement(w + "rPr", new XElement(w + "vanish"),
                    new XElement(w + "rFonts", new XAttribute(w + "ascii", "Arial"), new XAttribute(w + "hAnsi", "Arial")),
                    new XElement(w + "sz", new XAttribute(w + "val", "160"))), new XElement(w + "tab")));
        }, partOverride: container == "body" ? null : "word/math-tabs-" + container + ".xml");
}
