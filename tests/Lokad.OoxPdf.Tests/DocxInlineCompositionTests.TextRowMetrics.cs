using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxOrderedTextRowsRetainLocalFontMetricsAndPages()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var control = RenderTextRowMetrics("plain", "ordinary", mode);
            TestAssert.Equal(1, control.Length);
            foreach (string content in new[] { "picture", "math", "mixed" })
            {
                var actual = RenderTextRowMetrics(content, "ordinary", mode);
                TestAssert.Equal(1, actual.Length);
                AssertTextRowOwnersAndPitches(control, actual);
                TestAssert.Equal(content is "picture" or "mixed" ? 1 : 0,
                    actual.Sum(page => Regex.Matches(page.Content, @"/\w+\s+Do\b").Count));
                string math = string.Concat(actual.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true)
                    .Where(show => page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) &&
                        ReferenceEquals(font.Font, show.Font)))).Select(show => show.Text));
                TestAssert.Equal(content is "math" or "mixed" ? "x" : "", math.Normalize(System.Text.NormalizationForm.FormKC));
            }
        }
    }

    public static void DocxOrderedTextRowsShareKeepAndWidowAdmission()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (string admission in new[] { "keepLines", "keepNext", "widow" })
        {
            var control = RenderTextRowMetrics("plain", admission, mode);
            var actual = RenderTextRowMetrics("picture", admission, mode);
            TestAssert.Equal(admission == "keepLines" ? 1 : 2, actual.Length);
            TestAssert.Equal(control.Length, actual.Length);
            AssertTextRowOwnersAndPitches(control, actual);
            int headingPage = Array.FindIndex(actual, page => DocxPageFootnoteNumberingTests.DecodeText(page).Contains("Large heading", StringComparison.Ordinal));
            TestAssert.Equal(admission == "keepNext" ? 1 : 0, headingPage);
            TestAssert.Equal(1, actual.Sum(page => Regex.Matches(page.Content, @"/\w+\s+Do\b").Count));
            if (admission == "keepNext")
                TestAssert.True(DocxPageFootnoteNumberingTests.DecodeText(actual[1]).Contains("Afterword", StringComparison.Ordinal),
                    "The kept heading and its following paragraph occupy the same admitted page.");
        }
    }

    private static void AssertTextRowOwnersAndPitches(PdfPage[] control, PdfPage[] actual)
    {
        (int Page, double Y) Find(PdfPage[] pages, string text)
        {
            var matches = pages.SelectMany((page, index) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page)
                .GroupBy(show => show.Y).Where(group => string.Concat(group.Select(show => show.Text)).Trim() == text)
                .Select(group => (Page: index, Y: group.Key))).ToArray();
            TestAssert.Equal(1, matches.Length);
            return matches[0];
        }
        var previousControl = Find(control, "Small row 01");
        var previousActual = Find(actual, "Small row 01");
        TestAssert.Equal(previousControl.Page, previousActual.Page);
        for (int row = 2; row <= 7; row++)
        {
            var expected = Find(control, $"Small row {row:00}");
            var observed = Find(actual, $"Small row {row:00}");
            TestAssert.Equal(expected.Page, observed.Page);
            if (expected.Page == previousControl.Page)
                TestAssert.True(Math.Abs((expected.Y - previousControl.Y) - (observed.Y - previousActual.Y)) < .001d,
                    "Text-only rows retain ordinary run-local pitch independently of their paragraph's object row.");
            previousControl = expected; previousActual = observed;
        }
    }

    private static PdfPage[] RenderTextRowMetrics(string content, string admission, OoxPdfDocxMarkupMode mode)
    {
        var document = ReadMutation("Composition/report", mode, false, false, (xml, w, m) =>
        {
            var drawing = new XElement(xml.Descendants(w + "drawing").First());
            foreach (var extent in drawing.Descendants().Where(node => node.Name.LocalName is "extent" or "ext"))
                if (extent.Attribute("cx") is not null)
                { extent.SetAttributeValue("cx", "76200"); extent.SetAttributeValue("cy", "76200"); }
            XElement Run(string text, int size) => new(w + "r", new XElement(w + "rPr",
                new XElement(w + "rFonts", new XAttribute(w + "ascii", "DejaVu Sans"), new XAttribute(w + "hAnsi", "DejaVu Sans")),
                new XElement(w + "sz", new XAttribute(w + "val", size * 2))), new XElement(w + "t", text));
            XElement Spacing(int line, string rule) => new(w + "spacing", new XAttribute(w + "before", "0"),
                new XAttribute(w + "after", "0"), new XAttribute(w + "line", line), new XAttribute(w + "lineRule", rule));
            var body = xml.Root!.Element(w + "body")!; body.RemoveNodes();
            int filler = admission switch { "keepLines" => 1600, "keepNext" => 2000, "widow" => 3000, _ => 0 };
            if (filler != 0) body.Add(new XElement(w + "p", new XElement(w + "pPr", Spacing(filler, "exact")), Run("FILLER", 11)));
            var properties = new XElement(w + "pPr", Spacing(240, "auto"),
                new XElement(w + "widowControl", new XAttribute(w + "val", admission == "widow" ? "1" : "0")));
            if (admission is "keepLines" or "keepNext") properties.Add(new XElement(w + admission));
            var paragraph = new XElement(w + "p", properties, Run("Large heading", 36));
            if (content is "picture" or "mixed") paragraph.Add(new XElement(w + "r", drawing));
            if (content is "math" or "mixed") paragraph.Add(new XElement(m + "oMath", new XElement(m + "r", new XElement(m + "t", "x"))));
            for (int row = 1; row <= 7; row++)
                paragraph.Add(new XElement(w + "r", new XElement(w + "br")), Run($"Small row {row:00}", 11));
            body.Add(paragraph);
            if (admission == "keepNext") body.Add(new XElement(w + "p", new XElement(w + "pPr", Spacing(240, "auto")), Run("Afterword", 11)));
            body.Add(new XElement(w + "sectPr", new XElement(w + "pgSz", new XAttribute(w + "w", "7200"), new XAttribute(w + "h", "5200")),
                new XElement(w + "pgMar", new XAttribute(w + "top", "400"), new XAttribute(w + "bottom", "400"),
                    new XAttribute(w + "left", "400"), new XAttribute(w + "right", "400"), new XAttribute(w + "header", "0"),
                    new XAttribute(w + "footer", "0"), new XAttribute(w + "gutter", "0"))));
        });
        return new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
            mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();
    }
}
