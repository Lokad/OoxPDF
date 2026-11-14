using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxTabbedTextRowsRetainLocalMetricsAndVisibleContent()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (bool numbered in new[] { false, true })
        foreach (string content in new[] { "plain", "picture", "math", "mixed" })
        {
            var control = RenderTextRowMetricDocument(numbered ? ReadNumberedTextRows(content, "ordinary", mode) :
                ReadTextRowMetrics(content, "ordinary", mode), mode);
            var actual = RenderTextRowMetricDocument(ReadTabbedTextRows(content, "ordinary", mode, numbered), mode);
            TestAssert.Equal(1, actual.Length);
            AssertTextRowOwnersAndPitches(control, actual);
            string nativeText = string.Concat(actual.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page,
                true, honorActualText: false)).Select(show => show.Text));
            for (int row = 1; row <= 7; row++)
                TestAssert.Equal(1, Regex.Matches(nativeText, $"Small row {row:00}").Count);
            TestAssert.Equal(1, Regex.Matches(nativeText, "Large heading").Count);
            TestAssert.Equal(content is "picture" or "mixed" ? 1 : 0,
                actual.Sum(page => Regex.Matches(page.Content, @"/\w+\s+Do\b").Count));
            if (numbered) AssertNumberedTextRowContent(actual, content);
        }
    }

    public static void DocxTabbedTextRowsShareKeepAndWidowAdmission()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (bool numbered in new[] { false, true })
        foreach (string admission in new[] { "keepLines", "keepNext", "widow" })
        {
            var control = RenderTextRowMetricDocument(numbered ? ReadNumberedTextRows("picture", admission, mode) :
                ReadTextRowMetrics("picture", admission, mode), mode);
            var actual = RenderTextRowMetricDocument(ReadTabbedTextRows("picture", admission, mode, numbered), mode);
            TestAssert.Equal(admission == "keepLines" ? 1 : 2, actual.Length);
            AssertTextRowOwnersAndPitches(control, actual);
            int headingPage = Array.FindIndex(actual, page => DocxPageFootnoteNumberingTests.DecodeText(page)
                .Contains("Large heading", StringComparison.Ordinal));
            TestAssert.Equal(admission == "keepNext" ? 1 : 0, headingPage);
        }
    }

    public static void DocxTabbedTextRowsKeepOwnRunStrutsAndDeclaredSlots()
    {
        foreach (bool numbered in new[] { false, true })
        foreach (bool exact in new[] { false, true })
        {
            var source = ReadTabbedTextRows("picture", "ordinary", OoxPdfDocxMarkupMode.Final, numbered,
                largeFirstTab: true, exact: exact);
            var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
            var plan = DocxFontPlan.Create(source, resolver, default);
            var measurer = DocxFontPlanTextMeasurer.CreateAsync(plan, resolver.Resolve(new FontRequest("DejaVu Sans", false, false)),
                default, resolver).GetAwaiter().GetResult();
            var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(source, measurer, default);
            var rows = layout.Pages.SelectMany(page => page.Items).OfType<DocxTextLineLayout>()
                .Where(line => line.Text.TrimStart('\t').StartsWith("Small row", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(7, rows.Length);
            TestAssert.Equal(1, layout.Pages.Count);
            TestAssert.Equal(36d, rows[0].FontSize);
            TestAssert.True(rows.All(row => ReferenceEquals(row.SourceParagraph, source.Paragraphs[0])),
                "Tabs retain their original paragraph and run ownership.");
            if (exact)
                TestAssert.True(rows.All(row => row.LineHeight == 12d), "An authored exact slot remains exact despite the large tab run.");
            else
            {
                TestAssert.Equal(11d, rows[1].FontSize);
                TestAssert.True(rows[0].LineHeight > 2d * rows[1].LineHeight,
                    "The actual large tab run contributes its own strut; later short tabs do not inherit it.");
            }
        }
    }

    private static DocxDocument ReadTabbedTextRows(string content, string admission, OoxPdfDocxMarkupMode mode,
        bool numbered, bool largeFirstTab = false, bool exact = false) =>
        ReadTextRowMetrics(content, admission, mode, (xml, w, _) =>
        {
            var paragraph = xml.Descendants(w + "p").Single(p => p.Descendants(w + "t").Any(text => text.Value == "Large heading"));
            if (numbered) paragraph.Element(w + "pPr")!.Add(new XElement(w + "numPr",
                new XElement(w + "ilvl", new XAttribute(w + "val", "0")), new XElement(w + "numId", new XAttribute(w + "val", "1"))));
            if (exact)
            {
                var spacing = paragraph.Element(w + "pPr")!.Element(w + "spacing")!;
                spacing.SetAttributeValue(w + "lineRule", "exact"); spacing.SetAttributeValue(w + "line", "240");
            }
            foreach (var text in paragraph.Elements(w + "r").Elements(w + "t").Where(text => text.Value.StartsWith("Small row", StringComparison.Ordinal)).ToArray())
            {
                if (largeFirstTab && text.Value == "Small row 01")
                {
                    var properties = new XElement(text.Parent!.Element(w + "rPr")!);
                    properties.Element(w + "sz")!.SetAttributeValue(w + "val", "72");
                    text.Parent.AddBeforeSelf(new XElement(w + "r", properties, new XElement(w + "tab")));
                }
                else text.AddBeforeSelf(new XElement(w + "tab"));
            }
        });
}
