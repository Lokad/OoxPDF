using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxGrowingInlineBoxesRetainCompleteChildInkAcrossContainers()
    {
        foreach (string story in new[] { "body", "cell", "header", "footer" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = ReadGrowingInlineBox(story, mode, 20d, 15d);
            var page = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode, CachedMathGeometry(mode))
                .RenderPagesAsync(source, null, default).ToBlockingEnumerable().Single();
            AssertSourceLabelsPaint(page, mode, GrowthLabels(), "growing " + story);
        }
    }

    public static void DocxGrowingInlineBoxFrameOwnsBothVerticalInsets()
    {
        foreach (double scale in new[] { 1d, .75d })
        {
            var source = ReadGrowingInlineBox("body", OoxPdfDocxMarkupMode.Final, 0d, 0d);
            var box = source.Paragraphs.Single(paragraph => paragraph.InlineTextBoxes.Count != 0).InlineTextBoxes.Single();
            var measurer = new DocxTests.FamilyWidthTextMeasurer();
            DocxInlineTextBoxLayout Place(DocxInlineTextBox candidate) => TestAssert.NotNull(
                DocxLayoutEngine.CreateInlineTextBoxLayout(candidate, 0, 0d, 400d * scale, 0d,
                    DocxTextAlignment.Left, measurer, 36d, scale, 1, default));
            var original = Place(box);
            foreach (var (top, bottom) in new[] { (20d, 0d), (0d, 15d), (20d, 15d) })
            {
                var actual = Place(box with { TextBoxInsetTopValue = (top * 12700d).ToString(System.Globalization.CultureInfo.InvariantCulture),
                    TextBoxInsetBottomValue = (bottom * 12700d).ToString(System.Globalization.CultureInfo.InvariantCulture) });
                TestAssert.True(Math.Abs(actual.BoxHeight - original.BoxHeight - (top + bottom) * scale) < .000001d,
                    "Growing a frame retains each authored inset in the container's scale.");
                TestAssert.True(Math.Abs(actual.TextLines[0].BaselineY - original.TextLines[0].BaselineY + top * scale) < .000001d,
                    "The top inset moves children; the bottom inset owns following clearance.");
            }
        }
    }

    private static string[] GrowthLabels() => Enumerable.Range(1, 6).Select(index => "GROWTH LINE " + index + " pyg").ToArray();

    private static DocxDocument ReadGrowingInlineBox(string story, OoxPdfDocxMarkupMode mode, double top, double bottom,
        Action<XDocument, XNamespace, XNamespace>? mutate = null)
    {
        string fixture = story == "body" ? "inline-box-body-equation-inline-box-valid" : "inline-box-math-" + story;
        return ReadMutation("InlineBoxMath/" + fixture, mode, false, false, (xml, w, m) =>
        {
            XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
            XNamespace wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            var content = xml.Descendants(w + "txbxContent").Single();
            content.ReplaceNodes(GrowthLabels().Select(label => new XElement(w + "p", new XElement(w + "pPr",
                new XElement(w + "spacing", new XAttribute(w + "before", "0"), new XAttribute(w + "after", "0"),
                    new XAttribute(w + "line", "240"), new XAttribute(w + "lineRule", "auto"))),
                new XElement(w + "r", new XElement(w + "rPr", new XElement(w + "rFonts",
                    new XAttribute(w + "ascii", "Arial"), new XAttribute(w + "hAnsi", "Arial")),
                    new XElement(w + "sz", new XAttribute(w + "val", "20"))), new XElement(w + "t", label)))));
            var properties = xml.Descendants(wps + "bodyPr").Single();
            properties.SetAttributeValue("tIns", top * 12700d);
            properties.SetAttributeValue("bIns", bottom * 12700d);
            properties.Elements().Where(element => element.Name == a + "spAutoFit" || element.Name == a + "normAutofit" || element.Name == a + "noAutofit").Remove();
            properties.Add(new XElement(a + "spAutoFit"));
            xml.Descendants().Where(element => element.Name == m + "oMath" || element.Name == m + "oMathPara").Remove();
            foreach (var extent in xml.Descendants().Where(element => (element.Name == wp + "extent" || element.Name == a + "ext") && element.Attribute("cy") is not null))
                extent.SetAttributeValue("cy", 18 * 12700);
            mutate?.Invoke(xml, w, m);
        }, partOverride: story is "header" or "footer" ? "word/inline-box-" + story + ".xml" : null);
    }
}
