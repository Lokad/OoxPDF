using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxOrderedInlineParagraphsHonorDefaultAndDisabledWidowControl()
    {
        foreach (int count in new[] { 3, 4 })
        foreach (string content in new[] { "text", "picture", "math", "mixed" })
        foreach (bool? enabled in new bool?[] { null, true, false })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var pages = RenderWidowInlineSource(count, content, enabled, mode, counter: false);
            TestAssert.Equal(2, pages.Length);
            for (int line = 1; line <= count; line++)
            {
                int firstPageLines = count == 3 ? enabled != false ? 0 : 1 : enabled != false ? 2 : 3;
                int expected = line <= firstPageLines ? 0 : 1;
                int actual = Enumerable.Range(0, pages.Length).Single(index =>
                    DocxPageFootnoteNumberingTests.DecodeText(pages[index]).Contains("ROW " + line, StringComparison.Ordinal));
                TestAssert.True(expected == actual,
                    $"Shared head/tail admission retains row {line}: {count}/{content}/{enabled}/{mode}.");
            }
            if (content is "picture" or "mixed")
                TestAssert.Equal(1, pages.Sum(page => Regex.Matches(page.Content, @"/\w+\s+Do\b").Count));
            if (content is "math" or "mixed")
                TestAssert.Equal("x", string.Concat(pages.SelectMany(page =>
                    DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true).Where(show => page.Fonts.Any(font =>
                        font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(font.Font, show.Font))))
                    .Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
        }
    }

    public static void DocxOrderedWidowFragmentsRebindCountersAndRetainPictureNavigation()
    {
        foreach (bool enabled in new[] { true, false })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var pages = RenderWidowInlineSource(3, "mixed", enabled, mode, counter: true);
            int expected = enabled ? 1 : 0;
            string first = DocxPageFootnoteNumberingTests.DecodeText(pages[expected]);
            TestAssert.True(first.Contains("ROW 1 " + (expected + 1), StringComparison.Ordinal),
                "A deferred first fragment binds PAGE only to its admitted page.");
            TestAssert.Equal(1, Regex.Matches(pages[expected].Content, @"/\w+\s+Do\b").Count);
            var link = pages.SelectMany(page => page.Annotations).Single(annotation => annotation.IsDestination);
            TestAssert.Equal(expected, link.Destination!.Value.PageIndex);
        }
    }

    private static PdfPage[] RenderWidowInlineSource(int count, string content, bool? enabled,
        OoxPdfDocxMarkupMode mode, bool counter)
    {
        var document = ReadMutation("Composition/report", mode, false, false, (xml, w, m) =>
        {
            var drawing = new XElement(xml.Descendants(w + "drawing").First());
            foreach (var extent in drawing.Descendants().Where(element => element.Name.LocalName is "extent" or "ext"))
            {
                if (extent.Attribute("cx") is null) continue;
                extent.SetAttributeValue("cx", "76200"); extent.SetAttributeValue("cy", "76200");
            }
            var body = xml.Root!.Element(w + "body")!;
            body.RemoveNodes();
            XElement Run(string text) => new(w + "r", new XElement(w + "rPr",
                new XElement(w + "rFonts", new XAttribute(w + "ascii", "Calibri"), new XAttribute(w + "hAnsi", "Calibri")),
                new XElement(w + "sz", new XAttribute(w + "val", "22"))), new XElement(w + "t", text));
            XElement Paragraph(int height) => new(w + "p", new XElement(w + "pPr", new XElement(w + "spacing",
                new XAttribute(w + "before", "0"), new XAttribute(w + "after", "0"),
                new XAttribute(w + "line", height), new XAttribute(w + "lineRule", "exact"))));
            var filler = Paragraph(count == 3 ? 2600 : 1800); filler.Add(Run("FILLER")); body.Add(filler);
            var paragraph = Paragraph(400);
            if (enabled.HasValue) paragraph.Element(w + "pPr")!.Add(new XElement(w + "widowControl", new XAttribute(w + "val", enabled.Value ? "1" : "0")));
            if (counter) paragraph.Add(new XElement(w + "bookmarkStart", new XAttribute(w + "id", "8802"), new XAttribute(w + "name", "KeptParagraph")));
            if (content is "picture" or "mixed") paragraph.Add(new XElement(w + "r", drawing));
            for (int index = 1; index <= count; index++)
            {
                if (index > 1) paragraph.Add(new XElement(w + "r", new XElement(w + "br")));
                paragraph.Add(Run("ROW " + index + (index == 1 ? " " : "")));
                if (index != 1) continue;
                if (counter) paragraph.Add(new XElement(w + "fldSimple", new XAttribute(w + "instr", " PAGE "), Run("9")));
                if (content is "math" or "mixed") paragraph.Add(new XElement(m + "oMath", new XElement(m + "r", new XElement(m + "t", "x"))));
            }
            if (counter) paragraph.Add(new XElement(w + "bookmarkEnd", new XAttribute(w + "id", "8802")));
            body.Add(paragraph);
            if (counter)
            {
                var jump = Paragraph(400); jump.Add(new XElement(w + "hyperlink", new XAttribute(w + "anchor", "KeptParagraph"), Run("JUMP"))); body.Add(jump);
            }
            body.Add(new XElement(w + "sectPr",
                new XElement(w + "pgSz", new XAttribute(w + "w", "6000"), new XAttribute(w + "h", "4000")),
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
