using System.Xml.Linq;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxInheritanceTests
{
    public static void PptxInheritanceHighlightUsesDirectOrDefaultOwner()
    {
        const string yellow = "<a:highlight><a:srgbClr val=\"FFFF00\"/></a:highlight>";
        const string green = "<a:highlight><a:srgbClr val=\"00FF00\"/></a:highlight>";
        PdfPage flatYellow = RenderTextEffects("", yellow), flatGreen = RenderTextEffects("", green), plain = RenderTextEffects("", "");
        TestAssert.Contains("1 1 0 rg", flatYellow.Content);
        TestAssert.Contains("0 1 0 rg", flatGreen.Content);
        foreach (bool paragraphDefault in new[] { false, true })
        {
            TestAssert.Equal(flatYellow.Content, RenderTextEffects(yellow, "", paragraphDefault).Content);
            TestAssert.Equal(flatGreen.Content, RenderTextEffects(yellow, green, paragraphDefault).Content);
            // Unsupported direct highlight paint keeps its existing no-highlight
            // fallback; its presence must not resurrect an inherited color.
            TestAssert.Equal(plain.Content, RenderTextEffects(yellow,
                "<a:highlight><a:schemeClr val=\"accent1\"/></a:highlight>", paragraphDefault).Content);
        }
    }

    public static void PptxInheritanceOutlineRemovalDiffersFromAbsentOrUnsupportedPaint()
    {
        const string red = "<a:ln w=\"25400\"><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill></a:ln>";
        const string blue = "<a:ln w=\"25400\"><a:solidFill><a:srgbClr val=\"0000FF\"/></a:solidFill></a:ln>";
        PdfPage flatRed = RenderTextEffects("", red), flatBlue = RenderTextEffects("", blue), plain = RenderTextEffects("", "");
        TestAssert.Contains("1 0 0 RG", flatRed.Content);
        TestAssert.Contains("0 0 1 RG", flatBlue.Content);
        foreach (bool paragraphDefault in new[] { false, true })
        {
            TestAssert.Equal(flatRed.Content, RenderTextEffects(red, "", paragraphDefault).Content);
            TestAssert.Equal(plain.Content, RenderTextEffects(red, "<a:ln><a:noFill/></a:ln>", paragraphDefault).Content);
            TestAssert.Equal(flatBlue.Content, RenderTextEffects(red, blue, paragraphDefault).Content);
            TestAssert.Equal(flatRed.Content, RenderTextEffects(red, "<a:ln><a:gradFill/></a:ln>", paragraphDefault).Content);
        }
    }

    private static PdfPage RenderTextEffects(string defaults, string direct, bool paragraphDefault = false)
    {
        const string font = "<a:defRPr sz=\"2400\"><a:latin typeface=\"DejaVu Sans\"/>";
        var parts = TextDefaultParts(font + (paragraphDefault ? "" : defaults) + "</a:defRPr>", "", direct);
        if (paragraphDefault)
            Edit(parts, "ppt/slides/slide1.xml", xml => {
                foreach (var paragraph in xml.Descendants(A + "p"))
                    paragraph.AddFirst(XElement.Parse($"<a:pPr xmlns:a='{A}'><a:defRPr sz='2400'><a:latin typeface='DejaVu Sans'/>{defaults}</a:defRPr></a:pPr>"));
            });
        PdfPage? result = null;
        WithDeck(parts, (document, package) => result = Render(document, package).Single());
        PdfPage page = result!;
        string text = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).Select(show => show.Text));
        TestAssert.Contains("Shape default", text);
        TestAssert.Contains("Cell default", text);
        PdfContentValidator.ValidatePage(page, 0, default);
        return page;
    }
}
