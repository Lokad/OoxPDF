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
            // Unresolved direct highlight paint keeps its no-highlight
            // fallback; its presence must not resurrect an inherited color.
            TestAssert.Equal(plain.Content, RenderTextEffects(yellow,
                "<a:highlight><a:schemeClr val=\"unresolved\"/></a:highlight>", paragraphDefault).Content);
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
        var parts = TextEffectParts(defaults, direct, paragraphDefault);
        PdfPage? result = null;
        WithDeck(parts, (document, package) => result = Render(document, package).Single());
        PdfPage page = result ?? throw new InvalidOperationException("The text-effect control produced no page.");
        string text = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).Select(show => show.Text));
        TestAssert.Contains("Shape default", text);
        TestAssert.Contains("Cell default", text);
        PdfContentValidator.ValidatePage(page, 0, default);
        return page;
    }

    private static Dictionary<string, byte[]> TextEffectParts(string defaults, string direct, bool paragraphDefault = false)
    {
        const string font = "<a:defRPr sz=\"2400\"><a:latin typeface=\"DejaVu Sans\"/>";
        var parts = TextDefaultParts(font + (paragraphDefault ? "" : defaults) + "</a:defRPr>", "", direct);
        if (paragraphDefault)
            Edit(parts, "ppt/slides/slide1.xml", xml => {
                foreach (var paragraph in xml.Descendants(A + "p"))
                    paragraph.AddFirst(XElement.Parse($"<a:pPr xmlns:a='{A}'><a:defRPr sz='2400'><a:latin typeface='DejaVu Sans'/>{defaults}</a:defRPr></a:pPr>"));
            });
        // A declared palette makes scheme-color expectations independent from
        // renderer defaults; the theme-less fixture formerly left accent1 unresolved.
        parts["ppt/theme/paint.xml"] = TestFixtures.Utf8($"<a:theme xmlns:a='{A}' name='Paint'><a:themeElements><a:clrScheme name='Paint'><a:accent1><a:srgbClr val='FFFF00'/></a:accent1></a:clrScheme></a:themeElements></a:theme>");
        Edit(parts, "ppt/_rels/presentation.xml.rels", xml => xml.Root!.Add(new XElement(xml.Root.Name.Namespace + "Relationship",
            new XAttribute("Id", "theme"), new XAttribute("Type", R.NamespaceName + "/theme"), new XAttribute("Target", "theme/paint.xml"))));
        Edit(parts, "[Content_Types].xml", xml => xml.Root!.Add(new XElement(xml.Root.Name.Namespace + "Override",
            new XAttribute("PartName", "/ppt/theme/paint.xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.theme+xml"))));
        return parts;
    }
}
