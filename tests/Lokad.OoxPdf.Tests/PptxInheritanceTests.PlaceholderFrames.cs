using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

using static Lokad.OoxPdf.Tests.PptxInheritanceFixture;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxInheritanceTests
{
    public static void PptxPlaceholderFramesKeepInkAndLinksAcrossEmptyOwners()
    {
        foreach (bool grouped in new[] { false, true })
        foreach (int frameOwner in new[] { 0, 1, 2 })
        {
            int x = 40 + frameOwner * 40;
            string frame = PlaceholderFrame(x);
            PdfPage control = RenderPlaceholderFrames(
                PlaceholderShape("1", PlaceholderFrame(40)), PlaceholderShape("1", PlaceholderFrame(80)),
                PlaceholderShape("1", frame, "INHERITED", link: true), grouped);
            string?[] emptyOwners = frameOwner == 2 ? [""]
                : [null, "", "<a:xfrm><a:off x=\"0\" y=\"0\"/></a:xfrm>"];
            foreach (string? empty in emptyOwners)
            {
                PdfPage page = RenderPlaceholderFrames(
                    PlaceholderShape("1", PlaceholderFrame(40)),
                    PlaceholderShape("1", frameOwner == 0 ? empty : PlaceholderFrame(80)),
                    PlaceholderShape("1", frameOwner == 2 ? frame : empty, "INHERITED", link: true), grouped);
                TestAssert.Equal("INHERITED".Length, PortableRenderingTests.ReadPaintedSourceLabel(page, "INHERITED").Count);
                TestAssert.Equal(control.Content, page.Content);
                var annotation = page.Annotations.Single();
                TestAssert.Equal((double)x + (grouped ? 100 : 0), annotation.X);
                TestAssert.Equal(grouped ? 320d : 420d, annotation.Y);
                TestAssert.Equal(300d, annotation.Width);
                TestAssert.Equal(80d, annotation.Height);
                TestAssert.Equal("https://example.test/placeholder", annotation.Uri);
                PdfContentValidator.ValidatePage(page, 0, default);
            }
        }

        // A truly absent frame still cannot be placed. A complete zero-size
        // declaration stays an explicit frame rather than falling through.
        WithDeck(PlaceholderFrameParts(PlaceholderShape("1", ""), PlaceholderShape("1", ""),
            PlaceholderShape("1", "", "INHERITED")), (document, package) =>
                TestAssert.Equal(0, PptxRenderer.InspectTextFrameModels(document, package, 0).Count));
        WithDeck(PlaceholderFrameParts(PlaceholderShape("1", PlaceholderFrame(40)), PlaceholderShape("1", ""),
            PlaceholderShape("1", PlaceholderFrame(120, 0, 0), "INHERITED")), (document, package) =>
            {
                var frame = PptxRenderer.InspectTextFrameModels(document, package, 0).Single();
                TestAssert.Equal(120d, frame.FrameX);
                TestAssert.Equal(0d, frame.FrameWidth);
                TestAssert.Equal(0d, frame.FrameHeight);
                TestAssert.True(!frame.UsesInheritedShapeBounds, "A complete direct zero-size frame retains its owner.");
            });

        string master = PlaceholderShape("1", PlaceholderFrame(40));
        string slide = PlaceholderShape("1", "", "INHERITED", link: true);
        string layoutExtension = "<p:extLst><p:ext uri='{5D154C70-1063-4818-8DFE-C915C125A6AD}'>" +
            "<case:payload xmlns:case='urn:ooxpdf:test'>" + PlaceholderShape("1", PlaceholderFrame(270)) +
            "</case:payload></p:ext></p:extLst>";
        PdfPage noLayoutMatch = RenderPlaceholderFrames(master, "", slide);
        PdfPage extension = RenderPlaceholderFrames(master, layoutExtension, slide);
        TestAssert.Equal(noLayoutMatch.Content, extension.Content);
        TestAssert.Equal(noLayoutMatch.Annotations.Single(), extension.Annotations.Single());
        TestAssert.Equal(40d, extension.Annotations.Single().X);
    }

    public static void PptxPlaceholderIndicesKeepNumericOwnersAndExplicitFallbacks()
    {
        foreach (string index in new[] { "2", "4000000000" })
        {
            string peers = PlaceholderShape("1", PlaceholderFrame(40, 12, 6), color: "0000FF") +
                PlaceholderShape(index, PlaceholderFrame(40), id: 3, color: "FF0000");
            PdfPage RenderIndex(string? lexical) => RenderPlaceholderFrames("", peers,
                PlaceholderShape(lexical, "", "CROSSING"));
            PdfPage canonical = RenderIndex(index);
            TestAssert.Equal(8, PortableRenderingTests.ReadPaintedSourceLabel(canonical, "CROSSING").Count);
            TestAssert.Contains("1 0 0 rg", canonical.Content);
            foreach (string lexical in new[] { "0" + index, "  " + index + "  ", "0000" + index })
            {
                PdfPage page = RenderIndex(lexical);
                TestAssert.Equal(canonical.Content, page.Content);
                var ink = PortableRenderingTests.ReadPaintedSourceLabel(page, "CROSSING");
                TestAssert.Equal(8, ink.Count);
                TestAssert.True(ink.All(glyph => Math.Abs(glyph.Origin.Y - ink[0].Origin.Y) < .001),
                    "Equivalent numeric IDs retain the wide peer and one line of actual ink.");
            }
            PdfPage fallback = RenderIndex("1");
            TestAssert.True(fallback.Content != canonical.Content, "Distinct peer IDs retain independent frames and styles.");
            TestAssert.Contains("0 0 1 rg", fallback.Content);
            foreach (string? lexical in new string?[] { null, "3", "bad", "-2", "+2", "4294967296" })
                TestAssert.Equal(fallback.Content, RenderIndex(lexical).Content);
        }

        // Malformed exact-spelling recovery is independent of numeric matching.
        string malformedPeers = PlaceholderShape("1", PlaceholderFrame(80)) + PlaceholderShape("bad", PlaceholderFrame(40), id: 3);
        TestAssert.Equal(RenderPlaceholderFrames("", malformedPeers, PlaceholderShape("bad", "", "CROSSING")).Content,
            RenderPlaceholderFrames("", malformedPeers, PlaceholderShape("bad", PlaceholderFrame(40), "CROSSING")).Content);
    }

    private static PdfPage RenderPlaceholderFrames(string master, string layout, string slide, bool grouped = false)
    {
        PdfPage? page = null;
        WithDeck(PlaceholderFrameParts(master, layout, slide, grouped), (document, package) => page = Render(document, package).Single());
        return page ?? throw new InvalidOperationException("The placeholder control produced no page.");
    }

    private static Dictionary<string, byte[]> PlaceholderFrameParts(string master, string layout, string slide, bool grouped = false)
    {
        string Story(string kind, string shapes) => $"<p:{kind} xmlns:p='{P}' xmlns:a='{A}' xmlns:r='{R}'><p:cSld><p:spTree>{shapes}</p:spTree></p:cSld></p:{kind}>";
        if (grouped)
            slide = "<p:grpSp><p:nvGrpSpPr><p:cNvPr id='10' name='Group'/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>" +
                "<p:grpSpPr><a:xfrm><a:off x='1270000' y='1270000'/><a:ext cx='9144000' cy='6858000'/>" +
                "<a:chOff x='0' y='0'/><a:chExt cx='9144000' cy='6858000'/></a:xfrm></p:grpSpPr>" + slide + "</p:grpSp>";
        return new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(PptxTests.BasicContentTypes().Replace("</Types>",
                "<Override PartName='/ppt/slideLayouts/slideLayout1.xml' ContentType='application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'/>" +
                "<Override PartName='/ppt/slideMasters/slideMaster1.xml' ContentType='application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml'/></Types>", StringComparison.Ordinal)),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8(Story("sld", slide)),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8(Relationships(
                "<Relationship Id='layout' Type='" + R + "/slideLayout' Target='../slideLayouts/slideLayout1.xml'/>" +
                "<Relationship Id='link' Type='" + R + "/hyperlink' Target='https://example.test/placeholder' TargetMode='External'/>")),
            ["ppt/slideLayouts/slideLayout1.xml"] = TestFixtures.Utf8(Story("sldLayout", layout)),
            ["ppt/slideLayouts/_rels/slideLayout1.xml.rels"] = TestFixtures.Utf8(Relationships(
                "<Relationship Id='master' Type='" + R + "/slideMaster' Target='../slideMasters/slideMaster1.xml'/>")),
            ["ppt/slideMasters/slideMaster1.xml"] = TestFixtures.Utf8(Story("sldMaster", master)),
        };

        static string Relationships(string entries) => "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>" + entries + "</Relationships>";
    }

    private static string PlaceholderFrame(int x, int width = 300, int height = 80) =>
        $"<a:xfrm><a:off x='{x * 12700}' y='508000'/><a:ext cx='{width * 12700}' cy='{height * 12700}'/></a:xfrm>";

    private static string PlaceholderShape(string? index, string? frame, string text = "", int id = 2, bool link = false, string color = "000000")
    {
        string phIndex = index is null ? "" : " idx='" + index + "'";
        string properties = frame is null ? "" : "<p:spPr>" + frame + "<a:prstGeom prst='rect'/><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>";
        string click = link ? "<a:hlinkClick r:id='link'/>" : "";
        // The empty inherited paragraph owns defaults; the actual source run
        // does not repeat them, so matching affects typography as well as width.
        string listStyle = text.Length == 0 ? "<a:lvl1pPr><a:defRPr sz='1800'><a:solidFill><a:srgbClr val='" + color +
            "'/></a:solidFill><a:latin typeface='DejaVu Sans'/></a:defRPr></a:lvl1pPr>" : "";
        return $"<p:sp><p:nvSpPr><p:cNvPr id='{id}' name='Placeholder'>{click}</p:cNvPr><p:cNvSpPr/><p:nvPr><p:ph type='body'{phIndex}/></p:nvPr></p:nvSpPr>" +
            properties + "<p:txBody><a:bodyPr lIns='0' rIns='0' tIns='0' bIns='0'/><a:lstStyle>" + listStyle +
            "</a:lstStyle><a:p><a:r><a:t>" + text + "</a:t></a:r></a:p></p:txBody></p:sp>";
    }
}
