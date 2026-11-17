using System.Xml.Linq;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxTextHyperlinksTests
{
    public static void PptxRunHyperlinksKeepVisibleHorizontalAndSlideCrops()
    {
        foreach (var (left, clips, missing) in new[] {
            (72d, true, false), (72d, true, true),
            (72d, false, false), (72d, false, true),
            (715d, true, false), (750d, true, false) })
        {
            var page = RenderCroppedRunLink(left, clips, "", missing);
            if (left >= 720d)
            {
                TestAssert.Equal(0, page.Annotations.Count);
                continue;
            }
            TestAssert.True(page.Annotations.Count > 0, "Visible linked text retains navigation.");
            if (!clips)
            {
                TestAssert.True(page.Annotations.Max(link => link.X + link.Width) > left + 18d,
                    "Authored horizontal overflow retains the links on visible overflow text.");
                continue;
            }
            var annotation = page.Annotations.Single();
            TestAssert.True(Math.Abs(annotation.X - left) < .001d &&
                Math.Abs(annotation.Width - Math.Min(18d, 720d - left)) < .001d,
                "Navigation ends at the authored horizontal clip and physical slide edge.");
            TestAssert.Equal("https://example.invalid/cropped", annotation.Uri);
        }
    }

    public static void PptxTransformedRunHyperlinksStayInsidePhysicalPaintClips()
    {
        foreach (string transform in new[] { "flipH=\"1\"", "flipV=\"1\"", "rot=\"2700000\"" })
        {
            var page = RenderCroppedRunLink(72d, true, transform, missing: false);
            var frames = DocxNoteClippingTests.ReadPhysicalPaintFrames(page)
                .Where(frame => frame.Block.StartsWith("BT", StringComparison.Ordinal)).ToArray();
            TestAssert.True(frames.Length > 0 && page.Annotations.Count > 0, "The transformed crop keeps visible paint and a link.");
            foreach (var link in page.Annotations)
                TestAssert.True(frames.Any(frame => link.X >= frame.Clip.X - .001d && link.Y >= frame.Clip.Y - .001d &&
                    link.X + link.Width <= frame.Clip.X + frame.Clip.Width + .001d &&
                    link.Y + link.Height <= frame.Clip.Y + frame.Clip.Height + .001d),
                    "Each navigation rectangle stays inside an independently read physical paint clip.");
        }
    }

    private static PdfPage RenderCroppedRunLink(double left, bool clips, string transformAttributes, bool missing)
    {
        var parts = RunLinkPackageParts(
            """<a:p><a:r><a:rPr sz="1800"><a:hlinkClick r:id="rIdLink"/></a:rPr><a:t>LONG-LINK</a:t></a:r></a:p>""",
            """<Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/cropped" TargetMode="External"/>""",
            $"""<a:bodyPr horzOverflow="{(clips ? "clip" : "overflow")}" vertOverflow="overflow" wrap="none" lIns="0" rIns="0" tIns="0" bIns="0"/>""");
        var xml = XDocument.Parse(parts["ppt/slides/slide1.xml"]);
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var transform = xml.Descendants(a + "xfrm").Single();
        transform.Element(a + "off")!.SetAttributeValue("x", (long)(left * 12700d));
        transform.Element(a + "ext")!.SetAttributeValue("cx", 228600);
        foreach (var attribute in XElement.Parse("<x " + transformAttributes + "/>").Attributes())
            transform.SetAttributeValue(attribute.Name, attribute.Value);
        parts["ppt/slides/slide1.xml"] = xml.ToString();
        using var input = TestFixtures.CreateZipPackage(parts);
        using var package = OoxPackage.Open(input, default);
        var document = new PptxReader().Read(package, default);
        IFontResolver resolver = missing ? new MissingTestFontResolver() : new PortableRenderingTests.PinnedResolver();
        return new PptxRenderer(resolver).RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
    }
}
