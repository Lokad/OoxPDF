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

    public static void PptxOrthogonalTextClipsKeepAuthoredSourceInkAndLinks()
    {
        foreach (string transform in new[] { "", "rot=\"5400000\"", "rot=\"10800000\"", "rot=\"16200000\"" })
        {
            var page = RenderCroppedRunLink(72d, true, transform, missing: false, text: "ABCDEF");
            // A fits in the authored 18pt crop. Following letters retain their
            // authored crop; encoded/extractable text alone cannot establish ink.
            var first = PortableRenderingTests.ReadPaintedSourceLabel(page, "A").Single();
            TestAssert.True(first.Ink.Width > 0 && first.Ink.Height > 0, "A supported source glyph survives exact rectangular placement.");
            var link = page.Annotations.Single();
            TestAssert.True(first.Ink.X >= link.X - .001d && first.Ink.Y >= link.Y - .001d &&
                first.Ink.X + first.Ink.Width <= link.X + link.Width + .001d &&
                first.Ink.Y + first.Ink.Height <= link.Y + link.Height + .001d,
                "The resulting navigation frame covers the same visible glyph.");
        }
    }

    public static void PptxVerticalClipRetainsExplicitHorizontalOverflowInkAndLinks()
    {
        const string label = "ABCDEF";
        foreach (string transform in new[] { "", "rot=\"5400000\"" })
        {
            var page = RenderCroppedRunLink(72d, false, transform, missing: false, text: label, clipVertically: true);
            var glyphs = PortableRenderingTests.ReadPaintedSourceLabel(page, label);
            TestAssert.Equal(label.Length, glyphs.Count);
            TestAssert.DoesNotContain("LOWER LINE", DocxPageFootnoteNumberingTests.DecodeText(page));
            var last = glyphs.Last().Ink;
            TestAssert.True(page.Annotations.Any(link => link.Uri == "https://example.invalid/cropped" &&
                last.X >= link.X - .001d && last.Y >= link.Y - .001d &&
                last.X + last.Width <= link.X + link.Width + .001d &&
                last.Y + last.Height <= link.Y + link.Height + .001d),
                "The horizontal overflow carries its visible link through the vertical crop and shape transform.");
        }
    }

    private static PdfPage RenderCroppedRunLink(double left, bool clips, string transformAttributes, bool missing,
        string text = "LONG-LINK", bool clipVertically = false)
    {
        var parts = RunLinkPackageParts(
            """<a:p><a:r><a:rPr sz="1800"><a:hlinkClick r:id="rIdLink"/></a:rPr><a:t>LONG-LINK</a:t></a:r></a:p>""",
            """<Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/cropped" TargetMode="External"/>""",
            $"""<a:bodyPr horzOverflow="{(clips ? "clip" : "overflow")}" vertOverflow="{(clipVertically ? "clip" : "overflow")}" wrap="none" lIns="0" rIns="0" tIns="0" bIns="0"/>""");
        var xml = XDocument.Parse(parts["ppt/slides/slide1.xml"]);
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var transform = xml.Descendants(a + "xfrm").Single();
        xml.Descendants(a + "t").Single().Value = text;
        transform.Element(a + "off")!.SetAttributeValue("x", (long)(left * 12700d));
        transform.Element(a + "ext")!.SetAttributeValue("cx", 228600);
        if (clipVertically)
        {
            transform.Element(a + "ext")!.SetAttributeValue("cy", 304800);
            xml.Descendants(a + "p").Single().AddAfterSelf(new XElement(a + "p",
                new XElement(a + "r", new XElement(a + "rPr", new XAttribute("sz", "1800")),
                    new XElement(a + "t", "LOWER LINE"))));
        }
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
