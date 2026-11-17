using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests
{
    internal static class PptxPlaceholderBoundsTests
    {
        public static void PlaceholderBoundsInheritMasterThroughEmptyLayoutProperties()
        {
            AssertBounds("<p:spPr/>", "<p:spPr/>", 72d, true);
        }

        public static void PlaceholderBoundsInheritMasterThroughMissingLayoutProperties()
        {
            AssertBounds(string.Empty, "<p:spPr/>", 72d, true);
        }

        public static void PlaceholderBoundsInheritMasterThroughLayoutFill()
        {
            AssertBounds("<p:spPr><a:noFill/></p:spPr>", "<p:spPr/>", 72d, true);
        }

        public static void PlaceholderBoundsPreferLayoutOverMaster()
        {
            AssertBounds(ShapeProperties(1828800), "<p:spPr/>", 144d, true);
        }

        public static void PlaceholderBoundsPreferSlideOverLayoutAndMaster()
        {
            AssertBounds(ShapeProperties(1828800), ShapeProperties(2743200), 216d, false);
        }

        public static void PlaceholderBoundsStayAbsentWithoutAnyGeometry()
        {
            IReadOnlyList<PptxTextFrameModelSnapshot> frames = ReadFrames("<p:spPr/>", "<p:spPr/>", "<p:spPr/>");
            TestAssert.Equal(0, frames.Count);
        }

        private static void AssertBounds(string layoutProperties, string slideProperties, double expectedX, bool inherited)
        {
            PptxTextFrameModelSnapshot frame = ReadFrames(ShapeProperties(914400), layoutProperties, slideProperties).Single();
            TestAssert.Equal(expectedX, frame.FrameX);
            TestAssert.Equal(72d, frame.FrameY);
            TestAssert.Equal(432d, frame.FrameWidth);
            TestAssert.Equal(72d, frame.FrameHeight);
            TestAssert.Equal(inherited, frame.UsesInheritedShapeBounds);
            TestAssert.Equal(2, frame.InheritedPlaceholderCount);
            TestAssert.Equal("Slide title", frame.Paragraphs.Single().Runs.Single().Text);
        }

        private static IReadOnlyList<PptxTextFrameModelSnapshot> ReadFrames(string masterProperties, string layoutProperties, string slideProperties)
        {
            using MemoryStream stream = TestFixtures.CreateZipPackage(new Dictionary<string, string>
            {
                ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
                ["_rels/.rels"] = PptxTests.PackageRelationship(),
                ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
                ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
                ["ppt/slides/_rels/slide1.xml.rels"] = """
                    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
                    </Relationships>
                    """,
                ["ppt/slideLayouts/_rels/slideLayout1.xml.rels"] = """
                    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
                    </Relationships>
                    """,
                ["ppt/slideMasters/slideMaster1.xml"] = SlideXml("sldMaster", masterProperties, "Master title"),
                ["ppt/slideLayouts/slideLayout1.xml"] = SlideXml("sldLayout", layoutProperties, "Layout title"),
                ["ppt/slides/slide1.xml"] = SlideXml("sld", slideProperties, "Slide title"),
            });
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
            return PptxRenderer.InspectTextFrameModels(document, package, 0);
        }

        private static string ShapeProperties(long x)
        {
            return $"<p:spPr><a:xfrm><a:off x=\"{x}\" y=\"914400\"/><a:ext cx=\"5486400\" cy=\"914400\"/></a:xfrm></p:spPr>";
        }

        private static string SlideXml(string elementName, string properties, string text)
        {
            return $$"""
                <p:{{elementName}} xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
                    {{properties}}
                    <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400"/><a:t>{{text}}</a:t></a:r></a:p></p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:{{elementName}}>
                """;
        }
    }
}
