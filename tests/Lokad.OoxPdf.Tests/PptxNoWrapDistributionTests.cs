using System.Reflection;
using System.Xml.Linq;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxNoWrapDistributionTests
{
    public static void PptxNoWrapDistributionKeepsNominalTextAndActiveKerning()
    {
        foreach (string text in new[] { "ABAB", "ABAB ABAB", "ABAB  ABAB", "SKU-123 / A+B, C.D" })
        {
            foreach (string kern in new[] { "1", "10000" })
            {
                TestAssert.True(Render(Paragraph("dist", text, kern)).SequenceEqual(Render(Paragraph("l", text, kern))),
                    "Admitted distributed no-wrap text must match nominal left placement, retaining the authored positive kerning rule.");
            }
        }
        TestAssert.True(!Render(Paragraph("dist", "ABAB", "1")).SequenceEqual(Render(Paragraph("dist", "ABAB", "10000"))),
            "Disabling distribution must preserve the active AB kerning pair.");
    }

    public static void PptxNoWrapDistributionHonorsInheritedAlignmentAndMixedParagraphs()
    {
        string inherited = "<a:lvl1pPr algn=\"dist\"/>";
        TestAssert.True(Render(Paragraph(null), listStyle: inherited).SequenceEqual(Render(Paragraph("dist"))),
            "Inherited distributed alignment must share the resolved paragraph admission.");
        TestAssert.True(Render(Paragraph("l"), listStyle: inherited).SequenceEqual(Render(Paragraph("l"))),
            "Direct left alignment must override inherited distributed alignment.");
        TestAssert.True(Render(Paragraph("dist") + Paragraph("l")).SequenceEqual(Render(Paragraph("l") + Paragraph("l"))),
            "Admission must apply independently within mixed paragraphs.");
    }

    public static void PptxNoWrapDistributionPreservesXmlSceneGlyphPlacement()
    {
        string input = Package(Paragraph("dist", "ABAB ABAB"));
        try
        {
            using FileStream stream = File.OpenRead(input);
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
            PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
            PptxSceneNode node = scene.Slides[0].SlideNodes[0];
            var resolver = new PresentationFontResolver(new DistributionFontResolver());
            MethodInfo method = typeof(PptxRenderer).GetMethods(BindingFlags.NonPublic | BindingFlags.Static)
                .Single(m => m.Name == "ReadTextSpansForShape" && m.GetParameters().Length == 9);
            var xml = (IReadOnlyList<PptxRenderer.PptxPositionedTextSpan>)method.Invoke(null,
                [node.Source, document, scene.Theme, scene.Slides[0].SlideColorMap, 1, false, Array.Empty<XDocument>(), resolver, CancellationToken.None])!;
            var fed = PptxRenderer.BuildSceneFedTextSpans(node, document, scene.Theme, scene.Slides[0].SlideColorMap, 1, false, [], resolver, CancellationToken.None);
            TestAssert.Equal(xml.Count, fed.Count);
            TestAssert.True(xml.Any(x => x.Run.Text.Contains("ABAB", StringComparison.Ordinal)), "Nominal text must retain word spans.");
            for (int i = 0; i < xml.Count; i++)
            {
                TestAssert.Equal(xml[i].Run.Text, fed[i].Run.Text);
                TestAssert.Equal(xml[i].Run.X, fed[i].Run.X);
                TestAssert.Equal(xml[i].Run.Y, fed[i].Run.Y);
                TestAssert.Equal(xml[i].Run.KerningEnabled, fed[i].Run.KerningEnabled);
            }
        }
        finally { File.Delete(input); }
    }

    public static void PptxNoWrapDistributionRetainsExcludedGlyphDistribution()
    {
        foreach ((string body, string autofit, string shape, string marker) in new[]
        {
            ("wrap=\"square\"", "<a:noAutofit/>", "", ""),
            ("", "<a:noAutofit/>", "", ""),
            ("wrap=\"none\"", "", "", ""),
            ("wrap=\"none\"", "<a:spAutoFit/>", "", ""),
            ("wrap=\"none\"", "<a:normAutofit/>", "", ""),
            ("wrap=\"none\" numCol=\"2\"", "<a:noAutofit/>", "", ""),
            ("wrap=\"none\" vert=\"vert270\"", "<a:noAutofit/>", "", ""),
            ("wrap=\"none\"", "<a:noAutofit/>", "rot=\"900000\"", ""),
            ("wrap=\"none\"", "<a:noAutofit/>", "flipH=\"1\"", ""),
            ("wrap=\"none\"", "<a:noAutofit/>", "flipV=\"1\"", ""),
            ("wrap=\"none\"", "<a:noAutofit/>", "", "<a:buChar char=\"*\"/>"),
            ("wrap=\"none\"", "<a:noAutofit/>", "", "<a:tabLst><a:tab pos=\"914400\" algn=\"l\"/></a:tabLst>")
        })
        {
            TestAssert.True(!Render(Paragraph("dist", marker: marker), body, autofit, shape).SequenceEqual(Render(Paragraph("l", marker: marker), body, autofit, shape)),
                "Excluded settings must retain existing distributed glyph placement.");
        }
        TestAssert.True(!Render(Paragraph("dist", "ABAB\u00a0ABAB")).SequenceEqual(Render(Paragraph("l", "ABAB\u00a0ABAB"))),
            "Non-ASCII spacing must retain fallback.");
        string broken = Paragraph("dist").Replace("</a:r></a:p>", "</a:r><a:br/><a:r><a:rPr sz=\"2400\"><a:latin typeface=\"TestFont\"/></a:rPr><a:t>ABAB ABAB</a:t></a:r></a:p>", StringComparison.Ordinal);
        TestAssert.True(!Render(broken).SequenceEqual(Render(broken.Replace("algn=\"dist\"", "algn=\"l\"", StringComparison.Ordinal))),
            "Manual breaks must retain fallback.");
    }

    private static string Paragraph(string? alignment, string text = "ABAB ABAB", string kern = "1", string marker = "")
    {
        string algn = alignment is null ? "" : $" algn=\"{alignment}\"";
        return $"<a:p><a:pPr{algn}>{marker}</a:pPr><a:r><a:rPr sz=\"2400\" kern=\"{kern}\"><a:latin typeface=\"TestFont\"/></a:rPr><a:t>{text}</a:t></a:r></a:p>";
    }

    private static string Package(string paragraphs, string body = "wrap=\"none\"", string autofit = "<a:noAutofit/>", string shape = "", string listStyle = "")
    {
        return TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = $$"""
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Distribution"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
                    <p:spPr><a:xfrm {{shape}}><a:off x="914400" y="914400"/><a:ext cx="6350000" cy="4445000"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody><a:bodyPr {{body}} lIns="0" rIns="0" tIns="0" bIns="0">{{autofit}}</a:bodyPr><a:lstStyle>{{listStyle}}</a:lstStyle>{{paragraphs}}</p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
    }

    private static byte[] Render(string paragraphs, string body = "wrap=\"none\"", string autofit = "<a:noAutofit/>", string shape = "", string listStyle = "")
    {
        string input = Package(paragraphs, body, autofit, shape, listStyle);
        string output = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pdf");
        try
        {
            OoxPdfConverter.Convert(input, output, new OoxPdfOptions { FontResolver = new DistributionFontResolver() });
            return File.ReadAllBytes(output);
        }
        finally { File.Delete(input); File.Delete(output); }
    }

    private sealed class DistributionFontResolver : IFontResolver
    {
        private readonly MemoryFontProgramSource source = new("nowrap-distribution-test", TestFontBuilder.CreateTestFont());
        public FontFaceResolution Resolve(FontRequest request) => new(request.FamilyName, "TestFont",
            new FontStyleKey(false, false, 400, 0, false), source, IsFallback: false);
    }
}
