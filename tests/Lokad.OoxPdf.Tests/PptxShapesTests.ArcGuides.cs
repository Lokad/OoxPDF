using System.Text;
using System.Text.RegularExpressions;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxShapesTests
{
    public static void PptxPresetArcClampsFiniteGuidesBeforeTracing()
    {
        foreach (string tail in new[] { "none", "stealth" })
        foreach ((string start, string end, string expectedStart, string expectedEnd) in new[]
        {
            ("0", "21600000", "0", "21599999"),
            ("-5400000", "5400000", "0", "5400000"),
            ("0", "1e300", "0", "21599999"),
            ("1e300", "0", "21599999", "0"),
            ("-1e300", "5400000", "0", "5400000"),
            ("5400000", "-1e300", "5400000", "0"),
            ("-1.7976931348623157e308", "1.7976931348623157e308", "0", "21599999")
        })
        {
            string actual = RenderArcGuides(start, end, tail);
            string expected = RenderArcGuides(expectedStart, expectedEnd, tail);
            TestAssert.True(expected == actual, $"Arc {start}/{end}/{tail} must use the preset's pinned angles.");
        }
    }

    public static void PptxPresetArcNonfiniteGuidesUseDefaults()
    {
        foreach (string tail in new[] { "none", "stealth" })
        {
            string expected = RenderArcGuides("16200000", "0", tail);
            foreach (string invalid in new[] { "NaN", "Infinity", "-Infinity", "1e999", "-1e999", "not-a-number" })
            {
                TestAssert.True(expected == RenderArcGuides(invalid, "0", tail), "Invalid start guide must use its default.");
                TestAssert.True(expected == RenderArcGuides("16200000", invalid, tail), "Invalid end guide must use its default.");
            }
        }
    }

    public static void PptxPresetArcEqualAndReversedGuidesKeepTheirSweep()
    {
        string fullTurn = RenderArcGuides("0", "0", "none");
        TestAssert.True(Regex.Matches(fullTurn, @" c\b").Count == 4, "Equal angles describe one complete turn.");
        string reversed = RenderArcGuides("16200000", "5400000", "none");
        TestAssert.True(Regex.Matches(reversed, @" c\b").Count == 2, "Reversed angles cross zero once.");
    }

    private static string RenderArcGuides(string start, string end, string tail)
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = $$"""
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp><p:spPr>
                    <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="2743200"/></a:xfrm>
                    <a:prstGeom prst="arc"><a:avLst><a:gd name="adj1" fmla="val {{start}}"/><a:gd name="adj2" fmla="val {{end}}"/></a:avLst></a:prstGeom>
                    <a:noFill/><a:ln w="19050"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:tailEnd type="{{tail}}"/></a:ln>
                  </p:spPr></p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = input + ".pdf";
        try
        {
            OoxPdfConverter.Convert(input, output);
            return File.ReadAllText(output, Encoding.ASCII);
        }
        finally
        {
            File.Delete(input);
            File.Delete(output);
        }
    }
}
