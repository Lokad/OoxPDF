using System.Reflection;
using System.Xml.Linq;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxNumberingTests
{
    public static void PptxNumberingContinuesEqualSettingsAndRestartsChangedSettings()
    {
        AssertLabels(["9.", "10.", "11."], Paragraph(9) + Paragraph(9) + Paragraph(9));
        AssertLabels(["9.", "1.", "2."], Paragraph(9) + Paragraph() + Paragraph());
        AssertLabels(["1.", "2.", "10.", "1."], Paragraph(1) + Paragraph() + Paragraph(10) + Paragraph());
        AssertLabels(["1.", "a.", "1."], Paragraph() + Paragraph(kind: "alphaLcPeriod") + Paragraph());
        AssertLabels(["aa.", "ab.", "ac."], Paragraph(27, "alphaLcPeriod") + Paragraph(27, "alphaLcPeriod") + Paragraph(27, "alphaLcPeriod"));
        AssertLabels(["iv.", "v.", "vi."], Paragraph(4, "romanLcPeriod") + Paragraph(4, "romanLcPeriod") + Paragraph(4, "romanLcPeriod"));
        AssertLabels(["(9)", "(1)", "(2)"], Paragraph(9, "arabicParenBoth") + Paragraph(kind: "arabicParenBoth") + Paragraph(kind: "arabicParenBoth"));
    }

    public static void PptxNumberingKeepsLevelSequencesAndInheritedSettings()
    {
        foreach (int child in new[] { 1, 2, 8 })
        {
            AssertLabels(["1.", "1.", "2.", "2.", "1."], Paragraph() + Paragraph(level: child) + Paragraph(level: child) + Paragraph() + Paragraph(level: child));
        }
        AssertLabels(["9.", "1.", "1."], Paragraph(9) + Paragraph(level: 1) + Paragraph());
        string inherited = "<a:lvl1pPr><a:buAutoNum type=\"arabicPeriod\" startAt=\"9\"/></a:lvl1pPr>";
        AssertLabels(["9.", "10.", "11."], Paragraph(numbering: false) + Paragraph(numbering: false) + Paragraph(numbering: false), listStyle: inherited);
        AssertLabels(["1.", "1."], Paragraph() + Paragraph(numbering: false, marker: "<a:buNone/>") + Paragraph());
    }

    public static void PptxNumberingEmptyParagraphChangesSettingsWithoutAdvancing()
    {
        AssertLabels(["1.", "2."], Paragraph() + Paragraph(text: "") + Paragraph());
        AssertLabels(["9.", "10."], Paragraph(9) + Paragraph(9, text: "") + Paragraph(9));
        AssertLabels(["1.", "1."], Paragraph() + Paragraph(9, text: "") + Paragraph());
        AssertLabels(["9.", "1."], Paragraph(9) + Paragraph(text: "") + Paragraph());
        AssertLabels(["1.", "2.", "3."], Paragraph() + Paragraph(text: " ") + Paragraph());
    }

    public static void PptxNumberingLabelUsesBodyFontAndNominalAdvances()
    {
        var spans = Read(Paragraph(28, "alphaUcPeriod"));
        var label = spans.Single(s => s.Run.Text == "AB.");
        var body = spans.Single(s => s.Run.Text == "PublicItem");
        TestAssert.Equal("TestFont", label.Run.FontFamily);
        TestAssert.True(!label.Run.KerningEnabled, "Automatic labels must use nominal glyph advances.");
        // Synthetic AB has a -50-unit pair; nominal 600+620+500 at 24pt is 41.28pt.
        TestAssert.True(Math.Abs(body.Run.X - (90d + 41.28d)) < 0.001d, "First body fragment must start after the complete unkerned label.");
        foreach (double spacing in new[] { 1.5d, -0.5d })
        {
            spans = Read(Paragraph(100, spacing: spacing));
            body = spans.Single(s => s.Run.Text == "PublicItem");
            TestAssert.True(Math.Abs(body.Run.X - (90d + 48d + 4d * spacing)) < 0.001d,
                "The label advance includes the authored trailing character spacing.");
        }
    }

    public static void PptxNumberingRetainsContinuationIndentAndSceneAgreement()
    {
        var manual = Read(Paragraph(100, suffix: "<a:br/><a:r><a:rPr sz=\"2400\"><a:latin typeface=\"TestFont\"/></a:rPr><a:t>Continuation</a:t></a:r>"));
        TestAssert.Equal(138d, manual.Single(s => s.Run.Text == "PublicItem").Run.X);
        TestAssert.Equal(108d, manual.Single(s => s.Run.Text == "Continuation").Run.X);
        var wrapped = Read(Paragraph(100, text: "AAAA BBBB CCCC"), width: 180);
        var words = wrapped.Where(s => s.Run.Text.Contains('A') || s.Run.Text.Contains('B') || s.Run.Text.Contains('C')).ToArray();
        TestAssert.True(words.Length >= 2 && words[0].Run.X == 138d && words.Skip(1).Any(s => s.Run.X == 108d),
            "Wrapping must retain the authored continuation margin after the pushed first fragment.");
    }

    public static void PptxNumberingRetainsUnqualifiedBulletAndFrameFallbacks()
    {
        foreach (string marker in new[] { "<a:buSzPct val=\"150000\"/>", "<a:buFont typeface=\"TestFont\"/>" })
        {
            AssertLabels(["9.", "10.", "11."], Paragraph(9, marker: marker) + Paragraph(marker: marker) + Paragraph(marker: marker));
            TestAssert.Equal(108d, Read(Paragraph(100, marker: marker)).Single(s => s.Run.Text == "PublicItem").Run.X);
        }
        AssertLabels(["9.", "10.", "11."], Paragraph(9, "arabicPlain") + Paragraph(kind: "arabicPlain") + Paragraph(kind: "arabicPlain"));
        TestAssert.Equal(108d, Read(Paragraph(100, hanging: 18)).Single(s => s.Run.Text == "PublicItem").Run.X);
        TestAssert.Equal(108d, Read(Paragraph(100), body: "wrap=\"none\" numCol=\"2\"").Single(s => s.Run.Text == "PublicItem").Run.X);
        TestAssert.Equal(108d, Read(Paragraph(100), autofit: "<a:normAutofit/>").Single(s => s.Run.Text == "PublicItem").Run.X);
        TestAssert.Equal(108d, Read(Paragraph(100), shape: "flipH=\"1\"").Single(s => s.Run.Text == "PublicItem").Run.X);
        TestAssert.Equal(108d, Read(Paragraph(100), shape: "flipV=\"1\"").Single(s => s.Run.Text == "PublicItem").Run.X);
    }

    private static string Paragraph(int? start = null, string kind = "arabicPeriod", int level = 0,
        string text = "PublicItem", bool numbering = true, string marker = "", string suffix = "", double spacing = 0d, int hanging = -18)
    {
        string startAttribute = start.HasValue ? $" startAt=\"{start.Value}\"" : "";
        string number = numbering ? $"<a:buAutoNum type=\"{kind}\"{startAttribute}/>" : "";
        int spc = (int)(spacing * 100d);
        return $"<a:p><a:pPr algn=\"l\" lvl=\"{level}\" marL=\"457200\" indent=\"{hanging * 12700}\">{marker}{number}</a:pPr><a:r><a:rPr sz=\"2400\" kern=\"1\" spc=\"{spc}\"><a:latin typeface=\"TestFont\"/></a:rPr><a:t>{text}</a:t></a:r>{suffix}<a:endParaRPr sz=\"2400\"/></a:p>";
    }

    private static void AssertLabels(string[] expected, string paragraphs, string listStyle = "")
    {
        string[] actual = Read(paragraphs, listStyle: listStyle).Where(s => s.Run.Text != "PublicItem" && !string.IsNullOrWhiteSpace(s.Run.Text)).Select(s => s.Run.Text).ToArray();
        TestAssert.True(expected.SequenceEqual(actual), $"Expected labels {string.Join(',', expected)}, got {string.Join(',', actual)}.");
    }

    private static IReadOnlyList<PptxRenderer.PptxPositionedTextSpan> Read(string paragraphs,
        string body = "wrap=\"square\"", string autofit = "<a:noAutofit/>", string listStyle = "", int width = 500, string shape = "")
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = $$"""
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Numbering"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
                    <p:spPr><a:xfrm {{shape}}><a:off x="914400" y="914400"/><a:ext cx="{{width * 12700}}" cy="4445000"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody><a:bodyPr {{body}} lIns="0" rIns="0" tIns="0" bIns="0">{{autofit}}</a:bodyPr><a:lstStyle>{{listStyle}}</a:lstStyle>{{paragraphs}}</p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        try
        {
            using FileStream stream = File.OpenRead(input);
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
            PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
            PptxSceneNode node = scene.Slides[0].SlideNodes[0];
            var resolver = new PresentationFontResolver(new NumberingFontResolver());
            MethodInfo method = typeof(PptxRenderer).GetMethods(BindingFlags.NonPublic | BindingFlags.Static)
                .Single(m => m.Name == "ReadTextSpansForShape" && m.GetParameters().Length == 9);
            var xml = (IReadOnlyList<PptxRenderer.PptxPositionedTextSpan>)method.Invoke(null,
                [node.Source, document, scene.Theme, scene.Slides[0].SlideColorMap, 1, false, Array.Empty<XDocument>(), resolver, CancellationToken.None])!;
            var fed = PptxRenderer.BuildSceneFedTextSpans(node, document, scene.Theme, scene.Slides[0].SlideColorMap, 1, false, [], resolver, CancellationToken.None);
            TestAssert.Equal(xml.Count, fed.Count);
            for (int i = 0; i < xml.Count; i++)
            {
                TestAssert.Equal(xml[i].Run.Text, fed[i].Run.Text);
                TestAssert.Equal(xml[i].Run.X, fed[i].Run.X);
                TestAssert.Equal(xml[i].Run.Y, fed[i].Run.Y);
                TestAssert.Equal(xml[i].Run.FontFamily, fed[i].Run.FontFamily);
                TestAssert.Equal(xml[i].Run.KerningEnabled, fed[i].Run.KerningEnabled);
            }
            return xml;
        }
        finally { File.Delete(input); }
    }

    private sealed class NumberingFontResolver : IFontResolver
    {
        private readonly MemoryFontProgramSource source = new("numbering-test", TestFontBuilder.CreateTestFont());
        public FontFaceResolution Resolve(FontRequest request) => new(request.FamilyName, "TestFont",
            new FontStyleKey(false, false, 400, 0, false), source, IsFallback: false);
    }
}
