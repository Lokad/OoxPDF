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
        AssertLabels(["aa.", "bb.", "cc."], Paragraph(27, "alphaLcPeriod") + Paragraph(27, "alphaLcPeriod") + Paragraph(27, "alphaLcPeriod"));
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
        var spans = Read(Paragraph(27, "alphaUcPeriod"));
        var label = spans.Single(s => s.Run.Text == "AA.");
        var body = spans.Single(s => s.Run.Text == "PublicItem");
        TestAssert.Equal("TestFont", label.Run.FontFamily);
        TestAssert.True(!label.Run.KerningEnabled, "Automatic labels must use nominal glyph advances.");
        // Synthetic AA plus the period advances 600+600+500 units, or 40.8pt at 24pt.
        TestAssert.True(Math.Abs(body.Run.X - (90d + 40.8d)) < 0.001d, "First body fragment must start after the complete unkerned label.");
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
            AssertLabels(["9.", "1.", "2."], Paragraph(9, marker: marker) + Paragraph(marker: marker) + Paragraph(marker: marker));
            double expected = marker.Contains("buSzPct", StringComparison.Ordinal) ? 162d : 138d;
            TestAssert.Equal(expected, Read(Paragraph(100, marker: marker)).Single(s => s.Run.Text == "PublicItem").Run.X);
        }
        AssertLabels(["9.", "10.", "11."], Paragraph(9, "arabicPlain") + Paragraph(kind: "arabicPlain") + Paragraph(kind: "arabicPlain"));
        TestAssert.Equal(108d, Read(Paragraph(100, hanging: 18)).Single(s => s.Run.Text == "PublicItem").Run.X);
        TestAssert.Equal(108d, Read(Paragraph(100), body: "wrap=\"none\" numCol=\"2\"").Single(s => s.Run.Text == "PublicItem").Run.X);
        TestAssert.Equal(108d, Read(Paragraph(100), autofit: "<a:normAutofit/>").Single(s => s.Run.Text == "PublicItem").Run.X);
        TestAssert.Equal(108d, Read(Paragraph(100), shape: "flipH=\"1\"").Single(s => s.Run.Text == "PublicItem").Run.X);
        TestAssert.Equal(108d, Read(Paragraph(100), shape: "flipV=\"1\"").Single(s => s.Run.Text == "PublicItem").Run.X);
    }

    public static void PptxNumberingOverrideKeepsSettingsAndNestedSequences()
    {
        foreach (string marker in new[] { "<a:buFont typeface=\"LabelFont\"/>", "<a:buSzPct val=\"150000\"/>", "<a:buSzPts val=\"1200\"/>" })
        {
            AssertLabels(["9.", "10.", "11."], Paragraph(9, marker: marker) + Paragraph(9, marker: marker) + Paragraph(9, marker: marker));
            AssertLabels(["9.", "1.", "2."], Paragraph(9, marker: marker) + Paragraph(marker: marker) + Paragraph(marker: marker));
            AssertLabels(["1.", "1.", "2.", "2.", "1."], Paragraph(marker: marker) + Paragraph(level: 1, marker: marker) + Paragraph(level: 1, marker: marker) + Paragraph(marker: marker) + Paragraph(level: 1, marker: marker));
            AssertLabels(["aa.", "bb."], Paragraph(27, "alphaLcPeriod", marker: marker) + Paragraph(27, "alphaLcPeriod", marker: marker));
        }
    }

    public static void PptxNumberingOverrideHonorsInheritanceTransitionsAndEmptyParagraphs()
    {
        string font = "<a:buFont typeface=\"LabelFont\"/>";
        string size = "<a:buSzPct val=\"150000\"/>";
        string inherited = "<a:lvl1pPr><a:buSzPts val=\"1800\"/><a:buFont typeface=\"LabelFont\"/><a:buAutoNum type=\"arabicPeriod\" startAt=\"9\"/></a:lvl1pPr>";
        AssertLabels(["9.", "10.", "11."], Paragraph(numbering: false) + Paragraph(numbering: false) + Paragraph(numbering: false), listStyle: inherited);
        AssertLabels(["1.", "2.", "3."], Paragraph() + Paragraph(marker: font) + Paragraph(marker: size));
        AssertLabels(["9.", "10."], Paragraph(9, marker: font) + Paragraph(9, marker: font, text: "") + Paragraph(9, marker: font));
        AssertLabels(["1.", "1."], Paragraph(marker: size) + Paragraph(numbering: false, marker: "<a:buNone/>") + Paragraph(marker: size));
        string emptyPlain = "<a:p><a:pPr><a:buNone/></a:pPr><a:endParaRPr sz=\"2400\"/></a:p>";
        AssertLabels(["1.", "1."], Paragraph(marker: size) + emptyPlain + Paragraph(marker: size));
        AssertLabels(["1.", "1."], Paragraph() + emptyPlain + Paragraph());
    }

    public static void PptxNumberingOverrideFollowsOfficeLabelFontAndPlacement()
    {
        foreach ((string marker, string family, double size) in new[]
        {
            ("<a:buFont typeface=\"LabelFont\"/>", "TestFont", 24d),
            ("<a:buSzPct val=\"150000\"/>", "TestFont", 36d),
            ("<a:buSzPts val=\"1200\"/>", "TestFont", 12d),
            ("<a:buSzPct val=\"150000\"/><a:buFont typeface=\"LabelFont\"/>", "TestFont", 36d)
        })
        {
            var spans = Read(Paragraph(100, marker: marker));
            var label = spans.Single(s => s.Run.Text == "100.");
            TestAssert.Equal(family, label.Run.FontFamily);
            TestAssert.Equal(size, label.Run.FontSize);
            TestAssert.True(!label.Run.KerningEnabled, "Qualified override labels use Office nominal advances.");
            TestAssert.Equal(Math.Max(108d, 90d + 2d * size), spans.Single(s => s.Run.Text == "PublicItem").Run.X);
        }
    }

    public static void PptxNumberingOverrideMatchesOfficeAlphabeticRolloverAndCycle()
    {
        AssertLabels(["z.", "aa.", "bb."], Paragraph(26, "alphaLcPeriod") + Paragraph(26, "alphaLcPeriod") + Paragraph(26, "alphaLcPeriod"));
        AssertLabels(["ZZ.", "AAA.", "BBB."], Paragraph(52, "alphaUcPeriod") + Paragraph(52, "alphaUcPeriod") + Paragraph(52, "alphaUcPeriod"));
        AssertLabels(["zz)", "aaa)", "bbb)"], Paragraph(52, "alphaLcParenR") + Paragraph(52, "alphaLcParenR") + Paragraph(52, "alphaLcParenR"));
        AssertLabels(["ZZ)", "AAA)", "BBB)"], Paragraph(52, "alphaUcParenR") + Paragraph(52, "alphaUcParenR") + Paragraph(52, "alphaUcParenR"));
        AssertLabels([new string('Y', 30) + ".", new string('Z', 30) + ".", "A."], Paragraph(779, "alphaUcPeriod") + Paragraph(779, "alphaUcPeriod") + Paragraph(779, "alphaUcPeriod"));
        AssertLabels(["G.", "H.", "I."], Paragraph(32767, "alphaUcPeriod") + Paragraph(32767, "alphaUcPeriod") + Paragraph(32767, "alphaUcPeriod"));
    }

    public static void PptxNumberingOverrideKeepsParagraphPitchAtBodyFontSize()
    {
        double[] normal = Read(Paragraph() + Paragraph()).Where(s => s.Run.Text == "PublicItem").Select(s => s.Run.Y).ToArray();
        foreach (string marker in new[] { "<a:buSzPct val=\"150000\"/>", "<a:buSzPts val=\"3600\"/>", "<a:buSzPct val=\"125000\"/><a:buFont typeface=\"LabelFont\"/>" })
        {
            double[] styled = Read(Paragraph(marker: marker) + Paragraph(marker: marker)).Where(s => s.Run.Text == "PublicItem").Select(s => s.Run.Y).ToArray();
            TestAssert.True(normal.SequenceEqual(styled), "A larger automatic-number label must preserve the body-font paragraph pitch.");
        }
    }

    public static void PptxNumberingClearanceUsesBodyTypefaceAndAuthoredSize()
    {
        foreach ((string marker, double size) in new[]
        {
            ("<a:buFont typeface=\"LabelFont\"/>", 24d),
            ("<a:buSzPct val=\"150000\"/>", 36d),
            ("<a:buSzPts val=\"1200\"/>", 12d),
            ("<a:buSzPts val=\"3600\"/><a:buFont typeface=\"LabelFont\"/>", 36d)
        })
        {
            var spans = Read(Paragraph(100, marker: marker));
            var label = spans.Single(s => s.Run.Text == "100.").Run;
            TestAssert.Equal("TestFont", label.FontFamily);
            TestAssert.Equal(size, label.FontSize);
            TestAssert.Equal(90d, label.X);
            TestAssert.True(!label.KerningEnabled, "Office auto-number labels use nominal advances.");
            TestAssert.Equal(90d + 2d * size, spans.Single(s => s.Run.Text == "PublicItem").Run.X);
        }
    }

    public static void PptxNumberingClearanceIncludesTrailingSpacingAndNominalAdvances()
    {
        foreach (double spacing in new[] { 1.5d, -0.5d })
        {
            var spans = Read(Paragraph(100, marker: "<a:buSzPts val=\"3600\"/>", spacing: spacing));
            TestAssert.Equal(90d + 72d + 4d * spacing, spans.Single(s => s.Run.Text == "PublicItem").Run.X);
            TestAssert.True(!spans.Single(s => s.Run.Text == "100.").Run.KerningEnabled, "The label must not introduce pair kerning.");
        }
        var alpha = Read(Paragraph(27, "alphaUcPeriod", marker: "<a:buSzPct val=\"150000\"/><a:buFont typeface=\"LabelFont\"/>"));
        TestAssert.Equal(151.2d, alpha.Single(s => s.Run.Text == "PublicItem").Run.X);
        TestAssert.Equal("TestFont", alpha.Single(s => s.Run.Text == "AA.").Run.FontFamily);
    }

    public static void PptxNumberingClearanceKeepsContinuationIndentAndSceneAgreement()
    {
        string marker = "<a:buSzPts val=\"3600\"/>";
        var manual = Read(Paragraph(100, marker: marker, suffix: "<a:br/><a:r><a:rPr sz=\"2400\"><a:latin typeface=\"TestFont\"/></a:rPr><a:t>Continuation</a:t></a:r>"));
        TestAssert.Equal(162d, manual.Single(s => s.Run.Text == "PublicItem").Run.X);
        TestAssert.Equal(108d, manual.Single(s => s.Run.Text == "Continuation").Run.X);
        var wrapped = Read(Paragraph(100, marker: marker, text: "AAAA BBBB CCCC"), width: 180);
        var words = wrapped.Where(s => s.Run.Text.Contains('A') || s.Run.Text.Contains('B') || s.Run.Text.Contains('C')).ToArray();
        TestAssert.True(words.Length >= 2 && words[0].Run.X == 162d && words.Skip(1).Any(s => s.Run.X == 108d),
            "Only the first body fragment moves; wrapped continuation keeps the authored margin.");
    }

    public static void PptxNumberingClearanceRetainsExcludedFrameAndNarrowFallbacks()
    {
        string marker = "<a:buSzPts val=\"3600\"/><a:buFont typeface=\"LabelFont\"/>";
        var variants = new[]
        {
            Read(Paragraph(100, marker: marker, hanging: 18)),
            Read(Paragraph(100, marker: marker), body: "wrap=\"square\" numCol=\"2\""),
            Read(Paragraph(100, marker: marker), autofit: "<a:normAutofit/>"),
            Read(Paragraph(100, marker: marker), shape: "flipH=\"1\""),
            Read(Paragraph(100, marker: marker), shape: "flipV=\"1\""),
            Read(Paragraph(100, marker: marker), shape: "rot=\"900000\""),
            // The short label fits, but the first body fragment does not.
            Read(Paragraph(1, marker: marker), width: 56),
            Read(Paragraph(100, "arabicPlain", marker: marker)),
            Read(Paragraph(numbering: false, marker: marker + "<a:buChar char=\"*\"/>"))
        };
        foreach (var spans in variants)
        {
            var label = spans.First(s => s.Run.Text != "PublicItem").Run;
            TestAssert.Equal("LabelFont", label.FontFamily);
            TestAssert.True(label.KerningEnabled, "Excluded layouts retain their label emission rule.");
            TestAssert.Equal(108d, spans.Single(s => s.Run.Text == "PublicItem").Run.X);
        }
    }

    public static void PptxNumberingWrapKeepsWordsWithinAuthoredWidth()
    {
        string marker = "<a:buSzPts val=\"1200\"/>";
        var narrow = Read(Paragraph(1, marker: marker, text: "AA BB"), width: 104);
        var a = narrow.Single(s => s.Run.Text.Contains("AA", StringComparison.Ordinal)).Run;
        var bb = narrow.Single(s => s.Run.Text.Contains("BB", StringComparison.Ordinal)).Run;
        // AA + space + BB uses 70.56pt; the authored body width is 68pt.
        TestAssert.True(bb.Y < a.Y, "Automatic numbering must wrap the word beyond the authored edge.");
        TestAssert.Equal(108d, a.X);
        TestAssert.Equal(108d, bb.X);
        var fitting = Read(Paragraph(1, marker: marker, text: "AA BB"), width: 107);
        TestAssert.Equal(fitting.Single(s => s.Run.Text.Contains("AA", StringComparison.Ordinal)).Run.Y,
            fitting.Single(s => s.Run.Text.Contains("BB", StringComparison.Ordinal)).Run.Y);
    }

    public static void PptxNumberingWrapKeepsSplitWordsAndSceneAgreement()
    {
        string suffix = "<a:r><a:rPr sz=\"2400\" kern=\"1\"><a:latin typeface=\"TestFont\"/></a:rPr><a:t>B</a:t></a:r>";
        var split = Read(Paragraph(1, marker: "<a:buSzPts val=\"1200\"/>", text: "AA B", suffix: suffix), width: 104);
        var a = split.Single(s => s.Run.Text.Contains("AA", StringComparison.Ordinal)).Run;
        var pieces = split.Where(s => s.Run.Text.Contains('B')).Select(s => s.Run).ToArray();
        TestAssert.True(pieces.Length >= 1 && pieces.All(s => s.Y < a.Y), "A word split across runs must move together to the continuation line.");
        TestAssert.Equal(108d, pieces[0].X);
        TestAssert.Equal(2, string.Concat(pieces.Select(s => s.Text)).Count(c => c == 'B'));
        TestAssert.True(pieces.All(s => s.Y == pieces[0].Y), "Both run fragments must stay on one continuation baseline.");
    }

    public static void PptxNumberingWrapPreservesCountersPitchAndContinuationMargins()
    {
        string marker = "<a:buSzPts val=\"1200\"/>";
        var spans = Read(Paragraph(9, marker: marker, text: "AA BB") + Paragraph(9, marker: marker, text: "AA BB"), width: 104);
        string[] labels = spans.Where(s => s.Run.Text is "9." or "10.").Select(s => s.Run.Text).ToArray();
        TestAssert.True(labels.SequenceEqual(new[] { "9.", "10." }), "Wrapping must retain equal-setting counters.");
        var firstWords = spans.Where(s => s.Run.Text.Contains("AA", StringComparison.Ordinal)).Select(s => s.Run).ToArray();
        var lastWords = spans.Where(s => s.Run.Text.Contains("BB", StringComparison.Ordinal)).Select(s => s.Run).ToArray();
        TestAssert.Equal(2, firstWords.Length);
        TestAssert.Equal(2, lastWords.Length);
        TestAssert.True(Math.Abs(firstWords[0].Y - firstWords[1].Y - 57.6d) < .001d,
            "Two wrapped body-font rows must advance the next paragraph by 57.6pt.");
        for (int i = 0; i < 2; i++)
        {
            TestAssert.Equal(108d, firstWords[i].X);
            TestAssert.Equal(108d, lastWords[i].X);
            TestAssert.True(Math.Abs(firstWords[i].Y - lastWords[i].Y - 28.8d) < .001d, "Continuation pitch stays at the body font size.");
        }
    }

    public static void PptxNumberingWrapRetainsCharacterAndExcludedFallbacks()
    {
        string marker = "<a:buSzPts val=\"1200\"/>";
        string text = "AA BB";
        var variants = new[]
        {
            Read(Paragraph(numbering: false, marker: marker + "<a:buChar char=\"*\"/>", text: text), width: 104),
            Read(Paragraph(1, marker: marker, text: text, hanging: 18), width: 104),
            Read(Paragraph(1, marker: marker, text: text), width: 104, autofit: "<a:normAutofit/>"),
            Read(Paragraph(1, marker: marker, text: text), width: 104, autofit: "<a:spAutoFit/>"),
            Read(Paragraph(1, marker: marker, text: text), width: 104, body: "wrap=\"none\""),
            Read(Paragraph(1, marker: marker, text: text), width: 104, shape: "rot=\"900000\""),
            Read(Paragraph(1, marker: marker, text: text), width: 104, shape: "flipH=\"1\""),
            Read(Paragraph(1, marker: marker, text: text), width: 104, shape: "flipV=\"1\""),
            Read(Paragraph(1, "arabicPlain", marker: marker, text: text), width: 104),
            Read(Paragraph(1, marker: marker + "<a:tabLst><a:tab pos=\"914400\" algn=\"l\"/></a:tabLst>", text: text), width: 104),
            Read(Paragraph(1, marker: marker, text: "AA\u00a0BB"), width: 104),
            Read(Paragraph(1, marker: marker, text: text, suffix: "<a:br/><a:r><a:rPr sz=\"2400\"><a:latin typeface=\"TestFont\"/></a:rPr><a:t>Continuation</a:t></a:r>"), width: 104)
        };
        foreach (var spans in variants)
        {
            var a = spans.Single(s => s.Run.Text.Contains("AA", StringComparison.Ordinal)).Run;
            var bb = spans.Single(s => s.Run.Text.Contains("BB", StringComparison.Ordinal)).Run;
            TestAssert.Equal(a.Y, bb.Y);
        }
        var oversizedFirstWord = Read(Paragraph(1, text: "AA BB"), width: 69);
        var first = oversizedFirstWord.Single(s => s.Run.Text.Contains("AA", StringComparison.Ordinal)).Run;
        var number = oversizedFirstWord.Single(s => s.Run.Text == "1.").Run;
        TestAssert.Equal(number.Y, first.Y);
        TestAssert.Equal(114d, first.X);
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
