using System.Xml.Linq;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxInheritanceTests
{
    public static void PptxInheritanceLayersRetainLiveRunStylesAndFrames()
    {
        var presentation = XDocument.Parse(PptxTests.BasicPresentation());
        presentation.Root!.Add(XElement.Parse($"<p:defaultTextStyle xmlns:p='{P}' xmlns:a='{A}'><a:lvl2pPr><a:defRPr strike='sngStrike'/></a:lvl2pPr></p:defaultTextStyle>"));
        var parts = new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(PptxTests.BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(presentation.ToString()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="layout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>
                """),
            ["ppt/slideLayouts/_rels/slideLayout1.xml.rels"] = TestFixtures.Utf8("""
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="master" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>
                """),
            ["ppt/slideMasters/slideMaster1.xml"] = TestFixtures.Utf8($$"""
                <p:sldMaster xmlns:p="{{P}}" xmlns:a="{{A}}">
                  <p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="1" name="Master body"/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>
                    <p:spPr/><p:txBody><a:bodyPr/><a:lstStyle><a:lvl2pPr/></a:lstStyle><a:p/></p:txBody>
                  </p:sp></p:spTree></p:cSld>
                  <p:txStyles><p:bodyStyle><a:lvl2pPr algn="ctr"><a:defRPr sz="2800" b="1" spc="120"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:defRPr></a:lvl2pPr></p:bodyStyle></p:txStyles>
                </p:sldMaster>
                """),
            ["ppt/slideLayouts/slideLayout1.xml"] = TestFixtures.Utf8($$"""
                <p:sldLayout xmlns:p="{{P}}" xmlns:a="{{A}}"><p:cSld><p:spTree><p:sp>
                  <p:nvSpPr><p:cNvPr id="2" name="Layout body"/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>
                  <p:spPr/><p:txBody><a:bodyPr/><a:lstStyle><a:lvl2pPr><a:defRPr sz="2600" i="1"/></a:lvl2pPr></a:lstStyle><a:p/></p:txBody>
                </p:sp></p:spTree></p:cSld></p:sldLayout>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8($$"""
                <p:sld xmlns:p="{{P}}" xmlns:a="{{A}}"><p:cSld><p:spTree><p:sp>
                  <p:nvSpPr><p:cNvPr id="3" name="Slide body"/><p:cNvSpPr/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>
                  <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm></p:spPr>
                  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr lvl="1"/>
                    <a:r><a:rPr u="sng"><a:highlight><a:srgbClr val="FFFF00"/></a:highlight></a:rPr><a:t>Hello</a:t></a:r>
                    <a:br/><a:fld type="slidenum"><a:rPr sz="1200"/><a:t>1</a:t></a:fld>
                  </a:p></p:txBody>
                </p:sp></p:spTree></p:cSld></p:sld>
                """)
        };
        WithDeck(parts, (document, package) =>
        {
            var scene = new PptxSceneBuilder().Build(document, package, default);
            var fonts = new PresentationFontResolver(new TestFaceFontResolver());
            TestAssert.Equal("Hello\n1", string.Concat(scene.Slides[0].SlideNodes.Single().TextBody!.Paragraphs.Single().Runs.Select(run => run.Text)));

            PptxTextGlyphRunSnapshot[] glyphHelloRuns = PptxRenderer.InspectTextGlyphRuns(document, package, 0, scene, fonts)
                .Where(run => run.FrameIndex == 0 && run.ParagraphIndex == 0 && run.SourceRunIndex == 0)
                .ToArray();
            TestAssert.Equal("Hello", string.Concat(glyphHelloRuns.Select(run => run.Text)));
            TestAssert.True(glyphHelloRuns.All(run => run.Bold && run.Italic && run.Strike),
                "Live glyph styles retain master bold, presentation strike and the layout italic override.");
            TestAssert.True(glyphHelloRuns.Length != 0 && glyphHelloRuns.All(run => run.Underline),
                "The live glyph inspector retains underline on every emitted source segment.");
            TestAssert.True(glyphHelloRuns.All(run => run.HighlightColor == new RgbColor(255, 255, 0)),
                "The live glyph inspector retains the resolved source highlight.");
            TestAssert.True(
                glyphHelloRuns.All(run => run.UnderlineWidth is not null && run.UnderlineHeight is not null),
                "Expected glyph-run inspection to expose underline geometry for each emitted source-run segment.");
            TestAssert.True(
                glyphHelloRuns.All(run => Math.Abs(run.UnderlineWidth.GetValueOrDefault() - run.Width) < 0.01d),
                "Expected underline geometry to be owned by each emitted glyph-run segment width.");

            IReadOnlyList<PptxTextFrameModelSnapshot> textFrames = PptxRenderer.InspectTextFrameModels(document, package, 0, scene, fonts);
            TestAssert.True(
                textFrames.Any(frame => frame.Paragraphs.Any(paragraph => paragraph.Runs.Any(run => run.Text == "Hello"))),
                "Expected text-frame model inspection to include Hello. Runs: " + string.Join("|", textFrames.SelectMany(frame => frame.Paragraphs).SelectMany(paragraph => paragraph.Runs).Select(run => run.Text)));
            PptxTextFrameModelSnapshot textFrame = textFrames.Single(frame => frame.Paragraphs.Any(paragraph => paragraph.Runs.Any(run => run.Text == "Hello")));
            TestAssert.True(textFrame.InheritedPlaceholderCount >= 1, "Expected text model to expose inherited placeholder participation.");
            TestAssert.True(textFrame.HasInheritedTextBody, "Expected text model to expose inherited placeholder text body participation.");
            TestAssert.True(!textFrame.UsesInheritedShapeBounds, "Expected direct text-box geometry to stay distinct from inherited placeholder style participation.");
            TestAssert.Equal(1, textFrame.Paragraphs.Count);
            TestAssert.Equal(1, textFrame.Paragraphs[0].Level);
            TestAssert.Equal("lvl2pPr", textFrame.Paragraphs[0].CascadeLevelName);
            TestAssert.True(textFrame.Paragraphs[0].ResolvedCascadeSourceCount >= 2, "Expected text model to expose inherited cascade inputs before style resolution.");
            TestAssert.Contains("shape.lstStyle", string.Join("|", textFrame.Paragraphs[0].CascadeLayerNames));
            TestAssert.Contains("layout.placeholder.lstStyle", string.Join("|", textFrame.Paragraphs[0].CascadeLayerNames));
            TestAssert.Contains("master.placeholder.lstStyle", string.Join("|", textFrame.Paragraphs[0].CascadeLayerNames));
            TestAssert.Contains("inherited.txStyle", string.Join("|", textFrame.Paragraphs[0].CascadeLayerNames));
            TestAssert.Contains("defaultTextStyle", string.Join("|", textFrame.Paragraphs[0].CascadeLayerNames));
            TestAssert.Contains("ShapeListStyle", string.Join("|", textFrame.Paragraphs[0].CascadeLayerKinds));
            TestAssert.Contains("LayoutPlaceholderListStyle", string.Join("|", textFrame.Paragraphs[0].CascadeLayerKinds));
            TestAssert.Contains("MasterPlaceholderListStyle", string.Join("|", textFrame.Paragraphs[0].CascadeLayerKinds));
            TestAssert.Contains("InheritedTextStyle", string.Join("|", textFrame.Paragraphs[0].CascadeLayerKinds));
            TestAssert.Contains("DefaultTextStyle", string.Join("|", textFrame.Paragraphs[0].CascadeLayerKinds));
            TestAssert.True(textFrame.Paragraphs[0].ResolvedStyleSourceCount >= textFrame.Paragraphs[0].ResolvedCascadeSourceCount + 1, "Expected resolved paragraph style provenance to include direct paragraph properties after inherited defaults.");
            TestAssert.Contains("paragraph.pPr", string.Join("|", textFrame.Paragraphs[0].ResolvedStyleLayerNames));
            TestAssert.Contains("ParagraphProperties", string.Join("|", textFrame.Paragraphs[0].ResolvedStyleLayerKinds));
            TestAssert.Equal("Center", textFrame.Paragraphs[0].Alignment);
            TestAssert.Equal(26d, textFrame.Paragraphs[0].FontSize);
            TestAssert.Equal("Text", textFrame.Paragraphs[0].Runs[0].Kind);
            TestAssert.Equal("Break", textFrame.Paragraphs[0].Runs[1].Kind);
            TestAssert.Equal("Field", textFrame.Paragraphs[0].Runs[2].Kind);
            TestAssert.Equal("Hello", textFrame.Paragraphs[0].Runs[0].Text);
            TestAssert.True(textFrame.Paragraphs[0].Runs[0].ResolvedCascadeSourceCount >= 2, "Expected run model to expose direct and inherited run-style inputs.");
            TestAssert.Contains("run.rPr", string.Join("|", textFrame.Paragraphs[0].Runs[0].CascadeLayerNames));
            TestAssert.Contains("paragraph.defRPr", string.Join("|", textFrame.Paragraphs[0].Runs[0].CascadeLayerNames));
            TestAssert.Contains("layout.placeholder.lstStyle.defRPr", string.Join("|", textFrame.Paragraphs[0].Runs[0].CascadeLayerNames));
            TestAssert.Contains("inherited.txStyle.defRPr", string.Join("|", textFrame.Paragraphs[0].Runs[0].CascadeLayerNames));
            TestAssert.Contains("defaultTextStyle.defRPr", string.Join("|", textFrame.Paragraphs[0].Runs[0].CascadeLayerNames));
            TestAssert.Contains("RunProperties", string.Join("|", textFrame.Paragraphs[0].Runs[0].CascadeLayerKinds));
            TestAssert.Contains("ParagraphDefaultRunProperties", string.Join("|", textFrame.Paragraphs[0].Runs[0].CascadeLayerKinds));
            TestAssert.Contains("LayoutPlaceholderDefaultRunProperties", string.Join("|", textFrame.Paragraphs[0].Runs[0].CascadeLayerKinds));
            TestAssert.Contains("InheritedTextStyleDefaultRunProperties", string.Join("|", textFrame.Paragraphs[0].Runs[0].CascadeLayerKinds));
            TestAssert.Contains("DefaultTextStyleDefaultRunProperties", string.Join("|", textFrame.Paragraphs[0].Runs[0].CascadeLayerKinds));
            TestAssert.Equal(26d, textFrame.Paragraphs[0].Runs[0].FontSize);
            TestAssert.True(textFrame.Paragraphs[0].Runs[0].Underline, "Expected text model to preserve resolved run underline before layout.");
            TestAssert.Equal(new RgbColor(255, 255, 0), textFrame.Paragraphs[0].Runs[0].Highlight ?? default);

            PptxTextFlowSnapshot textFlow = PptxRenderer.InspectTextFlow(document, package, 0, scene, fonts);
            PptxTextFlowFrameSnapshot flowFrame = textFlow.Frames.Single(frame => frame.Paragraphs.Any(paragraph => paragraph.Runs.Any(run => run.SourceText == "Hello")));
            TestAssert.Equal(1, flowFrame.Paragraphs.Count);
            TestAssert.Equal(1, flowFrame.Paragraphs[0].Level);
            TestAssert.Equal("Center", flowFrame.Paragraphs[0].Alignment);
            TestAssert.Equal("Hello", flowFrame.Paragraphs[0].Runs[0].SourceText);
            TestAssert.Equal("Text", flowFrame.Paragraphs[0].Runs[0].Segments[0].Kind);
            TestAssert.Equal("Hello", flowFrame.Paragraphs[0].Runs[0].Segments[0].Text);
            TestAssert.Equal("Break", flowFrame.Paragraphs[0].Runs[1].SourceKind);
            TestAssert.Equal("Break", flowFrame.Paragraphs[0].Runs[1].Segments[0].Kind);
            TestAssert.True(flowFrame.TextWidth > 0d, "Expected text flow box to own text bounds before line layout.");

            PptxTextLayoutSnapshot textLayout = PptxRenderer.InspectTextLayout(document, package, 0, scene, fonts);
            PptxTextFrameLayoutSnapshot layoutFrame = textLayout.Frames.Single(frame => frame.Paragraphs.Any(paragraph => paragraph.Lines.Any(line => line.Spans.Any(span => span.SourceText == "Hello"))));
            TestAssert.Equal(1, layoutFrame.Paragraphs.Count);
            TestAssert.Equal(2, layoutFrame.Paragraphs[0].Lines.Count);
            PptxTextSpanLayoutSnapshot[] layoutHelloSpans = layoutFrame.Paragraphs[0].Lines[0].Spans
                .Where(span => span.SourceText == "Hello")
                .ToArray();
            TestAssert.Equal("Hello", string.Concat(layoutHelloSpans.Select(span => span.Text)));
            TestAssert.True(layoutHelloSpans.All(span => span.SourceText == "Hello"), "Expected split layout spans to retain source-run text ownership.");
            TestAssert.Equal("1", layoutFrame.Paragraphs[0].Lines[1].Spans[0].Text);
            TestAssert.True(layoutFrame.Paragraphs[0].Lines[0].EndX > layoutFrame.Paragraphs[0].Lines[0].StartX, "Expected layout line to own measured advance before PDF emission.");
            TestAssert.True(layoutFrame.Paragraphs[0].Lines[0].TopY > layoutFrame.Paragraphs[0].Lines[0].BaselineY, "Expected layout line boxes to expose top-to-baseline geometry before PDF emission.");
            TestAssert.True(layoutFrame.Paragraphs[0].Lines[0].Advance > 0d, "Expected layout line boxes to expose line advance before paragraph stacking.");
            TestAssert.True(layoutFrame.Paragraphs[0].Lines[0].BaselineOffset > 0d, "Expected layout line boxes to own baseline offset separately from text spans.");
            TestAssert.Equal("Default", layoutFrame.Paragraphs[0].Lines[0].LineSpacingKind);
        });
    }
}
