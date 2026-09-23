using System.Reflection;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxSceneTextAgreementTests
{
    public static void SceneParagraphRetainsCascadeDefaults()
    {
        // R14-deeper: the scene builder resolves the placeholder/master default chain once
        // per paragraph; the migration reuses these retained defaults instead of re-walking
        // the chain in the renderer.
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:sp><p:nvSpPr><p:cNvPr id="2" name="Defaults"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4572000" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                      <p:txBody>
                        <a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr sz="1600"/></a:lvl1pPr></a:lstStyle>
                        <a:p><a:r><a:rPr><a:latin typeface="Arial"/></a:rPr><a:t>Styled</a:t></a:r></a:p>
                      </p:txBody>
                    </p:sp>
                    <p:sp><p:nvSpPr><p:cNvPr id="3" name="Plain"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="1828800"/><a:ext cx="4572000" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                      <p:txBody>
                        <a:bodyPr/><a:lstStyle/>
                        <a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>Bare</a:t></a:r></a:p>
                      </p:txBody>
                    </p:sp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);

        PptxSceneTextParagraph styled = TestAssert.NotNull(scene.Slides[0].SlideNodes[0].TextBody).Paragraphs[0];
        XElement? styledDefaults = styled.DefaultParagraphProperties;
        TestAssert.True(styledDefaults is not null, "Expected the shape lstStyle level defaults to be retained for the migration.");
        TestAssert.Equal("1600", styledDefaults.Elements().FirstOrDefault(e => e.Name.LocalName == "defRPr")?.Attribute("sz")?.Value ?? string.Empty);
        TestAssert.Equal("1600", styled.DefaultRunProperties?.Attribute("sz")?.Value ?? string.Empty);

        PptxSceneTextParagraph bare = TestAssert.NotNull(scene.Slides[0].SlideNodes[1].TextBody).Paragraphs[0];
        TestAssert.True(bare.DefaultParagraphProperties is null, "Expected no retained defaults without any level overrides.");
    }

    public static void SceneCascadeLayersMatchRendererSnapshot()
    {
        // R14-deeper: scene-retained cascade layers must reproduce the renderer snapshot
        // layer names and kinds (the matched placeholder inherits layout lstStyle).
        string input = Path.Combine(Directory.GetCurrentDirectory(), "tests", "Lokad.OoxPdf.Tests", "Cases",
            "pptx-ladder-04-placeholder-inherit-matched.pptx");
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        PptxSceneTextParagraph sceneParagraph = scene.Slides[0].SlideNodes
            .Where(node => node.TextBody is not null)
            .SelectMany(node => node.TextBody!.Paragraphs)
            .First(paragraph => paragraph.Runs.Any(run => run.Text == "BodyText"));
        PptxTextParagraphModelSnapshot snapshotParagraph = PptxRenderer.InspectTextFrameModels(document, package, 0)
            .SelectMany(frame => frame.Paragraphs)
            .First(paragraph => paragraph.Runs.Any(run => run.Text == "BodyText"));
        TestAssert.Equal(snapshotParagraph.CascadeLayerNames.Count, sceneParagraph.CascadeLayers.Count);
        for (int index = 0; index < sceneParagraph.CascadeLayers.Count; index++)
        {
            TestAssert.Equal(snapshotParagraph.CascadeLayerNames[index], sceneParagraph.CascadeLayers[index].Name);
            TestAssert.Equal(snapshotParagraph.CascadeLayerKinds[index], sceneParagraph.CascadeLayers[index].Kind);
        }
        TestAssert.True(sceneParagraph.CascadeLayers.Any(layer => layer.Name == "layout.placeholder.lstStyle"),
            "Expected the matched placeholder to retain the layout placeholder list-style layer.");
    }

    public static void SceneFedSpansMatchXmlPathForPlainShape()
    {
        // R14-deeper: the scene-fed builder must reproduce the XML-path positioned
        // spans exactly for plain shapes (same retained inputs through shared readers).
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="21" name="Probe21"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4572000" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800" b="1"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:rPr><a:t>Alpha </a:t></a:r><a:r><a:rPr sz="1200" i="1"><a:solidFill><a:srgbClr val="445566"/></a:solidFill></a:rPr><a:t>Beta</a:t></a:r></a:p>
                      <a:p><a:r><a:rPr sz="1400" u="sng" strike="sngStrike"><a:solidFill><a:srgbClr val="778899"/></a:solidFill></a:rPr><a:t>Gamma</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        PptxSceneNode node = scene.Slides[0].SlideNodes[0];

        IReadOnlyList<PptxRenderer.PptxPositionedTextSpan> xmlSpans = ReadSpans(node, document, scene);
        var resolver = CreateCannedPresentationResolver("memory:r14-scene-fed");
        IReadOnlyList<PptxRenderer.PptxPositionedTextSpan> fedSpans = PptxRenderer.BuildSceneFedTextSpans(node, document, scene.Theme, scene.Slides[0].SlideColorMap, 1, false, InheritedSources(scene), resolver, CancellationToken.None);
        TestAssert.Equal(xmlSpans.Count, fedSpans.Count);
        for (int index = 0; index < xmlSpans.Count; index++)
        {
            TestAssert.Equal(xmlSpans[index].Run.Text, fedSpans[index].Run.Text);
            TestAssert.Equal(xmlSpans[index].Run.X, fedSpans[index].Run.X);
            TestAssert.Equal(xmlSpans[index].Run.Y, fedSpans[index].Run.Y);
            TestAssert.Equal(xmlSpans[index].Run.FontSize, fedSpans[index].Run.FontSize);
            TestAssert.Equal(xmlSpans[index].Run.Bold, fedSpans[index].Run.Bold);
            TestAssert.Equal(xmlSpans[index].Run.Italic, fedSpans[index].Run.Italic);
            TestAssert.Equal(xmlSpans[index].Run.Underline, fedSpans[index].Run.Underline);
            TestAssert.Equal(xmlSpans[index].ParagraphIndex, fedSpans[index].ParagraphIndex);
            TestAssert.Equal(xmlSpans[index].LineIndex, fedSpans[index].LineIndex);
        }
    }

    public static void SceneFedSpansMatchXmlPathForPlaceholders()
    {
        // R14-deeper: placeholder inheritance resolves through retained defaults on
        // both paths (unmatched body placeholder and layout-matched placeholder).
        string[] inputs = new[]
        {
            System.IO.Path.Combine(Directory.GetCurrentDirectory(), "tests", "Lokad.OoxPdf.Tests", "Cases", "pptx-ladder-04-placeholder-inherit.pptx"),
            System.IO.Path.Combine(Directory.GetCurrentDirectory(), "tests", "Lokad.OoxPdf.Tests", "Cases", "pptx-ladder-04-placeholder-inherit-matched.pptx"),
        };
        int comparedSpans = 0;
        foreach (string input in inputs)
        {
            using FileStream stream = File.OpenRead(input);
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
            PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
            var resolver = CreateCannedPresentationResolver("memory:r14-scene-fed-placeholders");
            foreach (PptxSceneNode node in scene.Slides[0].SlideNodes)
            {
                if (node.TextBody is null)
                {
                    continue;
                }
                IReadOnlyList<PptxRenderer.PptxPositionedTextSpan> xmlSpans = ReadSpans(node, document, scene, true);
                IReadOnlyList<PptxRenderer.PptxPositionedTextSpan> fedSpans = PptxRenderer.BuildSceneFedTextSpans(node, document, scene.Theme, scene.Slides[0].SlideColorMap, 1, true, InheritedSources(scene), resolver, CancellationToken.None);
                TestAssert.Equal(xmlSpans.Count, fedSpans.Count);
                comparedSpans += xmlSpans.Count;
                for (int index = 0; index < xmlSpans.Count; index++)
                {
                    TestAssert.Equal(xmlSpans[index].Run.Text, fedSpans[index].Run.Text);
                    TestAssert.Equal(xmlSpans[index].Run.X, fedSpans[index].Run.X);
                    TestAssert.Equal(xmlSpans[index].Run.Y, fedSpans[index].Run.Y);
                    TestAssert.Equal(xmlSpans[index].Run.FontSize, fedSpans[index].Run.FontSize);
                }
            }
        TestAssert.True(comparedSpans > 0, "Expected placeholder fixtures to exercise fed-path spans.");
        }
    }

    public static void PlainShapeRunTextAndStyleAgree()
    {
        // R14: first scene-fed-layout agreement gate. Plain-shape run text and core
        // styles must resolve identically in the scene model and the renderer spans
        // (the spans path re-clones node.Source today; the scene TextBody is the
        // migration target).
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="21" name="Probe21"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4572000" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800" b="1"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:rPr><a:t>Alpha </a:t></a:r><a:r><a:rPr sz="1200" i="1"><a:solidFill><a:srgbClr val="445566"/></a:solidFill></a:rPr><a:t>Beta</a:t></a:r></a:p>
                      <a:p><a:r><a:rPr sz="1400" u="sng" strike="sngStrike"><a:solidFill><a:srgbClr val="778899"/></a:solidFill></a:rPr><a:t>Gamma</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        PptxSceneNode node = scene.Slides[0].SlideNodes[0];
        PptxSceneTextBody body = TestAssert.NotNull(node.TextBody);
        PptxSceneTextRun[] sceneRuns = body.Paragraphs
            .SelectMany(paragraph => paragraph.Runs)
            .Where(run => run.Kind == PptxSceneTextRunKind.Text)
            .ToArray();

        IReadOnlyList<PptxRenderer.PptxPositionedTextSpan> positioned = ReadSpans(node, document, scene);
        var chunks = new List<(PptxRenderer.PptxTextRunModel Model, StringBuilder Text)>();
        foreach (PptxRenderer.PptxPositionedTextSpan span in positioned)
        {
            TestAssert.True(span.SourceRun is not null, "Expected every span of the plain shape to carry its source run.");
            PptxRenderer.PptxTextRunModel model = span.SourceRun!;
            if (chunks.Count > 0 && ReferenceEquals(chunks[^1].Model, model))
            {
                chunks[^1].Text.Append(span.Run.Text);
            }
            else
            {
                chunks.Add((model, new StringBuilder(span.Run.Text)));
            }
        }

        TestAssert.Equal(sceneRuns.Length, chunks.Count);
        for (int i = 0; i < sceneRuns.Length; i++)
        {
            TestAssert.Equal(sceneRuns[i].Text, chunks[i].Text.ToString());
            TestAssert.Equal(sceneRuns[i].ResolvedStyle.FontSize, chunks[i].Model.Style.FontSize);
            TestAssert.Equal(sceneRuns[i].ResolvedStyle.Bold, chunks[i].Model.Style.Bold);
            TestAssert.Equal(sceneRuns[i].ResolvedStyle.Italic, chunks[i].Model.Style.Italic);
            TestAssert.Equal(sceneRuns[i].ResolvedStyle.Underline, chunks[i].Model.Style.Underline);
            TestAssert.Equal(sceneRuns[i].ResolvedStyle.Strike, chunks[i].Model.Style.Strike);
            TestAssert.Equal(sceneRuns[i].ResolvedStyle.Color, chunks[i].Model.Style.Color);
        }
    }

    public static void PlainShapeFullRunStyleAgrees()
    {
        // R14: full run-style agreement gate. Every scene run-style field must
        // resolve identically in the renderer run models: caps, spacing, baseline,
        // highlight, explicit and fallback typefaces, underline/strike variants,
        // and fill alpha. This widens the migration-safe surface to the whole
        // PptxSceneRunStyle record.
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="22" name="Probe22"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="6400800" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800" cap="allCaps" spc="200"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:rPr><a:t>A</a:t></a:r><a:r><a:rPr sz="1800" baseline="30000" u="dbl"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:rPr><a:t>B</a:t></a:r><a:r><a:rPr sz="1800" strike="dblStrike"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>C</a:t></a:r><a:r><a:rPr sz="1800"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:highlight><a:srgbClr val="FFFF00"/></a:highlight></a:rPr><a:t>D</a:t></a:r><a:r><a:rPr sz="1800"><a:solidFill><a:srgbClr val="112233"><a:alpha val="50000"/></a:srgbClr></a:solidFill></a:rPr><a:t>E</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        PptxSceneNode node = scene.Slides[0].SlideNodes[0];
        PptxSceneTextBody body = TestAssert.NotNull(node.TextBody);
        PptxSceneTextRun[] sceneRuns = body.Paragraphs
            .SelectMany(paragraph => paragraph.Runs)
            .Where(run => run.Kind == PptxSceneTextRunKind.Text)
            .ToArray();
        TestAssert.Equal(5, sceneRuns.Length);

        IReadOnlyList<PptxRenderer.PptxPositionedTextSpan> positioned = ReadSpans(node, document, scene);
        PptxRenderer.PptxTextRunModel[] spanRuns = positioned
            .Select(span => span.SourceRun)
            .Where(run => run is not null)
            .Select(run => run!)
            .Distinct()
            .ToArray();
        TestAssert.Equal(5, spanRuns.Length);
        for (int i = 0; i < sceneRuns.Length; i++)
        {
            TestAssert.Equal(sceneRuns[i].Text, spanRuns[i].Text);
            // Size agrees at the nominal level: raised/lowered runs render scaled
            // (superscript/subscript rule), so emission FontSize differs by design.
            TestAssert.Equal(sceneRuns[i].ResolvedStyle.FontSize, spanRuns[i].Style.NominalFontSize);
            TestAssert.Equal(sceneRuns[i].ResolvedStyle.Color, spanRuns[i].Style.Color);
            TestAssert.Equal(sceneRuns[i].ResolvedStyle.Alpha, spanRuns[i].Style.Alpha);
            TestAssert.Equal(sceneRuns[i].ResolvedStyle.Typeface, spanRuns[i].Style.Typeface);
            TestAssert.Equal(sceneRuns[i].ResolvedStyle.TypefaceSource, spanRuns[i].Style.TypefaceSource);
            TestAssert.Equal(sceneRuns[i].ResolvedStyle.Bold, spanRuns[i].Style.Bold);
            TestAssert.Equal(sceneRuns[i].ResolvedStyle.Italic, spanRuns[i].Style.Italic);
            TestAssert.Equal(sceneRuns[i].ResolvedStyle.Underline, spanRuns[i].Style.Underline);
            TestAssert.Equal(sceneRuns[i].ResolvedStyle.UnderlineValue, spanRuns[i].Style.UnderlineValue);
            TestAssert.Equal(sceneRuns[i].ResolvedStyle.Strike, spanRuns[i].Style.Strike);
            TestAssert.Equal(sceneRuns[i].ResolvedStyle.StrikeValue, spanRuns[i].Style.StrikeValue);
            TestAssert.Equal(sceneRuns[i].ResolvedStyle.CapsValue, spanRuns[i].Style.CapsValue);
            TestAssert.Equal(sceneRuns[i].ResolvedStyle.BaselineOffset, spanRuns[i].Style.BaselineOffset);
            TestAssert.Equal(sceneRuns[i].ResolvedStyle.Highlight, spanRuns[i].Style.Highlight);
        }
    }

    public static void BreakRunsAgreeBetweenSceneAndSpans()
    {
        // R14: break-run agreement probe. The scene model carries an explicit Break
        // run; the spans encode it as a line boundary (breaks emit no span), so the
        // gate pins text runs plus the forced line break at the same position.
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="23" name="Probe23"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4572000" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"/><a:t>One</a:t></a:r><a:br/><a:r><a:rPr sz="1800"/><a:t>Two</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        PptxSceneNode node = scene.Slides[0].SlideNodes[0];
        PptxSceneTextBody body = TestAssert.NotNull(node.TextBody);
        PptxSceneTextRun[] sceneRuns = body.Paragraphs
            .SelectMany(paragraph => paragraph.Runs)
            .ToArray();
        TestAssert.Equal(3, sceneRuns.Length);
        TestAssert.Equal(PptxSceneTextRunKind.Break, sceneRuns[1].Kind);

        IReadOnlyList<PptxRenderer.PptxPositionedTextSpan> positioned = ReadSpans(node, document, scene);
        var lines = new List<(string Text, int Line)>();
        PptxRenderer.PptxTextRunModel? current = null;
        foreach (PptxRenderer.PptxPositionedTextSpan span in positioned)
        {
            TestAssert.True(span.SourceRun is not null, "Expected every span of the probe shape to carry its source run.");
            if (!ReferenceEquals(current, span.SourceRun))
            {
                current = span.SourceRun;
                lines.Add((span.SourceRun!.Text, span.LineIndex));
            }
        }

        TestAssert.Equal(2, lines.Count);
        TestAssert.Equal("One", lines[0].Text);
        TestAssert.Equal("Two", lines[1].Text);
        TestAssert.True(lines[1].Line > lines[0].Line, "Expected the scene break run to surface as a span line boundary at the same position.");
    }

    public static void FieldRunsAgreeBetweenSceneAndSpans()
    {
        // R14: field-run agreement probe. Slide-number fields must resolve to the
        // same text and style in the scene model and the renderer spans.
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="24" name="Probe24"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4572000" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"/><a:t>Page </a:t></a:r><a:fld type="slidenum"><a:rPr sz="1800"/><a:t>1</a:t></a:fld><a:r><a:rPr sz="1800"/><a:t> end</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        PptxSceneNode node = scene.Slides[0].SlideNodes[0];
        PptxSceneTextBody body = TestAssert.NotNull(node.TextBody);
        PptxSceneTextRun[] sceneRuns = body.Paragraphs
            .SelectMany(paragraph => paragraph.Runs)
            .ToArray();
        TestAssert.Equal(3, sceneRuns.Length);
        TestAssert.Equal(PptxSceneTextRunKind.Field, sceneRuns[1].Kind);

        IReadOnlyList<PptxRenderer.PptxPositionedTextSpan> positioned = ReadSpans(node, document, scene);
        var chunks = new List<(PptxRenderer.PptxTextRunModel Model, StringBuilder Text)>();
        foreach (PptxRenderer.PptxPositionedTextSpan span in positioned)
        {
            TestAssert.True(span.SourceRun is not null, "Expected every span of the probe shape to carry its source run.");
            PptxRenderer.PptxTextRunModel model = span.SourceRun!;
            if (chunks.Count > 0 && ReferenceEquals(chunks[^1].Model, model))
            {
                chunks[^1].Text.Append(span.Run.Text);
            }
            else
            {
                chunks.Add((model, new StringBuilder(span.Run.Text)));
            }
        }

        TestAssert.Equal(3, chunks.Count);
        TestAssert.Equal(PptxRenderer.PptxTextRunKind.Field, chunks[1].Model.Kind);
        for (int i = 0; i < sceneRuns.Length; i++)
        {
            TestAssert.Equal(sceneRuns[i].Text, chunks[i].Text.ToString());
            TestAssert.Equal(sceneRuns[i].ResolvedStyle.FontSize, chunks[i].Model.Style.FontSize);
            TestAssert.Equal(sceneRuns[i].ResolvedStyle.Color, chunks[i].Model.Style.Color);
        }
    }

    public static void ParagraphAlignmentAgreesBetweenSceneAndSpans()
    {
        // R14: paragraph-level agreement probe. Explicit paragraph alignment must
        // resolve identically in the scene model and the renderer line layout
        // (alignment applies at line level; fragments stay fragment-relative).
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="25" name="Probe25"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4572000" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="1800"/><a:t>Centered</a:t></a:r></a:p>
                      <a:p><a:pPr algn="r"/><a:r><a:rPr sz="1800"/><a:t>Right</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        PptxSceneNode node = scene.Slides[0].SlideNodes[0];
        PptxSceneTextBody body = TestAssert.NotNull(node.TextBody);
        string[] sceneAlignments = body.Paragraphs
            .Select(paragraph => paragraph.ResolvedStyle.Alignment)
            .ToArray();
        TestAssert.Equal(2, sceneAlignments.Length);

        PptxTextParagraphLayoutSnapshot[] layoutParagraphs = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames.SelectMany(frame => frame.Paragraphs)
            .ToArray();
        TestAssert.Equal(sceneAlignments.Length, layoutParagraphs.Length);
        // The scene preserves the raw alignment token while layout lines carry the
        // resolved name; the gate bridges them through the production parser so both
        // pipelines must interpret the token identically.
        MethodInfo parse = typeof(PptxRenderer).GetMethod("ParseAlignment", BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException("Expected alignment parser.");
        for (int i = 0; i < sceneAlignments.Length; i++)
        {
            string lineAlignment = layoutParagraphs[i].Lines[0].Alignment;
            object? parsed;
            try
            {
                parsed = parse.Invoke(null, [sceneAlignments[i]]);
            }
            catch (TargetInvocationException ex)
            {
                throw ex.InnerException ?? ex;
            }

            TestAssert.Equal(parsed?.ToString(), lineAlignment);
        }
    }

    public static void PlaceholderTextAgreesBetweenSceneAndSpans()
    {
        // R14: placeholder-text agreement probe. An unmatched body placeholder keeps
        // its own txBody text and inherits the master bodyStyle size on both pipelines.
        // (Fixtures must use p:ph: real Office files never contain a:ph, and the
        // placeholder readers only recognize the presentationml namespace.)
        string slide = """
            <?xml version="1.0" encoding="UTF-8"?>
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld><p:spTree><p:sp>
                <p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
                <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4572000" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                <p:txBody>
                  <a:bodyPr tIns="0" bIns="0"/><a:lstStyle/>
                  <a:p><a:r><a:t>BodyText</a:t></a:r></a:p>
                </p:txBody>
              </p:sp></p:spTree></p:cSld>
            </p:sld>
            """;
        string slideRels = """
            <?xml version="1.0" encoding="UTF-8"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
            </Relationships>
            """;
        string layout = """
            <?xml version="1.0" encoding="UTF-8"?>
            <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree/></p:cSld></p:sldLayout>
            """;
        string layoutRels = """
            <?xml version="1.0" encoding="UTF-8"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
            </Relationships>
            """;
        string master = """
            <?xml version="1.0" encoding="UTF-8"?>
            <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld><p:spTree/></p:cSld>
              <p:txStyles>
                <p:titleStyle><a:lvl1pPr><a:defRPr sz="3100"/></a:lvl1pPr></p:titleStyle>
                <p:bodyStyle><a:lvl1pPr><a:defRPr sz="2500"><a:solidFill><a:srgbClr val="AABBCC"/></a:solidFill></a:defRPr></a:lvl1pPr></p:bodyStyle>
                <p:otherStyle><a:lvl1pPr><a:defRPr sz="1900"/></a:lvl1pPr></p:otherStyle>
              </p:txStyles>
            </p:sldMaster>
            """;
        string types = """
            <?xml version="1.0" encoding="UTF-8"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
              <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
              <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
              <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
            </Types>
            """;
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = types,
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = slide,
            ["ppt/slides/_rels/slide1.xml.rels"] = slideRels,
            ["ppt/slideLayouts/slideLayout1.xml"] = layout,
            ["ppt/slideLayouts/_rels/slideLayout1.xml.rels"] = layoutRels,
            ["ppt/slideMasters/slideMaster1.xml"] = master,
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        PptxSceneNode node = scene.Slides[0].SlideNodes[0];
        PptxSceneTextBody body = TestAssert.NotNull(node.TextBody);
        PptxSceneTextRun[] sceneRuns = body.Paragraphs
            .SelectMany(paragraph => paragraph.Runs)
            .Where(run => run.Kind == PptxSceneTextRunKind.Text)
            .ToArray();
        TestAssert.Equal(1, sceneRuns.Length);
        string sceneText = string.Concat(sceneRuns.Select(run => run.Text));
        TestAssert.Equal(0, body.Paragraphs[0].Level);
        System.Xml.Linq.XElement slideSp = scene.Slides[0].SlideXml.Descendants("{http://schemas.openxmlformats.org/presentationml/2006/main}sp").First();
        TestAssert.True(object.ReferenceEquals(node.Source, slideSp), "node-source-detached");

        IReadOnlyList<PptxRenderer.PptxPositionedTextSpan> positioned = ReadSpans(node, document, scene, includePlaceholders: true);
        PptxRenderer.PptxTextRunModel[] spanRuns = positioned
            .Select(span => span.SourceRun)
            .Where(run => run is not null && run.Kind == PptxRenderer.PptxTextRunKind.Text)
            .Distinct()
            .Select(run => run!)
            .ToArray();
        TestAssert.Equal(1, spanRuns.Length);
        string spanText = string.Concat(spanRuns.Select(run => run.Text));
        TestAssert.Equal("BodyText", sceneText);
        TestAssert.Equal(sceneText, spanText);
        TestAssert.Equal(25d, sceneRuns[0].ResolvedStyle.FontSize);
        TestAssert.Equal(25d, spanRuns[0].Style.FontSize);
        TestAssert.Equal(sceneRuns[0].ResolvedStyle.Color, spanRuns[0].Style.Color);
        TestAssert.Equal(new RgbColor(0xAA, 0xBB, 0xCC), sceneRuns[0].ResolvedStyle.Color);
    }

    public static void MatchedPlaceholderInheritanceAgrees()
    {
        // R14: matched-placeholder inheritance probe. With a layout body placeholder
        // present, the slide placeholder run inherits the layout lstStyle size through
        // the placeholder chain on both pipelines (layout beats master txStyles).
        string slide = """
            <?xml version="1.0" encoding="UTF-8"?>
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld><p:spTree><p:sp>
                <p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
                <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4572000" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                <p:txBody>
                  <a:bodyPr tIns="0" bIns="0"/><a:lstStyle/>
                  <a:p><a:r><a:t>MatchedBody</a:t></a:r></a:p>
                </p:txBody>
              </p:sp></p:spTree></p:cSld>
            </p:sld>
            """;
        string slideRels = """
            <?xml version="1.0" encoding="UTF-8"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
            </Relationships>
            """;
        string layout = """
            <?xml version="1.0" encoding="UTF-8"?>
            <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld><p:spTree><p:sp>
                <p:nvSpPr><p:cNvPr id="2" name="Body Placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
                <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4572000" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                <p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr sz="2600"/></a:lvl1pPr></a:lstStyle><a:p/></p:txBody>
              </p:sp></p:spTree></p:cSld>
            </p:sldLayout>
            """;
        string layoutRels = """
            <?xml version="1.0" encoding="UTF-8"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
            </Relationships>
            """;
        string master = """
            <?xml version="1.0" encoding="UTF-8"?>
            <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld><p:spTree/></p:cSld>
              <p:txStyles>
                <p:bodyStyle><a:lvl1pPr><a:defRPr sz="2500"/></a:lvl1pPr></p:bodyStyle>
                <p:otherStyle><a:lvl1pPr><a:defRPr sz="1900"/></a:lvl1pPr></p:otherStyle>
              </p:txStyles>
            </p:sldMaster>
            """;
        string types = """
            <?xml version="1.0" encoding="UTF-8"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
              <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
              <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
              <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
            </Types>
            """;
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = types,
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = slide,
            ["ppt/slides/_rels/slide1.xml.rels"] = slideRels,
            ["ppt/slideLayouts/slideLayout1.xml"] = layout,
            ["ppt/slideLayouts/_rels/slideLayout1.xml.rels"] = layoutRels,
            ["ppt/slideMasters/slideMaster1.xml"] = master,
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        PptxSceneNode node = scene.Slides[0].SlideNodes[0];
        PptxSceneTextBody body = TestAssert.NotNull(node.TextBody);
        PptxSceneTextRun[] sceneRuns = body.Paragraphs
            .SelectMany(paragraph => paragraph.Runs)
            .Where(run => run.Kind == PptxSceneTextRunKind.Text)
            .ToArray();
        TestAssert.Equal(1, sceneRuns.Length);
        TestAssert.Equal(0, body.Paragraphs[0].Level);
        System.Xml.Linq.XElement slideSp = scene.Slides[0].SlideXml.Descendants("{http://schemas.openxmlformats.org/presentationml/2006/main}sp").First();
        TestAssert.True(object.ReferenceEquals(node.Source, slideSp), "node-source-detached");

        IReadOnlyList<PptxRenderer.PptxPositionedTextSpan> positioned = ReadSpans(node, document, scene, includePlaceholders: true);

        PptxRenderer.PptxTextRunModel[] spanRuns = positioned
            .Select(span => span.SourceRun)
            .Where(run => run is not null)
            .Select(run => run!)
            .Distinct()
            .ToArray();
        TestAssert.Equal(1, spanRuns.Length);
        TestAssert.Equal("MatchedBody", spanRuns[0].Text);
        TestAssert.Equal(sceneRuns[0].Text, spanRuns[0].Text);
        double sceneSize = sceneRuns[0].ResolvedStyle.FontSize;
        double spanSize = spanRuns[0].Style.FontSize;
        TestAssert.Equal(26d, sceneSize);
        TestAssert.Equal(26d, spanSize);
    }

    public static void HyperlinkRunAgreesBetweenSceneAndSpans()
    {
        // R14: hyperlink agreement probe. A linked run resolves the hyperlink
        // color and underline on both pipelines, and the renderer run model carries
        // the click id. The scene model records no click identity (only its color and
        // underline effects), so a scene-fed migration must add click id/action to
        // PptxSceneRunStyle to preserve link emission.
        string slideRels = """
            <?xml version="1.0" encoding="UTF-8"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/" TargetMode="External"/>
            </Relationships>
            """;
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="26" name="Probe26"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4572000" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"><a:hlinkClick r:id="rIdLink"/></a:rPr><a:t>Link</a:t></a:r><a:r><a:rPr sz="1800"/><a:t>Plain</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """,
            ["ppt/slides/_rels/slide1.xml.rels"] = slideRels,
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        PptxSceneNode node = scene.Slides[0].SlideNodes[0];
        PptxSceneTextBody body = TestAssert.NotNull(node.TextBody);
        PptxSceneTextRun[] sceneRuns = body.Paragraphs
            .SelectMany(paragraph => paragraph.Runs)
            .Where(run => run.Kind == PptxSceneTextRunKind.Text)
            .ToArray();
        TestAssert.Equal(2, sceneRuns.Length);

        IReadOnlyList<PptxRenderer.PptxPositionedTextSpan> positioned = ReadSpans(node, document, scene);
        PptxRenderer.PptxTextRunModel[] spanRuns = positioned
            .Select(span => span.SourceRun)
            .Where(run => run is not null)
            .Select(run => run!)
            .Distinct()
            .ToArray();
        TestAssert.Equal(2, spanRuns.Length);
        TestAssert.Equal("Link", spanRuns[0].Text);
        TestAssert.True(spanRuns[0].Style.HasHyperlinkClick, "Expected the linked run to carry a hyperlink click.");
        TestAssert.Equal("rIdLink", spanRuns[0].Style.HyperlinkClickId);
        TestAssert.True(!spanRuns[1].Style.HasHyperlinkClick, "Expected the plain run to carry no hyperlink click.");
        TestAssert.Equal(sceneRuns[0].ResolvedStyle.Color, spanRuns[0].Style.Color);
        TestAssert.Equal(sceneRuns[0].ResolvedStyle.Underline, spanRuns[0].Style.Underline);
    }

    public static void UnstyledRunsInheritParagraphDefaults()
    {
        // R14: paragraph-default agreement probe. Runs without direct formatting
        // inherit the paragraph defRPr size and color identically on both pipelines.
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="27" name="Probe27"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4572000" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:pPr><a:defRPr sz="2000"><a:solidFill><a:srgbClr val="993366"/></a:solidFill></a:defRPr></a:pPr><a:r><a:t>Bare</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        PptxSceneNode node = scene.Slides[0].SlideNodes[0];
        PptxSceneTextBody body = TestAssert.NotNull(node.TextBody);
        PptxSceneTextRun[] sceneRuns = body.Paragraphs
            .SelectMany(paragraph => paragraph.Runs)
            .Where(run => run.Kind == PptxSceneTextRunKind.Text)
            .ToArray();
        TestAssert.Equal(1, sceneRuns.Length);
        TestAssert.Equal(20d, sceneRuns[0].ResolvedStyle.FontSize);

        IReadOnlyList<PptxRenderer.PptxPositionedTextSpan> positioned = ReadSpans(node, document, scene);
        PptxRenderer.PptxTextRunModel[] spanRuns = positioned
            .Select(span => span.SourceRun)
            .Where(run => run is not null)
            .Select(run => run!)
            .Distinct()
            .ToArray();
        TestAssert.Equal(1, spanRuns.Length);
        TestAssert.Equal("Bare", spanRuns[0].Text);
        TestAssert.Equal(20d, spanRuns[0].Style.NominalFontSize);
        TestAssert.Equal(sceneRuns[0].ResolvedStyle.Color, spanRuns[0].Style.Color);
    }

    public static void UnstyledRunsInheritShapeListStyle()
    {
        // R14: list-style agreement probe. Paragraphs without properties inherit
        // the shape lstStyle level defaults identically on both pipelines.
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="28" name="Probe28"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4572000" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr tIns="0" bIns="0"/><a:lstStyle><a:lvl1pPr><a:defRPr sz="2200"><a:solidFill><a:srgbClr val="334455"/></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle>
                      <a:p><a:r><a:t>Styled</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        PptxSceneNode node = scene.Slides[0].SlideNodes[0];
        PptxSceneTextBody body = TestAssert.NotNull(node.TextBody);
        PptxSceneTextRun[] sceneRuns = body.Paragraphs
            .SelectMany(paragraph => paragraph.Runs)
            .Where(run => run.Kind == PptxSceneTextRunKind.Text)
            .ToArray();
        TestAssert.Equal(1, sceneRuns.Length);
        TestAssert.Equal(22d, sceneRuns[0].ResolvedStyle.FontSize);

        IReadOnlyList<PptxRenderer.PptxPositionedTextSpan> positioned = ReadSpans(node, document, scene);
        PptxRenderer.PptxTextRunModel[] spanRuns = positioned
            .Select(span => span.SourceRun)
            .Where(run => run is not null)
            .Select(run => run!)
            .Distinct()
            .ToArray();
        TestAssert.Equal(1, spanRuns.Length);
        TestAssert.Equal("Styled", spanRuns[0].Text);
        TestAssert.Equal(22d, spanRuns[0].Style.NominalFontSize);
        TestAssert.Equal(sceneRuns[0].ResolvedStyle.Color, spanRuns[0].Style.Color);
    }

    public static void UnstyledRunsInheritMasterDefaultStyle()
    {
        // R14: default-style agreement probe. With no shape, inherited, txStyles,
        // or paragraph defaults anywhere, runs inherit the master defaultTextStyle
        // identically on both pipelines.
        string slideRels = """
            <?xml version="1.0" encoding="UTF-8"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
            </Relationships>
            """;
        string layout = """
            <?xml version="1.0" encoding="UTF-8"?>
            <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree/></p:cSld></p:sldLayout>
            """;
        string layoutRels = """
            <?xml version="1.0" encoding="UTF-8"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
            </Relationships>
            """;
        string master = """
            <?xml version="1.0" encoding="UTF-8"?>
            <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld><p:spTree/></p:cSld>
              <p:defaultTextStyle><a:lvl1pPr><a:defRPr sz="2300"><a:solidFill><a:srgbClr val="556677"/></a:solidFill></a:defRPr></a:lvl1pPr></p:defaultTextStyle>
            </p:sldMaster>
            """;
        string types = """
            <?xml version="1.0" encoding="UTF-8"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
              <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
              <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
              <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
            </Types>
            """;
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = types,
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="29" name="Probe29"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4572000" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:r><a:t>Defaulted</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """,
            ["ppt/slides/_rels/slide1.xml.rels"] = slideRels,
            ["ppt/slideLayouts/slideLayout1.xml"] = layout,
            ["ppt/slideLayouts/_rels/slideLayout1.xml.rels"] = layoutRels,
            ["ppt/slideMasters/slideMaster1.xml"] = master,
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        PptxSceneNode node = scene.Slides[0].SlideNodes[0];
        PptxSceneTextBody body = TestAssert.NotNull(node.TextBody);
        PptxSceneTextRun[] sceneRuns = body.Paragraphs
            .SelectMany(paragraph => paragraph.Runs)
            .Where(run => run.Kind == PptxSceneTextRunKind.Text)
            .ToArray();
        TestAssert.Equal(1, sceneRuns.Length);
        TestAssert.Equal(23d, sceneRuns[0].ResolvedStyle.FontSize);

        IReadOnlyList<PptxRenderer.PptxPositionedTextSpan> positioned = ReadSpans(node, document, scene);
        PptxRenderer.PptxTextRunModel[] spanRuns = positioned
            .Select(span => span.SourceRun)
            .Where(run => run is not null)
            .Select(run => run!)
            .Distinct()
            .ToArray();
        TestAssert.Equal(1, spanRuns.Length);
        TestAssert.Equal("Defaulted", spanRuns[0].Text);
        TestAssert.Equal(23d, spanRuns[0].Style.NominalFontSize);
        TestAssert.Equal(sceneRuns[0].ResolvedStyle.Color, spanRuns[0].Style.Color);
    }

    public static void TableCellTextMatchesShapeText()
    {
        // R14: table/shape consistency probe. Identical paragraph XML in a shape
        // and in table cells must resolve to the same run text and core styles on
        // both pipelines, pinning that the table path adds no phantom styling over
        // the shared readers (table cells carry unresolved XML, so this gates the
        // renderer cascade against the scene-resolved shape).
        string bodyPara1 = """<a:p><a:r><a:rPr sz="1600" b="1"><a:solidFill><a:srgbClr val="334455"/></a:solidFill></a:rPr><a:t>One</a:t></a:r></a:p>""";
        string bodyPara2 = """<a:p><a:r><a:rPr sz="1600" i="1"><a:solidFill><a:srgbClr val="334455"/></a:solidFill></a:rPr><a:t>Two</a:t></a:r></a:p>""";
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = $"""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="30" name="Probe30"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr tIns="0" bIns="0"/><a:lstStyle/>
                      {bodyPara1}
                      {bodyPara2}
                    </p:txBody>
                  </p:sp>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="0" y="2743200"/><a:ext cx="3657600" cy="1828800"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                        <a:tbl>
                          <a:tblGrid><a:gridCol w="1828800"/><a:gridCol w="1828800"/></a:tblGrid>
                          <a:tr h="914400">
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/>
                            {bodyPara1}
                            </a:txBody><a:tcPr/></a:tc>
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/>
                            {bodyPara2}
                            </a:txBody><a:tcPr/></a:tc>
                          </a:tr>
                        </a:tbl>
                      </a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        PptxSceneNode shapeNode = scene.Slides[0].SlideNodes[0];
        PptxSceneNode tableNode = scene.Slides[0].SlideNodes[1];
        TestAssert.Equal(PptxSceneNodeKind.Shape, shapeNode.Kind);
        TestAssert.Equal(PptxSceneNodeKind.Table, tableNode.Kind);
        PptxSceneTextBody shapeBody = TestAssert.NotNull(shapeNode.TextBody);
        PptxSceneTextRun[] shapeRuns = shapeBody.Paragraphs
            .SelectMany(paragraph => paragraph.Runs)
            .Where(run => run.Kind == PptxSceneTextRunKind.Text)
            .ToArray();
        TestAssert.Equal(2, shapeRuns.Length);

        var tableResolver = CreateCannedPresentationResolver("memory:r14-table");
        var slideSource = new PptxRenderSource(
            PptxRenderSourceKind.Slide,
            scene.Slides[0].PartName,
            scene.Slides[0].SlideXml,
            new Dictionary<string, OoxRelationship>(),
            scene.Slides[0].SlideColorMap);
        var context = new PptxRenderContext(
            document,
            scene.Theme,
            document.Slides[0],
            scene.Slides[0],
            slideSource,
            [],
            tableResolver,
            new Dictionary<string, PdfImageXObject?>(),
            null,
            CancellationToken.None);
        MethodInfo readTableSpans = typeof(PptxRenderer).GetMethod("ReadTableFrameTextSpans", BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException("Expected table span reader.");
        object? tableSpans;
        try { tableSpans = readTableSpans.Invoke(null, [context, tableNode, PptxRenderer.GroupTransform.Identity, scene.Slides[0].SlideColorMap]); }
        catch (TargetInvocationException ex) { throw ex.InnerException ?? ex; }
        PptxRenderer.PptxTextRunModel[] cellRuns = ((System.Collections.IList)tableSpans!)
            .Cast<PptxRenderer.PptxPositionedTextSpan>()
            .Select(span => span.SourceRun)
            .Where(run => run is not null)
            .Select(run => run!)
            .Distinct()
            .ToArray();
        TestAssert.Equal(2, cellRuns.Length);
        for (int i = 0; i < shapeRuns.Length; i++)
        {
            TestAssert.Equal(shapeRuns[i].Text, cellRuns[i].Text);
            TestAssert.Equal(shapeRuns[i].ResolvedStyle.FontSize, cellRuns[i].Style.NominalFontSize);
            TestAssert.Equal(shapeRuns[i].ResolvedStyle.Color, cellRuns[i].Style.Color);
            TestAssert.Equal(shapeRuns[i].ResolvedStyle.Bold, cellRuns[i].Style.Bold);
            TestAssert.Equal(shapeRuns[i].ResolvedStyle.Italic, cellRuns[i].Style.Italic);
        }
    }

    public static void GroupedShapeTextAgreesBetweenSceneAndSpans()
    {
        // R14: grouped-shape agreement probe. Text in a grouped shape resolves
        // run text and core styles identically on both pipelines through the group
        // transform chain.
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:grpSp>
                    <p:nvGrpSpPr><p:cNvPr id="10" name="Group"/><p:nvPr/></p:nvGrpSpPr>
                    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="1828800"/><a:chOff x="0" y="0"/><a:chExt cx="2743200" cy="1828800"/></a:xfrm></p:grpSpPr>
                    <p:sp><p:nvSpPr><p:cNvPr id="11" name="GroupedShape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
                      <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                      <p:txBody>
                        <a:bodyPr tIns="0" bIns="0"/><a:lstStyle/>
                        <a:p><a:r><a:rPr sz="1800" b="1"><a:solidFill><a:srgbClr val="334455"/></a:solidFill></a:rPr><a:t>Grouped</a:t></a:r></a:p>
                      </p:txBody>
                    </p:sp>
                  </p:grpSp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        PptxSceneNode group = scene.Slides[0].SlideNodes[0];
        TestAssert.Equal(PptxSceneNodeKind.Group, group.Kind);
        PptxSceneNode node = group.Children[0];
        PptxSceneTextBody body = TestAssert.NotNull(node.TextBody);
        PptxSceneTextRun[] sceneRuns = body.Paragraphs
            .SelectMany(paragraph => paragraph.Runs)
            .Where(run => run.Kind == PptxSceneTextRunKind.Text)
            .ToArray();
        TestAssert.Equal(1, sceneRuns.Length);

        IReadOnlyList<PptxRenderer.PptxPositionedTextSpan> positioned = ReadSpans(node, document, scene);
        PptxRenderer.PptxTextRunModel[] spanRuns = positioned
            .Select(span => span.SourceRun)
            .Where(run => run is not null)
            .Select(run => run!)
            .Distinct()
            .ToArray();
        TestAssert.Equal(1, spanRuns.Length);
        TestAssert.Equal("Grouped", spanRuns[0].Text);
        TestAssert.Equal(sceneRuns[0].ResolvedStyle.FontSize, spanRuns[0].Style.NominalFontSize);
        TestAssert.Equal(sceneRuns[0].ResolvedStyle.Color, spanRuns[0].Style.Color);
        TestAssert.Equal(sceneRuns[0].ResolvedStyle.Bold, spanRuns[0].Style.Bold);
    }

    public static void TableStyleTextAgreesBetweenSceneAndSpans()
    {
        // R14: table-style agreement probe. Unstyled header-row runs under a table
        // style resolve the scene cell StyleText and the renderer run styles identically.
        string theme = """
            <?xml version="1.0" encoding="UTF-8"?>
            <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="TableTheme">
              <a:themeElements>
                <a:clrScheme name="TableTheme">
                  <a:dk1><a:srgbClr val="000000"/></a:dk1>
                  <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
                  <a:accent6><a:srgbClr val="336699"/></a:accent6>
                </a:clrScheme>
                <a:fontScheme name="TableTheme"><a:majorFont/><a:minorFont/></a:fontScheme>
              </a:themeElements>
            </a:theme>
            """;
        string presentationRels = """
            <?xml version="1.0" encoding="UTF-8"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
              <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
            </Relationships>
            """;
        string types = """
            <?xml version="1.0" encoding="UTF-8"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
              <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
              <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
            </Types>
            """;
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = types,
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = presentationRels,
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/theme/theme1.xml"] = theme,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="1828800"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                        <a:tbl>
                          <a:tblPr firstRow="1"><a:tableStyleId>{93296810-A885-4BE3-A3E7-6D5BEEA58F35}</a:tableStyleId></a:tblPr>
                          <a:tblGrid><a:gridCol w="1828800"/><a:gridCol w="1828800"/></a:tblGrid>
                          <a:tr h="914400">
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Head</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>Body</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                          </a:tr>
                        </a:tbl>
                      </a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        PptxSceneNode tableNode = scene.Slides[0].SlideNodes[0];
        TestAssert.Equal(PptxSceneNodeKind.Table, tableNode.Kind);
        PptxSceneTable table = TestAssert.NotNull(tableNode.Table);
        PptxSceneTableCell headerCell = table.Rows[0].Cells[0];
        PptxSceneTableCell bodyCell = table.Rows[0].Cells[1];
        var slideSource = new PptxRenderSource(
            PptxRenderSourceKind.Slide,
            scene.Slides[0].PartName,
            scene.Slides[0].SlideXml,
            new Dictionary<string, OoxRelationship>(),
            scene.Slides[0].SlideColorMap);
        var context = new PptxRenderContext(
            document,
            scene.Theme,
            document.Slides[0],
            scene.Slides[0],
            slideSource,
            [],
            CreateCannedPresentationResolver("memory:r14-tablestyle"),
            new Dictionary<string, PdfImageXObject?>(),
            null,
            CancellationToken.None);
        MethodInfo readTableSpans = typeof(PptxRenderer).GetMethod("ReadTableFrameTextSpans", BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException("Expected table span reader.");
        object? tableSpans;
        try { tableSpans = readTableSpans.Invoke(null, [context, tableNode, PptxRenderer.GroupTransform.Identity, scene.Slides[0].SlideColorMap]); }
        catch (TargetInvocationException ex) { throw ex.InnerException ?? ex; }
        PptxRenderer.PptxTextRunModel[] cellRuns = ((System.Collections.IList)tableSpans!)
            .Cast<PptxRenderer.PptxPositionedTextSpan>()
            .Select(span => span.SourceRun)
            .Where(run => run is not null)
            .Select(run => run!)
            .Distinct()
            .ToArray();
        TestAssert.Equal(2, cellRuns.Length);
        TestAssert.Equal("Head", cellRuns[0].Text);
        TestAssert.Equal("Body", cellRuns[1].Text);
        TestAssert.Equal(headerCell.StyleText.Color, cellRuns[0].Style.Color);
        TestAssert.Equal(headerCell.StyleText.Bold, cellRuns[0].Style.Bold);
        TestAssert.Equal(bodyCell.StyleText.Color, cellRuns[1].Style.Color);
        TestAssert.True(headerCell.StyleText.Color.HasValue || headerCell.StyleText.Bold, "Expected the table style to resolve header text styling.");
        TestAssert.Equal(bodyCell.StyleText.Bold, cellRuns[1].Style.Bold);
   }


    public static void SchemeColorRunsAgreeBetweenSceneAndSpans()
    {
        // R14: scheme-color agreement probe. Runs colored by theme scheme
        // references resolve identically on both pipelines through the slide color map.
        string theme = """
            <?xml version="1.0" encoding="UTF-8"?>
            <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="SchemeTheme">
              <a:themeElements>
                <a:clrScheme name="SchemeTheme">
                  <a:dk1><a:srgbClr val="111111"/></a:dk1>
                  <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
                  <a:accent6><a:srgbClr val="336699"/></a:accent6>
                </a:clrScheme>
                <a:fontScheme name="SchemeTheme"><a:majorFont/><a:minorFont/></a:fontScheme>
              </a:themeElements>
            </a:theme>
            """;
        string presentationRels = """
            <?xml version="1.0" encoding="UTF-8"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
              <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
            </Relationships>
            """;
        string types = """
            <?xml version="1.0" encoding="UTF-8"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
              <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
              <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
            </Types>
            """;
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = types,
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = presentationRels,
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/theme/theme1.xml"] = theme,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="31" name="Probe31"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4572000" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"><a:solidFill><a:schemeClr val="accent6"/></a:solidFill></a:rPr><a:t>Acc</a:t></a:r><a:r><a:rPr sz="1800"><a:solidFill><a:schemeClr val="dk1"/></a:solidFill></a:rPr><a:t>Dark</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        PptxSceneNode node = scene.Slides[0].SlideNodes[0];
        PptxSceneTextBody body = TestAssert.NotNull(node.TextBody);
        PptxSceneTextRun[] sceneRuns = body.Paragraphs
            .SelectMany(paragraph => paragraph.Runs)
            .Where(run => run.Kind == PptxSceneTextRunKind.Text)
            .ToArray();
        TestAssert.Equal(2, sceneRuns.Length);
        TestAssert.Equal(new RgbColor(0x33, 0x66, 0x99), sceneRuns[0].ResolvedStyle.Color);
        TestAssert.Equal(new RgbColor(0x11, 0x11, 0x11), sceneRuns[1].ResolvedStyle.Color);

        IReadOnlyList<PptxRenderer.PptxPositionedTextSpan> positioned = ReadSpans(node, document, scene);
        PptxRenderer.PptxTextRunModel[] spanRuns = positioned
            .Select(span => span.SourceRun)
            .Where(run => run is not null)
            .Select(run => run!)
            .Distinct()
            .ToArray();
        TestAssert.Equal(2, spanRuns.Length);
        TestAssert.Equal("Acc", spanRuns[0].Text);
        TestAssert.Equal("Dark", spanRuns[1].Text);
        TestAssert.Equal(sceneRuns[0].ResolvedStyle.Color, spanRuns[0].Style.Color);
        TestAssert.Equal(sceneRuns[1].ResolvedStyle.Color, spanRuns[1].Style.Color);
    }


    private static IReadOnlyList<PptxRenderer.PptxPositionedTextSpan> ReadSpans(
        PptxSceneNode node,
        PptxDocument document,
        PptxScene scene,
        bool includePlaceholders = false)
    {
        var resolver = CreateCannedPresentationResolver("memory:r14-agreement");
        MethodInfo spans = typeof(PptxRenderer).GetMethods(BindingFlags.NonPublic | BindingFlags.Static)
            .Single(candidate => candidate.Name == "ReadTextSpansForShape" && candidate.GetParameters().Length == 9);
        try
        {
            return (IReadOnlyList<PptxRenderer.PptxPositionedTextSpan>)spans.Invoke(null, [node.Source, document, scene.Theme, scene.Slides[0].SlideColorMap, 1, includePlaceholders, InheritedSources(scene), resolver, CancellationToken.None])!;
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }
    }

    private static IReadOnlyList<XDocument> InheritedSources(PptxScene scene)
    {
        // Mirrors the slide placeholder-source chain (master, then layout).
        var inherited = new List<XDocument>();
        if (scene.Slides[0].MasterXml is { } master)
        {
            inherited.Add(master);
        }

        if (scene.Slides[0].LayoutXml is { } layout)
        {
            inherited.Add(layout);
        }

        return inherited;
    }

    private static PresentationFontResolver CreateCannedPresentationResolver(string sourceId)
    {
        byte[] bytes = TestFontBuilder.CreateTestFont();
        OpenTypeFont font = OpenTypeFont.Load(bytes);
        var resolution = new FontFaceResolution(
            font.FamilyName,
            font.FamilyName,
            new FontStyleKey(),
            new MemoryFontProgramSource(sourceId, bytes),
            IsFallback: false);
        return new PresentationFontResolver(new CannedFontResolver(resolution));
    }

    private sealed class CannedFontResolver(FontFaceResolution resolution) : IFontResolver
    {
        public FontFaceResolution Resolve(FontRequest request)
        {
            return resolution;
        }
    }
}
