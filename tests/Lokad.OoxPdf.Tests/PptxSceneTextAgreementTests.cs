using System.Reflection;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Tests;

internal static class PptxSceneTextAgreementTests
{
    public static void SceneParagraphRetainsCascadeDefaults()
    {
        // Source cascade metadata stays available; only the live renderer resolves styles.
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
        XElement styledDefaults = TestAssert.NotNull(styled.CascadeLayers[0].Source);
        TestAssert.Equal("1600", styledDefaults.Element(OoxNamespaces.DrawingNamespace + "defRPr")?.Attribute("sz")?.Value ?? string.Empty);
        PptxSceneTextParagraph bare = TestAssert.NotNull(scene.Slides[0].SlideNodes[1].TextBody).Paragraphs[0];
        TestAssert.True(bare.CascadeLayers.All(layer => layer.Source is null), "Bare paragraphs have no inherited style source.");
        var frames = PptxRenderer.InspectTextFrameModels(document, package, 0);
        TestAssert.Equal(16d, frames[0].Paragraphs[0].Runs[0].FontSize);
        TestAssert.Equal(18d, frames[1].Paragraphs[0].Runs[0].FontSize);
    }

    public static void SceneCascadeLayersMatchRendererSnapshot()
    {
        // R14-deeper: scene-retained cascade layers must reproduce the renderer snapshot
        // layer names and kinds (the matched placeholder inherits layout lstStyle).
        string input = TestFixtures.PublicCasePath("pptx-ladder-04-placeholder-inherit-matched.pptx");
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
        using var stream = TestFixtures.CreateZipPackage(new Dictionary<string, string>
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

        using OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        PptxSceneNode node = scene.Slides[0].SlideNodes[0];

        IReadOnlyList<PptxRenderer.PptxPositionedTextSpan> xmlSpans = ReadXmlSpans(node, document, scene);
        var resolver = CreateCannedPresentationResolver("memory:r14-scene-fed");
        IReadOnlyList<PptxRenderer.PptxPositionedTextSpan> fedSpans = PptxRenderer.BuildSceneFedTextSpansAsync(node, document, scene.Slides[0].Theme, scene.Slides[0].SlideColorMap, 1, false, InheritedSources(scene), resolver, CancellationToken.None).GetAwaiter().GetResult();
        TestAssert.Equal("Alpha Beta\nGamma", string.Join("\n", xmlSpans
            .GroupBy(span => span.ParagraphIndex)
            .Select(paragraph => string.Concat(paragraph.Select(span => span.Run.Text)))));
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

        PptxSceneTextBody body = TestAssert.NotNull(node.TextBody);
        PptxSceneTextRun[] sceneRuns = body.Paragraphs
            .SelectMany(paragraph => paragraph.Runs)
            .ToArray();

        var chunks = new List<(PptxRenderer.PptxTextRunModel Model, StringBuilder Text)>();
        foreach (PptxRenderer.PptxPositionedTextSpan span in xmlSpans)
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

        TestAssert.Equal("Alpha |Beta|Gamma", string.Join("|", sceneRuns.Select(run => run.Text)));
        TestAssert.Equal(sceneRuns.Length, chunks.Count);
        for (int i = 0; i < sceneRuns.Length; i++)
        {
            TestAssert.Equal(sceneRuns[i].Text, chunks[i].Text.ToString());
        }
        TestAssert.Equal(18d, chunks[0].Model.Style.NominalFontSize);
        TestAssert.True(chunks[0].Model.Style.Bold, "Alpha is explicitly bold.");
        TestAssert.Equal(new RgbColor(0x11,0x22,0x33), chunks[0].Model.Style.Color);
        TestAssert.Equal(12d, chunks[1].Model.Style.NominalFontSize);
        TestAssert.True(chunks[1].Model.Style.Italic, "Beta is explicitly italic.");
        TestAssert.Equal(new RgbColor(0x44,0x55,0x66), chunks[1].Model.Style.Color);
        TestAssert.Equal(14d, chunks[2].Model.Style.NominalFontSize);
        TestAssert.True(chunks[2].Model.Style.Underline && chunks[2].Model.Style.Strike, "Gamma retains both decorations.");
        TestAssert.Equal(new RgbColor(0x77,0x88,0x99), chunks[2].Model.Style.Color);
    }

    public static void SceneFedSpansMatchXmlPathForPlaceholders()
    {
        // R14-deeper: placeholder inheritance resolves through retained defaults on
        // both paths (unmatched body placeholder and layout-matched placeholder).
        string[] inputs = new[]
        {
            TestFixtures.PublicCasePath("pptx-ladder-04-placeholder-inherit.pptx"),
            TestFixtures.PublicCasePath("pptx-ladder-04-placeholder-inherit-matched.pptx"),
        };
        foreach (string input in inputs)
        {
            using FileStream stream = File.OpenRead(input);
            using OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
            PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
            PptxSceneSlide slide = scene.Slides.Single();
            PptxSceneNode node = slide.SlideNodes.Single(shape => shape.IsPlaceholder);
            var resolver = CreateCannedPresentationResolver("memory:r14-scene-fed-placeholders");

            IReadOnlyList<PptxRenderer.PptxPositionedTextSpan> xmlSpans = ReadXmlSpans(node, document, scene,
                includePlaceholders: true, resolver: resolver);
            // The inherited bullet can be attached to the first text span.
            string layoutText = string.Concat(xmlSpans.Select(span => span.Run.Text));
            TestAssert.True(layoutText is "BodyText" or "\u2022BodyText",
                $"The XML layout must retain the authored BodyText in {Path.GetFileName(input)}. Found '{layoutText}'.");
            IReadOnlyList<PptxRenderer.PptxPositionedTextSpan> fedSpans = PptxRenderer.BuildSceneFedTextSpansAsync(
                node, document, slide.Theme, slide.SlideColorMap, 1, true, InheritedSources(scene), resolver,
                CancellationToken.None).GetAwaiter().GetResult();
            TestAssert.Equal(xmlSpans.Count, fedSpans.Count);
            for (int index = 0; index < xmlSpans.Count; index++)
            {
                TestAssert.Equal(xmlSpans[index].Run.Text, fedSpans[index].Run.Text);
                TestAssert.Equal(xmlSpans[index].Run.X, fedSpans[index].Run.X);
                TestAssert.Equal(xmlSpans[index].Run.Y, fedSpans[index].Run.Y);
                TestAssert.Equal(xmlSpans[index].Run.FontSize, fedSpans[index].Run.FontSize);
            }
        }
    }

    public static void SceneFedSpansMatchXmlPathForField()
    {
        string fieldInput = TestFixtures.PublicCasePath("pptx-ladder-04-field-text.pptx");
        CompareXmlAndFedSpans(fieldInput, ("1", PptxRenderer.PptxTextRunKind.Field, null));
    }

    private static (PptxSceneTextBody Body, IReadOnlyList<PptxRenderer.PptxPositionedTextSpan> XmlSpans) CompareXmlAndFedSpans(string input,
        params (string Text, PptxRenderer.PptxTextRunKind Kind, string? ClickId)[] expected)
    {
        using FileStream stream = File.OpenRead(input);
        using OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        PptxSceneSlide slide = scene.Slides.Single();
        PptxSceneNode node = slide.SlideNodes.Single(shape => shape.TextBody is not null);
        var resolver = CreateCannedPresentationResolver("memory:r14-scene-fed-links-fields");
        IReadOnlyList<PptxRenderer.PptxPositionedTextSpan> xmlSpans = ReadXmlSpans(node, document, scene, resolver: resolver);
        IReadOnlyList<PptxRenderer.PptxPositionedTextSpan> fedSpans = PptxRenderer.BuildSceneFedTextSpansAsync(
            node, document, slide.Theme, slide.SlideColorMap, 1, false, InheritedSources(scene),
            resolver, CancellationToken.None).GetAwaiter().GetResult();
        TestAssert.Equal(expected.Length, xmlSpans.Count);
        TestAssert.Equal(expected.Length, fedSpans.Count);
        for (int index = 0; index < expected.Length; index++)
        {
            var xml = xmlSpans[index];
            var fed = fedSpans[index];
            var xmlSource = TestAssert.NotNull(xml.SourceRun);
            var fedSource = TestAssert.NotNull(fed.SourceRun);
            TestAssert.Equal(expected[index], (xml.Run.Text, xmlSource.Kind, xmlSource.Style.HyperlinkClickId));
            TestAssert.Equal(expected[index], (fed.Run.Text, fedSource.Kind, fedSource.Style.HyperlinkClickId));
            TestAssert.Equal(xml.Run.X, fed.Run.X);
            TestAssert.Equal(xml.Run.Y, fed.Run.Y);
            TestAssert.Equal(xml.Run.FontSize, fed.Run.FontSize);
        }
        return (TestAssert.NotNull(node.TextBody), xmlSpans);
    }

    public static void SceneFedSpansMatchXmlPathForGroupedShape()
    {
        using var input = TestFixtures.CreateZipPackage(new Dictionary<string, string>
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
                    <p:grpSpPr><a:xfrm><a:off x="228600" y="0"/><a:ext cx="2743200" cy="1828800"/><a:chOff x="0" y="0"/><a:chExt cx="2743200" cy="1828800"/></a:xfrm></p:grpSpPr>
                    <p:sp><p:nvSpPr><p:cNvPr id="11" name="GroupedShape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
                      <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                      <p:txBody>
                        <a:bodyPr lIns="0" tIns="0" bIns="0"/><a:lstStyle/>
                        <a:p><a:r><a:rPr sz="1800" b="1"><a:solidFill><a:srgbClr val="334455"/></a:solidFill></a:rPr><a:t>Grouped</a:t></a:r></a:p>
                      </p:txBody>
                    </p:sp>
                  </p:grpSp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using var package = OoxPackage.Open(input, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        PptxSceneNode group = scene.Slides[0].SlideNodes.Single();
        TestAssert.Equal(PptxSceneNodeKind.Group, group.Kind);
        PptxSceneNode node = group.Children.Single();
        PptxSceneTextBody body = TestAssert.NotNull(node.TextBody);
        TestAssert.Equal("Grouped", body.Paragraphs.SelectMany(paragraph => paragraph.Runs).Single().Text);

        var resolver = CreateCannedPresentationResolver("memory:grouped-text-agreement");
        var xmlSpan = ReadXmlSpans(node, document, scene, resolver: resolver).Single();
        var fedSpan = PptxRenderer.BuildSceneFedTextSpansAsync(node, document, scene.Slides[0].Theme,
            scene.Slides[0].SlideColorMap, 1, false, InheritedSources(scene), resolver, CancellationToken.None)
            .GetAwaiter().GetResult().Single();
        var model = TestAssert.NotNull(xmlSpan.SourceRun);
        TestAssert.Equal("Grouped", model.Text);
        TestAssert.Equal(18d, model.Style.NominalFontSize);
        TestAssert.Equal(new RgbColor(0x33, 0x44, 0x55), model.Style.Color);
        TestAssert.True(model.Style.Bold, "The grouped run is explicitly bold.");
        TestAssert.Equal(18d, xmlSpan.Run.X);
        TestAssert.Equal(xmlSpan.Run.Text, fedSpan.Run.Text);
        TestAssert.Equal(xmlSpan.Run.X, fedSpan.Run.X);
        TestAssert.Equal(xmlSpan.Run.Y, fedSpan.Run.Y);
        TestAssert.Equal(xmlSpan.Run.FontSize, fedSpan.Run.FontSize);
    }

    public static void TableCellSpansPreserveSourceContentAndDeclaredFrame()
    {
        string input = TestFixtures.PublicCasePath("pptx-ladder-10-table-style-text.pptx");
        using FileStream stream = File.OpenRead(input);
        using OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        var fonts = new PresentationFontResolver(new PortableRenderingTests.PinnedResolver());
        var cells = PptxRenderer.InspectTextGlyphRuns(document, package, 0, inspectionFonts: fonts)
            .Where(run => run.TableRowIndex.HasValue)
            .GroupBy(run => (run.TableRowIndex, run.TableColumnIndex))
            .ToDictionary(group => group.Key, group => group.ToArray());

        // The authored grid has two 216pt cells starting at x=72pt and a 72pt row.
        TestAssert.Equal(2, cells.Count);
        foreach (var (column, left, text) in new[] { (0, 72d, "Head"), (1, 288d, "Body") })
        {
            PptxTextGlyphRunSnapshot[] runs = cells[(0, column)];
            TestAssert.Equal(text, string.Concat(runs.Select(run => run.Text)));
            foreach (PptxTextGlyphRunSnapshot run in runs)
            {
                TestAssert.True(run.X >= left && run.X + run.Width <= left + 216d,
                    "Cell text must stay within its authored horizontal frame.");
                TestAssert.True(run.BaselineY >= document.SlideHeightPoints - 144d &&
                    run.BaselineY <= document.SlideHeightPoints - 72d,
                    "Cell baselines must stay within the authored row.");
            }
        }
    }

    public static void PlainShapeFullRunStyleAgrees()
    {
        // Source text and live rendering retain the declared formatting.
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
            .ToArray();
        TestAssert.Equal(5, sceneRuns.Length);

        IReadOnlyList<PptxRenderer.PptxPositionedTextSpan> positioned = ReadXmlSpans(node, document, scene);
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
            TestAssert.Equal(18d, spanRuns[i].Style.NominalFontSize);
            TestAssert.Equal(new RgbColor(0x11,0x22,0x33), spanRuns[i].Style.Color);
        }
        TestAssert.Equal("allCaps", spanRuns[0].Style.CapsValue);
        TestAssert.Equal(2d, spanRuns[0].Style.CharacterSpacing);
        TestAssert.True(Math.Abs(5.4d - spanRuns[1].Style.BaselineOffset) < 1e-12,
            "An 18pt run with 30% baseline retains its 5.4pt offset within floating-point precision.");
        TestAssert.True(spanRuns[1].Style.Underline, "B retains double underline.");
        TestAssert.Equal("dbl", spanRuns[1].Style.UnderlineValue);
        TestAssert.True(spanRuns[2].Style.Strike, "C retains double strike.");
        TestAssert.Equal("dblStrike", spanRuns[2].Style.StrikeValue);
        TestAssert.Equal("Arial", spanRuns[2].Style.Typeface);
        TestAssert.Equal(new RgbColor(255,255,0), spanRuns[3].Style.Highlight?.Color ?? default);
        TestAssert.Equal(1d, spanRuns[3].Style.Highlight?.Alpha ?? 0d);
        TestAssert.Equal(0.5d, spanRuns[4].Style.Alpha);
    }

    public static void BreakRunsAgreeBetweenSceneAndSpans()
    {
        // Source text and live rendering retain the declared formatting.
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
        TestAssert.Equal("\n", sceneRuns[1].Text);

        IReadOnlyList<PptxRenderer.PptxPositionedTextSpan> positioned = ReadXmlSpans(node, document, scene);
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
        // Source text and live rendering retain the declared formatting.
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
        TestAssert.Equal("1", sceneRuns[1].Text);

        IReadOnlyList<PptxRenderer.PptxPositionedTextSpan> positioned = ReadXmlSpans(node, document, scene);
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
            TestAssert.Equal(18d, chunks[i].Model.Style.NominalFontSize);
            TestAssert.Equal(new RgbColor(0,0,0), chunks[i].Model.Style.Color);
        }
    }

    public static void ParagraphAlignmentAgreesBetweenSceneAndSpans()
    {
        // Source text and live rendering retain the declared formatting.
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
        TestAssert.Equal(2, body.Paragraphs.Count);
        PptxTextParagraphLayoutSnapshot[] layoutParagraphs = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames.SelectMany(frame => frame.Paragraphs).ToArray();
        TestAssert.Equal(2, layoutParagraphs.Length);
        TestAssert.Equal("Center", layoutParagraphs[0].Lines[0].Alignment);
        TestAssert.Equal("Right", layoutParagraphs[1].Lines[0].Alignment);
    }

    public static void PlaceholderTextAgreesBetweenSceneAndSpans()
    {
        var run = ReadPlaceholderAgreementRun("BodyText", matchLayoutPlaceholder: false);
        TestAssert.Equal(25d, run.Style.FontSize);
        TestAssert.Equal(new RgbColor(0xAA, 0xBB, 0xCC), run.Style.Color);
    }

    public static void MatchedPlaceholderInheritanceAgrees()
    {
        var run = ReadPlaceholderAgreementRun("MatchedBody", matchLayoutPlaceholder: true);
        TestAssert.Equal(26d, run.Style.FontSize);
    }

    private static PptxRenderer.PptxTextRunModel ReadPlaceholderAgreementRun(string caption, bool matchLayoutPlaceholder)
    {
        string slide = $"""
            <?xml version="1.0" encoding="UTF-8"?>
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld><p:spTree><p:sp>
                <p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
                <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4572000" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                <p:txBody>
                  <a:bodyPr tIns="0" bIns="0"/><a:lstStyle/>
                  <a:p><a:r><a:t>{caption}</a:t></a:r></a:p>
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
        string layoutShape = matchLayoutPlaceholder ? """
            <p:sp>
              <p:nvSpPr><p:cNvPr id="2" name="Body Placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
              <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4572000" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
              <p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr sz="2600"/></a:lvl1pPr></a:lstStyle><a:p/></p:txBody>
            </p:sp>
            """ : string.Empty;
        string layout = $"""
            <?xml version="1.0" encoding="UTF-8"?>
            <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld><p:spTree>{layoutShape}</p:spTree></p:cSld>
            </p:sldLayout>
            """;
        string layoutRels = """
            <?xml version="1.0" encoding="UTF-8"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
            </Relationships>
            """;
        string titleStyle = matchLayoutPlaceholder ? string.Empty :
            """<p:titleStyle><a:lvl1pPr><a:defRPr sz="3100"/></a:lvl1pPr></p:titleStyle>""";
        string bodyRunProperties = matchLayoutPlaceholder ? """<a:defRPr sz="2500"/>""" :
            """<a:defRPr sz="2500"><a:solidFill><a:srgbClr val="AABBCC"/></a:solidFill></a:defRPr>""";
        string master = $"""
            <?xml version="1.0" encoding="UTF-8"?>
            <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld><p:spTree/></p:cSld>
              <p:txStyles>
                {titleStyle}
                <p:bodyStyle><a:lvl1pPr>{bodyRunProperties}</a:lvl1pPr></p:bodyStyle>
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
        using OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        PptxSceneSlide sceneSlide = scene.Slides.Single();
        PptxSceneNode node = sceneSlide.SlideNodes.Single();
        PptxSceneTextParagraph paragraph = TestAssert.NotNull(node.TextBody).Paragraphs.Single();
        TestAssert.Equal(caption, paragraph.Runs.Single().Text);
        TestAssert.Equal(0, paragraph.Level);
        XElement slideShape = sceneSlide.SlideXml.Descendants("{http://schemas.openxmlformats.org/presentationml/2006/main}sp").Single();
        TestAssert.True(ReferenceEquals(node.Source, slideShape), "node-source-detached");

        PptxRenderer.PptxTextRunModel span = ReadXmlSpans(node, document, scene, includePlaceholders: true)
            .Select(positioned => positioned.SourceRun)
            .OfType<PptxRenderer.PptxTextRunModel>()
            .Distinct().Single();
        TestAssert.Equal(PptxRenderer.PptxTextRunKind.Text, span.Kind);
        TestAssert.Equal(caption, span.Text);
        return span;
    }

    public static void HyperlinkRunAgreesBetweenSceneAndSpans()
    {
        // Source text and live rendering retain the declared formatting.
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

        var (body, xmlSpans) = CompareXmlAndFedSpans(input,
            ("Link", PptxRenderer.PptxTextRunKind.Text, "rIdLink"),
            ("Plain", PptxRenderer.PptxTextRunKind.Text, null));
        TestAssert.Equal(2, body.Paragraphs.Sum(paragraph => paragraph.Runs.Count));
        var linked = TestAssert.NotNull(xmlSpans[0].SourceRun);
        var plain = TestAssert.NotNull(xmlSpans[1].SourceRun);
        TestAssert.Equal("Link", linked.Text);
        TestAssert.True(linked.Style.HasHyperlinkClick, "Expected the linked run to carry a hyperlink click.");
        TestAssert.True(!plain.Style.HasHyperlinkClick, "Expected the plain run to carry no hyperlink click.");
        TestAssert.True(linked.Style.Underline, "Hyperlinks receive the default underline.");
        // This fixture intentionally has no theme part: default underline is
        // preserved, while the unresolved theme hyperlink color falls back.
        TestAssert.Equal(PptxRenderer.PptxRunTextColorSource.FallbackBlack, linked.Style.ColorSource);
    }

    public static void UnstyledRunsInheritParagraphDefaults()
    {
        string slide = """
            <?xml version="1.0" encoding="UTF-8"?>
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="27" name="Probe27"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4572000" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                <p:txBody>
                  <a:bodyPr tIns="0" bIns="0"/><a:lstStyle/>
                  <a:p><a:pPr><a:defRPr sz="2000"><a:solidFill><a:srgbClr val="993366"/></a:solidFill></a:defRPr></a:pPr><a:r><a:t>Bare</a:t></a:r></a:p>
                </p:txBody>
              </p:sp></p:spTree></p:cSld>
            </p:sld>
            """;
        AssertUnstyledRunCascade(PptxTests.BasicSlidePackageEntries(slide), "Bare", 20d, new RgbColor(0x99, 0x33, 0x66));
    }

    public static void UnstyledRunsInheritShapeListStyle()
    {
        string slide = """
            <?xml version="1.0" encoding="UTF-8"?>
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="28" name="Probe28"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4572000" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                <p:txBody>
                  <a:bodyPr tIns="0" bIns="0"/><a:lstStyle><a:lvl1pPr><a:defRPr sz="2200"><a:solidFill><a:srgbClr val="334455"/></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle>
                  <a:p><a:r><a:t>Styled</a:t></a:r></a:p>
                </p:txBody>
              </p:sp></p:spTree></p:cSld>
            </p:sld>
            """;
        AssertUnstyledRunCascade(PptxTests.BasicSlidePackageEntries(slide), "Styled", 22d, new RgbColor(0x33, 0x44, 0x55));
    }

    public static void UnstyledRunsInheritPresentationDefaultStyle()
    {
        string layout = """
            <?xml version="1.0" encoding="UTF-8"?>
            <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree/></p:cSld></p:sldLayout>
            """;
        string master = """
            <?xml version="1.0" encoding="UTF-8"?>
            <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld><p:spTree/></p:cSld>
            </p:sldMaster>
            """;
        string slide = """
            <?xml version="1.0" encoding="UTF-8"?>
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="29" name="Probe29"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4572000" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                <p:txBody>
                  <a:bodyPr tIns="0" bIns="0"/><a:lstStyle/>
                  <a:p><a:r><a:t>Defaulted</a:t></a:r></a:p>
                </p:txBody>
              </p:sp></p:spTree></p:cSld>
            </p:sld>
            """;
        var entries = PptxTests.InheritedShapePackageEntries(master, layout, slide);
        const string defaultTextStyle = "<p:defaultTextStyle xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
            "<a:lvl1pPr><a:defRPr sz=\"2300\"><a:solidFill><a:srgbClr val=\"556677\"/></a:solidFill>" +
            "</a:defRPr></a:lvl1pPr></p:defaultTextStyle>";
        entries["ppt/presentation.xml"] = PptxTests.BasicPresentation().Replace("</p:presentation>",
            defaultTextStyle + "</p:presentation>", StringComparison.Ordinal);
        AssertUnstyledRunCascade(entries, "Defaulted", 23d, new RgbColor(0x55, 0x66, 0x77));
    }

    private static void AssertUnstyledRunCascade(IReadOnlyDictionary<string, string> entries, string text, double fontSize, RgbColor color)
    {
        using MemoryStream stream = TestFixtures.CreateZipPackage(entries);
        using OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        PptxSceneNode node = scene.Slides[0].SlideNodes[0];
        PptxSceneTextBody body = TestAssert.NotNull(node.TextBody);
        PptxSceneTextRun[] sceneRuns = body.Paragraphs
            .SelectMany(paragraph => paragraph.Runs)
            .ToArray();
        TestAssert.Equal(1, sceneRuns.Length);

        IReadOnlyList<PptxRenderer.PptxPositionedTextSpan> positioned = ReadXmlSpans(node, document, scene);
        PptxRenderer.PptxTextRunModel[] spanRuns = positioned
            .Select(span => span.SourceRun)
            .Where(run => run is not null)
            .Select(run => run!)
            .Distinct()
            .ToArray();
        TestAssert.Equal(1, spanRuns.Length);
        TestAssert.Equal(text, spanRuns[0].Text);
        TestAssert.Equal(fontSize, spanRuns[0].Style.NominalFontSize);
        TestAssert.Equal(color, spanRuns[0].Style.Color);
    }

    public static void TableCellTextMatchesShapeText()
    {
        // Source text and live rendering retain the declared formatting.
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
            .ToArray();
        TestAssert.Equal(2, shapeRuns.Length);

        var context = TestAssert.NotNull(PptxRenderer.TryLoadTextInspectionContext(document, package, 0, scene,
            CreateCannedPresentationResolver("memory:r14-table")));
        PptxRenderer.PptxTextRunModel[] cellRuns = ReadTableCellRuns(context);
        TestAssert.Equal(2, cellRuns.Length);
        var shapeModels = ReadXmlSpans(shapeNode, document, scene).Select(span => span.SourceRun!).Distinct().ToArray();
        for (int i = 0; i < shapeRuns.Length; i++)
        {
            TestAssert.Equal(shapeRuns[i].Text, cellRuns[i].Text);
            TestAssert.Equal(shapeModels[i].Style.NominalFontSize, cellRuns[i].Style.NominalFontSize);
            TestAssert.Equal(shapeModels[i].Style.Color, cellRuns[i].Style.Color);
            TestAssert.Equal(shapeModels[i].Style.Bold, cellRuns[i].Style.Bold);
            TestAssert.Equal(shapeModels[i].Style.Italic, cellRuns[i].Style.Italic);
        }
    }

    public static void TableStyleTextAgreesBetweenSceneAndSpans()
    {
        const string theme = """
            <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="TableTheme">
              <a:themeElements>
                <a:clrScheme name="TableTheme">
                  <a:dk1><a:srgbClr val="000000"/></a:dk1>
                  <a:lt1><a:srgbClr val="F2E6C9"/></a:lt1>
                  <a:accent6><a:srgbClr val="336699"/></a:accent6>
                </a:clrScheme>
                <a:fontScheme name="TableTheme"><a:majorFont/><a:minorFont/></a:fontScheme>
              </a:themeElements>
            </a:theme>
            """;
        using var input = PptxTests.CreateThemedSlidePackage(theme, """
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld><p:spTree>
                <p:graphicFrame>
                  <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="1828800"/></p:xfrm>
                  <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                    <a:tbl>
                      <a:tblPr firstRow="1"><a:tableStyleId>{93296810-A885-4BE3-A3E7-6D5BEEA58F35}</a:tableStyleId></a:tblPr>
                      <a:tblGrid><a:gridCol w="1828800"/></a:tblGrid>
                      <a:tr h="914400">
                        <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Head</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                      </a:tr>
                      <a:tr h="914400">
                        <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>Body</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                      </a:tr>
                    </a:tbl>
                  </a:graphicData></a:graphic>
                </p:graphicFrame>
              </p:spTree></p:cSld>
            </p:sld>
            """);

        using OoxPackage package = OoxPackage.Open(input, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        PptxSceneTable table = TestAssert.NotNull(scene.Slides.Single().SlideNodes.Single().Table);
        TestAssert.Equal(2, table.Rows.Count);
        var headerStyle = table.Rows[0].Cells.Single().StyleText;
        var bodyStyle = table.Rows[1].Cells.Single().StyleText;
        var headerColor = new RgbColor(242, 230, 201);
        TestAssert.Equal(new PptxSceneTableCellTextStyle(headerColor, true), headerStyle);
        TestAssert.Equal(default(PptxSceneTableCellTextStyle), bodyStyle);

        var fonts = new PresentationFontResolver(new PortableRenderingTests.PinnedResolver());
        var context = TestAssert.NotNull(PptxRenderer.TryLoadTextInspectionContext(document, package, 0, scene, fonts));
        var spans = PptxRenderer.ReadSceneTableLayoutsAsync(context).ToBlockingEnumerable().Single().TextSpans;
        TestAssert.Equal(2, spans.Count);
        AssertRow(spans[0], 0, "Head", headerColor, true);
        AssertRow(spans[1], 1, "Body", new RgbColor(0, 0, 0), false);
        TestAssert.Equal(14d, TestAssert.NotNull(spans[1].SourceRun).Style.NominalFontSize);

        static void AssertRow(PptxRenderer.PptxPositionedTextSpan span, int row, string text, RgbColor color, bool bold)
        {
            TestAssert.Equal((int?)row, span.TableRowIndex);
            TestAssert.Equal((int?)0, span.TableColumnIndex);
            TestAssert.Equal(text, span.Run.Text);
            TestAssert.Equal(color, span.Run.Color);
            TestAssert.Equal(bold, span.Run.Bold);
        }
    }

    private static PptxRenderer.PptxTextRunModel[] ReadTableCellRuns(PptxRenderContext context) =>
        PptxRenderer.ReadSceneTableLayoutsAsync(context).ToBlockingEnumerable().Single().TextSpans
            .Select(span => span.SourceRun)
            .OfType<PptxRenderer.PptxTextRunModel>()
            .Distinct()
            .ToArray();


    public static void SchemeColorRunsAgreeBetweenSceneAndSpans()
    {
        // Source text and live rendering retain the declared formatting.
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
            .ToArray();
        TestAssert.Equal(2, sceneRuns.Length);

        IReadOnlyList<PptxRenderer.PptxPositionedTextSpan> positioned = ReadXmlSpans(node, document, scene);
        PptxRenderer.PptxTextRunModel[] spanRuns = positioned
            .Select(span => span.SourceRun)
            .Where(run => run is not null)
            .Select(run => run!)
            .Distinct()
            .ToArray();
        TestAssert.Equal(2, spanRuns.Length);
        TestAssert.Equal("Acc", spanRuns[0].Text);
        TestAssert.Equal("Dark", spanRuns[1].Text);
        TestAssert.Equal(new RgbColor(0x33,0x66,0x99), spanRuns[0].Style.Color);
        TestAssert.Equal(new RgbColor(0x11,0x11,0x11), spanRuns[1].Style.Color);
    }


    internal static IReadOnlyList<PptxRenderer.PptxPositionedTextSpan> ReadXmlSpans(
        PptxSceneNode node,
        PptxDocument document,
        PptxScene scene,
        bool includePlaceholders = false,
        PresentationFontResolver? resolver = null)
    {
        resolver ??= CreateCannedPresentationResolver("memory:r14-agreement");
        const BindingFlags flags = BindingFlags.NonPublic | BindingFlags.Static;
        try
        {
            // Keep the independent XML text projection for the same admitted
            // shape. The scene-fed path resolves its own paragraph models.
            var context = new PptxRenderContext(document, document.Slides[0], scene.Slides[0], resolver, null, null, default);
            MethodInfo read = typeof(PptxRenderer).GetMethod("BuildTextFrameModelsAsync", flags)!;
            object frameModels = AwaitResult(read.Invoke(null,
                [context, new[] { node }, context.SceneSlide.SlideColorMap, includePlaceholders, context.InheritedXml])!);
            MethodInfo build = typeof(PptxRenderer).GetMethod("BuildTextLayoutModelAsync", flags)!;
            object layout = AwaitResult(build.Invoke(null, [context, frameModels])!);
            var frames = (System.Collections.IEnumerable)layout.GetType().GetProperty("Frames")!.GetValue(layout)!;
            TestAssert.True(frames.Cast<object>().All(frame =>
            {
                object model = frame.GetType().GetProperty("Model")!.GetValue(frame)!;
                return ReferenceEquals(model.GetType().GetProperty("Shape")!.GetValue(model), node.Source);
            }), "The XML agreement fixture must contain only the selected text shape.");
            MethodInfo flatten = typeof(PptxRenderer).GetMethod("FlattenTextLayoutToSpansAsync", flags)!;
            var spans = ((ValueTask<IReadOnlyList<PptxRenderer.PptxPositionedTextSpan>>)flatten.Invoke(null, [layout, resolver, document])!).GetAwaiter().GetResult();
            MethodInfo remap = typeof(PptxRenderer).GetMethod("RemapMongolianVerticalSpans", flags)!;
            return (IReadOnlyList<PptxRenderer.PptxPositionedTextSpan>)remap.Invoke(null, [spans, node.Source])!;
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }

        static object AwaitResult(object pending)
        {
            var task = (Task)pending.GetType().GetMethod("AsTask")!.Invoke(pending, null)!;
            task.GetAwaiter().GetResult();
            return task.GetType().GetProperty("Result")!.GetValue(task)!;
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
