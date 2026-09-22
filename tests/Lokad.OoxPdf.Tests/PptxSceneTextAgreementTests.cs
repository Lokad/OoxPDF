using System.Reflection;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxSceneTextAgreementTests
{
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
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4572000" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
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

        byte[] bytes = TestFontBuilder.CreateTestFont();
        OpenTypeFont font = OpenTypeFont.Load(bytes);
        var resolution = new FontFaceResolution(
            font.FamilyName,
            font.FamilyName,
            new FontStyleKey(),
            new MemoryFontProgramSource("memory:r14-agreement", bytes),
            IsFallback: false);
        var resolver = new PresentationFontResolver(new CannedFontResolver(resolution));
        MethodInfo spans = typeof(PptxRenderer).GetMethods(BindingFlags.NonPublic | BindingFlags.Static)
            .Single(candidate => candidate.Name == "ReadTextSpansForShape" && candidate.GetParameters().Length == 9);
        object? result;
        try
        {
            result = spans.Invoke(null, [node.Source, document, scene.Theme, scene.Slides[0].SlideColorMap, 1, false, Array.Empty<XDocument>(), resolver, CancellationToken.None]);
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }

        IReadOnlyList<PptxRenderer.PptxPositionedTextSpan> positioned = (IReadOnlyList<PptxRenderer.PptxPositionedTextSpan>)result!;
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

    private sealed class CannedFontResolver(FontFaceResolution resolution) : IFontResolver
    {
        public FontFaceResolution Resolve(FontRequest request)
        {
            return resolution;
        }
    }
}
