using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxDiagnosticsTests
{
    public static void PptxSyntheticLayoutAndMasterShapesRender()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
                  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/_rels/slide1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
                </Relationships>
                """,
            ["ppt/slideLayouts/_rels/slideLayout1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
                </Relationships>
                """,
            ["ppt/slideMasters/slideMaster1.xml"] = PptxTests.InheritedShapePart("FF0000", 914400, 914400),
            ["ppt/slideLayouts/slideLayout1.xml"] = PptxTests.InheritedShapePart("00FF00", 1828800, 1828800),
            ["ppt/slides/slide1.xml"] = PptxTests.InheritedShapePart("0000FF", 2743200, 2743200)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.Contains("0 1 0 rg", pdf);
        TestAssert.Contains("0 0 1 rg", pdf);
    }

    public static void PptxSyntheticLayoutPictureUsesLayoutRelationships()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
                  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
                </Relationships>
                """),
            ["ppt/slideLayouts/_rels/slideLayout1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
                  <Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/layout-image.png"/>
                </Relationships>
                """),
            ["ppt/slideMasters/slideMaster1.xml"] = TestFixtures.Utf8(PptxTests.InheritedShapePart("FF0000", 914400, 914400)),
            ["ppt/slideLayouts/slideLayout1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:pic>
                    <p:blipFill><a:blip r:embed="rIdImage"/></p:blipFill>
                    <p:spPr><a:xfrm><a:off x="1828800" y="1828800"/><a:ext cx="914400" cy="914400"/></a:xfrm></p:spPr>
                  </p:pic></p:spTree></p:cSld>
                </p:sldLayout>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8(PptxTests.InheritedShapePart("0000FF", 2743200, 2743200)),
            ["ppt/media/layout-image.png"] = TestFixtures.CreateRgbPng(1, 1, [255, 0, 0])
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/Subtype /Image", pdf);
        TestAssert.Contains("/Im1 Do", pdf);
    }

    public static void PptxSyntheticInheritedPlaceholderTextIsSkipped()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
                  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/_rels/slide1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
                </Relationships>
                """,
            ["ppt/slideLayouts/_rels/slideLayout1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
                </Relationships>
                """,
            ["ppt/slideMasters/slideMaster1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:nvSpPr><p:cNvPr id="2" name="Body Placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400"/><a:t>Click to edit Master text styles</a:t></a:r></a:p></p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """,
            ["ppt/slideLayouts/slideLayout1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree/></p:cSld></p:sld>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree/></p:cSld></p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(!pdf.Contains("/Subtype /Type0", StringComparison.Ordinal), "Inherited placeholder text should not create PDF font resources.");
    }

    public static void PptxSyntheticSlidePlaceholderTextUsesInheritedBounds()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
                  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/_rels/slide1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
                </Relationships>
                """,
            ["ppt/slideLayouts/_rels/slideLayout1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
                </Relationships>
                """,
            ["ppt/slideMasters/slideMaster1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="4000"/></a:lvl1pPr></p:titleStyle></p:txStyles>
                </p:sldMaster>
                """,
            ["ppt/slideLayouts/slideLayout1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:nvSpPr><p:cNvPr id="2" name="Title Placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400"/><a:t>Layout title</a:t></a:r></a:p></p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:nvSpPr><p:cNvPr id="3" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
                    <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>Slide title</a:t></a:r></a:p></p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/F1 39.96 Tf", pdf);
        PptxTests.AssertContainsTextMatrixAtX(pdf, 79.2d);

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxTextFrameModelSnapshot textFrame = PptxRenderer.InspectTextFrameModels(document, package, 0)
            .Single(frame => frame.Paragraphs.Any(paragraph => paragraph.Runs.Any(run => run.Text == "Slide title")));
        TestAssert.True(textFrame.InheritedPlaceholderCount >= 1, "Expected text model to expose inherited placeholder participation.");
        TestAssert.True(textFrame.HasInheritedTextBody, "Expected text model to expose inherited placeholder text body participation.");
        TestAssert.True(textFrame.UsesInheritedShapeBounds, "Expected text model to expose placeholder geometry fallback.");
    }

    public static void PptxUnsupportedMediaDiagnosticsUseScenePictureState()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:pic>
                      <p:nvPicPr><p:cNvPr id="2" name="Video"/><p:cNvPicPr/><p:nvPr><p:video/></p:nvPr></p:nvPicPr>
                      <p:blipFill><a:blip><a:videoFile/></a:blip></p:blipFill>
                    </p:pic>
                    <p:pic>
                      <p:nvPicPr><p:cNvPr id="3" name="Audio"/><p:cNvPicPr/><p:nvPr><p:audio/></p:nvPr></p:nvPicPr>
                      <p:blipFill><a:blip><a:audioFile/></a:blip></p:blipFill>
                    </p:pic>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        PptxSceneSlide sceneSlide = scene.Slides[0];
        TestAssert.True(sceneSlide.SlideNodes.Any(node => node.Picture?.HasVideo == true), "Expected video provenance to be owned by the scene picture.");
        TestAssert.True(sceneSlide.SlideNodes.Any(node => node.Picture?.HasAudio == true), "Expected audio provenance to be owned by the scene picture.");

        PptxSceneSnapshot snapshot = PptxRenderer.InspectScene(document, package);
        TestAssert.True(snapshot.Slides[0].SlideNodes.Any(node => node.PictureHasVideo), "Expected video provenance to be visible in scene inspection.");
        TestAssert.True(snapshot.Slides[0].SlideNodes.Any(node => node.PictureHasAudio), "Expected audio provenance to be visible in scene inspection.");

        var diagnostics = new List<OoxPdfDiagnostic>();
        XDocument slideXmlWithoutMedia = XDocument.Parse("""
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                   xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld><p:spTree/></p:cSld>
            </p:sld>
            """);
        System.Reflection.MethodInfo emitDiagnostics = typeof(PptxRenderer).GetMethod(
            "EmitUnsupportedFeatureDiagnostics",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected unsupported feature diagnostic emitter.");
        Action<OoxPdfDiagnostic> sink = diagnostics.Add;
        emitDiagnostics.Invoke(null, [sceneSlide, slideXmlWithoutMedia, "/ppt/slides/slide1.xml", 1, sink, new HashSet<string>(StringComparer.Ordinal)]);

        string ids = string.Join("|", diagnostics.Select(d => d.Id));
        TestAssert.Contains("PPTX_UNSUPPORTED_AUDIO", ids);
        TestAssert.Contains("PPTX_UNSUPPORTED_VIDEO", ids);
    }

    public static void PptxUnsupportedSmartArtDiagnosticsUseSceneGraphicFrameState()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <a:graphic>
                        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>
                      </a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        PptxSceneSlide sceneSlide = scene.Slides[0];
        TestAssert.True(sceneSlide.SlideNodes.Any(node => node.IsSmartArtGraphicFrame), "Expected SmartArt classification to be owned by the scene node.");

        PptxSceneSnapshot snapshot = PptxRenderer.InspectScene(document, package);
        TestAssert.True(snapshot.Slides[0].SlideNodes.Any(node => node.IsSmartArtGraphicFrame), "Expected SmartArt classification to be visible in scene inspection.");

        var diagnostics = new List<OoxPdfDiagnostic>();
        XDocument slideXmlWithoutSmartArt = XDocument.Parse("""
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                   xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld><p:spTree/></p:cSld>
            </p:sld>
            """);
        System.Reflection.MethodInfo emitDiagnostics = typeof(PptxRenderer).GetMethod(
            "EmitUnsupportedFeatureDiagnostics",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected unsupported feature diagnostic emitter.");
        Action<OoxPdfDiagnostic> sink = diagnostics.Add;
        emitDiagnostics.Invoke(null, [sceneSlide, slideXmlWithoutSmartArt, "/ppt/slides/slide1.xml", 1, sink, new HashSet<string>(StringComparer.Ordinal)]);

        string ids = string.Join("|", diagnostics.Select(d => d.Id));
        TestAssert.Contains("PPTX_UNSUPPORTED_SMARTART", ids);
        TestAssert.DoesNotContain("PPTX_UNSUPPORTED_GRAPHIC_FRAME", ids);
    }

    public static void PptxUnsupportedDynamicDiagnosticsUseSceneSlideState()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:oleObj/></p:spTree></p:cSld>
                  <p:transition/>
                  <p:timing/>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        PptxSceneSlide sceneSlide = scene.Slides[0];
        TestAssert.True(sceneSlide.HasTransition, "Expected transition provenance to be owned by the scene slide.");
        TestAssert.True(sceneSlide.HasTiming, "Expected timing provenance to be owned by the scene slide.");
        TestAssert.True(sceneSlide.HasOleObject, "Expected OLE object provenance to be owned by the scene slide.");

        PptxSceneSnapshot snapshot = PptxRenderer.InspectScene(document, package);
        TestAssert.True(snapshot.Slides[0].HasTransition, "Expected transition provenance to be visible in scene inspection.");
        TestAssert.True(snapshot.Slides[0].HasTiming, "Expected timing provenance to be visible in scene inspection.");
        TestAssert.True(snapshot.Slides[0].HasOleObject, "Expected OLE object provenance to be visible in scene inspection.");

        var diagnostics = new List<OoxPdfDiagnostic>();
        XDocument slideXmlWithoutDynamicFeatures = XDocument.Parse("""
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                   xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld><p:spTree/></p:cSld>
            </p:sld>
            """);
        System.Reflection.MethodInfo emitDiagnostics = typeof(PptxRenderer).GetMethod(
            "EmitUnsupportedFeatureDiagnostics",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected unsupported feature diagnostic emitter.");
        Action<OoxPdfDiagnostic> sink = diagnostics.Add;
        emitDiagnostics.Invoke(null, [sceneSlide, slideXmlWithoutDynamicFeatures, "/ppt/slides/slide1.xml", 1, sink, new HashSet<string>(StringComparer.Ordinal)]);

        string ids = string.Join("|", diagnostics.Select(d => d.Id));
        TestAssert.Contains("PPTX_UNSUPPORTED_ANIMATION", ids);
        TestAssert.Contains("PPTX_UNSUPPORTED_OLE_OBJECT", ids);
        TestAssert.Contains("PPTX_UNSUPPORTED_TRANSITION", ids);
    }

    public static void PptxUnsupportedTransparencyDiagnosticsUseSceneShapeState()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:sp>
                      <p:spPr>
                        <a:effectLst>
                          <a:reflection><a:srgbClr val="123456"><a:alpha val="50000"/></a:srgbClr></a:reflection>
                        </a:effectLst>
                      </p:spPr>
                    </p:sp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        PptxSceneSlide sceneSlide = scene.Slides[0];
        PptxSceneNode sceneNode = sceneSlide.SlideNodes[0];
        TestAssert.True(sceneNode.Shape?.HasUnsupportedTransparency == true, "Expected unsupported alpha provenance to be owned by the scene shape.");

        PptxSceneNodeSnapshot snapshot = PptxRenderer.InspectScene(document, package).Slides[0].SlideNodes[0];
        TestAssert.True(snapshot.ShapeHasUnsupportedTransparency, "Expected private-safe scene inspection to expose unsupported transparency state.");

        var diagnostics = new List<OoxPdfDiagnostic>();
        XDocument slideXmlWithoutAlpha = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                   xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld><p:spTree><p:sp><p:spPr/></p:sp></p:spTree></p:cSld>
            </p:sld>
            """);
        System.Reflection.MethodInfo emitDiagnostics = typeof(PptxRenderer).GetMethod(
            "EmitUnsupportedFeatureDiagnostics",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected unsupported feature diagnostic emitter.");
        Action<OoxPdfDiagnostic> sink = diagnostics.Add;
        emitDiagnostics.Invoke(null, [sceneSlide, slideXmlWithoutAlpha, "/ppt/slides/slide1.xml", 1, sink, new HashSet<string>(StringComparer.Ordinal)]);

        TestAssert.Contains("PPTX_UNSUPPORTED_TRANSPARENCY", string.Join("|", diagnostics.Select(d => d.Id)));
    }

    public static void PptxUnsupportedTransparencyDiagnosticsUseInheritedSceneShapeState()
    {
        string contentTypes = PptxTests.BasicContentTypes().Replace(
            "</Types>",
            """
              <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
              <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
            </Types>
            """);
        string presentationRelationships = PptxTests.PresentationRelationship().Replace(
            "</Relationships>",
            """
              <Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
            </Relationships>
            """);
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = contentTypes,
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = presentationRelationships,
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/_rels/slide1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
                </Relationships>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree/></p:cSld>
                </p:sld>
                """,
            ["ppt/slideLayouts/_rels/slideLayout1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
                </Relationships>
                """,
            ["ppt/slideLayouts/slideLayout1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:sp>
                      <p:nvSpPr><p:cNvPr id="2" name="LayoutTransparentEffect"/><p:nvPr/></p:nvSpPr>
                      <p:spPr>
                        <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                        <a:effectLst>
                          <a:reflection><a:srgbClr val="123456"><a:alpha val="50000"/></a:srgbClr></a:reflection>
                        </a:effectLst>
                      </p:spPr>
                    </p:sp>
                  </p:spTree></p:cSld>
                </p:sldLayout>
                """,
            ["ppt/slideMasters/slideMaster1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree/></p:cSld>
                </p:sldMaster>
                """
        });
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        PptxSceneSlide sceneSlide = scene.Slides[0];
        TestAssert.True(
            sceneSlide.LayoutNodes.Any(node => node.Shape?.HasUnsupportedTransparency == true),
            "Expected inherited layout transparency provenance to be owned by layout scene nodes.");

        var diagnostics = new List<OoxPdfDiagnostic>();
        XDocument slideXmlWithoutAlpha = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                   xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld><p:spTree/></p:cSld>
            </p:sld>
            """);
        System.Reflection.MethodInfo emitDiagnostics = typeof(PptxRenderer).GetMethod(
            "EmitUnsupportedFeatureDiagnostics",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected unsupported feature diagnostic emitter.");
        Action<OoxPdfDiagnostic> sink = diagnostics.Add;
        emitDiagnostics.Invoke(null, [sceneSlide, slideXmlWithoutAlpha, "/ppt/slides/slide1.xml", 1, sink, new HashSet<string>(StringComparer.Ordinal)]);

        TestAssert.Contains("PPTX_UNSUPPORTED_TRANSPARENCY", string.Join("|", diagnostics.Select(d => d.Id)));
    }

    public static void PptxUnsupportedGradientDiagnosticsUseSceneShapeGradient()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:sp><p:spPr><a:gradFill/></p:spPr></p:sp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
            PptxSceneShape shape = new PptxSceneBuilder().Build(document, package, CancellationToken.None).Slides[0].SlideNodes[0].Shape
                ?? throw new InvalidOperationException("Expected shape scene node.");
            PptxSceneNodeSnapshot snapshot = PptxRenderer.InspectScene(document, package).Slides[0].SlideNodes[0];

            TestAssert.True(shape.GradientFill.HasGradientSource, "Expected unsupported shape gradient provenance to remain scene-owned.");
            TestAssert.True(shape.GradientFill.HasUnsupportedGradient, "Expected unsupported shape gradient state to remain scene-owned.");
            TestAssert.True(snapshot.ShapeHasGradientSource, "Expected private-safe scene inspection to expose shape gradient provenance.");
            TestAssert.True(snapshot.ShapeHasUnsupportedGradient, "Expected private-safe scene inspection to expose unsupported shape gradient provenance.");
        }

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(diagnostics.Any(d => d.Id == "PPTX_UNSUPPORTED_GRADIENT_FILL"), "Unsupported shape gradients should be diagnostic-covered from scene-owned gradient provenance.");
    }

    public static void PptxUnsupportedPatternDiagnosticsUseSceneShapePattern()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:sp><p:spPr><a:pattFill prst="pct25"/></p:spPr></p:sp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
            PptxSceneShape shape = new PptxSceneBuilder().Build(document, package, CancellationToken.None).Slides[0].SlideNodes[0].Shape
                ?? throw new InvalidOperationException("Expected shape scene node.");
            PptxSceneNodeSnapshot snapshot = PptxRenderer.InspectScene(document, package).Slides[0].SlideNodes[0];

            TestAssert.True(shape.PatternFill.HasPatternSource, "Expected unsupported shape pattern provenance to remain scene-owned.");
            TestAssert.True(shape.PatternFill.HasUnsupportedPattern, "Expected unsupported shape pattern state to remain scene-owned.");
            TestAssert.True(snapshot.ShapeHasPatternSource, "Expected private-safe scene inspection to expose shape pattern provenance.");
            TestAssert.True(snapshot.ShapeHasUnsupportedPattern, "Expected private-safe scene inspection to expose unsupported shape pattern provenance.");
        }

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(diagnostics.Any(d => d.Id == "PPTX_UNSUPPORTED_PATTERN_FILL"), "Unsupported shape patterns should be diagnostic-covered from scene-owned pattern provenance.");
    }

    public static void PptxUnsupportedTextDiagnosticsUseSceneTextBodyProperties()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:sp>
                      <p:spPr><a:prstGeom prst="rect"/></p:spPr>
                      <p:txBody>
                        <a:bodyPr vert="futureVert" vertOverflow="ellipsis"/>
                        <a:p><a:r><a:t>Text</a:t></a:r></a:p>
                      </p:txBody>
                    </p:sp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
            PptxSceneTextBody textBody = new PptxSceneBuilder().Build(document, package, CancellationToken.None).Slides[0].SlideNodes[0].TextBody
                ?? throw new InvalidOperationException("Expected shape text body scene node.");
            PptxSceneNodeSnapshot snapshot = PptxRenderer.InspectScene(document, package).Slides[0].SlideNodes[0];

            TestAssert.True(textBody.HasUnsupportedTextOrientation, "Expected unsupported shape text orientation state to remain scene-owned.");
            TestAssert.True(textBody.HasUnsupportedVerticalOverflow, "Expected unsupported shape text overflow state to remain scene-owned.");
            TestAssert.True(snapshot.TextHasUnsupportedOrientation, "Expected private-safe scene inspection to expose unsupported shape text orientation.");
            TestAssert.True(snapshot.TextHasUnsupportedVerticalOverflow, "Expected private-safe scene inspection to expose unsupported shape text overflow.");
        }

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(diagnostics.Any(d => d.Id == "PPTX_UNSUPPORTED_TEXT_ORIENTATION"), "Unsupported shape text orientation should be diagnostic-covered from scene-owned text-body provenance.");
        TestAssert.True(diagnostics.Any(d => d.Id == "PPTX_UNSUPPORTED_TEXT_OVERFLOW"), "Unsupported shape text overflow should be diagnostic-covered from scene-owned text-body provenance.");
    }

    public static void PptxUnsupportedTableTextDiagnosticsUseSceneCellBodyProperties()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:nvGraphicFramePr>
                        <p:cNvPr id="2" name="Table"/>
                        <p:cNvGraphicFramePr/>
                        <p:nvPr/>
                      </p:nvGraphicFramePr>
                      <p:xfrm>
                        <a:off x="0" y="0"/>
                        <a:ext cx="914400" cy="457200"/>
                      </p:xfrm>
                      <a:graphic>
                        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                          <a:tbl>
                            <a:tblGrid><a:gridCol w="914400"/></a:tblGrid>
                            <a:tr h="457200">
                              <a:tc>
                                <a:txBody>
                                  <a:bodyPr vert="futureVert" vertOverflow="ellipsis"/>
                                  <a:p><a:r><a:t>Cell</a:t></a:r></a:p>
                                </a:txBody>
                                <a:tcPr/>
                              </a:tc>
                            </a:tr>
                          </a:tbl>
                        </a:graphicData>
                      </a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
            PptxSceneTableCell cell = new PptxSceneBuilder().Build(document, package, CancellationToken.None).Slides[0].SlideNodes[0].Table?.Rows[0].Cells[0]
                ?? throw new InvalidOperationException("Expected table cell scene node.");

            TestAssert.True(cell.HasUnsupportedTextOrientation, "Expected unsupported table-cell text orientation state to remain scene-owned.");
            TestAssert.True(cell.HasUnsupportedVerticalOverflow, "Expected unsupported table-cell text overflow state to remain scene-owned.");
        }

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(diagnostics.Any(d => d.Id == "PPTX_UNSUPPORTED_TEXT_ORIENTATION"), "Unsupported table-cell text orientation should be diagnostic-covered from scene-owned table cell provenance.");
        TestAssert.True(diagnostics.Any(d => d.Id == "PPTX_UNSUPPORTED_TEXT_OVERFLOW"), "Unsupported table-cell text overflow should be diagnostic-covered from scene-owned table cell provenance.");
    }

    public static void PptxUnsupportedCustomGeometryDiagnosticsUseSceneUnsupportedFlag()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:sp><p:spPr>
                      <a:custGeom>
                        <a:pathLst>
                          <a:path><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="21600" y="21600"/></a:lnTo></a:path>
                          <a:path><a:unknownPathCommand/></a:path>
                        </a:pathLst>
                      </a:custGeom>
                    </p:spPr></p:sp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
            PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
            PptxSceneShape shape = scene.Slides[0].SlideNodes[0].Shape ?? throw new InvalidOperationException("Expected shape scene node.");
            TestAssert.True(shape.CustomGeometry.HasGeometry, "Expected the supported custom path to remain parsed.");
            TestAssert.True(shape.CustomGeometry.HasUnsupportedGeometry, "Expected unsupported custom path provenance to remain scene-owned.");
        }

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(diagnostics.Any(d => d.Id == "PPTX_UNSUPPORTED_CUSTOM_GEOMETRY"), "Unsupported custom geometry should be diagnostic-covered from scene-owned geometry provenance.");
    }

    public static void PptxChartNumberFormatFallbackIsDiagnosedOncePerChart()
    {
        IReadOnlyList<OoxPdfDiagnostic> diagnostics = PptxTests.ConvertSingleChartAndCollectDiagnostics("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:plotArea><c:lineChart>
                <c:grouping val="standard"/>
                <c:dLbls><c:numFmt formatCode="[Blue]0%" sourceLinked="0"/><c:showVal val="1"/></c:dLbls>
                <c:ser>
                  <c:cat><c:strLit><c:ptCount val="2"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                  <c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numLit></c:val>
                </c:ser>
                <c:axId val="10"/><c:axId val="20"/>
              </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="20"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:numFmt formatCode="[Red]0.0" sourceLinked="0"/><c:tickLblPos val="nextTo"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart>
            </c:chartSpace>
            """);

        OoxPdfDiagnostic[] numberFormat = diagnostics.Where(d => d.Id == "PPTX_UNSUPPORTED_CHART_NUMBER_FORMAT").ToArray();
        TestAssert.Equal(1, numberFormat.Length);
        TestAssert.Contains("colors", numberFormat[0].Message);
        TestAssert.True(!numberFormat[0].Message.Contains("[Red]") && !numberFormat[0].Message.Contains("[Blue]"), "Format codes must not leak into diagnostics.");
    }

    public static void PptxChartNumberFormatWithoutFallbackGapsStaysQuiet()
    {
        IReadOnlyList<OoxPdfDiagnostic> diagnostics = PptxTests.ConvertSingleChartAndCollectDiagnostics("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:plotArea><c:lineChart>
                <c:grouping val="standard"/>
                <c:ser>
                  <c:cat><c:strLit><c:ptCount val="2"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                  <c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numLit></c:val>
                </c:ser>
                <c:axId val="10"/><c:axId val="20"/>
              </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="nextTo"/><c:crossAx val="20"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:numFmt formatCode="0.0" sourceLinked="0"/><c:tickLblPos val="nextTo"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart>
            </c:chartSpace>
            """);

        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_CHART_NUMBER_FORMAT"), "Supported number formats should not warn.");
    }

    public static void PptxChartTrendlineWarnsOnce()
    {
        IReadOnlyList<OoxPdfDiagnostic> diagnostics = PptxTests.ConvertSingleChartAndCollectDiagnostics("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:plotArea><c:lineChart>
                <c:grouping val="standard"/>
                <c:ser>
                  <c:cat><c:strLit><c:ptCount val="2"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                  <c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numLit></c:val>
                  <c:trendline><c:trendlineType val="linear"/></c:trendline>
                </c:ser>
                <c:ser>
                  <c:cat><c:strLit><c:ptCount val="2"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                  <c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>30</c:v></c:pt><c:pt idx="1"><c:v>40</c:v></c:pt></c:numLit></c:val>
                  <c:trendline><c:trendlineType val="polynomial"/><c:order val="2"/></c:trendline>
                </c:ser>
                <c:axId val="10"/><c:axId val="20"/>
              </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="nextTo"/><c:crossAx val="20"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="nextTo"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart>
            </c:chartSpace>
            """);

        TestAssert.Equal(1, diagnostics.Count(d => d.Id == "PPTX_UNSUPPORTED_CHART_TRENDLINE"));
    }

    public static void PptxChartWithoutTrendlineStaysQuiet()
    {
        IReadOnlyList<OoxPdfDiagnostic> diagnostics = PptxTests.ConvertSingleChartAndCollectDiagnostics("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:plotArea><c:lineChart>
                <c:grouping val="standard"/>
                <c:ser>
                  <c:cat><c:strLit><c:ptCount val="2"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                  <c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numLit></c:val>
                </c:ser>
                <c:axId val="10"/><c:axId val="20"/>
              </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="nextTo"/><c:crossAx val="20"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="nextTo"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart>
            </c:chartSpace>
            """);

        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_CHART_TRENDLINE"), "Charts without trendlines should not warn.");
    }
}
