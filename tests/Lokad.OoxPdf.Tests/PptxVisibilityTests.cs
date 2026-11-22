using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxVisibilityTests
{
    public static void PptxHiddenObjectsLeaveNoPaintResourcesOrLinks()
    {
        // A hidden group must exclude descendants before preflight, just as a
        // hidden leaf excludes its own text, picture or table resources.
        foreach (string kind in new[] { "shape", "picture", "table", "group" })
        foreach (string hidden in new[] { "", "0", "false", "1", "true", " true ", " 1 " })
        {
            bool visible = hidden is not ("1" or "true" or " true " or " 1 ");
            string flag = hidden.Length == 0 ? "" : " hidden=\"" + hidden + "\"";
            const string link = "<a:hlinkClick r:id=\"rIdLink\"/>";
            const string transform = "<a:xfrm><a:off x=\"914400\" y=\"914400\"/><a:ext cx=\"2743200\" cy=\"914400\"/></a:xfrm>";
            const string text = "<a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz=\"1200\"><a:latin typeface=\"Arial\"/></a:rPr><a:t>Visible text</a:t></a:r></a:p>";
            string shape = "<p:sp><p:nvSpPr><p:cNvPr id=\"2\" name=\"Text\"" + (kind == "shape" ? flag : "") + ">" + link + "</p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>" + transform + "<a:prstGeom prst=\"rect\"/><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill></p:spPr><p:txBody>" + text + "</p:txBody></p:sp>";
            string node = kind switch
            {
                "picture" => "<p:pic><p:nvPicPr><p:cNvPr id=\"2\" name=\"Picture\"" + flag + ">" + link + "</p:cNvPr><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed=\"rIdImage\"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>" + transform + "<a:prstGeom prst=\"rect\"/></p:spPr></p:pic>",
                "table" => "<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id=\"2\" name=\"Table\"" + flag + ">" + link + "</p:cNvPr><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x=\"914400\" y=\"914400\"/><a:ext cx=\"2743200\" cy=\"914400\"/></p:xfrm><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/table\"><a:tbl><a:tblPr/><a:tblGrid><a:gridCol w=\"2743200\"/></a:tblGrid><a:tr h=\"914400\"><a:tc><a:txBody>" + text + "</a:txBody><a:tcPr/></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>",
                "group" => "<p:grpSp><p:nvGrpSpPr><p:cNvPr id=\"3\" name=\"Group\"" + flag + "/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>" + shape + "</p:grpSp>",
                _ => shape
            };
            using var input = TestFixtures.CreateZipPackage(new Dictionary<string, byte[]>
            {
                ["[Content_Types].xml"] = TestFixtures.Utf8(PptxTests.BasicContentTypes().Replace("</Types>", "<Default Extension=\"png\" ContentType=\"image/png\"/></Types>", StringComparison.Ordinal)),
                ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
                ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
                ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
                ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><p:cSld><p:spTree>" + node + "</p:spTree></p:cSld></p:sld>"),
                ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rIdLink\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.invalid/visible\" TargetMode=\"External\"/><Relationship Id=\"rIdImage\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/image.png\"/></Relationships>"),
                ["ppt/media/image.png"] = TestFixtures.CreateTinyStoredRgbPng()
            });
            using OoxPackage package = OoxPackage.Open(input, default);
            PptxDocument document = new PptxReader().Read(package, default);
            PdfPage page = new PptxRenderer(new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
            TestAssert.Equal(visible ? 1 : 0, page.Annotations.Count);
            TestAssert.Equal(visible && kind == "picture" ? 1 : 0, page.Images.Count);
            TestAssert.Equal(visible && kind != "picture", page.Fonts.Count > 0);
            TestAssert.Equal(visible && kind is "shape" or "group", page.Content.Contains("1 0 0 rg", StringComparison.Ordinal));
            if (!visible)
            {
                TestAssert.Equal(0, PptxRenderer.InspectTextGlyphRuns(document, package, 0).Count);
                if (kind is "shape" or "group")
                {
                    TestAssert.Equal(0, PptxRenderer.InspectTextFrameModels(document, package, 0).Count);
                    TestAssert.Equal(0, PptxRenderer.InspectTextLayout(document, package, 0).Frames.Count);
                    TestAssert.Equal(0, PptxRenderer.InspectTextFlow(document, package, 0).Frames.Count);
                }
            }
        }
    }

    public static void PptxGlyphEmissionInspectionUsesSceneVisibilityAndMasterAdmission()
    {
        foreach (bool showMaster in new[] { false, true })
        foreach (bool hidden in new[] { false, true })
        {
            string input = WriteMasterShapePackage(showMaster ? "1" : "0", withText: true);
            if (hidden)
            {
                using var zip = System.IO.Compression.ZipFile.Open(input, System.IO.Compression.ZipArchiveMode.Update);
                var entry = zip.GetEntry("ppt/slides/slide1.xml")!;
                string xml;
                using (var reader = new StreamReader(entry.Open())) { xml = reader.ReadToEnd(); }
                entry.Delete();
                using var writer = new StreamWriter(zip.CreateEntry("ppt/slides/slide1.xml").Open());
                writer.Write(xml.Replace("<p:sp>", "<p:sp><p:nvSpPr><p:cNvPr id=\"9\" name=\"Hidden\" hidden=\"1\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>", StringComparison.Ordinal));
            }
            try
            {
                using OoxPackage package = OoxPackage.Open(input, default);
                PptxDocument document = new PptxReader().Read(package, default);
                var fonts = new PresentationFontResolver(new PortableRenderingTests.PinnedResolver());
                var models = PptxRenderer.InspectTextFrameModels(document, package, 0, inspectionFonts: fonts);
                var layout = PptxRenderer.InspectTextLayout(document, package, 0, inspectionFonts: fonts);
                var flow = PptxRenderer.InspectTextFlow(document, package, 0, inspectionFonts: fonts);
                var glyphs = PptxRenderer.InspectTextGlyphRuns(document, package, 0, inspectionFonts: fonts);
                int admittedFrames = 1 + (showMaster ? 1 : 0) + (hidden ? 0 : 1);
                TestAssert.Equal(admittedFrames, models.Count);
                TestAssert.Equal(admittedFrames, layout.Frames.Count);
                TestAssert.Equal(admittedFrames, flow.Frames.Count);
                PdfPage painted = new PptxRenderer(new PortableRenderingTests.PinnedResolver())
                    .RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
                foreach (string text in new[]
                {
                    string.Concat(models.SelectMany(frame => frame.Paragraphs).SelectMany(p => p.Runs).Select(run => run.Text)),
                    string.Concat(layout.Frames.SelectMany(frame => frame.Paragraphs).SelectMany(p => p.Lines).SelectMany(line => line.Spans).Select(span => span.Text)),
                    string.Concat(flow.Frames.SelectMany(frame => frame.Paragraphs).SelectMany(p => p.Runs).Select(run => run.SourceText)),
                    string.Concat(glyphs.Select(run => run.Text)),
                    DocxPageFootnoteNumberingTests.DecodeText(painted)
                })
                {
                    TestAssert.Equal(showMaster, text.Contains("Master text", StringComparison.Ordinal));
                    TestAssert.Equal(!hidden, text.Contains("Slide text", StringComparison.Ordinal));
                    TestAssert.Contains("Layout text", text);
                }
            }
            finally { File.Delete(input); }
        }
    }

    public static void PptxTableInspectionUsesMasterAdmission()
    {
        string fixture = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "pptx-table.pptx"));
        XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        XElement table;
        using (var source = ZipFile.OpenRead(fixture))
        using (var reader = new StreamReader(source.GetEntry("ppt/slides/slide1.xml")!.Open()))
            table = new XElement(XDocument.Load(reader).Descendants(p + "graphicFrame").First());
        table.Descendants(a + "t").First().Value = "Master cell";
        foreach (bool showMaster in new[] { false, true })
        {
            string input = WriteMasterShapePackage(showMaster ? "1" : "0", withText: true);
            try
            {
                using (var zip = ZipFile.Open(input, ZipArchiveMode.Update))
                {
                    var entry = zip.GetEntry("ppt/slideMasters/slideMaster1.xml")!;
                    XDocument master;
                    using (var reader = new StreamReader(entry.Open())) master = XDocument.Load(reader);
                    master.Descendants(p + "sp").Single().ReplaceWith(new XElement(table));
                    entry.Delete();
                    using var writer = new StreamWriter(zip.CreateEntry("ppt/slideMasters/slideMaster1.xml").Open());
                    master.Save(writer);
                }
                using OoxPackage package = OoxPackage.Open(input, default);
                PptxDocument document = new PptxReader().Read(package, default);
                var fonts = new PresentationFontResolver(new PortableRenderingTests.PinnedResolver());
                var models = PptxRenderer.InspectTableTextFrameModels(document, package, 0, inspectionFonts: fonts);
                TestAssert.Equal(showMaster, models.Count > 0);
                string modelText = string.Concat(models.SelectMany(frame => frame.Paragraphs).SelectMany(paragraph => paragraph.Runs).Select(run => run.Text));
                string glyphText = string.Concat(PptxRenderer.InspectTextGlyphRuns(document, package, 0, inspectionFonts: fonts).Select(run => run.Text));
                PdfPage painted = new PptxRenderer(new PortableRenderingTests.PinnedResolver())
                    .RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
                foreach (string text in new[] { modelText, glyphText, DocxPageFootnoteNumberingTests.DecodeText(painted) })
                    TestAssert.Equal(showMaster, text.Contains("Master cell", StringComparison.Ordinal));
            }
            finally { File.Delete(input); }
        }
    }

    public static void PptxHiddenObjectsAreExcludedFromEverySceneSource()
    {
        // The three source layers share node construction.
        string input = WriteMasterShapePackage(showMasterSp: null);
        using (var archive = System.IO.Compression.ZipFile.Open(input, System.IO.Compression.ZipArchiveMode.Update))
        {
            foreach (string part in new[] { "ppt/slides/slide1.xml", "ppt/slideLayouts/slideLayout1.xml", "ppt/slideMasters/slideMaster1.xml" })
            {
                var entry = archive.GetEntry(part)!;
                string xml;
                using (var reader = new StreamReader(entry.Open())) { xml = reader.ReadToEnd(); }
                xml = xml.Replace("<p:sp>", "<p:sp><p:nvSpPr><p:cNvPr id=\"9\" name=\"Hidden\" hidden=\"true\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>", StringComparison.Ordinal);
                entry.Delete();
                using var writer = new StreamWriter(archive.CreateEntry(part).Open());
                writer.Write(xml);
            }
        }
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, default);
        PptxDocument document = new PptxReader().Read(package, default);
        PptxSceneSlide scene = new PptxSceneBuilder().Build(document, package, default).Slides.Single();
        TestAssert.Equal(0, scene.SlideNodes.Count);
        TestAssert.Equal(0, scene.LayoutNodes.Count);
        TestAssert.Equal(0, scene.MasterNodes.Count);
    }

    public static void PptxHiddenMasterShapesStayUnpainted() => AssertMasterShapePaint("0", expectsMaster: false);
    public static void PptxShownMasterShapesPaintByDefault() => AssertMasterShapePaint(null, expectsMaster: true);
    public static void PptxExplicitShownMasterShapesPaint() => AssertMasterShapePaint("1", expectsMaster: true);

    private static void AssertMasterShapePaint(string? showMasterSp, bool expectsMaster)
    {
        string input = WriteMasterShapePackage(showMasterSp);
        string output = Path.ChangeExtension(input, ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Equal(expectsMaster, pdf.Contains("1 0 0 rg", StringComparison.Ordinal));
        TestAssert.Contains("0 1 0 rg", pdf);
        TestAssert.Contains("0 0 1 rg", pdf);
    }

    public static void PptxHiddenSlidesAreSkipped()
    {
        string input = WriteHiddenSlidePackage("0");
        string output = Path.ChangeExtension(input, ".pdf");
        try
        {
            OoxPdfConverter.Convert(input, output);
            AssertVisibleSlidePaint(File.ReadAllText(output, Encoding.ASCII));
        }
        finally
        {
            File.Delete(input);
            File.Delete(output);
        }
    }

    public static void PptxPaddedBooleanValuesPreserveExcludedSlides()
    {
        foreach (string show in new[] { "false", " false ", " 0 " })
        {
            string input = WriteHiddenSlidePackage(show);
            try
            {
                using var source = File.OpenRead(input);
                using var output = new MemoryStream();
                OoxPdfConverter.Convert(source, output, new OoxPdfOptions { InputKind = OoxPdfInputKind.Pptx });
                AssertVisibleSlidePaint(Encoding.ASCII.GetString(output.ToArray()));
            }
            finally { File.Delete(input); }
        }
    }

    private static void AssertVisibleSlidePaint(string pdf)
    {
        TestAssert.Equal(1, PptxTests.CountOccurrences(pdf, "/Type /Page "));
        TestAssert.Contains("0 0 1 rg", pdf);
        TestAssert.DoesNotContain("1 0 0 rg", pdf);
    }

    public static void PptxLinksToHiddenSlidesAreDiagnosed()
    {
        string input = PptxHyperlinksTests.WriteInternalSlideLinkPackage(PptxHyperlinksTests.ShapeWithInternalLink, hideTarget: true);
        var (pages, diagnostics) = PptxHyperlinksTests.RenderSlidePages(input);

        PdfPage page = pages.Single();
        TestAssert.Equal(0, page.Annotations.Count);
        TestAssert.True(Regex.IsMatch(page.Content, @"0 0 1 rg\s+72 396 144 72 re f\*"),
            "The linked shape keeps its blue fill at the authored bounds.");
        TestAssert.DoesNotContain("1 0 0 rg", page.Content);
        OoxPdfDiagnostic warning = diagnostics.Single(d => d.Id == "PPTX_UNSUPPORTED_HYPERLINK");
        TestAssert.Equal(OoxPdfSeverity.Warning, warning.Severity);
        TestAssert.Equal("/ppt/slides/slide1.xml", warning.PartName);
        TestAssert.Equal<int?>(1, warning.SlideIndex);
    }

    private static string WriteHiddenSlidePackage(string show)
    {
        return TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                  <Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
                </Relationships>
                """,
            ["ppt/presentation.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:sldSz cx="9144000" cy="6858000"/>
                  <p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst>
                </p:presentation>
                """,
            ["ppt/slides/slide1.xml"] = PptxTests.InheritedShapePart("0000FF", 914400, 914400),
            ["ppt/slides/slide2.xml"] = $$"""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" show="{{show}}">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr>
                      <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="rect"/>
                      <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
                    </p:spPr>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
    }

    private static string WriteMasterShapePackage(string? showMasterSp, bool withText = false)
    {
        string slide = PptxTests.InheritedShapePart("0000FF", 2743200, 2743200);
        if (showMasterSp is not null)
        {
            slide = slide.Replace("<p:sld ", "<p:sld showMasterSp=\"" + showMasterSp + "\" ");
        }

        return TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
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
            ["ppt/slideMasters/slideMaster1.xml"] = AddText(PptxTests.InheritedShapePart("FF0000", 914400, 914400), "Master text"),
            ["ppt/slideLayouts/slideLayout1.xml"] = AddText(PptxTests.InheritedShapePart("00FF00", 1828800, 1828800), "Layout text"),
            ["ppt/slides/slide1.xml"] = AddText(slide, "Slide text")
        });

        string AddText(string xml, string text) => withText
            ? xml.Replace("</p:spPr>", "</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz=\"1200\"><a:latin typeface=\"Arial\"/></a:rPr><a:t>" + text + "</a:t></a:r></a:p></p:txBody>", StringComparison.Ordinal)
            : xml;
    }
}

