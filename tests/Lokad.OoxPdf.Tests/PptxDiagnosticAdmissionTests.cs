using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxDiagnosticAdmissionTests
{
    public static void PptxDiagnosticAdmissionSuppressesUnusedMasterFeatures()
    {
        foreach (var feature in new[] {
            ("<a:prstGeom prst=\"wedgeRoundRectCallout\"/>", "PPTX_UNSUPPORTED_CALLOUT"),
            ("<a:prstGeom prst=\"rect\"/><a:effectDag/>", "PPTX_UNSUPPORTED_EFFECT"),
            ("<a:custGeom><a:pathLst><a:path><a:unknownPathCommand/></a:path></a:pathLst></a:custGeom>", "PPTX_UNSUPPORTED_CUSTOM_GEOMETRY") })
        {
            WithParts(master: Shape(feature.Item1), shown: false, check: path =>
            {
                var diagnostics = Convert(path);
                TestAssert.Equal(0, diagnostics.Count);
            });
            WithParts(master: Shape(feature.Item1), shown: true, check: path =>
            {
                var diagnostic = Convert(path).Single(item => item.Id == feature.Item2);
                TestAssert.Equal("/ppt/slideMasters/slideMaster1.xml", diagnostic.PartName);
                TestAssert.Equal(1, diagnostic.SlideIndex);
            });
        }
    }

    public static void PptxDiagnosticAdmissionMatchesGroupedMasterAndLayoutPlaceholders()
    {
        foreach (bool master in new[] { false, true })
        foreach (bool placeholder in new[] { false, true })
        {
            string shape = Shape("<a:prstGeom prst=\"wedgeRoundRectCallout\"/>", placeholder);
            string group = "<p:grpSp><p:nvGrpSpPr><p:cNvPr id=\"3\" name=\"Group\"/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>" + shape + "</p:grpSp>";
            WithParts(master: master ? group : "", layout: master ? "" : group, check: path =>
            {
                var diagnostics = Convert(path);
                if (placeholder) TestAssert.Equal(0, diagnostics.Count);
                else TestAssert.Equal(master ? "/ppt/slideMasters/slideMaster1.xml" : "/ppt/slideLayouts/slideLayout1.xml",
                    diagnostics.Single(item => item.Id == "PPTX_UNSUPPORTED_CALLOUT").PartName);
            });
        }
        // A slide's own placeholder is admitted, including inside a group.
        WithParts(slide: Shape("<a:prstGeom prst=\"wedgeRoundRectCallout\"/>", true), check: path =>
            TestAssert.Equal("/ppt/slides/slide1.xml", Convert(path).Single(item => item.Id == "PPTX_UNSUPPORTED_CALLOUT").PartName));
    }

    public static void PptxDiagnosticAdmissionKeepsMasterMarkupChecksAndCancellation()
    {
        WithParts(shown: false, master: "<x:unsupported xmlns:x=\"urn:public-review-unknown\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:MustUnderstand=\"x\"/>", check: path =>
        {
            var diagnostics = Convert(path);
            TestAssert.True(diagnostics.Any(item => item.Id == "OOXML_MUST_UNDERSTAND" && item.PartName == "/ppt/slideMasters/slideMaster1.xml"),
                "Master XML still supplies consumed inheritance and markup-compatibility warnings.");
        });
        WithParts(master: Shape("<a:prstGeom prst=\"wedgeRoundRectCallout\"/>"), check: path =>
        {
            using var cancellation = new CancellationTokenSource();
            using var input = File.OpenRead(path);
            using var output = new MemoryStream();
            bool observed = false;
            TestAssert.Throws<OperationCanceledException>(() => OoxPdfConverter.Convert(input, output,
                new OoxPdfOptions { InputKind = OoxPdfInputKind.Pptx, DiagnosticSink = diagnostic => {
                    if (diagnostic.Id == "PPTX_UNSUPPORTED_CALLOUT") { observed = true; cancellation.Cancel(); }
                } }, cancellation.Token));
            TestAssert.True(observed, "Cancellation follows an admitted source-owned diagnostic.");
        });
    }

    public static void PptxLiveStylesHonorNoFillAndRetainLayoutSource()
    {
        string text = "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz=\"1800\"><a:noFill/><a:latin typeface=\"Arial\"/></a:rPr><a:t>A</a:t></a:r></a:p></p:txBody>";
        WithParts(slide: Shape("<a:prstGeom prst=\"rect\"/>").Replace("</p:sp>", text + "</p:sp>"), check: path =>
        {
            using var input = File.OpenRead(path);
            OoxPackage package = OoxPackage.Open(input, default);
            var document = new PptxReader().Read(package, default);
            var scene = new PptxSceneBuilder().Build(document, package, default);
            var node = scene.Slides[0].SlideNodes.Single();
            string before = node.Source.ToString(SaveOptions.DisableFormatting);
            var spans = PptxSceneTextAgreementTests.ReadXmlSpans(node, document, scene);
            var run = spans.First().SourceRun!;
            TestAssert.Equal(0d, run.Style.Alpha);
            TestAssert.Equal(PptxRenderer.PptxRunTextColorSource.RunNoFill, run.Style.ColorSource);
            TestAssert.Equal("A", run.Text);
            TestAssert.True(spans.Sum(span => span.Run.Width) > 0, "No-fill text still occupies its declared layout advance.");
            var snapshot = PptxRenderer.InspectTextFrameModels(document, package, 0).Single().Paragraphs.Single().Runs.Single();
            TestAssert.Equal("RunNoFill", snapshot.ColorSource);
            var page = new PptxRenderer(new PortableRenderingTests.PinnedResolver()).RenderPages(document, package, null, default).Single();
            TestAssert.DoesNotContain("BT" + Environment.NewLine, page.Content);
            TestAssert.Equal(before, node.Source.ToString(SaveOptions.DisableFormatting));
        });
    }

    private static List<OoxPdfDiagnostic> Convert(string path)
    {
        var diagnostics = new List<OoxPdfDiagnostic>();
        using var input = File.OpenRead(path);
        using var output = new MemoryStream();
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { InputKind=OoxPdfInputKind.Pptx, DiagnosticSink=diagnostics.Add });
        return diagnostics;
    }

    private static string Shape(string geometry, bool placeholder = false) =>
        "<p:sp><p:nvSpPr><p:cNvPr id=\"2\" name=\"Public control\"/><p:cNvSpPr/><p:nvPr>" +
        (placeholder ? "<p:ph type=\"body\"/>" : "") + "</p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"914400\" y=\"914400\"/><a:ext cx=\"1828800\" cy=\"914400\"/></a:xfrm>" + geometry + "</p:spPr></p:sp>";

    private static void WithParts(Action<string> check, string slide = "", string layout = "", string master = "", bool shown = true)
    {
        string Story(string content) => "<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" showMasterSp=\"" + (shown ? "1" : "0") + "\"><p:cSld><p:spTree>" + content + "</p:spTree></p:cSld></p:sld>";
        string Relationship(string kind, string target) => "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/" + kind + "\" Target=\"" + target + "\"/></Relationships>";
        string types = PptxTests.BasicContentTypes().Replace("</Types>", "<Override PartName=\"/ppt/slideLayouts/slideLayout1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml\"/><Override PartName=\"/ppt/slideMasters/slideMaster1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml\"/></Types>");
        string path = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string,string> {
            ["[Content_Types].xml"]=types,
            ["_rels/.rels"]=PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"]=PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"]=PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"]=Story(slide),
            ["ppt/slides/_rels/slide1.xml.rels"]=Relationship("slideLayout","../slideLayouts/slideLayout1.xml"),
            ["ppt/slideLayouts/slideLayout1.xml"]=Story(layout),
            ["ppt/slideLayouts/_rels/slideLayout1.xml.rels"]=Relationship("slideMaster","../slideMasters/slideMaster1.xml"),
            ["ppt/slideMasters/slideMaster1.xml"]=Story(master)
        });
        try { check(path); } finally { File.Delete(path); }
    }
}
