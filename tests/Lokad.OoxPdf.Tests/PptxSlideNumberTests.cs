using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxSlideNumberTests
{
    public static void PptxSlideNumbersPreserveSourceOrderAndNumberingOrigin()
    {
        foreach (long first in new long[] { 0, 1, 10, uint.MaxValue })
        foreach (bool hidden in new[] { false, true })
            WithDeck(first.ToString(CultureInfo.InvariantCulture), hidden, (document, package) =>
            {
                TestAssert.Equal(first, document.FirstSlideNumber);
                TestAssert.Equal(hidden ? 2 : 3, document.Slides.Count);
                for (int index = 0; index < document.Slides.Count; index++)
                {
                    int sourceIndex = hidden && index == 1 ? 2 : index;
                    TestAssert.Equal(index, document.Slides[index].Index);
                    TestAssert.Equal(sourceIndex, document.Slides[index].SourceIndex ?? -1);
                    string expected = (first + sourceIndex).ToString(CultureInfo.InvariantCulture);
                    var fields = PptxRenderer.InspectTextFrameModels(document, package, index)
                        .SelectMany(frame => frame.Paragraphs).SelectMany(paragraph => paragraph.Runs)
                        .Where(run => run.Kind == "Field").ToArray();
                    TestAssert.Equal(3, fields.Length); // master, layout and slide
                    TestAssert.True(fields.All(run => run.Text == expected), "Every shape source must use the displayed slide number.");
                    var table = PptxRenderer.InspectTableTextFrameModels(document, package, index).Single();
                    TestAssert.Equal(expected, table.Paragraphs.Single().Runs.Single().Text);
                }
            });
    }

    public static void PptxSlideNumbersKeepOutputDiagnosticsAndNavigation()
    {
        WithDeck("10", true, (document, package) =>
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = new PptxRenderer(new PortableRenderingTests.PinnedResolver())
                .RenderPagesAsync(document, package, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(2, pages.Length);
            TestAssert.Equal(1, pages[0].Annotations.Single(annotation => annotation.IsDestination).Destination!.Value.PageIndex);
            var warning = diagnostics.Single(item => item.Id == "PPTX_UNSUPPORTED_CALLOUT");
            TestAssert.Equal(2, warning.SlideIndex);
            TestAssert.Equal("/ppt/slides/slide3.xml", warning.PartName);
        }, unsupportedLast: true);
    }

    public static void PptxSlideNumbersRetainOrderingAndAbsentListRepair()
    {
        foreach (bool absent in new[] { false, true })
            WithDeck("10", true, (document, package) =>
            {
                TestAssert.Equal(2, document.Slides.Count);
                TestAssert.Equal(absent ? "/ppt/slides/slide1.xml" : "/ppt/slides/slide3.xml", document.Slides[0].PartName);
                string[] expected = absent ? ["10", "12"] : ["10", "11"];
                for (int index = 0; index < 2; index++)
                    TestAssert.True(PptxRenderer.InspectTextFrameModels(document, package, index)
                        .SelectMany(frame => frame.Paragraphs).SelectMany(paragraph => paragraph.Runs)
                        .All(run => run.Text == expected[index]), "Fields follow the declared ordering or established absent-list repair order.");
            }, absentList: absent, reorder: !absent);
    }

    public static void PptxSlideNumbersKeepOriginFallbackAndCancellation()
    {
        foreach (string first in new[] { "bad", "-1", "4294967296" })
            WithDeck(first, false, (document, _) => TestAssert.Equal(1L, document.FirstSlideNumber));
        WithDeck("10", true, (_, package) =>
        {
            using var cancelled = new CancellationTokenSource();
            cancelled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => new PptxReader().Read(package, cancelled.Token));
        });
    }

    private static void WithDeck(string first, bool hidden, Action<PptxDocument,OoxPackage> check,
        bool unsupportedLast = false, bool absentList = false, bool reorder = false)
    {
        const string p = "http://schemas.openxmlformats.org/presentationml/2006/main";
        const string a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        const string r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        const string ns = "xmlns:p=\""+p+"\" xmlns:a=\""+a+"\" xmlns:r=\""+r+"\"";
        string Field() => "<a:p><a:fld id=\"{00000000-0000-0000-0000-000000000001}\" type=\"slidenum\"><a:rPr sz=\"1200\"><a:latin typeface=\"Arial\"/></a:rPr><a:t>99</a:t></a:fld></a:p>";
        string Shape(bool linked = false, bool unsupported = false) =>
            "<p:sp><p:nvSpPr><p:cNvPr id=\"2\" name=\"Number\">"+(linked ? "<a:hlinkClick r:id=\"jump\" action=\"ppaction://hlinksldjump\"/>" : "")+
            "</p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"914400\" y=\"914400\"/><a:ext cx=\"1828800\" cy=\"914400\"/></a:xfrm><a:prstGeom prst=\""+
            (unsupported ? "wedgeRoundRectCallout" : "rect")+"\"/></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>"+Field()+"</p:txBody></p:sp>";
        string Table() => "<p:graphicFrame><p:xfrm><a:off x=\"914400\" y=\"2743200\"/><a:ext cx=\"1828800\" cy=\"914400\"/></p:xfrm><a:graphic><a:graphicData uri=\""+a.Replace("/main","/table")+"\"><a:tbl><a:tblPr/><a:tblGrid><a:gridCol w=\"1828800\"/></a:tblGrid><a:tr h=\"914400\"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/>"+Field()+"</a:txBody><a:tcPr/></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>";
        string Rel(string id, string kind, string target) => "<Relationship Id=\""+id+"\" Type=\""+r+"/"+kind+"\" Target=\""+target+"\"/>";
        string Rels(string contents) => "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"+contents+"</Relationships>";
        string Story(string kind, string contents) => "<p:"+kind+" "+ns+"><p:cSld><p:spTree>"+contents+"</p:spTree></p:cSld></p:"+kind+">";
        string ids = absentList ? "" : "<p:sldIdLst>"+string.Concat((reorder ? new[] {3,1,2} : new[] {1,2,3}).Select(i => "<p:sldId id=\""+(255+i)+"\" r:id=\"s"+i+"\"/>"))+"</p:sldIdLst>";
        var parts = new Dictionary<string,string> {
            ["[Content_Types].xml"]=PptxTests.BasicContentTypes().Replace("</Types>",string.Concat(new[] {2,3}.Select(i=>"<Override PartName=\"/ppt/slides/slide"+i+".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/>"))+"<Override PartName=\"/ppt/slideLayouts/slideLayout1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml\"/><Override PartName=\"/ppt/slideMasters/slideMaster1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml\"/></Types>"),
            ["_rels/.rels"]=PptxTests.PackageRelationship(),
            ["ppt/presentation.xml"]="<p:presentation "+ns+" firstSlideNum=\""+first+"\">"+ids+"<p:sldSz cx=\"9144000\" cy=\"6858000\"/></p:presentation>",
            ["ppt/_rels/presentation.xml.rels"]=Rels(string.Concat(new[] {1,2,3}.Select(i=>Rel("s"+i,"slide","slides/slide"+i+".xml")))),
            ["ppt/slideLayouts/slideLayout1.xml"]=Story("sldLayout",Shape()),
            ["ppt/slideLayouts/_rels/slideLayout1.xml.rels"]=Rels(Rel("master","slideMaster","../slideMasters/slideMaster1.xml")),
            ["ppt/slideMasters/slideMaster1.xml"]=Story("sldMaster",Shape())
        };
        for (int i=1;i<=3;i++) {
            parts["ppt/slides/slide"+i+".xml"]=Story("sld",Shape(i==1,unsupportedLast && i==3)+Table()).Replace("<p:sld ","<p:sld show=\""+(hidden && i==2 ? "0" : "1")+"\" ");
            parts["ppt/slides/_rels/slide"+i+".xml.rels"]=Rels(Rel("layout","slideLayout","../slideLayouts/slideLayout1.xml")+Rel("jump","slide","slide3.xml"));
        }
        string path=TestFixtures.WriteTempPackage(".pptx",parts);
        try {
            using var input=File.OpenRead(path);
            var package=OoxPackage.Open(input,default);
            check(new PptxReader().Read(package,default),package);
        }
        finally {File.Delete(path);}
    }
}
