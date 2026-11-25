using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
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
            // Only the slide shape has this label. The inherited shapes and
            // table also render the number, with no surrounding text.
            PortableRenderingTests.ReadPaintedSourceLabel(pages[0], "Slide 10 end");
            PortableRenderingTests.ReadPaintedSourceLabel(pages[1], "Slide 12 end");
            TestAssert.Equal(1, pages[0].Annotations.Single(annotation => annotation.IsDestination).Destination!.Value.PageIndex);
            var warning = diagnostics.Single(item => item.Id == "PPTX_UNSUPPORTED_CALLOUT");
            TestAssert.Equal(2, warning.SlideIndex);
            TestAssert.Equal("/ppt/slides/slide3.xml", warning.PartName);
        }, unsupportedLast: true);
    }

    public static void PptxSlideNumbersReserveWrappedTableRowsForDisplayedFields()
    {
        PdfPage[] fields = Render(_ => WrappingNumberTable(null));
        PdfPage[] literals = Render(index => WrappingNumberTable((index + 7).ToString(CultureInfo.InvariantCulture)));

        double ordinaryNextRow = AssertTable(fields[0], "8", 22);
        double wrappedNextRow = AssertTable(fields[1], "10", 44);
        TestAssert.True(Math.Abs(ordinaryNextRow - wrappedNextRow - 22d) < .01d,
            "The added line advances the next row by one line height.");
        for (int index = 0; index < fields.Length; index++)
            TestAssert.Equal(literals[index].Content, fields[index].Content);

        static PdfPage[] Render(Func<int, string> tableForSlide)
        {
            PdfPage[] pages = [];
            WithDeck("8", true, (document, package) =>
            {
                pages = new PptxRenderer(new PortableRenderingTests.PinnedResolver())
                    .RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().ToArray();
            }, tableForSlide: tableForSlide);
            TestAssert.Equal(2, pages.Length);
            return pages;
        }

        static double AssertTable(PdfPage page, string number, int rowHeight)
        {
            PortableRenderingTests.ReadPaintedSourceLabel(page, $"Alpha {number} End");
            var nextRow = PortableRenderingTests.ReadPaintedSourceLabel(page, "Next row");
            TestAssert.Contains(FormattableString.Invariant($"72 {page.Height - 216 - rowHeight} 114 {rowHeight} re f*"), page.Content);
            return nextRow[0].Origin.Y;
        }
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
                {
                    var fields = PptxRenderer.InspectTextFrameModels(document, package, index)
                        .SelectMany(frame => frame.Paragraphs).SelectMany(paragraph => paragraph.Runs)
                        .Where(run => run.Kind == "Field").ToArray();
                    TestAssert.Equal(3, fields.Length);
                    TestAssert.True(fields.All(run => run.Text == expected[index]),
                        "Fields follow the declared ordering or established absent-list repair order.");
                }
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

    private static string WrappingNumberTable(string? literal)
    {
        const string runProperties = """<a:rPr sz="1800"><a:latin typeface="DejaVu Sans"/></a:rPr>""";
        string number = literal is null
            ? $$"""<a:fld id="{00000000-0000-0000-0000-000000000001}" type="slidenum">{{runProperties}}<a:t>1</a:t></a:fld>"""
            : Run(literal);
        return $$"""
            <p:graphicFrame><p:xfrm><a:off x="914400" y="2743200"/><a:ext cx="1447800" cy="1524000"/></p:xfrm>
              <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/>
                <a:tblGrid><a:gridCol w="1447800"/></a:tblGrid>
                {{Row(Run("Alpha ") + number + Run(" End"), "CCEECC")}}
                {{Row(Run("Next row"), "CCCCEE")}}
              </a:tbl></a:graphicData></a:graphic>
            </p:graphicFrame>
            """;

        static string Run(string text) => $"<a:r>{runProperties}<a:t>{text}</a:t></a:r>";

        static string Row(string text, string fill) => $$"""
            <a:tr h="152400"><a:tc><a:txBody><a:bodyPr wrap="square"/><a:lstStyle/>
              <a:p><a:pPr><a:lnSpc><a:spcPts val="2200"/></a:lnSpc></a:pPr>{{text}}</a:p></a:txBody>
              <a:tcPr marL="0" marR="0" marT="0" marB="0"><a:solidFill><a:srgbClr val="{{fill}}"/></a:solidFill></a:tcPr>
            </a:tc></a:tr>
            """;
    }

    private static void WithDeck(string first, bool hidden, Action<PptxDocument,OoxPackage> check,
        bool unsupportedLast = false, bool absentList = false, bool reorder = false, Func<int, string>? tableForSlide = null)
    {
        const string p = "http://schemas.openxmlformats.org/presentationml/2006/main";
        const string a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        const string r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        const string ns = "xmlns:p=\""+p+"\" xmlns:a=\""+a+"\" xmlns:r=\""+r+"\"";
        const string numberStyle = "<a:rPr sz=\"1200\"><a:latin typeface=\"Arial\"/></a:rPr>";
        string Field(bool slideLabel = false) => "<a:p>" + (slideLabel ? Run("Slide ") : "") +
            "<a:fld id=\"{00000000-0000-0000-0000-000000000001}\" type=\"slidenum\">" + numberStyle +
            "<a:t>99</a:t></a:fld>" + (slideLabel ? Run(" end") : "") + "</a:p>";
        string Run(string text) => "<a:r>" + numberStyle + "<a:t xml:space=\"preserve\">" + text + "</a:t></a:r>";
        string Shape(bool linked = false, bool unsupported = false, bool slideLabel = false) =>
            "<p:sp><p:nvSpPr><p:cNvPr id=\"2\" name=\"Number\">"+(linked ? "<a:hlinkClick r:id=\"jump\" action=\"ppaction://hlinksldjump\"/>" : "")+
            "</p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"914400\" y=\"914400\"/><a:ext cx=\"1828800\" cy=\"914400\"/></a:xfrm><a:prstGeom prst=\""+
            (unsupported ? "wedgeRoundRectCallout" : "rect")+"\"/></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>"+Field(slideLabel)+"</p:txBody></p:sp>";
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
            string slideContent = Shape(linked: i == 1, unsupported: unsupportedLast && i == 3, slideLabel: true) +
                (tableForSlide?.Invoke(i) ?? Table());
            parts["ppt/slides/slide"+i+".xml"] = Story("sld", slideContent)
                .Replace("<p:sld ", "<p:sld show=\"" + (hidden && i == 2 ? "0" : "1") + "\" ");
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
