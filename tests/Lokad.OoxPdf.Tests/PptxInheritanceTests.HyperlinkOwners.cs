using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxInheritanceTests
{
    public static void PptxHyperlinkDiagnosticsKeepNodeAndRunPartOwners()
    {
        foreach (bool action in new[] { false, true })
        {
            string? control = null;
            foreach (bool distinct in new[] { false, true })
            {
                var parts = HyperlinkOwnerParts(action, distinct);
                PdfPage page = CheckHyperlinkOwners(parts, action, charts: false);
                control ??= page.Content;
                TestAssert.Equal(control, page.Content);
            }
        }
    }

    public static void PptxHyperlinkDiagnosticsKeepChartOwnersAcrossTextRoles()
    {
        foreach (string role in new[] { "title", "axis", "label" })
        foreach (bool action in new[] { false, true })
        {
            string? control = null;
            foreach (bool distinct in new[] { false, true })
            {
                PdfPage page = CheckHyperlinkOwners(HyperlinkOwnerParts(action, distinct, role), action, charts: true);
                foreach (string word in new[] { "CHARTONE", "CHARTTWO" })
                    TestAssert.Equal(word.Length, PortableRenderingTests.ReadPaintedSourceLabel(page, word).Count);
                control ??= page.Content;
                TestAssert.Equal(control, page.Content);
            }
        }
    }

    private static PdfPage CheckHyperlinkOwners(Dictionary<string, byte[]> parts, bool action, bool charts)
    {
        PdfPage? page = null;
        var diagnostics = new List<OoxPdfDiagnostic>();
        WithDeck(parts, (document, package) => page = Render(document, package, diagnostics).Single());
        PdfPage actual = page ?? throw new InvalidOperationException("The hyperlink ownership source produced no page.");
        var warnings = diagnostics.Where(item => item.Feature == "hyperlink").ToArray();
        string[] owners = ["/ppt/slideMasters/slideMaster1.xml", "/ppt/slideLayouts/slideLayout1.xml", "/ppt/slides/slide1.xml"];
        if (charts) owners = [.. owners, "/ppt/charts/chart1.xml", "/ppt/charts/chart2.xml"];
        TestAssert.Equal(owners.Length, warnings.Length);
        TestAssert.True(warnings.All(item => item.Id == (action ? "PPTX_UNSUPPORTED_HYPERLINK_ACTION" : "PPTX_UNSUPPORTED_HYPERLINK") &&
            item.SlideIndex == 1 && item.Severity == OoxPdfSeverity.Warning && item.Fallback == "Ignored"),
            "Part ownership does not change diagnostic ID, location, severity or fallback.");
        foreach (string owner in owners)
            TestAssert.Equal(1, warnings.Count(item => item.PartName == owner));
        foreach (string word in new[] { "MASTER", "LAYOUT", "SLIDE" })
            TestAssert.Equal(word.Length, PortableRenderingTests.ReadPaintedSourceLabel(actual, word).Count);
        TestAssert.Equal(0, actual.Annotations.Count);
        PdfContentValidator.ValidatePage(actual, 0, default);
        return actual;
    }

    private static Dictionary<string, byte[]> HyperlinkOwnerParts(bool action, bool distinct, string? chartRole = null)
    {
        var parts = PlaceholderFrameParts("", "", "");
        string Click(string owner) => action ? "action='urn:public:unsupported:" + (distinct ? owner : "shared") + "'"
            : "r:id='missing" + (distinct ? owner : "") + "'";
        string[] names = ["ppt/slideMasters/slideMaster1.xml", "ppt/slideLayouts/slideLayout1.xml", "ppt/slides/slide1.xml"];
        string[] words = ["MASTER", "LAYOUT", "SLIDE"];
        for (int index = 0; index < names.Length; index++)
        {
            string click = Click(words[index]);
            XElement shape = XElement.Parse($"<root xmlns:p='{P}' xmlns:a='{A}' xmlns:r='{R}'>" +
                PlaceholderShape(null, PlaceholderFrame(40 + index * 220, 200, 80), words[index]) + "</root>").Elements().Single();
            shape.Descendants(P + "ph").Remove();
            shape.Descendants(P + "cNvPr").Single().Add(XElement.Parse($"<a:hlinkClick xmlns:a='{A}' xmlns:r='{R}' {click}/>"));
            shape.Descendants(A + "r").Single().AddFirst(XElement.Parse($"<a:rPr xmlns:a='{A}' xmlns:r='{R}' sz='1800'><a:latin typeface='DejaVu Sans'/><a:hlinkClick {click}/></a:rPr>"));
            Edit(parts, names[index], xml => xml.Descendants(P + "spTree").Single().Add(shape));
        }
        if (chartRole is null) return parts;

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        for (int index = 1; index <= 2; index++)
        {
            string word = index == 1 ? "CHARTONE" : "CHARTTWO";
            // This ownership control pins the regular face. Synthetic bold's
            // fill-and-stroke mode has separate native coverage; the shared
            // managed source-label oracle deliberately accepts fill-only text.
            string rich = "<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz='1200' b='0'><a:latin typeface='DejaVu Sans'/><a:hlinkClick " + Click(word) +
                "/></a:rPr><a:t>" + word + "</a:t></a:r></a:p></c:rich></c:tx>";
            string title = "<c:title>" + rich + "<c:overlay val='0'/></c:title>";
            string chart = $$"""
                <c:chartSpace xmlns:c="{{c}}" xmlns:a="{{A}}" xmlns:r="{{R}}"><c:chart>
                {{(chartRole == "title" ? title : "")}}<c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
                <c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Series</c:v></c:tx><c:cat><c:strLit><c:ptCount val="1"/><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser>
                {{(chartRole == "label" ? "<c:dLbls><c:dLbl><c:idx val='0'/>" + rich + "</c:dLbl></c:dLbls>" : "")}}
                <c:axId val="10"/><c:axId val="20"/></c:barChart>
                <c:catAx><c:axId val="10"/><c:scaling/><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="20"/></c:catAx>
                <c:valAx><c:axId val="20"/><c:scaling><c:min val="0"/><c:max val="5"/></c:scaling><c:axPos val="l"/>
                {{(chartRole == "axis" ? title : "")}}<c:tickLblPos val="none"/><c:crossAx val="10"/><c:majorUnit val="5"/></c:valAx>
                </c:plotArea></c:chart></c:chartSpace>
                """;
            parts[$"ppt/charts/chart{index}.xml"] = TestFixtures.Utf8(chart);
            parts[$"ppt/charts/_rels/chart{index}.xml.rels"] = TestFixtures.Utf8("<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'/>");
            int chartIndex = index;
            Edit(parts, "[Content_Types].xml", xml => xml.Root!.Add(new XElement(xml.Root.Name.Namespace + "Override",
                new XAttribute("PartName", $"/ppt/charts/chart{chartIndex}.xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"))));
            Edit(parts, "ppt/slides/_rels/slide1.xml.rels", xml => xml.Root!.Add(new XElement(xml.Root.Name.Namespace + "Relationship",
                new XAttribute("Id", $"chart{chartIndex}"), new XAttribute("Type", R.NamespaceName + "/chart"), new XAttribute("Target", $"../charts/chart{chartIndex}.xml"))));
            Edit(parts, "ppt/slides/slide1.xml", xml => xml.Descendants(P + "spTree").Single().Add(XElement.Parse($$"""
                <p:graphicFrame xmlns:p="{{P}}" xmlns:a="{{A}}" xmlns:c="{{c}}" xmlns:r="{{R}}"><p:nvGraphicFramePr><p:cNvPr id="{{chartIndex + 10}}" name="Public chart"/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="{{(chartIndex - 1) * 4572000}}" y="2540000"/><a:ext cx="4064000" cy="3556000"/></p:xfrm><a:graphic><a:graphicData uri="{{c}}"><c:chart r:id="chart{{chartIndex}}"/></a:graphicData></a:graphic></p:graphicFrame>
                """)));
        }
        return parts;
    }
}
