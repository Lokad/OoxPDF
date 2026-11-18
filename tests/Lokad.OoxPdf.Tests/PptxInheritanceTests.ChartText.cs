using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxInheritanceTests
{
    public static void PptxInheritanceChartTextFillChoicesReachTitlesAndLabels()
    {
        const string noFill = "<a:noFill/>";
        const string red = "<a:solidFill><a:srgbClr val=\"CC0000\"/></a:solidFill>";
        const string transparent = "<a:solidFill><a:srgbClr val=\"CC0000\"><a:alpha val=\"0\"/></a:srgbClr></a:solidFill>";
        foreach (string owner in new[] { "title", "label" })
        foreach (var (inheritedFill, directFill, visible) in new[] {
            ("", "", true), ("", noFill, false), (noFill, "", false),
            (red, noFill, false), (noFill, red, true), (red, "", true),
            (noFill, transparent, false), (transparent, red, true), ("", transparent, false) })
        {
            string rich = $"<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr>{directFill}</a:rPr><a:t>TARGET</a:t></a:r></a:p></c:rich></c:tx>";
            string title = owner == "title" ? $"<c:title>{rich}</c:title>" : "<c:autoTitleDeleted val=\"1\"/>";
            string labels = owner == "label" ? $"<c:dLbls><c:dLbl><c:idx val=\"0\"/>{rich}</c:dLbl><c:showVal val=\"1\"/></c:dLbls>" : "";
            string xml = $$"""
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <c:chart>{{title}}<c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/>
                    <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                    <c:val><c:numLit><c:pt idx="0"><c:v>3</c:v></c:pt></c:numLit></c:val>
                  </c:ser>{{labels}}</c:barChart></c:plotArea></c:chart>
                  <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr>{{inheritedFill}}</a:defRPr></a:pPr></a:p></c:txPr>
                </c:chartSpace>
                """;
            using var input = PptxTests.CreateSingleChartPackageStream(xml);
            using var package = OoxPackage.Open(input, default);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var page = new PptxRenderer(new PortableRenderingTests.PinnedResolver())
                .RenderPagesAsync(new PptxReader().Read(package, default), package, diagnostics.Add, default)
                .ToBlockingEnumerable().Single();
            string text = DocxPageFootnoteNumberingTests.DecodeText(page);
            TestAssert.True(text.Contains("TARGET", StringComparison.Ordinal) == visible,
                $"{owner} paint must preserve its effective fill: inherited {inheritedFill}, direct {directFill}.");
            TestAssert.Equal(0, diagnostics.Count);
        }
    }
}
