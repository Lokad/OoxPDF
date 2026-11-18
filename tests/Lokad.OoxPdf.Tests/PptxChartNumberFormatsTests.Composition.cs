using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxChartNumberFormatsTests
{
    public static void PptxChartNumberFractionLabelsKeepLargeMagnitudes()
    {
        string xml = """
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>-3000000000.5</c:v></c:pt><c:pt idx="1"><c:v>3000000000.5</c:v></c:pt></c:numLit></c:val>
              </c:ser><c:dLbls><c:numFmt formatCode="# ?/?" sourceLinked="0"/>
                <c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/>
              </c:dLbls></c:barChart></c:plotArea></c:chart>
            </c:chartSpace>
            """;
        // Wrapping may separate words or digit chunks; every painted label
        // digit and fraction must retain the source's independently known value.
        string labels = string.Concat(NumberChartLabels(xml, "CBD"));
        TestAssert.Equal("-30000000001/230000000001/2", new string(labels.Where(c => !char.IsWhiteSpace(c)).ToArray()));
    }

    public static void PptxChartNumberBubbleLabelsKeepSourceDateContext()
    {
        foreach (bool date1904 in new[] { false, true })
        foreach (var (sourceFormat, expected) in new[] {
            ("yyyy-mm-dd", date1904 ? "2027-03-16" : "2023-03-15"),
            ("0.0", "45000.0"), ("General", "45000"), ("", "45000") })
        foreach (bool explicitFormat in new[] { false, true })
        {
            string xml = $$"""
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
                  <c:date1904 val="{{(date1904 ? 1 : 0)}}"/><c:chart><c:plotArea><c:bubbleChart>
                    <c:ser><c:idx val="0"/><c:order val="0"/>
                      <c:xVal><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:xVal>
                      <c:yVal><c:numLit><c:pt idx="0"><c:v>3</c:v></c:pt></c:numLit></c:yVal>
                      <c:bubbleSize><c:numLit><c:formatCode>{{sourceFormat}}</c:formatCode><c:pt idx="0"><c:v>45000</c:v></c:pt></c:numLit></c:bubbleSize>
                    </c:ser><c:dLbls>{{(explicitFormat ? $"<c:numFmt formatCode=\"{sourceFormat}\" sourceLinked=\"0\"/>" : "")}}
                      <c:showBubbleSize val="1"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/>
                    </c:dLbls><c:axId val="10"/><c:axId val="20"/></c:bubbleChart>
                    <c:valAx><c:axId val="10"/><c:axPos val="b"/><c:crossAx val="20"/></c:valAx>
                    <c:valAx><c:axId val="20"/><c:axPos val="l"/><c:crossAx val="10"/></c:valAx>
                  </c:plotArea></c:chart>
                </c:chartSpace>
                """;
            TestAssert.Equal(expected, string.Concat(NumberChartLabels(xml, "CSD")));
        }
    }

    private static IReadOnlyList<string> NumberChartLabels(string xml, string fontPrefix)
    {
        using var input = PptxTests.CreateSingleChartPackageStream(xml);
        using var package = OoxPackage.Open(input, default);
        var diagnostics = new List<OoxPdfDiagnostic>();
        var page = new PptxRenderer(new PortableRenderingTests.PinnedResolver())
            .RenderPagesAsync(new PptxReader().Read(package, default), package, diagnostics.Add, default)
            .ToBlockingEnumerable().Single();
        TestAssert.Equal(0, diagnostics.Count);
        return DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
                fontResourceFilter: name => name.StartsWith(fontPrefix, StringComparison.Ordinal))
            .Select(show => show.Text).ToArray();
    }
}
