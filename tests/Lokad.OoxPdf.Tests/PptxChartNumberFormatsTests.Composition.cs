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

    public static void PptxChartNumberLabelFormatsKeepCodeAndWorkbookLinkOwners()
    {
        var workbookParts = new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(PptxTests.BasicContentTypes().Replace("</Types>", """
                  <Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>
                </Types>
                """, StringComparison.Ordinal)),
            ["ppt/embeddings/format-source.xlsx"] = PptxTests.EmbeddedChartWorkbook(),
            ["ppt/charts/_rels/chart1.xml.rels"] = TestFixtures.Utf8("""
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdBook" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"
                    Target="../embeddings/format-source.xlsx"/>
                </Relationships>
                """)
        };
        const string componentFlags = """
            <c:showVal val="1"/><c:showPercent val="0"/><c:showCatName val="0"/><c:showSerName val="0"/>
            <c:showLegendKey val="0"/><c:showBubbleSize val="0"/>
            """;
        foreach (var (date1904, date) in new[] { ("0", "3/15/23"), ("1", "3/16/27") })
        foreach (var (seriesFormat, pointFormat, expected) in new[]
        {
            ("", "", date),
            (Format("", "0"), "", "45000.000"),
            (Format("", "0"), Format("", "1"), date),
            (Format("", "0"), Format("General", "0"), "45000.00"),
            (Format("0.0", "0"), Format("", "0"), "45000.0"),
            (Format("0.0", "0"), Format("0.0000", "0"), "45000.0000")
        })
        {
            string chart = $$"""
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <c:date1904 val="{{date1904}}"/><c:chart><c:plotVisOnly val="0"/><c:plotArea>
                    <c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
                      <c:ser><c:idx val="0"/><c:order val="0"/>
                        <c:dLbls><c:dLbl><c:idx val="0"/>{{pointFormat}}{{componentFlags}}</c:dLbl>
                          {{seriesFormat}}{{componentFlags}}</c:dLbls>
                        <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                        <c:val><c:numRef><c:f>Sheet1!$B$2</c:f><c:numCache><c:formatCode>0.00</c:formatCode>
                          <c:ptCount val="1"/><c:pt idx="0"><c:v>45000</c:v></c:pt>
                        </c:numCache></c:numRef></c:val>
                      </c:ser>
                      <c:dLbls><c:numFmt formatCode="0.000" sourceLinked="1"/>{{componentFlags}}</c:dLbls>
                    </c:barChart></c:plotArea></c:chart><c:externalData r:id="rIdBook"/>
                </c:chartSpace>
                """;
            TestAssert.Equal(expected, string.Concat(NumberChartLabels(chart, "CBD", workbookParts)));
        }

        static string Format(string code, string linked) => $"""<c:numFmt formatCode="{code}" sourceLinked="{linked}"/>""";
    }

    private static IReadOnlyList<string> NumberChartLabels(string xml, string fontPrefix,
        IReadOnlyDictionary<string, byte[]>? additionalParts = null)
    {
        using var input = PptxTests.CreateSingleChartPackageStream(xml, additionalParts);
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
