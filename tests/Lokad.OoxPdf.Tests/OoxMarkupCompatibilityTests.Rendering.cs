using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxMarkupCompatibilityTests
{
    public static void IgnoredExtensionChartPayloadKeepsIndependentTextAndValues()
    {
        byte[]? plainPdf = null;
        foreach (string variant in new[] { "plain", "ignored", "processed" })
        {
            string extra = variant == "plain" ? "" : "<e:box><a:r><a:t>EXTRA</a:t></a:r></e:box>";
            string point = variant == "plain" ? "" : "<e:box><c:pt idx=\"1\"><c:v>3</c:v></c:pt></e:box>";
            string process = variant == "processed" ? " mc:ProcessContent=\"e:box\"" : "";
            string chart = $$"""
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                  xmlns:mc="{{Mc}}" xmlns:e="urn:extension" mc:Ignorable="e"{{process}}>
                  <c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p>
                    <a:r><a:rPr sz="1800"/><a:t>VISIBLE</a:t></a:r>{{extra}}
                  </a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:plotArea>
                  <c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
                    <c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Series</c:v></c:tx>
                    <c:cat><c:strLit><c:ptCount val="1"/><c:pt idx="0"><c:v>Category</c:v></c:pt></c:strLit></c:cat>
                    <c:val><c:numLit><c:ptCount val="{{(variant == "processed" ? 2 : 1)}}"/>
                      <c:pt idx="0"><c:v>2</c:v></c:pt>{{point}}</c:numLit></c:val></c:ser>
                    <c:axId val="1"/><c:axId val="2"/></c:barChart>
                  <c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
                  <c:valAx><c:axId val="2"/><c:scaling><c:min val="0"/><c:max val="4"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
                  </c:plotArea></c:chart></c:chartSpace>
                """;
            XDocument normalized = Load(chart);
            XNamespace c = OoxNamespaces.ChartNamespace;
            TestAssert.Equal(variant == "processed" ? "VISIBLEEXTRA" : "VISIBLE",
                PptxSceneBuilder.ReadChartText(normalized.Descendants(c + "title").Single().Element(c + "tx"), false));
            var series = PptxSceneBuilder.ReadChartSeries(normalized.Descendants(c + "barChart").Single(),
                PptxTheme.Empty, PptxColorMap.Default, PptxSceneChartPlotKind.Bar, false).Single();
            TestAssert.Equal(variant == "processed" ? 2 : 1, series.ValuePoints.Count);
            TestAssert.Equal(2d, series.ValuePoints[0].Value!.Value);
            if (variant == "processed") TestAssert.Equal(3d, series.ValuePoints[1].Value!.Value);

            using var input = PptxTests.CreateSingleChartPackageStream(chart);
            using var output = new MemoryStream();
            var diagnostics = new List<OoxPdfDiagnostic>();
            OoxPdfConverter.Convert(input, output, new OoxPdfOptions {
                InputKind = OoxPdfInputKind.Pptx, FontResolver = new PortableRenderingTests.PinnedResolver(), DiagnosticSink = diagnostics.Add
            });
            TestAssert.Equal(0, diagnostics.Count);
            if (variant == "plain") plainPdf = output.ToArray();
            else TestAssert.True((variant == "ignored") == output.ToArray().AsSpan().SequenceEqual(plainPdf),
                "Ignored payload preserves the complete ordinary PDF; explicitly processed payload remains visible.");
        }
    }

    public static void CompatibilityNormalizationPreservesSupportedCoreNamespaces()
    {
        foreach (string ns in new[] {
            OoxNamespaces.MathNamespace.NamespaceName, "http://schemas.openxmlformats.org/drawingml/2006/picture",
            OoxNamespaces.SpreadsheetNamespace.NamespaceName, OoxNamespaces.RelationshipsNamespace.NamespaceName,
            OoxNamespaces.PackageRelationshipsNamespace.NamespaceName, OoxNamespaces.ContentTypesNamespace.NamespaceName
        })
        {
            XDocument document = Load($$"""
                <root xmlns:mc="{{Mc}}" xmlns:s="{{ns}}" mc:Ignorable="s" mc:MustUnderstand="s">
                  <s:payload>Kept</s:payload>
                </root>
                """);
            TestAssert.Equal("Kept", document.Root?.Element(XName.Get("payload", ns))?.Value);
            TestAssert.True(!OoxMarkupCompatibility.HasUnrecognizedMustUnderstand(document), "Supported namespaces stay understood.");
        }
    }

    public static void ProcessContentRetainsDirectGalleryStyleWithoutUnderstandingOtherC14Content()
    {
        XNamespace c = OoxNamespaces.ChartNamespace;
        XDocument document = Load($$"""
            <c:chartSpace xmlns:c="{{c}}" xmlns:mc="{{Mc}}" xmlns:c14="{{OoxNamespaces.ChartStyle2010Namespace}}" mc:Ignorable="c14">
              <c14:style val="109"/>
              <c14:box><c:style val="7"/></c14:box>
              <mc:AlternateContent><mc:Choice Requires="c14"><c:style val="8"/></mc:Choice>
                <mc:Fallback><c:style val="2"/></mc:Fallback></mc:AlternateContent>
            </c:chartSpace>
            """);
        TestAssert.Equal<int?>(109, OoxMarkupCompatibility.ReadPreferredChartGalleryStyle(document.Root));
        TestAssert.Equal("2", document.Root?.Elements(c + "style").Single().Attribute("val")?.Value);
        TestAssert.Equal(1, document.Descendants(c + "style").Count());
    }

    public static void CompatibilityNormalizationHonorsCancellationDuringIgnoredRemoval()
    {
        var document = XDocument.Parse($$"""
            <w:document xmlns:w="{{Word}}" xmlns:mc="{{Mc}}" xmlns:e="urn:extension" mc:Ignorable="e">
              <e:box><w:p>Ignored</w:p></e:box><w:p>Next</w:p>
            </w:document>
            """);
        using var cancellation = new CancellationTokenSource();
        document.Changed += (_, _) => cancellation.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => OoxMarkupCompatibility.Normalize(document, cancellation.Token));
    }
}
