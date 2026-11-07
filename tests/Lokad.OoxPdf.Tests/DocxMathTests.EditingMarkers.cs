using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxMathTests
{
    public static void DocxMathEditingMarkersKeepControlAndArgumentContent()
    {
        foreach (string role in new[] { "oMath", "num", "den" })
        foreach (string markers in new[]
        {
            "<w:proofErr w:type='spellStart'/><w:proofErr w:type='spellEnd'/>",
            "<w:permStart w:id='12' w:edGrp='everyone'/><w:permEnd w:id='12'/>"
        })
        {
            var xml = EditingMarkerFraction();
            var owner = role == "oMath" ? xml : xml.Descendants(MathNamespace + role).Single();
            var pair = XElement.Parse($"<root xmlns:w='{WordprocessingNamespace}'>{markers}</root>").Elements().ToArray();
            owner.AddFirst(pair[0]); owner.Add(pair[1]);
            var fraction = (DocxMathFraction)Parse(xml).Children.Single();
            TestAssert.Equal("a+b", RowText(fraction.Numerator));
            TestAssert.Equal("c", RowText(fraction.Denominator));
        }
    }

    public static void DocxMathEditingMarkersKeepUnknownAndContentBearingFallback()
    {
        foreach (string marker in new[]
        {
            "<x:proofErr xmlns:x='urn:future'/>", "<w:future/>",
            "<w:proofErr w:type='spellStart'>Visible</w:proofErr>",
            "<w:permStart w:id='12'><m:r><m:t>Visible</m:t></m:r></w:permStart>"
        })
        {
            var xml = EditingMarkerFraction();
            var content = XElement.Parse($"<root xmlns:w='{WordprocessingNamespace}' xmlns:m='{MathNamespace}'>{marker}</root>");
            xml.AddFirst(content.Elements().Single());
            var parsed = Parse(xml);
            TestAssert.True(parsed.Children[0] is DocxMathUnsupported,
                "Unknown or content-bearing annotations must not be silently discarded.");
            TestAssert.Equal("a+b", RowText(((DocxMathFraction)parsed.Children[1]).Numerator));
        }
    }

    public static void DocxMathEditingMarkersKeepCooperativeTraversalCancellation()
    {
        var xml = EditingMarkerFraction();
        xml.Element(MathNamespace + "f")!.AddAfterSelf(
            new XElement(MathNamespace + "r", new XElement(MathNamespace + "t", "tail")));
        xml.AddFirst(new XElement(WordprocessingNamespace + "proofErr",
            new XAttribute(WordprocessingNamespace + "type", "spellStart")));
        using var cancellation = new CancellationTokenSource();
        int reads = 0;
        TestAssert.Throws<OperationCanceledException>(() => DocxMathParser.Read(xml, (_, text) =>
        {
            if (++reads == 1) cancellation.Cancel();
            return PlainRun(text);
        }, cancellation.Token));
        TestAssert.Equal(1, reads);
    }

    private static XElement EditingMarkerFraction() => XElement.Parse($"""
        <m:oMath xmlns:m="{MathNamespace}" xmlns:w="{WordprocessingNamespace}">
          <m:f><m:num><m:r><m:t>a</m:t></m:r><m:r><m:t>+</m:t></m:r><m:r><m:t>b</m:t></m:r></m:num>
          <m:den><m:r><m:t>c</m:t></m:r></m:den></m:f>
        </m:oMath>
        """);
}
