using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxMathTests
{
    public static void DocxMathRunNonPrintingMarkersRetainTextAndFallbacks()
    {
        foreach (string marker in new[] { "lastRenderedPageBreak", "proofErr", "permStart", "permEnd" })
        {
            var xml = EditingMarkerFraction();
            foreach (var run in xml.Descendants(MathNamespace + "r"))
            {
                run.AddFirst(new XElement(WordprocessingNamespace + marker));
                run.Add(new XElement(WordprocessingNamespace + marker));
            }
            var fraction = (DocxMathFraction)Parse(xml).Children.Single();
            TestAssert.Equal("a+b", RowText(fraction.Numerator));
            TestAssert.Equal("c", RowText(fraction.Denominator));
        }

        foreach (string marker in new[]
        {
            "<w:lastRenderedPageBreak extra='future'/>",
            "<w:lastRenderedPageBreak>Visible</w:lastRenderedPageBreak>",
            "<w:lastRenderedPageBreak><m:t>Visible</m:t></w:lastRenderedPageBreak>",
            "<x:lastRenderedPageBreak xmlns:x='urn:future'/>"
        })
        {
            var xml = EditingMarkerFraction();
            var content = XElement.Parse($"<root xmlns:w='{WordprocessingNamespace}' xmlns:m='{MathNamespace}'>{marker}</root>");
            xml.Descendants(MathNamespace + "r").First().Add(content.Elements().Single());
            var fraction = (DocxMathFraction)Parse(xml).Children.Single();
            TestAssert.True(fraction.Numerator.Children[0] is DocxMathUnsupported,
                "Unknown or content-bearing run metadata retains the explicit fallback.");
        }

        var misplaced = EditingMarkerFraction();
        misplaced.Descendants(MathNamespace + "num").Single().AddFirst(new XElement(WordprocessingNamespace + "lastRenderedPageBreak"));
        TestAssert.True(((DocxMathFraction)Parse(misplaced).Children.Single()).Numerator.Children[0] is DocxMathUnsupported,
            "A cached run position does not become a valid mathematical argument.");
    }
}
