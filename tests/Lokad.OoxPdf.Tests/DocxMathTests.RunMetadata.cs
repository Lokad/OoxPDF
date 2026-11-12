using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxMathTests
{
    public static void DocxMathRevisionAdmissionSharesReadableEmptyRunRules()
    {
        foreach (string kind in new[] { "ins", "del" })
        foreach (bool included in new[] { false, true })
        foreach (string contents in new[] { "", "<m:t/>", "<w:lastRenderedPageBreak/>",
            "<m:rPr><m:nor m:val='0'/></m:rPr>" })
        {
            var xml = EditingMarkerFraction();
            var run = xml.Descendants(MathNamespace + "r").First();
            var empty = XElement.Parse($"<m:r xmlns:m='{MathNamespace}' xmlns:w='{WordprocessingNamespace}'>{contents}</m:r>");
            run.ReplaceWith(new XElement(WordprocessingNamespace + kind, new XElement(run), empty));
            var parsed = DocxMathParser.Read(xml, (_, text) => PlainRun(text), default, _ => included);
            var fraction = (DocxMathFraction)parsed.Children.Single();
            var revision = fraction.Numerator.Children[0] as DocxMathRow;
            TestAssert.NotNull(revision, "A readable empty run does not invalidate its actual revision container.");
            TestAssert.Equal(included ? 2 : 0, revision!.Children.Count);
            if (included) TestAssert.Equal("a", RowText(revision));
            TestAssert.Equal("c", RowText(fraction.Denominator));

            var emptyOnly = EditingMarkerFraction();
            emptyOnly.Descendants(MathNamespace + "r").First().ReplaceWith(
                new XElement(WordprocessingNamespace + kind, new XElement(empty)));
            var unsupported = DocxMathParser.Read(emptyOnly, (_, text) => PlainRun(text), default, _ => included);
            TestAssert.True(((DocxMathFraction)unsupported.Children.Single()).Numerator.Children[0] is DocxMathUnsupported,
                "Empty-only wrappers retain the existing requirement for an actual revision payload.");
        }

        foreach (string kind in new[] { "ins", "del" })
        foreach (bool included in new[] { false, true })
        foreach (string contents in new[] { "<m:future/>", "<w:lastRenderedPageBreak extra='future'/>",
            "<m:rPr><m:future/></m:rPr>", "<m:rPr><m:nor m:val='future'/></m:rPr>" })
        {
            var xml = EditingMarkerFraction();
            var run = xml.Descendants(MathNamespace + "r").First();
            var unknown = XElement.Parse($"<m:r xmlns:m='{MathNamespace}' xmlns:w='{WordprocessingNamespace}'>{contents}</m:r>");
            run.ReplaceWith(new XElement(WordprocessingNamespace + kind, new XElement(run), unknown));
            var parsed = DocxMathParser.Read(xml, (_, text) => PlainRun(text), default, _ => included);
            TestAssert.True(((DocxMathFraction)parsed.Children.Single()).Numerator.Children[0] is DocxMathUnsupported,
                "An unknown empty payload retains its previous revision fallback in either view.");
        }

        foreach (bool included in new[] { false, true })
        {
            var xml = EditingMarkerFraction();
            var run = xml.Descendants(MathNamespace + "r").First();
            run.ReplaceWith(new XElement(WordprocessingNamespace + "del", new XElement(run),
                new XElement(MathNamespace + "r", new XElement(MathNamespace + "t", "Visible"), new XElement(MathNamespace + "future"))));
            var parsed = DocxMathParser.Read(xml, (_, text) => PlainRun(text), default, _ => included);
            var row = (DocxMathRow)((DocxMathFraction)parsed.Children.Single()).Numerator.Children[0];
            TestAssert.True(included ? row.Children.Last() is DocxMathUnsupported : row.Children.Count == 0,
                "Visible unknown content retains its existing deferred fallback and excluded-literal filtering.");
        }
    }

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
