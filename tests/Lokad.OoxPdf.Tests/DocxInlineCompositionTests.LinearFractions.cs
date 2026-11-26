using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxLinearFractionsKeepSlashRowsAcrossDocumentStories()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (string part in new[] { "word/document.xml", "word/footer1.xml" })
        {
            var renderer = MathContentControlRenderer(mode);
            var controlDiagnostics = new List<OoxPdfDiagnostic>();
            var diagnostics = new List<OoxPdfDiagnostic>();
            var control = renderer.RenderPagesAsync(ReadMutation("Composition/composed-static-cell-report", mode, false, false,
                (xml, w, m) => Rewrite(xml, w, m, linear: false), part), controlDiagnostics.Add, default)
                .ToBlockingEnumerable().ToArray();
            var actual = renderer.RenderPagesAsync(ReadMutation("Composition/composed-static-cell-report", mode, false, false,
                (xml, w, m) => Rewrite(xml, w, m, linear: true), part), diagnostics.Add, default)
                .ToBlockingEnumerable().ToArray();

            TestAssert.True(controlDiagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "The authored slash row must render before it can serve as a control.");
            TestAssert.True(controlDiagnostics.SequenceEqual(diagnostics), "Linear fractions keep the slash row diagnostics.");
            AssertMathContentControlPages(control, actual);
        }

        static void Rewrite(XDocument xml, XNamespace w, XNamespace m, bool linear)
        {
            XElement[] fractions = xml.Descendants(m + "f").ToArray();
            TestAssert.True(fractions.Length >= 2, "The composed source must include multiple equations.");
            foreach (XElement fraction in fractions)
            {
                TestAssert.Equal<XElement?>(null, fraction.Element(m + "fPr"));
                if (linear)
                {
                    fraction.AddFirst(new XElement(m + "fPr", new XElement(m + "type", new XAttribute(m + "val", "lin"))));
                }
                else
                {
                    fraction.ReplaceWith(fraction.Element(m + "num")!.Elements().Select(child => new XElement(child)).ToArray(),
                        new XElement(m + "r", new XElement(m + "t", "/")),
                        fraction.Element(m + "den")!.Elements().Select(child => new XElement(child)).ToArray());
                }
            }
        }
    }
}
