using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxDisplayMathKeepsOperandsAroundDrawingCompanions()
    {
        foreach (string container in new[] { "body", "cell" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (bool withBreak in new[] { false, true })
        {
            var expected = Render(packed: false);
            var actual = Render(packed: true);
            TestAssert.Equal(3, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "314", mathOnly: true).Count);
            TestAssert.Equal(3, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "793", mathOnly: true).Count);
            TestAssert.Equal(2, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "QA").Count);
            TestAssert.Equal(1, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "Ω").Count);
            TestAssert.Equal(1, actual.Page.Images.Count);
            TestAssert.Equal(expected.Page.Content, actual.Page.Content);
            TestAssert.True(expected.Diagnostics.SequenceEqual(actual.Diagnostics) &&
                !actual.Diagnostics.Contains("DOCX_UNSUPPORTED_EQUATION"),
                "Every projected equation fragment must have a matching source count.");
            TestAssert.Equal(expected.SourceEquationCount, actual.SourceEquationCount);

            (PdfPage Page, List<string> Diagnostics, int SourceEquationCount) Render(bool packed)
            {
                var diagnostics = new List<string>();
                void Record(Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic item) => diagnostics.Add(item.Id);
                var source = ReadMutation("CompositionProbes/math-display-image-" + container + "-between", mode,
                    false, false, (xml, w, m) =>
                    {
                        var display = xml.Descendants(m + "oMathPara").First();
                        var originalPictureRun = display.Parent!.Elements(w + "r").Single();
                        var drawing = originalPictureRun.Element(w + "drawing")!;
                        var mathProperties = new XElement(m + "rPr", new XElement(m + "sty", new XAttribute(m + "val", "p")));
                        var wordProperties = new XElement(w + "rPr", new XElement(w + "color", new XAttribute(w + "val", "005599")));
                        var leftOperand = new XElement(m + "t", "314");
                        var rightOperand = new XElement(m + "t", "793");
                        var label = new XElement(w + "t", "QA");
                        var symbol = new XElement(w + "sym", new XAttribute(w + "font", "Arial"), new XAttribute(w + "char", "03A9"));
                        var lineBreak = withBreak ? new XElement(w + "br") : null;
                        if (packed)
                        {
                            display.Add(new XElement(m + "r", mathProperties, wordProperties,
                                leftOperand, lineBreak, label, symbol, drawing, rightOperand));
                        }
                        else
                        {
                            display.Add(new XElement(m + "r", mathProperties, wordProperties, leftOperand),
                                new XElement(w + "r", wordProperties, lineBreak, label, symbol, drawing),
                                new XElement(m + "r", mathProperties, wordProperties, rightOperand));
                        }
                        originalPictureRun.Remove();
                    }, diagnosticSink: Record);

                var paragraph = DocxBlockTraversal.EnumerateBodyParagraphs(source).Single(p => p.Images.Count != 0);
                var items = TestAssert.NotNull(paragraph.InlineContent).ToArray();
                int first = Array.FindIndex(items, item =>
                    item is DocxInlineEquation { Equation.Content.Children: [DocxMathRun { Text.Text: "314" }] });
                int picture = Array.FindIndex(items, item => item is DocxInlinePicture);
                int last = Array.FindIndex(items, item =>
                    item is DocxInlineEquation { Equation.Content.Children: [DocxMathRun { Text.Text: "793" }] });
                TestAssert.True(first >= 0 && first < picture && picture < last,
                    "Both operands and the drawing retain their source order in the paragraph.");
                TestAssert.Equal(withBreak, items[(first + 1)..picture].Any(item => item is DocxInlineBreak));
                var page = PictureRangeRenderer(mode).RenderPagesAsync(source, Record, default).ToBlockingEnumerable().Single();
                return (page, diagnostics, source.EquationParts.Sum(part => part.SourceCount));
            }
        }
    }
}
