using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxDisplayMathKeepsOperandsAroundDrawingCompanions()
    {
        foreach (bool withBreak in new[] { false, true })
            CheckDisplayDrawingCompanions(withBreak, explicitRow: false);
    }

    public static void DocxDisplayRowKeepsAdjacentOperandsAroundDrawingCompanions() =>
        CheckDisplayDrawingCompanions(withBreak: false, explicitRow: true);

    public static void DocxDisplayRowKeepsBookmarkAcrossDrawingCompanions() =>
        CheckDisplayDrawingCompanions(withBreak: false, explicitRow: true, withBookmark: true);

    private static void CheckDisplayDrawingCompanions(bool withBreak, bool explicitRow, bool withBookmark = false)
    {
        foreach (string container in new[] { "body", "cell" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var expected = Render(packed: false);
            var actual = Render(packed: true);
            TestAssert.Equal(3, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "314", mathOnly: true).Count);
            TestAssert.Equal(3, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "793", mathOnly: true).Count);
            if (explicitRow)
            {
                TestAssert.Equal(2, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "17", mathOnly: true).Count);
                TestAssert.Equal(2, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "28", mathOnly: true).Count);
            }
            TestAssert.Equal(2, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "QA").Count);
            TestAssert.Equal(1, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "Ω").Count);
            TestAssert.Equal(1, actual.Page.Images.Count);
            TestAssert.Equal(expected.Page.Content, actual.Page.Content);
            if (withBookmark)
                TestAssert.Equal(expected.Page.Annotations.Single(), actual.Page.Annotations.Single());
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
                        var leftNeighbor = explicitRow ? MathRun(new XElement(m + "t", "17")) : null;
                        var rightNeighbor = explicitRow ? MathRun(new XElement(m + "t", "28")) : null;
                        XElement? bookmarkStart = null;
                        XElement? bookmarkEnd = null;
                        if (withBookmark)
                        {
                            const string name = "EquationDrawing";
                            bookmarkStart = new XElement(w + "bookmarkStart", new XAttribute(w + "id", "412"), new XAttribute(w + "name", name));
                            bookmarkEnd = new XElement(w + "bookmarkEnd", new XAttribute(w + "id", "412"));
                            xml.Descendants(w + "p").First().Add(
                                new XElement(w + "r", new XElement(w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), " ")),
                                new XElement(w + "hyperlink", new XAttribute(w + "anchor", name),
                                    new XElement(w + "r", new XElement(w + "t", "Jump row"))));
                        }
                        if (packed)
                        {
                            var run = MathRun(leftOperand, lineBreak, label, symbol, drawing, rightOperand);
                            display.Add(explicitRow ? new XElement(m + "oMath", bookmarkStart, leftNeighbor, run, rightNeighbor, bookmarkEnd) : run);
                        }
                        else
                        {
                            var leftRun = MathRun(leftOperand);
                            var rightRun = MathRun(rightOperand);
                            display.Add(bookmarkStart,
                                explicitRow ? new XElement(m + "oMath", leftNeighbor, leftRun) : leftRun,
                                new XElement(w + "r", wordProperties, lineBreak, label, symbol, drawing),
                                explicitRow ? new XElement(m + "oMath", rightRun, rightNeighbor) : rightRun,
                                bookmarkEnd);
                        }
                        originalPictureRun.Remove();

                        XElement MathRun(params XElement?[] content) => new(m + "r",
                            new XElement(mathProperties), new XElement(wordProperties), content);
                    }, diagnosticSink: Record);

                var paragraph = DocxBlockTraversal.EnumerateBodyParagraphs(source).Single(p => p.Images.Count != 0);
                var items = TestAssert.NotNull(paragraph.InlineContent).ToArray();
                string[] before = explicitRow ? ["17", "314"] : ["314"];
                string[] after = explicitRow ? ["793", "28"] : ["793"];
                int first = Array.FindIndex(items, item => Matches(item, before));
                int picture = Array.FindIndex(items, item => item is DocxInlinePicture);
                int last = Array.FindIndex(items, item => Matches(item, after));
                TestAssert.True(first >= 0 && first < picture && picture < last,
                    "Both operands and the drawing retain their source order in the paragraph.");
                TestAssert.Equal(withBreak, items[(first + 1)..picture].Any(item => item is DocxInlineBreak));
                var page = PictureRangeRenderer(mode).RenderPagesAsync(source, Record, default).ToBlockingEnumerable().Single();
                return (page, diagnostics, source.EquationParts.Sum(part => part.SourceCount));

                static bool Matches(DocxInlineItem item, string[] labels) => item is DocxInlineEquation math &&
                    math.Equation.Content.Children.Count == labels.Length &&
                    math.Equation.Content.Children.OfType<DocxMathRun>().Select(run => run.Text.Text).SequenceEqual(labels);
            }
        }
    }
}
