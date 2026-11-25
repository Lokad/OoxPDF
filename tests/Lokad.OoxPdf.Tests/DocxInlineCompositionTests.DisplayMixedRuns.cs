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

    public static void DocxInlineEquationKeepsDrawingCompanionsBetweenTextRuns() =>
        CheckDisplayDrawingCompanions(withBreak: false, explicitRow: true, inlineEquation: true);

    public static void DocxDisplayRowKeepsBookmarkAcrossDrawingCompanions() =>
        CheckDisplayDrawingCompanions(withBreak: false, explicitRow: true, withBookmark: true);

    public static void DocxDisplayRowKeepsDrawingCompanionsInTrackedChanges()
    {
        foreach (var revision in new[] { DocxRevisionKind.Insertion, DocxRevisionKind.Deletion })
            CheckDisplayDrawingCompanions(withBreak: false, explicitRow: true, revision: revision);
    }

    public static void DocxDisplayRowKeepsCommentRangeThroughDrawing()
    {
        foreach (string container in new[] { "body", "cell" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var expected = Render(inside: false);
            var actual = Render(inside: true);
            TestAssert.Equal(expected.Page.Content, actual.Page.Content);
            TestAssert.Equal(1, actual.Page.Images.Count);
            TestAssert.Equal(3, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "314", mathOnly: true).Count);
            TestAssert.Equal(3, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "793", mathOnly: true).Count);
            TestAssert.Equal(2, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "QA").Count);
            TestAssert.True(expected.Diagnostics.SequenceEqual(actual.Diagnostics) &&
                !actual.Diagnostics.Contains("DOCX_UNSUPPORTED_EQUATION"), "Range metadata does not drop any equation fragments.");
            string text = DocxPageFootnoteNumberingTests.DecodeText(actual.Page);
            TestAssert.Equal(mode == OoxPdfDocxMarkupMode.AllMarkup, text.Contains("Check printed geometry.", StringComparison.Ordinal));

            (PdfPage Page, List<string> Diagnostics) Render(bool inside)
            {
                var diagnostics = new List<string>();
                void Record(Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic item) => diagnostics.Add(item.Id);
                var document = ReadMutation("CompositionProbes/math-display-image-" + container + "-between", mode,
                    false, false, (xml, w, m) =>
                    {
                        var display = xml.Descendants(m + "oMathPara").First();
                        var originalPicture = display.Parent!.Elements(w + "r").Single();
                        var drawing = originalPicture.Element(w + "drawing")!;
                        var start = xml.Descendants(w + "commentRangeStart").Single();
                        var end = xml.Descendants(w + "commentRangeEnd").Single();
                        var referenceRun = xml.Descendants(w + "commentReference").Single().Parent!;
                        start.Remove();
                        end.Remove();
                        referenceRun.Remove();
                        var picture = new XElement((inside ? m : w) + "r", new XElement(w + "t", "QA"), drawing);
                        originalPicture.Remove();
                        if (inside)
                            display.Add(new XElement(m + "oMath", start, MathRun("314"), picture, end, MathRun("793")));
                        else
                            display.Add(new XElement(m + "oMath", start, MathRun("314")), picture, end,
                                new XElement(m + "oMath", MathRun("793")));
                        display.AddAfterSelf(referenceRun);

                        XElement MathRun(string text) => new(m + "r", new XElement(m + "t", text));
                    }, diagnosticSink: Record);
                var paragraph = DocxBlockTraversal.EnumerateBodyParagraphs(document).Single(p => p.Images.Count != 0);
                var range = paragraph.CommentRanges.Single();
                var items = TestAssert.NotNull(paragraph.InlineContent);
                int start = range.StartInlineItemIndex!.Value;
                int end = range.EndInlineItemIndex!.Value;
                TestAssert.True(items[start] is DocxInlineEquation { Equation.Content.Children: [DocxMathRun { Text.Text: "314" }] } &&
                    items[end - 1] is DocxInlinePicture &&
                    items[end] is DocxInlineEquation { Equation.Content.Children: [DocxMathRun { Text.Text: "793" }] },
                    "The range closes after the picture and before the following operand.");
                var page = PictureRangeRenderer(mode).RenderPagesAsync(document, Record, default).ToBlockingEnumerable().Single();
                return (page, diagnostics);
            }
        }
    }

    private static void CheckDisplayDrawingCompanions(bool withBreak, bool explicitRow, bool withBookmark = false,
        DocxRevisionKind? revision = null, bool inlineEquation = false)
    {
        OoxPdfDocxMarkupMode[] modes = revision is null
            ? [OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup]
            : [OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.Original, OoxPdfDocxMarkupMode.AllMarkup];
        foreach (string container in new[] { "body", "cell" })
        foreach (var mode in modes)
        {
            bool visible = (revision, mode) is not
                ((DocxRevisionKind.Insertion, OoxPdfDocxMarkupMode.Original) or (DocxRevisionKind.Deletion, OoxPdfDocxMarkupMode.Final));
            var expected = Render(packed: false);
            var actual = Render(packed: true);
            TestAssert.Equal(expected.Page.Content, actual.Page.Content);
            TestAssert.True(expected.Diagnostics.SequenceEqual(actual.Diagnostics),
                "Packing should preserve the diagnostics for the same revision view.");
            TestAssert.Equal(expected.SourceEquationCount, actual.SourceEquationCount);
            if (!visible)
            {
                TestAssert.Equal(0, actual.Page.Images.Count);
                string painted = DocxPageFootnoteNumberingTests.DecodeText(actual.Page);
                foreach (string label in new[] { "17", "314", "793", "28", "QA", "Ω" })
                    TestAssert.DoesNotContain(label, painted);
                continue;
            }

            TestAssert.Equal(3, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "314", mathOnly: true).Count);
            TestAssert.Equal(3, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "793", mathOnly: true).Count);
            if (explicitRow)
            {
                var leading = PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "17", mathOnly: true);
                TestAssert.Equal(2, leading.Count);
                TestAssert.Equal(2, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "28", mathOnly: true).Count);
                if (withBookmark)
                {
                    var link = actual.Page.Annotations.Single();
                    var outer = expected.Page.Annotations.Single();
                    TestAssert.Equal((outer.X, outer.Y, outer.Width, outer.Height),
                        (link.X, link.Y, link.Width, link.Height));
                    var ink = leading[0].Ink;
                    TestAssert.True(link.Destination is { PageIndex: 0, Left: double left, Top: double top } &&
                        outer.Destination is { Left: double outerLeft, Top: double outerTop } &&
                        Math.Abs(left - outerLeft) < .002d &&
                        top >= ink.Y + ink.Height - .002d && top <= outerTop + .002d,
                        "The inner bookmark opens above the first operand, within the containing line's viewport.");
                }
            }
            TestAssert.Equal(2, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "QA").Count);
            TestAssert.Equal(1, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "Ω").Count);
            if (inlineEquation)
            {
                TestAssert.Equal(6, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "before").Count);
                TestAssert.Equal(5, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "after").Count);
            }
            TestAssert.Equal(1, actual.Page.Images.Count);
            TestAssert.Contains($"/{actual.Page.Images.Single().ResourceName} Do", actual.Page.Content);
            TestAssert.True(!actual.Diagnostics.Contains("DOCX_UNSUPPORTED_EQUATION"),
                "Every projected equation fragment must have a matching source count.");

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
                        var label = new XElement(w + (revision == DocxRevisionKind.Deletion ? "delText" : "t"), "QA");
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
                        XElement owner = inlineEquation ? new XElement(w + "p") : display;
                        if (revision is not null)
                        {
                            owner = new XElement(w + (revision == DocxRevisionKind.Insertion ? "ins" : "del"),
                                new XAttribute(w + "id", "412"), new XAttribute(w + "author", "QA Reviewer"));
                            display.Add(owner);
                        }
                        if (packed)
                        {
                            var run = MathRun(leftOperand, lineBreak, label, symbol, drawing, rightOperand);
                            owner.Add(explicitRow ? new XElement(m + "oMath", bookmarkStart, leftNeighbor, run, rightNeighbor, bookmarkEnd) : run);
                        }
                        else
                        {
                            var leftRun = MathRun(leftOperand);
                            var rightRun = MathRun(rightOperand);
                            owner.Add(bookmarkStart,
                                explicitRow ? new XElement(m + "oMath", leftNeighbor, leftRun) : leftRun,
                                new XElement(w + "r", wordProperties, lineBreak, label, symbol, drawing),
                                explicitRow ? new XElement(m + "oMath", rightRun, rightNeighbor) : rightRun,
                                bookmarkEnd);
                        }
                        if (inlineEquation)
                            display.AddAfterSelf(new XElement(w + "r", new XElement(w + "t", "Run before")),
                                owner.Elements(), new XElement(w + "r", new XElement(w + "t", "Run after")));
                        originalPictureRun.Remove();

                        XElement MathRun(params XElement?[] content) => new(m + "r",
                            new XElement(mathProperties), new XElement(wordProperties), content);
                    }, diagnosticSink: Record);

                if (!visible)
                {
                    TestAssert.True(!DocxBlockTraversal.EnumerateBodyParagraphs(source).Any(p => p.Images.Count != 0),
                        "The excluded revision cannot contribute a picture to layout.");
                }
                else
                {
                    CheckInlineOrder(source, packed);
                }
                var page = PictureRangeRenderer(mode).RenderPagesAsync(source, Record, default).ToBlockingEnumerable().Single();
                return (page, diagnostics, source.EquationParts.Sum(part => part.SourceCount));
            }

            void CheckInlineOrder(DocxDocument source, bool packed)
            {
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
                if (revision is not null)
                {
                    TestAssert.Equal(revision, ((DocxInlineEquation)items[first]).Equation.Revision?.Kind);
                    TestAssert.Equal(revision, ((DocxInlineEquation)items[last]).Equation.Revision?.Kind);
                    var pictureRevision = ((DocxInlinePicture)items[picture]).Image.Revisions.Single(r => r.Id == "412");
                    TestAssert.Equal(revision, pictureRevision.Kind);
                }
                if (withBookmark && packed)
                {
                    var owner = ((DocxInlineEquation)items[first]).Equation.Content.Bookmarks.Single();
                    TestAssert.Equal(0, owner.ChildIndex);
                    var anchor = paragraph.BookmarkAnchors.Single(bookmark => bookmark.Name == "EquationDrawing");
                    TestAssert.True(ReferenceEquals(anchor.MathSource, owner.Source),
                        "The original bookmark remains bound to the mathematical source before the drawing.");
                }
                static bool Matches(DocxInlineItem item, string[] labels) => item is DocxInlineEquation math &&
                    math.Equation.Content.Children.Count == labels.Length &&
                    math.Equation.Content.Children.OfType<DocxMathRun>().Select(run => run.Text.Text).SequenceEqual(labels);
            }
        }
    }
}
