using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxKeepNextOverflowAdmitsTheWholeSourceChainOnce()
    {
        foreach (bool mixed in new[] { false, true })
        foreach (bool leading in new[] { false, true })
        foreach (bool keep in new[] { false, true })
        {
            var document = ReadKeepChain(mixed, keep, leading);
            var layout = LayoutKeepChain(document);
            int first = leading && keep ? 2 : 1;
            TestAssert.Equal(leading ? 3 : 2, layout.Pages.Count);
            foreach (char label in "ABCD")
            {
                int[] pages = KeepParagraphPages(document, layout, label);
                if (!leading || keep)
                    TestAssert.True(pages.SequenceEqual([first + (label - 'A') / 2]),
                        "An oversized source chain starts once and flows; its suffixes cannot each demand a fresh page.");
            }
            var lines = layout.Pages.SelectMany(page => page.Items.OfType<DocxTextLineLayout>()).ToArray();
            foreach (char label in "ABCD")
            foreach (int line in Enumerable.Range(1, 6))
                TestAssert.Equal(1, lines.Count(item => item.Text.StartsWith($"{label}{line}", StringComparison.Ordinal)));
            if (mixed) TestAssert.Equal(4, lines.Sum(line => line.InlineMath.Count));
        }
    }

    public static void DocxKeepChainMembersRetainIndependentKeepLinesAndExactFits()
    {
        foreach (int secondLines in new[] { 6, 8 })
        {
            var document = ReadKeepChain(mixed: true, keep: true, leading: false,
                secondKeepLines: true, secondLines: secondLines);
            var layout = LayoutKeepChain(document);
            TestAssert.True(KeepParagraphPages(document, layout, 'B').SequenceEqual([secondLines == 6 ? 1 : 2]),
                "A chain member still owns keep-lines; a complete frame that exactly fits is admitted.");
            TestAssert.Equal(secondLines == 6 ? 2 : 3, layout.Pages.Count);
        }
    }

    public static void DocxKeepChainsStopAtExplicitPageColumnAndSectionBoundaries()
    {
        foreach (string boundary in new[] { "page", "column", "section" })
        {
            var document = ReadKeepChain(mixed: true, keep: true, leading: false, boundary: boundary);
            var layout = LayoutKeepChain(document);
            TestAssert.Equal(2, layout.Pages.Count);
            TestAssert.True(KeepParagraphPages(document, layout, 'B').SequenceEqual([1]) &&
                KeepParagraphPages(document, layout, 'C').SequenceEqual([2]),
                "Explicit source boundaries terminate one keep chain before the next owner starts.");
            using var cancelled = new CancellationTokenSource();
            cancelled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() =>
                DocxLayoutEngine.KeepChainBlockIndexes(document.BodyElements, 0, cancelled.Token));
        }
    }

    private static DocxDocument ReadKeepChain(bool mixed, bool keep, bool leading,
        bool secondKeepLines = false, int secondLines = 6, string? boundary = null) =>
        ReadMutation("Composition/report", OoxPdfDocxMarkupMode.Final, false, false, (source, w, m) =>
        {
            var drawing = new XElement(source.Descendants(w + "drawing").Single());
            foreach (var extent in drawing.Descendants().Where(element => element.Name.LocalName is "extent" or "ext"))
            { extent.SetAttributeValue("cx", 18 * 12700); extent.SetAttributeValue("cy", 12 * 12700); }
            var equation = new XElement(source.Descendants(m + "oMath").First());
            var body = source.Root!.Element(w + "body")!;
            body.RemoveNodes();
            if (leading) body.Add(Paragraph('Z', 1, false));
            for (int index = 0; index < 4; index++)
            {
                if (index == 2 && boundary is not null)
                {
                    if (boundary == "section") body.Elements(w + "p").Last().Element(w + "pPr")!.Add(
                        Section(nextPage: true));
                    else body.Add(new XElement(w + "p", new XElement(w + "r",
                        new XElement(w + "br", new XAttribute(w + "type", boundary)))));
                }
                body.Add(Paragraph((char)('A' + index), index == 1 ? secondLines : 6, keep && index < 3));
            }
            body.Add(Section(nextPage: false));

            XElement Section(bool nextPage) => new(w + "sectPr",
                nextPage ? new XElement(w + "type", new XAttribute(w + "val", "nextPage")) : null,
                // Break-only paragraphs retain their own mark. Leave room for it;
                // exact frame admission is exercised separately above.
                new XElement(w + "pgSz", new XAttribute(w + "w", 12240), new XAttribute(w + "h", boundary is null ? 6960 : 7680)),
                new XElement(w + "pgMar", new XAttribute(w + "top", 600), new XAttribute(w + "bottom", 600),
                    new XAttribute(w + "left", 600), new XAttribute(w + "right", 600)));

            XElement Paragraph(char label, int count, bool kept)
            {
                var paragraph = new XElement(w + "p", new XElement(w + "pPr",
                    new XElement(w + "spacing", new XAttribute(w + "before", 0), new XAttribute(w + "after", 0),
                        new XAttribute(w + "line", 480), new XAttribute(w + "lineRule", "exact")),
                    new XElement(w + "widowControl", new XAttribute(w + "val", 0)),
                    new XElement(w + "keepNext", new XAttribute(w + "val", kept ? 1 : 0)),
                    label == 'B' && secondKeepLines ? new XElement(w + "keepLines") : null));
                for (int index = 1; index <= count; index++)
                {
                    paragraph.Add(new XElement(w + "r", new XElement(w + "t", $"{label}{index}")));
                    if (mixed && label != 'Z' && index == 3)
                        paragraph.Add(new XElement(equation), new XElement(w + "r", new XElement(drawing)));
                    if (index != count) paragraph.Add(new XElement(w + "r", new XElement(w + "br")));
                }
                return paragraph;
            }
        });

    private static DocxLayout LayoutKeepChain(DocxDocument document)
    {
        var math = new Dictionary<DocxEquation, DocxPreparedMathParagraph>(ReferenceEqualityComparer.Instance);
        foreach (var paragraph in document.Paragraphs)
        foreach (var equation in paragraph.Equations)
            math.Add(equation, new(equation, new DocxMathBox(18d, 10d, 4d, []), 11d, 16d));
        return new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: math);
    }

    private static int[] KeepParagraphPages(DocxDocument document, DocxLayout layout, char label)
    {
        int block = Enumerable.Range(0, document.BodyElements.Count).Single(index =>
            document.BodyElements[index] is DocxParagraphElement paragraph &&
            paragraph.Paragraph.Runs.Any(run => run.Text.StartsWith($"{label}1", StringComparison.Ordinal)));
        return Enumerable.Range(0, layout.Pages.Count).Where(page => layout.Pages[page].Items
            .OfType<DocxTextLineLayout>().Any(line => line.SourceBlockIndex == block)).Select(page => page + 1).ToArray();
    }
}

