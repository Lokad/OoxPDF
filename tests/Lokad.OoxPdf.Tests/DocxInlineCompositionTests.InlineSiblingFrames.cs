using System.Reflection;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineSiblingFramesPreserveIndependentWordContent()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native sibling-frame content uses installed Windows fonts.");
        CheckInlineBoxMath(null, "inline-sibling-frames-office.json");
    }

    public static void DocxInlineSiblingFrameOriginsMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent sibling-frame origins use installed Windows fonts.");
        CheckInlineBoxWordOrigins("inline-sibling-frames-office.json");
    }

    public static void DocxInlineSiblingSourceItemsRetainBoxOrderAndOwners()
    {
        using var gold = ReadInlineBoxMathGold("inline-sibling-frames-office.json");
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray()
            .Where(row => row.GetProperty("Mode").GetString() == "final-preserve"))
        {
            var document = ReadMutation("InlineBoxMath/" + row.GetProperty("Id").GetString(), OoxPdfDocxMarkupMode.Final,
                false, false, partOverride: row.GetProperty("MathPart").GetString());
            var sources = DocxBlockTraversal.EnumerateSourceParagraphs(document)
                .Where(paragraph => paragraph.Equations.Count != 0 && paragraph.InlineTextBoxes.Count != 0).ToArray();
            TestAssert.True(sources.Length != 0, "At least one source story owns the sibling frame.");
            foreach (var source in sources)
            {
                var items = TestAssert.NotNull(source.InlineContent);
                TestAssert.True(ReferenceEquals(source.InlineTextBoxes.Single(), items.OfType<DocxInlineBox>().Single().Box),
                    "The ordered frame retains its actual parsed owner.");
                TestAssert.Equal(source.Equations.Count, items.OfType<DocxInlineEquation>().Count());
                TestAssert.Equal(source.Images.Count, items.OfType<DocxInlinePicture>().Count());
                TestAssert.Equal(source.Equations.Count, DocxMathParagraphPolicy.GetPreparationEquations(source).Count);
                if (row.GetProperty("Id").GetString() == "inline-sibling-first") TestAssert.True(items[0] is DocxInlineBox,
                    "A box before the first text run retains its source position.");
                var replaced = source with { InlineTextBoxes = source.InlineTextBoxes.ToArray() };
                TestAssert.True(replaced.InlineContent is null, "Replacing box owners invalidates the derived source view.");
                TestAssert.Equal(0, DocxMathParagraphPolicy.GetPreparationEquations(replaced).Count);
            }
        }
    }

    public static void DocxInlineSiblingFramesReuseMeasuredChildrenDuringWrapping()
    {
        var source = ReadSiblingFrameSource();
        var prepared = PrepareSiblingFrameMath(source);
        var measurer = new SiblingFrameMeasurer();
        TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineAtoms(source, prepared, measurer,
            11d, 10d, 16d, 1d, default, out var atoms), "The complete source-owned frame measures.");
        var measured = TestAssert.NotNull(atoms);
        var frame = TestAssert.NotNull(measured.Fragments.Single(fragment => fragment.TextBox is not null).TextBox);
        TestAssert.True(frame.BoxHeight >= 60d && frame.BoxWidth == 216d, "Authored frame bounds survive measurement.");
        int childCalls = measurer.ChildCalls;
        TestAssert.True(childCalls > 0, "The child story was measured.");
        foreach (double width in new[] { 450d, 300d, 220d })
        {
            TestAssert.True(DocxLayoutEngine.TryWrapOrderedInlineAtoms(measured, width, width, default, out var lines),
                "A complete box fits or moves atomically to the next row.");
            var fragments = TestAssert.NotNull(lines).SelectMany(line => line.Fragments).ToArray();
            TestAssert.Equal(1, fragments.Count(fragment => ReferenceEquals(fragment.TextBox, frame)));
            TestAssert.Equal(1, fragments.Count(fragment => fragment.Math is not null));
        }
        TestAssert.Equal(childCalls, measurer.ChildCalls);
    }

    public static void DocxInlineSiblingSourceOffsetsOrderMixedDrawingsWithinOneRun()
    {
        var document = ReadMutation("InlineBoxMath/inline-sibling-mixed", OoxPdfDocxMarkupMode.Final, false, false,
            (xml, w, math) =>
            {
                XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
                var paragraph = xml.Descendants(w + "p").Single(p => p.Elements(math + "oMath").Any());
                var box = paragraph.Elements(w + "r").Single(run => run.Descendants(w + "txbxContent").Any());
                var picture = paragraph.Elements(w + "r").Single(run => run.Descendants(a + "blip").Any());
                picture.Add(new XElement(w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), " A "),
                    new XElement(box.Element(w + "drawing")!),
                    new XElement(w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), " B "));
                box.Remove();
            });
        var source = document.Paragraphs.Single(p => p.InlineTextBoxes.Count != 0);
        int owner = source.InlineTextBoxes.Single().SourceRunIndex;
        var items = TestAssert.NotNull(source.InlineContent).Where(item => item switch
        {
            DocxInlineText text => text.Run.SourceRunIndex == owner,
            DocxInlinePicture picture => picture.Image.SourceRunIndex == owner,
            DocxInlineBox box => box.Box.SourceRunIndex == owner,
            _ => false
        }).ToArray();
        TestAssert.Equal(4, items.Length);
        TestAssert.True(items[0] is DocxInlinePicture && items[1] is DocxInlineText &&
            items[2] is DocxInlineBox && items[3] is DocxInlineText, "Offset and drawing order preserve image/text/box/text within one run.");
        TestAssert.Equal(3, source.InlineTextBoxes.Single().SourceTextOffsetInRun);
        var bookmarked = ReadMutation("InlineBoxMath/inline-sibling-first", OoxPdfDocxMarkupMode.Final, false, false,
            (xml, w, math) =>
            {
                var paragraph = xml.Descendants(w + "p").Single(p => p.Elements(math + "oMath").Any());
                var box = paragraph.Elements(w + "r").Single(run => run.Descendants(w + "txbxContent").Any());
                box.AddBeforeSelf(new XElement(w + "bookmarkStart", new XAttribute(w + "id", "902"), new XAttribute(w + "name", "PublicBox")));
                box.AddAfterSelf(new XElement(w + "bookmarkEnd", new XAttribute(w + "id", "902")));
            }).Paragraphs.Single(p => p.InlineTextBoxes.Count != 0);
        TestAssert.Equal(0, bookmarked.BookmarkAnchors.Single(anchor => anchor.Name == "PublicBox").InlineItemIndex ?? -1);
        TestAssert.Equal(1, DocxMathParagraphPolicy.GetPreparationEquations(bookmarked).Count);
    }

    public static void DocxInlineSiblingFrameMeasurementObservesChildCancellationAndNestedFallback()
    {
        var source = ReadSiblingFrameSource();
        var prepared = PrepareSiblingFrameMath(source);
        using var cancellation = new CancellationTokenSource();
        TestAssert.Throws<OperationCanceledException>(() => DocxLayoutEngine.TryMeasureOrderedInlineAtoms(source,
            prepared, new SiblingFrameMeasurer(cancellation), 11d, 10d, 16d, 1d, cancellation.Token, out _));
        var original = source.InlineTextBoxes.Single();
        var nested = original with { BodyElements = [new DocxParagraphElement(source)] };
        var items = TestAssert.NotNull(source.InlineContent);
        var unsupported = source with { InlineTextBoxes = [nested], InlineContent = items.Select(item =>
            item is DocxInlineBox ? (DocxInlineItem)new DocxInlineBox(nested) : item).ToArray() };
        TestAssert.True(!DocxLayoutEngine.TryMeasureOrderedInlineAtoms(unsupported, prepared, new SiblingFrameMeasurer(),
            11d, 10d, 16d, 1d, default, out _), "Nested frames retain the explicit existing fallback.");
    }

    public static void DocxInlineSiblingCellMemoRelocatesCompleteFramesWithoutRemeasuringChildren()
    {
        var source = ReadSiblingFrameSource();
        var report = ReadMutation("InlineBoxMath/inline-box-mixed-source-report", OoxPdfDocxMarkupMode.Final, false, false);
        var box = source.InlineTextBoxes.Single() with
        { BodyElements = report.Paragraphs.Single(p => p.InlineTextBoxes.Count != 0).InlineTextBoxes.Single().BodyElements };
        var items = TestAssert.NotNull(source.InlineContent);
        source = source with { InlineTextBoxes = [box], InlineContent = items.Select(item =>
            item is DocxInlineBox ? (DocxInlineItem)new DocxInlineBox(box) : item).ToArray() };
        var prepared = PrepareSiblingFrameMath(source);
        var cell = new DocxTableCell(string.Empty, [source], null, null, null, null, [], DocxTableCellMargins.Empty);
        var measurer = new SiblingFrameMeasurer();
        var memo = new DocxLayoutEngine.DocxTableCellTextLinesMemo();
        var method = typeof(DocxLayoutEngine).GetMethod("LayoutTableCellTextLines", BindingFlags.NonPublic | BindingFlags.Static)!;
        (IReadOnlyList<DocxTextLineLayout> Lines, IReadOnlyList<DocxInlineImageLayout> Images,
            IReadOnlyList<DocxInlineTextBoxLayout> Boxes, double Height) Place(double x, double y, double height, int page,
                DocxLayoutEngine.DocxTableCellTextLinesMemo? cache) =>
            ((IReadOnlyList<DocxTextLineLayout>, IReadOnlyList<DocxInlineImageLayout>, IReadOnlyList<DocxInlineTextBoxLayout>, double))
                method.Invoke(null, [cell, x, y, 450d, height, 0d, measurer, 36d, page, 2, 1d, cache, page, prepared, CancellationToken.None])!;
        var first = Place(10d, 20d, 200d, 1, memo);
        int childCalls = measurer.Calls;
        var replay = Place(40d, 80d, 260d, 2, memo);
        TestAssert.Equal(childCalls, measurer.Calls);
        var fresh = Place(40d, 80d, 260d, 2, null);
        TestAssert.Equal(1, memo.Hits);
        TestAssert.Equal(1, memo.Misses);
        var firstBox = first.Boxes.Single(); var replayBox = replay.Boxes.Single(); var freshBox = fresh.Boxes.Single();
        TestAssert.Equal(firstBox.BoxX + 30d, replayBox.BoxX);
        TestAssert.Equal(firstBox.BoxTop + 120d, replayBox.BoxTop);
        TestAssert.Equal(freshBox.BoxX, replayBox.BoxX);
        TestAssert.Equal(freshBox.BoxTop, replayBox.BoxTop);
        TestAssert.Equal(freshBox.TextLines.Count, replayBox.TextLines.Count);
        for (int index = 0; index < freshBox.TextLines.Count; index++)
        {
            TestAssert.Equal(freshBox.TextLines[index].BaselineY, replayBox.TextLines[index].BaselineY);
            TestAssert.Equal(freshBox.TextLines[index].X, replayBox.TextLines[index].X);
        }
        TestAssert.Equal(2, replayBox.InlineImages.Single().PageIndex);
        TestAssert.Equal(freshBox.InlineImages.Single().X, replayBox.InlineImages.Single().X);
        TestAssert.Equal(freshBox.InlineImages.Single().Y, replayBox.InlineImages.Single().Y);
    }

    private static DocxParagraph ReadSiblingFrameSource() => ReadMutation("InlineBoxMath/inline-box-body-equation-inline-box-valid",
        OoxPdfDocxMarkupMode.Final, false, false).Paragraphs.Single(paragraph => paragraph.InlineTextBoxes.Count != 0);

    private static Dictionary<DocxEquation, DocxPreparedMathParagraph> PrepareSiblingFrameMath(DocxParagraph source)
    {
        var prepared = new Dictionary<DocxEquation, DocxPreparedMathParagraph>(ReferenceEqualityComparer.Instance);
        foreach (var equation in source.Equations.Concat(source.InlineTextBoxes.SelectMany(box =>
            DocxBlockTraversal.EnumerateBodyParagraphs(box.BodyElements)).SelectMany(paragraph => paragraph.Equations)))
            prepared.Add(equation, new(equation, new(18d, 8d, 2d, []), 10d, 16d) { UsesDisplayStyle = false });
        return prepared;
    }

    private sealed class SiblingFrameMeasurer(CancellationTokenSource? cancellation = null) : IDocxTextMeasurer
    {
        public int ChildCalls { get; private set; }
        public int Calls { get; private set; }
        public double MeasureText(DocxTextRun? run, string text, double fontSize)
        {
            Calls++;
            if (text.Contains("content", StringComparison.Ordinal) || text.StartsWith("Inline", StringComparison.Ordinal))
            { ChildCalls++; cancellation?.Cancel(); }
            return text.Length * 5d;
        }
    }
}
