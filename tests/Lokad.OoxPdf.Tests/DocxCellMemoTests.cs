using System.Reflection;
using System.Runtime.ExceptionServices;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxCellMemoTests
{
    public static void MemoReturnsHitForSameKey()
    {
        var cell = new DocxTableCell("sample", [], null, null, null, null, [], DocxTableCellMargins.Empty);
        var memo = new DocxLayoutEngine.DocxTableCellTextLinesMemo();
        var line = MakeLine("sample", 10d, 20d, 12d);
        memo.StoreRelativeLines(cell, 80d, null, 36d, 0d, 1d, null, null, DocxCounterDependencies.None, new[] { line }, Array.Empty<DocxInlineImageLayout>(), Array.Empty<DocxInlineTextBoxLayout>(), 10d, 20d, 30d);
        bool hit = memo.TryGetRelativeLines(cell, 80d, null, 36d, 0d, 1d, null, null, DocxCounterDependencies.None, out var relative, out _, out _, out double usedHeight);
        TestAssert.True(hit, "Same key must hit.");
        TestAssert.Equal(1, memo.Hits);
        TestAssert.Equal(0, memo.Misses);
        TestAssert.Equal(30d, usedHeight);
        TestAssert.Equal(1, relative.Count);
        TestAssert.Equal(0d, relative[0].X);
        TestAssert.Equal(0d, relative[0].BaselineY);
        var replayed = DocxLayoutEngine.DocxTableCellTextLinesMemo.ShiftLines(relative, 10d, 20d);
        TestAssert.Equal(10d, replayed[0].X);
        TestAssert.Equal(20d, replayed[0].BaselineY);
        TestAssert.Equal(12d - 10d + 10d, replayed[0].Segments[0].X);
    }

    public static void MemoReplaysAtChangedOrigin()
    {
        var cell = new DocxTableCell("sample", [], null, null, null, null, [], DocxTableCellMargins.Empty);
        var memo = new DocxLayoutEngine.DocxTableCellTextLinesMemo();
        var line = MakeLine("sample", 10d, 20d, 12d);
        memo.StoreRelativeLines(cell, 80d, null, 36d, 0d, 1d, null, null, DocxCounterDependencies.None, new[] { line }, Array.Empty<DocxInlineImageLayout>(), Array.Empty<DocxInlineTextBoxLayout>(), 10d, 20d, 30d);
        bool hit = memo.TryGetRelativeLines(cell, 80d, null, 36d, 0d, 1d, null, null, DocxCounterDependencies.None, out var relative, out _, out _, out double usedHeight);
        TestAssert.True(hit, "Same key must hit for origin replay.");
        TestAssert.Equal(30d, usedHeight);
        var replayed = DocxLayoutEngine.DocxTableCellTextLinesMemo.ShiftLines(relative, 30d, 40d);
        TestAssert.Equal(30d, replayed[0].X);
        TestAssert.Equal(40d, replayed[0].BaselineY);
        TestAssert.Equal(32d, replayed[0].Segments[0].X);
    }

    public static void MemoDistinguishesReferenceIdentity()
    {
        var cell = new DocxTableCell("sample", [], null, null, null, null, [], DocxTableCellMargins.Empty);
        var clone = cell with { };
        TestAssert.True(!ReferenceEquals(cell, clone), "Test requires distinct instances.");
        TestAssert.True(((object)cell).Equals((object)clone), "Test requires equal-valued clones.");
        var memo = new DocxLayoutEngine.DocxTableCellTextLinesMemo();
        var line = MakeLine("sample", 10d, 20d, 12d);
        memo.StoreRelativeLines(cell, 80d, null, 36d, 0d, 1d, null, null, DocxCounterDependencies.None, new[] { line }, Array.Empty<DocxInlineImageLayout>(), Array.Empty<DocxInlineTextBoxLayout>(), 10d, 20d, 30d);
        bool cloneHit = memo.TryGetRelativeLines(clone, 80d, null, 36d, 0d, 1d, null, null, DocxCounterDependencies.None, out _, out _, out _, out _);
        TestAssert.True(!cloneHit, "Equal-valued distinct cell must not hit reference-identity memo.");
        TestAssert.Equal(0, memo.Hits);
        TestAssert.Equal(1, memo.Misses);
        bool sameHit = memo.TryGetRelativeLines(cell, 80d, null, 36d, 0d, 1d, null, null, DocxCounterDependencies.None, out _, out _, out _, out _);
        TestAssert.True(sameHit, "Same reference must hit.");
        TestAssert.Equal(1, memo.Hits);
    }

    public static void MemoTiersPageStaticVsDynamic()
    {
        var cell = new DocxTableCell("sample", [], null, null, null, null, [], DocxTableCellMargins.Empty);
        var line = MakeLine("sample", 10d, 20d, 12d);
        (DocxCounterDependencies Dependencies, (int Number, int Count)[] Hits)[] cases =
        [
            (DocxCounterDependencies.None, [(1, 12), (1, 13), (2, 12), (2, 13)]),
            (DocxCounterDependencies.PageNumber, [(1, 12), (1, 13)]),
            (DocxCounterDependencies.PageCount, [(1, 12), (2, 12)]),
            (DocxCounterDependencies.PageNumber | DocxCounterDependencies.PageCount, [(1, 12)])
        ];
        foreach (var item in cases)
        {
            var memo = new DocxLayoutEngine.DocxTableCellTextLinesMemo();
            memo.StoreRelativeLines(cell, 80d, null, 36d, 0d, 1d, 1, 12, item.Dependencies,
                [line], [], [], 10d, 20d, 30d);
            foreach (int number in new[] { 1, 2 })
            foreach (int count in new[] { 12, 13 })
            {
                TestAssert.Equal(item.Hits.Contains((number, count)), memo.TryGetRelativeLines(cell, 80d, null, 36d, 0d, 1d,
                    number, count, item.Dependencies, out _, out _, out _, out _));
            }
        }
    }

    public static void CellMemoCounterInputsRemainIndependentDuringPlacement()
    {
        foreach (var kind in new[] { DocxFieldKind.Page, DocxFieldKind.NumPages })
        {
            var paragraph = DocxTests.CreateDocxLayoutParagraph("Field", 10d, 10d);
            var run = paragraph.Runs[0] with { Text = DocxFieldText.Placeholder(kind)!, FieldKind = kind };
            paragraph = paragraph with { Runs = [run] };
            var cell = new DocxTableCell(run.Text, [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
            var memo = new DocxLayoutEngine.DocxTableCellTextLinesMemo();
            var measurer = new CountingMeasurer();
            var first = PlaceCounterCell(cell, measurer, memo, 1, 12);
            long calls = measurer.Calls;
            var moved = PlaceCounterCell(cell, measurer, memo, kind == DocxFieldKind.Page ? 1 : 2,
                kind == DocxFieldKind.Page ? 13 : 12, x: 30d, y: 40d, height: 280d);
            TestAssert.True(calls != 0 && measurer.Calls == calls && memo.Hits == 1,
                kind + ": an unrelated page input must reuse measured source lines.");
            TestAssert.Equal(kind == DocxFieldKind.Page ? "1" : "12", moved.Lines.Single().Text);
            TestAssert.True(Math.Abs(moved.Lines[0].X - first.Lines[0].X - 20d) < .000001d &&
                Math.Abs(moved.Lines[0].BaselineY - first.Lines[0].BaselineY - 60d) < .000001d,
                "Memo reuse applies the caller's current origin/height.");
            var changed = PlaceCounterCell(cell, measurer, memo, 2, 13);
            TestAssert.True(measurer.Calls > calls && memo.Misses == 2, "Changing the owned input must remeasure.");
            TestAssert.Equal(kind == DocxFieldKind.Page ? "2" : "13", changed.Lines.Single().Text);
            TestAssert.Equal(kind == DocxFieldKind.Page ? "1" : "12", first.Lines.Single().Text);
            TestAssert.Equal(DocxFieldText.Placeholder(kind), run.Text);
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => PlaceCounterCell(cell, measurer, memo, 2, 13, token: cancellation.Token));
            TestAssert.Equal(1, memo.Hits);
        }
    }

    public static void CellMemoCounterInputsFollowChildBoxOwnership()
    {
        string input = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "InlineBoxMath",
            "inline-box-child-total-counter.docx"));
        var document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
        var paragraph = document.Paragraphs.Single(paragraph => paragraph.InlineTextBoxes.Count != 0);
        var child = DocxMathParagraphPolicy.EnumerateInlineBoxPreparationCandidates(document, default).Single();
        var math = child.Equations.ToDictionary(equation => equation,
            equation => new DocxPreparedMathParagraph(equation, new(18d, 10d, 4d, []), 10d, 16d));
        var cell = new DocxTableCell("", [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var memo = new DocxLayoutEngine.DocxTableCellTextLinesMemo();
        var measurer = new CountingMeasurer();
        var first = PlaceCounterCell(cell, measurer, memo, 1, 12, width: 320d, math: math);
        long calls = measurer.Calls;
        var moved = PlaceCounterCell(cell, measurer, memo, 2, 12, x: 30d, y: 40d, height: 280d, width: 320d, math: math);
        TestAssert.True(calls != 0 && calls == measurer.Calls && memo.Hits == 1,
            "NUMPAGES-only child content can reuse its source frame on another page.");
        var box = moved.Boxes.Single();
        TestAssert.Equal("12", string.Concat(box.TextLines.SelectMany(line => line.Segments)
            .Where(segment => segment.StyleRun.FieldKind == DocxFieldKind.NumPages).Select(segment => segment.Text)));
        TestAssert.True(ReferenceEquals(first.Boxes.Single().TextBox, box.TextBox), "Replayed frames keep the source box owner.");
        TestAssert.Equal(1, box.TextLines.Sum(line => line.InlineMath.Count));
    }

    private static (IReadOnlyList<DocxTextLineLayout> Lines, IReadOnlyList<DocxInlineImageLayout> Images,
        IReadOnlyList<DocxInlineTextBoxLayout> Boxes, double Height) PlaceCounterCell(
        DocxTableCell cell, IDocxTextMeasurer measurer, DocxLayoutEngine.DocxTableCellTextLinesMemo memo,
        int number, int count, double x = 10d, double y = 20d, double height = 240d, double width = 160d,
        IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph>? math = null, CancellationToken token = default)
    {
        var method = typeof(DocxLayoutEngine).GetMethod("LayoutTableCellTextLines", BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException("Expected live cell-layout method.");
        try
        {
            return ((IReadOnlyList<DocxTextLineLayout>, IReadOnlyList<DocxInlineImageLayout>, IReadOnlyList<DocxInlineTextBoxLayout>, double))
                method.Invoke(null, [cell, x, y, width, height, 0d, measurer, 36d, number, count, 1d, memo, number - 1, math, token])!;
        }
        catch (TargetInvocationException exception) when (exception.InnerException is not null)
        {
            ExceptionDispatchInfo.Capture(exception.InnerException).Throw();
            throw;
        }
    }

    public static void MemoKeyIncludesMeasurerReference()
    {
        var cell = new DocxTableCell("sample", [], null, null, null, null, [], DocxTableCellMargins.Empty);
        var memo = new DocxLayoutEngine.DocxTableCellTextLinesMemo();
        var line = MakeLine("sample", 10d, 20d, 12d);
        var firstMeasurer = new CountingMeasurer();
        var secondMeasurer = new CountingMeasurer();
        memo.StoreRelativeLines(cell, 80d, firstMeasurer, 36d, 0d, 1d, null, null, DocxCounterDependencies.None, new[] { line }, Array.Empty<DocxInlineImageLayout>(), Array.Empty<DocxInlineTextBoxLayout>(), 10d, 20d, 30d);
        bool sameHit = memo.TryGetRelativeLines(cell, 80d, firstMeasurer, 36d, 0d, 1d, null, null, DocxCounterDependencies.None, out _, out _, out _, out _);
        TestAssert.True(sameHit, "Same measurer reference must hit.");
        bool otherMiss = memo.TryGetRelativeLines(cell, 80d, secondMeasurer, 36d, 0d, 1d, null, null, DocxCounterDependencies.None, out _, out _, out _, out _);
        TestAssert.True(!otherMiss, "Different measurer reference must miss.");
    }

    private static DocxTextLineLayout MakeLine(string text, double x, double baselineY, double segmentX)
    {
        var run = new DocxTextRun(text, 12d, null, false, false, false, null, null);
        ConstructorInfo segCtor = typeof(DocxTextSegmentLayout).GetConstructors().Single();
        object?[] segArgs = segCtor.GetParameters().Select(p => p.ParameterType.IsValueType ? Activator.CreateInstance(p.ParameterType) : null).ToArray();
        segArgs[0] = text;
        segArgs[1] = run;
        segArgs[2] = segmentX;
        var segment = (DocxTextSegmentLayout)segCtor.Invoke(segArgs);
        ConstructorInfo lineCtor = typeof(DocxTextLineLayout).GetConstructors().Single();
        object?[] lineArgs = lineCtor.GetParameters().Select(p => p.ParameterType.IsValueType ? Activator.CreateInstance(p.ParameterType) : null).ToArray();
        lineArgs[0] = text;
        lineArgs[1] = run;
        lineArgs[3] = x;
        lineArgs[4] = baselineY;
        lineArgs[6] = new[] { segment };
        return (DocxTextLineLayout)lineCtor.Invoke(lineArgs);
    }

    private sealed class CountingMeasurer : IDocxTextMeasurer
    {
        public long Calls;
        public double MeasureText(DocxTextRun? run, string text, double fontSize)
        {
            Calls++;
            return text.Length;
        }
    }
}
