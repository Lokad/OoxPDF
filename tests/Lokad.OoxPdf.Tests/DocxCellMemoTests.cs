using System.Reflection;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxCellMemoTests
{
    public static void MemoReturnsHitForSameKey()
    {
        var cell = new DocxTableCell("sample", [], null, null, null, null, [], DocxTableCellMargins.Empty);
        var memo = new DocxLayoutEngine.DocxTableCellTextLinesMemo();
        var line = MakeLine("sample", 10d, 20d, 12d);
        memo.StoreRelativeLines(cell, 80d, null, 36d, 0d, 1d, null, null, true, new[] { line }, 10d, 20d, 30d);
        bool hit = memo.TryGetRelativeLines(cell, 80d, null, 36d, 0d, 1d, null, null, true, out var relative, out double usedHeight);
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
        memo.StoreRelativeLines(cell, 80d, null, 36d, 0d, 1d, null, null, true, new[] { line }, 10d, 20d, 30d);
        bool hit = memo.TryGetRelativeLines(cell, 80d, null, 36d, 0d, 1d, null, null, true, out var relative, out double usedHeight);
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
        memo.StoreRelativeLines(cell, 80d, null, 36d, 0d, 1d, null, null, true, new[] { line }, 10d, 20d, 30d);
        bool cloneHit = memo.TryGetRelativeLines(clone, 80d, null, 36d, 0d, 1d, null, null, true, out _, out _);
        TestAssert.True(!cloneHit, "Equal-valued distinct cell must not hit reference-identity memo.");
        TestAssert.Equal(0, memo.Hits);
        TestAssert.Equal(1, memo.Misses);
        bool sameHit = memo.TryGetRelativeLines(cell, 80d, null, 36d, 0d, 1d, null, null, true, out _, out _);
        TestAssert.True(sameHit, "Same reference must hit.");
        TestAssert.Equal(1, memo.Hits);
    }

    public static void MemoTiersPageStaticVsDynamic()
    {
        var cell = new DocxTableCell("sample", [], null, null, null, null, [], DocxTableCellMargins.Empty);
        var memo = new DocxLayoutEngine.DocxTableCellTextLinesMemo();
        var line = MakeLine("sample", 10d, 20d, 12d);
        memo.StoreRelativeLines(cell, 80d, null, 36d, 0d, 1d, 1, 2, true, new[] { line }, 10d, 20d, 30d);
        bool staticHit = memo.TryGetRelativeLines(cell, 80d, null, 36d, 0d, 1d, 2, 5, true, out _, out _);
        TestAssert.True(staticHit, "Static tier must omit page args.");

        var dynamicMemo = new DocxLayoutEngine.DocxTableCellTextLinesMemo();
        dynamicMemo.StoreRelativeLines(cell, 80d, null, 36d, 0d, 1d, 1, 2, false, new[] { line }, 10d, 20d, 30d);
        bool dynamicMiss = dynamicMemo.TryGetRelativeLines(cell, 80d, null, 36d, 0d, 1d, 2, 2, false, out _, out _);
        TestAssert.True(!dynamicMiss, "Dynamic tier must include page number.");
        bool dynamicHit = dynamicMemo.TryGetRelativeLines(cell, 80d, null, 36d, 0d, 1d, 1, 2, false, out _, out _);
        TestAssert.True(dynamicHit, "Dynamic tier must hit same page args.");
    }

    public static void MemoKeyIncludesMeasurerReference()
    {
        var cell = new DocxTableCell("sample", [], null, null, null, null, [], DocxTableCellMargins.Empty);
        var memo = new DocxLayoutEngine.DocxTableCellTextLinesMemo();
        var line = MakeLine("sample", 10d, 20d, 12d);
        var firstMeasurer = new CountingMeasurer();
        var secondMeasurer = new CountingMeasurer();
        memo.StoreRelativeLines(cell, 80d, firstMeasurer, 36d, 0d, 1d, null, null, true, new[] { line }, 10d, 20d, 30d);
        bool sameHit = memo.TryGetRelativeLines(cell, 80d, firstMeasurer, 36d, 0d, 1d, null, null, true, out _, out _);
        TestAssert.True(sameHit, "Same measurer reference must hit.");
        bool otherMiss = memo.TryGetRelativeLines(cell, 80d, secondMeasurer, 36d, 0d, 1d, null, null, true, out _, out _);
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
