using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionAutomaticSpacingFollowsTheCompletedRow()
    {
        var document = ReadMutation("CompositionProbes/math-minimum-line-72", OoxPdfDocxMarkupMode.Final, false, false);
        var mixed = document.Paragraphs.First(p => p.Equations.Count != 0);
        var display = document.Paragraphs.First(p => p.Equations.Any(e => e.Display));
        foreach (double scale in new[] { 1d, .75d })
        foreach (string kind in new[] { "text", "math", "picture", "picture-math", "display" })
        {
            var single = Measure(1d);
            var expanded = Measure(1.15d);
            TestAssert.True(Math.Abs(expanded.OccupiedHeight - single.OccupiedHeight - 2.16d * scale) < 1e-9,
                "The ordinary 14.4pt line adds 2.16pt at115%, after the complete row, including a picture-only row.");
            TestAssert.Equal(single.BaselineInset, expanded.BaselineInset);
            TestAssert.Equal(single.Advance, expanded.Advance);
            TestAssert.True(single.Fragments.Select(f => f.Source).SequenceEqual(expanded.Fragments.Select(f => f.Source)),
                "Spacing preserves source order and horizontal geometry.");

            DocxMeasuredInlineLine Measure(double factor)
            {
                var source = kind == "display" ? display : mixed;
                bool math = kind is "math" or "picture-math" or "display";
                bool picture = kind is "picture" or "picture-math";
                var image = new DocxInlineImage(18d, 72d, "image/png", null, "frame.png") { SourceRunIndex = 0 };
                var items = source.InlineContent!.Where(item => math || item is not DocxInlineEquation).ToList();
                if (kind == "picture") items.Clear();
                if (picture) items.Add(new DocxInlinePicture(image));
                var paragraph = source with { LineSpacingPoints = null, LineSpacingFactor = factor,
                    Equations = math ? source.Equations : [], Images = picture ? [image] : [], InlineContent = items,
                    Runs = kind == "picture" ? [] : source.Runs,
                    Spacing = source.Spacing with { LineRuleValue = "auto", LineValue = factor == 1d ? "240" : "276" } };
                var equations = new Dictionary<DocxEquation, DocxPreparedMathParagraph>(ReferenceEqualityComparer.Instance);
                if (math)
                {
                    var equation = paragraph.Equations.Single();
                    equations[equation] = new(equation, new(18d, 20d, 20d, []), 20d,
                        40d + (kind == "display" ? (factor - 1d) * 14.4d : 0d)) { UsesDisplayStyle = kind == "display" };
                }
                TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineAtoms(paragraph, equations, new TrackingMeasurer(),
                    12d, 9d * scale, 14.4d * factor * scale, scale, default, out var atoms, allowStaticFallback: kind == "text"),
                    $"Supported {kind} source content shares one ordered measurement path at scale {scale}.");
                TestAssert.True(DocxLayoutEngine.TryWrapOrderedInlineAtoms(TestAssert.NotNull(atoms), 600d, 600d,
                    default, out var lines), "The complete row fits without changing its content.");
                return TestAssert.NotNull(lines).Single();
            }
        }
    }

    public static void DocxInlineCompositionAutomaticSpacingSurvivesResumedWrapping()
    {
        var document = ReadMutation("CompositionProbes/math-minimum-line-72", OoxPdfDocxMarkupMode.Final, false, false);
        var source = document.Paragraphs.First(p => p.Equations.Count != 0);
        var paragraph = source with { LineSpacingPoints = null, LineSpacingFactor = 1.15d,
            Spacing = source.Spacing with { LineRuleValue = "auto", LineValue = "276" } };
        var equation = paragraph.Equations.Single();
        var prepared = new DocxPreparedMathParagraph(equation, new(18d, 20d, 20d, []), 20d, 40d) { UsesDisplayStyle = false };
        var equations = new Dictionary<DocxEquation, DocxPreparedMathParagraph>(ReferenceEqualityComparer.Instance) { [equation] = prepared };
        TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineAtoms(paragraph, equations, new TrackingMeasurer(),
            12d, 9d, 16.56d, 1d, default, out var result), "Measure the paragraph once before placement.");
        var atoms = TestAssert.NotNull(result);
        TestAssert.True(DocxLayoutEngine.TryWrapOrderedInlineAtoms(atoms, 80d, 80d, default, out var wrapped),
            "Mixed content supports ordinary soft wrapping.");
        var expected = TestAssert.NotNull(wrapped);
        TestAssert.True(expected.Count > 1 && expected.Any(l => l.Fragments.Any(f => f.Math is not null)),
            "The control must exercise both wrapping and the occupied math row.");
        var cursor = default(DocxInlineSourceCursor);
        foreach (var line in expected)
        {
            TestAssert.True(DocxLayoutEngine.TryGetNextOrderedInlineLine(atoms, cursor, 80d, 80d, default, out var next),
                "Placement can resume from the last admitted source position.");
            var actual = TestAssert.NotNull(next);
            TestAssert.Equal(line.OccupiedHeight, actual.OccupiedHeight);
            TestAssert.Equal(line.BaselineInset, actual.BaselineInset);
            TestAssert.Equal(line.EndCursor, actual.EndCursor);
            if (actual.Fragments.Any(f => f.Math is not null))
                TestAssert.True(Math.Abs(actual.OccupiedHeight - 42.16d) < 1e-9,
                    "A resumed tall row consumes its ordinary spacing exactly once.");
            cursor = actual.EndCursor;
        }
        TestAssert.True(cursor.IsComplete(atoms), "Every source item is consumed once.");
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxLayoutEngine.TryGetNextOrderedInlineLine(
            atoms, default, 80d, 80d, cancelled.Token, out _));
    }
}
