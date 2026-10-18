using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionMinimumLineSlotOwnsLeadingAndRetainsSpacingFallbacks()
    {
        var document = ReadMutation("CompositionProbes/math-minimum-line-72", OoxPdfDocxMarkupMode.Final, false, false);
        var source = document.Paragraphs.First(p => p.Equations.Count != 0);
        var equation = source.Equations.Single();
        var prepared = new DocxPreparedMathParagraph(equation, new(18d, 10d, 4d, []), 12d, 16d) { UsesDisplayStyle = false };
        var equations = new Dictionary<DocxEquation, DocxPreparedMathParagraph>(ReferenceEqualityComparer.Instance) { [equation] = prepared };
        foreach (double scale in new[] { 1d, .75d })
        {
            var natural = Measure(0d);
            var minimum = Measure(60d);
            double extra = 60d * scale - natural.OccupiedHeight;
            TestAssert.True(extra > 0d && Math.Abs(minimum.OccupiedHeight - 60d * scale) < 1e-9,
                "The completed line consumes its authored minimum once in physical coordinates.");
            TestAssert.True(Math.Abs(minimum.BaselineInset - natural.BaselineInset - extra) < 1e-9 &&
                Math.Abs(minimum.OccupiedHeight - minimum.BaselineInset - natural.OccupiedHeight + natural.BaselineInset) < 1e-9,
                "Only empty leading before the baseline grows; intrinsic descent remains unchanged.");
            TestAssert.Equal(natural.Advance, minimum.Advance);
            TestAssert.True(minimum.Fragments.Where(f => f.Math is not null).All(f => ReferenceEquals(f.Math, prepared) &&
                f.Source is DocxInlineEquation item && ReferenceEquals(item.Equation, equation)),
                "Minimum placement retains the original measured and source owners.");

            DocxMeasuredInlineLine Measure(double minimumHeight)
            {
                var paragraph = source with { LineSpacingPoints = minimumHeight,
                    Spacing = source.Spacing with { LineValue = (minimumHeight * 20d).ToString(System.Globalization.CultureInfo.InvariantCulture) } };
                TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineAtoms(paragraph, equations, new TrackingMeasurer(),
                    12d, 9d * scale, 60d * scale, scale, default, out var atoms), "A valid minimum supports the shared ordered path.");
                TestAssert.True(DocxLayoutEngine.TryWrapOrderedInlineAtoms(TestAssert.NotNull(atoms), 600d, 600d, default, out var lines),
                    "The complete source row remains measurable.");
                return TestAssert.NotNull(lines).Single();
            }
        }
        foreach (var invalid in new[] {
            source with { Spacing = source.Spacing with { LineRuleValue = "future" } },
            source with { LineSpacingPoints = null },
            source with { LineSpacingPoints = -1d },
            source with { LineSpacingPoints = double.NaN },
            source with { LineSpacingPoints = double.PositiveInfinity } })
            TestAssert.Equal(0, DocxMathParagraphPolicy.GetPreparationEquations(invalid).Count);
        var display = document.Paragraphs.First(p => p.Equations.Any(e => e.Display));
        var compressed = display with { LineSpacingPoints = null, LineSpacingFactor = .5d,
            Spacing = display.Spacing with { LineRuleValue = "auto", LineValue = "120" } };
        TestAssert.True(!DocxMathParagraphPolicy.TryGetDisplayLineSpacingFactor(compressed, out _),
            "Unqualified compressed automatic display slots retain their existing fallback.");
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxLayoutEngine.TryMeasureOrderedInlineAtoms(source,
            equations, new TrackingMeasurer(), 12d, 9d, 72d, 1d, cancelled.Token, out _));
        TestAssert.Equal(72d, source.LineSpacingPoints);
    }
}
