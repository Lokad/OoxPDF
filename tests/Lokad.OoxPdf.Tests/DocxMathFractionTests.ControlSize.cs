using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxMathFractionTests
{
    public static void MathFontFractionAutomaticSpacingUsesNominalControlAcrossChildSizes()
    {
        var resource = Resource(TestFontBuilder.CreateMathFont(Constants()));
        var font = resource.Font.Font;
        var source = Fraction();
        var fraction = (DocxMathFraction)source.Content.Children.Single();
        double? firstGrowth = null;
        foreach (double ownerSize in new[] { 12d, 24d })
        {
            double? ownerGrowth = null;
            foreach (double childSize in new[] { 9d, 18d, 30d })
            {
                DocxMathRow Resize(DocxMathRow row) => row with { Children = row.Children.Select(node =>
                    node is DocxMathRun run ? run with { Text = run.Text with { FontSize = childSize } } : node).ToArray() };
                var equation = source with { ParentFontSize = ownerSize, Content = new([fraction with
                    { Numerator = Resize(fraction.Numerator), Denominator = Resize(fraction.Denominator) }]) };
                TestAssert.True(DocxMathLayout.TryMeasure(equation, _ => resource, default, out var measured, out var reason,
                    resolveControlMetrics: _ => font), reason ?? "Expected a fraction with independent child sizes.");
                var box = TestAssert.NotNull(measured);
                TestAssert.True(DocxMathLayout.TryGetLineMetrics(box, default, out double singleAbove, out double singleHeight),
                    "The ordinary slot retains the complete math box.");
                TestAssert.True(DocxMathLayout.TryGetLineMetrics(box, default, out double expandedAbove, out double expandedHeight, 1.15d),
                    "Both automatic slots preserve the same complete math box.");
                TestAssert.Equal(singleAbove, expandedAbove);
                double growth = expandedHeight - singleHeight;
                TestAssert.True(growth > 0d, "Expanded spacing adds occupied height without moving the intrinsic baseline.");
                if (ownerGrowth is { } expected)
                    TestAssert.True(Math.Abs(growth - expected) < 1e-9,
                        "Changing painted child sizes must not change nominal automatic spacing.");
                ownerGrowth = growth;
            }
            if (firstGrowth is { } smaller)
                TestAssert.True(Math.Abs(ownerGrowth!.Value - smaller * 2d) < 1e-9,
                    "Doubling the nominal owner doubles ordinary line expansion, independently of tall ink.");
            firstGrowth = ownerGrowth;
        }
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryGetLineMetrics(
            new(1d, 1d, 1d, [new(65, 65, resource, "0001", 12d, 0d, 0d, 1d, null)]), cancelled.Token, out _, out _, 1.15d));
    }

    public static void MathFontFractionControlSizeOwnsGeometryWithoutRescalingChildren()
    {
        var resource = Resource(TestFontBuilder.CreateMathFont(Constants()));
        var source = Fraction() with { Display = false, ParentFontSize = 12d };
        var fraction = (DocxMathFraction)source.Content.Children.Single();
        DocxMathBox Measure(double? size)
        {
            var equation = source with { Content = new([fraction with { ControlFontSize = size }]) };
            TestAssert.True(DocxMathLayout.TryMeasure(equation, _ => resource, default, out var box, out var reason), reason ?? "Expected a supported controlled fraction.");
            return TestAssert.NotNull(box);
        }
        var implicitSize = Measure(null);
        var neutral = Measure(12d);
        TestAssert.True(implicitSize.Glyphs.SequenceEqual(neutral.Glyphs) && implicitSize.Rules.SequenceEqual(neutral.Rules) &&
            implicitSize.ParentMetrics == neutral.ParentMetrics && implicitSize.AboveBaseline == neutral.AboveBaseline &&
            implicitSize.BelowBaseline == neutral.BelowBaseline,
            "An explicit neutral declaration preserves the enclosing metric inheritance.");
        var larger = Measure(24d);
        TestAssert.True(larger.Glyphs.Select(g => g.FontSize).SequenceEqual(implicitSize.Glyphs.Select(g => g.FontSize)) &&
            larger.Rules.Single().Thickness == implicitSize.Rules.Single().Thickness * 2d && larger.ParentMetrics?.Size == 24d,
            "The fraction rule and line owner use the authored control; child glyph sizes remain independent.");
        TestAssert.True(DocxMathLayout.TryGetLineMetrics(larger, default, out _, out double largerHeight) &&
            DocxMathLayout.TryGetLineMetrics(implicitSize, default, out _, out double originalHeight) && largerHeight > originalHeight,
            "The authored control's larger frame owns sufficient line leading.");
        foreach (double invalid in new[] { 0d, -1d, double.NaN, double.PositiveInfinity })
            TestAssert.True(!DocxMathLayout.TryMeasure(source with { Content = new([fraction with { ControlFontSize = invalid }]) },
                _ => resource, default, out var absent, out _) && absent is null,
                "An invalid metric owner cannot publish a partial fraction.");
    }
}
