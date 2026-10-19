using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxMathFractionTests
{
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
            DocxMathLayout.TryGetLineMetrics(implicitSize, default, out _, out double originalHeight) && largerHeight > originalHeight);
        foreach (double invalid in new[] { 0d, -1d, double.NaN, double.PositiveInfinity })
            TestAssert.True(!DocxMathLayout.TryMeasure(source with { Content = new([fraction with { ControlFontSize = invalid }]) },
                _ => resource, default, out var absent, out _) && absent is null,
                "An invalid metric owner cannot publish a partial fraction.");
    }
}
