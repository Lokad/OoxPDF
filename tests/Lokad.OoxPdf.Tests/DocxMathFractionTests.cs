using System.Buffers.Binary;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxMathFractionTests
{
    public static void MathFontInlineFractionRetainsParentMetricsAndScaledChildGlyphs()
    {
        var resource = Resource(TestFontBuilder.CreateMathFont(Constants()));
        TestAssert.True(DocxMathLayout.TryMeasure(Fraction() with { Display = false }, _ => resource, CancellationToken.None,
            out DocxMathBox? box, out string? reason), reason ?? "Expected an inline fraction.");
        DocxMathBox measured = TestAssert.NotNull(box);
        TestAssert.True(measured.Glyphs.All(glyph => Math.Abs(glyph.FontSize - 8.52d) < 1e-9),
            "Numerator and denominator glyphs own the script-style font size.");
        TestAssert.True(Math.Abs(measured.Rules[0].Thickness - 40d * 12d / resource.Font.Font.UnitsPerEm) < 1e-9,
            "The fraction axis/rule metrics remain owned by the parent math size, independently of child scaling.");
        TestAssert.True(measured.Glyphs.Take(3).All(glyph => glyph.BaselineY > measured.Rules[0].Y + measured.Rules[0].Thickness) &&
            measured.Glyphs[3].BaselineY < measured.Rules[0].Y, "Scaled child roles remain on opposite sides of the parent rule.");
    }

    public static void MathFontFractionConstantsDriveStoredRuleAndChildBaselines()
    {
        var resource = Resource(TestFontBuilder.CreateMathFont(Constants()));
        DocxEquation equation = Fraction();
        TestAssert.True(DocxMathLayout.TryMeasure(equation, _ => resource, CancellationToken.None, out DocxMathBox? box, out string? reason),
            reason ?? "Expected a display fraction.");
        DocxMathBox measured = TestAssert.NotNull(box);
        TestAssert.Equal(4, measured.Glyphs.Count);
        TestAssert.Equal(1, measured.Rules.Count);
        DocxMathRulePlacement rule = measured.Rules[0];
        TestAssert.True(measured.Glyphs.Take(3).All(glyph => glyph.BaselineY > rule.Y + rule.Thickness),
            "The numerator must retain its own baseline above the fraction rule.");
        TestAssert.True(measured.Glyphs[3].BaselineY < rule.Y,
            "The denominator must retain its own baseline below the fraction rule.");
        TestAssert.True(Math.Abs(rule.Thickness - 40d * 12d / resource.Font.Font.UnitsPerEm) < 1e-9,
            "Rule thickness must come from the bounded font MATH constants.");
        TestAssert.True(Math.Abs(rule.Width - measured.Width) < 1e-9,
            "The stored rule must cover the measured child-box width.");
        var graphics = new PdfGraphicsBuilder();
        DocxMathBoxPainter.Paint(measured, graphics, 20d, 50d, CancellationToken.None);
        TestAssert.Contains(" re f", graphics.ToString());
        TestAssert.True(graphics.UsedFontResourceNames.SetEquals(new[] { "Math1" }), "Fraction painting must use the measured font resource.");
        using var output = new MemoryStream();
        PdfDocumentWriter.WriteBlank(output, [new PdfPage(200d, 150d, graphics.ToString(), [resource])], CancellationToken.None);
        TestAssert.True(output.Length > 0, "Stored fraction geometry must produce a validated PDF.");
    }

    public static void MathFontFractionMissingInvalidOrMixedMetricsKeepWholeBoxFallback()
    {
        DocxEquation equation = Fraction();
        foreach (var font in new[] { TestFontBuilder.LoadTestFont(), OpenTypeFont.Load(TestFontBuilder.CreateMathFont(Constants(-1))) })
        {
            var resource = new PdfFontResource("Math1", PdfEmbeddedFont.Create(font, "A+BC".Select(c => (int)c), CancellationToken.None));
            TestAssert.True(!DocxMathLayout.TryMeasure(equation, _ => resource, CancellationToken.None, out DocxMathBox? box, out _),
                "Absent or invalid MATH metrics cannot publish an approximate fraction box.");
            TestAssert.True(box is null, "A failed fraction cannot retain a partially measured box.");
        }
        var valid = Resource(TestFontBuilder.CreateMathFont(Constants()));
        DocxMathFraction fraction = (DocxMathFraction)equation.Content.Children[0];
        DocxMathRun denominator = (DocxMathRun)fraction.Denominator.Children[0];
        DocxEquation mixed = equation with { Content = new([fraction with
        { Denominator = new([denominator with { Text = denominator.Text with { FontSize = 10d } }]) }]) };
        TestAssert.True(!DocxMathLayout.TryMeasure(mixed, _ => valid, CancellationToken.None, out _, out _),
            "Unowned mixed-size fraction metrics require an explicit fallback.");
    }

    public static void MathFontFractionTallInkConsumesFontLeadingOnce()
    {
        var resource = Resource(TestFontBuilder.CreateMathFont(Constants()));
        TestAssert.True(DocxMathLayout.TryMeasure(Fraction(), _ => resource, CancellationToken.None, out DocxMathBox? box, out _),
            "Expected a measured fraction.");
        DocxMathBox measured = TestAssert.NotNull(box);
        TestAssert.True(DocxMathLayout.TryGetLineMetrics(measured, CancellationToken.None, out double above, out double height),
            "Composed fraction line metrics must be finite.");
        double leading = 120d * 12d / resource.Font.Font.UnitsPerEm;
        TestAssert.True(Math.Abs(above - measured.AboveBaseline - leading) < 1e-9,
            "Tall composed ink consumes one font-leading term, shared by placement and keep measurement.");
        TestAssert.True(height >= above + measured.BelowBaseline, "The line extent must also contain the denominator.");
    }

    public static void MathFontFractionSourceRolesStylesAndCancellationRemainExplicit()
    {
        var resource = Resource(TestFontBuilder.CreateMathFont(Constants()));
        DocxEquation equation = Fraction();
        TestAssert.True(DocxMathLayout.TryGetRuns(equation, CancellationToken.None, out IReadOnlyList<DocxMathRunSource> runs),
            "Display fractions must discover their source runs without flattening numerator/denominator roles.");
        TestAssert.Equal("A+BC", string.Concat(runs.Select(source => source.Run.Text.Text)));
        TestAssert.True(DocxMathLayout.TryGetRuns(equation with { Display = false }, CancellationToken.None, out var inline) &&
            inline.All(source => source.ScriptLevel == 1), "Inline fraction arguments retain their script-style ownership.");
        DocxMathFraction fraction = (DocxMathFraction)equation.Content.Children[0];
        foreach (var unsupported in new[] { fraction with { TypeValue = "skw" }, fraction with { HasUnsupportedProperties = true },
            fraction with { Numerator = new([fraction]) } })
            TestAssert.True(!DocxMathLayout.TryGetRuns(equation with { Content = new([unsupported]) }, CancellationToken.None, out _),
                "Unsupported fraction type, formatting or nested structure remains explicit.");
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryGetRuns(equation, cancelled.Token, out _));
        TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryMeasure(equation, _ => resource, cancelled.Token, out _, out _));
        TestAssert.True(DocxMathLayout.TryMeasure(equation, _ => resource, CancellationToken.None, out DocxMathBox? box, out _), "Expected valid fraction.");
        TestAssert.Throws<OperationCanceledException>(() => DocxMathBoxPainter.Paint(TestAssert.NotNull(box), new PdfGraphicsBuilder(), 0d, 0d, cancelled.Token));
    }

    public static void MathFontFractionUsesCumulativeWorkAndRejectsInvalidSourceControls()
    {
        var resource = Resource(TestFontBuilder.CreateMathFont(Constants()));
        using (OoxConversionBudget.Scope scope = OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxSceneNodesPerConversion = 4 }))
            TestAssert.Throws<OoxPdfLimitExceededException>(() => DocxMathLayout.TryMeasure(Fraction(), _ => resource, CancellationToken.None, out _, out _));
        foreach ((string properties, bool accepted) in new[]
        {
            ("<m:fPr><m:type m:val='bar'/></m:fPr>", true),
            ("<m:fPr><m:ctrlPr/></m:fPr>", false),
            ("<m:fPr><m:type m:val='skw'/></m:fPr>", false)
        })
        {
            string xml = "<m:oMath xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math'><m:f>" + properties +
                "<m:num><m:r><m:rPr><m:sty m:val='p'/></m:rPr><m:t>A</m:t></m:r></m:num>" +
                "<m:den><m:r><m:rPr><m:sty m:val='p'/></m:rPr><m:t>C</m:t></m:r></m:den></m:f></m:oMath>";
            DocxMathRow row = DocxMathParser.Read(System.Xml.Linq.XElement.Parse(xml),
                (_, text) => new(text, 12d, null, false, false, false, null, "TestFont"), CancellationToken.None);
            TestAssert.Equal(accepted, DocxMathLayout.TryGetRuns(Fraction() with { Content = row }, CancellationToken.None, out _));
            TestAssert.True(row.Children[0] is DocxMathFraction, "Fallback must retain the typed fraction source.");
        }
    }

    private static DocxEquation Fraction() => new(new([new DocxMathFraction(Row("A+B"), Row("C"), null)]), 0, 0, true, "center", null, null);
    private static DocxMathRow Row(string text) => new([new DocxMathRun(new(text, 12d, null, false, false, false, null, "TestFont"), "p", null)]);
    private static PdfFontResource Resource(byte[] bytes) => new("Math1", PdfEmbeddedFont.Create(OpenTypeFont.Load(bytes), "A+BC".Select(c => (int)c), CancellationToken.None));
    private static byte[] Constants(short thickness = 40)
    {
        byte[] bytes = new byte[224];
        BinaryPrimitives.WriteUInt16BigEndian(bytes, 1);
        BinaryPrimitives.WriteUInt16BigEndian(bytes.AsSpan(4), 10);
        BinaryPrimitives.WriteInt16BigEndian(bytes.AsSpan(10), 73);
        Set(OpenTypeMathConstant.AxisHeight, 250);
        Set(OpenTypeMathConstant.MathLeading, 120);
        Set(OpenTypeMathConstant.FractionRuleThickness, thickness);
        Set(OpenTypeMathConstant.FractionNumeratorDisplayStyleShiftUp, 700);
        Set(OpenTypeMathConstant.FractionDenominatorDisplayStyleShiftDown, 500);
        Set(OpenTypeMathConstant.FractionNumDisplayStyleGapMin, 100);
        Set(OpenTypeMathConstant.FractionDenomDisplayStyleGapMin, 100);
        Set(OpenTypeMathConstant.FractionNumeratorShiftUp, 500);
        Set(OpenTypeMathConstant.FractionDenominatorShiftDown, 400);
        Set(OpenTypeMathConstant.FractionNumeratorGapMin, 100);
        Set(OpenTypeMathConstant.FractionDenominatorGapMin, 100);
        return bytes;
        void Set(OpenTypeMathConstant field, short value) => BinaryPrimitives.WriteInt16BigEndian(bytes.AsSpan(18 + (int)field * 4), value);
    }
}
