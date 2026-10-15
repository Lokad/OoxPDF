using System.Buffers.Binary;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxMathFractionTests
{
    public static void MathFontStructuredSequencesRetainIndependentChildrenAndRules()
    {
        var resource = Resource(TestFontBuilder.CreateMathFont(Constants()));
        DocxEquation child = Fraction();
        var fraction = (DocxMathFraction)child.Content.Children[0];
        var scripts = new DocxMathScripts(Row("B"), null, Row("C"));
        var source = child with { Content = new([Row("A+"), fraction, Row("+"), scripts, Row("+"), fraction, Row("+B")]) };
        TestAssert.True(DocxMathLayout.TryMeasure(source, _ => resource, CancellationToken.None, out var box, out var reason),
            reason ?? "A supported sequence must not inherit the former single-structure guard.");
        DocxMathBox measured = TestAssert.NotNull(box);
        TestAssert.Equal("A+A+BC+BC+A+BC+B", string.Concat(measured.Glyphs.Select(g => char.ConvertFromUtf32(g.SourceCodePoint))));
        TestAssert.Equal(2, measured.Rules.Count);
        TestAssert.True(measured.Rules[0].X + measured.Rules[0].Width < measured.Rules[1].X,
            "Independent child rules retain distinct source positions in the sequence.");
        TestAssert.True(DocxMathLayout.TryMeasure(child with { Content = new([scripts]) }, _ => resource,
            CancellationToken.None, out var scriptBox, out _), "Expected the independently measured script.");
        TestAssert.True(measured.Glyphs.Single(g => g.SourceCodePoint == 'C' && g.FontSize < 12d).BaselineY ==
            TestAssert.NotNull(scriptBox).Glyphs[1].BaselineY,
            "The script role retains its independently measured style/baseline, including zero-shift synthetic metrics.");
        var graphics = new PdfGraphicsBuilder();
        DocxMathBoxPainter.Paint(measured, graphics, 20d, 50d, CancellationToken.None);
        using var output = new MemoryStream();
        PdfDocumentWriter.WriteBlank(output, [new PdfPage(300d, 150d, graphics.ToString(), [resource])], CancellationToken.None);
        TestAssert.True(output.Length > 0, "All composed glyphs/rules retain valid owned PDF resources.");
    }

    public static void MathFontStructuredSequencesConsumeParentLeadingOnceForNormalChildren()
    {
        OpenTypeFont parent = OpenTypeFont.Load(TestFontBuilder.CreateMathFont(Constants()));
        var text = new PdfFontResource("Text1", PdfEmbeddedFont.Create(TestFontBuilder.LoadTestFont(),
            "A+BC".Select(c => (int)c), CancellationToken.None));
        DocxMathRow Normal(string value) => new([new DocxMathRun(new(value, 12d, null, false, false, false, null, "TestFont"),
            null, null) { NormalTextValue = true }]);
        var fraction = new DocxMathFraction(Normal("A+B"), Normal("C"), "bar");
        var one = new DocxEquation(new([fraction]), 0, 0, true, null, null, null) { ParentFontSize = 12d };
        var source = one with { Content = new([fraction, Normal("+"), fraction]) };
        TestAssert.True(DocxMathLayout.TryMeasure(one, _ => text, CancellationToken.None, out var child, out _, parent), "Expected existing normal-child fraction.");
        TestAssert.True(DocxMathLayout.TryMeasure(source, _ => text, CancellationToken.None, out var box, out _, parent), "Expected normal-child sequence.");
        TestAssert.True(DocxMathLayout.TryGetLineMetrics(TestAssert.NotNull(child), CancellationToken.None, out double expectedAbove, out double expectedHeight) &&
            DocxMathLayout.TryGetLineMetrics(TestAssert.NotNull(box), CancellationToken.None, out double actualAbove, out double actualHeight) &&
            Math.Abs(actualAbove - expectedAbove) < 1e-9 && Math.Abs(actualHeight - expectedHeight) < 1e-9,
            "Side-by-side boxes consume one parent-leading/occupied line frame, not one term per fraction.");
        TestAssert.True(box!.ParentMetrics is { } owner && ReferenceEquals(owner.Font, parent) && box.Glyphs.All(g => g.Font == text),
            "The hidden parent metric owner survives independently of actual text-face glyph resources.");
    }

    public static void MathFontScriptRowsUseEquationSpacingWithoutChangingChildMetrics()
    {
        var resource = Resource(TestFontBuilder.CreateMathFont(Constants()));
        DocxMathRow Sized(string text) => new([new DocxMathRun(new(text, 20d, null, false, false, false, null, "TestFont"), "p", null)]);
        var scripts = new DocxMathScripts(Sized("B"), null, Sized("C"));
        DocxEquation child = Fraction() with { Content = new([scripts]), ParentFontSize = 12d };
        DocxMathRow prefix = Row("A+");
        TestAssert.True(DocxMathRowLayout.TryMeasure(prefix, _ => resource, CancellationToken.None, out var first, out _), "Expected literal prefix.");
        TestAssert.True(DocxMathLayout.TryMeasure(child, _ => resource, CancellationToken.None, out var basis, out _), "Expected authored20pt child.");
        TestAssert.True(DocxMathLayout.TryMeasure(child with { Content = new([prefix, scripts]) }, _ => resource,
            CancellationToken.None, out var box, out _), "Expected mixed-size structured row.");
        DocxMathBox measured = TestAssert.NotNull(box), standalone = TestAssert.NotNull(basis);
        TestAssert.True(Math.Abs(measured.Glyphs[2].X - TestAssert.NotNull(first).AdvanceWidth - 8d * 12d / 18d) < 1e-9,
            "Both sides of the binary boundary use the equation's12pt spacing, independently of20pt child glyphs.");
        TestAssert.True(measured.Glyphs[2].FontSize == standalone.Glyphs[0].FontSize &&
            measured.Glyphs[3].FontSize == standalone.Glyphs[1].FontSize &&
            measured.Glyphs[3].BaselineY == standalone.Glyphs[1].BaselineY,
            "Equation glue cannot rescale or reposition the already measured script role internally.");
    }

    public static void MathFontScriptRowsRetainChildStylesBaselinesAndSourceOwners()
    {
        var resource = Resource(TestFontBuilder.CreateMathFont(Constants()));
        foreach (bool superscript in new[] { true, false })
        {
            var scripts = new DocxMathScripts(Row("B"), superscript ? null : Row("C"), superscript ? Row("C") : null);
            DocxEquation child = Fraction() with { Content = new([scripts]) };
            DocxEquation source = child with { Content = new([Row("A+"), scripts, Row("+B")]) };
            TestAssert.True(DocxMathLayout.TryGetRuns(source, CancellationToken.None, out var runs) &&
                runs.Count == 4 && runs[0].ScriptLevel == 0 && runs[1].ScriptLevel == 0 &&
                runs[2].ScriptLevel == 1 && runs[3].ScriptLevel == 0,
                "Literal siblings and base retain text style while the script role retains script style.");
            TestAssert.True(DocxMathLayout.TryMeasure(child, _ => resource, CancellationToken.None, out var basis, out _), "Expected existing script box.");
            TestAssert.True(DocxMathLayout.TryMeasure(source, _ => resource, CancellationToken.None, out var box, out var reason),
                reason ?? "One supported script must compose alongside literal atoms.");
            DocxMathBox measured = TestAssert.NotNull(box), scriptBox = TestAssert.NotNull(basis);
            TestAssert.Equal("A+BC+B", string.Concat(measured.Glyphs.Select(g => char.ConvertFromUtf32(g.SourceCodePoint))));
            for (int i = 0; i < scriptBox.Glyphs.Count; i++)
            {
                var expected = scriptBox.Glyphs[i]; var actual = measured.Glyphs[i + 2];
                TestAssert.True(actual.Font == expected.Font && actual.FontSize == expected.FontSize &&
                    actual.OriginalGlyph == expected.OriginalGlyph && actual.BaselineY == expected.BaselineY,
                    "Composition cannot change the already qualified child's glyph/style/baseline ownership.");
            }
            TestAssert.True(measured.AdvanceWidth > scriptBox.AdvanceWidth && measured.Width >= measured.AdvanceWidth,
                "Literal siblings extend logical advance while occupied ink remains bounded by the stored frame.");
            var graphics = new PdfGraphicsBuilder();
            DocxMathBoxPainter.Paint(measured, graphics, 20d, 50d, CancellationToken.None);
            using var output = new MemoryStream();
            PdfDocumentWriter.WriteBlank(output, [new PdfPage(200d, 150d, graphics.ToString(), [resource])], CancellationToken.None);
            TestAssert.True(output.Length > 0, "The composed row retains valid source encodings/resources.");
        }
    }

    public static void MathFontScriptRowsKeepUnknownDeeperFallbackAndCancellation()
    {
        var resource = Resource(TestFontBuilder.CreateMathFont(Constants()));
        var scripts = new DocxMathScripts(Row("B"), null, Row("C"));
        DocxEquation source = Fraction() with { Content = new([Row("A+"), scripts]) };
        foreach (var content in new[] { new DocxMathRow([Row("A+"), scripts with { PropertiesXml = ["<m:sSupPr/>"] }]),
            new DocxMathRow([Row("A+"), scripts with { Base = new([scripts]) }]) })
            TestAssert.True(!DocxMathLayout.TryGetRuns(source with { Content = content }, CancellationToken.None, out _) &&
                !DocxMathLayout.TryMeasure(source with { Content = content }, _ => resource, CancellationToken.None, out var box, out _) && box is null,
                "Unknown/deeper structures retain the complete source fallback.");
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryMeasure(source, _ => resource, cancelled.Token, out _, out _));
        using (OoxConversionBudget.Scope scope = OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxSceneNodesPerConversion = 3 }))
            TestAssert.Throws<OoxPdfLimitExceededException>(() => DocxMathLayout.TryMeasure(source, _ => resource, CancellationToken.None, out _, out _));
    }

    public static void MathFontFractionRowsShareBinaryBoundarySpacingAndRetainSourceOrder()
    {
        var resource = Resource(TestFontBuilder.CreateMathFont(Constants()));
        DocxEquation child = Fraction();
        var fraction = (DocxMathFraction)child.Content.Children[0];
        DocxMathRow before = Row("A+"), after = Row("+B");
        var source = child with { Content = new([before, fraction, after]), ParentFontSize = 12d };
        TestAssert.True(DocxMathLayout.TryMeasure(child, _ => resource, CancellationToken.None, out var basis, out _), "Expected existing fraction.");
        TestAssert.True(DocxMathRowLayout.TryMeasure(before, _ => resource, CancellationToken.None, out var first, out _), "Expected standalone prefix.");
        TestAssert.True(DocxMathRowLayout.TryMeasure(after, _ => resource, CancellationToken.None, out var last, out _), "Expected standalone suffix.");
        TestAssert.True(DocxMathLayout.TryMeasure(source, _ => resource, CancellationToken.None, out var box, out var reason),
            reason ?? "Existing source atoms and one fraction must form a shared row.");
        DocxMathBox measured = TestAssert.NotNull(box);
        TestAssert.Equal("A+A+BC+B", string.Concat(measured.Glyphs.Select(g => char.ConvertFromUtf32(g.SourceCodePoint))));
        DocxMathBox firstBox = TestAssert.NotNull(first), fractionBox = TestAssert.NotNull(basis), lastBox = TestAssert.NotNull(last);
        double standalone = firstBox.AdvanceWidth + fractionBox.AdvanceWidth + lastBox.AdvanceWidth;
        TestAssert.True(Math.Abs(measured.AdvanceWidth - standalone - 16d * 12d / 18d) < 1e-9,
            "A binary operator around an ordinary fraction receives the shared four-unit glue on both sides, without unary boundary reclassification.");
        TestAssert.Equal(1, measured.Rules.Count);
        TestAssert.True(measured.Rules[0].X > firstBox.AdvanceWidth && measured.Rules[0].Width == fractionBox.Rules[0].Width,
            "The child rule keeps its width while moving with its source-owned row position.");
        var graphics = new PdfGraphicsBuilder();
        DocxMathBoxPainter.Paint(measured, graphics, 20d, 50d, CancellationToken.None);
        using var output = new MemoryStream();
        PdfDocumentWriter.WriteBlank(output, [new PdfPage(200d, 150d, graphics.ToString(), [resource])], CancellationToken.None);
        TestAssert.True(output.Length > 0, "Shared row encodings/resources must remain valid.");
    }

    public static void MathFontFractionRowsKeepDeeperStructureFallbackCancellationAndWorkBounds()
    {
        var resource = Resource(TestFontBuilder.CreateMathFont(Constants()));
        DocxEquation source = Fraction();
        var fraction = (DocxMathFraction)source.Content.Children[0];
        source = source with { Content = new([Row("A+"), fraction]) };
        var deeper = source with { Content = new([Row("A+"), fraction with { Numerator = new([fraction]) }]) };
        TestAssert.True(!DocxMathLayout.TryGetRuns(deeper, CancellationToken.None, out _) &&
            !DocxMathLayout.TryMeasure(deeper, _ => resource, CancellationToken.None, out var box, out _) && box is null,
            "Unsupported deeper structures cannot publish partial formula text.");
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryMeasure(source, _ => resource, cancelled.Token, out _, out _));
        using (OoxConversionBudget.Scope scope = OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxSceneNodesPerConversion = 3 }))
            TestAssert.Throws<OoxPdfLimitExceededException>(() => DocxMathLayout.TryMeasure(source, _ => resource, CancellationToken.None, out _, out _));
    }

    public static void MathFontFractionScriptBasesRetainChildRulesRolesAndSelectedStyles()
    {
        var resource = Resource(TestFontBuilder.CreateMathFont(Constants()));
        DocxEquation source = Fraction();
        var fraction = (DocxMathFraction)source.Content.Children[0];
        foreach (bool display in new[] { true, false })
        foreach (string type in new[] { "bar", "noBar" })
        {
            var scripts = new DocxMathScripts(new([fraction with { TypeValue = type }]), Row("B"), Row("B"));
            source = source with { Display = display, Content = new([scripts]) };
            TestAssert.True(DocxMathLayout.TryGetRuns(source, CancellationToken.None, out var runs), "Expected typed nested-base discovery.");
            TestAssert.True(runs.Count == 4 && runs[0].ScriptLevel == (display ? 0 : 1) && runs[2].ScriptLevel == 1,
                "Base fraction arguments inherit the equation style while script arguments own script style.");
            TestAssert.True(DocxMathLayout.TryMeasure(source, _ => resource, CancellationToken.None, out var box, out var reason),
                reason ?? "The existing boxes must compose without dropping the base.");
            DocxMathBox measured = TestAssert.NotNull(box);
            TestAssert.Equal("A+BCBB", string.Concat(measured.Glyphs.Select(g => char.ConvertFromUtf32(g.SourceCodePoint))));
            TestAssert.Equal(type == "bar" ? 1 : 0, measured.Rules.Count);
            TestAssert.True(measured.Glyphs[^2].BaselineY <= measured.Glyphs[3].BaselineY &&
                measured.Glyphs[^1].BaselineY >= measured.Glyphs[0].BaselineY &&
                measured.Glyphs[^2].BaselineY < measured.Glyphs[^1].BaselineY,
                "Nested-base scripts retain their lower/upper roles; zero minimum-drop constants permit equal child baselines.");
            var graphics = new PdfGraphicsBuilder();
            DocxMathBoxPainter.Paint(measured, graphics, 20d, 50d, CancellationToken.None);
            using var output = new MemoryStream();
            PdfDocumentWriter.WriteBlank(output, [new PdfPage(200d, 150d, graphics.ToString(), [resource])], CancellationToken.None);
            TestAssert.True(output.Length > 0, "Nested glyphs and rules must retain valid owned resources/encodings.");
        }
    }

    public static void MathFontFractionScriptBasesRetainDeeperFallbackCancellationAndWorkBounds()
    {
        var resource = Resource(TestFontBuilder.CreateMathFont(Constants()));
        DocxEquation source = Fraction();
        var fraction = (DocxMathFraction)source.Content.Children[0];
        var scripts = new DocxMathScripts(new([fraction]), null, Row("B"));
        source = source with { Content = new([scripts]) };
        var deeper = source with { Content = new([scripts with
            { Base = new([fraction with { Numerator = new([fraction]) }]) }]) };
        TestAssert.True(!DocxMathLayout.TryGetRuns(deeper, CancellationToken.None, out _) &&
            !DocxMathLayout.TryMeasure(deeper, _ => resource, CancellationToken.None, out var box, out _) && box is null,
            "The bounded first nesting grammar cannot admit deeper structures through recursive composition.");
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryMeasure(source, _ => resource, cancelled.Token, out _, out _));
        using (OoxConversionBudget.Scope scope = OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxSceneNodesPerConversion = 3 }))
            TestAssert.Throws<OoxPdfLimitExceededException>(() => DocxMathLayout.TryMeasure(source, _ => resource, CancellationToken.None, out _, out _));
    }

    public static void MathFontNoBarStacksKeepSeparatedSourceRolesWithoutPaintingARule()
    {
        byte[] constants = Constants(thickness: -1);
        Set(OpenTypeMathConstant.StackTopDisplayStyleShiftUp, 600);
        Set(OpenTypeMathConstant.StackBottomDisplayStyleShiftDown, 400);
        Set(OpenTypeMathConstant.StackDisplayStyleGapMin, 1800);
        var resource = Resource(TestFontBuilder.CreateMathFont(constants));
        DocxEquation source = Fraction();
        var fraction = (DocxMathFraction)source.Content.Children[0];
        source = source with { Content = new([fraction with { TypeValue = "noBar" }]) };
        TestAssert.True(DocxMathLayout.TryMeasure(source, _ => resource, CancellationToken.None, out var box, out var reason),
            reason ?? "Stack metrics do not depend on an unused fraction-rule thickness.");
        DocxMathBox measured = TestAssert.NotNull(box);
        TestAssert.Equal(0, measured.Rules.Count);
        TestAssert.Equal("A+BC", string.Concat(measured.Glyphs.Select(g => char.ConvertFromUtf32(g.SourceCodePoint))));
        double bottomOfTop = double.PositiveInfinity, topOfBottom = double.NegativeInfinity;
        foreach (DocxMathGlyphPlacement glyph in measured.Glyphs)
        {
            TestAssert.True(resource.Font.Font.TryGetGlyphBounds(glyph.OriginalGlyph, out var bounds), "Expected source ink.");
            double scale = glyph.FontSize / resource.Font.Font.UnitsPerEm;
            if (glyph.SourceCodePoint == 'C') topOfBottom = glyph.BaselineY + bounds.YMax * scale;
            else bottomOfTop = Math.Min(bottomOfTop, glyph.BaselineY + bounds.YMin * scale);
        }
        TestAssert.True(bottomOfTop - topOfBottom >= 1800d * 12d / resource.Font.Font.UnitsPerEm - 1e-9,
            "The independently stored ink bounds must respect the font's minimum stack gap.");
        var graphics = new PdfGraphicsBuilder();
        DocxMathBoxPainter.Paint(measured, graphics, 20d, 50d, CancellationToken.None);
        TestAssert.True(!graphics.ToString().Contains(" re f", StringComparison.Ordinal), "A stack must not acquire a visible fraction bar.");
        using var output = new MemoryStream();
        PdfDocumentWriter.WriteBlank(output, [new PdfPage(200d, 150d, graphics.ToString(), [resource])], CancellationToken.None);
        TestAssert.True(output.Length > 0, "Owned source encodings/resources must form a valid PDF.");
        void Set(OpenTypeMathConstant field, short value) => BinaryPrimitives.WriteInt16BigEndian(constants.AsSpan(18 + (int)field * 4), value);
    }

    public static void MathFontNoBarStacksRejectInvalidMetricsAndRetainInlineScriptOwnership()
    {
        DocxEquation source = Fraction();
        var fraction = (DocxMathFraction)source.Content.Children[0];
        source = source with { Display = false, Content = new([fraction with { TypeValue = "noBar" }]) };
        var valid = Resource(TestFontBuilder.CreateMathFont(Constants()));
        TestAssert.True(DocxMathLayout.TryMeasure(source, _ => valid, CancellationToken.None, out var box, out _), "Expected inline stack.");
        TestAssert.True(TestAssert.NotNull(box).Glyphs.All(g => Math.Abs(g.FontSize - 8.52d) < 1e-9),
            "Stack children retain the existing parent-owned inline script scale.");
        byte[] invalid = Constants();
        BinaryPrimitives.WriteInt16BigEndian(invalid.AsSpan(18 + (int)OpenTypeMathConstant.StackGapMin * 4), -1);
        var bad = Resource(TestFontBuilder.CreateMathFont(invalid));
        TestAssert.True(!DocxMathLayout.TryMeasure(source, _ => bad, CancellationToken.None, out box, out _)
            && box is null, "Invalid stack metrics cannot publish a partial box.");
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryMeasure(source, _ => valid, cancelled.Token, out _, out _));
    }

    public static void MathFontNormalFractionChildrenRetainParentConstantsAndLeadingWithoutChildMathTables()
    {
        OpenTypeFont parent = OpenTypeFont.Load(TestFontBuilder.CreateMathFont(Constants()));
        OpenTypeFont text = TestFontBuilder.LoadTestFont();
        var child = new PdfFontResource("Text1", PdfEmbeddedFont.Create(text, "A+BC".Select(c => (int)c), CancellationToken.None));
        DocxMathRow Normal(string value) => new([new DocxMathRun(new(value, 12d, null, false, false, false, null, "TestFont"),
            null, null) { NormalTextValue = true }]);
        var equation = new DocxEquation(new([new DocxMathFraction(Normal("A+B"), Normal("C"), null)]), 0, 0, false, null, null, null)
            { ParentFontSize = 12d };
        TestAssert.True(!DocxMathLayout.TryMeasure(equation, _ => child, CancellationToken.None, out _, out _),
            "Text faces cannot invent parent MATH fraction constants.");
        TestAssert.True(DocxMathLayout.TryMeasure(equation, _ => child, CancellationToken.None, out var box, out var reason, parent),
            reason ?? "A separately owned parent supplies script scale and fraction constants.");
        DocxMathBox measured = TestAssert.NotNull(box);
        TestAssert.True(measured.Glyphs.All(glyph => ReferenceEquals(glyph.Font, child) && Math.Abs(glyph.FontSize - 8.52d) < 1e-9),
            "Scaled normal text retains its text-face subset and ordinary glyphs.");
        TestAssert.True(DocxMathLayout.TryGetLineMetrics(measured, CancellationToken.None, out double above, out _),
            "Parent-owned leading survives composition even when no child has MATH.");
        TestAssert.True(Math.Abs(above - measured.AboveBaseline - 120d * 12d / parent.UnitsPerEm) < 1e-9,
            "The parent font contributes one mathematical leading term, independently of child font tables.");
        TestAssert.True(measured.ParentMetrics is { } owner && ReferenceEquals(owner.Font, parent) && owner.Size == 12d,
            "Font and size remain one typed metric owner.");
    }

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
