using System.Buffers.Binary;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxMathFunctionLayoutTests
{
    public static void MathFontDelimitedSubformulasSharePaintedThresholdAndSemanticGlue()
    {
        foreach (ushort minimum in new ushort[] { 0, 1500, 2500 })
        {
            byte[] bytes = MathFontBytes();
            var tables = TestFontBuilder.GetDirectoryEntries(bytes);
            int math = tables.Single(t => t.Tag == "MATH").Offset, glyf = tables.Single(t => t.Tag == "glyf").Offset;
            BinaryPrimitives.WriteUInt16BigEndian(bytes.AsSpan(math + 14), minimum);
            foreach (int glyph in new[] { 9, 10, 127, 128 })
            {
                bool variant = glyph >= 127;
                BinaryPrimitives.WriteInt16BigEndian(bytes.AsSpan(glyf + glyph * 10 + 4), (short)(variant ? -750 : -250));
                BinaryPrimitives.WriteInt16BigEndian(bytes.AsSpan(glyf + glyph * 10 + 8), (short)(variant ? 1250 : 750));
            }
            var font = OpenTypeFont.Load(bytes);
            var resource = new PdfFontResource("Subformula", PdfEmbeddedFont.CreateForGlyphMappings(font,
                new[] { ((ushort)9, 40), ((ushort)10, 41), ((ushort)128, 40), ((ushort)127, 41) }
                    .Concat("sinBC".Select(c => (font.MapCodePoint(c), (int)c))), CancellationToken.None));
            foreach (bool fixedSize in new[] { false, true })
            foreach (double childSize in new[] { 12d, 20d })
            {
                var child = Run("B") with { Text = Run("B").Text with { FontSize = childSize } };
                var delimiter = new DocxMathDelimiter([new([child])]) { GrowValue = !fixedSize };
                var function = Function(new([Run("sin")]), new([delimiter]));
                var equation = Equation(function);
                bool inner = minimum == 1500 && !fixedSize && childSize == 20d;
                TestAssert.True(DocxMathLayout.TryMeasure(equation, _ => resource, CancellationToken.None, out var box, out var reason,
                    resolveControlFont: _ => resource, resolveControlMetrics: _ => font), reason ?? "Expected bounded function delimiter.");
                var measured = TestAssert.NotNull(box); double nameAdvance = measured.Glyphs.Take(3).Sum(g => g.Advance);
                TestAssert.True(Math.Abs(measured.Glyphs[3].X - nameAdvance - (inner ? 2d : 0d)) < 1e-9,
                    "Only a painted mark meeting its own font's positive subformula threshold consumes application glue; fixed marks and unknown thresholds retain opening semantics.");
                var row = equation with { Content = new([Run("B"), delimiter, Run("C")]) };
                TestAssert.True(DocxMathLayout.TryMeasure(row, _ => resource, CancellationToken.None, out var mixed, out reason), reason ?? "Expected ordinary subformula boundaries.");
                var sequence = TestAssert.NotNull(mixed);
                TestAssert.True(Math.Abs(sequence.Glyphs[1].X - sequence.Glyphs[0].Advance - (inner ? 2d : 0d)) < 1e-9 &&
                    Math.Abs(sequence.Glyphs[^1].X - sequence.Glyphs[^2].X - sequence.Glyphs[^2].Advance - (inner ? 2d : 0d)) < 1e-9,
                    "The same inner edge supplies incoming and outgoing ordinary spacing, rather than a function-only correction.");
            }
        }
        TestAssert.Equal(4, DocxMathRowLayout.Glue(DocxMathRowLayout.AtomKind.Binary, DocxMathRowLayout.AtomKind.Inner));
        TestAssert.Equal(5, DocxMathRowLayout.Glue(DocxMathRowLayout.AtomKind.Inner, DocxMathRowLayout.AtomKind.Relation));
    }

    public static void MathFontMixedFunctionRowsShareParentFrameAndSemanticEdgesWithoutPartialPaint()
    {
        var font = OpenTypeFont.Load(MathFontBytes()); var resource = Resource("Rows", font, "BcustomCsin=Z");
        foreach (bool reverse in new[] { false, true })
        foreach (double parentSize in new[] { 12d, 18d })
        {
            var first = Function(new([Run("custom", color: "0000FF")]), new([Run("C")]), 20d);
            var lastArgument = Run("Z", color: "008000");
            var second = Function(new([Run("sin")]), new([lastArgument]), 8d);
            DocxMathNode[] functions = reverse ? [second, first] : [first, second];
            var equation = Equation(first) with { Content = new([Run("B"), .. functions, Run("=B")]), ParentFontSize = parentSize };
            TestAssert.True(DocxMathLayout.TryGetRuns(equation, CancellationToken.None, out var sources) &&
                sources.Count == 6 && sources.Count(s => s.ControlOwner is DocxMathFunction) == 2,
                "Mixed discovery retains literal order and two independent invisible control owners.");
            TestAssert.True(DocxMathLayout.TryMeasure(equation, _ => resource, CancellationToken.None, out var box, out var reason,
                resolveControlMetrics: _ => font), reason ?? "Expected contextual function row.");
            var measured = TestAssert.NotNull(box);
            TestAssert.True(measured.ParentMetrics is { IncludeTypographicFrame: true } owner && owner.Size == parentSize,
                "The surrounding row owns its frame in either control order, independent of authored standalone control sizes.");
            TestAssert.True(Math.Abs(measured.Glyphs[1].X - measured.Glyphs[0].Advance - parentSize / 6d) < 1e-9,
                "An ordinary atom before a function consumes one thin math space.");
            int nameLength = reverse ? 3 : 6;
            var nameLast = measured.Glyphs[nameLength]; var argument = measured.Glyphs[nameLength + 1];
            TestAssert.True(Math.Abs(argument.X - nameLast.X - nameLast.Advance - parentSize / 6d) < 1e-9,
                "Application spacing uses the same parent context while painted name glyphs keep their authored size.");
            TestAssert.True(measured.Glyphs.Where(g => g.SourceCodePoint == 'Z').All(g => g.Color.Green == 128),
                "Per-node argument paint remains independent of the row frame.");
            var missing = Resource("MissingLateArgument", font, "BcustomCsin=");
            TestAssert.True(!DocxMathLayout.TryMeasure(equation, r => ReferenceEquals(r, lastArgument) ? missing : resource,
                CancellationToken.None, out var absent, out _, resolveControlMetrics: _ => font) && absent is null,
                "A missing encoding in either function rejects the complete row without publishing a partial prefix.");
            var nested = equation with { Content = new([Run("B"), first with { Argument = new([second]) }]) };
            TestAssert.True(!DocxMathLayout.TryGetRuns(nested, CancellationToken.None, out _),
                "Mixed eligibility cannot recursively expand nested application arguments.");
        }
        var bounded = Equation(Function(new([Run("sin")]), new([Run("C")]))) with
            { Content = new([Run("B"), Function(new([Run("sin")]), new([Run("C")]))]) };
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryMeasure(bounded, _ => resource, cancelled.Token, out _, out _));
        using var scope = OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxSceneNodesPerConversion = 3 });
        TestAssert.Throws<OoxPdfLimitExceededException>(() => DocxMathLayout.TryMeasure(bounded, _ => resource, CancellationToken.None, out _, out _, resolveControlMetrics: _ => font));
    }

    public static void MathFontFunctionMetricOnlyControlsKeepIndependentFontsAndWholeLineFrame()
    {
        var metric = OpenTypeFont.Load(MathFontBytes()); var text = TestFontBuilder.LoadTestFont();
        var nameRun = Run("custom", normal: true, color: "0000FF"); var argumentRun = Run("B", normal: true, color: "008000");
        var names = Resource("Names", text, "custom"); var arguments = Resource("Arguments", text, "B");
        var function = Function(new([nameRun]), new([argumentRun]), 20d);
        TestAssert.True(DocxMathLayout.GetControlCodePoints(function).Count == 0, "An invisible application control owns metrics without requesting glyph mappings.");
        TestAssert.True(DocxMathLayout.TryMeasure(Equation(function), r => ReferenceEquals(r, nameRun) ? names : arguments,
            CancellationToken.None, out var box, out var reason, resolveControlMetrics: _ => metric), reason ?? "Expected metric-only function control.");
        var measured = TestAssert.NotNull(box); double nameAdvance = measured.Glyphs.Take(6).Sum(g => g.Advance);
        TestAssert.True(Math.Abs(measured.Glyphs[^1].X - nameAdvance - 20d / 6d) < 1e-9 &&
            measured.Glyphs.Take(6).All(g => ReferenceEquals(g.Font, names) && g.Color.Blue == 255) &&
            ReferenceEquals(measured.Glyphs[^1].Font, arguments) && measured.Glyphs[^1].Color.Green == 128,
            "Arbitrary literal names, argument fonts/colors and control-owned spacing remain independent.");
        TestAssert.True(measured.ParentMetrics is { IncludeTypographicFrame: true } owner && ReferenceEquals(owner.Font, metric),
            "The unpainted control explicitly owns its ordinary typographic frame.");
        TestAssert.True(DocxMathLayout.TryGetLineMetrics(measured, CancellationToken.None, out double above, out double height) &&
            height - above >= -metric.Os2.TypographicDescender * OfficePdfTextEmissionProfile.FontSize(20d) / metric.UnitsPerEm - 1e-9,
            "A larger invisible control must retain its descent even though every painted glyph is smaller.");
        var graphics = new PdfGraphicsBuilder(); DocxMathBoxPainter.Paint(measured, graphics, 20d, 60d, CancellationToken.None);
        using var output = new MemoryStream();
        PdfDocumentWriter.WriteBlank(output, [new PdfPage(250d, 180d, graphics.ToString(), [names, arguments])], CancellationToken.None);
        TestAssert.True(output.Length > 0, "Only painted resources are required; a metric-only font never needs an empty subset.");
    }

    public static void MathFontFunctionOpeningArgumentsKeepNestedControlAndZeroApplicationGap()
    {
        var font = OpenTypeFont.Load(MathFontBytes()); var plain = TestFontBuilder.LoadTestFont();
        var names = Resource("Names", plain, "sin");
        var arguments = new PdfFontResource("Fence", PdfEmbeddedFont.CreateForGlyphMappings(font,
            new[] { ((ushort)9, 40), ((ushort)10, 41), ((ushort)128, 40), ((ushort)127, 41), (font.MapCodePoint('B'), (int)'B') }, CancellationToken.None));
        var name = new DocxMathRow([Run("sin", normal: true)]);
        var delimiter = new DocxMathDelimiter([new([Run("B")])]); var function = Function(name, new([delimiter]));
        TestAssert.True(DocxMathLayout.TryGetRuns(Equation(function), CancellationToken.None, out var sources) && sources.Count == 2 &&
            ReferenceEquals(sources[0].ControlOwner, function) && ReferenceEquals(sources[1].ControlOwner, delimiter),
            "Function metrics and nested delimiter glyphs retain separate source owners.");
        TestAssert.True(DocxMathLayout.TryMeasure(Equation(function), r => r.NormalText ? names : arguments, CancellationToken.None,
            out var box, out var reason, resolveControlFont: c => ReferenceEquals(c, delimiter) ? arguments : null,
            resolveControlMetrics: _ => font), reason ?? "Expected fenced argument.");
        var measured = TestAssert.NotNull(box);
        TestAssert.Equal("sin(B)", string.Concat(measured.Glyphs.Select(g => char.ConvertFromUtf32(g.SourceCodePoint))));
        TestAssert.True(Math.Abs(measured.Glyphs[3].X - measured.Glyphs.Take(3).Sum(g => g.Advance)) < 1e-9,
            "An opening argument consumes no extra application space.");
    }

    public static void MathFontFunctionFractionAndScriptArgumentsRetainRulesAndSourceLevels()
    {
        var font = OpenTypeFont.Load(MathFontBytes()); var resource = Resource("Children", font, "sinBC");
        foreach (bool display in new[] { true, false })
        foreach (bool fraction in new[] { true, false })
        {
            var name = new DocxMathRow([Run("sin")]); var basis = new DocxMathRow([Run("B")]); var child = new DocxMathRow([Run("C")]);
            DocxMathNode argument = fraction ? new DocxMathFraction(basis, child, "bar") : new DocxMathScripts(basis, null, child);
            var function = Function(name, new([argument])); var equation = Equation(function) with { Display = display };
            TestAssert.True(DocxMathLayout.TryGetRuns(equation, CancellationToken.None, out var sources) && sources.Count == 3 &&
                sources[0].ScriptLevel == 0 && sources[^1].ScriptLevel == (fraction && display ? 0 : 1),
                "The name and bounded argument retain independent text/script roles.");
            TestAssert.True(DocxMathLayout.TryMeasure(equation, _ => resource, CancellationToken.None, out var box, out var reason,
                resolveControlMetrics: _ => font), reason ?? "Expected bounded argument boxes.");
            TestAssert.Equal(fraction ? 1 : 0, TestAssert.NotNull(box).Rules.Count);
            TestAssert.True(TestAssert.NotNull(box).Rules.All(r => r.X > 0d), "Inner fraction rules move with their argument box.");
        }
    }

    public static void MathFontFunctionFallbackCancellationAndWorkRemainBounded()
    {
        var font = OpenTypeFont.Load(MathFontBytes()); var resource = Resource("Function", font, "sinBC");
        var function = Function(new([Run("sin")]), new([Run("B")]));
        foreach (var invalid in new[] { function with { HasUnsupportedProperties = true }, function with { Name = new([]) },
            function with { Name = new([new DocxMathFraction(new([Run("B")]), new([Run("C")]), "bar")]) },
            function with { Argument = new([function]) } })
            TestAssert.True(!DocxMathLayout.TryMeasure(Equation(invalid), _ => resource, CancellationToken.None, out var absent, out _,
                resolveControlMetrics: _ => font) && absent is null, "Unsupported name/property/deeper roles keep complete fallback.");
        var missing = Resource("MissingName", font, "BC");
        TestAssert.True(!DocxMathLayout.TryMeasure(Equation(function), _ => missing, CancellationToken.None, out var box, out _,
            resolveControlMetrics: _ => font) && box is null, "A missing name CID cannot publish an argument-only function.");
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryGetRuns(Equation(function), cancelled.Token, out _));
        TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryMeasure(Equation(function), _ => resource, cancelled.Token, out _, out _));
        using var scope = OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxSceneNodesPerConversion = 3 });
        TestAssert.Throws<OoxPdfLimitExceededException>(() => DocxMathLayout.TryMeasure(Equation(function), _ => resource, CancellationToken.None, out _, out _, resolveControlMetrics: _ => font));
    }

    private static DocxMathRun Run(string text, bool normal = false, string? color = null) =>
        new(new(text, 12d, color, false, false, false, null, "TestFont"), "p", null) { NormalTextValue = normal };
    private static DocxMathFunction Function(DocxMathRow name, DocxMathRow argument, double size = 12d) =>
        new(name, argument) { ControlText = new("", size, "FF0000", false, false, false, null, "TestFont") };
    private static DocxEquation Equation(DocxMathFunction function) => new(new([function]), 0, 0, true, null, null, null) { ParentFontSize = 12d };
    private static PdfFontResource Resource(string name, OpenTypeFont font, string text) =>
        new(name, PdfEmbeddedFont.Create(font, text.Select(c => (int)c), CancellationToken.None));
    private static byte[] MathFontBytes()
    {
        byte[] table = new byte[270];
        void Put(int at, ushort value) => BinaryPrimitives.WriteUInt16BigEndian(table.AsSpan(at), value);
        void Set(OpenTypeMathConstant field, ushort value) => Put(18 + (int)field * 4, value);
        Put(0, 1); Put(4, 10); Put(10, 73); Put(12, 50); Set(OpenTypeMathConstant.AxisHeight, 250); Set(OpenTypeMathConstant.MathLeading, 120);
        Set(OpenTypeMathConstant.SuperscriptShiftUp, 500); Set(OpenTypeMathConstant.FractionRuleThickness, 40);
        Set(OpenTypeMathConstant.FractionNumeratorDisplayStyleShiftUp, 700); Set(OpenTypeMathConstant.FractionDenominatorDisplayStyleShiftDown, 500);
        Set(OpenTypeMathConstant.FractionNumeratorShiftUp, 500); Set(OpenTypeMathConstant.FractionDenominatorShiftDown, 400);
        foreach (var field in new[] { OpenTypeMathConstant.FractionNumDisplayStyleGapMin, OpenTypeMathConstant.FractionDenomDisplayStyleGapMin,
            OpenTypeMathConstant.FractionNumeratorGapMin, OpenTypeMathConstant.FractionDenominatorGapMin }) Set(field, 100);
        Put(8, 224); Put(226, 14); Put(230, 2); Put(234, 22); Put(236, 34); Put(238, 1); Put(240, 2); Put(242, 9); Put(244, 10);
        Put(248, 2); Put(250, 9); Put(252, 1000); Put(254, 128); Put(256, 2000);
        Put(260, 2); Put(262, 10); Put(264, 1000); Put(266, 127); Put(268, 2000);
        return TestFontBuilder.CreateMathFont(table);
    }
}
