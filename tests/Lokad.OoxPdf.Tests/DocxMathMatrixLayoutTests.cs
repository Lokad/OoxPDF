using System.Buffers.Binary;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxMathMatrixLayoutTests
{
    public static void MathFontMatrixMixedCellsKeepOrderedSourcesAndCompleteFallbacks()
    {
        var font = MathFont(); var resource = Resource("MixedCells", font, "BC+");
        var leading = Run("B+"); var numerator = Run("B"); var denominator = Run("C"); var trailing = Run("+C");
        var between = Run("B"); var nextLeading = Run("C+"); var basis = Run("B"); var power = Run("C");
        var nextTrailing = Run("+C"); var last = Run("C");
        var fraction = new DocxMathFraction(new([numerator]), new([denominator]), "bar");
        var scripts = new DocxMathScripts(new([basis]), null, new([power]));
        var matrix = new DocxMathMatrix([new([new([leading, fraction, trailing]), new([between])]),
            new([new([nextLeading, scripts, nextTrailing]), new([last])])]);
        var equation = Equation(matrix) with { Display = false };
        TestAssert.True(DocxMathLayout.TryGetRuns(equation, default, out var sources), "Mixed cells retain their supported complete source.");
        TestAssert.True(sources.Select(source => source.Run).SequenceEqual(new[]
            { leading, numerator, denominator, trailing, between, nextLeading, basis, power, nextTrailing, last }),
            "Mixed cells and intervening ordinary cells keep source identity in row/cell order.");
        TestAssert.Equal("0,1,1,0,0,0,0,1,0,0", string.Join(',', sources.Select(source => source.ScriptLevel)));
        TestAssert.True(ReferenceEquals(sources[0].ControlOwner, matrix), "The grid retains its outer metric owner.");
        TestAssert.True(DocxMathLayout.TryMeasure(equation, _ => resource, default, out var measured, out var reason,
            resolveControlMetrics: _ => font), reason ?? "Expected a complete mixed-cell grid.");
        var box = TestAssert.NotNull(measured);
        TestAssert.Equal(14, box.Glyphs.Count); TestAssert.Equal(1, box.Rules.Count);
        foreach (DocxMathNode child in new DocxMathNode[] { matrix,
            new DocxMathRadical(new([basis]), new([]), true),
            fraction with { HasUnsupportedProperties = true }, fraction with { Numerator = new([fraction]) } })
        {
            var invalid = matrix with { Rows = [matrix.Rows[0], new([new([nextLeading, child, nextTrailing]), new([last])])] };
            TestAssert.True(!DocxMathLayout.TryGetRuns(Equation(invalid), default, out _) &&
                !DocxMathLayout.TryMeasure(Equation(invalid), _ => resource, default, out var absent, out _, resolveControlMetrics: _ => font) && absent is null,
                "A mixed cell cannot admit another grid, deeper container or unsupported child source.");
        }
        var missing = Resource("MissingLastCell", font, "B+");
        TestAssert.True(!DocxMathLayout.TryMeasure(equation, run => ReferenceEquals(run, last) ? missing : resource,
            default, out var partial, out _, resolveControlMetrics: _ => font) && partial is null,
            "A late missing CID cannot publish an earlier mixed row.");
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryGetRuns(equation, cancelled.Token, out _));
        TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryMeasure(equation, _ => resource, cancelled.Token, out _, out _));
        using var scope = OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxSceneNodesPerConversion = 1 });
        TestAssert.Throws<OoxPdfLimitExceededException>(() => DocxMathLayout.TryMeasure(equation, _ => resource, default,
            out _, out _, resolveControlMetrics: _ => font));
    }

    public static void MathFontMatrixFunctionArgumentsRetainOwnersIndependentPaintAndBoundedFallback()
    {
        var outerFont = MathFont(); var innerFont = MathFont(); var plain = TestFontBuilder.LoadTestFont();
        var resource = Resource("FunctionCells", plain, "BC");
        var matrix = Matrix([Run("B"), Run("C")], [Run("C"), Run("B")]) with { ControlText = new("", 20d, null, false, false, false, null, "TestFont") };
        var function = new DocxMathFunction(new([Run("B")]), new([matrix])) { ControlText = new("", 12d, null, false, false, false, null, "TestFont") };
        var equation = Equation(matrix) with { Content = new([function]) };
        TestAssert.True(DocxMathLayout.TryGetRuns(equation, CancellationToken.None, out var sources) && sources.Count == 5 &&
            ReferenceEquals(sources[0].ControlOwner, function) && ReferenceEquals(sources[1].ControlOwner, matrix),
            "Function names and matrix arguments retain distinct metric owners without duplicating their painted source.");
        OpenTypeFont? Metrics(IDocxMathControl control) => ReferenceEquals(control, matrix) ? innerFont : ReferenceEquals(control, function) ? outerFont : null;
        TestAssert.True(DocxMathLayout.TryMeasure(equation, _ => resource, CancellationToken.None, out var composed, out var reason,
            resolveControlMetrics: Metrics), reason ?? "Expected bounded matrix argument.");
        TestAssert.True(DocxMathLayout.TryMeasure(Equation(matrix), _ => resource, CancellationToken.None, out var alone, out reason,
            resolveControlMetrics: Metrics, rowControlSize: 12d), reason ?? "Expected argument in surrounding function context.");
        var box = TestAssert.NotNull(composed); var child = TestAssert.NotNull(alone); double x = box.Glyphs[1].X - child.Glyphs[0].X;
        TestAssert.True(x > box.Glyphs[0].Advance && child.Glyphs.Select((g, i) => (g, i)).All(pair =>
            Math.Abs(box.Glyphs[pair.i + 1].X - pair.g.X - x) < 1e-9 && box.Glyphs[pair.i + 1].BaselineY == pair.g.BaselineY),
            "The existing function gap translates a complete matrix while surrounding context owns the effective control size.");
        var graphics = new PdfGraphicsBuilder(); DocxMathBoxPainter.Paint(box, graphics, 20d, 80d, CancellationToken.None);
        using var output = new MemoryStream();
        PdfDocumentWriter.WriteBlank(output, [new PdfPage(300d, 200d, graphics.ToString(), [resource])], CancellationToken.None);
        TestAssert.True(output.Length > 0, "Both invisible owners use metrics without adding unused PDF font resources.");
        TestAssert.True(!DocxMathLayout.TryMeasure(equation, _ => resource, CancellationToken.None, out var absent, out _,
            resolveControlMetrics: control => ReferenceEquals(control, matrix) ? null : outerFont) && absent is null,
            "A missing matrix metric owner rejects the complete application.");
        var deeper = function with { Argument = new([function]) };
        TestAssert.True(!DocxMathLayout.TryGetRuns(equation with { Content = new([deeper]) }, CancellationToken.None, out _),
            "Adding a matrix argument cannot admit arbitrary function recursion.");
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryGetRuns(equation, cancelled.Token, out _));
        using var scope = OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxSceneNodesPerConversion = 3 });
        TestAssert.Throws<OoxPdfLimitExceededException>(() => DocxMathLayout.TryMeasure(equation, _ => resource, CancellationToken.None,
            out _, out _, resolveControlMetrics: Metrics));
    }

    public static void MathFontMatrixFencesRetainBothControlOwnersAndBoundedArgumentRoles()
    {
        var outerFont = MathFont(); var innerFont = MathFont(); var plain = TestFontBuilder.LoadTestFont();
        var painted = Resource("Cells", plain, "BC"); var marks = Resource("Marks", outerFont, "()");
        var matrix = Matrix([Run("B"), Run("C")], [Run("C"), Run("B")]) with { ControlText = new("", 20d, null, false, false, false, null, "TestFont") };
        var delimiter = new DocxMathDelimiter([new([matrix])]) { GrowValue = false,
            ControlText = new("", 12d, "FF0000", false, false, false, null, "TestFont") };
        var equation = Equation(matrix) with { Content = new([delimiter]) };
        TestAssert.True(DocxMathLayout.TryGetRuns(equation, CancellationToken.None, out var sources) && sources.Count == 4 &&
            sources[0].ControlOwners.Count == 2 && ReferenceEquals(sources[0].ControlOwners[0], delimiter) &&
            ReferenceEquals(sources[0].ControlOwners[1], matrix) && sources.Skip(1).All(s => s.ControlOwners.Count == 0),
            "One source run retains ordered outer painted and inner metric-only owners without duplicating painted cell sources.");
        TestAssert.True(DocxMathLayout.GetControlCodePoints(delimiter).Count == 2 && DocxMathLayout.GetControlCodePoints(matrix).Count == 0,
            "Only the enclosing marks require glyph mappings.");
        bool innerResolved = false;
        TestAssert.True(DocxMathLayout.TryMeasure(equation, _ => painted, CancellationToken.None, out var box, out var reason,
            resolveControlFont: control => ReferenceEquals(control, delimiter) ? marks : null,
            resolveControlMetrics: control => { innerResolved |= ReferenceEquals(control, matrix); return ReferenceEquals(control, matrix) ? innerFont : null; }),
            reason ?? "Expected bounded fenced grid.");
        var fenced = TestAssert.NotNull(box);
        TestAssert.True(innerResolved && fenced.Glyphs.Count == 6 && fenced.Glyphs[0].Color.Red == 255 &&
            ReferenceEquals(fenced.Glyphs[0].Font, marks) && ReferenceEquals(fenced.Glyphs[^1].Font, marks) &&
            fenced.Glyphs.Skip(1).Take(4).All(g => ReferenceEquals(g.Font, painted)),
            "Separate control metrics and painted font resources survive the same bounded argument box.");
        var graphics = new PdfGraphicsBuilder(); DocxMathBoxPainter.Paint(fenced, graphics, 20d, 80d, CancellationToken.None);
        using var output = new MemoryStream();
        PdfDocumentWriter.WriteBlank(output, [new PdfPage(300d, 200d, graphics.ToString(), [marks, painted])], CancellationToken.None);
        TestAssert.True(output.Length > 0, "The inner metric font needs no PDF resource or unused subset.");
        TestAssert.True(!DocxMathLayout.TryMeasure(equation, _ => painted, CancellationToken.None, out var absent, out _,
            resolveControlFont: _ => marks, resolveControlMetrics: _ => null) && absent is null,
            "A missing nested metric owner cannot publish enclosing marks or partial cells.");
        var deeper = delimiter with { Arguments = [new([delimiter])] };
        TestAssert.True(!DocxMathLayout.TryGetRuns(equation with { Content = new([deeper]) }, CancellationToken.None, out _),
            "General delimiter recursion stays outside this bounded matrix argument grammar.");
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryGetRuns(equation, cancelled.Token, out _));
        using var scope = OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxSceneNodesPerConversion = 3 });
        TestAssert.Throws<OoxPdfLimitExceededException>(() => DocxMathLayout.TryMeasure(equation, _ => painted, CancellationToken.None,
            out _, out _, resolveControlFont: _ => marks, resolveControlMetrics: _ => innerFont));
    }

    public static void MathFontMatrixMixedRowsRetainIndependentOwnersAndWholeFallback()
    {
        var font = MathFont(); var resource = Resource("Mixed", font, "BC+=");
        var first = Matrix([Run("B"), Run("C")], [Run("C"), Run("B")]);
        var second = first with { ControlText = new("", 20d, null, false, false, false, null, "TestFont") };
        var prefix = Run("B+") with { NormalTextValue = false }; var suffix = Run("=C") with { NormalTextValue = false };
        var equation = Equation(first) with { Content = new([prefix, first, second, suffix]) };
        TestAssert.True(DocxMathLayout.TryGetRuns(equation, CancellationToken.None, out var sources) &&
            sources.Count(s => ReferenceEquals(s.ControlOwner, first)) == 1 && sources.Count(s => ReferenceEquals(s.ControlOwner, second)) == 1,
            "Adjacent matrices retain independent metric-only control owners in the shared structured row.");
        TestAssert.True(DocxMathLayout.TryMeasure(equation, _ => resource, CancellationToken.None, out var measured, out var reason,
            resolveControlMetrics: _ => font), reason ?? "Expected bounded mixed matrix row.");
        var row = TestAssert.NotNull(measured); int offset = 2;
        foreach (var matrix in new[] { first, second })
        {
            TestAssert.True(DocxMathLayout.TryMeasure(Equation(matrix), _ => resource, CancellationToken.None, out var single, out reason,
                resolveControlMetrics: _ => font, rowControlSize: 12d), reason ?? "Expected matrix in the shared row context.");
            var child = TestAssert.NotNull(single); double x = row.Glyphs[offset].X - child.Glyphs[0].X;
            TestAssert.True(child.Glyphs.Select((g, i) => (g, i)).All(pair =>
                Math.Abs(row.Glyphs[offset + pair.i].X - pair.g.X - x) < 1e-9 && row.Glyphs[offset + pair.i].BaselineY == pair.g.BaselineY),
                "Composition translates complete matrix boxes while the surrounding row owns their effective control size.");
            if (matrix.ControlText is { FontSize: 20d })
            {
                TestAssert.True(DocxMathLayout.TryMeasure(Equation(matrix), _ => resource, CancellationToken.None, out var standalone, out _,
                    resolveControlMetrics: _ => font) && TestAssert.NotNull(standalone).Glyphs[0].BaselineY != child.Glyphs[0].BaselineY,
                    "An authored standalone control keeps its earlier geometry; row context does not mutate source formatting.");
            }
            offset += child.Glyphs.Count;
        }
        TestAssert.Equal("B+BCCBBCCB=C", string.Concat(row.Glyphs.Select(g => char.ConvertFromUtf32(g.SourceCodePoint))));
        var invalid = equation with { Content = new([prefix, first, second with { HasUnsupportedProperties = true }, suffix]) };
        TestAssert.True(!DocxMathLayout.TryGetRuns(invalid, CancellationToken.None, out _) &&
            !DocxMathLayout.TryMeasure(invalid, _ => resource, CancellationToken.None, out var rejected, out _, resolveControlMetrics: _ => font) && rejected is null,
            "An unsupported late matrix rejects the whole mixed row without publishing a partial prefix or first grid.");
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryGetRuns(equation, cancelled.Token, out _));
        using var scope = OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxSceneNodesPerConversion = 4 });
        TestAssert.Throws<OoxPdfLimitExceededException>(() => DocxMathLayout.TryMeasure(equation, _ => resource, CancellationToken.None,
            out _, out _, resolveControlMetrics: _ => font));
    }

    public static void MathFontMatrixColumnSpacingSharesSlotsAndRetainsSourceUnits()
    {
        var font = MathFont(); var resource = Resource("Spacing", font, "BC");
        var matrix = Matrix([Run("B"), Run("CC")], [Run("BBB"), Run("C")]) with { ColumnGroups = [new(1, "left", false), new(1, "right", false)] };
        double b = font.GetAdvanceWidth(font.MapCodePoint('B')) * 12d / font.UnitsPerEm;
        double c = font.GetAdvanceWidth(font.MapCodePoint('C')) * 12d / font.UnitsPerEm;
        foreach (double size in new[] { 12d, 20d })
        foreach (uint rule in new uint[] { 0, 1, 2, 3, 4 })
        {
            uint amount = rule == 3 ? 120u : rule == 4 ? 3u : 777u;
            var declared = matrix with { MinimumColumnWidthTwips = 480, ColumnGapRuleValue = rule, ColumnGapValue = amount,
                ControlText = new("", size, null, false, false, false, null, "TestFont") };
            TestAssert.True(DocxMathLayout.TryMeasure(Equation(declared), _ => resource, CancellationToken.None, out var box, out var reason,
                resolveControlMetrics: _ => font), reason ?? "Expected authored column spacing.");
            double gap = rule == 3 ? 6d : rule == 4 ? size * 1.5d : size * (1d + rule / 2d);
            double first = Math.Max(24d, 3d * b), second = Math.Max(24d, 2d * c); var grid = TestAssert.NotNull(box);
            TestAssert.True(Math.Abs(grid.AdvanceWidth - first - second - gap) < 1e-9 && grid.Glyphs[0].X == 0d &&
                Math.Abs(grid.Glyphs[1].X - first - gap - second + 2d * c) < 1e-9 &&
                Math.Abs(grid.Glyphs[^1].X - first - gap - second + c) < 1e-9,
                "Point-based minimum widths and exact gaps stay fixed while em gaps follow the control size; grouped columns share those slots.");
        }
        var zero = matrix with { ColumnGapRuleValue = 3, ColumnGapValue = 0 };
        TestAssert.True(DocxMathLayout.TryMeasure(Equation(zero), _ => resource, CancellationToken.None, out var touching, out _, resolveControlMetrics: _ => font) &&
            Math.Abs(TestAssert.NotNull(touching).AdvanceWidth - 3d * b - 2d * c) < 1e-9, "An exact zero gap retains adjacent logical columns.");
        var centered = matrix with { MinimumColumnWidthTwips = 480, ColumnGroups = null };
        TestAssert.True(DocxMathLayout.TryMeasure(Equation(centered), _ => resource, CancellationToken.None, out var centeredBox, out _, resolveControlMetrics: _ => font) &&
            Math.Abs(TestAssert.NotNull(centeredBox).AdvanceWidth - Math.Max(24d, 3d * b) - 12d - (Math.Max(24d, 2d * c) + 2d * c) / 2d) < 1e-9,
            "A widened centered final slot contributes its logical argument edge while unused trailing padding cannot displace the centered formula.");
        foreach (var invalid in new[] { matrix with { ColumnGapRuleValue = 5 }, matrix with { ColumnGapValue = 65536 } })
            TestAssert.True(!DocxMathLayout.TryGetRuns(Equation(invalid), CancellationToken.None, out _), "Invalid typed spacing cannot silently select the default profile.");
        var wide = matrix with { MinimumColumnWidthTwips = uint.MaxValue };
        TestAssert.True(DocxMathLayout.TryMeasure(Equation(wide), _ => resource, CancellationToken.None, out var enormous, out _, resolveControlMetrics: _ => font) &&
            TestAssert.NotNull(enormous).Glyphs.Count == 7 && double.IsFinite(TestAssert.NotNull(enormous).AdvanceWidth),
            "A format-valid large width changes only scalar geometry; it cannot allocate additional slots or glyphs.");
    }

    public static void MathFontMatrixHiddenEmptyCellsKeepSlotsRowFramesAndFirstPaintedOwner()
    {
        var font = MathFont(withPlaceholder: true); var resource = Resource("Sparse", font, "BC"); var empty = new DocxMathRow([]);
        var b = new DocxMathRow([Run("B")]); var c = new DocxMathRow([Run("C")]);
        var matrix = new DocxMathMatrix([new([empty, b, empty]), new([c, empty, empty])]) { HidePlaceholdersValue = true };
        TestAssert.True(DocxMathLayout.TryGetRuns(Equation(matrix), CancellationToken.None, out var runs) && runs.Count == 2 &&
            ReferenceEquals(runs[0].ControlOwner, matrix) && runs[1].ControlOwner is null,
            "An initial empty cell cannot take over the metric-only owner from the first painted run.");
        TestAssert.True(DocxMathLayout.TryMeasure(Equation(matrix), _ => resource, CancellationToken.None, out var box, out var reason,
            resolveControlMetrics: _ => font), reason ?? "Expected sparse grid.");
        var grid = TestAssert.NotNull(box); double bWidth = font.GetAdvanceWidth(font.MapCodePoint('B')) * 12d / font.UnitsPerEm;
        double cWidth = font.GetAdvanceWidth(font.MapCodePoint('C')) * 12d / font.UnitsPerEm;
        double placeholderWidth = font.GetAdvanceWidth(font.MapCodePoint(0x2B1A)) * 12d / font.UnitsPerEm;
        double firstWidth = Math.Max(cWidth, placeholderWidth), secondWidth = Math.Max(bWidth, placeholderWidth);
        TestAssert.True(grid.Glyphs.Count == 2 && grid.Rules.Count == 0 && Math.Abs(grid.AdvanceWidth - firstWidth - secondWidth - placeholderWidth - 24d) < 1e-9 &&
            Math.Abs(grid.Glyphs[0].X - firstWidth - 12d - (secondWidth - bWidth) / 2d) < 1e-9 &&
            Math.Abs(grid.Glyphs[1].X - (firstWidth - cWidth) / 2d) < 1e-9,
            "Hidden cells retain metric-only dotted-square widths, including an empty column and default inter-column gaps.");
        var separated = matrix with { Rows = [new([b, c]), new([empty, empty]), new([c, b])] };
        TestAssert.True(DocxMathLayout.TryMeasure(Equation(separated), _ => resource, CancellationToken.None, out box, out reason,
            resolveControlMetrics: _ => font), reason ?? "Expected empty internal row.");
        grid = TestAssert.NotNull(box);
        TestAssert.True(grid.Glyphs.Count == 4 && Math.Abs(grid.Glyphs[0].BaselineY - grid.Glyphs[^1].BaselineY -
            2d * DocxLineMetrics.MeasureTypographicLineHeight(font, 12d)) < 1e-9,
            "A hidden empty row retains one control-owned row frame without publishing placeholder glyphs.");
    }

    public static void MathFontMatrixHiddenEmptyCellsPreserveVisibilityFallbackCancellationAndWork()
    {
        var font = MathFont(withPlaceholder: true); var resource = Resource("Sparse", font, "B"); var empty = new DocxMathRow([]); var literal = new DocxMathRow([Run("B")]);
        var matrix = new DocxMathMatrix([new([empty, empty, literal])]) { HidePlaceholdersValue = true };
        DocxMathBox? reference = null;
        foreach (bool? visibility in new bool?[] { null, false, true })
        {
            var known = matrix with { HidePlaceholdersValue = visibility };
            TestAssert.True(DocxMathLayout.TryGetRuns(Equation(known), CancellationToken.None, out _),
                "Known placeholder visibility profiles retain source discovery.");
            TestAssert.True(DocxMathLayout.TryMeasure(Equation(known), _ => resource, CancellationToken.None, out var measured, out _,
                resolveControlMetrics: _ => font), "Known placeholder visibility profiles retain the printed grid.");
            var box = TestAssert.NotNull(measured);
            TestAssert.Equal(1, box.Glyphs.Count);
            if (reference is not null)
                TestAssert.True(reference.Width == box.Width && reference.AboveBaseline == box.AboveBaseline &&
                    reference.BelowBaseline == box.BelowBaseline && reference.Glyphs[0].X == box.Glyphs[0].X &&
                    reference.Glyphs[0].BaselineY == box.Glyphs[0].BaselineY,
                    "Known visibility settings preserve the independent Word-equivalent empty slots.");
            reference = box;
        }
        foreach (var invalid in new[] { matrix with { Rows = [new([empty with { PropertiesXml = ["<future/>"] }, empty, literal])] } })
            TestAssert.True(!DocxMathLayout.TryGetRuns(Equation(invalid), CancellationToken.None, out _) &&
                !DocxMathLayout.TryMeasure(Equation(invalid), _ => resource, CancellationToken.None, out var absent, out _, resolveControlMetrics: _ => font) && absent is null,
                "Unknown empty-cell properties retain complete fallback.");
        var allEmpty = Equation(matrix with { Rows = [new([empty, empty, empty])] });
        TestAssert.True(DocxMathLayout.TryGetRuns(allEmpty, default, out var emptySources) && emptySources.Count == 0 &&
            DocxMathLayout.TryMeasure(allEmpty, _ => throw new InvalidOperationException("No literal font source exists."),
                default, out var emptyBox, out _, resolveControlMetrics: _ => font) && emptyBox is { Glyphs.Count: 0, Width: > 0d },
            "All-empty slots use their actual control/paragraph metrics without inventing painted source runs.");
        TestAssert.True(!DocxMathLayout.TryMeasure(allEmpty, _ => resource, default, out _, out _),
            "An all-empty grid still requires a metric owner, not a fabricated literal font lookup.");
        TestAssert.True(!DocxMathLayout.TryMeasure(Equation(matrix), _ => resource, CancellationToken.None, out var missing, out _,
            resolveControlMetrics: _ => MathFont()) && missing is null, "Missing placeholder metrics cannot invent an empty width or publish partial cells.");
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryGetRuns(Equation(matrix), cancelled.Token, out _));
        using var scope = OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxSceneNodesPerConversion = 1 });
        TestAssert.Throws<OoxPdfLimitExceededException>(() => DocxMathLayout.TryMeasure(Equation(matrix), _ => resource,
            CancellationToken.None, out _, out _, resolveControlMetrics: _ => font));
    }

    public static void MathFontMatrixStructuredCellsTranslateChildRulesAndKeepControlMetrics()
    {
        var font = MathFont(); var resource = Resource("Structured", font, "BC");
        var basis = new DocxMathRow([Run("B")]); var argument = new DocxMathRow([Run("C")]);
        foreach (bool display in new[] { true, false })
        foreach (double size in new[] { 12d, 20d })
        foreach (DocxMathNode structure in new DocxMathNode[] { new DocxMathFraction(basis, argument, "bar"),
            new DocxMathFraction(basis, argument, "noBar"), new DocxMathScripts(basis, argument, argument) })
        {
            var cell = new DocxMathRow([structure]);
            var matrix = Matrix([Run("B"), Run("C")], [Run("C"), Run("B")]) with
            { Rows = [new([basis, cell]), new([argument, basis])], ControlText = new("", size, null, false, false, false, null, "TestFont") };
            var equation = Equation(matrix) with { Display = display };
            TestAssert.True(DocxMathLayout.TryGetRuns(equation, CancellationToken.None, out var sources) &&
                ReferenceEquals(sources[0].ControlOwner, matrix) && sources.Skip(1).All(s => s.ControlOwner is null),
                "Bounded structured cells retain one metric-only matrix owner and ordered child source roles.");
            TestAssert.True(DocxMathLayout.TryMeasure(equation, _ => resource, CancellationToken.None, out var box, out var reason,
                resolveControlMetrics: _ => font), reason ?? "Expected bounded structured grid.");
            TestAssert.True(DocxMathLayout.TryMeasure(equation with { Content = cell, ParentFontSize = size }, _ => resource,
                CancellationToken.None, out var inner, out reason, parentMetricsFont: font), reason ?? "Expected existing child box.");
            var grid = TestAssert.NotNull(box); var child = TestAssert.NotNull(inner);
            double x = grid.Glyphs[1].X - child.Glyphs[0].X, y = grid.Glyphs[1].BaselineY - child.Glyphs[0].BaselineY;
            TestAssert.True(x > 0d && child.Glyphs.Select((g, i) => (g, i)).All(pair =>
                Math.Abs(grid.Glyphs[pair.i + 1].X - pair.g.X - x) < 1e-9 &&
                Math.Abs(grid.Glyphs[pair.i + 1].BaselineY - pair.g.BaselineY - y) < 1e-9),
                "Existing child geometry moves as a whole into the matrix slot, including script baselines.");
            TestAssert.True(grid.Rules.Count == child.Rules.Count && grid.Rules.Zip(child.Rules).All(pair =>
                Math.Abs(pair.First.X - pair.Second.X - x) < 1e-9 && Math.Abs(pair.First.Y - pair.Second.Y - y) < 1e-9 &&
                pair.First.Thickness == pair.Second.Thickness), "Fraction rules retain child geometry and move with the same cell origin.");
            TestAssert.True(grid.ParentMetrics?.Size == size && grid.Glyphs[0].FontSize == 12d,
                "Control metrics remain independent of painted cell sizes.");
        }
    }

    public static void MathFontMatrixStructuredCellsKeepDeeperAndLateFailureFallback()
    {
        var font = MathFont(); var resource = Resource("Grid", font, "BC");
        var fraction = new DocxMathFraction(new([Run("B")]), new([Run("C")]), "bar");
        var matrix = Matrix([Run("B"), Run("C")], [Run("C"), Run("B")]);
        foreach (DocxMathNode invalid in new DocxMathNode[] { fraction with { Numerator = new([fraction]) },
            fraction with { TypeValue = "future" }, new DocxMathScripts(new([matrix]), null, new([Run("B")])) })
        {
            var bad = matrix with { Rows = [matrix.Rows[0], new([new([Run("B")]), new([invalid])])] };
            TestAssert.True(!DocxMathLayout.TryGetRuns(Equation(bad), CancellationToken.None, out _) &&
                !DocxMathLayout.TryMeasure(Equation(bad), _ => resource, CancellationToken.None, out var absent, out _, resolveControlMetrics: _ => font) && absent is null,
                "Deeper, unknown and unsupported cell roles reject the complete grid before publication.");
        }
        var last = Run("C"); var structured = matrix with { Rows = [new([new([fraction]), new([Run("B")])]), new([new([Run("B")]), new([last])])] };
        var missing = Resource("Missing", font, "B");
        TestAssert.True(!DocxMathLayout.TryMeasure(Equation(structured), r => ReferenceEquals(r, last) ? missing : resource,
            CancellationToken.None, out var rejected, out _, resolveControlMetrics: _ => font) && rejected is null,
            "A late missing glyph cannot publish an earlier fractional cell or its rule.");
    }

    public static void MathFontMatrixBaseJustificationAnchorsRowsWithoutChangingSlots()
    {
        var font = MathFont(); var resource = Resource("Aligned", font, "BC");
        var matrix = Matrix([Run("B"), Run("CC")], [Run("BBB"), Run("C")]);
        matrix = matrix with { Rows = [matrix.Rows[0], matrix.Rows[1], matrix.Rows[0]],
            ColumnGroups = [new(1, "left", false), new(1, "right", false)] };
        DocxMathBox Measure(string? alignment)
        {
            TestAssert.True(DocxMathLayout.TryMeasure(Equation(matrix with { BaseJustificationValue = alignment }), _ => resource,
                CancellationToken.None, out var box, out var reason, resolveControlMetrics: _ => font), reason ?? "Expected matrix baseline alignment.");
            return TestAssert.NotNull(box);
        }
        var center = Measure(null); var explicitCenter = Measure("center"); var top = Measure("top"); var bottom = Measure("bot");
        TestAssert.True(Math.Abs(top.Glyphs[0].BaselineY) < 1e-9 && Math.Abs(bottom.Glyphs[^1].BaselineY) < 1e-9,
            "Top and bottom anchor the first and last row baselines, including grids with more than two rows.");
        foreach (var aligned in new[] { explicitCenter, top, bottom })
        {
            double shift = aligned.Glyphs[0].BaselineY - center.Glyphs[0].BaselineY;
            TestAssert.True(aligned.Width == center.Width && aligned.AdvanceWidth == center.AdvanceWidth &&
                aligned.Glyphs.Zip(center.Glyphs).All(pair => pair.First.X == pair.Second.X &&
                    Math.Abs(pair.First.BaselineY - pair.Second.BaselineY - shift) < 1e-9),
                "Base justification translates the complete grid while preserving row distances, grouped column slots and logical width.");
        }
        TestAssert.True(explicitCenter.Glyphs.SequenceEqual(center.Glyphs), "Explicit and omitted center retain identical geometry.");
        foreach (string invalid in new[] { "bottom", "future", "" })
            TestAssert.True(!DocxMathLayout.TryGetRuns(Equation(matrix with { BaseJustificationValue = invalid }), CancellationToken.None, out _),
                "Unknown typed base values retain complete fallback rather than silently centering.");
    }

    public static void MathFontMatrixColumnGroupsShareSlotsAndBoundDeclaredCounts()
    {
        var font = MathFont(); var resource = Resource("Columns", font, "BC");
        var matrix = Matrix([Run("B"), Run("CC")], [Run("BBB"), Run("C")]);
        double b = font.GetAdvanceWidth(font.MapCodePoint('B')) * 12d / font.UnitsPerEm;
        double c = font.GetAdvanceWidth(font.MapCodePoint('C')) * 12d / font.UnitsPerEm;
        foreach (string alignment in new[] { "left", "center", "right" })
        {
            var declared = matrix with { ColumnGroups = [new(2, alignment, false)] };
            TestAssert.True(DocxMathLayout.TryMeasure(Equation(declared), _ => resource, CancellationToken.None, out var box, out var reason,
                resolveControlMetrics: _ => font), reason ?? "Expected declared column group.");
            var measured = TestAssert.NotNull(box); double factor = alignment == "left" ? 0d : alignment == "right" ? 1d : 0.5d;
            TestAssert.True(Math.Abs(measured.Glyphs[0].X - 2d * b * factor) < 1e-9 &&
                Math.Abs(measured.Glyphs[^1].X - 3d * b - 12d - c * factor) < 1e-9,
                "One declared group applies the same alignment to each actual logical column slot.");
        }
        var grouped = matrix with { ColumnGroups = [new(null, "left", false), new(null, "right", false)] };
        TestAssert.True(DocxMathLayout.TryMeasure(Equation(grouped), _ => resource, CancellationToken.None, out var mixed, out _, resolveControlMetrics: _ => font) &&
            TestAssert.NotNull(mixed).Glyphs[0].X == 0d && Math.Abs(TestAssert.NotNull(mixed).Glyphs[^1].X - 3d * b - 12d - c) < 1e-9,
            "Absent counts default to one while grouped columns retain independent alignment.");
        foreach (IReadOnlyList<DocxMathMatrixColumn> invalid in new IReadOnlyList<DocxMathMatrixColumn>[] {
            [], [new(1, "left", false)], [new(3, "left", false)], [new(int.MaxValue, "left", false)],
            [new(0, "left", false)], [new(2, "future", false)], [new(2, "left", true)] })
            TestAssert.True(!DocxMathLayout.TryMeasure(Equation(matrix with { ColumnGroups = invalid }), _ => resource, CancellationToken.None,
                out var rejected, out _, resolveControlMetrics: _ => font) && rejected is null,
                "Invalid/incomplete/overdeclared groups reject the complete grid before any count-driven expansion.");
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryGetRuns(Equation(grouped), cancelled.Token, out _));
    }

    public static void MathFontMatrixLiteralCellsShareSlotsAxisAndMetricOnlyControl()
    {
        var metric = MathFont(); var plain = TestFontBuilder.LoadTestFont();
        var first = Resource("First", plain, "BC"); var rest = Resource("Rest", plain, "BC");
        var a = Run("B", "0000FF"); var b = Run("CC"); var c = Run("BBB", "008000"); var d = Run("C");
        foreach (double size in new[] { 12d, 20d })
        {
            var matrix = Matrix([a, b], [c, d]) with { ControlText = new("", size, "FF0000", false, false, false, null, "TestFont") };
            TestAssert.True(DocxMathLayout.TryGetRuns(Equation(matrix), CancellationToken.None, out var sources) && sources.Count == 4 &&
                ReferenceEquals(sources[0].ControlOwner, matrix) && sources.Skip(1).All(s => s.ControlOwner is null),
                "One invisible owner retains ordered cell source without taking over painted fonts.");
            TestAssert.True(DocxMathLayout.GetControlCodePoints(matrix).Count == 0, "A matrix control requests metrics without glyph mappings.");
            TestAssert.True(DocxMathLayout.TryMeasure(Equation(matrix), r => ReferenceEquals(r, a) ? first : rest, CancellationToken.None,
                out var box, out var reason, resolveControlMetrics: _ => metric), reason ?? "Expected literal grid.");
            var measured = TestAssert.NotNull(box); double advance = plain.GetAdvanceWidth(plain.MapCodePoint('B')) * 12d / plain.UnitsPerEm;
            double otherAdvance = plain.GetAdvanceWidth(plain.MapCodePoint('C')) * 12d / plain.UnitsPerEm;
            TestAssert.Equal("BCCBBBC", string.Concat(measured.Glyphs.Select(g => char.ConvertFromUtf32(g.SourceCodePoint))));
            TestAssert.True(Math.Abs(measured.Glyphs[0].X - advance) < 1e-9 && measured.Glyphs[3].X == 0d &&
                Math.Abs(measured.Glyphs[1].X - 3d * advance - size) < 1e-9 &&
                Math.Abs(measured.Glyphs[6].X - 3d * advance - otherAdvance / 2d - size) < 1e-9,
                "Each column centers its complete logical cell width and consumes one control-sized em between slots.");
            double export = OfficePdfTextEmissionProfile.FontSize(size);
            double line = DocxLineMetrics.MeasureTypographicLineHeight(metric, export), above = DocxLineMetrics.MeasureTypographicBaselineInset(metric, export);
            TestAssert.True(Math.Abs(measured.Glyphs[0].BaselineY - measured.Glyphs[3].BaselineY - line) < 1e-9,
                "Single row spacing reuses typographic frames instead of independent per-cell baseline corrections.");
            double frameTop = measured.Glyphs[0].BaselineY + above, frameBottom = measured.Glyphs[3].BaselineY - (line - above);
            TestAssert.True(Math.Abs((frameTop + frameBottom) / 2d - size / 4d) < 1e-9,
                "The whole grid frame centers on the metric font's mathematical axis.");
            TestAssert.True(measured.ParentMetrics is { IncludeTypographicFrame: false } owner && ReferenceEquals(owner.Font, metric) && owner.Size == size &&
                ReferenceEquals(measured.Glyphs[0].Font, first) && measured.Glyphs[0].Color.Blue == 255 &&
                measured.Glyphs.Skip(3).Take(3).All(g => ReferenceEquals(g.Font, rest) && g.Color.Green == 128) && measured.Rules.Count == 0,
                "Internal control frames, source paint and occupied outer ink retain distinct ownership, without implicit delimiters.");
            var graphics = new PdfGraphicsBuilder(); DocxMathBoxPainter.Paint(measured, graphics, 20d, 80d, CancellationToken.None);
            using var output = new MemoryStream();
            PdfDocumentWriter.WriteBlank(output, [new PdfPage(250d, 200d, graphics.ToString(), [first, rest])], CancellationToken.None);
            TestAssert.True(output.Length > 0, "Only painted cell resources are required; the metric font needs no empty subset.");
        }
    }

    public static void MathFontMatrixDefaultControlUsesParentFrameAndRetainsIndependentCellSizes()
    {
        var font = MathFont(); var resource = Resource("Cells", font, "BC");
        var first = Run("B") with { Text = Run("B").Text with { FontSize = 18d } };
        var matrix = Matrix([first, Run("C")], [Run("C"), Run("B")]);
        TestAssert.True(DocxMathLayout.TryMeasure(Equation(matrix), _ => resource, CancellationToken.None, out var box, out var reason,
            resolveControlMetrics: _ => font), reason ?? "Expected surrounding control inheritance.");
        var measured = TestAssert.NotNull(box);
        TestAssert.True(measured.ParentMetrics?.Size == 12d && measured.Glyphs[0].FontSize == 18d && measured.Glyphs.Skip(1).All(g => g.FontSize == 12d),
            "An absent matrix control uses the surrounding size while every painted cell keeps its authored size.");
    }

    public static void MathFontMatrixFallbackCancellationAndWorkRejectCompleteGrids()
    {
        var font = MathFont(); var resource = Resource("Grid", font, "BC"); var last = Run("C");
        var matrix = Matrix([Run("B"), Run("C")], [Run("B"), last]);
        var missing = Resource("MissingLast", font, "B");
        TestAssert.True(!DocxMathLayout.TryMeasure(Equation(matrix), r => ReferenceEquals(r, last) ? missing : resource, CancellationToken.None,
            out var absent, out _, resolveControlMetrics: _ => font) && absent is null,
            "A late missing cell encoding cannot publish earlier cells or rows.");
        foreach (var invalid in new[] { matrix with { HasUnsupportedProperties = true }, matrix with { Rows = [matrix.Rows[0], new([new([Run("B")])])] },
            matrix with { Rows = [new([new([]), new([Run("B")])])] },
            matrix with { Rows = [new([new([new DocxMathUnsupported("future", "test", "<future/>", [])])])] },
            matrix with { ControlText = new("", double.NaN, null, false, false, false, null, "TestFont") } })
            TestAssert.True(!DocxMathLayout.TryMeasure(Equation(invalid), _ => resource, CancellationToken.None, out var rejected, out _,
                resolveControlMetrics: _ => font) && rejected is null, "Unsupported rows, empty placeholders, properties and invalid metrics retain complete fallback.");
        TestAssert.True(!DocxMathLayout.TryMeasure(Equation(matrix), _ => resource, CancellationToken.None, out absent, out _,
            resolveControlMetrics: _ => TestFontBuilder.LoadTestFont()) && absent is null, "A missing mathematical metric owner cannot publish a default-shaped grid.");
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryGetRuns(Equation(matrix), cancelled.Token, out _));
        TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryMeasure(Equation(matrix), _ => resource, cancelled.Token, out _, out _));
        using var scope = OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxSceneNodesPerConversion = 3 });
        TestAssert.Throws<OoxPdfLimitExceededException>(() => DocxMathLayout.TryMeasure(Equation(matrix), _ => resource, CancellationToken.None, out _, out _, resolveControlMetrics: _ => font));
    }

    private static DocxMathRun Run(string text, string? color = null) => new(new(text, 12d, color, false, false, false, null, "TestFont"), "p", null) { NormalTextValue = true };
    private static DocxMathMatrix Matrix(DocxMathRun[] first, DocxMathRun[] second) => new([new(first.Select(r => new DocxMathRow([r])).ToArray()), new(second.Select(r => new DocxMathRow([r])).ToArray())]);
    private static DocxEquation Equation(DocxMathMatrix matrix) => new(new([matrix]), 0, 0, true, null, null, null) { ParentFontSize = 12d };
    private static PdfFontResource Resource(string name, OpenTypeFont font, string text) => new(name, PdfEmbeddedFont.Create(font, text.Select(c => (int)c), CancellationToken.None));
    private static OpenTypeFont MathFont(bool withPlaceholder = false)
    {
        byte[] table = new byte[224]; BinaryPrimitives.WriteUInt16BigEndian(table, 1); BinaryPrimitives.WriteUInt16BigEndian(table.AsSpan(4), 10);
        BinaryPrimitives.WriteUInt16BigEndian(table.AsSpan(10), 73); BinaryPrimitives.WriteUInt16BigEndian(table.AsSpan(12), 50);
        BinaryPrimitives.WriteInt16BigEndian(table.AsSpan(18 + (int)OpenTypeMathConstant.AxisHeight * 4), 250);
        BinaryPrimitives.WriteInt16BigEndian(table.AsSpan(18 + (int)OpenTypeMathConstant.MathLeading * 4), 120);
        void Set(OpenTypeMathConstant field, short value) => BinaryPrimitives.WriteInt16BigEndian(table.AsSpan(18 + (int)field * 4), value);
        Set(OpenTypeMathConstant.FractionRuleThickness, 40); Set(OpenTypeMathConstant.SuperscriptShiftUp, 500);
        Set(OpenTypeMathConstant.FractionNumeratorDisplayStyleShiftUp, 700); Set(OpenTypeMathConstant.FractionDenominatorDisplayStyleShiftDown, 500);
        Set(OpenTypeMathConstant.FractionNumeratorShiftUp, 500); Set(OpenTypeMathConstant.FractionDenominatorShiftDown, 400);
        Set(OpenTypeMathConstant.StackTopDisplayStyleShiftUp, 700); Set(OpenTypeMathConstant.StackBottomDisplayStyleShiftDown, 500);
        Set(OpenTypeMathConstant.StackTopShiftUp, 500); Set(OpenTypeMathConstant.StackBottomShiftDown, 400);
        byte[] bytes = TestFontBuilder.CreateMathFont(table);
        if (withPlaceholder)
        {
            var cmap = TestFontBuilder.GetTableRange(bytes, "cmap");
            int format12 = cmap.Offset + (int)BinaryPrimitives.ReadUInt32BigEndian(bytes.AsSpan(cmap.Offset + 16));
            // Replace the unused accent singleton with a dotted-square mapping
            // to the existing B outline; production font coverage is unchanged.
            BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(format12 + 28), 0x2B1A);
            BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(format12 + 32), 0x2B1A);
            BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(format12 + 36), TestFontBuilder.LoadTestFont().MapCodePoint('B'));
        }
        return OpenTypeFont.Load(bytes);
    }
}
