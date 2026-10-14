using System.Buffers.Binary;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class MathScriptStyleTests
{
    public static void MathFontNormalScriptChildrenUseParentScaleWithoutChildMathTables()
    {
        byte[] constants = new byte[224];
        BinaryPrimitives.WriteUInt16BigEndian(constants, 1);
        BinaryPrimitives.WriteUInt16BigEndian(constants.AsSpan(4), 10);
        BinaryPrimitives.WriteInt16BigEndian(constants.AsSpan(10), 73);
        OpenTypeFont parentFont = OpenTypeFont.Load(TestFontBuilder.CreateMathFont(constants));
        OpenTypeFont textFont = TestFontBuilder.LoadTestFont();
        var parentResource = new PdfFontResource("Math1", PdfEmbeddedFont.Create(parentFont, ['A'], CancellationToken.None));
        var childResource = new PdfFontResource("Text1", PdfEmbeddedFont.Create(textFont, ['B'], CancellationToken.None));
        var argument = new DocxMathRow([new DocxMathRun(new("B", 12d, null, false, false, false, null, "TestFont"), null, null)
            { NormalTextValue = true }]);
        TestAssert.True(!DocxMathRowLayout.TryMeasure(argument, _ => childResource, CancellationToken.None,
            out _, out _, scriptLevel: 1), "An unowned script scale cannot be guessed from a text face without MATH.");
        TestAssert.True(DocxMathRowLayout.TryMeasure(argument, _ => childResource, CancellationToken.None,
            out var box, out _, scriptLevel: 1, scriptMetricsFont: parentFont),
            "The parent math face can own the script scale while the child owns ordinary text glyphs.");
        DocxMathGlyphPlacement glyph = TestAssert.NotNull(box).Glyphs.Single();
        TestAssert.True(ReferenceEquals(glyph.Font, childResource) && Math.Abs(glyph.FontSize - 8.52d) < 1e-9,
            "Parent scaling cannot replace the child's text face or subset ownership.");
        var basis = new DocxMathRow([new DocxMathRun(new("A", 12d, null, false, false, false, null, "TestFont"), "p", null)]);
        var equation = new DocxEquation(new([new DocxMathScripts(basis, argument, null)]), 0, 0, true, null, null, null);
        TestAssert.True(DocxMathLayout.TryMeasure(equation, run => run.NormalText ? childResource : parentResource,
            CancellationToken.None, out var composed, out _), "Mixed normal-text child faces share the parent script geometry.");
        TestAssert.True(TestAssert.NotNull(composed).Glyphs[1].BaselineY < 0d, "The normal-text child retains its subscript role.");
    }

    public static void MathFontScriptArgumentIncrementsOwnFontStyleAndPreserveParentMetrics()
    {
        byte[] constants = new byte[224];
        BinaryPrimitives.WriteUInt16BigEndian(constants, 1);
        BinaryPrimitives.WriteUInt16BigEndian(constants.AsSpan(4), 10);
        BinaryPrimitives.WriteInt16BigEndian(constants.AsSpan(10), 73);
        BinaryPrimitives.WriteInt16BigEndian(constants.AsSpan(12), 60);
        OpenTypeFont font = OpenTypeFont.Load(TestFontBuilder.CreateMathFont(constants));
        var resource = new PdfFontResource("Math1", PdfEmbeddedFont.Create(font, ['A', 'B'], CancellationToken.None));
        foreach ((int increment, int expectedLevel, double expectedSize) in new[]
            { (-2, 2, 6.96d), (-1, 2, 6.96d), (0, 1, 8.52d), (1, 0, 12d), (2, 0, 12d) })
        {
            string xml = "<m:oMath xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math'><m:sSub>" +
                "<m:e><m:r><m:rPr><m:sty m:val='p'/></m:rPr><m:t>A</m:t></m:r></m:e>" +
                $"<m:sub><m:argPr><m:argSz m:val='{increment}'/></m:argPr>" +
                "<m:r><m:rPr><m:sty m:val='p'/></m:rPr><m:t>B</m:t></m:r></m:sub></m:sSub></m:oMath>";
            DocxMathRow row = DocxMathParser.Read(System.Xml.Linq.XElement.Parse(xml),
                (_, text) => new(text, 12d, null, false, false, false, null, "TestFont"), CancellationToken.None);
            var equation = new DocxEquation(row, 0, 0, true, null, null, null);
            TestAssert.True(DocxMathLayout.TryGetRuns(equation, CancellationToken.None, out var runs),
                "Known source increments must discover the role's actual style.");
            TestAssert.Equal(expectedLevel, runs.Single(source => source.Run.Text.Text == "B").ScriptLevel);
            TestAssert.True(DocxMathLayout.TryMeasure(equation, _ => resource, CancellationToken.None, out var box, out var reason),
                reason ?? "Expected measured script arguments.");
            DocxMathBox measured = TestAssert.NotNull(box);
            TestAssert.True(Math.Abs(measured.Glyphs[0].FontSize - 12d) < 1e-9 &&
                Math.Abs(measured.Glyphs[1].FontSize - expectedSize) < 1e-9,
                "The argument owns its resolved style size while the base retains parent metrics.");
            TestAssert.Equal(font.MapCodePoint('B'), measured.Glyphs[1].OriginalGlyph);
            TestAssert.True(!DocxMathLayout.TryGetRuns(equation with { Content = ((DocxMathScripts)row.Children[0]).Subscript! },
                CancellationToken.None, out _) || increment == 0,
                "Signed increments are admitted only on supported script argument roles.");
        }
    }

    public static void MathFontScriptArgumentsUseSemanticOperatorSpacing()
    {
        byte[] constants = new byte[224];
        BinaryPrimitives.WriteUInt16BigEndian(constants, 1);
        BinaryPrimitives.WriteUInt16BigEndian(constants.AsSpan(4), 10);
        BinaryPrimitives.WriteInt16BigEndian(constants.AsSpan(10), 100);
        OpenTypeFont font = OpenTypeFont.Load(TestFontBuilder.CreateMathFont(constants));
        var resource = new PdfFontResource("Math1", PdfEmbeddedFont.Create(font, "A+B=C".Select(c => (int)c), CancellationToken.None));
        var row = new DocxMathRow([new DocxMathRun(new("A+B=C", 12d, null, false, false, false, null, "TestFont"), "p", null)]);
        TestAssert.True(DocxMathRowLayout.TryMeasure(row, _ => resource, CancellationToken.None, out DocxMathBox? normal, out _), "Expected ordinary math row.");
        TestAssert.True(DocxMathRowLayout.TryMeasure(row, _ => resource, CancellationToken.None, out DocxMathBox? script, out _, scriptLevel: 1),
            "Expected script-style argument row.");
        DocxMathBox argument = TestAssert.NotNull(script), ordinary = TestAssert.NotNull(normal);
        TestAssert.True(argument.Width < ordinary.Width, "Script-style binary/relation operators do not inherit ordinary-row glue.");
        TestAssert.True(Math.Abs(argument.Glyphs[1].X - argument.Glyphs[0].Advance) < 1e-9,
            "The argument's first binary operator follows the stored glyph advance without ordinary-style glue.");
        TestAssert.True(argument.Glyphs.Select(g => g.SourceCodePoint).SequenceEqual(ordinary.Glyphs.Select(g => g.SourceCodePoint)),
            "The semantic style changes spacing while preserving mathematical source characters.");
    }

    public static void MathFontCornerKernSignedThresholdsHonorEqualityAndAbsentCorners()
    {
        foreach (bool ranges in new[] { false, true })
        {
            OpenTypeFont font = OpenTypeFont.Load(TestFontBuilder.CreateMathFont(KernTable(ranges)));
            foreach ((double height, short expected) in new[] { (-101d, (short)-70), (-100d, (short)20), (499.99d, (short)20), (500d, (short)90) })
            {
                TestAssert.True(font.TryGetMathKern(34, MathKernCorner.TopRight, height, out short actual),
                    "Signed corner corrections must resolve at a bounded font-relative height.");
                TestAssert.Equal(expected, actual);
            }
            TestAssert.True(font.TryGetMathKern(34, MathKernCorner.BottomLeft, 0d, out short absent) && absent == 0,
                "Absent corners contribute zero to the paired-height sum.");
            TestAssert.True(font.TryGetMathKern(35, MathKernCorner.TopRight, 0d, out absent) && absent == 0,
                "Uncovered glyphs contribute zero without invalidating ordinary metrics.");
        }
    }

    public static void MathFontCornerKernTruncationBadOffsetsAndNonfiniteHeightsKeepFallback()
    {
        byte[] complete = KernTable(false);
        for (int length = 1; length < complete.Length; length++)
        {
            OpenTypeFont font = OpenTypeFont.Load(TestFontBuilder.CreateMathFont(complete[..length]));
            TestAssert.True(!font.TryGetMathKern(34, MathKernCorner.TopRight, 0d, out _),
                "Truncated selected corner data cannot publish a correction.");
            TestAssert.True(font.MapCodePoint('A') != 0, "Optional MATH corner defects preserve ordinary face use.");
        }
        foreach (int offset in new[] { 6, 16, 18, 22, 36 })
        {
            byte[] invalid = (byte[])complete.Clone();
            BinaryPrimitives.WriteUInt16BigEndian(invalid.AsSpan(offset), ushort.MaxValue);
            OpenTypeFont font = OpenTypeFont.Load(TestFontBuilder.CreateMathFont(invalid));
            TestAssert.True(!font.TryGetMathKern(34, MathKernCorner.TopRight, 0d, out _), "Malformed selected offsets/counts keep fallback.");
        }
        OpenTypeFont valid = OpenTypeFont.Load(TestFontBuilder.CreateMathFont(complete));
        TestAssert.True(!valid.TryGetMathKern(34, MathKernCorner.TopRight, double.NaN, out _), "Nonfinite query heights keep fallback.");
        TestAssert.True(TestFontBuilder.LoadTestFont().TryGetMathKern(34, MathKernCorner.TopRight, 0d, out short missing) && missing == 0,
            "A missing MATH table contributes no corner kerning.");
    }

    private static byte[] KernTable(bool ranges)
    {
        int kern = ranges ? 40 : 36;
        byte[] data = new byte[kern + 22];
        void U(int at, ushort value) => BinaryPrimitives.WriteUInt16BigEndian(data.AsSpan(at), value);
        void I(int at, short value) => BinaryPrimitives.WriteInt16BigEndian(data.AsSpan(at), value);
        U(0, 1); U(6, 10); U(16, 8); U(18, 12); U(20, 1); U(22, (ushort)(kern - 18));
        U(30, ranges ? (ushort)2 : (ushort)1); U(32, 1); U(34, 34);
        if (ranges) { U(36, 34); U(38, 0); }
        U(kern, 2); I(kern + 2, -100); I(kern + 6, 500);
        I(kern + 10, -70); I(kern + 14, 20); I(kern + 18, 90);
        return data;
    }

    public static void MathFontExtendedShapeCoverageDistinguishesGlyphAndBoxOwners()
    {
        foreach (bool ranges in new[] { false, true })
        {
            byte[] table = new byte[ranges ? 28 : 24];
            void U(int at, ushort value) => BinaryPrimitives.WriteUInt16BigEndian(table.AsSpan(at), value);
            U(0, 1); U(6, 10); U(14, 8); U(18, ranges ? (ushort)2 : (ushort)1); U(20, 1); U(22, 34);
            if (ranges) { U(24, 34); U(26, 0); }
            OpenTypeFont font = OpenTypeFont.Load(TestFontBuilder.CreateMathFont(table));
            TestAssert.True(font.TryIsMathExtendedShape(34, out bool extended) && extended,
                "Covered font glyphs own the extended-shape baseline contract.");
            TestAssert.True(font.TryIsMathExtendedShape(35, out extended) && !extended,
                "Ordinary glyphs cannot inherit composed-box baseline-drop constraints.");
            for (int length = 1; length < table.Length; length++)
            {
                font = OpenTypeFont.Load(TestFontBuilder.CreateMathFont(table[..length]));
                TestAssert.True(!font.TryIsMathExtendedShape(34, out _), "Truncated selected coverage keeps explicit fallback.");
                TestAssert.True(font.MapCodePoint('A') != 0, "Optional math-shape defects preserve ordinary face use.");
            }
        }
        TestAssert.True(TestFontBuilder.LoadTestFont().TryIsMathExtendedShape(34, out bool missing) && !missing,
            "Absent MATH coverage keeps ordinary glyph ownership.");
    }

    public static void MathFontScriptAlternatesHonorMathLanguageAndBothCoverageFormats()
    {
        foreach (bool ranges in new[] { false, true })
        foreach (bool extension in new[] { false, true })
        {
            byte[] table = Gsub(ranges, extension);
            OpenTypeMathScriptStyles styles = OpenTypeMathScriptStyles.Read(table, CancellationToken.None);
            TestAssert.True(styles.IsValid, "The selected math/dflt/ssty feature must be readable.");
            TestAssert.True(styles.TryGetGlyph(34, 1, CancellationToken.None, out ushort glyph), "First script style must be readable.");
            TestAssert.Equal((ushort)36, glyph);
            TestAssert.True(styles.TryGetGlyph(34, 2, CancellationToken.None, out glyph), "Second script style must be readable.");
            TestAssert.Equal((ushort)37, glyph);
            TestAssert.True(styles.TryGetGlyph(35, 1, CancellationToken.None, out glyph), "Uncovered glyphs retain the original shape.");
            TestAssert.Equal((ushort)35, glyph);
            "latn"u8.CopyTo(table.AsSpan(12));
            styles = OpenTypeMathScriptStyles.Read(table, CancellationToken.None);
            TestAssert.True(styles.TryGetGlyph(34, 1, CancellationToken.None, out glyph), "A missing math feature must preserve the base glyph.");
            TestAssert.Equal((ushort)34, glyph);
        }
    }

    public static void MathFontScriptTruncationBadOffsetsAndContextualLookupsKeepFallback()
    {
        byte[] complete = Gsub(false, true);
        for (int length = 0; length < complete.Length; length++)
        {
            OpenTypeMathScriptStyles styles = OpenTypeMathScriptStyles.Read(complete.AsMemory(0, length), CancellationToken.None);
            TestAssert.True(!styles.IsValid || !styles.TryGetGlyph(34, 1, CancellationToken.None, out _),
                "A truncated selected feature cannot publish a substituted glyph.");
        }
        foreach ((int offset, ushort value) in new[] { (4, ushort.MaxValue), (28, ushort.MaxValue), (44, ushort.MaxValue),
            (48, (ushort)5), (50, (ushort)8), (60, (ushort)7) })
        {
            byte[] malformed = (byte[])complete.Clone();
            BinaryPrimitives.WriteUInt16BigEndian(malformed.AsSpan(offset), value);
            OpenTypeMathScriptStyles styles = OpenTypeMathScriptStyles.Read(malformed, CancellationToken.None);
            TestAssert.True(!styles.IsValid || !styles.TryGetGlyph(34, 1, CancellationToken.None, out _),
                "Invalid offsets, contextual lookups and recursive extensions retain explicit fallback.");
        }
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => OpenTypeMathScriptStyles.Read(complete, cancelled.Token));
        TestAssert.Throws<OperationCanceledException>(() => OpenTypeMathScriptStyles.Read(complete, CancellationToken.None)
            .TryGetGlyph(34, 1, cancelled.Token, out _));
    }

    public static void MathFontScriptSingleSubstitutionsRetainTheirSignedDeltaAndExplicitMap()
    {
        foreach (bool delta in new[] { false, true })
        {
            byte[] table = Gsub(false, false);
            BinaryPrimitives.WriteUInt16BigEndian(table.AsSpan(48), 1);
            BinaryPrimitives.WriteUInt16BigEndian(table.AsSpan(58), delta ? (ushort)1 : (ushort)2);
            BinaryPrimitives.WriteUInt16BigEndian(table.AsSpan(62), delta ? (ushort)0xFFFF : (ushort)1);
            BinaryPrimitives.WriteUInt16BigEndian(table.AsSpan(64), 38);
            OpenTypeMathScriptStyles styles = OpenTypeMathScriptStyles.Read(table, CancellationToken.None);
            TestAssert.True(styles.TryGetGlyph(34, 2, CancellationToken.None, out ushort glyph),
                "Single-substitution script forms work at both script levels without contextual matching.");
            TestAssert.Equal(delta ? (ushort)33 : (ushort)38, glyph);
        }
    }

    public static void MathFontScriptSizesSeparateMathScaleHalfPointsAndPdfGrid()
    {
        byte[] table = new byte[224];
        BinaryPrimitives.WriteUInt16BigEndian(table, 1);
        BinaryPrimitives.WriteUInt16BigEndian(table.AsSpan(4), 10);
        BinaryPrimitives.WriteInt16BigEndian(table.AsSpan(10), 73);
        OpenTypeFont font = OpenTypeFont.Load(TestFontBuilder.CreateMathFont(table));
        foreach ((double parent, double expected) in new[] { (8d, 5.52d), (12d, 8.52d), (16d, 11.52d), (20d, 14.52d), (24d, 17.52d) })
        {
            TestAssert.True(DocxMathArgumentMetrics.TryGetScriptFontSize(font, parent, 1, out double size),
                "Known font-scale/Word export contrasts must resolve a finite script size.");
            TestAssert.True(Math.Abs(size - expected) < 1e-9, "The resolved script size must match the independent Word contrast.");
        }
        TestAssert.True(!DocxMathArgumentMetrics.TryGetScriptFontSize(TestFontBuilder.LoadTestFont(), 12d, 1, out _),
            "Missing MATH scale cannot claim a native argument size.");
        TestAssert.True(!DocxMathArgumentMetrics.TryGetScriptFontSize(font, double.PositiveInfinity, 1, out _),
            "Nonfinite source sizes keep the whole-object fallback.");
    }

    public static void MathFontExplicitGlyphSubsetsPreserveUnicodeOwnershipAndResourceIdentity()
    {
        OpenTypeFont font = TestFontBuilder.LoadTestFont();
        ushort normal = font.MapCodePoint('A'), alternate = font.MapCodePoint('B');
        var map = new List<(ushort Glyph, int Unicode)> { (normal, 'A'), (alternate, 'A') };
        PdfEmbeddedFont embedded = PdfEmbeddedFont.CreateForGlyphMappings(font, map, CancellationToken.None);
        map.Clear();
        TestAssert.True(embedded.TryGetEncodedCid(alternate, out ushort cid), "The selected alternate must own an encoded subset CID.");
        TestAssert.Equal((int)'A', embedded.UnicodeByCid[cid]);
        TestAssert.True(embedded.ResourceKey != PdfEmbeddedFont.Create(font, new[] { (int)'A' }, CancellationToken.None).ResourceKey,
            "Distinct glyph remaps require distinct resource identities even when their Unicode sets match.");
        PdfEmbeddedFont canonical = PdfEmbeddedFont.Create(font, new[] { (int)'A' }, CancellationToken.None);
        PdfEmbeddedFont explicitCanonical = PdfEmbeddedFont.CreateForGlyphMappings(font, new[] { (normal, (int)'A') }, CancellationToken.None);
        TestAssert.Equal(canonical.ResourceKey, explicitCanonical.ResourceKey);
        TestAssert.True(canonical.FontProgramBytes.Span.SequenceEqual(explicitCanonical.FontProgramBytes.Span),
            "Canonical mappings retain their established subset bytes and identity.");
        TestAssert.Equal(embedded.ResourceKey, PdfEmbeddedFont.Merge(new[] { canonical, embedded }, CancellationToken.None).ResourceKey);
        TestAssert.Throws<ArgumentException>(() => PdfEmbeddedFont.CreateForGlyphMappings(font, new[] { ((ushort)0, (int)'A') }, CancellationToken.None));
        TestAssert.Throws<ArgumentException>(() => PdfEmbeddedFont.CreateForGlyphMappings(font, new[] { (normal, 0xD800) }, CancellationToken.None));
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => PdfEmbeddedFont.CreateForGlyphMappings(font, new[] { (normal, (int)'A') }, cancelled.Token));
    }

    public static void MathFontScriptBoxesPreserveSourceRolesSelectedGlyphsAndSeparatedBaselines()
    {
        byte[] table = new byte[224];
        BinaryPrimitives.WriteUInt16BigEndian(table, 1);
        BinaryPrimitives.WriteUInt16BigEndian(table.AsSpan(4), 10);
        BinaryPrimitives.WriteInt16BigEndian(table.AsSpan(10), 73);
        void Set(OpenTypeMathConstant field, short value) => BinaryPrimitives.WriteInt16BigEndian(table.AsSpan(18 + (int)field * 4), value);
        Set(OpenTypeMathConstant.SubscriptShiftDown, 200);
        Set(OpenTypeMathConstant.SubscriptTopMax, 500);
        Set(OpenTypeMathConstant.SubscriptBaselineDropMin, 100);
        Set(OpenTypeMathConstant.SuperscriptShiftUp, 350);
        Set(OpenTypeMathConstant.SuperscriptBottomMin, 100);
        Set(OpenTypeMathConstant.SuperscriptBaselineDropMax, 500);
        Set(OpenTypeMathConstant.SubSuperscriptGapMin, 150);
        Set(OpenTypeMathConstant.SuperscriptBottomMaxWithSubscript, 450);
        Set(OpenTypeMathConstant.SpaceAfterScript, 50);
        OpenTypeFont font = OpenTypeFont.Load(TestFontBuilder.CreateMathFont(table, Gsub(false, true)));
        var resource = new PdfFontResource("Math1", PdfEmbeddedFont.CreateForGlyphMappings(font,
            new[] { (font.MapCodePoint('A'), (int)'A'), ((ushort)36, (int)'A'), (font.MapCodePoint('B'), (int)'B') }, CancellationToken.None));
        DocxMathRow Row(string text) => new([new DocxMathRun(new(text, 12d, null, false, false, false, null, "TestFont"), "p", null)]);
        var scripts = new DocxMathScripts(Row("A"), Row("A"), Row("B"));
        var equation = new DocxEquation(new([scripts]), 0, 0, true, "center", null, null);
        TestAssert.True(DocxMathLayout.TryMeasure(equation, _ => resource, CancellationToken.None, out DocxMathBox? box, out string? reason),
            reason ?? "Expected a script box.");
        DocxMathBox measured = TestAssert.NotNull(box);
        TestAssert.Equal(3, measured.Glyphs.Count);
        TestAssert.True(measured.Glyphs[0].EncodedGlyph != measured.Glyphs[1].EncodedGlyph &&
            measured.Glyphs[0].SourceCodePoint == measured.Glyphs[1].SourceCodePoint,
            "Script style changes the selected glyph while preserving the same mathematical source character.");
        TestAssert.True(measured.Glyphs[0].BaselineY == 0d && measured.Glyphs[1].BaselineY < 0d && measured.Glyphs[2].BaselineY > 0d,
            "Base, subscript and superscript must retain distinct stored baselines.");
        TestAssert.True(measured.Glyphs.Skip(1).All(glyph => Math.Abs(glyph.FontSize - 8.52d) < 1e-9),
            "Script arguments use the independently measured font scale and Office grid.");
        TestAssert.Equal("AAB", string.Concat(new[] { scripts.Base, scripts.Subscript!, scripts.Superscript! }
            .SelectMany(row => row.Children.Cast<DocxMathRun>()).Select(run => run.Text.Text)));
        TestAssert.True(!DocxMathLayout.TryGetRuns(equation with { Content = new([scripts with { Subscript = new([scripts]) }]) },
            CancellationToken.None, out _), "Nested script structure keeps its explicit fallback.");
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryMeasure(equation, _ => resource, cancelled.Token, out _, out _));
        using OoxConversionBudget.Scope scope = OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxSceneNodesPerConversion = 2 });
        TestAssert.Throws<OoxPdfLimitExceededException>(() => DocxMathLayout.TryMeasure(equation, _ => resource, CancellationToken.None, out _, out _));
    }

    public static void MathFontScriptBadTargetsAndOptionalGsubPreserveOrdinaryFontUse()
    {
        byte[] constants = new byte[224];
        BinaryPrimitives.WriteUInt16BigEndian(constants, 1);
        BinaryPrimitives.WriteUInt16BigEndian(constants.AsSpan(4), 10);
        foreach (ushort target in new ushort[] { 0, ushort.MaxValue })
        {
            byte[] gsub = Gsub(false, true);
            BinaryPrimitives.WriteUInt16BigEndian(gsub.AsSpan(78), target);
            OpenTypeFont font = OpenTypeFont.Load(TestFontBuilder.CreateMathFont(constants, gsub));
            TestAssert.True(!font.TryGetMathScriptGlyph(34, 1, CancellationToken.None, out _),
                "Malformed script targets cannot publish .notdef or out-of-font glyphs.");
            TestAssert.True(font.TryGetGlyphBounds(34, out _), "Optional script defects cannot invalidate ordinary glyph metrics.");
        }
        OpenTypeFont malformed = OpenTypeFont.Load(TestFontBuilder.CreateMathFont(constants, new byte[4]));
        TestAssert.True(!malformed.TryGetMathScriptGlyph(34, 1, CancellationToken.None, out _), "A malformed selected GSUB keeps script fallback.");
        TestAssert.True(malformed.MapCodePoint('A') != 0, "Ordinary face loading remains usable with malformed optional GSUB.");
    }

    private static byte[] Gsub(bool ranges, bool extension)
    {
        int subtable = extension ? 66 : 58;
        byte[] data = new byte[subtable + (ranges ? 28 : 24)];
        void U(int at, ushort value) => BinaryPrimitives.WriteUInt16BigEndian(data.AsSpan(at), value);
        U(0, 1); U(4, 10); U(6, 30); U(8, 44);
        U(10, 1); "math"u8.CopyTo(data.AsSpan(12)); U(16, 8);
        U(18, 4); U(22, 0); U(24, 0xFFFF); U(26, 1); U(28, 0);
        U(30, 1); "ssty"u8.CopyTo(data.AsSpan(32)); U(36, 8); U(38, 0); U(40, 1); U(42, 0);
        U(44, 1); U(46, 4); U(48, extension ? (ushort)7 : (ushort)3); U(50, 0); U(52, 1); U(54, 10);
        if (extension) { U(58, 1); U(60, 3); BinaryPrimitives.WriteUInt32BigEndian(data.AsSpan(62), 8); }
        U(subtable, 1); U(subtable + 2, 18); U(subtable + 4, 1); U(subtable + 6, 10);
        U(subtable + 10, 2); U(subtable + 12, 36); U(subtable + 14, 37);
        U(subtable + 18, ranges ? (ushort)2 : (ushort)1); U(subtable + 20, 1); U(subtable + 22, 34);
        if (ranges) { U(subtable + 24, 34); U(subtable + 26, 0); }
        return data;
    }
}
