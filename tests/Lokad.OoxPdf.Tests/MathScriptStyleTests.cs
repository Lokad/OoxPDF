using System.Buffers.Binary;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class MathScriptStyleTests
{
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
