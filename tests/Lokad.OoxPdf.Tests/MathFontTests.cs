using System.Buffers.Binary;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static class MathFontTests
{
    public static void MathFontVerticalNearestVariantsUseDistanceTallerTiesAndAssemblyFallback()
    {
        foreach(bool range in new[]{true,false})
        {
            var font=OpenTypeFont.Load(TestFontBuilder.CreateMathFont(VariantTable(range)));
            foreach((double height,ushort glyph,ushort advance) in new (double,ushort,ushort)[]{
                (0,34,500),(749.99,34,500),(750,36,1000),(1000,36,1000),(1250,35,1500),(1500,35,1500)})
            {
                TestAssert.True(font.TryGetNearestMathVerticalVariant(34,height,CancellationToken.None,out var selected),"Expected a ready-made nearest variant.");
                TestAssert.Equal(new OpenTypeMathGlyphVariant(glyph,advance),selected);
            }
            foreach(double height in new[]{-1d,1500.01d,double.NaN,double.PositiveInfinity})
                TestAssert.True(!font.TryGetNearestMathVerticalVariant(34,height,CancellationToken.None,out var selected) && selected==default,
                    "Nearest fitting cannot publish a largest inadequate variant when only an assembly can reach the target.");
        }
        byte[] malformed=VariantTable(false);BinaryPrimitives.WriteUInt16BigEndian(malformed.AsSpan(40),ushort.MaxValue);
        var bad=OpenTypeFont.Load(TestFontBuilder.CreateMathFont(malformed));
        TestAssert.True(!bad.TryGetNearestMathVerticalVariant(34,500d,CancellationToken.None,out var absent) && absent==default,
            "A malformed later record discards even an exact earlier nearest candidate.");
        var valid=OpenTypeFont.Load(TestFontBuilder.CreateMathFont(VariantTable(false)));
        using var cancelled=new CancellationTokenSource();cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(()=>valid.TryGetNearestMathVerticalVariant(34,500d,cancelled.Token,out _));
        using(OoxConversionBudget.Scope scope=OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxFontWorkPerConversion=1 }))
        {
            TestAssert.True(valid.TryGetNearestMathVerticalVariant(34,500d,CancellationToken.None,out _),"One bounded query is available.");
            TestAssert.Throws<OoxPdfLimitExceededException>(()=>valid.TryGetMathVerticalVariant(34,500d,CancellationToken.None,out _));
        }
    }

    public static void MathFontVerticalVariantsChooseTheSmallestAdequateOwnedGlyph()
    {
        foreach (bool rangeCoverage in new[] { false, true })
        {
            var font = OpenTypeFont.Load(TestFontBuilder.CreateMathFont(VariantTable(rangeCoverage)));
            foreach ((double height, ushort glyph, ushort advance) in new (double, ushort, ushort)[]
                { (0d,34,500), (500d,34,500), (500.01d,36,1000), (1000d,36,1000), (1000.01d,35,1500), (1500d,35,1500) })
            {
                TestAssert.True(font.TryGetMathVerticalVariant(34,height,CancellationToken.None,out var selected), "Expected a ready-made vertical variant.");
                TestAssert.Equal(new OpenTypeMathGlyphVariant(glyph,advance),selected);
            }
            foreach (double height in new[] { -1d,1500.01d,double.NaN,double.PositiveInfinity })
                TestAssert.True(!font.TryGetMathVerticalVariant(34,height,CancellationToken.None,out var selected) && selected==default,
                    "Invalid sizes and heights requiring an assembly cannot publish a ready-made glyph.");
            TestAssert.True(!font.TryGetMathVerticalVariant(35,0d,CancellationToken.None,out _), "Uncovered glyphs retain explicit fallback.");
        }
    }

    public static void MathFontVerticalVariantsRejectTruncationOffsetsAndInvalidLaterRecords()
    {
        byte[] table=VariantTable(false);
        for(int length=0;length<table.Length;length++)
        {
            var font=OpenTypeFont.Load(TestFontBuilder.CreateMathFont(table[..length]));
            TestAssert.True(!font.TryGetMathVerticalVariant(34,0d,CancellationToken.None,out var selected) && selected==default,
                "A truncated MATH subtable cannot publish a partial variant selection.");
            TestAssert.Equal((ushort)34,font.MapCodePoint('A'));
        }
        foreach((int offset,ushort value) in new (int,ushort)[]
            { (2,1),(8,ushort.MaxValue),(12,1),(16,ushort.MaxValue),(20,1),(22,3),(30,ushort.MaxValue),(40,ushort.MaxValue),(42,0) })
        {
            byte[] malformed=(byte[])table.Clone();BinaryPrimitives.WriteUInt16BigEndian(malformed.AsSpan(offset),value);
            var font=OpenTypeFont.Load(TestFontBuilder.CreateMathFont(malformed));
            TestAssert.True(!font.TryGetMathVerticalVariant(34,0d,CancellationToken.None,out var selected) && selected==default,
                "Malformed offsets, coverage and later records must discard earlier candidate glyphs.");
            TestAssert.Equal((ushort)34,font.MapCodePoint('A'));
        }
    }

    public static void MathFontVerticalVariantsKeepCancellationAndCumulativeFontWork()
    {
        var font=OpenTypeFont.Load(TestFontBuilder.CreateMathFont(VariantTable(false)));
        using var cancelled=new CancellationTokenSource();cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(()=>font.TryGetMathVerticalVariant(34,0d,cancelled.Token,out _));
        using(OoxConversionBudget.Scope scope=OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxFontWorkPerConversion=1 }))
        {
            TestAssert.True(font.TryGetMathVerticalVariant(34,0d,CancellationToken.None,out _),"One bounded font-selection operation is available.");
            TestAssert.Throws<OoxPdfLimitExceededException>(()=>font.TryGetMathVerticalVariant(34,500d,CancellationToken.None,out _));
        }
        TestAssert.True(!TestFontBuilder.LoadTestFont().TryGetMathVerticalVariant(34,0d,CancellationToken.None,out _),
            "An ordinary face remains usable without optional MATH variants.");
    }

    private static byte[] VariantTable(bool range)
    {
        int construction=12+(range?10:6);
        byte[] bytes=new byte[10+construction+16];
        void Set(int offset,ushort value)=>BinaryPrimitives.WriteUInt16BigEndian(bytes.AsSpan(offset),value);
        Set(0,1);Set(8,10);Set(12,12);Set(16,1);Set(20,(ushort)construction);
        Set(22,(ushort)(range?2:1));Set(24,1);Set(26,34);
        if(range){Set(28,34);Set(30,0);}
        int at=10+construction;Set(at+2,3);
        Set(at+4,34);Set(at+6,500);Set(at+8,35);Set(at+10,1500);Set(at+12,36);Set(at+14,1000);
        return bytes;
    }

    public static void MathFontConstantsReadEveryRecordInDesignUnits()
    {
        byte[] table = CreateTable(10);
        TestAssert.True(OpenTypeMathConstants.TryRead(table, out OpenTypeMathConstants constants), "Complete MATH constants must be readable.");
        TestAssert.Equal((short)73, constants.ScriptPercentScaleDown);
        TestAssert.Equal((short)60, constants.ScriptScriptPercentScaleDown);
        TestAssert.Equal((ushort)3000, constants.DelimitedSubFormulaMinHeight);
        TestAssert.Equal((ushort)4000, constants.DisplayOperatorMinHeight);
        TestAssert.Equal((short)65, constants.RadicalDegreeBottomRaisePercent);
        foreach (OpenTypeMathConstant key in Enum.GetValues<OpenTypeMathConstant>())
            TestAssert.Equal((short)((int)key - 25), constants.GetValue(key));
        TestAssert.Throws<ArgumentOutOfRangeException>(() => constants.GetValue((OpenTypeMathConstant)(-1)));
        TestAssert.Throws<ArgumentOutOfRangeException>(() => constants.GetValue((OpenTypeMathConstant)51));
    }

    public static void MathFontConstantsRejectTruncationAndInvalidOffsets()
    {
        byte[] table = CreateTable(32);
        for (int length = 0; length < table.Length; length++)
            TestAssert.True(!OpenTypeMathConstants.TryRead(table.AsMemory(0, length), out _), "Truncated MATH constants must not escape their table.");
        foreach (ushort offset in new ushort[] { 0, 1, 9, 33, ushort.MaxValue })
        {
            byte[] changed = (byte[])table.Clone();
            BinaryPrimitives.WriteUInt16BigEndian(changed.AsSpan(4), offset);
            TestAssert.True(!OpenTypeMathConstants.TryRead(changed, out _), "Invalid constants offset must take the explicit missing-metrics path.");
        }
        TestAssert.True(OpenTypeMathConstants.TryRead(table, out _), "Valid non-default offset must remain readable.");
    }

    public static void MathFontConstantsRetainSignedValuesAndIgnoreDeviceCorrections()
    {
        byte[] table = CreateTable(10);
        BinaryPrimitives.WriteInt16BigEndian(table.AsSpan(10), short.MinValue);
        BinaryPrimitives.WriteInt16BigEndian(table.AsSpan(18 + (int)OpenTypeMathConstant.RadicalKernAfterDegree * 4), short.MinValue);
        TestAssert.True(OpenTypeMathConstants.TryRead(table, out OpenTypeMathConstants constants), "Signed values do not make a structurally valid constants table unreadable.");
        TestAssert.Equal(short.MinValue, constants.ScriptPercentScaleDown);
        TestAssert.Equal(short.MinValue, constants.GetValue(OpenTypeMathConstant.RadicalKernAfterDegree));
        // Every test MathValueRecord has an out-of-range device offset. PDF
        // design-unit metrics intentionally do not dereference those offsets.
        TestAssert.Equal((short)-24, constants.GetValue(OpenTypeMathConstant.AxisHeight));
    }

    public static void MathFontConstantsAreOptionalAndStayInsideTheirFontTable()
    {
        OpenTypeFont ordinary = TestFontBuilder.LoadTestFont();
        TestAssert.True(!ordinary.TryGetMathConstants(out _), "An ordinary face remains a usable font without math constants.");
        foreach (byte[] table in new[] { Array.Empty<byte>(), CreateTable(10)[..^1], CreateTable(10) })
        {
            OpenTypeFont font = OpenTypeFont.Load(TestFontBuilder.CreateMathFont(table));
            TestAssert.Equal(ordinary.MapCodePoint('A'), font.MapCodePoint('A'));
            TestAssert.Equal(ordinary.GetAdvanceWidth(ordinary.MapCodePoint('A')), font.GetAdvanceWidth(font.MapCodePoint('A')));
            bool present = font.TryGetMathConstants(out OpenTypeMathConstants constants);
            TestAssert.Equal(table.Length == 224, present);
            if (present) { TestAssert.Equal((short)9, constants.GetValue(OpenTypeMathConstant.FractionRuleThickness)); }
        }
    }

    public static void MathFontConstantsRejectUnknownVersionsWithoutAffectingOrdinaryGlyphs()
    {
        foreach ((ushort major, ushort minor) in new[] { ((ushort)0, (ushort)0), ((ushort)2, (ushort)0), ((ushort)1, (ushort)1) })
        {
            byte[] table = CreateTable(10);
            BinaryPrimitives.WriteUInt16BigEndian(table, major);
            BinaryPrimitives.WriteUInt16BigEndian(table.AsSpan(2), minor);
            OpenTypeFont font = OpenTypeFont.Load(TestFontBuilder.CreateMathFont(table));
            TestAssert.True(!font.TryGetMathConstants(out _), "Unknown MATH versions must not imply supported math metrics.");
            TestAssert.Equal((ushort)34, font.MapCodePoint('A'));
        }
    }

    private static byte[] CreateTable(ushort offset)
    {
        byte[] bytes = new byte[offset + 214];
        BinaryPrimitives.WriteUInt16BigEndian(bytes, 1);
        BinaryPrimitives.WriteUInt16BigEndian(bytes.AsSpan(4), offset);
        BinaryPrimitives.WriteInt16BigEndian(bytes.AsSpan(offset), 73);
        BinaryPrimitives.WriteInt16BigEndian(bytes.AsSpan(offset + 2), 60);
        BinaryPrimitives.WriteUInt16BigEndian(bytes.AsSpan(offset + 4), 3000);
        BinaryPrimitives.WriteUInt16BigEndian(bytes.AsSpan(offset + 6), 4000);
        for (int i = 0; i < 51; i++)
        {
            BinaryPrimitives.WriteInt16BigEndian(bytes.AsSpan(offset + 8 + i * 4), (short)(i - 25));
            BinaryPrimitives.WriteUInt16BigEndian(bytes.AsSpan(offset + 10 + i * 4), ushort.MaxValue);
        }
        BinaryPrimitives.WriteInt16BigEndian(bytes.AsSpan(offset + 212), 65);
        return bytes;
    }
}
