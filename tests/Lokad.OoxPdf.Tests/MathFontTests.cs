using System.Buffers.Binary;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static class MathFontTests
{
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
