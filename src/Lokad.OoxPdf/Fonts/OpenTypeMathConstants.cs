using System.Buffers.Binary;

namespace Lokad.OoxPdf.Fonts;

// OpenType MATH 1.0: four initial scalar fields, 51 MathValueRecords,
// and one final scalar. Retain only the fixed constants window, with no
// copying, mutable publication, per-glyph allocation or device-table traversal.
internal readonly struct OpenTypeMathConstants
{
    private const int ByteLength = 214;
    private readonly ReadOnlyMemory<byte> data;
    private OpenTypeMathConstants(ReadOnlyMemory<byte> data) => this.data = data;

    public short ScriptPercentScaleDown => I16(0);
    public short ScriptScriptPercentScaleDown => I16(2);
    public ushort DelimitedSubFormulaMinHeight => U16(4);
    public ushort DisplayOperatorMinHeight => U16(6);
    public short RadicalDegreeBottomRaisePercent => I16(212);

    public short GetValue(OpenTypeMathConstant constant)
    {
        int index = (int)constant;
        if ((uint)index >= 51u) { throw new ArgumentOutOfRangeException(nameof(constant)); }
        return I16(8 + index * 4);
    }

    public static bool TryRead(ReadOnlyMemory<byte> table, out OpenTypeMathConstants constants)
    {
        constants = default;
        ReadOnlySpan<byte> header = table.Span;
        if (header.Length < 10 || BinaryPrimitives.ReadUInt16BigEndian(header) != 1 ||
            BinaryPrimitives.ReadUInt16BigEndian(header.Slice(2)) != 0)
        {
            return false;
        }
        int offset = BinaryPrimitives.ReadUInt16BigEndian(header.Slice(4));
        if (offset < 10 || table.Length - offset < ByteLength) { return false; }
        constants = new OpenTypeMathConstants(table.Slice(offset, ByteLength));
        return true;
    }

    private short I16(int offset) => BinaryPrimitives.ReadInt16BigEndian(data.Span.Slice(offset, 2));
    private ushort U16(int offset) => BinaryPrimitives.ReadUInt16BigEndian(data.Span.Slice(offset, 2));
}

internal enum OpenTypeMathConstant
{
    MathLeading,
    AxisHeight,
    AccentBaseHeight,
    FlattenedAccentBaseHeight,
    SubscriptShiftDown,
    SubscriptTopMax,
    SubscriptBaselineDropMin,
    SuperscriptShiftUp,
    SuperscriptShiftUpCramped,
    SuperscriptBottomMin,
    SuperscriptBaselineDropMax,
    SubSuperscriptGapMin,
    SuperscriptBottomMaxWithSubscript,
    SpaceAfterScript,
    UpperLimitGapMin,
    UpperLimitBaselineRiseMin,
    LowerLimitGapMin,
    LowerLimitBaselineDropMin,
    StackTopShiftUp,
    StackTopDisplayStyleShiftUp,
    StackBottomShiftDown,
    StackBottomDisplayStyleShiftDown,
    StackGapMin,
    StackDisplayStyleGapMin,
    StretchStackTopShiftUp,
    StretchStackBottomShiftDown,
    StretchStackGapAboveMin,
    StretchStackGapBelowMin,
    FractionNumeratorShiftUp,
    FractionNumeratorDisplayStyleShiftUp,
    FractionDenominatorShiftDown,
    FractionDenominatorDisplayStyleShiftDown,
    FractionNumeratorGapMin,
    FractionNumDisplayStyleGapMin,
    FractionRuleThickness,
    FractionDenominatorGapMin,
    FractionDenomDisplayStyleGapMin,
    SkewedFractionHorizontalGap,
    SkewedFractionVerticalGap,
    OverbarVerticalGap,
    OverbarRuleThickness,
    OverbarExtraAscender,
    UnderbarVerticalGap,
    UnderbarRuleThickness,
    UnderbarExtraDescender,
    RadicalVerticalGap,
    RadicalDisplayStyleVerticalGap,
    RadicalRuleThickness,
    RadicalExtraAscender,
    RadicalKernBeforeDegree,
    RadicalKernAfterDegree
}
