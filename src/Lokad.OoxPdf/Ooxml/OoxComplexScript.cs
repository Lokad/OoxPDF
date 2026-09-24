using System.Globalization;
using System.Text;

namespace Lokad.OoxPdf.Ooxml;

// RV02: complex-script detection for truthful diagnostics. Rendering maps one
// glyph per codepoint without joining, reordering, bidi resolution, or mark
// positioning, so text needing those behaviors must diagnose instead of silently
// misrendering. Buckets group by unimplemented behavior; scripts outside the
// listed ranges stay undetected until measured.
[Flags]
internal enum OoxComplexScriptKind
{
    None = 0,
    Joining = 1,
    Reordering = 2,
    Bidirectional = 4,
    CombiningMark = 8
}

internal static class OoxComplexScript
{
    public static OoxComplexScriptKind DetectNeeds(string? text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return OoxComplexScriptKind.None;
        }

        OoxComplexScriptKind kinds = OoxComplexScriptKind.None;
        foreach (Rune rune in text.EnumerateRunes())
        {
            int value = rune.Value;
            if (IsJoiningScript(value))
            {
                kinds |= OoxComplexScriptKind.Joining;
            }

            if (IsReorderingScript(value))
            {
                kinds |= OoxComplexScriptKind.Reordering;
            }

            if (IsBidirectionalScript(value))
            {
                kinds |= OoxComplexScriptKind.Bidirectional;
            }

            UnicodeCategory category = Rune.GetUnicodeCategory(rune);
            if (category == UnicodeCategory.NonSpacingMark || category == UnicodeCategory.EnclosingMark)
            {
                kinds |= OoxComplexScriptKind.CombiningMark;
            }
        }

        return kinds;
    }

    private static bool IsJoiningScript(int value)
    {
        return (value >= 0x0600 && value <= 0x06FF) || (value >= 0x0700 && value <= 0x074F) ||
            (value >= 0x0750 && value <= 0x077F) || (value >= 0x08A0 && value <= 0x08FF) ||
            (value >= 0xFB50 && value <= 0xFDFF) || (value >= 0xFE70 && value <= 0xFEFF);
    }

    private static bool IsReorderingScript(int value)
    {
        return (value >= 0x0900 && value <= 0x0DFF) || (value >= 0x0F00 && value <= 0x0FFF) ||
            (value >= 0x1000 && value <= 0x109F);
    }

    private static bool IsBidirectionalScript(int value)
    {
        return (value >= 0x0590 && value <= 0x05FF) || (value >= 0x0780 && value <= 0x07BF) ||
            (value >= 0x07C0 && value <= 0x07FF) || value == 0x061C || value == 0x200E || value == 0x200F ||
            (value >= 0x202A && value <= 0x202E) || (value >= 0x2066 && value <= 0x2069);
    }
}
