using System.Collections.Generic;
using System.Globalization;
using System.Text;
using Lokad.OoxPdf.Diagnostics;

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
            kinds |= ClassifyScalar(rune.Value, Rune.GetUnicodeCategory(rune));
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

    private static OoxComplexScriptKind ClassifyScalar(int value, UnicodeCategory category)
    {
        OoxComplexScriptKind kinds = OoxComplexScriptKind.None;
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

        if (category == UnicodeCategory.NonSpacingMark || category == UnicodeCategory.EnclosingMark)
        {
            kinds |= OoxComplexScriptKind.CombiningMark;
        }

        return kinds;
    }

    public static OoxComplexScriptKind DetectCodepoints(IEnumerable<int> codepoints)
    {
        OoxComplexScriptKind kinds = OoxComplexScriptKind.None;
        foreach (int codepoint in codepoints)
        {
            if (codepoint < 0 || codepoint > 0x10FFFF || (codepoint >= 0xD800 && codepoint <= 0xDFFF))
            {
                continue;
            }

            kinds |= ClassifyScalar(codepoint, Rune.GetUnicodeCategory(new Rune(codepoint)));
        }

        return kinds;
    }

    public static OoxPdfDiagnostic CreateApproximationDiagnostic(OoxComplexScriptKind kind)
    {
        string feature;
        string behavior;
        switch (kind)
        {
            case OoxComplexScriptKind.Joining:
                feature = "joining scripts";
                behavior = "cursive joining";
                break;
            case OoxComplexScriptKind.Reordering:
                feature = "reordering scripts";
                behavior = "syllable reordering";
                break;
            case OoxComplexScriptKind.Bidirectional:
                feature = "bidirectional text";
                behavior = "bidirectional reordering";
                break;
            case OoxComplexScriptKind.CombiningMark:
                feature = "combining marks";
                behavior = "mark positioning";
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(kind), "Complex-script diagnostics take one behavior flag at a time.");
        }

        return new OoxPdfDiagnostic(
            "COMPLEX_SCRIPT_APPROXIMATION",
            OoxPdfSeverity.Warning,
            "Complex script renders without " + behavior + "; glyphs emit in source order.",
            PartName: null,
            SlideIndex: null,
            PageIndex: null,
            Feature: feature,
            Fallback: "Unshaped glyphs");
    }
}
