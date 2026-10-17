using System.Diagnostics.CodeAnalysis;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal static partial class DocxMathLayout
{
    // Word retains the dotted-square's metrics while suppressing its paint.
    // Matrices and empty fraction arguments share this metric-only box.
    private static bool TryMeasureHiddenPlaceholder(OpenTypeFont font, double size, [NotNullWhen(true)] out DocxMathBox? box)
    {
        box = null;
        ushort glyph = font.MapCodePoint(0x2B1A);
        if (glyph == 0 || !font.TryGetGlyphBounds(glyph, out var bounds)) return false;
        double scale = size / font.UnitsPerEm;
        double width = font.GetAdvanceWidth(glyph) * scale;
        if (!double.IsFinite(width) || width <= 0d) return false;
        box = new(width, Math.Max(0d, bounds.YMax * scale), Math.Max(0d, -bounds.YMin * scale), []);
        return true;
    }

    private static bool IsEmptyArgument(DocxMathRow row) => row.HasSupportedProperties &&
        row.Children.All(node => node is DocxMathRun run && run.Text.Text.Length == 0 &&
            !run.HasUnsupportedProperties && DocxMathRowLayout.HasSupportedTextFormatting(run.Text, run.NormalText));

    internal static bool IsEmptyFraction(DocxEquation equation) => GetFraction(equation) is { ControlText: not null, TypeValue: null or "bar" } fraction &&
        IsEmptyArgument(fraction.Numerator) && IsEmptyArgument(fraction.Denominator);

    internal static DocxMathRow? RemoveDeletedFractionText(DocxEquation equation, CancellationToken token)
    {
        if (GetFraction(equation) is not { TypeValue: null or "bar" } fraction || !TryGetRuns(equation, token, out _)) return null;
        return new([fraction with { Numerator = new([]), Denominator = new([]) }]);
    }

    private static bool TryMeasureEmptyFraction(DocxEquation equation, DocxMathFraction fraction, OpenTypeFont? font,
        CancellationToken token, out DocxMathBox? box, out string? fallback)
    {
        box = null;
        token.ThrowIfCancellationRequested();
        double size = fraction.ControlText!.FontSize;
        if (font is null || !font.TryGetMathConstants(out var constants) || !double.IsFinite(size) || size <= 0d ||
            !DocxMathRowLayout.HasSupportedTextFormatting(fraction.ControlText, allowBold: false))
        { fallback = "MissingEmptyFractionMetrics"; return false; }
        double argumentSize = size;
        if (!equation.Display && !DocxMathArgumentMetrics.TryGetScriptFontSize(font, size, 1, out argumentSize) ||
            !TryMeasureHiddenPlaceholder(font, OfficePdfTextEmissionProfile.FontSize(argumentSize), out var argument) || argument is null)
        { fallback = "MissingEmptyFractionPlaceholder"; return false; }
        OoxConversionBudget.Current?.ChargeSceneNodes(2);
        RgbColor.TryParse(fraction.ControlText.ColorHex, out var color);
        return TryComposeFraction(equation, fraction, font, constants, size, argument, argument, color,
            useParentMetrics: true, token, out box, out fallback);
    }
}
