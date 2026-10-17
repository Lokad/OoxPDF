using System.Diagnostics.CodeAnalysis;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal static partial class DocxMathLayout
{
    private const int HiddenPlaceholderCodePoint = 0x2B1A;
    // Word retains the dotted-square's metrics while suppressing its paint.
    // Supported empty argument roles share this metric-only box.
    private static bool TryMeasureHiddenPlaceholder(OpenTypeFont font, double size, [NotNullWhen(true)] out DocxMathBox? box)
    {
        box = null;
        ushort glyph = font.MapCodePoint(HiddenPlaceholderCodePoint);
        if (glyph == 0 || !font.TryGetGlyphBounds(glyph, out var bounds)) return false;
        double scale = size / font.UnitsPerEm;
        double width = font.GetAdvanceWidth(glyph) * scale;
        if (!double.IsFinite(width) || width <= 0d) return false;
        box = new(width, Math.Max(0d, bounds.YMax * scale), Math.Max(0d, -bounds.YMin * scale), []);
        return true;
    }

    private static bool IsEmptyArgument(DocxMathRow row, CancellationToken token = default, bool allowArgumentSize = false)
    {
        if (!row.HasSupportedProperties && !(allowArgumentSize && row.HasKnownArgumentProperties)) return false;
        if (row.Children.Count == 0) return true;
        // Ordinary visible rows exit without allocating a traversal stack.
        if (row.Children.Any(node => node is DocxMathRun run && run.Text.Text.Length != 0)) return false;
        var pending = new Stack<DocxMathNode>(row.Children);
        while (pending.TryPop(out var node))
        {
            token.ThrowIfCancellationRequested();
            if (node is DocxMathRow nested && nested.HasSupportedProperties)
                foreach (var child in nested.Children) pending.Push(child);
            else if (node is not DocxMathRun run || run.Text.Text.Length != 0 || !DocxMathRowLayout.HasSupportedStyle(run)) return false;
        }
        return true;
    }

    private static bool TryMeasureLiteralArgument(DocxMathRow row, Func<DocxMathRun, PdfFontResource?> resolveFont,
        OpenTypeFont font, double size, int scriptLevel, CancellationToken token,
        out DocxMathBox? box, out string? fallback, out bool empty,
        bool allowArgumentSize = false, bool exportFontSize = false, double? operatorSpacingFontSize = null,
        DocxMathRowLayout.Neighbors neighbors = default)
    {
        empty = IsEmptyArgument(row, token, allowArgumentSize);
        if (!empty)
            return DocxMathRowLayout.TryMeasure(row, resolveFont, token, out box, out fallback, scriptLevel,
                operatorSpacingFontSize: operatorSpacingFontSize, scriptMetricsFont: font, allowArgumentSize: allowArgumentSize, exportFontSize: exportFontSize, neighbors: neighbors);
        box = null;
        fallback = "MissingMathArgumentPlaceholder";
        double argumentSize = size;
        if (scriptLevel != 0 && !DocxMathArgumentMetrics.TryGetScriptFontSize(font, size, scriptLevel, out argumentSize) ||
            !TryMeasureHiddenPlaceholder(font, OfficePdfTextEmissionProfile.FontSize(argumentSize), out box)) return false;
        OoxConversionBudget.Current?.ChargeSceneNodes(1);
        fallback = null;
        return true;
    }

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
