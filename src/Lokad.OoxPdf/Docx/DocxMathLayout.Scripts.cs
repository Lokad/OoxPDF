using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal static partial class DocxMathLayout
{
    private static DocxMathScripts? GetScripts(DocxEquation equation) =>
        GetStandaloneChild(equation) is DocxMathScripts { PropertiesXml.Count: 0 } scripts &&
        (scripts.Subscript is not null || scripts.Superscript is not null) ? scripts : null;

    private static bool TryMeasureScripts(DocxEquation equation, DocxMathScripts scripts,
        Func<DocxMathRun, PdfFontResource?> resolveFont, CancellationToken token, out DocxMathBox? box, out string? fallback)
    {
        box = null; fallback = null;
        if (!TryGetRuns(equation, token, out IReadOnlyList<DocxMathRunSource> sources) || sources.Count == 0 || sources[0].ScriptLevel != 0)
        { fallback = "UnsupportedMathScriptSource"; return false; }
        DocxMathRun owner = sources[0].Run;
        PdfFontResource? ownerResource = resolveFont(owner);
        if (ownerResource is null || !ownerResource.Font.Font.TryGetMathConstants(out OpenTypeMathConstants constants))
        { fallback = "MissingMathScriptConstants"; return false; }
        OpenTypeFont font = ownerResource.Font.Font;
        ReadOnlySpan<OpenTypeMathConstant> required = [OpenTypeMathConstant.SubscriptShiftDown, OpenTypeMathConstant.SubscriptTopMax,
            OpenTypeMathConstant.SubscriptBaselineDropMin, OpenTypeMathConstant.SuperscriptShiftUp, OpenTypeMathConstant.SuperscriptBottomMin,
            OpenTypeMathConstant.SuperscriptBaselineDropMax, OpenTypeMathConstant.SubSuperscriptGapMin,
            OpenTypeMathConstant.SuperscriptBottomMaxWithSubscript, OpenTypeMathConstant.SpaceAfterScript];
        foreach (OpenTypeMathConstant field in required)
            if (constants.GetValue(field) < 0) { fallback = "InvalidMathScriptConstants"; return false; }
        double size = owner.Text.FontSize;
        foreach (DocxMathRunSource source in sources)
        {
            token.ThrowIfCancellationRequested();
            if (source.Run.Text.FontSize != size || !ReferenceEquals(resolveFont(source.Run)?.Font.Font, font))
            { fallback = "MixedMathScriptMetrics"; return false; }
        }
        if (!DocxMathRowLayout.TryMeasure(scripts.Base, resolveFont, token, out DocxMathBox? basis, out fallback, exportFontSize: true) ||
            basis is null || basis.Glyphs.Count == 0) { fallback ??= "EmptyMathScriptBase"; return false; }
        DocxMathBox? subscript = null, superscript = null;
        if (scripts.Subscript is { } sub && (!DocxMathRowLayout.TryMeasure(sub, resolveFont, token, out subscript, out fallback, scriptLevel: 1) ||
            subscript is null || subscript.Glyphs.Count == 0)) { fallback ??= "EmptyMathSubscript"; return false; }
        if (scripts.Superscript is { } sup && (!DocxMathRowLayout.TryMeasure(sup, resolveFont, token, out superscript, out fallback, scriptLevel: 1) ||
            superscript is null || superscript.Glyphs.Count == 0)) { fallback ??= "EmptyMathSuperscript"; return false; }
        double scale = size / font.UnitsPerEm;
        double Value(OpenTypeMathConstant field) => constants.GetValue(field) * scale;
        double up = 0d, down = 0d;
        if (superscript is not null)
            up = Math.Max(Value(OpenTypeMathConstant.SuperscriptShiftUp), Math.Max(
                basis.AboveBaseline - Value(OpenTypeMathConstant.SuperscriptBaselineDropMax),
                superscript.BelowBaseline + Value(OpenTypeMathConstant.SuperscriptBottomMin)));
        if (subscript is not null)
            down = Math.Max(Value(OpenTypeMathConstant.SubscriptShiftDown), Math.Max(
                subscript.AboveBaseline - Value(OpenTypeMathConstant.SubscriptTopMax),
                basis.BelowBaseline + Value(OpenTypeMathConstant.SubscriptBaselineDropMin)));
        if (subscript is not null && superscript is not null)
        {
            double missing = Math.Max(0d, Value(OpenTypeMathConstant.SubSuperscriptGapMin) -
                (up + down - superscript.BelowBaseline - subscript.AboveBaseline));
            double raise = Math.Min(missing, Math.Max(0d,
                Value(OpenTypeMathConstant.SuperscriptBottomMaxWithSubscript) - (up - superscript.BelowBaseline)));
            up += raise;
            down += missing - raise;
        }
        DocxMathGlyphPlacement last = basis.Glyphs[^1];
        ushort lastGlyph = font.MapCodePoint(last.RenderCodePoint);
        if (!font.TryGetMathItalicCorrection(lastGlyph, out short correction))
        { fallback = "InvalidMathScriptItalicCorrection"; return false; }
        double subX = basis.Width - correction * last.FontSize / font.UnitsPerEm;
        double supX = basis.Width;
        double width = Math.Max(basis.Width, Math.Max(subscript is null ? 0d : subX + subscript.Width,
            superscript is null ? 0d : supX + superscript.Width)) + Value(OpenTypeMathConstant.SpaceAfterScript);
        double above = Math.Max(basis.AboveBaseline, Math.Max(subscript is null ? 0d : subscript.AboveBaseline - down,
            superscript is null ? 0d : superscript.AboveBaseline + up));
        double below = Math.Max(basis.BelowBaseline, Math.Max(subscript is null ? 0d : subscript.BelowBaseline + down,
            superscript is null ? 0d : superscript.BelowBaseline - up));
        if (!double.IsFinite(width) || !double.IsFinite(above) || !double.IsFinite(below) ||
            !double.IsFinite(up) || !double.IsFinite(down) || width < 0d || above < 0d || below < 0d || up < 0d || down < 0d || subX < 0d)
        { fallback = "InvalidMathScriptExtent"; return false; }
        var glyphs = new List<DocxMathGlyphPlacement>(basis.Glyphs.Count + (subscript?.Glyphs.Count ?? 0) + (superscript?.Glyphs.Count ?? 0));
        Append(basis, 0d, 0d);
        if (subscript is not null) { Append(subscript, subX, -down); }
        if (superscript is not null) { Append(superscript, supX, up); }
        box = new DocxMathBox(width, above, below, glyphs);
        fallback = null;
        return true;

        void Append(DocxMathBox child, double x, double baseline)
        {
            foreach (DocxMathGlyphPlacement glyph in child.Glyphs)
            {
                token.ThrowIfCancellationRequested();
                glyphs.Add(glyph with { X = glyph.X + x, BaselineY = glyph.BaselineY + baseline });
            }
        }
    }
}
