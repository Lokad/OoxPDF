using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal static partial class DocxMathLayout
{
    private static DocxMathFunction? GetFunction(DocxEquation equation) =>
        GetSingleSupportedChild(equation) is { } node ? GetSupportedFunction(node) : null;

    private static DocxMathFunction? GetSupportedFunction(DocxMathNode node) =>
        node is DocxMathFunction { HasUnsupportedProperties: false } function &&
        (function.ControlText is null || DocxMathRowLayout.HasSupportedTextFormatting(function.ControlText, allowBold: false)) ? function : null;

    private static bool TryGetFunctionRuns(DocxEquation equation, DocxMathFunction function, CancellationToken token,
        out IReadOnlyList<DocxMathRunSource> runs)
    {
        runs = [];
        if (!DocxMathRowLayout.TryGetRenderCodePoints(function.Name, token, out var namePoints, out _) || namePoints.Count == 0) { return false; }
        DocxEquation argument = equation with { Content = function.Argument };
        // Validate the bounded argument role before delegating discovery, so a
        // nested function cannot recursively expand arbitrary authored depth.
        if (GetDelimiter(argument) is null && GetFraction(argument) is null && GetScripts(argument) is null && GetMatrix(argument, token) is null &&
            !DocxMathRowLayout.TryGetRenderCodePoints(function.Argument, token, out _, out _)) { return false; }
        if (!TryGetRuns(equation with { Content = function.Name }, token, out var nameRuns) || nameRuns.Count == 0 ||
            !TryGetRuns(argument, token, out var argumentRuns) || argumentRuns.Count == 0) { return false; }
        var result = new List<DocxMathRunSource>(nameRuns.Count + argumentRuns.Count);
        result.AddRange(nameRuns); result[0] = result[0] with { ControlOwners = [function] };
        result.AddRange(argumentRuns); runs = result;
        return true;
    }

    private static bool TryMeasureFunction(DocxEquation equation, DocxMathFunction function,
        Func<DocxMathRun, PdfFontResource?> resolveFont, CancellationToken token,
        out DocxMathBox? box, out string? fallback, Func<IDocxMathControl, PdfFontResource?>? resolveControlFont,
        Func<IDocxMathControl, OpenTypeFont?>? resolveControlMetrics, double? rowControlSize = null)
    {
        box = null;
        if (!TryGetRuns(equation, token, out var sources) || sources.Count == 0)
        { fallback = "UnsupportedMathFunctionSource"; return false; }
        OpenTypeFont? font = resolveControlMetrics is not null ? resolveControlMetrics(function) :
            resolveControlFont is not null ? resolveControlFont(function)?.Font.Font : resolveFont(sources[0].Run)?.Font.Font;
        if (font is null || !font.TryGetMathConstants(out _))
        { fallback = "MissingMathFunctionMetrics"; return false; }
        // A mixed row owns application spacing/frame size. A standalone
        // application retains its authored invisible control size.
        double size = rowControlSize ?? function.ControlText?.FontSize ?? equation.ParentFontSize ?? sources[0].Run.Text.FontSize;
        if (!double.IsFinite(size) || size <= 0d)
        { fallback = "InvalidMathFunctionSize"; return false; }
        if (!DocxMathRowLayout.TryMeasure(function.Name, resolveFont, token, out var name, out fallback, exportFontSize: true) || name is null || name.Glyphs.Count == 0)
        { fallback ??= "UnsupportedMathFunctionName"; return false; }
        DocxEquation argumentEquation = equation with { Content = function.Argument };
        bool measured = GetDelimiter(argumentEquation) is not null
            ? TryMeasure(argumentEquation, resolveFont, token, out var argument, out fallback, font, resolveControlFont, resolveControlMetrics)
            : TryMeasureBoxArgument(argumentEquation, resolveFont, font, size, token, out argument, out fallback,
                resolveControlFont: resolveControlFont, resolveControlMetrics: resolveControlMetrics);
        if (!measured || argument is null || argument.Glyphs.Count == 0)
        { fallback ??= "UnsupportedMathFunctionArgument"; return false; }
        double gap = argument.Edges?.First == DocxMathRowLayout.AtomKind.Open ? 0d : size / 6d;
        double argumentX = name.AdvanceWidth + gap;
        double advance = argumentX + argument.AdvanceWidth;
        double width = Math.Max(advance, Math.Max(name.Width, argumentX + argument.Width));
        double above = Math.Max(name.AboveBaseline, argument.AboveBaseline), below = Math.Max(name.BelowBaseline, argument.BelowBaseline);
        if (!double.IsFinite(width) || !double.IsFinite(above) || !double.IsFinite(below) || width <= 0d)
        { fallback = "InvalidMathFunctionExtent"; return false; }
        var glyphs = new List<DocxMathGlyphPlacement>(name.Glyphs);
        foreach (var glyph in argument.Glyphs) { token.ThrowIfCancellationRequested(); glyphs.Add(glyph with { X = glyph.X + argumentX }); }
        var rules = new List<DocxMathRulePlacement>();
        foreach (var rule in argument.Rules) { token.ThrowIfCancellationRequested(); rules.Add(rule with { X = rule.X + argumentX }); }
        box = new(width, above, below, glyphs)
        { LogicalAdvance = advance == width ? null : advance, Rules = rules, ParentMetrics = new(font, size) { IncludeTypographicFrame = true },
            Edges = new(DocxMathRowLayout.AtomKind.LargeOperator, argument.Edges?.Last ?? DocxMathRowLayout.AtomKind.Ordinary, false) };
        fallback = null;
        return true;
    }
}
