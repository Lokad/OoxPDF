using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal static partial class DocxMathLayout
{
    // A supported control owns metrics and any structural paint even when its
    // literal children are empty. No synthetic source run or recursive discovery.
    internal static IDocxMathControl? GetStandaloneControl(DocxEquation equation, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        if (GetRadical(equation) is { } radical) return radical;
        if (GetNary(equation, token) is { } nary) return nary;
        if (GetMatrix(equation, token) is { } matrix) return matrix;
        if (GetFraction(equation) is { } fraction) return fraction;
        if (GetScripts(equation) is { } scripts) return scripts;
        return GetDelimiter(equation);
    }

    private static PdfFontResource? ResolveControlResource(IDocxMathControl control,
        IReadOnlyList<DocxMathRunSource> sources, Func<DocxMathRun, PdfFontResource?> resolveFont,
        Func<IDocxMathControl, PdfFontResource?>? resolveControlFont) =>
        resolveControlFont is not null ? resolveControlFont(control) :
            sources.Count == 0 ? null : resolveFont(sources[0].Run);

    private static double ResolveControlSize(IDocxMathControl control, DocxEquation equation,
        IReadOnlyList<DocxMathRunSource> sources) => control.ControlText?.FontSize ?? equation.ParentFontSize ??
            (sources.Count == 0 ? double.NaN : sources[0].Run.Text.FontSize);
}
