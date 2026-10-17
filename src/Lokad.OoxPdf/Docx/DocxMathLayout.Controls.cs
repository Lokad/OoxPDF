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

    internal static IReadOnlyList<IDocxMathControl> GetControlSources(DocxEquation equation, CancellationToken token)
    {
        var controls = new List<IDocxMathControl>();
        if (GetStandaloneControl(equation, token) is { } single)
        { Append(single); return controls; }
        if (GetStructuredRow(equation, token) is not { } row) return [];
        foreach (var part in row.Parts)
        {
            token.ThrowIfCancellationRequested();
            if (part.Structure is not IDocxMathControl control) return [];
            Append(control);
        }
        return controls;

        void Append(IDocxMathControl control)
        {
            token.ThrowIfCancellationRequested();
            controls.Add(control);
            // The admitted fraction grammar has at most one literal script
            // per argument. Publish its actual owner even without source runs;
            // no recursive discovery or invented source/paint is required.
            if (control is DocxMathFraction fraction)
            {
                if (GetScriptArgument(fraction.Numerator) is { } numerator) controls.Add(numerator);
                if (GetScriptArgument(fraction.Denominator) is { } denominator) controls.Add(denominator);
            }
        }
    }

    internal static IReadOnlyList<IDocxMathControl> GetMetricOnlyControls(DocxEquation equation,
        IReadOnlyList<DocxMathRunSource> runs, CancellationToken token)
    {
        if (MetricOnly(equation)) return GetControlSources(equation, token);
        var controls = GetControlSources(equation, token);
        var metricOnly = new List<IDocxMathControl>();
        foreach (var control in controls)
        {
            token.ThrowIfCancellationRequested();
            if (MetricOnly(equation with { Content = new([(DocxMathNode)control]) })) metricOnly.Add(control);
            else if (runs.Count == 0) return [];
        }
        return metricOnly;

        bool MetricOnly(DocxEquation source) => IsEmptyFraction(source) || IsEmptyMatrix(source, token) || IsEmptyScripts(source, token);
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
