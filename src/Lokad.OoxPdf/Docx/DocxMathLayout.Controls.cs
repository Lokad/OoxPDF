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
        // Callers have admitted the source grammar through TryGetRuns. Walk its
        // actual roles in source order; painted runs cannot stand in for an
        // empty child's metric owner. Hidden operator roles remain unvisited.
        var controls = new List<IDocxMathControl>();
        var pending = new Stack<DocxMathNode>();
        pending.Push(equation.Content);
        while (pending.TryPop(out var node))
        {
            token.ThrowIfCancellationRequested();
            if (node is IDocxMathControl control) controls.Add(control);
            switch (node)
            {
                case DocxMathRow row:
                    for (int i = row.Children.Count - 1; i >= 0; i--) pending.Push(row.Children[i]);
                    break;
                case DocxMathFraction fraction:
                    pending.Push(fraction.Denominator); pending.Push(fraction.Numerator);
                    break;
                case DocxMathScripts scripts:
                    if (scripts.Superscript is { } sup) pending.Push(sup);
                    if (scripts.Subscript is { } sub) pending.Push(sub);
                    pending.Push(scripts.Base);
                    break;
                case DocxMathRadical radical:
                    pending.Push(radical.Radicand);
                    if (radical.DegreeHidden != true) pending.Push(radical.Degree);
                    break;
                case DocxMathDelimiter delimiter:
                    for (int i = delimiter.Arguments.Count - 1; i >= 0; i--) pending.Push(delimiter.Arguments[i]);
                    break;
                case DocxMathNary nary:
                    pending.Push(nary.Body);
                    if (nary.SuperscriptHiddenValue != true) pending.Push(nary.Superscript);
                    if (nary.SubscriptHiddenValue != true) pending.Push(nary.Subscript);
                    break;
                case DocxMathFunction function:
                    pending.Push(function.Argument); pending.Push(function.Name);
                    break;
                case DocxMathMatrix matrix:
                    for (int r = matrix.Rows.Count - 1; r >= 0; r--)
                        for (int c = matrix.Rows[r].Cells.Count - 1; c >= 0; c--) pending.Push(matrix.Rows[r].Cells[c]);
                    break;
                case DocxMathRun: break;
                default: return [];
            }
        }
        return controls;
    }

    internal static IReadOnlyList<IDocxMathControl> GetMetricOnlyControls(DocxEquation equation,
        IReadOnlyList<DocxMathRunSource> runs, CancellationToken token,
        IReadOnlyList<IDocxMathControl>? controlSources = null)
    {
        token.ThrowIfCancellationRequested();
        var controls = controlSources ?? GetControlSources(equation, token);
        var root = GetSingleSupportedChild(equation) as IDocxMathControl;
        bool hasLiteralSources = runs.Any(item => item.Run.Text.Text.Length != 0);
        var metricOnly = new List<IDocxMathControl>();
        foreach (var control in controls)
        {
            token.ThrowIfCancellationRequested();
            if (control is (DocxMathFraction or DocxMathMatrix or DocxMathScripts) &&
                (ReferenceEquals(control, root) ? !hasLiteralSources :
                TryGetRuns(equation with { Content = new([(DocxMathNode)control]) }, token, out var sources) &&
                sources.All(item => item.Run.Text.Text.Length == 0))) metricOnly.Add(control);
        }
        return metricOnly;
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
