using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal static partial class DocxMathLayout
{
    // Discovery and measurement share the existing structured argument kinds.
    // Each child's source-role checks retain the bounded deeper fallback.
    private static bool HasStructuredBoxArgument(DocxEquation equation, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        return GetFraction(equation) is not null || GetScripts(equation) is not null || GetMatrix(equation, token) is not null ||
            HasMixedBoxArgument(equation, token);
    }

    private static bool HasMixedBoxArgument(DocxEquation equation, CancellationToken token)
    {
        // Reuse the ordered row composer for literal neighbors and the existing
        // bounded boxes. Container children are excluded: this role cannot
        // recursively admit another radical, n-ary, delimiter or function.
        if (!equation.Content.Children.Any(node => node is not DocxMathRun)) return false;
        return GetStructuredRow(equation, token) is { } row &&
            row.Parts.All(part => part.Structure is DocxMathFraction or DocxMathScripts or DocxMathMatrix);
    }

    private static bool TryMeasureBoxArgument(DocxEquation equation, Func<DocxMathRun, PdfFontResource?> resolveFont,
        OpenTypeFont metricFont, double spacingSize, CancellationToken token, out DocxMathBox? child, out string? fallback,
        DocxMathRowLayout.Neighbors neighbors = default, Func<IDocxMathControl, PdfFontResource?>? resolveControlFont = null,
        Func<IDocxMathControl, OpenTypeFont?>? resolveControlMetrics = null) =>
        HasStructuredBoxArgument(equation, token)
            ? TryMeasure(equation, resolveFont, token, out child, out fallback, metricFont, resolveControlFont, resolveControlMetrics, rowControlSize: spacingSize)
            : TryMeasureLiteralArgument(equation.Content, resolveFont, metricFont, spacingSize, 0, token,
                out child, out fallback, out _, exportFontSize: true, operatorSpacingFontSize: spacingSize, neighbors: neighbors);
}
