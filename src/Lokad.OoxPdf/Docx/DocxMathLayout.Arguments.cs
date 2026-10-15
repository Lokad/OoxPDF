using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal static partial class DocxMathLayout
{
    // Discovery admits literals, one fraction/stack or a supported script box.
    // Its literal child checks bound composition before any measurement recurs.
    private static bool TryMeasureBoxArgument(DocxEquation equation, Func<DocxMathRun, PdfFontResource?> resolveFont,
        OpenTypeFont metricFont, double spacingSize, CancellationToken token, out DocxMathBox? child, out string? fallback,
        DocxMathRowLayout.Neighbors neighbors = default, Func<IDocxMathControl, PdfFontResource?>? resolveControlFont = null,
        Func<IDocxMathControl, OpenTypeFont?>? resolveControlMetrics = null) =>
        GetFraction(equation) is not null || GetScripts(equation) is not null || GetMatrix(equation, token) is not null
            ? TryMeasure(equation, resolveFont, token, out child, out fallback, metricFont, resolveControlFont, resolveControlMetrics, rowControlSize: spacingSize)
            : DocxMathRowLayout.TryMeasure(equation.Content, resolveFont, token, out child, out fallback,
                exportFontSize: true, operatorSpacingFontSize: spacingSize, neighbors: neighbors);
}
