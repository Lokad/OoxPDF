using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal static partial class DocxMathLayout
{
    private static DocxMathMatrix? GetMatrix(DocxEquation equation) =>
        GetSingleSupportedChild(equation) is DocxMathMatrix { HasUnsupportedProperties: false, Rows.Count: > 0 } matrix &&
        matrix.Rows[0].Cells.Count > 0 && matrix.Rows.All(r => !r.HasUnsupportedProperties && r.PropertiesXml.Count == 0 &&
            r.Cells.Count == matrix.Rows[0].Cells.Count) &&
        (matrix.ControlText is null || DocxMathRowLayout.HasSupportedTextFormatting(matrix.ControlText, allowBold: false)) ? matrix : null;

    private static bool TryMeasureMatrix(DocxEquation equation, DocxMathMatrix matrix,
        Func<DocxMathRun, PdfFontResource?> resolveFont, CancellationToken token,
        out DocxMathBox? box, out string? fallback, Func<IDocxMathControl, PdfFontResource?>? resolveControlFont,
        Func<IDocxMathControl, OpenTypeFont?>? resolveControlMetrics)
    {
        box = null;
        if (!TryGetRuns(equation, token, out var sources) || sources.Count == 0)
        { fallback = "UnsupportedMathMatrixSource"; return false; }
        OpenTypeFont? font = resolveControlMetrics is not null ? resolveControlMetrics(matrix) :
            resolveControlFont is not null ? resolveControlFont(matrix)?.Font.Font : resolveFont(sources[0].Run)?.Font.Font;
        if (font is null || !font.TryGetMathConstants(out var constants))
        { fallback = "MissingMathMatrixMetrics"; return false; }
        double size = matrix.ControlText?.FontSize ?? equation.ParentFontSize ?? sources[0].Run.Text.FontSize;
        if (!double.IsFinite(size) || size <= 0d)
        { fallback = "InvalidMathMatrixSize"; return false; }
        double glyphSize = OfficePdfTextEmissionProfile.FontSize(size);
        double baselineFloor = DocxLineMetrics.MeasureTypographicBaselineInset(font, glyphSize);
        double descentFloor = DocxLineMetrics.MeasureTypographicLineHeight(font, glyphSize) - baselineFloor;
        int columns = matrix.Rows[0].Cells.Count;
        var widths = new double[columns]; var rowAbove = new double[matrix.Rows.Count]; var rowBelow = new double[matrix.Rows.Count];
        var cells = new List<DocxMathBox[]>(matrix.Rows.Count);
        double height = 0d;
        for (int r = 0; r < matrix.Rows.Count; r++)
        {
            token.ThrowIfCancellationRequested(); var measured = new DocxMathBox[columns];
            rowAbove[r] = baselineFloor; rowBelow[r] = descentFloor;
            for (int c = 0; c < columns; c++)
            {
                token.ThrowIfCancellationRequested();
                if (!DocxMathRowLayout.TryMeasure(matrix.Rows[r].Cells[c], resolveFont, token, out var cell, out fallback,
                    exportFontSize: true, operatorSpacingFontSize: size, neighbors: new(After: DocxMathRowLayout.AtomKind.Ordinary)) ||
                    cell is null || cell.Glyphs.Count == 0)
                { fallback ??= "UnsupportedMathMatrixCell"; return false; }
                measured[c] = cell; widths[c] = Math.Max(widths[c], cell.AdvanceWidth);
                // The control owns each ordinary row frame. Cell ink can grow
                // it; a painted run's unrelated font-line gap cannot. MathLeading
                // belongs to the composed formula, not each internal matrix row.
                rowAbove[r] = Math.Max(rowAbove[r], cell.AboveBaseline); rowBelow[r] = Math.Max(rowBelow[r], cell.BelowBaseline);
            }
            cells.Add(measured); height += rowAbove[r] + rowBelow[r];
        }
        double advance = widths.Sum() + (columns - 1d) * size;
        double axis = constants.GetValue(OpenTypeMathConstant.AxisHeight) * size / font.UnitsPerEm;
        if (!double.IsFinite(height) || height <= 0d || !double.IsFinite(advance) || advance <= 0d || !double.IsFinite(axis))
        { fallback = "InvalidMathMatrixExtent"; return false; }
        var glyphs = new List<DocxMathGlyphPlacement>();
        double top = axis + height / 2d, width = advance, inkAbove = 0d, inkBelow = 0d;
        for (int r = 0; r < cells.Count; r++)
        {
            double baseline = top - rowAbove[r], x = 0d;
            for (int c = 0; c < columns; c++)
            {
                token.ThrowIfCancellationRequested(); var child = cells[r][c];
                double cellX = x + (widths[c] - child.AdvanceWidth) / 2d;
                width = Math.Max(width, cellX + child.Width); inkAbove = Math.Max(inkAbove, baseline + child.AboveBaseline);
                inkBelow = Math.Max(inkBelow, child.BelowBaseline - baseline);
                foreach (var glyph in child.Glyphs)
                { token.ThrowIfCancellationRequested(); glyphs.Add(glyph with { X = glyph.X + cellX, BaselineY = glyph.BaselineY + baseline }); }
                x += widths[c] + size;
            }
            top -= rowAbove[r] + rowBelow[r];
        }
        box = new(width, inkAbove, inkBelow, glyphs)
        { LogicalAdvance = advance == width ? null : advance, ParentMetrics = new(font, size) };
        fallback = null;
        return true;
    }
}
