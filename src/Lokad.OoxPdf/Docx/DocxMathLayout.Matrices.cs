using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal static partial class DocxMathLayout
{
    private enum MatrixColumnAlignment { Center, Left, Right }

    // Word's PDF hides placeholder paint for every known plcHide value while
    // retaining the same metric box. Unknown properties remain rejected above.
    private static bool IsEmptyMatrixCell(DocxMathRow cell, CancellationToken token) => IsEmptyArgument(cell, token);

    internal static bool IsEmptyMatrix(DocxEquation equation, CancellationToken token) =>
        GetMatrix(equation, token) is { } matrix &&
        matrix.Rows.All(row => row.Cells.All(cell => IsEmptyMatrixCell(cell, token)));

    private static DocxMathMatrix? GetMatrix(DocxEquation equation, CancellationToken token) =>
        GetSupportedMatrix(GetSingleSupportedChild(equation), token);

    private static DocxMathMatrix? GetSupportedMatrix(DocxMathNode? node, CancellationToken token)
    {
        if (node is not DocxMathMatrix { HasUnsupportedProperties: false, Rows.Count: > 0 } matrix ||
            matrix.BaseJustificationValue is not (null or "center" or "top" or "bot") ||
            !TryResolveMatrixSpacing(matrix, 1d, out _, out _) ||
            matrix.Rows[0].Cells.Count == 0 || matrix.ControlText is { } control && !DocxMathRowLayout.HasSupportedTextFormatting(control, allowBold: false)) { return null; }
        foreach (var row in matrix.Rows)
        {
            token.ThrowIfCancellationRequested();
            if (row.HasUnsupportedProperties || row.PropertiesXml.Count != 0 || row.Cells.Count != matrix.Rows[0].Cells.Count) { return null; }
        }
        return TryResolveMatrixColumns(matrix.ColumnGroups, matrix.Rows[0].Cells.Count, [], token) ? matrix : null;
    }

    private static bool TryResolveMatrixSpacing(DocxMathMatrix matrix, double size, out double minimumWidth, out double gap)
    {
        uint rule = matrix.ColumnGapRuleValue ?? 0, amount = matrix.ColumnGapValue ?? 0;
        minimumWidth = (matrix.MinimumColumnWidthTwips ?? 0) / 20d;
        gap = rule switch { 0 or 1 or 2 => size * (1d + rule / 2d), 3 => amount / 20d, 4 => amount * size / 2d, _ => double.NaN };
        return rule <= 4 && amount <= ushort.MaxValue && double.IsFinite(size) && size > 0d && double.IsFinite(gap) && gap >= 0d;
    }

    private static bool TryResolveMatrixColumns(IReadOnlyList<DocxMathMatrixColumn>? groups, int columns, Span<MatrixColumnAlignment> destination, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        if (!destination.IsEmpty && destination.Length != columns) { return false; }
        if (groups is null) { destination.Fill(MatrixColumnAlignment.Center); return true; }
        int used = 0;
        foreach (var group in groups)
        {
            token.ThrowIfCancellationRequested();
            int count = group.CountValue ?? 1;
            if (group.HasUnsupportedProperties || count <= 0 || count > columns - used ||
                group.JustificationValue is not (null or "center" or "left" or "right")) { return false; }
            var alignment = group.JustificationValue switch { "left" => MatrixColumnAlignment.Left, "right" => MatrixColumnAlignment.Right, _ => MatrixColumnAlignment.Center };
            if (!destination.IsEmpty) { destination.Slice(used, count).Fill(alignment); }
            used += count;
        }
        return used == columns;
    }

    private static bool TryMeasureMatrix(DocxEquation equation, DocxMathMatrix matrix,
        Func<DocxMathRun, PdfFontResource?> resolveFont, CancellationToken token,
        out DocxMathBox? box, out string? fallback, Func<IDocxMathControl, PdfFontResource?>? resolveControlFont,
        Func<IDocxMathControl, OpenTypeFont?>? resolveControlMetrics, double? rowControlSize = null)
    {
        box = null;
        if (!TryGetRuns(equation, token, out var sources))
        { fallback = "UnsupportedMathMatrixSource"; return false; }
        OpenTypeFont? font = resolveControlMetrics is not null ? resolveControlMetrics(matrix) :
            resolveControlFont is not null ? resolveControlFont(matrix)?.Font.Font :
            sources.Count == 0 ? null : resolveFont(sources[0].Run)?.Font.Font;
        if (font is null || !font.TryGetMathConstants(out var constants))
        { fallback = "MissingMathMatrixMetrics"; return false; }
        double size = rowControlSize ?? ResolveControlSize(matrix, equation, sources);
        if (!double.IsFinite(size) || size <= 0d)
        { fallback = "InvalidMathMatrixSize"; return false; }
        if (!TryResolveMatrixSpacing(matrix, size, out double minimumWidth, out double gap))
        { fallback = "InvalidMathMatrixSpacing"; return false; }
        double glyphSize = OfficePdfTextEmissionProfile.FontSize(size);
        double baselineFloor = DocxLineMetrics.MeasureTypographicBaselineInset(font, glyphSize);
        double descentFloor = DocxLineMetrics.MeasureTypographicLineHeight(font, glyphSize) - baselineFloor;
        int columns = matrix.Rows[0].Cells.Count;
        // Resolve into the actual source width, never a declared group count.
        var alignments = new MatrixColumnAlignment[columns];
        if (!TryResolveMatrixColumns(matrix.ColumnGroups, columns, alignments, token))
        { fallback = "UnsupportedMathMatrixColumns"; return false; }
        var widths = new double[columns]; var rowAbove = new double[matrix.Rows.Count]; var rowBelow = new double[matrix.Rows.Count];
        Array.Fill(widths, minimumWidth);
        var cells = new List<DocxMathBox[]>(matrix.Rows.Count);
        DocxMathBox? placeholder = null;
        double height = 0d;
        for (int r = 0; r < matrix.Rows.Count; r++)
        {
            token.ThrowIfCancellationRequested(); var measured = new DocxMathBox[columns];
            rowAbove[r] = baselineFloor; rowBelow[r] = descentFloor;
            for (int c = 0; c < columns; c++)
            {
                token.ThrowIfCancellationRequested();
                if (IsEmptyMatrixCell(matrix.Rows[r].Cells[c], token))
                {
                    OoxConversionBudget.Current?.ChargeSceneNodes(1);
                    if (placeholder is null)
                    {
                        if (!TryMeasureHiddenPlaceholder(font, glyphSize, out placeholder))
                        { fallback = "MissingMathMatrixPlaceholderMetrics"; return false; }
                    }
                    measured[c] = placeholder; widths[c] = Math.Max(widths[c], placeholder.AdvanceWidth);
                    rowAbove[r] = Math.Max(rowAbove[r], placeholder.AboveBaseline); rowBelow[r] = Math.Max(rowBelow[r], placeholder.BelowBaseline);
                    continue;
                }
                if (!TryMeasureBoxArgument(equation with { Content = matrix.Rows[r].Cells[c], ParentFontSize = size },
                    resolveFont, font, size, token, out var cell, out fallback, neighbors: new(After: DocxMathRowLayout.AtomKind.Ordinary),
                    resolveControlFont: resolveControlFont, resolveControlMetrics: resolveControlMetrics) ||
                    cell is null || cell.Glyphs.Count == 0 && cell.AdvanceWidth <= 0d && cell.AboveBaseline + cell.BelowBaseline <= 0d)
                { fallback ??= "UnsupportedMathMatrixCell"; return false; }
                measured[c] = cell; widths[c] = Math.Max(widths[c], cell.AdvanceWidth);
                // The control owns each ordinary row frame. Cell ink can grow
                // it; a painted run's unrelated font-line gap cannot. MathLeading
                // belongs to the composed formula, not each internal matrix row.
                rowAbove[r] = Math.Max(rowAbove[r], cell.AboveBaseline); rowBelow[r] = Math.Max(rowBelow[r], cell.BelowBaseline);
            }
            cells.Add(measured); height += rowAbove[r] + rowBelow[r];
        }
        double frameWidth = widths.Sum() + (columns - 1d) * gap;
        double axis = constants.GetValue(OpenTypeMathConstant.AxisHeight) * size / font.UnitsPerEm;
        if (!double.IsFinite(height) || height <= 0d || !double.IsFinite(frameWidth) || frameWidth <= 0d || !double.IsFinite(axis))
        { fallback = "InvalidMathMatrixExtent"; return false; }
        var glyphs = new List<DocxMathGlyphPlacement>(); var rules = new List<DocxMathRulePlacement>();
        // Top/bottom alignment anchors the corresponding row baseline. Center
        // alignment keeps the complete row frame centered on the math axis.
        double top = matrix.BaseJustificationValue switch
        {
            "top" => rowAbove[0],
            "bot" => height - rowBelow[^1],
            _ => axis + height / 2d
        };
        double advance = 0d, width = 0d, inkAbove = 0d, inkBelow = 0d;
        for (int r = 0; r < cells.Count; r++)
        {
            double baseline = top - rowAbove[r], x = 0d;
            for (int c = 0; c < columns; c++)
            {
                token.ThrowIfCancellationRequested(); var child = cells[r][c];
                double spare = widths[c] - child.AdvanceWidth;
                double cellX = x + (alignments[c] switch { MatrixColumnAlignment.Left => 0d, MatrixColumnAlignment.Right => spare, _ => spare / 2d });
                // Slots control alignment and the next column origin. The
                // formula ends at its last argument's logical edge, including
                // hidden placeholder boxes, rather than unused trailing padding.
                advance = Math.Max(advance, cellX + child.AdvanceWidth);
                width = Math.Max(width, cellX + child.Width); inkAbove = Math.Max(inkAbove, baseline + child.AboveBaseline);
                inkBelow = Math.Max(inkBelow, child.BelowBaseline - baseline);
                foreach (var glyph in child.Glyphs)
                { token.ThrowIfCancellationRequested(); glyphs.Add(glyph with { X = glyph.X + cellX, BaselineY = glyph.BaselineY + baseline }); }
                foreach (var rule in child.Rules)
                { token.ThrowIfCancellationRequested(); rules.Add(rule with { X = rule.X + cellX, Y = rule.Y + baseline }); }
                x += widths[c] + gap;
            }
            top -= rowAbove[r] + rowBelow[r];
        }
        box = new(width, inkAbove, inkBelow, glyphs)
        { LogicalAdvance = advance == width ? null : advance, ParentMetrics = new(font, size), Rules = rules,
            EnclosureHeight = inkAbove + inkBelow };
        fallback = null;
        return true;
    }
}
