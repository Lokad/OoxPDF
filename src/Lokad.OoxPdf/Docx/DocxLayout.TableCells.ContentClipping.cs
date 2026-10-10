namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    private static bool IsEligibleDeclaredTopContentTable(
        DocxTable table, IDocxTextMeasurer? measurer, double fixedScale)
    {
        if (fixedScale != 1d || table.Rows.Count is < 1 or > 2 ||
            table.ColumnWidthsPoints.Count is < 1 or > 2)
            return false;
        return table.Rows.All(row =>
            row.Cells.Count == table.ColumnWidthsPoints.Count &&
            row.HeightPoints is > 0d && double.IsFinite(row.HeightPoints.Value) &&
            (string.Equals(row.HeightRuleValue, "exact", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(row.HeightRuleValue, "atLeast", StringComparison.OrdinalIgnoreCase)) &&
            row.Cells.All(cell => cell.VerticalAlignment == DocxTableCellVerticalAlignment.Top &&
                (cell.VerticalAlignmentValue is null ||
                    string.Equals(cell.VerticalAlignmentValue, "top", StringComparison.OrdinalIgnoreCase)) &&
                cell.Borders.All(border => DocxTableBorderGeometry.ResolveVisibleWidth(border) <= 0d ||
                    string.Equals(border.Value, "single", StringComparison.OrdinalIgnoreCase) ||
                    border.UseDoubleBandWidth && string.Equals(border.Value, "double", StringComparison.OrdinalIgnoreCase)) &&
                IsEligibleTableMinimumCell(table, row, cell, measurer) &&
                cell.Paragraphs.All(p => p.Runs.Count != 0) &&
                cell.Paragraphs.Any(p => p.EffectiveProperties.LineSpacingPoints is > 0d &&
                    string.Equals(p.EffectiveProperties.Spacing.LineRuleValue, "atLeast", StringComparison.OrdinalIgnoreCase))));
    }

    private static bool IsEligibleAutomaticTopContentTable(
        DocxTable table, IDocxTextMeasurer? measurer, double fixedScale)
    {
        if (fixedScale != 1d || table.Rows.Count is < 1 or > 2 ||
            table.ColumnWidthsPoints.Count is < 1 or > 2)
            return false;
        // Top masks have independent admission. They never enable the
        // centered/bottom paragraph-slot alignment calculation.
        return table.Rows.All(row =>
            row.Cells.Count == table.ColumnWidthsPoints.Count &&
            HasAutomaticTableRowHeight(row.HeightPoints, row.HeightRuleValue) &&
            row.Cells.All(cell => cell.VerticalAlignment == DocxTableCellVerticalAlignment.Top &&
                (cell.VerticalAlignmentValue is null ||
                    string.Equals(cell.VerticalAlignmentValue, "top", StringComparison.OrdinalIgnoreCase)) &&
                IsEligibleTableMinimumCell(table, row, cell, measurer) &&
                cell.Paragraphs.All(p => p.Runs.Count != 0) &&
                cell.Paragraphs.Any(p => p.EffectiveProperties.LineSpacingPoints is > 0d &&
                    string.Equals(p.EffectiveProperties.Spacing.LineRuleValue, "atLeast", StringComparison.OrdinalIgnoreCase))));
    }
}
