namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
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
