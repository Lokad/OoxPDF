namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    private static IReadOnlyList<DocxTextLineLayout> AdjustTableMinimumLinePlacement(
        DocxTable table, DocxTableRow row, DocxTableCell cell,
        IReadOnlyList<DocxTextLineLayout> lines, IDocxTextMeasurer? measurer,
        double cellHeight, double rowTopPadding, double usedHeight, bool isLastRow, double marginAdvance,
        bool automaticMinimumPlacement)
    {
        if (!lines.Any(l => l.LineHeightSource == DocxLineHeightSource.AtLeastLineSpacing)) return lines;
        if (!IsEligibleTableMinimumCell(table, row, cell, measurer)) return lines;
        var metrics = (IDocxLineMetricsProvider)measurer!;
        double alignmentShift = 0d;
        // Single-cell automatic tables already have qualified placement. Grid
        // alignment must not move their existing exact/minimum overflow slots.
        bool alignAutomaticGrid = automaticMinimumPlacement &&
            (table.Rows.Count > 1 || row.Cells.Count > 1);
        if (cell.VerticalAlignment != DocxTableCellVerticalAlignment.Top &&
            (!automaticMinimumPlacement || alignAutomaticGrid))
        {
            // Word aligns the paragraph slot extent. The terminal bottom border
            // hangs below it; the first-baseline inset is already applied.
            if (!automaticMinimumPlacement && (row.HeightPoints is not > 0d ||
                string.Equals(row.HeightRuleValue, "auto", StringComparison.OrdinalIgnoreCase) ||
                cell.Paragraphs.Any(p => p.Runs.Count == 0))) return lines;
            double topPadding = rowTopPadding + ResolveTableCellTopBorderContentInset(cell, 1d);
            double bottomPadding = ResolveTableCellVerticalPadding(cell.Margins.BottomPoints, 1d);
            double baselineInset = ResolveTableCellFirstBaselineInset(cell.Paragraphs, measurer);
            double borderAdvance = isLastRow ? row.Cells
                .Select(c => DocxTableBorderGeometry.ResolveVisibleWidth(DocxTableBorderGeometry.Find(c.Borders, "bottom")))
                .DefaultIfEmpty(0d).Max() : 0d;
            if (!automaticMinimumPlacement && bottomPadding > 0d && marginAdvance == 0d) return lines;
            double oldExtra = Math.Max(0d, cellHeight - topPadding - bottomPadding - baselineInset - usedHeight);
            // Align the complete slot extent using the row's maximum bottom
            // margin. A shorter cell's own padding does not enlarge that extent.
            double alignmentBottom = automaticMinimumPlacement
                ? row.Cells.Select(c => ResolveTableCellVerticalPadding(c.Margins.BottomPoints, 1d)).Max()
                : marginAdvance;
            double newExtra = Math.Max(0d, cellHeight - topPadding - alignmentBottom - borderAdvance - usedHeight);
            alignmentShift = ResolveVerticalContentOffset(cell, newExtra) - ResolveVerticalContentOffset(cell, oldExtra);
        }
        var byParagraph = new Dictionary<DocxParagraph, List<DocxTextLineLayout>>(ReferenceEqualityComparer.Instance);
        foreach (DocxTextLineLayout line in lines)
        {
            if (line.SourceParagraph is not { } paragraph) continue;
            if (!byParagraph.TryGetValue(paragraph, out var paragraphLines))
                byParagraph.Add(paragraph, paragraphLines = []);
            paragraphLines.Add(line);
        }
        var adjustments = new Dictionary<DocxParagraph, (double Excess, double Inset)>(ReferenceEqualityComparer.Instance);
        foreach ((DocxParagraph paragraph, List<DocxTextLineLayout> paragraphLines) in byParagraph)
        {
            if (paragraph.Runs.Count == 0 ||
                paragraphLines[0].LineHeightSource != DocxLineHeightSource.AtLeastLineSpacing)
                continue;
            double natural = metrics.MeasureSingleLineHeight(paragraph.Runs[0], paragraphLines[0].FontSize);
            double slotHeight = paragraphLines[0].LineHeight ?? 0d;
            if (!(natural > 0d) || !double.IsFinite(natural) || !(slotHeight > 0d) || !double.IsFinite(slotHeight)) continue;
            // A natural-height floor has no excess and keeps its existing emission inset.
            if (automaticMinimumPlacement && !(paragraph.EffectiveProperties.LineSpacingPoints is { } minimum && minimum > natural))
                continue;
            double excess = Math.Max(0d, slotHeight - natural);
            double inset = DocxLineMetrics.ResolveTableCellBodyBaselineOffset(
                paragraphLines[0].FontSize, slotHeight, false, cell.Paragraphs, paragraph, measurer) + excess;
            adjustments.Add(paragraph, (excess, inset));
        }
        if (adjustments.Count == 0 && alignmentShift == 0d) return lines;
        // Word puts the excess minimum-height slot above the first baseline.
        // Cursor advances and row measurement already include this space.
        return lines.Select(l => l.SourceParagraph is { } p && adjustments.TryGetValue(p, out var adjustment) ? l with
        {
            BaselineY = l.BaselineY - adjustment.Excess - alignmentShift,
            BodyLineBoxBaselineInsetPoints = adjustment.Inset,
            BodyLineBoxHeightPoints = l.LineHeight
        } : alignmentShift == 0d ? l : l with { BaselineY = l.BaselineY - alignmentShift }).ToArray();
    }
    private static bool IsEligibleTableMinimumCell(
        DocxTable table, DocxTableRow row, DocxTableCell cell, IDocxTextMeasurer? measurer)
    {
        // Explicit automatic height ignores its authored value, as does row measurement.
        if (row.HeightPoints is { } height &&
            !string.Equals(row.HeightRuleValue, "auto", StringComparison.OrdinalIgnoreCase) &&
            (!(height > 0d) || !double.IsFinite(height) ||
            row.HeightRuleValue is not null &&
            !string.Equals(row.HeightRuleValue, "exact", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(row.HeightRuleValue, "atLeast", StringComparison.OrdinalIgnoreCase)))
            return false;
        if (row.IsHeader ||
            table.Revisions.Count != 0 || row.Revisions.Count != 0 || cell.Revisions.Count != 0 ||
            table.CellSpacingPoints is > 0d || table.UseLegacyTableGrid || cell.GridSpan != 1 ||
            cell.HasVerticalMerge || cell.NoWrap || cell.FitText ||
            cell.TextDirectionValue is not (null or "lrTb") ||
            cell.Paragraphs.Count == 0 || cell.BodyElements.Count != cell.Paragraphs.Count ||
            cell.BodyElements.Any(e => e is not DocxParagraphElement || e.Revisions.Count != 0) ||
            cell.Paragraphs.Any(p => !IsPlainContextualSpacingParagraph(p)) ||
            measurer is not IDocxLineMetricsProvider ||
            measurer is not IDocxStaticTextMetricsProvider ||
            measurer is not IDocxHheaDescenderProvider || measurer is not IDocxHheaLineGapProvider)
            return false;
        // Font changes between plain paragraphs have independent line metrics.
        // Rich runs within one paragraph retain the earlier placement model.
        if (cell.Paragraphs.Any(p => p.Runs.Count > 0 && p.Runs.Skip(1).Any(r =>
            !string.Equals(r.EffectiveProperties.FontFamily, p.Runs[0].EffectiveProperties.FontFamily, StringComparison.OrdinalIgnoreCase) ||
            r.EffectiveProperties.FontSize != p.Runs[0].EffectiveProperties.FontSize ||
            r.EffectiveProperties.Bold != p.Runs[0].EffectiveProperties.Bold ||
            r.EffectiveProperties.Italic != p.Runs[0].EffectiveProperties.Italic)))
            return false;
        return true;
    }

    private static double ResolveAlignedMinimumRowBottomMargin(
        DocxTable table, DocxTableRow row, IDocxTextMeasurer? measurer,
        double fixedScale, double? measuredContentHeight = null)
    {
        if (fixedScale != 1d || row.HeightPoints is not > 0d ||
            string.Equals(row.HeightRuleValue, "auto", StringComparison.OrdinalIgnoreCase) ||
            row.Cells.Count == 0 || row.Cells.Any(cell =>
                cell.VerticalAlignment == DocxTableCellVerticalAlignment.Top ||
                !IsEligibleTableMinimumCell(table, row, cell, measurer) ||
                cell.Paragraphs.Any(p => p.Runs.Count == 0) ||
                !cell.Paragraphs.Any(p => p.EffectiveProperties.LineSpacingPoints is > 0d &&
                    string.Equals(p.EffectiveProperties.Spacing.LineRuleValue, "atLeast", StringComparison.OrdinalIgnoreCase))))
            return 0d;
        // Content-dominated minimum rows retain their earlier geometry.
        if (!string.Equals(row.HeightRuleValue, "exact", StringComparison.OrdinalIgnoreCase) &&
            measuredContentHeight is { } content && content > row.HeightPoints.Value + ResolveTableRowTopPadding(row, fixedScale))
            return 0d;
        return row.Cells.Select(c => ResolveTableCellVerticalPadding(c.Margins.BottomPoints, fixedScale)).Max();
    }

    internal static bool HasAutomaticTableRowHeight(double? height, string? rule)
    {
        return (height is null && rule is null) ||
            string.Equals(rule, "auto", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsEligibleAutomaticMinimumTable(
        DocxTable table, IDocxTextMeasurer? measurer, double fixedScale, bool allowGrid)
    {
        if (fixedScale != 1d || table.Rows.Count == 0 ||
            (!allowGrid && (table.Rows.Count != 1 || table.Rows[0].Cells.Count != 1)))
            return false;
        bool allAutomatic = table.Rows.All(row => HasAutomaticTableRowHeight(row.HeightPoints, row.HeightRuleValue));
        // Mixed-height admission is qualified only for complete two-by-two
        // body grids. Other composed shapes retain their previous fallback.
        if (!allAutomatic && (!allowGrid || table.Rows.Count != 2 || table.Rows.Any(row => row.Cells.Count != 2)))
            return false;
        return table.Rows.Any(row => HasAutomaticTableRowHeight(row.HeightPoints, row.HeightRuleValue)) &&
            table.Rows.All(row => row.Cells.Count > 0 &&
            (HasAutomaticTableRowHeight(row.HeightPoints, row.HeightRuleValue) ||
                row.HeightPoints is > 0d &&
                (string.Equals(row.HeightRuleValue, "exact", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(row.HeightRuleValue, "atLeast", StringComparison.OrdinalIgnoreCase))) &&
            row.Cells.All(cell => cell.VerticalAlignment != DocxTableCellVerticalAlignment.Top &&
                IsEligibleTableMinimumCell(table, row, cell, measurer) &&
                cell.Paragraphs.All(p => p.Runs.Count != 0) &&
                cell.Paragraphs.Any(p => p.EffectiveProperties.LineSpacingPoints is > 0d &&
                    string.Equals(p.EffectiveProperties.Spacing.LineRuleValue, "atLeast", StringComparison.OrdinalIgnoreCase))));
    }

}
