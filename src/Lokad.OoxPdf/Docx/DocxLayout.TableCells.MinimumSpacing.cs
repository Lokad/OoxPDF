namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    private static IReadOnlyList<DocxTextLineLayout> AdjustTableMinimumLinePlacement(
        DocxTable table, DocxTableRow row, DocxTableCell cell,
        IReadOnlyList<DocxTextLineLayout> lines, IDocxTextMeasurer? measurer,
        double cellHeight, double rowTopPadding, double usedHeight, bool isLastRow)
    {
        if (!lines.Any(l => l.LineHeightSource == DocxLineHeightSource.AtLeastLineSpacing)) return lines;
        // Explicit automatic height ignores its authored value, as does row measurement.
        if (row.HeightPoints is { } height &&
            !string.Equals(row.HeightRuleValue, "auto", StringComparison.OrdinalIgnoreCase) &&
            (!(height > 0d) || !double.IsFinite(height) ||
            row.HeightRuleValue is not null &&
            !string.Equals(row.HeightRuleValue, "exact", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(row.HeightRuleValue, "atLeast", StringComparison.OrdinalIgnoreCase)))
            return lines;
        if (row.IsHeader ||
            table.Revisions.Count != 0 || row.Revisions.Count != 0 || cell.Revisions.Count != 0 ||
            table.CellSpacingPoints is > 0d || table.UseLegacyTableGrid || cell.GridSpan != 1 ||
            cell.HasVerticalMerge || cell.NoWrap || cell.FitText ||
            cell.TextDirectionValue is not (null or "lrTb") ||
            cell.Paragraphs.Count == 0 || cell.BodyElements.Count != cell.Paragraphs.Count ||
            cell.BodyElements.Any(e => e is not DocxParagraphElement || e.Revisions.Count != 0) ||
            cell.Paragraphs.Any(p => !IsPlainContextualSpacingParagraph(p)) ||
            measurer is not IDocxLineMetricsProvider metrics ||
            measurer is not IDocxStaticTextMetricsProvider ||
            measurer is not IDocxHheaDescenderProvider || measurer is not IDocxHheaLineGapProvider)
            return lines;
        // Font changes between plain paragraphs have independent line metrics.
        // Rich runs within one paragraph retain the earlier placement model.
        if (cell.Paragraphs.Any(p => p.Runs.Count > 0 && p.Runs.Skip(1).Any(r =>
            !string.Equals(r.EffectiveProperties.FontFamily, p.Runs[0].EffectiveProperties.FontFamily, StringComparison.OrdinalIgnoreCase) ||
            r.EffectiveProperties.FontSize != p.Runs[0].EffectiveProperties.FontSize ||
            r.EffectiveProperties.Bold != p.Runs[0].EffectiveProperties.Bold ||
            r.EffectiveProperties.Italic != p.Runs[0].EffectiveProperties.Italic)))
            return lines;
        double alignmentShift = 0d;
        if (cell.VerticalAlignment != DocxTableCellVerticalAlignment.Top)
        {
            // Word aligns the paragraph slot extent. The terminal bottom border
            // hangs below it; the first-baseline inset is already applied.
            if (row.HeightPoints is not > 0d ||
                string.Equals(row.HeightRuleValue, "auto", StringComparison.OrdinalIgnoreCase) ||
                cell.Margins.TopPoints is > 0d || cell.Margins.BottomPoints is > 0d ||
                cell.Paragraphs.Any(p => p.Runs.Count == 0)) return lines;
            double topPadding = rowTopPadding + ResolveTableCellTopBorderContentInset(cell, 1d);
            double bottomPadding = ResolveTableCellVerticalPadding(cell.Margins.BottomPoints, 1d);
            double baselineInset = ResolveTableCellFirstBaselineInset(cell.Paragraphs, measurer);
            double borderAdvance = isLastRow ? row.Cells
                .Select(c => DocxTableBorderGeometry.ResolveVisibleWidth(DocxTableBorderGeometry.Find(c.Borders, "bottom")))
                .DefaultIfEmpty(0d).Max() : 0d;
            double oldExtra = Math.Max(0d, cellHeight - topPadding - bottomPadding - baselineInset - usedHeight);
            double newExtra = Math.Max(0d, cellHeight - topPadding - bottomPadding - borderAdvance - usedHeight);
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
            double excess = Math.Max(0d, slotHeight - natural);
            double inset = DocxLineMetrics.ResolveTableCellBodyBaselineOffset(
                paragraphLines[0].FontSize, slotHeight, false, cell.Paragraphs, paragraph, measurer) + excess;
            adjustments.Add(paragraph, (excess, inset));
        }
        if (adjustments.Count == 0) return lines;
        // Word puts the excess minimum-height slot above the first baseline.
        // Cursor advances and row measurement already include this space.
        return lines.Select(l => l.SourceParagraph is { } p && adjustments.TryGetValue(p, out var adjustment) ? l with
        {
            BaselineY = l.BaselineY - adjustment.Excess - alignmentShift,
            BodyLineBoxBaselineInsetPoints = adjustment.Inset,
            BodyLineBoxHeightPoints = l.LineHeight
        } : alignmentShift == 0d ? l : l with { BaselineY = l.BaselineY - alignmentShift }).ToArray();
    }
}
