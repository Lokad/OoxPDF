namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    private static IReadOnlyList<DocxTextLineLayout> AdjustTableMinimumLinePlacement(
        DocxTable table, DocxTableRow row, DocxTableCell cell,
        IReadOnlyList<DocxTextLineLayout> lines, IDocxTextMeasurer? measurer)
    {
        if (row.HeightPoints is not null || row.IsHeader ||
            table.Revisions.Count != 0 || row.Revisions.Count != 0 || cell.Revisions.Count != 0 ||
            table.CellSpacingPoints is > 0d || table.UseLegacyTableGrid || cell.GridSpan != 1 ||
            cell.HasVerticalMerge || cell.NoWrap || cell.FitText ||
            cell.VerticalAlignment != DocxTableCellVerticalAlignment.Top ||
            cell.TextDirectionValue is not (null or "lrTb") ||
            cell.Paragraphs.Count is < 1 or > 2 || cell.BodyElements.Count != cell.Paragraphs.Count ||
            cell.BodyElements.Any(e => e is not DocxParagraphElement || e.Revisions.Count != 0) ||
            cell.Paragraphs.Any(p => !IsPlainContextualSpacingParagraph(p)) ||
            !DocxLineMetrics.TableCellFontSizesUniform(cell.Paragraphs) ||
            measurer is not IDocxLineMetricsProvider metrics ||
            measurer is not IDocxStaticTextMetricsProvider ||
            measurer is not IDocxHheaDescenderProvider || measurer is not IDocxHheaLineGapProvider)
            return lines;
        DocxTextRun? firstRun = cell.Paragraphs.SelectMany(p => p.Runs).FirstOrDefault();
        if (firstRun is null || cell.Paragraphs.SelectMany(p => p.Runs).Any(r =>
            !string.Equals(r.EffectiveProperties.FontFamily, firstRun.EffectiveProperties.FontFamily, StringComparison.OrdinalIgnoreCase) ||
            r.EffectiveProperties.Bold != firstRun.EffectiveProperties.Bold ||
            r.EffectiveProperties.Italic != firstRun.EffectiveProperties.Italic))
            return lines;
        DocxParagraph current = cell.Paragraphs[^1];
        DocxTextLineLayout[] currentLines = lines.Where(l => ReferenceEquals(l.SourceParagraph, current)).ToArray();
        if (currentLines.Length <= 1 || currentLines[0].LineHeightSource != DocxLineHeightSource.AtLeastLineSpacing ||
            current.Runs.Count == 0 || cell.Paragraphs.Count == 2 &&
            lines.Count(l => ReferenceEquals(l.SourceParagraph, cell.Paragraphs[0])) != 1)
            return lines;
        double natural = metrics.MeasureSingleLineHeight(current.Runs[0], currentLines[0].FontSize);
        if (!(natural > 0d) || !double.IsFinite(natural)) return lines;
        double slotHeight = currentLines[0].LineHeight ?? 0d;
        if (!(slotHeight > 0d) || !double.IsFinite(slotHeight)) return lines;
        double excess = Math.Max(0d, slotHeight - natural);
        double inset = DocxLineMetrics.ResolveTableCellFirstBaselineInset([current], measurer) + excess;
        // Word puts the excess minimum-height slot above the first baseline.
        // Cursor advances and row measurement already include this space.
        return lines.Select(l => ReferenceEquals(l.SourceParagraph, current) ? l with
        {
            BaselineY = l.BaselineY - excess,
            BodyLineBoxBaselineInsetPoints = inset,
            BodyLineBoxHeightPoints = l.LineHeight
        } : l).ToArray();
    }
}
