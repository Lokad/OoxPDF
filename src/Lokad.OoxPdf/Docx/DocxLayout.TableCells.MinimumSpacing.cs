namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    private static IReadOnlyList<DocxTextLineLayout> AdjustTableMinimumLinePlacement(
        DocxTable table, DocxTableRow row, DocxTableCell cell,
        IReadOnlyList<DocxTextLineLayout> lines, IDocxTextMeasurer? measurer)
    {
        if (!lines.Any(l => l.LineHeightSource == DocxLineHeightSource.AtLeastLineSpacing)) return lines;
        if (row.HeightPoints is not null || row.IsHeader ||
            table.Revisions.Count != 0 || row.Revisions.Count != 0 || cell.Revisions.Count != 0 ||
            table.CellSpacingPoints is > 0d || table.UseLegacyTableGrid || cell.GridSpan != 1 ||
            cell.HasVerticalMerge || cell.NoWrap || cell.FitText ||
            cell.VerticalAlignment != DocxTableCellVerticalAlignment.Top ||
            cell.TextDirectionValue is not (null or "lrTb") ||
            cell.Paragraphs.Count == 0 || cell.BodyElements.Count != cell.Paragraphs.Count ||
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
            double inset = DocxLineMetrics.ResolveTableCellFirstBaselineInset([paragraph], measurer) + excess;
            adjustments.Add(paragraph, (excess, inset));
        }
        if (adjustments.Count == 0) return lines;
        // Word puts the excess minimum-height slot above the first baseline.
        // Cursor advances and row measurement already include this space.
        return lines.Select(l => l.SourceParagraph is { } p && adjustments.TryGetValue(p, out var adjustment) ? l with
        {
            BaselineY = l.BaselineY - adjustment.Excess,
            BodyLineBoxBaselineInsetPoints = adjustment.Inset,
            BodyLineBoxHeightPoints = l.LineHeight
        } : l).ToArray();
    }
}
