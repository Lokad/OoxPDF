namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    private static (IReadOnlyList<DocxTextLineLayout> Lines, IReadOnlyList<DocxInlineImageLayout> Images) AdjustTableMinimumLinePlacement(
        DocxTable table, DocxTableRow row, DocxTableCell cell,
        IReadOnlyList<DocxTextLineLayout> lines, IReadOnlyList<DocxInlineImageLayout> images, IDocxTextMeasurer? measurer,
        double cellHeight, double rowTopPadding, double usedHeight, bool isLastRow, double marginAdvance,
        bool automaticMinimumPlacement)
    {
        if (!lines.Any(l => l.LineHeightSource == DocxLineHeightSource.AtLeastLineSpacing)) return (lines, images);
        if (!IsEligibleTableMinimumCell(table, row, cell, measurer)) return (lines, images);
        // Every admitted picture must already belong to a resolved inline line.
        // Block pictures and incomplete source ownership retain their prior model.
        if (images.Count != cell.Paragraphs.Sum(p => p.Images.Count) || images.Any(image =>
            image.LineFrame is not { IsUsable: true } || image.SourceParagraphIndex is not { } index ||
            index < 0 || index >= cell.Paragraphs.Count ||
            !cell.Paragraphs[index].Images.Any(source => ReferenceEquals(source, image.Image))))
            return (lines, images);
        var metrics = (IDocxLineMetricsProvider)measurer!;
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
        var lineOwners = images.Count == 0 ? null : lines.ToLookup(line => (line.SourceParagraphIndex, line.BaselineY));
        var imageOwners = new Dictionary<DocxInlineImageLayout, DocxTextLineLayout>(ReferenceEqualityComparer.Instance);
        var absorbedByLine = new Dictionary<DocxTextLineLayout, double>(ReferenceEqualityComparer.Instance);
        foreach (DocxInlineImageLayout image in images)
        {
            DocxTextLineLayout[] owners = lineOwners![(image.SourceParagraphIndex, image.Y)].ToArray();
            if (owners.Length != 1 || !imageOwners.TryAdd(image, owners[0])) return (lines, images);
            DocxTextLineLayout owner = owners[0];
            if (!ReferenceEquals(owner.SourceParagraph, cell.Paragraphs[image.SourceParagraphIndex!.Value]))
                return (lines, images);
            if (owner.SourceParagraph is not { } paragraph || !adjustments.TryGetValue(paragraph, out var adjustment)) continue;
            // Minimum leading and picture growth occupy the same space above
            // the baseline. A tall picture consumes that headroom only once.
            double after = owner.SourceLineIndex == byParagraph[paragraph][^1].SourceLineIndex
                ? byParagraph[paragraph][0].ParagraphAfterSpacing ?? 0d : 0d;
            double growth = Math.Max(0d, image.LineFrame!.Value.Height - (owner.LineHeight ?? 0d) - after);
            double absorbed = Math.Min(adjustment.Excess, growth);
            absorbedByLine[owner] = Math.Max(absorbedByLine.GetValueOrDefault(owner), absorbed);
        }
        double absorbedHeight = absorbedByLine.Values.Sum();
        double alignmentShift = 0d;
        // Single-cell automatic tables retain their qualified alignment rule.
        bool alignAutomaticGrid = automaticMinimumPlacement && (table.Rows.Count > 1 || row.Cells.Count > 1);
        if (cell.VerticalAlignment != DocxTableCellVerticalAlignment.Top && (!automaticMinimumPlacement || alignAutomaticGrid))
        {
            if (!automaticMinimumPlacement && (row.HeightPoints is not > 0d ||
                string.Equals(row.HeightRuleValue, "auto", StringComparison.OrdinalIgnoreCase) ||
                cell.Paragraphs.Any(p => p.Runs.Count == 0))) return (lines, images);
            double topPadding = rowTopPadding + ResolveTableCellTopBorderContentInset(cell, 1d);
            double bottomPadding = ResolveTableCellVerticalPadding(cell.Margins.BottomPoints, 1d);
            double baselineInset = ResolveTableCellFirstBaselineInset(cell.Paragraphs, measurer);
            double borderAdvance = isLastRow ? row.Cells
                .Select(c => DocxTableBorderGeometry.ResolveVisibleWidth(DocxTableBorderGeometry.Find(c.Borders, "bottom")))
                .DefaultIfEmpty(0d).Max() : 0d;
            if (!automaticMinimumPlacement && bottomPadding > 0d && marginAdvance == 0d) return (lines, images);
            double oldExtra = Math.Max(0d, cellHeight - topPadding - bottomPadding - baselineInset - usedHeight);
            double alignmentBottom = automaticMinimumPlacement
                ? row.Cells.Select(c => ResolveTableCellVerticalPadding(c.Margins.BottomPoints, 1d)).Max() : marginAdvance;
            double newExtra = Math.Max(0d, cellHeight - topPadding - alignmentBottom - borderAdvance - usedHeight + absorbedHeight);
            alignmentShift = ResolveVerticalContentOffset(cell, newExtra) - ResolveVerticalContentOffset(cell, oldExtra);
        }
        if (adjustments.Count == 0 && alignmentShift == 0d) return (lines, images);
        var shifts = images.Count == 0 ? null : new Dictionary<DocxTextLineLayout, double>(ReferenceEqualityComparer.Instance);
        double consumed = 0d;
        IReadOnlyList<DocxTextLineLayout> adjustedLines = lines.Select(line =>
        {
            double absorbed = absorbedByLine.GetValueOrDefault(line);
            consumed += absorbed;
            bool adjusted = line.SourceParagraph is { } paragraph && adjustments.TryGetValue(paragraph, out _);
            var adjustment = adjusted ? adjustments[line.SourceParagraph!] : default;
            double shift = (adjusted ? adjustment.Excess : 0d) + alignmentShift - consumed;
            shifts?.Add(line, shift);
            return adjusted ? line with
            {
                BaselineY = line.BaselineY - shift,
                BodyLineBoxBaselineInsetPoints = adjustment.Inset - absorbed,
                BodyLineBoxHeightPoints = line.LineHeight
            } : shift == 0d ? line : line with { BaselineY = line.BaselineY - shift };
        }).ToArray();
        IReadOnlyList<DocxInlineImageLayout> adjustedImages = images.Count == 0 ? images : images.Select(image =>
        {
            double shift = shifts![imageOwners[image]];
            return shift == 0d ? image : image.Translate(0d, -shift);
        }).ToArray();
        return (adjustedLines, adjustedImages);
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
            cell.Paragraphs.Any(p => !IsPlainContextualSpacingParagraph(p, allowSourceAffinedImages: true)) ||
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
        double fixedScale, double? measuredContentHeight = null, bool allowTopRow = false)
    {
        if (fixedScale != 1d || row.HeightPoints is not > 0d ||
            string.Equals(row.HeightRuleValue, "auto", StringComparison.OrdinalIgnoreCase) ||
            row.Cells.Count == 0 || row.Cells.Any(cell =>
                (cell.VerticalAlignment == DocxTableCellVerticalAlignment.Top && !allowTopRow) ||
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
